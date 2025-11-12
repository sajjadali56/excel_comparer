import pandas as pd
import json
from flask import Flask, render_template, request, send_from_directory, jsonify
import os
from datetime import datetime

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_FOLDER = os.path.join(BASE_DIR, "reports")
os.makedirs(REPORT_FOLDER, exist_ok=True)

ALLOWED_EXT = {".xls", ".xlsx"}

def allowed_filename(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT

def compare_excel_stats(file1, file2):
    print("Comparing files:", file1.filename, file2.filename)
    
    # Save files temporarily for processing
    temp_file1 = os.path.join(REPORT_FOLDER, f"temp1_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    temp_file2 = os.path.join(REPORT_FOLDER, f"temp2_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    file1.save(temp_file1)
    file2.save(temp_file2)
    
    try:
        xl1 = pd.ExcelFile(temp_file1)
        xl2 = pd.ExcelFile(temp_file2)

        common_sheets = sorted(set(xl1.sheet_names).intersection(xl2.sheet_names))
        print("Common sheets:", common_sheets)

        force_object_cols = {"UW_Year", "Loss_Period"}
        
        comparison_results = {
            "file1_name": file1.filename,
            "file2_name": file2.filename,
            "comparison_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_sheets": len(common_sheets),
            "sheets": []
        }

        if not common_sheets:
            comparison_results["error"] = "No common sheets found"
            return comparison_results

        for sheet in common_sheets:
            df1 = xl1.parse(sheet)
            df2 = xl2.parse(sheet)
            common_cols = list(df1.columns.intersection(df2.columns))
            
            sheet_data = {
                "sheet_name": sheet,
                "total_columns": len(common_cols),
                "matching_columns": 0,
                "different_columns": 0,
                "columns": []
            }

            if not common_cols:
                sheet_data["error"] = "No common columns found"
                comparison_results["sheets"].append(sheet_data)
                continue

            df1 = df1[common_cols]
            df2 = df2[common_cols]

            for col in common_cols:
                col1 = df1[col]
                col2 = df2[col]
                col_data = {
                    "name": col,
                    "type": "numeric" if (pd.api.types.is_numeric_dtype(col1) and pd.api.types.is_numeric_dtype(col2)) and col not in force_object_cols else "text",
                    "status": "matching",
                    "differences": []
                }

                if col in force_object_cols or not (
                    pd.api.types.is_numeric_dtype(col1)
                    and pd.api.types.is_numeric_dtype(col2)
                ):
                    # Text column comparison
                    vc1 = col1.astype(str).value_counts().sort_index()
                    vc2 = col2.astype(str).value_counts().sort_index()
                    combined_index = vc1.index.union(vc2.index)

                    diffs = [
                        {"value": str(val), "file1_count": int(vc1.get(val, 0)), "file2_count": int(vc2.get(val, 0))}
                        for val in combined_index
                        if vc1.get(val, 0) != vc2.get(val, 0)
                    ]

                    if diffs:
                        col_data["status"] = "different"
                        col_data["differences"] = diffs[:10]  # Limit to first 10 differences
                        sheet_data["different_columns"] += 1
                    else:
                        sheet_data["matching_columns"] += 1
                else:
                    # Numeric column comparison
                    col1 = pd.to_numeric(col1, errors="coerce")
                    col2 = pd.to_numeric(col2, errors="coerce")

                    stats = {
                        "sum": {"file1": float(col1.sum()), "file2": float(col2.sum())},
                        "mean": {"file1": float(col1.mean()), "file2": float(col2.mean())},
                        "min": {"file1": float(col1.min()), "file2": float(col2.min())},
                        "max": {"file1": float(col1.max()), "file2": float(col2.max())},
                        "std": {"file1": float(col1.std()), "file2": float(col2.std())}
                    }

                    # Check for differences
                    differences_found = []
                    for stat, values in stats.items():
                        v1, v2 = values["file1"], values["file2"]
                        if stat in ["mean", "sum"] and round(v1, 4) != round(v2, 4):
                            differences_found.append({
                                "statistic": stat,
                                "file1_value": round(v1, 4),
                                "file2_value": round(v2, 4),
                                "difference": round(v2 - v1, 4)
                            })
                        elif stat in ["min", "max"] and v1 != v2:
                            differences_found.append({
                                "statistic": stat,
                                "file1_value": v1,
                                "file2_value": v2,
                                "difference": round(v2 - v1, 4)
                            })

                    if differences_found:
                        col_data["status"] = "different"
                        col_data["differences"] = differences_found
                        col_data["statistics"] = stats
                        sheet_data["different_columns"] += 1
                    else:
                        sheet_data["matching_columns"] += 1
                        col_data["statistics"] = stats

                sheet_data["columns"].append(col_data)

            comparison_results["sheets"].append(sheet_data)

        return comparison_results

    except Exception as e:
        return {"error": str(e)}
    finally:
        # Clean up temporary files
        try:
            os.remove(temp_file1)
            os.remove(temp_file2)
        except:
            pass

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", results=None)

@app.route("/process", methods=["POST"])
def process():
    uploaded_pairs = []
    indices = set()
    
    for key in request.files.keys():
        if key.startswith("actual_") or key.startswith("expected_"):
            try:
                idx = int(key.split("_")[-1])
                indices.add(idx)
            except ValueError:
                pass
    indices = sorted(indices)

    try:
        for i in indices:
            actual_file = request.files.get(f"actual_{i}")
            expected_file = request.files.get(f"expected_{i}")

            if not actual_file or not expected_file:
                continue

            if not allowed_filename(actual_file.filename) or not allowed_filename(expected_file.filename):
                continue

            # Run comparison
            comparison_results = compare_excel_stats(actual_file, expected_file)
            
            # Save JSON report
            report_filename = f"report_{actual_file.filename.split('.')[0]}_VS_{expected_file.filename.split('.')[0]}.json"
            report_path = os.path.join(REPORT_FOLDER, report_filename)
            
            with open(report_path, "w", encoding="utf-8") as rf:
                json.dump(comparison_results, rf, indent=2)

            uploaded_pairs.append({
                "report_file": report_filename,
                "pair": f"{actual_file.filename} vs {expected_file.filename}",
                "results": comparison_results
            })

        return jsonify(uploaded_pairs)
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/download/<folder>/<filename>")
def download_file(folder, filename):
    if folder not in ("uploads", "reports"):
        return "Not allowed", 403
    directory = REPORT_FOLDER
    return send_from_directory(directory, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)