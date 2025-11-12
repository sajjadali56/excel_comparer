import pandas as pd
import numpy as np
import os
from datetime import datetime
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_FOLDER = os.path.join(BASE_DIR, "reports")
os.makedirs(REPORT_FOLDER, exist_ok=True)

ALLOWED_EXT = {".xls", ".xlsx"}

def allowed_filename(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT

def safe_parse_excel(file_path, sheet_name):
    """Safely parse an Excel sheet with comprehensive error handling"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        # Handle completely empty sheets
        if df.empty:
            return pd.DataFrame()
        # Handle sheets with only headers
        if len(df) == 0:
            return df
        return df
    except Exception as e:
        print(f"Warning: Could not parse sheet '{sheet_name}' from {file_path}: {str(e)}")
        return pd.DataFrame()

def safe_column_comparison(col1, col2, col_name, force_object_cols):
    """Safely compare two columns with comprehensive error handling"""
    try:
        col_data = {
            "name": col_name,
            "type": "unknown",
            "status": "error",
            "differences": [],
            "error": None
        }
        
        # Handle empty columns
        if (col1.empty and col2.empty):
            col_data.update({
                "type": "empty",
                "status": "matching",
                "error": "Both columns are empty"
            })
            return col_data
        elif col1.empty or col2.empty:
            col_data.update({
                "type": "empty",
                "status": "different",
                "error": f"One column is empty (col1: {len(col1)} rows, col2: {len(col2)} rows)"
            })
            return col_data

        # Handle completely null columns
        if col1.isna().all() and col2.isna().all():
            col_data.update({
                "type": "null",
                "status": "matching",
                "error": "Both columns contain only null values"
            })
            return col_data
        elif col1.isna().all() or col2.isna().all():
            col_data.update({
                "type": "null",
                "status": "different",
                "error": "One column contains only null values"
            })
            return col_data

        # Determine column type with error handling
        is_numeric = False
        try:
            is_numeric = (pd.api.types.is_numeric_dtype(col1) and 
                         pd.api.types.is_numeric_dtype(col2) and 
                         col_name not in force_object_cols)
        except Exception:
            is_numeric = False

        col_data["type"] = "numeric" if is_numeric else "text"

        if col_name in force_object_cols or not is_numeric:
            # Text column comparison with error handling
            try:
                # Handle mixed data types by converting to string
                vc1 = col1.astype(str).value_counts().sort_index()
                vc2 = col2.astype(str).value_counts().sort_index()
                combined_index = vc1.index.union(vc2.index)

                diffs = []
                for val in combined_index:
                    try:
                        count1 = int(vc1.get(val, 0))
                        count2 = int(vc2.get(val, 0))
                        if count1 != count2:
                            diffs.append({
                                "value": str(val),
                                "file1_count": count1,
                                "file2_count": count2
                            })
                    except (ValueError, TypeError):
                        continue

                if diffs:
                    col_data.update({
                        "status": "different",
                        "differences": diffs[:10]  # Limit to first 10 differences
                    })
                else:
                    col_data.update({
                        "status": "matching"
                    })
                    
            except Exception as e:
                col_data.update({
                    "status": "error",
                    "error": f"Text comparison failed: {str(e)}"
                })
                
        else:
            # Numeric column comparison with error handling
            try:
                # Safe conversion to numeric
                col1_numeric = pd.to_numeric(col1, errors='coerce')
                col2_numeric = pd.to_numeric(col2, errors='coerce')
                
                # Handle cases where conversion fails for all values
                if col1_numeric.isna().all() or col2_numeric.isna().all():
                    col_data.update({
                        "type": "text",  # Fall back to text comparison
                        "status": "error",
                        "error": "Failed to convert to numeric values"
                    })
                    return col_data

                # Calculate statistics with safe defaults
                stats = {}
                differences_found = []
                
                try:
                    stats = {
                        "sum": {
                            "file1": float(col1_numeric.sum()) if not col1_numeric.isna().all() else 0,
                            "file2": float(col2_numeric.sum()) if not col2_numeric.isna().all() else 0
                        },
                        "mean": {
                            "file1": float(col1_numeric.mean()) if not col1_numeric.isna().all() else 0,
                            "file2": float(col2_numeric.mean()) if not col2_numeric.isna().all() else 0
                        },
                        "min": {
                            "file1": float(col1_numeric.min()) if not col1_numeric.isna().all() else 0,
                            "file2": float(col2_numeric.min()) if not col2_numeric.isna().all() else 0
                        },
                        "max": {
                            "file1": float(col1_numeric.max()) if not col1_numeric.isna().all() else 0,
                            "file2": float(col2_numeric.max()) if not col2_numeric.isna().all() else 0
                        },
                        "std": {
                            "file1": float(col1_numeric.std()) if not col1_numeric.isna().all() and len(col1_numeric) > 1 else 0,
                            "file2": float(col2_numeric.std()) if not col2_numeric.isna().all() and len(col2_numeric) > 1 else 0
                        }
                    }
                    
                    # Check for differences with tolerance for floating point errors
                    for stat, values in stats.items():
                        v1, v2 = values["file1"], values["file2"]
                        
                        # Handle NaN and Inf values
                        if np.isnan(v1) or np.isnan(v2) or np.isinf(v1) or np.isinf(v2):
                            differences_found.append({
                                "statistic": stat,
                                "file1_value": v1,
                                "file2_value": v2,
                                "difference": "NaN/Inf detected"
                            })
                            continue
                            
                        if stat in ["mean", "sum"]:
                            # Use relative tolerance for large numbers, absolute for small
                            if abs(v1) > 1 or abs(v2) > 1:
                                relative_diff = abs(v1 - v2) / max(abs(v1), abs(v2))
                                if relative_diff > 1e-10:  # 0.00000001% tolerance
                                    differences_found.append({
                                        "statistic": stat,
                                        "file1_value": round(v1, 6),
                                        "file2_value": round(v2, 6),
                                        "difference": round(v2 - v1, 6)
                                    })
                            else:
                                if abs(v1 - v2) > 1e-10:  # Absolute tolerance for small numbers
                                    differences_found.append({
                                        "statistic": stat,
                                        "file1_value": round(v1, 6),
                                        "file2_value": round(v2, 6),
                                        "difference": round(v2 - v1, 6)
                                    })
                        elif stat in ["min", "max"]:
                            if abs(v1 - v2) > 1e-10:
                                differences_found.append({
                                    "statistic": stat,
                                    "file1_value": v1,
                                    "file2_value": v2,
                                    "difference": round(v2 - v1, 6)
                                })
                                
                except Exception as stats_error:
                    col_data.update({
                        "status": "error",
                        "error": f"Statistical calculation failed: {str(stats_error)}"
                    })
                    return col_data

                if differences_found:
                    col_data.update({
                        "status": "different",
                        "differences": differences_found,
                        "statistics": stats
                    })
                else:
                    col_data.update({
                        "status": "matching",
                        "statistics": stats
                    })
                    
            except Exception as e:
                col_data.update({
                    "status": "error",
                    "error": f"Numeric comparison failed: {str(e)}"
                })

        return col_data

    except Exception as e:
        return {
            "name": col_name,
            "type": "unknown",
            "status": "error",
            "differences": [],
            "error": f"Unexpected error in column comparison: {str(e)}"
        }

def compare_excel_stats(file1, file2):
    print("Comparing files:", file1.filename, file2.filename)
    
    # Save files temporarily for processing
    temp_file1 = os.path.join(REPORT_FOLDER, f"temp1_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    temp_file2 = os.path.join(REPORT_FOLDER, f"temp2_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    
    try:
        file1.save(temp_file1)
        file2.save(temp_file2)
    except Exception as e:
        return {"error": f"Failed to save temporary files: {str(e)}"}

    try:
        # Safely load Excel files
        try:
            xl1 = pd.ExcelFile(temp_file1, engine='openpyxl')
            xl2 = pd.ExcelFile(temp_file2, engine='openpyxl')
        except Exception as e:
            return {"error": f"Failed to read Excel files: {str(e)}"}

        # Get sheet names with error handling
        try:
            sheets1 = xl1.sheet_names
            sheets2 = xl2.sheet_names
            common_sheets = sorted(set(sheets1).intersection(sheets2))
        except Exception as e:
            return {"error": f"Failed to read sheet names: {str(e)}"}

        print("Common sheets:", common_sheets)

        force_object_cols = {"UW_Year", "Loss_Period"}
        
        comparison_results = {
            "file1_name": file1.filename,
            "file2_name": file2.filename,
            "comparison_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_sheets": len(common_sheets),
            "sheets_processed": 0,
            "sheets_failed": 0,
            "sheets": []
        }

        if not common_sheets:
            comparison_results["warning"] = "No common sheets found between files"
            return comparison_results

        for sheet in common_sheets:
            sheet_data = {
                "sheet_name": sheet,
                "status": "processed",
                "error": None,
                "total_columns": 0,
                "matching_columns": 0,
                "different_columns": 0,
                "error_columns": 0,
                "columns": []
            }

            try:
                # Safely parse sheets
                df1 = safe_parse_excel(temp_file1, sheet)
                df2 = safe_parse_excel(temp_file2, sheet)
                
                if df1.empty or df2.empty:
                    sheet_data.update({
                        "status": "warning",
                        "error": "One or both sheets are empty",
                        "total_columns": 0
                    })
                    comparison_results["sheets"].append(sheet_data)
                    continue

                # Get common columns safely
                try:
                    common_cols = list(df1.columns.intersection(df2.columns))
                    sheet_data["total_columns"] = len(common_cols)
                except Exception as e:
                    sheet_data.update({
                        "status": "error",
                        "error": f"Failed to identify common columns: {str(e)}"
                    })
                    comparison_results["sheets"].append(sheet_data)
                    continue

                if not common_cols:
                    sheet_data.update({
                        "status": "warning",
                        "error": "No common columns found in this sheet"
                    })
                    comparison_results["sheets"].append(sheet_data)
                    continue

                # Process each column with individual error handling
                for col in common_cols:
                    try:
                        col1 = df1[col]
                        col2 = df2[col]
                        
                        col_result = safe_column_comparison(col1, col2, col, force_object_cols)
                        
                        # Update counters based on column status
                        if col_result.get("status") == "matching":
                            sheet_data["matching_columns"] += 1
                        elif col_result.get("status") == "different":
                            sheet_data["different_columns"] += 1
                        elif col_result.get("status") == "error":
                            sheet_data["error_columns"] += 1
                            
                        sheet_data["columns"].append(col_result)
                        
                    except Exception as col_error:
                        # Individual column failure shouldn't stop the entire process
                        error_col_data = {
                            "name": col,
                            "type": "unknown",
                            "status": "error",
                            "differences": [],
                            "error": f"Column processing failed: {str(col_error)}"
                        }
                        sheet_data["columns"].append(error_col_data)
                        sheet_data["error_columns"] += 1
                        print(f"Warning: Failed to process column '{col}' in sheet '{sheet}': {str(col_error)}")

                comparison_results["sheets_processed"] += 1
                
            except Exception as sheet_error:
                # Sheet-level error - continue with next sheet
                sheet_data.update({
                    "status": "error",
                    "error": f"Sheet processing failed: {str(sheet_error)}",
                    "total_columns": 0,
                    "columns": []
                })
                comparison_results["sheets_failed"] += 1
                print(f"Warning: Failed to process sheet '{sheet}': {str(sheet_error)}")

            comparison_results["sheets"].append(sheet_data)

        # Add summary statistics
        total_columns = sum(sheet["total_columns"] for sheet in comparison_results["sheets"])
        matching_columns = sum(sheet["matching_columns"] for sheet in comparison_results["sheets"])
        different_columns = sum(sheet["different_columns"] for sheet in comparison_results["sheets"])
        error_columns = sum(sheet["error_columns"] for sheet in comparison_results["sheets"])
        
        comparison_results["summary"] = {
            "total_columns_compared": total_columns,
            "matching_columns": matching_columns,
            "different_columns": different_columns,
            "error_columns": error_columns,
            "success_rate": round((matching_columns + different_columns) / total_columns * 100, 2) if total_columns > 0 else 0
        }

        return comparison_results

    except Exception as e:
        error_trace = traceback.format_exc()
        print(f"Critical error in comparison: {error_trace}")
        return {"error": f"Unexpected error during comparison: {str(e)}"}

    finally:
        # Clean up temporary files
        try:
            if os.path.exists(temp_file1):
                os.remove(temp_file1)
            if os.path.exists(temp_file2):
                os.remove(temp_file2)
        except Exception as cleanup_error:
            print(f"Warning: Failed to clean up temporary files: {cleanup_error}")