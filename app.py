
import json
from flask import Flask, render_template, request, send_from_directory, jsonify
import os

from app.services.compare_logic import *
from app.services.pdf import generate_pdf_report

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", results=None)

# @app.route("/process", methods=["POST"])
# def process():
#     uploaded_pairs = []
#     indices = set()
    
#     for key in request.files.keys():
#         if key.startswith("actual_") or key.startswith("expected_"):
#             try:
#                 idx = int(key.split("_")[-1])
#                 indices.add(idx)
#             except ValueError:
#                 pass
#     indices = sorted(indices)

#     try:
#         for i in indices:
#             actual_file = request.files.get(f"actual_{i}")
#             expected_file = request.files.get(f"expected_{i}")

#             if not actual_file or not expected_file:
#                 continue

#             if not allowed_filename(actual_file.filename) or not allowed_filename(expected_file.filename):
#                 continue

#             # Run comparison
#             comparison_results = compare_excel_stats(actual_file, expected_file)
            
#             # Save JSON report
#             report_filename = f"report_{actual_file.filename.split('.')[0]}_VS_{expected_file.filename.split('.')[0]}.json"
#             report_path = os.path.join(REPORT_FOLDER, report_filename)
            
#             with open(report_path, "w", encoding="utf-8") as rf:
#                 json.dump(comparison_results, rf, indent=2)

#             uploaded_pairs.append({
#                 "report_file": report_filename,
#                 "pair": f"{actual_file.filename} vs {expected_file.filename}",
#                 "results": comparison_results
#             })

#         return jsonify(uploaded_pairs)
#     except Exception as e:
#         return jsonify({"error": str(e)})

# @app.route("/download/<folder>/<filename>")
# def download_file(folder, filename):
#     if folder not in ("uploads", "reports"):
#         return "Not allowed", 403
#     directory = REPORT_FOLDER
#     return send_from_directory(directory, filename, as_attachment=True)


# Update your Flask routes
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
            json_report_filename = f"report_{actual_file.filename.split('.')[0]}_VS_{expected_file.filename.split('.')[0]}.json"
            json_report_path = os.path.join(REPORT_FOLDER, json_report_filename)
            
            with open(json_report_path, "w", encoding="utf-8") as rf:
                json.dump(comparison_results, rf, indent=2)
            
            # Generate PDF report
            pdf_report_filename = f"report_{actual_file.filename.split('.')[0]}_VS_{expected_file.filename.split('.')[0]}.pdf"
            pdf_report_path = os.path.join(REPORT_FOLDER, pdf_report_filename)
            
            pdf_success = generate_pdf_report(comparison_results, pdf_report_path)
            
            uploaded_pairs.append({
                "report_file": pdf_report_filename if pdf_success else json_report_filename,
                "json_report_file": json_report_filename,
                "pdf_report_file": pdf_report_filename if pdf_success else None,
                "pair": f"{actual_file.filename} vs {expected_file.filename}",
                "results": comparison_results,
                "has_pdf": pdf_success
            })

        return jsonify(uploaded_pairs)
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/download/<folder>/<filename>")
def download_file(folder, filename):
    if folder not in ("uploads", "reports"):
        return "Not allowed", 403
    directory = REPORT_FOLDER
    
    # Determine content type based on file extension
    if filename.lower().endswith('.pdf'):
        mimetype = 'application/pdf'
    elif filename.lower().endswith('.json'):
        mimetype = 'application/json'
    else:
        mimetype = 'text/plain'
    
    return send_from_directory(directory, filename, as_attachment=True, mimetype=mimetype)

if __name__ == "__main__":
    app.run(debug=True)
