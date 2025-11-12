
import json
from flask import Flask, render_template, request, send_from_directory, jsonify
import os

from compare_logic import *

app = Flask(__name__)

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