import os
from json import JSONEncoder, JSONDecoder
from flask import Flask, render_template, request, send_from_directory, jsonify

from compare_logic import compare_excel_stats

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_FOLDER = os.path.join(BASE_DIR, "reports")

os.makedirs(REPORT_FOLDER, exist_ok=True)

ALLOWED_EXT = {".xls", ".xlsx"}

def allowed_filename(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", results=None)

@app.route("/process", methods=["POST"])
def process():
    uploaded_pairs = []

    # Determine indices dynamically from request.files keys (handles gaps or variable counts)
    indices = set()
    print(request.files.keys())
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
                # skip incomplete pair
                continue

            # basic allowed check
            if not allowed_filename(actual_file.filename) or not allowed_filename(expected_file.filename):
                # skip or handle error; here we'll skip this pair
                continue

            

            # Create unique filenames to avoid collisions
            actual_orig, _ = os.path.splitext(actual_file.filename.lower())
            expected_orig, _ = os.path.splitext(expected_file.filename.lower())

            # Run comparison and save human-readable report
            report_contents = compare_excel_stats(actual_file, expected_file)

            # Create a plain text report file summarizing results
            report_filename = f"report_{actual_file.filename.split('.')[0]} VS {expected_file.filename.split('.')[0]}.txt"
            report_path = os.path.join(REPORT_FOLDER, report_filename)

            # If compare_excel_stats returns the same text structure you had earlier, convert to text:
            with open(report_path, "w", encoding="utf-8") as rf:
                rf.write(f"Comparison: {actual_orig}  VS  {expected_orig}\n\n")
                if not report_contents:
                    rf.write("No output from compare function.\n")
                else:
                    # report_contents might be a list of dicts with 'sheet' and 'message'
                    for entry in report_contents:
                        sheet = entry.get("sheet", "Sheet")
                        msg = entry.get("message", "")
                        rf.write(f"--- {sheet} ---\n")
                        rf.write(msg + "\n\n")

            uploaded_pairs.append({
                "report_file": report_filename,
                "pair": f"{actual_orig} vs {expected_orig}",
                "results": report_contents

            })

        return jsonify(uploaded_pairs)
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/download/<folder>/<filename>")
def download_file(folder, filename):
    # restrict to our two folders for safety
    if folder not in ("uploads", "reports"):
        return "Not allowed", 403
    directory = REPORT_FOLDER
    # Security: ensure file exists and prevent path traversal by using send_from_directory
    return send_from_directory(directory, filename, as_attachment=True, mimetype="text/plain", download_name=filename)

if __name__ == "__main__":
    app.run(debug=True)
