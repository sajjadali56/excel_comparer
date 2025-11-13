
import json
from flask import Flask, render_template, request, send_from_directory, jsonify
import os

from waitress import serve

from app.formatter import format_comparison_results
from app.services.compare_logic import *
from app.services.pdf import generate_pdf_report

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", results=None)

@app.route("/process", methods=["POST"])
def process():
    start_time = time.time()
    uploaded_pairs = []
    indices = set()
    
    logger.info("Starting file processing request")
    
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
            pair_start_time = time.time()
            actual_file = request.files.get(f"actual_{i}")
            expected_file = request.files.get(f"expected_{i}")

            if not actual_file or not expected_file:
                logger.warning(f"Pair {i}: Incomplete file pair")
                continue

            if not allowed_filename(actual_file.filename) or not allowed_filename(expected_file.filename):
                logger.warning(f"Pair {i}: Invalid file types")
                continue

            logger.info(f"Processing pair {i}: {actual_file.filename} vs {expected_file.filename}")

            # Run comparison
            comparison_results = compare_excel_stats(actual_file, expected_file)
            
            # Generate unique base name for reports
            base_name = f"report_{actual_file.filename.split('.')[0]}_VS_{expected_file.filename.split('.')[0]}"
            
            # Save JSON report only if needed
            json_report_filename = f"{base_name}.json"
            json_report_path = os.path.join(REPORT_FOLDER, json_report_filename)
            
            with open(json_report_path, "w", encoding="utf-8") as rf:
                json.dump(comparison_results, rf, indent=2)
            
            # Generate PDF report asynchronously or in background if needed
            pdf_report_filename = f"{base_name}.pdf"
            pdf_report_path = os.path.join(REPORT_FOLDER, pdf_report_filename)
            
            pdf_success = generate_pdf_report(comparison_results, pdf_report_path)
            
            pair_data = {
                "report_file": pdf_report_filename if pdf_success else json_report_filename,
                "json_report_file": json_report_filename,
                "pdf_report_file": pdf_report_filename if pdf_success else None,
                "pair": f"{actual_file.filename} vs {expected_file.filename}",
                "results": format_comparison_results(comparison_results),
                "has_pdf": pdf_success
            }
            
            uploaded_pairs.append(pair_data)
            logger.info(f"Pair {i} completed in {time.time() - pair_start_time:.2f}s")

        total_time = time.time() - start_time
        logger.info(f"All pairs processed in {total_time:.2f}s")
        return jsonify(uploaded_pairs)
        
    except Exception as e:
        logger.error(f"Process route error: {str(e)}")
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

def run_browser(message):
    print(message)
    # app.run(debug=True, use_reloader=True)  # Run with debug mode
    serve(app, host='127.0.0.1', port=5000) # Run with production mode 


if __name__ == "__main__":

    PRODUCTION = True
    message = "Welcome to the tool. Please access the tool using the link: http://127.0.0.1:5000/"

    if PRODUCTION:
        run_browser(message)
    else:
        app.run(debug=True)
