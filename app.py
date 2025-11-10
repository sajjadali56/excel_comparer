from flask import Flask, render_template, request
import os
from compare_logic import compare_excel_stats

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["REPORT_FOLDER"] = "reports"

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["REPORT_FOLDER"], exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded_pairs = []
        for i in range(len(request.files)//2):
            actual = request.files.get(f"actual_{i}")
            expected = request.files.get(f"expected_{i}")
            if actual and expected:
                actual_path = os.path.join(app.config["UPLOAD_FOLDER"], actual.filename)
                expected_path = os.path.join(app.config["UPLOAD_FOLDER"], expected.filename)
                actual.save(actual_path)
                expected.save(expected_path)
                results = compare_excel_stats(actual_path, expected_path)
                uploaded_pairs.append({
                    "pair": actual.filename + " vs " + expected.filename,
                    "results": results
                })

        return render_template("index.html", results=uploaded_pairs)
    return render_template("index.html")
    
if __name__ == "__main__":
    app.run(debug=True)
