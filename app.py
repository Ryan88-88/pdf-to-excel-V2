from flask import Flask, render_template, request, send_file
import os
from process_pdf import process_pdf_file

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files["file"]

    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(OUTPUT_FOLDER, "final_output.xlsx")

    file.save(pdf_path)

    process_pdf_file(pdf_path, output_path)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run()
