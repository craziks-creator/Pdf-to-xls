import os
from flask import Flask, request, render_template, send_file
import fitz  # PyMuPDF
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_questions_and_options(pdf_path):
    # Implement the same extract_questions_and_options function from the previous response

def save_to_excel(questions, options, output_excel_path):
    # Implement the same save_to_excel function from the previous response

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Check if a file was uploaded
        if "file" not in request.files:
            return render_template("index.html", error="No file part")

        file = request.files["file"]

        # Check if a file was uploaded and it has an allowed extension
        if file.filename == "" or not allowed_file(file.filename):
            return render_template("index.html", error="Invalid file")

        # Save the uploaded file
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(file_path)

        # Extract questions and options
        questions, options = extract_questions_and_options(file_path)

        # Create Excel file
        output_excel_path = "output_questions.xlsx"
        save_to_excel(questions, options, output_excel_path)

        # Provide the Excel file for download
        return send_file(output_excel_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    app.run(debug=True)
