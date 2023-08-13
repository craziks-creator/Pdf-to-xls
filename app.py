import os
from flask import Flask, request, render_template, send_file
from pdf_processor import extract_questions_and_options, save_to_excel

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

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
        output_excel_path = os.path.join(app.config["UPLOAD_FOLDER"], "output_questions.xlsx")
        save_to_excel(questions, options, output_excel_path)

        # Provide the Excel file for download
        return send_file(output_excel_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    app.run(debug=True)
