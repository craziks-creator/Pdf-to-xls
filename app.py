import os
import fitz  # PyMuPDF
import pandas as pd
from flask import Flask, request, render_template, send_file

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_questions_and_options(pdf_path):
    questions = []
    options = []
    
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text = page.get_text()
        lines = text.split('\n')
        
        question = ""
        option_group = []
        is_question = True
        
        for line in lines:
            stripped_line = line.strip()
            if not stripped_line:
                continue
            
            if is_question:
                question += stripped_line + " "
                is_question = False
            else:
                option_group.append(stripped_line)
            
            # Assuming that an empty line indicates the end of an option group
            if not stripped_line:
                questions.append(question)
                options.append(option_group)
                question = ""
                option_group = []
                is_question = True
    
    pdf_document.close()
    return questions, options

def save_to_excel(questions, options, output_excel_path):
    data = []
    for i in range(len(questions)):
        data.append({"Question": questions[i], "Options": options[i]})
    
    df = pd.DataFrame(data)
    df.to_excel(output_excel_path, index=False)

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
