import fitz  # PyMuPDF
import pandas as pd

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
