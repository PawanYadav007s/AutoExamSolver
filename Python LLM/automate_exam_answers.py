import os
import requests
import PyPDF2
from docx import Document

# Function to read PDF and extract text
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text() + '\n'
    return text

# Function to call Google Gemini API
def get_answers_from_gemini(prompt):
    api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?"  # Replace with actual API endpoint
    headers = {
        "Authorization": "AIzaSyAWCcJyC8xRozqpPDGEzL7KTRTFThkiRSY",  # Replace with your API key
        "Content-Type": "application/json"
    }
    data = {
        "prompt": prompt,
        "max_tokens": 1000  # Adjust as needed
    }
    response = requests.post(api_url, headers=headers, json=data)
    return response.json().get('answer', '')


# Function to create a Word document with answers
def create_word_document(output_path, answers):
    doc = Document()
    doc.add_paragraph(answers)
    doc.save(output_path)

# Main function to process all exam papers
def process_exam_papers(input_folder):
    for filename in os.listdir(input_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(input_folder, filename)
            text = extract_text_from_pdf(pdf_path)
            answers = get_answers_from_gemini(text)
            output_filename = f"{os.path.splitext(filename)[0]}_ANSWERS.docx"
            output_path = os.path.join(input_folder, output_filename)
            create_word_document(output_path, answers)
            print(f"Processed {filename} -> {output_filename}")

# Specify the folder containing the exam papers
input_folder = 'QuestionPaper'  # Change this to your folder path
process_exam_papers(input_folder)
