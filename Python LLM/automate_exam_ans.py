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

# response json to text 
def extract_text_from_response(api_response):
    """
    Extracts the 'text' part from the given API response.

    :param api_response: A dictionary containing the API response.
    :return: The extracted text if found, otherwise None.
    """
    try:
        # Navigate through the nested structure to extract the text
        candidates = api_response.get('candidates', [])
        if candidates:
            content = candidates[0].get('content', {})
            parts = content.get('parts', [])
            if parts:
                return parts[0].get('text', "").strip()  # Strip to remove trailing newline or whitespace
        return None
    except (KeyError, IndexError):
        return None



# Function to call Google Gemini API
def get_answers_from_gemini(prompt):
    api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=AIzaSyAWCcJyC8xRozqpPDGEzL7KTRTFThkiRSY"  # Replace with actual API key
    headers = {
        "Content-Type": "application/json"
    }
    data = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ]
    }
    try:
        response = requests.post(api_url, headers=headers, json=data)
        response.raise_for_status()  # Raise an error for bad responses
        response_data = response.json()
        # Assuming the API response contains a field for generated content
        # return response_data.get('generatedContent', 'No answer generated.')
        response_data = response.json()
        print("API Response:", response_data)  # Debugging: print the full response
        text_data = extract_text_from_response(response_data)
        print(text_data)
        # return response_data.get(text_data, 'No answer generated.')
        return text_data

    except requests.exceptions.RequestException as e:
        print(f"Error calling the Gemini API: {e}")
        return "Error generating answer."



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
            print(f"Processing {filename}...")
            text = extract_text_from_pdf(pdf_path)
            # text = "Who is won 2011 cricket world cup?"
            print(text)
            answers = get_answers_from_gemini(text)
            print(answers)
            output_filename = f"{os.path.splitext(filename)[0]}_ANSWERS.docx"
            output_path = os.path.join(input_folder, output_filename)
            create_word_document(output_path, answers)
            print(f"Processed {filename} -> {output_filename}")

# Specify the folder containing the exam papers
input_folder = 'QuestionPaper'  # Change this to your folder path
process_exam_papers(input_folder)
