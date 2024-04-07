import os
import docx2txt
import PyPDF2
from openpyxl import Workbook

# Function to extract text from a Word document
def extract_text_from_docx(file_path):
    text = docx2txt.process(file_path)
    return text

# Function to extract text from a PDF
def extract_text_from_pdf(file_path):
    text = ""
    with open(file_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text()
    return text

# Function to extract email ID and contact number from text
def extract_contact_info(text):
    # Code to extract email ID and contact number from text
    # You can use regular expressions or any other method to extract this information
    # For demonstration purpose, I'm leaving this part as pseudocode
    email = "example@example.com"
    contact_number = "1234567890"
    return email, contact_number

# Function to process CV files and extract information
def process_cv_files(cv_folder):
    cv_data = []
    for filename in os.listdir(cv_folder):
        file_path = os.path.join(cv_folder, filename)
        if filename.endswith(".docx"):
            text = extract_text_from_docx(file_path)
        elif filename.endswith(".pdf"):
            text = extract_text_from_pdf(file_path)
        else:
            continue

        email, contact_number = extract_contact_info(text)
        cv_data.append({"File": filename, "Email": email, "Contact Number": contact_number, "Text": text})

    return cv_data

# Function to create Excel file from extracted data
def create_excel(cv_data):
    wb = Workbook()
    ws = wb.active
    ws.append(["File", "Email", "Contact Number", "Text"])
    for cv in cv_data:
        ws.append([cv["File"], cv["Email"], cv["Contact Number"], cv["Text"]])
    excel_file_path = "cv_data.xlsx"
    wb.save(excel_file_path)
    return excel_file_path

# Sample usage
# cv_folder = "C:/Users/HP/Desktop/Sample2/AarushiRohatgi.pdf"
# cv_folder = "C:/Users/HP/Desktop/Sample2/"
cv_folder = "C:/Users/HP/Desktop/Amit"




cv_data = process_cv_files(cv_folder)
excel_file_path = create_excel(cv_data)
print("Excel file created:", excel_file_path)
