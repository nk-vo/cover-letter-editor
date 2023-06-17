import json
import os
import docx2pdf
from docx import Document
from docx.shared import Pt
import webbrowser
from dotenv import load_dotenv
load_dotenv()

# Load the existing company names from the JSON file
json_file = os.getenv('JSON_FILE')
if os.path.exists(json_file):
    with open(json_file) as f:
        data = json.load(f)
    last_entry = data[-1]
    default_old_name = last_entry['new_name']
else:
    default_old_name = ''

# Get the PDF file name from the environment variable
pdf_dir, pdf_file_name = os.path.split(os.getenv('PDF_FILE'))

while True:
    # Prompt the user to enter the old and new company names
    old_name = input("Enter the old company name (or press Enter to use the last new company name: {} or \'exit\' to exit): ".format(default_old_name))
    if not old_name:
        old_name = default_old_name
    elif old_name.lower() == 'exit':
        break
    new_name = input("Enter the new company name: ")

    # Open the Word document
    doc = Document(os.getenv('DOCX_FILE'))

    # Create a dictionary to store the old and new company names
    company_names = {"old_name": old_name, "new_name": new_name}

    # Find and replace the old company name with the new company name
    for paragraph in doc.paragraphs:
        if old_name in paragraph.text:
            for run in paragraph.runs:
                font = run.font
                font_size = font.size
                font_name = font.name
                bold = font.bold
                italic = font.italic
                underline = font.underline
                strike = font.strike
                run.text = run.text.replace(old_name, new_name)
                run.font.name = font_name
                run.font.size = font_size
                run.bold = bold
                run.italic = italic
                run.underline = underline
                run.strike = strike

    # Save the updated Word document
    doc.save(os.getenv('DOCX_FILE'))

    # Convert the Word document to PDF
    docx_file = os.getenv('DOCX_FILE')
    pdf_file = os.path.join(pdf_dir, f'{pdf_file_name.split(".")[0]}_{new_name.replace(" ", "")}.pdf')
    docx2pdf.convert(docx_file, pdf_file)

    # Store the old and new company names in a JSON file
    if os.path.exists(json_file):
        with open(json_file) as f:
            data = json.load(f)
        data.append(company_names)
    else:
        data = [company_names]

    with open(json_file, 'w') as f:
        json.dump(data, f)

    # Ask the user if they want to open the new PDF file
    open_file = input("Do you want to open the new PDF file? (y/n)")
    if open_file.lower() == 'y':
        webbrowser.open(pdf_file)
