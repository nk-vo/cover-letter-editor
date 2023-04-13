## Description

This Python script allows you to replace company names in a Microsoft Word document and save the updated document as a PDF file with the new company name. It also stores the old and new company names in a JSON file for future reference.

## Requirements

- Python 3
- The `docx2pdf`, `python-docx`, `webbrowser`, and `dotenv` packages

## Installation

1. Clone the repository or download the script file
2. Install the required packages by running `pip install docx2pdf python-docx webbrowser python-dotenv`

## Usage

1. Set the environment variables in a `.env` file in the same directory as the script:
```
DOCX_FILE='YOUR_INPUT_DOCX_FILEPATH'
PDF_FILE='YOUR_OUTPUT_PDF_FILEPATH'
JSON_FILE='YOUR_OUTPUT_JSON_FILEPATH'
```
2. Run the script by running `python replace_company_names.py` in the command line
3. When prompted, enter the old company name and the new company name
4. The script will replace all instances of the old company name with the new company name in the Word document, save the updated document, convert it to a PDF file with the new company name, and store the old and new company names in the JSON file
5. If you choose to open the new PDF file when prompted, it will be opened in your default PDF viewer
6. The script will continue to run until you enter "exit" as the old company name

## License

This script is released under the MIT License.
