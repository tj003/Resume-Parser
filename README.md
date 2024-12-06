###Resume Parsing Project
#Overview
This project is designed to automate the extraction of key information from resumes. The script can process multiple resumes from a specified input folder, extract relevant details such as name, contact information, skills, education, work experience, and more, and save the parsed data into an Excel file.

Features
Bulk Resume Processing: Process multiple resumes in various formats (.pdf, .docx, .doc) from a specified folder.
Custom Extraction: Extracts key details such as name, email, contact number, skills, education, work experience, current company, and designation.
Save to Excel: Parsed data is saved into an Excel file with each resume's information neatly organized into columns.
Error Handling: Includes robust error handling for unsupported formats and potential issues during text extraction.

Requirements
Python 3.7 or later
Required Python packages (install via pip):

pip install spacy
pip install pdfminer.six
pip install PyPDF2
pip install python-docx
pip install textract
pip install pandas
pip install openpyxl
pip install python-dateutil
