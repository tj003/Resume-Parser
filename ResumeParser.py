import re
import spacy
from spacy.matcher import Matcher
from datetime import datetime
from dateutil import relativedelta
import pandas as pd
from resume_parser import resumeparse
from docx import Document
import os
from pdfminer.high_level import extract_text as extract_pdf_text
from docx import Document
import time
import pandas as pd
import textract
import re
from datetime import datetime
from dateutil import relativedelta
import spacy
from spacy.matcher import Matcher
from openpyxl import Workbook
from openpyxl.styles import Font
from resume_parser import resumeparse

import PyPDF2

nlp = spacy.load('en_core_web_sm')

# Initialize spacy and matcher
nlp = spacy.load('en_core_web_sm')
matcher = Matcher(nlp.vocab)

def extract_text(file_path):
    if file_path.endswith('.pdf'):
        return extract_text_from_pdf_pypdf2(file_path)
    elif file_path.endswith('.docx'):
        return extract_text_from_docx(file_path)
    elif file_path.endswith('.doc'):
        return extract_text_from_doc(file_path)
    else:
        raise ValueError("Unsupported file format")

def extract_text_from_pdf_pypdf2(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        for page in range(len(reader.pages)):
            text += reader.pages[page].extract_text()
        return text

def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        full_text = [para.text for para in doc.paragraphs]
        text = '\n'.join(full_text).strip()
        if not text:
            print(f"No text found in DOCX {file_path}. Likely an image-based DOCX.")
            return None
        return text
    except Exception as e:
        print(f"Error processing DOCX file {file_path}: {e}")
        return None

def extract_text_from_doc(file_path):
    try:
        text = textract.process(file_path).decode('utf-8').strip()
        if not text:
            print(f"No text found in DOC {file_path}. Likely an image-based DOC.")
            return None
        return text
    except Exception as e:
        print(f"Error processing DOC file {file_path}: {e}")
        return None


# Custom model functions
def extract_email(text):
    email = re.findall(r"([^@|\s]+@[^@]+\.[^@|\s]+)", text)
    if email:
        return email[0].split()[0].strip(';')

def extract_name(nlp_text, matcher):
    patterns = [
       [{'POS': 'PROPN'}, {'POS': 'PROPN'}],  # First name and Last name
       [{'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}],  # First, Middle, and Last name
       [{'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}, {'POS': 'PROPN'}]  # First, two Middle, and Last name
    ]

    matcher.add('NAME', patterns=patterns)

    doc = nlp(nlp_text)
    matches = matcher(doc)

    for match_id, start, end in matches:
        span = doc[start:end]
        return span.text

    return None

def extract_mobile_number(text, custom_regex=None):
    if not custom_regex:
        mob_num_regex = r'''(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)
                        [-\.\s]*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})'''
        phone = re.findall(re.compile(mob_num_regex), text)
    else:
        phone = re.findall(re.compile(custom_regex), text)
    if phone:
        number = ''.join(phone[0])
        return number

def extract_skills(nlp_text, noun_chunks, skills_file=None):
    tokens = [token.text for token in nlp_text if not token.is_stop]
    if not skills_file:
        data = pd.read_csv('/content/skills.csv')
    else:
        data = pd.read_csv(skills_file)
    skills = list(data.columns.values)
    skillset = []
    for token in tokens:
        if token.lower() in skills:
            skillset.append(token)

    for token in noun_chunks:
        token = token.text.lower().strip()
        if token in skills:
            skillset.append(token)
    return [i.capitalize() for i in set([i.lower() for i in skillset])]

def extract_education(nlp_text):
    education = []
    try:
        resume_text = nlp_text.text.lower()
        for edu in ["Bachelor", "Master", "PhD", "High School", "B.Sc", "M.Sc"]:  # Example educational degrees
            if edu.lower() in resume_text:
                education.append(edu)
    except Exception as e:
        print(e)
    return education

def extract_experience_from_summary(resume_text):
    summary_exp_match = re.search(r'(\d+(\.\d+)?)(\+|-)?\s*year(s)?\s*of\s*experience', resume_text, re.IGNORECASE)
    if summary_exp_match:
        return float(summary_exp_match.group(1))
    return None

def extract_experience_section(resume_text):
    experience_pattern = r'(PROFESSIONAL SUMMARY.*?|WORK EXPERIENCE.*?|PROFESSIONAL EXPERIENCE.*?|EXPERIENCE.*?)(?:EDUCATION|CERTIFICATION|SKILLS|PROJECTS|AWARDS|ACTIVITIES|LANGUAGES|REFERENCES|$)'
    experience_match = re.search(experience_pattern, resume_text, re.IGNORECASE | re.DOTALL)
    if experience_match:
        return experience_match.group(1).strip()
    return None

def parse_date(date_str):
    date_str = re.sub(r'\b(\d+)(st|nd|rd|th)\b', r'\1', date_str)
    date_formats = ['%d %b %Y', '%b %Y', '%m/%Y', '%Y']
    for fmt in date_formats:
        try:
            parsed_date = datetime.strptime(date_str.replace('.', ''), fmt)
            return parsed_date
        except ValueError:
            continue
    raise ValueError(f"Date format for {date_str} not recognized")

def extract_experience(resume_text):
    if not resume_text:
        return 0

    total_experience = 0
    date_patterns = [
        r'(\b(?:\d{1,2}\s)?(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)|Apr(?:il)|May|Jun(?:e)|Jul(?:y)|Aug(?:ust)|Sep(?:tember)|Oct(?:ober)|Nov(?:ember)?|Dec(?:ember)?)(?:\.|\s)?\s+\d{4})\s*[-–—to]+\s*(\b(?:\d{1,2}\s)?(?:Jan(?:uary)?|Feb(?:ruary)|Mar(?:ch)|Apr(?:il)|May|Jun(?:e)|Jul(?:y)|Aug(?:ust)|Sep(?:tember)|Oct(?:ober)|Nov(?:ember)?|Dec(?:ember)?)(?:\.|\s)?\s+\d{4}|\bPresent\b|\bTill\b|\bTill date\b)',
        r'(\b\d{1,2}/\d{4})\s*[-–—to]+\s*(\b\d{1,2}/\d{4}|\bPresent\b|\bTill\b|\bTill date\b)',
        r'(\b\d{4})\s*[-–—to]+\s*(\b\d{4}|\bPresent\b|\bTill\b|\bTill date\b)',
        r'(\b(?:\d{1,2}(st|nd|rd|th)?\s)?\b(?:Jan(?:uary)?|Feb(?:ruary)|Mar(?:ch)|Apr(?:il)|May|Jun(?:e)|Jul(?:y)|Aug(?:ust)|Sep(?:tember)|Oct(?:ober)|Nov(?:ember)?)\s+\d{4})\s*(to|[-–—])\s*(\b(?:\d{1,2}(st|nd|rd|th)?\s)?(?:Jan(?:uary)|Feb(?:ruary)|Mar(?:ch)|Apr(?:il)|May|Jun(?:e)|Jul(?:y)|Aug(?:ust)|Sep(?:tember)|Oct(?:ober)|Nov(?:ember)?)\s+\d{4}|\bPresent\b|\bTill\b|\bTill date\b)'
    ]

    experience_entries = []
    for pattern in date_patterns:
        matches = re.findall(pattern, resume_text, re.IGNORECASE)
        experience_entries.extend(matches)

    for entry in experience_entries:
        start_date_str, end_date_str = entry[0], entry[1]
        try:
            start_date = parse_date(start_date_str)
            end_date = parse_date(end_date_str) if 'present' not in end_date_str.lower() and 'till' not in end_date_str.lower() else datetime.now()
            experience = relativedelta.relativedelta(end_date, start_date).years + (relativedelta.relativedelta(end_date, start_date).months / 12)
            total_experience += experience
        except ValueError as e:
            print(f"Error parsing dates: {start_date_str} - {end_date_str}, {e}")

    summary_experience = extract_experience_from_summary(resume_text)
    if summary_experience is not None:
        return summary_experience

    return round(total_experience, 1)

# Custom extraction logic if certain fields are empty
def map_parsed_data(parsed_data, resume_text):
    mapped_data = {
        "Name": parsed_data.get('name') or extract_name(resume_text, matcher),
        "Email": parsed_data.get('email') or extract_email(resume_text),
        "Contact Number": parsed_data.get('phone') or extract_mobile_number(resume_text),
        "Skills": parsed_data.get('skills') or extract_skills(nlp(resume_text), nlp(resume_text).noun_chunks),
        "Education": parsed_data.get('degree') or extract_education(nlp(resume_text)),
        "Address": None,  # Implement a custom function if needed
        "Total Work Experience": parsed_data.get('total_exp') or extract_experience(resume_text),
        "Work Summary": extract_experience_section(resume_text),
        "Current Company": parsed_data.get('Companies worked at')[0] if parsed_data.get('Companies worked at') else None,
        "Current Designation": parsed_data.get('designition')[0] if parsed_data.get('designition') else None
    }

    return mapped_data

def save_to_excel(data, output_file='parsed_resume_data.xlsx'):
    # Convert data to DataFrame
    df = pd.DataFrame([data])
    
    # Save to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        for col in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in col)
            max_length = max(max_length, len(col[0].value))
            worksheet.column_dimensions[col[0].column_letter].width = max_length + 2

# Example usage:
file_path = r'D:\Language_Processing-master\Language_Processing-master\resumes\01 Resume Mrinalika.pdf'
resume_text = extract_text(file_path)  # Extract text from the resume
if resume_text:
    data = resumeparse.read_file(file_path)
    mapped_data = map_parsed_data(data, resume_text)
    print(mapped_data)
    
    # Save the mapped data to an Excel file
    save_to_excel(mapped_data, output_file='parsed_resume_data.xlsx')
else:
    print("Failed to extract text from the resume.")
