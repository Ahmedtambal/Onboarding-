import io
import re
import pandas as pd
import numpy as np
import docx
import PyPDF2

# Mapping from possible field labels in the employee form to keys in the parsed data.
FIELD_MAP = {
    "Title": "Title",
    "Full Name": "Full Name",
    "Home Address": "Home Address",
    "Home Telephone Number": "Home Telephone Number",
    "Mobile Telephone Number": "Mobile Telephone Number",
    "Telephone Number": "Telephone Number",  # Alternative if only one telephone provided.
    "Personal Email Address": "Personal Email Address",
    "Date of Birth": "Date of Birth",
    "Pronouns": "Pronouns",
    "National Insurance Number": "National Insurance Number",
    "National Insurance No.": "National Insurance Number",
    "Job Title": "Job Title",
    "Start Date": "Start Date",
    "Date Employment Commenced": "Start Date",  # Alternative label.
    "Basic Salary": "Basic Salary",
    "Pension Contribution": "Pension Contribution",
    "Marital Status": "Marital Status",
    "Nationality": "Nationality",
    "Country of Residence": "Country of Residence",
    "Name of an Emergency Contact": "Emergency Contact Name",
    "Emergency Contact Name": "Emergency Contact Name",
    "Emergency Contact Number": "Emergency Contact Number",
    "Telephone Number of Emergency Contact": "Emergency Contact Number",
    "Emergency Contact Address": "Emergency Contact Address",
    "Emergency Contact Email": "Emergency Contact Email",
    "Relationship to Emergency Contact": "Emergency Contact Relationship",
    "Employment Location Postcode": "Employment Location Postcode",
    "Notes": "Notes"
}

def parse_docx(file_bytes, debug=False):
    doc = docx.Document(io.BytesIO(file_bytes))
    lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    if debug:
         print("DEBUG: Raw DOCX lines:", lines)
    data = {}
    for i, line in enumerate(lines):
        # First, check if line starts with one of our keys
        matched = False
        for key in FIELD_MAP:
            if line.lower().startswith(key.lower()):
                potential_value = line[len(key):].strip(" :")
                if potential_value:
                    data[FIELD_MAP[key]] = potential_value
                    if debug:
                        print(f"DEBUG: Found field label on same line: '{key}', value: '{potential_value}'")
                    matched = True
                    break
        if matched:
            continue
        # Fallback: if the cleaned line exactly equals a key, take next line as value.
        clean_line = re.sub(r'\(.*?\)', '', line).strip()
        if clean_line in FIELD_MAP:
            if debug:
                 print(f"DEBUG: Found field label on separate line: '{clean_line}'")
            if i + 1 < len(lines):
                value = lines[i+1].strip()
                data[FIELD_MAP[clean_line]] = value
                if debug:
                    print(f"DEBUG: Setting {FIELD_MAP[clean_line]} = '{value}'")
    return data

def parse_pdf(file_bytes, debug=False):
    pdf_reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
    text = ""
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    if debug:
         print("DEBUG: Raw PDF lines:", lines)
    data = {}
    for i, line in enumerate(lines):
        # Check if the line starts with one of the keys (same line extraction)
        matched = False
        for key in FIELD_MAP:
            if line.lower().startswith(key.lower()):
                potential_value = line[len(key):].strip(" :")
                if potential_value:
                    data[FIELD_MAP[key]] = potential_value
                    if debug:
                        print(f"DEBUG: Found field label on same line: '{key}', value: '{potential_value}'")
                    matched = True
                    break
        if matched:
            continue
        # Fallback: if line exactly equals a key (after cleaning), use the next line.
        clean_line = re.sub(r'\(.*?\)', '', line).strip()
        if clean_line in FIELD_MAP:
            if debug:
                 print(f"DEBUG: Found field label on separate line: '{clean_line}'")
            if i + 1 < len(lines):
                value = lines[i+1].strip()
                data[FIELD_MAP[clean_line]] = value
                if debug:
                    print(f"DEBUG: Setting {FIELD_MAP[clean_line]} = '{value}'")
    return data

def load_master_file(file_obj, file_name):
    if file_name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj)
    else:
        raise ValueError("Unsupported file type. Please upload an Excel file.")
    return df

def map_employee_data(emp_data, debug=False):
    mapped = {}
    if debug:
         print("DEBUG: Mapping employee data:", emp_data)
    mapped["Title"] = emp_data.get("Title", np.nan)
    
    full_name = emp_data.get("Full Name", "").strip()
    if full_name:
        parts = full_name.split()
        mapped["First Name"] = parts[0]
        mapped["Surname"] = parts[-1] if len(parts) > 1 else np.nan
    else:
        mapped["First Name"] = np.nan
        mapped["Surname"] = np.nan

    mapped["Legal Gender"] = np.nan  # Not provided in the form.
    mapped["Marital Status"] = emp_data.get("Marital Status", np.nan)

    home_addr = emp_data.get("Home Address", "")
    addr_parts = []
    if home_addr:
        addr_parts = [part.strip() for part in home_addr.split("\n") if part.strip()]
        if len(addr_parts) == 1:
            addr_parts = [part.strip() for part in home_addr.split(",") if part.strip()]
    mapped["Address 1"] = addr_parts[0] if len(addr_parts) >= 1 else np.nan
    mapped["Address 2"] = addr_parts[1] if len(addr_parts) >= 2 else np.nan
    mapped["Address 3"] = addr_parts[2] if len(addr_parts) >= 3 else np.nan
    mapped["Address 4"] = addr_parts[3] if len(addr_parts) >= 4 else np.nan
    mapped["Post Code"] = addr_parts[-1] if len(addr_parts) >= 5 else np.nan

    mapped["Date of Birth"] = emp_data.get("Date of Birth", np.nan)
    mapped["NI Number"] = emp_data.get("National Insurance Number", np.nan)
    mapped["Start Date"] = emp_data.get("Start Date", np.nan)
    mapped["Job Title"] = emp_data.get("Job Title", np.nan)
    mapped["Basic Annual Salary"] = emp_data.get("Basic Salary", np.nan)
    mapped["Nationality"] = emp_data.get("Nationality", np.nan)
    mapped["Email Address"] = emp_data.get("Personal Email Address", np.nan)
    mapped["Any Other Information"] = emp_data.get("Notes", np.nan)
    
    if debug:
         print("DEBUG: Mapped data:", mapped)
    return mapped

def append_employee_record(df, emp_data, debug=False):
    mapped_data = map_employee_data(emp_data, debug=debug)
    master_columns = [
        "Title", "First Name", "Surname", "Legal Gender", "Marital Status",
        "Address 1", "Address 2", "Address 3", "Address 4", "Post Code",
        "Date of Birth", "NI Number", "Start Date", "Job Title",
        "Basic Annual Salary", "Nationality", "Email Address", "Any Other Information"
    ]
    new_record = {col: mapped_data.get(col, np.nan) for col in master_columns}
    if debug:
         print("DEBUG: New record to append:", new_record)
    new_row_df = pd.DataFrame([new_record])
    df = pd.concat([df, new_row_df], ignore_index=True)
    return df

def export_master_file(df, file_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    file_ext = "xlsx"
    output.seek(0)
    return output, mime, file_ext
