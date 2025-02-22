import io
import re
import pandas as pd
import numpy as np
import docx
import PyPDF2
import datetime

# =========================
# 1) Field Map
# =========================
FIELD_MAP = {
    "Title": "Title",
    "Full Name": "Full Name",
    "Home Address": "Home Address",
    "Home Telephone Number": "Home Telephone Number",
    "Mobile Telephone Number": "Mobile Telephone Number",
    "Telephone Number": "Telephone Number",
    "Personal Email Address": "Personal Email Address",
    "Date of Birth": "Date of Birth",
    "Pronouns": "Pronouns",
    "National Insurance Number": "National Insurance Number",
    "National Insurance No.": "National Insurance Number",
    "Job Title": "Job Title",
    "Start Date": "Start Date",
    "Date Employment Commenced": "Start Date",
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

# =========================
# 2) Parsing DOCX & PDF
# =========================

def parse_docx(file_bytes, debug=False):
    """
    Parse a DOCX file, extracting text line by line.
    If a line starts with a known key, we take the rest of the line as the value.
    If a line exactly matches a key (and no value on same line), 
    we look at the next line for the value.
    """
    doc = docx.Document(io.BytesIO(file_bytes))
    lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    if debug:
        print("DEBUG: Raw DOCX lines:", lines)
    data = {}

    for i, line in enumerate(lines):
        found_key = False
        for key in FIELD_MAP:
            # Check if line starts with key (case-insensitive)
            if line.lower().startswith(key.lower()):
                # Extract the remainder of the line
                potential_value = line[len(key):].strip(" :")
                if potential_value:
                    # We have "Key: Value" on the same line
                    data[FIELD_MAP[key]] = potential_value
                    if debug:
                        print(f"DEBUG: Found '{key}' on same line -> {potential_value}")
                    found_key = True
                    break
                else:
                    # The line might exactly match the key, with the value on the next line
                    if line.strip().lower() == key.lower():
                        if (i + 1) < len(lines):
                            fallback_value = lines[i + 1].strip()
                            data[FIELD_MAP[key]] = fallback_value
                            if debug:
                                print(f"DEBUG: Found '{key}' on separate line -> {fallback_value}")
                        found_key = True
                        break
        # If not found_key, just move on to the next line
    return data

def parse_pdf(file_bytes, debug=False):
    """
    Parse a PDF file, extracting text line by line.
    If a line starts with a known key, we take the rest of the line as the value.
    If a line exactly matches a key (and no value on same line), 
    we look at the next line for the value.
    """
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
        found_key = False
        for key in FIELD_MAP:
            # Check if line starts with key (case-insensitive)
            if line.lower().startswith(key.lower()):
                # Extract the remainder of the line
                potential_value = line[len(key):].strip(" :")
                if potential_value:
                    # We have "Key: Value" on the same line
                    data[FIELD_MAP[key]] = potential_value
                    if debug:
                        print(f"DEBUG: Found '{key}' on same line -> {potential_value}")
                    found_key = True
                    break
                else:
                    # The line might exactly match the key, with the value on the next line
                    if line.strip().lower() == key.lower():
                        if (i + 1) < len(lines):
                            fallback_value = lines[i + 1].strip()
                            data[FIELD_MAP[key]] = fallback_value
                            if debug:
                                print(f"DEBUG: Found '{key}' on separate line -> {fallback_value}")
                        found_key = True
                        break
        # If not found_key, just move on to the next line
    return data

# =========================
# 3) Load Master File
# =========================

def load_master_file(file_obj, file_name):
    """
    Loads an Excel file (XLS or XLSX) into a Pandas DataFrame,
    stripping whitespace from column names to avoid duplicates.
    """
    if file_name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj)
        df.columns = df.columns.str.strip()  # remove leading/trailing spaces
        return df
    else:
        raise ValueError("Unsupported file type. Please upload an Excel file.")

# =========================
# 4) Robust Date Parsing
# =========================

def remove_ordinal_suffixes(s: str) -> str:
    """
    Remove ordinal suffixes like 'st', 'nd', 'rd', 'th' from day numbers.
    E.g. '1st March 2025' -> '1 March 2025'
    """
    pattern = r'(\d+)(st|nd|rd|th)\b'
    return re.sub(pattern, r'\1', s, flags=re.IGNORECASE)

def fix_common_numeric_typos(s: str) -> str:
    """
    Attempt to fix common typos in numeric contexts:
      - 'O' or 'o' -> '0'
      - 'l' or 'I' -> '1'
    But only when they appear between digits or date delimiters.
    """
    text = s.lower()
    # Replace 'o' with '0' if between digits or date delimiters
    text = re.sub(r'(?<=[0-9./\- ])o(?=[0-9./\- ])', '0', text)
    # Replace 'l' or 'i' with '1' if between digits or date delimiters
    text = re.sub(r'(?<=[0-9./\- ])[li](?=[0-9./\- ])', '1', text)
    return text

def fix_missing_slash_between_month_and_year(s: str) -> str:
    """
    If the user typed something like '01/031987' (missing slash before year),
    insert a slash to make it '01/03/1987'.
    We'll detect patterns like ^DD/MMYYYY$ or ^D/MYYYY$.
    """
    pattern = r'^(\d{1,2})/(\d{1,2})(\d{4})$'
    replacement = r'\1/\2/\3'
    return re.sub(pattern, replacement, s)

def robust_parse_date_str(date_str: str) -> pd.Timestamp:
    """
    Attempt to parse a variety of date formats, correcting typical user typos.
    Returns NaT if parsing fails.
    """
    s = date_str.strip()

    # 1) Remove ordinal suffixes like '1st', '2nd', '3rd', '4th'...
    s = remove_ordinal_suffixes(s)
    # 2) Fix common numeric typos (e.g. 'o' -> '0', 'l' -> '1' in numeric context)
    s = fix_common_numeric_typos(s)
    # 3) Fix missing slash between month and year (e.g. '01/031987' -> '01/03/1987')
    s = fix_missing_slash_between_month_and_year(s)

    # 4) Finally, parse with dayfirst=True and coerce errors to NaT
    parsed = pd.to_datetime(s, errors='coerce', dayfirst=True)
    return parsed if not pd.isnull(parsed) else pd.NaT

# =========================
# 5) Map Employee Data
# =========================

def map_employee_data(emp_data, debug=False):
    """
    Maps the extracted data into a consistent set of fields
    and uses robust date parsing for Date of Birth and Start Date.
    """
    if debug:
        print("DEBUG: Mapping employee data:", emp_data)

    mapped = {}
    mapped["Title"] = emp_data.get("Title", np.nan)

    # Split Full Name into first/last
    full_name = emp_data.get("Full Name", "").strip()
    if full_name:
        parts = full_name.split()
        mapped["First Name"] = parts[0]
        mapped["Surname"] = parts[-1] if len(parts) > 1 else np.nan
    else:
        mapped["First Name"] = np.nan
        mapped["Surname"] = np.nan

    mapped["Legal Gender"] = np.nan
    mapped["Marital Status"] = emp_data.get("Marital Status", np.nan)

    # Address
    home_addr = emp_data.get("Home Address", "")
    addr_parts = [part.strip() for part in home_addr.split("\n") if part.strip()]
    if len(addr_parts) == 1:
        addr_parts = [x.strip() for x in addr_parts[0].split(",") if x.strip()]
    mapped["Address 1"] = addr_parts[0] if len(addr_parts) >= 1 else np.nan
    mapped["Address 2"] = addr_parts[1] if len(addr_parts) >= 2 else np.nan
    mapped["Address 3"] = addr_parts[2] if len(addr_parts) >= 3 else np.nan
    mapped["Address 4"] = addr_parts[3] if len(addr_parts) >= 4 else np.nan
    mapped["Post Code"] = addr_parts[-1] if len(addr_parts) >= 5 else np.nan

    # Date of Birth
    dob = robust_parse_date_str(emp_data.get("Date of Birth", ""))
    mapped["Date of Birth"] = dob if not pd.isnull(dob) else np.nan

    # National Insurance
    mapped["NI Number"] = emp_data.get("National Insurance Number", np.nan)

    # Start Date
    start_date_raw = emp_data.get("Start Date", "")
    start_date_parsed = robust_parse_date_str(start_date_raw)
    mapped["Start Date"] = start_date_parsed if not pd.isnull(start_date_parsed) else np.nan

    mapped["Job Title"] = emp_data.get("Job Title", np.nan)
    mapped["Basic Annual Salary"] = emp_data.get("Basic Salary", np.nan)
    mapped["Nationality"] = emp_data.get("Nationality", np.nan)
    mapped["Email Address"] = emp_data.get("Personal Email Address", np.nan)
    mapped["Any Other Information"] = emp_data.get("Notes", np.nan)

    if debug:
        print("DEBUG: Mapped data:", mapped)
    return mapped

# =========================
# 6) Append Employee Record
# =========================

def append_employee_record(df, emp_data, debug=False):
    """
    Append a new row with automatically generated:
      - Id (increment from current max or start at 1 if none)
      - Start time (now)
      - Completion time (now)
    and mapped employee data including 'Start Date' from "Date Employment Commenced".
    """
    mapped_data = map_employee_data(emp_data, debug=debug)

    # 1) Generate new Id
    if 'Id' in df.columns and not df['Id'].isnull().all():
        max_id = df['Id'].max()
        if pd.isnull(max_id):
            max_id = 0
        new_id = max_id + 1
    else:
        # If 'Id' col doesn't exist or is empty
        new_id = 1
        if 'Id' not in df.columns:
            df['Id'] = np.nan

    # 2) Current timestamps for Start/Completion
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Build the new row dict
    new_record = {
        "Id": new_id,
        "Start time": now_str,
        "Completion time": now_str
    }

    # Add mapped data to new_record
    for key, val in mapped_data.items():
        new_record[key] = val

    # The columns we want to ensure exist
    master_columns = [
        "Id", "Start time", "Completion time",
        "Title", "First Name", "Surname", "Legal Gender", "Marital Status",
        "Address 1", "Address 2", "Address 3", "Address 4", "Post Code",
        "Date of Birth", "NI Number", "Start Date", "Job Title",
        "Basic Annual Salary", "Nationality", "Email Address", "Any Other Information"
    ]

    # Ensure these columns exist in df
    for col in master_columns:
        if col not in df.columns:
            df[col] = np.nan

    # Convert new_record into a 1-row DataFrame
    new_row_df = pd.DataFrame([new_record])

    # Append it
    df = pd.concat([df, new_row_df], ignore_index=True)

    return df

# =========================
# 7) Export Master File
# =========================

def export_master_file(df, file_name):
    """
    Export DataFrame to Excel with a consistent date/time format
    for the 'Start Date', 'Start time', and 'Completion time' columns.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

        # Apply formatting
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd HH:MM:SS"})

        for col_name in ["Start Date", "Start time", "Completion time"]:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name)
                worksheet.set_column(col_idx, col_idx, 25, date_format)

    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    file_ext = "xlsx"
    output.seek(0)
    return output, mime, file_ext
