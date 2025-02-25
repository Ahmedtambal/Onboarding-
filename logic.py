import io
import re
import pandas as pd
import numpy as np
import docx2txt
import docx
import PyPDF2
import datetime

# =========================
# 1) Field Map (for non-Excel employee files)
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
# 2) Parsing Employee Files (DOCX, PDF, CSV/TXT, Excel)
# =========================

def parse_docx(file_bytes, debug=False):
    # Use docx2txt to extract all text from the DOCX file
    text = docx2txt.process(io.BytesIO(file_bytes))
    # Split the text into lines and remove empty ones
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    if debug:
        print("DEBUG: Raw DOCX lines:", lines)
    data = {}
    for i, line in enumerate(lines):
        for key in FIELD_MAP:
            if line.lower().startswith(key.lower()):
                potential_value = line[len(key):].strip(" :")
                if potential_value:
                    data[FIELD_MAP[key]] = potential_value
                    if debug:
                        print(f"DEBUG: Found '{key}' on same line -> {potential_value}")
                    break
                else:
                    if line.strip().lower() == key.lower() and (i + 1) < len(lines):
                        fallback_value = lines[i + 1].strip()
                        data[FIELD_MAP[key]] = fallback_value
                        if debug:
                            print(f"DEBUG: Found '{key}' on separate line -> {fallback_value}")
                        break
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
        for key in FIELD_MAP:
            if line.lower().startswith(key.lower()):
                potential_value = line[len(key):].strip(" :")
                if potential_value:
                    data[FIELD_MAP[key]] = potential_value
                    if debug:
                        print(f"DEBUG: Found '{key}' on same line -> {potential_value}")
                    break
                else:
                    if line.strip().lower() == key.lower() and (i + 1) < len(lines):
                        fallback_value = lines[i + 1].strip()
                        data[FIELD_MAP[key]] = fallback_value
                        if debug:
                            print(f"DEBUG: Found '{key}' on separate line -> {fallback_value}")
                        break
    return data

def parse_csv_employee(file_bytes, debug=False):
    text = file_bytes.decode('utf-8')
    try:
        df = pd.read_csv(io.StringIO(text))
    except Exception as e:
        if debug:
            print("DEBUG: Error parsing CSV/TXT employee file:", e)
        return {}
    if len(df) > 0:
        row = df.iloc[0].to_dict()
        if debug:
            print("DEBUG: Parsed CSV/TXT employee row:", row)
        return row
    return {}

# New: Map Excel employee row according to the desired master mapping.
def map_excel_employee_data(row, debug=False):
    mapped = {}
    mapped["Surname*"] = row.get("Surname", np.nan)
    mapped["FirstName*"] = row.get("First Name", np.nan)
    mapped["SchemeRef*"] = np.nan
    mapped["CategoryName"] = np.nan
    mapped["Title"] = row.get("Title", np.nan)
    mapped["AddressLine1"] = row.get("Address 1", np.nan)
    mapped["AddressLine2"] = row.get("Address 2", np.nan)
    mapped["AddressLine3"] = row.get("Address 3", np.nan)
    mapped["AddressLine4"] = row.get("Address 4", np.nan)
    mapped["CityTown"] = row.get("City", np.nan)
    mapped["County"] = row.get("county", np.nan)
    mapped["Country"] = row.get("Country of Residence", np.nan)
    mapped["PostCode"] = row.get("Post Code", np.nan)
    mapped["AdviceType*"] = row.get("AdviceType", np.nan)
    # Robustly parse dates from Excel (convert to string first)
    mapped["DateJoinedScheme"] = robust_parse_date_str(str(row.get("Start Date", "")))
    mapped["DateofBirth*"] = robust_parse_date_str(str(row.get("Date of Birth", "")))
    mapped["EmailAddress"] = row.get("Email Address", np.nan)
    mapped["Gender"] = row.get("Legal Gender", np.nan)
    mapped["HomeNumber"] = row.get("Home Telephone Number", np.nan)
    mapped["MobileNumber"] = row.get("Mobile Telephone Number", np.nan)
    mapped["NINumber"] = row.get("NI Number", np.nan)
    mapped["PensionableSalary"] = row.get("Basic Annual Salary", np.nan)
    mapped["PensionableSalaryStartDate"] = mapped["DateJoinedScheme"]
    mapped["SalaryPostSacrifice"] = np.nan
    mapped["PolicyNumber"] = np.nan
    mapped["SellingAdviserId*"] = np.nan
    mapped["SplitTemplateGroupName"] = np.nan
    mapped["SplitTemplateGroupSource"] = np.nan
    mapped["ServiceStatus"] = np.nan
    mapped["ClientCategory"] = np.nan
    if debug:
        print("DEBUG: Mapped Excel row:", mapped)
    return mapped

# New: Parse Excel employee file, returning a list of employee data dictionaries.
def parse_excel_employee(file_bytes, debug=False):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        df.columns = df.columns.str.strip()  # Remove extra spaces from column names
    except Exception as e:
        if debug:
            print("DEBUG: Error parsing Excel employee file:", e)
        return []
    emp_data_list = []
    for index, row in df.iterrows():
        row_dict = row.to_dict()
        mapped = map_excel_employee_data(row_dict, debug=debug)
        emp_data_list.append(mapped)
    return emp_data_list



# =========================
# 3) Load Master File (Excel, CSV, or TXT)
# =========================

def load_master_file(file_obj, file_name):
    if file_name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_obj)
    elif file_name.lower().endswith((".csv", ".txt")):
        df = pd.read_csv(file_obj)
    else:
        raise ValueError("Unsupported master file type. Please upload an Excel, CSV, or TXT file.")
    df.columns = df.columns.str.strip()
    return df

# =========================
# 4) Robust Date Parsing
# =========================

def remove_ordinal_suffixes(s: str) -> str:
    pattern = r'(\d+)(st|nd|rd|th)\b'
    return re.sub(pattern, r'\1', s, flags=re.IGNORECASE)

def fix_common_numeric_typos(s: str) -> str:
    text = s.lower()
    text = re.sub(r'(?<=[0-9./\- ])o(?=[0-9./\- ])', '0', text)
    text = re.sub(r'(?<=[0-9./\- ])[li](?=[0-9./\- ])', '1', text)
    return text

def fix_missing_slash_between_month_and_year(s: str) -> str:
    pattern = r'^(\d{1,2})/(\d{1,2})(\d{4})$'
    replacement = r'\1/\2/\3'
    return re.sub(pattern, replacement, s)
def robust_parse_date_str(date_str) -> object:
    # If the input is not a string, try to convert it directly.
    if not isinstance(date_str, str):
        try:
            parsed = pd.to_datetime(date_str, errors='coerce', dayfirst=True)
            return parsed if not pd.isnull(parsed) else str(date_str)
        except Exception:
            return str(date_str)
    # Handle strings like "Timestamp('2000-02-01 00:00:00')"
    if date_str.startswith("Timestamp("):
        inner = date_str[len("Timestamp("):].rstrip(")")
        inner = inner.replace("'", "").replace('"', "")
        try:
            parsed = pd.to_datetime(inner)
            return parsed if not pd.isnull(parsed) else date_str
        except Exception:
            return date_str
    s = date_str.strip()
    s = remove_ordinal_suffixes(s)
    s = fix_common_numeric_typos(s)
    s = fix_missing_slash_between_month_and_year(s)
    parsed = pd.to_datetime(s, errors='coerce', dayfirst=True)
    return parsed if not pd.isnull(parsed) else s



# =========================
# 5) Map Employee Data (for non-Excel files)
# =========================

def safe_str(val):
    if isinstance(val, str):
        return val
    elif pd.isnull(val):
        return ""
    else:
        return str(val)

def map_employee_data(emp_data, debug=False):
    if debug:
        print("DEBUG: Mapping employee data:", emp_data)
    mapped = {}
    full_name = safe_str(emp_data.get("Full Name", "") or emp_data.get("Name", "")).strip()
    if full_name:
        parts = full_name.split()
        mapped["FirstName*"] = parts[0]
        mapped["Surname*"] = " ".join(parts[1:]) if len(parts) > 1 else np.nan
    else:
        mapped["FirstName*"] = np.nan
        mapped["Surname*"] = np.nan
    mapped["SchemeRef*"] = np.nan
    mapped["CategoryName"] = np.nan
    mapped["Title"] = safe_str(emp_data.get("Title", "")).strip() or np.nan
    home_addr = safe_str(emp_data.get("Home Address", "") or emp_data.get("Address", "")).strip()
    addr_parts = [part.strip() for part in home_addr.split("\n") if part.strip()]
    if len(addr_parts) == 1:
        addr_parts = [x.strip() for x in addr_parts[0].split(",") if x.strip()]
    mapped["AddressLine1"] = addr_parts[0] if len(addr_parts) >= 1 else np.nan
    mapped["AddressLine2"] = addr_parts[1] if len(addr_parts) >= 2 else np.nan
    mapped["AddressLine3"] = addr_parts[2] if len(addr_parts) >= 3 else np.nan
    mapped["AddressLine4"] = addr_parts[3] if len(addr_parts) >= 4 else np.nan
    mapped["CityTown"] = np.nan
    mapped["County"] = np.nan
    mapped["Country"] = np.nan
    mapped["PostCode"] = addr_parts[4] if len(addr_parts) >= 5 else np.nan
    mapped["AdviceType*"] = np.nan

    # Corrected: Use emp_data instead of row.
    mapped["DateJoinedScheme"] = robust_parse_date_str(safe_str(emp_data.get("Start Date", "")))
    dob_raw = safe_str(emp_data.get("Date of Birth", "") or emp_data.get("DOB", "")).strip()
    dob = robust_parse_date_str(dob_raw)
    mapped["DateofBirth*"] = dob if not pd.isnull(dob) else np.nan

    mapped["EmailAddress"] = safe_str(emp_data.get("Personal Email Address", "") or emp_data.get("Email", "")).strip() or np.nan
    mapped["Gender"] = safe_str(emp_data.get("Gender", "")).strip() or np.nan
    home_num = safe_str(emp_data.get("Home Telephone Number", "") or emp_data.get("Telephone Number", "")).strip()
    mapped["HomeNumber"] = home_num if home_num != "" else np.nan
    mapped["MobileNumber"] = safe_str(emp_data.get("Mobile Telephone Number", "")).strip() or np.nan
    mapped["NINumber"] = safe_str(emp_data.get("National Insurance Number", "") or emp_data.get("NI Number", "")).strip() or np.nan
    mapped["PensionableSalary"] = safe_str(emp_data.get("Basic Salary", "")).strip() or np.nan
    mapped["PensionableSalaryStartDate"] = mapped["DateJoinedScheme"]
    mapped["SalaryPostSacrifice"] = np.nan
    mapped["PolicyNumber"] = np.nan
    mapped["SellingAdviserId*"] = np.nan
    mapped["SplitTemplateGroupName"] = np.nan
    mapped["SplitTemplateGroupSource"] = np.nan
    mapped["ServiceStatus"] = np.nan
    mapped["ClientCategory"] = np.nan

    if debug:
        print("DEBUG: Mapped data:", mapped)
    return mapped


# =========================
# 6) Append Employee Record to Master DataFrame
# =========================

def append_employee_record(df, emp_data, debug=False):
    # If the employee data is already mapped (e.g. from an Excel file), don't re-map it.
    if "Surname*" in emp_data:
        mapped_data = emp_data
    else:
        mapped_data = map_employee_data(emp_data, debug=debug)
    master_columns = [
        "Surname*", "FirstName*", "SchemeRef*", "CategoryName", "Title",
        "AddressLine1", "AddressLine2", "AddressLine3", "AddressLine4",
        "CityTown", "County", "Country", "PostCode", "AdviceType*",
        "DateJoinedScheme", "DateofBirth*", "EmailAddress", "Gender",
        "HomeNumber", "MobileNumber", "NINumber", "PensionableSalary",
        "PensionableSalaryStartDate", "SalaryPostSacrifice", "PolicyNumber",
        "SellingAdviserId*", "SplitTemplateGroupName", "SplitTemplateGroupSource",
        "ServiceStatus", "ClientCategory"
    ]
    for col in master_columns:
        if col not in df.columns:
            df[col] = np.nan
    new_row_df = pd.DataFrame([mapped_data])
    df = pd.concat([df, new_row_df], ignore_index=True)
    return df

# =========================
# 7) Export Master File
# =========================

def export_master_file(df, file_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        for col_name in ["DateJoinedScheme", "DateofBirth*"]:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name)
                worksheet.set_column(col_idx, col_idx, 20, date_format)
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    file_ext = "xlsx"
    output.seek(0)
    return output, mime, file_ext
