import streamlit as st
import pandas as pd
from logic import (
    parse_docx, parse_pdf, parse_csv_employee, parse_excel_employee, 
    load_master_file, append_employee_record, export_master_file
)

# Enable debugging output
DEBUG = True

# Inject custom CSS for a modern, stylish UI.
st.markdown(
    """
    <style>
    body {
         background: #f4f4f9;
         color: #333;
         font-family: 'Roboto', sans-serif;
    }
    .stButton>button {
         background-color: #4CAF50;
         color: white;
         border: none;
         padding: 0.75rem 1.5rem;
         border-radius: 8px;
         font-size: 1rem;
         font-weight: bold;
         transition: background-color 0.3s ease;
    }
    .stButton>button:hover {
         background-color: #45a049;
    }
    .stFileUploader > label {
         font-size: 1.1rem;
         font-weight: bold;
    }
    .header {
         text-align: center;
         margin-bottom: 2rem;
    }
    .header h1 {
         font-size: 3rem;
         color: #2c3e50;
         margin-bottom: 0;
    }
    .header p {
         font-size: 1.2rem;
         color: #34495e;
         margin-top: 0.5rem;
    }
    .upload-section {
         background-color: #fff;
         border-radius: 8px;
         padding: 2rem;
         box-shadow: 0 2px 10px rgba(0,0,0,0.1);
         margin-bottom: 2rem;
    }
    </style>
    """, unsafe_allow_html=True
)

# Header section
st.markdown('<div class="header"><h1>New Employee Data Uploader</h1><p>Streamlined Onboarding</p></div>', unsafe_allow_html=True)

st.markdown(
    """
    <div class="upload-section">
      <p>
         This tool lets you upload one or more new employee details files (DOCX, PDF, CSV, TXT, or Excel) along with a master record file.
         Any missing fields will be set as NA.
      </p>
    </div>
    """, unsafe_allow_html=True
)

# Two-column layout for file uploads.
col1, col2 = st.columns(2)
with col1:
    # Allow multiple employee files (DOCX, PDF, CSV, TXT, Excel)
    emp_files = st.file_uploader("Upload Employee Details File(s)", 
                                 type=["docx", "pdf", "csv", "txt", "xlsx", "xls"], 
                                 accept_multiple_files=True)
with col2:
    # Allow master file in Excel, CSV, or TXT format
    master_file = st.file_uploader("Upload the Master File", type=["xlsx", "xls", "csv", "txt"])

if emp_files is not None and len(emp_files) > 0 and master_file is not None:
    try:
        df = load_master_file(master_file, master_file.name)
    except Exception as e:
        st.error(f"Error reading master file: {e}")
        df = pd.DataFrame()
    
    for emp_file in emp_files:
        file_bytes = emp_file.read()
        emp_data_list = []
        if emp_file.name.lower().endswith(".docx"):
            emp_data = parse_docx(file_bytes, debug=DEBUG)
            emp_data_list.append(emp_data)
        elif emp_file.name.lower().endswith(".pdf"):
            emp_data = parse_pdf(file_bytes, debug=DEBUG)
            emp_data_list.append(emp_data)
        elif emp_file.name.lower().endswith((".csv", ".txt")):
            emp_data = parse_csv_employee(file_bytes, debug=DEBUG)
            emp_data_list.append(emp_data)
        elif emp_file.name.lower().endswith((".xlsx", ".xls")):
            # For Excel employee files, treat each row as a new employee.
            emp_data_list = parse_excel_employee(file_bytes, debug=DEBUG)
        else:
            st.error(f"Unsupported employee file format: {emp_file.name}")
            continue

        for idx, emp_data in enumerate(emp_data_list):
            st.subheader(f"Extracted Data from {emp_file.name} - Employee {idx+1}")
            # st.write(emp_data)  # Commented out to hide JSON output
            df = append_employee_record(df, emp_data, debug=DEBUG)
    
    st.subheader("Current Master Record")
    st.dataframe(df)
    
    output, mime, file_ext = export_master_file(df, master_file.name)
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center;'>Download Updated Master File</h3>", unsafe_allow_html=True)
    st.download_button(
        label="Download Updated Master File",
        data=output,
        file_name=f"Updated_Master_File.{file_ext}",
        mime=mime
    )
