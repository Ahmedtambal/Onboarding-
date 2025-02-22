import streamlit as st
import pandas as pd
from logic import parse_docx, parse_pdf, load_master_file, append_employee_record, export_master_file

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
st.markdown('<div class="header"><h1>New Employee Data Uploader</h1><p>Streamlined Onboarding for Modern HR</p></div>', unsafe_allow_html=True)

st.markdown(
    """
    <div class="upload-section">
      <p>
         This tool lets you upload a new employee details file (DOCX or PDF) and an Excel master record.
         Any missing fields will be set as NaN.
      </p>
      <h4>Required Columns</h4>
      <ul>
         <li>Title</li>
         <li>Full Name</li>
         <li>Home Address</li>
         <li>Home Telephone Number</li>
         <li>Mobile Telephone Number</li>
         <li>Telephone Number</li>
         <li>Personal Email Address</li>
         <li>Date of Birth</li>
         <li>Pronouns</li>
         <li>National Insurance Number (or National Insurance No.)</li>
         <li>Job Title</li>
         <li>Start Date (or Date Employment Commenced)</li>
         <li>Basic Salary</li>
         <li>Pension Contribution</li>
         <li>Marital Status</li>
         <li>Nationality</li>
         <li>Country of Residence</li>
         <li>Name of an Emergency Contact (or Emergency Contact Name)</li>
         <li>Emergency Contact Number (or Telephone Number of Emergency Contact)</li>
         <li>Emergency Contact Address</li>
         <li>Emergency Contact Email</li>
         <li>Relationship to Emergency Contact</li>
         <li>Employment Location Postcode</li>
         <li>Notes</li>
      </ul>
      <p>
         If you want to add more columns, please contact Ahmed.
      </p>
    </div>
    """, unsafe_allow_html=True
)
# Two-column layout for file uploads.
col1, col2 = st.columns(2)
with col1:
    emp_file = st.file_uploader("Upload Employee Details File (DOCX or PDF)", type=["docx", "pdf"])
with col2:
    master_file = st.file_uploader("Upload the Excel Master File", type=["xlsx", "xls"])

if emp_file is not None and master_file is not None:
    file_bytes = emp_file.read()
    if emp_file.name.lower().endswith(".docx"):
        emp_data = parse_docx(file_bytes, debug=DEBUG)
    elif emp_file.name.lower().endswith(".pdf"):
        emp_data = parse_pdf(file_bytes, debug=DEBUG)
    else:
        st.error("Unsupported employee file format.")
        emp_data = {}

    st.subheader("Extracted Employee Data")
    st.write(emp_data)
    
    try:
        df = load_master_file(master_file, master_file.name)
    except Exception as e:
        st.error(f"Error reading master file: {e}")
        df = pd.DataFrame()
    
    st.subheader("Current Master Record")
    st.dataframe(df)
    
    updated_df = append_employee_record(df, emp_data, debug=DEBUG)
    
    st.subheader("Updated Master Record")
    st.dataframe(updated_df)
    
    output, mime, file_ext = export_master_file(updated_df, master_file.name)
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center;'>Download Updated Master File</h3>", unsafe_allow_html=True)
    st.download_button(
        label="Download Updated Master File",
        data=output,
        file_name=f"Updated_Master_File.{file_ext}",
        mime=mime
    )
