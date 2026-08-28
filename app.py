import streamlit as st
import pandas as pd
import re
import io

# App Layout Configuration
st.set_page_config(page_title="Recruitment Funnel Builder", page_icon="📊", layout="wide")

st.title("📊 Recruitment Funnel & Data Cleaner")
st.markdown("Upload your raw extraction spreadsheet to dynamically compute funnel metrics and export tailored data sets.")

# File Uploader
uploaded_file = st.file_uploader("Upload your raw CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        # Load File Safely
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        # --- THE UNDERLYING CLEANING CORE ---
        # Strip invisible spaces and fix cases across the core dataset
        strings_to_clean = ['Candidate Name', 'Email Address', 'NRIC Number', 'Application Status', 'Job Name', 'Citizenship', 'Country Of Birth']
        for col in strings_to_clean:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                
        if 'Candidate Name' in df.columns:
            df['Candidate Name'] = df['Candidate Name'].str.title()
        if 'Email Address' in df.columns:
            df['Email Address'] = df['Email Address'].str.lower()
        if 'NRIC Number' in df.columns:
            df['NRIC Number'] = df['NRIC Number'].str.upper()
        if 'Citizenship' in df.columns:
            df['Citizenship'] = df['Citizenship'].str.title()
        if 'Country Of Birth' in df.columns:
            df['Country Of Birth'] = df['Country Of Birth'].str.title()

        # Handle numeric errors safely
        if 'X0PA Score' in df.columns:
            df['X0PA Score'] = pd.to_numeric(df['X0PA Score'], errors='coerce').fillna(0)
        if 'Total Exp' in df.columns:
            df['Total Exp'] = pd.to_numeric(df['Total Exp'], errors='coerce').fillna(0)

        # Parse Employer From Complex Work Experience Strings
        def get_current_company(txt):
            if pd.isna(txt) or not isinstance(txt, str):
                return "Not Listed"
            match = re.search(r'C:\s*([^.\n]+)', txt)
            return match.group(1).strip() if match else "Not Listed"

        df['Current_Company'] = df['Work Experience'].apply(get_current_company) if 'Work Experience' in df.columns else "Not Listed"

        # Determine Baseline Funnel Stages From Status Strings
        def evaluate_stage(status):
            status_str = str(status)
            if "Slot In Progress" in status_str:
                return "Shortlisted"
            elif "Reject" in status_str:
                return "Rejected"
            else:
                return "Screening Pool"

        df['Funnel_Stage'] = df['Application Status'].apply(evaluate_stage) if 'Application Status' in df.columns else "Screening Pool"

        # --- ADVANCED CUSTOM EDUCATION PARSER ---
        def parse_education_funnel(edu_text):
            if pd.isna(edu_text) or not isinstance(edu_text, str) or edu_text.strip() == "":
                return "Not Provided"
            
            # Extract standard key degree indicators from raw pipe-delimited text blocks
            text_upper = edu_text.upper()
            
            has_phd = "PHD" in text_upper or "DOCTOR" in text_upper
            has_master = "MASTER" in text_upper or "MSC" in text_upper or "MBA" in text_upper
            has_bachelor = "BACHELOR" in text_upper or "DEGREE" in text_upper or "BSC" in text_upper or "BENG" in text_upper
            has_diploma = "DIPLOMA" in text_upper or "POLYTECHNIC" in text_upper
            has_alevels = "A LEVEL" in text_upper or "ADVANCED LEVEL" in text_upper or "JUNIOR COLLEGE" in text_upper
            
            # Apply your tiered visibility truncation rules
            if has_phd:
                output = ["PhD"]
                if has_master: output.append("Master")
                if has_bachelor: output.append("Bachelor")
                return " ➔ ".join(output)
            
            elif has_master:
                output = ["Master"]
                if has_bachelor: output.append("Bachelor")
                return " ➔ ".join(output)
            
            elif has_bachelor:
                return "Bachelor"
            
            elif has_diploma or has_alevels:
                output = []
                if has_diploma: output.append("Diploma")
                if has_alevels: output.append("A-Levels")
                return " & ".join(output)
                
            return "Other / Lower Secondary"

        df['Highest_Education'] = df['Candidate Education'].apply(parse_education_funnel) if 'Candidate Education' in df.columns else "Not Provided"

        # --- EXPLICIT COLUMN FILTERING ---
        desired_columns = [
            'Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 
            'Citizenship', 'Country Of Birth', 'Highest_Education', 'Current_Company',
            'Total Exp', 'X0PA Score', 'Funnel_Stage', 'Application Status', 'App Date',
            'Job Id', 'Job Name', 'Job Status', 'Application Source'
        ]
        
        # Only keep columns that actually exist in the uploaded file to avoid errors
        available_cols = [c for c in desired_columns if c in df.columns or c in ['Current_Company', 'Funnel_Stage', 'Highest_Education']]
        master_cleaned_df = df[available_cols]

        # --- DYNAMIC PROCESSOR FOR THE TWO VIEWS ---
        # View A: Total Applications (Keep everything)
        apps_df = master_cleaned_df.copy()

        # View B: Smart Unique Applicants Filter (Email or NRIC Check)
        applicants_df = master_cleaned_df.copy().sort_values(by='X0PA Score', ascending=False)
        
        for col in ['Email Address', 'NRIC Number']:
            if col in applicants_df.columns:
                applicants_df[col] = applicants_df[col].replace(['nan', 'none', 'na', ''], pd.NA)

        if 'Email Address' in applicants_df.columns or 'NRIC Number' in applicants_df.columns:
            has_nric = applicants_df[applicants_df['NRIC Number'].notna()]
            no_nric = applicants_df[applicants_df['NRIC Number'].isna()]
            
            if 'NRIC Number' in applicants_df.columns:
                has_nric = has_nric.drop_duplicates(subset=['NRIC Number'], keep='first')
            
            combined_pass1 = pd.concat([has_nric, no_nric])
            
            if 'Email Address' in applicants_df.columns:
                applicants_df = combined_pass1.drop_duplicates(subset=['Email Address'], keep='first')
            else:
                applicants_df = combined_pass1

        # Fill NAs back with clean display values for the UI tables
        apps_df = apps_df.fillna("Not Provided")
        applicants_df = applicants_df.fillna("Not Provided")

        # --- USER INTERFACE TABS ---
        tab1, tab2 = st.tabs(["📈 View A: Total Applications Funnel", "👥 View B: Unique Applicants Funnel"])

        # TAB 1: TOTAL APPLICATIONS VIEW
        with tab1:
            st.subheader("Application Funnel Summary")
            total_apps = len(apps_df)
            shortlisted_apps = len(apps_df[apps_df['Funnel_Stage'] == "Shortlisted"])
            screening_apps = len(apps_df[apps_df['Funnel_Stage'] == "Screening Pool"])
            rejected_apps = len(apps_df[apps_df['Funnel_Stage'] == "Rejected"])

            # Metric blocks display
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Applications", total_apps)
            c2.metric("Shortlisted / In-Progress", shortlisted_apps)
            c3.metric("Awaiting Screening", screening_apps)
            c4.metric("Rejected Pool", rejected_apps)

            st.markdown("### Cleaned Application Data Table")
            st.dataframe(apps_df)

            # Excel download buffer
            buf_a = io.BytesIO()
            with pd.ExcelWriter(buf_a, engine='openpyxl') as w:
                apps_df.to_excel(w, index=False, sheet_name='All Applications')
            
            st.download_button(
                label="📥 Export Total Applications Workbook",
                data=buf_a.getvalue(),
                file_name="total_applications_funnel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # TAB 2: UNIQUE APPLICANTS VIEW
        with tab2:
            st.subheader("Unique Applicant Funnel Summary")
            total_uniq = len(applicants_df)
            shortlisted_uniq = len(applicants_df[applicants_df['Funnel_Stage'] == "Shortlisted"])
            screening_uniq = len(applicants_df[applicants_df['Funnel_Stage'] == "Screening Pool"])
            rejected_uniq = len(applicants_df[applicants_df['Funnel_Stage'] == "Rejected"])

            # Metric blocks display
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Unique Talent", total_uniq)
            m2.metric("Shortlisted Candidates", shortlisted_uniq)
            m3.metric("Awaiting Screening", screening_uniq)
            m4.metric("Rejected Candidates", rejected_uniq)

            st.markdown("### Cleaned Unique Candidate Data Table")
            st.dataframe(applicants_df)

            # Excel download buffer
            buf_b = io.BytesIO()
            with pd.ExcelWriter(buf_b, engine='openpyxl') as w:
                applicants_df.to_excel(w, index=False, sheet_name='Unique Applicants')
            
            st.download_button(
                label="📥 Export Unique Applicants Workbook",
                data=buf_b.getvalue(),
                file_name="unique_applicants_funnel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while compiling your funnel: {e}")
else:
    st.info("Awaiting raw dataset upload to generate funnel pipelines.")
