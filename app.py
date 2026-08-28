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
        for col in ['Candidate Name', 'Email Address', 'Application Status', 'Job Name']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                
        if 'Candidate Name' in df.columns:
            df['Candidate Name'] = df['Candidate Name'].str.title()
        if 'Email Address' in df.columns:
            df['Email Address'] = df['Email Address'].str.lower()

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

        # --- EXPLICIT COLUMN FILTERING ---
        # Define the exact columns needed for the analysis
        desired_columns = [
            'Candidate Name', 'Email Address', 'Mobile Number', 'Current_Company',
            'Total Exp', 'X0PA Score', 'Funnel_Stage', 'Application Status', 'App Date',
            'Job Id', 'Job Name', 'Job Status', 'Application Source', 'Recruiter Name'
        ]
        
        # Only keep columns that actually exist in the uploaded file to avoid errors
        available_cols = [c for c in desired_columns if c in df.columns or c in ['Current_Company', 'Funnel_Stage']]
        master_cleaned_df = df[available_cols]

        # --- DYNAMIC PROCESSOR FOR THE TWO VIEWS ---
        # View A: Total Applications (Keep everything)
        apps_df = master_cleaned_df.copy()

        # View B: Unique Applicants (Remove duplicates by Email, keeping highest X0PA Score row)
        applicants_df = master_cleaned_df.copy()
        if 'Email Address' in applicants_df.columns and 'X0PA Score' in applicants_df.columns:
            applicants_df = applicants_df.sort_values(by='X0PA Score', ascending=False)
            applicants_df = applicants_df.drop_duplicates(subset=['Email Address'], keep='first')

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
