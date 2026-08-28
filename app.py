import streamlit as st
import pandas as pd
import re
import io

# Set up clean Web App configuration
st.set_page_config(page_title="Recruitment Data Cleaner", page_icon="💼", layout="wide")

st.title("💼 Automated Recruitment Data Cleaner")
st.markdown("Upload your raw extraction spreadsheet to automatically standardize text, parse current employers, and tag candidate metrics.")

# 1. File Uploader Component
uploaded_file = st.file_uploader("Upload your raw CSV or Excel file here", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        # Check file extension and load appropriately
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success(f"Successfully loaded {len(df)} candidate entries!")
        
        # --- THE CLEANING ENGINE ---
        # Standardize basic text case structures
        if 'Candidate Name' in df.columns:
            df['Candidate Name'] = df['Candidate Name'].astype(str).str.title().str.strip()
        if 'Email Address' in df.columns:
            df['Email Address'] = df['Email Address'].astype(str).str.lower().str.strip()
        if 'Application Status' in df.columns:
            df['Application Status'] = df['Application Status'].astype(str).str.strip()
            
        # Standardize numeric columns safely
        if 'X0PA Score' in df.columns:
            df['X0PA Score'] = pd.to_numeric(df['X0PA Score'], errors='coerce').fillna(0)
        if 'Total Exp' in df.columns:
            df['Total Exp'] = pd.to_numeric(df['Total Exp'], errors='coerce').fillna(0)
            
        # Extract Current Company function
        def extract_company(text):
            if pd.isna(text) or not isinstance(text, str):
                return "Not Listed"
            match = re.search(r'C:\s*([^.\n]+)', text)
            return match.group(1).strip() if match else "Not Listed"

        if 'Work Experience' in df.columns:
            df['Current_Company'] = df['Work Experience'].apply(extract_company)
        else:
            df['Current_Company'] = "Not Listed"

        # Apply Recruitment Experience Tags
        def tag_experience(years):
            if years <= 2: return "Junior"
            elif years <= 7: return "Mid-Level"
            elif years <= 12: return "Senior"
            else: return "Lead / Principal"

        if 'Total Exp' in df.columns:
            df['Experience_Level'] = df['Total Exp'].apply(tag_experience)
            
        # Apply Candidate Priorities Tags
        def tag_priority(row):
            score = row.get('X0PA Score', 0)
            status = str(row.get('Application Status', ''))
            if score >= 65: return "High Priority"
            elif "Slot In Progress" in status: return "Medium Priority"
            else: return "Standard Review"

        df['Application_Priority'] = df.apply(tag_priority, axis=1)
        
        # Organize columns layout to display the most critical information first
        core_cols = [
            'Candidate Name', 'Email Address', 'Current_Company', 'Experience_Level', 
            'Total Exp', 'X0PA Score', 'Application_Priority', 'Application Status', 
            'Job Name'
        ]
        # Keep any other original metrics appended at the end
        remaining_cols = [c for c in df.columns if c not in core_cols]
        final_df = df[core_cols + remaining_cols]
        
        # --- USER INTERFACE ELEMENTS ---
        st.subheader("👀 Cleaned Data Preview")
        st.markdown("Here is how your data looks right now. Scroll to check the new tagged columns.")
        st.dataframe(final_df)
        
        # --- EXPORT TO EXCEL SYSTEM ---
        # Generate an in-memory stream for Excel download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Cleaned Candidates')
        
        st.markdown("---")
        st.subheader("💾 Export Cleaned Workbook")
        
        st.download_button(
            label="📥 Download Processed Excel Sheet",
            data=buffer.getvalue(),
            file_name="cleaned_recruitment_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred while processing the app: {e}")
else:
    st.info("Awaiting raw candidate data file. Please upload an Excel sheet or CSV above.")
