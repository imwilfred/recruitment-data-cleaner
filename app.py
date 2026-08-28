import streamlit as st, pandas as pd, re, io
st.set_page_config(page_title="Funnel Builder", page_icon="📊", layout="wide")
st.title("📊 Recruitment Funnel & Data Cleaner")
uploaded_file = st.file_uploader("Upload your raw CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        clean_cols = ['Candidate Name', 'Email Address', 'NRIC Number', 'Application Status', 'Job Name', 'Citizenship', 'Country Of Birth']
        for col in clean_cols:
            if col in df.columns: df[col] = df[col].astype(str).str.strip()
        if 'Candidate Name' in df.columns: df['Candidate Name'] = df['Candidate Name'].str.title()
        if 'Email Address' in df.columns: df['Email Address'] = df['Email Address'].str.lower()
        if 'NRIC Number' in df.columns: df['NRIC Number'] = df['NRIC Number'].str.upper()
        if 'Citizenship' in df.columns: df['Citizenship'] = df['Citizenship'].str.title()
        if 'Country Of Birth' in df.columns: df['Country Of Birth'] = df['Country Of Birth'].str.title()
        if 'X0PA Score' in df.columns: df['X0PA Score'] = pd.to_numeric(df['X0PA Score'], errors='coerce').fillna(0)
        if 'Total Exp' in df.columns: df['Total Exp'] = pd.to_numeric(df['Total Exp'], errors='coerce').fillna(0)

        # Parse current employer strings
        def get_company(txt):
            if pd.isna(txt) or not isinstance(txt, str): return "Not Listed"
            m = re.search(r'C:\s*([^.\n]+)', txt)
            return m.group(1).strip() if m else "Not Listed"
        df['Current_Company'] = df['Work Experience'].apply(get_company) if 'Work Experience' in df.columns else "Not Listed"

        # Determine funnel pipeline stages
        def get_stage(status):
            s = str(status)
            return "Shortlisted" if "Slot In Progress" in s else ("Rejected" if "Reject" in s else "Screening Pool")
        df['Funnel_Stage'] = df['Application Status'].apply(get_stage) if 'Application Status' in df.columns else "Screening Pool"

        # Evaluate recruitment eligibility parameters
        def check_eligibility(row):
            cz = str(row.get('Citizenship', '')).strip().lower()
            cob = str(row.get('Country Of Birth', '')).strip().lower()
            if 'citizen' in cz or cz == 'singapore':
                if any(x in cob for x in ['china', 'myanmar', 'myanmr']): return "Ineligible (Exception)"
                return "Eligible"
            return "Ineligible"
        df['Eligibility_Status'] = df.apply(check_eligibility, axis=1)

        # Content-Aware Education, Field, and Institution Parser
        def get_edu_details(edu_text):
            defaults = {"l": "Not Provided", "d": "Not Listed", "s": "Not Listed"}
            if pd.isna(edu_text) or not isinstance(edu_text, str) or edu_text.strip() == "": return defaults
            parts = [p.strip() for p in edu_text.split('|') if p.strip()]
            if not parts: return defaults
            u = edu_text.upper()
            phd, master, bach, dip, alev = "PHD" in u or "DOCTOR" in u, "MASTER" in u or "MSC" in u or "MBA" in u, "BACHELOR" in u or "DEGREE" in u or "BSC" in u or "BENG" in u, "DIPLOMA" in u or "POLYTECHNIC" in u, "A LEVEL" in u or "ADVANCED LEVEL" in u or "JUNIOR COLLEGE" in u
            lvl = "PhD" + (" ➔ Master" if master else "") + (" ➔ Bachelor" if bach else "") if phd else ("Master" + (" ➔ Bachelor" if bach else "") if master else ("Bachelor" if bach else (" & ".join([w for w, c in [("Diploma", dip), ("A-Levels", alev)] if c]) if (dip or alev) else "Other / School")))
            t_kw = "PHD" if phd else ("MASTER" if master else ("BACHELOR" if bach else "DIPLOMA"))
            t_entries = []
            for i, p in enumerate(parts):
                if t_kw in p.upper():
                    t_entries = parts[max(0, i - 1):min(len(parts), i + 5)]
                    break
            if not t_entries: t_entries = parts
            sch, disc = "Not Listed", "Not Listed"
            school_keywords = ["UNIVERSITY", "POLYTECHNIC", "INSTITUTE", "COLLEGE", "SCHOOL", "NUS", "NTU", "SMU", "SIT", "SUTD", "SUSS", "ACADEMY"]
            ignore_keywords = ["BACHELOR", "MASTER", "PHD", "DIPLOMA", "DEGREE", "HONOURS", "HONORS", "DISTINCTION", "CERTIFICATE", "BSC", "BENG", "MSC", "MBA", "CERTIFICATION"]
            for item in t_entries:
                item_upper = item.upper()
                if re.search(r'\d{4}', item) or re.match(r'^\d+(\.\d+)?$', item): continue
                if any(k in item_upper for k in school_keywords): sch = item; continue
                if not any(k in item_upper for k in ignore_keywords) and len(item) > 2 and disc == "Not Listed": disc = item
            if disc == "Not Listed" and len(t_entries) > 0:
                for item in t_entries:
                    if not any(k in item.upper() for k in school_keywords + ignore_keywords) and not re.search(r'\d', item): disc = item; break
            disc = re.sub(r'(Bachelor of|Master of|BSc|BEng|Diploma in|BSc Hons|Degree in|With Honours|With Distinction)\s*', '', disc, flags=re.IGNORECASE).strip()
            return {"l": lvl, "d": disc.title() if disc != "Not Listed" else "Not Listed", "s": sch.title() if sch != "Not Listed" else "Not Listed"}

        if 'Candidate Education' in df.columns:
            edu_mapped = df['Candidate Education'].apply(get_edu_details)
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = [i['l'] for i in edu_mapped], [i['d'] for i in edu_mapped], [i['s'] for i in edu_mapped]
        else:
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = "Not Provided", "Not Listed", "Not Listed"

        # Organize screen view layout columns
        layout = ['Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 'Citizenship', 'Country Of Birth', 'Eligibility_Status', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Current_Company', 'Total Exp', 'X0PA Score', 'Funnel_Stage', 'Application Status', 'App Date', 'Job Id', 'Job Name', 'Job Status', 'Application Source']
        available_cols = [c for c in layout if c in df.columns or c in ['Current_Company', 'Funnel_Stage', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Eligibility_Status']]
        master_df = df[available_cols]
        apps_df = master_df.copy().fillna("Not Provided")

        uniq_df = master_df.copy().sort_values(by='X0PA Score', ascending=False)
        for col in ['Email Address', 'NRIC Number']:
            if col in uniq_df.columns: uniq_df[col] = uniq_df[col].replace(['nan', 'none', 'na', ''], pd.NA)
        if 'Email Address' in uniq_df.columns or 'NRIC Number' in uniq_df.columns:
            has_n = uniq_df[uniq_df['NRIC Number'].notna()]
            no_n = uniq_df[uniq_df['NRIC Number'].isna()]
            if 'NRIC Number' in uniq_df.columns: has_n = has_n.drop_duplicates(subset=['NRIC Number'], keep='first')
            comb = pd.concat([has_n, no_n])
            uniq_df = comb.drop_duplicates(subset=['Email Address'], keep='first') if 'Email Address' in comb.columns else comb
        uniq_df = uniq_df.fillna("Not Provided")

        # --- RE-CONFIGURED METRIC PANEL DESIGN ---
        tab1, tab2 = st.tabs(["📈 View A: Total Applications Funnel", "👥 View B: Unique Applicants Funnel"])
        
        with tab1:
            st.subheader("Application Funnel Summary")
            t_a = len(apps_df)
            e_a = len(apps_df[apps_df['Eligibility_Status'] == "Eligible"])
            s_a = len(apps_df[apps_df['Funnel_Stage'] == "Shortlisted"])
            r_a = len(apps_df[apps_df['Funnel_Stage'] == "Rejected"])
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Applications", t_a)
            col2.metric("Eligible Applications", e_a)
            col3.metric("Shortlisted / In-Progress", s_a)
            col4.metric("Rejected Pool", r_a)
            
            st.dataframe(apps_df)
            buf_a = io.BytesIO()
            with pd.ExcelWriter(buf_a, engine='openpyxl') as w: apps_df.to_excel(w, index=False, sheet_name='All Applications')
            st.download_button("📥 Export Total Applications", data=buf_a.getvalue(), file_name="total_applications.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        with tab2:
            st.subheader("Unique Applicant Funnel Summary")
            t_u = len(uniq_df)
            e_u = len(uniq_df[uniq_df['Eligibility_Status'] == "Eligible"])
            s_u = len(uniq_df[uniq_df['Funnel_Stage'] == "Shortlisted"])
            r_u = len(uniq_df[uniq_df['Funnel_Stage'] == "Rejected"])
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Unique Talent", t_u)
            m2.metric("Eligible Unique Talent", e_u)
            m3.metric("Shortlisted Candidates", s_u)
            m4.metric("Rejected Candidates", r_u)
            
            st.dataframe(uniq_df)
            buf_b = io.BytesIO()
            with pd.ExcelWriter(buf_b, engine='openpyxl') as w: uniq_df.to_excel(w, index=False, sheet_name='Unique Applicants')
            st.download_button("📥 Export Unique Applicants", data=buf_b.getvalue(), file_name="unique_applicants.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
    except Exception as e: st.error(f"An error occurred while compiling your funnel: {e}")
else: st.info("Awaiting raw dataset upload to generate funnel pipelines.")
