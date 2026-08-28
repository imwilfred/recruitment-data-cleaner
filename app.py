import streamlit as st, pandas as pd, re, io
st.set_page_config(page_title="Funnel Builder", page_icon="📊", layout="wide")
st.title("Advanced Recruitment Funnel & Data Cleaner")

uploaded_file = st.file_uploader("Upload your raw CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        
        # Clean text columns
        cc = ['Candidate Name', 'Email Address', 'NRIC Number', 'Application Status', 'Job Name', 'Citizenship', 'Country Of Birth']
        for c in cc:
            if c in df.columns: df[c] = df[c].astype(str).str.strip()
        if 'Candidate Name' in df.columns: df['Candidate Name'] = df['Candidate Name'].str.title()
        if 'Email Address' in df.columns: df['Email Address'] = df['Email Address'].str.lower()
        if 'NRIC Number' in df.columns: df['NRIC Number'] = df['NRIC Number'].str.upper()
        if 'Citizenship' in df.columns: df['Citizenship'] = df['Citizenship'].str.title()
        if 'Country Of Birth' in df.columns: df['Country Of Birth'] = df['Country Of Birth'].str.title()
        if 'X0PA Score' in df.columns: df['X0PA Score'] = pd.to_numeric(df['X0PA Score'], errors='coerce').fillna(0)
        if 'Total Exp' in df.columns: df['Total Exp'] = pd.to_numeric(df['Total Exp'], errors='coerce').fillna(0)

        # Parse employer text blocks
        def get_company(txt):
            if pd.isna(txt) or not isinstance(txt, str): return "Not Listed"
            m = re.search(r'C:\s*([^.\n]+)', txt)
            return m.group(1).strip() if m else "Not Listed"
        df['Current_Company'] = df['Work Experience'].apply(get_company) if 'Work Experience' in df.columns else "Not Listed"

        # Check SG citizenship eligibility criteria
        def get_eligibility(row):
            cz = str(row.get('Citizenship', '')).lower()
            cob = str(row.get('Country Of Birth', '')).lower()
            if 'citizen' in cz or cz == 'singapore':
                if any(x in cob for x in ['china', 'myanmar', 'myanmr']): return "Ineligible (Exception)"
                return "Eligible"
            return "Ineligible"
        df['Eligibility_Status'] = df.apply(get_eligibility, axis=1)

        # Precise 11-Tier Status Rankings mapping
        st_map = {
            "Hired": 1, "Hire in Progress": 2, "Offer in Progress": 3,
            "Verbal Offer in Progress": 4, "Salary Proposal in Progress": 5,
            "Interview in Progress": 6, "Interview Reject": 7,
            "Post Screening Slot in Progress": 8, "Post Screening Slot Reject": 9,
            "Screening in Progress": 10, "Screening Reject": 11
        }
        df['Rank'] = df['Application Status'].apply(lambda x: st_map.get(str(x).strip(), 12))

        # Classify high level funnel milestones using safe numeric comparisons
        def get_stage(r):
            if r == 1: return "Hired"
            elif r >= 2 and r <= 5: return "Offered Stage"
            elif r >= 6 and r <= 7: return "Interview Stage"
            elif r >= 8 and r <= 9: return "Shortlisted Stage"
            elif r == 11: return "Rejected Baseline"
            return "Screening Pool"
        df['Funnel_Stage'] = df['Rank'].apply(get_stage)

        # Extract Tiered Education details
        def get_edu(edu_text):
            defaults = {"l": "Not Provided", "d": "Not Listed", "s": "Not Listed"}
            if pd.isna(edu_text) or not isinstance(edu_text, str) or edu_text.strip() == "": return defaults
            parts = [p.strip() for p in edu_text.split('|') if p.strip()]
            if not parts: return defaults
            u = edu_text.upper()
            phd, master, bach, dip, alev = "PHD" in u or "DOCTOR" in u, "MASTER" in u or "MSC" in u or "MBA" in u, "BACHELOR" in u or "DEGREE" in u or "BSC" in u or "BENG" in u, "DIPLOMA" in u or "POLYTECHNIC" in u, "A LEVEL" in u or "ADVANCED LEVEL" in u or "JUNIOR COLLEGE" in u
            lvl = "PhD" + (" ➔ Master" if master else "") + (" ➔ Bachelor" if bach else "") if phd else ("Master" + (" ➔ Bachelor" if bach else "") if master else ("Bachelor" if bach else (" & ".join([w for w, c in [("Diploma", dip), ("A-Levels", alev)] if c]) if (dip or alev) else "Other / School")))
            t_idx = 0
            t_kw = "PHD" if phd else ("MASTER" if master else ("BACHELOR" if bach else "DIPLOMA"))
            for i, p in enumerate(parts):
                if t_kw in p.upper(): t_idx = i; break
            try:
                t_arr = parts[max(0, t_idx-1):min(len(parts), t_idx+5)]
                sch, disc = "Not Listed", "Not Listed"
                skw = ["UNIVERSITY", "POLYTECHNIC", "INSTITUTE", "COLLEGE", "SCHOOL", "NUS", "NTU", "SMU", "SIT", "SUTD", "SUSS", "ACADEMY"]
                ikw = ["BACHELOR", "MASTER", "PHD", "DIPLOMA", "DEGREE", "HONOURS", "HONORS", "DISTINCTION", "CERTIFICATE", "BSC", "BENG", "MSC", "MBA", "CERTIFICATION"]
                for item in t_arr:
                    if re.search(r'\d{4}', item) or re.match(r'^\d+(\.\d+)?$', item): continue
                    if any(k in item.upper() for k in skw): sch = item; continue
                    if not any(k in item.upper() for k in ikw) and len(item) > 2 and disc == "Not Listed": disc = item
                disc = re.sub(r'(Bachelor of|Master of|BSc|BEng|Diploma in|BSc Hons|Degree in)\s*', '', disc, flags=re.IGNORECASE).strip()
                return {"l": lvl, "d": disc.title(), "s": sch.title()}
            except:
                return {"l": lvl, "d": "Not Listed", "s": "Not Listed"}

        if 'Candidate Education' in df.columns:
            edu_mapped = df['Candidate Education'].apply(get_edu)
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = [i['l'] for i in edu_mapped], [i['d'] for i in edu_mapped], [i['s'] for i in edu_mapped]
        else:
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = "Not Provided", "Not Listed", "Not Listed"

        # Structural Layout Filter
        layout = ['Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 'Citizenship', 'Country Of Birth', 'Eligibility_Status', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Current_Company', 'Total Exp', 'X0PA Score', 'Funnel_Stage', 'Application Status', 'App Date', 'Job Id', 'Job Name', 'Job Status', 'Application Source', 'Rank']
        av_cols = [c for c in layout if c in df.columns or c in ['Current_Company', 'Funnel_Stage', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Eligibility_Status', 'Rank']]
        master_df = df[av_cols]

        # VIEW A: Total Applications
        apps_df = master_df.copy().fillna("Not Provided")
        if 'Rank' in apps_df.columns: apps_df = apps_df.drop(columns=['Rank'])

        # VIEW B: Unique Applicants Hierarchical Deduplication
        uniq_df = master_df.copy().sort_values(by=['Rank', 'X0PA Score'], ascending=[True, False])
        for col in ['Email Address', 'NRIC Number']:
            if col in uniq_df.columns: uniq_df[col] = uniq_df[col].replace(['nan', 'none', 'na', ''], pd.NA)
        if 'Email Address' in uniq_df.columns or 'NRIC Number' in uniq_df.columns:
            has_n = uniq_df[uniq_df['NRIC Number'].notna()]
            no_n = uniq_df[uniq_df['NRIC Number'].isna()]
            if 'NRIC Number' in uniq_df.columns: has_n = has_n.drop_duplicates(subset=['NRIC Number'], keep='first')
            comb = pd.concat([has_n, no_n])
            uniq_df = comb.drop_duplicates(subset=['Email Address'], keep='first') if 'Email Address' in comb.columns else comb
        if 'Rank' in uniq_df.columns: uniq_df = uniq_df.drop(columns=['Rank'])
        uniq_df = uniq_df.fillna("Not Provided")

        # UI Panels rendering
        tab1, tab2 = st.tabs(["📈 View A: Total Applications Funnel", "👥 View B: Unique Applicants Funnel"])
        
        with tab1:
            st.subheader("Application Funnel Summary")
            ta, ea, sa = len(apps_df), len(apps_df[apps_df['Eligibility_Status'] == "Eligible"]), len(apps_df[apps_df['Funnel_Stage'].isin(["Shortlisted Stage", "Interview Stage", "Offered Stage", "Hired"])])
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Applications", ta)
            col2.metric("Eligible Applications", ea)
            col3.metric("Shortlisted Applications", sa)
            st.dataframe(apps_df)
            buf_a = io.BytesIO()
            with pd.ExcelWriter(buf_a, engine='openpyxl') as w: apps_df.to_excel(w, index=False, sheet_name='All Applications')
            st.download_button("📥 Export Total Applications", data=buf_a.getvalue(), file_name="total_applications.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        with tab2:
            st.subheader("Unique Applicant Funnel Summary")
            tu, eu, su = len(uniq_df), len(uniq_df[uniq_df['Eligibility_Status'] == "Eligible"]), len(uniq_df[uniq_df['Funnel_Stage'].isin(["Shortlisted Stage", "Interview Stage", "Offered Stage", "Hired"])])
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Unique Talent", tu)
            m2.metric("Eligible Unique Talent", eu)
            m3.metric("Shortlisted Candidates", su)
            st.dataframe(uniq_df)
            buf_b = io.BytesIO()
            with pd.ExcelWriter(buf_b, engine='openpyxl') as w: uniq_df.to_excel(w, index=False, sheet_name='Unique Applicants')
            st.download_button("📥 Export Unique Applicants", data=buf_b.getvalue(), file_name="unique_applicants.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
    except Exception as e: st.error(f"Error compiling funnel: {e}")
else: st.info("Awaiting raw dataset upload to generate funnel pipelines.")
