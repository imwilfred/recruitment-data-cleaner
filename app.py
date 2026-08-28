import streamlit as st, pandas as pd, re, io
st.set_page_config(page_title="Funnel Builder", layout="wide")
st.title("📊 Interactive Recruitment Funnel & Data Cleaner")

# Helper to normalize standard text strings safely
def clean_str(series, case="title"):
    if case == "title": return series.astype(str).str.strip().str.title()
    if case == "upper": return series.astype(str).str.strip().str.upper()
    return series.astype(str).str.strip().str.lower()

# Shared UI Render Block to display metrics, dataframes, and excel downloads
def render_funnel_tab(title, data, job_name):
    st.subheader(f"{title}: {job_name}")
    t = len(data)
    e = len(data[data['Eligibility_Status'] == "Eligible"])
    s = len(data[data['Funnel_Stage'].isin(["Offered Stage", "Interview Stage", "Shortlisted Stage", "Hired"])])
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Count", t)
    c2.metric("Eligible (SG Citizens)", e)
    c3.metric("Shortlisted & Beyond", s)
    
    st.dataframe(data)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w: 
        data.to_excel(w, index=False, sheet_name='Data Export')
    st.download_button(f"📥 Export {title} to Excel", data=buf.getvalue(), file_name=f"{title.lower().replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

file = st.file_uploader("Upload raw CSV or Excel file", type=["csv", "xlsx"])

if file is not None:
    try:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        
        # Clean text columns and spaces
        for col, mode in [('Candidate Name','title'), ('Email Address','lower'), ('NRIC Number','upper'), ('Citizenship','title'), ('Country Of Birth','title'), ('Application Status','none')]:
            if col in df.columns: df[col] = clean_str(df[col], mode)
            
        df['Job Name'] = df['Job Name'].fillna("Unknown Role").astype(str).str.strip()
        df['X0PA Score'] = pd.to_numeric(df['X0PA Score'], errors='coerce').fillna(0)
        df['Total Exp'] = pd.to_numeric(df['Total Exp'], errors='coerce').fillna(0)

        # Parse current employer text strings
        def parse_co(x):
            if pd.isna(x) or not isinstance(x, str): return "Not Listed"
            m = re.search(r'C:\s*([^.\n]+)', x)
            return m.group(1).strip() if m else "Not Listed"
        df['Current_Company'] = df['Work Experience'].apply(parse_co) if 'Work Experience' in df.columns else "Not Listed"

        # SG Citizenship Eligibility Mapping
        def parse_el(r):
            cz, cob = str(r.get('Citizenship', '')).lower(), str(r.get('Country Of Birth', '')).lower()
            if 'citizen' in cz or cz == 'singapore':
                return "Ineligible (Exception)" if any(x in cob for x in ['china', 'myanmar', 'myanmr']) else "Eligible"
            return "Ineligible"
        df['Eligibility_Status'] = df.apply(parse_el, axis=1)

        # Linear Lifecycle Status Priority Mapping
        st_map = {
            "Hired": 1, "Hire in Progress": 2, "Offer in Progress": 3,
            "Verbal Offer in Progress": 4, "Salary Proposal in Progress": 5,
            "Interview in Progress": 6, "Interview Reject": 7,
            "Post Screening Slot in Progress": 8, "Post Screening Slot Reject": 9,
            "Screening in Progress": 10, "Screening Reject": 11
        }
        df['Rank'] = df['Application Status'].apply(lambda x: st_map.get(str(x).strip(), 12))

        # Classify lifecycle categories
        def parse_st(r):
            s = str(r.get('Application Status', '')).strip()
            if s == "Hired": return "Hired"
            if "Offer" in s or "Salary" in s or s == "Hire in Progress": return "Offered Stage"
            if "Interview" in s: return "Interview Stage"
            if "Slot" in s: return "Shortlisted Stage"
            return "Rejected Pool" if "Reject" in s else "Screening Pool"
        df['Funnel_Stage'] = df.apply(parse_st, axis=1)

        # Education, Field and School Extractor
        def parse_edu(txt):
            defaults = {"l": "Not Provided", "d": "Not Listed", "s": "Not Listed"}
            if pd.isna(txt) or not isinstance(txt, str) or txt.strip() == "": return defaults
            parts = [p.strip() for p in txt.split('|') if p.strip()]
            u = txt.upper()
            phd, master, bach, dip, alev = "PHD" in u or "DOCTOR" in u, "MASTER" in u or "MSC" in u or "MBA" in u, "BACHELOR" in u or "DEGREE" in u or "BSC" in u or "BENG" in u, "DIPLOMA" in u or "POLYTECHNIC" in u, "A LEVEL" in u or "ADVANCED LEVEL" in u or "JUNIOR COLLEGE" in u
            lvl = "PhD" + (" ➔ Master" if master else "") + (" ➔ Bachelor" if bach else "") if phd else ("Master" + (" ➔ Bachelor" if bach else "") if master else ("Bachelor" if bach else (" & ".join([w for w, c in [("Diploma", dip), ("A-Levels", alev)] if c]) if (dip or alev) else "Other / School")))
            t_kw = "PHD" if phd else ("MASTER" if master else ("BACHELOR" if bach else "DIPLOMA"))
            t_idx = next((i for i, p in enumerate(parts) if t_kw in p.upper()), 0)
            try:
                t_arr = parts[max(0, t_idx-1):min(len(parts), t_idx+5)]
                sch, disc = "Not Listed", "Not Listed"
                for item in t_arr:
                    if re.search(r'\d{4}', item) or re.match(r'^\d+(\.\d+)?$', item): continue
                    if any(k in item.upper() for k in ["UNIVERSITY", "POLYTECHNIC", "INSTITUTE", "COLLEGE", "SCHOOL", "NUS", "NTU", "SMU", "SIT", "SUTD", "SUSS"]): sch = item
                    elif not any(k in item.upper() for k in ["BACHELOR", "MASTER", "PHD", "DIPLOMA", "DEGREE", "HONOURS", "DISTINCTION"]) and len(item) > 2 and disc == "Not Listed": disc = item
                disc = re.sub(r'(Bachelor of|Master of|BSc|BEng|Diploma in|BSc Hons|Degree in)\s*', '', disc, flags=re.IGNORECASE).strip()
                return {"l": lvl, "d": disc.title(), "s": sch.title()}
            except: return {"l": lvl, "d": "Not Listed", "s": "Not Listed"}

        if 'Candidate Education' in df.columns:
            mapped = df['Candidate Education'].apply(parse_edu)
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = [i['l'] for i in mapped], [i['d'] for i in mapped], [i['s'] for i in mapped]
        else:
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = "Not Provided", "Not Listed", "Not Listed"

        # Clean display layout
        layout = ['Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 'Citizenship', 'Country Of Birth', 'Eligibility_Status', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Current_Company', 'Total Exp', 'X0PA Score', 'Funnel_Stage', 'Application Status', 'App Date', 'Job Id', 'Job Name', 'Job Status', 'Application Source', 'Rank']
        av_cols = [c for c in layout if c in df.columns or c in ['Current_Company', 'Funnel_Stage', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Eligibility_Status', 'Rank']]
        master_df = df[av_cols]

        # --- SIDEBAR FILTERS CONTROLS ---
        st.sidebar.header("🔍 Funnel Controls")
        job_list = ["All Jobs"] + sorted(list(master_df['Job Name'].unique()))
        selected_job = st.sidebar.selectbox("Filter by Specific Job Requisition", job_list)
        max_exp = float(master_df['Total Exp'].max()) if 'Total Exp' in master_df.columns else 30.0
        min_exp_input = st.sidebar.slider("Minimum Years of Experience", 0.0, max_exp, 0.0, step=0.5)

        # --- DUAL CORE DATA MATRIX BUILDER ---
        apps_df = master_df.copy().fillna("Not Provided")
        
        uniq_df = master_df.copy().sort_values(by=['Rank', 'X0PA Score'], ascending=[True, False])
        for col in ['Email Address', 'NRIC Number']:
            if col in uniq_df.columns: uniq_df[col] = uniq_df[col].replace(['nan', 'none', 'na', ''], pd.NA)
        if 'Email Address' in uniq_df.columns or 'NRIC Number' in uniq_df.columns:
            has_n = uniq_df[uniq_df['NRIC Number'].notna()]
            no_n = uniq_df[uniq_df['NRIC Number'].isna()]
            if 'NRIC Number' in uniq_df.columns: has_n = has_n.drop_duplicates(subset=['NRIC Number'], keep='first')
            uniq_df = pd.concat([has_n, no_n]).drop_duplicates(subset=['Email Address'], keep='first')

        # Apply Global Interactive Controls
        if selected_job != "All Jobs":
            apps_df = apps_df[apps_df['Job Name'] == selected_job]
            uniq_df = uniq_df[uniq_df['Job Name'] == selected_job]
        apps_df = apps_df[apps_df['Total Exp'] >= min_exp_input]
        uniq_df = uniq_df[uniq_df['Total Exp'] >= min_exp_input].fillna("Not Provided")

        if 'Rank' in apps_df.columns: apps_df = apps_df.drop(columns=['Rank'])
        if 'Rank' in uniq_df.columns: uniq_df = uniq_df.drop(columns=['Rank'])

        # --- GRAPHICAL UI TABS ---
        t1, t2 = st.tabs(["📈 View A: Total Applications Funnel", "👥 View B: Unique Applicants Funnel"])
        with t1: render_funnel_tab("Total Applications Funnel", apps_df, selected_job)
        with t2: render_funnel_tab("Unique Applicants Funnel", uniq_df, selected_job)
            
    except Exception as e: st.error(f"Error compiling funnel: {e}")
else: st.info("Awaiting raw dataset upload to generate funnel pipelines.")
