import streamlit as st, pandas as pd, re, io
st.set_page_config(page_title="Funnel", layout="wide")
st.title("📊 Interactive Recruitment Funnel & Data Cleaner")

if "ukey" not in st.session_state: st.session_state["ukey"] = 0

def draw_funnel(t, e, s, i, o, h):
    st.markdown("### 🗺️ Visual Pipeline Funnel (Strict Sequential Step-Down)")
    p = lambda v: (v / t * 100) if t > 0 else 0
    stg = [
        {"n": "1. Total Inflow / Intake Pool", "v": t, "w": "100%", "c": "#1E88E5"},
        {"n": "2. Eligible Volume (Confirmed SG Citizens)", "v": e, "w": "85%", "c": "#2196F3"},
        {"n": "3. Advanced to Shortlist Stage", "v": s, "w": "70%", "c": "#42A5F5"},
        {"n": "4. Advanced to Interview Loop", "v": i, "w": "55%", "c": "#64B5F6"},
        {"n": "5. Advanced to Offer/Clearance", "v": o, "w": "40%", "c": "#90CAF9"},
        {"n": "6. Hired / Cleared Pool", "v": h, "w": "25%", "c": "#BBDEFB"}
    ]
    st.markdown("<div style='display:flex; flex-direction:column; align-items:center; justify-content:center; width:100%; margin:20px 0;'>", unsafe_allow_html=True)
    for s in stg:
        st.markdown(f"<div style='background-color:{s['c']}; width:{s['w']}; max-width:600px; margin:4px auto; padding:12px; border-radius:8px; text-align:center; color:#0D47A1; box-shadow:0 2px 4px rgba(0,0,0,0.1);'><strong style='font-size:15px;'>{s['n']}</strong><br/><span style='font-size:20px; font-weight:bold;'>{s['v']}</span> <span style='font-size:12px; font-style:italic;'>({p(s['v']):.1f}%)</span></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

def render_tab(title, data, job):
    t = len(data)
    el_df = data[data['Eligibility_Status'] == "Eligible"]
    e = len(el_df)
    s = len(el_df[el_df['Rank'] <= 9])
    i = len(el_df[el_df['Rank'] <= 7])
    o = len(el_df[el_df['Rank'] <= 5])
    h = len(el_df[el_df['Rank'] == 1])
    st.subheader(f"{title}: {'All Positions' if job == 'All Jobs' else job}")
    f_view, d_view = st.tabs(["🗺️ View Graphical Funnel Map", "📋 View Detailed Data Table"])
    with f_view: draw_funnel(t, e, s, i, o, h)
    with d_view:
        g = data.drop(columns=['Rank']) if 'Rank' in data.columns else data
        st.dataframe(g)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: g.to_excel(w, index=False, sheet_name='Data')
        st.download_button(f"📥 Export {title}", data=buf.getvalue(), file_name=f"{title.lower().replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

file = st.file_uploader("Upload raw file", type=["csv", "xlsx"], key=f"up_{st.session_state['ukey']}")

if file is not None:
    if st.button("🗑️ Clear Current File & Restart", type="primary"):
        st.session_state["ukey"] += 1
        st.rerun()
    try:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        cc = ['Candidate Name', 'Email Address', 'NRIC Number', 'Application Status', 'Job Name', 'Citizenship', 'Country Of Birth']
        for c in cc:
            if c in df.columns: df[c] = df[c].fillna("").astype(str).str.strip()
        if 'Candidate Name' in df.columns: df['Candidate Name'] = df['Candidate Name'].str.title()
        if 'Email Address' in df.columns: df['Email Address'] = df['Email Address'].str.lower()
        if 'NRIC Number' in df.columns: df['NRIC Number'] = df['NRIC Number'].str.upper()
        if 'Citizenship' in df.columns: df['Citizenship'] = df['Citizenship'].str.title()
        if 'Country Of Birth' in df.columns: df['Country Of Birth'] = df['Country Of Birth'].str.title()
        df['Job Name'] = df['Job Name'].replace("", "Unknown Role")
        df['X0PA Score'] = pd.to_numeric(df['X0PA Score'], errors='coerce').fillna(0)
        df['Total Exp'] = pd.to_numeric(df['Total Exp'], errors='coerce').fillna(0)

        def parse_co(x):
            if pd.isna(x) or not isinstance(x, str): return "Not Listed"
            m = re.search(r'C:\s*([^.\n]+)', x)
            return m.group(1).strip() if m else "Not Listed"
        df['Current_Company'] = df['Work Experience'].apply(parse_co) if 'Work Experience' in df.columns else "Not Listed"

        def parse_el(r):
            cz, cob = str(r.get('Citizenship', '')).lower(), str(r.get('Country Of Birth', '')).lower()
            if 'citizen' in cz or cz == 'singapore':
                if any(x in cob for x in ['china', 'myanmar', 'myanmr']): return "Ineligible (Exception)"
                return "Eligible"
            return "Ineligible"
        df['Eligibility_Status'] = df.apply(parse_el, axis=1)

        st_map = {
            "Hired": 1, "Hire in Progress": 2, "Offer in Progress": 3, "Verbal Offer in Progress": 4,
            "Salary Proposal in Progress": 5, "Interview in Progress": 6, "Interview Reject": 7,
            "Post Screening Slot in Progress": 8, "Post Screening Slot Reject": 9, "Screening in Progress": 10, "Screening Reject": 11
        }
        df['Rank'] = df['Application Status'].apply(lambda x: st_map.get(str(x).strip(), 12))

        def parse_edu(txt):
            defaults = {"l": "Not Provided", "d": "Not Listed", "s": "Not Listed"}
            if pd.isna(txt) or not isinstance(txt, str) or txt.strip() == "": return defaults
            parts = [p.strip() for p in txt.split('|') if p.strip()]
            u = txt.upper()
            phd, master, bach, dip, alev = "PHD" in u or "DOCTOR" in u, "MASTER" in u or "MSC" in u or "MBA" in u, "BACHELOR" in u or "DEGREE" in u or "BSC" in u or "BENG" in u, "DIPLOMA" in u or "POLYTECHNIC" in u, "A LEVEL" in u or "ADVANCED LEVEL" in u or "JUNIOR COLLEGE" in u
            lvl = "PhD" + (" ➔ Master" if master else "") + (" ➔ Bachelor" if bach else "") if phd else ("Master" + (" ➔ Bachelor" if bach else "") if master else ("Bachelor" if bach else (" & ".join([w for w, c in [("Diploma", dip), ("A-Levels", alev)] if c]) if (dip or alev) else "Other / School")))
            t_idx = next((i for i, p in enumerate(parts) if ("PHD" if phd else ("MASTER" if master else ("BACHELOR" if bach else "DIPLOMA"))) in p.upper()), 0)
            try:
                t_arr = parts[max(0, t_idx-1):min(len(parts), t_idx+5)]
                sch, disc = "Not Listed", "Not Listed"
                for item in t_arr:
                    if re.search(r'\d{4}', item) or re.match(r'^\d+(\.\d+)?$', item): continue
                    if any(k in item.upper() for k in ["UNIVERSITY", "POLYTECHNIC", "INSTITUTE", "COLLEGE", "SCHOOL", "NUS", "NTU", "SMU", "SIT", "SUTD", "SUSS"]): sch = item
                    elif not any(k in item.upper() for k in ["BACHELOR", "MASTER", "PHD", "DIPLOMA", "DEGREE", "HONOURS"]) and len(item) > 2 and disc == "Not Listed": disc = item
                return {"l": lvl, "d": re.sub(r'(Bachelor of|Master of|BSc|BEng|Diploma in|BSc Hons|Degree in)\s*', '', disc, flags=re.IGNORECASE).strip().title(), "s": sch.title()}
            except: return {"l": lvl, "d": "Not Listed", "s": "Not Listed"}

        if 'Candidate Education' in df.columns:
            mapped = df['Candidate Education'].apply(parse_edu)
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = [i['l'] for i in mapped], [i['d'] for i in mapped], [i['s'] for i in mapped]
        else: df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = "Not Provided", "Not Listed", "Not Listed"

        layout = ['Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 'Citizenship', 'Country Of Birth', 'Eligibility_Status', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Current_Company', 'Total Exp', 'X0PA Score', 'Application Status', 'App Date', 'Job Id', 'Job Name', 'Job Status', 'Application Source', 'Rank']
        av_cols = [c for c in layout if c in df.columns or c in ['Current_Company', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Eligibility_Status', 'Rank']]
        master_df = df[av_cols]

        st.sidebar.header("🔍 Funnel Controls")
        selected_job = st.sidebar.selectbox("Filter by Job Requisition", ["All Jobs"] + sorted(list(master_df['Job Name'].unique())))
        max_exp = float(master_df['Total Exp'].max()) if 'Total Exp' in master_df.columns and len(master_df) > 0 else 10.0
        min_exp_input = st.sidebar.slider("Minimum Years of Experience", 0.0, max_exp if not pd.isna(max_exp) else 10.0, 0.0, step=0.5)

        apps_df = master_df.copy().fillna("Not Provided")
        uniq_df = master_df.copy().sort_values(by=['Rank', 'X0PA Score'], ascending=[True, False])
        for col in ['Email Address', 'NRIC Number']:
            if col in uniq_df.columns: uniq_df[col] = uniq_df[col].replace(['nan', 'none', 'na', ''], pd.NA)
        if 'Email Address' in uniq_df.columns or 'NRIC Number' in uniq_df.columns:
            has_n = uniq_df[uniq_df['NRIC Number'].notna()]
            no_n = uniq_df[uniq_df['NRIC Number'].isna()]
            if 'NRIC Number' in uniq_df.columns: has_n = has_n.drop_duplicates(subset=['NRIC Number'], keep='first')
            uniq_df = pd.concat([has_n, no_n]).drop_duplicates(subset=['Email Address'], keep='first')

        if selected_job != "All Jobs":
            apps_df, uniq_df = apps_df[apps_df['Job Name'] == selected_job], uniq_df[uniq_df['Job Name'] == selected_job]
        apps_df, uniq_df = apps_df[apps_df['Total Exp'] >= min_exp_input], uniq_df[uniq_df['Total Exp'] >= min_exp_input].fillna("Not Provided")

        t1, t2 = st.tabs(["📈 View A: Total Applications Funnel", "👥 View B: Unique Applicants Funnel"])
        with t1: render_tab("Total Applications Funnel", apps_df, selected_job)
        with t2: render_tab("Unique Applicants Funnel", uniq_df, selected_job)
    except Exception as e: st.error(f"Error compiling funnel: {e}")
else: st.info("Awaiting raw dataset upload to generate funnel pipelines.")
