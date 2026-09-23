import streamlit as st, pandas as pd, re, io
st.set_page_config(page_title="Funnel", layout="wide")
st.title("📊 Interactive Recruitment Funnel & Data Cleaner")

# Initialize persistent background session memory trackers for hot reloads
if "ukey" not in st.session_state: st.session_state["ukey"] = 0
if "map_df_stored" not in st.session_state: st.session_state["map_df_stored"] = None

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

def render_tab(title, data, job, domain):
    t = len(data)
    el_df = data[data['Eligibility_Status'] == "Eligible"]
    e = len(el_df)
    s = len(el_df[el_df['Rank'] <= 9]) if 'Rank' in el_df.columns else 0
    i = len(el_df[el_df['Rank'] <= 7]) if 'Rank' in el_df.columns else 0
    o = len(el_df[el_df['Rank'] <= 5]) if 'Rank' in el_df.columns else 0
    h = len(el_df[el_df['Rank'] == 1]) if 'Rank' in el_df.columns else 0
    st.subheader(f"{title}: {f'All Positions ({domain})' if job == 'All Jobs' else f'{job} ({domain})'}")
    f_view, d_view = st.tabs(["🗺️ View Graphical Funnel Map", "📋 View Detailed Data Table"])
    with f_view: draw_funnel(t, e, s, i, o, h)
    with d_view:
        g = data.drop(columns=['Rank']) if 'Rank' in data.columns else data
        st.dataframe(g)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: g.to_excel(w, index=False, sheet_name='Data')
        st.download_button(f"📥 Export {title}", data=buf.getvalue(), file_name=f"{title.lower().replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.sidebar.header("📁 Reference Uploads")
ref_file = st.sidebar.file_uploader("1. Upload Job Mapping File", type=["csv", "xlsx"])

if ref_file is not None:
    try:
        ref_df = pd.read_csv(ref_file) if ref_file.name.endswith('.csv') else pd.read_excel(ref_file)
        ref_df.columns = [str(c).replace('\xa0', ' ').strip().title() for c in ref_df.columns]
        
        dom_col = next((c for c in ref_df.columns if 'DOMAIN' in c.upper()), None)
        job_col = next((c for c in ref_df.columns if 'JOB' in c.upper()), None)
        
        if job_col and dom_col:
            ref_df['Job_Clean'] = ref_df[job_col].astype(str).str.replace('\xa0', ' ').str.replace('\u200b', ' ')
            ref_df['Job_Clean'] = ref_df['Job_Clean'].str.replace('–', '-').str.replace('—', '-').str.replace('‒', '-')
            ref_df['Job_Clean'] = ref_df['Job_Clean'].str.replace('[', '').str.replace(']', '').str.replace('(', '').str.replace(')', '')
            ref_df['Job_Clean'] = ref_df['Job_Clean'].str.replace(' ', '').str.replace('/', '').str.strip().str.upper()
            
            ref_df['Key_Len'] = ref_df['Job_Clean'].str.len()
            ref_df = ref_df.sort_values(by='Key_Len', ascending=False)
            
            st.session_state["map_df_stored"] = ref_df[['Job_Clean', dom_col]].rename(columns={dom_col: 'Extracted_Domain'}).drop_duplicates()
            st.sidebar.success("✅ Job Master Reference Linked!")
        else: st.sidebar.error("Error: Could not identify 'Job' or 'Domain' column headers.")
    except Exception as err: st.sidebar.error(f"Error: {err}")

if st.session_state["map_df_stored"] is None:
    st.sidebar.info("💡 Tip: Upload reference map above to unlock structural domain filters.")

st.sidebar.markdown("---")
st.sidebar.header("🔍 Funnel Controls")
file = st.file_uploader("2. Upload Raw Candidate File", type=["csv", "xlsx"], key=f"up_{st.session_state['ukey']}")
if file is not None:
    if st.button("🗑️ Clear Candidate File & Restart", type="primary"):
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
            "hired": 1, "hire in progress": 2, "offer in progress": 3, "verbal offer in progress": 4,
            "salary proposal in progress": 5, "interview in progress": 6, "interview reject": 7,
            "post screening slot in progress": 8, "post screening slot reject": 9, "screening in progress": 10, "screening reject": 11
        }
        df['Rank'] = df['Application Status'].apply(lambda x: st_map.get(str(x).strip().lower(), 12))

        # --- BULLETPROOF EXACT-FIRST SUBSTRING CONTAINER LOOKUP ENGINE ---
        def match_domain_row(jname):
            if st.session_state["map_df_stored"] is not None:
                c_key = str(jname).replace('\xa0', ' ').replace('\u200b', ' ')
                c_key = c_key.replace('–', '-').replace('—', '-').replace('‒', '-')
                c_key = c_key.replace('[', '').replace(']', '').replace('(', '').replace(')', '')
                c_key = c_key.replace(' ', '').replace('/', '').strip().upper()
                
                # FIRST PASS: Check for an absolute character-for-character EXACT match to preserve direct mappings
                for _, row in st.session_state["map_df_stored"].iterrows():
                    ref_key = str(row['Job_Clean']).replace('/', '')
                    if c_key == ref_key:
                        val = str(row['Extracted_Domain']).strip()
                        if val.upper() in ["PRM", "PCP", "ESP", "AESP", "IT"]: return val.upper()
                        return val.title()
                
                # SECOND PASS: Only check for partial substring containment if exact match fails
                for _, row in st.session_state["map_df_stored"].iterrows():
                    ref_key = str(row['Job_Clean']).replace('/', '')
                    if c_key in ref_key or ref_key in c_key:
                        val = str(row['Extracted_Domain']).strip()
                        if val.upper() in ["PRM", "PCP", "ESP", "AESP", "IT"]: return val.upper()
                        return val.title()
                        
                return "Unmapped Role"
            return "Map File Missing"
            
        df['Job_Domain'] = df['Job Name'].apply(match_domain_row)

        def parse_edu(txt):
            defaults = {"l": "Not Provided", "d": "Not Listed", "s": "Not Listed"}
            if pd.isna(txt) or not isinstance(txt, str) or txt.strip() == "": return defaults
            parts = [p.strip() for p in txt.split('|') if p.strip()]
            u = txt.upper()
            phd, master, bach, dip, alev = "PHD" in u or "DOCTOR" in u, "MASTER" in u or "MSC" in u or "MBA" in u, "BACHELOR" in u or "DEGREE" in u or "BSC" in u or "BENG" in u, "DIPLOMA" in u or "POLYTECHNIC" in u, "A LEVEL" in u or "ADVANCED LEVEL" in u or "JUNIOR COLLEGE" in u
            lvl = "PhD" + (" ➔ Master" if master else "") + (" ➔ Bachelor" if bach else "") if phd else ("Master" + (" ➔ Bachelor" if bach else "") if master else ("Bachelor" if bach else (" & ".join([w for w, c in [("Diploma", dip), ("A-Levels", alev)] if c]) if (dip or alev) else "Other / School")))
            sch, disc = "Not Listed", "Not Listed"
            skw = ["UNIVERSITY", "POLYTECHNIC", "INSTITUTE", "COLLEGE", "SCHOOL", "NUS", "NTU", "SMU", "SIT", "SUTD", "SUSS", "ACADEMY", "CENTRE", "CENTER", "FACULTY", "UNIVERSIDADE", "UNIVERSIDAD", "ECOLE", "UPF"]
            ikw = ["BACHELOR", "MASTER", "PHD", "DIPLOMA", "DEGREE", "HONOURS", "HONORS", "DISTINCTION", "CERTIFICATE", "BSC", "BENG", "MSC", "MBA", "CERTIFICATION", "GRADUATE", "EQUIVALENT"]
            for p in parts:
                if any(k in p.upper() for k in skw): sch = p; break
            longest_len = 0
            for p in parts:
                pu = p.upper()
                if p == sch or re.search(r'\d{4}', p) or re.match(r'^\d+(\.\d+)?$', p): continue
                if any(k in pu for k in ikw) or any(k in pu for k in skw) or len(p) <= 2: continue
                if len(p) > longest_len: longest_len = len(p); disc = p
            if disc != "Not Listed":
                disc = re.sub(r'^(Bachelor of|Master of|BSc|BEng|Diploma in|BSc Hons|Degree in|Tecnologo Em|Tecnólogo Em)\s*', '', disc, flags=re.IGNORECASE)
                disc = re.sub(r'[\-,]\s*(Honours|Honors|Distinction|Graduation).*$', '', disc, flags=re.IGNORECASE).strip()
                for k in ["NTU", "NUS", "SMU", "SIT", "SUSS", "SUTD"]:
                    if disc.endswith(k): disc = disc[:-len(k)].strip()
            return {"l": lvl, "d": disc.title() if disc != "Not Listed" else "Not Listed", "s": sch.title() if sch != "Not Listed" else "Not Listed"}

        if 'Candidate Education' in df.columns:
            mapped = df['Candidate Education'].apply(parse_edu)
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = [i['l'] for i in mapped], [i['d'] for i in mapped], [i['s'] for i in mapped]
        else: df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = "Not Provided", "Not Listed", "Not Listed"

        layout = ['Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 'Citizenship', 'Country Of Birth', 'Eligibility_Status', 'Job_Domain', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Current_Company', 'Total Exp', 'X0PA Score', 'Application Status', 'App Date', 'Job Id', 'Job Name', 'Job Status', 'Application Source', 'Rank']
        av_cols = [c for c in layout if c in df.columns or c in ['Current_Company', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Eligibility_Status', 'Rank', 'Job_Domain']]
        master_df = df[av_cols]

        domain_list = ["All Domains"] + sorted(list(master_df['Job_Domain'].unique()))
        selected_domain = st.sidebar.selectbox("Filter by Domain Track", domain_list)
        
        filtered_job_source = master_df if selected_domain == "All Domains" else master_df[master_df['Job_Domain'] == selected_domain]
        job_list = ["All Jobs"] + sorted(list(filtered_job_source['Job Name'].unique()))
        selected_job = st.sidebar.selectbox("Filter by Job Requisition", job_list)
        
        max_exp = float(master_df['Total Exp'].max()) if 'Total Exp' in master_df.columns and len(master_df) > 0 else 10.0
        min_exp_input = st.sidebar.slider("Minimum Years of Experience", 0.0, max_exp if not pd.isna(max_exp) else 10.0, 0.0, step=0.5)

        apps_df = master_df.copy().fillna("Not Provided")
        u_build = master_df.copy().sort_values(by=['Rank', 'X0PA Score'], ascending=[True, False])
        u_build['NRIC Number'] = u_build['NRIC Number'].replace("", pd.NA)
        u_build['Comp_Key'] = u_build['Candidate Name'].astype(str).str.lower().str.replace(" ", "") + "_" + u_build['NRIC Number'].astype(str)
        
        has_nric = u_build[u_build['NRIC Number'].notna()].drop_duplicates(subset=['Comp_Key'], keep='first')
        no_nric = u_build[u_build['NRIC Number'].isna()]
        pass1_df = pd.concat([has_nric, no_nric])
        pass1_df['Email Address'] = pass1_df['Email Address'].replace("", pd.NA)
        uniq_df = pass1_df.drop_duplicates(subset=['Email Address'], keep='first')
        if 'Comp_Key' in uniq_df.columns: uniq_df = uniq_df.drop(columns=['Comp_Key'])

        if selected_domain != "All Domains":
            apps_df, uniq_df = apps_df[apps_df['Job_Domain'] == selected_domain], uniq_df[uniq_df['Job_Domain'] == selected_domain]
        if selected_job != "All Jobs":
            apps_df, uniq_df = apps_df[apps_df['Job Name'] == selected_job], uniq_df[uniq_df['Job Name'] == selected_job]
        apps_df, uniq_df = apps_df[apps_df['Total Exp'] >= min_exp_input], uniq_df[uniq_df['Total Exp'] >= min_exp_input].fillna("Not Provided")

        t1, t2 = st.tabs(["📈 View A: Total Applications Funnel", "👥 View B: Unique Applicants Funnel"])
        with t1: render_tab("Total Applications Funnel", apps_df, selected_job, selected_domain)
        with t2: render_tab("Unique Applicants Funnel", uniq_df, selected_job, selected_domain)
