import streamlit as st, pandas as pd, re, io
st.set_page_config(page_title="Funnel", layout="wide")
st.title("?? Interactive Recruitment Funnel & Data Cleaner")

if "ukey" not in st.session_state: st.session_state["ukey"] = 0
if "m_df" not in st.session_state: st.session_state["m_df"] = None

def fnl(t, e, s, i, o, h):
    st.markdown("### ??? Visual Pipeline Funnel (Strict Sequential Step-Down)")
    p = lambda v: (v / t * 100) if t > 0 else 0
    stg = [
        {"n": "1. Total Inflow / Intake Pool", "v": t, "w": "100%", "c": "#1E88E5"},
        {"n": "2. Eligible Volume (Confirmed SG Citizens)", "v": e, "w": "85%", "c": "#2196F3"},
        {"n": "3. Advanced to Shortlist Stage", "v": s, "w": "70%", "c": "#42A5F5"},
        {"n": "4. Advanced to Interview Loop", "v": i, "w": "55%", "c": "#64B5F6"},
        {"n": "5. Advanced to Offer/Clearance", "v": o, "w": "40%", "c": "#90CAF9"},
        {"n": "6. Hired / Cleared Pool", "v": h, "w": "25%", "c": "#BBDEFB"}
    ]
    st.markdown("<div style='display:flex; flex-direction:column; align-items:center; width:100%;'>", unsafe_allow_html=True)
    for s in stg:
        st.markdown(f"<div style='background-color:{s['c']}; width:{s['w']}; max-width:600px; margin:4px auto; padding:12px; border-radius:8px; text-align:center; color:#0D47A1; box-shadow:0 2px 4px rgba(0,0,0,0.1);'><strong style='font-size:15px;'>{s['n']}</strong><br/><span style='font-size:20px; font-weight:bold;'>{s['v']}</span> <span style='font-size:12px;'>({p(s['v']):.1f}%)</span></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

def edu(x):
    d = {"l": "Not Provided", "d": "Not Listed", "s": "Not Listed"}
    if pd.isna(x) or not isinstance(x, str) or x.strip() == "": return d
    pt = [p.strip() for p in x.split('|') if p.strip()]
    u = x.upper()
    phd, mstr, bach, dip, alev = "PHD" in u or "DOCTOR" in u, "MASTER" in u or "MSC" in u or "MBA" in u, "BACHELOR" in u or "DEGREE" in u or "BSC" in u or "BENG" in u, "DIPLOMA" in u or "POLYTECHNIC" in u, "A LEVEL" in u or "ADVANCED LEVEL" in u or "JUNIOR COLLEGE" in u
    lvl = "PhD" + (" ? Master" if mstr else "") + (" ? Bachelor" if bach else "") if phd else ("Master" + (" ? Bachelor" if bach else "") if mstr else ("Bachelor" if bach else (" & ".join([w for w, c in [("Diploma", dip), ("A-Levels", alev)] if c]) if (dip or alev) else "Other / School")))
    s, ds = "Not Listed", "Not Listed"
    sk = ["UNIVERSITY", "POLYTECHNIC", "INSTITUTE", "COLLEGE", "SCHOOL", "NUS", "NTU", "SMU", "SIT", "SUTD", "SUSS", "ACADEMY", "CENTRE", "CENTER", "FACULTY", "UNIVERSIDADE", "UNIVERSIDAD", "ECOLE", "UPF"]
    ik = ["BACHELOR", "MASTER", "PHD", "DIPLOMA", "DEGREE", "HONOURS", "HONORS", "DISTINCTION", "CERTIFICATE", "BSC", "BENG", "MSC", "MBA", "CERTIFICATION", "GRADUATE", "EQUIVALENT"]
    for p in pt:
        if any(k in p.upper() for k in sk): s = p; break
    l_len = 0
    for p in pt:
        pu = p.upper()
        if p == s or re.search(r'\d{4}', p) or re.match(r'^\d+(\.\d+)?$', p): continue
        if any(k in pu for k in ik) or any(k in pu for k in sk) or len(p) <= 2: continue
        if len(p) > l_len: l_len = len(p); ds = p
    if ds != "Not Listed":
        ds = re.sub(r'^(Bachelor of|Master of|BSc|BEng|Diploma in|BSc Hons|Degree in|Tecnologo Em|Tecnólogo Em)\s*', '', ds, flags=re.IGNORECASE)
        ds = re.sub(r'[\-,]\s*(Honours|Honors|Distinction|Graduation).*$', '', ds, flags=re.IGNORECASE).strip()
        for k in ["NTU", "NUS", "SMU", "SIT", "SUSS", "SUTD"]:
            if ds.endswith(k): ds = ds[:-len(k)].strip()
    return {"l": lvl, "d": ds.title() if ds != "Not Listed" else "Not Listed", "s": s.title() if s != "Not Listed" else "Not Listed"}

def tab(title, data, job, dom):
    t = len(data)
    el = data[data['Eligibility_Status'] == "Eligible"]
    e = len(el)
    s = len(el[el['Rank'] <= 9]) if 'Rank' in el.columns else 0
    i = len(el[el['Rank'] <= 7]) if 'Rank' in el.columns else 0
    o = len(el[el['Rank'] <= 5]) if 'Rank' in el.columns else 0
    h = len(el[el['Rank'] == 1]) if 'Rank' in el.columns else 0
    st.subheader(f"{title}: {f'All Positions ({dom})' if job == 'All Jobs' else f'{job} ({dom})'}")
    fv, dv = st.tabs(["??? View Graphical Funnel Map", "?? View Detailed Data Table"])
    with fv: fnl(t, e, s, i, o, h)
    with dv:
        g = data.drop(columns=['Rank']) if 'Rank' in data.columns else data
        st.dataframe(g)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: g.to_excel(w, index=False, sheet_name='Data')
        st.download_button(f"?? Export {title}", data=buf.getvalue(), file_name=f"{title.lower().replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def clean_txt_key(val_str):
    k = str(val_str).replace('\xa0', ' ').replace('\u200b', ' ')
    k = k.replace('–', '-').replace('—', '-').replace('-', '-')
    k = k.replace('[', '').replace(']', '').replace('(', '').replace(')', '')
    return k.replace(' ', '').replace('/', '').strip().upper()
st.sidebar.header("?? Reference Uploads")
r_file = st.sidebar.file_uploader("1. Upload Job Mapping File", type=["csv", "xlsx"])

if r_file is not None:
    try:
        r_df = pd.read_csv(r_file) if r_file.name.endswith('.csv') else pd.read_excel(r_file)
        rcnt = {}
        new_cols = []
        for c in r_df.columns:
            cn = str(c).replace('\xa0', ' ').strip().title()
            rcnt[cn] = rcnt.get(cn, 0) + 1
            new_cols.append(f"{cn}_{rcnt[cn]}" if rcnt[cn] > 1 else cn)
        r_df.columns = new_cols
        dm_c = next((c for c in r_df.columns if 'DOMAIN' in c.upper()), None)
        job_header = next((c for c in r_df.columns if 'JOB' in c.upper()), None)
        if job_header and dm_c:
            clean_keys = [clean_txt_key(str(v)) for v in r_df[job_header].tolist()]
            domain_vals = [str(v).strip() for v in r_df[dm_c].tolist()]
            st.session_state["m_df"] = dict(zip(clean_keys, domain_vals))
            st.sidebar.success("? Job Master Reference Linked!")
    except Exception as err: st.sidebar.error(f"Error reading map: {err}")

if st.session_state["m_df"] is None:
    st.sidebar.info("?? Tip: Upload reference map above to unlock structural domain filters.")

st.sidebar.markdown("---")
st.sidebar.header("?? Funnel Controls")
file = st.file_uploader("2. Upload Raw Candidate File", type=["csv", "xlsx"], key=f"up_{st.session_state['ukey']}")

if file is not None:
    if st.button("??? Clear Candidate File & Restart", type="primary"):
        st.session_state["ukey"] += 1
        st.rerun()
    try:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        ccnt = {}
        clist = []
        for c in df.columns:
            c_clean = str(c).strip()
            ccnt[c_clean] = ccnt.get(c_clean, 0) + 1
            clist.append(f"{c_clean}.{ccnt[c_clean]-1}" if ccnt[c_clean] > 1 else c_clean)
        df.columns = clist
        t_jb = next((c for c in df.columns if 'JOB' in c.upper() and 'DOMAIN' not in c.upper()), 'Job Name')
        if t_jb != 'Job Name': df = df.rename(columns={t_jb: 'Job Name'})
        
        if 'Candidate Name' in df.columns:
            df['Candidate Name'] = [str(v).strip().title() if pd.notna(v) else "" for v in df['Candidate Name'].tolist()]
        if 'Email Address' in df.columns:
            df['Email Address'] = [str(v).strip().lower() if pd.notna(v) else "" for v in df['Email Address'].tolist()]
        if 'NRIC Number' in df.columns:
            df['NRIC Number'] = [str(v).strip().upper() if pd.notna(v) else "" for v in df['NRIC Number'].tolist()]
        if 'Citizenship' in df.columns:
            df['Citizenship'] = [str(v).strip().title() if pd.notna(v) else "" for v in df['Citizenship'].tolist()]
        if 'Country Of Birth' in df.columns:
            df['Country Of Birth'] = [str(v).strip().title() if pd.notna(v) else "" for v in df['Country Of Birth'].tolist()]
        if 'Application Status' in df.columns:
            df['Application Status'] = [str(v).strip() if pd.notna(v) else "" for v in df['Application Status'].tolist()]
        if 'Job Name' in df.columns:
            df['Job Name'] = [str(v).strip() if pd.notna(v) else "Unknown Role" for v in df['Job Name'].tolist()]

        sm = {
            "hired": 1, "hire in progress": 2, "offer in progress": 3, "verbal offer in progress": 4,
            "salary proposal in progress": 5, "interview in progress": 6, "interview reject": 7,
            "post screening slot in progress": 8, "post screening slot reject": 9, "screening in progress": 10, "screening reject": 11
        }
        df['Rank'] = [sm.get(str(x).lower(), 12) for x in df['Application Status'].tolist()] if 'Application Status' in df.columns else 12

        def get_dom(jn):
            ck = clean_txt_key(jn)
            if st.session_state["m_df"] is not None:
                if ck in st.session_state["m_df"]:
                    v = str(st.session_state["m_df"][ck]).strip()
                    return v.upper() if v.upper() in ["PRM", "PCP", "ESP", "AESP", "IT"] else v.title()
                for ref_k, domain_v in st.session_state["m_df"].items():
                    rk = str(ref_k).replace('/', '')
                    if rk in ck or ck in rk:
                        v = str(domain_v).strip()
                        return v.upper() if v.upper() in ["PRM", "PCP", "ESP", "AESP", "IT"] else v.title()
            ju = str(jn).upper()
            if any(k in ju for k in ["SOFTWARE", "DEVELOPER", "CYBER", "CLOUD", "DATA", "AI", "ROBOTICS", "NETWORK", "SERVER", "UX", "DEVSECOPS"]): return "Digital"
            if any(k in ju for k in ["ARCHITECT", "BUILDING", "INFRASTRUCTURE", "TECHNICAL", "SURVEYOR", "AEROSPACE", "NAVAL", "MARINE", "ARMAMENT", "RADAR", "SENSOR", "VEHICLE"]): return "Engineering"
            if any(k in ju for k in ["HUMAN RESOURCE", "CORPORATE PLANNING", "FINANCIAL", "AUDIT", "OUTREACH", "COMMUNICATIONS", "BENEFITS", "SCHOLARSHIP"]): return "Corporate"
            return "Engineering" if ("ENGINEER" in ju or "ANALYST" in ju) else "Corporate"
            
        df['Eligibility_Status'] = ["Eligible" if "CITIZEN" in str(cz).upper() or str(cz).upper() == 'SINGAPORE' else "Ineligible" for cz in df['Citizenship'].tolist()]
        df['Job_Domain'] = [get_dom(j) for j in df['Job Name'].tolist()]

        if 'Candidate Education' in df.columns:
            df['Highest_Education'] = [edu(x)['l'] for x in df['Candidate Education'].tolist()]
            df['Primary_Discipline'] = [edu(x)['d'] for x in df['Candidate Education'].tolist()]
            df['Institution'] = [edu(x)['s'] for x in df['Candidate Education'].tolist()]
        else: df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = "Not Provided", "Not Listed", "Not Listed"

        lo = ['Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 'Citizenship', 'Country Of Birth', 'Eligibility_Status', 'Job Name', 'Job_Domain', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Application Status', 'Rank']
        av = [c for c in lo if c in df.columns or c in ['Highest_Education', 'Primary_Discipline', 'Institution', 'Eligibility_Status', 'Rank', 'Job_Domain', 'Job Name']]
        master_df = df[av]

        domain_list = ["All Domains"] + sorted(list(master_df['Job_Domain'].unique()))
        selected_domain = st.sidebar.selectbox("Filter by Domain Track", domain_list)
        filtered_job_source = master_df if selected_domain == "All Domains" else master_df[master_df['Job_Domain'] == selected_domain]
        job_list = ["All Jobs"] + sorted(list(filtered_job_source['Job Name'].unique())) if 'Job Name' in filtered_job_source.columns else ["All Jobs"]
        selected_job = st.sidebar.selectbox("Filter by Job Requisition", job_list)

        apps_df = master_df.copy().fillna("Not Provided")
        u_build = master_df.copy().sort_values(by=['Rank'], ascending=[True])
        u_build['NRIC Number'] = u_build['NRIC Number'].replace("", pd.NA)
        u_build['Comp_Key'] = [f"{str(n).lower().replace(' ', '')}_{str(i)}" for n, i in zip(u_build['Candidate Name'].tolist(), u_build['NRIC Number'].tolist())]
        has_nric = u_build[u_build['NRIC Number'].notna()].drop_duplicates(subset=['Comp_Key'], keep='first')
        no_nric = u_build[u_build['NRIC Number'].isna()]
        pass1_df = pd.concat([has_nric, no_nric])
        pass1_df['Email Address'] = pass1_df['Email Address'].replace("", pd.NA)
        uniq_df = pass1_df.drop_duplicates(subset=['Email Address'], keep='first')
        if 'Comp_Key' in uniq_df.columns: uniq_df = uniq_df.drop(columns=['Comp_Key'])

        if selected_domain != "All Domains":
            apps_df = apps_df[apps_df['Job_Domain'] == selected_domain]
            uniq_df = uniq_df[uniq_df['Job_Domain'] == selected_domain]
        if 'Job Name' in apps_df.columns and selected_job != "All Jobs":
            apps_df = apps_df[apps_df['Job Name'] == selected_job]
            uniq_df = uniq_df[uniq_df['Job Name'] == selected_job]

        if 'Rank' in apps_df.columns: apps_df = apps_df.drop(columns=['Rank'])
        if 'Rank' in uniq_df.columns: uniq_df = uniq_df.drop(columns=['Rank'])

        t1, t2 = st.tabs(["?? View A: Total Applications Funnel", "?? View B: Unique Applicants Funnel"])
        with t1: tab("Total Applications Funnel", apps_df, selected_job, selected_domain)
        with t2: tab("Unique Applicants Funnel", uniq_df, selected_job, selected_domain)
    except Exception as e: st.error(f"Error compiling funnel: {e}")
else: st.info("Awaiting raw dataset upload to generate funnel pipelines.")
