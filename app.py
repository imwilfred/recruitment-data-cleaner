import streamlit as st, pandas as pd, re, io

st.set_page_config(page_title="Funnel", layout="wide")
st.title("📊 Interactive Recruitment Funnel & Data Cleaner")

if "ukey" not in st.session_state: st.session_state["ukey"] = 0
if "m_df" not in st.session_state: st.session_state["m_df"] = None

DOMAINS = {"digital": "Digital", "engineering": "Engineering", "prm": "PRM", "corporate": "Corporate"}


# ---------- helpers ----------
def job_key(s):
    """Letters/digits only, upper-case: ignores spaces, '/', brackets, commas, dashes, nbsp."""
    return re.sub(r'[\W_]+', '', str(s)).upper()


def is_eligible(cz):
    """EDIT HERE if your Citizenship values differ."""
    u = str(cz).strip().upper()
    if u in ("SINGAPORE", "SINGAPOREAN", "SG"): return True
    return "CITIZEN" in u and not any(w in u for w in ("NON", "FOREIGN"))


def tidy_domain(v):
    """Clean stray spaces / zero-width characters; standardise the four main domains, keep anything else as written."""
    v = re.sub(r'[\u200b\xa0\s]+', ' ', str(v)).strip()
    return DOMAINS.get(v.lower(), v)


def get_dom(jn):
    """Returns (domain, mapping title used). Title is None for an exact match, so partial matches can be reviewed."""
    m, t = st.session_state["m_df"], st.session_state.get("m_titles", {})
    if not m: return "Unmapped", None
    ck = job_key(jn)
    if ck in m: return m[ck], None                                 # 1. exact match after cleaning
    inside = [k for k in m if len(k) >= 12 and k in ck]            # 2. mapping title contained in raw title (e.g. ", CIO Office" added)
    if inside:
        k = max(inside, key=len)
        return m[k], t.get(k)
    return "Unmapped", None


def dedupe(d):
    """Keep one row per applicant: furthest-progressed application (lowest Rank).
    Pass 1 = NRIC, Pass 2 = Email. Rows with a blank identifier are never merged on it."""
    d = d.sort_values('Rank', kind='stable')
    for col in ['NRIC Number', 'Email Address']:
        if col in d.columns:
            blank = d[col].isna() | (d[col].astype(str).str.strip() == "")
            d = pd.concat([d[~blank].drop_duplicates(subset=[col], keep='first'), d[blank]])
            d = d.sort_values('Rank', kind='stable')
    return d


def fnl(t, e, s, i, o, h):
    st.markdown("### 🔻 Visual Pipeline Funnel (Strict Sequential Step-Down)")
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
    for x in stg:
        st.markdown(f"<div style='background-color:{x['c']}; width:{x['w']}; max-width:600px; margin:4px auto; padding:12px; border-radius:8px; text-align:center; color:#0D47A1; box-shadow:0 2px 4px rgba(0,0,0,0.1);'><strong style='font-size:15px;'>{x['n']}</strong><br/><span style='font-size:20px; font-weight:bold;'>{x['v']}</span> <span style='font-size:12px;'>({p(x['v']):.1f}%)</span></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def edu(x):
    d = {"l": "Not Provided", "d": "Not Listed", "s": "Not Listed"}
    if pd.isna(x) or not isinstance(x, str) or x.strip() == "": return d
    pt = [p.strip() for p in x.split('|') if p.strip()]
    u = x.upper()
    phd, mstr, bach, dip, alev = "PHD" in u or "DOCTOR" in u, "MASTER" in u or "MSC" in u or "MBA" in u, "BACHELOR" in u or "DEGREE" in u or "BSC" in u or "BENG" in u, "DIPLOMA" in u or "POLYTECHNIC" in u, "A LEVEL" in u or "ADVANCED LEVEL" in u or "JUNIOR COLLEGE" in u
    lvl = "PhD" + (" > Master" if mstr else "") + (" > Bachelor" if bach else "") if phd else ("Master" + (" > Bachelor" if bach else "") if mstr else ("Bachelor" if bach else (" & ".join([w for w, c in [("Diploma", dip), ("A-Levels", alev)] if c]) if (dip or alev) else "Other / School")))
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


def tab(title, data, job, dom, key):
    t = len(data)
    el = data[data['Eligibility_Status'] == "Eligible"]
    e = len(el)
    has_rank = 'Rank' in el.columns
    s = len(el[el['Rank'] <= 9]) if has_rank else 0
    i = len(el[el['Rank'] <= 7]) if has_rank else 0
    o = len(el[el['Rank'] <= 5]) if has_rank else 0
    h = len(el[el['Rank'] == 1]) if has_rank else 0
    st.subheader(f"{title}: {f'All Positions ({dom})' if job == 'All Jobs' else f'{job} ({dom})'}")
    fv, dv = st.tabs(["🔻 View Graphical Funnel Map", "📋 View Detailed Data Table"])
    with fv: fnl(t, e, s, i, o, h)
    with dv:
        g = data.drop(columns=['Rank']) if 'Rank' in data.columns else data
        st.dataframe(g)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: g.to_excel(w, index=False, sheet_name='Data')
        st.download_button(f"📥 Export {title}", data=buf.getvalue(), file_name=f"{title.lower().replace(' ', '_')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key)


def clean_str_col(df, col, fn):
    if col in df.columns:
        df[col] = [fn(str(v).strip()) if pd.notna(v) else "" for v in df[col].tolist()]


# ---------- job mapping upload ----------
st.info(
    "**How to use this tool**  \n"
    "**Step 1** - Upload the **Job Mapping file** (left box). All its sheets are read; the newest year wins.  \n"
    "**Step 2** - Upload the **Raw Candidate file** (right box). Only its first sheet is read.  \n"
    "**Step 3** - Use the filters in the left sidebar, switch between **View A (Total)** and **View B (Unique)**, and export any table to Excel.  \n"
    "Accepted formats: .xlsx, .xls, .csv"
)
c1, c2 = st.columns(2)
r_file = c1.file_uploader("Step 1: Upload the Job Mapping file", type=["csv", "xlsx", "xls"])

if r_file is not None:
    try:
        sheets = {"csv": pd.read_csv(r_file)} if r_file.name.endswith('.csv') else pd.read_excel(r_file, sheet_name=None)
        year = lambda n: int((re.findall(r'(?:19|20)\d{2}', n) or ['0'])[0])
        mp, titles, used = {}, {}, 0
        for sn in sorted(sheets, key=year):            # oldest first, so the newest sheet wins any conflict
            sd = sheets[sn]
            sd.columns = [str(c).replace('\xa0', ' ').strip() for c in sd.columns]
            jcol = next((c for c in sd.columns if 'JOB' in c.upper() and 'DOMAIN' not in c.upper()), None)
            dcol = next((c for c in sd.columns if 'DOMAIN' in c.upper()), None)
            if not (jcol and dcol): continue
            used += 1
            for jb, dm in sd[[jcol, dcol]].dropna().itertuples(index=False):
                k, d = job_key(jb), tidy_domain(dm)
                if k and d: mp[k], titles[k] = d, str(jb).strip()
        if used:
            st.session_state["m_df"], st.session_state["m_titles"] = mp, titles
            c1.success(f"✅ Job mapping loaded ({used} sheet(s); newest sheet overrides older)")
            odd = sorted(set(mp.values()) - set(DOMAINS.values()))
            if odd: c1.info(f"Other domain values in mapping file: {', '.join(odd)}")
        else:
            c1.error("No sheet has both a 'Job' column and a 'Domain' column")
    except Exception as err: c1.error(f"Error reading map: {err}")

if st.session_state["m_df"] is None:
    c1.info("💡 Upload the Job Mapping file to fill in the Job Domain column.")

st.sidebar.header("🎛️ Filters")
st.sidebar.caption("Filters apply to both views once the Raw Candidate file is loaded.")
file = c2.file_uploader("Step 2: Upload the Raw Candidate file", type=["csv", "xlsx", "xls"], key=f"up_{st.session_state['ukey']}")

# ---------- main ----------
if file is not None:
    if st.button("🗑️ Clear Candidate File & Restart", type="primary"):
        st.session_state["ukey"] += 1
        st.rerun()
    try:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        ccnt, clist = {}, []
        for c in df.columns:
            c_clean = str(c).strip()
            ccnt[c_clean] = ccnt.get(c_clean, 0) + 1
            clist.append(f"{c_clean}.{ccnt[c_clean]-1}" if ccnt[c_clean] > 1 else c_clean)
        df.columns = clist

        if 'Job Name' not in df.columns:
            t_jb = next((c for c in df.columns if 'JOB' in c.upper() and 'DOMAIN' not in c.upper()), None)
            if t_jb: df = df.rename(columns={t_jb: 'Job Name'})
        if 'Job Name' not in df.columns: df['Job Name'] = "Unknown Role"
        if 'Citizenship' not in df.columns: df['Citizenship'] = ""

        clean_str_col(df, 'Candidate Name', str.title)
        clean_str_col(df, 'Email Address', str.lower)
        clean_str_col(df, 'NRIC Number', str.upper)
        clean_str_col(df, 'Citizenship', str.title)
        clean_str_col(df, 'Country Of Birth', str.title)
        clean_str_col(df, 'Application Status', str)
        df['Job Name'] = [str(v).strip() if pd.notna(v) else "Unknown Role" for v in df['Job Name'].tolist()]

        sm = {
            "hired": 1, "hire in progress": 2, "offer in progress": 3, "verbal offer in progress": 4,
            "salary proposal in progress": 5, "interview in progress": 6, "interview reject": 7,
            "post screening slot in progress": 8, "post screening slot reject": 9, "screening in progress": 10, "screening reject": 11
        }
        df['Rank'] = [sm.get(str(x).strip().lower(), 12) for x in df['Application Status'].tolist()] if 'Application Status' in df.columns else 12

        df['Eligibility_Status'] = ["Eligible" if is_eligible(cz) else "Ineligible" for cz in df['Citizenship'].tolist()]
        res = {j: get_dom(j) for j in df['Job Name'].unique()}
        df['Job Domain'] = [res[j][0] for j in df['Job Name'].tolist()]

        if st.session_state["m_df"] is not None:
            um = sorted(j for j, r in res.items() if r[0] == "Unmapped")
            pm = sorted(f"{j}  →  {r[1]}" for j, r in res.items() if r[1])
            if um:
                with st.sidebar.expander(f"⚠️ {len(um)} unmapped job names"): st.write(um)
            if pm:
                with st.sidebar.expander(f"🔎 {len(pm)} matched by partial title - please review"): st.write(pm)

        if 'Candidate Education' in df.columns:
            ed = [edu(x) for x in df['Candidate Education'].tolist()]
            df['Highest_Education'] = [x['l'] for x in ed]
            df['Primary_Discipline'] = [x['d'] for x in ed]
            df['Institution'] = [x['s'] for x in ed]
        else:
            df['Highest_Education'], df['Primary_Discipline'], df['Institution'] = "Not Provided", "Not Listed", "Not Listed"

        lo = ['Candidate Name', 'Email Address', 'NRIC Number', 'Mobile Number', 'Citizenship', 'Country Of Birth', 'Eligibility_Status',
              'Job Name', 'Job Domain', 'Highest_Education', 'Primary_Discipline', 'Institution', 'Application Status', 'Rank']
        master_df = df[[c for c in lo if c in df.columns]]

        # sidebar filters
        domain_list = ["All Domains"] + sorted(master_df['Job Domain'].unique().tolist())
        selected_domain = st.sidebar.selectbox("Filter by Domain Track", domain_list)
        src = master_df if selected_domain == "All Domains" else master_df[master_df['Job Domain'] == selected_domain]
        selected_job = st.sidebar.selectbox("Filter by Job Requisition", ["All Jobs"] + sorted(src['Job Name'].unique().tolist()))

        # two outcomes
        apps_df = master_df.copy()
        uniq_df = dedupe(master_df.copy())

        if selected_domain != "All Domains":
            apps_df = apps_df[apps_df['Job Domain'] == selected_domain]
            uniq_df = uniq_df[uniq_df['Job Domain'] == selected_domain]
        if selected_job != "All Jobs":
            apps_df = apps_df[apps_df['Job Name'] == selected_job]
            uniq_df = uniq_df[uniq_df['Job Name'] == selected_job]

        apps_df, uniq_df = apps_df.fillna("Not Provided"), uniq_df.fillna("Not Provided")

        t1, t2 = st.tabs(["📄 View A: Total Applications Funnel", "👤 View B: Unique Applicants Funnel"])
        with t1: tab("Total Applications Funnel", apps_df, selected_job, selected_domain, "dl_total")
        with t2: tab("Unique Applicants Funnel", uniq_df, selected_job, selected_domain, "dl_unique")
    except Exception as e: st.error(f"Error compiling funnel: {e}")
else: st.info("Awaiting raw dataset upload to generate funnel pipelines.")
