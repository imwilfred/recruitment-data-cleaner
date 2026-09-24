st.sidebar.markdown("---")
st.sidebar.header("🔍 Funnel Controls")
file = st.file_uploader("2. Upload Raw Candidate File", type=["csv", "xlsx"], key=f"up_{st.session_state['ukey']}")

if file is not None:
    if st.button("🗑️ Clear Candidate File & Restart", type="primary"):
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
        
        # FIX: Swapped out pandas vector syntax for absolute element-by-element parsing to eliminate DataFrame str clashes
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
        
        if 'Application Status' in df.columns:
            df['Rank'] = [sm.get(str(x).lower(), 12) for x in df['Application Status'].tolist()]
        else:
            df['Rank'] = 12

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
        
        names_list = [str(n).lower().replace(" ", "") for n in u_build['Candidate Name'].tolist()]
        nrics_list = [str(i) for i in u_build['NRIC Number'].tolist()]
        u_build['Comp_Key'] = [f"{names_list[idx]}_{nrics_list[idx]}" for idx in range(len(u_build))]
        
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

        t1, t2 = st.tabs(["📈 View A: Total Applications Funnel", "👥 View B: Unique Applicants Funnel"])
        with t1: tab("Total Applications Funnel", apps_df, selected_job, selected_domain)
        with t2: tab("Unique Applicants Funnel", uniq_df, selected_job, selected_domain)
    except Exception as e: st.error(f"Error compiling funnel: {e}")
else: st.info("Awaiting raw dataset upload to generate funnel pipelines.")
