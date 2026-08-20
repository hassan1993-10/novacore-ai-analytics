from __future__ import annotations
import io, os, re, json
import pandas as pd
import streamlit as st
import excel_mcp
from copilot_agent import ask

st.set_page_config(page_title="NovaCore Copilot", page_icon="✦", layout="wide", initial_sidebar_state="expanded")
if "GITHUB_TOKEN" in st.secrets: os.environ["GITHUB_TOKEN"]=st.secrets["GITHUB_TOKEN"]
if "COPILOT_GITHUB_TOKEN" in st.secrets: os.environ["COPILOT_GITHUB_TOKEN"]=st.secrets["COPILOT_GITHUB_TOKEN"]

AR=re.compile(r"[\u0600-\u06FF]")
def is_ar(t): return len(AR.findall(str(t)))>len(re.findall(r"[A-Za-z]",str(t)))
def xlsx(df):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w: df.to_excel(w,index=False,sheet_name="Analysis")
    return b.getvalue()

st.markdown("""
<style>
#MainMenu,footer{visibility:hidden}[data-testid="stAppViewContainer"]{background:#f7f9fc}
[data-testid="stHeader"]{background:#fff;border-bottom:1px solid #e7ebf0}
[data-testid="stSidebar"]{background:#071a31;border-right:1px solid #102b4d}
[data-testid="stSidebar"] *{color:white}.block-container{max-width:1380px;padding-top:1.4rem}
.brand{font-size:20px;font-weight:800}.brand b{color:#2d8cff}.sub{font-size:12px;color:#9eb2c8!important;margin:4px 0 28px}
.nav{padding:12px 14px;border:1px solid #183754;border-radius:12px;margin:8px 0}.nav.active{background:#1769e8;border-color:#3385ff}
.hero{padding:45px 0 24px;text-align:center}.mark{width:66px;height:66px;margin:auto;border-radius:18px;background:linear-gradient(135deg,#1769e8,#22a6f2);display:flex;align-items:center;justify-content:center;color:#fff;font-size:28px;font-weight:900}
.hero h1{font-size:30px;margin:16px 0 5px}.hero p{color:#6b778c}.card,.side,.answer{background:white;border:1px solid #e1e6ed;border-radius:16px}
.card{padding:18px;min-height:118px}.side{padding:18px;margin-bottom:14px}.answer{padding:15px 18px;line-height:1.8}
.krow{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid #f0f2f5}.rtl{direction:rtl;text-align:right}.ltr{direction:ltr;text-align:left}
</style>""",unsafe_allow_html=True)

if "messages" not in st.session_state: st.session_state.messages=[]
if "last_df" not in st.session_state: st.session_state.last_df=None

with st.sidebar:
    st.markdown('<div class="brand">NovaCore Solutions <b>Copilot</b></div><div class="sub">Enterprise Analytics Platform</div>',unsafe_allow_html=True)
    if st.button("＋  New Chat",use_container_width=True):
        st.session_state.messages=[];st.session_state.last_df=None;st.rerun()
    st.markdown('<div class="nav active">✦ &nbsp; Chat</div><div class="nav">▦ &nbsp; Data Explorer</div><div class="nav">▥ &nbsp; Reports</div><div class="nav">◇ &nbsp; Insights</div>',unsafe_allow_html=True)
    st.markdown("<br><br><br><small>NovaCore Solutions<br><br>v1.1.0 · Copilot</small>",unsafe_allow_html=True)

main,side=st.columns([4.6,1.25],gap="large")
with main:
    if not st.session_state.messages:
        st.markdown('<div class="hero"><div class="mark">N</div><h1>How can I help with your data?</h1><p>Ask questions across sales, finance, HR, IT, procurement and operations.</p></div>',unsafe_allow_html=True)
        qs=["Compare revenue by year","Top vendors by spend","Headcount by department","IT SLA performance"]
        cs=st.columns(4)
        for i,q0 in enumerate(qs):
            with cs[i]:
                st.markdown(f'<div class="card"><b>{q0}</b><br><small>Analyze enterprise data</small></div>',unsafe_allow_html=True)
                if st.button("Ask",key=f"q{i}",use_container_width=True): st.session_state.pending=q0;st.rerun()
    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            d="rtl" if is_ar(m["content"]) else "ltr"
            st.markdown(f'<div class="answer {d}">{m["content"]}</div>',unsafe_allow_html=True)
            if m.get("data") is not None: st.dataframe(m["data"],use_container_width=True,hide_index=True)
    q=st.chat_input("Ask NovaCore Copilot...")
    if not q and st.session_state.get("pending"): q=st.session_state.pop("pending")
    if q:
        st.session_state.messages.append({"role":"user","content":q})
        hist=[{"role":m["role"],"content":m["content"]} for m in st.session_state.messages[-7:]]
        with st.spinner("Analyzing your data..."):
            try:
                out=ask(q,hist,excel_mcp);answer,df=out["answer"],out["data"]
            except Exception as e:
                answer="تعذر الاتصال بـ GitHub Copilot حاليًا. راجع صلاحية التوكن وسجل التطبيق." if is_ar(q) else "Could not connect to GitHub Copilot. Check token permissions and app logs."
                st.error(str(e));df=None
        st.session_state.messages.append({"role":"assistant","content":answer,"data":df})
        st.session_state.last_df=df;st.rerun()
    if st.session_state.last_df is not None and not st.session_state.last_df.empty:
        st.divider();a,b=st.columns(2)
        with a: st.download_button("Download CSV",st.session_state.last_df.to_csv(index=False).encode("utf-8-sig"),"novacore_analysis.csv","text/csv",use_container_width=True)
        with b: st.download_button("Download Excel",xlsx(st.session_state.last_df),"novacore_analysis.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

with side:
    try:
        ov=json.loads(excel_mcp.get_model_overview("novacore_enterprise_sample_data")).get("data",{})
        tables=ov.get("tables",[]);rows=sum(int(x.get("row_count",0)) for x in tables)
        emp=len(excel_mcp.load_table("novacore_enterprise_sample_data","Employees"))
        dep=len(excel_mcp.load_table("novacore_enterprise_sample_data","Departments"))
    except Exception: tables=[];rows=emp=dep=0
    st.markdown('<div class="side"><b>Data Overview</b>'
                f'<div class="krow"><span>Tables</span><b>{len(tables)}</b></div>'
                f'<div class="krow"><span>Departments</span><b>{dep}</b></div>'
                f'<div class="krow"><span>Employees</span><b>{emp}</b></div>'
                f'<div class="krow"><span>Records</span><b>{rows:,}</b></div></div>',unsafe_allow_html=True)
    st.markdown('<div class="side"><b>Copilot</b><br><br>● Ready<br><br><small>Model: Auto · Verified Excel calculations</small></div>',unsafe_allow_html=True)
    if st.button("Refresh Data",use_container_width=True):
        st.cache_data.clear()
        try: excel_mcp.clear_cache()
        except Exception: pass
        st.rerun()
