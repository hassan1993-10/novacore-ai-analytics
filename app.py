from __future__ import annotations

import html
import base64
import io
import json
import os
import re
import smtplib
import traceback
from email.message import EmailMessage
from pathlib import Path
from typing import Optional
from urllib.parse import quote

import pandas as pd
import streamlit as st

import excel_mcp
from copilot_agent import ask

try:
    import plotly.express as px
except Exception:
    px = None


# =========================================================
# Page / environment
# =========================================================

st.set_page_config(
    page_title="NovaCore Solutions Copilot",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "GITHUB_TOKEN" in st.secrets:
    os.environ["GITHUB_TOKEN"] = str(st.secrets["GITHUB_TOKEN"])
if "COPILOT_GITHUB_TOKEN" in st.secrets:
    os.environ["COPILOT_GITHUB_TOKEN"] = str(st.secrets["COPILOT_GITHUB_TOKEN"])

DATASET = "novacore_enterprise_sample_data"
APP_DIR = Path(__file__).resolve().parent
ASSETS_DIR = APP_DIR / "assets"

def _data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")

LOGO_URI = _data_uri(ASSETS_DIR / "novacore_logo.png")
BUILDING_URI = _data_uri(ASSETS_DIR / "company_building.png")



# =========================================================
# Helpers
# =========================================================

ARABIC_RE = re.compile(r"[\u0600-\u06FF]")


def is_arabic(text: str) -> bool:
    text = str(text or "")
    return len(ARABIC_RE.findall(text)) > len(re.findall(r"[A-Za-z]", text))


def direction(text: str) -> str:
    return "rtl" if is_arabic(text) else "ltr"


def formatted_html(text: str) -> str:
    """Safe lightweight markdown-like formatting for assistant text."""
    value = html.escape(str(text or ""))
    value = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", value)
    value = re.sub(r"`(.+?)`", r"<code>\1</code>", value)
    lines = value.splitlines()

    output = []
    in_list = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith(("- ", "• ")):
            if not in_list:
                output.append("<ul>")
                in_list = True
            output.append(f"<li>{stripped[2:]}</li>")
        else:
            if in_list:
                output.append("</ul>")
                in_list = False
            if stripped:
                output.append(f"<div>{stripped}</div>")
            else:
                output.append("<div style='height:6px'></div>")
    if in_list:
        output.append("</ul>")
    return "".join(output)


def render_message_text(text: str) -> None:
    d = direction(text)
    st.markdown(
        f'<div class="message-text {d}" dir="{d}">{formatted_html(text)}</div>',
        unsafe_allow_html=True,
    )


def excel_bytes(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Analysis")
    return out.getvalue()


def csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def is_time_dimension(col: str) -> bool:
    key = str(col).casefold()
    return any(x in key for x in ["year", "month", "quarter", "date", "time", "period"])


def numeric_columns(df: pd.DataFrame) -> list[str]:
    """Numeric measures. Time dimensions such as Year stay dimensions even if numeric."""
    result = []
    for col in df.columns:
        if is_time_dimension(col):
            continue
        if pd.api.types.is_numeric_dtype(df[col]):
            result.append(col)
            continue
        converted = pd.to_numeric(df[col], errors="coerce")
        if len(df) and converted.notna().mean() >= 0.85:
            result.append(col)
    return result


def dimension_columns(df: pd.DataFrame) -> list[str]:
    nums = set(numeric_columns(df))
    return [c for c in df.columns if c not in nums]


def render_kpis(df: pd.DataFrame) -> bool:
    """Render KPI cards for a one-row result with numeric values."""
    if df is None or df.empty or len(df) != 1:
        return False

    nums = numeric_columns(df)
    if not nums:
        return False

    cols = st.columns(min(len(nums), 4))
    for idx, col in enumerate(nums[:4]):
        val = pd.to_numeric(df.iloc[0][col], errors="coerce")
        if pd.isna(val):
            display = "-"
        elif abs(float(val)) >= 1000:
            display = f"{float(val):,.0f}"
        else:
            display = f"{float(val):,.2f}"
        with cols[idx]:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">{html.escape(str(col).replace("_", " "))}</div>
                    <div class="kpi-value">{display}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    return True


def render_chart(df: pd.DataFrame, user_question: str = "") -> bool:
    """Infer a useful chart from the verified result."""
    if df is None or df.empty or len(df) < 2:
        return False

    columns = list(df.columns)

    # 1) Prefer an explicit time field as X even when stored as integer (e.g. Year=2024).
    time_dims = [c for c in columns if is_time_dimension(c)]

    # 2) Find numeric candidates from every non-X column.
    def looks_numeric(col: str) -> bool:
        if pd.api.types.is_numeric_dtype(df[col]):
            return True
        converted = pd.to_numeric(df[col], errors="coerce")
        return bool(len(df) and converted.notna().mean() >= 0.85)

    if time_dims:
        x = time_dims[0]
        y_candidates = [c for c in columns if c != x and looks_numeric(c)]
    else:
        categorical = [c for c in columns if not looks_numeric(c)]
        x = categorical[0] if categorical else (columns[0] if len(columns) >= 2 else None)
        y_candidates = [c for c in columns if c != x and looks_numeric(c)]

    if not x or not y_candidates:
        return False

    # Usually the final numeric column is the aggregated measure returned by the agent.
    y = y_candidates[-1]

    chart_df = df.copy()
    chart_df[y] = pd.to_numeric(chart_df[y], errors="coerce")
    chart_df = chart_df.dropna(subset=[y])
    if chart_df.empty:
        return False

    # Keep chronological order for time series.
    if is_time_dimension(x):
        try:
            chart_df = chart_df.sort_values(x, ascending=True)
        except Exception:
            pass

    question = str(user_question or "")
    arabic = is_arabic(question)
    title = "التحليل المرئي" if arabic else "Visual Analysis"

    if px is None:
        st.line_chart(chart_df.set_index(x)[y]) if is_time_dimension(x) else st.bar_chart(chart_df.set_index(x)[y])
        return True

    is_trend = is_time_dimension(x) or any(
        token in question.casefold()
        for token in ["trend", "over time", "year", "month", "quarter", "سن", "شهر", "ربع", "اتجاه"]
    )

    if is_trend:
        fig = px.line(chart_df, x=x, y=y, markers=True, title=title)
        fig.update_traces(line_width=3, marker_size=9)
    elif len(chart_df) > 8:
        chart_df = chart_df.sort_values(y, ascending=True)
        fig = px.bar(chart_df, x=y, y=x, orientation="h", title=title)
    else:
        fig = px.bar(chart_df, x=x, y=y, title=title)

    fig.update_layout(
        margin=dict(l=20, r=20, t=48, b=20),
        height=390,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(size=13),
        title_font=dict(size=16),
        xaxis_title=str(x).replace("_", " "),
        yaxis_title=str(y).replace("_", " "),
        showlegend=False,
        hovermode="x unified" if is_trend else "closest",
    )
    fig.update_yaxes(tickformat=",")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    return True


def email_is_configured() -> bool:
    try:
        e = st.secrets["email"]
        return all(
            e.get(k)
            for k in ["smtp_host", "smtp_port", "username", "password", "from_email"]
        )
    except Exception:
        return False


def send_email(
    recipient: str,
    subject: str,
    body: str,
    df: Optional[pd.DataFrame],
) -> tuple[bool, str]:
    if not email_is_configured():
        return False, "Email service is not configured in Streamlit Secrets."

    recipient = recipient.strip()
    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", recipient):
        return False, "Invalid recipient email."

    cfg = st.secrets["email"]

    msg = EmailMessage()
    msg["From"] = str(cfg["from_email"])
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    if df is not None and not df.empty:
        msg.add_attachment(
            excel_bytes(df),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="novacore_analysis.xlsx",
        )

    try:
        with smtplib.SMTP(str(cfg["smtp_host"]), int(cfg["smtp_port"])) as server:
            server.starttls()
            server.login(str(cfg["username"]), str(cfg["password"]))
            server.send_message(msg)
        return True, "Email sent successfully."
    except Exception as exc:
        return False, f"Email failed: {exc}"


# =========================================================
# Styling — closer to approved mockup
# =========================================================

st.markdown(
    """
    <style>
    :root{--navy:#06182E;--blue:#176BFF;--blue2:#38B6FF;--text:#101828;--muted:#667085;--line:#E2E8F0;--bg:#F8FAFC}
    html,body,[data-testid="stAppViewContainer"]{background:var(--bg);color:var(--text)}
    #MainMenu,footer,[data-testid="stDecoration"]{visibility:hidden}
    [data-testid="stHeader"]{background:#fff;border-bottom:1px solid var(--line)}
    .block-container{max-width:1540px;padding-top:.55rem;padding-bottom:4.5rem}
    [data-testid="stSidebar"]{background:linear-gradient(180deg,#06182E,#051529);border-right:1px solid #12304E;min-width:292px}
    [data-testid="stSidebar"] *{color:#fff}
    [data-testid="stSidebar"] .stButton>button{background:linear-gradient(135deg,#125ED5,#2479FF)!important;color:#fff!important;border:1px solid #3185FF!important;border-radius:10px!important;min-height:52px!important;font-weight:750!important;font-size:15px!important}
    .brand-wrap{display:flex;align-items:center;gap:12px;margin:7px 3px 23px}.brand-logo{width:48px;height:41px;object-fit:contain;filter:saturate(1.25) contrast(1.08)}.brand-name{font-size:17px;font-weight:800;line-height:1.25}.brand-name span{color:#58B4FF}
    .sidebar-nav{display:flex;align-items:center;gap:13px;padding:12px 10px;margin:7px 1px;font-size:15px;color:#EAF2FC}.sidebar-nav .ico{width:26px;text-align:center;font-size:21px}
    .company-card{position:relative;margin-top:120px;min-height:165px;border:1px solid rgba(129,171,220,.22);border-radius:12px;padding:18px;background:rgba(5,25,48,.55);overflow:hidden}.company-card-title{font-size:14px;font-weight:800}.company-card-sub{font-size:11px;color:#C2D1E1!important;margin-top:6px}.company-building{position:absolute;right:4px;bottom:-18px;height:120px}.version-pill{position:absolute;left:18px;bottom:14px;font-size:10px;border:1px solid rgba(255,255,255,.28);border-radius:999px;padding:4px 8px}.copyright{font-size:10px;color:#9FB0C4!important;margin:14px 4px}
    .app-header{display:flex;justify-content:space-between;align-items:center;padding:4px 8px 15px}.header-left{display:flex;align-items:center;gap:20px}.hamb{font-size:24px;color:#17304D}.header-title{font-size:22px;font-weight:800}.header-title span{color:var(--blue)}.header-actions{display:flex;align-items:center;gap:18px}.status{display:flex;align-items:center;gap:7px;border:1px solid var(--line);border-radius:999px;padding:8px 13px;background:#fff;font-size:12px}.dot{width:7px;height:7px;border-radius:50%;background:#12B76A}.moon{font-size:22px}.avatar{width:38px;height:38px;border-radius:50%;background:#071B35;color:#fff;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800}
    .hero{text-align:center;padding:72px 0 42px}.hero-logo{width:95px;height:auto;filter:saturate(1.3) contrast(1.08);margin:auto}.hero h1{font-size:31px;margin:15px 0 5px;font-weight:800}.hero p{color:var(--muted);font-size:15px}
    .suggestion-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:8px 0 65px}.suggestion-card{display:block;text-decoration:none!important;color:var(--text)!important;background:#fff;border:1px solid var(--line);border-radius:14px;padding:22px 18px;min-height:174px;box-shadow:0 4px 12px rgba(16,24,40,.025)}.sicon{width:46px;height:46px;border-radius:11px;display:flex;align-items:center;justify-content:center;font-size:23px;margin-bottom:28px}.s1{background:#EDF5FF;color:#176BFF}.s2{background:#EAFBF2;color:#18B36B}.s3{background:#F3EDFF;color:#7B45E5}.s4{background:#FFF5E6;color:#F59E0B}.stitle{font-size:14px;font-weight:650;line-height:1.5}
    .right-card{background:#fff;border:1px solid var(--line);border-radius:14px;margin-bottom:15px;overflow:hidden}.right-head{display:flex;align-items:center;gap:10px;padding:14px 17px;border-bottom:1px solid var(--line);font-weight:800;font-size:14px}.rbody{padding:9px 17px}.metric{display:flex;align-items:center;gap:12px;padding:8px 0}.metric-icon{width:44px;height:44px;border-radius:11px;background:#F0F6FF;color:#176BFF;display:flex;align-items:center;justify-content:center;font-size:20px}.metric-num{font-size:15px;font-weight:800}.metric-lbl{font-size:11px;color:var(--muted)}.activity{display:flex;justify-content:space-between;gap:8px;padding:10px 0;font-size:11px;border-bottom:1px solid #F2F4F7}.activity:last-child{border-bottom:0}.activity-time{color:#7C8AA0;white-space:nowrap}
    [data-testid="stChatMessage"]{background:#fff;border:1px solid var(--line);border-radius:14px;padding:.55rem .72rem;box-shadow:0 4px 12px rgba(16,24,40,.025);margin-bottom:.65rem}.answer-text,.message-text{font-size:14px;line-height:1.9}.answer-text.rtl,.message-text.rtl{direction:rtl;text-align:right;font-family:Tahoma,Arial,sans-serif}.answer-text.ltr,.message-text.ltr{direction:ltr;text-align:left}.answer-text strong,.message-text strong{font-weight:800}.answer-text ul{margin:6px 18px;padding:0}
    [data-testid="stChatInput"]{max-width:1000px;margin:0 auto}[data-testid="stChatInput"]>div{background:#fff;border:1px solid #D9E2EC;border-radius:13px;min-height:64px}[data-testid="stChatInput"] textarea{unicode-bidi:plaintext;font-size:14px}.chat-tools{max-width:1000px;margin:-46px auto 5px;padding-left:16px;position:relative;z-index:5;display:flex;gap:17px;width:calc(100% - 20px);pointer-events:none}.chat-tools span{font-size:19px;color:#17304D}.chips{display:flex;justify-content:center;gap:10px;flex-wrap:wrap;margin:28px 0 34px}.chip{font-size:11px;color:#5F7188;border:1px solid #DFE7EF;background:#fff;border-radius:999px;padding:8px 14px}.security{text-align:center;color:#75849A;font-size:11px;margin-top:48px}
    .kpi-card{background:#fff;border:1px solid var(--line);border-radius:13px;padding:15px 17px}.kpi-name,.kpi-label{font-size:11px;color:var(--muted);margin-bottom:6px}.kpi-number,.kpi-value{font-size:24px;font-weight:850;direction:ltr}div[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:12px;overflow:hidden}div[data-testid="stPlotlyChart"]{background:#fff;border:1px solid var(--line);border-radius:13px;padding:8px}
    @media(max-width:1200px){.suggestion-row{grid-template-columns:repeat(2,1fr)}.company-card{margin-top:45px}}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# Session state
# =========================================================

if "messages" not in st.session_state:
    st.session_state.messages = []
if "last_df" not in st.session_state:
    st.session_state.last_df = None
if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""
if "recent_queries" not in st.session_state:
    st.session_state.recent_queries = []
if "email_open" not in st.session_state:
    st.session_state.email_open = False


prompt_param = st.query_params.get("prompt")
if prompt_param:
    st.session_state.pending_question = str(prompt_param)
    st.query_params.clear()

# =========================================================
# Sidebar
# =========================================================

with st.sidebar:
    st.markdown(f'<div class="brand-wrap"><img class="brand-logo" src="{LOGO_URI}"><div class="brand-name">NovaCore Solutions<br><span>Copilot</span></div></div>', unsafe_allow_html=True)
    if st.button("▢   New Chat     +", use_container_width=True):
        st.session_state.messages=[]; st.session_state.last_df=None; st.session_state.last_answer=""; st.rerun()
    st.markdown('<div class="sidebar-nav"><span class="ico">▦</span>Dashboard</div><div class="sidebar-nav"><span class="ico">◉</span>Data Explorer</div><div class="sidebar-nav"><span class="ico">▥</span>Reports</div><div class="sidebar-nav"><span class="ico">◌</span>Insights</div><div class="sidebar-nav"><span class="ico">♡</span>Saved Chats</div><div class="sidebar-nav"><span class="ico">⚙</span>Settings</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="company-card"><div class="company-card-title">NovaCore Solutions</div><div class="company-card-sub">Enterprise Analytics Platform</div><img class="company-building" src="{BUILDING_URI}"><div class="version-pill">v1.0.0</div></div><div class="copyright">© 2026 NovaCore Solutions</div>', unsafe_allow_html=True)


st.markdown('<div class="app-header"><div class="header-left"><span class="hamb">☰</span><div class="header-title">NovaCore Solutions <span>Copilot</span></div></div><div class="header-actions"><div class="status"><span class="dot"></span>All Systems Operational</div><span class="moon">◔</span><div class="avatar">NC</div></div></div>', unsafe_allow_html=True)

main_col, right_col = st.columns([4.7, 1.28], gap="large")

# =========================================================
# Main
# =========================================================

with main_col:
    if not st.session_state.messages:
        st.markdown(
            f"""
            <div class="hero">
                <img class="hero-logo" src="{LOGO_URI}">
                <h1>Hello! I’m NovaCore Copilot</h1>
                <p>Your AI assistant for enterprise data analysis</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        cards=[("▥","s1","Show total revenue<br>by department","Show total revenue by department"),("⌑","s2","Top 10 vendors by<br>purchase amount","Top 10 vendors by purchase amount"),("♙","s3","Headcount by<br>department","Headcount by department"),("◉","s4","Monthly expense<br>trend for this year","Monthly expense trend for this year")]
        cards_html='<div class="suggestion-row">'
        for icon,cls,title,q0 in cards:
            cards_html += f'<a class="suggestion-card" href="?prompt={quote(q0)}"><div class="sicon {cls}">{icon}</div><div class="stitle">{title}</div></a>'
        cards_html += '</div>'
        st.markdown(cards_html, unsafe_allow_html=True)

    # Conversation
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            render_message_text(message["content"])

            df = message.get("data")
            if df is not None and not df.empty:
                render_kpis(df)
                render_chart(df, message.get("question", ""))
                st.dataframe(
                    df,
                    use_container_width=True,
                    hide_index=True,
                    height=min(420, 38 + len(df) * 35),
                )

    question = st.chat_input("Ask a question about your data...")
    st.markdown('<div class="chat-tools"><span>⌕</span><span>▦</span></div>', unsafe_allow_html=True)
    if not st.session_state.messages:
        st.markdown('<div class="chips"><span class="chip">Sales this month</span><span class="chip">IT tickets by priority</span><span class="chip">Purchase by category</span><span class="chip">Employee turnover rate</span><span class="chip">Cash flow overview</span></div><div class="security">♙ &nbsp; Your data is secure and private. All analysis is performed using NovaCore enterprise data.</div>', unsafe_allow_html=True)

    if not question and st.session_state.get("pending_question"):
        question = st.session_state.pop("pending_question")

    if question:
        st.session_state.messages.append(
            {"role": "user", "content": question, "question": question}
        )
        st.session_state.recent_queries.insert(0, question)
        st.session_state.recent_queries = st.session_state.recent_queries[:5]

        history = [
            {"role": x["role"], "content": x["content"]}
            for x in st.session_state.messages[-8:]
        ]

        with st.spinner("Analyzing your data..."):
            try:
                output = ask(question, history, excel_mcp)
                answer = output["answer"]
                result_df = output["data"]
            except Exception as exc:
                print("\n=== NOVACORE COPILOT ERROR ===")
                traceback.print_exc()
                print("=== END NOVACORE COPILOT ERROR ===\n")
                answer = (
                    "تعذر إكمال التحليل حاليًا. راجع سجل التطبيق."
                    if is_arabic(question)
                    else "The analysis could not be completed. Check the app logs."
                )
                st.error(f"{type(exc).__name__}: {exc}")
                result_df = None

        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": answer,
                "data": result_df,
                "question": question,
            }
        )
        st.session_state.last_df = result_df
        st.session_state.last_answer = answer
        st.rerun()

    # Result actions
    if st.session_state.last_df is not None and not st.session_state.last_df.empty:
        st.divider()
        st.markdown(
            '<div class="result-actions">Export or share the latest verified result</div>',
            unsafe_allow_html=True,
        )

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button(
                "↓  Download CSV",
                csv_bytes(st.session_state.last_df),
                file_name="novacore_analysis.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "↓  Download Excel",
                excel_bytes(st.session_state.last_df),
                file_name="novacore_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c3:
            if st.button("✉  Send by Email", use_container_width=True):
                st.session_state.email_open = not st.session_state.email_open

        if st.session_state.email_open:
            with st.form("email_form"):
                email_to = st.text_input("Recipient email")
                email_subject = st.text_input(
                    "Subject", value="NovaCore Copilot Analysis"
                )
                email_body = st.text_area(
                    "Message",
                    value=st.session_state.last_answer,
                    height=120,
                )
                if st.form_submit_button("Send Email", use_container_width=True):
                    ok, msg = send_email(
                        email_to,
                        email_subject,
                        email_body,
                        st.session_state.last_df,
                    )
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)


# =========================================================
# Right panel
# =========================================================

with right_col:
    try:
        overview = json.loads(excel_mcp.get_model_overview(DATASET))
        od = overview.get("data", {}) if overview.get("success") else {}
        tables = od.get("tables", [])
        record_count = sum(int(x.get("row_count",0)) for x in tables)
        employee_count = len(excel_mcp.load_table(DATASET,"Employees"))
        department_count = len(excel_mcp.load_table(DATASET,"Departments"))
    except Exception:
        record_count=employee_count=department_count=0
    st.markdown(f'<div class="right-card"><div class="right-head"><span style="color:#176BFF;font-size:20px">▦</span>Data Overview</div><div class="rbody"><div class="metric"><div class="metric-icon">▣</div><div><div class="metric-num">1</div><div class="metric-lbl">Datasets</div></div></div><div class="metric"><div class="metric-icon">♧</div><div><div class="metric-num">{department_count}</div><div class="metric-lbl">Departments</div></div></div><div class="metric"><div class="metric-icon">♙</div><div><div class="metric-num">{employee_count:,}</div><div class="metric-lbl">Employees</div></div></div><div class="metric"><div class="metric-icon">☷</div><div><div class="metric-num">{record_count:,}</div><div class="metric-lbl">Records</div></div></div></div></div>', unsafe_allow_html=True)
    recent = st.session_state.get("recent_queries", [])
    defaults=[("Revenue by Department","2 min ago"),("Top Vendors Analysis","15 min ago"),("Headcount Overview","1 hour ago"),("IT Tickets Summary","2 hours ago"),("Monthly Expenses","3 hours ago")]
    if recent:
        activity=''.join(f'<div class="activity"><span>{html.escape(q[:34])}</span><span class="activity-time">just now</span></div>' for q in recent[:5])
    else:
        activity=''.join(f'<div class="activity"><span>{n}</span><span class="activity-time">{t}</span></div>' for n,t in defaults)
    st.markdown(f'<div class="right-card"><div class="right-head"><span style="color:#176BFF;font-size:20px">◷</span>Recent Activity</div><div class="rbody">{activity}</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="right-card"><div class="right-head"><span style="color:#176BFF;font-size:20px">ϟ</span>Quick Actions</div><div class="rbody">', unsafe_allow_html=True)
    st.button("Upload Data", use_container_width=True, disabled=True)
    if st.button("Data Dictionary", use_container_width=True):
        try:
            dd = excel_mcp.load_table(DATASET,"Data_Dictionary")
            st.dataframe(dd,use_container_width=True,hide_index=True)
        except Exception as exc: st.error(str(exc))
    st.markdown('</div></div>', unsafe_allow_html=True)
