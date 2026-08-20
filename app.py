
from __future__ import annotations

import io
import json
import re
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import streamlit as st

import excel_mcp


# =========================================================
# Page config
# =========================================================

st.set_page_config(
    page_title="NovaCore Solutions Copilot",
    page_icon="N",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_DATASET = "novacore_enterprise_sample_data"


# =========================================================
# Language / direction helpers
# =========================================================

ARABIC_RE = re.compile(r"[\u0600-\u06FF]")


def detect_language(text: str) -> str:
    text = str(text or "")
    arabic_chars = len(ARABIC_RE.findall(text))
    latin_chars = len(re.findall(r"[A-Za-z]", text))

    if arabic_chars > latin_chars:
        return "ar"
    return "en"


def text_direction(text: str) -> str:
    return "rtl" if detect_language(text) == "ar" else "ltr"


def safe_html(value: Any) -> str:
    import html
    return html.escape(str(value or ""))


def render_text(text: str, role: str = "assistant") -> None:
    direction = text_direction(text)
    align = "right" if direction == "rtl" else "left"

    st.markdown(
        f"""
        <div class="answer-block {direction}" dir="{direction}" style="text-align:{align}">
            {safe_html(text).replace(chr(10), "<br>")}
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# Excel helpers
# =========================================================

@st.cache_data(show_spinner=False)
def load_dataset_overview(dataset_name: str) -> dict:
    payload = json.loads(excel_mcp.get_model_overview(dataset_name))
    return payload if isinstance(payload, dict) else {}


@st.cache_data(show_spinner=False)
def load_sheet(dataset_name: str, table_name: str) -> pd.DataFrame:
    return excel_mcp.load_table(dataset_name, table_name).copy()


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Result") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


# =========================================================
# Email
# =========================================================

def email_configured() -> bool:
    try:
        return bool(
            st.secrets["email"]["smtp_host"]
            and st.secrets["email"]["smtp_port"]
            and st.secrets["email"]["username"]
            and st.secrets["email"]["password"]
            and st.secrets["email"]["from_email"]
        )
    except Exception:
        return False


def recipient_allowed(recipient: str) -> tuple[bool, str]:
    recipient = str(recipient or "").strip()

    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", recipient):
        return False, "Invalid email address."

    try:
        allowed_domains = list(st.secrets["email"].get("allowed_domains", []))
    except Exception:
        allowed_domains = []

    if allowed_domains:
        domain = recipient.rsplit("@", 1)[-1].lower()
        normalized = {str(d).strip().lower() for d in allowed_domains}
        if domain not in normalized:
            return False, "This email domain is not allowed."

    return True, ""


def send_result_email(
    recipient: str,
    subject: str,
    body: str,
    result_df: Optional[pd.DataFrame] = None,
) -> tuple[bool, str]:

    if not email_configured():
        return False, "Email service is not configured."

    valid, message = recipient_allowed(recipient)
    if not valid:
        return False, message

    cfg = st.secrets["email"]

    msg = EmailMessage()
    msg["From"] = cfg["from_email"]
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    if result_df is not None and not result_df.empty:
        attachment = dataframe_to_excel_bytes(result_df, "Analysis")
        msg.add_attachment(
            attachment,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="novacore_analysis.xlsx",
        )

    try:
        with smtplib.SMTP(cfg["smtp_host"], int(cfg["smtp_port"])) as server:
            server.starttls()
            server.login(cfg["username"], cfg["password"])
            server.send_message(msg)

        return True, "Email sent successfully."

    except Exception as exc:
        return False, f"Email could not be sent: {exc}"


# =========================================================
# Simple pre-LLM demo analytics
# This will be replaced by GitHub Models Agent in the next step.
# =========================================================

def demo_analysis(question: str) -> tuple[str, Optional[pd.DataFrame]]:
    q = question.casefold()
    lang = detect_language(question)

    try:
        if ("revenue" in q or "sales" in q or "مبيعات" in q or "إيراد" in q or "ايراد" in q):
            df = load_sheet(DEFAULT_DATASET, "Sales_Orders")
            won = df[df["Order_Status"].astype(str).eq("Won")].copy()
            result = (
                won.groupby("Region", dropna=False)["Net_Revenue_SAR"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
                .rename(columns={"Net_Revenue_SAR": "Revenue_SAR"})
            )

            total = float(result["Revenue_SAR"].sum())
            if lang == "ar":
                summary = f"إجمالي الإيرادات للطلبات المكتسبة هو {total:,.0f} ريال. الجدول أدناه يوضح الإيرادات حسب المنطقة."
            else:
                summary = f"Total revenue from won orders is SAR {total:,.0f}. The table below shows revenue by region."

            return summary, result

        if ("ticket" in q or "sla" in q or "تذكرة" in q or "تذاكر" in q):
            df = load_sheet(DEFAULT_DATASET, "IT_Tickets")
            result = (
                df["SLA_Met"]
                .fillna("Unknown")
                .value_counts()
                .rename_axis("SLA_Status")
                .reset_index(name="Tickets")
            )

            if lang == "ar":
                summary = "هذا ملخص حالة الالتزام باتفاقية مستوى الخدمة لتذاكر تقنية المعلومات."
            else:
                summary = "Here is the IT ticket SLA compliance summary."

            return summary, result

        if ("vendor" in q or "supplier" in q or "purchase" in q or "مورد" in q or "مشتريات" in q):
            df = load_sheet(DEFAULT_DATASET, "Purchases")
            result = (
                df.groupby("Supplier_Name", dropna=False)["Total_Amount_SAR"]
                .sum()
                .sort_values(ascending=False)
                .head(10)
                .reset_index()
                .rename(columns={"Total_Amount_SAR": "Purchase_Amount_SAR"})
            )

            if lang == "ar":
                summary = "هذه أعلى 10 جهات توريد حسب إجمالي قيمة أوامر الشراء."
            else:
                summary = "These are the top 10 suppliers by total purchase amount."

            return summary, result

        if ("headcount" in q or "employee" in q or "موظف" in q or "الموظفين" in q):
            df = load_sheet(DEFAULT_DATASET, "Employees")
            result = (
                df.groupby("Department_Name", dropna=False)["Employee_ID"]
                .nunique()
                .sort_values(ascending=False)
                .reset_index(name="Headcount")
            )

            if lang == "ar":
                summary = "هذا توزيع عدد الموظفين حسب الإدارة."
            else:
                summary = "Here is the employee headcount by department."

            return summary, result

        if lang == "ar":
            return (
                "تم تجهيز الواجهة. في الخطوة التالية سنربط GitHub Models حتى أتمكن من فهم أي سؤال تحليلي وتنفيذ الأدوات المناسبة تلقائيًا.",
                None,
            )

        return (
            "The interface is ready. In the next step, GitHub Models will be connected so I can understand any analytical question and automatically use the appropriate tools.",
            None,
        )

    except Exception as exc:
        if lang == "ar":
            return f"حدث خطأ أثناء قراءة البيانات: {exc}", None
        return f"An error occurred while reading the data: {exc}", None


# =========================================================
# Styling
# =========================================================

st.markdown(
    """
    <style>
    :root {
        --navy: #07182E;
        --navy-2: #0B2445;
        --blue: #176BFF;
        --blue-soft: #EDF5FF;
        --text: #101828;
        --muted: #667085;
        --line: #E4E7EC;
        --surface: #FFFFFF;
        --bg: #F8FAFC;
        --green: #12B76A;
    }

    html, body, [data-testid="stAppViewContainer"] {
        background: var(--bg);
        color: var(--text);
    }

    [data-testid="stHeader"] {
        background: rgba(255,255,255,.95);
        border-bottom: 1px solid var(--line);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, var(--navy), #061326);
        border-right: 1px solid #102C50;
    }

    [data-testid="stSidebar"] * {
        color: #FFFFFF;
    }

    .block-container {
        max-width: 1480px;
        padding-top: 1.1rem;
        padding-bottom: 4rem;
    }

    .brand {
        font-size: 1.2rem;
        font-weight: 800;
        margin-bottom: .15rem;
    }

    .brand span {
        color: #56A8FF;
    }

    .brand-sub {
        color: #9FB4CC;
        font-size: .78rem;
        margin-bottom: 1.25rem;
    }

    .nav-card {
        border: 1px solid rgba(255,255,255,.10);
        background: rgba(255,255,255,.04);
        border-radius: 14px;
        padding: .72rem .85rem;
        margin-bottom: .45rem;
    }

    .nav-card.active {
        background: linear-gradient(135deg, #0F56CC, #2478FF);
        border-color: #2C7EFF;
    }

    .top-title {
        font-size: 1.45rem;
        font-weight: 800;
        margin-bottom: .1rem;
    }

    .top-title span {
        color: var(--blue);
    }

    .hero {
        text-align: center;
        padding: 3.2rem 1rem 1.5rem;
    }

    .logo-mark {
        display: inline-flex;
        width: 74px;
        height: 74px;
        align-items: center;
        justify-content: center;
        border-radius: 20px;
        color: #FFFFFF;
        font-weight: 900;
        font-size: 2rem;
        background: linear-gradient(135deg,#0C55E8,#22B5FF);
        box-shadow: 0 15px 35px rgba(23,107,255,.22);
    }

    .hero h1 {
        font-size: 2.05rem;
        margin: .9rem 0 .25rem;
    }

    .hero p {
        color: var(--muted);
        font-size: 1rem;
    }

    .suggestion {
        min-height: 142px;
        padding: 1rem;
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 16px;
        box-shadow: 0 6px 18px rgba(16,24,40,.04);
    }

    .side-panel {
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 16px;
        padding: 1rem;
        margin-bottom: .8rem;
    }

    .panel-title {
        font-weight: 800;
        margin-bottom: .85rem;
    }

    .metric-row {
        display:flex;
        justify-content:space-between;
        align-items:center;
        border-bottom:1px solid #F1F3F5;
        padding:.5rem 0;
    }

    .metric-row:last-child {
        border-bottom:0;
    }

    .metric-label {
        color:var(--muted);
        font-size:.82rem;
    }

    .metric-value {
        font-weight:800;
    }

    .answer-block {
        background:#FFFFFF;
        border:1px solid var(--line);
        border-radius:15px;
        padding:1rem 1.1rem;
        line-height:1.75;
        margin-bottom:.6rem;
        unicode-bidi: plaintext;
    }

    .answer-block.rtl {
        direction: rtl;
        font-family: Arial, "Tahoma", sans-serif;
    }

    .answer-block.ltr {
        direction: ltr;
    }

    div[data-testid="stDataFrame"] {
        direction:ltr;
        border:1px solid var(--line);
        border-radius:14px;
        overflow:hidden;
    }

    [data-testid="stChatInput"] textarea {
        unicode-bidi: plaintext;
    }

    .security-note {
        text-align:center;
        color:var(--muted);
        font-size:.78rem;
        padding:1rem;
    }

    .status-pill {
        display:inline-block;
        border:1px solid #DDE5ED;
        border-radius:999px;
        padding:.4rem .7rem;
        background:#FFF;
        font-size:.8rem;
    }

    .status-dot {
        display:inline-block;
        width:7px;
        height:7px;
        border-radius:50%;
        background:var(--green);
        margin-right:6px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# State
# =========================================================

if "messages" not in st.session_state:
    st.session_state.messages = []

if "last_result_df" not in st.session_state:
    st.session_state.last_result_df = None

if "last_result_text" not in st.session_state:
    st.session_state.last_result_text = ""


# =========================================================
# Sidebar
# =========================================================

with st.sidebar:
    st.markdown('<div class="brand">NovaCore Solutions <span>Copilot</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-sub">Enterprise Analytics Platform</div>', unsafe_allow_html=True)

    if st.button("＋  New Chat", use_container_width=True):
        st.session_state.messages = []
        st.session_state.last_result_df = None
        st.session_state.last_result_text = ""
        st.rerun()

    st.markdown('<div class="nav-card active">💬 &nbsp; Chat</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-card">◫ &nbsp; Data Explorer</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-card">▥ &nbsp; Reports</div>', unsafe_allow_html=True)
    st.markdown('<div class="nav-card">✦ &nbsp; Insights</div>', unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.caption("NovaCore Solutions")
    st.caption("v1.0.0")


# =========================================================
# Main layout
# =========================================================

header_left, header_right = st.columns([4, 1])
with header_left:
    st.markdown('<div class="top-title">NovaCore Solutions <span>Copilot</span></div>', unsafe_allow_html=True)
with header_right:
    st.markdown('<div class="status-pill"><span class="status-dot"></span>All Systems Operational</div>', unsafe_allow_html=True)

main_col, right_col = st.columns([4.3, 1.25], gap="large")


with main_col:
    if not st.session_state.messages:
        st.markdown(
            """
            <div class="hero">
                <div class="logo-mark">N</div>
                <h1>Hello! I’m NovaCore Copilot</h1>
                <p>Your AI assistant for enterprise data analysis</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        suggestions = [
            "Show total revenue by region",
            "Top 10 vendors by purchase amount",
            "Headcount by department",
            "IT tickets SLA summary",
        ]

        cols = st.columns(4)
        for idx, label in enumerate(suggestions):
            with cols[idx]:
                st.markdown(f'<div class="suggestion"><b>{label}</b></div>', unsafe_allow_html=True)
                if st.button("Ask", key=f"suggest_{idx}", use_container_width=True):
                    st.session_state.pending_question = label
                    st.rerun()

    # Chat history
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            render_text(msg["content"], msg["role"])

            if msg.get("data") is not None:
                df = msg["data"]
                st.dataframe(df, use_container_width=True, hide_index=True)

    question = st.chat_input("Ask a question about your data...")

    if not question and st.session_state.get("pending_question"):
        question = st.session_state.pop("pending_question")

    if question:
        st.session_state.messages.append({"role": "user", "content": question})
        with st.chat_message("user"):
            render_text(question, "user")

        answer, result_df = demo_analysis(question)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer,
            "data": result_df,
        })
        st.session_state.last_result_text = answer
        st.session_state.last_result_df = result_df

        with st.chat_message("assistant"):
            render_text(answer, "assistant")

            if result_df is not None and not result_df.empty:
                st.dataframe(result_df, use_container_width=True, hide_index=True)

        st.rerun()

    # Export and email tools
    if st.session_state.last_result_text:
        st.divider()

        action1, action2, action3 = st.columns(3)

        result_df = st.session_state.last_result_df

        with action1:
            if result_df is not None and not result_df.empty:
                st.download_button(
                    "Download CSV",
                    dataframe_to_csv_bytes(result_df),
                    file_name="novacore_analysis.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

        with action2:
            if result_df is not None and not result_df.empty:
                st.download_button(
                    "Download Excel",
                    dataframe_to_excel_bytes(result_df),
                    file_name="novacore_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        with action3:
            if st.button("Send by Email", use_container_width=True):
                st.session_state.show_email_box = not st.session_state.get("show_email_box", False)

        if st.session_state.get("show_email_box", False):
            with st.form("email_form"):
                recipient = st.text_input("Recipient email")
                subject = st.text_input("Subject", value="NovaCore Copilot Analysis")
                body = st.text_area(
                    "Message",
                    value=st.session_state.last_result_text,
                    height=130,
                )
                send_clicked = st.form_submit_button("Send Email", use_container_width=True)

                if send_clicked:
                    ok, message = send_result_email(
                        recipient=recipient,
                        subject=subject,
                        body=body,
                        result_df=result_df,
                    )
                    if ok:
                        st.success(message)
                    else:
                        st.error(message)

    st.markdown(
        '<div class="security-note">Your data is processed through NovaCore enterprise analytics tools.</div>',
        unsafe_allow_html=True,
    )


with right_col:
    overview = load_dataset_overview(DEFAULT_DATASET)
    overview_data = overview.get("data", {}) if overview.get("success") else {}

    tables = overview_data.get("tables", [])
    table_count = overview_data.get("table_count", len(tables))
    total_rows = sum(int(t.get("row_count", 0)) for t in tables)

    try:
        employees_count = len(load_sheet(DEFAULT_DATASET, "Employees"))
        departments_count = len(load_sheet(DEFAULT_DATASET, "Departments"))
    except Exception:
        employees_count = 0
        departments_count = 0

    st.markdown('<div class="side-panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Data Overview</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-row"><span class="metric-label">Tables</span><span class="metric-value">{table_count:,}</span></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-row"><span class="metric-label">Departments</span><span class="metric-value">{departments_count:,}</span></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-row"><span class="metric-label">Employees</span><span class="metric-value">{employees_count:,}</span></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-row"><span class="metric-label">Records</span><span class="metric-value">{total_rows:,}</span></div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="side-panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Quick Actions</div>', unsafe_allow_html=True)

    if st.button("Refresh Data", use_container_width=True):
        st.cache_data.clear()
        try:
            excel_mcp.clear_cache()
        except Exception:
            pass
        st.rerun()

    if st.button("Data Dictionary", use_container_width=True):
        try:
            dd = load_sheet(DEFAULT_DATASET, "Data_Dictionary")
            st.dataframe(dd, use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(str(exc))

    st.markdown('</div>', unsafe_allow_html=True)
