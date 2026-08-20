from __future__ import annotations

import html
import io
import json
import os
import re
import smtplib
import traceback
from email.message import EmailMessage
from typing import Optional

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


def numeric_columns(df: pd.DataFrame) -> list[str]:
    result = []
    for col in df.columns:
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


def is_time_dimension(col: str) -> bool:
    key = str(col).casefold()
    return any(x in key for x in ["year", "month", "quarter", "date", "time", "period"])


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

    nums = numeric_columns(df)
    dims = dimension_columns(df)
    if not nums:
        return False

    y = nums[-1]
    x = dims[0] if dims else None
    if not x:
        return False

    chart_df = df.copy()
    chart_df[y] = pd.to_numeric(chart_df[y], errors="coerce")
    chart_df = chart_df.dropna(subset=[y])
    if chart_df.empty:
        return False

    question = str(user_question or "")
    arabic = is_arabic(question)
    title = "التحليل المرئي" if arabic else "Visual Analysis"

    if px is None:
        st.bar_chart(chart_df.set_index(x)[y])
        return True

    # Trend/date/year -> line chart
    if is_time_dimension(x) or any(
        token in question.casefold()
        for token in ["trend", "over time", "year", "month", "quarter", "سن", "شهر", "ربع", "اتجاه"]
    ):
        fig = px.line(chart_df, x=x, y=y, markers=True, title=title)
    else:
        # Long categorical list -> horizontal bar
        if len(chart_df) > 8:
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
    :root {
        --navy: #071A31;
        --navy-soft: #0C2544;
        --blue: #176BFF;
        --blue2: #26A6F3;
        --bg: #F7F9FC;
        --white: #FFFFFF;
        --text: #101828;
        --muted: #667085;
        --line: #E2E8F0;
        --green: #12B76A;
    }

    #MainMenu, footer { visibility: hidden; }

    [data-testid="stAppViewContainer"] {
        background: var(--bg);
        color: var(--text);
    }

    [data-testid="stHeader"] {
        background: rgba(255,255,255,.97);
        border-bottom: 1px solid var(--line);
    }

    .block-container {
        max-width: 1540px;
        padding-top: .8rem;
        padding-bottom: 4.5rem;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #071A31 0%, #06152A 100%);
        border-right: 1px solid #102E50;
        min-width: 270px;
    }

    [data-testid="stSidebar"] * { color: #FFFFFF; }

    .side-brand {
        display:flex; align-items:center; gap:11px;
        margin: 4px 0 26px;
    }
    .mini-logo {
        width:38px;height:38px;border-radius:11px;
        background:linear-gradient(135deg,var(--blue),var(--blue2));
        display:flex;align-items:center;justify-content:center;
        font-weight:900;font-size:19px;
        box-shadow:0 8px 20px rgba(23,107,255,.28);
    }
    .brand-title {font-size:17px;font-weight:800;line-height:1.2}
    .brand-title span {color:#55A9FF}
    .brand-sub {font-size:11px;color:#9EB2C8!important;margin-top:3px}

    .nav-item {
        display:flex;align-items:center;gap:10px;
        padding:12px 13px;margin:8px 0;
        border:1px solid #183A59;border-radius:12px;
        font-size:14px;font-weight:650;
    }
    .nav-item.active {
        background:linear-gradient(135deg,#125ECF,#2479FF);
        border-color:#3185FF;
        box-shadow:0 6px 16px rgba(0,0,0,.14);
    }

    .topbar {
        display:flex;justify-content:space-between;align-items:center;
        padding:4px 2px 12px;
    }
    .product-title {font-size:22px;font-weight:800}
    .product-title span {color:var(--blue)}
    .status-pill {
        display:inline-flex;align-items:center;gap:7px;
        padding:7px 12px;border:1px solid var(--line);
        border-radius:999px;background:white;
        font-size:12px;color:#344054;
    }
    .green-dot {width:7px;height:7px;background:var(--green);border-radius:50%}

    .hero {
        text-align:center;
        padding:58px 0 28px;
    }
    .hero-logo {
        width:74px;height:74px;border-radius:20px;margin:auto;
        display:flex;align-items:center;justify-content:center;
        color:white;font-size:31px;font-weight:900;
        background:linear-gradient(135deg,var(--blue),var(--blue2));
        box-shadow:0 15px 34px rgba(23,107,255,.22);
    }
    .hero h1 {font-size:31px;margin:17px 0 5px;font-weight:800}
    .hero p {color:var(--muted);font-size:15px}

    .suggestion-card {
        background:white;border:1px solid var(--line);border-radius:15px;
        padding:17px;min-height:112px;
        box-shadow:0 4px 14px rgba(16,24,40,.035);
    }
    .suggestion-icon {
        width:34px;height:34px;border-radius:9px;
        display:flex;align-items:center;justify-content:center;
        background:#EDF5FF;color:var(--blue);
        margin-bottom:16px;font-size:17px;
    }
    .suggestion-title {font-size:13px;font-weight:750;line-height:1.45}
    .suggestion-sub {font-size:11px;color:var(--muted);margin-top:4px}

    .panel {
        background:white;border:1px solid var(--line);border-radius:15px;
        padding:16px 17px;margin-bottom:13px;
        box-shadow:0 4px 12px rgba(16,24,40,.025);
    }
    .panel-title {font-size:14px;font-weight:800;margin-bottom:12px}
    .overview-row {
        display:flex;justify-content:space-between;align-items:center;
        padding:9px 0;border-bottom:1px solid #F0F2F5;
        font-size:12px;
    }
    .overview-row:last-child {border-bottom:0}
    .overview-label {color:var(--muted)}
    .overview-value {font-size:14px;font-weight:800}

    .activity-row {
        padding:8px 0;border-bottom:1px solid #F2F4F7;
        font-size:11.5px;line-height:1.4;
    }
    .activity-row:last-child {border-bottom:0}

    .message-text {
        line-height:1.9;font-size:14px;
        color:#172033;
    }
    .message-text.rtl {
        direction:rtl;text-align:right;
        font-family:Tahoma, Arial, sans-serif;
    }
    .message-text.ltr {direction:ltr;text-align:left}
    .message-text ul {margin:6px 18px;padding:0}
    .message-text code {
        direction:ltr;background:#F2F4F7;padding:2px 5px;border-radius:5px;
    }

    [data-testid="stChatMessage"] {
        background:white;border:1px solid var(--line);
        border-radius:15px;padding:.55rem .75rem;
        box-shadow:0 4px 12px rgba(16,24,40,.025);
        margin-bottom:.7rem;
    }

    [data-testid="stChatInput"] {
        max-width:1100px;margin:0 auto;
    }
    [data-testid="stChatInput"] > div {
        border:1px solid #D9E1EA;border-radius:14px;
        background:white;
    }
    [data-testid="stChatInput"] textarea {
        unicode-bidi:plaintext;
    }

    .kpi-card {
        background:white;border:1px solid var(--line);border-radius:14px;
        padding:16px 18px;
    }
    .kpi-label {font-size:12px;color:var(--muted);margin-bottom:7px}
    .kpi-value {font-size:24px;font-weight:850;direction:ltr}

    div[data-testid="stDataFrame"] {
        border:1px solid var(--line);
        border-radius:13px;overflow:hidden;
    }

    .result-actions {
        font-size:12px;color:var(--muted);
        margin:5px 0;
    }

    @media(max-width: 1050px) {
        [data-testid="stSidebar"] {min-width:235px}
        .hero {padding-top:30px}
    }
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


# =========================================================
# Sidebar
# =========================================================

with st.sidebar:
    st.markdown(
        """
        <div class="side-brand">
            <div class="mini-logo">N</div>
            <div>
                <div class="brand-title">NovaCore Solutions <span>Copilot</span></div>
                <div class="brand-sub">Enterprise Analytics Platform</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("＋  New Chat", use_container_width=True):
        st.session_state.messages = []
        st.session_state.last_df = None
        st.session_state.last_answer = ""
        st.rerun()

    st.markdown(
        """
        <div class="nav-item active">✦ &nbsp; New Chat</div>
        <div class="nav-item">▦ &nbsp; Dashboard</div>
        <div class="nav-item">◉ &nbsp; Data Explorer</div>
        <div class="nav-item">▥ &nbsp; Reports</div>
        <div class="nav-item">◇ &nbsp; Insights</div>
        <div class="nav-item">♡ &nbsp; Saved Chats</div>
        <div class="nav-item">⚙ &nbsp; Settings</div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<br><br><br>", unsafe_allow_html=True)
    st.caption("NovaCore Solutions")
    st.caption("v1.2.0 · Copilot")


# =========================================================
# Header + layout
# =========================================================

st.markdown(
    """
    <div class="topbar">
        <div class="product-title">NovaCore Solutions <span>Copilot</span></div>
        <div class="status-pill"><span class="green-dot"></span> All Systems Operational</div>
    </div>
    """,
    unsafe_allow_html=True,
)

main_col, right_col = st.columns([4.5, 1.35], gap="large")


# =========================================================
# Main
# =========================================================

with main_col:
    if not st.session_state.messages:
        st.markdown(
            """
            <div class="hero">
                <div class="hero-logo">N</div>
                <h1>Hello! I’m NovaCore Copilot</h1>
                <p>Your AI assistant for enterprise data analysis</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        suggestions = [
            ("↗", "Compare revenue by year", "Trend and performance"),
            ("◎", "Top vendors by spend", "Procurement analysis"),
            ("♙", "Headcount by department", "Workforce overview"),
            ("⚡", "IT SLA performance", "Service performance"),
        ]

        cols = st.columns(4)
        for i, (icon, title, subtitle) in enumerate(suggestions):
            with cols[i]:
                st.markdown(
                    f"""
                    <div class="suggestion-card">
                        <div class="suggestion-icon">{icon}</div>
                        <div class="suggestion-title">{title}</div>
                        <div class="suggestion-sub">{subtitle}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Ask", key=f"suggestion_{i}", use_container_width=True):
                    st.session_state.pending_question = title
                    st.rerun()

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
        data = overview.get("data", {})
        tables = data.get("tables", [])
        table_count = len(tables)
        record_count = sum(int(x.get("row_count", 0)) for x in tables)
        employee_count = len(excel_mcp.load_table(DATASET, "Employees"))
        department_count = len(excel_mcp.load_table(DATASET, "Departments"))
    except Exception:
        table_count = record_count = employee_count = department_count = 0

    st.markdown(
        f"""
        <div class="panel">
            <div class="panel-title">Data Overview</div>
            <div class="overview-row"><span class="overview-label">Tables</span><span class="overview-value">{table_count}</span></div>
            <div class="overview-row"><span class="overview-label">Departments</span><span class="overview-value">{department_count}</span></div>
            <div class="overview-row"><span class="overview-label">Employees</span><span class="overview-value">{employee_count}</span></div>
            <div class="overview-row"><span class="overview-label">Records</span><span class="overview-value">{record_count:,}</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    activity_html = ""
    if st.session_state.recent_queries:
        for q in st.session_state.recent_queries[:5]:
            activity_html += f'<div class="activity-row">{html.escape(q[:70])}</div>'
    else:
        activity_html = '<div class="activity-row" style="color:#98A2B3">No recent activity</div>'

    st.markdown(
        f"""
        <div class="panel">
            <div class="panel-title">Recent Activity</div>
            {activity_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="panel">
            <div class="panel-title">Copilot Status</div>
            <div class="overview-row"><span class="overview-label">Status</span><span class="overview-value" style="color:#12B76A">● Ready</span></div>
            <div class="overview-row"><span class="overview-label">Model</span><span class="overview-value">GitHub Copilot</span></div>
            <div class="overview-row"><span class="overview-label">Calculation</span><span class="overview-value">Verified Excel</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="panel"><div class="panel-title">Quick Actions</div>', unsafe_allow_html=True)

    if st.button("↻  Refresh Data", use_container_width=True):
        st.cache_data.clear()
        try:
            excel_mcp.clear_cache()
        except Exception:
            pass
        st.rerun()

    if st.button("▤  Data Dictionary", use_container_width=True):
        try:
            dictionary = excel_mcp.load_table(DATASET, "Data_Dictionary")
            st.dataframe(dictionary, use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(str(exc))

    st.markdown("</div>", unsafe_allow_html=True)
