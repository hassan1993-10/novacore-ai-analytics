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

import pandas as pd
import streamlit as st

import excel_mcp
from copilot_agent import ask, clear_schema_cache

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


def svg_icon(name: str, size: int = 22, stroke: str = "currentColor") -> str:
    paths = {
        "grid": '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
        "database": '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"/>',
        "chart": '<path d="M4 20V10"/><path d="M10 20V4"/><path d="M16 20v-7"/><path d="M22 20H2"/>',
        "bulb": '<path d="M9 18h6"/><path d="M10 22h4"/><path d="M8.5 14.5A7 7 0 1 1 15.5 14.5c-.9.8-1.5 1.7-1.5 3.5h-4c0-1.8-.6-2.7-1.5-3.5z"/>',
        "bookmark": '<path d="M6 3h12v18l-6-4-6 4z"/>',
        "settings": '<circle cx="12" cy="12" r="3"/><path d="M19 12a7 7 0 1 1-14 0 7 7 0 0 1 14 0z"/>',
        "trend": '<path d="M3 17l6-6 4 4 8-9"/><path d="M15 6h6v6"/>',
        "cart": '<circle cx="9" cy="20" r="1"/><circle cx="18" cy="20" r="1"/><path d="M3 4h2l2.4 10.2a2 2 0 0 0 2 1.6h7.7a2 2 0 0 0 2-1.6L21 8H6"/>',
        "users": '<circle cx="9" cy="8" r="3"/><path d="M3 20a6 6 0 0 1 12 0"/><circle cx="17" cy="9" r="2.5"/><path d="M15 15.5a5 5 0 0 1 6 4.5"/>',
        "coins": '<ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v5c0 1.7 3.1 3 7 3s7-1.3 7-3V6"/><path d="M5 11v5c0 1.7 3.1 3 7 3s7-1.3 7-3v-5"/>',
        "briefcase": '<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M8 7V4h8v3"/><path d="M3 12h18"/>',
        "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v6l4 2"/>',
        "zap": '<path d="M13 2L4 14h7l-1 8 9-12h-7z"/>',
        "search": '<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>',
        "table": '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M9 4v16M15 4v16"/>',
        "lock": '<rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/>',
        "menu": '<path d="M4 6h16M4 12h16M4 18h16"/>',
        "moon": '<path d="M20 15.5A8.5 8.5 0 0 1 8.5 4 8.5 8.5 0 1 0 20 15.5z"/>',
    }
    body = paths.get(name, paths["grid"])
    return f'<svg class="svg-ico" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{stroke}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">{body}</svg>'




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
        height=360,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(size=13),
        title_font=dict(size=16),
        xaxis_title=str(x).replace("_", " "),
        yaxis_title=str(y).replace("_", " "),
        showlegend=False,
        hovermode="x unified" if is_trend else "closest",
    )
    x_key = str(x).casefold()
    if any(k in x_key for k in ["year", "quarter", "month", "period"]):
        fig.update_xaxes(
            type="category",
            tickmode="array",
            tickvals=chart_df[x].tolist(),
            ticktext=[str(v) for v in chart_df[x].tolist()],
            categoryorder="array",
            categoryarray=chart_df[x].tolist(),
        )
    fig.update_yaxes(tickformat=",")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False, "responsive": True})
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
    *{font-family:"Segoe UI",Arial,Tahoma,sans-serif}
    .svg-ico{display:block;flex:0 0 auto}
    .message-text{unicode-bidi:plaintext;overflow-wrap:anywhere;word-break:normal}
    .message-text.rtl{font-family:Tahoma,"Segoe UI",Arial,sans-serif!important;font-size:15px!important;line-height:2.05!important}
    .message-text.ltr{font-family:"Segoe UI",Arial,sans-serif!important}
    .message-text code{direction:ltr;unicode-bidi:isolate;display:inline-block;font-family:Consolas,monospace;background:#F2F4F7;padding:1px 5px;border-radius:5px}
    .sidebar-nav .ico,.sicon,.metric-icon,.right-head .head-icon{display:flex;align-items:center;justify-content:center}
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

    .suggestion-slot{
        background:#fff;border:1px solid var(--line);border-radius:14px;
        padding:18px;min-height:138px;box-shadow:0 4px 12px rgba(16,24,40,.025);
        transition:.16s ease;
    }
    .suggestion-slot:hover{border-color:#B9D1F7;box-shadow:0 8px 22px rgba(16,24,40,.06);transform:translateY(-1px)}
    .suggestion-slot .sicon{margin-bottom:18px}
    .ask-label{
        max-width:1000px;margin:18px auto 8px;font-size:13px;font-weight:800;color:#22354C;
        display:flex;align-items:center;gap:8px;
    }
    .ask-label .ask-dot{width:8px;height:8px;border-radius:50%;background:#176BFF;box-shadow:0 0 0 4px #EAF2FF}
    [data-testid="stChatInput"]{max-width:1000px;margin:0 auto;filter:drop-shadow(0 8px 20px rgba(16,24,40,.055))}
    [data-testid="stChatInput"]>div{
        background:#fff!important;border:2px solid #C9D9EC!important;border-radius:14px!important;min-height:70px!important;
    }
    [data-testid="stChatInput"]>div:focus-within{
        border-color:#176BFF!important;box-shadow:0 0 0 4px rgba(23,107,255,.08)!important;
    }
    [data-testid="stChatInput"] button{background:#176BFF!important;color:#fff!important;border-radius:10px!important}
    .ask-help{max-width:1000px;margin:7px auto 0;color:#7B8BA0;font-size:10.5px}
    .tools-note{font-size:10.5px;color:#7B8BA0;line-height:1.5;margin-bottom:8px}
    .app-header{
        background:rgba(255,255,255,.98);border-bottom:1px solid #E4EAF1;
        margin:-8px -24px 12px;padding:10px 24px 14px!important;
        position:sticky;top:0;z-index:20;backdrop-filter:blur(10px);
    }

    @media(max-width:1200px){
        .company-card{margin-top:45px}
        .header-actions .moon{display:none}
    }
    @media(max-width:900px){
        .block-container{padding-left:.7rem!important;padding-right:.7rem!important}
        .app-header{margin:-8px -12px 10px;padding:9px 12px 12px!important}
        .header-title{font-size:18px!important}
        .status{font-size:10px!important;padding:6px 8px!important}
        .hero{padding:35px 0 26px!important}
        .hero h1{font-size:25px!important}
        .hero-logo{width:78px!important}
        .suggestion-slot{min-height:125px!important;padding:14px!important}
        .sicon{margin-bottom:14px!important}
        .stitle{font-size:12px!important}
        .chips{justify-content:flex-start!important;overflow-x:auto;flex-wrap:nowrap!important;padding-bottom:5px}
        .chip{white-space:nowrap}
    }
    @media(max-width:600px){
        [data-testid="stSidebar"]{min-width:240px}
        .header-left .hamb{display:none}
        .header-title{font-size:16px!important}
        .header-actions{gap:7px!important}
        .status{display:none!important}
        .avatar{width:32px!important;height:32px!important;font-size:10px!important}
        .hero{padding:25px 0 18px!important}
        .hero h1{font-size:22px!important}
        .hero p{font-size:12px!important}
        .hero-logo{width:68px!important}
        [data-testid="stChatInput"]>div{min-height:64px!important}
        .ask-label,.ask-help{padding-left:3px;padding-right:3px}
        [data-testid="stDataFrame"]{font-size:11px}
        div[data-testid="stPlotlyChart"]{padding:2px!important}
    }


    /* Final responsive/mobile polish */
    .suggestion-slot{height:100%;}
    .suggestion-slot + div [data-testid="stBaseButton-secondary"]{min-height:40px;border-radius:10px;font-weight:650}
    [data-testid="stChatMessage"] .message-text{max-width:100%;}
    [data-testid="stChatInput"] textarea{min-height:52px!important;line-height:1.55!important;}
    [data-testid="stChatInput"] textarea::placeholder{color:#6B7C93!important;opacity:1!important;}
    .result-actions{font-weight:750;margin:4px 0 10px;color:#22354C}

    @media(max-width:900px){
      [data-testid="stSidebar"]{min-width:260px!important;}
      .block-container{padding-top:.25rem!important;padding-bottom:6.5rem!important;}
      .app-header{position:relative!important;top:auto!important;}
      .header-title{white-space:nowrap;}
      .suggestion-row{grid-template-columns:repeat(2,minmax(0,1fr))!important;}
      div[data-testid="stHorizontalBlock"]{gap:.65rem!important;}
      [data-testid="stChatInput"]{max-width:100%!important;}
      .ask-label,.ask-help{max-width:100%!important;}
      .message-text.rtl{font-size:14px!important;line-height:1.9!important;}
      div[data-testid="stPlotlyChart"]{width:100%!important;overflow:hidden!important;}
      .right-card{margin-top:4px;}
    }
    @media(max-width:600px){
      .app-header{margin:-4px -6px 8px!important;padding:8px 6px 10px!important;}
      .header-left{gap:8px!important;}
      .header-title{font-size:14px!important;}
      .header-title span{display:inline;}
      .hero{padding:18px 0 15px!important;}
      .hero-logo{width:62px!important;}
      .hero h1{font-size:20px!important;line-height:1.25!important;margin-top:10px!important;}
      .hero p{font-size:11.5px!important;margin-top:6px!important;}
      .suggestion-slot{min-height:108px!important;padding:12px!important;}
      .suggestion-slot .sicon{width:38px!important;height:38px!important;margin-bottom:10px!important;}
      .stitle{font-size:11.5px!important;line-height:1.4!important;}
      .ask-label{font-size:12.5px!important;margin-top:14px!important;}
      [data-testid="stChatInput"]>div{min-height:68px!important;border-width:2px!important;}
      [data-testid="stChatInput"] textarea{font-size:14px!important;}
      .ask-help{font-size:10px!important;line-height:1.45!important;}
      .chips{margin:18px 0 22px!important;}
      .security{margin-top:24px!important;font-size:10px!important;line-height:1.5!important;}
      [data-testid="stChatMessage"]{padding:.45rem .5rem!important;border-radius:12px!important;}
      .message-text.rtl,.message-text.ltr{font-size:13px!important;line-height:1.8!important;}
      .kpi-card{padding:11px 12px!important;}
      .kpi-value{font-size:20px!important;}
      div[data-testid="stDataFrame"]{overflow-x:auto!important;}
      .right-head{padding:11px 13px!important;}
      .rbody{padding:7px 13px!important;}
      .metric{padding:6px 0!important;}
      .metric-icon{width:38px!important;height:38px!important;}
    }

    /* ===== Responsive V2 + clearer Arabic ===== */
    html,body,[data-testid="stAppViewContainer"]{
        max-width:100vw!important;overflow-x:hidden!important;
        -webkit-text-size-adjust:100%;text-size-adjust:100%;
    }
    [data-testid="stAppViewContainer"] *{box-sizing:border-box}
    .message-text,.message-text *{color:#172033!important}
    .message-text.rtl{
        font-family:"Noto Sans Arabic","Segoe UI",Tahoma,Arial,sans-serif!important;
        font-size:15.5px!important;line-height:2!important;letter-spacing:0!important;
        color:#172033!important;overflow:visible!important;
    }
    .message-text.rtl div,.message-text.rtl li,.message-text.rtl strong{
        color:#172033!important;font-family:inherit!important;
    }
    .message-text.ltr{
        font-family:"Segoe UI",Arial,sans-serif!important;color:#172033!important;
        overflow:visible!important;
    }
    .message-text code{unicode-bidi:isolate!important;direction:ltr!important}
    [data-testid="stChatMessage"],[data-testid="stChatMessageContent"]{
        min-width:0!important;max-width:100%!important;overflow:visible!important;height:auto!important;
    }
    [data-testid="stDataFrame"],div[data-testid="stPlotlyChart"]{width:100%!important;max-width:100%!important}
    .ask-label{color:#172033!important}.ask-help{color:#667085!important}

    /* V5: executive text hierarchy — keep the existing clean palette */
    .message-text{color:#1F2937!important;font-weight:400!important}
    .message-text.rtl,.message-text.ltr{color:#1F2937!important}
    .message-text.rtl div,.message-text.rtl li,.message-text.ltr div,.message-text.ltr li{
        color:#1F2937!important;font-weight:400!important;
    }
    .message-text strong{
        color:#0B1F3A!important;font-weight:800!important;
    }
    .message-text ul{margin-top:8px!important;margin-bottom:4px!important}
    .message-text li{margin-bottom:5px!important}
    .message-text code{color:#0B1F3A!important}
    .message-text.rtl{line-height:1.95!important}
    .message-text.ltr{line-height:1.8!important}

    @media(max-width:1024px){
        .block-container{max-width:100%!important;padding-left:1rem!important;padding-right:1rem!important}
        .hero{padding:32px 0 22px!important}
        .suggestion-row{grid-template-columns:repeat(2,minmax(0,1fr))!important}
    }

    @media(max-width:640px){
        .block-container{width:100%!important;max-width:100%!important;padding:.4rem .65rem 5rem!important}
        [data-testid="stSidebar"]{width:min(86vw,300px)!important;min-width:min(86vw,300px)!important}
        .app-header{position:relative!important;margin:-6px -.65rem 8px!important;padding:9px .7rem!important;width:calc(100% + 1.3rem)!important}
        .header-title{font-size:15px!important;white-space:nowrap}.header-actions .status,.header-actions .moon{display:none!important}
        .avatar{width:32px!important;height:32px!important}
        .hero{padding:18px 0 12px!important}.hero-logo{width:62px!important}
        .hero h1{font-size:21px!important;line-height:1.25!important;margin:9px 0 5px!important}.hero p{font-size:12.5px!important}
        .suggestion-row{grid-template-columns:1fr!important;gap:9px!important;margin-bottom:18px!important}
        .suggestion-slot,.suggestion-card{width:100%!important;min-height:104px!important}
        [data-testid="stHorizontalBlock"]{flex-wrap:wrap!important;gap:.55rem!important}
        [data-testid="column"]{min-width:100%!important;width:100%!important;flex:1 1 100%!important}
        .message-text.rtl{font-size:16px!important;line-height:2.05!important}
        .message-text.ltr{font-size:15px!important;line-height:1.75!important}
        [data-testid="stChatInput"]{width:100%!important;max-width:100%!important;margin:0!important}
        [data-testid="stChatInput"]>div{min-height:62px!important;width:100%!important}
        [data-testid="stChatInput"] textarea{font-size:16px!important;line-height:1.45!important}
        .ask-label{font-size:13px!important;margin-top:11px!important}.ask-help{font-size:11px!important;line-height:1.55!important}
        [data-testid="stDownloadButton"] button,.stButton button{min-height:44px!important}
        div[data-testid="stPlotlyChart"]{margin:0!important;padding:2px!important}
        .right-card{margin:8px 0!important}.company-card{margin-top:25px!important}
        .security{font-size:10.5px!important;line-height:1.5!important;padding:0 8px}
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# Cached UI metadata
# =========================================================
@st.cache_data(ttl=1800, show_spinner=False)
def cached_overview_metrics():
    overview = json.loads(excel_mcp.get_model_overview(DATASET))
    od = overview.get("data", {}) if overview.get("success") else {}
    tables = od.get("tables", [])
    record_count = sum(int(x.get("row_count", 0) or 0) for x in tables)
    try: employee_count = len(excel_mcp.load_table(DATASET, "Employees"))
    except Exception: employee_count = 0
    try: department_count = len(excel_mcp.load_table(DATASET, "Departments"))
    except Exception: department_count = 0
    return record_count, employee_count, department_count

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
    st.markdown(f'<div class="brand-wrap"><img class="brand-logo" src="{LOGO_URI}"><div class="brand-name">NovaCore Solutions<br><span>Copilot</span></div></div>', unsafe_allow_html=True)
    if st.button("New Chat                                      +", use_container_width=True):
        st.session_state.messages=[]; st.session_state.last_df=None; st.session_state.last_answer=""; st.rerun()
    st.markdown(
        f'<div class="sidebar-nav"><span class="ico">{svg_icon("grid")}</span>Dashboard</div>'
        f'<div class="sidebar-nav"><span class="ico">{svg_icon("database")}</span>Data Explorer</div>'
        f'<div class="sidebar-nav"><span class="ico">{svg_icon("chart")}</span>Reports</div>'
        f'<div class="sidebar-nav"><span class="ico">{svg_icon("bulb")}</span>Insights</div>'
        f'<div class="sidebar-nav"><span class="ico">{svg_icon("bookmark")}</span>Saved Chats</div>'
        f'<div class="sidebar-nav"><span class="ico">{svg_icon("settings")}</span>Settings</div>',
        unsafe_allow_html=True,
    )
    st.markdown(f'<div class="company-card"><div class="company-card-title">NovaCore Solutions</div><div class="company-card-sub">Enterprise Analytics Platform</div><img class="company-building" src="{BUILDING_URI}"><div class="version-pill">Version 1</div></div><div class="copyright">© 2026 NovaCore Solutions</div>', unsafe_allow_html=True)


st.markdown(
    f'<div class="app-header"><div class="header-left"><span class="hamb">{svg_icon("menu",24)}</span>'
    f'<div class="header-title">NovaCore Solutions <span>Copilot</span></div></div>'
    f'<div class="header-actions"><div class="status"><span class="dot"></span>All Systems Operational</div>'
    f'<span class="moon">{svg_icon("moon",22)}</span><div class="avatar">NC</div></div></div>',
    unsafe_allow_html=True,
)

# Responsive layout: keep the main conversation dominant.
# Streamlit columns stack automatically on narrow/mobile screens.
main_col, right_col = st.columns([5.15, 1.15], gap="medium")

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

        cards=[
            ("trend","s1","Show total revenue by department"),
            ("cart","s2","Top 10 vendors by purchase amount"),
            ("users","s3","Headcount by department"),
            ("coins","s4","Monthly expense trend for this year"),
        ]
        prompt_cols = st.columns(4)
        for idx, (icon, cls, q0) in enumerate(cards):
            with prompt_cols[idx]:
                st.markdown(
                    f'<div class="suggestion-slot"><div class="sicon {cls}">{svg_icon(icon,23)}</div>'
                    f'<div class="stitle">{html.escape(q0)}</div></div>',
                    unsafe_allow_html=True,
                )
                if st.button("Ask this", key=f"suggested_prompt_{idx}", use_container_width=True):
                    st.session_state.pending_question = q0
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

    st.markdown(
        '<div class="ask-label"><span class="ask-dot"></span>Ask NovaCore Copilot · اكتب سؤالك هنا</div>',
        unsafe_allow_html=True,
    )
    question = st.chat_input("Ask about your data / اكتب سؤالك عن البيانات هنا...")
    st.markdown(
        '<div class="ask-help">Example: Compare net revenue by year · مثال: قارن صافي الإيرادات حسب السنة</div>',
        unsafe_allow_html=True,
    )
    if not st.session_state.messages:
        st.markdown('<div class="chips"><span class="chip">Sales this month</span><span class="chip">IT tickets by priority</span><span class="chip">Purchase by category</span><span class="chip">Employee turnover rate</span><span class="chip">Cash flow overview</span></div><div class="security">🔒 &nbsp; Your data is secure and private. All analysis is performed using NovaCore enterprise data.</div>', unsafe_allow_html=True)

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
        record_count, employee_count, department_count = cached_overview_metrics()
    except Exception:
        record_count=employee_count=department_count=0
    st.markdown(f'<div class="right-card"><div class="right-head"><span class="head-icon" style="color:#176BFF">{svg_icon("database",20)}</span>Data Overview</div><div class="rbody"><div class="metric"><div class="metric-icon">{svg_icon("briefcase",20)}</div><div><div class="metric-num">1</div><div class="metric-lbl">Datasets</div></div></div><div class="metric"><div class="metric-icon">{svg_icon("grid",20)}</div><div><div class="metric-num">{department_count}</div><div class="metric-lbl">Departments</div></div></div><div class="metric"><div class="metric-icon">{svg_icon("users",20)}</div><div><div class="metric-num">{employee_count:,}</div><div class="metric-lbl">Employees</div></div></div><div class="metric"><div class="metric-icon">{svg_icon("table",20)}</div><div><div class="metric-num">{record_count:,}</div><div class="metric-lbl">Records</div></div></div></div></div>', unsafe_allow_html=True)
    recent = st.session_state.get("recent_queries", [])
    defaults=[("Revenue by Department","2 min ago"),("Top Vendors Analysis","15 min ago"),("Headcount Overview","1 hour ago"),("IT Tickets Summary","2 hours ago"),("Monthly Expenses","3 hours ago")]
    if recent:
        activity=''.join(f'<div class="activity"><span>{html.escape(q[:34])}</span><span class="activity-time">just now</span></div>' for q in recent[:5])
    else:
        activity=''.join(f'<div class="activity"><span>{n}</span><span class="activity-time">{t}</span></div>' for n,t in defaults)
    st.markdown(f'<div class="right-card"><div class="right-head"><span class="head-icon" style="color:#176BFF">{svg_icon("clock",20)}</span>Recent Activity</div><div class="rbody">{activity}</div></div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="right-card"><div class="right-head"><span class="head-icon" style="color:#176BFF">{svg_icon("zap",20)}</span>Quick Tools</div>'
        f'<div class="rbody"><div class="tools-note">Utilities for the current dataset.</div>',
        unsafe_allow_html=True,
    )
    if st.button("Refresh Data", key="quick_refresh", use_container_width=True):
        st.cache_data.clear()
        clear_schema_cache()
        try:
            excel_mcp.clear_cache()
        except Exception:
            pass
        st.rerun()
    if st.button("Data Dictionary", use_container_width=True):
        try:
            dd = excel_mcp.load_table(DATASET,"Data_Dictionary")
            st.dataframe(dd,use_container_width=True,hide_index=True)
        except Exception as exc: st.error(str(exc))
    st.markdown('</div></div>', unsafe_allow_html=True)
