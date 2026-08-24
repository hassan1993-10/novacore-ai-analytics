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
    initial_sidebar_state="auto",
)

if "GITHUB_TOKEN" in st.secrets:
    os.environ["GITHUB_TOKEN"] = str(st.secrets["GITHUB_TOKEN"])
if "COPILOT_GITHUB_TOKEN" in st.secrets:
    os.environ["COPILOT_GITHUB_TOKEN"] = str(st.secrets["COPILOT_GITHUB_TOKEN"])

DATASET = "novacore_enterprise_sample_data"
APP_DIR = Path(__file__).resolve().parent
ASSETS_DIR = APP_DIR / "assets"


@st.cache_data(show_spinner=False)
def _cached_data_uri(path_value: str, modified_ns: int) -> str:
    path = Path(path_value)
    if not path.is_file():
        return ""
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")


def _data_uri(path: Path) -> str:
    modified_ns = path.stat().st_mtime_ns if path.is_file() else 0
    return _cached_data_uri(str(path), modified_ns)


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
ANALYSIS_HEADINGS = {
    "executive summary", "summary", "key insights", "insights", "drivers",
    "risks", "risks / anomalies", "risks and anomalies", "anomalies",
    "recommendations", "visual analysis", "supporting data",
    "الملخص التنفيذي", "الملخص", "أهم الرؤى", "الرؤى الرئيسية", "الرؤى",
    "المحركات", "العوامل المؤثرة", "المخاطر", "المخاطر والشذوذ",
    "الانحرافات", "التوصيات", "التحليل المرئي", "البيانات الداعمة",
}


def is_arabic(text: str) -> bool:
    text = str(text or "")
    return len(ARABIC_RE.findall(text)) > len(re.findall(r"[A-Za-z]", text))


def direction(text: str) -> str:
    return "rtl" if is_arabic(text) else "ltr"


def formatted_html(text: str) -> str:
    """Safely format only the analytical structure present in the response."""
    value = html.escape(str(text or ""))
    value = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", value)
    value = re.sub(r"`(.+?)`", r"<code>\1</code>", value)
    lines = value.splitlines()

    output = []
    in_list = False
    for line in lines:
        stripped = line.strip()
        plain = re.sub(r"</?strong>", "", stripped).rstrip(":：").strip().casefold()
        if stripped.startswith(("- ", "• ")):
            if not in_list:
                output.append("<ul>")
                in_list = True
            output.append(f"<li>{stripped[2:]}</li>")
        else:
            if in_list:
                output.append("</ul>")
                in_list = False
            if plain in ANALYSIS_HEADINGS:
                output.append(f'<div class="analysis-heading">{stripped.rstrip(":：")}</div>')
            elif stripped:
                output.append(f"<div>{stripped}</div>")
            else:
                output.append('<div class="message-gap"></div>')
    if in_list:
        output.append("</ul>")
    return "".join(output)


def render_message_text(text: str, role: str = "assistant") -> None:
    d = direction(text)
    role_name = "You · أنت" if role == "user" else "NovaCore Copilot"
    st.markdown(
        f'<div class="message-text {role} {d}" dir="{d}">'
        f'<div class="message-role">{role_name}</div>{formatted_html(text)}</div>',
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


def _human_label(name: str) -> str:
    return str(name).replace("_", " ").strip()

def _short_number(v):
    try:
        n = float(v)
    except Exception:
        return str(v)
    if abs(n) >= 1_000_000_000:
        return f"{n/1_000_000_000:.1f}B"
    if abs(n) >= 1_000_000:
        return f"{n/1_000_000:.1f}M"
    if abs(n) >= 1_000:
        return f"{n/1_000:.1f}K"
    return f"{n:,.0f}" if n.is_integer() else f"{n:,.2f}"


def _nova_chart_layout(
    fig,
    *,
    height: int = 360,
    showlegend: bool = False,
    hovermode: str = "closest",
):
    """Apply the single NovaCore chart theme."""
    fig.update_layout(
        height=height,
        autosize=True,
        margin=dict(l=18, r=22, t=18, b=42),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(size=12, color="#344054", family="Segoe UI, Tahoma, Arial"),
        showlegend=showlegend,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02,
            xanchor="left", x=0, title_text="",
        ),
        hovermode=hovermode,
        bargap=0.28,
        separators=".,",
    )
    fig.update_xaxes(
        showgrid=False, zeroline=False, linecolor="#DCE5EF",
        tickfont=dict(color="#667085"), automargin=True,
    )
    fig.update_yaxes(
        showgrid=True, gridcolor="#E9EFF6", zeroline=False,
        tickfont=dict(color="#667085"), automargin=True,
    )
    return fig

def render_chart(df: pd.DataFrame, user_question: str = "") -> bool:
    """Smart chart renderer aligned to the verified result grain."""
    if df is None or df.empty:
        return False

    chart_df = df.copy()
    columns = list(chart_df.columns)
    time_dims = [c for c in columns if is_time_dimension(c)]
    nums = numeric_columns(chart_df)
    dims = [c for c in columns if c not in nums and c not in time_dims]
    if not nums:
        return False

    q = str(user_question or "")
    arabic = is_arabic(q)
    q_key = q.casefold()
    actual_cols = [c for c in nums if any(k in str(c).casefold() for k in ["actual", "فعلي", "الفعلية"])]
    target_cols = [c for c in nums if any(k in str(c).casefold() for k in ["target", "budget", "هدف", "المستهدف", "الميزانية"])]
    is_variance = bool(actual_cols and target_cols)

    if len(chart_df) < 2 and not is_variance:
        return False

    for col in nums:
        chart_df[col] = pd.to_numeric(chart_df[col], errors="coerce")

    if px is None:
        if time_dims:
            st.line_chart(chart_df.set_index(time_dims[0])[nums[:4]])
        elif dims:
            st.bar_chart(chart_df.set_index(dims[0])[nums[-1]])
        return True

    palette = ["#176BFF", "#35B9F2", "#158A55", "#667085"]
    heading_dir = "rtl" if arabic else "ltr"

    if is_variance:
        x = dims[0] if dims else "_Result"
        if not dims:
            chart_df[x] = "النتيجة" if arabic else "Result"
        value_cols = [actual_cols[0], target_cols[0]]
        melted = chart_df.melt(
            id_vars=[x], value_vars=value_cols,
            var_name="Series", value_name="Value",
        ).dropna(subset=["Value"])
        title = "مقارنة الفعلي بالمستهدف" if arabic else "Actual vs target"
        fig = px.bar(
            melted, x="Value", y=x, color="Series", orientation="h",
            barmode="group", text_auto=".3s",
            color_discrete_sequence=["#176BFF", "#667085"],
        )
        fig.update_xaxes(title_text="القيمة" if arabic else "Value", tickformat="~s")
        fig.update_yaxes(title_text="")
        showlegend = True
        hovermode = "closest"
    elif time_dims:
        x = time_dims[0]
        chart_df = chart_df.sort_values(x)
        series = nums[:4]
        title = f"اتجاه {_human_label(series[-1])}" if arabic else f"{_human_label(series[-1])} trend"
        fig = px.line(
            chart_df, x=x, y=series, markers=len(chart_df) <= 18,
            color_discrete_sequence=palette,
        )
        fig.update_traces(line_width=2.5, marker_size=7)
        if any(k in str(x).casefold() for k in ["year", "quarter", "month", "period"]):
            vals = chart_df[x].drop_duplicates().tolist()
            fig.update_xaxes(
                type="category", tickmode="array", tickvals=vals,
                ticktext=[str(int(v)) if isinstance(v, (int, float)) and float(v).is_integer() else str(v) for v in vals],
                categoryorder="array", categoryarray=vals,
            )
        fig.update_xaxes(title_text=_human_label(x))
        fig.update_yaxes(title_text="", tickformat="~s")
        showlegend = len(series) > 1
        hovermode = "x unified"
    else:
        dimension = dims[0] if dims else columns[0]
        measure = nums[-1]
        chart_df = chart_df.dropna(subset=[measure])
        composition_terms = ["share", "composition", "distribution", "percentage", "حصة", "توزيع", "نسبة"]
        is_composition = any(term in q_key for term in composition_terms) and 2 <= chart_df[dimension].nunique() <= 5
        title = f"{_human_label(measure)} حسب {_human_label(dimension)}" if arabic else f"{_human_label(measure)} by {_human_label(dimension)}"

        if is_composition:
            fig = px.pie(
                chart_df, names=dimension, values=measure, hole=0.52,
                color_discrete_sequence=palette,
            )
            fig.update_traces(textposition="inside", textinfo="percent+label")
            showlegend = True
        else:
            chart_df = chart_df.sort_values(measure, ascending=True)
            fig = px.bar(
                chart_df, x=measure, y=dimension, orientation="h",
                color_discrete_sequence=["#176BFF"],
                text=chart_df[measure].map(_short_number),
            )
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_xaxes(title_text=_human_label(measure), tickformat="~s")
            fig.update_yaxes(title_text="")
            showlegend = False
        hovermode = "closest"

    st.markdown(
        f'<div class="visual-section-title" dir="{heading_dir}">'
        f'{html.escape(title)}</div>',
        unsafe_allow_html=True,
    )
    _nova_chart_layout(fig, showlegend=showlegend, hovermode=hovermode)

    st.plotly_chart(
        fig, width="stretch",
        theme=None,
        config={"displayModeBar": False, "responsive": True}
    )
    return True


def render_supporting_data(df: pd.DataFrame) -> None:
    """Brand-consistent supporting-data preview instead of a dark dataframe block."""
    if df is None or df.empty:
        return

    preview = df.head(20).copy()
    headers = "".join(
        f'<th dir="{direction(c)}">{html.escape(_human_label(c))}</th>'
        for c in preview.columns
    )
    rows = []
    for _, row in preview.iterrows():
        cells = []
        for c in preview.columns:
            val = row[c]
            if pd.isna(val):
                display = "—"
                cell_class = ""
            elif isinstance(val, (int, float)):
                cell_class = "numeric"
                if isinstance(val, float) and val.is_integer():
                    display = f"{int(val):,}"
                else:
                    display = f"{val:,.2f}" if isinstance(val, float) else f"{val:,}"
            else:
                cell_class = ""
                display = str(val)
            cells.append(
                f'<td class="{cell_class}" dir="{direction(display)}">'
                f'{html.escape(display)}</td>'
            )
        rows.append("<tr>" + "".join(cells) + "</tr>")

    st.markdown(
        f"""
        <div class="supporting-data">
          <div class="supporting-title">Supporting Data · البيانات الداعمة</div>
          <div class="table-scroll">
            <table class="nova-table">
              <thead><tr>{headers}</tr></thead>
              <tbody>{''.join(rows)}</tbody>
            </table>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_driver_charts(driver_evidence: dict | None, user_question: str = "") -> bool:
    """Render verified contribution-to-change charts for root-cause/driver questions."""
    if not driver_evidence or not driver_evidence.get("drivers"):
        return False

    arabic = is_arabic(user_question)
    rendered = False

    family_labels_ar = {
        "region": "المناطق الأكثر تأثيرًا",
        "product": "المنتجات الأكثر تأثيرًا",
        "department": "الأقسام الأكثر تأثيرًا",
        "category": "الفئات الأكثر تأثيرًا",
        "vendor": "الموردون الأكثر تأثيرًا",
    }
    family_labels_en = {
        "region": "Regions contributing to change",
        "product": "Products contributing to change",
        "department": "Departments contributing to change",
        "category": "Categories contributing to change",
        "vendor": "Vendors contributing to change",
    }

    for family, payload in list(driver_evidence.get("drivers", {}).items())[:2]:
        rows = payload.get("rows") or []
        if not rows:
            continue

        d = pd.DataFrame(rows)
        if d.empty or "name" not in d.columns or "delta" not in d.columns:
            continue

        d["magnitude"] = d["delta"].abs()
        d = d.nlargest(8, "magnitude").sort_values("delta", ascending=True)

        if d.empty:
            continue

        title = family_labels_ar.get(family, "العوامل الأكثر تأثيرًا") if arabic else family_labels_en.get(family, "Main contributors")
        value_label = "المساهمة في التغير" if arabic else "Contribution to change"

        if px is not None:
            fig = px.bar(
                d,
                x="delta",
                y="name",
                orientation="h",
                text=d["delta"].map(_short_number),
                color=d["delta"].ge(0).map({True: "Positive", False: "Negative"}),
                color_discrete_map={"Positive": "#158A55", "Negative": "#C4323C"},
            )
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_xaxes(
                title_text=value_label, tickformat="~s", zeroline=True,
                zerolinecolor="#10243E", zerolinewidth=1,
            )
            fig.update_yaxes(title_text="")
            st.markdown(
                f'<div class="visual-section-title" dir="{"rtl" if arabic else "ltr"}">'
                f'{html.escape(title)}</div>',
                unsafe_allow_html=True,
            )
            _nova_chart_layout(fig, height=max(320, min(460, 48 * len(d))), showlegend=False)
            st.plotly_chart(
                fig,
                width="stretch",
                theme=None,
                config={"displayModeBar": False, "responsive": True},
            )
        else:
            st.bar_chart(d.set_index("name")["impact"])

        rendered = True

    return rendered


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
# NovaCore design system
# =========================================================

st.markdown(
    """
    <style>
    :root {
        --nc-navy: #071A2F;
        --nc-blue: #176BFF;
        --nc-cyan: #35B9F2;
        --nc-strong: #10243E;
        --nc-text: #344054;
        --nc-muted: #667085;
        --nc-bg: #F5F8FC;
        --nc-surface: #FFFFFF;
        --nc-border: #DCE5EF;
        --nc-soft-blue: #EDF5FF;
        --nc-success: #158A55;
        --nc-warning: #B66A00;
        --nc-critical: #C4323C;
        --nc-shadow: 0 8px 24px rgba(16, 36, 62, .055);
        --nc-radius: 12px;
    }

    html, body, [data-testid="stAppViewContainer"] {
        max-width: 100%;
        overflow-x: hidden;
        background: var(--nc-bg);
        color: var(--nc-text);
        font-family: "Segoe UI Variable", "Segoe UI", Tahoma, Arial, sans-serif;
        -webkit-text-size-adjust: 100%;
        text-size-adjust: 100%;
    }
    [data-testid="stAppViewContainer"] *,
    [data-testid="stSidebar"] * { box-sizing: border-box; }
    button, input, textarea, select {
        font-family: "Segoe UI Variable", "Segoe UI", Tahoma, Arial, sans-serif !important;
        font-size: 14px !important;
    }
    #MainMenu, footer, [data-testid="stDecoration"] { display: none; }
    [data-testid="stHeader"] {
        height: 48px;
        background: rgba(255, 255, 255, .96);
        border-bottom: 1px solid var(--nc-border);
        backdrop-filter: blur(10px);
    }
    [data-testid="stHeader"] button {
        min-width: 44px;
        min-height: 44px;
        color: var(--nc-navy);
    }
    [data-testid="stMain"] { width: 100%; min-width: 0; }
    [data-testid="stMainBlockContainer"], .block-container {
        width: 100%;
        max-width: none;
        margin: 0;
        padding: .65rem 1.25rem 5rem;
    }
    .svg-ico { display: block; flex: 0 0 auto; }
    .conversation-anchor, .context-anchor, .suggestion-anchor, .export-anchor {
        display: none;
    }
    [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) { display: none; }
    .mobile-menu-brand { margin-bottom: 12px; color: var(--nc-navy); font-size: 15px; font-weight: 750; }
    .mobile-menu-brand span { color: var(--nc-blue); }
    .mobile-nav > div {
        display: flex; align-items: center; gap: 10px; min-height: 44px;
        padding: 9px 7px; border-bottom: 1px solid #EDF1F5; color: var(--nc-text); font-size: 14px;
    }
    .mobile-nav > div:last-child { border-bottom: 0; }

    /* Sidebar and native drawer */
    [data-testid="stSidebar"] {
        width: 260px !important;
        min-width: 260px !important;
        background: var(--nc-navy);
        border-right: 1px solid #17324F;
    }
    [data-testid="stSidebar"] > div { width: 260px !important; }
    [data-testid="stSidebar"] * { color: #F3F7FC; }
    [data-testid="stSidebar"] .block-container { padding: 1rem .9rem 1.25rem; }
    [data-testid="stSidebar"] .stButton > button {
        min-height: 48px;
        border: 1px solid #397FD9;
        border-radius: 8px;
        background: var(--nc-blue);
        color: #FFFFFF;
        font-weight: 700;
        box-shadow: none;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background: #0E5DDF;
        border-color: #75A9FF;
    }
    .brand-wrap { display: flex; align-items: center; gap: 12px; margin: 2px 2px 22px; }
    .brand-logo { width: 44px; height: 40px; object-fit: contain; }
    .brand-name { font-size: 16px; font-weight: 750; line-height: 1.35; }
    .brand-name span { color: var(--nc-cyan); }
    .sidebar-nav {
        display: flex; align-items: center; gap: 12px;
        min-height: 44px; padding: 10px 9px; margin: 3px 0;
        border-radius: 8px; color: #DCE8F5; font-size: 14px;
    }
    .sidebar-nav:first-child { background: rgba(53, 185, 242, .12); color: #FFFFFF; }
    .sidebar-nav .ico { display: flex; align-items: center; justify-content: center; width: 24px; }
    .company-card {
        position: relative; min-height: 150px; margin-top: 44px; padding: 16px;
        overflow: hidden; border: 1px solid rgba(220, 229, 239, .18);
        border-radius: 10px; background: rgba(255, 255, 255, .045);
    }
    .company-card-title { font-size: 14px; font-weight: 700; }
    .company-card-sub, .copyright { color: #B8C7D8 !important; font-size: 12px; }
    .company-card-sub { margin-top: 5px; }
    .company-building { position: absolute; right: 0; bottom: -20px; height: 112px; opacity: .9; }
    .version-pill {
        position: absolute; left: 16px; bottom: 13px; padding: 4px 8px;
        border: 1px solid rgba(255, 255, 255, .28); border-radius: 999px; font-size: 12px;
    }
    .copyright { margin: 12px 3px 0; }

    /* Application header */
    .app-header {
        display: flex; align-items: center; justify-content: space-between;
        min-height: 58px; margin-bottom: 18px; padding: 5px 4px 12px;
        border-bottom: 1px solid var(--nc-border);
    }
    .header-left, .header-actions { display: flex; align-items: center; }
    .header-left { gap: 10px; }
    .header-actions { gap: 12px; }
    .header-logo { width: 34px; height: 32px; object-fit: contain; }
    .header-title { color: var(--nc-navy); font-size: 20px; font-weight: 750; }
    .header-title span { color: var(--nc-blue); }
    .status {
        display: flex; align-items: center; gap: 7px; min-height: 36px;
        padding: 7px 12px; border: 1px solid var(--nc-border);
        border-radius: 999px; background: var(--nc-surface); color: var(--nc-text); font-size: 12px;
    }
    .dot { width: 8px; height: 8px; border-radius: 50%; background: var(--nc-success); }
    .avatar {
        display: flex; align-items: center; justify-content: center;
        width: 38px; height: 38px; border-radius: 50%;
        background: var(--nc-navy); color: #FFFFFF; font-size: 12px; font-weight: 750;
    }

    /* Desktop workspace: flexible conversation plus fixed contextual rail */
    [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) {
        align-items: flex-start;
        gap: 24px;
    }
    [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor)
    > [data-testid="stColumn"]:has(.conversation-anchor) {
        flex: 1 1 0 !important; width: 0 !important; min-width: 0 !important;
    }
    [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor)
    > [data-testid="stColumn"]:has(.context-anchor) {
        flex: 0 0 280px !important; width: 280px !important; min-width: 280px !important;
    }
    [data-testid="stColumn"]:has(.conversation-anchor) > div { max-width: 1120px; margin-inline: auto; }

    /* Welcome and suggestions */
    .hero { text-align: center; padding: 30px 0 24px; }
    .hero-logo { width: 76px; height: auto; }
    .hero h1 { margin: 12px 0 5px; color: var(--nc-navy); font-size: 28px; line-height: 1.25; font-weight: 750; }
    .hero p { margin: 0; color: var(--nc-muted); font-size: 15px; }
    [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) { gap: 12px; }
    [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) > [data-testid="stColumn"] { min-width: 0; }
    [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) .stButton > button {
        position: relative; align-items: flex-end; justify-content: flex-start;
        min-height: 122px; padding: 18px; overflow: hidden;
        border: 1px solid var(--nc-border); border-radius: 10px;
        background: var(--nc-surface); color: var(--nc-strong);
        text-align: left; white-space: normal; font-weight: 700;
        box-shadow: var(--nc-shadow); transition: border-color .15s ease, transform .15s ease, box-shadow .15s ease;
    }
    [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) .stButton > button::before {
        content: ""; position: absolute; left: 18px; top: 18px;
        width: 34px; height: 6px; border-radius: 999px; background: var(--nc-blue);
        box-shadow: 15px 0 0 var(--nc-cyan);
    }
    [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) .stButton > button:hover {
        transform: translateY(-1px); border-color: #9EC0F8; box-shadow: 0 12px 26px rgba(23, 107, 255, .09);
    }
    [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) .stButton > button:focus-visible,
    button:focus-visible, textarea:focus-visible, input:focus-visible {
        outline: 3px solid rgba(23, 107, 255, .25) !important;
        outline-offset: 2px;
    }

    /* Copilot conversation */
    [data-testid="stChatMessage"] {
        width: 100%; min-width: 0; margin-bottom: 12px; padding: 16px 18px;
        border: 1px solid var(--nc-border); border-radius: var(--nc-radius);
        background: var(--nc-surface); box-shadow: 0 5px 18px rgba(16, 36, 62, .04);
    }
    [data-testid="stChatMessage"]:has(.message-text.user) {
        width: min(78%, 760px); margin-left: auto;
        border-color: #BDD3F7; background: var(--nc-soft-blue); box-shadow: none;
    }
    [data-testid="stChatMessage"]:has(.message-text.user.rtl) { margin-right: 0; }
    .message-text {
        color: var(--nc-text); unicode-bidi: plaintext;
        overflow-wrap: anywhere; word-break: normal; font-weight: 400;
    }
    .message-text.ltr {
        direction: ltr; text-align: left; font-size: 15px; line-height: 1.75;
        font-family: "Segoe UI Variable", "Segoe UI", Tahoma, Arial, sans-serif;
    }
    .message-text.rtl {
        direction: rtl; text-align: right; font-size: 16px; line-height: 1.85;
        font-family: Tahoma, "Segoe UI", Arial, sans-serif;
    }
    .message-role {
        margin-bottom: 8px; color: var(--nc-blue); font-size: 12px;
        line-height: 1.4; font-weight: 750;
    }
    .message-text.user .message-role { color: var(--nc-navy); }
    .message-text strong { color: var(--nc-strong); font-weight: 750; }
    .message-text ul { margin: 8px 1.2rem 4px; padding: 0; }
    .message-text li { margin-bottom: 6px; }
    .message-gap { height: 8px; }
    .analysis-heading, .visual-section-title {
        margin: 20px 0 8px; padding-top: 14px;
        border-top: 1px solid var(--nc-border); color: var(--nc-strong);
        font-size: 18px; line-height: 1.45; font-weight: 750;
    }
    .message-role + .analysis-heading { margin-top: 4px; padding-top: 0; border-top: 0; }
    .message-text code {
        direction: ltr; unicode-bidi: isolate; display: inline-block;
        padding: 1px 5px; border-radius: 5px; background: #EEF2F6;
        color: var(--nc-strong); font-family: Consolas, monospace;
    }

    /* Input, chips, and result actions */
    .ask-label, .ask-help, [data-testid="stChatInput"] { width: 100%; max-width: 1040px; margin-inline: auto; }
    .ask-label {
        display: flex; align-items: center; gap: 9px; margin-top: 20px; margin-bottom: 8px;
        color: var(--nc-strong); font-size: 14px; font-weight: 750;
    }
    .ask-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--nc-blue); box-shadow: 0 0 0 4px var(--nc-soft-blue); }
    [data-testid="stChatInput"] > div {
        min-height: 68px; border: 1.5px solid #B8CDF0; border-radius: 12px;
        background: var(--nc-surface); box-shadow: var(--nc-shadow);
    }
    [data-testid="stChatInput"] > div:focus-within { border-color: var(--nc-blue); box-shadow: 0 0 0 4px rgba(23, 107, 255, .09); }
    [data-testid="stChatInput"] textarea { min-height: 52px; color: var(--nc-text); font-size: 15px !important; line-height: 1.5; unicode-bidi: plaintext; }
    [data-testid="stChatInput"] button { min-width: 44px; min-height: 44px; border-radius: 8px; background: var(--nc-blue); color: #FFFFFF; }
    .ask-help, .tools-note { margin-top: 7px; color: var(--nc-muted); font-size: 12px; line-height: 1.5; }
    .chips { display: flex; justify-content: center; gap: 8px; flex-wrap: wrap; margin: 22px 0 26px; }
    .chip { padding: 7px 11px; border: 1px solid var(--nc-border); border-radius: 999px; background: var(--nc-surface); color: var(--nc-muted); font-size: 12px; }
    .security { margin-top: 28px; text-align: center; color: var(--nc-muted); font-size: 12px; }
    .result-actions { margin: 4px 0 10px; color: var(--nc-strong); font-size: 14px; font-weight: 750; }
    [data-testid="stDownloadButton"] button, [data-testid="stForm"] button { min-height: 44px; border-radius: 8px; }

    /* KPIs, charts, and supporting data */
    [data-testid="stHorizontalBlock"]:has(.kpi-card):not(:has(.conversation-anchor)) { gap: 10px; margin: 10px 0 14px; }
    .kpi-card { min-height: 104px; padding: 15px 16px; border: 1px solid var(--nc-border); border-radius: 10px; background: var(--nc-surface); }
    .kpi-label { margin-bottom: 7px; color: var(--nc-muted); font-size: 12px; }
    .kpi-value { direction: ltr; color: var(--nc-strong); font-size: 24px; font-weight: 750; font-variant-numeric: tabular-nums; }
    div[data-testid="stPlotlyChart"] { width: 100%; max-width: 100%; min-width: 0; margin-bottom: 12px; overflow: hidden; }
    .supporting-data { margin: 3px 0 8px; overflow: hidden; border: 1px solid var(--nc-border); border-radius: 10px; background: var(--nc-surface); }
    .supporting-title { padding: 12px 14px; border-bottom: 1px solid var(--nc-border); background: #F9FBFD; color: var(--nc-strong); font-size: 12px; font-weight: 750; }
    .table-scroll { width: 100%; overflow-x: auto; overscroll-behavior-inline: contain; }
    .nova-table { width: 100%; min-width: max-content; border-collapse: collapse; color: var(--nc-text); font-size: 12px; }
    .nova-table th { position: sticky; top: 0; padding: 10px 12px; border-bottom: 1px solid var(--nc-border); background: #F3F7FB; color: var(--nc-muted); text-align: start; white-space: nowrap; font-weight: 700; }
    .nova-table td { padding: 10px 12px; border-bottom: 1px solid #EDF1F5; text-align: start; white-space: nowrap; }
    .nova-table td.numeric { direction: ltr; text-align: end; font-variant-numeric: tabular-nums; }
    .nova-table tr:last-child td { border-bottom: 0; }
    [data-testid="stExpander"] { border: 1px solid var(--nc-border); border-radius: 10px; background: var(--nc-surface); }

    /* Contextual rail */
    [data-testid="stColumn"]:has(.context-anchor) [data-testid="stExpander"] { margin-bottom: 12px; box-shadow: 0 5px 18px rgba(16, 36, 62, .035); }
    [data-testid="stColumn"]:has(.context-anchor) [data-testid="stExpander"] summary { min-height: 48px; color: var(--nc-strong); font-size: 14px; font-weight: 750; }
    .context-body { padding: 0 2px 3px; }
    .metric { display: flex; align-items: center; gap: 11px; padding: 8px 0; }
    .metric-icon { display: flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 8px; background: var(--nc-soft-blue); color: var(--nc-blue); }
    .metric-num { color: var(--nc-strong); font-size: 15px; font-weight: 750; font-variant-numeric: tabular-nums; }
    .metric-lbl, .activity-time { color: var(--nc-muted); font-size: 12px; }
    .activity { display: flex; justify-content: space-between; gap: 10px; padding: 10px 0; border-bottom: 1px solid #EDF1F5; color: var(--nc-text); font-size: 12px; }
    .activity:last-child { border-bottom: 0; }
    .activity span:first-child { min-width: 0; overflow-wrap: anywhere; }
    .activity-time { flex: 0 0 auto; }
    [data-testid="stColumn"]:has(.context-anchor) .stButton > button { min-height: 44px; border: 1px solid var(--nc-border); border-radius: 8px; background: var(--nc-surface); color: var(--nc-strong); }

    /* Tablet: drawer navigation, primary conversation, context below */
    @media (min-width: 768px) and (max-width: 1279px) {
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="stHeader"] { height: 0; min-height: 0; border: 0; }
        [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) {
            display: flex; width: auto; min-height: 44px;
            position: absolute; top: .55rem; left: 1rem; z-index: 100;
        }
        [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) [data-testid="stPopoverButton"] {
            height: 44px !important; min-height: 44px; border: 1px solid var(--nc-border); border-radius: 8px;
            background: var(--nc-surface); color: var(--nc-navy); font-weight: 700;
        }
        [data-testid="stMain"] { width: 100% !important; margin-left: 0 !important; }
        [data-testid="stMainBlockContainer"], .block-container { padding-inline: 1rem; }
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) { display: block; }
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) > [data-testid="stColumn"]:has(.conversation-anchor),
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) > [data-testid="stColumn"]:has(.context-anchor) {
            display: block; flex: 1 1 auto !important; width: 100% !important; min-width: 0 !important;
        }
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) > [data-testid="stColumn"]:has(.context-anchor) { margin-top: 22px; }
        [data-testid="stColumn"]:has(.conversation-anchor) > div { max-width: 1080px; }
        [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) { flex-wrap: wrap; }
        [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) > [data-testid="stColumn"] { flex: 1 1 calc(50% - 8px) !important; width: calc(50% - 8px) !important; min-width: 280px !important; }
        .hero { padding: 22px 0 20px; }
        .hero h1 { font-size: 25px; }
    }

    /* Mobile: drawer, single-column content, compact visuals */
    @media (max-width: 767px) {
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="stHeader"] { height: 0; min-height: 0; border: 0; }
        [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) {
            display: flex; width: 44px; min-height: 44px;
            position: absolute; top: .45rem; left: .65rem; z-index: 100;
        }
        [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) [data-testid="stPopoverButton"] {
            width: 44px; min-width: 44px; height: 44px !important; min-height: 44px; padding: 0;
            border: 1px solid var(--nc-border); border-radius: 8px;
            background: var(--nc-surface); color: var(--nc-navy); font-size: 0 !important;
        }
        [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) [data-testid="stPopoverButton"] [data-testid="stMarkdownContainer"],
        [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) [data-testid="stPopoverButton"] [aria-hidden="true"] { display: none; }
        [data-testid="stHorizontalBlock"]:has(.mobile-menu-anchor) [data-testid="stPopoverButton"] [data-testid="stIconMaterial"] { display: inline; font-size: 20px !important; }
        [data-testid="stMain"] { width: 100% !important; margin-left: 0 !important; }
        [data-testid="stMainBlockContainer"], .block-container { width: 100%; padding: .45rem .65rem 5.25rem; }
        .app-header { min-height: 46px; margin-bottom: 8px; padding: 2px 0 8px 52px; }
        .header-logo { width: 30px; height: 28px; }
        .header-title { font-size: 15px; white-space: nowrap; }
        .status { display: none; }
        .avatar { width: 34px; height: 34px; }
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) { display: block; }
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) > [data-testid="stColumn"]:has(.conversation-anchor),
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) > [data-testid="stColumn"]:has(.context-anchor) {
            display: block; flex: 1 1 auto !important; width: 100% !important; min-width: 0 !important;
        }
        [data-testid="stHorizontalBlock"]:has(.conversation-anchor):has(.context-anchor) > [data-testid="stColumn"]:has(.context-anchor) { margin-top: 18px; }
        .hero { padding: 12px 0 14px; }
        .hero-logo { width: 56px; }
        .hero h1 { margin-top: 8px; font-size: 21px; }
        .hero p { font-size: 14px; }
        [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) { display: block; }
        [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) > [data-testid="stColumn"] { width: 100% !important; min-width: 0 !important; margin-bottom: 9px; }
        [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) .stButton > button { min-height: 88px; padding: 15px; }
        [data-testid="stHorizontalBlock"]:has(.suggestion-anchor):not(:has(.conversation-anchor)) .stButton > button::before { left: 15px; top: 14px; }
        [data-testid="stChatMessage"] { padding: 14px; border-radius: 10px; }
        [data-testid="stChatMessage"]:has(.message-text.user) { width: 92%; }
        .message-text.rtl { font-size: 16px; line-height: 1.85; }
        .message-text.ltr { font-size: 15px; line-height: 1.7; }
        .analysis-heading, .visual-section-title { font-size: 17px; }
        [data-testid="stChatInput"] textarea { font-size: 16px !important; }
        .chips { justify-content: flex-start; flex-wrap: nowrap; overflow-x: auto; padding-bottom: 4px; }
        .chip { flex: 0 0 auto; }
        [data-testid="stHorizontalBlock"]:has(.kpi-card):not(:has(.conversation-anchor)) { flex-wrap: wrap; }
        [data-testid="stHorizontalBlock"]:has(.kpi-card):not(:has(.conversation-anchor)) > [data-testid="stColumn"] { flex: 1 1 calc(50% - 6px) !important; width: calc(50% - 6px) !important; min-width: 145px !important; }
        [data-testid="stHorizontalBlock"]:has(.export-anchor):not(:has(.conversation-anchor)) { display: block; }
        [data-testid="stHorizontalBlock"]:has(.export-anchor):not(:has(.conversation-anchor)) > [data-testid="stColumn"] { width: 100% !important; min-width: 0 !important; margin-bottom: 8px; }
        div[data-testid="stPlotlyChart"] { height: 300px !important; min-height: 280px !important; }
        div[data-testid="stPlotlyChart"] .js-plotly-plot,
        div[data-testid="stPlotlyChart"] .plot-container,
        div[data-testid="stPlotlyChart"] .svg-container { height: 100% !important; }
        .table-scroll { -webkit-overflow-scrolling: touch; }
        .nova-table { font-size: 12px; }
        .nova-table th, .nova-table td { padding: 9px 10px; }
        .security { margin-top: 20px; }
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
    if st.button("New Chat  +", use_container_width=True):
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


mobile_menu_col = st.columns(1)[0]
with mobile_menu_col:
    st.markdown('<span class="mobile-menu-anchor"></span>', unsafe_allow_html=True)
    with st.popover("Navigation", icon=":material/menu:", use_container_width=False):
        st.markdown(
            '<div class="mobile-menu-brand">NovaCore Solutions <span>Copilot</span></div>',
            unsafe_allow_html=True,
        )
        if st.button("New Chat  +", key="mobile_new_chat", use_container_width=True):
            st.session_state.messages = []
            st.session_state.last_df = None
            st.session_state.last_answer = ""
            st.rerun()
        st.markdown(
            f'<div class="mobile-nav"><div>{svg_icon("grid", 19)}<span>Dashboard</span></div>'
            f'<div>{svg_icon("database", 19)}<span>Data Explorer</span></div>'
            f'<div>{svg_icon("chart", 19)}<span>Reports</span></div>'
            f'<div>{svg_icon("bulb", 19)}<span>Insights</span></div>'
            f'<div>{svg_icon("bookmark", 19)}<span>Saved Chats</span></div>'
            f'<div>{svg_icon("settings", 19)}<span>Settings</span></div></div>',
            unsafe_allow_html=True,
        )


st.markdown(
    f'<div class="app-header"><div class="header-left"><img class="header-logo" src="{LOGO_URI}" alt="NovaCore">'
    f'<div class="header-title">NovaCore Solutions <span>Copilot</span></div></div>'
    f'<div class="header-actions"><div class="status"><span class="dot"></span>All Systems Operational</div>'
    f'<div class="avatar">NC</div></div></div>',
    unsafe_allow_html=True,
)

# Scoped anchors let the workspace respond without changing unrelated columns.
main_col, right_col = st.columns([1, 0.28], gap="large")

# =========================================================
# Main
# =========================================================

with main_col:
    st.markdown('<span class="conversation-anchor"></span>', unsafe_allow_html=True)
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
        for idx, (_, _, q0) in enumerate(cards):
            with prompt_cols[idx]:
                st.markdown('<span class="suggestion-anchor"></span>', unsafe_allow_html=True)
                if st.button(q0, key=f"suggested_prompt_{idx}", use_container_width=True):
                    st.session_state.pending_question = q0
                    st.rerun()

    # Conversation
    data_message_indexes = [
        idx for idx, item in enumerate(st.session_state.messages)
        if item.get("role") == "assistant" and item.get("data") is not None
    ]
    latest_data_index = data_message_indexes[-1] if data_message_indexes else -1
    for message_index, message in enumerate(st.session_state.messages):
        with st.chat_message(message["role"]):
            render_message_text(message["content"], message["role"])

            df = message.get("data")
            if df is not None and not df.empty:
                render_kpis(df)

                if message_index == latest_data_index:
                    evidence = message.get("driver_evidence") or {}
                    driver_rendered = render_driver_charts(
                        evidence,
                        message.get("question", "")
                    )
                    if not driver_rendered:
                        render_chart(df, message.get("question", ""))

                with st.expander("View supporting data · عرض البيانات الداعمة", expanded=False):
                    render_supporting_data(df)

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

        with st.chat_message("user"):
            render_message_text(question, "user")

        with st.spinner("Analyzing your data..."):
            try:
                output = ask(question, history, excel_mcp)
                answer = output["answer"]
                result_df = output["data"]
                result_plan = output.get("plan", {})
                driver_evidence = output.get("driver_evidence", {})
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
                result_plan = {}
                driver_evidence = {}

        st.session_state.messages.append(
            {
                "role": "assistant",
                "content": answer,
                "data": result_df,
                "question": question,
                "plan": result_plan,
                "driver_evidence": driver_evidence,
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
            st.markdown('<span class="export-anchor"></span>', unsafe_allow_html=True)
            st.download_button(
                "↓  Download CSV",
                csv_bytes(st.session_state.last_df),
                file_name="novacore_analysis.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            st.markdown('<span class="export-anchor"></span>', unsafe_allow_html=True)
            st.download_button(
                "↓  Download Excel",
                excel_bytes(st.session_state.last_df),
                file_name="novacore_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c3:
            st.markdown('<span class="export-anchor"></span>', unsafe_allow_html=True)
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
    st.markdown('<span class="context-anchor"></span>', unsafe_allow_html=True)
    try:
        record_count, employee_count, department_count = cached_overview_metrics()
    except Exception:
        record_count=employee_count=department_count=0
    with st.expander("Data Overview", expanded=True, icon=":material/database:"):
        st.markdown(f'<div class="context-body"><div class="metric"><div class="metric-icon">{svg_icon("briefcase",20)}</div><div><div class="metric-num">1</div><div class="metric-lbl">Datasets</div></div></div><div class="metric"><div class="metric-icon">{svg_icon("grid",20)}</div><div><div class="metric-num">{department_count}</div><div class="metric-lbl">Departments</div></div></div><div class="metric"><div class="metric-icon">{svg_icon("users",20)}</div><div><div class="metric-num">{employee_count:,}</div><div class="metric-lbl">Employees</div></div></div><div class="metric"><div class="metric-icon">{svg_icon("table",20)}</div><div><div class="metric-num">{record_count:,}</div><div class="metric-lbl">Records</div></div></div></div>', unsafe_allow_html=True)
    recent = st.session_state.get("recent_queries", [])
    if recent:
        activity=''.join(f'<div class="activity"><span>{html.escape(q[:34])}</span><span class="activity-time">just now</span></div>' for q in recent[:5])
    else:
        activity='<div class="activity"><span>No recent analysis yet</span><span class="activity-time">—</span></div>'
    with st.expander("Recent Activity", expanded=False, icon=":material/history:"):
        st.markdown(f'<div class="context-body">{activity}</div>', unsafe_allow_html=True)
    with st.expander("Quick Tools", expanded=False, icon=":material/bolt:"):
        st.markdown('<div class="tools-note">Utilities for the current dataset.</div>', unsafe_allow_html=True)
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
            except Exception as exc:
                st.error(str(exc))
