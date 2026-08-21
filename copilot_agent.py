from __future__ import annotations
import asyncio, json, re, threading
from typing import Any
import pandas as pd
from copilot import CopilotClient
from copilot.session import PermissionHandler
from copilot.session_events import AssistantMessageData, SessionIdleData
import os

DATASET = "novacore_enterprise_sample_data"
_SCHEMA_CACHE = None
_SCHEMA_LOCK = threading.Lock()
SYSTEM = """You are NovaCore Solutions Copilot, an enterprise analytics assistant.
Reply in the same language as the user's latest message.
For Arabic use natural Arabic; for English use professional concise English.
Never invent data. Calculations are performed by the application, not by you.
When asked who you are, say you are NovaCore Solutions Copilot.
Keep answers concise, decision-oriented, and suitable for executives."""

def _run(coro):
    """Run one async Copilot request exactly once."""
    return asyncio.run(coro)

async def _ask(prompt: str) -> str:
    token = os.getenv("COPILOT_GITHUB_TOKEN") or os.getenv("GITHUB_TOKEN")
    if not token:
        raise RuntimeError("No Copilot GitHub token is configured.")

    # GitHub Copilot SDK official PAT/OAuth pattern:
    # pass the token directly to CopilotClient.
    client = CopilotClient(
        github_token=token,
        use_logged_in_user=False,
        base_directory="/tmp/novacore-copilot",
        log_level="debug",
    )

    await client.start()
    try:
        session = await client.create_session(
            on_permission_request=PermissionHandler.approve_all,
            model="gpt-5.4",
        )
        try:
            response = await session.send_and_wait(prompt, timeout=120)

            if response is None:
                raise RuntimeError("Copilot completed without an assistant message.")

            content = getattr(response.data, "content", "") or ""
            if not str(content).strip():
                raise RuntimeError("Copilot returned an empty assistant message.")

            return str(content).strip()
        finally:
            await session.disconnect()
    finally:
        await client.stop()


def _schema(excel_mcp, force_refresh: bool = False):
    """Cache schema so every question does not re-read all Excel tables."""
    global _SCHEMA_CACHE
    if _SCHEMA_CACHE is not None and not force_refresh:
        return _SCHEMA_CACHE
    with _SCHEMA_LOCK:
        if _SCHEMA_CACHE is not None and not force_refresh:
            return _SCHEMA_CACHE
        overview = json.loads(excel_mcp.get_model_overview(DATASET))
        tables = overview.get("data", {}).get("tables", [])
        out = {}
        for t in tables:
            name = t.get("table_name") or t.get("name")
            if not name:
                continue
            raw_cols = t.get("columns") or []
            cols = []
            for c in raw_cols:
                value = (c.get("column_name") or c.get("name")) if isinstance(c, dict) else c
                if value:
                    cols.append(str(value))
            if not cols:
                try:
                    cols = [str(c) for c in excel_mcp.load_table(DATASET, name).columns]
                except Exception:
                    cols = []
            if cols:
                out[name] = cols
        _SCHEMA_CACHE = out
        return _SCHEMA_CACHE

def clear_schema_cache():
    global _SCHEMA_CACHE
    _SCHEMA_CACHE = None

def _extract_json(text: str):
    text = re.sub(r"^```(?:json)?\s*", "", text.strip())
    text = re.sub(r"\s*```$", "", text)
    a, b = text.find("{"), text.rfind("}")
    if a < 0 or b < a:
        raise ValueError("Copilot did not return a valid analysis plan.")
    return json.loads(text[a:b+1])

def _build_plan(question, schema, history):
    prompt = f"""{SYSTEM}
You are the PLANNER. Convert the latest user question into ONE safe JSON analytics plan.
Available Excel schema:
{json.dumps(schema, ensure_ascii=False)}
Recent conversation:
{json.dumps((history or [])[-6:], ensure_ascii=False)}
Latest question:
{question}
Return JSON only:
{{
 "mode":"chat|analysis","table":null,"group_by":[],"metric_column":null,
 "aggregation":"sum|mean|count|nunique|min|max","filters":[],
 "date_column":null,"date_grain":null,"sort":"desc","limit":20,"intent":""
}}
Filters use: {{"column":"exact column","op":"eq|neq|contains|gt|gte|lt|lte","value":"value"}}.
Use only exact schema names. For year/month/quarter set date_column + date_grain.
mode=chat for greetings, identity, help, or non-data questions. Never output Python."""
    return _extract_json(_run(_ask(prompt)))

def _apply_filter(df, f):
    col, op, value = f.get("column"), f.get("op"), f.get("value")
    if col not in df.columns: return df
    s = df[col]
    if op == "contains": return df[s.astype(str).str.contains(str(value), case=False, na=False)]
    if op in {"gt","gte","lt","lte"}:
        n, v = pd.to_numeric(s, errors="coerce"), float(value)
        return df[{"gt":n>v,"gte":n>=v,"lt":n<v,"lte":n<=v}[op]]
    if op == "neq": return df[s.astype(str).str.casefold()!=str(value).casefold()]
    return df[s.astype(str).str.casefold()==str(value).casefold()]

def _execute(plan, excel_mcp):
    table = plan.get("table")
    if not table: raise ValueError("No table selected.")
    df = excel_mcp.load_table(DATASET, table).copy()
    for f in plan.get("filters") or []: df = _apply_filter(df, f)
    groups = [g for g in (plan.get("group_by") or []) if g in df.columns]
    dc, grain = plan.get("date_column"), plan.get("date_grain")
    if dc in df.columns and grain in {"year","quarter","month"}:
        dt = pd.to_datetime(df[dc], errors="coerce")
        new = {"year":"Year","quarter":"Year_Quarter","month":"Year_Month"}[grain]
        df[new] = dt.dt.year.astype("Int64") if grain=="year" else dt.dt.to_period("Q" if grain=="quarter" else "M").astype(str)
        groups = [new] + groups
    metric, agg = plan.get("metric_column"), plan.get("aggregation","sum")
    if agg=="count" and not metric:
        result = df.groupby(groups,dropna=False).size().reset_index(name="Count") if groups else pd.DataFrame({"Count":[len(df)]})
    else:
        if metric not in df.columns: raise ValueError(f"Metric column '{metric}' was not found.")
        if agg in {"sum","mean","min","max"}: df[metric]=pd.to_numeric(df[metric],errors="coerce")
        if groups: result=df.groupby(groups,dropna=False)[metric].agg(agg).reset_index()
        else:
            value=df[metric].nunique() if agg=="nunique" else getattr(df[metric],agg)()
            result=pd.DataFrame({f"{agg}_{metric}":[value]})
    last=result.columns[-1]
    if pd.api.types.is_numeric_dtype(result[last]):
        result=result.sort_values(last,ascending=plan.get("sort")=="asc")
    return result.head(int(plan.get("limit") or 20)).reset_index(drop=True)

def _fmt_value(value):
    if pd.isna(value): return "-"
    if isinstance(value, (int, float)):
        n=float(value)
        if abs(n)>=1_000_000_000: return f"{n/1_000_000_000:.2f}B"
        if abs(n)>=1_000_000: return f"{n/1_000_000:.2f}M"
        if abs(n)>=1_000: return f"{n:,.0f}"
        return f"{n:,.2f}".rstrip("0").rstrip(".")
    return str(value)

def _local_summary(question, result, plan):
    """Fast narrative from verified dataframe; removes the second LLM call."""
    ar=bool(re.search(r"[\u0600-\u06FF]", str(question or "")))
    if result is None or result.empty:
        return "لا توجد بيانات مطابقة للسؤال." if ar else "No matching data was found."
    if len(result)==1:
        col=result.columns[-1]; val=_fmt_value(result.iloc[0,-1])
        return f"النتيجة المحسوبة هي **{val}** ({col})." if ar else f"The calculated result is **{val}** ({col})."
    valcol=result.columns[-1]
    if pd.api.types.is_numeric_dtype(result[valcol]):
        label=result.columns[0]
        ranked=result.sort_values(valcol,ascending=False).reset_index(drop=True)
        hi,lo=ranked.iloc[0],ranked.iloc[-1]
        time_like=any(k in str(label).casefold() for k in ["year","month","quarter","date","period","time"])
        if time_like:
            ordered=result.reset_index(drop=True); first,last=ordered.iloc[0][valcol],ordered.iloc[-1][valcol]
            try: pct=((float(last)-float(first))/abs(float(first))*100) if float(first)!=0 else None
            except Exception: pct=None
            if ar:
                trend="ارتفع" if pct and pct>0 else "انخفض" if pct and pct<0 else "استقر"
                extra=f" بنسبة **{abs(pct):.1f}%**" if pct is not None else ""
                return f"الاتجاه العام **{trend}** من {_fmt_value(first)} إلى {_fmt_value(last)}{extra}. أعلى قيمة كانت **{_fmt_value(hi[valcol])}**."
            trend="increased" if pct and pct>0 else "decreased" if pct and pct<0 else "was broadly stable"
            extra=f" by **{abs(pct):.1f}%**" if pct is not None else ""
            return f"The overall trend **{trend}** from {_fmt_value(first)} to {_fmt_value(last)}{extra}. Peak: **{_fmt_value(hi[valcol])}**."
        return (f"أعلى نتيجة هي **{hi[label]}** بقيمة **{_fmt_value(hi[valcol])}**، والأقل **{lo[label]}** بقيمة **{_fmt_value(lo[valcol])}**."
                if ar else f"**{hi[label]}** is highest at **{_fmt_value(hi[valcol])}**; **{lo[label]}** is lowest at **{_fmt_value(lo[valcol])}**.")
    return "تم تنفيذ التحليل وإرجاع النتائج الموثقة أدناه." if ar else "The analysis was completed and the verified results are shown below."

def ask(question, history, excel_mcp):
    schema=_schema(excel_mcp)
    plan=_build_plan(question,schema,history)
    if plan.get("mode")=="chat":
        p=f"""{SYSTEM}
Recent conversation: {json.dumps((history or [])[-6:],ensure_ascii=False)}
User: {question}
Answer directly. Do not mention implementation details, demo mode, or GitHub Models."""
        return {"answer":_run(_ask(p)),"data":None,"plan":plan}
    result=_execute(plan,excel_mcp)
    # Performance V2: one Copilot request for analytics (planning), then local verified summary.
    return {"answer":_local_summary(question,result,plan),"data":result,"plan":plan}
