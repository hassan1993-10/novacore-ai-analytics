from __future__ import annotations
import asyncio, json, re
from typing import Any
import pandas as pd
from copilot import CopilotClient
from copilot.session import PermissionHandler
from copilot.session_events import AssistantMessageData, SessionIdleData
import os

DATASET = "novacore_enterprise_sample_data"
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
    client = CopilotClient({
        "github_token": token,
        "use_logged_in_user": False,
    })

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


def _schema(excel_mcp):
    overview = json.loads(excel_mcp.get_model_overview(DATASET))
    tables = overview.get("data", {}).get("tables", [])
    out = {}
    for t in tables:
        name = t.get("table_name") or t.get("name")
        if name:
            try:
                out[name] = [str(c) for c in excel_mcp.load_table(DATASET, name).columns]
            except Exception:
                pass
    return out

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
    rows=result.head(20).where(pd.notna(result),None).to_dict("records")
    p=f"""{SYSTEM}
User question: {question}
Verified calculation intent: {plan.get("intent")}
Verified result: {json.dumps(rows,ensure_ascii=False,default=str)}
Explain in 2-5 concise sentences. Mention strongest/weakest or trend when relevant.
Do not invent or recalculate. Do not render a markdown table."""
    return {"answer":_run(_ask(p)),"data":result,"plan":plan}
