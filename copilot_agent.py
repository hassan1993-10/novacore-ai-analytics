from __future__ import annotations
import asyncio, json, logging, re, threading, time
from typing import Any
import pandas as pd
from copilot import CopilotClient
from copilot.session import PermissionHandler
from copilot.session_events import AssistantMessageData, SessionIdleData
import os

DATASET = "novacore_enterprise_sample_data"
_SCHEMA_CACHE = None
_SCHEMA_LOCK = threading.Lock()
_LOGGER = logging.getLogger("novacore.copilot")
_ANALYTICAL_INTENTS = {
    "single_kpi", "trend", "comparison", "ranking", "root_cause",
    "domain_overview", "multi_domain_overview", "enterprise_overview",
}
_OVERVIEW_INTENTS = {
    "domain_overview", "multi_domain_overview", "enterprise_overview",
}
_DOMAIN_ALIASES = {
    "sales": ["sales", "sale", "orders", "commercial", "مبيعات", "المبيعات", "طلب", "الطلبات"],
    "finance": ["finance", "financial", "accounting", "budget", "مالية", "المالية", "تمويل", "ميزانية"],
    "procurement": ["procurement", "purchasing", "purchase", "vendor", "supplier", "مشتريات", "المشتريات", "مورد", "الموردين"],
    "hr": ["hr", "human resources", "workforce", "employee", "employees", "موارد بشرية", "الموارد البشرية", "موظف", "الموظفين"],
    "it": ["it", "information technology", "technology", "ticket", "service desk", "تقنية المعلومات", "تقنية", "تذاكر", "الدعم الفني"],
}
_METRIC_CONCEPTS = {
    "net_revenue": ["net revenue", "net sales", "صافي الايرادات", "صافي الإيرادات", "صافي المبيعات"],
    "revenue": ["revenue", "sales amount", "الايرادات", "الإيرادات", "المبيعات"],
    "gross_profit": ["gross profit", "profit", "اجمالي الربح", "إجمالي الربح", "الربح"],
    "cost": ["cost", "expense", "expenses", "التكلفة", "تكلفة", "المصروفات"],
    "amount": ["amount", "value", "المبلغ", "القيمة"],
    "quantity": ["quantity", "units", "الكمية", "الوحدات"],
    "budget": ["budget", "الميزانية"],
    "spend": ["spend", "spent", "الإنفاق", "الانفاق"],
    "headcount": ["headcount", "employees", "employee count", "عدد الموظفين"],
    "tickets": ["tickets", "ticket count", "عدد التذاكر"],
    "orders": ["orders", "order count", "عدد الطلبات", "الطلبات"],
    "resolution_time": ["resolution time", "resolve time", "زمن الحل", "وقت الحل"],
    "salary": ["salary", "payroll", "compensation", "راتب", "الرواتب", "الأجور"],
    "margin": ["margin", "هامش"],
    "discount": ["discount", "خصم", "الخصومات"],
    "balance": ["balance", "رصيد"],
    "cash_flow": ["cash flow", "cashflow", "التدفق النقدي"],
    "variance": ["variance", "فرق", "انحراف"],
    "savings": ["savings", "saving", "وفورات", "توفير"],
    "cycle_time": ["cycle time", "lead time", "زمن الدورة", "مدة التوريد"],
    "satisfaction": ["satisfaction", "csat", "رضا"],
    "uptime": ["uptime", "availability", "الإتاحة", "التوافر"],
    "downtime": ["downtime", "outage", "التوقف", "الانقطاع"],
}
SYSTEM = """You are NovaCore Solutions Copilot, an enterprise analytics assistant.
Reply in the same language as the user's latest message.
For Arabic use natural Arabic; for English use professional concise English.
Never invent data. Calculations are performed by the application, not by you.
When asked who you are, say you are NovaCore Solutions Copilot.
Keep answers concise, decision-oriented, and suitable for executives."""


class _RequestCache:
    """Request-scoped data cache; business data is never retained across requests."""
    def __init__(self, excel_mcp):
        self.excel_mcp = excel_mcp
        self.frames: dict[str, pd.DataFrame] = {}
        self.data_loading_ms = 0.0

    def load(self, table: str) -> pd.DataFrame:
        if table not in self.frames:
            started = time.perf_counter()
            self.frames[table] = self.excel_mcp.load_table(DATASET, table)
            self.data_loading_ms += (time.perf_counter() - started) * 1000
        return self.frames[table]

class _CopilotRuntime:
    """Keep one CopilotClient alive for the lifetime of the Streamlit worker."""
    def __init__(self):
        self._loop = asyncio.new_event_loop()
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        self._client = None
        self._token = None

    def _run_loop(self):
        asyncio.set_event_loop(self._loop)
        self._loop.run_forever()

    async def _ensure_client(self):
        token = os.getenv("COPILOT_GITHUB_TOKEN") or os.getenv("GITHUB_TOKEN")
        if not token:
            raise RuntimeError("No Copilot GitHub token is configured.")

        if self._client is not None and self._token == token:
            return self._client

        if self._client is not None:
            try:
                await self._client.stop()
            except Exception:
                pass

        self._client = CopilotClient(
            github_token=token,
            use_logged_in_user=False,
            base_directory="/tmp/novacore-copilot",
            log_level="error",
        )
        await self._client.start()
        self._token = token
        return self._client

    async def _ask_once(self, prompt: str) -> str:
        client = await self._ensure_client()
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
            try:
                await session.disconnect()
            except Exception:
                pass

    async def _ask_with_recovery(self, prompt: str) -> str:
        try:
            return await self._ask_once(prompt)
        except Exception:
            # Recreate only when the persistent client has actually failed.
            if self._client is not None:
                try:
                    await self._client.stop()
                except Exception:
                    pass
            self._client = None
            self._token = None
            return await self._ask_once(prompt)

    def ask(self, prompt: str) -> str:
        fut = asyncio.run_coroutine_threadsafe(
            self._ask_with_recovery(prompt), self._loop
        )
        return fut.result(timeout=135)


_RUNTIME = _CopilotRuntime()


def _ask_sync(prompt: str) -> str:
    return _RUNTIME.ask(prompt)


def _schema(excel_mcp, force_refresh: bool = False, request_cache: _RequestCache | None = None):
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
                    frame = request_cache.load(name) if request_cache is not None else excel_mcp.load_table(DATASET, name)
                    cols = [str(c) for c in frame.columns]
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
You are the INTENT-FIRST ANALYTICS PLANNER. Understand the question and its conversational context before selecting a table, metric, grouping, or filter.
Available Excel schema:
{json.dumps(schema, ensure_ascii=False)}
Recent conversation:
{json.dumps((history or [])[-6:], ensure_ascii=False)}
Latest question:
{question}
Return JSON only:
{{
 "mode":"chat|analysis",
 "intent":"single_kpi|trend|comparison|ranking|root_cause|domain_overview|multi_domain_overview|enterprise_overview",
 "scope":"single_domain|domain|multi_domain|enterprise",
 "domains":[],"table":null,"group_by":[],"metric_column":null,"metric_semantic":null,
 "aggregation":"sum|mean|count|nunique|min|max","filters":[],
 "date_column":null,"date_grain":null,"sort":"desc|asc","limit":20,
 "analysis_depth":"standard|deep","clarification":null
}}
Filters use: {{"column":"exact column","op":"eq|neq|contains|gt|gte|lt|lte","value":"value"}}.
INTENT DEFINITIONS:
- single_kpi: one requested verified KPI with no grouping.
- trend: one metric over time.
- comparison: grouped comparison by a non-time dimension.
- ranking: explicit top/bottom N request.
- root_cause: why, decline drivers, contribution, diagnosis, or evidence-based recommendations.
- domain_overview: broad performance summary for one business domain; this is multi-metric and table may initially be null.
- multi_domain_overview: broad summary of only the explicitly requested domains; table should be null.
- enterprise_overview: broad company/executive/overall performance summary across all meaningful available domains; table should be null.
STRICT RULES:
- First resolve intent, scope, and domains semantically. Do not classify broad performance summaries as single_kpi.
- Use the recent conversation to resolve elliptical follow-ups such as 'by region?', 'what about products?', or 'focus on finance'.
- Use actual schema table and column names only. For overview intents, Python will discover metrics and tables, so do not force one metric.
- group_by must contain ONLY dimensions explicitly requested by the user.
- Never add unrelated dimensions.
- "by year / حسب السنة / حسب السنوات" means aggregate ONLY by year unless another dimension is explicitly requested.
- Filters are not grouping dimensions.
- Set analysis_depth=deep only for root_cause and overview intents; otherwise standard.
- For year/month/quarter requests set date_column and date_grain.
- If a data request is genuinely ambiguous after using history and schema, set clarification to one concise question.
- mode=chat only for greetings, identity, help, or non-data questions. Never output Python."""
    return _extract_json(_ask_sync(prompt))

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


def _qnorm(text: str) -> str:
    return re.sub(r"[\W_]+", " ", str(text or "").casefold(), flags=re.UNICODE).strip()

def _mentions_any(question: str, tokens) -> bool:
    q = f" {_qnorm(question)} "
    return any(f" {_qnorm(t)} " in q for t in tokens)


def _semantic_metric(question: str) -> str | None:
    q = _qnorm(question)
    # Specific concepts must win over generic ones (net revenue before revenue).
    for concept, aliases in _METRIC_CONCEPTS.items():
        if any(_qnorm(alias) in q for alias in aliases):
            return concept
    return None


def _domain_from_text(text: str) -> list[str]:
    found = []
    for domain, aliases in _DOMAIN_ALIASES.items():
        if _mentions_any(text, aliases):
            found.append(domain)
    return found


def _domain_label(domain: str) -> str:
    labels = {"sales":"Sales", "finance":"Finance", "procurement":"Procurement", "hr":"HR", "it":"IT"}
    return labels.get(str(domain).casefold(), str(domain).replace("_", " ").title())


def _normalize_intent(plan: dict) -> str:
    intent = str(plan.get("intent") or "").casefold().strip()
    aliases = {
        "scalar":"single_kpi", "kpi":"single_kpi", "temporal":"trend",
        "driver":"root_cause", "diagnosis":"root_cause", "recommendation":"root_cause",
    }
    intent = aliases.get(intent, intent)
    if intent in _ANALYTICAL_INTENTS:
        return intent
    if plan.get("date_grain"):
        return "trend"
    if plan.get("group_by"):
        return "comparison"
    return "single_kpi"

_DIMENSION_HINTS = {
    "region": ["region", "regions", "منطقة", "المناطق", "المنطقة"],
    "status": ["status", "order status", "حالة", "الحالة", "حالات"],
    "product": ["product", "products", "منتج", "منتجات", "المنتج", "المنتجات"],
    "department": ["department", "departments", "قسم", "القسم", "الأقسام", "الاقسام"],
    "vendor": ["vendor", "vendors", "supplier", "suppliers", "مورد", "المورد", "الموردين"],
    "employee": ["employee", "employees", "موظف", "الموظف", "الموظفين"],
    "category": ["category", "categories", "فئة", "الفئة", "الفئات"],
    "priority": ["priority", "priorities", "أولوية", "الاولوية", "الأولوية"],
    "service": ["service", "services", "خدمة", "الخدمة", "الخدمات"],
}

_TIME_HINTS = {
    "year": ["year", "years", "annual", "سنة", "السنة", "سنوات", "السنوات", "سنوي"],
    "quarter": ["quarter", "quarters", "ربع", "ربع سنوي", "أرباع", "الارباع"],
    "month": ["month", "months", "monthly", "شهر", "الشهر", "أشهر", "الاشهر", "شهري"],
}

def _column_semantic_family(col: str) -> str | None:
    c = _qnorm(col)
    for family, hints in _DIMENSION_HINTS.items():
        if any(_qnorm(h) in c for h in hints):
            return family
    return None

def _sanitize_plan(question: str, plan: dict) -> dict:
    """Prevent planner over-grouping and keep result grain aligned to the question."""
    plan = dict(plan or {})
    if str(plan.get("mode") or "analysis").casefold() == "chat":
        plan["mode"] = "chat"
        return plan

    plan["mode"] = "analysis"
    plan["intent"] = _normalize_intent(plan)
    plan["domains"] = [str(x).casefold() for x in (plan.get("domains") or []) if x]
    if not plan["domains"]:
        plan["domains"] = _domain_from_text(question)

    intent = plan["intent"]
    if intent == "enterprise_overview":
        plan["scope"], plan["table"] = "enterprise", None
    elif intent == "multi_domain_overview":
        plan["scope"], plan["table"] = "multi_domain", None
    elif intent == "domain_overview":
        plan["scope"], plan["table"] = "domain", None
    else:
        plan.setdefault("scope", "single_domain")

    groups = list(plan.get("group_by") or [])

    # Deterministic business-term resolver: do not let the LLM switch measures
    # between equivalent questions. Exact schema names are resolved later.
    requested_metric = _semantic_metric(question)
    if requested_metric:
        plan["metric_semantic"] = requested_metric
    q = str(question or "")

    # Keep only categorical dimensions explicitly requested in the question.
    kept = []
    for g in groups:
        family = _column_semantic_family(g)
        if family is None:
            # Unknown dimensions are kept only when their column label itself is mentioned.
            label = str(g).replace("_", " ")
            if _qnorm(label) and _qnorm(label) in _qnorm(q):
                kept.append(g)
        elif _mentions_any(q, _DIMENSION_HINTS[family]):
            kept.append(g)

    # For an explicit time-only request, unrelated groups must not leak into result grain.
    explicit_time = None
    for grain, hints in _TIME_HINTS.items():
        if _mentions_any(q, hints):
            explicit_time = grain
            break

    if explicit_time:
        plan["date_grain"] = explicit_time
        # preserve only explicitly requested categorical groups
        plan["group_by"] = kept
    else:
        plan["group_by"] = kept

    if intent == "root_cause" and plan.get("date_column") and not plan.get("date_grain"):
        plan["date_grain"] = "year"

    # Ranking direction and N remain deterministic guardrails after semantic planning.
    ranking_match = re.search(r"(?:top|bottom|أعلى|اعلى|أقل|اقل)\s*(\d+)", str(question), re.IGNORECASE)
    if intent == "ranking":
        if ranking_match:
            plan["limit"] = max(1, min(int(ranking_match.group(1)), 100))
        bottom = _mentions_any(question, ["bottom", "lowest", "least", "أقل", "اقل", "الأدنى", "الادنى"])
        plan["sort"] = "asc" if bottom else "desc"

    # Deep analysis signal can be corrected deterministically.
    deep_tokens = [
        "why","cause","causes","reason","reasons","analyze","analyse","analysis",
        "recommend","recommendation","recommendations","risk","risks","opportunity",
        "opportunities","driver","drivers","root cause",
        "لماذا","سبب","اسباب","أسباب","حلل","تحليل","فسر","اشرح","توصية","توصيات",
        "مخاطر","فرص","محركات","جذر"
    ]
    if intent in _OVERVIEW_INTENTS or intent == "root_cause" or _mentions_any(q, deep_tokens):
        plan["analysis_depth"] = "deep"
    else:
        plan["analysis_depth"] = "standard"
    return plan

def _resolve_metric(plan: dict, df: pd.DataFrame) -> dict:
    """Resolve business metric semantics against real columns, preferring exact matches."""
    plan = dict(plan)
    semantic = plan.get("metric_semantic")
    aliases = {
        "net_revenue": ["Net_Revenue_SAR", "Net Revenue SAR", "Net_Revenue", "Net Revenue"],
        "revenue": ["Revenue_SAR", "Revenue SAR", "Revenue"],
        "gross_profit": ["Gross_Profit_SAR", "Gross Profit SAR", "Gross_Profit", "Gross Profit"],
        "cost": ["Cost_SAR", "Cost SAR", "Total_Cost_SAR", "Total Cost", "Cost", "Expenses"],
        "amount": ["Amount_SAR", "Amount SAR", "Amount"],
        "quantity": ["Quantity"],
        "budget": ["Budget_SAR", "Budget SAR", "Budget"],
        "spend": ["Spend_SAR", "Spend SAR", "Total_Spend", "Spend"],
        "resolution_time": ["Resolution_Time_Hours", "Resolution Time Hours", "Resolution_Time", "Resolution Time"],
    }
    if semantic:
        norm_cols = {_qnorm(c): c for c in df.columns}
        resolved = None
        for candidate in aliases.get(semantic, []):
            if _qnorm(candidate) in norm_cols:
                resolved = norm_cols[_qnorm(candidate)]
                break
        # Schema-specific names may include currency/unit suffixes. Match the full
        # semantic token sequence, but never fall back to another concept.
        if resolved is None:
            concept_tokens = _qnorm(semantic).split()
            matches = [c for c in df.columns if all(t in _qnorm(c).split() for t in concept_tokens)]
            if matches:
                resolved = matches[0]
        plan["metric_column"] = resolved

    if plan.get("metric_column") not in df.columns and semantic in {"orders", "headcount", "tickets"}:
        id_terms = {
            "orders": ["order id", "order number"],
            "headcount": ["employee id", "employee number"],
            "tickets": ["ticket id", "ticket number"],
        }[semantic]
        match = next((c for c in df.columns if any(_qnorm(t) in _qnorm(c) for t in id_terms)), None)
        if match:
            plan["metric_column"] = match
            plan["aggregation"] = "nunique"
    return plan

def _request_frame(excel_mcp, table: str, request_cache: _RequestCache | None = None) -> pd.DataFrame:
    return request_cache.load(table) if request_cache is not None else excel_mcp.load_table(DATASET, table)


def _parse_datetime_series(series: pd.Series, column: str = "") -> pd.Series:
    if "year" in _qnorm(column).split() or "سنة" in _qnorm(column).split():
        return pd.to_datetime(series.astype("string"), format="%Y", errors="coerce")
    return pd.to_datetime(series, errors="coerce")


def _driver_evidence(question: str, plan: dict, excel_mcp, request_cache: _RequestCache | None = None) -> dict:
    """Compute verified contribution-to-change evidence for deep/root-cause questions.
    This identifies data drivers (not causal proof) across explicitly requested dimensions.
    """
    table = plan.get("table")
    if not table:
        return {}
    df = _request_frame(excel_mcp, table, request_cache)
    plan = _resolve_metric(plan, df)
    for f in plan.get("filters") or []:
        df = _apply_filter(df, f)
    metric = plan.get("metric_column")
    if metric not in df.columns:
        return {}
    df = df.copy()
    df[metric] = pd.to_numeric(df[metric], errors="coerce")

    dc = plan.get("date_column")
    if dc not in df.columns:
        return {}
    dt = _parse_datetime_series(df[dc], dc)
    df["__Year"] = dt.dt.year.astype("Int64")
    years = sorted([int(x) for x in df["__Year"].dropna().unique()])
    if len(years) < 2:
        return {}
    prev_y, curr_y = years[-2], years[-1]
    totals = df[df["__Year"].isin([prev_y,curr_y])].groupby("__Year")[metric].sum()
    prev_total, curr_total = float(totals.get(prev_y,0)), float(totals.get(curr_y,0))
    total_delta = curr_total-prev_total

    requested=[]
    q=str(question or "")
    for family,hints in _DIMENSION_HINTS.items():
        if _mentions_any(q,hints): requested.append(family)
    # For why/root-cause questions, test the most useful driver dimensions even if not named.
    if not requested and _needs_deep_insight(question, plan):
        requested=["region","product","department","category"]

    evidence={"period":{"from":prev_y,"to":curr_y,"from_value":prev_total,"to_value":curr_total,"delta":total_delta,"change_pct":((total_delta/abs(prev_total))*100 if prev_total else None)},"drivers":{}}
    for family in requested[:3]:
        candidates=[c for c in df.columns if _column_semantic_family(c)==family]
        if not candidates: continue
        col=candidates[0]
        g=(df[df["__Year"].isin([prev_y,curr_y])].groupby([col,"__Year"],dropna=False)[metric].sum().unstack(fill_value=0))
        if prev_y not in g.columns: g[prev_y]=0
        if curr_y not in g.columns: g[curr_y]=0
        g["delta"]=g[curr_y]-g[prev_y]
        g["contribution_pct"]=(g["delta"]/abs(total_delta)*100) if total_delta else 0
        ranked = pd.concat([g.nsmallest(4, "delta"), g.nlargest(4, "delta")]).drop_duplicates().sort_values("delta")
        rows=[]
        for idx,row in ranked.iterrows():
            rows.append({"name":str(idx),"from":float(row[prev_y]),"to":float(row[curr_y]),"delta":float(row["delta"]),"contribution_pct":float(row["contribution_pct"])})
        evidence["drivers"][family]={"column":col,"rows":rows}
    return evidence

def _execute(plan, excel_mcp, request_cache: _RequestCache | None = None):
    table = plan.get("table")
    if not table:
        raise ValueError("No table selected.")

    df = _request_frame(excel_mcp, table, request_cache)
    plan = _resolve_metric(plan, df)
    for f in plan.get("filters") or []:
        df = _apply_filter(df, f)

    groups = [g for g in (plan.get("group_by") or []) if g in df.columns]
    dc, grain = plan.get("date_column"), plan.get("date_grain")
    time_group = None

    if dc in df.columns and grain in {"year","quarter","month"}:
        df = df.copy()
        dt = _parse_datetime_series(df[dc], dc)
        time_group = {"year":"Year","quarter":"Year_Quarter","month":"Year_Month"}[grain]
        if grain == "year":
            df[time_group] = dt.dt.year.astype("Int64")
        else:
            df[time_group] = dt.dt.to_period("Q" if grain=="quarter" else "M").astype(str)

        # Never duplicate the source date column or the derived time label.
        groups = [g for g in groups if g not in {dc, time_group}]
        groups = [time_group] + groups

    metric = plan.get("metric_column")
    agg = plan.get("aggregation", "sum")

    if agg == "count" and not metric:
        result = (
            df.groupby(groups, dropna=False).size().reset_index(name="Count")
            if groups else pd.DataFrame({"Count":[len(df)]})
        )
    else:
        if metric not in df.columns:
            requested = plan.get("metric_semantic") or metric or "requested metric"
            raise ValueError(f"The requested metric '{requested}' is not available in the selected data.")
        if agg in {"sum","mean","min","max"}:
            if dc not in df.columns or grain not in {"year", "quarter", "month"}:
                df = df.copy()
            df[metric] = pd.to_numeric(df[metric], errors="coerce")

        if groups:
            if agg == "nunique":
                result = df.groupby(groups, dropna=False)[metric].nunique().reset_index(name=metric)
            else:
                result = df.groupby(groups, dropna=False)[metric].agg(agg).reset_index()
        else:
            value = df[metric].nunique() if agg=="nunique" else getattr(df[metric], agg)()
            result = pd.DataFrame({f"{agg}_{metric}":[value]})

    # Time results must remain chronological; rankings remain metric-sorted.
    if time_group and time_group in result.columns:
        try:
            result = result.sort_values(time_group, ascending=True)
        except Exception:
            pass
    else:
        last = result.columns[-1]
        if pd.api.types.is_numeric_dtype(result[last]):
            result = result.sort_values(last, ascending=plan.get("sort")=="asc")

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

def _fmt_label(value):
    if pd.isna(value):
        return "-"
    if isinstance(value, (int, float)):
        n = float(value)
        if n.is_integer():
            return str(int(n))
    return str(value)


def _column_metric_concept(column: str) -> str | None:
    name = _qnorm(column)
    for concept, aliases in _METRIC_CONCEPTS.items():
        if concept in {"orders", "headcount", "tickets"}:
            continue
        if any(_qnorm(alias) in name for alias in aliases):
            return concept
    return None


def _technical_column(column: str) -> bool:
    name = _qnorm(column)
    tokens = set(name.split())
    return bool(tokens & {"id", "key", "code", "flag", "index", "row", "number", "no"}) or name.endswith(" id")


def _discover_domains(schema: dict) -> dict[str, list[str]]:
    """Map actual workbook tables to semantic business domains without assuming a fixed workbook."""
    discovered: dict[str, list[str]] = {}
    for table, columns in schema.items():
        haystack = " ".join([str(table)] + [str(c) for c in columns])
        scores = {domain: sum(1 for alias in aliases if _mentions_any(haystack, [alias])) for domain, aliases in _DOMAIN_ALIASES.items()}
        best = max(scores, key=scores.get) if scores and max(scores.values()) > 0 else None
        key = best or _qnorm(table).replace(" ", "_")
        discovered.setdefault(key, []).append(table)
    return discovered


def _resolve_requested_domains(plan: dict, schema: dict) -> list[str]:
    discovered = _discover_domains(schema)
    requested = [str(x).casefold() for x in (plan.get("domains") or []) if x]
    resolved = []
    for value in requested:
        if value in discovered:
            resolved.append(value)
            continue
        match = next((d for d, aliases in _DOMAIN_ALIASES.items() if value == d or _mentions_any(value, aliases)), None)
        if match in discovered:
            resolved.append(match)
            continue
        table_match = next((d for d, tables in discovered.items() if any(_qnorm(value) in _qnorm(t) for t in tables)), None)
        if table_match:
            resolved.append(table_match)
    return list(dict.fromkeys(resolved))


def _select_date_column(df: pd.DataFrame) -> str | None:
    named = [c for c in df.columns if any(t in _qnorm(c).split() for t in ["date", "time", "month", "year", "period"])]
    for col in named:
        parsed = _parse_datetime_series(df[col], col)
        if len(df) == 0 or parsed.notna().mean() >= 0.6:
            return col
    for col in df.select_dtypes(include=["datetime", "datetimetz"]).columns:
        return str(col)
    return None


def _select_dimensions(df: pd.DataFrame) -> list[str]:
    scored = []
    for col in df.columns:
        if _technical_column(col):
            continue
        unique = int(df[col].nunique(dropna=True))
        if unique < 2 or unique > min(100, max(20, len(df) // 2)):
            continue
        family = _column_semantic_family(col)
        semantic = 2 if family else 0
        text_type = 1 if not pd.api.types.is_numeric_dtype(df[col]) else 0
        scored.append((semantic + text_type, unique, str(col)))
    scored.sort(key=lambda x: (-x[0], x[1], x[2]))
    return [x[2] for x in scored[:4]]


def _select_business_metrics(df: pd.DataFrame, domain: str) -> list[dict]:
    """Select 3–6 meaningful metrics from real columns; identifiers remain excluded."""
    selected = []
    seen_concepts = set()
    priority = {
        "net_revenue": 100, "revenue": 95, "gross_profit": 90, "amount": 80,
        "spend": 80, "budget": 75, "cost": 70, "quantity": 60,
        "balance": 80, "cash_flow": 80, "salary": 75, "margin": 70,
        "savings": 70, "variance": 65, "discount": 60, "resolution_time": 60,
        "cycle_time": 60, "satisfaction": 60, "uptime": 60, "downtime": 60,
    }
    candidates = []
    for col in df.columns:
        concept = _column_metric_concept(col)
        if not concept or _technical_column(col):
            continue
        numeric = pd.to_numeric(df[col], errors="coerce")
        if len(df) and numeric.notna().mean() < 0.6:
            continue
        mean_concepts = {"resolution_time", "cycle_time", "satisfaction", "margin", "uptime", "downtime"}
        average_named = any(t in _qnorm(col) for t in ["rate", "average", "avg", "time", "duration", "score", "percent", "pct"])
        salary_average = concept == "salary" and "payroll" not in _qnorm(col)
        agg = "mean" if concept in mean_concepts or average_named or salary_average else "sum"
        candidates.append((priority.get(concept, 50), str(col), concept, agg, numeric))
    for _, col, concept, agg, numeric in sorted(candidates, reverse=True):
        if concept in seen_concepts:
            continue
        value = numeric.mean() if agg == "mean" else numeric.sum(min_count=1)
        if pd.isna(value):
            continue
        selected.append({"label": col.replace("_", " "), "value": float(value), "aggregation": agg, "column": col, "concept": concept})
        seen_concepts.add(concept)
        if len(selected) >= 5:
            break

    count_specs = {
        "sales": ("Orders", ["order id", "order number"]),
        "hr": ("Headcount", ["employee id", "employee number"]),
        "it": ("Tickets", ["ticket id", "ticket number"]),
        "procurement": ("Purchase Orders", ["purchase order id", "po id", "order id"]),
        "finance": ("Transactions", ["transaction id", "invoice id", "journal id"]),
    }
    if domain in count_specs:
        label, aliases = count_specs[domain]
        id_col = next((c for c in df.columns if any(_qnorm(a) in _qnorm(c) for a in aliases)), None)
        if id_col is not None:
            distinct_count = int(df[id_col].nunique(dropna=True))
            selected.append({"label": label, "value": distinct_count, "aggregation": "nunique", "column": str(id_col), "concept": label.casefold().replace(" ", "_")})
            if domain == "sales" and distinct_count:
                revenue = next((m for m in selected if m["concept"] in {"net_revenue", "revenue"}), None)
                if revenue:
                    selected.append({"label": "Average Order Value", "value": float(revenue["value"]) / distinct_count, "aggregation": "derived", "column": "", "concept": "average_order_value"})
    return selected[:6]


def _aggregate_series(df: pd.DataFrame, metric: dict | None):
    if metric is None:
        return None, "Records"
    col, agg = metric["column"], metric["aggregation"]
    values = df[col] if agg == "nunique" else pd.to_numeric(df[col], errors="coerce")
    return values, metric["label"]


def _build_main_trend(df: pd.DataFrame, metric: dict | None, date_col: str | None) -> dict:
    if not date_col or metric is None:
        return {}
    dt = _parse_datetime_series(df[date_col], date_col)
    valid = dt.notna()
    if valid.sum() < 2:
        return {}
    values, label = _aggregate_series(df, metric)
    work = pd.DataFrame({"date": dt[valid], "value": values[valid]})
    years = int(work["date"].dt.year.nunique())
    if years >= 2:
        work["period"] = work["date"].dt.year.astype("Int64")
        grain = "year"
    else:
        work["period"] = work["date"].dt.to_period("M").astype(str)
        grain = "month"
    grouped = work.groupby("period", dropna=False)["value"].agg(metric["aggregation"]).dropna()
    if len(grouped) < 2:
        return {}
    grouped = grouped.tail(12)
    first, last = float(grouped.iloc[0]), float(grouped.iloc[-1])
    change_pct = ((last - first) / abs(first) * 100) if first else None
    points = [{"period": _fmt_label(idx), "value": float(value)} for idx, value in grouped.items()]
    return {
        "metric": label, "grain": grain, "points": points,
        "first_value": first, "last_value": last, "change": last - first,
        "change_pct": change_pct,
        "peak": {"period": _fmt_label(grouped.idxmax()), "value": float(grouped.max())},
        "low": {"period": _fmt_label(grouped.idxmin()), "value": float(grouped.min())},
    }


def _build_breakdowns(df: pd.DataFrame, metric: dict | None, dimensions: list[str]) -> dict:
    if metric is None:
        return {}
    values, label = _aggregate_series(df, metric)
    work = df.assign(__metric=values)
    out = {}
    for dim in dimensions[:2]:
        grouped = work.groupby(dim, dropna=False)["__metric"].agg(metric["aggregation"]).dropna().sort_values(ascending=False).head(5)
        if len(grouped) < 2:
            continue
        out[dim] = {
            "metric": label,
            "top": [{"name": _fmt_label(idx), "value": float(value)} for idx, value in grouped.items()],
        }
    return out


def _snapshot_observations(trend: dict, breakdowns: dict) -> list[dict]:
    observations = []
    if trend:
        observations.append({
            "type": "trend", "metric": trend["metric"], "change": trend["change"],
            "change_pct": trend["change_pct"], "from": trend["points"][0]["period"],
            "to": trend["points"][-1]["period"],
        })
    for dim, detail in breakdowns.items():
        if detail.get("top"):
            observations.append({"type": "top_contributor", "dimension": dim, "metric": detail["metric"], **detail["top"][0]})
    return observations[:3]


def _build_domain_snapshot(domain: str, schema: dict, request_cache: _RequestCache) -> dict:
    discovered = _discover_domains(schema)
    tables = discovered.get(domain, [])
    # Schema-only preselection avoids loading date/dimension tables that cannot
    # yield a meaningful KPI. Keep at most two strong analytical candidates.
    table_scores = []
    for table in tables:
        columns = schema.get(table) or []
        metric_count = sum(1 for col in columns if _column_metric_concept(col))
        identifier_count = sum(1 for col in columns if any(t in _qnorm(col) for t in ["order id", "employee id", "ticket id", "transaction id", "invoice id", "purchase order id"]))
        table_scores.append((metric_count * 4 + identifier_count * 2, table))
    analytical = [table for score, table in sorted(table_scores, reverse=True) if score > 0]
    tables = analytical[:2]
    best = None
    for table in tables:
        df = request_cache.load(table)
        metrics = _select_business_metrics(df, domain)
        dimensions = _select_dimensions(df)
        date_col = _select_date_column(df)
        score = len(metrics) * 4 + len(dimensions) + (2 if date_col else 0)
        candidate = (score, table, df, metrics, dimensions, date_col)
        if best is None or score > best[0]:
            best = candidate
    if best is None or not best[3]:
        return {}
    _, table, df, metrics, dimensions, date_col = best
    main_metric = metrics[0] if metrics else None
    trend = _build_main_trend(df, main_metric, date_col)
    breakdowns = _build_breakdowns(df, main_metric, dimensions)
    return {
        "domain": _domain_label(domain), "table": table,
        "kpis": [{k: v for k, v in metric.items() if k in {"label", "value", "aggregation"}} for metric in metrics],
        "trend": trend, "breakdowns": breakdowns,
        "observations": _snapshot_observations(trend, breakdowns),
    }


def _build_enterprise_snapshot(schema: dict, request_cache: _RequestCache, requested_domains: list[str] | None = None) -> dict:
    discovered = _discover_domains(schema)
    domains = requested_domains if requested_domains is not None else list(discovered)
    snapshots = {}
    for domain in domains:
        snapshot = _build_domain_snapshot(domain, schema, request_cache)
        if snapshot:
            snapshots[snapshot["domain"]] = snapshot
    return snapshots


def _build_overview_dataframe(snapshots: dict) -> pd.DataFrame:
    rows = []
    for domain, snapshot in snapshots.items():
        trend = snapshot.get("trend") or {}
        for metric in snapshot.get("kpis") or []:
            rows.append({
                "Domain": domain, "KPI": metric["label"], "Value": metric["value"],
                "Trend Change %": trend.get("change_pct") if trend.get("metric") == metric["label"] else None,
            })
    return pd.DataFrame(rows, columns=["Domain", "KPI", "Value", "Trend Change %"])

def _local_summary(question, result, plan):
    """Fast insight engine from verified dataframe; no second LLM call."""
    ar = bool(re.search(r"[\u0600-\u06FF]", str(question or "")))
    if result is None or result.empty:
        return "لا توجد بيانات مطابقة للسؤال." if ar else "No matching data was found."

    if len(result) == 1:
        col = result.columns[-1]
        val = _fmt_value(result.iloc[0, -1])
        return (
            f"**النتيجة:** {val}\n\n"
            f"القيمة محسوبة مباشرة من البيانات باستخدام الحقل **{col}**."
            if ar else
            f"**Result:** {val}\n\n"
            f"The value was calculated directly from the data using **{col}**."
        )

    value_col = result.columns[-1]
    label_col = result.columns[0]

    if not pd.api.types.is_numeric_dtype(result[value_col]):
        return (
            "تم تنفيذ التحليل وإرجاع النتائج الموثقة أدناه."
            if ar else
            "The analysis was completed and the verified results are shown below."
        )

    ranked = result.sort_values(value_col, ascending=False).reset_index(drop=True)
    hi = ranked.iloc[0]
    lo = ranked.iloc[-1]
    time_like = any(
        k in str(label_col).casefold()
        for k in ["year", "month", "quarter", "date", "period", "time"]
    )

    insights = []

    if time_like:
        ordered = result.reset_index(drop=True)
        first_val = float(ordered.iloc[0][value_col])
        last_val = float(ordered.iloc[-1][value_col])

        pct = None
        if first_val != 0:
            pct = ((last_val - first_val) / abs(first_val)) * 100

        # Period-over-period changes
        changes = []
        for i in range(1, len(ordered)):
            prev = float(ordered.iloc[i - 1][value_col])
            cur = float(ordered.iloc[i][value_col])
            delta = cur - prev
            pct_delta = (delta / abs(prev) * 100) if prev != 0 else None
            changes.append(
                {
                    "from": ordered.iloc[i - 1][label_col],
                    "to": ordered.iloc[i][label_col],
                    "delta": delta,
                    "pct": pct_delta,
                }
            )

        biggest_drop = min(changes, key=lambda x: x["delta"]) if changes else None
        biggest_rise = max(changes, key=lambda x: x["delta"]) if changes else None

        if ar:
            if pct is not None:
                direction = "انخفاض" if pct < 0 else "ارتفاع" if pct > 0 else "استقرار"
                insights.append(
                    f"**الاتجاه العام:** {direction} بنسبة **{abs(pct):.1f}%** "
                    f"من **{_fmt_value(first_val)}** إلى **{_fmt_value(last_val)}**."
                )
            insights.append(
                f"**أعلى قيمة:** **{_fmt_value(hi[value_col])}** في **{_fmt_label(hi[label_col])}**، "
                f"وأقل قيمة **{_fmt_value(lo[value_col])}** في **{_fmt_label(lo[label_col])}**."
            )
            if biggest_drop and biggest_drop["delta"] < 0:
                pct_txt = (
                    f" (**{abs(biggest_drop['pct']):.1f}%**)"
                    if biggest_drop["pct"] is not None else ""
                )
                insights.append(
                    f"**أكبر تراجع بين فترتين:** من **{_fmt_label(biggest_drop['from'])}** إلى "
                    f"**{_fmt_label(biggest_drop['to'])}** بمقدار **{_fmt_value(abs(biggest_drop['delta']))}**{pct_txt}."
                )
            if biggest_rise and biggest_rise["delta"] > 0:
                pct_txt = (
                    f" (**{abs(biggest_rise['pct']):.1f}%**)"
                    if biggest_rise["pct"] is not None else ""
                )
                insights.append(
                    f"**أكبر تحسن بين فترتين:** من **{_fmt_label(biggest_rise['from'])}** إلى "
                    f"**{_fmt_label(biggest_rise['to'])}** بمقدار **{_fmt_value(biggest_rise['delta'])}**{pct_txt}."
                )
            if pct is not None and pct <= -20:
                insights.append(
                    "**ملاحظة:** الاتجاه يستحق المتابعة؛ الانخفاض الكلي يتجاوز 20% خلال الفترة المعروضة."
                )
        else:
            if pct is not None:
                direction = "decrease" if pct < 0 else "increase" if pct > 0 else "stable movement"
                insights.append(
                    f"**Overall trend:** {direction} of **{abs(pct):.1f}%**, "
                    f"from **{_fmt_value(first_val)}** to **{_fmt_value(last_val)}**."
                )
            insights.append(
                f"**Peak / low:** **{_fmt_label(hi[label_col])}** is highest at **{_fmt_value(hi[value_col])}**; "
                f"**{_fmt_label(lo[label_col])}** is lowest at **{_fmt_value(lo[value_col])}**."
            )
            if biggest_drop and biggest_drop["delta"] < 0:
                insights.append(
                    f"**Largest period drop:** {biggest_drop['from']} → {biggest_drop['to']}, "
                    f"down **{_fmt_value(abs(biggest_drop['delta']))}**"
                    + (f" (**{abs(biggest_drop['pct']):.1f}%**)." if biggest_drop["pct"] is not None else ".")
                )
            if biggest_rise and biggest_rise["delta"] > 0:
                insights.append(
                    f"**Largest period improvement:** {biggest_rise['from']} → {biggest_rise['to']}, "
                    f"up **{_fmt_value(biggest_rise['delta'])}**"
                    + (f" (**{abs(biggest_rise['pct']):.1f}%**)." if biggest_rise["pct"] is not None else ".")
                )
            if pct is not None and pct <= -20:
                insights.append(
                    "**Watchpoint:** the cumulative decline exceeds 20% across the displayed period."
                )

    else:
        # Category comparison insights
        total = float(pd.to_numeric(result[value_col], errors="coerce").sum())
        hi_val = float(hi[value_col])
        share = (hi_val / total * 100) if total else None

        if ar:
            insights.append(
                f"**الأعلى:** **{_fmt_label(hi[label_col])}** بقيمة **{_fmt_value(hi[value_col])}**."
            )
            insights.append(
                f"**الأقل:** **{_fmt_label(lo[label_col])}** بقيمة **{_fmt_value(lo[value_col])}**."
            )
            if share is not None:
                insights.append(
                    f"**حصة الأعلى:** تمثل **{_fmt_label(hi[label_col])}** حوالي **{share:.1f}%** من إجمالي النتائج المعروضة."
                )
            if len(result) >= 3:
                top3 = ranked.head(3)[value_col].sum()
                if total:
                    insights.append(
                        f"**التركيز:** أعلى 3 فئات تمثل **{(float(top3)/total*100):.1f}%** من الإجمالي."
                    )
        else:
            insights.append(
                f"**Highest:** **{_fmt_label(hi[label_col])}** at **{_fmt_value(hi[value_col])}**."
            )
            insights.append(
                f"**Lowest:** **{_fmt_label(lo[label_col])}** at **{_fmt_value(lo[value_col])}**."
            )
            if share is not None:
                insights.append(
                    f"**Top share:** {hi[label_col]} accounts for **{share:.1f}%** of the displayed total."
                )
            if len(result) >= 3:
                top3 = ranked.head(3)[value_col].sum()
                if total:
                    insights.append(
                        f"**Concentration:** the top 3 categories represent **{(float(top3)/total*100):.1f}%** of the total."
                    )

    heading = "**أبرز الاستنتاجات:**" if ar else "**Key insights:**"
    return heading + "\n\n" + "\n".join(f"- {x}" for x in insights[:5])



def _needs_deep_insight(question: str, plan: dict) -> bool:
    if str(plan.get("analysis_depth", "")).casefold() == "deep":
        return True
    q = _qnorm(question)
    tokens = [
        "why","cause","causes","reason","analyze","analysis","recommend",
        "risk","opportunity","driver","root cause",
        "لماذا","سبب","اسباب","أسباب","حلل","تحليل","فسر","اشرح",
        "توصية","توصيات","مخاطر","فرص","محركات"
    ]
    return any(_qnorm(t) in q for t in tokens)

def _executive_insight(question: str, result: pd.DataFrame, plan: dict, driver_evidence: dict | None = None) -> str:
    """Use Copilot only when the user asks for deeper interpretation."""
    rows = result.head(20).where(pd.notna(result), None).to_dict("records")
    prompt = f"""{SYSTEM}
You are the EXECUTIVE INSIGHT layer.
The application has already performed the verified calculation.

User question:
{question}

Verified analysis plan:
{json.dumps(plan, ensure_ascii=False, default=str)}

Verified result rows:
{json.dumps(rows, ensure_ascii=False, default=str)}

Verified driver/contribution evidence (may be empty):
{json.dumps(driver_evidence or {}, ensure_ascii=False, default=str)}

Write a concise executive answer in the SAME LANGUAGE as the user's question.

Rules:
- Answer the actual question first.
- Then provide 3 to 5 meaningful insights, not a description of columns.
- For WHY/root-cause questions, use the verified driver evidence to identify which region/product/category contributed most to the latest decline. Quantify contribution when available.
- Distinguish a DATA DRIVER from a proven BUSINESS CAUSE. Never claim causation unless the data proves it.
- If the user requests recommendations, give exactly 3 practical recommendations tied to the verified drivers.
- Explain trends, concentration, turning points, anomalies or operational implications only when supported.
- If evidence is insufficient, explicitly say what additional field/dimension is required to validate the cause.
- Never invent a cause, recommendation, number or business fact not supported by the verified result.
- Do not render a markdown table.
- Keep it decision-oriented and suitable for management.
"""
    return _ask_sync(prompt)


def _compact_snapshots(snapshots: dict) -> dict:
    """Bound synthesis payload size without removing verified executive evidence."""
    compact = {}
    for domain, snapshot in snapshots.items():
        trend = dict(snapshot.get("trend") or {})
        trend["points"] = (trend.get("points") or [])[-8:]
        compact[domain] = {
            "kpis": (snapshot.get("kpis") or [])[:6],
            "trend": trend,
            "breakdowns": {
                dim: {"metric": detail.get("metric"), "top": (detail.get("top") or [])[:5]}
                for dim, detail in list((snapshot.get("breakdowns") or {}).items())[:2]
            },
            "observations": (snapshot.get("observations") or [])[:3],
        }
    return compact


def _synthesize_overview(question: str, plan: dict, snapshots: dict) -> str:
    prompt = f"""{SYSTEM}
You are the EXECUTIVE SYNTHESIS layer. The application has already calculated the compact verified snapshot below.

User question:
{question}

Intent and scope:
{json.dumps({k: plan.get(k) for k in ['intent','scope','domains']}, ensure_ascii=False)}

Compact verified snapshot:
{json.dumps(_compact_snapshots(snapshots), ensure_ascii=False, default=str)}

Write a concise executive answer in the SAME LANGUAGE as the user's question.
Use only sections supported by the snapshot: Executive Summary, Key KPIs, Performance Trend, Key Drivers, Risks/Anomalies, Management Attention, and Recommendations.
For an enterprise overview, summarize each available domain first, then give only verified enterprise-level observations.
For a multi-domain overview, discuss only the supplied requested domains.
Never calculate, infer, or invent a number. Never invent a cross-domain relationship.
Do not describe a data driver as a proven cause. If a possible relationship needs validation, say so explicitly.
Do not render a markdown table. Keep recommendations practical and tied to verified observations.
"""
    return _ask_sync(prompt)


def _local_overview_summary(question: str, snapshots: dict) -> str:
    ar = bool(re.search(r"[\u0600-\u06FF]", str(question or "")))
    if not snapshots:
        return "لا تتوفر بيانات أعمال كافية لبناء الملخص المطلوب." if ar else "There is not enough business data to build the requested overview."
    lines = ["**الملخص التنفيذي**" if ar else "**Executive summary**"]
    for domain, snapshot in snapshots.items():
        lines.append(f"\n**{domain}**")
        for metric in (snapshot.get("kpis") or [])[:4]:
            lines.append(f"- {metric['label']}: **{_fmt_value(metric['value'])}**")
        trend = snapshot.get("trend") or {}
        if trend and trend.get("change_pct") is not None:
            direction = ("ارتفع" if trend["change_pct"] > 0 else "انخفض") if ar else ("increased" if trend["change_pct"] > 0 else "decreased")
            lines.append(
                f"- {trend['metric']} {direction} **{abs(trend['change_pct']):.1f}%** "
                f"({trend['points'][0]['period']} → {trend['points'][-1]['period']})."
            )
    return "\n".join(lines)


def _prepare_simple_plan(plan: dict, schema: dict) -> dict:
    prepared = dict(plan)
    if prepared.get("table") not in schema:
        domains = _resolve_requested_domains(prepared, schema)
        discovered = _discover_domains(schema)
        if domains:
            tables = discovered.get(domains[0]) or []
            if tables:
                prepared["table"] = tables[0]
    if prepared.get("intent") in {"trend", "root_cause"} and not prepared.get("date_column"):
        columns = schema.get(prepared.get("table")) or []
        prepared["date_column"] = next(
            (c for c in columns if any(t in _qnorm(c).split() for t in ["date", "year", "month", "period"])),
            None,
        )
        if prepared.get("date_column") and not prepared.get("date_grain"):
            prepared["date_grain"] = "year"
    return prepared


def _execute_plan(plan: dict, question: str, schema: dict, excel_mcp, request_cache: _RequestCache):
    """Route intent to deterministic execution; overview intents never use table=None in _execute."""
    intent = plan.get("intent")
    if intent in _OVERVIEW_INTENTS:
        if intent == "enterprise_overview":
            domains = None
        else:
            domains = _resolve_requested_domains(plan, schema)
            if not domains:
                raise ValueError("The requested business domain could not be resolved from the available workbook schema.")
            if intent == "domain_overview":
                domains = domains[:1]
        snapshots = _build_enterprise_snapshot(schema, request_cache, domains)
        if not snapshots:
            raise ValueError("No meaningful business metrics are available for the requested domain scope.")
        return _build_overview_dataframe(snapshots), {}, snapshots

    prepared = _prepare_simple_plan(plan, schema)
    if not prepared.get("table"):
        raise ValueError("The analytical data source could not be resolved from the question and available schema.")
    for key in ["table", "date_column", "date_grain"]:
        if prepared.get(key) is not None:
            plan[key] = prepared[key]
    prepared = _resolve_metric(prepared, request_cache.load(prepared["table"]))
    for key in ["metric_column", "metric_semantic", "aggregation"]:
        if prepared.get(key) is not None:
            plan[key] = prepared[key]
    result = _execute(prepared, excel_mcp, request_cache)
    evidence = {}
    if intent == "root_cause":
        evidence = _driver_evidence(question, prepared, excel_mcp, request_cache)
    return result, evidence, {}


def _safe_failure_message(question: str, exc: Exception) -> str:
    ar = bool(re.search(r"[\u0600-\u06FF]", str(question or "")))
    detail = str(exc)
    if ar:
        if "domain" in detail.casefold() or "source" in detail.casefold():
            return "تعذر تحديد نطاق الأعمال المطلوب من البيانات المتاحة. يرجى تحديد المجال، مثل المبيعات أو المالية أو الموارد البشرية."
        if "metric" in detail.casefold():
            return "المؤشر المطلوب غير متوفر في البيانات المحددة. يرجى اختيار مؤشر موجود أو طلب نظرة عامة على المجال."
        return "تعذر إكمال التحليل من البيانات المتاحة. يرجى توضيح المجال أو المؤشر المطلوب."
    if "domain" in detail.casefold() or "source" in detail.casefold():
        return "The requested business scope could not be resolved from the available data. Specify a domain such as Sales, Finance, or HR."
    if "metric" in detail.casefold():
        return "The requested metric is not available in the selected data. Choose an available metric or request a domain overview."
    return "The analysis could not be completed from the available data. Clarify the business domain or metric."

def ask(question, history, excel_mcp):
    total_started = time.perf_counter()
    timings = {"planning_ms": 0.0, "data_loading_ms": 0.0, "execution_ms": 0.0, "synthesis_ms": 0.0, "total_ms": 0.0}
    request_cache = _RequestCache(excel_mcp)
    schema = {}
    plan: dict[str, Any] = {"mode": "analysis", "intent": "single_kpi"}
    result = None
    evidence = {}
    try:
        schema = _schema(excel_mcp, request_cache=request_cache)
        started = time.perf_counter()
        plan = _sanitize_plan(question, _build_plan(question, schema, history))
        timings["planning_ms"] = (time.perf_counter() - started) * 1000

        clarification = plan.get("clarification")
        if clarification:
            return {"answer": str(clarification), "data": None, "plan": plan, "driver_evidence": {}}

        if plan.get("mode") == "chat":
            started = time.perf_counter()
            p = f"""{SYSTEM}
Recent conversation: {json.dumps((history or [])[-6:],ensure_ascii=False)}
User: {question}
Answer directly. Do not mention implementation details, demo mode, or GitHub Models."""
            answer = _ask_sync(p)
            timings["synthesis_ms"] = (time.perf_counter() - started) * 1000
            return {"answer": answer, "data": None, "plan": plan, "driver_evidence": {}}

        loading_before_execution = request_cache.data_loading_ms
        started = time.perf_counter()
        result, evidence, snapshots = _execute_plan(plan, question, schema, excel_mcp, request_cache)
        routed_ms = (time.perf_counter() - started) * 1000
        loading_during_execution = request_cache.data_loading_ms - loading_before_execution
        timings["execution_ms"] = max(0.0, routed_ms - loading_during_execution)

        started = time.perf_counter()
        if plan.get("intent") in _OVERVIEW_INTENTS:
            try:
                answer = _synthesize_overview(question, plan, snapshots)
            except Exception:
                _LOGGER.exception("NovaCore overview synthesis failed; using verified local fallback")
                answer = _local_overview_summary(question, snapshots)
        elif plan.get("intent") == "root_cause":
            try:
                answer = _executive_insight(question, result, plan, evidence)
            except Exception:
                _LOGGER.exception("NovaCore root-cause synthesis failed; using verified local fallback")
                answer = _local_summary(question, result, plan)
        else:
            answer = _local_summary(question, result, plan)
        timings["synthesis_ms"] = (time.perf_counter() - started) * 1000
        return {"answer": answer, "data": result, "plan": plan, "driver_evidence": evidence}
    except Exception as exc:
        _LOGGER.exception("NovaCore analytical request failed")
        return {"answer": _safe_failure_message(question, exc), "data": result, "plan": plan, "driver_evidence": evidence}
    finally:
        timings["data_loading_ms"] = request_cache.data_loading_ms
        timings["total_ms"] = (time.perf_counter() - total_started) * 1000
        _LOGGER.info(
            "NovaCore Performance: planning_ms=%.1f data_loading_ms=%.1f execution_ms=%.1f synthesis_ms=%.1f total_ms=%.1f intent=%s",
            timings["planning_ms"], timings["data_loading_ms"], timings["execution_ms"],
            timings["synthesis_ms"], timings["total_ms"], plan.get("intent"),
        )
