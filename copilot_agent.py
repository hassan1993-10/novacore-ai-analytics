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
STRICT SEMANTIC RULES:
- group_by must contain ONLY dimensions explicitly requested by the user.
- Never add unrelated dimensions.
- "by year / حسب السنة / حسب السنوات" means aggregate ONLY by year unless another dimension is explicitly requested.
- Filters are not grouping dimensions.
- Set analysis_depth="deep" for why/causes/analyze/recommendations/risks/opportunities/root-cause questions.
Use only exact schema names. For year/month/quarter set date_column + date_grain.
mode=chat for greetings, identity, help, or non-data questions. Never output Python."""
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
    q = _qnorm(question)
    return any(_qnorm(t) in q for t in tokens)

_DIMENSION_HINTS = {
    "region": ["region", "regions", "منطقة", "المناطق", "المنطقة"],
    "status": ["status", "order status", "حالة", "الحالة", "حالات"],
    "product": ["product", "products", "منتج", "المنتج", "المنتجات"],
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
    groups = list(plan.get("group_by") or [])

    # Deterministic business-term resolver: do not let the LLM switch measures
    # between equivalent questions. Exact schema names are resolved later.
    qn = _qnorm(question)
    metric_aliases = {
        "net_revenue": ["صافي الايرادات", "صافي الإيرادات", "net revenue"],
        "revenue": ["الايرادات", "الإيرادات", "revenue"],
        "gross_profit": ["اجمالي الربح", "إجمالي الربح", "gross profit"],
        "amount": ["amount", "المبلغ", "القيمة"],
        "quantity": ["quantity", "الكمية"],
    }
    requested_metric = next((k for k,v in metric_aliases.items() if any(_qnorm(x) in qn for x in v)), None)
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

    # Deep analysis signal can be corrected deterministically.
    deep_tokens = [
        "why","cause","causes","reason","reasons","analyze","analyse","analysis",
        "recommend","recommendation","recommendations","risk","risks","opportunity",
        "opportunities","driver","drivers","root cause",
        "لماذا","سبب","اسباب","أسباب","حلل","تحليل","فسر","اشرح","توصية","توصيات",
        "مخاطر","فرص","محركات","جذر"
    ]
    if _mentions_any(q, deep_tokens):
        plan["analysis_depth"] = "deep"
    else:
        plan.setdefault("analysis_depth", "standard")
    return plan

def _resolve_metric(plan: dict, df: pd.DataFrame) -> dict:
    """Resolve business metric semantics against real columns, preferring exact matches."""
    plan = dict(plan)
    semantic = plan.get("metric_semantic")
    aliases = {
        "net_revenue": ["Net_Revenue_SAR", "Net Revenue SAR", "Net_Revenue", "Net Revenue"],
        "revenue": ["Revenue_SAR", "Revenue SAR", "Revenue"],
        "gross_profit": ["Gross_Profit_SAR", "Gross Profit SAR", "Gross_Profit", "Gross Profit"],
        "amount": ["Amount_SAR", "Amount SAR", "Amount"],
        "quantity": ["Quantity"],
    }
    if semantic:
        norm_cols = {_qnorm(c): c for c in df.columns}
        for candidate in aliases.get(semantic, []):
            if _qnorm(candidate) in norm_cols:
                plan["metric_column"] = norm_cols[_qnorm(candidate)]
                break
    return plan

def _driver_evidence(question: str, plan: dict, excel_mcp) -> dict:
    """Compute verified contribution-to-change evidence for deep/root-cause questions.
    This identifies data drivers (not causal proof) across explicitly requested dimensions.
    """
    table = plan.get("table")
    if not table:
        return {}
    df = excel_mcp.load_table(DATASET, table).copy()
    plan = _resolve_metric(plan, df)
    for f in plan.get("filters") or []:
        df = _apply_filter(df, f)
    metric = plan.get("metric_column")
    if metric not in df.columns:
        return {}
    df[metric] = pd.to_numeric(df[metric], errors="coerce")

    dc = plan.get("date_column")
    if dc not in df.columns:
        return {}
    dt = pd.to_datetime(df[dc], errors="coerce")
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

    evidence={"period":{"from":prev_y,"to":curr_y,"from_value":prev_total,"to_value":curr_total,"delta":total_delta},"drivers":{}}
    for family in requested[:3]:
        candidates=[c for c in df.columns if _column_semantic_family(c)==family]
        if not candidates: continue
        col=candidates[0]
        g=(df[df["__Year"].isin([prev_y,curr_y])].groupby([col,"__Year"],dropna=False)[metric].sum().unstack(fill_value=0))
        if prev_y not in g.columns: g[prev_y]=0
        if curr_y not in g.columns: g[curr_y]=0
        g["delta"]=g[curr_y]-g[prev_y]
        g["contribution_pct"]=(g["delta"]/abs(total_delta)*100) if total_delta else 0
        rows=[]
        for idx,row in g.sort_values("delta").head(8).iterrows():
            rows.append({"name":str(idx),"from":float(row[prev_y]),"to":float(row[curr_y]),"delta":float(row["delta"]),"contribution_pct":float(row["contribution_pct"])})
        evidence["drivers"][family]={"column":col,"rows":rows}
    return evidence

def _execute(plan, excel_mcp):
    table = plan.get("table")
    if not table:
        raise ValueError("No table selected.")

    df = excel_mcp.load_table(DATASET, table).copy()
    plan = _resolve_metric(plan, df)
    for f in plan.get("filters") or []:
        df = _apply_filter(df, f)

    groups = [g for g in (plan.get("group_by") or []) if g in df.columns]
    dc, grain = plan.get("date_column"), plan.get("date_grain")
    time_group = None

    if dc in df.columns and grain in {"year","quarter","month"}:
        dt = pd.to_datetime(df[dc], errors="coerce")
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
            raise ValueError(f"Metric column '{metric}' was not found.")
        if agg in {"sum","mean","min","max"}:
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

def ask(question, history, excel_mcp):
    schema = _schema(excel_mcp)
    plan = _build_plan(question, schema, history)
    plan = _sanitize_plan(question, plan)

    if plan.get("mode") == "chat":
        p = f"""{SYSTEM}
Recent conversation: {json.dumps((history or [])[-6:],ensure_ascii=False)}
User: {question}
Answer directly. Do not mention implementation details, demo mode, or GitHub Models."""
        return {"answer":_ask_sync(p),"data":None,"plan":plan}

    result = _execute(plan, excel_mcp)

    # Hybrid mode:
    # standard questions stay fast; deep analytical questions get one verified executive-insight call.
    if _needs_deep_insight(question, plan):
        try:
            evidence = _driver_evidence(question, plan, excel_mcp)
            answer = _executive_insight(question, result, plan, evidence)
        except Exception:
            answer = _local_summary(question, result, plan)
    else:
        answer = _local_summary(question, result, plan)

    return {"answer":answer, "data":result, "plan":plan}
