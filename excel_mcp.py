"""
NovaCore Excel MCP
==================
Generic MCP server for Excel analytics.

Architecture:
    LLM / Agent
        -> MCP tools
        -> Pandas
        -> Excel workbooks in ./data

Each Excel file is treated as a dataset.
Each worksheet is treated as a table.

The module is intentionally read-only.
"""

from __future__ import annotations

import json
import math
import re
from functools import lru_cache
from pathlib import Path
from typing import Any, Optional

import pandas as pd
from mcp.server.fastmcp import FastMCP


# =========================================================
# Configuration
# =========================================================

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"

mcp = FastMCP("novacore-excel-mcp")


# =========================================================
# Helpers
# =========================================================

def json_safe(value: Any) -> Any:
    """Convert pandas/numpy/date values into JSON-safe Python values."""
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    if isinstance(value, (pd.Timestamp,)):
        return value.isoformat()

    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass

    return value


def records(df: pd.DataFrame) -> list[dict]:
    if df.empty:
        return []

    output = []
    for row in df.to_dict(orient="records"):
        output.append({str(k): json_safe(v) for k, v in row.items()})
    return output


def success_json(data: Any, row_count: Optional[int] = None, **extra: Any) -> str:
    payload = {
        "success": True,
        "data": data,
    }
    if row_count is not None:
        payload["row_count"] = int(row_count)

    payload.update(extra)
    return json.dumps(payload, ensure_ascii=False, indent=2, default=str)


def error_json(error: Any, error_type: str = "excel_error", **extra: Any) -> str:
    payload = {
        "success": False,
        "error_type": error_type,
        "message": str(error),
    }
    payload.update(extra)
    return json.dumps(payload, ensure_ascii=False, indent=2, default=str)


def normalize_name(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def dataset_path(dataset_name: str) -> Path:
    requested = normalize_name(dataset_name)

    for path in discover_datasets():
        if normalize_name(path.name) == requested:
            return path
        if normalize_name(path.stem) == requested:
            return path

    raise ValueError(f"Dataset not found: {dataset_name}")


@lru_cache(maxsize=1)
def discover_datasets() -> tuple[Path, ...]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    files = []
    for pattern in ("*.xlsx", "*.xlsm", "*.xls"):
        files.extend(DATA_DIR.glob(pattern))

    files = [
        path for path in files
        if not path.name.startswith("~$")
    ]

    return tuple(sorted(files, key=lambda p: p.name.casefold()))


@lru_cache(maxsize=64)
def workbook_sheet_names(dataset_name: str) -> tuple[str, ...]:
    path = dataset_path(dataset_name)
    xls = pd.ExcelFile(path)
    return tuple(xls.sheet_names)


def resolve_sheet(dataset_name: str, table_name: str) -> str:
    requested = normalize_name(table_name)
    sheets = workbook_sheet_names(dataset_name)

    for sheet in sheets:
        if normalize_name(sheet) == requested:
            return sheet

    # Light fuzzy/contains matching
    candidates = [
        sheet for sheet in sheets
        if requested in normalize_name(sheet) or normalize_name(sheet) in requested
    ]

    if len(candidates) == 1:
        return candidates[0]

    raise ValueError(
        f"Table/sheet not found: {table_name}. "
        f"Available sheets: {', '.join(sheets)}"
    )


@lru_cache(maxsize=128)
def load_table(dataset_name: str, table_name: str) -> pd.DataFrame:
    path = dataset_path(dataset_name)
    sheet = resolve_sheet(dataset_name, table_name)

    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    return df


def column_profile(series: pd.Series) -> dict:
    non_null = series.dropna()

    result = {
        "column_name": str(series.name),
        "dtype": str(series.dtype),
        "row_count": int(len(series)),
        "null_count": int(series.isna().sum()),
        "distinct_count": int(non_null.nunique(dropna=True)),
    }

    if not non_null.empty:
        result["sample_values"] = [
            json_safe(v) for v in non_null.drop_duplicates().head(5).tolist()
        ]

    if pd.api.types.is_numeric_dtype(series):
        numeric = pd.to_numeric(series, errors="coerce").dropna()
        if not numeric.empty:
            result.update({
                "min": json_safe(numeric.min()),
                "max": json_safe(numeric.max()),
                "mean": json_safe(numeric.mean()),
                "sum": json_safe(numeric.sum()),
            })

    elif pd.api.types.is_datetime64_any_dtype(series):
        dates = pd.to_datetime(series, errors="coerce").dropna()
        if not dates.empty:
            result.update({
                "min": dates.min().isoformat(),
                "max": dates.max().isoformat(),
            })

    return result


def infer_key_candidates(df: pd.DataFrame) -> list[str]:
    result = []

    for col in df.columns:
        name = str(col)
        non_null = df[col].dropna()

        if non_null.empty:
            continue

        uniqueness = non_null.nunique(dropna=True) / max(len(non_null), 1)

        if (
            name.casefold().endswith("_id")
            or name.casefold() == "id"
            or name.casefold().endswith("_key")
        ):
            if uniqueness >= 0.90:
                result.append(name)

    return result


def infer_relationships_internal(dataset_name: str) -> list[dict]:
    """
    Infer simple relationships using matching *_ID / *_Key column names.
    This is heuristic metadata, not a database-enforced relationship.
    """
    sheets = workbook_sheet_names(dataset_name)
    table_columns: dict[str, set[str]] = {}
    key_candidates: dict[str, set[str]] = {}

    for sheet in sheets:
        df = load_table(dataset_name, sheet)
        table_columns[sheet] = set(map(str, df.columns))
        key_candidates[sheet] = set(infer_key_candidates(df))

    rels: list[dict] = []
    seen = set()

    for parent in sheets:
        for key in key_candidates[parent]:
            for child in sheets:
                if child == parent:
                    continue

                if key in table_columns[child]:
                    sig = (parent, key, child, key)
                    if sig in seen:
                        continue
                    seen.add(sig)

                    rels.append({
                        "parent_table": parent,
                        "parent_column": key,
                        "child_table": child,
                        "child_column": key,
                        "relationship_type": "inferred",
                    })

    return rels


def apply_filters(df: pd.DataFrame, filters: list[dict]) -> pd.DataFrame:
    result = df.copy()

    for item in filters or []:
        column = str(item.get("column", "")).strip()
        operator = str(item.get("operator", "eq")).strip().casefold()
        value = item.get("value")

        if column not in result.columns:
            raise ValueError(f"Filter column not found: {column}")

        series = result[column]

        if operator == "eq":
            result = result[series == value]
        elif operator == "ne":
            result = result[series != value]
        elif operator == "gt":
            result = result[series > value]
        elif operator == "gte":
            result = result[series >= value]
        elif operator == "lt":
            result = result[series < value]
        elif operator == "lte":
            result = result[series <= value]
        elif operator == "contains":
            result = result[
                series.astype(str).str.contains(str(value), case=False, na=False, regex=False)
            ]
        elif operator == "in":
            values = value if isinstance(value, list) else [value]
            result = result[series.isin(values)]
        elif operator == "not_in":
            values = value if isinstance(value, list) else [value]
            result = result[~series.isin(values)]
        elif operator == "is_blank":
            result = result[series.isna()]
        elif operator == "not_blank":
            result = result[series.notna()]
        else:
            raise ValueError(
                f"Unsupported filter operator: {operator}. "
                "Supported: eq, ne, gt, gte, lt, lte, contains, in, not_in, is_blank, not_blank"
            )

    return result


def aggregate_series(series: pd.Series, operation: str) -> Any:
    operation = operation.casefold()

    if operation == "sum":
        return json_safe(pd.to_numeric(series, errors="coerce").sum())
    if operation == "mean" or operation == "avg":
        return json_safe(pd.to_numeric(series, errors="coerce").mean())
    if operation == "min":
        return json_safe(series.min())
    if operation == "max":
        return json_safe(series.max())
    if operation == "count":
        return int(series.count())
    if operation in ("nunique", "distinct_count"):
        return int(series.nunique(dropna=True))

    raise ValueError(f"Unsupported aggregation: {operation}")


# =========================================================
# MCP Tools - Discovery
# =========================================================

@mcp.tool()
def list_datasets() -> str:
    """Return Excel datasets available inside the ./data folder."""
    try:
        data = []

        for path in discover_datasets():
            data.append({
                "dataset_id": path.stem,
                "dataset_name": path.name,
                "file_name": path.name,
                "size_bytes": path.stat().st_size,
            })

        return success_json(data, row_count=len(data))

    except Exception as exc:
        return error_json(exc, "dataset_discovery_error")


@mcp.tool()
def get_model_overview(dataset_name: str) -> str:
    """Return workbook sheets, row counts, column counts and inferred relationships."""
    try:
        sheets = workbook_sheet_names(dataset_name)

        tables = []
        for sheet in sheets:
            df = load_table(dataset_name, sheet)
            tables.append({
                "table_name": sheet,
                "row_count": int(len(df)),
                "column_count": int(len(df.columns)),
            })

        rels = infer_relationships_internal(dataset_name)

        return success_json(
            {
                "dataset": dataset_path(dataset_name).name,
                "table_count": len(tables),
                "tables": tables,
                "relationship_count": len(rels),
                "relationships": rels,
            },
            dataset_name=dataset_name,
        )

    except Exception as exc:
        return error_json(exc, "model_overview_error", dataset_name=dataset_name)


@mcp.tool()
def get_tables(dataset_name: str, search_text: str = "") -> str:
    """Return worksheet/table names for one Excel dataset."""
    try:
        names = list(workbook_sheet_names(dataset_name))

        if search_text.strip():
            q = normalize_name(search_text)
            names = [name for name in names if q in normalize_name(name)]

        data = [{"table_name": name} for name in names]
        return success_json(data, row_count=len(data), dataset_name=dataset_name)

    except Exception as exc:
        return error_json(exc, "table_discovery_error", dataset_name=dataset_name)


@mcp.tool()
def get_columns(
    dataset_name: str,
    table_name: str,
    search_text: str = "",
) -> str:
    """Return detailed column profiles for one worksheet."""
    try:
        table = resolve_sheet(dataset_name, table_name)
        df = load_table(dataset_name, table)

        columns = []
        for col in df.columns:
            if search_text.strip() and normalize_name(search_text) not in normalize_name(col):
                continue
            columns.append(column_profile(df[col]))

        return success_json(
            columns,
            row_count=len(columns),
            dataset_name=dataset_name,
            table_name=table,
        )

    except Exception as exc:
        return error_json(
            exc,
            "column_discovery_error",
            dataset_name=dataset_name,
            table_name=table_name,
        )


@mcp.tool()
def inspect_table(
    dataset_name: str,
    table_name: str,
    max_rows: int = 5,
) -> str:
    """Return columns plus a small row sample for one worksheet."""
    try:
        table = resolve_sheet(dataset_name, table_name)
        df = load_table(dataset_name, table)

        return success_json(
            {
                "table_name": table,
                "row_count": int(len(df)),
                "column_count": int(len(df.columns)),
                "key_candidates": infer_key_candidates(df),
                "columns": [column_profile(df[col]) for col in df.columns],
                "sample_rows": records(df.head(max(1, min(int(max_rows), 25)))),
            },
            dataset_name=dataset_name,
            table_name=table,
        )

    except Exception as exc:
        return error_json(
            exc,
            "inspect_table_error",
            dataset_name=dataset_name,
            table_name=table_name,
        )


@mcp.tool()
def get_distinct_values(
    dataset_name: str,
    table_name: str,
    column_name: str,
    max_values: int = 100,
) -> str:
    """Return distinct nonblank values for a confirmed table column."""
    try:
        table = resolve_sheet(dataset_name, table_name)
        df = load_table(dataset_name, table)

        if column_name not in df.columns:
            raise ValueError(f"Column not found: {table}[{column_name}]")

        values = (
            df[column_name]
            .dropna()
            .drop_duplicates()
            .head(max(1, min(int(max_values), 1000)))
            .tolist()
        )

        values = [json_safe(v) for v in values]

        return success_json(
            values,
            row_count=len(values),
            dataset_name=dataset_name,
            table_name=table,
            column_name=column_name,
        )

    except Exception as exc:
        return error_json(
            exc,
            "distinct_values_error",
            dataset_name=dataset_name,
            table_name=table_name,
            column_name=column_name,
        )


@mcp.tool()
def get_value_distribution(
    dataset_name: str,
    table_name: str,
    column_name: str,
    max_values: int = 50,
) -> str:
    """Return most common values and counts for a column."""
    try:
        table = resolve_sheet(dataset_name, table_name)
        df = load_table(dataset_name, table)

        if column_name not in df.columns:
            raise ValueError(f"Column not found: {table}[{column_name}]")

        counts = (
            df[column_name]
            .fillna("<BLANK>")
            .value_counts(dropna=False)
            .head(max(1, min(int(max_values), 500)))
            .rename_axis(column_name)
            .reset_index(name="row_count")
        )

        return success_json(
            records(counts),
            row_count=len(counts),
            dataset_name=dataset_name,
            table_name=table,
            column_name=column_name,
        )

    except Exception as exc:
        return error_json(
            exc,
            "value_distribution_error",
            dataset_name=dataset_name,
            table_name=table_name,
            column_name=column_name,
        )


@mcp.tool()
def get_relationships(dataset_name: str) -> str:
    """Return simple inferred relationships based on shared ID/key columns."""
    try:
        rels = infer_relationships_internal(dataset_name)
        return success_json(
            rels,
            row_count=len(rels),
            dataset_name=dataset_name,
        )

    except Exception as exc:
        return error_json(
            exc,
            "relationship_discovery_error",
            dataset_name=dataset_name,
        )


# =========================================================
# MCP Tools - Safe Analytics
# =========================================================

@mcp.tool()
def run_analysis(
    dataset_name: str,
    table_name: str,
    group_by: Optional[list[str]] = None,
    metrics: Optional[list[dict]] = None,
    filters: Optional[list[dict]] = None,
    sort_by: str = "",
    sort_desc: bool = True,
    limit: int = 100,
) -> str:
    """
    Execute a structured, read-only Pandas analysis.

    metrics example:
    [
      {"column": "Net_Revenue_SAR", "operation": "sum", "alias": "Revenue"},
      {"column": "Order_ID", "operation": "nunique", "alias": "Orders"}
    ]

    filters example:
    [
      {"column": "Order_Status", "operator": "eq", "value": "Won"}
    ]

    group_by example:
    ["Region"]
    """
    try:
        table = resolve_sheet(dataset_name, table_name)
        df = load_table(dataset_name, table).copy()

        group_by = group_by or []
        metrics = metrics or []
        filters = filters or []

        for col in group_by:
            if col not in df.columns:
                raise ValueError(f"Group-by column not found: {col}")

        df = apply_filters(df, filters)

        if not metrics:
            # Default output: filtered rows only
            result = df.copy()

        elif group_by:
            named_aggs = {}

            for metric in metrics:
                col = str(metric.get("column", "")).strip()
                op = str(metric.get("operation", "")).strip().casefold()
                alias = str(metric.get("alias", "")).strip() or f"{op}_{col}"

                if col not in df.columns:
                    raise ValueError(f"Metric column not found: {col}")

                pandas_op = {
                    "sum": "sum",
                    "mean": "mean",
                    "avg": "mean",
                    "min": "min",
                    "max": "max",
                    "count": "count",
                    "nunique": "nunique",
                    "distinct_count": "nunique",
                }.get(op)

                if not pandas_op:
                    raise ValueError(f"Unsupported aggregation: {op}")

                named_aggs[alias] = pd.NamedAgg(column=col, aggfunc=pandas_op)

            result = (
                df.groupby(group_by, dropna=False)
                .agg(**named_aggs)
                .reset_index()
            )

        else:
            row = {}

            for metric in metrics:
                col = str(metric.get("column", "")).strip()
                op = str(metric.get("operation", "")).strip().casefold()
                alias = str(metric.get("alias", "")).strip() or f"{op}_{col}"

                if col not in df.columns:
                    raise ValueError(f"Metric column not found: {col}")

                row[alias] = aggregate_series(df[col], op)

            result = pd.DataFrame([row])

        if sort_by:
            if sort_by not in result.columns:
                raise ValueError(f"Sort column not found in result: {sort_by}")
            result = result.sort_values(sort_by, ascending=not bool(sort_desc))

        limit = max(1, min(int(limit), 1000))
        result = result.head(limit)

        return success_json(
            records(result),
            row_count=len(result),
            dataset_name=dataset_name,
            table_name=table,
            source_rows_after_filter=int(len(df)),
        )

    except Exception as exc:
        return error_json(
            exc,
            "analysis_error",
            dataset_name=dataset_name,
            table_name=table_name,
        )


@mcp.tool()
def clear_cache() -> str:
    """Clear cached Excel metadata/data after files are updated."""
    try:
        discover_datasets.cache_clear()
        workbook_sheet_names.cache_clear()
        load_table.cache_clear()

        return success_json({"cache_cleared": True})

    except Exception as exc:
        return error_json(exc, "cache_clear_error")


# =========================================================
# Entry point
# =========================================================

if __name__ == "__main__":
    mcp.run()
