from __future__ import annotations

import io
import math
import re
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# -----------------------------------------------------------------------------
# App configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Oilfield Inventory Consolidation Dashboard",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_TITLE = "Oilfield Inventory Consolidation Value Dashboard"
DEFAULT_ECONOMICS_FILE = "economics.xlsx"
DEFAULT_CONSOLIDATION_FILE = "consol.xlsx"

KEY_CONSOL_COLUMNS = {
    "consolidation_number": "Consolidation #",
    "well_1": "well 1 entity",
    "well_2": "well 2 entity",
    "enhanced": "enhanced well entity",
}

DEFAULT_METRIC_CANDIDATES = [
    "Npv Cash Flow BTax 10.0% (M$)",
    "Npv Investment BTax 10.0% (M$)",
    "Boe Company Share Total (boe)",
    "1st Year Production (boe)",
    "Total Op Costs (M$)",
    "Combined CO2e Gross Total (t)",
]

PREFERRED_VALUE_METRICS = [
    "Npv Cash Flow BTax 10.0% (M$)",
    "Npv Cash Flow ATax 10.0% (M$)",
    "Npv Cash Flow BTax 12.0% (M$)",
    "Npv Cash Flow BTax 15.0% (M$)",
    "Npv Cash Flow BTax 5.0% (M$)",
    "Npv Cash Flow BTax 0.0% (M$)",
]

PREFERRED_INVESTMENT_METRICS = [
    "Npv Investment BTax 10.0% (M$)",
    "Npv Investment ATax 10.0% (M$)",
    "Npv Investment BTax 0.0% (M$)",
]

PREFERRED_RESERVE_METRICS = [
    "Boe Company Share Total (boe)",
    "Boe WI Total (boe)",
    "Boe Gross Total (boe)",
    "Boe Company Net/NRI Total (boe)",
]

PREFERRED_PRODUCTION_METRICS = [
    "1st Year Production (boe)",
    "1st Year Production Rate (boepd)",
]

CASE_ORDER = ["Consolidation", "Extension", "Creation"]


# -----------------------------------------------------------------------------
# Styling
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    [data-testid="stMetricValue"] {font-size: 1.65rem;}
    [data-testid="stMetricDelta"] {font-size: 0.95rem;}
    .small-note {color: #6b7280; font-size: 0.88rem;}
    .status-good {padding: .65rem .85rem; border-radius: .5rem; background: #ecfdf5; color: #065f46;}
    .status-warn {padding: .65rem .85rem; border-radius: .5rem; background: #fffbeb; color: #92400e;}
    .status-bad {padding: .65rem .85rem; border-radius: .5rem; background: #fef2f2; color: #991b1b;}
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# Utility functions
# -----------------------------------------------------------------------------
def normalize_header(value: object) -> str:
    """Normalize a column header for resilient matching."""
    text = "" if value is None else str(value)
    text = text.replace("\u00a0", " ").strip().lower()
    return re.sub(r"\s+", " ", text)


def normalize_entity(value: object) -> object:
    """Normalize lookup keys without changing meaningful internal characters."""
    if pd.isna(value):
        return pd.NA
    text = str(value).replace("\u00a0", " ").strip()
    if not text:
        return pd.NA
    return re.sub(r"\s+", " ", text).casefold()


def clean_display_entity(value: object) -> object:
    if pd.isna(value):
        return pd.NA
    text = str(value).replace("\u00a0", " ").strip()
    return re.sub(r"\s+", " ", text) if text else pd.NA


def find_actual_column(columns: Iterable[object], desired: str) -> str | None:
    lookup = {normalize_header(c): str(c) for c in columns}
    return lookup.get(normalize_header(desired))


def first_existing(columns: Iterable[str], candidates: Iterable[str]) -> str | None:
    normalized = {normalize_header(c): c for c in columns}
    for candidate in candidates:
        found = normalized.get(normalize_header(candidate))
        if found is not None:
            return found
    return None


def is_blank_series(series: pd.Series) -> pd.Series:
    return series.isna() | series.astype("string").str.strip().eq("")


def classify_case(w1: pd.Series, w2: pd.Series) -> pd.Series:
    has_w1 = ~is_blank_series(w1)
    has_w2 = ~is_blank_series(w2)
    return pd.Series(
        np.select(
            [has_w1 & has_w2, has_w1 ^ has_w2, ~has_w1 & ~has_w2],
            CASE_ORDER,
            default="Invalid",
        ),
        index=w1.index,
        dtype="string",
    )


def case_source_count(case_type: str) -> int:
    return {"Consolidation": 2, "Extension": 1, "Creation": 0}.get(case_type, 0)


def safe_numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = (
        series.astype("string")
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce")


def infer_additive_metric(name: str) -> bool:
    """Conservative test: totals and NPVs are additive; rates/ratios/dates are not."""
    n = normalize_header(name)
    non_additive_tokens = [
        "%", "$/", "per boe", "intensity", "years", "date", "rate", "fraction",
        "npv /", "payout", "ror", "life", "latitude", "longitude", "pos", "coe",
        "cost of production", "cost of reserves", "netback /", "avg.", "average",
        "boepd", "cum steam oil ratio", "initial wi",
    ]
    if any(token in n for token in non_additive_tokens):
        return False
    additive_tokens = [
        "(m$)", "total (", "production (boe)", "cum (boe)", "tech ult rec",
        "tech rem rec", "npv cash flow", "npv op income", "npv investment",
    ]
    return any(token in n for token in additive_tokens)


def metric_unit(metric: str) -> str:
    match = re.search(r"\(([^()]*)\)\s*$", metric)
    return match.group(1) if match else ""


def format_number(value: float | int | None, unit: str = "", decimals: int = 1) -> str:
    if value is None or pd.isna(value):
        return "—"
    value = float(value)
    unit_lower = unit.lower()
    if unit_lower == "m$":
        return f"${value:,.{decimals}f}M"
    if unit_lower in {"boe", "bbl", "mcf", "t", "st"}:
        abs_value = abs(value)
        if abs_value >= 1_000_000:
            return f"{value / 1_000_000:,.{decimals}f} MM{unit.upper()}"
        if abs_value >= 1_000:
            return f"{value / 1_000:,.{decimals}f} M{unit.upper()}"
        return f"{value:,.0f} {unit}"
    if "$" in unit:
        return f"${value:,.{decimals}f}"
    if "%" in unit:
        return f"{value:,.{decimals}f}%"
    return f"{value:,.{decimals}f}" + (f" {unit}" if unit else "")


def metric_delta_label(value: float, unit: str) -> str:
    if pd.isna(value):
        return ""
    prefix = "+" if value > 0 else ""
    return f"{prefix}{format_number(value, unit)}"


def dataframe_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = re.sub(r"[\\/*?:\[\]]", "_", sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
            ws = writer.book[safe_name]
            ws.freeze_panes = "A2"
            ws.sheet_view.showGridLines = False
            for col_cells in ws.columns:
                max_len = min(max(len(str(cell.value or "")) for cell in col_cells), 45)
                ws.column_dimensions[col_cells[0].column_letter].width = max(11, max_len + 2)
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="1F4E78")
    return output.getvalue()


@st.cache_data(show_spinner=False)
def read_excel_sheets(file_bytes: bytes) -> tuple[list[str], dict[str, pd.DataFrame]]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return xls.sheet_names, {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}


def source_bytes(uploaded_file, local_filename: str) -> bytes | None:
    if uploaded_file is not None:
        return uploaded_file.getvalue()
    local_path = Path(local_filename)
    if local_path.exists() and local_path.is_file():
        return local_path.read_bytes()
    return None


def recommend_sheet(sheet_frames: dict[str, pd.DataFrame], required_columns: list[str]) -> str:
    best_sheet = next(iter(sheet_frames))
    best_score = -1
    required_norm = {normalize_header(c) for c in required_columns}
    for name, df in sheet_frames.items():
        score = len(required_norm.intersection({normalize_header(c) for c in df.columns}))
        if score > best_score:
            best_sheet = name
            best_score = score
    return best_sheet


def calculate_event_model(
    consol: pd.DataFrame,
    economics: pd.DataFrame,
    entity_col: str,
    selected_metrics: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Create event-level before/after economics and entity-level lookup diagnostics."""
    actual_cols = {}
    for semantic, desired in KEY_CONSOL_COLUMNS.items():
        actual = find_actual_column(consol.columns, desired)
        if actual is None:
            raise ValueError(f'Missing required consolidation column: "{desired}"')
        actual_cols[semantic] = actual

    work = consol.copy()
    work = work.rename(
        columns={
            actual_cols["consolidation_number"]: "Consolidation #",
            actual_cols["well_1"]: "well 1 entity",
            actual_cols["well_2"]: "well 2 entity",
            actual_cols["enhanced"]: "enhanced well entity",
        }
    )
    work["Source Row"] = np.arange(2, len(work) + 2)

    for col in ["well 1 entity", "well 2 entity", "enhanced well entity"]:
        work[col] = work[col].map(clean_display_entity)
        work[f"__key_{col}"] = work[col].map(normalize_entity)

    work["Case Type"] = classify_case(work["well 1 entity"], work["well 2 entity"])
    work["Before Well Count"] = work["Case Type"].map({"Consolidation": 2, "Extension": 1, "Creation": 0}).fillna(0).astype(int)
    work["After Well Count"] = 1
    work["Net Well Count Change"] = work["After Well Count"] - work["Before Well Count"]
    work["Before Lateral Miles"] = work["Before Well Count"].astype(float)
    work["After Lateral Miles"] = 2.0
    work["Net Lateral Miles"] = work["After Lateral Miles"] - work["Before Lateral Miles"]

    econ = economics.copy()
    econ[entity_col] = econ[entity_col].map(clean_display_entity)
    econ["__entity_key"] = econ[entity_col].map(normalize_entity)

    for metric in selected_metrics:
        econ[metric] = safe_numeric(econ[metric])

    duplicate_mask = econ["__entity_key"].notna() & econ["__entity_key"].duplicated(keep=False)
    duplicates = econ.loc[duplicate_mask, [entity_col, "__entity_key"] + selected_metrics].sort_values("__entity_key")

    # Keep the first record only to permit diagnostics display; results are invalidated in UI if duplicates exist.
    lookup = econ.dropna(subset=["__entity_key"]).drop_duplicates("__entity_key", keep="first").set_index("__entity_key")

    source_key_columns = ["__key_well 1 entity", "__key_well 2 entity"]
    for metric in selected_metrics:
        w1 = work[source_key_columns[0]].map(lookup[metric])
        w2 = work[source_key_columns[1]].map(lookup[metric])
        enhanced = work["__key_enhanced well entity"].map(lookup[metric])

        # Blank source entities contribute zero; populated unmatched entities remain NaN.
        w1_contribution = w1.where(work[source_key_columns[0]].notna(), 0.0)
        w2_contribution = w2.where(work[source_key_columns[1]].notna(), 0.0)
        work[f"Before | {metric}"] = w1_contribution + w2_contribution
        work[f"After | {metric}"] = enhanced
        work[f"Change | {metric}"] = work[f"After | {metric}"] - work[f"Before | {metric}"]

    lookup_key_set = set(lookup.index)
    unmatched_records: list[dict[str, object]] = []
    for role, key_col, display_col in [
        ("Well 1", "__key_well 1 entity", "well 1 entity"),
        ("Well 2", "__key_well 2 entity", "well 2 entity"),
        ("Enhanced", "__key_enhanced well entity", "enhanced well entity"),
    ]:
        mask = work[key_col].notna() & ~work[key_col].isin(lookup_key_set)
        for _, row in work.loc[mask].iterrows():
            unmatched_records.append(
                {
                    "Source Row": row["Source Row"],
                    "Consolidation #": row["Consolidation #"],
                    "Case Type": row["Case Type"],
                    "Entity Role": role,
                    "Unmatched Entity": row[display_col],
                }
            )
    unmatched = pd.DataFrame(unmatched_records)

    helper_cols = [c for c in work.columns if c.startswith("__key_")]
    clean_work = work.drop(columns=helper_cols)
    return clean_work, duplicates, unmatched


def summarize_by_case(event_df: pd.DataFrame, metrics: list[str]) -> pd.DataFrame:
    aggregations = {
        "Consolidation #": "count",
        "Before Well Count": "sum",
        "After Well Count": "sum",
        "Net Well Count Change": "sum",
        "Before Lateral Miles": "sum",
        "After Lateral Miles": "sum",
        "Net Lateral Miles": "sum",
    }
    for metric in metrics:
        aggregations[f"Before | {metric}"] = "sum"
        aggregations[f"After | {metric}"] = "sum"
        aggregations[f"Change | {metric}"] = "sum"

    summary = event_df.groupby("Case Type", dropna=False).agg(aggregations).reset_index()
    summary = summary.rename(columns={"Consolidation #": "Events"})
    summary["Case Type"] = pd.Categorical(summary["Case Type"], CASE_ORDER + ["Invalid"], ordered=True)
    return summary.sort_values("Case Type").reset_index(drop=True)


def build_entity_usage(event_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for role_col, scenario in [
        ("well 1 entity", "Before"),
        ("well 2 entity", "Before"),
        ("enhanced well entity", "After"),
    ]:
        part = event_df[["Consolidation #", "Case Type", role_col]].dropna(subset=[role_col]).copy()
        part = part.rename(columns={role_col: "Entity"})
        part["Role"] = role_col
        part["Scenario"] = scenario
        rows.append(part)
    if not rows:
        return pd.DataFrame(columns=["Entity", "Scenario", "Case Type", "Uses"])
    usage = pd.concat(rows, ignore_index=True)
    return (
        usage.groupby(["Entity", "Scenario", "Case Type"], dropna=False)
        .size()
        .rename("Uses")
        .reset_index()
        .sort_values(["Uses", "Entity"], ascending=[False, True])
    )


# -----------------------------------------------------------------------------
# Header and input controls
# -----------------------------------------------------------------------------
st.title(APP_TITLE)
st.caption(
    "Classifies inventory changes, joins entity type averages, and calculates event- and portfolio-level before/after economics."
)

with st.sidebar:
    st.header("1. Load workbooks")
    econ_upload = st.file_uploader("Entity economics workbook", type=["xlsx", "xlsm"], key="economics")
    consol_upload = st.file_uploader("Consolidation workbook", type=["xlsx", "xlsm"], key="consolidation")
    st.caption(
        f"When no files are uploaded, the app looks for `{DEFAULT_ECONOMICS_FILE}` and "
        f"`{DEFAULT_CONSOLIDATION_FILE}` beside the Python script."
    )

    economics_bytes = source_bytes(econ_upload, DEFAULT_ECONOMICS_FILE)
    consolidation_bytes = source_bytes(consol_upload, DEFAULT_CONSOLIDATION_FILE)

if economics_bytes is None or consolidation_bytes is None:
    st.info(
        "Upload both workbooks in the sidebar, or place `economics.xlsx` and `consol.xlsx` in the same folder as this app."
    )
    st.stop()

try:
    econ_sheet_names, econ_sheets = read_excel_sheets(economics_bytes)
    consol_sheet_names, consol_sheets = read_excel_sheets(consolidation_bytes)
except Exception as exc:
    st.error(f"The workbooks could not be read: {exc}")
    st.stop()

with st.sidebar:
    st.header("2. Select sheets")
    suggested_econ_sheet = recommend_sheet(econ_sheets, ["Entity"] + DEFAULT_METRIC_CANDIDATES)
    suggested_consol_sheet = recommend_sheet(consol_sheets, list(KEY_CONSOL_COLUMNS.values()))
    econ_sheet = st.selectbox(
        "Economics sheet",
        econ_sheet_names,
        index=econ_sheet_names.index(suggested_econ_sheet),
    )
    consol_sheet = st.selectbox(
        "Consolidation sheet",
        consol_sheet_names,
        index=consol_sheet_names.index(suggested_consol_sheet),
    )

raw_economics = econ_sheets[econ_sheet].copy()
raw_consol = consol_sheets[consol_sheet].copy()

with st.sidebar:
    st.header("3. Configure columns")
    entity_guess = first_existing(raw_economics.columns.astype(str), ["Entity", "Entity Name", "Well Entity"])
    entity_col = st.selectbox(
        "Economics entity key",
        options=list(raw_economics.columns.astype(str)),
        index=list(raw_economics.columns.astype(str)).index(entity_guess) if entity_guess else 0,
    )

numeric_candidates = []
for col in raw_economics.columns.astype(str):
    if col == entity_col:
        continue
    converted = safe_numeric(raw_economics[col])
    if converted.notna().sum() > 0:
        numeric_candidates.append(col)

additive_candidates = [c for c in numeric_candidates if infer_additive_metric(c)]
default_metrics = [c for c in DEFAULT_METRIC_CANDIDATES if c in additive_candidates]
if not default_metrics:
    default_metrics = additive_candidates[: min(6, len(additive_candidates))]

with st.sidebar:
    st.header("4. Select additive metrics")
    selected_metrics = st.multiselect(
        "Portfolio metrics to aggregate",
        options=additive_candidates,
        default=default_metrics,
        help="Only additive metrics should be selected. Rates, payout, ROR, intensities, and per-unit values should not be summed.",
    )

    value_metric_default = first_existing(selected_metrics, PREFERRED_VALUE_METRICS)
    value_metric = st.selectbox(
        "Primary value metric",
        options=selected_metrics,
        index=selected_metrics.index(value_metric_default) if value_metric_default in selected_metrics else 0,
    ) if selected_metrics else None

    investment_options = ["None"] + selected_metrics
    investment_guess = first_existing(selected_metrics, PREFERRED_INVESTMENT_METRICS)
    investment_metric = st.selectbox(
        "Investment metric",
        options=investment_options,
        index=investment_options.index(investment_guess) if investment_guess in investment_options else 0,
    )
    investment_metric = None if investment_metric == "None" else investment_metric

    reserve_options = ["None"] + selected_metrics
    reserve_guess = first_existing(selected_metrics, PREFERRED_RESERVE_METRICS)
    reserve_metric = st.selectbox(
        "Reserve metric",
        options=reserve_options,
        index=reserve_options.index(reserve_guess) if reserve_guess in reserve_options else 0,
    )
    reserve_metric = None if reserve_metric == "None" else reserve_metric

    production_options = ["None"] + selected_metrics
    production_guess = first_existing(selected_metrics, PREFERRED_PRODUCTION_METRICS)
    production_metric = st.selectbox(
        "Production metric",
        options=production_options,
        index=production_options.index(production_guess) if production_guess in production_options else 0,
    )
    production_metric = None if production_metric == "None" else production_metric

if not selected_metrics:
    st.warning("Select at least one additive metric in the sidebar.")
    st.stop()

try:
    event_df, duplicate_entities, unmatched_entities = calculate_event_model(
        raw_consol,
        raw_economics,
        entity_col,
        selected_metrics,
    )
except ValueError as exc:
    st.error(str(exc))
    st.stop()
except Exception as exc:
    st.exception(exc)
    st.stop()

invalid_cases = event_df[event_df["Case Type"].eq("Invalid")]
missing_enhanced = event_df[event_df["enhanced well entity"].isna()]
blocking_errors = len(duplicate_entities) > 0 or len(unmatched_entities) > 0 or len(invalid_cases) > 0 or len(missing_enhanced) > 0

case_summary = summarize_by_case(event_df, selected_metrics)
entity_usage = build_entity_usage(event_df)

# Filters
with st.sidebar:
    st.header("5. Dashboard filters")
    available_cases = [c for c in CASE_ORDER if c in set(event_df["Case Type"].dropna())]
    selected_cases = st.multiselect("Case types", available_cases, default=available_cases)
    search_text = st.text_input("Entity or consolidation search", value="")

filtered_events = event_df[event_df["Case Type"].isin(selected_cases)].copy()
if search_text.strip():
    search = search_text.strip().casefold()
    searchable = (
        filtered_events[["Consolidation #", "well 1 entity", "well 2 entity", "enhanced well entity"]]
        .astype("string")
        .fillna("")
        .agg(" | ".join, axis=1)
        .str.casefold()
    )
    filtered_events = filtered_events[searchable.str.contains(re.escape(search), na=False)]

filtered_case_summary = summarize_by_case(filtered_events, selected_metrics) if not filtered_events.empty else pd.DataFrame()

# -----------------------------------------------------------------------------
# Quality banner
# -----------------------------------------------------------------------------
if blocking_errors:
    st.markdown(
        '<div class="status-bad"><b>Results require review.</b> One or more duplicate keys, unmatched entities, invalid cases, or missing enhanced entities were detected. Portfolio calculations containing unmatched values may be incomplete.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="status-good"><b>Data validation passed.</b> Entity keys are unique, all populated entities matched, and every event has an enhanced entity.</div>',
        unsafe_allow_html=True,
    )

# -----------------------------------------------------------------------------
# Dashboard tabs
# -----------------------------------------------------------------------------
tab_exec, tab_cases, tab_metrics, tab_qa, tab_data = st.tabs(
    ["Executive Summary", "Case Analysis", "Metric Explorer", "Data Quality", "Event Detail"]
)

with tab_exec:
    st.subheader("Portfolio transformation")

    before_wells = filtered_events["Before Well Count"].sum()
    after_wells = filtered_events["After Well Count"].sum()
    net_wells = filtered_events["Net Well Count Change"].sum()
    before_miles = filtered_events["Before Lateral Miles"].sum()
    after_miles = filtered_events["After Lateral Miles"].sum()
    net_miles = filtered_events["Net Lateral Miles"].sum()

    value_unit = metric_unit(value_metric) if value_metric else ""
    before_value = filtered_events[f"Before | {value_metric}"].sum(min_count=1) if value_metric else np.nan
    after_value = filtered_events[f"After | {value_metric}"].sum(min_count=1) if value_metric else np.nan
    value_change = filtered_events[f"Change | {value_metric}"].sum(min_count=1) if value_metric else np.nan

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Inventory events", f"{len(filtered_events):,}")
    k2.metric("Wells before", f"{before_wells:,.0f}")
    k3.metric("Wells after", f"{after_wells:,.0f}", f"{net_wells:+,.0f}")
    k4.metric("Lateral miles after", f"{after_miles:,.0f}", f"{net_miles:+,.0f}")
    k5.metric(
        value_metric or "Value change",
        format_number(after_value, value_unit),
        metric_delta_label(value_change, value_unit),
    )

    c1, c2 = st.columns([1.05, 1])
    with c1:
        physical_bridge = pd.DataFrame(
            {
                "Scenario": ["Before", "After"],
                "Well Count": [before_wells, after_wells],
                "Lateral Miles": [before_miles, after_miles],
            }
        )
        physical_long = physical_bridge.melt("Scenario", var_name="Measure", value_name="Amount")
        fig = px.bar(
            physical_long,
            x="Measure",
            y="Amount",
            color="Scenario",
            barmode="group",
            text_auto=",.0f",
            title="Physical inventory bridge",
        )
        fig.update_layout(legend_title_text="", yaxis_title="Count / miles")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        case_counts = (
            filtered_events.groupby("Case Type", observed=False)
            .size()
            .reindex(CASE_ORDER, fill_value=0)
            .rename("Events")
            .reset_index()
        )
        fig = px.bar(
            case_counts,
            x="Case Type",
            y="Events",
            text_auto=True,
            title="Events by case type",
            category_orders={"Case Type": CASE_ORDER},
        )
        fig.update_layout(xaxis_title="", yaxis_title="Events", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Value bridge by case")
    if not filtered_case_summary.empty and value_metric:
        value_bridge = filtered_case_summary[["Case Type", f"Change | {value_metric}"]].copy()
        value_bridge = value_bridge.rename(columns={f"Change | {value_metric}": "Value Change"})
        fig = px.bar(
            value_bridge,
            x="Case Type",
            y="Value Change",
            text_auto=",.1f",
            category_orders={"Case Type": CASE_ORDER},
            title=f"Contribution to change in {value_metric}",
        )
        fig.add_hline(y=0, line_width=1)
        fig.update_layout(xaxis_title="", yaxis_title=value_unit or "Value", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Key economic outputs")
    display_metrics = [m for m in [value_metric, investment_metric, reserve_metric, production_metric] if m]
    # Preserve order and remove duplicates.
    display_metrics = list(dict.fromkeys(display_metrics))
    cols = st.columns(max(1, len(display_metrics)))
    for col, metric in zip(cols, display_metrics):
        unit = metric_unit(metric)
        before = filtered_events[f"Before | {metric}"].sum(min_count=1)
        after = filtered_events[f"After | {metric}"].sum(min_count=1)
        change = filtered_events[f"Change | {metric}"].sum(min_count=1)
        with col:
            st.metric(metric, format_number(after, unit), metric_delta_label(change, unit))
            st.caption(f"Before: {format_number(before, unit)}")

    if value_metric and investment_metric:
        delta_investment = filtered_events[f"Change | {investment_metric}"].sum(min_count=1)
        if pd.notna(delta_investment) and not math.isclose(delta_investment, 0.0, abs_tol=1e-12):
            incremental_efficiency = value_change / abs(delta_investment)
            st.info(
                f"Incremental value / absolute incremental investment: **{incremental_efficiency:,.2f}x**. "
                "Review the sign convention of the selected investment field before using this ratio externally."
            )

    st.subheader("Automated management summary")
    value_text = format_number(value_change, value_unit)
    reserve_text = ""
    if reserve_metric:
        reserve_change = filtered_events[f"Change | {reserve_metric}"].sum(min_count=1)
        reserve_text = f" and changed {reserve_metric} by {format_number(reserve_change, metric_unit(reserve_metric))}"
    summary_text = (
        f"The selected inventory redesign includes {len(filtered_events):,} events and changes the planned well count "
        f"from {before_wells:,.0f} to {after_wells:,.0f} ({net_wells:+,.0f}). Planned lateral inventory changes "
        f"from {before_miles:,.0f} to {after_miles:,.0f} miles ({net_miles:+,.0f}). The resulting change in "
        f"{value_metric} is {value_text}{reserve_text}."
    )
    st.text_area("Copy-ready summary", summary_text, height=125)

with tab_cases:
    st.subheader("Case-level economics")
    if filtered_case_summary.empty:
        st.info("No events match the current filters.")
    else:
        overview_cols = [
            "Case Type", "Events", "Before Well Count", "After Well Count", "Net Well Count Change",
            "Before Lateral Miles", "After Lateral Miles", "Net Lateral Miles",
        ]
        if value_metric:
            overview_cols += [f"Before | {value_metric}", f"After | {value_metric}", f"Change | {value_metric}"]
        st.dataframe(filtered_case_summary[overview_cols], use_container_width=True, hide_index=True)

        metric_for_case = st.selectbox("Case comparison metric", selected_metrics, index=selected_metrics.index(value_metric) if value_metric in selected_metrics else 0)
        case_chart_data = filtered_case_summary[
            ["Case Type", f"Before | {metric_for_case}", f"After | {metric_for_case}"]
        ].melt("Case Type", var_name="Scenario", value_name="Value")
        case_chart_data["Scenario"] = case_chart_data["Scenario"].str.split(" | ", regex=False).str[0]
        fig = px.bar(
            case_chart_data,
            x="Case Type",
            y="Value",
            color="Scenario",
            barmode="group",
            text_auto=",.1f",
            category_orders={"Case Type": CASE_ORDER},
            title=f"Before and after: {metric_for_case}",
        )
        fig.update_layout(xaxis_title="", yaxis_title=metric_unit(metric_for_case) or "Value", legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Largest event-level contributors")
        top_n = st.slider("Events shown", 5, 50, 15)
        contributor_data = filtered_events[
            ["Consolidation #", "Case Type", "well 1 entity", "well 2 entity", "enhanced well entity", f"Change | {metric_for_case}"]
        ].copy()
        contributor_data = contributor_data.sort_values(f"Change | {metric_for_case}", ascending=False)
        st.dataframe(contributor_data.head(top_n), use_container_width=True, hide_index=True)

        chart_top = contributor_data.dropna(subset=[f"Change | {metric_for_case}"]).head(top_n).copy()
        chart_top["Event"] = chart_top["Consolidation #"].astype("string")
        fig = px.bar(
            chart_top.sort_values(f"Change | {metric_for_case}"),
            x=f"Change | {metric_for_case}",
            y="Event",
            color="Case Type",
            orientation="h",
            title=f"Top positive contributors to {metric_for_case}",
            category_orders={"Case Type": CASE_ORDER},
        )
        fig.update_layout(xaxis_title=metric_unit(metric_for_case) or "Change", yaxis_title="Consolidation #", legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)

with tab_metrics:
    st.subheader("Metric explorer")
    explorer_metric = st.selectbox("Metric", selected_metrics, index=selected_metrics.index(value_metric) if value_metric in selected_metrics else 0, key="explorer")
    explorer_unit = metric_unit(explorer_metric)

    before_total = filtered_events[f"Before | {explorer_metric}"].sum(min_count=1)
    after_total = filtered_events[f"After | {explorer_metric}"].sum(min_count=1)
    change_total = filtered_events[f"Change | {explorer_metric}"].sum(min_count=1)
    pct_change = change_total / abs(before_total) if pd.notna(before_total) and not math.isclose(before_total, 0.0, abs_tol=1e-12) else np.nan

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Before", format_number(before_total, explorer_unit))
    m2.metric("After", format_number(after_total, explorer_unit))
    m3.metric("Change", format_number(change_total, explorer_unit))
    m4.metric("Change vs. before", f"{pct_change:.1%}" if pd.notna(pct_change) else "—")

    waterfall_values = []
    for case in CASE_ORDER:
        case_value = filtered_events.loc[
            filtered_events["Case Type"].eq(case), f"Change | {explorer_metric}"
        ].sum(min_count=1)
        waterfall_values.append(0.0 if pd.isna(case_value) else case_value)

    fig = go.Figure(
        go.Waterfall(
            name="Portfolio bridge",
            orientation="v",
            measure=["absolute", "relative", "relative", "relative", "total"],
            x=["Before"] + CASE_ORDER + ["After"],
            y=[before_total] + waterfall_values + [0],
            connector={"line": {"width": 1}},
            textposition="outside",
            text=[f"{before_total:,.1f}"] + [f"{v:+,.1f}" for v in waterfall_values] + [f"{after_total:,.1f}"],
        )
    )
    fig.update_layout(title=f"Portfolio waterfall: {explorer_metric}", yaxis_title=explorer_unit or "Value", showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Distribution of event value changes")
    fig = px.histogram(
        filtered_events,
        x=f"Change | {explorer_metric}",
        color="Case Type",
        marginal="box",
        nbins=min(40, max(10, int(math.sqrt(max(len(filtered_events), 1))) * 2)),
        category_orders={"Case Type": CASE_ORDER},
    )
    fig.update_layout(xaxis_title=f"Change ({explorer_unit})" if explorer_unit else "Change", yaxis_title="Events", legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Entity usage")
    st.dataframe(entity_usage, use_container_width=True, hide_index=True)

with tab_qa:
    st.subheader("Validation results")
    qa1, qa2, qa3, qa4 = st.columns(4)
    qa1.metric("Economics rows", f"{len(raw_economics):,}")
    qa2.metric("Consolidation rows", f"{len(raw_consol):,}")
    qa3.metric("Duplicate entity rows", f"{len(duplicate_entities):,}")
    qa4.metric("Unmatched references", f"{len(unmatched_entities):,}")

    if len(duplicate_entities):
        st.error(
            "Duplicate normalized entity keys exist in the economics workbook. The app displays a provisional result using the first occurrence, but the economics key should be unique before relying on totals."
        )
        st.dataframe(duplicate_entities, use_container_width=True, hide_index=True)
    else:
        st.success("No duplicate normalized entity keys were found in the economics data.")

    if len(unmatched_entities):
        st.error("The following populated consolidation entities were not found in the economics lookup.")
        st.dataframe(unmatched_entities, use_container_width=True, hide_index=True)
    else:
        st.success("Every populated consolidation entity matched an economics entity.")

    if len(missing_enhanced):
        st.error("Every event must have an enhanced well entity. The following rows are missing one.")
        st.dataframe(
            missing_enhanced[["Source Row", "Consolidation #", "Case Type", "well 1 entity", "well 2 entity", "enhanced well entity"]],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success("Every consolidation row has an enhanced well entity.")

    if len(invalid_cases):
        st.error("Some rows could not be classified using the supplied rules.")
        st.dataframe(invalid_cases, use_container_width=True, hide_index=True)
    else:
        st.success("All rows were classified as Consolidation, Extension, or Creation.")

    st.subheader("Classification reconciliation")
    recon = (
        event_df.groupby("Case Type", observed=False)
        .agg(
            Events=("Consolidation #", "size"),
            Before_Wells=("Before Well Count", "sum"),
            After_Wells=("After Well Count", "sum"),
            Net_Well_Change=("Net Well Count Change", "sum"),
            Net_Lateral_Miles=("Net Lateral Miles", "sum"),
        )
        .reindex(CASE_ORDER + ["Invalid"])
        .dropna(how="all")
        .reset_index()
    )
    st.dataframe(recon, use_container_width=True, hide_index=True)

    with st.expander("Metric classification and numeric completeness"):
        completeness = []
        for metric in selected_metrics:
            completeness.append(
                {
                    "Metric": metric,
                    "Unit": metric_unit(metric),
                    "Economics Non-Null": safe_numeric(raw_economics[metric]).notna().sum(),
                    "Economics Rows": len(raw_economics),
                    "Additive Inference": infer_additive_metric(metric),
                    "Event Before Non-Null": event_df[f"Before | {metric}"].notna().sum(),
                    "Event After Non-Null": event_df[f"After | {metric}"].notna().sum(),
                }
            )
        st.dataframe(pd.DataFrame(completeness), use_container_width=True, hide_index=True)

with tab_data:
    st.subheader("Event-level calculation table")
    base_cols = [
        "Source Row", "Consolidation #", "Case Type", "well 1 entity", "well 2 entity", "enhanced well entity",
        "Before Well Count", "After Well Count", "Net Well Count Change",
        "Before Lateral Miles", "After Lateral Miles", "Net Lateral Miles",
    ]
    metric_cols = []
    for metric in selected_metrics:
        metric_cols.extend([f"Before | {metric}", f"After | {metric}", f"Change | {metric}"])
    display_event_df = filtered_events[base_cols + metric_cols]
    st.dataframe(display_event_df, use_container_width=True, hide_index=True, height=600)

    export_sheets = {
        "Executive Summary": filtered_case_summary,
        "Event Detail": display_event_df,
        "Entity Usage": entity_usage,
        "Unmatched Entities": unmatched_entities,
        "Duplicate Economics Keys": duplicate_entities,
    }
    export_bytes = dataframe_to_excel_bytes(export_sheets)
    st.download_button(
        "Download dashboard results (.xlsx)",
        data=export_bytes,
        file_name="inventory_consolidation_dashboard_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()
st.caption(
    "Method: source entities represent before-state type averages; each enhanced entity represents one after-state two-mile well. "
    "Consolidations have two source entities, extensions have one, and creations have none. Blank source entities contribute zero; populated unmatched entities remain missing and are flagged."
)
