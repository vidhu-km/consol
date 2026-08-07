from __future__ import annotations

import warnings
from io import BytesIO
from pathlib import Path
from typing import Any
import zipfile

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ────────────────────────────────────────────────────────────────────────────────
# PATHS & CONSTANTS
# ────────────────────────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = APP_DIR / "economics.xlsx"

SHEET_WELLS = "wells"
SHEET_ECONOMIC_INDICATORS = "economic indicators"
SHEET_FORECASTS = "forecasts"

# Column names — wells
COL_EVENT = "event #"
COL_OLD_UWI_1 = "old well 1 UWI"
COL_OLD_UWI_2 = "old well 2 UWI"
COL_OLD_CURVE_1 = "old well 1 type curve"
COL_OLD_CURVE_2 = "old well 2 type curve"
COL_NEW_CURVE = "new well type curve"

# Column names — indicators
COL_TYPE_CURVE = "type curve"
COL_NPV10 = "Npv Cash Flow BTax 10.0% (M$)"
COL_NPV_INVESTMENT_RATIO = "NPV / Disc. Invest BTax"
COL_PAYOUT = "Payout BTax (years)"
COL_RESERVES = "Boe WI Total (boe)"
COL_FIRST_YEAR_RATE = "1st Year Production Rate (boepd)"
COL_COST_OF_RESERVES = "Cost of Reserves ($/boe)"
COL_IP30 = "IP30 Cum (boe)"
COL_CAPEX = "Npv Investment BTax  0.0% (M$)"
COL_ROR = "BTax Disc. CF. ROR (%)"
COL_INITIAL_WI = "Initial WI (%)"
COL_THREE_MONTH_RATE = "3 Month Avg Production (boepd)"

# Column names — forecasts
COL_FORECAST_MONTH = "month #"
COL_FORECAST_YEAR = "year"
COL_FORECAST_REVENUE = "total_revenue ($M)"
COL_FORECAST_OPERATING_INCOME = "operating_income ($M)"
COL_FORECAST_VOLUME = "boe"

MONEY_TO_DOLLARS = 1000.0

REQUIRED_WELLS_COLS = (
    COL_EVENT, COL_OLD_UWI_1, COL_OLD_UWI_2,
    COL_OLD_CURVE_1, COL_OLD_CURVE_2, COL_NEW_CURVE,
)
REQUIRED_INDICATOR_COLS = (
    COL_TYPE_CURVE, COL_NPV10, COL_NPV_INVESTMENT_RATIO, COL_PAYOUT,
    COL_RESERVES, COL_FIRST_YEAR_RATE, COL_COST_OF_RESERVES, COL_IP30,
    COL_CAPEX, COL_ROR, COL_INITIAL_WI, COL_THREE_MONTH_RATE,
)
REQUIRED_FORECAST_COLS = (
    COL_TYPE_CURVE, COL_FORECAST_MONTH, COL_FORECAST_YEAR,
    COL_FORECAST_VOLUME, COL_FORECAST_REVENUE, COL_FORECAST_OPERATING_INCOME,
)

EVENT_TYPE_MAP = {
    "Consolidation": "Existing Plan Optimization",
    "Extension": "Inventory Enhancement Identified",
    "Creation": "New Inventory Identified",
}

# ── Palette ──────────────────────────────────────────────────────────────────
PAL_CONSOLIDATION = "#2563EB"   # blue-600
PAL_EXTENSION = "#0D9488"       # teal-600
PAL_CREATION = "#7C3AED"        # violet-600
PAL_OLD = "#94A3B8"             # slate-400
PAL_NEW = "#1E40AF"             # blue-800
PAL_POSITIVE = "#16A34A"        # green-600
PAL_NEGATIVE = "#DC2626"        # red-600
PAL_BG = "#F8FAFC"              # slate-50

# Executive-summary development ladder
LADDER_CURVE_125 = "VF_BAK_2.0M_125 (June 2026)"
LADDER_CURVE_100 = "VF_BAK_2.0M_100 (June 2026)"
LADDER_CURVE_75 = "VF_BAK_2.0M_75 (June 2026)"
LADDER_WELLS_PER_YEAR = 22

CATEGORY_COLORS = {
    "Consolidation": PAL_CONSOLIDATION,
    "Extension": PAL_EXTENSION,
    "Creation": PAL_CREATION,
}


# ────────────────────────────────────────────────────────────────────────────────
# DATA LOADING & VALIDATION
# ────────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading economics.xlsx …")
def load_workbook(path: str, modified_time_ns: int) -> dict[str, pd.DataFrame]:
    return pd.read_excel(
        path,
        sheet_name=[SHEET_WELLS, SHEET_ECONOMIC_INDICATORS, SHEET_FORECASTS],
        engine="openpyxl",
    )


def validate_workbook_schema(
    wells: pd.DataFrame,
    indicators: pd.DataFrame,
    forecasts: pd.DataFrame,
) -> pd.DataFrame:
    """Run blocking & warning checks.  Stops the app on any BLOCKING failure."""
    rows: list[dict[str, Any]] = []
    blocking = False

    def _add(sev: str, check: str, status: str, n: int = 0, details: str = ""):
        nonlocal blocking
        rows.append(dict(severity=sev, check=check, status=status,
                         affected_count=n, details=details))
        if sev == "BLOCKING" and status == "FAIL":
            blocking = True

    # Column presence
    for label, df, req in [
        (SHEET_WELLS, wells, REQUIRED_WELLS_COLS),
        (SHEET_ECONOMIC_INDICATORS, indicators, REQUIRED_INDICATOR_COLS),
        (SHEET_FORECASTS, forecasts, REQUIRED_FORECAST_COLS),
    ]:
        missing = [c for c in req if c not in df.columns]
        if missing:
            _add("BLOCKING", f"{label} columns", "FAIL", len(missing),
                 f"Missing: {missing}")
        else:
            _add("INFO", f"{label} columns", "PASS")

    # event # integrity
    if wells[COL_EVENT].isna().any():
        _add("BLOCKING", "event # no nulls", "FAIL",
             int(wells[COL_EVENT].isna().sum()))
    else:
        _add("INFO", "event # no nulls", "PASS")

    if wells[COL_EVENT].duplicated().any():
        _add("BLOCKING", "event # unique", "FAIL",
             int(wells[COL_EVENT].duplicated().sum()))
    else:
        _add("INFO", "event # unique", "PASS")

    try:
        wells[COL_EVENT].astype(int)
        _add("INFO", "event # numeric", "PASS")
    except (ValueError, TypeError):
        _add("BLOCKING", "event # numeric", "FAIL",
             details="Cannot convert to int")

    if wells[COL_NEW_CURVE].isna().any():
        _add("BLOCKING", "new well type curve no nulls", "FAIL",
             int(wells[COL_NEW_CURVE].isna().sum()))
    else:
        _add("INFO", "new well type curve no nulls", "PASS")

    if indicators[COL_TYPE_CURVE].isna().any():
        _add("BLOCKING", "indicator type curve no nulls", "FAIL")
    if indicators[COL_TYPE_CURVE].duplicated().any():
        _add("BLOCKING", "indicator type curve unique", "FAIL")
    else:
        _add("INFO", "indicator type curve unique", "PASS")

    for c in REQUIRED_FORECAST_COLS:
        nn = int(forecasts[c].isna().sum())
        if nn > 0:
            _add("BLOCKING", f"forecasts {c} no nulls", "FAIL", nn)

    # Cross-sheet referential integrity
    ref_curves: set[str] = set()
    ref_curves.update(wells[COL_NEW_CURVE].dropna().unique())
    ref_curves.update(wells[COL_OLD_CURVE_1].dropna().unique())
    ref_curves.update(wells[COL_OLD_CURVE_2].dropna().unique())
    ind_curves = set(indicators[COL_TYPE_CURVE].dropna().unique())
    fc_curves = set(forecasts[COL_TYPE_CURVE].dropna().unique())

    missing_ind = ref_curves - ind_curves
    missing_fc = ref_curves - fc_curves
    if missing_ind:
        _add("BLOCKING", "ref curves in indicators", "FAIL",
             len(missing_ind), str(missing_ind))
    else:
        _add("INFO", "ref curves in indicators", "PASS")
    if missing_fc:
        _add("BLOCKING", "ref curves in forecasts", "FAIL",
             len(missing_fc), str(missing_fc))
    else:
        _add("INFO", "ref curves in forecasts", "PASS")

    # Forecast key uniqueness
    fc_key = forecasts[[COL_TYPE_CURVE, COL_FORECAST_YEAR, COL_FORECAST_MONTH]]
    dup_fc = fc_key.duplicated().sum()
    if dup_fc > 0:
        _add("BLOCKING", "forecast key unique", "FAIL", int(dup_fc))
    else:
        _add("INFO", "forecast key unique", "PASS")

    # Uniform forecast period counts
    ref_fc = forecasts[forecasts[COL_TYPE_CURVE].isin(ref_curves)]
    curve_counts = ref_fc.groupby(COL_TYPE_CURVE).size()
    if curve_counts.nunique() > 1:
        _add("BLOCKING", "forecast period count uniform", "FAIL",
             details=str(curve_counts.value_counts().to_dict()))
    else:
        cnt = curve_counts.iloc[0] if len(curve_counts) else 0
        _add("INFO", "forecast period count uniform", "PASS",
             details=f"{cnt} periods each")

    # Numeric checks on indicator columns
    num_cols = [
        COL_NPV10, COL_NPV_INVESTMENT_RATIO, COL_PAYOUT, COL_RESERVES,
        COL_FIRST_YEAR_RATE, COL_COST_OF_RESERVES, COL_IP30, COL_CAPEX,
        COL_ROR, COL_INITIAL_WI, COL_THREE_MONTH_RATE,
    ]
    for c in num_cols:
        if c in indicators.columns:
            converted = pd.to_numeric(indicators[c], errors="coerce")
            new_nulls = converted.isna().sum() - indicators[c].isna().sum()
            if new_nulls > 0:
                _add("BLOCKING", f"indicators {c} numeric", "FAIL",
                     int(new_nulls))

    report = pd.DataFrame(rows)
    if blocking:
        st.error("❌ Blocking validation errors — see details below.")
        for r in rows:
            if r["severity"] == "BLOCKING" and r["status"] == "FAIL":
                st.error(f"**{r['check']}**: {r['details']}")
        st.dataframe(report)
        st.stop()
    return report


# ────────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION & ENRICHMENT
# ────────────────────────────────────────────────────────────────────────────────
def _is_populated(val) -> bool:
    return pd.notna(val) and str(val).strip() != ""


def classify_events(wells: pd.DataFrame) -> pd.DataFrame:
    df = wells.copy()
    df[COL_EVENT] = df[COL_EVENT].astype(int)
    types, stories, old_counts, pop_curves = [], [], [], []
    for _, row in df.iterrows():
        c1 = _is_populated(row[COL_OLD_CURVE_1])
        c2 = _is_populated(row[COL_OLD_CURVE_2])
        if c1 and c2:
            t, cnt, curves = "Consolidation", 2, [row[COL_OLD_CURVE_1], row[COL_OLD_CURVE_2]]
        elif c1 or c2:
            t, cnt = "Extension", 1
            curves = [row[COL_OLD_CURVE_1]] if c1 else [row[COL_OLD_CURVE_2]]
        else:
            t, cnt, curves = "Creation", 0, []
        types.append(t)
        stories.append(EVENT_TYPE_MAP[t])
        old_counts.append(cnt)
        pop_curves.append(curves)
    df["event_type"] = types
    df["event_story"] = stories
    df["old_curve_count"] = old_counts
    df["populated_old_curves"] = pop_curves
    return df


def prepare_forecasts(forecasts: pd.DataFrame, ref_curves: set[str]) -> pd.DataFrame:
    df = forecasts[forecasts[COL_TYPE_CURVE].isin(ref_curves)].copy()
    df = df.sort_values([COL_TYPE_CURVE, COL_FORECAST_YEAR, COL_FORECAST_MONTH]).reset_index(drop=True)
    df["producing_month"] = df.groupby(COL_TYPE_CURVE).cumcount() + 1
    df["volume_boe"] = pd.to_numeric(df[COL_FORECAST_VOLUME], errors="coerce")
    df["total_revenue_dollars"] = pd.to_numeric(df[COL_FORECAST_REVENUE], errors="coerce") * MONEY_TO_DOLLARS
    df["operating_income_dollars"] = pd.to_numeric(df[COL_FORECAST_OPERATING_INCOME], errors="coerce") * MONEY_TO_DOLLARS
    for m in ["volume_boe", "total_revenue_dollars", "operating_income_dollars"]:
        df[f"cumulative_{m}"] = df.groupby(COL_TYPE_CURVE)[m].cumsum()
    return df


def build_type_curve_lifetime(fc: pd.DataFrame) -> pd.DataFrame:
    return fc.groupby(COL_TYPE_CURVE).agg(
        lifetime_volume_boe=("volume_boe", "sum"),
        lifetime_revenue_dollars=("total_revenue_dollars", "sum"),
        lifetime_operating_income_dollars=("operating_income_dollars", "sum"),
        forecast_periods=("producing_month", "count"),
    ).reset_index()


def prepare_indicators(indicators: pd.DataFrame) -> pd.DataFrame:
    df = indicators.copy()
    df["npv10_dollars"] = pd.to_numeric(df[COL_NPV10], errors="coerce") * MONEY_TO_DOLLARS
    df["investment_dollars"] = pd.to_numeric(df[COL_CAPEX], errors="coerce") * MONEY_TO_DOLLARS
    return df


# ────────────────────────────────────────────────────────────────────────────────
# CORE ECONOMICS BUILD
# ────────────────────────────────────────────────────────────────────────────────
def _safe_div(num: float, den: float) -> float:
    if den == 0 or np.isnan(den) or np.isnan(num):
        return np.nan
    return num / den


@st.cache_data(show_spinner="Calculating event economics …")
def build_event_economics(
    wells_classified: pd.DataFrame,
    indicators_clean: pd.DataFrame,
    lifetime: pd.DataFrame,
    _forecasts_clean: pd.DataFrame,          # unused but kept for cache key
) -> pd.DataFrame:
    ind = indicators_clean.set_index(COL_TYPE_CURVE)
    lt = lifetime.set_index(COL_TYPE_CURVE)
    records: list[dict[str, Any]] = []

    for _, row in wells_classified.iterrows():
        ev: dict[str, Any] = {}
        ev["event"] = int(row[COL_EVENT])
        ev["event_type"] = row["event_type"]
        ev["event_story"] = row["event_story"]
        ev[COL_OLD_UWI_1] = row[COL_OLD_UWI_1] if pd.notna(row[COL_OLD_UWI_1]) else None
        ev[COL_OLD_UWI_2] = row[COL_OLD_UWI_2] if pd.notna(row[COL_OLD_UWI_2]) else None
        ev[COL_OLD_CURVE_1] = row[COL_OLD_CURVE_1] if pd.notna(row[COL_OLD_CURVE_1]) else None
        ev[COL_OLD_CURVE_2] = row[COL_OLD_CURVE_2] if pd.notna(row[COL_OLD_CURVE_2]) else None
        ev[COL_NEW_CURVE] = row[COL_NEW_CURVE]

        old_curves: list[str] = row["populated_old_curves"]
        new_curve: str = row[COL_NEW_CURVE]

        # Old-plan aggregation
        old_capex = sum(ind.loc[c, "investment_dollars"] for c in old_curves) if old_curves else 0.0
        old_npv10 = sum(ind.loc[c, "npv10_dollars"] for c in old_curves) if old_curves else 0.0
        old_reserves = sum(ind.loc[c, COL_RESERVES] for c in old_curves) if old_curves else 0.0
        old_lt_rev = sum(lt.loc[c, "lifetime_revenue_dollars"] for c in old_curves) if old_curves else 0.0
        old_lt_oi = sum(lt.loc[c, "lifetime_operating_income_dollars"] for c in old_curves) if old_curves else 0.0
        old_lt_vol = sum(lt.loc[c, "lifetime_volume_boe"] for c in old_curves) if old_curves else 0.0

        # New-plan values
        new_capex = ind.loc[new_curve, "investment_dollars"]
        new_npv10 = ind.loc[new_curve, "npv10_dollars"]
        new_reserves = ind.loc[new_curve, COL_RESERVES]
        new_lt_rev = lt.loc[new_curve, "lifetime_revenue_dollars"]
        new_lt_oi = lt.loc[new_curve, "lifetime_operating_income_dollars"]
        new_lt_vol = lt.loc[new_curve, "lifetime_volume_boe"]
        new_payout = ind.loc[new_curve, COL_PAYOUT]
        new_ror = ind.loc[new_curve, COL_ROR]
        new_npv_inv_ratio = ind.loc[new_curve, COL_NPV_INVESTMENT_RATIO]
        new_cor = ind.loc[new_curve, COL_COST_OF_RESERVES]

        # Deltas
        ev["old_capex_dollars"] = old_capex
        ev["new_capex_dollars"] = new_capex
        ev["capex_delta_dollars"] = new_capex - old_capex

        ev["old_npv10_dollars"] = old_npv10
        ev["new_npv10_dollars"] = new_npv10
        ev["npv10_delta_dollars"] = new_npv10 - old_npv10

        ev["old_reserves_boe"] = old_reserves
        ev["new_reserves_boe"] = new_reserves
        ev["reserves_delta_boe"] = new_reserves - old_reserves

        ev["old_lifetime_revenue_dollars"] = old_lt_rev
        ev["new_lifetime_revenue_dollars"] = new_lt_rev

        ev["old_lifetime_oi_dollars"] = old_lt_oi
        ev["new_lifetime_oi_dollars"] = new_lt_oi

        ev["old_lifetime_volume_boe"] = old_lt_vol
        ev["new_lifetime_volume_boe"] = new_lt_vol
        ev["lifetime_volume_delta_boe"] = new_lt_vol - old_lt_vol

        ev["new_payout_years"] = new_payout
        ev["new_ror_pct"] = new_ror
        ev["new_npv_inv_ratio"] = new_npv_inv_ratio
        ev["new_cost_of_reserves"] = new_cor

        # Derived per-unit metrics
        ev["old_cost_of_reserves_derived"] = _safe_div(old_capex, old_reserves)
        ev["new_cost_of_reserves_derived"] = _safe_div(new_capex, new_reserves)
        ev["cost_of_reserves_delta"] = (
            _safe_div(new_capex, new_reserves) - _safe_div(old_capex, old_reserves)
            if old_reserves > 0 else np.nan
        )
        ev["cost_of_reserves_improvement"] = (
            _safe_div(old_capex, old_reserves) - _safe_div(new_capex, new_reserves)
            if old_reserves > 0 and new_reserves > 0 else np.nan
        )

        # Consolidation specifics
        is_consol = ev["event_type"] == "Consolidation"
        ev["capital_saved_dollars"] = max(old_capex - new_capex, 0.0) if is_consol else 0.0
        ev["npv10_sacrificed_dollars"] = max(old_npv10 - new_npv10, 0.0) if is_consol else 0.0

        # Extension incremental efficiency
        inc_cap = new_capex - old_capex
        inc_npv = new_npv10 - old_npv10
        inc_res = new_reserves - old_reserves
        ev["incremental_capex_dollars"] = inc_cap
        ev["incremental_npv10_dollars"] = inc_npv
        ev["incremental_reserves_boe"] = inc_res
        ev["marginal_npv_per_inc_capital"] = _safe_div(inc_npv, inc_cap) if inc_cap > 0 else np.nan
        ev["inc_capital_per_inc_boe"] = _safe_div(inc_cap, inc_res) if inc_res > 0 else np.nan

        # Ratio helpers
        ev["old_npv_inv_ratio_derived"] = _safe_div(old_npv10, old_capex)
        ev["new_npv_inv_ratio_derived"] = _safe_div(new_npv10, new_capex)

        ev["old_type_curves_used"] = " | ".join(old_curves) if old_curves else ""
        ev["new_type_curve_used"] = new_curve
        records.append(ev)

    return pd.DataFrame(records)


# ────────────────────────────────────────────────────────────────────────────────
# EVENT FORECASTS
# ────────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Building event forecasts …")
def build_event_forecasts(
    wells_classified: pd.DataFrame,
    forecasts_clean: pd.DataFrame,
) -> pd.DataFrame:
    all_series: list[pd.DataFrame] = []
    base_cols = [
        "producing_month", COL_FORECAST_YEAR, COL_FORECAST_MONTH,
        "volume_boe", "total_revenue_dollars", "operating_income_dollars",
    ]

    for _, row in wells_classified.iterrows():
        ev_num = int(row[COL_EVENT])
        ev_type = row["event_type"]
        ev_story = row["event_story"]
        old_curves: list[str] = row["populated_old_curves"]
        new_curve: str = row[COL_NEW_CURVE]

        # Old-plan series
        if ev_type == "Consolidation":
            dfs_old = []
            for idx, c in enumerate(old_curves, 1):
                fc = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == c].sort_values("producing_month")[base_cols].copy()
                fc["series_key"] = f"old_{idx}"
                fc["series_label"] = f"Old Well {idx}: {c}"
                fc["plan_side"] = "old"
                dfs_old.append(fc)
            if len(dfs_old) == 2:
                comb = dfs_old[0][base_cols].copy().set_index("producing_month")
                comb2 = dfs_old[1][base_cols].copy().set_index("producing_month")
                combined = comb.copy()
                for mc in ["volume_boe", "total_revenue_dollars", "operating_income_dollars"]:
                    combined[mc] = comb[mc].values + comb2[mc].values
                combined = combined.reset_index()
                combined["series_key"] = "combined_old"
                combined["series_label"] = "Combined Old Plan"
                combined["plan_side"] = "old"
                dfs_old.append(combined)
            for d in dfs_old:
                d["event"] = ev_num
                d["event_type"] = ev_type
                d["event_story"] = ev_story
            all_series.extend(dfs_old)

        elif ev_type == "Extension":
            fc = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == old_curves[0]].sort_values("producing_month")[base_cols].copy()
            fc["series_key"] = "old_ref"
            fc["series_label"] = f"1-Mile Reference: {old_curves[0]}"
            fc["plan_side"] = "old"
            fc["event"] = ev_num
            fc["event_type"] = ev_type
            fc["event_story"] = ev_story
            all_series.append(fc)

        # New-plan series
        fc_new = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == new_curve].sort_values("producing_month")[base_cols].copy()
        if ev_type == "Consolidation":
            fc_new["series_key"] = "new_plan"
            fc_new["series_label"] = f"2-Mile Consolidated: {new_curve}"
        elif ev_type == "Extension":
            fc_new["series_key"] = "new_opp"
            fc_new["series_label"] = f"2-Mile Extended: {new_curve}"
        else:
            fc_new["series_key"] = "new_inventory"
            fc_new["series_label"] = f"New Well: {new_curve}"
        fc_new["plan_side"] = "new"
        fc_new["event"] = ev_num
        fc_new["event_type"] = ev_type
        fc_new["event_story"] = ev_story
        all_series.append(fc_new)

    if not all_series:
        return pd.DataFrame()
    result = pd.concat(all_series, ignore_index=True)
    for metric in ["volume_boe", "total_revenue_dollars", "operating_income_dollars"]:
        result[f"cumulative_{metric}"] = result.groupby(["event", "series_key"])[metric].cumsum()
    return result


def build_event_annual_forecasts(event_forecasts: pd.DataFrame) -> pd.DataFrame:
    if event_forecasts.empty:
        return pd.DataFrame()
    group_cols = [
        "event", "event_type", "event_story",
        "series_key", "series_label", "plan_side", COL_FORECAST_YEAR,
    ]
    return event_forecasts.groupby(group_cols, as_index=False).agg(
        volume_boe=("volume_boe", "sum"),
        total_revenue_dollars=("total_revenue_dollars", "sum"),
        operating_income_dollars=("operating_income_dollars", "sum"),
    )


# ────────────────────────────────────────────────────────────────────────────────
# FORMATTING HELPERS
# ────────────────────────────────────────────────────────────────────────────────
def _f_mm(v, prefix="$", suffix="MM"):
    """Format dollars → $X.X MM"""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    mm = v / 1_000_000
    return f"({prefix}{abs(mm):,.1f} {suffix})" if mm < 0 else f"{prefix}{mm:,.1f} {suffix}"


def _f_mm_signed(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    mm = v / 1_000_000
    sign = "+" if mm >= 0 else ""
    return f"{sign}${mm:,.1f} MM"


def _f_mboe(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v / 1_000:,.1f} Mboe"


def _f_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:,.1f}%"


def _f_ratio(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/M"
    return f"{v:,.2f}x"


def _f_dollarboe(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"${v:,.2f}/boe"


def _f_years(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:,.1f} yrs"


def _f_rate(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:,.0f} boe/d"


# ────────────────────────────────────────────────────────────────────────────────
# CHART HELPERS
# ────────────────────────────────────────────────────────────────────────────────
def _apply_layout(fig: go.Figure, title: str = "", xaxis: str = "", yaxis: str = ""):
    fig.update_layout(
        title=dict(text=title, font=dict(size=16)),
        xaxis_title=xaxis,
        yaxis_title=yaxis,
        template="plotly_white",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
        margin=dict(l=60, r=30, t=50, b=40),
        plot_bgcolor="white",
        font=dict(size=12),
    )
    return fig


# ────────────────────────────────────────────────────────────────────────────────
# EXPORT HELPERS
# ────────────────────────────────────────────────────────────────────────────────
def _df_to_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def _dfs_to_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False, freeze_panes=(1, 0))
    buf.seek(0)
    return buf.read()


def _fig_to_html(fig: go.Figure) -> bytes:
    return fig.to_html(include_plotlyjs="cdn").encode("utf-8")


def _fig_to_png(fig: go.Figure) -> bytes | None:
    try:
        return fig.to_image(format="png", width=1200, height=600, scale=2)
    except Exception:
        return None


def _build_zip(files: dict[str, bytes]) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    buf.seek(0)
    return buf.read()


# ════════════════════════════════════════════════════════════════════════════════
# PAGE RENDERERS
# ════════════════════════════════════════════════════════════════════════════════


# ── 1. PORTFOLIO OVERVIEW ─────────────────────────────────────────────────────
def render_portfolio_overview(econ: pd.DataFrame):
    """Executive summary focused on annual capital savings and value/production uplift."""
    st.title("📊 Executive Summary")
    st.caption("Inventory sequencing and annualized value impact from the VFB 2026 inventory overhaul")

    # Default user view excludes the 75 curve, while keeping it available in the filter.
    all_curves = sorted(econ[COL_NEW_CURVE].dropna().unique())
    default_curves = [c for c in all_curves if c != LADDER_CURVE_75]
    with st.expander("🔍 Filters", expanded=False):
        sel_curves = st.multiselect(
            "New Type Curves", all_curves, default=default_curves, key="po_curves",
            help=f"{LADDER_CURVE_75} is excluded by default but can be added here.",
        )
    filtered = econ[econ[COL_NEW_CURVE].isin(sel_curves)].copy()

    st.subheader("Development Ladder — 22 Wells / Year")
    st.markdown(
        f"**Sequence:** exhaust **{LADDER_CURVE_125}** first, then develop "
        f"**{LADDER_CURVE_100}** at **{LADDER_WELLS_PER_YEAR} wells/year**. "
        f"**{LADDER_CURVE_75}** is held outside the default user view."
    )

    ladder_rows = []
    for priority, curve in enumerate([LADDER_CURVE_125, LADDER_CURVE_100], 1):
        d = filtered[filtered[COL_NEW_CURVE] == curve]
        wells = len(d)
        ladder_rows.append({
            "Priority": priority,
            "Type Curve": curve,
            "Inventory (wells)": wells,
            "Years @ 22 wells/yr": wells / LADDER_WELLS_PER_YEAR if wells else 0.0,
            "Capital Saved ($MM)": d["capital_saved_dollars"].sum() / 1e6,
            "NPV10 Added ($MM)": d["npv10_delta_dollars"].sum() / 1e6,
            "Production Added (Mboe)": d["lifetime_volume_delta_boe"].sum() / 1e3,
        })
    ladder = pd.DataFrame(ladder_rows)
    st.dataframe(
        ladder, use_container_width=True, hide_index=True,
        column_config={
            "Years @ 22 wells/yr": st.column_config.NumberColumn(format="%.1f"),
            "Capital Saved ($MM)": st.column_config.NumberColumn(format="$%.1f"),
            "NPV10 Added ($MM)": st.column_config.NumberColumn(format="$%.1f"),
            "Production Added (Mboe)": st.column_config.NumberColumn(format="%.1f"),
        },
    )

    # Annualized impact at the stated 22-well development pace.
    annual_rows = []
    for curve in [LADDER_CURVE_125, LADDER_CURVE_100]:
        d = filtered[filtered[COL_NEW_CURVE] == curve]
        if d.empty:
            continue
        annual_rows.append({
            "Type Curve": curve,
            "Capital Saved / Year ($MM)": d["capital_saved_dollars"].mean() * LADDER_WELLS_PER_YEAR / 1e6,
            "NPV10 Added / Year ($MM)": d["npv10_delta_dollars"].mean() * LADDER_WELLS_PER_YEAR / 1e6,
            "Production Added / Year (Mboe)": d["lifetime_volume_delta_boe"].mean() * LADDER_WELLS_PER_YEAR / 1e3,
        })
    annual = pd.DataFrame(annual_rows)

    if not annual.empty:
        active_curve = annual.iloc[0]["Type Curve"]
        active = annual.iloc[0]
        st.markdown(f"#### Annual Impact at 22 Wells — {active_curve}")
        k1, k2, k3 = st.columns(3)
        k1.metric("Capital Saved / Year", f"${active['Capital Saved / Year ($MM)']:,.1f} MM")
        k2.metric("NPV10 Added / Year", f"{active['NPV10 Added / Year ($MM)']:+,.1f} MM")
        k3.metric("Production Added / Year", f"{active['Production Added / Year (Mboe)']:+,.1f} Mboe")

        fig = go.Figure()
        fig.add_trace(go.Bar(name="Capital Saved / Year", x=annual["Type Curve"], y=annual["Capital Saved / Year ($MM)"]))
        fig.add_trace(go.Bar(name="NPV10 Added / Year", x=annual["Type Curve"], y=annual["NPV10 Added / Year ($MM)"]))
        _apply_layout(fig, "Annual Capital Saved and NPV10 Added at 22 Wells / Year", "Development Tier", "$MM / year")
        fig.update_layout(barmode="group")
        st.plotly_chart(fig, use_container_width=True)

        fig_prod = go.Figure(go.Bar(
            x=annual["Type Curve"], y=annual["Production Added / Year (Mboe)"],
            text=annual["Production Added / Year (Mboe)"].map(lambda x: f"{x:+,.1f}"),
            textposition="outside",
        ))
        _apply_layout(fig_prod, "Production Added at 22 Wells / Year", "Development Tier", "Mboe / year")
        st.plotly_chart(fig_prod, use_container_width=True)

    with st.expander("Portfolio totals for selected curves", expanded=False):
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Capital Saved", _f_mm(filtered["capital_saved_dollars"].sum()))
        c2.metric("Total Δ NPV10", _f_mm_signed(filtered["npv10_delta_dollars"].sum()))
        c3.metric("Total Δ Production", _f_mboe(filtered["lifetime_volume_delta_boe"].sum()))

# ── 2. CONSOLIDATION DEEP-DIVE ───────────────────────────────────────────────
def render_consolidation(econ: pd.DataFrame, event_forecasts: pd.DataFrame):
    st.title("🔵 Consolidation — Capital Removal")
    st.markdown(
        "> **Thesis:** A single 2-mile well is cheaper on a per-barrel basis "
        "than drilling its two constituent 1-mile wells."
    )

    consol = econ[econ["event_type"] == "Consolidation"].copy()
    if consol.empty:
        st.info("No consolidation events found.")
        return

    # Filters
    with st.expander("🔍 Filters", expanded=False):
        curves = sorted(consol[COL_NEW_CURVE].dropna().unique())
        sel_curves = st.multiselect("New Type Curves", curves, default=curves, key="con_nc")
        events = sorted(consol["event"].unique())
        sel_events = st.multiselect("Event #s (leave blank = all)", events, default=[], key="con_ev")
    mask = consol[COL_NEW_CURVE].isin(sel_curves)
    if sel_events:
        mask = mask & consol["event"].isin(sel_events)
    df = consol[mask]

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Events", f"{len(df)}")
    c2.metric("Total Capital Saved", _f_mm(df["capital_saved_dollars"].sum()))
    c3.metric("Avg $/boe Improvement", _f_dollarboe(df["cost_of_reserves_improvement"].mean()))
    c4.metric("Avg NPV/Investment", _f_ratio(df["new_npv_inv_ratio"].mean()))

    # Cost of reserves comparison chart
    st.markdown("#### Cost of Reserves: Old Plan vs. New Plan")
    bar_df = df[["event", "old_cost_of_reserves_derived", "new_cost_of_reserves_derived"]].copy()
    bar_df = bar_df.sort_values("old_cost_of_reserves_derived", ascending=False)
    bar_df["event_label"] = bar_df["event"].astype(str)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Old Plan (2 × 1-mile)",
        x=bar_df["event_label"],
        y=bar_df["old_cost_of_reserves_derived"],
        marker_color=PAL_OLD,
    ))
    fig.add_trace(go.Bar(
        name="New Plan (1 × 2-mile)",
        x=bar_df["event_label"],
        y=bar_df["new_cost_of_reserves_derived"],
        marker_color=PAL_CONSOLIDATION,
    ))
    _apply_layout(fig, "Cost of Reserves by Event ($/boe)", "Event #", "$/boe")
    fig.update_layout(barmode="group")
    st.plotly_chart(fig, use_container_width=True)

    # Capital saved bar chart
    st.markdown("#### Capital Saved per Event")
    cap_df = df[["event", "capital_saved_dollars"]].copy().sort_values("capital_saved_dollars", ascending=False)
    cap_df["event_label"] = cap_df["event"].astype(str)
    fig2 = go.Figure(go.Bar(
        x=cap_df["event_label"],
        y=cap_df["capital_saved_dollars"] / 1e6,
        marker_color=PAL_POSITIVE,
        text=cap_df["capital_saved_dollars"].apply(lambda v: f"${v/1e6:,.1f}M"),
        textposition="outside",
    ))
    _apply_layout(fig2, "Capital Saved ($MM)", "Event #", "$MM")
    st.plotly_chart(fig2, use_container_width=True)

    # Detail table
    with st.expander("📋 Event Detail Table", expanded=False):
        detail = df[[
            "event", COL_OLD_UWI_1, COL_OLD_UWI_2, COL_NEW_CURVE,
            "old_capex_dollars", "new_capex_dollars", "capital_saved_dollars",
            "old_cost_of_reserves_derived", "new_cost_of_reserves_derived",
            "cost_of_reserves_improvement",
            "new_npv_inv_ratio", "new_reserves_boe",
        ]].copy()
        detail.columns = [
            "Event #", "Old UWI 1", "Old UWI 2", "New Curve",
            "Old Capex ($)", "New Capex ($)", "Capital Saved ($)",
            "Old $/boe", "New $/boe", "$/boe Improvement",
            "NPV/Invest", "Reserves (boe)",
        ]
        st.dataframe(detail.sort_values("Event #"), use_container_width=True, hide_index=True)


# ── 3. EXTENSION DEEP-DIVE ───────────────────────────────────────────────────
def render_extension(econ: pd.DataFrame, event_forecasts: pd.DataFrame):
    st.title("🟢 Extension — Capital Efficiency")
    st.markdown(
        "> **Thesis:** Extending a 1-mile well to 2 miles generates outsized "
        "incremental value for relatively little additional capital."
    )

    ext = econ[econ["event_type"] == "Extension"].copy()
    if ext.empty:
        st.info("No extension events found.")
        return

    # Filters
    with st.expander("🔍 Filters", expanded=False):
        curves = sorted(ext[COL_NEW_CURVE].dropna().unique())
        sel_curves = st.multiselect("New Type Curves", curves, default=curves, key="ext_nc")
        events = sorted(ext["event"].unique())
        sel_events = st.multiselect("Event #s (leave blank = all)", events, default=[], key="ext_ev")
    mask = ext[COL_NEW_CURVE].isin(sel_curves)
    if sel_events:
        mask = mask & ext["event"].isin(sel_events)
    df = ext[mask]

    # KPIs
    inc_npv = df["incremental_npv10_dollars"].sum()
    inc_cap = df["incremental_capex_dollars"].sum()
    inc_res = df["incremental_reserves_boe"].sum()
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Events", f"{len(df)}")
    e2.metric("Incremental NPV10", _f_mm_signed(inc_npv))
    e3.metric("Incremental Capex", _f_mm(inc_cap))
    e4.metric(
        "Marginal NPV / $ Invested",
        _f_ratio(_safe_div(inc_npv, inc_cap)),
        help="Every $1 of incremental capital generates this much NPV10",
    )

    # Stacked bar: old NPV + incremental NPV
    st.markdown("#### Value Build-Up: Base 1-Mile NPV10 + Incremental from Extension")
    chart_df = df[["event", "old_npv10_dollars", "incremental_npv10_dollars"]].copy()
    chart_df = chart_df.sort_values("incremental_npv10_dollars", ascending=False)
    chart_df["event_label"] = chart_df["event"].astype(str)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="1-Mile Base NPV10",
        x=chart_df["event_label"],
        y=chart_df["old_npv10_dollars"] / 1e6,
        marker_color=PAL_OLD,
    ))
    fig.add_trace(go.Bar(
        name="Incremental NPV10 from Extension",
        x=chart_df["event_label"],
        y=chart_df["incremental_npv10_dollars"] / 1e6,
        marker_color=PAL_EXTENSION,
    ))
    _apply_layout(fig, "NPV10 Build-Up ($MM)", "Event #", "$MM")
    fig.update_layout(barmode="stack")
    st.plotly_chart(fig, use_container_width=True)

    # Detail table
    with st.expander("📋 Event Detail Table", expanded=False):
        detail = df[[
            "event", COL_OLD_CURVE_1, COL_NEW_CURVE,
            "old_capex_dollars", "new_capex_dollars", "incremental_capex_dollars",
            "old_npv10_dollars", "new_npv10_dollars", "incremental_npv10_dollars",
            "marginal_npv_per_inc_capital",
            "old_reserves_boe", "new_reserves_boe", "incremental_reserves_boe",
        ]].copy()
        detail.columns = [
            "Event #", "Old Curve", "New Curve",
            "Old Capex ($)", "New Capex ($)", "Δ Capex ($)",
            "Old NPV10 ($)", "New NPV10 ($)", "Δ NPV10 ($)",
            "Marginal NPV/$",
            "Old Reserves", "New Reserves", "Δ Reserves",
        ]
        st.dataframe(detail.sort_values("Event #"), use_container_width=True, hide_index=True)


# ── 4. CREATION DEEP-DIVE ────────────────────────────────────────────────────
def render_creation(econ: pd.DataFrame):
    st.title("🟣 Creation — New Value")
    st.markdown(
        "> **Thesis:** New well locations create incremental value "
        "for the portfolio — pure additions to inventory."
    )

    cre = econ[econ["event_type"] == "Creation"].copy()
    if cre.empty:
        st.info("No creation events found.")
        return

    # Filters
    with st.expander("🔍 Filters", expanded=False):
        curves = sorted(cre[COL_NEW_CURVE].dropna().unique())
        sel_curves = st.multiselect("New Type Curves", curves, default=curves, key="cre_nc")
        events = sorted(cre["event"].unique())
        sel_events = st.multiselect("Event #s (leave blank = all)", events, default=[], key="cre_ev")
    mask = cre[COL_NEW_CURVE].isin(sel_curves)
    if sel_events:
        mask = mask & cre["event"].isin(sel_events)
    df = cre[mask]

    # KPIs
    n1, n2, n3, n4 = st.columns(4)
    n1.metric("New Locations", f"{len(df)}")
    n2.metric("Total NPV10 Added", _f_mm(df["new_npv10_dollars"].sum()))
    n3.metric("Total Reserves Added", _f_mboe(df["new_reserves_boe"].sum()))
    n4.metric("Avg ROR", _f_pct(df["new_ror_pct"].mean()))

    # NPV10 bar chart
    st.markdown("#### NPV10 by New Well")
    chart_df = df[["event", "new_npv10_dollars", COL_NEW_CURVE]].copy()
    chart_df = chart_df.sort_values("new_npv10_dollars", ascending=False)
    chart_df["event_label"] = chart_df["event"].astype(str)
    fig = go.Figure(go.Bar(
        x=chart_df["event_label"],
        y=chart_df["new_npv10_dollars"] / 1e6,
        marker_color=PAL_CREATION,
        text=chart_df["new_npv10_dollars"].apply(lambda v: f"${v/1e6:,.1f}M"),
        textposition="outside",
    ))
    _apply_layout(fig, "NPV10 Added per New Well ($MM)", "Event #", "$MM")
    st.plotly_chart(fig, use_container_width=True)

    # Detail table
    with st.expander("📋 Event Detail Table", expanded=False):
        detail = df[[
            "event", COL_NEW_CURVE,
            "new_capex_dollars", "new_npv10_dollars",
            "new_npv_inv_ratio", "new_ror_pct",
            "new_payout_years", "new_reserves_boe",
            "new_cost_of_reserves",
        ]].copy()
        detail.columns = [
            "Event #", "Type Curve",
            "Capex ($)", "NPV10 ($)",
            "NPV/Invest", "ROR (%)",
            "Payout (yrs)", "Reserves (boe)",
            "Cost of Reserves ($/boe)",
        ]
        st.dataframe(detail.sort_values("Event #"), use_container_width=True, hide_index=True)


# ── 5. EVENT EXPLORER ─────────────────────────────────────────────────────────
def render_event_explorer(
    econ: pd.DataFrame,
    wells_classified: pd.DataFrame,
    indicators_clean: pd.DataFrame,
    forecasts_clean: pd.DataFrame,
    event_forecasts: pd.DataFrame,
    event_annual: pd.DataFrame,
):
    st.title("🔎 Event Explorer")
    st.caption("Drill into any individual event: forecasts, economics, and source data.")

    fc1, fc2 = st.columns([1, 3])
    event_type_filter = fc1.selectbox(
        "Event Type", ["All", "Consolidation", "Extension", "Creation"], key="ee_et"
    )
    if event_type_filter == "All":
        avail = sorted(econ["event"].unique())
    else:
        avail = sorted(econ[econ["event_type"] == event_type_filter]["event"].unique())
    if not avail:
        st.info("No events match filter.")
        return

    def _label(ev_num: int) -> str:
        r = econ[econ["event"] == ev_num].iloc[0]
        et = r["event_type"]
        u1, u2 = r.get(COL_OLD_UWI_1), r.get(COL_OLD_UWI_2)
        uwis = [str(v) for v in (u1, u2) if pd.notna(v) and str(v).strip()]
        if et == "Consolidation":
            return f"#{ev_num} Consol — {' + '.join(uwis) if uwis else r['new_type_curve_used']}"
        if et == "Extension":
            return f"#{ev_num} Ext — {uwis[0] if uwis else r['new_type_curve_used']}"
        return f"#{ev_num} New — {r['new_type_curve_used']}"

    sel_event = fc2.selectbox("Select Event", avail, format_func=_label, key="ee_ev")
    ev = econ[econ["event"] == sel_event].iloc[0]

    # Header
    color = CATEGORY_COLORS.get(ev["event_type"], "#333")
    st.markdown(f"### Event #{sel_event} — {ev['event_type']}")
    st.caption(ev["event_story"])

    # Identity row
    id1, id2, id3 = st.columns(3)
    id1.markdown(f"**Old Curves:** {ev['old_type_curves_used'] or '—'}")
    id2.markdown(f"**New Curve:** {ev['new_type_curve_used']}")
    uwi_str = ", ".join(
        str(v) for v in [ev[COL_OLD_UWI_1], ev[COL_OLD_UWI_2]]
        if v is not None
    )
    id3.markdown(f"**UWIs:** {uwi_str or '—'}")

    # Narrative-specific KPIs
    st.markdown("---")
    if ev["event_type"] == "Consolidation":
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Capital Saved", _f_mm(ev["capital_saved_dollars"]))
        k2.metric("$/boe Improvement", _f_dollarboe(ev["cost_of_reserves_improvement"]))
        k3.metric("NPV/Investment", _f_ratio(ev["new_npv_inv_ratio"]))
        k4.metric("Payout", _f_years(ev["new_payout_years"]))
    elif ev["event_type"] == "Extension":
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Δ NPV10", _f_mm_signed(ev["incremental_npv10_dollars"]))
        k2.metric("Δ Capex", _f_mm(ev["incremental_capex_dollars"]))
        k3.metric("Marginal NPV / $", _f_ratio(ev["marginal_npv_per_inc_capital"]))
        k4.metric("Δ Reserves", _f_mboe(ev["incremental_reserves_boe"]))
    else:
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("NPV10", _f_mm(ev["new_npv10_dollars"]))
        k2.metric("Capex", _f_mm(ev["new_capex_dollars"]))
        k3.metric("ROR", _f_pct(ev["new_ror_pct"]))
        k4.metric("Payout", _f_years(ev["new_payout_years"]))

    # Forecast charts
    ev_fc = event_forecasts[event_forecasts["event"] == sel_event]
    ev_ann = event_annual[event_annual["event"] == sel_event]

    if not ev_fc.empty:
        st.markdown("---")
        st.markdown("#### Production & Revenue Forecasts")
        mc1, mc2 = st.columns(2)
        metric_map = {
            "Volume (boe)": ("volume_boe", "cumulative_volume_boe", "boe"),
            "Revenue ($)": ("total_revenue_dollars", "cumulative_total_revenue_dollars", "$MM"),
            "Operating Income ($)": ("operating_income_dollars", "cumulative_operating_income_dollars", "$MM"),
        }
        sel_metric = mc1.selectbox("Metric", list(metric_map.keys()), key="ee_fm")
        sel_view = mc2.selectbox("View", ["Monthly", "Cumulative", "Annual"], key="ee_fv")
        monthly_col, cum_col, unit = metric_map[sel_metric]
        scale = 1.0 if unit == "boe" else 1e6

        if sel_view == "Monthly":
            fig = px.line(
                ev_fc, x="producing_month", y=ev_fc[monthly_col] / scale,
                color="series_label",
                labels={"producing_month": "Producing Month", "y": f"{sel_metric}"},
            )
            _apply_layout(fig, f"Monthly {sel_metric}", "Producing Month", f"{sel_metric}")
        elif sel_view == "Cumulative":
            fig = px.line(
                ev_fc, x="producing_month", y=ev_fc[cum_col] / scale,
                color="series_label",
            )
            _apply_layout(fig, f"Cumulative {sel_metric}", "Producing Month", f"Cum. {sel_metric}")
        else:
            fig = px.bar(
                ev_ann, x=COL_FORECAST_YEAR, y=ev_ann[monthly_col] / scale,
                color="series_label", barmode="group",
            )
            _apply_layout(fig, f"Annual {sel_metric}", "Year", f"{sel_metric}")

        # Style traces
        for trace in fig.data:
            name = trace.name
            if "Combined" in name or "Old" in name or "1-Mile" in name:
                trace.update(line=dict(dash="dash", width=1.5))
            elif "New" in name or "2-Mile" in name:
                trace.update(line=dict(width=3))
        st.plotly_chart(fig, use_container_width=True)

        # Chart download
        col_dl1, col_dl2, _ = st.columns([1, 1, 4])
        col_dl1.download_button(
            "📥 Chart HTML", _fig_to_html(fig),
            f"event_{sel_event}_chart.html", "text/html", key="ee_html",
        )
        img = _fig_to_png(fig)
        if img:
            col_dl2.download_button(
                "📥 Chart PNG", img,
                f"event_{sel_event}_chart.png", "image/png", key="ee_png",
            )

    # Source data
    with st.expander("📄 Source Data", expanded=False):
        curves_used = {ev[COL_NEW_CURVE]}
        if pd.notna(ev.get(COL_OLD_CURVE_1)):
            curves_used.add(ev[COL_OLD_CURVE_1])
        if pd.notna(ev.get(COL_OLD_CURVE_2)):
            curves_used.add(ev[COL_OLD_CURVE_2])
        curves_used.discard(None)

        st.markdown("**Wells Row**")
        st.dataframe(
            wells_classified[wells_classified[COL_EVENT] == sel_event],
            use_container_width=True, hide_index=True,
        )
        st.markdown("**Economic Indicators**")
        st.dataframe(
            indicators_clean[indicators_clean[COL_TYPE_CURVE].isin(curves_used)],
            use_container_width=True, hide_index=True,
        )
        st.markdown("**Forecast Series**")
        st.dataframe(ev_fc, use_container_width=True, hide_index=True)

    # Event downloads
    with st.expander("📥 Event Downloads", expanded=False):
        ev_record = econ[econ["event"] == sel_event]
        d1, d2, d3 = st.columns(3)
        d1.download_button(
            "Summary CSV", _df_to_csv(ev_record),
            f"event_{sel_event}_summary.csv", "text/csv", key="ee_dl_sum",
        )
        d2.download_button(
            "Monthly Forecast CSV", _df_to_csv(ev_fc),
            f"event_{sel_event}_monthly.csv", "text/csv", key="ee_dl_mfc",
        )
        ind_rows = indicators_clean[indicators_clean[COL_TYPE_CURVE].isin(curves_used)]
        sheets = {
            "Summary": ev_record,
            "Monthly Forecasts": ev_fc,
            "Annual Forecasts": ev_ann,
            "Source Indicators": ind_rows,
        }
        d3.download_button(
            "Workbook XLSX", _dfs_to_xlsx(sheets),
            f"event_{sel_event}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="ee_dl_xlsx",
        )


# ── 6. DATA & DOWNLOADS ──────────────────────────────────────────────────────
def render_downloads(
    econ: pd.DataFrame,
    event_forecasts: pd.DataFrame,
    event_annual: pd.DataFrame,
    validation_report: pd.DataFrame,
    wells_classified: pd.DataFrame,
    indicators_clean: pd.DataFrame,
    forecasts_clean: pd.DataFrame,
):
    st.title("📥 Data & Downloads")

    st.subheader("Validation Report")
    with st.expander("Show validation results", expanded=False):
        st.dataframe(validation_report, use_container_width=True, hide_index=True)

    st.subheader("Portfolio Exports")
    csv_files: dict[str, bytes] = {}
    c1, c2, c3 = st.columns(3)

    b1 = _df_to_csv(econ)
    csv_files["event_economics.csv"] = b1
    c1.download_button("Event Economics CSV", b1, "event_economics.csv", "text/csv", key="dl_ee")

    b2 = _df_to_csv(event_forecasts)
    csv_files["monthly_forecasts.csv"] = b2
    c2.download_button("Monthly Forecasts CSV", b2, "monthly_forecasts.csv", "text/csv", key="dl_mf")

    b3 = _df_to_csv(event_annual)
    csv_files["annual_forecasts.csv"] = b3
    c3.download_button("Annual Forecasts CSV", b3, "annual_forecasts.csv", "text/csv", key="dl_af")

    c4, c5, _ = st.columns(3)
    xlsx_sheets = {
        "Event Economics": econ,
        "Monthly Forecasts": event_forecasts,
        "Annual Forecasts": event_annual,
        "Validation": validation_report,
        "Source Wells": wells_classified.drop(columns=["populated_old_curves"], errors="ignore"),
        "Source Indicators": indicators_clean,
        "Source Forecasts": forecasts_clean,
    }
    xlsx_bytes = _dfs_to_xlsx(xlsx_sheets)
    csv_files["all_data.xlsx"] = xlsx_bytes
    c4.download_button(
        "Full Workbook XLSX", xlsx_bytes, "portfolio_workbook.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx",
    )
    zip_bytes = _build_zip(csv_files)
    c5.download_button("All Outputs ZIP", zip_bytes, "portfolio_outputs.zip", "application/zip", key="dl_zip")


# ════════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="VFB 2026 Inventory Overhaul",
        page_icon="📈",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # Minimal custom CSS for metric card borders
    st.html("""
    <style>
        [data-testid="stMetric"] {
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 12px 16px;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.85rem;
        }
    </style>
    """)

    st.sidebar.title("📈 VFB 2026")
    st.sidebar.caption("Inventory Overhaul Dashboard")

    if not WORKBOOK_PATH.exists():
        st.error(f"`economics.xlsx` not found beside `app.py`.\n\nExpected: `{WORKBOOK_PATH}`")
        st.stop()

    mtime_ns = WORKBOOK_PATH.stat().st_mtime_ns
    sheets = load_workbook(str(WORKBOOK_PATH), mtime_ns)
    wells_raw = sheets[SHEET_WELLS]
    indicators_raw = sheets[SHEET_ECONOMIC_INDICATORS]
    forecasts_raw = sheets[SHEET_FORECASTS]

    validation_report = validate_workbook_schema(wells_raw, indicators_raw, forecasts_raw)
    wells_classified = classify_events(wells_raw)
    indicators_clean = prepare_indicators(indicators_raw)

    ref_curves: set[str] = set()
    ref_curves.update(wells_classified[COL_NEW_CURVE].dropna().unique())
    ref_curves.update(wells_classified[COL_OLD_CURVE_1].dropna().unique())
    ref_curves.update(wells_classified[COL_OLD_CURVE_2].dropna().unique())
    ref_curves.discard("")

    forecasts_clean = prepare_forecasts(forecasts_raw, ref_curves)
    lifetime = build_type_curve_lifetime(forecasts_clean)
    econ = build_event_economics(wells_classified, indicators_clean, lifetime, forecasts_clean)
    event_forecasts = build_event_forecasts(wells_classified, forecasts_clean)
    event_annual = build_event_annual_forecasts(event_forecasts)

    # Navigation
    page = st.sidebar.radio(
        "Navigate",
        [
            "Executive Summary",
            "Consolidation",
            "Extension",
            "Creation",
            "Event Explorer",
            "Data & Downloads",
        ],
        captions=[
            "Development ladder & annual value",
            "Capital removal — 2-mile vs. 2×1-mile",
            "Capital efficiency — 1→2 mile uplift",
            "New inventory value creation",
            "Individual event drill-down",
            "Export tables & charts",
        ],
    )

    if page == "Executive Summary":
        render_portfolio_overview(econ)
    elif page == "Consolidation":
        render_consolidation(econ, event_forecasts)
    elif page == "Extension":
        render_extension(econ, event_forecasts)
    elif page == "Creation":
        render_creation(econ)
    elif page == "Event Explorer":
        render_event_explorer(
            econ, wells_classified, indicators_clean,
            forecasts_clean, event_forecasts, event_annual,
        )
    elif page == "Data & Downloads":
        render_downloads(
            econ, event_forecasts, event_annual, validation_report,
            wells_classified, indicators_clean, forecasts_clean,
        )


if __name__ == "__main__":
    main()