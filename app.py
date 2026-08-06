"""
Viewfield Bakken Exploitation  Dashboard
================================================
A production-quality Streamlit dashboard that reads economics.xlsx
and presents the results of a Viewfield Bakken exploitation .
"""
from __future__ import annotations

import warnings
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Literal
import zipfile

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ════════════════════════════════════════════════════════════════════════════════
# 1. CONSTANTS AND CONFIGURATION
# ════════════════════════════════════════════════════════════════════════════════

APP_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = APP_DIR / "economics.xlsx"

SHEET_WELLS = "wells"
SHEET_ECONOMIC_INDICATORS = "economic indicators"
SHEET_FORECASTS = "forecasts"

COL_EVENT = "event #"
COL_OLD_UWI_1 = "old well 1 UWI"
COL_OLD_UWI_2 = "old well 2 UWI"
COL_OLD_CURVE_1 = "old well 1 type curve"
COL_OLD_CURVE_2 = "old well 2 type curve"
COL_NEW_CURVE = "new well type curve"

COL_TYPE_CURVE = "type curve"
COL_NPV10 = "Npv Cash Flow BTax 10.0% (M$)"
COL_NPV_INVESTMENT_RATIO = "NPV / Disc. Invest BTax"
COL_PAYOUT = "Payout BTax (years)"
COL_RESERVES = "Boe WI Total (boe)"
COL_FIRST_YEAR_RATE = "1st Year Production Rate (boepd)"
COL_COST_OF_RESERVES = "Cost of Reserves ($/boe)"
COL_IP30 = "IP30 Cum (boe)"
COL_CAPEX = "Npv Investment BTax  0.0% (M$)"  # two spaces between BTax and 0.0%
COL_ROR = "BTax Disc. CF. ROR (%)"
COL_INITIAL_WI = "Initial WI (%)"
COL_THREE_MONTH_RATE = "3 Month Avg Production (boepd)"

COL_FORECAST_MONTH = "month #"
COL_FORECAST_YEAR = "year"
COL_FORECAST_REVENUE = "total_revenue ($M)"
COL_FORECAST_OPERATING_INCOME = "operating_income ($M)"
COL_FORECAST_CASH_FLOW = "cash_flow ($M)"

MONEY_TO_DOLLARS = 1_000.0

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
    COL_FORECAST_REVENUE, COL_FORECAST_OPERATING_INCOME, COL_FORECAST_CASH_FLOW,
)

SUM_INDICATOR_COLS = [
    COL_NPV10, COL_RESERVES, COL_FIRST_YEAR_RATE, COL_IP30,
    COL_CAPEX, COL_THREE_MONTH_RATE,
]

EVENT_TYPE_MAP = {
    "Consolidation": "Existing Plan Optimization",
    "Extension": "Inventory Enhancement Identified",
    "Creation": "New Inventory Identified",
}

FAVORABLE_HIGHER = {
    COL_NPV10, COL_NPV_INVESTMENT_RATIO, COL_RESERVES,
    COL_FIRST_YEAR_RATE, COL_IP30, COL_ROR, COL_THREE_MONTH_RATE,
}
FAVORABLE_LOWER = {COL_PAYOUT, COL_COST_OF_RESERVES}
CONTEXT_DEPENDENT = {COL_CAPEX}
NEUTRAL = {COL_INITIAL_WI}

# Color palette
COLOR_OLD = "#8e99a4"
COLOR_NEW = "#1f77b4"
COLOR_POSITIVE = "#2ca02c"
COLOR_NEGATIVE = "#d62728"
COLOR_CREATION = "#7f3fbf"
COLOR_EXTENSION = "#17becf"
COLOR_CONSOLIDATION = "#1f77b4"

# ════════════════════════════════════════════════════════════════════════════════
# 2. WORKBOOK LOADING
# ════════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Loading economics.xlsx …")
def load_workbook(path: str, modified_time_ns: int) -> dict[str, pd.DataFrame]:
    """Read the three required worksheets from the workbook."""
    sheets = pd.read_excel(
        path,
        sheet_name=[SHEET_WELLS, SHEET_ECONOMIC_INDICATORS, SHEET_FORECASTS],
        engine="openpyxl",
    )
    return sheets


# ════════════════════════════════════════════════════════════════════════════════
# 3. SCHEMA VALIDATION
# ════════════════════════════════════════════════════════════════════════════════

def _check_columns(df: pd.DataFrame, required: tuple[str, ...], sheet_name: str) -> list[str]:
    """Return list of missing columns."""
    actual = set(df.columns)
    return [c for c in required if c not in actual]


def validate_workbook_schema(
    wells: pd.DataFrame,
    indicators: pd.DataFrame,
    forecasts: pd.DataFrame,
) -> pd.DataFrame:
    """Run all validation checks. Returns a validation report DataFrame.
    Raises SystemExit via st.stop() on blocking errors.
    """
    rows: list[dict[str, Any]] = []

    def _add(severity: str, check: str, status: str, affected: int = 0, details: str = ""):
        rows.append({
            "severity": severity,
            "check": check,
            "status": status,
            "affected_count": affected,
            "details": details,
        })

    blocking = False

    # -- Columns ---
    for label, df, req in [
        (SHEET_WELLS, wells, REQUIRED_WELLS_COLS),
        (SHEET_ECONOMIC_INDICATORS, indicators, REQUIRED_INDICATOR_COLS),
        (SHEET_FORECASTS, forecasts, REQUIRED_FORECAST_COLS),
    ]:
        missing = _check_columns(df, req, label)
        if missing:
            _add("BLOCKING", f"{label} columns", "FAIL", len(missing), f"Missing: {missing}")
            blocking = True
        else:
            _add("INFO", f"{label} columns", "PASS")

    # -- event # no nulls, unique, numeric ---
    if wells[COL_EVENT].isna().any():
        _add("BLOCKING", "event # no nulls", "FAIL", int(wells[COL_EVENT].isna().sum()))
        blocking = True
    else:
        _add("INFO", "event # no nulls", "PASS")

    if wells[COL_EVENT].duplicated().any():
        _add("BLOCKING", "event # unique", "FAIL", int(wells[COL_EVENT].duplicated().sum()))
        blocking = True
    else:
        _add("INFO", "event # unique", "PASS")

    try:
        wells[COL_EVENT].astype(int)
        _add("INFO", "event # numeric", "PASS")
    except (ValueError, TypeError):
        _add("BLOCKING", "event # numeric", "FAIL", details="Cannot convert to int")
        blocking = True

    # -- new well type curve no nulls ---
    if wells[COL_NEW_CURVE].isna().any():
        n = int(wells[COL_NEW_CURVE].isna().sum())
        _add("BLOCKING", "new well type curve no nulls", "FAIL", n)
        blocking = True
    else:
        _add("INFO", "new well type curve no nulls", "PASS")

    # -- indicators unique type curve ---
    if indicators[COL_TYPE_CURVE].isna().any():
        _add("BLOCKING", "indicator type curve no nulls", "FAIL")
        blocking = True
    if indicators[COL_TYPE_CURVE].duplicated().any():
        _add("BLOCKING", "indicator type curve unique", "FAIL")
        blocking = True
    else:
        _add("INFO", "indicator type curve unique", "PASS")

    # -- forecasts no nulls in required cols ---
    for c in REQUIRED_FORECAST_COLS:
        nn = int(forecasts[c].isna().sum())
        if nn > 0:
            _add("BLOCKING", f"forecasts {c} no nulls", "FAIL", nn)
            blocking = True

    # -- Referenced type curves ---
    ref_curves = set()
    ref_curves.update(wells[COL_NEW_CURVE].dropna().unique())
    ref_curves.update(wells[COL_OLD_CURVE_1].dropna().unique())
    ref_curves.update(wells[COL_OLD_CURVE_2].dropna().unique())

    ind_curves = set(indicators[COL_TYPE_CURVE].dropna().unique())
    fc_curves = set(forecasts[COL_TYPE_CURVE].dropna().unique())

    missing_ind = ref_curves - ind_curves
    missing_fc = ref_curves - fc_curves
    if missing_ind:
        _add("BLOCKING", "ref curves in indicators", "FAIL", len(missing_ind), str(missing_ind))
        blocking = True
    else:
        _add("INFO", "ref curves in indicators", "PASS")
    if missing_fc:
        _add("BLOCKING", "ref curves in forecasts", "FAIL", len(missing_fc), str(missing_fc))
        blocking = True
    else:
        _add("INFO", "ref curves in forecasts", "PASS")

    # -- Forecast uniqueness by type curve + year + month # ---
    fc_key = forecasts[[COL_TYPE_CURVE, COL_FORECAST_YEAR, COL_FORECAST_MONTH]]
    dup_fc = fc_key.duplicated().sum()
    if dup_fc > 0:
        _add("BLOCKING", "forecast key unique", "FAIL", int(dup_fc))
        blocking = True
    else:
        _add("INFO", "forecast key unique", "PASS")

    # -- Each referenced forecast curve has identical period count ---
    ref_fc = forecasts[forecasts[COL_TYPE_CURVE].isin(ref_curves)]
    curve_counts = ref_fc.groupby(COL_TYPE_CURVE).size()
    if curve_counts.nunique() > 1:
        _add("BLOCKING", "forecast period count uniform", "FAIL",
             details=str(curve_counts.value_counts().to_dict()))
        blocking = True
    else:
        _add("INFO", "forecast period count uniform", "PASS",
             details=f"{curve_counts.iloc[0] if len(curve_counts) else 0} periods each")

    # -- Numeric conversion ---
    numeric_cols_ind = [COL_NPV10, COL_NPV_INVESTMENT_RATIO, COL_PAYOUT, COL_RESERVES,
                        COL_FIRST_YEAR_RATE, COL_COST_OF_RESERVES, COL_IP30,
                        COL_CAPEX, COL_ROR, COL_INITIAL_WI, COL_THREE_MONTH_RATE]
    for c in numeric_cols_ind:
        if c in indicators.columns:
            converted = pd.to_numeric(indicators[c], errors="coerce")
            new_nulls = converted.isna().sum() - indicators[c].isna().sum()
            if new_nulls > 0:
                _add("BLOCKING", f"indicators {c} numeric", "FAIL", int(new_nulls))
                blocking = True

    # -- Classification resolvability ---
    def _classify_check(row):
        c1 = pd.notna(row[COL_OLD_CURVE_1]) and str(row[COL_OLD_CURVE_1]).strip() != ""
        c2 = pd.notna(row[COL_OLD_CURVE_2]) and str(row[COL_OLD_CURVE_2]).strip() != ""
        if c1 and c2:
            return "Consolidation"
        elif c1 or c2:
            return "Extension"
        else:
            return "Creation"

    classifications = wells.apply(_classify_check, axis=1)
    unresolved = classifications.isna().sum()
    if unresolved > 0:
        _add("BLOCKING", "event classification resolvable", "FAIL", int(unresolved))
        blocking = True
    else:
        _add("INFO", "event classification resolvable", "PASS",
             details=classifications.value_counts().to_dict().__str__())

    # -- Nonblocking: UWI present but curve blank ---
    for uwi_col, curve_col, label in [
        (COL_OLD_UWI_1, COL_OLD_CURVE_1, "old well 1"),
        (COL_OLD_UWI_2, COL_OLD_CURVE_2, "old well 2"),
    ]:
        mask_uwi = wells[uwi_col].notna() & (wells[uwi_col].astype(str).str.strip() != "")
        mask_curve = wells[curve_col].isna() | (wells[curve_col].astype(str).str.strip() == "")
        n = int((mask_uwi & mask_curve).sum())
        if n > 0:
            _add("WARNING", f"{label} UWI present but curve blank", "WARN", n,
                 "Informational; classification uses type-curve columns only")

        mask_curve2 = wells[curve_col].notna() & (wells[curve_col].astype(str).str.strip() != "")
        mask_uwi2 = wells[uwi_col].isna() | (wells[uwi_col].astype(str).str.strip() == "")
        n2 = int((mask_curve2 & mask_uwi2).sum())
        if n2 > 0:
            _add("WARNING", f"{label} curve present but UWI blank", "WARN", n2)

    # -- Nonblocking: extra type curves ---
    extra_ind = ind_curves - ref_curves
    if extra_ind:
        _add("WARNING", "extra type curves in indicators", "INFO", len(extra_ind),
             str(extra_ind))
    extra_fc = fc_curves - ref_curves
    if extra_fc:
        _add("WARNING", "extra type curves in forecasts", "INFO", len(extra_fc),
             str(extra_fc))

    report = pd.DataFrame(rows)

    if blocking:
        st.error("❌ Blocking validation errors found. See the validation details below.")
        for r in rows:
            if r["severity"] == "BLOCKING" and r["status"] == "FAIL":
                st.error(f"**{r['check']}**: {r['details']}")
        st.dataframe(report)
        st.stop()

    return report


# ════════════════════════════════════════════════════════════════════════════════
# 4–5. DATA NORMALIZATION & EVENT CLASSIFICATION
# ════════════════════════════════════════════════════════════════════════════════

def classify_events(wells: pd.DataFrame) -> pd.DataFrame:
    """Add event_type, event_story, old_curve_count, populated_old_curves columns."""
    df = wells.copy()
    df[COL_EVENT] = df[COL_EVENT].astype(int)

    def _is_populated(val) -> bool:
        return pd.notna(val) and str(val).strip() != ""

    types, stories, old_counts, pop_curves = [], [], [], []
    for _, row in df.iterrows():
        c1 = _is_populated(row[COL_OLD_CURVE_1])
        c2 = _is_populated(row[COL_OLD_CURVE_2])
        if c1 and c2:
            t = "Consolidation"
            cnt = 2
            curves = [row[COL_OLD_CURVE_1], row[COL_OLD_CURVE_2]]
        elif c1 or c2:
            t = "Extension"
            cnt = 1
            curves = [row[COL_OLD_CURVE_1]] if c1 else [row[COL_OLD_CURVE_2]]
        else:
            t = "Creation"
            cnt = 0
            curves = []
        types.append(t)
        stories.append(EVENT_TYPE_MAP[t])
        old_counts.append(cnt)
        pop_curves.append(curves)

    df["event_type"] = types
    df["event_story"] = stories
    df["old_curve_count"] = old_counts
    df["populated_old_curves"] = pop_curves
    return df


# ════════════════════════════════════════════════════════════════════════════════
# 6. FORECAST PREPARATION
# ════════════════════════════════════════════════════════════════════════════════

def prepare_forecasts(forecasts: pd.DataFrame, ref_curves: set[str]) -> pd.DataFrame:
    """Filter to referenced curves, add producing_month & dollar columns."""
    df = forecasts[forecasts[COL_TYPE_CURVE].isin(ref_curves)].copy()
    df = df.sort_values([COL_TYPE_CURVE, COL_FORECAST_YEAR, COL_FORECAST_MONTH]).reset_index(drop=True)
    df["producing_month"] = df.groupby(COL_TYPE_CURVE).cumcount() + 1

    for src, dst in [
        (COL_FORECAST_REVENUE, "total_revenue_dollars"),
        (COL_FORECAST_OPERATING_INCOME, "operating_income_dollars"),
        (COL_FORECAST_CASH_FLOW, "cash_flow_dollars"),
    ]:
        df[dst] = pd.to_numeric(df[src], errors="coerce") * MONEY_TO_DOLLARS

    for metric in ["total_revenue_dollars", "operating_income_dollars", "cash_flow_dollars"]:
        cum_name = f"cumulative_{metric.replace('_dollars', '')}_dollars"
        df[cum_name] = df.groupby(COL_TYPE_CURVE)[metric].cumsum()

    return df


def build_type_curve_lifetime(forecasts_clean: pd.DataFrame) -> pd.DataFrame:
    """Lifetime sums by type curve."""
    agg = forecasts_clean.groupby(COL_TYPE_CURVE).agg(
        lifetime_revenue_dollars=("total_revenue_dollars", "sum"),
        lifetime_operating_income_dollars=("operating_income_dollars", "sum"),
        lifetime_cash_flow_dollars=("cash_flow_dollars", "sum"),
        forecast_periods=("producing_month", "count"),
    ).reset_index()
    return agg


# ════════════════════════════════════════════════════════════════════════════════
# 7. INDICATOR PREPARATION
# ════════════════════════════════════════════════════════════════════════════════

def prepare_indicators(indicators: pd.DataFrame) -> pd.DataFrame:
    """Add actual-dollar columns."""
    df = indicators.copy()
    df["npv10_dollars"] = pd.to_numeric(df[COL_NPV10], errors="coerce") * MONEY_TO_DOLLARS
    df["investment_dollars"] = pd.to_numeric(df[COL_CAPEX], errors="coerce") * MONEY_TO_DOLLARS
    return df


# ════════════════════════════════════════════════════════════════════════════════
# 8. PAYOUT HELPER
# ════════════════════════════════════════════════════════════════════════════════

def calculate_payout_years(monthly_cash_flow: pd.Series) -> float:
    """Derive payout in years from a monthly cash-flow series (chronological)."""
    cum = monthly_cash_flow.values.cumsum()
    if len(cum) == 0:
        return np.nan
    if cum[0] >= 0:
        return 0.0
    for i in range(1, len(cum)):
        if cum[i] >= 0:
            prev = cum[i - 1]
            month_cf = monthly_cash_flow.values[i]
            if month_cf <= 0:
                return np.nan
            fraction = abs(prev) / month_cf
            payout_months = i + fraction
            return payout_months / 12.0
    return np.nan


# ════════════════════════════════════════════════════════════════════════════════
# 9. EVENT-LEVEL ECONOMIC CALCULATIONS
# ════════════════════════════════════════════════════════════════════════════════

def _safe_div(num: float, den: float) -> float:
    if den == 0 or np.isnan(den) or np.isnan(num):
        return np.nan
    return num / den


@st.cache_data(show_spinner="Calculating event economics …")
def build_event_economics(
    wells_classified: pd.DataFrame,
    indicators_clean: pd.DataFrame,
    lifetime: pd.DataFrame,
    forecasts_clean: pd.DataFrame,
) -> pd.DataFrame:
    """One row per event with old/new/delta financials."""
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

        # --- Old plan metrics ---
        old_capex = sum(ind.loc[c, "investment_dollars"] for c in old_curves) if old_curves else 0.0
        old_npv10 = sum(ind.loc[c, "npv10_dollars"] for c in old_curves) if old_curves else 0.0
        old_reserves = sum(ind.loc[c, COL_RESERVES] for c in old_curves) if old_curves else 0.0
        old_lt_rev = sum(lt.loc[c, "lifetime_revenue_dollars"] for c in old_curves) if old_curves else 0.0
        old_lt_oi = sum(lt.loc[c, "lifetime_operating_income_dollars"] for c in old_curves) if old_curves else 0.0
        old_lt_cf = sum(lt.loc[c, "lifetime_cash_flow_dollars"] for c in old_curves) if old_curves else 0.0
        old_first_yr = sum(ind.loc[c, COL_FIRST_YEAR_RATE] for c in old_curves) if old_curves else 0.0
        old_ip30 = sum(ind.loc[c, COL_IP30] for c in old_curves) if old_curves else 0.0
        old_3mo = sum(ind.loc[c, COL_THREE_MONTH_RATE] for c in old_curves) if old_curves else 0.0

        # --- New plan metrics ---
        new_capex = ind.loc[new_curve, "investment_dollars"]
        new_npv10 = ind.loc[new_curve, "npv10_dollars"]
        new_reserves = ind.loc[new_curve, COL_RESERVES]
        new_lt_rev = lt.loc[new_curve, "lifetime_revenue_dollars"]
        new_lt_oi = lt.loc[new_curve, "lifetime_operating_income_dollars"]
        new_lt_cf = lt.loc[new_curve, "lifetime_cash_flow_dollars"]
        new_first_yr = ind.loc[new_curve, COL_FIRST_YEAR_RATE]
        new_ip30 = ind.loc[new_curve, COL_IP30]
        new_3mo = ind.loc[new_curve, COL_THREE_MONTH_RATE]
        new_payout_source = ind.loc[new_curve, COL_PAYOUT]
        new_ror = ind.loc[new_curve, COL_ROR]
        new_npv_inv_ratio = ind.loc[new_curve, COL_NPV_INVESTMENT_RATIO]
        new_cor = ind.loc[new_curve, COL_COST_OF_RESERVES]
        new_wi = ind.loc[new_curve, COL_INITIAL_WI]

        # Old-plan payout from combined cash flow
        if old_curves:
            old_fc_list = []
            for c in old_curves:
                fc = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == c].sort_values("producing_month")
                old_fc_list.append(fc[["producing_month", "cash_flow_dollars"]].set_index("producing_month"))
            combined_old_cf = pd.concat(old_fc_list, axis=1).sum(axis=1).sort_index()
            old_payout = calculate_payout_years(combined_old_cf)
        else:
            old_payout = np.nan

        # New-plan payout (derived)
        new_fc = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == new_curve].sort_values("producing_month")
        new_payout_derived = calculate_payout_years(new_fc["cash_flow_dollars"])

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
        ev["revenue_delta_dollars"] = new_lt_rev - old_lt_rev

        ev["old_lifetime_operating_income_dollars"] = old_lt_oi
        ev["new_lifetime_operating_income_dollars"] = new_lt_oi
        ev["operating_income_delta_dollars"] = new_lt_oi - old_lt_oi

        ev["old_lifetime_cash_flow_dollars"] = old_lt_cf
        ev["new_lifetime_cash_flow_dollars"] = new_lt_cf
        ev["cash_flow_delta_dollars"] = new_lt_cf - old_lt_cf

        ev["old_first_year_rate_boepd"] = old_first_yr
        ev["new_first_year_rate_boepd"] = new_first_yr
        ev["old_ip30_boe"] = old_ip30
        ev["new_ip30_boe"] = new_ip30
        ev["old_3mo_rate_boepd"] = old_3mo
        ev["new_3mo_rate_boepd"] = new_3mo

        ev["old_payout_years"] = old_payout
        ev["new_payout_source"] = new_payout_source
        ev["new_payout_derived"] = new_payout_derived

        ev["new_ror_pct"] = new_ror
        ev["new_npv_inv_ratio"] = new_npv_inv_ratio
        ev["new_cost_of_reserves"] = new_cor
        ev["new_initial_wi_pct"] = new_wi

        # Derived efficiency metrics
        ev["old_cost_of_reserves"] = _safe_div(old_capex, old_reserves)
        ev["new_cost_of_reserves_derived"] = _safe_div(new_capex, new_reserves)
        ev["cost_of_reserves_delta"] = (
            _safe_div(new_capex, new_reserves) - _safe_div(old_capex, old_reserves)
            if old_reserves > 0 else np.nan
        )

        # Capital-efficiency fields. These preserve the existing old/new calculation approach
        # while expressing favorable outcomes as positive values for dashboard reporting.
        is_consolidation = ev["event_type"] == "Consolidation"
        is_inventory_add = ev["event_type"] in {"Extension", "Creation"}
        ev["capital_saved_dollars"] = max(old_capex - new_capex, 0.0) if is_consolidation else 0.0
        ev["boe_lost"] = max(old_reserves - new_reserves, 0.0) if is_consolidation else 0.0
        ev["capital_saved_per_boe_lost"] = (
            _safe_div(ev["capital_saved_dollars"], ev["boe_lost"])
            if is_consolidation and ev["boe_lost"] > 0 else np.nan
        )
        ev["cost_of_reserves_improvement"] = (
            _safe_div(old_capex, old_reserves) - _safe_div(new_capex, new_reserves)
            if is_consolidation and old_reserves > 0 and new_reserves > 0 else np.nan
        )
        ev["locations_eliminated"] = max(int(row["old_curve_count"]) - 1, 0) if is_consolidation else 0
        ev["npv_sacrificed_dollars"] = max(old_npv10 - new_npv10, 0.0) if is_consolidation else 0.0
        ev["incremental_npv10_added_dollars"] = (
            (new_npv10 - old_npv10) if ev["event_type"] == "Extension"
            else new_npv10 if ev["event_type"] == "Creation" else 0.0
        )
        ev["old_npv_inv_ratio_derived"] = _safe_div(old_npv10, old_capex)
        ev["new_npv_inv_ratio_derived"] = _safe_div(new_npv10, new_capex)
        ev["old_lifetime_oi_per_boe"] = _safe_div(old_lt_oi, old_reserves)
        ev["new_lifetime_oi_per_boe"] = _safe_div(new_lt_oi, new_reserves)
        ev["old_npv10_per_boe"] = _safe_div(old_npv10, old_reserves)
        ev["new_npv10_per_boe"] = _safe_div(new_npv10, new_reserves)

        # Extension-specific marginal metrics
        inc_cap = new_capex - old_capex
        inc_npv = new_npv10 - old_npv10
        inc_res = new_reserves - old_reserves
        ev["marginal_npv_to_inc_capital"] = _safe_div(inc_npv, inc_cap) if inc_cap > 0 else np.nan
        ev["inc_capital_per_inc_boe"] = _safe_div(inc_cap, inc_res) if inc_res > 0 else np.nan
        ev["inc_npv_per_inc_boe"] = _safe_div(inc_npv, inc_res) if inc_res > 0 else np.nan

        ev["old_type_curves_used"] = " | ".join(old_curves) if old_curves else ""
        ev["new_type_curve_used"] = new_curve

        records.append(ev)

    return pd.DataFrame(records)


# ════════════════════════════════════════════════════════════════════════════════
# 10. EVENT FORECAST SERIES
# ════════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Building event forecasts …")
def build_event_forecasts(
    wells_classified: pd.DataFrame,
    forecasts_clean: pd.DataFrame,
) -> pd.DataFrame:
    """Long-form event forecast table."""
    all_series: list[pd.DataFrame] = []

    for _, row in wells_classified.iterrows():
        ev_num = int(row[COL_EVENT])
        ev_type = row["event_type"]
        ev_story = row["event_story"]
        old_curves: list[str] = row["populated_old_curves"]
        new_curve: str = row[COL_NEW_CURVE]

        base_cols = ["producing_month", COL_FORECAST_YEAR, COL_FORECAST_MONTH,
                     "total_revenue_dollars", "operating_income_dollars", "cash_flow_dollars"]

        if ev_type == "Consolidation":
            dfs_old = []
            for idx, c in enumerate(old_curves, 1):
                fc = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == c].sort_values("producing_month")[base_cols].copy()
                fc["series_key"] = f"old_{idx}"
                fc["series_label"] = f"Old Well {idx}: {c}"
                fc["plan_side"] = "old"
                dfs_old.append(fc)
            # Combined old
            if len(dfs_old) == 2:
                comb = dfs_old[0][base_cols].copy().set_index("producing_month")
                comb2 = dfs_old[1][base_cols].copy().set_index("producing_month")
                combined = comb.copy()
                for mc in ["total_revenue_dollars", "operating_income_dollars", "cash_flow_dollars"]:
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
            c = old_curves[0]
            fc = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == c].sort_values("producing_month")[base_cols].copy()
            fc["series_key"] = "old_ref"
            fc["series_label"] = f"One-Mile Reference: {c}"
            fc["plan_side"] = "old"
            fc["event"] = ev_num
            fc["event_type"] = ev_type
            fc["event_story"] = ev_story
            all_series.append(fc)

        # New curve (all types)
        fc_new = forecasts_clean[forecasts_clean[COL_TYPE_CURVE] == new_curve].sort_values("producing_month")[base_cols].copy()
        if ev_type == "Consolidation":
            fc_new["series_key"] = "new_plan"
            fc_new["series_label"] = f"New Two-Mile Plan: {new_curve}"
        elif ev_type == "Extension":
            fc_new["series_key"] = "new_opp"
            fc_new["series_label"] = f"Two-Mile Opportunity: {new_curve}"
        else:
            fc_new["series_key"] = "new_inventory"
            fc_new["series_label"] = f"New Inventory Opportunity: {new_curve}"
        fc_new["plan_side"] = "new"
        fc_new["event"] = ev_num
        fc_new["event_type"] = ev_type
        fc_new["event_story"] = ev_story
        all_series.append(fc_new)

    if not all_series:
        return pd.DataFrame()

    result = pd.concat(all_series, ignore_index=True)

    # Add cumulative columns per event+series
    for metric in ["total_revenue_dollars", "operating_income_dollars", "cash_flow_dollars"]:
        cum_name = f"cumulative_{metric.replace('_dollars', '')}_dollars"
        result[cum_name] = result.groupby(["event", "series_key"])[metric].cumsum()

    return result


def build_event_annual_forecasts(event_forecasts: pd.DataFrame) -> pd.DataFrame:
    """Annual aggregation of event forecasts."""
    if event_forecasts.empty:
        return pd.DataFrame()
    group_cols = ["event", "event_type", "event_story", "series_key", "series_label", "plan_side", COL_FORECAST_YEAR]
    agg = event_forecasts.groupby(group_cols, as_index=False).agg(
        total_revenue_dollars=("total_revenue_dollars", "sum"),
        operating_income_dollars=("operating_income_dollars", "sum"),
        cash_flow_dollars=("cash_flow_dollars", "sum"),
    )
    return agg


# ════════════════════════════════════════════════════════════════════════════════
# 11. EXECUTIVE SUMMARY
# ════════════════════════════════════════════════════════════════════════════════

def build_executive_summary(econ: pd.DataFrame) -> pd.DataFrame:
    """One-row-per-category + combined total."""
    rows = []
    for et in ["Consolidation", "Extension", "Creation"]:
        sub = econ[econ["event_type"] == et]
        r: dict[str, Any] = {"category": et, "event_story": EVENT_TYPE_MAP[et], "events": len(sub)}
        if et == "Consolidation":
            r["primary_value_dollars"] = sub["npv10_delta_dollars"].sum()
            r["capital_dollars"] = sub["capex_delta_dollars"].sum()
            r["reserves_boe"] = sub["reserves_delta_boe"].sum()
            r["capital_saved_dollars"] = sub["capital_saved_dollars"].sum()
            r["boe_lost"] = sub["boe_lost"].sum()
            r["capital_saved_per_boe_lost"] = _safe_div(r["capital_saved_dollars"], r["boe_lost"])
            r["cost_of_reserves_improvement"] = (
                _safe_div(sub["old_capex_dollars"].sum(), sub["old_reserves_boe"].sum())
                - _safe_div(sub["new_capex_dollars"].sum(), sub["new_reserves_boe"].sum())
            )
            r["locations_eliminated"] = int(sub["locations_eliminated"].sum())
            r["incremental_npv10_added_dollars"] = 0.0
        elif et == "Extension":
            r["primary_value_dollars"] = sub["npv10_delta_dollars"].sum()
            r["capital_dollars"] = sub["capex_delta_dollars"].sum()
            r["reserves_boe"] = sub["reserves_delta_boe"].sum()
            r["capital_saved_dollars"] = 0.0
            r["boe_lost"] = 0.0
            r["capital_saved_per_boe_lost"] = np.nan
            r["cost_of_reserves_improvement"] = np.nan
            r["locations_eliminated"] = 0
            r["incremental_npv10_added_dollars"] = sub["incremental_npv10_added_dollars"].sum()
        else:
            r["primary_value_dollars"] = sub["new_npv10_dollars"].sum()
            r["capital_dollars"] = sub["new_capex_dollars"].sum()
            r["reserves_boe"] = sub["new_reserves_boe"].sum()
            r["capital_saved_dollars"] = 0.0
            r["boe_lost"] = 0.0
            r["capital_saved_per_boe_lost"] = np.nan
            r["cost_of_reserves_improvement"] = np.nan
            r["locations_eliminated"] = 0
            r["incremental_npv10_added_dollars"] = sub["incremental_npv10_added_dollars"].sum()
        rows.append(r)

    total = {
        "category": "Combined",
        "event_story": "Total  Impact",
        "events": sum(r["events"] for r in rows),
        "primary_value_dollars": sum(r["primary_value_dollars"] for r in rows),
        "capital_dollars": sum(r["capital_dollars"] for r in rows),
        "reserves_boe": sum(r["reserves_boe"] for r in rows),
        "capital_saved_dollars": sum(r["capital_saved_dollars"] for r in rows),
        "boe_lost": sum(r["boe_lost"] for r in rows),
        "capital_saved_per_boe_lost": _safe_div(
            sum(r["capital_saved_dollars"] for r in rows), sum(r["boe_lost"] for r in rows)
        ),
        "cost_of_reserves_improvement": rows[0]["cost_of_reserves_improvement"] if rows else np.nan,
        "locations_eliminated": sum(r["locations_eliminated"] for r in rows),
        "incremental_npv10_added_dollars": sum(r["incremental_npv10_added_dollars"] for r in rows),
    }
    rows.append(total)
    return pd.DataFrame(rows)


# ════════════════════════════════════════════════════════════════════════════════
# 12. FORMATTING HELPERS
# ════════════════════════════════════════════════════════════════════════════════

def fmt_mm(val: float | None, prefix: str = "$", suffix: str = " MM") -> str:
    """Format actual dollars as $X.X MM."""
    if val is None or np.isnan(val):
        return "N/A"
    mm = val / 1_000_000
    if mm < 0:
        return f"({prefix}{abs(mm):,.1f}{suffix})"
    return f"{prefix}{mm:,.1f}{suffix}"


def fmt_signed_mm(val: float | None, suffix: str = " MM") -> str:
    """Format a change in actual dollars with an explicit + / - sign."""
    if val is None or np.isnan(val):
        return "N/A"
    return f"{val / 1_000_000:+,.1f}{suffix}".replace("+", "+$", 1).replace("-", "-$", 1)


def fmt_mboe(val: float | None) -> str:
    if val is None or np.isnan(val):
        return "N/A"
    return f"{val / 1_000:,.1f} Mboe"


def fmt_signed_mboe(val: float | None) -> str:
    """Format a reserves change with an explicit + / - sign."""
    if val is None or np.isnan(val):
        return "N/A"
    return f"{val / 1_000:+,.1f} Mboe"


def fmt_signed_dollarboe(val: float | None) -> str:
    """Format a $/boe change with an explicit + / - sign."""
    if val is None or np.isnan(val):
        return "N/A"
    return f"{val:+,.2f} $/boe"


def fmt_pct(val: float | None) -> str:
    if val is None or np.isnan(val):
        return "N/A"
    return f"{val:,.1f}%"


def fmt_ratio(val: float | None) -> str:
    if val is None or np.isnan(val):
        return "N/M"
    return f"{val:,.2f}x"


def fmt_dollarboe(val: float | None) -> str:
    if val is None or np.isnan(val):
        return "N/A"
    return f"${val:,.2f}/boe"


def fmt_years(val: float | None) -> str:
    if val is None or np.isnan(val):
        return "No payout"
    return f"{val:,.2f} years"


def fmt_rate(val: float | None) -> str:
    if val is None or np.isnan(val):
        return "N/A"
    return f"{val:,.0f} boe/d"


def fmt_boe(val: float | None) -> str:
    if val is None or np.isnan(val):
        return "N/A"
    return f"{val:,.0f} boe"


# ════════════════════════════════════════════════════════════════════════════════
# 13. DOWNLOAD HELPERS
# ════════════════════════════════════════════════════════════════════════════════

def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def dataframes_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            sheet_name = name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
    buf.seek(0)
    return buf.read()


def plotly_to_html_bytes(fig: go.Figure) -> bytes:
    return fig.to_html(include_plotlyjs="cdn").encode("utf-8")


def plotly_to_image_bytes(fig: go.Figure, fmt: str = "png") -> bytes | None:
    try:
        return fig.to_image(format=fmt, width=1200, height=600, scale=2)
    except Exception:
        return None


def build_zip_bytes(files: dict[str, bytes]) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    buf.seek(0)
    return buf.read()


# ════════════════════════════════════════════════════════════════════════════════
# 14. CHART HELPERS
# ════════════════════════════════════════════════════════════════════════════════

def _base_layout(fig: go.Figure, title: str, xaxis: str = "", yaxis: str = ""):
    fig.update_layout(
        title=title,
        xaxis_title=xaxis,
        yaxis_title=yaxis,
        template="plotly_white",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.3),
        margin=dict(l=60, r=30, t=50, b=40),
    )
    return fig


# ════════════════════════════════════════════════════════════════════════════════
# PAGE RENDERERS
# ════════════════════════════════════════════════════════════════════════════════

def render_executive_summary(econ: pd.DataFrame, exec_summary: pd.DataFrame, event_forecasts: pd.DataFrame):
    st.header("Capital Efficiency")
    st.caption("How many dollars of capital did we free up while preserving development value?")

    consol = econ[econ["event_type"] == "Consolidation"]
    ext = econ[econ["event_type"] == "Extension"]
    cre = econ[econ["event_type"] == "Creation"]

    total_capital_saved = consol["capital_saved_dollars"].sum()
    total_boe_lost = consol["boe_lost"].sum()
    portfolio_saved_per_boe_lost = _safe_div(total_capital_saved, total_boe_lost)
    portfolio_old_cor = _safe_div(consol["old_capex_dollars"].sum(), consol["old_reserves_boe"].sum())
    portfolio_new_cor = _safe_div(consol["new_capex_dollars"].sum(), consol["new_reserves_boe"].sum())
    portfolio_cor_improvement = (portfolio_old_cor - portfolio_new_cor
                                 if pd.notna(portfolio_old_cor) and pd.notna(portfolio_new_cor) else np.nan)
    total_locations_eliminated = int(consol["locations_eliminated"].sum())
    total_incremental_npv_added = (
        ext["incremental_npv10_added_dollars"].sum()
        + cre["incremental_npv10_added_dollars"].sum()
    )

    st.subheader("Capital Efficiency")
    h1, h2, h3, h4, h5 = st.columns(5)
    h1.metric("Capital Saved", fmt_mm(total_capital_saved))
    h3.metric("Cost of Reserves Improvement", fmt_dollarboe(portfolio_cor_improvement))
    h4.metric("Locations Eliminated", f"{total_locations_eliminated:,}")
    h5.metric("Incremental NPV10 Added", fmt_mm(total_incremental_npv_added))

    st.caption(
        "Capital Saved, BOE Lost, cost-of-reserves improvement, and locations eliminated are based on consolidations. "
        "Incremental NPV10 Added combines extension NPV10 deltas and creation NPV10."
    )

    total_val = (
        consol["npv10_delta_dollars"].sum()
        + ext["npv10_delta_dollars"].sum()
        + cre["new_npv10_dollars"].sum()
    )
    total_cap = (
        consol["capex_delta_dollars"].sum()
        + ext["capex_delta_dollars"].sum()
        + cre["new_capex_dollars"].sum()
    )

    # --- Section A: Overall  ---
    st.subheader("Portfolio Changes")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Events", f"{len(econ):,}")
    c2.metric("Existing Plan Optimizations", f"{len(consol):,}")
    c3.metric("Inventory Enhancements", f"{len(ext):,}")
    c4.metric("New Inventory Locations", f"{len(cre):,}")
    c5.metric(" Value Change ", fmt_signed_mm(total_val))
    c6.metric(" Capital Change (+ added / - saved)", fmt_signed_mm(total_cap))
    st.caption(
        " Value = consolidation NPV10 change + extension incremental NPV10 + creation NPV10. "
        " Capital = consolidation capital change + extension incremental capital + creation capital. "
        "Positive values are added; negative values are removed or saved."
    )

    # --- Section B:  Value Mix ---
    st.subheader(" Value Breakdown")
    value_mix = pd.DataFrame({
        "Category": [
            "Existing Plan Optimization",
            "Inventory Enhancement",
            "New Inventory",
        ],
        "Value": [
            consol["npv10_delta_dollars"].sum(),
            ext["npv10_delta_dollars"].sum(),
            cre["new_npv10_dollars"].sum(),
        ],
    })
    pie_values = value_mix["Value"].clip(lower=0)
    if pie_values.sum() > 0:
        fig_mix = go.Figure(go.Pie(
            labels=value_mix["Category"],
            values=pie_values,
            customdata=value_mix["Value"] / 1e6,
            texttemplate="%{label}<br>%{percent}<br>$%{customdata:.1f} MM",
            hovertemplate="%{label}<br> value: $%{customdata:.2f} MM<extra></extra>",
            hole=0.35,
        ))
        fig_mix.update_layout(
            template="plotly_white",
            margin=dict(l=30, r=30, t=60, b=30),
            legend=dict(orientation="h", yanchor="bottom", y=-0.15),
        )
        st.plotly_chart(fig_mix, use_container_width=True)
        if (value_mix["Value"] < 0).any():
            st.caption("Negative category values are excluded from pie-slice sizing and remain visible in the signed summary metrics above.")
    else:
        st.info("No positive  value is available for the pie chart.")

    # --- Section C: Existing Plan Optimization ---
    st.subheader("Existing Plan Optimization (Consolidations)")
    cb1, cb2, cb3, cb4 = st.columns(4)
    cb1.metric("Consolidation Events", f"{len(consol):,}")
    cb2.metric("Capital Saved", fmt_mm(consol["capital_saved_dollars"].sum()))
    cb4.metric("Locations Eliminated", f"{int(consol['locations_eliminated'].sum()):,}")

    cd1, cd2, cd3, cd4 = st.columns(4)
    cd1.metric("NPV10 Change", fmt_signed_mm(consol["npv10_delta_dollars"].sum()))
    cd2.metric("NPV10 Sacrificed", fmt_mm(consol["npv_sacrificed_dollars"].sum()))
    cd3.metric("Cost of Reserves Improvement", fmt_dollarboe(portfolio_cor_improvement))
    cd4.metric("BOE Lost", fmt_boe(consol["boe_lost"].sum()))

    # --- Section D: Inventory Enhancements ---
    st.subheader("Inventory Enhancements Identified (Extensions)")
    if len(ext) > 0:
        ce1, ce2, ce3, ce4, ce5 = st.columns(5)
        ce1.metric("Extension Opportunities", f"{len(ext):,}")
        ce2.metric("Capital Change", fmt_signed_mm(ext["capex_delta_dollars"].sum()))
        ce3.metric("Incremental NPV10 Added", fmt_signed_mm(ext["incremental_npv10_added_dollars"].sum()))
        ce4.metric("Reserves Change", fmt_signed_mboe(ext["reserves_delta_boe"].sum()))
        valid_marginal = ext[ext["marginal_npv_to_inc_capital"].notna()]
        if len(valid_marginal) > 0:
            weighted = ext["npv10_delta_dollars"].sum() / ext["capex_delta_dollars"].sum() if ext["capex_delta_dollars"].sum() != 0 else np.nan
            ce5.metric("Portfolio Marginal NPV/Cap", fmt_ratio(weighted))
        else:
            ce5.metric("Portfolio Marginal NPV/Cap", "N/M")

        fig_ext = go.Figure(go.Bar(
            x=ext.sort_values("npv10_delta_dollars", ascending=False)["event"].astype(str),
            y=ext.sort_values("npv10_delta_dollars", ascending=False)["npv10_delta_dollars"] / 1e6,
            marker_color=COLOR_EXTENSION,
        ))
        _base_layout(fig_ext, "Incremental NPV10 by Extension Event", "Event #", "Incremental NPV10 ($MM)")
        fig_ext.add_hline(y=0, line_dash="dash", line_color="gray")
        st.plotly_chart(fig_ext, use_container_width=True)
    else:
        st.info("No extension events.")

    # --- Section E: New Inventory ---
    st.subheader("New Inventory Identified (Creations)")
    if len(cre) > 0:
        cc1, cc2, cc3, cc4 = st.columns(4)
        cc1.metric("New Locations Identified", f"{len(cre):,}")
        cc2.metric("New Inventory Capital Added", fmt_signed_mm(cre["new_capex_dollars"].sum()))
        cc3.metric("Incremental NPV10 Added", fmt_signed_mm(cre["incremental_npv10_added_dollars"].sum()))
        cc4.metric("New Inventory Reserves Added", fmt_signed_mboe(cre["new_reserves_boe"].sum()))

        fig_cre = go.Figure(go.Bar(
            x=cre.sort_values("new_npv10_dollars", ascending=False)["event"].astype(str),
            y=cre.sort_values("new_npv10_dollars", ascending=False)["new_npv10_dollars"] / 1e6,
            marker_color=COLOR_CREATION,
        ))
        _base_layout(fig_cre, "New Inventory NPV10 by Event", "Event #", "NPV10 ($MM)")
        st.plotly_chart(fig_cre, use_container_width=True)
    else:
        st.info("No creation events.")

    



def render_existing_plan_optimization(econ: pd.DataFrame, event_forecasts: pd.DataFrame):
    st.header("Existing Plan Optimization: One-Mile to Two-Mile Consolidations")

    consol = econ[econ["event_type"] == "Consolidation"].copy()

    if len(consol) == 0:
        st.info("No consolidation events found.")
        return

    # Filters
    with st.expander("Filters", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        new_curves_avail = sorted(consol[COL_NEW_CURVE].unique())
        sel_new_curve = fc1.multiselect("New Type Curve", new_curves_avail, default=new_curves_avail, key="opt_nc")
        npv_dir = fc2.radio("NPV Uplift Direction", ["All", "Positive", "Negative"], key="opt_npv_dir")
        cap_dir = fc3.radio("Capital Direction", ["All", "Reduction", "Increase"], key="opt_cap_dir")

    mask = consol[COL_NEW_CURVE].isin(sel_new_curve)
    if npv_dir == "Positive":
        mask &= consol["npv10_delta_dollars"] >= 0
    elif npv_dir == "Negative":
        mask &= consol["npv10_delta_dollars"] < 0
    if cap_dir == "Reduction":
        mask &= consol["capex_delta_dollars"] < 0
    elif cap_dir == "Increase":
        mask &= consol["capex_delta_dollars"] >= 0
    filtered = consol[mask]

    # Summary metrics
    mc1, mc2, mc3, mc4, mc5 = st.columns(5)
    mc1.metric("Events", f"{len(filtered):,}")
    mc2.metric("Capital Saved", fmt_mm(filtered["capital_saved_dollars"].sum()))
    filtered_old_cor = _safe_div(filtered["old_capex_dollars"].sum(), filtered["old_reserves_boe"].sum())
    filtered_new_cor = _safe_div(filtered["new_capex_dollars"].sum(), filtered["new_reserves_boe"].sum())
    mc4.metric("Cost of Reserves Improvement", fmt_dollarboe(
        filtered_old_cor - filtered_new_cor if pd.notna(filtered_old_cor) and pd.notna(filtered_new_cor) else np.nan))
    mc5.metric("Locations Eliminated", f"{int(filtered['locations_eliminated'].sum()):,}")

    mc6, mc7, mc8, mc9, mc10 = st.columns(5)
    mc6.metric("Old Plan NPV10", fmt_mm(filtered["old_npv10_dollars"].sum()))
    mc7.metric("New Plan NPV10", fmt_mm(filtered["new_npv10_dollars"].sum()))
    mc8.metric("Old Reserves", fmt_mboe(filtered["old_reserves_boe"].sum()))
    mc9.metric("New Reserves", fmt_mboe(filtered["new_reserves_boe"].sum()))
    mc10.metric("Reserves Change", fmt_signed_mboe(filtered["reserves_delta_boe"].sum()))

    # Charts
    col_l, col_r = st.columns(2)
    with col_l:
        df_r = filtered.sort_values("npv10_delta_dollars", ascending=False)
        colors = [COLOR_POSITIVE if v >= 0 else COLOR_NEGATIVE for v in df_r["npv10_delta_dollars"]]
        fig = go.Figure(go.Bar(
            x=df_r["event"].astype(str), y=df_r["npv10_delta_dollars"] / 1e6,
            marker_color=colors,
            hovertemplate="Event %{x}<br>NPV Uplift: $%{y:.2f} MM<extra></extra>",
        ))
        _base_layout(fig, "NPV10 Uplift by Event", "Event #", "NPV10 Uplift ($MM)")
        fig.add_hline(y=0, line_dash="dash", line_color="gray")
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        df_r2 = filtered.sort_values("capex_delta_dollars", ascending=True)
        colors2 = [COLOR_POSITIVE if v < 0 else COLOR_NEGATIVE for v in df_r2["capex_delta_dollars"]]
        fig2 = go.Figure(go.Bar(
            x=df_r2["event"].astype(str), y=df_r2["capex_delta_dollars"] / 1e6,
            marker_color=colors2,
        ))
        _base_layout(fig2, "Capital Delta by Event", "Event #", "Capital Delta ($MM)")
        fig2.add_hline(y=0, line_dash="dash", line_color="gray")
        st.plotly_chart(fig2, use_container_width=True)


    # Event table
    st.subheader("Event Detail Table")
    display_cols = [
        "event", COL_OLD_UWI_1, COL_OLD_UWI_2, COL_OLD_CURVE_1, COL_OLD_CURVE_2, COL_NEW_CURVE,
        "old_capex_dollars", "new_capex_dollars", "capex_delta_dollars",
        "old_npv10_dollars", "new_npv10_dollars", "npv10_delta_dollars",
        "old_reserves_boe", "new_reserves_boe", "reserves_delta_boe",
        "old_cost_of_reserves", "new_cost_of_reserves_derived",
        "capital_saved_dollars", "boe_lost", "capital_saved_per_boe_lost",
        "cost_of_reserves_improvement", "locations_eliminated", "npv_sacrificed_dollars",
    ]
    avail = [c for c in display_cols if c in filtered.columns]
    st.dataframe(filtered[avail].sort_values("event"), use_container_width=True, hide_index=True)


def render_inventory_opportunities(econ: pd.DataFrame):
    st.header("Inventory Opportunities")

    tab_ext, tab_cre = st.tabs(["Extensions", "Creations"])

    with tab_ext:
        st.subheader("Inventory Enhancements Identified")
        st.caption("Newly identified opportunities compared against a one-mile reference. These are not modifications to a committed old plan.")
        ext = econ[econ["event_type"] == "Extension"].copy()

        if len(ext) == 0:
            st.info("No extension events.")
        else:
            me1, me2, me3, me4 = st.columns(4)
            me1.metric("Extension Opportunities", f"{len(ext):,}")
            me2.metric("Capital Change", fmt_signed_mm(ext["capex_delta_dollars"].sum()))
            me3.metric("Incremental NPV10 Added", fmt_signed_mm(ext["incremental_npv10_added_dollars"].sum()))
            me4.metric("Reserves Change", fmt_signed_mboe(ext["reserves_delta_boe"].sum()))

            me5, me6, me7, me8 = st.columns(4)
            me5.metric("OI Change", fmt_signed_mm(ext["operating_income_delta_dollars"].sum()))
            total_inc_cap = ext["capex_delta_dollars"].sum()
            total_inc_npv = ext["npv10_delta_dollars"].sum()
            me6.metric("Portfolio Marginal NPV/Cap", fmt_ratio(_safe_div(total_inc_npv, total_inc_cap)))
            total_inc_res = ext["reserves_delta_boe"].sum()
            me7.metric("Inc Cap/Inc BOE", fmt_dollarboe(_safe_div(total_inc_cap, total_inc_res)))
            me8.metric("Inc NPV/Inc BOE", fmt_dollarboe(_safe_div(total_inc_npv, total_inc_res)))

            col_l, col_r = st.columns(2)
            with col_l:
                df_r = ext.sort_values("npv10_delta_dollars", ascending=False)
                fig = go.Figure(go.Bar(
                    x=df_r["event"].astype(str), y=df_r["npv10_delta_dollars"] / 1e6,
                    marker_color=COLOR_EXTENSION,
                ))
                _base_layout(fig, "Incremental NPV10 by Event", "Event #", "Incremental NPV10 ($MM)")
                fig.add_hline(y=0, line_dash="dash", line_color="gray")
                st.plotly_chart(fig, use_container_width=True)

            with col_r:
                fig2 = px.scatter(
                    ext, x=ext["capex_delta_dollars"] / 1e6,
                    y=ext["npv10_delta_dollars"] / 1e6,
                    size=ext["reserves_delta_boe"].clip(lower=1).values,
                    hover_data={"event": True},
                )
                _base_layout(fig2, "Incremental Capital vs Incremental NPV", "Incremental Capital ($MM)", "Incremental NPV10 ($MM)")
                st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Extension Detail Table")
            display_cols = [
                "event", COL_OLD_CURVE_1, COL_OLD_CURVE_2, COL_NEW_CURVE,
                "old_capex_dollars", "new_capex_dollars", "capex_delta_dollars",
                "old_npv10_dollars", "new_npv10_dollars", "npv10_delta_dollars",
                "old_reserves_boe", "new_reserves_boe", "reserves_delta_boe",
                "marginal_npv_to_inc_capital", "incremental_npv10_added_dollars",
            ]
            avail = [c for c in display_cols if c in ext.columns]
            st.dataframe(ext[avail].sort_values("event"), use_container_width=True, hide_index=True)

    with tab_cre:
        st.subheader("New Inventory Identified")
        st.caption("New opportunities where no prior type-curve opportunity existed. Gross opportunity economics are shown.")
        cre = econ[econ["event_type"] == "Creation"].copy()

        if len(cre) == 0:
            st.info("No creation events.")
        else:
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("New Locations", f"{len(cre):,}")
            mc2.metric("New Inventory Capital Added", fmt_signed_mm(cre["new_capex_dollars"].sum()))
            mc3.metric("Incremental NPV10 Added", fmt_signed_mm(cre["incremental_npv10_added_dollars"].sum()))
            mc4.metric("New Inventory Reserves Added", fmt_signed_mboe(cre["new_reserves_boe"].sum()))

            mc5, mc6, mc7, mc8 = st.columns(4)
            avg_npv_inv = _safe_div(cre["new_npv10_dollars"].sum(), cre["new_capex_dollars"].sum())
            mc5.metric("Avg NPV/Investment", fmt_ratio(avg_npv_inv))
            avg_cor = _safe_div(cre["new_capex_dollars"].sum(), cre["new_reserves_boe"].sum())
            mc6.metric("Avg Cost of Reserves", fmt_dollarboe(avg_cor))
            mc7.metric("Lifetime Revenue", fmt_mm(cre["new_lifetime_revenue_dollars"].sum()))
            mc8.metric("Lifetime OI", fmt_mm(cre["new_lifetime_operating_income_dollars"].sum()))

            col_l, col_r = st.columns(2)
            with col_l:
                df_r = cre.sort_values("new_npv10_dollars", ascending=False)
                fig = go.Figure(go.Bar(
                    x=df_r["event"].astype(str), y=df_r["new_npv10_dollars"] / 1e6,
                    marker_color=COLOR_CREATION,
                ))
                _base_layout(fig, "New Inventory NPV10 by Event", "Event #", "NPV10 ($MM)")
                st.plotly_chart(fig, use_container_width=True)

            with col_r:
                fig2 = px.scatter(
                    cre, x=cre["new_capex_dollars"] / 1e6,
                    y=cre["new_reserves_boe"] / 1e3,
                    color=cre[COL_NEW_CURVE],
                    hover_data={"event": True},
                    labels={"x": "Capital ($MM)", "y": "Reserves (Mboe)", "color": "Type Curve"},
                )
                _base_layout(fig2, "New Inventory: Capital vs Reserves", "Capital ($MM)", "Reserves (Mboe)")
                st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Creation Detail Table")
            display_cols = [
                "event", COL_NEW_CURVE,
                "new_capex_dollars", "new_npv10_dollars", "new_reserves_boe",
                "new_lifetime_revenue_dollars", "new_lifetime_operating_income_dollars",
                "new_payout_source", "new_ror_pct", "new_npv_inv_ratio", "new_cost_of_reserves",
                "incremental_npv10_added_dollars",
            ]
            avail = [c for c in display_cols if c in cre.columns]
            st.dataframe(cre[avail].sort_values("event"), use_container_width=True, hide_index=True)



def render_event_explorer(
    econ: pd.DataFrame,
    wells_classified: pd.DataFrame,
    indicators_clean: pd.DataFrame,
    forecasts_clean: pd.DataFrame,
    event_forecasts: pd.DataFrame,
    event_annual: pd.DataFrame,
):
    st.header("Event Explorer")

    # Event selector
    fc1, fc2 = st.columns([1, 3])
    event_type_filter = fc1.selectbox("Event Type", ["All", "Consolidation", "Extension", "Creation"], key="ee_et")
    if event_type_filter == "All":
        avail_events = sorted(econ["event"].unique())
    else:
        avail_events = sorted(econ[econ["event_type"] == event_type_filter]["event"].unique())

    if not avail_events:
        st.info("No events match filter.")
        return

    def _event_selector_label(event_number: int) -> str:
        row = econ[econ["event"] == event_number].iloc[0]
        event_type = row["event_type"]
        uwi_1 = row.get(COL_OLD_UWI_1)
        uwi_2 = row.get(COL_OLD_UWI_2)
        uwis = [str(v) for v in (uwi_1, uwi_2) if pd.notna(v) and str(v).strip()]
        if event_type == "Consolidation":
            detail = " + ".join(uwis) if uwis else row["new_type_curve_used"]
            return f"UWIs: {detail}"
        if event_type == "Extension":
            detail = uwis[0] if uwis else row["new_type_curve_used"]
            return f"Extension: {detail}"
        return f"Creation: {row['new_type_curve_used']}"

    sel_event = fc2.selectbox(
        "UWIs / Creation / Extension",
        avail_events,
        format_func=_event_selector_label,
        key="ee_ev",
    )

    ev = econ[econ["event"] == sel_event].iloc[0]

    # --- Event Header ---
    st.subheader(f"Event #{sel_event} — {ev['event_type']}")
    st.caption(ev["event_story"])

    hc1, hc2, hc3, hc4 = st.columns(4)
    hc1.markdown(f"**Old Well 1 UWI:** {ev[COL_OLD_UWI_1] or 'Not applicable'}")
    hc2.markdown(f"**Old Well 2 UWI:** {ev[COL_OLD_UWI_2] or 'Not applicable'}")
    hc3.markdown(f"**Old Curves:** {ev['old_type_curves_used'] or 'Not applicable'}")
    hc4.markdown(f"**New Curve:** {ev['new_type_curve_used']}")

    # --- Capital Efficiency ---
    st.subheader("Capital Efficiency")
    if ev["event_type"] == "Consolidation":
        ec1, ec2, ec3, ec4, ec5 = st.columns(5)
        ec1.metric("Capital Saved", fmt_mm(ev["capital_saved_dollars"]))
        ec3.metric("Cost of Reserves Improvement", fmt_dollarboe(ev["cost_of_reserves_improvement"]))
        ec4.metric("Locations Eliminated", f"{int(ev['locations_eliminated']):,}")
        ec5.metric("NPV10 Sacrificed", fmt_mm(ev["npv_sacrificed_dollars"]))
    else:
        ec1, ec2 = st.columns(2)
        ec1.metric("Incremental NPV10 Added", fmt_signed_mm(ev["incremental_npv10_added_dollars"]))
        ec2.metric("Event Type", ev["event_type"])

    # --- Financial Bridge ---
    st.subheader("Financial Bridge")

    old_label = "No Prior Opportunity" if ev["event_type"] == "Creation" else "Old / Reference"

    bridge_data = [
        ("Capital", fmt_mm(ev["old_capex_dollars"]), fmt_mm(ev["new_capex_dollars"]), fmt_signed_mm(ev["capex_delta_dollars"])),
        ("NPV10", fmt_mm(ev["old_npv10_dollars"]), fmt_mm(ev["new_npv10_dollars"]), fmt_signed_mm(ev["npv10_delta_dollars"])),
        ("Lifetime Revenue", fmt_mm(ev["old_lifetime_revenue_dollars"]), fmt_mm(ev["new_lifetime_revenue_dollars"]), fmt_signed_mm(ev["revenue_delta_dollars"])),
        ("Lifetime OI", fmt_mm(ev["old_lifetime_operating_income_dollars"]), fmt_mm(ev["new_lifetime_operating_income_dollars"]), fmt_signed_mm(ev["operating_income_delta_dollars"])),
        ("Lifetime Cash Flow", fmt_mm(ev["old_lifetime_cash_flow_dollars"]), fmt_mm(ev["new_lifetime_cash_flow_dollars"]), fmt_signed_mm(ev["cash_flow_delta_dollars"])),
        ("Reserves", fmt_mboe(ev["old_reserves_boe"]), fmt_mboe(ev["new_reserves_boe"]), fmt_signed_mboe(ev["reserves_delta_boe"])),
        ("Cost of Reserves", fmt_dollarboe(ev["old_cost_of_reserves"]), fmt_dollarboe(ev["new_cost_of_reserves_derived"]),
         fmt_signed_dollarboe(ev["cost_of_reserves_delta"]) if pd.notna(ev.get("cost_of_reserves_delta")) else "N/A"),
        ("Payout", fmt_years(ev["old_payout_years"]), fmt_years(ev["new_payout_derived"]), "—"),
        ("NPV / Investment", fmt_ratio(ev["old_npv_inv_ratio_derived"]), fmt_ratio(ev["new_npv_inv_ratio"]), "—"),
        ("OI Margin $/boe", fmt_dollarboe(ev["old_lifetime_oi_per_boe"]), fmt_dollarboe(ev["new_lifetime_oi_per_boe"]), "—"),
        ("NPV10 per BOE", fmt_dollarboe(ev["old_npv10_per_boe"]), fmt_dollarboe(ev["new_npv10_per_boe"]), "—"),
        ("Capital Saved", "—", "—", fmt_mm(ev["capital_saved_dollars"]) if ev["event_type"] == "Consolidation" else "N/A"),
        ("Cost of Reserves Improvement", "—", "—", fmt_dollarboe(ev["cost_of_reserves_improvement"]) if ev["event_type"] == "Consolidation" else "N/A"),
        ("Locations Eliminated", "—", "—", f"{int(ev['locations_eliminated']):,}" if ev["event_type"] == "Consolidation" else "N/A"),
        ("Incremental NPV10 Added", "—", "—", fmt_signed_mm(ev["incremental_npv10_added_dollars"]) if ev["event_type"] != "Consolidation" else "N/A"),
    ]

    if ev["event_type"] == "Creation":
        bridge_df = pd.DataFrame(bridge_data, columns=["Metric", old_label, "New / Opportunity", "Change "])
        # Clear old column for creations
        bridge_df[old_label] = "—"
        bridge_df["Change "] = "—"
        bridge_df.loc[bridge_df["Metric"].isin(["Capital", "NPV10", "Lifetime Revenue", "Lifetime OI",
                                                  "Lifetime Cash Flow", "Reserves"]), "Change "] = "N/A (new inventory)"
    else:
        bridge_df = pd.DataFrame(bridge_data, columns=["Metric", old_label, "New / Opportunity", "Change "])

    st.dataframe(bridge_df, use_container_width=True, hide_index=True)

    # --- Forecast Charts ---
    st.subheader("Forecast Charts")
    ev_fc = event_forecasts[event_forecasts["event"] == sel_event]
    ev_ann = event_annual[event_annual["event"] == sel_event]

    if len(ev_fc) > 0:
        mc1, mc2 = st.columns(2)
        metric_map = {
            "Revenue": ("total_revenue_dollars", "cumulative_total_revenue_dollars"),
            "Operating Income": ("operating_income_dollars", "cumulative_operating_income_dollars"),
            "Cash Flow": ("cash_flow_dollars", "cumulative_cash_flow_dollars"),
        }
        sel_fc_metric = mc1.selectbox("Forecast Metric", list(metric_map.keys()), key="ee_fm")
        sel_fc_view = mc2.selectbox("View", ["Monthly", "Cumulative", "Annual"], key="ee_fv")

        monthly_col, cum_col = metric_map[sel_fc_metric]

        if sel_fc_view == "Monthly":
            fig = px.line(
                ev_fc, x="producing_month", y=ev_fc[monthly_col] / 1e6,
                color="series_label",
                labels={"producing_month": "Producing Month", "y": f"{sel_fc_metric} ($MM)", "color": "Series"},
            )
            _base_layout(fig, f"Monthly {sel_fc_metric} — Event #{sel_event}", "Producing Month", f"{sel_fc_metric} ($MM)")
            # Style: make combined old and new plan bolder
            for trace in fig.data:
                if "Combined" in trace.name or "New" in trace.name:
                    trace.update(line=dict(width=3))
                elif "Old Well" in trace.name or "One-Mile" in trace.name:
                    trace.update(line=dict(width=1, dash="dash"))
            st.plotly_chart(fig, use_container_width=True)

        elif sel_fc_view == "Cumulative":
            fig = px.line(
                ev_fc, x="producing_month", y=ev_fc[cum_col] / 1e6,
                color="series_label",
            )
            _base_layout(fig, f"Cumulative {sel_fc_metric} — Event #{sel_event}", "Producing Month", f"Cumulative {sel_fc_metric} ($MM)")
            for trace in fig.data:
                if "Combined" in trace.name or "New" in trace.name:
                    trace.update(line=dict(width=3))
                elif "Old Well" in trace.name or "One-Mile" in trace.name:
                    trace.update(line=dict(width=1, dash="dash"))
            st.plotly_chart(fig, use_container_width=True)

        else:  # Annual
            fig = px.bar(
                ev_ann, x=COL_FORECAST_YEAR, y=ev_ann[monthly_col] / 1e6,
                color="series_label", barmode="group",
                labels={COL_FORECAST_YEAR: "Year", "y": f"{sel_fc_metric} ($MM)", "color": "Series"},
            )
            _base_layout(fig, f"Annual {sel_fc_metric} — Event #{sel_event}", "Year", f"{sel_fc_metric} ($MM)")
            st.plotly_chart(fig, use_container_width=True)

        # Chart download
        with st.expander("Download Current Chart"):
            html_bytes = plotly_to_html_bytes(fig)
            st.download_button("Download Chart HTML", html_bytes, f"event_{sel_event}_chart.html", "text/html", key="ee_dl_html")
            img = plotly_to_image_bytes(fig, "png")
            if img:
                st.download_button("Download Chart PNG", img, f"event_{sel_event}_chart.png", "image/png", key="ee_dl_png")
            else:
                st.caption("PNG export not available (Kaleido may not be installed).")
    else:
        st.info("No forecast data for this event.")

    # --- Backend Rollup Tables ---
    st.subheader("Backend Rollup Tables")

    with st.expander("Source Wells Row"):
        w_row = wells_classified[wells_classified[COL_EVENT] == sel_event]
        st.dataframe(w_row, use_container_width=True, hide_index=True)

    with st.expander("Source Economic Indicator Rows Used"):
        curves_used = set()
        if pd.notna(ev.get(COL_OLD_CURVE_1)):
            curves_used.add(ev[COL_OLD_CURVE_1])
        if pd.notna(ev.get(COL_OLD_CURVE_2)):
            curves_used.add(ev[COL_OLD_CURVE_2])
        curves_used.add(ev[COL_NEW_CURVE])
        curves_used.discard(None)
        ind_rows = indicators_clean[indicators_clean[COL_TYPE_CURVE].isin(curves_used)]
        st.dataframe(ind_rows, use_container_width=True, hide_index=True)

    with st.expander("Source Forecast Rows Used"):
        fc_rows = forecasts_clean[forecasts_clean[COL_TYPE_CURVE].isin(curves_used)]
        st.dataframe(fc_rows, use_container_width=True, hide_index=True)

    with st.expander("Event Monthly Forecast Series"):
        st.dataframe(ev_fc, use_container_width=True, hide_index=True)

    with st.expander("Event Economic Output Record"):
        ev_record = econ[econ["event"] == sel_event]
        st.dataframe(ev_record, use_container_width=True, hide_index=True)

    # --- Event Downloads ---
    st.subheader("Event Downloads")
    dl1, dl2, dl3, dl4 = st.columns(4)
    ev_record_df = econ[econ["event"] == sel_event]
    dl1.download_button("Event Summary CSV", dataframe_to_csv_bytes(ev_record_df),
                        f"event_{sel_event}_summary.csv", "text/csv", key="ee_dl_sum")
    dl2.download_button("Event Monthly Forecast CSV", dataframe_to_csv_bytes(ev_fc),
                        f"event_{sel_event}_monthly.csv", "text/csv", key="ee_dl_mfc")
    dl3.download_button("Event Annual Forecast CSV", dataframe_to_csv_bytes(ev_ann),
                        f"event_{sel_event}_annual.csv", "text/csv", key="ee_dl_afc")

    sheets = {
        "Summary": ev_record_df,
        "Monthly Forecasts": ev_fc,
        "Annual Forecasts": ev_ann,
        "Source Indicators": ind_rows,
    }
    dl4.download_button("Event Workbook XLSX", dataframes_to_excel_bytes(sheets),
                        f"event_{sel_event}_workbook.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ee_dl_xlsx")


def render_downloads(
    econ: pd.DataFrame,
    exec_summary: pd.DataFrame,
    event_forecasts: pd.DataFrame,
    event_annual: pd.DataFrame,
    validation_report: pd.DataFrame,
    wells_classified: pd.DataFrame,
    indicators_clean: pd.DataFrame,
    forecasts_clean: pd.DataFrame,
):
    st.header("Data & Downloads")

    st.subheader("Portfolio Exports")

    csv_files: dict[str, bytes] = {}

    c1, c2, c3, c4 = st.columns(4)
    b = dataframe_to_csv_bytes(econ)
    csv_files["event_economics.csv"] = b
    c1.download_button("Event Economics CSV", b, "event_economics.csv", "text/csv", key="dl_ee")

    b2 = dataframe_to_csv_bytes(exec_summary)
    csv_files["executive_summary.csv"] = b2
    c2.download_button("Executive Summary CSV", b2, "executive_summary.csv", "text/csv", key="dl_es")

    b3 = dataframe_to_csv_bytes(event_forecasts)
    csv_files["event_monthly_forecasts.csv"] = b3
    c3.download_button("Monthly Forecasts CSV", b3, "event_monthly_forecasts.csv", "text/csv", key="dl_mf")

    b4 = dataframe_to_csv_bytes(event_annual)
    csv_files["event_annual_forecasts.csv"] = b4
    c4.download_button("Annual Forecasts CSV", b4, "event_annual_forecasts.csv", "text/csv", key="dl_af")

    c5, c6, c7, _ = st.columns(4)
    b5 = dataframe_to_csv_bytes(validation_report)
    csv_files["validation_report.csv"] = b5
    c5.download_button("Validation Report CSV", b5, "validation_report.csv", "text/csv", key="dl_vr")

    # XLSX workbook
    xlsx_sheets = {
        "Executive Summary": exec_summary,
        "Event Economics": econ,
        "Monthly Forecasts": event_forecasts,
        "Annual Forecasts": event_annual,
        "Validation Report": validation_report,
        "Source Wells": wells_classified.drop(columns=["populated_old_curves"], errors="ignore"),
        "Source Indicators": indicators_clean,
        "Source Forecasts": forecasts_clean,
    }
    xlsx_bytes = dataframes_to_excel_bytes(xlsx_sheets)
    csv_files["calculated_outputs.xlsx"] = xlsx_bytes
    c6.download_button("Calculated Outputs XLSX", xlsx_bytes, "calculated_outputs.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")

    # ZIP
    zip_bytes = build_zip_bytes(csv_files)
    c7.download_button("All Outputs ZIP", zip_bytes, "calculated_outputs.zip", "application/zip", key="dl_zip")



# ════════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Viewfield Bakken Capital Efficiency",
        page_icon="📈",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.sidebar.title("Viewfield Bakken ")

    # -- Check workbook exists --
    if not WORKBOOK_PATH.exists():
        st.error(f"The required file `economics.xlsx` was not found beside `app.py`.\n\nExpected: `{WORKBOOK_PATH}`")
        st.stop()

    # -- Load --
    mtime_ns = WORKBOOK_PATH.stat().st_mtime_ns
    sheets = load_workbook(str(WORKBOOK_PATH), mtime_ns)

    wells_raw = sheets[SHEET_WELLS]
    indicators_raw = sheets[SHEET_ECONOMIC_INDICATORS]
    forecasts_raw = sheets[SHEET_FORECASTS]

    # -- Validate --
    validation_report = validate_workbook_schema(wells_raw, indicators_raw, forecasts_raw)

    # -- Classify & prepare --
    wells_classified = classify_events(wells_raw)

    indicators_clean = prepare_indicators(indicators_raw)

    # Referenced curves
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
    exec_summary = build_executive_summary(econ)

    # -- Navigation --
    page = st.sidebar.radio(
        "Navigate",
        [
            "Capital Efficiency",
            "Existing Plan Optimization",
            "Inventory Opportunities",
            "Event Explorer",
            "Data & Downloads",
        ],
    )

    if page == "Capital Efficiency":
        render_executive_summary(econ, exec_summary, event_forecasts)
    elif page == "Existing Plan Optimization":
        render_existing_plan_optimization(econ, event_forecasts)
    elif page == "Inventory Opportunities":
        render_inventory_opportunities(econ)
    elif page == "Event Explorer":
        render_event_explorer(econ, wells_classified, indicators_clean, forecasts_clean, event_forecasts, event_annual)
    elif page == "Data & Downloads":
        render_downloads(econ, exec_summary, event_forecasts, event_annual, validation_report,
                         wells_classified, indicators_clean, forecasts_clean)


if __name__ == "__main__":
    main()