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
    """Clean category summary containing only the approved KPIs and counts."""
    rows: list[dict[str, Any]] = []

    consol = econ[econ["event_type"] == "Consolidation"]
    ext = econ[econ["event_type"] == "Extension"]
    cre = econ[econ["event_type"] == "Creation"]

    rows.append({
        "category": "Consolidation",
        "event_story": EVENT_TYPE_MAP["Consolidation"],
        "count": len(consol),
        "total_capital_saved_dollars": consol["capital_saved_dollars"].sum(),
        "average_cost_of_reserves_improvement": consol["cost_of_reserves_improvement"].mean(),
        "total_npv10_to_total_capex": total_npv_to_capex(consol),
        "incremental_npv10_dollars": np.nan,
        "incremental_npv10_pct": np.nan,
        "total_npv10_added_dollars": np.nan,
    })
    rows.append({
        "category": "Extension",
        "event_story": EVENT_TYPE_MAP["Extension"],
        "count": len(ext),
        "total_capital_saved_dollars": np.nan,
        "average_cost_of_reserves_improvement": np.nan,
        "total_npv10_to_total_capex": np.nan,
        "incremental_npv10_dollars": ext["npv10_delta_dollars"].sum(),
        "incremental_npv10_pct": extension_npv_uplift_pct(ext),
        "total_npv10_added_dollars": np.nan,
    })
    rows.append({
        "category": "Creation",
        "event_story": EVENT_TYPE_MAP["Creation"],
        "count": len(cre),
        "total_capital_saved_dollars": np.nan,
        "average_cost_of_reserves_improvement": np.nan,
        "total_npv10_to_total_capex": np.nan,
        "incremental_npv10_dollars": np.nan,
        "incremental_npv10_pct": np.nan,
        "total_npv10_added_dollars": cre["new_npv10_dollars"].sum(),
    })
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


def extension_npv_uplift_pct(df: pd.DataFrame) -> float:
    if df.empty:
        return np.nan
    return 100.0 * _safe_div(
        df["npv10_delta_dollars"].sum(),
        df["old_npv10_dollars"].sum(),
    )


def total_npv_to_capex(df: pd.DataFrame) -> float:
    if df.empty:
        return np.nan
    return _safe_div(
        df["new_npv10_dollars"].sum(),
        df["new_capex_dollars"].sum(),
    )


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
    st.header("Portfolio Summary")
    st.caption("Approved KPIs only: category counts and the specified economic measures.")

    consol = econ[econ["event_type"] == "Consolidation"]
    ext = econ[econ["event_type"] == "Extension"]
    cre = econ[econ["event_type"] == "Creation"]

    st.subheader("Consolidation")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Count", f"{len(consol):,}")
    c2.metric("Total Capital Saved", fmt_mm(consol["capital_saved_dollars"].sum()))
    c3.metric(
        "Average Cost of Reserves Improvement",
        fmt_dollarboe(consol["cost_of_reserves_improvement"].mean()),
    )
    c4.metric("Total NPV10 / Total Capex", fmt_ratio(total_npv_to_capex(consol)))

    st.subheader("Extension / Enhancement")
    e1, e2, e3 = st.columns(3)
    e1.metric("Count", f"{len(ext):,}")
    e2.metric("Incremental NPV10", fmt_signed_mm(ext["npv10_delta_dollars"].sum()))
    e3.metric("Incremental NPV10 %", fmt_pct(extension_npv_uplift_pct(ext)))

    st.subheader("Creation")
    n1, n2 = st.columns(2)
    n1.metric("Count", f"{len(cre):,}")
    n2.metric("Total NPV10 Added", fmt_mm(cre["new_npv10_dollars"].sum()))

    st.subheader("Category Summary")
    display = exec_summary.copy()
    display = display.rename(columns={
        "category": "Category",
        "event_story": "Description",
        "count": "Count",
        "total_capital_saved_dollars": "Total Capital Saved",
        "average_cost_of_reserves_improvement": "Average Cost of Reserves Improvement",
        "total_npv10_to_total_capex": "Total NPV10 / Total Capex",
        "incremental_npv10_dollars": "Incremental NPV10",
        "incremental_npv10_pct": "Incremental NPV10 %",
        "total_npv10_added_dollars": "Total NPV10 Added",
    })
    for col in ["Total Capital Saved", "Incremental NPV10", "Total NPV10 Added"]:
        display[col] = display[col].map(fmt_mm)
    display["Average Cost of Reserves Improvement"] = display["Average Cost of Reserves Improvement"].map(fmt_dollarboe)
    display["Total NPV10 / Total Capex"] = display["Total NPV10 / Total Capex"].map(fmt_ratio)
    display["Incremental NPV10 %"] = display["Incremental NPV10 %"].map(fmt_pct)
    st.dataframe(display, use_container_width=True, hide_index=True)


def render_existing_plan_optimization(econ: pd.DataFrame, event_forecasts: pd.DataFrame):
    st.header("Existing Plan Optimization: Consolidations")
    consol = econ[econ["event_type"] == "Consolidation"].copy()
    if consol.empty:
        st.info("No consolidation events found.")
        return

    with st.expander("Filters", expanded=False):
        curves = sorted(consol[COL_NEW_CURVE].dropna().unique())
        selected = st.multiselect("New Type Curve", curves, default=curves, key="opt_nc")
    filtered = consol[consol[COL_NEW_CURVE].isin(selected)]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Count", f"{len(filtered):,}")
    m2.metric("Total Capital Saved", fmt_mm(filtered["capital_saved_dollars"].sum()))
    m3.metric("Average Cost of Reserves Improvement", fmt_dollarboe(filtered["cost_of_reserves_improvement"].mean()))
    m4.metric("Total NPV10 / Total Capex", fmt_ratio(total_npv_to_capex(filtered)))

    st.subheader("Event Detail")
    detail = filtered[[
        "event", COL_OLD_UWI_1, COL_OLD_UWI_2, COL_OLD_CURVE_1, COL_OLD_CURVE_2, COL_NEW_CURVE,
        "capital_saved_dollars", "cost_of_reserves_improvement",
        "new_npv10_dollars", "new_capex_dollars",
    ]].copy()
    detail["npv10_to_capex"] = detail.apply(
        lambda r: _safe_div(r["new_npv10_dollars"], r["new_capex_dollars"]), axis=1
    )
    st.dataframe(detail.sort_values("event"), use_container_width=True, hide_index=True)


def render_inventory_opportunities(econ: pd.DataFrame):
    st.header("Inventory Opportunities")
    tab_ext, tab_cre = st.tabs(["Extensions / Enhancements", "Creations"])

    with tab_ext:
        ext = econ[econ["event_type"] == "Extension"].copy()
        if ext.empty:
            st.info("No extension events.")
        else:
            a1, a2, a3 = st.columns(3)
            a1.metric("Count", f"{len(ext):,}")
            a2.metric("Incremental NPV10", fmt_signed_mm(ext["npv10_delta_dollars"].sum()))
            a3.metric("Incremental NPV10 %", fmt_pct(extension_npv_uplift_pct(ext)))

            chart_df = ext.sort_values("npv10_delta_dollars", ascending=False)
            fig = go.Figure(go.Bar(
                x=chart_df["event"].astype(str),
                y=chart_df["npv10_delta_dollars"] / 1e6,
                marker_color=COLOR_EXTENSION,
            ))
            _base_layout(fig, "Incremental NPV10 by Extension Event", "Event #", "Incremental NPV10 ($MM)")
            fig.add_hline(y=0, line_dash="dash", line_color="gray")
            st.plotly_chart(fig, use_container_width=True)

            detail = ext[[
                "event", COL_OLD_CURVE_1, COL_OLD_CURVE_2, COL_NEW_CURVE,
                "old_npv10_dollars", "new_npv10_dollars", "npv10_delta_dollars",
            ]].copy()
            detail["incremental_npv10_pct"] = 100.0 * detail.apply(
                lambda r: _safe_div(r["npv10_delta_dollars"], r["old_npv10_dollars"]), axis=1
            )
            st.dataframe(detail.sort_values("event"), use_container_width=True, hide_index=True)

    with tab_cre:
        cre = econ[econ["event_type"] == "Creation"].copy()
        if cre.empty:
            st.info("No creation events.")
        else:
            b1, b2 = st.columns(2)
            b1.metric("Count", f"{len(cre):,}")
            b2.metric("Total NPV10 Added", fmt_mm(cre["new_npv10_dollars"].sum()))

            chart_df = cre.sort_values("new_npv10_dollars", ascending=False)
            fig = go.Figure(go.Bar(
                x=chart_df["event"].astype(str),
                y=chart_df["new_npv10_dollars"] / 1e6,
                marker_color=COLOR_CREATION,
            ))
            _base_layout(fig, "NPV10 Added by Creation Event", "Event #", "NPV10 Added ($MM)")
            st.plotly_chart(fig, use_container_width=True)

            detail = cre[["event", COL_NEW_CURVE, "new_npv10_dollars"]].copy()
            st.dataframe(detail.sort_values("event"), use_container_width=True, hide_index=True)


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

    # --- Approved event KPIs ---
    st.subheader("Approved Event KPIs")
    if ev["event_type"] == "Consolidation":
        k1, k2, k3 = st.columns(3)
        k1.metric("Capital Saved", fmt_mm(ev["capital_saved_dollars"]))
        k2.metric("Cost of Reserves Improvement", fmt_dollarboe(ev["cost_of_reserves_improvement"]))
        k3.metric("NPV10 / Capex", fmt_ratio(_safe_div(ev["new_npv10_dollars"], ev["new_capex_dollars"])))
    elif ev["event_type"] == "Extension":
        k1, k2 = st.columns(2)
        k1.metric("Incremental NPV10", fmt_signed_mm(ev["npv10_delta_dollars"]))
        k2.metric("Incremental NPV10 %", fmt_pct(100.0 * _safe_div(ev["npv10_delta_dollars"], ev["old_npv10_dollars"])))
    else:
        k1, k2 = st.columns(2)
        k1.metric("Count", "1")
        k2.metric("NPV10 Added", fmt_mm(ev["new_npv10_dollars"]))

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
        page_title="VFB 2026 Inventory Overhaul",
        page_icon="📈",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.sidebar.title("Contents")

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
            "Portfolio Summary",
            "Existing Plan Optimization",
            "Inventory Opportunities",
            "Event Explorer",
            "Data & Downloads",
        ],
    )

    if page == "Portfolio Summary":
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