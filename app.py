from __future__ import annotations

import warnings
from io import BytesIO
from pathlib import Path
from typing import Any

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
COL_PAYOUT = "Payout BTax (years)"
COL_RESERVES = "Boe WI Total (boe)"
COL_FIRST_YEAR_RATE = "1st Year Production Rate (boepd)"
COL_COST_OF_RESERVES = "Cost of Reserves ($/boe)"
COL_IP30 = "IP30 Cum (boe)"
COL_CAPEX = "Npv Investment BTax  0.0% (M$)"
COL_ROR = "BTax Disc. CF. ROR (%)"
COL_INITIAL_WI = "Initial WI (%)"
COL_THREE_MONTH_RATE = "3 Month Avg Production (boepd)"
COL_DISC_INVEST_10 = "Npv Investment BTax 10.0% (M$)"

# Column names — forecasts
COL_FORECAST_MONTH = "month #"
COL_FORECAST_YEAR = "year"
COL_FORECAST_REVENUE = "total_revenue ($M)"
COL_FORECAST_OPERATING_INCOME = "operating_income ($M)"
COL_FORECAST_VOLUME = "boe"

MONEY_TO_DOLLARS = 1000.0

DEFAULT_EXCLUDED_CURVES = {"VF_BAK_2.0M_75 (June 2026)"}

REQUIRED_WELLS_COLS = (
    COL_EVENT, COL_OLD_UWI_1, COL_OLD_UWI_2,
    COL_OLD_CURVE_1, COL_OLD_CURVE_2, COL_NEW_CURVE,
)
REQUIRED_INDICATOR_COLS = (
    COL_TYPE_CURVE, COL_NPV10, COL_PAYOUT,
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
PAL_CONSOLIDATION = "#2563EB"
PAL_EXTENSION = "#0D9488"
PAL_CREATION = "#7C3AED"
PAL_OLD = "#94A3B8"
PAL_NEW = "#1E40AF"
PAL_POSITIVE = "#16A34A"
PAL_NEGATIVE = "#DC2626"

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
    rows: list[dict[str, Any]] = []
    blocking = False

    def _add(sev: str, check: str, status: str, n: int = 0, details: str = ""):
        nonlocal blocking
        rows.append(dict(severity=sev, check=check, status=status,
                         affected_count=n, details=details))
        if sev == "BLOCKING" and status == "FAIL":
            blocking = True

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

    fc_key = forecasts[[COL_TYPE_CURVE, COL_FORECAST_YEAR, COL_FORECAST_MONTH]]
    dup_fc = fc_key.duplicated().sum()
    if dup_fc > 0:
        _add("BLOCKING", "forecast key unique", "FAIL", int(dup_fc))
    else:
        _add("INFO", "forecast key unique", "PASS")

    ref_fc = forecasts[forecasts[COL_TYPE_CURVE].isin(ref_curves)]
    curve_counts = ref_fc.groupby(COL_TYPE_CURVE).size()
    if curve_counts.nunique() > 1:
        _add("BLOCKING", "forecast period count uniform", "FAIL",
             details=str(curve_counts.value_counts().to_dict()))
    else:
        cnt = curve_counts.iloc[0] if len(curve_counts) else 0
        _add("INFO", "forecast period count uniform", "PASS",
             details=f"{cnt} periods each")

    num_cols = [
        COL_NPV10, COL_PAYOUT, COL_RESERVES,
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
    # Discounted investment at 10% for NPV/Invest ratio
    if COL_DISC_INVEST_10 in df.columns:
        df["disc_invest_10_dollars"] = pd.to_numeric(df[COL_DISC_INVEST_10], errors="coerce") * MONEY_TO_DOLLARS
    else:
        # Fallback: use undiscounted if 10% discount column missing
        df["disc_invest_10_dollars"] = df["investment_dollars"]
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
    _forecasts_clean: pd.DataFrame,
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
        old_lt_vol = sum(lt.loc[c, "lifetime_volume_boe"] for c in old_curves) if old_curves else 0.0
        old_lt_rev = sum(lt.loc[c, "lifetime_revenue_dollars"] for c in old_curves) if old_curves else 0.0
        old_lt_oi = sum(lt.loc[c, "lifetime_operating_income_dollars"] for c in old_curves) if old_curves else 0.0
        old_disc_invest = sum(ind.loc[c, "disc_invest_10_dollars"] for c in old_curves) if old_curves else 0.0

        # New-plan values
        new_capex = ind.loc[new_curve, "investment_dollars"]
        new_npv10 = ind.loc[new_curve, "npv10_dollars"]
        new_reserves = ind.loc[new_curve, COL_RESERVES]
        new_lt_vol = lt.loc[new_curve, "lifetime_volume_boe"]
        new_lt_rev = lt.loc[new_curve, "lifetime_revenue_dollars"]
        new_lt_oi = lt.loc[new_curve, "lifetime_operating_income_dollars"]
        new_payout = ind.loc[new_curve, COL_PAYOUT]
        new_ror = ind.loc[new_curve, COL_ROR]
        new_cor = ind.loc[new_curve, COL_COST_OF_RESERVES]
        new_disc_invest = ind.loc[new_curve, "disc_invest_10_dollars"]

        # ── THE BIG THREE: Capital, NPV10, EUR ──
        ev["old_capex"] = old_capex
        ev["new_capex"] = new_capex
        ev["delta_capex"] = new_capex - old_capex

        ev["old_npv10"] = old_npv10
        ev["new_npv10"] = new_npv10
        ev["delta_npv10"] = new_npv10 - old_npv10

        ev["old_eur"] = old_lt_vol
        ev["new_eur"] = new_lt_vol
        ev["delta_eur"] = new_lt_vol - old_lt_vol

        # ── INVESTMENT EFFICIENCY: NPV / Disc. Invest BTax ──
        ev["old_npv_invest_ratio"] = _safe_div(old_npv10, abs(old_disc_invest)) if old_disc_invest != 0 else np.nan
        ev["new_npv_invest_ratio"] = _safe_div(new_npv10, abs(new_disc_invest)) if new_disc_invest != 0 else np.nan
        ev["delta_npv_invest_ratio"] = (
            ev["new_npv_invest_ratio"] - ev["old_npv_invest_ratio"]
            if not (np.isnan(ev["new_npv_invest_ratio"]) or np.isnan(ev["old_npv_invest_ratio"]))
            else ev["new_npv_invest_ratio"]  # creation: old is NaN, show new
        )

        # ── ADDITIONAL MANAGEMENT METRICS ──
        # ROR (%)
        old_ror_vals = [ind.loc[c, COL_ROR] for c in old_curves] if old_curves else []
        ev["old_ror"] = np.mean(old_ror_vals) if old_ror_vals else np.nan
        ev["new_ror"] = new_ror
        ev["delta_ror"] = new_ror - ev["old_ror"] if not np.isnan(ev["old_ror"]) else np.nan

        # Payout (years)
        old_payout_vals = [ind.loc[c, COL_PAYOUT] for c in old_curves] if old_curves else []
        ev["old_payout"] = np.mean(old_payout_vals) if old_payout_vals else np.nan
        ev["new_payout"] = new_payout
        ev["delta_payout"] = new_payout - ev["old_payout"] if not np.isnan(ev["old_payout"]) else np.nan

        # Capital Efficiency: $/boe EUR
        ev["old_capex_per_eur"] = _safe_div(old_capex, old_lt_vol)
        ev["new_capex_per_eur"] = _safe_div(new_capex, new_lt_vol)
        ev["delta_capex_per_eur"] = (
            ev["new_capex_per_eur"] - ev["old_capex_per_eur"]
            if not np.isnan(ev["old_capex_per_eur"]) else np.nan
        )

        # Revenue
        ev["old_revenue"] = old_lt_rev
        ev["new_revenue"] = new_lt_rev
        ev["delta_revenue"] = new_lt_rev - old_lt_rev

        # Operating Income
        ev["old_oi"] = old_lt_oi
        ev["new_oi"] = new_lt_oi
        ev["delta_oi"] = new_lt_oi - old_lt_oi

        # Reserves (workbook field)
        ev["old_reserves"] = old_reserves
        ev["new_reserves"] = new_reserves
        ev["delta_reserves"] = new_reserves - old_reserves

        # Cost of Reserves (workbook field, new only since it's a per-well metric)
        ev["new_cost_of_reserves"] = new_cor

        # Disc invest (for reference)
        ev["old_disc_invest"] = old_disc_invest
        ev["new_disc_invest"] = new_disc_invest

        # Traceability
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
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    mm = v / 1_000_000
    if mm < 0:
        return f"({prefix}{abs(mm):,.1f} {suffix})"
    return f"{prefix}{mm:,.1f} {suffix}"


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


def _f_mboe_signed(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    mb = v / 1_000
    sign = "+" if mb >= 0 else ""
    return f"{sign}{mb:,.1f} Mboe"


def _f_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:,.1f}%"


def _f_pct_signed(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:,.1f}%"


def _f_dollarboe(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"${v:,.2f}/boe"


def _f_years(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:,.1f} yrs"


def _f_ratio(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    return f"{v:,.2f}x"


def _f_ratio_signed(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:,.2f}x"


# ────────────────────────────────────────────────────────────────────────────────
# DELTA COLOR HELPERS — green=good, red=bad
# ────────────────────────────────────────────────────────────────────────────────
def _delta_color_higher_better(val) -> str:
    """For metrics where higher is better (NPV10, EUR, ROR, NPV/Invest)."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "off"
    return "normal"  # streamlit: green=positive, red=negative


def _delta_color_lower_better(val) -> str:
    """For metrics where lower is better (Capex, Payout, $/boe)."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "off"
    return "inverse"  # streamlit: green=negative, red=positive


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
# SHARED: BIG 3 METRICS DISPLAY (Capital, NPV10, EUR) + deltas
# ────────────────────────────────────────────────────────────────────────────────
def render_big_three_metrics(df: pd.DataFrame, label_prefix: str = ""):
    """Show Old/New/Δ for Capital, NPV10, EUR — the required top-3 metrics.
    Expects df with columns: old_capex, new_capex, delta_capex,
    old_npv10, new_npv10, delta_npv10, old_eur, new_eur, delta_eur.
    Sums across all rows."""

    old_cap = df["old_capex"].sum()
    new_cap = df["new_capex"].sum()
    d_cap = df["delta_capex"].sum()

    old_npv = df["old_npv10"].sum()
    new_npv = df["new_npv10"].sum()
    d_npv = df["delta_npv10"].sum()

    old_eur = df["old_eur"].sum()
    new_eur = df["new_eur"].sum()
    d_eur = df["delta_eur"].sum()

    st.markdown("##### 💰 Capital")
    c1, c2, c3 = st.columns(3)
    c1.metric("Old Capital", _f_mm(old_cap))
    c2.metric("New Capital", _f_mm(new_cap))
    c3.metric("Δ Capital", _f_mm_signed(d_cap),
              delta=_f_mm_signed(d_cap),
              delta_color=_delta_color_lower_better(d_cap))

    st.markdown("##### 📈 NPV10")
    n1, n2, n3 = st.columns(3)
    n1.metric("Old NPV10", _f_mm(old_npv))
    n2.metric("New NPV10", _f_mm(new_npv))
    n3.metric("Δ NPV10", _f_mm_signed(d_npv),
              delta=_f_mm_signed(d_npv),
              delta_color=_delta_color_higher_better(d_npv))

    st.markdown("##### 🛢️ EUR (Forecast Lifetime Production)")
    e1, e2, e3 = st.columns(3)
    e1.metric("Old EUR", _f_mboe(old_eur))
    e2.metric("New EUR", _f_mboe(new_eur))
    e3.metric("Δ EUR", _f_mboe_signed(d_eur),
              delta=_f_mboe_signed(d_eur),
              delta_color=_delta_color_higher_better(d_eur))


def render_additional_metrics(df: pd.DataFrame, event_type: str):
    """Show additional management-critical metrics AFTER the big 3.
    Adapts by event type for relevance."""

    st.markdown("---")
    st.markdown("##### 📊 Additional Metrics")

    if event_type == "Consolidation":
        # For consolidation: show efficiency ratios, retention %s
        old_eur_total = df["old_eur"].sum()
        new_eur_total = df["new_eur"].sum()
        old_npv_total = df["old_npv10"].sum()
        new_npv_total = df["new_npv10"].sum()
        eur_retention = _safe_div(new_eur_total, old_eur_total) * 100 if old_eur_total > 0 else np.nan
        npv_retention = _safe_div(new_npv_total, old_npv_total) * 100 if old_npv_total > 0 else np.nan
        avg_new_ratio = df["new_npv_invest_ratio"].mean()
        avg_old_ratio = df["old_npv_invest_ratio"].mean()
        avg_new_ror = df["new_ror"].mean()
        avg_old_ror = df["old_ror"].mean()
        avg_new_payout = df["new_payout"].mean()
        avg_old_payout = df["old_payout"].mean()

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("EUR Retention", _f_pct(eur_retention))
        r2.metric("NPV10 Retention", _f_pct(npv_retention))
        r3.metric("Avg NPV/Invest (New)", _f_ratio(avg_new_ratio))
        r4.metric("Avg NPV/Invest (Old)", _f_ratio(avg_old_ratio))

        r5, r6, r7, r8 = st.columns(4)
        r5.metric("Avg ROR (New)", _f_pct(avg_new_ror),
                  delta=_f_pct_signed(avg_new_ror - avg_old_ror),
                  delta_color=_delta_color_higher_better(avg_new_ror - avg_old_ror))
        r6.metric("Avg ROR (Old)", _f_pct(avg_old_ror))
        r7.metric("Avg Payout (New)", _f_years(avg_new_payout),
                  delta=f"{avg_new_payout - avg_old_payout:+,.1f} yrs" if not np.isnan(avg_old_payout) else None,
                  delta_color=_delta_color_lower_better(avg_new_payout - avg_old_payout) if not np.isnan(avg_old_payout) else "off")
        r8.metric("Avg Payout (Old)", _f_years(avg_old_payout))

    elif event_type == "Extension":
        avg_new_ratio = df["new_npv_invest_ratio"].mean()
        avg_old_ratio = df["old_npv_invest_ratio"].mean()
        avg_new_ror = df["new_ror"].mean()
        avg_old_ror = df["old_ror"].mean()
        avg_new_payout = df["new_payout"].mean()
        avg_old_payout = df["old_payout"].mean()
        avg_new_cpe = df["new_capex_per_eur"].mean()
        avg_old_cpe = df["old_capex_per_eur"].mean()

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Avg NPV/Invest (New)", _f_ratio(avg_new_ratio),
                  delta=_f_ratio_signed(avg_new_ratio - avg_old_ratio),
                  delta_color=_delta_color_higher_better(avg_new_ratio - avg_old_ratio))
        r2.metric("Avg NPV/Invest (Old)", _f_ratio(avg_old_ratio))
        r3.metric("Avg $/EUR (New)", _f_dollarboe(avg_new_cpe),
                  delta=f"${avg_new_cpe - avg_old_cpe:+,.2f}/boe" if not np.isnan(avg_old_cpe) else None,
                  delta_color=_delta_color_lower_better(avg_new_cpe - avg_old_cpe) if not np.isnan(avg_old_cpe) else "off")
        r4.metric("Avg $/EUR (Old)", _f_dollarboe(avg_old_cpe))

        r5, r6, r7, r8 = st.columns(4)
        r5.metric("Avg ROR (New)", _f_pct(avg_new_ror),
                  delta=_f_pct_signed(avg_new_ror - avg_old_ror),
                  delta_color=_delta_color_higher_better(avg_new_ror - avg_old_ror))
        r6.metric("Avg ROR (Old)", _f_pct(avg_old_ror))
        r7.metric("Avg Payout (New)", _f_years(avg_new_payout),
                  delta=f"{avg_new_payout - avg_old_payout:+,.1f} yrs" if not np.isnan(avg_old_payout) else None,
                  delta_color=_delta_color_lower_better(avg_new_payout - avg_old_payout) if not np.isnan(avg_old_payout) else "off")
        r8.metric("Avg Payout (Old)", _f_years(avg_old_payout))

    elif event_type == "Creation":
        avg_ratio = df["new_npv_invest_ratio"].mean()
        avg_ror = df["new_ror"].mean()
        avg_payout = df["new_payout"].mean()
        avg_cpe = df["new_capex_per_eur"].mean()

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Avg NPV/Invest", _f_ratio(avg_ratio))
        r2.metric("Avg ROR", _f_pct(avg_ror))
        r3.metric("Avg Payout", _f_years(avg_payout))
        r4.metric("Avg Capital/EUR", _f_dollarboe(avg_cpe))


# ────────────────────────────────────────────────────────────────────────────────
# FILTER HELPER
# ────────────────────────────────────────────────────────────────────────────────
def _default_curve_selection(curves) -> list[str]:
    return [c for c in curves if c not in DEFAULT_EXCLUDED_CURVES]


def _apply_filters(df: pd.DataFrame, key_prefix: str) -> pd.DataFrame:
    with st.expander("🔍 Filters", expanded=False):
        fc1, fc2 = st.columns(2)
        curves = sorted(df[COL_NEW_CURVE].dropna().unique())
        sel_curves = fc1.multiselect(
            "New Type Curves", curves,
            default=_default_curve_selection(curves),
            key=f"{key_prefix}_nc",
        )
        events = sorted(df["event"].unique())
        sel_events = fc2.multiselect(
            "Event #s (leave blank = all)", events,
            default=[], key=f"{key_prefix}_ev",
        )
    mask = df[COL_NEW_CURVE].isin(sel_curves)
    if sel_events:
        mask = mask & df["event"].isin(sel_events)
    return df[mask]


# ════════════════════════════════════════════════════════════════════════════════
# PAGE RENDERERS
# ════════════════════════════════════════════════════════════════════════════════


# ── 1. EXECUTIVE SUMMARY ─────────────────────────────────────────────────────
def render_executive_summary(econ: pd.DataFrame):
    """Executive summary: net add/removed for Capital, NPV10, EUR across all 3 strategies.
    Tells the full portfolio story at a glance."""
    st.title("📊 Executive Summary")
    st.caption("Net impact of the inventory overhaul across all strategies")

    # ── Filters ───────────────────────────────────────────────────────────
    with st.expander("🔍 Filters", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        all_types = ["Consolidation", "Extension", "Creation"]
        sel_types = fc1.multiselect("Event Types", all_types, default=all_types, key="es_types")
        all_curves = sorted(econ[COL_NEW_CURVE].dropna().unique())
        sel_curves = fc2.multiselect("New Type Curves", all_curves,
                                     default=_default_curve_selection(all_curves), key="es_curves")
        all_events = sorted(econ["event"].unique())
        sel_events = fc3.multiselect("Event #s (leave blank = all)", all_events, default=[], key="es_events")

    mask = econ["event_type"].isin(sel_types) & econ[COL_NEW_CURVE].isin(sel_curves)
    if sel_events:
        mask = mask & econ["event"].isin(sel_events)
    filtered = econ[mask]

    consol = filtered[filtered["event_type"] == "Consolidation"]
    ext = filtered[filtered["event_type"] == "Extension"]
    cre = filtered[filtered["event_type"] == "Creation"]

    # ── NET TOTALS: The 3 numbers that matter ─────────────────────────────
    st.markdown("---")
    st.subheader("Net Portfolio Impact")

    net_cap = filtered["delta_capex"].sum()
    net_npv = filtered["delta_npv10"].sum()
    net_eur = filtered["delta_eur"].sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Net Δ Capital", _f_mm_signed(net_cap),
              delta=_f_mm_signed(net_cap),
              delta_color=_delta_color_lower_better(net_cap),
              help="Total new-plan capital minus total old-plan capital. Negative = savings.")
    k2.metric("Net Δ NPV10", _f_mm_signed(net_npv),
              delta=_f_mm_signed(net_npv),
              delta_color=_delta_color_higher_better(net_npv),
              help="Total new-plan NPV10 minus total old-plan NPV10. Positive = value created.")
    k3.metric("Net Δ EUR", _f_mboe_signed(net_eur),
              delta=_f_mboe_signed(net_eur),
              delta_color=_delta_color_higher_better(net_eur),
              help="Total new-plan EUR minus total old-plan EUR.")
    k4.metric("Total Events", f"{len(filtered):,}",
              help=f"Consolidation: {len(consol)} · Extension: {len(ext)} · Creation: {len(cre)}")

    # ── BREAKDOWN BY STRATEGY ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Δ Capital · Δ NPV10 · Δ EUR — by Strategy")

    summary_rows = []
    for label, subset, color_icon in [
        ("🔵 Consolidation", consol, PAL_CONSOLIDATION),
        ("🟢 Extension", ext, PAL_EXTENSION),
        ("🟣 Creation", cre, PAL_CREATION),
    ]:
        if subset.empty:
            summary_rows.append({
                "Strategy": label, "Events": 0,
                "Old Capital": "—", "New Capital": "—", "Δ Capital": "—",
                "Old NPV10": "—", "New NPV10": "—", "Δ NPV10": "—",
                "Old EUR": "—", "New EUR": "—", "Δ EUR": "—",
            })
        else:
            summary_rows.append({
                "Strategy": label,
                "Events": len(subset),
                "Old Capital": _f_mm(subset["old_capex"].sum()),
                "New Capital": _f_mm(subset["new_capex"].sum()),
                "Δ Capital": _f_mm_signed(subset["delta_capex"].sum()),
                "Old NPV10": _f_mm(subset["old_npv10"].sum()),
                "New NPV10": _f_mm(subset["new_npv10"].sum()),
                "Δ NPV10": _f_mm_signed(subset["delta_npv10"].sum()),
                "Old EUR": _f_mboe(subset["old_eur"].sum()),
                "New EUR": _f_mboe(subset["new_eur"].sum()),
                "Δ EUR": _f_mboe_signed(subset["delta_eur"].sum()),
            })

    # Add totals row
    summary_rows.append({
        "Strategy": "**TOTAL**",
        "Events": len(filtered),
        "Old Capital": _f_mm(filtered["old_capex"].sum()),
        "New Capital": _f_mm(filtered["new_capex"].sum()),
        "Δ Capital": _f_mm_signed(filtered["delta_capex"].sum()),
        "Old NPV10": _f_mm(filtered["old_npv10"].sum()),
        "New NPV10": _f_mm(filtered["new_npv10"].sum()),
        "Δ NPV10": _f_mm_signed(filtered["delta_npv10"].sum()),
        "Old EUR": _f_mboe(filtered["old_eur"].sum()),
        "New EUR": _f_mboe(filtered["new_eur"].sum()),
        "Δ EUR": _f_mboe_signed(filtered["delta_eur"].sum()),
    })

    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    # ── Strategy-level metric cards ───────────────────────────────────────
    for label, subset, icon in [
        ("Consolidation", consol, "🔵"),
        ("Extension", ext, "🟢"),
        ("Creation", cre, "🟣"),
    ]:
        if subset.empty:
            continue
        st.markdown("---")
        st.markdown(f"#### {icon} {label} ({len(subset)} events)")
        desc_map = {
            "Consolidation": "_Two 1-miles → one 2-mile. Save capital, retain value._",
            "Extension": "_One 1-mile → one 2-mile. Invest more capital, gain production & value._",
            "Creation": "_Net-new 2-mile. Fresh capital for new production & value._",
        }
        st.markdown(desc_map[label])

        d_cap = subset["delta_capex"].sum()
        d_npv = subset["delta_npv10"].sum()
        d_eur = subset["delta_eur"].sum()

        m1, m2, m3 = st.columns(3)
        m1.metric(f"Δ Capital", _f_mm_signed(d_cap),
                  delta=_f_mm_signed(d_cap),
                  delta_color=_delta_color_lower_better(d_cap))
        m2.metric(f"Δ NPV10", _f_mm_signed(d_npv),
                  delta=_f_mm_signed(d_npv),
                  delta_color=_delta_color_higher_better(d_npv))
        m3.metric(f"Δ EUR", _f_mboe_signed(d_eur),
                  delta=_f_mboe_signed(d_eur),
                  delta_color=_delta_color_higher_better(d_eur))

    # ── VALUE BRIDGE WATERFALL ────────────────────────────────────────────
    st.markdown("---")
    st.subheader("NPV10 Value Bridge")

    consol_d = consol["delta_npv10"].sum() if not consol.empty else 0
    ext_d = ext["delta_npv10"].sum() if not ext.empty else 0
    cre_d = cre["delta_npv10"].sum() if not cre.empty else 0
    total_d = consol_d + ext_d + cre_d

    fig_wf = go.Figure(go.Waterfall(
        orientation="v",
        measure=["relative", "relative", "relative", "total"],
        x=["Consolidation Δ", "Extension Δ", "Creation Δ", "Net NPV10 Δ"],
        y=[consol_d / 1e6, ext_d / 1e6, cre_d / 1e6, 0],
        text=[_f_mm_signed(consol_d), _f_mm_signed(ext_d),
              _f_mm_signed(cre_d), _f_mm_signed(total_d)],
        textposition="outside",
        connector=dict(line=dict(color="#CBD5E1", width=1)),
        increasing=dict(marker=dict(color=PAL_POSITIVE)),
        decreasing=dict(marker=dict(color=PAL_NEGATIVE)),
        totals=dict(marker=dict(color="#1E3A5F")),
    ))
    _apply_layout(fig_wf, "Net NPV10 Created by Inventory Overhaul ($MM)", yaxis="Δ NPV10 ($MM)")
    fig_wf.update_layout(showlegend=False)
    st.plotly_chart(fig_wf, use_container_width=True)

    # ── CAPITAL BRIDGE WATERFALL ──────────────────────────────────────────
    consol_cap_d = consol["delta_capex"].sum() if not consol.empty else 0
    ext_cap_d = ext["delta_capex"].sum() if not ext.empty else 0
    cre_cap_d = cre["delta_capex"].sum() if not cre.empty else 0
    total_cap_d = consol_cap_d + ext_cap_d + cre_cap_d

    fig_cap_wf = go.Figure(go.Waterfall(
        orientation="v",
        measure=["relative", "relative", "relative", "total"],
        x=["Consolidation Δ", "Extension Δ", "Creation Δ", "Net Capital Δ"],
        y=[consol_cap_d / 1e6, ext_cap_d / 1e6, cre_cap_d / 1e6, 0],
        text=[_f_mm_signed(consol_cap_d), _f_mm_signed(ext_cap_d),
              _f_mm_signed(cre_cap_d), _f_mm_signed(total_cap_d)],
        textposition="outside",
        connector=dict(line=dict(color="#CBD5E1", width=1)),
        increasing=dict(marker=dict(color=PAL_NEGATIVE)),   # more capital = red
        decreasing=dict(marker=dict(color=PAL_POSITIVE)),   # less capital = green
        totals=dict(marker=dict(color="#1E3A5F")),
    ))
    _apply_layout(fig_cap_wf, "Net Capital Change by Strategy ($MM)", yaxis="Δ Capital ($MM)")
    fig_cap_wf.update_layout(showlegend=False)
    st.plotly_chart(fig_cap_wf, use_container_width=True)


# ── 2. CONSOLIDATION DEEP-DIVE ───────────────────────────────────────────────
def render_consolidation(econ: pd.DataFrame, event_forecasts: pd.DataFrame):
    st.title("🔵 Consolidation — Capital Removal")
    st.markdown(
        "> **Thesis:** A single 2-mile replaces two 1-miles. For X amount *less* capital, "
        "you only lose X amount of production — and the value retained is compelling."
    )

    consol = econ[econ["event_type"] == "Consolidation"].copy()
    if consol.empty:
        st.info("No consolidation events found.")
        return

    df = _apply_filters(consol, "con")
    if df.empty:
        st.warning("No events match the selected filters.")
        return

    # ── BIG 3: Capital, NPV10, EUR (Old → New → Δ) ───────────────────────
    render_big_three_metrics(df)

    # ── ADDITIONAL METRICS ────────────────────────────────────────────────
    render_additional_metrics(df, "Consolidation")

    # ── AVERAGE TRADEOFF CHART ────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Average Capital vs. EUR Tradeoff")
    st.caption(
        "For how much less capital on average, how much less production do you give up?"
    )

    avg_old_cap = df["old_capex"].mean()
    avg_new_cap = df["new_capex"].mean()
    avg_old_eur = df["old_eur"].mean()
    avg_new_eur = df["new_eur"].mean()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Old Plan (2×1-mi avg)",
        x=["Capital ($MM)", "EUR (Mboe)"],
        y=[avg_old_cap / 1e6, avg_old_eur / 1e3],
        marker_color=PAL_OLD,
        text=[f"${avg_old_cap/1e6:,.1f}M", f"{avg_old_eur/1e3:,.0f}"],
        textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="New Plan (1×2-mi avg)",
        x=["Capital ($MM)", "EUR (Mboe)"],
        y=[avg_new_cap / 1e6, avg_new_eur / 1e3],
        marker_color=PAL_CONSOLIDATION,
        text=[f"${avg_new_cap/1e6:,.1f}M", f"{avg_new_eur/1e3:,.0f}"],
        textposition="outside",
    ))
    _apply_layout(fig, "Average per Event: Old Plan vs. New Plan", yaxis="")
    fig.update_layout(barmode="group")
    st.plotly_chart(fig, use_container_width=True)

    # ── AVERAGE NPV10 COMPARISON ──────────────────────────────────────────
    st.subheader("Average Capital vs. NPV10")
    st.caption("Capital saved vs. NPV10 retained — the value proposition.")

    avg_old_npv = df["old_npv10"].mean()
    avg_new_npv = df["new_npv10"].mean()
    avg_new_ratio = df["new_npv_invest_ratio"].mean()
    avg_old_ratio = df["old_npv_invest_ratio"].mean()

    fig2 = go.Figure()
    # Capital bars
    fig2.add_trace(go.Bar(
        name="Old Capital",
        x=["Capital ($MM)", "NPV10 ($MM)", "NPV/Invest (x)"],
        y=[avg_old_cap / 1e6, avg_old_npv / 1e6, avg_old_ratio],
        marker_color=PAL_OLD,
        text=[f"${avg_old_cap/1e6:,.1f}M", f"${avg_old_npv/1e6:,.1f}M", f"{avg_old_ratio:,.2f}x"],
        textposition="outside",
    ))
    fig2.add_trace(go.Bar(
        name="New Capital",
        x=["Capital ($MM)", "NPV10 ($MM)", "NPV/Invest (x)"],
        y=[avg_new_cap / 1e6, avg_new_npv / 1e6, avg_new_ratio],
        marker_color=PAL_CONSOLIDATION,
        text=[f"${avg_new_cap/1e6:,.1f}M", f"${avg_new_npv/1e6:,.1f}M", f"{avg_new_ratio:,.2f}x"],
        textposition="outside",
    ))
    _apply_layout(fig2, "Average per Event: Capital, NPV10, and Investment Efficiency", yaxis="")
    fig2.update_layout(barmode="group")
    st.plotly_chart(fig2, use_container_width=True)

    # ── DETAIL TABLE ──────────────────────────────────────────────────────
    with st.expander("📋 Event Detail Table", expanded=False):
        detail = df[[
            "event", COL_OLD_UWI_1, COL_OLD_UWI_2, COL_NEW_CURVE,
            "old_capex", "new_capex", "delta_capex",
            "old_npv10", "new_npv10", "delta_npv10",
            "old_eur", "new_eur", "delta_eur",
            "new_npv_invest_ratio", "new_ror", "new_payout",
        ]].copy()
        detail.columns = [
            "Event #", "Old UWI 1", "Old UWI 2", "New Curve",
            "Old Capital ($)", "New Capital ($)", "Δ Capital ($)",
            "Old NPV10 ($)", "New NPV10 ($)", "Δ NPV10 ($)",
            "Old EUR (boe)", "New EUR (boe)", "Δ EUR (boe)",
            "NPV/Invest (x)", "ROR (%)", "Payout (yrs)",
        ]
        st.dataframe(detail.sort_values("Event #"), use_container_width=True, hide_index=True)


# ── 3. EXTENSION DEEP-DIVE ───────────────────────────────────────────────────
def render_extension(econ: pd.DataFrame, event_forecasts: pd.DataFrame):
    st.title("🟢 Extension — Inventory Enhancement")
    st.markdown(
        "> **Thesis:** Extend a 1-mile to a 2-mile. For X amount *more* capital, "
        "you get X amount *more* production and value."
    )

    ext = econ[econ["event_type"] == "Extension"].copy()
    if ext.empty:
        st.info("No extension events found.")
        return

    df = _apply_filters(ext, "ext")
    if df.empty:
        st.warning("No events match the selected filters.")
        return

    # ── BIG 3 ─────────────────────────────────────────────────────────────
    render_big_three_metrics(df)

    # ── ADDITIONAL METRICS ────────────────────────────────────────────────
    render_additional_metrics(df, "Extension")

    # ── AVERAGE TRADEOFF CHART ────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Average Incremental Capital vs. EUR")
    st.caption(
        "For how much more capital on average, how much more production do you get?"
    )

    avg_old_cap = df["old_capex"].mean()
    avg_new_cap = df["new_capex"].mean()
    avg_old_eur = df["old_eur"].mean()
    avg_new_eur = df["new_eur"].mean()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Old Plan (1-mi avg)",
        x=["Capital ($MM)", "EUR (Mboe)"],
        y=[avg_old_cap / 1e6, avg_old_eur / 1e3],
        marker_color=PAL_OLD,
        text=[f"${avg_old_cap/1e6:,.1f}M", f"{avg_old_eur/1e3:,.0f}"],
        textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="New Plan (2-mi avg)",
        x=["Capital ($MM)", "EUR (Mboe)"],
        y=[avg_new_cap / 1e6, avg_new_eur / 1e3],
        marker_color=PAL_EXTENSION,
        text=[f"${avg_new_cap/1e6:,.1f}M", f"{avg_new_eur/1e3:,.0f}"],
        textposition="outside",
    ))
    _apply_layout(fig, "Average per Event: Old Plan vs. New Plan", yaxis="")
    fig.update_layout(barmode="group")
    st.plotly_chart(fig, use_container_width=True)

    # ── AVERAGE NPV10 COMPARISON ──────────────────────────────────────────
    st.subheader("Average Capital vs. NPV10")
    st.caption("Incremental capital deployed vs. incremental NPV10 earned.")

    avg_old_npv = df["old_npv10"].mean()
    avg_new_npv = df["new_npv10"].mean()
    avg_new_ratio = df["new_npv_invest_ratio"].mean()
    avg_old_ratio = df["old_npv_invest_ratio"].mean()

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        name="Old Plan (1-mi)",
        x=["Capital ($MM)", "NPV10 ($MM)", "NPV/Invest (x)"],
        y=[avg_old_cap / 1e6, avg_old_npv / 1e6, avg_old_ratio],
        marker_color=PAL_OLD,
        text=[f"${avg_old_cap/1e6:,.1f}M", f"${avg_old_npv/1e6:,.1f}M", f"{avg_old_ratio:,.2f}x"],
        textposition="outside",
    ))
    fig2.add_trace(go.Bar(
        name="New Plan (2-mi)",
        x=["Capital ($MM)", "NPV10 ($MM)", "NPV/Invest (x)"],
        y=[avg_new_cap / 1e6, avg_new_npv / 1e6, avg_new_ratio],
        marker_color=PAL_EXTENSION,
        text=[f"${avg_new_cap/1e6:,.1f}M", f"${avg_new_npv/1e6:,.1f}M", f"{avg_new_ratio:,.2f}x"],
        textposition="outside",
    ))
    _apply_layout(fig2, "Average per Event: Capital, NPV10, and Investment Efficiency", yaxis="")
    fig2.update_layout(barmode="group")
    st.plotly_chart(fig2, use_container_width=True)

    # ── DETAIL TABLE ──────────────────────────────────────────────────────
    with st.expander("📋 Event Detail Table", expanded=False):
        detail = df[[
            "event", COL_OLD_CURVE_1, COL_NEW_CURVE,
            "old_capex", "new_capex", "delta_capex",
            "old_npv10", "new_npv10", "delta_npv10",
            "old_eur", "new_eur", "delta_eur",
            "new_npv_invest_ratio", "new_ror", "new_payout",
        ]].copy()
        detail.columns = [
            "Event #", "Old Curve", "New Curve",
            "Old Capital ($)", "New Capital ($)", "Δ Capital ($)",
            "Old NPV10 ($)", "New NPV10 ($)", "Δ NPV10 ($)",
            "Old EUR (boe)", "New EUR (boe)", "Δ EUR (boe)",
            "NPV/Invest (x)", "ROR (%)", "Payout (yrs)",
        ]
        st.dataframe(detail.sort_values("Event #"), use_container_width=True, hide_index=True)


# ── 4. CREATION DEEP-DIVE ────────────────────────────────────────────────────
def render_creation(econ: pd.DataFrame):
    st.title("🟣 Creation — New Inventory")
    st.markdown(
        "> **Thesis:** Net-new 2-mile locations. For X amount of capital, "
        "you get X amount of production and value."
    )

    cre = econ[econ["event_type"] == "Creation"].copy()
    if cre.empty:
        st.info("No creation events found.")
        return

    df = _apply_filters(cre, "cre")
    if df.empty:
        st.warning("No events match the selected filters.")
        return

    # ── BIG 3 (Creation has no "old", so Old=0, New=actual, Δ=actual) ────
    render_big_three_metrics(df)

    # ── ADDITIONAL METRICS ────────────────────────────────────────────────
    render_additional_metrics(df, "Creation")

    # ── AVERAGE ECONOMICS CHART ───────────────────────────────────────────
    st.markdown("---")
    st.subheader("Average Capital, EUR, and NPV10 per New Well")
    st.caption("For X capital invested, you get X production and X value.")

    avg_cap = df["new_capex"].mean()
    avg_eur = df["new_eur"].mean()
    avg_npv = df["new_npv10"].mean()
    avg_ratio = df["new_npv_invest_ratio"].mean()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=["Capital ($MM)", "NPV10 ($MM)", "EUR (Mboe)", "NPV/Invest (x)"],
        y=[avg_cap / 1e6, avg_npv / 1e6, avg_eur / 1e3, avg_ratio],
        marker_color=PAL_CREATION,
        text=[f"${avg_cap/1e6:,.1f}M", f"${avg_npv/1e6:,.1f}M",
              f"{avg_eur/1e3:,.0f}", f"{avg_ratio:,.2f}x"],
        textposition="outside",
    ))
    _apply_layout(fig, "Average New Well Economics", yaxis="")
    fig.update_layout(showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    # ── BY-CURVE BREAKDOWN (if multiple curves) ──────────────────────────
    unique_curves = df[COL_NEW_CURVE].nunique()
    if unique_curves > 1:
        st.subheader("Average Economics by Type Curve")
        curve_agg = df.groupby(COL_NEW_CURVE).agg(
            count=("event", "count"),
            avg_capex=("new_capex", "mean"),
            avg_npv10=("new_npv10", "mean"),
            avg_eur=("new_eur", "mean"),
            avg_ratio=("new_npv_invest_ratio", "mean"),
            avg_ror=("new_ror", "mean"),
            avg_payout=("new_payout", "mean"),
        ).reset_index()

        fig_curve = go.Figure()
        for _, r in curve_agg.iterrows():
            fig_curve.add_trace(go.Bar(
                name=f"{r[COL_NEW_CURVE]} (n={r['count']})",
                x=["Avg Capital ($MM)", "Avg NPV10 ($MM)", "Avg EUR (Mboe)"],
                y=[r["avg_capex"] / 1e6, r["avg_npv10"] / 1e6, r["avg_eur"] / 1e3],
                text=[f"${r['avg_capex']/1e6:,.1f}M",
                      f"${r['avg_npv10']/1e6:,.1f}M",
                      f"{r['avg_eur']/1e3:,.0f}"],
                textposition="outside",
            ))
        _apply_layout(fig_curve, "Average Economics by Type Curve", yaxis="")
        fig_curve.update_layout(barmode="group")
        st.plotly_chart(fig_curve, use_container_width=True)

    # ── DETAIL TABLE ──────────────────────────────────────────────────────
    with st.expander("📋 Event Detail Table", expanded=False):
        detail = df[[
            "event", COL_NEW_CURVE,
            "new_capex", "new_npv10", "new_eur",
            "new_npv_invest_ratio", "new_ror", "new_payout",
            "new_capex_per_eur",
        ]].copy()
        detail.columns = [
            "Event #", "Type Curve",
            "Capital ($)", "NPV10 ($)", "EUR (boe)",
            "NPV/Invest (x)", "ROR (%)", "Payout (yrs)",
            "Capital/EUR ($/boe)",
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
    st.caption("Drill into any individual event: Old → New → Δ for all metrics, plus forecasts and source data.")

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

    # ── BIG 3: Old → New → Δ ─────────────────────────────────────────────
    st.markdown("---")

    st.markdown("##### 💰 Capital")
    c1, c2, c3 = st.columns(3)
    c1.metric("Old Capital", _f_mm(ev["old_capex"]))
    c2.metric("New Capital", _f_mm(ev["new_capex"]))
    c3.metric("Δ Capital", _f_mm_signed(ev["delta_capex"]),
              delta=_f_mm_signed(ev["delta_capex"]),
              delta_color=_delta_color_lower_better(ev["delta_capex"]))

    st.markdown("##### 📈 NPV10")
    n1, n2, n3 = st.columns(3)
    n1.metric("Old NPV10", _f_mm(ev["old_npv10"]))
    n2.metric("New NPV10", _f_mm(ev["new_npv10"]))
    n3.metric("Δ NPV10", _f_mm_signed(ev["delta_npv10"]),
              delta=_f_mm_signed(ev["delta_npv10"]),
              delta_color=_delta_color_higher_better(ev["delta_npv10"]))

    st.markdown("##### 🛢️ EUR")
    e1, e2, e3 = st.columns(3)
    e1.metric("Old EUR", _f_mboe(ev["old_eur"]))
    e2.metric("New EUR", _f_mboe(ev["new_eur"]))
    e3.metric("Δ EUR", _f_mboe_signed(ev["delta_eur"]),
              delta=_f_mboe_signed(ev["delta_eur"]),
              delta_color=_delta_color_higher_better(ev["delta_eur"]))

    # ── ADDITIONAL METRICS ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("##### 📊 Additional Metrics")

    a1, a2, a3 = st.columns(3)
    a1.metric("NPV/Invest (New)", _f_ratio(ev["new_npv_invest_ratio"]),
              delta=_f_ratio_signed(ev["delta_npv_invest_ratio"]) if not np.isnan(ev["delta_npv_invest_ratio"]) else None,
              delta_color=_delta_color_higher_better(ev["delta_npv_invest_ratio"]) if not np.isnan(ev["delta_npv_invest_ratio"]) else "off")
    a2.metric("ROR (New)", _f_pct(ev["new_ror"]),
              delta=_f_pct_signed(ev["delta_ror"]) if not np.isnan(ev["delta_ror"]) else None,
              delta_color=_delta_color_higher_better(ev["delta_ror"]) if not np.isnan(ev["delta_ror"]) else "off")
    a3.metric("Payout (New)", _f_years(ev["new_payout"]),
              delta=f"{ev['delta_payout']:+,.1f} yrs" if not np.isnan(ev["delta_payout"]) else None,
              delta_color=_delta_color_lower_better(ev["delta_payout"]) if not np.isnan(ev["delta_payout"]) else "off")

    b1, b2, b3 = st.columns(3)
    b1.metric("Capital/EUR (New)", _f_dollarboe(ev["new_capex_per_eur"]),
              delta=f"${ev['delta_capex_per_eur']:+,.2f}/boe" if not np.isnan(ev["delta_capex_per_eur"]) else None,
              delta_color=_delta_color_lower_better(ev["delta_capex_per_eur"]) if not np.isnan(ev["delta_capex_per_eur"]) else "off")
    b2.metric("Lifetime Revenue (New)", _f_mm(ev["new_revenue"]),
              delta=_f_mm_signed(ev["delta_revenue"]) if ev["delta_revenue"] != 0 else None,
              delta_color=_delta_color_higher_better(ev["delta_revenue"]))
    b3.metric("Lifetime Op. Income (New)", _f_mm(ev["new_oi"]),
              delta=_f_mm_signed(ev["delta_oi"]) if ev["delta_oi"] != 0 else None,
              delta_color=_delta_color_higher_better(ev["delta_oi"]))

    # ── FORECAST CHARTS ───────────────────────────────────────────────────
    ev_fc = event_forecasts[event_forecasts["event"] == sel_event]
    ev_ann = event_annual[event_annual["event"] == sel_event]

    if not ev_fc.empty:
        st.markdown("---")
        st.markdown("#### Production & Revenue Forecasts")
        mc1, mc2 = st.columns(2)
        metric_map = {
            "Production (boe)": ("volume_boe", "cumulative_volume_boe", "boe"),
            "Operating Income ($)": ("operating_income_dollars", "cumulative_operating_income_dollars", "$MM"),
            "Revenue ($)": ("total_revenue_dollars", "cumulative_total_revenue_dollars", "$MM"),
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

        for trace in fig.data:
            name = trace.name
            if "Combined" in name or "Old" in name or "1-Mile" in name:
                trace.update(line=dict(dash="dash", width=1.5))
            elif "New" in name or "2-Mile" in name:
                trace.update(line=dict(width=3))
        st.plotly_chart(fig, use_container_width=True)

    # ── SOURCE DATA ───────────────────────────────────────────────────────
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

    # Show validation warnings if any non-blocking issues
    warn_rows = validation_report[
        (validation_report["severity"] == "WARNING") & (validation_report["status"] == "FAIL")
    ]
    if not warn_rows.empty:
        with st.sidebar.expander("⚠️ Validation Warnings"):
            for _, r in warn_rows.iterrows():
                st.warning(f"{r['check']}: {r['details']}")

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

    # Navigation — streamlined: removed Development Timeline and Data & Downloads
    page = st.sidebar.radio(
        "Navigate",
        [
            "Executive Summary",
            "Consolidation",
            "Extension",
            "Creation",
            "Event Explorer",
        ],
        captions=[
            "Net Δ Capital · NPV10 · EUR across all strategies",
            "2-mile vs 2×1-mile: less capital, retain production",
            "1→2 mile: more capital, more production",
            "New 2-mile: capital → production → value",
            "Individual event drill-down",
        ],
    )

    if page == "Executive Summary":
        render_executive_summary(econ)
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


if __name__ == "__main__":
    main()