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

# -----------------------------
# Configuration
# -----------------------------
APP_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOKS = [
    APP_DIR / "economics.xlsx",
    APP_DIR / "economics(4).xlsx",
]

SHEET_WELLS = "wells"
SHEET_INDICATORS = "economic indicators"
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
COL_CAPEX = "Npv Investment BTax  0.0% (M$)"
COL_ROR = "BTax Disc. CF. ROR (%)"
COL_INITIAL_WI = "Initial WI (%)"
COL_THREE_MONTH_RATE = "3 Month Avg Production (boepd)"

COL_FORECAST_MONTH = "month #"
COL_FORECAST_YEAR = "year"
COL_FORECAST_BOE = "boe"
COL_FORECAST_REVENUE = "total_revenue ($M)"
COL_FORECAST_OPERATING_INCOME = "operating_income ($M)"

M_TO_DOLLARS = 1000.0  # Workbook columns labelled M$ are actually thousands of dollars.

EVENT_TYPE_MAP = {
    "Consolidation": "Existing Plan Optimization",
    "Extension": "Inventory Enhancement Identified",
    "Creation": "New Inventory Identified",
}

COLORS = {
    "navy": "#16324F",
    "blue": "#1F6E8C",
    "teal": "#2E8B8B",
    "green": "#2E7D32",
    "amber": "#D98E04",
    "red": "#B23A48",
    "purple": "#7157A8",
    "gray": "#6B7280",
    "light": "#F4F7FA",
}
EVENT_COLORS = {
    "Consolidation": COLORS["blue"],
    "Extension": COLORS["teal"],
    "Creation": COLORS["purple"],
}

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
    COL_FORECAST_BOE, COL_FORECAST_REVENUE, COL_FORECAST_OPERATING_INCOME,
)


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: #F7F9FC; }
        [data-testid="stSidebar"] { background: #16324F; }
        [data-testid="stSidebar"] * { color: white; }
        [data-testid="stMetric"] {
            background: white; border: 1px solid #E2E8F0; border-radius: 12px;
            padding: 14px 16px; box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
        }
        [data-testid="stMetricLabel"] { color: #475569; }
        h1, h2, h3 { color: #16324F; letter-spacing: -0.02em; }
        .hero {
            background: linear-gradient(120deg, #16324F, #1F6E8C);
            color: white; border-radius: 16px; padding: 24px 28px; margin-bottom: 18px;
        }
        .hero h1 { color: white; margin: 0; font-size: 2rem; }
        .hero p { margin: 6px 0 0; opacity: .88; }
        .callout {
            background: white; border-left: 5px solid #2E8B8B; border-radius: 10px;
            padding: 12px 16px; margin: 8px 0 16px; box-shadow: 0 2px 8px rgba(15,23,42,.04);
        }
        .small-muted { color: #64748B; font-size: 0.9rem; }
        div[data-testid="stDataFrame"] { background: white; border-radius: 10px; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _populated(value: Any) -> bool:
    return pd.notna(value) and str(value).strip() != ""


def _safe_div(num: float, den: float) -> float:
    if pd.isna(num) or pd.isna(den) or den == 0:
        return np.nan
    return float(num) / float(den)


def _weighted_average(values: pd.Series, weights: pd.Series) -> float:
    mask = values.notna() & weights.notna() & (weights > 0)
    if not mask.any():
        return np.nan
    return float(np.average(values[mask], weights=weights[mask]))


def find_default_workbook() -> Path | None:
    return next((p for p in DEFAULT_WORKBOOKS if p.exists()), None)


@st.cache_data(show_spinner="Loading workbook…")
def load_workbook(path: str, modified_time_ns: int) -> dict[str, pd.DataFrame]:
    del modified_time_ns
    return pd.read_excel(
        path,
        sheet_name=[SHEET_WELLS, SHEET_INDICATORS, SHEET_FORECASTS],
        engine="openpyxl",
    )


def validate_workbook_schema(
    wells: pd.DataFrame,
    indicators: pd.DataFrame,
    forecasts: pd.DataFrame,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    def add(severity: str, check: str, status: str, affected: int = 0, details: str = "") -> None:
        rows.append({
            "Severity": severity, "Check": check, "Status": status,
            "Affected": affected, "Details": details,
        })

    blocking = False
    for label, frame, required in [
        (SHEET_WELLS, wells, REQUIRED_WELLS_COLS),
        (SHEET_INDICATORS, indicators, REQUIRED_INDICATOR_COLS),
        (SHEET_FORECASTS, forecasts, REQUIRED_FORECAST_COLS),
    ]:
        missing = [c for c in required if c not in frame.columns]
        if missing:
            add("BLOCKING", f"{label}: required columns", "FAIL", len(missing), f"Missing: {missing}")
            blocking = True
        else:
            add("INFO", f"{label}: required columns", "PASS")

    if blocking:
        report = pd.DataFrame(rows)
        st.error("The workbook is missing required columns.")
        st.dataframe(report, use_container_width=True, hide_index=True)
        st.stop()

    checks = [
        (wells[COL_EVENT].isna(), "Event number is populated"),
        (wells[COL_EVENT].duplicated(), "Event number is unique"),
        (wells[COL_NEW_CURVE].isna(), "New type curve is populated"),
        (indicators[COL_TYPE_CURVE].isna(), "Indicator type curve is populated"),
        (indicators[COL_TYPE_CURVE].duplicated(), "Indicator type curve is unique"),
    ]
    for mask, label in checks:
        n = int(mask.sum())
        add("BLOCKING" if n else "INFO", label, "FAIL" if n else "PASS", n)
        blocking = blocking or n > 0

    try:
        pd.to_numeric(wells[COL_EVENT], errors="raise").astype(int)
        add("INFO", "Event number is numeric", "PASS")
    except (ValueError, TypeError):
        add("BLOCKING", "Event number is numeric", "FAIL")
        blocking = True

    numeric_indicator_cols = list(REQUIRED_INDICATOR_COLS[1:])
    for col in numeric_indicator_cols:
        converted = pd.to_numeric(indicators[col], errors="coerce")
        invalid = int((converted.isna() & indicators[col].notna()).sum())
        if invalid:
            add("BLOCKING", f"Indicator numeric: {col}", "FAIL", invalid)
            blocking = True

    numeric_forecast_cols = [
        COL_FORECAST_MONTH, COL_FORECAST_YEAR, COL_FORECAST_BOE,
        COL_FORECAST_REVENUE, COL_FORECAST_OPERATING_INCOME,
    ]
    for col in numeric_forecast_cols:
        converted = pd.to_numeric(forecasts[col], errors="coerce")
        invalid = int((converted.isna() & forecasts[col].notna()).sum())
        if invalid:
            add("BLOCKING", f"Forecast numeric: {col}", "FAIL", invalid)
            blocking = True

    ref_curves = (
        set(wells[COL_NEW_CURVE].dropna())
        | set(wells[COL_OLD_CURVE_1].dropna())
        | set(wells[COL_OLD_CURVE_2].dropna())
    )
    ref_curves.discard("")
    indicator_curves = set(indicators[COL_TYPE_CURVE].dropna())
    forecast_curves = set(forecasts[COL_TYPE_CURVE].dropna())
    for label, missing in [
        ("Referenced curves exist in indicators", ref_curves - indicator_curves),
        ("Referenced curves exist in forecasts", ref_curves - forecast_curves),
    ]:
        if missing:
            add("BLOCKING", label, "FAIL", len(missing), ", ".join(sorted(map(str, missing))))
            blocking = True
        else:
            add("INFO", label, "PASS")

    duplicate_key = forecasts.duplicated([COL_TYPE_CURVE, COL_FORECAST_YEAR, COL_FORECAST_MONTH])
    n_dup = int(duplicate_key.sum())
    add("BLOCKING" if n_dup else "INFO", "Forecast curve/year/month key is unique", "FAIL" if n_dup else "PASS", n_dup)
    blocking = blocking or n_dup > 0

    period_counts = forecasts[forecasts[COL_TYPE_CURVE].isin(ref_curves)].groupby(COL_TYPE_CURVE).size()
    if period_counts.nunique() > 1:
        add("WARNING", "Forecast period count is uniform", "WARN", details=str(period_counts.value_counts().to_dict()))
    else:
        periods = int(period_counts.iloc[0]) if len(period_counts) else 0
        add("INFO", "Forecast period count is uniform", "PASS", details=f"{periods} periods per curve")

    for uwi_col, curve_col, label in [
        (COL_OLD_UWI_1, COL_OLD_CURVE_1, "Old well 1"),
        (COL_OLD_UWI_2, COL_OLD_CURVE_2, "Old well 2"),
    ]:
        uwi_present = wells[uwi_col].map(_populated)
        curve_present = wells[curve_col].map(_populated)
        n1 = int((uwi_present & ~curve_present).sum())
        n2 = int((curve_present & ~uwi_present).sum())
        if n1:
            add("WARNING", f"{label}: UWI present but curve blank", "WARN", n1)
        if n2:
            add("WARNING", f"{label}: curve present but UWI blank", "WARN", n2)

    report = pd.DataFrame(rows)
    if blocking:
        st.error("Blocking validation errors were found.")
        st.dataframe(report, use_container_width=True, hide_index=True)
        st.stop()
    return report


def classify_events(wells: pd.DataFrame) -> pd.DataFrame:
    df = wells.copy()
    df[COL_EVENT] = pd.to_numeric(df[COL_EVENT]).astype(int)

    def classify(row: pd.Series) -> tuple[str, list[str]]:
        curves = [c for c in [row[COL_OLD_CURVE_1], row[COL_OLD_CURVE_2]] if _populated(c)]
        event_type = "Consolidation" if len(curves) == 2 else "Extension" if len(curves) == 1 else "Creation"
        return event_type, curves

    classified = df.apply(classify, axis=1)
    df["event_type"] = classified.map(lambda x: x[0])
    df["old_curves"] = classified.map(lambda x: x[1])
    df["old_curve_count"] = df["old_curves"].map(len)
    df["event_story"] = df["event_type"].map(EVENT_TYPE_MAP)
    return df


def prepare_indicators(indicators: pd.DataFrame) -> pd.DataFrame:
    df = indicators.copy()
    for col in REQUIRED_INDICATOR_COLS[1:]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["npv10_dollars"] = df[COL_NPV10] * M_TO_DOLLARS
    df["capex_dollars"] = df[COL_CAPEX] * M_TO_DOLLARS
    df["npv10_per_boe"] = df["npv10_dollars"] / df[COL_RESERVES].replace(0, np.nan)
    return df


def prepare_forecasts(forecasts: pd.DataFrame, ref_curves: set[str]) -> pd.DataFrame:
    df = forecasts[forecasts[COL_TYPE_CURVE].isin(ref_curves)].copy()
    for col in [COL_FORECAST_MONTH, COL_FORECAST_YEAR, COL_FORECAST_BOE, COL_FORECAST_REVENUE, COL_FORECAST_OPERATING_INCOME]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.sort_values([COL_TYPE_CURVE, COL_FORECAST_YEAR, COL_FORECAST_MONTH]).reset_index(drop=True)
    df["producing_month"] = df.groupby(COL_TYPE_CURVE).cumcount() + 1
    df["revenue_dollars"] = df[COL_FORECAST_REVENUE] * M_TO_DOLLARS
    df["operating_income_dollars"] = df[COL_FORECAST_OPERATING_INCOME] * M_TO_DOLLARS
    df["cumulative_boe"] = df.groupby(COL_TYPE_CURVE)[COL_FORECAST_BOE].cumsum()
    df["cumulative_revenue_dollars"] = df.groupby(COL_TYPE_CURVE)["revenue_dollars"].cumsum()
    df["cumulative_operating_income_dollars"] = df.groupby(COL_TYPE_CURVE)["operating_income_dollars"].cumsum()
    return df


def build_type_curve_lifetime(forecasts: pd.DataFrame) -> pd.DataFrame:
    result = forecasts.groupby(COL_TYPE_CURVE, as_index=False).agg(
        forecast_boe=(COL_FORECAST_BOE, "sum"),
        lifetime_revenue_dollars=("revenue_dollars", "sum"),
        lifetime_operating_income_dollars=("operating_income_dollars", "sum"),
        forecast_periods=("producing_month", "count"),
    )
    result["operating_margin_pct"] = 100 * result["lifetime_operating_income_dollars"] / result["lifetime_revenue_dollars"].replace(0, np.nan)
    result["revenue_per_boe"] = result["lifetime_revenue_dollars"] / result["forecast_boe"].replace(0, np.nan)
    result["operating_income_per_boe"] = result["lifetime_operating_income_dollars"] / result["forecast_boe"].replace(0, np.nan)
    return result


def _sum_lookup(indexed: pd.DataFrame, curves: list[str], column: str) -> float:
    return float(sum(indexed.loc[c, column] for c in curves)) if curves else 0.0


@st.cache_data(show_spinner="Calculating event economics…")
def build_event_economics(
    wells: pd.DataFrame,
    indicators: pd.DataFrame,
    lifetime: pd.DataFrame,
) -> pd.DataFrame:
    ind = indicators.set_index(COL_TYPE_CURVE)
    lt = lifetime.set_index(COL_TYPE_CURVE)
    records: list[dict[str, Any]] = []

    additive_indicators = {
        "capex_dollars": "capex_dollars",
        "npv10_dollars": "npv10_dollars",
        "reserves_boe": COL_RESERVES,
        "first_year_rate_boepd": COL_FIRST_YEAR_RATE,
        "ip30_boe": COL_IP30,
        "three_month_rate_boepd": COL_THREE_MONTH_RATE,
    }
    additive_lifetime = {
        "forecast_boe": "forecast_boe",
        "lifetime_revenue_dollars": "lifetime_revenue_dollars",
        "lifetime_operating_income_dollars": "lifetime_operating_income_dollars",
    }

    for _, row in wells.iterrows():
        old_curves: list[str] = row["old_curves"]
        new_curve = row[COL_NEW_CURVE]
        rec: dict[str, Any] = {
            "event": int(row[COL_EVENT]),
            "event_type": row["event_type"],
            "event_story": row["event_story"],
            COL_OLD_UWI_1: row[COL_OLD_UWI_1] if _populated(row[COL_OLD_UWI_1]) else None,
            COL_OLD_UWI_2: row[COL_OLD_UWI_2] if _populated(row[COL_OLD_UWI_2]) else None,
            COL_OLD_CURVE_1: row[COL_OLD_CURVE_1] if _populated(row[COL_OLD_CURVE_1]) else None,
            COL_OLD_CURVE_2: row[COL_OLD_CURVE_2] if _populated(row[COL_OLD_CURVE_2]) else None,
            COL_NEW_CURVE: new_curve,
            "old_type_curves_used": " | ".join(old_curves),
            "new_type_curve_used": new_curve,
            "old_curve_count": len(old_curves),
        }

        for output, source in additive_indicators.items():
            rec[f"old_{output}"] = _sum_lookup(ind, old_curves, source)
            rec[f"new_{output}"] = float(ind.loc[new_curve, source])
            rec[f"{output}_delta"] = rec[f"new_{output}"] - rec[f"old_{output}"]

        for output, source in additive_lifetime.items():
            rec[f"old_{output}"] = _sum_lookup(lt, old_curves, source)
            rec[f"new_{output}"] = float(lt.loc[new_curve, source])
            rec[f"{output}_delta"] = rec[f"new_{output}"] - rec[f"old_{output}"]

        rec["new_payout_years"] = float(ind.loc[new_curve, COL_PAYOUT])
        rec["new_ror_pct"] = float(ind.loc[new_curve, COL_ROR])
        rec["new_initial_wi_pct"] = float(ind.loc[new_curve, COL_INITIAL_WI])
        rec["new_npv_investment_ratio_source"] = float(ind.loc[new_curve, COL_NPV_INVESTMENT_RATIO])
        rec["new_cost_of_reserves_source"] = float(ind.loc[new_curve, COL_COST_OF_RESERVES])

        old_capex = rec["old_capex_dollars"]
        new_capex = rec["new_capex_dollars"]
        old_npv = rec["old_npv10_dollars"]
        new_npv = rec["new_npv10_dollars"]
        old_reserves = rec["old_reserves_boe"]
        new_reserves = rec["new_reserves_boe"]
        old_revenue = rec["old_lifetime_revenue_dollars"]
        new_revenue = rec["new_lifetime_revenue_dollars"]
        old_oi = rec["old_lifetime_operating_income_dollars"]
        new_oi = rec["new_lifetime_operating_income_dollars"]

        rec["old_cost_of_reserves"] = _safe_div(old_capex, old_reserves)
        rec["new_cost_of_reserves"] = _safe_div(new_capex, new_reserves)
        rec["cost_of_reserves_improvement"] = rec["old_cost_of_reserves"] - rec["new_cost_of_reserves"]
        rec["old_npv_to_capex"] = _safe_div(old_npv, old_capex)
        rec["new_npv_to_capex"] = _safe_div(new_npv, new_capex)
        rec["npv_efficiency_change"] = rec["new_npv_to_capex"] - rec["old_npv_to_capex"]
        rec["old_npv_per_boe"] = _safe_div(old_npv, old_reserves)
        rec["new_npv_per_boe"] = _safe_div(new_npv, new_reserves)
        rec["npv_per_boe_change"] = rec["new_npv_per_boe"] - rec["old_npv_per_boe"]
        rec["old_revenue_per_boe"] = _safe_div(old_revenue, rec["old_forecast_boe"])
        rec["new_revenue_per_boe"] = _safe_div(new_revenue, rec["new_forecast_boe"])
        rec["old_oi_per_boe"] = _safe_div(old_oi, rec["old_forecast_boe"])
        rec["new_oi_per_boe"] = _safe_div(new_oi, rec["new_forecast_boe"])
        rec["old_operating_margin_pct"] = 100 * _safe_div(old_oi, old_revenue)
        rec["new_operating_margin_pct"] = 100 * _safe_div(new_oi, new_revenue)

        rec["capital_saved_dollars"] = max(old_capex - new_capex, 0.0) if row["event_type"] == "Consolidation" else 0.0
        rec["locations_eliminated"] = max(len(old_curves) - 1, 0) if row["event_type"] == "Consolidation" else 0
        rec["npv_retention_pct"] = 100 * _safe_div(new_npv, old_npv) if old_npv > 0 else np.nan
        rec["reserve_retention_pct"] = 100 * _safe_div(new_reserves, old_reserves) if old_reserves > 0 else np.nan
        rec["capital_retention_pct"] = 100 * _safe_div(new_capex, old_capex) if old_capex > 0 else np.nan
        rec["capital_saved_per_boe_lost"] = _safe_div(
            max(old_capex - new_capex, 0), max(old_reserves - new_reserves, 0)
        )

        incremental_capex = new_capex - old_capex
        incremental_npv = new_npv - old_npv
        incremental_reserves = new_reserves - old_reserves
        rec["marginal_npv_to_incremental_capital"] = _safe_div(incremental_npv, incremental_capex) if incremental_capex > 0 else np.nan
        rec["incremental_capital_per_incremental_boe"] = _safe_div(incremental_capex, incremental_reserves) if incremental_reserves > 0 else np.nan
        rec["incremental_npv_per_incremental_boe"] = _safe_div(incremental_npv, incremental_reserves) if incremental_reserves > 0 else np.nan
        rec["first_year_rate_uplift_pct"] = 100 * _safe_div(
            rec["first_year_rate_boepd_delta"], rec["old_first_year_rate_boepd"]
        ) if rec["old_first_year_rate_boepd"] > 0 else np.nan
        rec["three_month_rate_uplift_pct"] = 100 * _safe_div(
            rec["three_month_rate_boepd_delta"], rec["old_three_month_rate_boepd"]
        ) if rec["old_three_month_rate_boepd"] > 0 else np.nan

        records.append(rec)
    return pd.DataFrame(records)


@st.cache_data(show_spinner="Building forecast comparisons…")
def build_event_forecasts(wells: pd.DataFrame, forecasts: pd.DataFrame) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    measure_cols = [
        "producing_month", COL_FORECAST_YEAR, COL_FORECAST_MONTH, COL_FORECAST_BOE,
        "revenue_dollars", "operating_income_dollars",
    ]
    for _, row in wells.iterrows():
        event = int(row[COL_EVENT])
        old_curves: list[str] = row["old_curves"]
        new_curve = row[COL_NEW_CURVE]

        old_parts = []
        for n, curve in enumerate(old_curves, start=1):
            frame = forecasts[forecasts[COL_TYPE_CURVE] == curve][measure_cols].copy()
            frame["series_key"] = f"old_{n}"
            frame["series_label"] = f"Old {n}: {curve}"
            frame["plan_side"] = "Old"
            old_parts.append(frame)
            frames.append(frame.assign(event=event, event_type=row["event_type"]))

        if old_parts:
            combined = pd.concat(old_parts).groupby("producing_month", as_index=False).agg({
                COL_FORECAST_YEAR: "first",
                COL_FORECAST_MONTH: "first",
                COL_FORECAST_BOE: "sum",
                "revenue_dollars": "sum",
                "operating_income_dollars": "sum",
            })
            combined["series_key"] = "combined_old"
            combined["series_label"] = "Combined Old Plan"
            combined["plan_side"] = "Old"
            frames.append(combined.assign(event=event, event_type=row["event_type"]))

        new_frame = forecasts[forecasts[COL_TYPE_CURVE] == new_curve][measure_cols].copy()
        new_frame["series_key"] = "new_plan"
        new_frame["series_label"] = f"New Plan: {new_curve}"
        new_frame["plan_side"] = "New"
        frames.append(new_frame.assign(event=event, event_type=row["event_type"]))

    result = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if not result.empty:
        for col in [COL_FORECAST_BOE, "revenue_dollars", "operating_income_dollars"]:
            result[f"cumulative_{col}"] = result.groupby(["event", "series_key"])[col].cumsum()
    return result


def build_annual_forecasts(event_forecasts: pd.DataFrame) -> pd.DataFrame:
    if event_forecasts.empty:
        return pd.DataFrame()
    return event_forecasts.groupby(
        ["event", "event_type", "series_key", "series_label", "plan_side", COL_FORECAST_YEAR],
        as_index=False,
    ).agg({
        COL_FORECAST_BOE: "sum",
        "revenue_dollars": "sum",
        "operating_income_dollars": "sum",
    })


def fmt_mm(value: float, signed: bool = False) -> str:
    if pd.isna(value):
        return "N/A"
    mm = value / 1_000_000
    if signed:
        return f"{mm:+,.1f} MM".replace("+", "+$", 1).replace("-", "-$", 1)
    return f"${mm:,.1f} MM" if mm >= 0 else f"(${abs(mm):,.1f} MM)"


def fmt_mboe(value: float, signed: bool = False) -> str:
    if pd.isna(value):
        return "N/A"
    return f"{value / 1_000:+,.1f} Mboe" if signed else f"{value / 1_000:,.1f} Mboe"


def fmt_pct(value: float, signed: bool = False) -> str:
    if pd.isna(value):
        return "N/A"
    return f"{value:+,.1f}%" if signed else f"{value:,.1f}%"


def fmt_ratio(value: float) -> str:
    return "N/M" if pd.isna(value) else f"{value:,.2f}x"


def fmt_dollar_boe(value: float, signed: bool = False) -> str:
    if pd.isna(value):
        return "N/A"
    return f"{value:+,.2f} $/boe" if signed else f"${value:,.2f}/boe"


def fmt_years(value: float) -> str:
    return "N/A" if pd.isna(value) else f"{value:,.2f} yrs"


def style_figure(fig: go.Figure, title: str, x_title: str = "", y_title: str = "") -> go.Figure:
    fig.update_layout(
        title=dict(text=title, x=0.01, font=dict(size=18, color=COLORS["navy"])),
        template="plotly_white",
        xaxis_title=x_title,
        yaxis_title=y_title,
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.28, xanchor="left", x=0),
        margin=dict(l=50, r=20, t=60, b=60),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="white",
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E8EEF5")
    return fig


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def dataframes_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False, freeze_panes=(1, 0))
            ws = writer.book[name[:31]]
            ws.sheet_view.showGridLines = False
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="16324F")
            for col_cells in ws.columns:
                width = min(max(len(str(c.value)) if c.value is not None else 0 for c in col_cells) + 2, 38)
                ws.column_dimensions[col_cells[0].column_letter].width = max(width, 10)
    buffer.seek(0)
    return buffer.read()


def render_header() -> None:
    st.markdown(
        """
        <div class="hero">
          <h1>VFB 2026 Inventory Economics</h1>
          <p>Portfolio optimization, inventory additions, and event-level economic comparisons.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def apply_filters(econ: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar.expander("Portfolio filters", expanded=True):
        selected_types = st.multiselect(
            "Event type",
            ["Consolidation", "Extension", "Creation"],
            default=["Consolidation", "Extension", "Creation"],
        )
        curves = sorted(econ[COL_NEW_CURVE].dropna().unique())
        selected_curves = st.multiselect("New type curve", curves, default=curves)
        min_npv = float(econ["new_npv10_dollars"].min() / 1_000_000)
        max_npv = float(econ["new_npv10_dollars"].max() / 1_000_000)
        npv_range = st.slider("New NPV10 range ($MM)", min_npv, max_npv, (min_npv, max_npv))
    return econ[
        econ["event_type"].isin(selected_types)
        & econ[COL_NEW_CURVE].isin(selected_curves)
        & econ["new_npv10_dollars"].between(npv_range[0] * 1_000_000, npv_range[1] * 1_000_000)
    ].copy()


def render_portfolio_summary(econ: pd.DataFrame) -> None:
    st.header("Portfolio Summary")
    if econ.empty:
        st.warning("No events match the active filters.")
        return

    new_npv = econ["new_npv10_dollars"].sum()
    new_capex = econ["new_capex_dollars"].sum()
    new_reserves = econ["new_reserves_boe"].sum()
    weighted_payout = _weighted_average(econ["new_payout_years"], econ["new_capex_dollars"])
    weighted_ror = _weighted_average(econ["new_ror_pct"], econ["new_capex_dollars"])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Events", f"{len(econ):,}")
    c2.metric("New Plan NPV10", fmt_mm(new_npv))
    c3.metric("New Plan Capex", fmt_mm(new_capex))
    c4.metric("New Plan Reserves", fmt_mboe(new_reserves))
    c5.metric("Portfolio NPV / Capex", fmt_ratio(_safe_div(new_npv, new_capex)))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Capital-weighted Payout", fmt_years(weighted_payout))
    c2.metric("Capital-weighted ROR", fmt_pct(weighted_ror))
    c3.metric("NPV10 per boe", fmt_dollar_boe(_safe_div(new_npv, new_reserves)))
    c4.metric("Cost of Reserves", fmt_dollar_boe(_safe_div(new_capex, new_reserves)))

    summary = econ.groupby("event_type", as_index=False).agg(
        Events=("event", "count"),
        New_NPV10=("new_npv10_dollars", "sum"),
        New_Capex=("new_capex_dollars", "sum"),
        New_Reserves=("new_reserves_boe", "sum"),
        Capital_Saved=("capital_saved_dollars", "sum"),
    )
    summary["NPV_to_Capex"] = summary["New_NPV10"] / summary["New_Capex"].replace(0, np.nan)
    summary["NPV_per_boe"] = summary["New_NPV10"] / summary["New_Reserves"].replace(0, np.nan)

    left, right = st.columns([1.2, 1])
    with left:
        fig = px.bar(
            summary, x="event_type", y="New_NPV10",
            color="event_type", color_discrete_map=EVENT_COLORS,
            text=summary["New_NPV10"].map(lambda x: f"${x / 1e6:,.1f}MM"),
        )
        fig.update_traces(textposition="outside")
        style_figure(fig, "NPV10 by Event Type", "", "NPV10 ($)")
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    with right:
        mix = econ["event_type"].value_counts().rename_axis("event_type").reset_index(name="count")
        fig = px.pie(mix, names="event_type", values="count", hole=0.55, color="event_type", color_discrete_map=EVENT_COLORS)
        fig.update_layout(title="Event Mix", legend=dict(orientation="h", y=-0.15), margin=dict(t=50, b=50))
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Decision Metrics by Event Type")
    display = summary.rename(columns={
        "event_type": "Event Type", "New_NPV10": "New NPV10", "New_Capex": "New Capex",
        "New_Reserves": "New Reserves", "Capital_Saved": "Capital Saved",
        "NPV_to_Capex": "NPV / Capex", "NPV_per_boe": "NPV / boe",
    })
    display["New NPV10"] = display["New NPV10"].map(fmt_mm)
    display["New Capex"] = display["New Capex"].map(fmt_mm)
    display["New Reserves"] = display["New Reserves"].map(fmt_mboe)
    display["Capital Saved"] = display["Capital Saved"].map(fmt_mm)
    display["NPV / Capex"] = display["NPV / Capex"].map(fmt_ratio)
    display["NPV / boe"] = display["NPV / boe"].map(fmt_dollar_boe)
    st.dataframe(display, use_container_width=True, hide_index=True)


def render_optimization(econ: pd.DataFrame) -> None:
    st.header("Existing Plan Optimization")
    df = econ[econ["event_type"] == "Consolidation"].copy()
    if df.empty:
        st.info("No consolidation events match the filters.")
        return

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Consolidations", f"{len(df):,}")
    c2.metric("Capital Saved", fmt_mm(df["capital_saved_dollars"].sum()))
    c3.metric("Locations Eliminated", f"{int(df['locations_eliminated'].sum()):,}")
    c4.metric("Weighted NPV Retention", fmt_pct(_weighted_average(df["npv_retention_pct"], df["old_npv10_dollars"])))
    c5.metric("Weighted Reserve Retention", fmt_pct(_weighted_average(df["reserve_retention_pct"], df["old_reserves_boe"])))

    st.markdown(
        '<div class="callout"><b>Interpretation:</b> Strong consolidations save capital while preserving a high share of old-plan NPV and reserves, and improving cost of reserves.</div>',
        unsafe_allow_html=True,
    )

    x = df["npv_retention_pct"]
    y = df["capital_saved_dollars"] / 1_000_000
    fig = px.scatter(
        df, x=x, y=y, size="old_capex_dollars", color="cost_of_reserves_improvement",
        hover_name=df["event"].map(lambda e: f"Event #{e}"),
        hover_data={"reserve_retention_pct": ":.1f", "new_npv_to_capex": ":.2f"},
        color_continuous_scale="RdYlGn",
        labels={"x": "NPV Retention (%)", "y": "Capital Saved ($MM)", "color": "CoR Improvement"},
    )
    fig.add_vline(x=90, line_dash="dash", line_color=COLORS["gray"])
    style_figure(fig, "Capital Saved vs. NPV Retention", "NPV Retention (%)", "Capital Saved ($MM)")
    st.plotly_chart(fig, use_container_width=True)

    detail_cols = [
        "event", COL_OLD_UWI_1, COL_OLD_UWI_2, COL_NEW_CURVE,
        "capital_saved_dollars", "npv_retention_pct", "reserve_retention_pct",
        "cost_of_reserves_improvement", "new_npv_to_capex", "locations_eliminated",
    ]
    detail = df[detail_cols].sort_values("capital_saved_dollars", ascending=False).copy()
    detail.columns = [
        "Event", "Old Well 1", "Old Well 2", "New Curve", "Capital Saved ($)",
        "NPV Retention (%)", "Reserve Retention (%)", "CoR Improvement ($/boe)",
        "New NPV / Capex", "Locations Eliminated",
    ]
    st.dataframe(
        detail,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Capital Saved ($)": st.column_config.NumberColumn(format="$%,.0f"),
            "NPV Retention (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "Reserve Retention (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "CoR Improvement ($/boe)": st.column_config.NumberColumn(format="$%.2f"),
            "New NPV / Capex": st.column_config.NumberColumn(format="%.2fx"),
        },
    )


def render_inventory(econ: pd.DataFrame) -> None:
    st.header("Inventory Opportunities")
    extension, creation = st.tabs(["Extensions / Enhancements", "Creations"])

    with extension:
        df = econ[econ["event_type"] == "Extension"].copy()
        if df.empty:
            st.info("No extension events match the filters.")
        else:
            inc_npv = df["npv10_dollars_delta"].sum()
            inc_capex = df["capex_dollars_delta"].sum()
            inc_reserves = df["reserves_boe_delta"].sum()
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Extensions", f"{len(df):,}")
            c2.metric("Incremental NPV10", fmt_mm(inc_npv, signed=True))
            c3.metric("Incremental Capex", fmt_mm(inc_capex, signed=True))
            c4.metric("Incremental Reserves", fmt_mboe(inc_reserves, signed=True))
            c5.metric("Marginal NPV / Capital", fmt_ratio(_safe_div(inc_npv, inc_capex)))

            chart = df.sort_values("npv10_dollars_delta", ascending=True)
            fig = px.bar(
                chart, y=chart["event"].astype(str), x=chart["npv10_dollars_delta"] / 1_000_000,
                orientation="h", color=chart["npv10_dollars_delta"], color_continuous_scale="RdYlGn",
                labels={"x": "Incremental NPV10 ($MM)", "y": "Event"},
            )
            fig.add_vline(x=0, line_color=COLORS["gray"])
            style_figure(fig, "Incremental NPV10 by Extension", "Incremental NPV10 ($MM)", "Event")
            st.plotly_chart(fig, use_container_width=True)

            cols = [
                "event", COL_OLD_UWI_1, COL_NEW_CURVE, "npv10_dollars_delta", "capex_dollars_delta",
                "reserves_boe_delta", "marginal_npv_to_incremental_capital",
                "incremental_capital_per_incremental_boe", "first_year_rate_uplift_pct",
            ]
            st.dataframe(df[cols].sort_values("npv10_dollars_delta", ascending=False), use_container_width=True, hide_index=True)

    with creation:
        df = econ[econ["event_type"] == "Creation"].copy()
        if df.empty:
            st.info("No creation events match the filters.")
        else:
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Creations", f"{len(df):,}")
            c2.metric("NPV10 Added", fmt_mm(df["new_npv10_dollars"].sum()))
            c3.metric("Capex Added", fmt_mm(df["new_capex_dollars"].sum()))
            c4.metric("Reserves Added", fmt_mboe(df["new_reserves_boe"].sum()))
            c5.metric("NPV / Capex", fmt_ratio(_safe_div(df["new_npv10_dollars"].sum(), df["new_capex_dollars"].sum())))

            fig = px.scatter(
                df, x="new_cost_of_reserves", y="new_npv_to_capex", size="new_reserves_boe",
                color="new_ror_pct", hover_name=df["event"].map(lambda e: f"Event #{e}"),
                labels={
                    "new_cost_of_reserves": "Cost of Reserves ($/boe)",
                    "new_npv_to_capex": "NPV / Capex (x)",
                    "new_ror_pct": "ROR (%)",
                },
                color_continuous_scale="Viridis",
            )
            style_figure(fig, "Creation Quality Matrix", "Cost of Reserves ($/boe)", "NPV / Capex (x)")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(
                df[["event", COL_NEW_CURVE, "new_npv10_dollars", "new_capex_dollars", "new_reserves_boe", "new_npv_to_capex", "new_payout_years", "new_ror_pct"]]
                .sort_values("new_npv10_dollars", ascending=False),
                use_container_width=True, hide_index=True,
            )


def render_rankings(econ: pd.DataFrame) -> None:
    st.header("Rankings & Concentration")
    if econ.empty:
        st.info("No events match the filters.")
        return
    metric_options = {
        "New NPV10": "new_npv10_dollars",
        "New NPV / Capex": "new_npv_to_capex",
        "New NPV per boe": "new_npv_per_boe",
        "New ROR": "new_ror_pct",
        "Capital Saved": "capital_saved_dollars",
        "Incremental NPV10": "npv10_dollars_delta",
    }
    selected = st.selectbox("Ranking metric", list(metric_options))
    metric = metric_options[selected]
    top_n = st.slider("Number of events", 5, min(30, len(econ)), min(15, len(econ)))
    ranked = econ.nlargest(top_n, metric).copy().sort_values(metric)
    fig = px.bar(
        ranked, x=metric, y=ranked["event"].astype(str), orientation="h",
        color="event_type", color_discrete_map=EVENT_COLORS,
        hover_data=[COL_NEW_CURVE], labels={metric: selected, "y": "Event"},
    )
    style_figure(fig, f"Top {top_n} Events — {selected}", selected, "Event")
    st.plotly_chart(fig, use_container_width=True)

    total_npv = econ["new_npv10_dollars"].sum()
    sorted_npv = econ.sort_values("new_npv10_dollars", ascending=False).copy()
    sorted_npv["cumulative_share"] = 100 * sorted_npv["new_npv10_dollars"].cumsum() / total_npv
    top10_share = 100 * sorted_npv.head(10)["new_npv10_dollars"].sum() / total_npv if total_npv else np.nan
    top20_share = 100 * sorted_npv.head(20)["new_npv10_dollars"].sum() / total_npv if total_npv else np.nan
    c1, c2, c3 = st.columns(3)
    c1.metric("Top 10 NPV Concentration", fmt_pct(top10_share))
    c2.metric("Top 20 NPV Concentration", fmt_pct(top20_share))
    c3.metric("Median Event NPV10", fmt_mm(econ["new_npv10_dollars"].median()))


def render_event_explorer(econ: pd.DataFrame, event_forecasts: pd.DataFrame, annual: pd.DataFrame) -> None:
    st.header("Event Explorer")
    if econ.empty:
        st.info("No events match the filters.")
        return

    left, right = st.columns([1, 2.5])
    event_type = left.selectbox("Event type", ["All", "Consolidation", "Extension", "Creation"])
    options = econ if event_type == "All" else econ[econ["event_type"] == event_type]

    def event_label(event: int) -> str:
        row = options[options["event"] == event].iloc[0]
        uwis = [str(row[c]) for c in [COL_OLD_UWI_1, COL_OLD_UWI_2] if _populated(row[c])]
        detail = " + ".join(uwis) if uwis else row[COL_NEW_CURVE]
        return f"#{event} · {row['event_type']} · {detail}"

    selected_event = right.selectbox("Event", sorted(options["event"].unique()), format_func=event_label)
    ev = options[options["event"] == selected_event].iloc[0]
    st.subheader(f"Event #{selected_event} — {ev['event_type']}")
    st.caption(ev["event_story"])

    st.markdown(
        f"**Old curves:** {ev['old_type_curves_used'] or 'None'}  &nbsp;&nbsp;|&nbsp;&nbsp; **New curve:** {ev['new_type_curve_used']}",
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("New NPV10", fmt_mm(ev["new_npv10_dollars"]), fmt_mm(ev["npv10_dollars_delta"], signed=True))
    c2.metric("New Capex", fmt_mm(ev["new_capex_dollars"]), fmt_mm(ev["capex_dollars_delta"], signed=True))
    c3.metric("New Reserves", fmt_mboe(ev["new_reserves_boe"]), fmt_mboe(ev["reserves_boe_delta"], signed=True))
    c4.metric("NPV / Capex", fmt_ratio(ev["new_npv_to_capex"]))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Payout", fmt_years(ev["new_payout_years"]))
    c2.metric("ROR", fmt_pct(ev["new_ror_pct"]))
    c3.metric("Cost of Reserves", fmt_dollar_boe(ev["new_cost_of_reserves"]))
    c4.metric("NPV per boe", fmt_dollar_boe(ev["new_npv_per_boe"]))

    ev_fc = event_forecasts[event_forecasts["event"] == selected_event]
    ev_annual = annual[annual["event"] == selected_event]
    if ev_fc.empty:
        st.info("No forecast rows found for this event.")
        return

    metric_map = {
        "Production": (COL_FORECAST_BOE, f"cumulative_{COL_FORECAST_BOE}", "boe"),
        "Revenue": ("revenue_dollars", "cumulative_revenue_dollars", "$"),
        "Operating Income": ("operating_income_dollars", "cumulative_operating_income_dollars", "$"),
    }
    m1, m2 = st.columns(2)
    metric_name = m1.selectbox("Forecast metric", list(metric_map))
    view = m2.selectbox("View", ["Monthly", "Cumulative", "Annual"])
    monthly_col, cumulative_col, unit = metric_map[metric_name]

    if view == "Annual":
        plot_df = ev_annual.copy()
        y = monthly_col
        scale = 1_000_000 if unit == "$" else 1_000
        plot_df["plot_value"] = plot_df[y] / scale
        fig = px.bar(plot_df, x=COL_FORECAST_YEAR, y="plot_value", color="series_label", barmode="group")
        y_label = f"{metric_name} ({'$MM' if unit == '$' else 'Mboe'})"
        style_figure(fig, f"Annual {metric_name}", "Year", y_label)
    else:
        plot_df = ev_fc.copy()
        y = monthly_col if view == "Monthly" else cumulative_col
        scale = 1_000_000 if unit == "$" else 1_000
        plot_df["plot_value"] = plot_df[y] / scale
        fig = px.line(plot_df, x="producing_month", y="plot_value", color="series_label")
        for trace in fig.data:
            trace.update(line=dict(width=3 if "New Plan" in trace.name or "Combined" in trace.name else 1.5, dash="solid" if "New Plan" in trace.name or "Combined" in trace.name else "dash"))
        y_label = f"{'Cumulative ' if view == 'Cumulative' else ''}{metric_name} ({'$MM' if unit == '$' else 'Mboe'})"
        style_figure(fig, f"{view} {metric_name}", "Producing Month", y_label)
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("Event economics record"):
        st.dataframe(pd.DataFrame([ev]), use_container_width=True, hide_index=True)


def render_data_downloads(
    econ: pd.DataFrame,
    event_forecasts: pd.DataFrame,
    annual: pd.DataFrame,
    validation: pd.DataFrame,
    wells: pd.DataFrame,
    indicators: pd.DataFrame,
    forecasts: pd.DataFrame,
) -> None:
    st.header("Data & Downloads")
    st.caption("Exports use the active portfolio filters where applicable.")
    files: dict[str, bytes] = {
        "event_economics.csv": dataframe_to_csv_bytes(econ),
        "event_monthly_forecasts.csv": dataframe_to_csv_bytes(event_forecasts[event_forecasts["event"].isin(econ["event"])]),
        "event_annual_forecasts.csv": dataframe_to_csv_bytes(annual[annual["event"].isin(econ["event"])]),
        "validation_report.csv": dataframe_to_csv_bytes(validation),
    }
    workbook = dataframes_to_excel_bytes({
        "Event Economics": econ,
        "Monthly Forecasts": event_forecasts[event_forecasts["event"].isin(econ["event"])],
        "Annual Forecasts": annual[annual["event"].isin(econ["event"])],
        "Validation": validation,
        "Source Wells": wells.drop(columns=["old_curves"], errors="ignore"),
        "Source Indicators": indicators,
        "Source Forecasts": forecasts,
    })
    files["calculated_outputs.xlsx"] = workbook

    c1, c2, c3 = st.columns(3)
    c1.download_button("Event Economics CSV", files["event_economics.csv"], "event_economics.csv", "text/csv")
    c2.download_button("Calculated Outputs XLSX", workbook, "calculated_outputs.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)
    c3.download_button("All Outputs ZIP", zip_buffer.getvalue(), "calculated_outputs.zip", "application/zip")

    with st.expander("Workbook validation report", expanded=False):
        st.dataframe(validation, use_container_width=True, hide_index=True)


def main() -> None:
    st.set_page_config(
        page_title="VFB 2026 Inventory Economics",
        page_icon="📈",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_css()
    render_header()

    st.sidebar.title("VFB Economics")
    default_workbook = find_default_workbook()
    upload = st.sidebar.file_uploader("Workbook", type=["xlsx"], help="Optional: upload a workbook with the required three sheets.")

    if upload is not None:
        uploaded_path = APP_DIR / ".uploaded_economics.xlsx"
        uploaded_path.write_bytes(upload.getvalue())
        workbook_path = uploaded_path
    elif default_workbook is not None:
        workbook_path = default_workbook
    else:
        st.error("Place `economics.xlsx` beside this script or upload the workbook in the sidebar.")
        st.stop()

    st.sidebar.caption(f"Using: {workbook_path.name}")
    sheets = load_workbook(str(workbook_path), workbook_path.stat().st_mtime_ns)
    wells_raw = sheets[SHEET_WELLS]
    indicators_raw = sheets[SHEET_INDICATORS]
    forecasts_raw = sheets[SHEET_FORECASTS]

    validation = validate_workbook_schema(wells_raw, indicators_raw, forecasts_raw)
    wells = classify_events(wells_raw)
    indicators = prepare_indicators(indicators_raw)
    ref_curves = set(wells[COL_NEW_CURVE].dropna()) | set(wells[COL_OLD_CURVE_1].dropna()) | set(wells[COL_OLD_CURVE_2].dropna())
    ref_curves.discard("")
    forecasts = prepare_forecasts(forecasts_raw, ref_curves)
    lifetime = build_type_curve_lifetime(forecasts)
    econ = build_event_economics(wells, indicators, lifetime)
    event_forecasts = build_event_forecasts(wells, forecasts)
    annual = build_annual_forecasts(event_forecasts)

    filtered = apply_filters(econ)
    page = st.sidebar.radio(
        "Navigate",
        [
            "Portfolio Summary",
            "Existing Plan Optimization",
            "Inventory Opportunities",
            "Rankings & Concentration",
            "Event Explorer",
            "Data & Downloads",
        ],
    )

    if page == "Portfolio Summary":
        render_portfolio_summary(filtered)
    elif page == "Existing Plan Optimization":
        render_optimization(filtered)
    elif page == "Inventory Opportunities":
        render_inventory(filtered)
    elif page == "Rankings & Concentration":
        render_rankings(filtered)
    elif page == "Event Explorer":
        render_event_explorer(filtered, event_forecasts, annual)
    else:
        render_data_downloads(filtered, event_forecasts, annual, validation, wells, indicators, forecasts)

    st.sidebar.markdown("---")
    st.sidebar.caption("All economics are before-tax and based on workbook type-curve inputs.")


if __name__ == "__main__":
    main()