"""
=============================================================================
 INVENTORY REDESIGN  —  CAPITAL EFFICIENCY DASHBOARD
=============================================================================

 Compares BEFORE inventory (legacy wells) with AFTER inventory (enhanced wells)
 across three event types:

     Consolidation : 2 legacy wells (w1 + w2) -> 1 enhanced well  (net -1 well)
     Extension     : 1 legacy well  (w1)      -> 1 enhanced well  (net  0 wells)
     Creation      : 0 legacy wells           -> 1 enhanced well  (net +1 well)

 "Enhanced entity" = what the combined/extended/created well becomes.
 w1_ent, w2_ent, and enh_ent values all exist in the other spreadsheets
 (economics.xlsx, capex.xlsx, forecast.xlsx) and are joined by entity name.

 ─────────────────────────────────────────────────────────────────────────────
 EXPECTED XLSX FORMATS (place next to app.py or in ./data):
 ─────────────────────────────────────────────────────────────────────────────

 economics.xlsx  (one row per entity — Aries/PHDWin economics export)
 ┌─────────────────────────────────────┬────────────────────────────────────┐
 │ Column                              │ Notes                              │
 ├─────────────────────────────────────┼────────────────────────────────────┤
 │ Entity                              │ entity name (text, join key)       │
 │ Npv Cash Flow BTax 10.0% (M$)      │ NPV before-tax at 10% disc (M$)   │
 │ Npv Investment BTax 0.0% (M$)      │ total investment before-tax (M$)   │
 │ Payout BTax (years)                 │ payout in years                    │
 │ Boe WI Total (boe)                 │ net reserves in BOE                │
 │ 1st Year Production Rate (boepd)   │ first-year avg rate                │
 │ Cost of Reserves ($/boe)           │ reported cost of reserves          │
 │ IP30 Cum (boe)                     │ IP30 cumulative production         │
 │ BTax Disc. CF. ROR (%)             │ rate of return %                   │
 │ Initial WI (%)                     │ working interest %                 │
 │ 3 Month Avg Production (boepd)     │ 3-month average rate               │
 └─────────────────────────────────────┴────────────────────────────────────┘
   NOTE: NPV and Investment columns are in M$ (thousands). The app multiplies
   by 1,000 to get base dollars. All other columns are in native units.

 consol.xlsx  (one row per redesign event — the mapping file)
 ┌─────────────────────────────────────┬────────────────────────────────────┐
 │ Column                              │ Notes                              │
 ├─────────────────────────────────────┼────────────────────────────────────┤
 │ consolidation #                     │ event ID (text, optional)          │
 │ well 1 name                         │ display name for well 1 (optional) │
 │ well 2 name                         │ display name for well 2 (optional) │
 │ well 1 entity                       │ entity name for before-well 1      │
 │ well 2 entity                       │ entity name for before-well 2      │
 │ enhanced well entity                │ entity name for the enhanced well  │
 └─────────────────────────────────────┴────────────────────────────────────┘
   - Consolidation: both well 1 entity and well 2 entity populated
   - Extension: only well 1 entity populated (well 2 blank)
   - Creation: both well 1 and well 2 blank (only enhanced populated)

 forecast.xlsx  (one row per entity per month — cash flow forecast)
 ┌─────────────────────────────────────┬────────────────────────────────────┐
 │ Column                              │ Notes                              │
 ├─────────────────────────────────────┼────────────────────────────────────┤
 │ entity_name                         │ entity name (text, join key)       │
 │ year                                │ calendar year (int)                │
 │ month                               │ calendar month 1-12 (int)         │
 │ total_revenue                       │ monthly revenue ($)                │
 │ operating_income                    │ monthly operating income ($)       │
 │ cash_flow                           │ monthly net cash flow ($)          │
 └─────────────────────────────────────┴────────────────────────────────────┘

 capex.xlsx  (one row per entity — capital cost)
 ┌─────────────────────────────────────┬────────────────────────────────────┐
 │ Column                              │ Notes                              │
 ├─────────────────────────────────────┼────────────────────────────────────┤
 │ entity                              │ entity name (text, join key)       │
 │ capex                               │ total capex ($)                    │
 └─────────────────────────────────────┴────────────────────────────────────┘

 Run:  streamlit run app.py
=============================================================================
"""

from __future__ import annotations

import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

# ─── CONFIG ──────────────────────────────────────────────────────────────────

APP_TITLE = "Inventory Redesign — Capital Efficiency"
APP_SUBTITLE = "Upstream Portfolio Value Creation Review"
BASE_DIR = Path(__file__).resolve().parent
SEARCH_DIRS = [BASE_DIR, BASE_DIR / "data"]

ECON_MONEY_SCALE = 1_000.0  # Aries M$ -> base $

CASE_ORDER = ["Consolidation", "Extension", "Creation"]

CLR = {
    "navy": "#12314F", "blue": "#0E6BA8", "teal": "#0E9AA7",
    "amber": "#C98A1B", "green": "#2E7D5B", "red": "#B3403A",
    "slate": "#8A9AA9", "light": "#D8E1E8", "ink": "#12222E",
    "before": "#8A9AA9", "after": "#0E6BA8", "incr": "#0E9AA7",
}
CASE_CLR = {
    "Consolidation": CLR["navy"],
    "Extension": CLR["teal"],
    "Creation": CLR["amber"],
}

st.set_page_config(page_title=APP_TITLE, page_icon="⛽", layout="wide",
                   initial_sidebar_state="expanded")

# ─── STYLING ─────────────────────────────────────────────────────────────────

CSS = """
<style>
  .block-container {padding-top:1.6rem; padding-bottom:3rem; max-width:1500px;}
  h1,h2,h3,h4 {color:#12222E;}
  .app-header {border-left:5px solid #0E6BA8; padding:.35rem 0 .35rem .9rem; margin-bottom:1.1rem;}
  .app-header .t {font-size:1.55rem; font-weight:700; color:#12314F; line-height:1.15;}
  .app-header .s {font-size:.86rem; color:#66788A; text-transform:uppercase; letter-spacing:.10em; margin-top:.15rem;}
  .sec {display:flex; align-items:center; gap:.6rem; margin:1.7rem 0 .7rem 0; padding-bottom:.4rem; border-bottom:1px solid #E3E9ED;}
  .sec .n {font-size:.72rem; font-weight:700; color:#FFF; background:#12314F; border-radius:3px; padding:.10rem .42rem;}
  .sec .h {font-size:1.05rem; font-weight:650; color:#12314F;}
  .sec .d {font-size:.80rem; color:#8A9AA9; margin-left:auto;}
  .kpi {background:#FFF; border:1px solid #E3E9ED; border-radius:8px; padding:.75rem .85rem; height:100%; box-shadow:0 1px 2px rgba(18,49,79,.05); border-top:3px solid #D8E1E8;}
  .kpi.pos {border-top-color:#2E7D5B;} .kpi.neg {border-top-color:#B3403A;} .kpi.acc {border-top-color:#0E6BA8;}
  .kpi .l {font-size:.685rem; font-weight:600; color:#7A8A99; text-transform:uppercase; letter-spacing:.055em; line-height:1.25; min-height:2.0em;}
  .kpi .v {font-size:1.42rem; font-weight:700; color:#12222E; line-height:1.25; margin-top:.18rem;}
  .kpi .d {font-size:.775rem; font-weight:600; margin-top:.10rem;}
  .kpi .d.up {color:#2E7D5B;} .kpi .d.dn {color:#B3403A;} .kpi .d.nt {color:#8A9AA9;}
  .kpi .s {font-size:.705rem; color:#98A6B3; margin-top:.18rem;}
  .callout {border-radius:8px; padding:.85rem 1rem; margin:.5rem 0 1rem 0; font-size:.90rem; line-height:1.55; border:1px solid; background:#F7FAFC;}
  .callout.good {border-color:#BEDCCB; background:#F2F9F5; color:#1F5A41;}
  .callout.bad  {border-color:#E7C4C2; background:#FDF5F4; color:#8A2F2A;}
  .callout.info {border-color:#C9DCEA; background:#F4F9FD; color:#154B72;}
  .pill {display:inline-block; font-size:.70rem; font-weight:600; padding:.10rem .48rem; border-radius:10px; margin-right:.3rem; color:#FFF;}
  footer, #MainMenu {visibility:hidden;}
</style>
"""

pio.templates["execdash"] = go.layout.Template(
    layout=dict(
        font=dict(family="Inter, Segoe UI, Helvetica, sans-serif", size=12, color=CLR["ink"]),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        colorway=[CLR["navy"], CLR["blue"], CLR["teal"], CLR["amber"], CLR["green"], CLR["red"]],
        margin=dict(l=60, r=25, t=55, b=55),
        title=dict(font=dict(size=15, color=CLR["navy"]), x=0.0, xanchor="left"),
        xaxis=dict(gridcolor="#EEF2F5", zerolinecolor="#D8E1E8", linecolor="#D8E1E8"),
        yaxis=dict(gridcolor="#EEF2F5", zerolinecolor="#D8E1E8", linecolor="#D8E1E8"),
        legend=dict(orientation="h", y=-0.18, x=0, font=dict(size=11), bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(font_size=12, bgcolor="#FFFFFF", bordercolor="#D8E1E8"),
    )
)
pio.templates.default = "plotly_white+execdash"


# ─── FORMATTERS ──────────────────────────────────────────────────────────────

DASH = "—"


def _bad(v) -> bool:
    if v is None:
        return True
    try:
        return not np.isfinite(float(v))
    except (TypeError, ValueError):
        return True


def safe_div(a, b):
    if _bad(a) or _bad(b) or float(b) == 0:
        return np.nan
    return float(a) / float(b)


def fmt_money(v, signed=False):
    if _bad(v): return DASH
    v = float(v)
    s = "-" if v < 0 else ("+" if signed and v > 0 else "")
    a = abs(v)
    if a >= 1e9: return f"{s}${a/1e9:,.1f}B"
    if a >= 1e6: return f"{s}${a/1e6:,.1f}MM"
    if a >= 1e3: return f"{s}${a/1e3:,.1f}K"
    return f"{s}${a:,.0f}"


def fmt_vol(v, signed=False):
    if _bad(v): return DASH
    v = float(v)
    s = "-" if v < 0 else ("+" if signed and v > 0 else "")
    a = abs(v)
    if a >= 1e6: return f"{s}{a/1e6:,.1f} MMboe"
    if a >= 1e3: return f"{s}{a/1e3:,.1f} Mboe"
    return f"{s}{a:,.0f} boe"


def fmt_rate(v, signed=False):
    if _bad(v): return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.0f} boe/d"


def fmt_ratio(v, signed=False):
    if _bad(v): return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.2f}x"


def fmt_pct(v, signed=False):
    if _bad(v): return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.1f}%"


def fmt_years(v, signed=False):
    if _bad(v): return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.2f} yr"


def fmt_usdboe(v, signed=False):
    if _bad(v): return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}${v:,.2f}/boe"


def fmt_int(v, signed=False):
    if _bad(v): return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.0f}"


# ─── FILE LOADING ────────────────────────────────────────────────────────────

def _find(stem: str) -> Path | None:
    for d in SEARCH_DIRS:
        if not d.exists():
            continue
        for p in sorted(d.iterdir()):
            if (p.is_file() and p.suffix.lower() in (".xlsx", ".csv")
                    and p.stem.strip().lower() == stem.lower()
                    and not p.name.startswith("~$")):
                return p
    return None


def _norm(s) -> str:
    """Lowercase, strip punctuation, collapse whitespace — for column matching."""
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def _resolve_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """Find the first column in df whose normalised header matches any alias."""
    lookup = {_norm(c): c for c in df.columns}
    for a in aliases:
        if _norm(a) in lookup:
            return lookup[_norm(a)]
    return None


def _to_num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s.astype("float64")
    t = (s.astype(str).str.replace(r"[,$\s%]", "", regex=True)
         .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
         .replace({"": None, "-": None, "NA": None, "N/A": None, "nan": None, "None": None}))
    return pd.to_numeric(t, errors="coerce")


def _clean(v) -> str:
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return ""
    s = re.sub(r"\s+", " ", str(v)).strip()
    return "" if s.lower() in {"nan", "none", "null", "n/a", "na", "#n/a", "-", "."} else s


def _key(v) -> str:
    return _clean(v).upper()


def _maybe_pct(s: pd.Series) -> pd.Series:
    v = s.dropna()
    if len(v) and v.abs().max() <= 1.5:
        return s * 100.0
    return s


def _read(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=0)


# ─── COLUMN ALIAS MAPS (hardcoded) ──────────────────────────────────────────

ECON_COLS = {
    "entity":  ["Entity", "entity_name", "Entity Name", "Well", "Propnum"],
    "npv":     ["Npv Cash Flow BTax 10.0% (M$)", "NPV Cash Flow BTax 10% (M$)", "NPV BTax 10%", "npv"],
    "invest":  ["Npv Investment BTax 0.0% (M$)", "NPV Investment BTax 0% (M$)", "Investment BTax", "Total Investment"],
    "payout":  ["Payout BTax (years)", "Payout BTax", "Payout"],
    "boe":     ["Boe WI Total (boe)", "BOE WI Total", "Net BOE", "boe"],
    "fy_rate": ["1st Year Production Rate (boepd)", "First Year Production Rate", "1st Yr Rate"],
    "cor":     ["Cost of Reserves ($/boe)", "Cost of Reserves", "COR"],
    "ip30":    ["IP30 Cum (boe)", "IP30 Cum", "IP30"],
    "ror":     ["BTax Disc. CF. ROR (%)", "BTax Disc CF ROR", "ROR", "IRR"],
    "wi":      ["Initial WI (%)", "Initial WI", "WI"],
    "avg3":    ["3 Month Avg Production (boepd)", "3 Mo Avg Production", "3 Month Avg Rate"],
}

CONSOL_COLS = {
    "event_id": ["consolidation #", "consolidation number", "consol #", "event", "event id", "id"],
    "w1_name":  ["well 1 name", "well1 name", "well_1_name"],
    "w2_name":  ["well 2 name", "well2 name", "well_2_name"],
    "w1_ent":   ["well 1 entity", "well1 entity", "well_1_entity"],
    "w2_ent":   ["well 2 entity", "well2 entity", "well_2_entity"],
    "enh_ent":  ["enhanced well entity", "enhanced entity", "enhanced well"],
}

FORECAST_COLS = {
    "entity":  ["entity_name", "Entity", "Entity Name", "Well"],
    "year":    ["year", "yr", "cal_year"],
    "month":   ["month", "mo", "cal_month"],
    "revenue": ["total_revenue", "revenue", "Total Revenue"],
    "opinc":   ["operating_income", "op_income", "Operating Income"],
    "cf":      ["cash_flow", "cashflow", "Cash Flow", "net_cash_flow"],
}

CAPEX_COLS = {
    "entity": ["entity", "Entity", "entity_name", "Well"],
    "capex":  ["capex", "CAPEX", "Capital", "total_capex"],
}


def _load_and_map(stem: str, col_map: dict[str, list[str]]) -> pd.DataFrame:
    """Find file, read it, resolve columns by alias, return standardised df."""
    p = _find(stem)
    if p is None:
        dirs = ", ".join(str(d) for d in SEARCH_DIRS if d.exists())
        raise FileNotFoundError(f"{stem}.xlsx not found. Searched: {dirs}")
    raw = _read(p)
    out = pd.DataFrame(index=raw.index)
    for key, aliases in col_map.items():
        src = _resolve_col(raw, aliases)
        if src is None:
            out[key] = np.nan
        else:
            out[key] = raw[src]
    return out


# ─── LOAD ALL DATA ──────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Loading data…")
def load_all():
    """Returns (econ_df, consol_df, forecast_df, capex_df) with canonical columns."""
    econ = _load_and_map("economics", ECON_COLS)
    consol = _load_and_map("consol", CONSOL_COLS)
    forecast = _load_and_map("forecast", FORECAST_COLS)
    capex = _load_and_map("capex", CAPEX_COLS)

    # --- economics: clean & scale ---
    econ["entity"] = econ["entity"].map(_clean)
    econ["entity_key"] = econ["entity"].map(_key)
    econ = econ[econ["entity_key"] != ""].drop_duplicates("entity_key", keep="first")
    for c in ["npv", "invest", "payout", "boe", "fy_rate", "cor", "ip30", "ror", "wi", "avg3"]:
        econ[c] = _to_num(econ[c])
    econ["npv"] = econ["npv"] * ECON_MONEY_SCALE
    econ["invest"] = econ["invest"] * ECON_MONEY_SCALE
    econ["ror"] = _maybe_pct(econ["ror"])
    econ["wi"] = _maybe_pct(econ["wi"])
    # derived
    econ["npv_inv"] = econ.apply(lambda r: safe_div(r["npv"], r["invest"]), axis=1)
    econ["inv_boe"] = econ.apply(lambda r: safe_div(r["invest"], r["boe"]), axis=1)
    econ["npv_boe"] = econ.apply(lambda r: safe_div(r["npv"], r["boe"]), axis=1)

    # --- capex: clean & aggregate ---
    capex["entity"] = capex["entity"].map(_clean)
    capex["entity_key"] = capex["entity"].map(_key)
    capex["capex"] = _to_num(capex["capex"])
    capex = capex[capex["entity_key"] != ""]
    capex = capex.groupby("entity_key", as_index=False).agg(capex=("capex", "sum"))

    # --- consol: clean & classify ---
    for c in ["w1_ent", "w2_ent", "enh_ent", "w1_name", "w2_name", "event_id"]:
        consol[c] = consol[c].map(_clean)
    # auto-generate event IDs where missing
    consol["event_id"] = [
        v if v else f"ROW-{i+2:04d}" for i, v in enumerate(consol["event_id"])
    ]
    n_src = (consol["w1_ent"] != "").astype(int) + (consol["w2_ent"] != "").astype(int)
    consol["case"] = np.select(
        [n_src == 2, n_src == 1, n_src == 0],
        ["Consolidation", "Extension", "Creation"], default="Unclassified")
    consol.loc[consol["enh_ent"] == "", "case"] = "Unclassified"
    # drop unclassified (no enhanced entity = unusable)
    consol = consol[consol["case"] != "Unclassified"].reset_index(drop=True)

    # --- forecast: clean & parse dates ---
    forecast["entity"] = forecast["entity"].map(_clean)
    forecast["entity_key"] = forecast["entity"].map(_key)
    forecast["year"] = _to_num(forecast["year"])
    forecast["month"] = _to_num(forecast["month"])
    for c in ["revenue", "opinc", "cf"]:
        forecast[c] = _to_num(forecast[c])
    ok = forecast["year"].between(1900, 2200) & forecast["month"].between(1, 12)
    forecast = forecast[ok].copy()
    forecast["date"] = pd.to_datetime(dict(
        year=forecast["year"].astype(int), month=forecast["month"].astype(int), day=1))
    forecast = forecast[forecast["entity_key"] != ""]

    return econ, consol, forecast, capex


# ─── BUILD EVENT ECONOMICS ──────────────────────────────────────────────────

# The additive metrics we sum across wells on a side
ADD_KEYS = ["npv", "invest", "capex", "boe", "fy_rate", "ip30", "avg3"]
# Metrics where we show weighted average or recompute from aggregates
ALL_ECON_KEYS = ADD_KEYS + ["payout", "ror", "cor", "wi", "npv_inv", "inv_boe", "npv_boe"]


def _lookup_entity(econ: pd.DataFrame, capex: pd.DataFrame, entity_key: str) -> dict:
    """Pull a single entity's metrics from econ + capex. Returns dict or all-NaN."""
    out = {k: np.nan for k in ALL_ECON_KEYS}
    er = econ[econ["entity_key"] == entity_key]
    if not er.empty:
        r = er.iloc[0]
        for k in ALL_ECON_KEYS:
            if k != "capex":
                out[k] = r.get(k, np.nan)
    cr = capex[capex["entity_key"] == entity_key]
    if not cr.empty:
        out["capex"] = float(cr.iloc[0]["capex"])
    return out


def _sum_dicts(dicts: list[dict]) -> dict:
    """Sum additive keys across dicts; recompute derived."""
    out = {k: 0.0 for k in ADD_KEYS}
    n = len(dicts)
    # Sum additive
    for d in dicts:
        for k in ADD_KEYS:
            v = d.get(k, np.nan)
            if not _bad(v):
                out[k] = out.get(k, 0.0) + float(v)
    # Weighted averages (invest-weighted for payout/ror, boe-weighted for cor/wi)
    def _wavg(key, weight_key):
        num = sum(d[key] * d[weight_key] for d in dicts
                  if not _bad(d.get(key)) and not _bad(d.get(weight_key)))
        den = sum(d[weight_key] for d in dicts
                  if not _bad(d.get(key)) and not _bad(d.get(weight_key)))
        return safe_div(num, den)
    out["payout"] = _wavg("payout", "invest")
    out["ror"] = _wavg("ror", "invest")
    out["cor"] = _wavg("cor", "boe")
    out["wi"] = _wavg("wi", "boe")
    # Derived
    out["npv_inv"] = safe_div(out["npv"], out["invest"])
    out["inv_boe"] = safe_div(out["invest"], out["boe"])
    out["npv_boe"] = safe_div(out["npv"], out["boe"])
    out["n_wells"] = n
    return out


@st.cache_data(show_spinner="Building event model…")
def build_events():
    econ, consol, forecast, capex = load_all()
    rows = []
    for _, ev in consol.iterrows():
        eid = ev["event_id"]
        case = ev["case"]
        w1k = _key(ev["w1_ent"])
        w2k = _key(ev["w2_ent"])
        enhk = _key(ev["enh_ent"])
        if not enhk:
            continue  # skip if no enhanced entity

        # before side
        before_entities = []
        if w1k:
            before_entities.append(_lookup_entity(econ, capex, w1k))
        if w2k and w2k != w1k:
            before_entities.append(_lookup_entity(econ, capex, w2k))

        if before_entities:
            before = _sum_dicts(before_entities)
        else:
            before = {k: 0.0 for k in ALL_ECON_KEYS}
            before["n_wells"] = 0

        # after side = the single enhanced entity
        after_raw = _lookup_entity(econ, capex, enhk)
        after = dict(after_raw)
        after["n_wells"] = 1

        # build row
        row = {"event_id": eid, "case": case,
               "w1_name": ev["w1_name"], "w2_name": ev["w2_name"],
               "w1_ent": ev["w1_ent"], "w2_ent": ev["w2_ent"], "enh_ent": ev["enh_ent"]}
        row["wells_before"] = before["n_wells"]
        row["wells_after"] = 1
        row["wells_net"] = 1 - before["n_wells"]
        for k in ALL_ECON_KEYS:
            row[f"before_{k}"] = before[k]
            row[f"after_{k}"] = after[k]
            b, a = before[k], after[k]
            row[f"delta_{k}"] = (a - b) if not (_bad(a) or _bad(b)) else np.nan
        rows.append(row)

    ev_df = pd.DataFrame(rows)
    if ev_df.empty:
        return ev_df, consol, econ, forecast, capex

    # computed columns
    ev_df["invest_saved"] = -ev_df["delta_invest"]
    ev_df["capex_saved"] = -ev_df["delta_capex"]
    ev_df["value_created"] = ev_df["delta_npv"]
    return ev_df, consol, econ, forecast, capex


# ─── AGGREGATE BY CASE & PORTFOLIO ──────────────────────────────────────────

def _aggregate_events(ev: pd.DataFrame) -> pd.Series:
    """Aggregate a set of event rows into a single summary series."""
    out = {}
    out["n_events"] = len(ev)
    out["wells_before"] = ev["wells_before"].sum()
    out["wells_after"] = ev["wells_after"].sum()
    out["wells_net"] = ev["wells_net"].sum()
    for side in ("before", "after"):
        for k in ADD_KEYS:
            out[f"{side}_{k}"] = ev[f"{side}_{k}"].sum()
        # weighted averages
        def _wa(metric, weight):
            m = ev[f"{side}_{metric}"]
            w = ev[f"{side}_{weight}"]
            valid = m.notna() & w.notna()
            if valid.sum() == 0:
                return np.nan
            return (m[valid] * w[valid]).sum() / w[valid].sum()
        out[f"{side}_payout"] = _wa("payout", "invest")
        out[f"{side}_ror"] = _wa("ror", "invest")
        out[f"{side}_cor"] = _wa("cor", "boe")
        out[f"{side}_wi"] = _wa("wi", "boe")
        # derived
        out[f"{side}_npv_inv"] = safe_div(out[f"{side}_npv"], out[f"{side}_invest"])
        out[f"{side}_inv_boe"] = safe_div(out[f"{side}_invest"], out[f"{side}_boe"])
        out[f"{side}_npv_boe"] = safe_div(out[f"{side}_npv"], out[f"{side}_boe"])
    for k in ALL_ECON_KEYS:
        b, a = out[f"before_{k}"], out[f"after_{k}"]
        out[f"delta_{k}"] = (a - b) if not (_bad(a) or _bad(b)) else np.nan
    out["invest_saved"] = -out["delta_invest"]
    out["capex_saved"] = -out["delta_capex"]
    out["value_created"] = out["delta_npv"]
    out["cap_productivity_idx"] = safe_div(out["after_npv_inv"], out["before_npv_inv"])
    # per well
    for k in ADD_KEYS:
        out[f"before_{k}_pw"] = safe_div(out[f"before_{k}"], out["wells_before"])
        out[f"after_{k}_pw"] = safe_div(out[f"after_{k}"], out["wells_after"])
        d_pw_b, d_pw_a = out[f"before_{k}_pw"], out[f"after_{k}_pw"]
        out[f"delta_{k}_pw"] = (d_pw_a - d_pw_b) if not (_bad(d_pw_a) or _bad(d_pw_b)) else np.nan
    return pd.Series(out)


# ─── FORECAST PANEL ──────────────────────────────────────────────────────────

def _build_forecast_panel(consol: pd.DataFrame, forecast: pd.DataFrame,
                          event_ids: set | None = None) -> pd.DataFrame:
    """Monthly before/after/incremental panel for a set of events."""
    if forecast.empty or consol.empty:
        return pd.DataFrame()

    rows_before, rows_after = [], []
    for _, ev in consol.iterrows():
        if event_ids is not None and ev["event_id"] not in event_ids:
            continue
        w1k, w2k, enhk = _key(ev["w1_ent"]), _key(ev["w2_ent"]), _key(ev["enh_ent"])
        before_keys = [k for k in [w1k, w2k] if k]
        after_keys = [enhk] if enhk else []
        for k in before_keys:
            f = forecast[forecast["entity_key"] == k]
            if not f.empty:
                rows_before.append(f[["date", "revenue", "opinc", "cf"]])
        for k in after_keys:
            f = forecast[forecast["entity_key"] == k]
            if not f.empty:
                rows_after.append(f[["date", "revenue", "opinc", "cf"]])

    if not rows_before and not rows_after:
        return pd.DataFrame()

    vals = ["revenue", "opinc", "cf"]

    def _agg(parts):
        if not parts:
            return pd.DataFrame()
        df = pd.concat(parts, ignore_index=True)
        return df.groupby("date", as_index=False)[vals].sum()

    b = _agg(rows_before)
    a = _agg(rows_after)

    # merge on date
    if b.empty and a.empty:
        return pd.DataFrame()
    if b.empty:
        panel = a.rename(columns={v: f"after_{v}" for v in vals})
        for v in vals:
            panel[f"before_{v}"] = 0.0
    elif a.empty:
        panel = b.rename(columns={v: f"before_{v}" for v in vals})
        for v in vals:
            panel[f"after_{v}"] = 0.0
    else:
        bm = b.rename(columns={v: f"before_{v}" for v in vals})
        am = a.rename(columns={v: f"after_{v}" for v in vals})
        panel = bm.merge(am, on="date", how="outer").fillna(0.0).sort_values("date")

    for v in vals:
        panel[f"incr_{v}"] = panel[f"after_{v}"] - panel[f"before_{v}"]
    for pre in ("before", "after", "incr"):
        for v in vals:
            panel[f"cum_{pre}_{v}"] = panel[f"{pre}_{v}"].cumsum()
    panel["year"] = panel["date"].dt.year
    panel["month"] = panel["date"].dt.month
    return panel.reset_index(drop=True)


# ─── UI HELPERS ──────────────────────────────────────────────────────────────

def header():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        f'<div class="app-header"><div class="t">{APP_TITLE}</div>'
        f'<div class="s">{APP_SUBTITLE}</div></div>', unsafe_allow_html=True)


def section(n, title, note=""):
    st.markdown(
        f'<div class="sec"><span class="n">{n}</span><span class="h">{title}</span>'
        f'<span class="d">{note}</span></div>', unsafe_allow_html=True)


def callout(kind, html):
    st.markdown(f'<div class="callout {kind}">{html}</div>', unsafe_allow_html=True)


def kpi(label, value, delta=None, dir_=0, sub=None, tone=""):
    dcls = "up" if dir_ > 0 else ("dn" if dir_ < 0 else "nt")
    arrow = "▲ " if dir_ > 0 else ("▼ " if dir_ < 0 else "")
    d = f'<div class="d {dcls}">{arrow}{delta}</div>' if delta else ""
    s = f'<div class="s">{sub}</div>' if sub else ""
    st.markdown(f'<div class="kpi {tone}"><div class="l">{label}</div>'
                f'<div class="v">{value}</div>{d}{s}</div>', unsafe_allow_html=True)


def kpi_grid(cards, ncols=4):
    for i in range(0, len(cards), ncols):
        for col, c in zip(st.columns(ncols), cards[i:i+ncols]):
            with col:
                kpi(**c)


def show_fig(fig, key=None):
    try:
        st.plotly_chart(fig, use_container_width=True,
                        config={"displaylogo": False}, key=key)
    except TypeError:
        st.plotly_chart(fig)


def show_df(df, height=None):
    kw = dict(hide_index=True, use_container_width=True)
    if height:
        kw["height"] = height
    try:
        st.dataframe(df, **kw)
    except TypeError:
        st.dataframe(df, hide_index=True)


def _dir(delta, good=1):
    """Return +1/-1/0 direction for a KPI arrow."""
    if _bad(delta) or abs(delta) < 1e-12:
        return 0
    return int(np.sign(delta) * good)


# ─── METRIC DEFINITIONS (for table display) ─────────────────────────────────

METRIC_DEFS = [
    # (key, label, fmt_fn, good: +1=higher better, -1=lower better)
    ("npv",     "NPV BTax @10%",        fmt_money,  +1),
    ("invest",  "Investment BTax @0%",   fmt_money,  -1),
    ("capex",   "Capex",                 fmt_money,  -1),
    ("boe",     "Net Reserves (BOE)",    fmt_vol,    +1),
    ("fy_rate", "1st-Year Rate",         fmt_rate,   +1),
    ("ip30",    "IP30 Cumulative",       fmt_vol,    +1),
    ("avg3",    "3-Month Avg Rate",      fmt_rate,   +1),
    ("npv_inv", "NPV / Investment",      fmt_ratio,  +1),
    ("inv_boe", "Investment / BOE",      fmt_usdboe, -1),
    ("npv_boe", "NPV / BOE",            fmt_usdboe, +1),
    ("payout",  "Payout BTax",           fmt_years,  -1),
    ("ror",     "BTax ROR",              fmt_pct,    +1),
    ("cor",     "Cost of Reserves",      fmt_usdboe, -1),
    ("wi",      "Initial WI",            fmt_pct,     0),
]


def _metric_card(key, before, after, label=None):
    fmap = {k: f for k, _, f, _ in METRIC_DEFS}
    gmap = {k: g for k, _, _, g in METRIC_DEFS}
    lmap = {k: l for k, l, _, _ in METRIC_DEFS}
    f = fmap.get(key, fmt_money)
    g = gmap.get(key, 0)
    lab = label or lmap.get(key, key)
    d = (after - before) if not (_bad(after) or _bad(before)) else np.nan
    return dict(label=lab, value=f(after),
                delta=f"{f(d, signed=True)} vs before" if not _bad(d) else None,
                dir_=_dir(d, g),
                sub=f"Before {f(before)}",
                tone="pos" if _dir(d, g) > 0 else ("neg" if _dir(d, g) < 0 else ""))


# ─── CHARTS ─────────────────────────────────────────────────────────────────

def _fin(fig, title, h=420, ytitle="", tickpre=""):
    fig.update_layout(title=title, height=h, yaxis_title=ytitle, bargap=0.28, hovermode="x unified")
    if tickpre:
        fig.update_yaxes(tickprefix=tickpre)
    return fig


def chart_bridge(ev):
    wb = float(ev["wells_before"].sum())
    per = {c: float(ev.loc[ev["case"] == c, "wells_net"].sum()) for c in CASE_ORDER}
    labels = ["Before inventory"] + [f"{c}\nnet" for c in CASE_ORDER] + ["After inventory"]
    vals = [wb] + [per[c] for c in CASE_ORDER] + [float(ev["wells_after"].sum())]
    meas = ["absolute"] + ["relative"] * 3 + ["total"]
    fig = go.Figure(go.Waterfall(
        orientation="v", measure=meas, x=labels, y=vals,
        text=[f"{v:+,.0f}" if m == "relative" else f"{v:,.0f}" for v, m in zip(vals, meas)],
        textposition="outside",
        increasing=dict(marker_color=CLR["green"]),
        decreasing=dict(marker_color=CLR["red"]),
        totals=dict(marker_color=CLR["navy"]),
        connector=dict(line=dict(color=CLR["light"], dash="dot")),
    ))
    return _fin(fig, "Inventory Bridge — Well Count", 400, "Wells")


def chart_waterfall_metric(ev, key, title, fmt_fn=fmt_money, topn=12):
    d = ev[["event_id", "case", f"delta_{key}"]].dropna(subset=[f"delta_{key}"]).copy()
    d = d.rename(columns={f"delta_{key}": "d"})
    d = d.reindex(d["d"].abs().sort_values(ascending=False).index)
    head = d.head(topn)

    labels = [f"Before ({fmt_fn(ev[f'before_{key}'].sum())})"]
    vals = [float(ev[f"before_{key}"].sum())]
    meas = ["absolute"]
    for r in head.itertuples():
        labels.append(str(r.event_id))
        vals.append(float(r.d))
        meas.append("relative")
    tail_sum = d.iloc[topn:]["d"].sum() if len(d) > topn else 0
    if abs(tail_sum) > 0:
        labels.append(f"Other ({len(d) - topn})")
        vals.append(float(tail_sum))
        meas.append("relative")
    labels.append("After")
    vals.append(float(ev[f"after_{key}"].sum()))
    meas.append("total")

    fig = go.Figure(go.Waterfall(
        orientation="v", measure=meas, x=labels, y=vals,
        text=[fmt_fn(v, signed=(m == "relative")) for v, m in zip(vals, meas)],
        textposition="outside", textfont=dict(size=10),
        increasing=dict(marker_color=CLR["green"]),
        decreasing=dict(marker_color=CLR["red"]),
        totals=dict(marker_color=CLR["navy"]),
        connector=dict(line=dict(color=CLR["light"], dash="dot")),
    ))
    fig.update_xaxes(tickangle=-40)
    return _fin(fig, title, 470, tickpre="$")


def chart_case_bars(case_data, key, title, fmt_fn=fmt_money):
    cases = list(case_data.keys())
    before_vals = [case_data[c][f"before_{key}"] for c in cases]
    after_vals = [case_data[c][f"after_{key}"] for c in cases]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=cases, y=before_vals, name="Before", marker_color=CLR["before"],
                         text=[fmt_fn(v) for v in before_vals], textposition="outside", textfont_size=10))
    fig.add_trace(go.Bar(x=cases, y=after_vals, name="After", marker_color=CLR["after"],
                         text=[fmt_fn(v) for v in after_vals], textposition="outside", textfont_size=10))
    fig.update_layout(barmode="group")
    return _fin(fig, title, 400)


def chart_monthly(panel, val, title, show=("before", "after")):
    fig = go.Figure()
    style = {"before": (CLR["before"], "dash", "Before"),
             "after": (CLR["after"], "solid", "After"),
             "incr": (CLR["incr"], "solid", "Incremental")}
    for pre in show:
        c, dash, nm = style[pre]
        col = f"{pre}_{val}"
        if col not in panel.columns:
            continue
        fig.add_trace(go.Scatter(
            x=panel["date"], y=panel[col], name=nm, mode="lines",
            line=dict(color=c, width=2.4, dash=dash),
            fill="tozeroy" if pre == "incr" else None,
            fillcolor="rgba(14,154,167,.14)" if pre == "incr" else None))
    fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
    return _fin(fig, title, 400, "$", tickpre="$")


def chart_incremental_bars(panel, val, title):
    y = panel[f"incr_{val}"]
    fig = go.Figure(go.Bar(
        x=panel["date"], y=y,
        marker_color=np.where(y >= 0, CLR["green"], CLR["red"])))
    fig.add_trace(go.Scatter(x=panel["date"], y=panel[f"cum_incr_{val}"],
                             name="Cumulative", yaxis="y2", mode="lines",
                             line=dict(color=CLR["navy"], width=2.4)))
    fig.update_layout(yaxis2=dict(overlaying="y", side="right", showgrid=False,
                                  tickprefix="$", title="Cumulative"))
    return _fin(fig, title, 420, "$ / month", tickpre="$")


def chart_migration(ev):
    d = ev.dropna(subset=["before_invest", "before_npv", "after_invest", "after_npv"])
    fig = go.Figure()
    # connecting lines
    xs, ys = [], []
    for r in d.itertuples():
        xs += [r.before_invest, r.after_invest, None]
        ys += [r.before_npv, r.after_npv, None]
    if xs:
        fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines", showlegend=False,
                                 line=dict(color=CLR["light"], width=1.2), hoverinfo="skip"))
    fig.add_trace(go.Scatter(
        x=d["before_invest"], y=d["before_npv"], mode="markers", name="Before",
        marker=dict(size=9, color=CLR["before"], symbol="circle-open", line_width=2),
        text=d["event_id"],
        hovertemplate="<b>%{text}</b><br>Inv %{x:$,.0f} · NPV %{y:$,.0f}<extra>Before</extra>"))
    fig.add_trace(go.Scatter(
        x=d["after_invest"], y=d["after_npv"], mode="markers", name="After",
        marker=dict(size=11, color=[CASE_CLR.get(c, CLR["slate"]) for c in d["case"]],
                    line=dict(color="#FFF", width=1)),
        text=d["event_id"],
        hovertemplate="<b>%{text}</b><br>Inv %{x:$,.0f} · NPV %{y:$,.0f}<extra>After</extra>"))
    fig.update_xaxes(tickprefix="$", title="Investment")
    return _fin(fig, "Capital Efficiency Migration — Investment vs NPV", 480, "NPV", tickpre="$")


def chart_quadrant(ev):
    d = ev.dropna(subset=["delta_npv", "delta_invest"])
    quads = [
        ("↑NPV ↓Capital", lambda r: r.delta_npv >= 0 and r.delta_invest <= 0, CLR["green"]),
        ("↑NPV ↑Capital", lambda r: r.delta_npv >= 0 and r.delta_invest > 0, CLR["blue"]),
        ("↓NPV ↓Capital", lambda r: r.delta_npv < 0 and r.delta_invest <= 0, CLR["amber"]),
        ("↓NPV ↑Capital", lambda r: r.delta_npv < 0 and r.delta_invest > 0, CLR["red"]),
    ]
    fig = go.Figure()
    for qname, qfilt, qcol in quads:
        mask = d.apply(qfilt, axis=1)
        g = d[mask]
        if g.empty:
            continue
        fig.add_trace(go.Scatter(
            x=g["delta_invest"], y=g["delta_npv"], mode="markers", name=qname,
            marker=dict(size=12, color=qcol, opacity=.85, line=dict(color="#FFF", width=1)),
            text=g["event_id"],
            hovertemplate="<b>%{text}</b><br>ΔInv %{x:$,.0f}<br>ΔNPV %{y:$,.0f}<extra></extra>"))
    fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
    fig.add_vline(x=0, line_color=CLR["slate"], line_width=1)
    fig.update_xaxes(title="Δ Investment", tickprefix="$")
    return _fin(fig, "Capital Quadrant — Value vs Capital Intensity", 480, "Δ NPV", tickpre="$")


# ─── PAGES ───────────────────────────────────────────────────────────────────

def page_summary(ev, port, case_agg, consol, forecast, panel):
    section("01", "Headline Outcome",
            f"{len(ev):,} events · {' / '.join(f'{c[:5]} {int((ev.case==c).sum())}' for c in CASE_ORDER if (ev.case==c).any())}")

    vc = port["delta_npv"]
    inv_saved = port["invest_saved"]
    eff_b, eff_a = port["before_npv_inv"], port["after_npv_inv"]

    verdict = "good" if not _bad(vc) and vc > 0 else ("bad" if not _bad(vc) and vc < 0 else "info")
    bits = [
        f"Inventory moves from <b>{fmt_int(port['wells_before'])}</b> wells to "
        f"<b>{fmt_int(port['wells_after'])}</b> (<b>{fmt_int(port['wells_net'], signed=True)}</b> net)",
        f"portfolio NPV changes by <b>{fmt_money(vc, signed=True)}</b>",
    ]
    if not _bad(inv_saved):
        bits.append(f"{'releases' if inv_saved > 0 else 'consumes'} "
                    f"<b>{fmt_money(abs(inv_saved))}</b> of investment")
    if not _bad(eff_a) and not _bad(eff_b):
        bits.append(f"capital efficiency moves from <b>{fmt_ratio(eff_b)}</b> to "
                    f"<b>{fmt_ratio(eff_a)}</b>")
    callout(verdict, " · ".join(bits) + ".")

    kpi_grid([
        dict(label="Events", value=fmt_int(len(ev)), tone="acc"),
        dict(label="Before Wells", value=fmt_int(port["wells_before"])),
        dict(label="After Wells", value=fmt_int(port["wells_after"])),
        dict(label="Net Wells", value=fmt_int(port["wells_net"], signed=True),
             tone="acc" if (port["wells_net"] or 0) <= 0 else ""),
    ], 4)
    st.write("")
    kpi_grid([
        dict(label="NPV Before", value=fmt_money(port["before_npv"])),
        dict(label="NPV After", value=fmt_money(port["after_npv"]), tone="acc"),
        dict(label="Value Created", value=fmt_money(vc, signed=True),
             dir_=_dir(vc, 1), tone="pos" if _dir(vc, 1) > 0 else "neg"),
        dict(label="NPV/Investment", value=fmt_ratio(eff_a),
             delta=f"{fmt_ratio(eff_a - eff_b, signed=True)} vs before" if not (_bad(eff_a) or _bad(eff_b)) else None,
             dir_=_dir(eff_a - eff_b if not (_bad(eff_a) or _bad(eff_b)) else np.nan, 1),
             sub=f"Before {fmt_ratio(eff_b)}", tone="acc"),
    ], 4)
    st.write("")
    kpi_grid([
        dict(label="Investment Before → After",
             value=f"{fmt_money(port['before_invest'])} → {fmt_money(port['after_invest'])}"),
        dict(label="Investment Saved", value=fmt_money(inv_saved, signed=True),
             dir_=_dir(inv_saved, 1),
             tone="pos" if _dir(inv_saved, 1) > 0 else "neg"),
        dict(label="Reserves After", value=fmt_vol(port["after_boe"]),
             delta=fmt_vol(port["delta_boe"], signed=True),
             dir_=_dir(port["delta_boe"], 1)),
        dict(label="1st-Year Rate After", value=fmt_rate(port["after_fy_rate"]),
             delta=fmt_rate(port["delta_fy_rate"], signed=True),
             dir_=_dir(port["delta_fy_rate"], 1)),
    ], 4)

    section("02", "Inventory & Value Bridges")
    c1, c2 = st.columns([1, 1.35])
    with c1:
        show_fig(chart_bridge(ev), "sum_bridge")
    with c2:
        show_fig(chart_waterfall_metric(ev, "npv", "NPV Waterfall — Before to After"), "sum_npvwf")

    section("03", "Cash-Flow Trajectory", "Portfolio, undiscounted")
    if not panel.empty:
        c1, c2 = st.columns(2)
        with c1:
            show_fig(chart_monthly(panel, "cf", "Monthly Cash Flow — Before vs After"), "sum_mcf")
        with c2:
            fig = go.Figure()
            for pre, col, nm, dash in (("before", CLR["before"], "Before", "dash"),
                                       ("after", CLR["after"], "After", "solid"),
                                       ("incr", CLR["incr"], "Incremental", "solid")):
                fig.add_trace(go.Scatter(x=panel["date"], y=panel[f"cum_{pre}_cf"],
                                         name=nm, mode="lines",
                                         line=dict(color=col, width=2.4, dash=dash)))
            fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
            show_fig(_fin(fig, "Cumulative Cash Flow", 400, "$", tickpre="$"), "sum_ccf")
    else:
        st.info("No forecast data matched to current events.")


def page_economics(ev, port, case_agg, consol, forecast, panel):
    section("01", "Portfolio Economics — Before vs After",
            "Ratios recomputed from aggregates, never summed")

    rows = []
    for key, label, fmt_fn, good in METRIC_DEFS:
        b = port.get(f"before_{key}", np.nan)
        a = port.get(f"after_{key}", np.nan)
        d = (a - b) if not (_bad(a) or _bad(b)) else np.nan
        pct_chg = safe_div(d, abs(b)) * 100 if not (_bad(d) or _bad(b)) else np.nan
        verdict = DASH
        if not _bad(d) and good and abs(d) > 1e-12:
            verdict = "✅ better" if np.sign(d) * good > 0 else "❌ worse"
        rows.append({
            "Metric": label,
            "Before": fmt_fn(b), "After": fmt_fn(a),
            "Delta": fmt_fn(d, signed=True),
            "% Change": fmt_pct(pct_chg, signed=True) if not _bad(pct_chg) else DASH,
            "Direction": verdict,
        })
    show_df(pd.DataFrame(rows), height=540)

    section("02", "Per-Well Normalisation",
            "Comparing 2 before-wells to 1 after-well requires per-well metrics")
    pw = []
    for k in ADD_KEYS:
        fmap = {m[0]: m[2] for m in METRIC_DEFS}
        f = fmap.get(k, fmt_money)
        lmap = {m[0]: m[1] for m in METRIC_DEFS}
        lab = lmap.get(k, k)
        b_pw = port.get(f"before_{k}_pw", np.nan)
        a_pw = port.get(f"after_{k}_pw", np.nan)
        d_pw = (a_pw - b_pw) if not (_bad(a_pw) or _bad(b_pw)) else np.nan
        pw.append({
            "Metric": f"{lab} / well",
            "Before / well": f(b_pw), "After / well": f(a_pw),
            "Delta / well": f(d_pw, signed=True),
        })
    show_df(pd.DataFrame(pw))

    section("03", "Case Contribution")
    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_case_bars(case_agg, "npv", "NPV by Case"), "ec_cnpv")
    with c2:
        show_fig(chart_case_bars(case_agg, "invest", "Investment by Case"), "ec_cinv")

    section("04", "Event Detail")
    disp = pd.DataFrame({
        "Event": ev["event_id"], "Case": ev["case"],
        "Wells": ev["wells_before"].map(lambda x: f"{x:.0f}") + " → " + ev["wells_after"].map(lambda x: f"{x:.0f}"),
        "Before NPV": ev["before_npv"].map(fmt_money),
        "After NPV": ev["after_npv"].map(fmt_money),
        "Δ NPV": ev["delta_npv"].map(lambda x: fmt_money(x, signed=True)),
        "Δ Investment": ev["delta_invest"].map(lambda x: fmt_money(x, signed=True)),
        "NPV/I After": ev["after_npv_inv"].map(fmt_ratio),
    })
    show_df(disp.sort_values("Δ NPV", key=lambda s: pd.to_numeric(s.str.replace(r"[^0-9.\-+]", "", regex=True), errors="coerce"), ascending=False), height=460)


def page_capital(ev, port, case_agg, consol, forecast, panel):
    section("01", "Capital Position")

    kpi_grid([
        dict(label="Investment Before", value=fmt_money(port["before_invest"])),
        dict(label="Investment After", value=fmt_money(port["after_invest"]), tone="acc"),
        dict(label="Investment Saved", value=fmt_money(port["invest_saved"], signed=True),
             dir_=_dir(port["invest_saved"], 1),
             tone="pos" if _dir(port["invest_saved"], 1) > 0 else "neg"),
        dict(label="Capex Saved", value=fmt_money(port["capex_saved"], signed=True),
             dir_=_dir(port["capex_saved"], 1)),
    ], 4)
    st.write("")
    kpi_grid([
        dict(label="NPV/Investment Before", value=fmt_ratio(port["before_npv_inv"])),
        dict(label="NPV/Investment After", value=fmt_ratio(port["after_npv_inv"]), tone="acc"),
        dict(label="Capital Productivity Index", value=fmt_ratio(port["cap_productivity_idx"]),
             sub="After NPV/I ÷ Before NPV/I · >1.00x = improved"),
        dict(label="Investment / BOE", value=fmt_usdboe(port["after_inv_boe"]),
             delta=fmt_usdboe(port["delta_inv_boe"], signed=True),
             dir_=_dir(port["delta_inv_boe"], -1),
             sub=f"Before {fmt_usdboe(port['before_inv_boe'])}"),
    ], 4)

    section("02", "Capital Waterfalls")
    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_waterfall_metric(ev, "invest", "Investment Waterfall"), "cap_iwf")
    with c2:
        show_fig(chart_waterfall_metric(ev, "capex", "Capex Waterfall"), "cap_cwf")

    section("03", "Capital Quadrant")
    show_fig(chart_quadrant(ev), "cap_quad")

    section("04", "Migration — Investment vs NPV")
    show_fig(chart_migration(ev), "cap_mig")


def page_cases(ev, port, case_agg, consol, forecast, panel):
    section("01", "Case Archetypes", "Consolidation · Extension · Creation")
    present = [c for c in CASE_ORDER if c in case_agg]
    if not present:
        st.info("No classified events found.")
        return

    # Summary table
    summary = []
    for c in present:
        r = case_agg[c]
        summary.append({
            "Case": c, "Events": fmt_int(r["n_events"]),
            "Wells Before": fmt_int(r["wells_before"]),
            "Wells After": fmt_int(r["wells_after"]),
            "Net Wells": fmt_int(r["wells_net"], signed=True),
            "NPV Before": fmt_money(r["before_npv"]),
            "NPV After": fmt_money(r["after_npv"]),
            "Value Created": fmt_money(r["delta_npv"], signed=True),
            "Investment Saved": fmt_money(r["invest_saved"], signed=True),
            "NPV/I Before": fmt_ratio(r["before_npv_inv"]),
            "NPV/I After": fmt_ratio(r["after_npv_inv"]),
        })
    show_df(pd.DataFrame(summary))

    tabs = st.tabs(present)
    for tab, case in zip(tabs, present):
        with tab:
            r = case_agg[case]
            sub = ev[ev["case"] == case]
            st.markdown(f'<span class="pill" style="background:{CASE_CLR[case]}">{case}</span>'
                        f'<span style="color:#7A8A99;font-size:.85rem">{len(sub):,} events</span>',
                        unsafe_allow_html=True)

            kpi_grid([
                dict(label="Events", value=fmt_int(r["n_events"]), tone="acc"),
                dict(label="Wells Before → After",
                     value=f"{fmt_int(r['wells_before'])} → {fmt_int(r['wells_after'])}",
                     sub=f"Net {fmt_int(r['wells_net'], signed=True)}"),
                _metric_card("npv", r["before_npv"], r["after_npv"]),
                _metric_card("invest", r["before_invest"], r["after_invest"]),
                _metric_card("npv_inv", r["before_npv_inv"], r["after_npv_inv"]),
                _metric_card("boe", r["before_boe"], r["after_boe"]),
                _metric_card("fy_rate", r["before_fy_rate"], r["after_fy_rate"]),
                _metric_card("inv_boe", r["before_inv_boe"], r["after_inv_boe"]),
            ], 4)

            # Economics table for this case
            st.markdown("**Economics Comparison**")
            rows = []
            for key, label, fmt_fn, good in METRIC_DEFS:
                b, a = r.get(f"before_{key}", np.nan), r.get(f"after_{key}", np.nan)
                d = (a - b) if not (_bad(a) or _bad(b)) else np.nan
                rows.append({"Metric": label, "Before": fmt_fn(b), "After": fmt_fn(a),
                             "Delta": fmt_fn(d, signed=True)})

            c1, c2 = st.columns([1, 1.25])
            with c1:
                show_df(pd.DataFrame(rows), height=430)
            with c2:
                # case-level forecast
                case_events = set(sub["event_id"])
                cpanel = _build_forecast_panel(consol, forecast, case_events)
                if not cpanel.empty:
                    show_fig(chart_monthly(cpanel, "cf", f"{case} — Monthly Cash Flow"), f"cs_cf_{case}")
                else:
                    st.info("No forecast coverage for this case.")

            # Top contributors
            st.markdown("**Top Contributors**")
            disp = pd.DataFrame({
                "Event": sub["event_id"], "Wells": sub["wells_before"].map(lambda x: f"{x:.0f}") + " → 1",
                "Before NPV": sub["before_npv"].map(fmt_money),
                "After NPV": sub["after_npv"].map(fmt_money),
                "Δ NPV": sub["delta_npv"].map(lambda x: fmt_money(x, signed=True)),
                "Δ Investment": sub["delta_invest"].map(lambda x: fmt_money(x, signed=True)),
                "NPV/I After": sub["after_npv_inv"].map(fmt_ratio),
            })
            show_df(disp.sort_values("Δ NPV", key=lambda s: pd.to_numeric(
                s.str.replace(r"[^0-9.\-+]", "", regex=True), errors="coerce"), ascending=False).head(20))


def page_forecast(ev, port, case_agg, consol, forecast, panel):
    section("01", "Forecast Analysis", "Monthly cash flow, revenue, operating income")
    if panel.empty:
        st.warning("No forecast data matched to events.")
        return

    section("02", "Before vs After — Monthly")
    tabs = st.tabs(["Cash Flow", "Revenue", "Operating Income"])
    for tab, val, lab in zip(tabs, ["cf", "revenue", "opinc"],
                             ["Cash Flow", "Revenue", "Operating Income"]):
        with tab:
            c1, c2 = st.columns(2)
            with c1:
                show_fig(chart_monthly(panel, val, f"Monthly {lab}"), f"fc_m_{val}")
            with c2:
                show_fig(chart_incremental_bars(panel, val, f"Incremental {lab} + Cumulative"),
                         f"fc_i_{val}")

    section("03", "Cumulative Cash Flow")
    fig = go.Figure()
    for pre, col, nm, dash in (("before", CLR["before"], "Before", "dash"),
                               ("after", CLR["after"], "After", "solid"),
                               ("incr", CLR["incr"], "Incremental", "solid")):
        fig.add_trace(go.Scatter(x=panel["date"], y=panel[f"cum_{pre}_cf"], name=nm,
                                 mode="lines", line=dict(color=col, width=2.4, dash=dash)))
    fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
    show_fig(_fin(fig, "Cumulative Cash Flow", 420, "$", tickpre="$"), "fc_cum")

    section("04", "Annual Summary")
    ann = panel.groupby("year").agg(
        before_revenue=("before_revenue", "sum"), after_revenue=("after_revenue", "sum"),
        before_cf=("before_cf", "sum"), after_cf=("after_cf", "sum"),
        incr_cf=("incr_cf", "sum")).reset_index()
    ann["cum_incr_cf"] = ann["incr_cf"].cumsum()
    disp = pd.DataFrame({
        "Year": ann["year"].astype(int),
        "Revenue Before": ann["before_revenue"].map(fmt_money),
        "Revenue After": ann["after_revenue"].map(fmt_money),
        "CF Before": ann["before_cf"].map(fmt_money),
        "CF After": ann["after_cf"].map(fmt_money),
        "Incremental CF": ann["incr_cf"].map(lambda x: fmt_money(x, signed=True)),
        "Cumulative Incr.": ann["cum_incr_cf"].map(lambda x: fmt_money(x, signed=True)),
    })
    show_df(disp, height=420)


def page_event_explorer(ev, port, case_agg, consol, forecast, panel):
    section("01", "Event Explorer")
    if ev.empty:
        st.info("No events loaded.")
        return

    lbl_map = {r["event_id"]: f"{r['event_id']}  ·  {r['case']}  ·  {fmt_money(r['delta_npv'], signed=True)}"
               for _, r in ev.iterrows()}
    pick = st.selectbox("Event", list(ev["event_id"]), format_func=lambda k: lbl_map.get(k, k))
    r = ev[ev["event_id"] == pick].iloc[0]

    st.markdown(
        f'<span class="pill" style="background:{CASE_CLR.get(r["case"], CLR["slate"])}">{r["case"]}</span>',
        unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Before — Well 1**")
        st.write(r["w1_name"] or DASH)
        st.caption(f"Entity: `{r['w1_ent'] or DASH}`")
    with c2:
        st.markdown("**Before — Well 2**")
        st.write(r["w2_name"] or DASH)
        st.caption(f"Entity: `{r['w2_ent'] or DASH}`")
    with c3:
        st.markdown("**After — Enhanced Well**")
        st.write(f"`{r['enh_ent']}`")
        st.caption(f"Wells: {r['wells_before']:.0f} → {r['wells_after']:.0f} ({fmt_int(r['wells_net'], signed=True)})")

    kpi_grid([
        _metric_card("npv", r["before_npv"], r["after_npv"]),
        _metric_card("invest", r["before_invest"], r["after_invest"]),
        _metric_card("capex", r["before_capex"], r["after_capex"]),
        _metric_card("npv_inv", r["before_npv_inv"], r["after_npv_inv"]),
        _metric_card("boe", r["before_boe"], r["after_boe"]),
        _metric_card("fy_rate", r["before_fy_rate"], r["after_fy_rate"]),
        _metric_card("payout", r["before_payout"], r["after_payout"]),
        _metric_card("inv_boe", r["before_inv_boe"], r["after_inv_boe"]),
    ], 4)

    # Full metrics table
    st.markdown("**Full Economics**")
    rows = []
    for key, label, fmt_fn, good in METRIC_DEFS:
        b, a = r.get(f"before_{key}", np.nan), r.get(f"after_{key}", np.nan)
        d = (a - b) if not (_bad(a) or _bad(b)) else np.nan
        rows.append({"Metric": label, "Before": fmt_fn(b), "After": fmt_fn(a),
                     "Delta": fmt_fn(d, signed=True)})
    show_df(pd.DataFrame(rows), height=430)

    # Event-level forecast
    section("02", "Event Cash Flow")
    ep = _build_forecast_panel(consol, forecast, {pick})
    if ep.empty:
        st.info("No forecast rows matched to this event.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            show_fig(chart_monthly(ep, "cf", "Monthly Cash Flow"), "ee_cf")
        with c2:
            show_fig(chart_incremental_bars(ep, "cf", "Incremental Cash Flow"), "ee_icf")


# ─── PAGE REGISTRY ──────────────────────────────────────────────────────────

PAGES = {
    "Executive Summary": page_summary,
    "Portfolio Economics": page_economics,
    "Capital Efficiency": page_capital,
    "Case Analysis": page_cases,
    "Forecast Analysis": page_forecast,
    "Event Explorer": page_event_explorer,
}


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    header()

    # load data
    try:
        ev_df, consol, econ, forecast, capex = build_events()
    except FileNotFoundError as e:
        st.error(str(e))
        st.info("Place `economics.xlsx`, `consol.xlsx`, `forecast.xlsx`, and `capex.xlsx` "
                "next to `app.py` or in a `./data` subfolder.")
        return

    if ev_df.empty:
        st.warning("No classifiable events found in consol.xlsx. Check that `well 1 entity`, "
                   "`well 2 entity`, and `enhanced well entity` columns are populated.")
        return

    # sidebar
    sb = st.sidebar
    sb.markdown("### ⛽ Navigation")
    page = sb.radio("Page", list(PAGES), label_visibility="collapsed")
    sb.markdown("---")
    sb.markdown("### Scope")
    cases = sb.multiselect("Event Cases", CASE_ORDER, default=CASE_ORDER)
    sb.markdown("---")
    if sb.button("🔄 Reload Data"):
        st.cache_data.clear()
        st.rerun()

    # filter
    ev = ev_df[ev_df["case"].isin(cases)].copy()
    if ev.empty:
        st.warning("No events match the selected cases.")
        return

    # portfolio aggregate
    port = _aggregate_events(ev)

    # case aggregates (dict of case -> series)
    case_agg = {}
    for c in CASE_ORDER:
        sub = ev[ev["case"] == c]
        if not sub.empty:
            case_agg[c] = _aggregate_events(sub)

    # forecast panel
    panel = _build_forecast_panel(consol, forecast, set(ev["event_id"]))

    # render page
    PAGES[page](ev, port, case_agg, consol, forecast, panel)

    # footer
    st.markdown("---")
    st.caption(f"{len(ev):,} events · money in base dollars (economics ×{ECON_MONEY_SCALE:g})")


if __name__ == "__main__":
    main()