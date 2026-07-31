"""
=============================================================================
 INVENTORY REDESIGN  —  PROJECT ECONOMICS DASHBOARD
=============================================================================

 Compares BEFORE inventory (legacy wells) with AFTER inventory (enhanced wells)
 across three event types:

     Consolidation : 2 legacy wells (w1 + w2) -> 1 enhanced well  (net -1 well)
     Extension     : 1 legacy well  (w1)      -> 1 enhanced well  (net  0 wells)
     Creation      : 0 legacy wells           -> 1 enhanced well  (net +1 well)

 "Enhanced entity" = what the combined/extended/created well becomes.
 w1_ent, w2_ent, and enh_ent values all exist as Entity values in economics.xlsx,
 capex.xlsx, and forecast.xlsx.

 ─────────────────────────────────────────────────────────────────────────────
 EXPECTED XLSX FORMATS (place next to app.py or in ./data):
 ─────────────────────────────────────────────────────────────────────────────

 economics.xlsx  (one row per entity)
   Entity                              text — the type-curve / entity name
   Npv Cash Flow BTax 10.0% (M$)      NPV in thousands of dollars
   Npv Investment BTax  0.0% (M$)     investment in thousands of dollars
   Payout BTax (years)                payout in years
   Boe WI Total (boe)                net reserves BOE
   1st Year Production Rate (boepd)   first-year rate
   Cost of Reserves ($/boe)           $/boe
   IP30 Cum (boe)                     IP30 cum production
   BTax Disc. CF. ROR (%)             ROR percent
   Initial WI (%)                     working interest percent
   3 Month Avg Production (boepd)     3-month avg rate

 consol.xlsx  (one row per redesign event)
   consolidation #                     event number
   well 1 name                         UWI / display name for well 1
   well 2 name                         UWI / display name for well 2
   well 1 entity                       type-curve entity for before-well 1
   well 2 entity                       type-curve entity for before-well 2
   enhanced well entity                type-curve entity for the enhanced well

 forecast.xlsx  (one row per entity per month)
   entity_name                         type-curve entity name
   year                                calendar year
   month                               calendar month 1-12
   total_revenue                       monthly revenue $
   operating_income                    monthly operating income $
   cash_flow                           monthly net cash flow $

 capex.xlsx  (one row per entity)
   entity                              type-curve entity name
   capex                               total capex $

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

APP_TITLE = "Inventory Redesign — Project Economics"
APP_SUBTITLE = "Before vs After · 3 Cases · Combined Summary"
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
  .callout {border-radius:8px; padding:.85rem 1rem; margin:.5rem 0 1rem 0; font-size:.90rem; line-height:1.55; border:1px solid;}
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
    if _bad(v):
        return DASH
    v = float(v)
    s = "-" if v < 0 else ("+" if signed and v > 0 else "")
    a = abs(v)
    if a >= 1e9:
        return f"{s}${a / 1e9:,.1f}B"
    if a >= 1e6:
        return f"{s}${a / 1e6:,.1f}MM"
    if a >= 1e3:
        return f"{s}${a / 1e3:,.1f}K"
    return f"{s}${a:,.0f}"


def fmt_vol(v, signed=False):
    if _bad(v):
        return DASH
    v = float(v)
    s = "-" if v < 0 else ("+" if signed and v > 0 else "")
    a = abs(v)
    if a >= 1e6:
        return f"{s}{a / 1e6:,.1f} MMboe"
    if a >= 1e3:
        return f"{s}{a / 1e3:,.1f} Mboe"
    return f"{s}{a:,.0f} boe"


def fmt_rate(v, signed=False):
    if _bad(v):
        return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.0f} boe/d"


def fmt_ratio(v, signed=False):
    if _bad(v):
        return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.2f}x"


def fmt_pct(v, signed=False):
    if _bad(v):
        return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.1f}%"


def fmt_years(v, signed=False):
    if _bad(v):
        return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.2f} yr"


def fmt_usdboe(v, signed=False):
    if _bad(v):
        return DASH
    s = "+" if signed and v > 0 else ""
    return f"{s}${v:,.2f}/boe"


def fmt_int(v, signed=False):
    if _bad(v):
        return DASH
    v = float(v)
    s = "+" if signed and v > 0 else ""
    return f"{s}{v:,.0f}"


# ─── METRIC DEFINITIONS ─────────────────────────────────────────────────────

METRIC_DEFS = [
    # (key, label, formatter, good: +1=higher better, -1=lower better, 0=neutral)
    ("npv", "NPV BTax @10%", fmt_money, +1),
    ("invest", "Investment BTax @0%", fmt_money, -1),
    ("capex", "Capex", fmt_money, -1),
    ("boe", "Net Reserves (BOE)", fmt_vol, +1),
    ("fy_rate", "1st-Year Rate", fmt_rate, +1),
    ("ip30", "IP30 Cumulative", fmt_vol, +1),
    ("avg3", "3-Month Avg Rate", fmt_rate, +1),
    ("npv_inv", "NPV / Investment", fmt_ratio, +1),
    ("inv_boe", "Investment / BOE", fmt_usdboe, -1),
    ("npv_boe", "NPV / BOE", fmt_usdboe, +1),
    ("payout", "Payout BTax", fmt_years, -1),
    ("ror", "BTax ROR", fmt_pct, +1),
    ("cor", "Cost of Reserves", fmt_usdboe, -1),
    ("wi", "Initial WI", fmt_pct, 0),
]

ADD_KEYS = ["npv", "invest", "capex", "boe", "fy_rate", "ip30", "avg3"]
ALL_ECON_KEYS = [m[0] for m in METRIC_DEFS]


# ─── FILE LOADING (hardcoded columns) ───────────────────────────────────────

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


def _read(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=0)


def _strip(v) -> str:
    """Convert any cell to a stripped string. NaN/None -> empty."""
    if v is None:
        return ""
    if isinstance(v, float) and not np.isfinite(v):
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "none", "null", "n/a", "na", "#n/a"):
        return ""
    return s


def _num(v) -> float:
    """Convert a single cell to float. Bad values -> NaN."""
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v) if np.isfinite(float(v)) else np.nan
    s = str(v).replace(",", "").replace("$", "").replace("%", "").strip()
    if not s or s.lower() in ("nan", "none", "null", "n/a", "na", "#n/a", "-", "."):
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


def _maybe_pct(vals: list[float]) -> list[float]:
    """If all non-NaN values are <= 1.5, assume 0-1 fraction and multiply by 100."""
    finite = [v for v in vals if not _bad(v)]
    if finite and max(abs(v) for v in finite) <= 1.5:
        return [v * 100 if not _bad(v) else np.nan for v in vals]
    return vals


def _col(df: pd.DataFrame, name: str) -> str | None:
    """Find column by exact match on stripped/lowered name."""
    target = name.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == target:
            return c
    return None


@st.cache_data(show_spinner="Loading data…")
def load_all():
    """Load all 4 workbooks with hardcoded column names. Returns dicts keyed by entity."""

    # ── ECONOMICS ────────────────────────────────────────────────────────
    p = _find("economics")
    if p is None:
        raise FileNotFoundError("economics.xlsx not found in " + ", ".join(str(d) for d in SEARCH_DIRS))
    raw = _read(p)

    # Hardcoded column names from your file
    C_ENT = _col(raw, "Entity")
    C_NPV = _col(raw, "Npv Cash Flow BTax 10.0% (M$)")
    C_INV = _col(raw, "Npv Investment BTax  0.0% (M$)")  # note: two spaces before 0.0%
    C_PAY = _col(raw, "Payout BTax (years)")
    C_BOE = _col(raw, "Boe WI Total (boe)")
    C_FYR = _col(raw, "1st Year Production Rate (boepd)")
    C_COR = _col(raw, "Cost of Reserves ($/boe)")
    C_IP3 = _col(raw, "IP30 Cum (boe)")
    C_ROR = _col(raw, "BTax Disc. CF. ROR (%)")
    C_WI = _col(raw, "Initial WI (%)")
    C_AV3 = _col(raw, "3 Month Avg Production (boepd)")

    # If the two-space version didn't match, try single space
    if C_INV is None:
        C_INV = _col(raw, "Npv Investment BTax 0.0% (M$)")

    # Debug: show what we found
    found = {
        "Entity": C_ENT, "NPV": C_NPV, "Invest": C_INV, "Payout": C_PAY,
        "BOE": C_BOE, "FY Rate": C_FYR, "COR": C_COR, "IP30": C_IP3,
        "ROR": C_ROR, "WI": C_WI, "Avg3": C_AV3,
    }
    econ_debug = {
        "columns_in_file": list(raw.columns),
        "columns_matched": {k: v for k, v in found.items() if v is not None},
        "columns_missing": [k for k, v in found.items() if v is None],
        "row_count": len(raw),
    }

    if C_ENT is None:
        raise ValueError(f"economics.xlsx: no 'Entity' column found. Columns are: {list(raw.columns)}")

    # Build entity dict: entity_string -> metrics dict
    econ = {}
    ror_vals, wi_vals, ror_keys, wi_keys = [], [], [], []

    for _, row in raw.iterrows():
        ent = _strip(row[C_ENT])
        if not ent:
            continue

        d = {
            "npv":     _num(row[C_NPV]) * ECON_MONEY_SCALE if C_NPV else np.nan,
            "invest":  _num(row[C_INV]) * ECON_MONEY_SCALE if C_INV else np.nan,
            "payout":  _num(row[C_PAY]) if C_PAY else np.nan,
            "boe":     _num(row[C_BOE]) if C_BOE else np.nan,
            "fy_rate": _num(row[C_FYR]) if C_FYR else np.nan,
            "cor":     _num(row[C_COR]) if C_COR else np.nan,
            "ip30":    _num(row[C_IP3]) if C_IP3 else np.nan,
            "ror":     _num(row[C_ROR]) if C_ROR else np.nan,
            "wi":      _num(row[C_WI]) if C_WI else np.nan,
            "avg3":    _num(row[C_AV3]) if C_AV3 else np.nan,
            "capex":   np.nan,  # filled from capex.xlsx
        }
        d["npv_inv"] = safe_div(d["npv"], d["invest"])
        d["inv_boe"] = safe_div(d["invest"], d["boe"])
        d["npv_boe"] = safe_div(d["npv"], d["boe"])
        econ[ent] = d

        ror_vals.append(d["ror"])
        ror_keys.append(ent)
        wi_vals.append(d["wi"])
        wi_keys.append(ent)

    # Fix ROR/WI if stored as fractions
    ror_vals = _maybe_pct(ror_vals)
    wi_vals = _maybe_pct(wi_vals)
    for i, ent in enumerate(ror_keys):
        econ[ent]["ror"] = ror_vals[i]
    for i, ent in enumerate(wi_keys):
        econ[ent]["wi"] = wi_vals[i]

    # ── CAPEX ────────────────────────────────────────────────────────────
    p = _find("capex")
    if p is None:
        raise FileNotFoundError("capex.xlsx not found")
    raw = _read(p)

    C_CENT = None
    C_CCAP = None
    for c in raw.columns:
        cl = str(c).strip().lower()
        if cl in ("entity", "entity_name", "well"):
            C_CENT = c
        if cl in ("capex", "capital", "total_capex"):
            C_CCAP = c

    if C_CENT and C_CCAP:
        for _, row in raw.iterrows():
            ent = _strip(row[C_CENT])
            if not ent:
                continue
            val = _num(row[C_CCAP])
            if ent in econ:
                existing = econ[ent].get("capex", np.nan)
                if _bad(existing):
                    econ[ent]["capex"] = val
                else:
                    econ[ent]["capex"] = existing + (val if not _bad(val) else 0)
            # Also store for entities that might only be in capex
            if ent not in econ:
                econ[ent] = {k: np.nan for k in ALL_ECON_KEYS}
                econ[ent]["capex"] = val

    # ── CONSOL ───────────────────────────────────────────────────────────
    p = _find("consol")
    if p is None:
        raise FileNotFoundError("consol.xlsx not found")
    raw = _read(p)

    # Hardcoded column names
    C_CID = _col(raw, "consolidation #")
    C_W1N = _col(raw, "well 1 name")
    C_W2N = _col(raw, "well 2 name")
    C_W1E = _col(raw, "well 1 entity")
    C_W2E = _col(raw, "well 2 entity")
    C_ENH = _col(raw, "enhanced well entity")

    if C_W1E is None:
        raise ValueError(f"consol.xlsx: no 'well 1 entity' column. Columns: {list(raw.columns)}")
    if C_ENH is None:
        raise ValueError(f"consol.xlsx: no 'enhanced well entity' column. Columns: {list(raw.columns)}")

    consol_rows = []
    for i, row in raw.iterrows():
        eid = _strip(row[C_CID]) if C_CID else ""
        if not eid:
            eid = f"EVT-{i + 1:04d}"
        w1_name = _strip(row[C_W1N]) if C_W1N else ""
        w2_name = _strip(row[C_W2N]) if C_W2N else ""
        w1_ent = _strip(row[C_W1E])
        w2_ent = _strip(row[C_W2E]) if C_W2E else ""
        enh_ent = _strip(row[C_ENH])

        if not enh_ent:
            continue  # skip rows with no enhanced entity

        n_src = (1 if w1_ent else 0) + (1 if w2_ent else 0)
        if n_src == 2:
            case = "Consolidation"
        elif n_src == 1:
            case = "Extension"
        else:
            case = "Creation"

        consol_rows.append({
            "event_id": eid, "case": case,
            "w1_name": w1_name, "w2_name": w2_name,
            "w1_ent": w1_ent, "w2_ent": w2_ent, "enh_ent": enh_ent,
        })

    consol = consol_rows  # list of dicts

    # ── FORECAST ─────────────────────────────────────────────────────────
    p = _find("forecast")
    if p is None:
        raise FileNotFoundError("forecast.xlsx not found")
    raw = _read(p)

    C_FENT = None
    C_FYR2 = None
    C_FMO = None
    C_FREV = None
    C_FOPI = None
    C_FCF = None
    for c in raw.columns:
        cl = str(c).strip().lower()
        if cl in ("entity_name", "entity", "well"):
            C_FENT = c
        elif cl in ("year", "yr", "cal_year"):
            C_FYR2 = c
        elif cl in ("month", "mo", "cal_month"):
            C_FMO = c
        elif cl in ("total_revenue", "revenue"):
            C_FREV = c
        elif cl in ("operating_income", "op_income"):
            C_FOPI = c
        elif cl in ("cash_flow", "cashflow", "net_cash_flow"):
            C_FCF = c

    fc_rows = []
    if C_FENT and C_FYR2 and C_FMO:
        for _, row in raw.iterrows():
            ent = _strip(row[C_FENT])
            if not ent:
                continue
            yr = _num(row[C_FYR2])
            mo = _num(row[C_FMO])
            if _bad(yr) or _bad(mo):
                continue
            yr, mo = int(yr), int(mo)
            if not (1900 <= yr <= 2200 and 1 <= mo <= 12):
                continue
            fc_rows.append({
                "entity": ent,
                "date": pd.Timestamp(yr, mo, 1),
                "revenue": _num(row[C_FREV]) if C_FREV else 0.0,
                "opinc": _num(row[C_FOPI]) if C_FOPI else 0.0,
                "cf": _num(row[C_FCF]) if C_FCF else 0.0,
            })

    fc = pd.DataFrame(fc_rows) if fc_rows else pd.DataFrame(
        columns=["entity", "date", "revenue", "opinc", "cf"])
    for c in ["revenue", "opinc", "cf"]:
        fc[c] = fc[c].fillna(0.0)

    debug = {
        "econ_entities": len(econ),
        "econ_columns": econ_debug,
        "consol_events": len(consol),
        "fc_rows": len(fc),
        "sample_econ_keys": list(econ.keys())[:10],
        "sample_consol_enh": [r["enh_ent"] for r in consol[:5]],
        "sample_consol_w1": [r["w1_ent"] for r in consol[:5]],
    }

    return econ, consol, fc, debug


# ─── BUILD EVENT ECONOMICS ──────────────────────────────────────────────────

def _get(econ: dict, entity: str) -> dict:
    """Look up an entity in the econ dict. Returns metrics or all-NaN."""
    if not entity:
        return {k: np.nan for k in ALL_ECON_KEYS}
    # Exact match first
    if entity in econ:
        return dict(econ[entity])
    # Try case-insensitive
    entity_upper = entity.upper()
    for k, v in econ.items():
        if k.upper() == entity_upper:
            return dict(v)
    return {k: np.nan for k in ALL_ECON_KEYS}


def _sum_sides(dicts: list[dict]) -> dict:
    """Aggregate multiple entity dicts into one side."""
    if not dicts:
        out = {k: 0.0 for k in ALL_ECON_KEYS}
        out["n_wells"] = 0
        return out

    out = {}
    for k in ADD_KEYS:
        vals = [d[k] for d in dicts if not _bad(d[k])]
        out[k] = sum(vals) if vals else np.nan

    def _wavg(metric, wt):
        pairs = [(d[metric], d[wt]) for d in dicts
                 if not _bad(d.get(metric)) and not _bad(d.get(wt))]
        if not pairs:
            return np.nan
        num = sum(m * w for m, w in pairs)
        den = sum(w for _, w in pairs)
        return safe_div(num, den)

    out["payout"] = _wavg("payout", "invest")
    out["ror"] = _wavg("ror", "invest")
    out["cor"] = _wavg("cor", "boe")
    out["wi"] = _wavg("wi", "boe")
    out["npv_inv"] = safe_div(out["npv"], out["invest"])
    out["inv_boe"] = safe_div(out["invest"], out["boe"])
    out["npv_boe"] = safe_div(out["npv"], out["boe"])
    out["n_wells"] = len(dicts)
    return out


@st.cache_data(show_spinner="Building event model…")
def build_events():
    econ, consol, fc, debug = load_all()

    rows = []
    for ev in consol:
        w1_ent = ev["w1_ent"]
        w2_ent = ev["w2_ent"]
        enh_ent = ev["enh_ent"]

        # Before side
        before_list = []
        if w1_ent:
            before_list.append(_get(econ, w1_ent))
        if w2_ent and w2_ent != w1_ent:
            before_list.append(_get(econ, w2_ent))

        before = _sum_sides(before_list)

        # After side = the single enhanced entity
        after = _get(econ, enh_ent)
        after["n_wells"] = 1

        # Check what matched
        enh_found = enh_ent in econ or any(k.upper() == enh_ent.upper() for k in econ)
        w1_found = (not w1_ent) or w1_ent in econ or any(k.upper() == w1_ent.upper() for k in econ)
        w2_found = (not w2_ent) or w2_ent in econ or any(k.upper() == w2_ent.upper() for k in econ)

        row = dict(ev)  # event_id, case, w1_name, w2_name, w1_ent, w2_ent, enh_ent
        row["wells_before"] = before["n_wells"]
        row["wells_after"] = 1
        row["wells_net"] = 1 - before["n_wells"]
        row["enh_found"] = enh_found
        row["w1_found"] = w1_found
        row["w2_found"] = w2_found

        for k in ALL_ECON_KEYS:
            bv = before.get(k, np.nan)
            av = after.get(k, np.nan)
            row[f"before_{k}"] = bv
            row[f"after_{k}"] = av
            row[f"delta_{k}"] = (av - bv) if not (_bad(av) or _bad(bv)) else np.nan

        rows.append(row)

    ev_df = pd.DataFrame(rows) if rows else pd.DataFrame()
    if not ev_df.empty:
        ev_df["invest_saved"] = -ev_df["delta_invest"]
        ev_df["capex_saved"] = -ev_df["delta_capex"]

    return ev_df, consol, fc, econ, debug


# ─── AGGREGATE ───────────────────────────────────────────────────────────────

def _aggregate(ev: pd.DataFrame) -> dict:
    """Aggregate event rows into a summary dict."""
    out = {
        "n_events": len(ev),
        "wells_before": ev["wells_before"].sum(),
        "wells_after": ev["wells_after"].sum(),
        "wells_net": ev["wells_net"].sum(),
    }
    for side in ("before", "after"):
        for k in ADD_KEYS:
            out[f"{side}_{k}"] = ev[f"{side}_{k}"].sum(min_count=1)

        def _wa(metric, weight):
            m = ev[f"{side}_{metric}"]
            w = ev[f"{side}_{weight}"]
            valid = m.notna() & w.notna() & (w != 0)
            if valid.sum() == 0:
                return np.nan
            return float((m[valid] * w[valid]).sum() / w[valid].sum())

        out[f"{side}_payout"] = _wa("payout", "invest")
        out[f"{side}_ror"] = _wa("ror", "invest")
        out[f"{side}_cor"] = _wa("cor", "boe")
        out[f"{side}_wi"] = _wa("wi", "boe")
        out[f"{side}_npv_inv"] = safe_div(out[f"{side}_npv"], out[f"{side}_invest"])
        out[f"{side}_inv_boe"] = safe_div(out[f"{side}_invest"], out[f"{side}_boe"])
        out[f"{side}_npv_boe"] = safe_div(out[f"{side}_npv"], out[f"{side}_boe"])

    for k in ALL_ECON_KEYS:
        b, a = out.get(f"before_{k}", np.nan), out.get(f"after_{k}", np.nan)
        out[f"delta_{k}"] = (a - b) if not (_bad(a) or _bad(b)) else np.nan

    out["invest_saved"] = -out["delta_invest"] if not _bad(out.get("delta_invest")) else np.nan
    out["capex_saved"] = -out["delta_capex"] if not _bad(out.get("delta_capex")) else np.nan
    out["cap_productivity_idx"] = safe_div(out.get("after_npv_inv"), out.get("before_npv_inv"))

    for k in ADD_KEYS:
        out[f"before_{k}_pw"] = safe_div(out.get(f"before_{k}"), out["wells_before"])
        out[f"after_{k}_pw"] = safe_div(out.get(f"after_{k}"), out["wells_after"])
        bpw, apw = out[f"before_{k}_pw"], out[f"after_{k}_pw"]
        out[f"delta_{k}_pw"] = (apw - bpw) if not (_bad(apw) or _bad(bpw)) else np.nan

    return out


# ─── FORECAST PANEL ──────────────────────────────────────────────────────────

def _build_panel(consol_rows, fc: pd.DataFrame,
                 event_ids: set | None = None) -> pd.DataFrame:
    if fc.empty:
        return pd.DataFrame()

    vals = ["revenue", "opinc", "cf"]
    parts_before, parts_after = [], []

    for ev in consol_rows:
        if event_ids is not None and ev["event_id"] not in event_ids:
            continue
        for bk in [ev["w1_ent"], ev["w2_ent"]]:
            if bk:
                f = fc[fc["entity"] == bk]
                if f.empty:
                    f = fc[fc["entity"].str.upper() == bk.upper()]
                if not f.empty:
                    parts_before.append(f[["date"] + vals])
        ek = ev["enh_ent"]
        if ek:
            f = fc[fc["entity"] == ek]
            if f.empty:
                f = fc[fc["entity"].str.upper() == ek.upper()]
            if not f.empty:
                parts_after.append(f[["date"] + vals])

    def _agg(parts):
        if not parts:
            return pd.DataFrame()
        return pd.concat(parts, ignore_index=True).groupby("date", as_index=False)[vals].sum()

    b, a = _agg(parts_before), _agg(parts_after)
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
    return panel.reset_index(drop=True)


def _entity_forecast(fc: pd.DataFrame, entity: str) -> pd.DataFrame:
    """Get monthly forecast for a single entity."""
    f = fc[fc["entity"] == entity]
    if f.empty:
        f = fc[fc["entity"].str.upper() == entity.upper()]
    return f.sort_values("date").copy() if not f.empty else pd.DataFrame()


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
        for col, c in zip(st.columns(ncols), cards[i:i + ncols]):
            with col:
                kpi(**c)


def show_fig(fig, key=None):
    st.plotly_chart(fig, use_container_width=True,
                    config={"displaylogo": False}, key=key)


def show_df(df, height=None):
    kw = dict(hide_index=True, use_container_width=True)
    if height:
        kw["height"] = height
    st.dataframe(df, **kw)


def _dir(delta, good=1):
    if _bad(delta) or abs(float(delta)) < 1e-12:
        return 0
    return int(np.sign(float(delta)) * good)


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


# ─── CHARTS ──────────────────────────────────────────────────────────────────

def _fin(fig, title, h=420, ytitle="", tickpre=""):
    fig.update_layout(title=title, height=h, yaxis_title=ytitle,
                      bargap=0.28, hovermode="x unified")
    if tickpre:
        fig.update_yaxes(tickprefix=tickpre)
    return fig


def chart_case_bars(case_agg, key, title):
    fmap = {k: f for k, _, f, _ in METRIC_DEFS}
    fmt_fn = fmap.get(key, fmt_money)
    cases = [c for c in CASE_ORDER if c in case_agg]
    bv = [case_agg[c][f"before_{key}"] for c in cases]
    av = [case_agg[c][f"after_{key}"] for c in cases]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=cases, y=bv, name="Before", marker_color=CLR["before"],
                         text=[fmt_fn(v) for v in bv], textposition="outside", textfont_size=10))
    fig.add_trace(go.Bar(x=cases, y=av, name="After", marker_color=CLR["after"],
                         text=[fmt_fn(v) for v in av], textposition="outside", textfont_size=10))
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


def chart_entity_cf(fc: pd.DataFrame, entities: dict[str, str], title: str) -> go.Figure:
    """One line per entity on the same chart. entities = {entity_name: display_label}."""
    palette = [CLR["navy"], CLR["blue"], CLR["teal"], CLR["amber"], CLR["green"],
               CLR["red"], CLR["slate"]]
    fig = go.Figure()
    for i, (ent, label) in enumerate(entities.items()):
        ef = fc[fc["entity"] == ent]
        if ef.empty:
            ef = fc[fc["entity"].str.upper() == ent.upper()]
        if ef.empty:
            continue
        ef = ef.sort_values("date")
        fig.add_trace(go.Scatter(
            x=ef["date"], y=ef["cf"], name=label, mode="lines",
            line=dict(color=palette[i % len(palette)], width=2.2)))
    fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
    return _fin(fig, title, 420, "$", tickpre="$")


def chart_migration(ev):
    d = ev.dropna(subset=["before_invest", "before_npv", "after_invest", "after_npv"])
    fig = go.Figure()
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
    return _fin(fig, "Investment vs NPV Migration", 480, "NPV", tickpre="$")


# ─── PAGES ───────────────────────────────────────────────────────────────────

def page_summary(ev, port, case_agg, consol, fc, panel):
    section("01", "Headline",
            f"{len(ev):,} events · "
            + " / ".join(f"{c[:5]} {int((ev.case == c).sum())}"
                         for c in CASE_ORDER if (ev.case == c).any()))

    vc = port["delta_npv"]
    inv_saved = port["invest_saved"]
    eff_b, eff_a = port["before_npv_inv"], port["after_npv_inv"]

    verdict = "good" if not _bad(vc) and vc > 0 else (
        "bad" if not _bad(vc) and vc < 0 else "info")
    bits = [
        f"Inventory moves from <b>{fmt_int(port['wells_before'])}</b> wells to "
        f"<b>{fmt_int(port['wells_after'])}</b> "
        f"(<b>{fmt_int(port['wells_net'], signed=True)}</b> net)",
        f"portfolio NPV changes by <b>{fmt_money(vc, signed=True)}</b>",
    ]
    if not _bad(inv_saved):
        bits.append(f"{'releases' if inv_saved > 0 else 'consumes'} "
                    f"<b>{fmt_money(abs(inv_saved))}</b> of investment")
    if not _bad(eff_a) and not _bad(eff_b):
        bits.append(f"capital efficiency: <b>{fmt_ratio(eff_b)}</b> → <b>{fmt_ratio(eff_a)}</b>")
    callout(verdict, " · ".join(bits) + ".")

    kpi_grid([
        dict(label="Events", value=fmt_int(len(ev)), tone="acc"),
        dict(label="Before Wells", value=fmt_int(port["wells_before"])),
        dict(label="After Wells", value=fmt_int(port["wells_after"])),
        dict(label="Net Wells", value=fmt_int(port["wells_net"], signed=True)),
    ], 4)
    st.write("")
    kpi_grid([
        dict(label="NPV Before", value=fmt_money(port["before_npv"])),
        dict(label="NPV After", value=fmt_money(port["after_npv"]), tone="acc"),
        dict(label="Value Created", value=fmt_money(vc, signed=True),
             dir_=_dir(vc, 1), tone="pos" if _dir(vc, 1) > 0 else "neg"),
        dict(label="NPV/Investment", value=fmt_ratio(eff_a),
             delta=(f"{fmt_ratio(eff_a - eff_b, signed=True)} vs before"
                    if not (_bad(eff_a) or _bad(eff_b)) else None),
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

    section("02", "Cash-Flow Trajectory", "Portfolio, undiscounted")
    if not panel.empty:
        c1, c2 = st.columns(2)
        with c1:
            show_fig(chart_monthly(panel, "cf", "Monthly Cash Flow — Before vs After"),
                     "sum_mcf")
        with c2:
            fig = go.Figure()
            for pre, col, nm, dash in (("before", CLR["before"], "Before", "dash"),
                                       ("after", CLR["after"], "After", "solid"),
                                       ("incr", CLR["incr"], "Incremental", "solid")):
                fig.add_trace(go.Scatter(
                    x=panel["date"], y=panel[f"cum_{pre}_cf"],
                    name=nm, mode="lines",
                    line=dict(color=col, width=2.4, dash=dash)))
            fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
            show_fig(_fin(fig, "Cumulative Cash Flow", 400, "$", tickpre="$"), "sum_ccf")
    else:
        st.info("No forecast data matched to events.")

    section("03", "Case Contribution")
    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_case_bars(case_agg, "npv", "NPV by Case"), "sum_cnpv")
    with c2:
        show_fig(chart_case_bars(case_agg, "invest", "Investment by Case"), "sum_cinv")


def page_economics(ev, port, case_agg, consol, fc, panel):
    section("01", "Portfolio Economics — Before vs After")

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
            "Metric": label, "Before": fmt_fn(b), "After": fmt_fn(a),
            "Delta": fmt_fn(d, signed=True),
            "% Change": fmt_pct(pct_chg, signed=True) if not _bad(pct_chg) else DASH,
            "Direction": verdict,
        })
    show_df(pd.DataFrame(rows), height=540)

    section("02", "Per-Well Normalisation")
    pw = []
    fmap = {k: f for k, _, f, _ in METRIC_DEFS}
    lmap = {k: l for k, l, _, _ in METRIC_DEFS}
    for k in ADD_KEYS:
        f = fmap.get(k, fmt_money)
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

    section("03", "All Events")
    disp = pd.DataFrame({
        "Event": ev["event_id"], "Case": ev["case"],
        "Wells": ev["wells_before"].map(lambda x: f"{x:.0f}") + " → "
                 + ev["wells_after"].map(lambda x: f"{x:.0f}"),
        "Before NPV": ev["before_npv"].map(fmt_money),
        "After NPV": ev["after_npv"].map(fmt_money),
        "Δ NPV": ev["delta_npv"].map(lambda x: fmt_money(x, signed=True)),
        "Δ Investment": ev["delta_invest"].map(lambda x: fmt_money(x, signed=True)),
        "NPV/I After": ev["after_npv_inv"].map(fmt_ratio),
    })
    show_df(disp, height=460)

    section("04", "Migration — Investment vs NPV")
    show_fig(chart_migration(ev), "ec_mig")


def page_cases(ev, port, case_agg, consol, fc, panel):
    section("01", "Case Archetypes", "Consolidation · Extension · Creation")
    present = [c for c in CASE_ORDER if c in case_agg]
    if not present:
        st.info("No classified events found.")
        return

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
            "Value Created": fmt_money(r.get("delta_npv"), signed=True),
            "Investment Saved": fmt_money(r.get("invest_saved"), signed=True),
            "NPV/I Before": fmt_ratio(r["before_npv_inv"]),
            "NPV/I After": fmt_ratio(r["after_npv_inv"]),
        })
    show_df(pd.DataFrame(summary))

    tabs = st.tabs(present)
    for tab, case in zip(tabs, present):
        with tab:
            r = case_agg[case]
            sub = ev[ev["case"] == case]
            st.markdown(
                f'<span class="pill" style="background:{CASE_CLR[case]}">{case}</span>'
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

            st.markdown("**Economics Comparison**")
            rows = []
            for key, label, fmt_fn, good in METRIC_DEFS:
                b = r.get(f"before_{key}", np.nan)
                a = r.get(f"after_{key}", np.nan)
                d = (a - b) if not (_bad(a) or _bad(b)) else np.nan
                rows.append({"Metric": label, "Before": fmt_fn(b), "After": fmt_fn(a),
                             "Delta": fmt_fn(d, signed=True)})

            c1, c2 = st.columns([1, 1.25])
            with c1:
                show_df(pd.DataFrame(rows), height=430)
            with c2:
                case_eids = set(sub["event_id"])
                cpanel = _build_panel(consol, fc, case_eids)
                if not cpanel.empty:
                    show_fig(chart_monthly(cpanel, "cf", f"{case} — Monthly Cash Flow"),
                             f"cs_cf_{case}")
                else:
                    st.info("No forecast coverage for this case.")

            st.markdown("**Events**")
            top = sub.sort_values("delta_npv", ascending=False)
            disp = pd.DataFrame({
                "Event": top["event_id"],
                "W1": top["w1_ent"], "W2": top["w2_ent"], "Enhanced": top["enh_ent"],
                "Before NPV": top["before_npv"].map(fmt_money),
                "After NPV": top["after_npv"].map(fmt_money),
                "Δ NPV": top["delta_npv"].map(lambda x: fmt_money(x, signed=True)),
                "Δ Investment": top["delta_invest"].map(lambda x: fmt_money(x, signed=True)),
            })
            show_df(disp)


def page_forecast(ev, port, case_agg, consol, fc, panel):
    section("01", "Forecast Analysis")
    if panel.empty:
        st.warning("No forecast data matched to events.")
        return

    tabs = st.tabs(["Cash Flow", "Revenue", "Operating Income"])
    for tab, val, lab in zip(tabs, ["cf", "revenue", "opinc"],
                             ["Cash Flow", "Revenue", "Operating Income"]):
        with tab:
            c1, c2 = st.columns(2)
            with c1:
                show_fig(chart_monthly(panel, val, f"Monthly {lab}"), f"fc_m_{val}")
            with c2:
                show_fig(chart_incremental_bars(panel, val,
                                                f"Incremental {lab} + Cumulative"),
                         f"fc_i_{val}")

    section("02", "Cumulative Cash Flow")
    fig = go.Figure()
    for pre, col, nm, dash in (("before", CLR["before"], "Before", "dash"),
                               ("after", CLR["after"], "After", "solid"),
                               ("incr", CLR["incr"], "Incremental", "solid")):
        fig.add_trace(go.Scatter(
            x=panel["date"], y=panel[f"cum_{pre}_cf"], name=nm, mode="lines",
            line=dict(color=col, width=2.4, dash=dash)))
    fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
    show_fig(_fin(fig, "Cumulative Cash Flow", 420, "$", tickpre="$"), "fc_cum")

    section("03", "Annual Summary")
    ann = panel.groupby("year").agg(
        before_revenue=("before_revenue", "sum"),
        after_revenue=("after_revenue", "sum"),
        before_cf=("before_cf", "sum"),
        after_cf=("after_cf", "sum"),
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


def page_event_explorer(ev, port, case_agg, consol, fc, panel):
    section("01", "Event Explorer")
    if ev.empty:
        st.info("No events loaded.")
        return

    # Dropdown label: event# · well1name + well2name
    def _label(r):
        parts = [f"#{r['event_id']}"]
        names = []
        if r["w1_name"]:
            names.append(r["w1_name"])
        if r["w2_name"]:
            names.append(r["w2_name"])
        if names:
            parts.append(" + ".join(names))
        parts.append(f"→ {r['enh_ent']}")
        parts.append(f"· {r['case']}")
        return "  ".join(parts)

    lbl_map = {r["event_id"]: _label(r) for _, r in ev.iterrows()}
    pick = st.selectbox("Event", list(ev["event_id"]),
                        format_func=lambda k: lbl_map.get(k, k))
    r = ev[ev["event_id"] == pick].iloc[0]

    st.markdown(
        f'<span class="pill" style="background:{CASE_CLR.get(r["case"], CLR["slate"])}">'
        f'{r["case"]}</span>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Before — Well 1**")
        st.write(r["w1_name"] or DASH)
        st.caption(f"Entity: `{r['w1_ent'] or DASH}`")
        if r["w1_ent"]:
            st.caption(f"Found in econ: {'✅' if r.get('w1_found', False) else '❌'}")
    with c2:
        st.markdown("**Before — Well 2**")
        st.write(r["w2_name"] or DASH)
        st.caption(f"Entity: `{r['w2_ent'] or DASH}`")
        if r["w2_ent"]:
            st.caption(f"Found in econ: {'✅' if r.get('w2_found', False) else '❌'}")
    with c3:
        st.markdown("**After — Enhanced Well**")
        st.write(f"`{r['enh_ent']}`")
        st.caption(f"Found in econ: {'✅' if r.get('enh_found', False) else '❌'}")
        st.caption(f"Wells: {r['wells_before']:.0f} → {r['wells_after']:.0f} "
                   f"({fmt_int(r['wells_net'], signed=True)})")

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

    st.markdown("**Full Economics**")
    rows = []
    for key, label, fmt_fn, good in METRIC_DEFS:
        b = r.get(f"before_{key}", np.nan)
        a = r.get(f"after_{key}", np.nan)
        d = (a - b) if not (_bad(a) or _bad(b)) else np.nan
        rows.append({"Metric": label, "Before": fmt_fn(b), "After": fmt_fn(a),
                     "Delta": fmt_fn(d, signed=True)})
    show_df(pd.DataFrame(rows), height=430)

    # Per-entity cash flow chart
    section("02", "Type Curve Cash Flows")
    entities = {}
    if r["w1_ent"]:
        entities[r["w1_ent"]] = f"W1: {r['w1_ent']}"
    if r["w2_ent"] and r["w2_ent"] != r["w1_ent"]:
        entities[r["w2_ent"]] = f"W2: {r['w2_ent']}"
    entities[r["enh_ent"]] = f"Enhanced: {r['enh_ent']}"

    if not fc.empty and entities:
        show_fig(chart_entity_cf(fc, entities, "Monthly Cash Flow by Type Curve"), "ee_tcf")

    section("03", "Event Combined Cash Flow")
    ev_consol = [e for e in consol if e["event_id"] == pick]
    ep = _build_panel(ev_consol, fc)
    if ep.empty:
        st.info("No forecast rows matched to this event.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            show_fig(chart_monthly(ep, "cf", "Monthly Cash Flow — Before vs After"), "ee_cf")
        with c2:
            show_fig(chart_incremental_bars(ep, "cf", "Incremental Cash Flow"), "ee_icf")


# ─── PAGE REGISTRY ──────────────────────────────────────────────────────────

PAGES = {
    "Executive Summary": page_summary,
    "Portfolio Economics": page_economics,
    "Case Analysis": page_cases,
    "Forecast Analysis": page_forecast,
    "Event Explorer": page_event_explorer,
}


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    header()

    try:
        ev_df, consol, fc, econ, debug = build_events()
    except (FileNotFoundError, ValueError) as e:
        st.error(str(e))
        st.info("Place `economics.xlsx`, `consol.xlsx`, `forecast.xlsx`, and `capex.xlsx` "
                "next to `app.py` or in a `./data` subfolder.")
        return

    if ev_df.empty:
        st.warning("No events built from consol.xlsx.")
        with st.expander("🔍 Debug — why no events?"):
            st.write(f"**Economics entities loaded:** {debug['econ_entities']}")
            st.write(f"**Consol events parsed:** {debug['consol_events']}")
            st.write(f"**Forecast rows:** {debug['fc_rows']}")
            st.write("**Economics column matching:**")
            st.json(debug["econ_columns"])
            st.write("**Sample entity keys from economics:**", debug["sample_econ_keys"])
            st.write("**Sample enhanced entities from consol:**", debug["sample_consol_enh"])
            st.write("**Sample w1 entities from consol:**", debug["sample_consol_w1"])
            # Check overlap
            econ_set = set(debug["sample_econ_keys"])
            consol_enh = debug["sample_consol_enh"]
            for e in consol_enh:
                st.write(f"  `{e}` in economics? {e in econ_set}")
        return

    # Check join health
    n_enh_found = ev_df["enh_found"].sum()
    n_enh_miss = len(ev_df) - n_enh_found
    if n_enh_miss > 0:
        callout("bad", f"<b>{n_enh_miss}</b> of {len(ev_df)} enhanced entities not found "
                       f"in economics.xlsx — their After values will be blank.")

    # Sidebar
    sb = st.sidebar
    sb.markdown("### ⛽ Navigation")
    page = sb.radio("Page", list(PAGES), label_visibility="collapsed")
    sb.markdown("---")
    sb.markdown("### Scope")
    cases = sb.multiselect("Event Cases", CASE_ORDER, default=CASE_ORDER)
    sb.markdown("---")
    sb.caption(f"**{debug['econ_entities']}** economics entities")
    sb.caption(f"**{len(ev_df)}** events ({n_enh_found} enhanced matched)")
    sb.caption(f"**{debug['fc_rows']:,}** forecast rows")
    sb.markdown("---")
    if sb.button("🔄 Reload Data"):
        st.cache_data.clear()
        st.rerun()

    # Show debug expander on every page
    with st.sidebar.expander("🔍 Debug"):
        st.write("**Econ columns found:**")
        st.json(debug["econ_columns"]["columns_matched"])
        if debug["econ_columns"]["columns_missing"]:
            st.warning(f"Missing: {debug['econ_columns']['columns_missing']}")
        st.write("**Sample econ keys:**", debug["sample_econ_keys"][:5])
        st.write("**Sample consol enh:**", debug["sample_consol_enh"][:5])
        st.write("**Sample consol w1:**", debug["sample_consol_w1"][:5])

    # Filter
    ev = ev_df[ev_df["case"].isin(cases)].copy()
    if ev.empty:
        st.warning("No events match the selected cases.")
        return

    port = _aggregate(ev)

    case_agg = {}
    for c in CASE_ORDER:
        sub = ev[ev["case"] == c]
        if not sub.empty:
            case_agg[c] = _aggregate(sub)

    panel = _build_panel(consol, fc, set(ev["event_id"]))

    PAGES[page](ev, port, case_agg, consol, fc, panel)

    st.markdown("---")
    st.caption(f"{len(ev):,} events · economics ×{ECON_MONEY_SCALE:g}")


if __name__ == "__main__":
    main()