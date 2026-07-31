"""
=============================================================================
 INVENTORY REDESIGN  —  CAPITAL EFFICIENCY EXECUTIVE DASHBOARD
=============================================================================
 Quantifies capital efficiency gains from redesigning drilling inventory by
 comparing a BEFORE inventory (existing wells) with an AFTER inventory
 (enhanced wells) across three event archetypes:

     Consolidation : 2 legacy wells -> 1 enhanced well   (net -1 well)
     Extension     : 1 legacy well  -> 1 enhanced well   (net  0 wells)
     Creation      : 0 legacy wells -> 1 enhanced well   (net +1 well)

 Inputs (auto-loaded from the app directory or ./data, no upload widgets):
     economics.xlsx, consol.xlsx, forecast.xlsx, capex.xlsx

 Run:  streamlit run app.py
=============================================================================
"""

from __future__ import annotations

import inspect
import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

# =============================================================================
# SECTION 1 — APP CONFIGURATION
# =============================================================================

APP_TITLE = "Inventory Redesign — Capital Efficiency"
APP_SUBTITLE = "Upstream Portfolio Value Creation Review"
APP_ICON = "⛽"

BASE_DIR = Path(__file__).resolve().parent
SEARCH_DIRS = [BASE_DIR, BASE_DIR / "data", BASE_DIR / "inputs", Path.cwd()]
READABLE_EXT = (".xlsx", ".xlsm", ".xls", ".csv", ".parquet")

FILE_STEMS = {
    "economics": "economics",
    "consol": "consol",
    "forecast": "forecast",
    "capex": "capex",
}

# --- Unit scaling ------------------------------------------------------------
# Everything is normalised to BASE DOLLARS internally. Aries/PHDWin "M$" means
# THOUSANDS of dollars, hence 1_000.0. Override in the sidebar if your exports
# differ; get this wrong and every dollar figure is off by 1000x.
DEFAULT_SCALES = {
    "economics": 1_000.0,   # M$  -> $
    "capex": 1.0,           # $   -> $
    "forecast": 1.0,        # $   -> $
}

CASE_ORDER = ["Consolidation", "Extension", "Creation", "Unclassified"]

# --- Palette -----------------------------------------------------------------
CLR = {
    "navy": "#12314F",
    "blue": "#0E6BA8",
    "teal": "#0E9AA7",
    "amber": "#C98A1B",
    "green": "#2E7D5B",
    "red": "#B3403A",
    "slate": "#8A9AA9",
    "light": "#D8E1E8",
    "ink": "#12222E",
    "before": "#8A9AA9",
    "after": "#0E6BA8",
    "incr": "#0E9AA7",
}
CASE_CLR = {
    "Consolidation": CLR["navy"],
    "Extension": CLR["teal"],
    "Creation": CLR["amber"],
    "Unclassified": CLR["red"],
}

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# SECTION 2 — STREAMLIT COMPAT + STYLING
# =============================================================================

_WIDTH_CACHE: dict[int, dict] = {}


def _wkw(fn: Callable) -> dict:
    """Return the correct full-width kwarg for the installed Streamlit build.

    `use_container_width` was deprecated in favour of `width="stretch"`;
    this keeps the app working on either side of that change.
    """
    fid = id(fn)
    if fid not in _WIDTH_CACHE:
        try:
            params = inspect.signature(fn).parameters
        except (TypeError, ValueError):
            params = {}
        _WIDTH_CACHE[fid] = (
            {"width": "stretch"} if "width" in params else {"use_container_width": True}
        )
    return _WIDTH_CACHE[fid]


def show_fig(fig: go.Figure, key: str | None = None) -> None:
    kw = dict(_wkw(st.plotly_chart))
    kw["config"] = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]}
    if key:
        kw["key"] = key
    try:
        st.plotly_chart(fig, **kw)
    except TypeError:                                  # very old / very new build
        st.plotly_chart(fig)


def show_df(df: pd.DataFrame, height: int | None = None, **kwargs) -> None:
    kw = dict(_wkw(st.dataframe))
    kw.update(hide_index=True)
    if height:
        kw["height"] = height
    kw.update(kwargs)
    try:
        st.dataframe(df, **kw)
    except TypeError:
        st.dataframe(df, hide_index=True)


CSS = """
<style>
  .block-container {padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1500px;}
  h1, h2, h3, h4 {color: #12222E; letter-spacing: -0.01em;}

  .app-header {
      border-left: 5px solid #0E6BA8; padding: .35rem 0 .35rem .9rem;
      margin-bottom: 1.1rem;
  }
  .app-header .t {font-size: 1.55rem; font-weight: 700; color:#12314F; line-height:1.15;}
  .app-header .s {font-size: .86rem; color:#66788A; text-transform: uppercase;
                  letter-spacing:.10em; margin-top:.15rem;}

  .sec {
      display:flex; align-items:center; gap:.6rem;
      margin: 1.7rem 0 .7rem 0; padding-bottom:.4rem;
      border-bottom: 1px solid #E3E9ED;
  }
  .sec .n {font-size:.72rem; font-weight:700; color:#FFF; background:#12314F;
           border-radius:3px; padding:.10rem .42rem;}
  .sec .h {font-size:1.05rem; font-weight:650; color:#12314F;}
  .sec .d {font-size:.80rem; color:#8A9AA9; margin-left:auto;}

  .kpi {
      background:#FFF; border:1px solid #E3E9ED; border-radius:8px;
      padding:.75rem .85rem; height:100%;
      box-shadow: 0 1px 2px rgba(18,49,79,.05);
      border-top:3px solid #D8E1E8;
  }
  .kpi.pos {border-top-color:#2E7D5B;} .kpi.neg {border-top-color:#B3403A;}
  .kpi.acc {border-top-color:#0E6BA8;} .kpi.warn{border-top-color:#C98A1B;}
  .kpi .l {font-size:.685rem; font-weight:600; color:#7A8A99;
           text-transform:uppercase; letter-spacing:.055em; line-height:1.25;
           min-height:2.0em;}
  .kpi .v {font-size:1.42rem; font-weight:700; color:#12222E; line-height:1.25;
           margin-top:.18rem; font-variant-numeric: tabular-nums;}
  .kpi .d {font-size:.775rem; font-weight:600; margin-top:.10rem;}
  .kpi .d.up {color:#2E7D5B;} .kpi .d.dn {color:#B3403A;} .kpi .d.nt{color:#8A9AA9;}
  .kpi .s {font-size:.705rem; color:#98A6B3; margin-top:.18rem;}

  .callout {
      border-radius:8px; padding:.85rem 1rem; margin:.5rem 0 1rem 0;
      font-size:.90rem; line-height:1.55; border:1px solid; background:#F7FAFC;
  }
  .callout.good {border-color:#BEDCCB; background:#F2F9F5; color:#1F5A41;}
  .callout.bad  {border-color:#E7C4C2; background:#FDF5F4; color:#8A2F2A;}
  .callout.info {border-color:#C9DCEA; background:#F4F9FD; color:#154B72;}
  .callout.warn {border-color:#EBD8AE; background:#FDFAF1; color:#7A5510;}

  .pill {display:inline-block; font-size:.70rem; font-weight:600; padding:.10rem .48rem;
         border-radius:10px; margin-right:.3rem; color:#FFF;}

  div[data-testid="stMetricValue"] {font-size:1.3rem;}
  section[data-testid="stSidebar"] {background:#F7F9FB; border-right:1px solid #E3E9ED;}
  section[data-testid="stSidebar"] .block-container {padding-top:1.2rem;}
  .stTabs [data-baseweb="tab"] {font-size:.88rem; font-weight:600;}
  footer, #MainMenu {visibility:hidden;}
</style>
"""

pio.templates["execdash"] = go.layout.Template(
    layout=dict(
        font=dict(family="Inter, Segoe UI, Helvetica, sans-serif", size=12, color=CLR["ink"]),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        colorway=[CLR["navy"], CLR["blue"], CLR["teal"], CLR["amber"],
                  CLR["green"], CLR["red"], CLR["slate"]],
        margin=dict(l=60, r=25, t=55, b=55),
        title=dict(font=dict(size=15, color=CLR["navy"]), x=0.0, xanchor="left"),
        xaxis=dict(gridcolor="#EEF2F5", zerolinecolor="#D8E1E8", linecolor="#D8E1E8"),
        yaxis=dict(gridcolor="#EEF2F5", zerolinecolor="#D8E1E8", linecolor="#D8E1E8"),
        legend=dict(orientation="h", y=-0.18, x=0, font=dict(size=11),
                    bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(font_size=12, bgcolor="#FFFFFF", bordercolor="#D8E1E8"),
    )
)
pio.templates.default = "plotly_white+execdash"


# =============================================================================
# SECTION 3 — FORMATTERS
# =============================================================================

DASH = "—"


def _bad(v) -> bool:
    return v is None or (isinstance(v, (float, np.floating)) and not np.isfinite(v)) or pd.isna(v)


def fmt_money(v, dp: int = 1, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    v = float(v)
    sign = "-" if v < 0 else ("+" if signed and v > 0 else "")
    a = abs(v)
    for div, suf in ((1e9, "B"), (1e6, "MM"), (1e3, "K")):
        if a >= div:
            return f"{sign}${a / div:,.{dp}f}{suf}"
    return f"{sign}${a:,.0f}"


def fmt_vol(v, unit: str = "boe", dp: int = 1, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    v = float(v)
    sign = "-" if v < 0 else ("+" if signed and v > 0 else "")
    a = abs(v)
    if a >= 1e6:
        return f"{sign}{a / 1e6:,.{dp}f} MM{unit}"
    if a >= 1e3:
        return f"{sign}{a / 1e3:,.{dp}f} M{unit}"
    return f"{sign}{a:,.0f} {unit}"


def fmt_rate(v, dp: int = 0, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    sign = "+" if signed and v > 0 else ""
    return f"{sign}{v:,.{dp}f} boe/d"


def fmt_ratio(v, dp: int = 2, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    sign = "+" if signed and v > 0 else ""
    return f"{sign}{v:,.{dp}f}x"


def fmt_pct(v, dp: int = 1, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    sign = "+" if signed and v > 0 else ""
    return f"{sign}{v:,.{dp}f}%"


def fmt_years(v, dp: int = 2, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    sign = "+" if signed and v > 0 else ""
    return f"{sign}{v:,.{dp}f} yr"


def fmt_usdboe(v, dp: int = 2, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    sign = "+" if signed and v > 0 else ""
    return f"{sign}${v:,.{dp}f}/boe"


def fmt_int(v, signed: bool = False) -> str:
    if _bad(v):
        return DASH
    sign = "+" if signed and v > 0 else ""
    return f"{sign}{v:,.0f}"


FMT: dict[str, Callable] = {
    "money": fmt_money,
    "boe": lambda v, **k: fmt_vol(v, "boe", **k),
    "boepd": fmt_rate,
    "ratio": fmt_ratio,
    "pct": fmt_pct,
    "years": fmt_years,
    "usdboe": fmt_usdboe,
    "int": fmt_int,
}


def safe_div(a, b):
    """Element-wise or scalar division that returns NaN instead of inf/ZeroDivision."""
    if isinstance(a, pd.Series) or isinstance(b, pd.Series):
        a = pd.Series(a) if not isinstance(a, pd.Series) else a
        out = a / pd.Series(b).replace(0, np.nan)
        return out.replace([np.inf, -np.inf], np.nan)
    if _bad(a) or _bad(b) or float(b) == 0.0:
        return np.nan
    return float(a) / float(b)


# =============================================================================
# SECTION 4 — METRIC REGISTRY
#   kind = additive | weighted | derived
#   THIS is where the spec's "sum everything" flaw is corrected.
# =============================================================================

@dataclass(frozen=True)
class Metric:
    key: str
    label: str
    kind: str
    fmt: str
    good: int = 1                  # +1 higher is better, -1 lower is better, 0 neutral
    weight: str | None = None      # weighted: metric key used as the weight
    num: str | None = None         # derived: numerator metric key
    den: str | None = None         # derived: denominator metric key
    scale: float = 1.0
    group: str = "economics"
    note: str = ""


METRICS: tuple[Metric, ...] = (
    # ---- additive -----------------------------------------------------------
    Metric("npv", "NPV BTax @10%", "additive", "money", +1),
    Metric("invest", "Investment BTax @0%", "additive", "money", -1),
    Metric("capex", "Capex", "additive", "money", -1, group="capital"),
    Metric("boe", "Net Reserves", "additive", "boe", +1),
    Metric("fy_rate", "1st-Year Rate", "additive", "boepd", +1),
    Metric("ip30", "IP30 Cumulative", "additive", "boe", +1),
    Metric("avg3", "3-Month Avg Rate", "additive", "boepd", +1),
    # ---- derived (recomputed from aggregates, NEVER summed) ------------------
    Metric("npv_inv", "NPV / Investment", "derived", "ratio", +1,
           num="npv", den="invest",
           note="Recomputed as Σ NPV ÷ Σ Investment."),
    Metric("inv_boe", "Investment / BOE", "derived", "usdboe", -1,
           num="invest", den="boe",
           note="Recomputed as Σ Investment ÷ Σ BOE."),
    Metric("npv_boe", "NPV / BOE", "derived", "usdboe", +1,
           num="npv", den="boe"),
    # ---- weighted averages --------------------------------------------------
    Metric("payout", "Payout BTax", "weighted", "years", -1, weight="invest",
           note="Capital-weighted. Cross-check against forecast-implied payback."),
    Metric("ror", "BTax Disc. CF ROR", "weighted", "pct", +1, weight="invest",
           note="Capital-weighted approximation; a true blended IRR requires "
                "combined cash flow (see Forecast Analysis)."),
    Metric("cor", "Cost of Reserves (rept.)", "weighted", "usdboe", -1, weight="boe",
           note="Volume-weighted average of the reported $/boe."),
    Metric("wi", "Initial WI", "weighted", "pct", 0, weight="boe",
           note="Volume-weighted average working interest."),
)

M_BY_KEY = {m.key: m for m in METRICS}
ADDITIVE = [m for m in METRICS if m.kind == "additive"]
WEIGHTED = [m for m in METRICS if m.kind == "weighted"]
DERIVED = [m for m in METRICS if m.kind == "derived"]
ADD_KEYS = [m.key for m in ADDITIVE]
ALL_KEYS = [m.key for m in METRICS]
ECON_KEYS = [m.key for m in METRICS if m.group == "economics"]

SUM_COLS = (
    ADD_KEYS
    + [f"_num_{m.key}" for m in WEIGHTED]
    + [f"_den_{m.key}" for m in WEIGHTED]
)

PER_WELL = {  # metric key -> label used in the per-well normalisation table
    "npv": "NPV / well",
    "invest": "Investment / well",
    "capex": "Capex / well",
    "boe": "Reserves / well",
    "fy_rate": "1st-Yr Rate / well",
    "ip30": "IP30 / well",
    "avg3": "3-Mo Rate / well",
}


def fmt_metric(key: str, v, signed: bool = False) -> str:
    m = M_BY_KEY.get(key)
    f = FMT[m.fmt] if m else fmt_int
    try:
        return f(v, signed=signed)
    except TypeError:
        return f(v)


# =============================================================================
# SECTION 5 — INPUT SCHEMA (alias-tolerant column resolution)
# =============================================================================

@dataclass(frozen=True)
class Fld:
    key: str
    aliases: tuple[str, ...]
    kind: str = "num"       # num | txt | int
    required: bool = True


SCHEMA: dict[str, tuple[Fld, ...]] = {
    "economics": (
        Fld("entity", ("Entity", "entity_name", "Entity Name", "Well", "Case", "Propnum"), "txt"),
        Fld("npv", ("Npv Cash Flow BTax 10.0% (M$)", "NPV Cash Flow BTax 10% (M$)",
                    "NPV BTax 10%", "NPV 10", "npv")),
        Fld("npv_inv_rep", ("NPV / Disc. Invest BTax", "NPV/Disc Invest BTax",
                            "NPV per Disc Invest"), required=False),
        Fld("payout", ("Payout BTax (years)", "Payout BTax", "Payout")),
        Fld("boe", ("Boe WI Total (boe)", "BOE WI Total", "Net BOE", "boe")),
        Fld("fy_rate", ("1st Year Production Rate (boepd)", "First Year Production Rate",
                        "1st Yr Rate")),
        Fld("cor", ("Cost of Reserves ($/boe)", "Cost of Reserves", "COR")),
        Fld("ip30", ("IP30 Cum (boe)", "IP30 Cum", "IP30")),
        Fld("invest", ("Npv Investment BTax 0.0% (M$)", "NPV Investment BTax 0% (M$)",
                       "Investment BTax", "Total Investment")),
        Fld("ror", ("BTax Disc. CF. ROR (%)", "BTax Disc CF ROR", "ROR", "IRR")),
        Fld("wi", ("Initial WI (%)", "Initial WI", "WI")),
        Fld("avg3", ("3 Month Avg Production (boepd)", "3 Mo Avg Production",
                     "3 Month Avg Rate")),
    ),
    "consol": (
        Fld("event_id", ("consolidation #", "consolidation number", "consol #",
                         "event", "event id", "id"), "txt", required=False),
        Fld("w1_name", ("well 1 name", "well1 name", "well_1_name"), "txt", required=False),
        Fld("w2_name", ("well 2 name", "well2 name", "well_2_name"), "txt", required=False),
        Fld("w1_ent", ("well 1 entity", "well1 entity", "well_1_entity"), "txt"),
        Fld("w2_ent", ("well 2 entity", "well2 entity", "well_2_entity"), "txt"),
        Fld("enh_ent", ("enhanced well entity", "enhanced entity", "enhanced well"), "txt"),
    ),
    "forecast": (
        Fld("entity", ("entity_name", "Entity", "Entity Name", "Well"), "txt"),
        Fld("year", ("year", "yr", "cal_year"), "int"),
        Fld("month", ("month", "mo", "cal_month"), "int"),
        Fld("revenue", ("total_revenue", "revenue", "Total Revenue")),
        Fld("opinc", ("operating_income", "op_income", "Operating Income")),
        Fld("cf", ("cash_flow", "cashflow", "Cash Flow", "net_cash_flow")),
    ),
    "capex": (
        Fld("entity", ("entity", "Entity", "entity_name", "Well"), "txt"),
        Fld("capex", ("capex", "CAPEX", "Capital", "total_capex")),
    ),
}


def norm_header(s) -> str:
    """lower / de-punctuate / collapse spaces  ->  comparable header token."""
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def clean_text(v) -> str:
    """Trim, collapse repeated internal whitespace, blank-out null sentinels."""
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return ""
    s = re.sub(r"\s+", " ", str(v)).strip()
    return "" if s.lower() in {"nan", "none", "null", "n/a", "na", "#n/a", "-", "."} else s


def norm_key(v) -> str:
    """Canonical entity join key: trimmed, single-spaced, case-insensitive."""
    return clean_text(v).upper()


def to_num(s: pd.Series) -> pd.Series:
    """Coerce messy Excel numerics: $ , % blanks and (123) negatives."""
    if pd.api.types.is_numeric_dtype(s):
        return s.astype("float64")
    t = (s.astype("string")
           .str.replace(r"[,$\s%]", "", regex=True)
           .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
           .replace({"": None, "-": None, "NA": None, "N/A": None, "#N/A": None}))
    return pd.to_numeric(t, errors="coerce")


def _score(target: str, cand: str) -> float:
    a, b = set(target.split()), set(cand.split())
    if not a or not b:
        return 0.0
    inter = len(a & b)
    return 2 * inter / (len(a) + len(b))


def standardize(df: pd.DataFrame, flds: Sequence[Fld], label: str):
    """Rename source columns to canonical keys; coerce dtypes; report issues."""
    lookup = {norm_header(c): c for c in df.columns}
    out, notes, missing = pd.DataFrame(index=df.index), [], []

    for f in flds:
        src = None
        for a in f.aliases:                                    # 1) exact (normalised)
            if norm_header(a) in lookup:
                src = lookup[norm_header(a)]
                break
        if src is None:                                        # 2) fuzzy token overlap
            best, bs = None, 0.0
            for a in f.aliases:
                for nh, orig in lookup.items():
                    sc = _score(norm_header(a), nh)
                    if sc > bs:
                        best, bs = orig, sc
            if bs >= 0.62:
                src, _ = best, notes.append(
                    f"`{label}`: matched **{f.key}** to column *{best}* (fuzzy {bs:.0%})."
                )
        if src is None:
            out[f.key] = "" if f.kind == "txt" else np.nan
            if f.required:
                missing.append(f.key)
            continue

        col = df[src]
        if f.kind == "txt":
            out[f.key] = col.map(clean_text)
        else:
            num = to_num(col)
            bad = int(col.notna().sum() - num.notna().sum())
            if bad:
                notes.append(f"`{label}`: {bad:,} non-numeric value(s) in *{src}* set to null.")
            out[f.key] = num.astype("Int64").astype("float64") if f.kind == "int" else num
    return out, notes, missing


# =============================================================================
# SECTION 6 — FILE DISCOVERY & LOADING
# =============================================================================

def find_file(stem: str) -> Path | None:
    for d in SEARCH_DIRS:
        if not d.exists():
            continue
        for p in sorted(d.iterdir()):
            if (p.is_file()
                    and p.suffix.lower() in READABLE_EXT
                    and p.stem.strip().lower() == stem.lower()
                    and not p.name.startswith("~$")):
                return p
    return None


def file_signature() -> tuple:
    sig = []
    for name, stem in FILE_STEMS.items():
        p = find_file(stem)
        sig.append((name, str(p) if p else None,
                    p.stat().st_mtime_ns if p else 0, p.stat().st_size if p else 0))
    return tuple(sig)


@st.cache_data(show_spinner=False)
def read_any(path_str: str, _sig: tuple) -> pd.DataFrame:
    p = Path(path_str)
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p)
    if p.suffix.lower() == ".parquet":
        return pd.read_parquet(p)
    return pd.read_excel(p, sheet_name=0)


@dataclass
class Raw:
    econ: pd.DataFrame
    consol: pd.DataFrame
    forecast: pd.DataFrame
    capex: pd.DataFrame
    notes: list[str] = field(default_factory=list)
    missing: dict[str, list[str]] = field(default_factory=dict)
    paths: dict[str, str] = field(default_factory=dict)


@st.cache_data(show_spinner="Loading workbooks…")
def load_raw(_sig: tuple) -> Raw:
    frames, notes, missing, paths = {}, [], {}, {}
    for name, stem in FILE_STEMS.items():
        p = find_file(stem)
        if p is None:
            raise FileNotFoundError(
                f"{stem}.xlsx not found. Searched: "
                + ", ".join(str(d) for d in SEARCH_DIRS if d.exists())
            )
        paths[name] = str(p)
        df = read_any(str(p), _sig)
        std, n, miss = standardize(df, SCHEMA[name], name)
        frames[name], missing[name] = std, miss
        notes += n
        notes.append(f"`{name}` ← **{p.name}** · {len(df):,} rows × {df.shape[1]} cols")
    return Raw(frames["economics"], frames["consol"], frames["forecast"],
               frames["capex"], notes, missing, paths)


# =============================================================================
# SECTION 7 — ENTITY MASTER
# =============================================================================

def maybe_pct(s: pd.Series) -> pd.Series:
    """WI/ROR stored as a 0-1 fraction is rescaled to 0-100."""
    v = s.dropna()
    if len(v) and v.abs().max() <= 1.5:
        return s * 100.0
    return s


@st.cache_data(show_spinner=False)
def build_entities(_sig: tuple, scales: tuple, dup_policy: str):
    """One row per normalised entity key, carrying every metric + join flags."""
    raw = load_raw(_sig)
    s_econ, s_capex, _ = scales
    issues: list[dict] = []

    # ---- economics ----------------------------------------------------------
    e = raw.econ.copy()
    e["entity_key"] = e["entity"].map(norm_key)
    e = e[e["entity_key"] != ""]

    dup = e["entity_key"].duplicated(keep=False)
    if dup.any():
        for k, g in e[dup].groupby("entity_key"):
            issues.append(dict(severity="High", category="Duplicate entity",
                               scope="economics.xlsx", ref=g["entity"].iloc[0],
                               detail=f"{len(g)} rows for one entity; policy = {dup_policy}."))
    if dup_policy == "sum":
        num = [c for c in e.columns if c not in ("entity", "entity_key")]
        agg = e.groupby("entity_key")[num].sum(min_count=1)
        agg["entity"] = e.groupby("entity_key")["entity"].first()
        e = agg.reset_index()
    elif dup_policy == "mean":
        num = [c for c in e.columns if c not in ("entity", "entity_key")]
        agg = e.groupby("entity_key")[num].mean()
        agg["entity"] = e.groupby("entity_key")["entity"].first()
        e = agg.reset_index()
    else:
        e = e.drop_duplicates("entity_key", keep="first")

    e["npv"] = e["npv"] * s_econ
    e["invest"] = e["invest"] * s_econ
    e["ror"] = maybe_pct(e["ror"])
    e["wi"] = maybe_pct(e["wi"])
    e["has_econ"] = True

    # ---- capex --------------------------------------------------------------
    c = raw.capex.copy()
    c["entity_key"] = c["entity"].map(norm_key)
    c = c[c["entity_key"] != ""]
    cdup = c["entity_key"].duplicated(keep=False)
    if cdup.any():
        for k, g in c[cdup].groupby("entity_key"):
            issues.append(dict(severity="Medium", category="Duplicate entity",
                               scope="capex.xlsx", ref=g["entity"].iloc[0],
                               detail=f"{len(g)} capex rows summed."))
    c = (c.groupby("entity_key", as_index=False)
           .agg(capex=("capex", "sum"), entity_capex=("entity", "first")))
    c["capex"] = c["capex"] * s_capex
    c["has_capex"] = True

    ents = e.merge(c[["entity_key", "capex", "has_capex", "entity_capex"]],
                   on="entity_key", how="outer")
    ents["entity"] = ents["entity"].fillna(ents.get("entity_capex"))
    ents = ents.drop(columns=[x for x in ("entity_capex",) if x in ents])
    for f in ("has_econ", "has_capex"):
        ents[f] = ents[f].fillna(False).astype(bool)
    for k in ALL_KEYS + ["npv_inv_rep"]:
        if k not in ents:
            ents[k] = np.nan

    # ---- weighted-average helper columns -----------------------------------
    for m in WEIGHTED:
        w = ents[m.weight]
        v = ents[m.key]
        ents[f"_num_{m.key}"] = v * w
        ents[f"_den_{m.key}"] = w.where(v.notna() & w.notna())

    keep = (["entity_key", "entity", "has_econ", "has_capex", "npv_inv_rep"]
            + ALL_KEYS + [c for c in ents.columns if c.startswith(("_num_", "_den_"))])
    return ents[keep].reset_index(drop=True), issues


@st.cache_data(show_spinner=False)
def build_forecast(_sig: tuple, scales: tuple):
    """Monthly forecast, one row per (entity_key, date). Returns (df, issues)."""
    raw = load_raw(_sig)
    s_fc = scales[2]
    f = raw.forecast.copy()
    f["entity_key"] = f["entity"].map(norm_key)
    f = f[f["entity_key"] != ""]

    issues = []
    y = pd.to_numeric(f["year"], errors="coerce")
    mo = pd.to_numeric(f["month"], errors="coerce")
    ok = y.between(1900, 2200) & mo.between(1, 12)
    if (~ok).any():
        issues.append(dict(severity="High", category="Invalid date",
                           scope="forecast.xlsx", ref="year/month",
                           detail=f"{int((~ok).sum()):,} row(s) dropped — "
                                  "year/month out of range or null."))
    f = f[ok].copy()
    f["date"] = pd.to_datetime(
        dict(year=y[ok].astype(int), month=mo[ok].astype(int), day=1))

    for c in ("revenue", "opinc", "cf"):
        f[c] = f[c] * s_fc

    dup = f.duplicated(["entity_key", "date"], keep=False)
    if dup.any():
        issues.append(dict(severity="Medium", category="Duplicate period",
                           scope="forecast.xlsx", ref="entity × month",
                           detail=f"{int(dup.sum()):,} duplicate rows aggregated by sum."))
    out = (f.groupby(["entity_key", "date"], as_index=False)[["revenue", "opinc", "cf"]]
             .sum(min_count=1)
             .sort_values(["entity_key", "date"]))
    out["year"] = out["date"].dt.year
    out["month"] = out["date"].dt.month
    return out, issues


# =============================================================================
# SECTION 8 — EVENT MEMBERSHIP & CLASSIFICATION
# =============================================================================

SLOTS = (("well1", "Before", "w1_ent", "w1_name"),
         ("well2", "Before", "w2_ent", "w2_name"),
         ("enhanced", "After", "enh_ent", None))


@st.cache_data(show_spinner=False)
def build_membership(_sig: tuple, scales: tuple, dup_policy: str):
    """Long membership table + event metadata + structural issues."""
    raw = load_raw(_sig)
    ents, _ = build_entities(_sig, scales, dup_policy)
    fc, _ = build_forecast(_sig, scales)

    cs = raw.consol.copy().reset_index(drop=True)
    ids = cs["event_id"].map(clean_text)
    cs["event_id"] = [
        v if v else f"ROW-{i + 2:04d}" for i, v in enumerate(ids)
    ]
    if cs["event_id"].duplicated().any():
        cs["event_id"] = [
            f"{v}#{i}" if d else v
            for i, (v, d) in enumerate(zip(cs["event_id"], cs["event_id"].duplicated(keep=False)))
        ]
    for c in ("w1_ent", "w2_ent", "enh_ent", "w1_name", "w2_name"):
        cs[c] = cs[c].map(clean_text)

    # ---- classification -----------------------------------------------------
    n_src = (cs["w1_ent"] != "").astype(int) + (cs["w2_ent"] != "").astype(int)
    cs["n_src"] = n_src
    cs["case"] = np.select(
        [n_src == 2, n_src == 1, n_src == 0],
        ["Consolidation", "Extension", "Creation"], default="Unclassified")
    cs.loc[cs["enh_ent"] == "", "case"] = "Unclassified"

    # ---- long membership ----------------------------------------------------
    parts = []
    for slot, side, ecol, ncol in SLOTS:
        sub = cs[["event_id", "case", ecol]].rename(columns={ecol: "entity_raw"})
        sub["slot"], sub["side"] = slot, side
        sub["well_name"] = cs[ncol] if ncol else ""
        parts.append(sub[sub["entity_raw"] != ""])
    mem = pd.concat(parts, ignore_index=True)
    mem["entity_key"] = mem["entity_raw"].map(norm_key)

    ekeys = set(ents.loc[ents["has_econ"], "entity_key"])
    ckeys = set(ents.loc[ents["has_capex"], "entity_key"])
    fkeys = set(fc["entity_key"].unique())
    mem["in_econ"] = mem["entity_key"].isin(ekeys)
    mem["in_capex"] = mem["entity_key"].isin(ckeys)
    mem["in_fc"] = mem["entity_key"].isin(fkeys)

    # ---- structural data-quality checks ------------------------------------
    iss: list[dict] = []
    for r in cs.itertuples():
        if r.enh_ent == "":
            iss.append(dict(severity="Critical", category="Missing enhanced entity",
                            scope=r.event_id, ref="enhanced well entity",
                            detail="No enhanced entity — event cannot be classified "
                                   "or valued; excluded from After case."))
        if r.n_src == 2 and norm_key(r.w1_ent) == norm_key(r.w2_ent):
            iss.append(dict(severity="High", category="Self-consolidation",
                            scope=r.event_id, ref=r.w1_ent,
                            detail="Well 1 and Well 2 are the same entity; "
                                   "de-duplicated to 1 before well."))
    for r in mem.itertuples():
        if not r.in_econ:
            iss.append(dict(severity="Critical" if r.slot == "enhanced" else "High",
                            category="Missing economics join",
                            scope=r.event_id, ref=r.entity_raw,
                            detail=f"{r.slot} populated but absent from economics.xlsx."))
        if not r.in_capex:
            iss.append(dict(severity="Medium", category="Missing capex join",
                            scope=r.event_id, ref=r.entity_raw,
                            detail=f"{r.slot} absent from capex.xlsx."))
        if not r.in_fc:
            iss.append(dict(severity="Medium", category="Missing forecast join",
                            scope=r.event_id, ref=r.entity_raw,
                            detail=f"{r.slot} absent from forecast.xlsx."))

    reuse = (mem.drop_duplicates(["event_id", "side", "entity_key"])
                .groupby(["side", "entity_key"])["event_id"]
                .agg(list))
    for (side, k), evs in reuse.items():
        if len(evs) > 1:
            iss.append(dict(severity="High", category="Entity reused across events",
                            scope=", ".join(map(str, evs[:6])), ref=k,
                            detail=f"Appears on the {side} side of {len(evs)} events. "
                                   "Portfolio rollups count it once; event rollups "
                                   "count it per event."))
    both = (set(mem.loc[mem.side == "Before", "entity_key"])
            & set(mem.loc[mem.side == "After", "entity_key"]))
    for k in sorted(both):
        iss.append(dict(severity="Medium", category="Chained event",
                        scope="portfolio", ref=k,
                        detail="Entity is a source in one event and the enhanced "
                               "well in another — verify the chain is intended."))

    events_meta = cs[["event_id", "case", "w1_name", "w2_name",
                      "w1_ent", "w2_ent", "enh_ent", "n_src"]].copy()
    return mem, events_meta, iss


# =============================================================================
# SECTION 9 — AGGREGATION ENGINE  (single code path: event / case / portfolio)
# =============================================================================

def _finalize(agg: pd.DataFrame) -> pd.DataFrame:
    """Resolve weighted + derived metrics from the summed helper columns."""
    out = agg.copy()
    for m in WEIGHTED:
        out[m.key] = safe_div(out[f"_num_{m.key}"], out[f"_den_{m.key}"])
    for m in DERIVED:
        out[m.key] = safe_div(out[m.num], out[m.den]) * m.scale
    return out.drop(columns=[c for c in out.columns if c.startswith(("_num_", "_den_"))])


def wide_from_membership(mem: pd.DataFrame, ents: pd.DataFrame,
                         index_cols: Sequence[str], full_index,
                         treat_missing_as_zero: bool) -> pd.DataFrame:
    """Before / After / Delta wide table at an arbitrary grain.

    De-duplicates (grain, side, entity) so a well shared by two events is
    counted once per grain — the double-counting guard.
    """
    idx = list(index_cols) or ["_all"]
    m = mem.copy()
    if not index_cols:
        m["_all"] = "Portfolio"
    m = m.drop_duplicates(subset=idx + ["side", "entity_key"])

    j = m.merge(ents, on="entity_key", how="left", suffixes=("", "_e"))
    for f in ("has_econ", "has_capex"):
        j[f] = j[f].fillna(False).astype(bool)
    j["_fc"] = j["in_fc"].astype(bool)

    g = j.groupby(idx + ["side"], dropna=False)
    agg = g[SUM_COLS].sum(min_count=1)
    agg["n_slots"] = g.size()
    agg["n_econ"] = g["has_econ"].sum()
    agg["n_capex"] = g["has_capex"].sum()
    agg["n_fc"] = g["_fc"].sum()
    agg = _finalize(agg)

    wide = agg.unstack("side")
    wide.columns = [f"{str(sd).lower()}_{col}" for col, sd in wide.columns]

    want = [f"{sd}_{c}" for sd in ("before", "after")
            for c in ALL_KEYS + ["n_slots", "n_econ", "n_capex", "n_fc"]]
    wide = wide.reindex(columns=sorted(set(wide.columns) | set(want)))
    if full_index is not None:
        wide = wide.reindex(pd.Index(full_index, name=idx[0] if len(idx) == 1 else None))

    # ---- completeness gating ------------------------------------------------
    for sd in ("before", "after"):
        ns = wide[f"{sd}_n_slots"].fillna(0)
        ne = wide[f"{sd}_n_econ"].fillna(0)
        nc = wide[f"{sd}_n_capex"].fillna(0)
        wide[f"{sd}_n_slots"], wide[f"{sd}_n_econ"] = ns, ne
        wide[f"{sd}_n_capex"] = nc
        wide[f"{sd}_n_fc"] = wide[f"{sd}_n_fc"].fillna(0)

        econ_ok = ns.eq(0) | ne.eq(ns) | bool(treat_missing_as_zero)
        cap_ok = ns.eq(0) | nc.eq(ns) | bool(treat_missing_as_zero)
        wide[f"{sd}_econ_complete"] = ns.eq(0) | ne.eq(ns)
        wide[f"{sd}_capex_complete"] = ns.eq(0) | nc.eq(ns)

        for k in ECON_KEYS:                       # unknown, not zero
            wide.loc[~econ_ok, f"{sd}_{k}"] = np.nan
        wide.loc[~cap_ok, f"{sd}_capex"] = np.nan
        for k in ADD_KEYS:                        # genuinely zero wells -> zero
            wide.loc[ns.eq(0), f"{sd}_{k}"] = 0.0

    for k in ALL_KEYS:
        wide[f"delta_{k}"] = wide[f"after_{k}"] - wide[f"before_{k}"]

    wide["wells_before"] = wide["before_n_slots"]
    wells_after = wide["after_n_slots"]
    wide["wells_after"] = wells_after
    wide["wells_net"] = wells_after - wide["before_n_slots"]
    wide["complete"] = (wide["before_econ_complete"] & wide["after_econ_complete"]
                        & wells_after.gt(0))

    # ---- capital-efficiency framing (the ΔNPV/ΔInvest fix) ------------------
    dnpv, dinv = wide["delta_npv"], wide["delta_invest"]
    wide["invest_saved"] = -dinv
    wide["capex_saved"] = -wide["delta_capex"]
    wide["value_created"] = dnpv
    wide["incr_cap_eff"] = np.where(dinv > 0, safe_div(dnpv, dinv), np.nan)
    wide["value_per_dollar_released"] = np.where(
        dinv < 0, safe_div(dnpv, -dinv), np.nan)
    wide["cap_productivity_idx"] = safe_div(wide["after_npv_inv"], wide["before_npv_inv"])
    wide["quadrant"] = np.select(
        [(dnpv >= 0) & (dinv <= 0), (dnpv >= 0) & (dinv > 0),
         (dnpv < 0) & (dinv <= 0), (dnpv < 0) & (dinv > 0)],
        ["Accretive · Capital Released", "Accretive · Capital Deployed",
         "Dilutive · Capital Released", "Dilutive · Capital Deployed"],
        default="n/a")
    for k in ("npv", "boe", "fy_rate", "invest", "capex", "ip30", "avg3"):
        wide[f"before_{k}_pw"] = safe_div(wide[f"before_{k}"], wide["wells_before"])
        wide[f"after_{k}_pw"] = safe_div(wide[f"after_{k}"], wide["wells_after"])
        wide[f"delta_{k}_pw"] = wide[f"after_{k}_pw"] - wide[f"before_{k}_pw"]
    return wide


@st.cache_data(show_spinner="Building event model…")
def build_events(_sig: tuple, scales: tuple, dup_policy: str, treat_zero: bool):
    ents, e_iss = build_entities(_sig, scales, dup_policy)
    mem, meta, m_iss = build_membership(_sig, scales, dup_policy)
    _, f_iss = build_forecast(_sig, scales)

    wide = wide_from_membership(mem, ents, ["event_id"],
                                meta["event_id"].tolist(), treat_zero)
    ev = meta.merge(wide.reset_index(), on="event_id", how="left")

    # a couple of value-level sanity checks
    iss = list(e_iss) + list(m_iss) + list(f_iss)
    for r in ev.itertuples():
        if r.case != "Creation" and pd.notna(r.delta_invest) and abs(r.delta_invest) < 1e-9 \
                and pd.notna(r.delta_npv) and abs(r.delta_npv) > 1e-6:
            iss.append(dict(severity="Low", category="Zero capital delta",
                            scope=r.event_id, ref=r.enh_ent,
                            detail="Investment unchanged but NPV moved — check that "
                                   "the enhanced case carries its own capital."))
        if pd.notna(r.after_boe) and r.after_boe <= 0:
            iss.append(dict(severity="Low", category="Non-positive reserves",
                            scope=r.event_id, ref=r.enh_ent,
                            detail="Enhanced well BOE ≤ 0; $/boe metrics suppressed."))
    issues = pd.DataFrame(iss, columns=["severity", "category", "scope", "ref", "detail"])
    if not issues.empty:
        order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
        issues["_o"] = issues["severity"].map(order).fillna(9)
        issues = (issues.sort_values(["_o", "category", "scope"])
                        .drop(columns="_o").reset_index(drop=True))
    return ev, mem, ents, issues


# =============================================================================
# SECTION 10 — FORECAST PANELS
# =============================================================================

VALS = ["revenue", "opinc", "cf"]


def build_panel(mem: pd.DataFrame, fc: pd.DataFrame,
                group: Sequence[str] | None = None) -> pd.DataFrame:
    """Monthly Before / After / Incremental panel with cumulative columns."""
    grp = list(group or [])
    m = mem.drop_duplicates(subset=grp + ["side", "entity_key"])
    j = m[grp + ["side", "entity_key"]].merge(fc, on="entity_key", how="inner")
    if j.empty:
        cols = grp + ["date"] + [f"{p}_{v}" for p in ("before", "after", "incr") for v in VALS]
        return pd.DataFrame(columns=cols)

    a = j.groupby(grp + ["side", "date"], as_index=False)[VALS].sum(min_count=1)
    p = a.pivot_table(index=grp + ["date"], columns="side", values=VALS,
                      aggfunc="sum", fill_value=0.0)
    p.columns = [f"{str(sd).lower()}_{v}" for v, sd in p.columns]
    for sd in ("before", "after"):
        for v in VALS:
            if f"{sd}_{v}" not in p:
                p[f"{sd}_{v}"] = 0.0
    for v in VALS:
        p[f"incr_{v}"] = p[f"after_{v}"] - p[f"before_{v}"]

    p = p.reset_index().sort_values(grp + ["date"])
    cum_src = [f"{pre}_{v}" for pre in ("before", "after", "incr") for v in VALS]
    if grp:
        p[[f"cum_{c}" for c in cum_src]] = p.groupby(grp)[cum_src].cumsum()
    else:
        p[[f"cum_{c}" for c in cum_src]] = p[cum_src].cumsum()
    p["year"] = p["date"].dt.year
    p["month"] = p["date"].dt.month
    return p


def zero_crossing(dates: Sequence, cum: Sequence) -> float | None:
    """Months (interpolated) until cumulative series turns positive → years."""
    cum = np.asarray(cum, dtype="float64")
    if cum.size == 0 or np.all(~np.isfinite(cum)):
        return None
    pos = np.where(cum > 0)[0]
    if pos.size == 0:
        return None
    i = int(pos[0])
    if i == 0:
        return round(1 / 12, 3)
    prev, cur = cum[i - 1], cum[i]
    frac = 0.0 if cur == prev else (0 - prev) / (cur - prev)
    return round((i - 1 + frac + 1) / 12.0, 3)


def irr_monthly(cf: Sequence, lo: float = -0.95, hi: float = 3.0) -> float | None:
    """Annualised IRR from a monthly cash-flow stream (bisection, no deps)."""
    c = np.asarray([x for x in cf if np.isfinite(x)], dtype="float64")
    if c.size < 2 or not (np.any(c > 0) and np.any(c < 0)):
        return None

    def npv(r):
        t = np.arange(c.size)
        return float(np.sum(c / (1.0 + r) ** t))

    f_lo, f_hi = npv(lo), npv(hi)
    if not np.isfinite(f_lo) or not np.isfinite(f_hi) or f_lo * f_hi > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        fm = npv(mid)
        if abs(fm) < 1e-9:
            break
        if f_lo * fm <= 0:
            hi, f_hi = mid, fm
        else:
            lo, f_lo = mid, fm
    r = (lo + hi) / 2
    return ((1 + r) ** 12 - 1) * 100.0


# =============================================================================
# SECTION 11 — UI PRIMITIVES
# =============================================================================

def header():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        f'<div class="app-header"><div class="t">{APP_TITLE}</div>'
        f'<div class="s">{APP_SUBTITLE}</div></div>', unsafe_allow_html=True)


def section(n: str, title: str, note: str = ""):
    st.markdown(
        f'<div class="sec"><span class="n">{n}</span><span class="h">{title}</span>'
        f'<span class="d">{note}</span></div>', unsafe_allow_html=True)


def callout(kind: str, html: str):
    st.markdown(f'<div class="callout {kind}">{html}</div>', unsafe_allow_html=True)


def kpi(label: str, value: str, delta: str | None = None,
        dir_: int = 0, sub: str | None = None, tone: str = ""):
    dcls = "up" if dir_ > 0 else ("dn" if dir_ < 0 else "nt")
    arrow = "▲ " if dir_ > 0 else ("▼ " if dir_ < 0 else "")
    d = f'<div class="d {dcls}">{arrow}{delta}</div>' if delta else ""
    s = f'<div class="s">{sub}</div>' if sub else ""
    st.markdown(f'<div class="kpi {tone}"><div class="l">{label}</div>'
                f'<div class="v">{value}</div>{d}{s}</div>', unsafe_allow_html=True)


def kpi_grid(cards: list[dict], ncols: int = 4):
    for i in range(0, len(cards), ncols):
        for col, c in zip(st.columns(ncols), cards[i:i + ncols]):
            with col:
                kpi(**c)


def metric_card(key: str, before, after, label: str | None = None) -> dict:
    """KPI card for a registry metric, with good/bad direction resolved."""
    m = M_BY_KEY[key]
    d = (after - before) if not (_bad(after) or _bad(before)) else np.nan
    dir_ = 0
    if not _bad(d) and abs(d) > 1e-12 and m.good:
        dir_ = int(np.sign(d) * m.good)
    return dict(label=label or m.label, value=fmt_metric(key, after),
                delta=f"{fmt_metric(key, d, signed=True)} vs before" if not _bad(d) else None,
                dir_=dir_, sub=f"Before {fmt_metric(key, before)}",
                tone="pos" if dir_ > 0 else ("neg" if dir_ < 0 else ""))


def dl(df: pd.DataFrame, fname: str, label: str, key: str):
    st.download_button(label, df.to_csv(index=False).encode("utf-8"),
                       file_name=fname, mime="text/csv", key=key)


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    engine = None
    for cand in ("xlsxwriter", "openpyxl"):
        try:
            __import__(cand)
            engine = cand
            break
        except ImportError:
            continue
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as xw:
        for name, df in sheets.items():
            (df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)) \
                .to_excel(xw, sheet_name=str(name)[:31], index=False)
    return buf.getvalue()


# =============================================================================
# SECTION 12 — CHART LIBRARY
# =============================================================================

def _fin(fig: go.Figure, title: str, h: int = 420, ytitle: str = "",
         tickpre: str = "", ticksuf: str = "") -> go.Figure:
    fig.update_layout(title=title, height=h, yaxis_title=ytitle,
                      bargap=0.28, hovermode="x unified")
    if tickpre or ticksuf:
        fig.update_yaxes(tickprefix=tickpre, ticksuffix=ticksuf)
    return fig


def chart_bridge(ev: pd.DataFrame) -> go.Figure:
    wb = float(ev["wells_before"].sum())
    per = {c: float(ev.loc[ev["case"] == c, "wells_net"].sum()) for c in CASE_ORDER}
    labels = ["Before inventory"] + [f"{c}<br>net" for c in CASE_ORDER if per[c] or c != "Unclassified"] \
             + ["After inventory"]
    vals = [wb] + [per[c] for c in CASE_ORDER if per[c] or c != "Unclassified"] \
           + [float(ev["wells_after"].sum())]
    meas = ["absolute"] + ["relative"] * (len(vals) - 2) + ["total"]
    fig = go.Figure(go.Waterfall(
        orientation="v", measure=meas, x=labels, y=vals,
        text=[f"{v:+,.0f}" if m == "relative" else f"{v:,.0f}"
              for v, m in zip(vals, meas)],
        textposition="outside",
        increasing=dict(marker_color=CLR["green"]),
        decreasing=dict(marker_color=CLR["red"]),
        totals=dict(marker_color=CLR["navy"]),
        connector=dict(line=dict(color=CLR["light"], dash="dot")),
    ))
    return _fin(fig, "Inventory bridge — well count", 400, "Wells")


def chart_waterfall(ev: pd.DataFrame, key: str, title: str, topn: int = 12) -> go.Figure:
    m = M_BY_KEY[key]
    d = ev[["event_id", "case", f"delta_{key}"]].dropna(subset=[f"delta_{key}"]).copy()
    d = d.rename(columns={f"delta_{key}": "d"})
    d = d.reindex(d["d"].abs().sort_values(ascending=False).index)
    head, tail = d.head(topn), d.iloc[topn:]

    labels = [f"Before ({fmt_metric(key, ev[f'before_{key}'].sum())})"]
    vals = [float(ev[f"before_{key}"].sum())]
    meas = ["absolute"]
    for r in head.itertuples():
        labels.append(f"{r.event_id}")
        vals.append(float(r.d))
        meas.append("relative")
    if len(tail):
        labels.append(f"Other ({len(tail)})")
        vals.append(float(tail["d"].sum()))
        meas.append("relative")
    labels.append("After")
    vals.append(float(ev[f"after_{key}"].sum()))
    meas.append("total")

    fig = go.Figure(go.Waterfall(
        orientation="v", measure=meas, x=labels, y=vals,
        text=[fmt_metric(key, v, signed=(m_ == "relative")) for v, m_ in zip(vals, meas)],
        textposition="outside", textfont=dict(size=10),
        increasing=dict(marker_color=CLR["green"] if m.good > 0 else CLR["red"]),
        decreasing=dict(marker_color=CLR["red"] if m.good > 0 else CLR["green"]),
        totals=dict(marker_color=CLR["navy"]),
        connector=dict(line=dict(color=CLR["light"], dash="dot")),
    ))
    fig.update_xaxes(tickangle=-40)
    return _fin(fig, title, 470, m.label)


def chart_monthly(panel: pd.DataFrame, val: str, title: str,
                  show: tuple[str, ...] = ("before", "after")) -> go.Figure:
    fig = go.Figure()
    style = {"before": (CLR["before"], "dash", "Before"),
             "after": (CLR["after"], "solid", "After"),
             "incr": (CLR["incr"], "solid", "Incremental")}
    for pre in show:
        c, dash, nm = style[pre]
        fig.add_trace(go.Scatter(
            x=panel["date"], y=panel[f"{pre}_{val}"], name=nm, mode="lines",
            line=dict(color=c, width=2.4, dash=dash),
            fill="tozeroy" if pre == "incr" else None,
            fillcolor="rgba(14,154,167,.14)" if pre == "incr" else None,
            hovertemplate="%{x|%b %Y}: %{y:$,.0f}<extra>" + nm + "</extra>"))
    fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
    return _fin(fig, title, 400, "$", tickpre="$")


def chart_incremental_bars(panel: pd.DataFrame, val: str, title: str) -> go.Figure:
    y = panel[f"incr_{val}"]
    fig = go.Figure(go.Bar(
        x=panel["date"], y=y,
        marker_color=np.where(y >= 0, CLR["green"], CLR["red"]),
        hovertemplate="%{x|%b %Y}: %{y:$,.0f}<extra>Incremental</extra>"))
    fig.add_trace(go.Scatter(x=panel["date"], y=panel[f"cum_incr_{val}"],
                             name="Cumulative", yaxis="y2", mode="lines",
                             line=dict(color=CLR["navy"], width=2.4)))
    fig.update_layout(yaxis2=dict(overlaying="y", side="right", showgrid=False,
                                  tickprefix="$", title="Cumulative"))
    return _fin(fig, title, 420, "$ / month", tickpre="$")


def chart_pareto(ev: pd.DataFrame, key: str = "npv", topn: int = 25) -> go.Figure:
    d = (ev[["event_id", "case", f"delta_{key}"]]
         .dropna().rename(columns={f"delta_{key}": "d"})
         .sort_values("d", ascending=False))
    pos = d[d["d"] > 0]
    tot = pos["d"].sum()
    d = d.head(topn)
    cum = d["d"].clip(lower=0).cumsum() / tot * 100 if tot else d["d"] * 0
    fig = go.Figure()
    fig.add_trace(go.Bar(x=d["event_id"], y=d["d"], name="Value created",
                         marker_color=[CASE_CLR.get(c, CLR["slate"]) for c in d["case"]],
                         hovertemplate="%{x}: %{y:$,.0f}<extra></extra>"))
    fig.add_trace(go.Scatter(x=d["event_id"], y=cum, name="Cumulative % of gross gain",
                             yaxis="y2", mode="lines+markers",
                             line=dict(color=CLR["amber"], width=2.2),
                             marker=dict(size=5)))
    fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 105],
                                  ticksuffix="%", showgrid=False))
    fig.update_xaxes(tickangle=-45, type="category")
    return _fin(fig, f"Pareto — where the value is created (top {topn})", 460,
                "Δ NPV", tickpre="$")


def chart_migration(ev: pd.DataFrame) -> go.Figure:
    """Investment vs NPV, with a segment per event showing before → after."""
    d = ev.dropna(subset=["before_invest", "before_npv", "after_invest", "after_npv"])
    fig = go.Figure()
    xs, ys = [], []
    for r in d.itertuples():
        xs += [r.before_invest, r.after_invest, None]
        ys += [r.before_npv, r.after_npv, None]
    if xs:
        fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines", showlegend=False,
                                 line=dict(color=CLR["light"], width=1.2),
                                 hoverinfo="skip"))
    fig.add_trace(go.Scatter(
        x=d["before_invest"], y=d["before_npv"], mode="markers", name="Before",
        marker=dict(size=9, color=CLR["before"], symbol="circle-open", line_width=2),
        text=d["event_id"],
        hovertemplate="<b>%{text}</b><br>Before · Inv %{x:$,.0f} · NPV %{y:$,.0f}<extra></extra>"))
    fig.add_trace(go.Scatter(
        x=d["after_invest"], y=d["after_npv"], mode="markers", name="After",
        marker=dict(size=11, color=[CASE_CLR.get(c, CLR["slate"]) for c in d["case"]],
                    line=dict(color="#FFF", width=1)),
        text=d["event_id"],
        hovertemplate="<b>%{text}</b><br>After · Inv %{x:$,.0f} · NPV %{y:$,.0f}<extra></extra>"))
    lim = float(np.nanmax([d["before_invest"].max(), d["after_invest"].max(), 1]))
    for mult, col in ((1.0, CLR["slate"]), (2.0, CLR["light"])):
        fig.add_trace(go.Scatter(x=[0, lim], y=[0, lim * mult], mode="lines",
                                 name=f"NPV/I = {mult:.0f}x", hoverinfo="skip",
                                 line=dict(color=col, dash="dot", width=1)))
    fig.update_xaxes(tickprefix="$")
    return _fin(fig, "Capital efficiency migration — investment vs NPV", 480,
                "NPV", tickpre="$")


def chart_bubble(ev: pd.DataFrame) -> go.Figure:
    d = ev.dropna(subset=["after_boe", "after_npv"]).copy()
    size = d["after_invest"].abs().fillna(0)
    ref = size.max() or 1
    fig = go.Figure()
    for case, g in d.groupby("case"):
        gs = g["after_invest"].abs().fillna(0)
        fig.add_trace(go.Scatter(
            x=g["after_boe"], y=g["after_npv"], mode="markers", name=case,
            marker=dict(size=gs, sizemode="area", sizeref=2 * ref / (44 ** 2), sizemin=5,
                        color=CASE_CLR.get(case, CLR["slate"]), opacity=.78,
                        line=dict(color="#FFF", width=1)),
            customdata=np.c_[g["event_id"], g["after_invest"], g["after_npv_inv"]],
            hovertemplate=("<b>%{customdata[0]}</b><br>Reserves %{x:,.0f} boe"
                           "<br>NPV %{y:$,.0f}<br>Investment %{customdata[1]:$,.0f}"
                           "<br>NPV/I %{customdata[2]:.2f}x<extra></extra>")))
    fig.update_xaxes(title="Enhanced well reserves (boe)")
    return _fin(fig, "Reserve quality — bubble area = investment", 470,
                "Enhanced well NPV", tickpre="$")


def chart_hist(ev: pd.DataFrame, key: str = "npv") -> go.Figure:
    d = ev[f"delta_{key}"].dropna()
    fig = go.Figure(go.Histogram(x=d, nbinsx=max(8, min(30, len(d) // 2 or 8)),
                                 marker=dict(color=CLR["blue"], line=dict(color="#FFF", width=1))))
    if len(d):
        fig.add_vline(x=0, line_color=CLR["slate"], line_width=1)
        fig.add_vline(x=float(d.median()), line_color=CLR["amber"], line_dash="dash",
                      annotation_text=f"median {fmt_metric(key, d.median())}",
                      annotation_position="top")
    fig.update_xaxes(tickprefix="$")
    return _fin(fig, "Distribution of value created per event", 380, "Events")


def chart_heatmap(panel: pd.DataFrame, col: str = "incr_cf",
                  title: str = "Incremental cash flow heatmap") -> go.Figure:
    if panel.empty:
        return _fin(go.Figure(), title, 360)
    p = panel.pivot_table(index="year", columns="month", values=col, aggfunc="sum")
    p = p.reindex(columns=range(1, 13))
    fig = go.Figure(go.Heatmap(
        z=p.values, x=[pd.Timestamp(2000, m, 1).strftime("%b") for m in p.columns],
        y=p.index.astype(str),
        colorscale=[[0, CLR["red"]], [0.5, "#F7F9FB"], [1, CLR["green"]]],
        zmid=0, colorbar=dict(title="$", tickprefix="$"),
        hovertemplate="%{y} %{x}: %{z:$,.0f}<extra></extra>"))
    fig.update_layout(yaxis=dict(autorange="reversed"))
    return _fin(fig, title, 40 + 34 * max(len(p.index), 4))


def chart_case_bars(cw: pd.DataFrame, key: str, title: str) -> go.Figure:
    m = M_BY_KEY[key]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=cw.index, y=cw[f"before_{key}"], name="Before",
                         marker_color=CLR["before"],
                         text=[fmt_metric(key, v) for v in cw[f"before_{key}"]],
                         textposition="outside", textfont_size=10))
    fig.add_trace(go.Bar(x=cw.index, y=cw[f"after_{key}"], name="After",
                         marker_color=CLR["after"],
                         text=[fmt_metric(key, v) for v in cw[f"after_{key}"]],
                         textposition="outside", textfont_size=10))
    fig.update_layout(barmode="group")
    return _fin(fig, title, 400, m.label)


def chart_quadrant(ev: pd.DataFrame) -> go.Figure:
    d = ev.dropna(subset=["delta_npv", "delta_invest"])
    fig = go.Figure()
    for q, col in (("Accretive · Capital Released", CLR["green"]),
                   ("Accretive · Capital Deployed", CLR["blue"]),
                   ("Dilutive · Capital Released", CLR["amber"]),
                   ("Dilutive · Capital Deployed", CLR["red"])):
        g = d[d["quadrant"] == q]
        if g.empty:
            continue
        fig.add_trace(go.Scatter(
            x=g["delta_invest"], y=g["delta_npv"], mode="markers", name=q,
            marker=dict(size=12, color=col, opacity=.85, line=dict(color="#FFF", width=1)),
            text=g["event_id"],
            hovertemplate=("<b>%{text}</b><br>Δ Investment %{x:$,.0f}"
                           "<br>Δ NPV %{y:$,.0f}<extra></extra>")))
    fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
    fig.add_vline(x=0, line_color=CLR["slate"], line_width=1)
    fig.add_annotation(text="BEST — more value, less capital", x=0.02, y=0.97,
                       xref="paper", yref="paper", showarrow=False,
                       font=dict(size=10, color=CLR["green"]), align="left")
    fig.add_annotation(text="WORST — less value, more capital", x=0.98, y=0.03,
                       xref="paper", yref="paper", showarrow=False,
                       font=dict(size=10, color=CLR["red"]), align="right")
    fig.update_xaxes(title="Δ Investment (negative = capital released)", tickprefix="$")
    return _fin(fig, "Capital quadrant — value vs capital intensity", 480,
                "Δ NPV", tickpre="$")


def chart_rolling(e: pd.DataFrame, col: str, title: str) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(x=e["date"], y=e[col], name="Monthly",
                         marker_color=CLR["light"]))
    for w, c in ((3, CLR["teal"]), (6, CLR["blue"]), (12, CLR["navy"])):
        fig.add_trace(go.Scatter(x=e["date"], y=e[col].rolling(w, min_periods=1).mean(),
                                 name=f"{w}-mo avg", mode="lines",
                                 line=dict(color=c, width=2)))
    return _fin(fig, title, 400, "$", tickpre="$")


# =============================================================================
# SECTION 13 — VIEW (filters applied)
# =============================================================================

@dataclass
class View:
    events: pd.DataFrame
    mem: pd.DataFrame
    ents: pd.DataFrame
    fc: pd.DataFrame
    issues: pd.DataFrame
    port: pd.Series          # portfolio Before/After/Delta wide row
    case_wide: pd.DataFrame
    panel: pd.DataFrame
    case_panel: pd.DataFrame
    excluded: int
    treat_zero: bool
    raw: Raw


def build_view(sig, scales, dup_policy, treat_zero, cases, only_complete,
               date_range) -> View:
    ev, mem, ents, issues = build_events(sig, scales, dup_policy, treat_zero)
    fc, _ = build_forecast(sig, scales)

    sel = ev[ev["case"].isin(cases)].copy()
    excluded = 0
    if only_complete:
        n0 = len(sel)
        sel = sel[sel["complete"].fillna(False)]
        excluded = n0 - len(sel)

    keep = set(sel["event_id"])
    m = mem[mem["event_id"].isin(keep)].copy()
    m = m.merge(sel[["event_id", "case"]].rename(columns={"case": "case_r"}),
                on="event_id", how="left")
    m["case"] = m["case_r"].fillna(m["case"])
    m = m.drop(columns="case_r")

    if date_range and len(fc):
        lo, hi = date_range
        fc = fc[(fc["date"] >= pd.Timestamp(lo)) & (fc["date"] <= pd.Timestamp(hi))]

    port = wide_from_membership(m, ents, [], ["Portfolio"], treat_zero).iloc[0]
    cases_present = [c for c in CASE_ORDER if c in set(sel["case"])]
    case_wide = wide_from_membership(m, ents, ["case"], cases_present, treat_zero)
    panel = build_panel(m, fc)
    case_panel = build_panel(m, fc, ["case"])
    return View(sel, m, ents, fc, issues, port, case_wide, panel, case_panel,
                excluded, treat_zero, load_raw(sig))


# =============================================================================
# SECTION 14 — PAGES
# =============================================================================

def page_exec(v: View):
    p, ev = v.port, v.events
    section("01", "Headline outcome",
            f"{len(ev):,} events in scope · {v.excluded} incomplete excluded")

    vc = p["delta_npv"]
    inv_saved = p["invest_saved"]
    cap_saved = p["capex_saved"]
    eff_b, eff_a = p["before_npv_inv"], p["after_npv_inv"]

    verdict = "good" if (not _bad(vc) and vc > 0) else ("bad" if not _bad(vc) and vc < 0 else "info")
    bits = [
        f"The redesign moves the inventory from <b>{fmt_int(p['wells_before'])}</b> wells to "
        f"<b>{fmt_int(p['wells_after'])}</b> (<b>{fmt_int(p['wells_net'], signed=True)}</b> net)",
        f"changes portfolio NPV by <b>{fmt_money(vc, signed=True)}</b>",
        f"and {'releases' if (not _bad(inv_saved) and inv_saved > 0) else 'consumes'} "
        f"<b>{fmt_money(abs(inv_saved) if not _bad(inv_saved) else np.nan)}</b> of investment",
    ]
    if not _bad(eff_a) and not _bad(eff_b):
        bits.append(f"lifting capital efficiency from <b>{fmt_ratio(eff_b)}</b> to "
                    f"<b>{fmt_ratio(eff_a)}</b> NPV per dollar invested")
    callout(verdict, " · ".join(bits) + ".")

    kpi_grid([
        dict(label="Events in scope", value=fmt_int(len(ev)),
             sub=" / ".join(f"{c[:5]} {int((ev['case'] == c).sum())}"
                            for c in CASE_ORDER if (ev["case"] == c).any()), tone="acc"),
        dict(label="Before wells", value=fmt_int(p["wells_before"])),
        dict(label="After wells", value=fmt_int(p["wells_after"])),
        dict(label="Net wells", value=fmt_int(p["wells_net"], signed=True),
             dir_=0, tone="warn" if (p["wells_net"] or 0) < 0 else "acc",
             sub="Negative = inventory consolidated"),
    ], 4)
    st.write("")
    kpi_grid([
        dict(label="Portfolio NPV — before", value=fmt_money(p["before_npv"])),
        dict(label="Portfolio NPV — after", value=fmt_money(p["after_npv"]), tone="acc"),
        dict(label="Value created", value=fmt_money(vc, signed=True),
             dir_=1 if (not _bad(vc) and vc > 0) else (-1 if not _bad(vc) and vc < 0 else 0),
             sub=f"{fmt_pct(safe_div(vc, abs(p['before_npv'])) * 100 if not _bad(p['before_npv']) else np.nan)} vs before",
             tone="pos" if (not _bad(vc) and vc > 0) else "neg"),
        dict(label="Capital efficiency (NPV/I)", value=fmt_ratio(eff_a),
             delta=f"{fmt_ratio(eff_a - eff_b, signed=True) if not (_bad(eff_a) or _bad(eff_b)) else DASH} vs before",
             dir_=1 if (not (_bad(eff_a) or _bad(eff_b)) and eff_a > eff_b) else -1,
             sub=f"Before {fmt_ratio(eff_b)}", tone="acc"),
    ], 4)
    st.write("")
    kpi_grid([
        dict(label="Investment — before / after",
             value=f"{fmt_money(p['before_invest'])} → {fmt_money(p['after_invest'])}"),
        dict(label="Investment released", value=fmt_money(inv_saved, signed=True),
             dir_=1 if (not _bad(inv_saved) and inv_saved > 0) else -1,
             tone="pos" if (not _bad(inv_saved) and inv_saved > 0) else "neg"),
        dict(label="Capex — before / after",
             value=f"{fmt_money(p['before_capex'])} → {fmt_money(p['after_capex'])}"),
        dict(label="Capex released", value=fmt_money(cap_saved, signed=True),
             dir_=1 if (not _bad(cap_saved) and cap_saved > 0) else -1,
             tone="pos" if (not _bad(cap_saved) and cap_saved > 0) else "neg"),
    ], 4)
    st.write("")
    fcf = v.panel["incr_cf"].sum() if len(v.panel) else np.nan
    kpi_grid([
        dict(label="Portfolio reserves — after", value=fmt_vol(p["after_boe"]),
             delta=f"{fmt_vol(p['delta_boe'], signed=True)}",
             dir_=1 if (not _bad(p["delta_boe"]) and p["delta_boe"] > 0) else -1),
        dict(label="1st-year rate — after", value=fmt_rate(p["after_fy_rate"]),
             delta=f"{fmt_rate(p['delta_fy_rate'], signed=True)}",
             dir_=1 if (not _bad(p["delta_fy_rate"]) and p["delta_fy_rate"] > 0) else -1),
        dict(label="Cost of reserves", value=fmt_usdboe(p["after_inv_boe"]),
             delta=f"{fmt_usdboe(p['delta_inv_boe'], signed=True)}",
             dir_=-1 if (not _bad(p["delta_inv_boe"]) and p["delta_inv_boe"] > 0) else 1,
             sub=f"Before {fmt_usdboe(p['before_inv_boe'])}"),
        dict(label="Forecast incremental cash flow", value=fmt_money(fcf, signed=True),
             dir_=1 if (not _bad(fcf) and fcf > 0) else -1,
             sub="Undiscounted, full forecast horizon", tone="acc"),
    ], 4)

    section("02", "Inventory and value bridges")
    c1, c2 = st.columns([1, 1.35])
    with c1:
        show_fig(chart_bridge(ev), "ex_bridge")
    with c2:
        show_fig(chart_waterfall(ev, "npv", "NPV waterfall — before to after"), "ex_npvwf")

    section("03", "Cash-flow trajectory", "Portfolio, undiscounted")
    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_monthly(v.panel, "cf", "Monthly cash flow — before vs after"), "ex_mcf")
    with c2:
        fig = go.Figure()
        for pre, col, nm, dash in (("before", CLR["before"], "Before", "dash"),
                                   ("after", CLR["after"], "After", "solid")):
            fig.add_trace(go.Scatter(x=v.panel["date"], y=v.panel[f"cum_{pre}_cf"],
                                     name=nm, mode="lines",
                                     line=dict(color=col, width=2.4, dash=dash)))
        fig.add_trace(go.Scatter(x=v.panel["date"], y=v.panel["cum_incr_cf"],
                                 name="Incremental", mode="lines",
                                 line=dict(color=CLR["incr"], width=2.4)))
        fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
        show_fig(_fin(fig, "Cumulative cash flow", 400, "$", tickpre="$"), "ex_ccf")

    section("04", "Where the value comes from")
    show_fig(chart_pareto(ev), "ex_pareto")

    top = ev.sort_values("delta_npv", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Top 10 value creators**")
        show_df(_event_display(top.head(10)))
    with c2:
        st.markdown("**Bottom 10 — review or reject**")
        show_df(_event_display(top.tail(10).iloc[::-1]))


def _event_display(ev: pd.DataFrame) -> pd.DataFrame:
    d = pd.DataFrame({
        "Event": ev["event_id"],
        "Case": ev["case"],
        "Wells": ev["wells_before"].map(lambda x: f"{x:,.0f}") + " → "
                 + ev["wells_after"].map(lambda x: f"{x:,.0f}"),
        "Before NPV": ev["before_npv"].map(fmt_money),
        "After NPV": ev["after_npv"].map(fmt_money),
        "Δ NPV": ev["delta_npv"].map(lambda x: fmt_money(x, signed=True)),
        "Δ Investment": ev["delta_invest"].map(lambda x: fmt_money(x, signed=True)),
        "NPV/I after": ev["after_npv_inv"].map(fmt_ratio),
        "Quadrant": ev["quadrant"],
        "Flag": np.where(ev["complete"].fillna(False), "", "⚠ incomplete"),
    })
    return d


def page_portfolio(v: View):
    p, ev = v.port, v.events
    section("01", "Portfolio economics — before, after, delta",
            "Ratios recomputed from aggregates, never summed")

    rows = []
    for m in METRICS:
        b, a = p[f"before_{m.key}"], p[f"after_{m.key}"]
        d = a - b if not (_bad(a) or _bad(b)) else np.nan
        pct = safe_div(d, abs(b)) * 100 if not (_bad(d) or _bad(b)) else np.nan
        verdict = DASH
        if not _bad(d) and m.good and abs(d) > 1e-12:
            verdict = "✅ better" if np.sign(d) * m.good > 0 else "❌ worse"
        rows.append({
            "Metric": m.label,
            "Aggregation": {"additive": "Sum", "weighted": f"Wtd by {M_BY_KEY[m.weight].label}"
                            if m.weight else "Weighted", "derived": "Recomputed"}[m.kind],
            "Before": fmt_metric(m.key, b),
            "After": fmt_metric(m.key, a),
            "Delta": fmt_metric(m.key, d, signed=True),
            "% Δ": fmt_pct(pct, signed=True) if not _bad(pct) else DASH,
            "Direction": verdict,
            "Note": m.note,
        })
    show_df(pd.DataFrame(rows), height=560)

    section("02", "Per-well normalisation",
            "A 2-well before case vs a 1-well after case is not comparable un-normalised")
    pw = []
    for k, lab in PER_WELL.items():
        b, a = p[f"before_{k}_pw"], p[f"after_{k}_pw"]
        d = a - b if not (_bad(a) or _bad(b)) else np.nan
        pw.append({"Metric": lab, "Before / well": fmt_metric(k, b),
                   "After / well": fmt_metric(k, a),
                   "Delta / well": fmt_metric(k, d, signed=True),
                   "Uplift": fmt_pct(safe_div(d, abs(b)) * 100, signed=True)
                   if not (_bad(d) or _bad(b)) else DASH})
    show_df(pd.DataFrame(pw))

    section("03", "Case contribution")
    cw = v.case_wide
    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_case_bars(cw, "npv", "NPV by case"), "pf_cnpv")
    with c2:
        show_fig(chart_case_bars(cw, "invest", "Investment by case"), "pf_cinv")

    section("04", "Quality of the after inventory")
    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_bubble(ev), "pf_bub")
    with c2:
        show_fig(chart_migration(ev), "pf_mig")

    section("05", "Dispersion of outcomes")
    c1, c2 = st.columns([1.1, 1])
    with c1:
        show_fig(chart_hist(ev, "npv"), "pf_hist")
    with c2:
        d = ev["delta_npv"].dropna()
        if len(d):
            stats = pd.DataFrame({
                "Statistic": ["Events valued", "Value accretive", "Value dilutive",
                              "Gross uplift", "Gross erosion", "Net value created",
                              "Mean per event", "Median per event",
                              "P10 (high)", "P90 (low)", "Best event", "Worst event"],
                "Value": [fmt_int(len(d)), fmt_int((d > 0).sum()), fmt_int((d < 0).sum()),
                          fmt_money(d[d > 0].sum()), fmt_money(d[d < 0].sum()),
                          fmt_money(d.sum(), signed=True), fmt_money(d.mean(), signed=True),
                          fmt_money(d.median(), signed=True),
                          fmt_money(d.quantile(.9)), fmt_money(d.quantile(.1)),
                          f"{ev.loc[d.idxmax(), 'event_id']} ({fmt_money(d.max())})",
                          f"{ev.loc[d.idxmin(), 'event_id']} ({fmt_money(d.min())})"],
            })
            show_df(stats, height=460)

    section("06", "Event detail")
    show_df(_event_display(ev.sort_values("delta_npv", ascending=False)), height=460)


def page_capital(v: View):
    p, ev = v.port, v.events
    section("01", "Capital position", "Investment = economics BTax @0%; Capex = capex.xlsx")

    kpi_grid([
        dict(label="Investment before", value=fmt_money(p["before_invest"])),
        dict(label="Investment after", value=fmt_money(p["after_invest"]), tone="acc"),
        dict(label="Investment saved", value=fmt_money(p["invest_saved"], signed=True),
             dir_=1 if (not _bad(p["invest_saved"]) and p["invest_saved"] > 0) else -1,
             tone="pos" if (not _bad(p["invest_saved"]) and p["invest_saved"] > 0) else "neg"),
        dict(label="Capex saved", value=fmt_money(p["capex_saved"], signed=True),
             dir_=1 if (not _bad(p["capex_saved"]) and p["capex_saved"] > 0) else -1),
        dict(label="NPV / Investment — before", value=fmt_ratio(p["before_npv_inv"])),
        dict(label="NPV / Investment — after", value=fmt_ratio(p["after_npv_inv"]), tone="acc"),
        dict(label="Capital productivity index",
             value=fmt_ratio(p["cap_productivity_idx"]),
             sub="After NPV/I ÷ Before NPV/I; >1.00 = more value per dollar"),
        dict(label="Investment per BOE", value=fmt_usdboe(p["after_inv_boe"]),
             delta=f"{fmt_usdboe(p['delta_inv_boe'], signed=True)}",
             dir_=-1 if (not _bad(p["delta_inv_boe"]) and p["delta_inv_boe"] > 0) else 1,
             sub=f"Before {fmt_usdboe(p['before_inv_boe'])}"),
    ], 4)

    callout("info",
            "<b>Why ΔNPV ÷ ΔInvestment is not the headline metric.</b> Most consolidations "
            "<i>release</i> capital, so ΔInvestment is negative and the ratio flips sign on "
            "your best outcomes. Incremental efficiency is therefore reported only where "
            "capital was deployed; where capital was released we report value per dollar "
            "released and the capital productivity index.")

    section("02", "Capital waterfalls")
    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_waterfall(ev, "invest", "Investment waterfall"), "cp_iwf")
    with c2:
        show_fig(chart_waterfall(ev, "capex", "Capex waterfall"), "cp_cwf")

    section("03", "Capital quadrant")
    show_fig(chart_quadrant(ev), "cp_quad")
    q = (ev.groupby("quadrant")
           .agg(Events=("event_id", "count"), dNPV=("delta_npv", "sum"),
                dInvest=("delta_invest", "sum"), Wells=("wells_net", "sum"))
           .reset_index())
    if not q.empty:
        q["dNPV"] = q["dNPV"].map(lambda x: fmt_money(x, signed=True))
        q["dInvest"] = q["dInvest"].map(lambda x: fmt_money(x, signed=True))
        q["Wells"] = q["Wells"].map(lambda x: fmt_int(x, signed=True))
        show_df(q.rename(columns={"quadrant": "Quadrant", "dNPV": "Δ NPV",
                                  "dInvest": "Δ Investment", "Wells": "Net wells"}))

    section("04", "Efficiency league table")
    t = ev.copy()
    disp = pd.DataFrame({
        "Event": t["event_id"], "Case": t["case"],
        "Δ NPV": t["delta_npv"].map(lambda x: fmt_money(x, signed=True)),
        "Δ Investment": t["delta_invest"].map(lambda x: fmt_money(x, signed=True)),
        "Investment saved": t["invest_saved"].map(lambda x: fmt_money(x, signed=True)),
        "Capex saved": t["capex_saved"].map(lambda x: fmt_money(x, signed=True)),
        "NPV/I before": t["before_npv_inv"].map(fmt_ratio),
        "NPV/I after": t["after_npv_inv"].map(fmt_ratio),
        "Cap. productivity": t["cap_productivity_idx"].map(fmt_ratio),
        "ΔNPV/ΔInv (deployed)": t["incr_cap_eff"].map(fmt_ratio),
        "$ value / $ released": t["value_per_dollar_released"].map(fmt_ratio),
        "$/boe after": t["after_inv_boe"].map(fmt_usdboe),
    })
    show_df(disp, height=480)

    section("05", "Reserve and production quality")
    c1, c2 = st.columns(2)
    with c1:
        rows = []
        for k in ("cor", "inv_boe", "npv_boe"):
            rows.append({"Metric": M_BY_KEY[k].label,
                         "Before": fmt_metric(k, p[f"before_{k}"]),
                         "After": fmt_metric(k, p[f"after_{k}"]),
                         "Delta": fmt_metric(k, p[f"delta_{k}"], signed=True)})
        st.markdown("**Cost of reserves comparison**")
        show_df(pd.DataFrame(rows))
        st.caption("Reported COR is the volume-weighted average from economics.xlsx; "
                   "Investment/BOE is recomputed from aggregates. Divergence usually "
                   "means the reported COR uses a different capital basis.")
    with c2:
        rows = []
        for k in ("fy_rate", "ip30", "avg3"):
            rows.append({"Metric": M_BY_KEY[k].label,
                         "Before total": fmt_metric(k, p[f"before_{k}"]),
                         "After total": fmt_metric(k, p[f"after_{k}"]),
                         "Before / well": fmt_metric(k, p[f"before_{k}_pw"]),
                         "After / well": fmt_metric(k, p[f"after_{k}_pw"])})
        st.markdown("**Production quality**")
        show_df(pd.DataFrame(rows))


def page_forecast(v: View):
    panel = v.panel
    section("01", "Forecast-implied returns",
            "Computed from monthly cash flow — independent of headline ratios")
    if panel.empty:
        st.warning("No forecast rows for the current selection.")
        return

    pb_b = zero_crossing(panel["date"], panel["cum_before_cf"])
    pb_a = zero_crossing(panel["date"], panel["cum_after_cf"])
    pb_i = zero_crossing(panel["date"], panel["cum_incr_cf"])
    irr_b = irr_monthly(panel["before_cf"])
    irr_a = irr_monthly(panel["after_cf"])
    irr_i = irr_monthly(panel["incr_cf"])

    kpi_grid([
        dict(label="Forecast payback — before", value=fmt_years(pb_b) if pb_b else "Not reached"),
        dict(label="Forecast payback — after", value=fmt_years(pb_a) if pb_a else "Not reached",
             tone="acc"),
        dict(label="Payback on the increment",
             value=fmt_years(pb_i) if pb_i else "Not reached"),
        dict(label="Headline payout (wtd)", value=fmt_years(v.port["after_payout"]),
             sub=f"Before {fmt_years(v.port['before_payout'])}"),
        dict(label="Forecast IRR — before", value=fmt_pct(irr_b) if irr_b else DASH),
        dict(label="Forecast IRR — after", value=fmt_pct(irr_a) if irr_a else DASH, tone="acc"),
        dict(label="IRR on the increment", value=fmt_pct(irr_i) if irr_i else DASH,
             tone="pos" if (irr_i or 0) > 0 else ""),
        dict(label="Headline ROR (wtd)", value=fmt_pct(v.port["after_ror"]),
             sub=f"Before {fmt_pct(v.port['before_ror'])}"),
    ], 4)
    callout("warn", "Forecast-implied payback and IRR assume <code>cash_flow</code> in "
                    "forecast.xlsx is <b>net of capital</b>. If it is operating cash flow only, "
                    "these will be optimistic — treat the headline payout/ROR as authoritative "
                    "and use these as a directional cross-check.")

    section("02", "Before vs after — monthly")
    tabs = st.tabs(["Cash flow", "Revenue", "Operating income"])
    for tab, val, lab in zip(tabs, VALS[::-1] if False else ["cf", "revenue", "opinc"],
                             ["Cash flow", "Revenue", "Operating income"]):
        with tab:
            c1, c2 = st.columns(2)
            with c1:
                show_fig(chart_monthly(panel, val, f"Monthly {lab.lower()}"), f"fc_m_{val}")
            with c2:
                show_fig(chart_incremental_bars(panel, val,
                                                f"Incremental {lab.lower()} + cumulative"),
                         f"fc_i_{val}")

    section("03", "Cumulative and seasonality")
    c1, c2 = st.columns([1.1, 1])
    with c1:
        fig = go.Figure()
        for pre, col, nm, dash in (("before", CLR["before"], "Before", "dash"),
                                   ("after", CLR["after"], "After", "solid"),
                                   ("incr", CLR["incr"], "Incremental", "solid")):
            fig.add_trace(go.Scatter(x=panel["date"], y=panel[f"cum_{pre}_cf"], name=nm,
                                     mode="lines", line=dict(color=col, width=2.4, dash=dash)))
        fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
        show_fig(_fin(fig, "Cumulative cash flow", 420, "$", tickpre="$"), "fc_cum")
    with c2:
        show_fig(chart_heatmap(panel, "incr_cf"), "fc_heat")

    section("04", "Annual summary")
    ann = panel.groupby("year").agg(
        before_revenue=("before_revenue", "sum"), after_revenue=("after_revenue", "sum"),
        before_opinc=("before_opinc", "sum"), after_opinc=("after_opinc", "sum"),
        before_cf=("before_cf", "sum"), after_cf=("after_cf", "sum"),
        incr_cf=("incr_cf", "sum")).reset_index()
    ann["cum_incr_cf"] = ann["incr_cf"].cumsum()
    disp = pd.DataFrame({
        "Year": ann["year"].astype(int),
        "Revenue before": ann["before_revenue"].map(fmt_money),
        "Revenue after": ann["after_revenue"].map(fmt_money),
        "Op income before": ann["before_opinc"].map(fmt_money),
        "Op income after": ann["after_opinc"].map(fmt_money),
        "Cash flow before": ann["before_cf"].map(fmt_money),
        "Cash flow after": ann["after_cf"].map(fmt_money),
        "Incremental CF": ann["incr_cf"].map(lambda x: fmt_money(x, signed=True)),
        "Cumulative incr.": ann["cum_incr_cf"].map(lambda x: fmt_money(x, signed=True)),
    })
    show_df(disp, height=420)

    section("05", "Revenue bridge")
    rev = (v.mem.drop_duplicates(["event_id", "side", "entity_key"])
             .merge(v.fc, on="entity_key", how="inner")
             .groupby(["event_id", "side"], as_index=False)["revenue"].sum())
    if not rev.empty:
        w = rev.pivot_table(index="event_id", columns="side", values="revenue",
                            aggfunc="sum", fill_value=0.0)
        for sd in ("Before", "After"):
            if sd not in w:
                w[sd] = 0.0
        w["d"] = w["After"] - w["Before"]
        tmp = pd.DataFrame({"event_id": w.index, "case": "",
                            "before_revenue": w["Before"], "after_revenue": w["After"],
                            "delta_revenue": w["d"]})
        labels = ["Before revenue"]
        vals = [float(w["Before"].sum())]
        meas = ["absolute"]
        top = tmp.reindex(tmp["delta_revenue"].abs().sort_values(ascending=False).index).head(12)
        for r in top.itertuples():
            labels.append(str(r.event_id)); vals.append(float(r.delta_revenue)); meas.append("relative")
        rest = tmp.loc[~tmp["event_id"].isin(top["event_id"]), "delta_revenue"].sum()
        if abs(rest) > 0:
            labels.append("Other"); vals.append(float(rest)); meas.append("relative")
        labels.append("After revenue"); vals.append(float(w["After"].sum())); meas.append("total")
        fig = go.Figure(go.Waterfall(
            orientation="v", measure=meas, x=labels, y=vals,
            text=[fmt_money(x) for x in vals], textposition="outside", textfont_size=10,
            increasing=dict(marker_color=CLR["green"]),
            decreasing=dict(marker_color=CLR["red"]),
            totals=dict(marker_color=CLR["navy"]),
            connector=dict(line=dict(color=CLR["light"], dash="dot"))))
        fig.update_xaxes(tickangle=-40)
        show_fig(_fin(fig, "Life-of-forecast revenue waterfall", 460, "Revenue", tickpre="$"),
                 "fc_revwf")


def page_cases(v: View):
    section("01", "Case archetypes", "Consolidation · Extension · Creation")
    cw = v.case_wide
    present = [c for c in CASE_ORDER if c in cw.index]
    if not present:
        st.info("No events in scope.")
        return

    summary = []
    for c in present:
        r = cw.loc[c]
        summary.append({
            "Case": c, "Events": fmt_int((v.events["case"] == c).sum()),
            "Wells before": fmt_int(r["wells_before"]), "Wells after": fmt_int(r["wells_after"]),
            "Net wells": fmt_int(r["wells_net"], signed=True),
            "NPV before": fmt_money(r["before_npv"]), "NPV after": fmt_money(r["after_npv"]),
            "Value created": fmt_money(r["delta_npv"], signed=True),
            "Investment saved": fmt_money(r["invest_saved"], signed=True),
            "NPV/I before": fmt_ratio(r["before_npv_inv"]),
            "NPV/I after": fmt_ratio(r["after_npv_inv"]),
        })
    show_df(pd.DataFrame(summary))

    tabs = st.tabs(present)
    for tab, case in zip(tabs, present):
        with tab:
            r = cw.loc[case]
            sub = v.events[v.events["case"] == case]
            st.markdown(f'<span class="pill" style="background:{CASE_CLR[case]}">{case}</span>'
                        f'<span style="color:#7A8A99;font-size:.85rem">{len(sub):,} events</span>',
                        unsafe_allow_html=True)
            kpi_grid([
                dict(label="Events", value=fmt_int(len(sub)), tone="acc"),
                dict(label="Wells before → after",
                     value=f"{fmt_int(r['wells_before'])} → {fmt_int(r['wells_after'])}",
                     sub=f"Net {fmt_int(r['wells_net'], signed=True)}"),
                metric_card("npv", r["before_npv"], r["after_npv"]),
                metric_card("invest", r["before_invest"], r["after_invest"]),
                metric_card("npv_inv", r["before_npv_inv"], r["after_npv_inv"]),
                metric_card("boe", r["before_boe"], r["after_boe"]),
                metric_card("fy_rate", r["before_fy_rate"], r["after_fy_rate"]),
                metric_card("inv_boe", r["before_inv_boe"], r["after_inv_boe"]),
            ], 4)

            st.markdown("**Economics comparison**")
            rows = []
            for m in METRICS:
                b, a = r[f"before_{m.key}"], r[f"after_{m.key}"]
                rows.append({"Metric": m.label, "Before": fmt_metric(m.key, b),
                             "After": fmt_metric(m.key, a),
                             "Delta": fmt_metric(m.key, (a - b) if not (_bad(a) or _bad(b))
                                                 else np.nan, signed=True)})
            c1, c2 = st.columns([1, 1.25])
            with c1:
                show_df(pd.DataFrame(rows), height=430)
            with c2:
                cp = v.case_panel
                cpp = cp[cp["case"] == case] if "case" in cp and len(cp) else cp.iloc[0:0]
                if len(cpp):
                    show_fig(chart_monthly(cpp, "cf", f"{case} — monthly cash flow"),
                             f"cs_cf_{case}")
                else:
                    st.info("No forecast coverage for this case.")

            if len(v.case_panel) and case in set(v.case_panel.get("case", [])):
                cpp = v.case_panel[v.case_panel["case"] == case]
                c1, c2 = st.columns(2)
                with c1:
                    show_fig(chart_incremental_bars(cpp, "cf", f"{case} — incremental CF"),
                             f"cs_icf_{case}")
                with c2:
                    show_fig(chart_heatmap(cpp, "incr_cf", f"{case} — incremental CF heatmap"),
                             f"cs_hm_{case}")

            st.markdown("**Top contributors**")
            show_df(_event_display(sub.sort_values("delta_npv", ascending=False).head(15)))
            dl(sub, f"case_{case.lower()}_events.csv", f"⬇ {case} events (CSV)",
               f"dl_case_{case}")


def page_event(v: View):
    section("01", "Event explorer")
    if v.events.empty:
        st.info("No events in scope.")
        return
    ev = v.events
    lbl = ev.apply(lambda r: f"{r['event_id']}  ·  {r['case']}  ·  "
                             f"{fmt_money(r['delta_npv'], signed=True)}", axis=1)
    pick = st.selectbox("Event", options=list(ev["event_id"]),
                        format_func=lambda k: lbl[ev.index[ev["event_id"] == k][0]])
    r = ev[ev["event_id"] == pick].iloc[0]

    st.markdown(
        f'<span class="pill" style="background:{CASE_CLR.get(r["case"], CLR["slate"])}">'
        f'{r["case"]}</span>'
        f'<span class="pill" style="background:{CLR["slate"]}">{r["quadrant"]}</span>'
        + ("" if r["complete"] else
           f'<span class="pill" style="background:{CLR["red"]}">INCOMPLETE JOIN</span>'),
        unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Before — well 1**")
        st.write(f"{r['w1_name'] or DASH}")
        st.caption(f"Entity: `{r['w1_ent'] or DASH}`")
    with c2:
        st.markdown("**Before — well 2**")
        st.write(f"{r['w2_name'] or DASH}")
        st.caption(f"Entity: `{r['w2_ent'] or DASH}`")
    with c3:
        st.markdown("**After — enhanced well**")
        st.write(f"{r['enh_ent'] or DASH}")
        st.caption(f"Wells: {fmt_int(r['wells_before'])} → {fmt_int(r['wells_after'])} "
                   f"({fmt_int(r['wells_net'], signed=True)})")

    kpi_grid([
        metric_card("npv", r["before_npv"], r["after_npv"]),
        metric_card("invest", r["before_invest"], r["after_invest"]),
        metric_card("capex", r["before_capex"], r["after_capex"]),
        metric_card("npv_inv", r["before_npv_inv"], r["after_npv_inv"]),
        metric_card("boe", r["before_boe"], r["after_boe"]),
        metric_card("fy_rate", r["before_fy_rate"], r["after_fy_rate"]),
        metric_card("payout", r["before_payout"], r["after_payout"]),
        metric_card("inv_boe", r["before_inv_boe"], r["after_inv_boe"]),
    ], 4)

    section("02", "Well-level build-up")
    slots = v.mem[v.mem["event_id"] == pick].merge(v.ents, on="entity_key", how="left")
    keys = ["npv", "invest", "capex", "boe", "fy_rate", "ip30", "avg3",
            "payout", "ror", "cor", "wi", "npv_inv"]
    tbl = pd.DataFrame({
        "Slot": slots["slot"], "Side": slots["side"], "Entity": slots["entity_raw"],
        "In econ": np.where(slots["in_econ"], "✅", "❌"),
        "In capex": np.where(slots["in_capex"], "✅", "❌"),
        "In forecast": np.where(slots["in_fc"], "✅", "❌"),
        **{M_BY_KEY[k].label: [fmt_metric(k, x) for x in slots[k]] for k in keys},
    })
    show_df(tbl)

    rows = []
    for m in METRICS:
        b, a = r[f"before_{m.key}"], r[f"after_{m.key}"]
        rows.append({"Metric": m.label, "Aggregation": m.kind.title(),
                     "Before": fmt_metric(m.key, b), "After": fmt_metric(m.key, a),
                     "Delta": fmt_metric(m.key, (a - b) if not (_bad(a) or _bad(b))
                                         else np.nan, signed=True)})
    show_df(pd.DataFrame(rows), height=430)

    section("03", "Event cash flow")
    ep = build_panel(v.mem[v.mem["event_id"] == pick], v.fc)
    if ep.empty:
        st.info("No forecast rows joined to this event.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            show_fig(chart_monthly(ep, "cf", "Monthly cash flow"), "ee_cf")
        with c2:
            show_fig(chart_incremental_bars(ep, "cf", "Incremental cash flow"), "ee_icf")
        c1, c2 = st.columns(2)
        with c1:
            show_fig(chart_monthly(ep, "revenue", "Monthly revenue"), "ee_rev")
        with c2:
            show_fig(chart_monthly(ep, "opinc", "Monthly operating income"), "ee_oi")
        pb = zero_crossing(ep["date"], ep["cum_incr_cf"])
        st.caption(f"Incremental payback from forecast: "
                   f"**{fmt_years(pb) if pb else 'not reached in forecast horizon'}**")


def page_entity(v: View):
    section("01", "Entity explorer")
    keys = sorted(set(v.ents["entity_key"]) | set(v.fc["entity_key"]))
    if not keys:
        st.info("No entities available.")
        return
    disp = dict(zip(v.ents["entity_key"], v.ents["entity"].fillna(v.ents["entity_key"])))
    pick = st.selectbox("Entity", keys, format_func=lambda k: disp.get(k, k))

    row = v.ents[v.ents["entity_key"] == pick]
    role = v.mem[v.mem["entity_key"] == pick]
    if not role.empty:
        st.markdown("**Role in the redesign**")
        show_df(pd.DataFrame({
            "Event": role["event_id"], "Case": role["case"], "Slot": role["slot"],
            "Side": role["side"], "As written": role["entity_raw"]}))
    else:
        st.caption("This entity is not referenced by any in-scope event.")

    if not row.empty:
        r = row.iloc[0]
        kpi_grid([
            dict(label=M_BY_KEY[k].label, value=fmt_metric(k, r[k]))
            for k in ("npv", "invest", "capex", "npv_inv", "boe", "fy_rate",
                      "ip30", "avg3", "payout", "ror", "cor", "wi")
        ], 4)
    else:
        st.warning("No economics row for this entity.")

    section("02", "Monthly forecast")
    e = v.fc[v.fc["entity_key"] == pick].sort_values("date").copy()
    if e.empty:
        st.info("No forecast rows for this entity.")
        return
    for c in VALS:
        e[f"cum_{c}"] = e[c].cumsum()

    kpi_grid([
        dict(label="Forecast months", value=fmt_int(len(e)),
             sub=f"{e['date'].min():%b %Y} → {e['date'].max():%b %Y}"),
        dict(label="Life revenue", value=fmt_money(e["revenue"].sum())),
        dict(label="Life operating income", value=fmt_money(e["opinc"].sum())),
        dict(label="Life cash flow", value=fmt_money(e["cf"].sum()), tone="acc"),
    ], 4)

    c1, c2 = st.columns(2)
    with c1:
        show_fig(chart_rolling(e, "cf", "Cash flow + rolling averages"), "en_cf")
        show_fig(chart_rolling(e, "revenue", "Revenue + rolling averages"), "en_rev")
    with c2:
        show_fig(chart_rolling(e, "opinc", "Operating income + rolling averages"), "en_oi")
        fig = go.Figure()
        for c, col, nm in (("cum_revenue", CLR["teal"], "Cum revenue"),
                           ("cum_opinc", CLR["blue"], "Cum operating income"),
                           ("cum_cf", CLR["navy"], "Cum cash flow")):
            fig.add_trace(go.Scatter(x=e["date"], y=e[c], name=nm, mode="lines",
                                     line=dict(width=2.3, color=col)))
        fig.add_hline(y=0, line_color=CLR["slate"], line_width=1)
        show_fig(_fin(fig, "Cumulative performance", 400, "$", tickpre="$"), "en_cum")

    with st.expander("Monthly detail table"):
        d = e[["date", "revenue", "opinc", "cf", "cum_cf"]].copy()
        d["date"] = d["date"].dt.strftime("%Y-%m")
        for c in ("revenue", "opinc", "cf", "cum_cf"):
            d[c] = d[c].map(fmt_money)
        show_df(d.rename(columns={"date": "Month", "revenue": "Revenue",
                                  "opinc": "Operating income", "cf": "Cash flow",
                                  "cum_cf": "Cumulative CF"}), height=420)


def page_quality(v: View):
    section("01", "Validation summary")
    iss = v.issues
    counts = {s: int((iss["severity"] == s).sum()) if not iss.empty else 0
              for s in ("Critical", "High", "Medium", "Low")}
    kpi_grid([
        dict(label="Critical", value=fmt_int(counts["Critical"]),
             tone="neg" if counts["Critical"] else "pos",
             sub="Blocks valuation of an event"),
        dict(label="High", value=fmt_int(counts["High"]),
             tone="warn" if counts["High"] else "pos", sub="Materially affects numbers"),
        dict(label="Medium", value=fmt_int(counts["Medium"]), sub="Partial coverage"),
        dict(label="Low", value=fmt_int(counts["Low"]), sub="Informational"),
    ], 4)

    if counts["Critical"] or counts["High"]:
        callout("bad", "Resolve <b>Critical</b> and <b>High</b> findings before circulating "
                       "these numbers. Events with unmatched populated entities are excluded "
                       "from portfolio rollups by default (see sidebar).")
    elif iss.empty:
        callout("good", "No validation findings. All populated entities joined cleanly across "
                        "economics, capex and forecast.")

    section("02", "Coverage")
    ev_all, mem_all, ents_all, _ = build_events(
        file_signature(), st.session_state["_scales"], st.session_state["_dup"],
        st.session_state["_tz"])
    cov = pd.DataFrame({
        "Check": ["Consolidation rows", "Distinct events", "Classified events",
                  "Unclassified events", "Distinct entities referenced",
                  "Referenced ∩ economics", "Referenced ∩ capex", "Referenced ∩ forecast",
                  "Economics rows", "Capex rows", "Forecast rows", "Forecast months"],
        "Value": [fmt_int(len(v.raw.consol)), fmt_int(ev_all["event_id"].nunique()),
                  fmt_int((ev_all["case"] != "Unclassified").sum()),
                  fmt_int((ev_all["case"] == "Unclassified").sum()),
                  fmt_int(mem_all["entity_key"].nunique()),
                  fmt_int(mem_all.loc[mem_all["in_econ"], "entity_key"].nunique()),
                  fmt_int(mem_all.loc[mem_all["in_capex"], "entity_key"].nunique()),
                  fmt_int(mem_all.loc[mem_all["in_fc"], "entity_key"].nunique()),
                  fmt_int(len(v.raw.econ)), fmt_int(len(v.raw.capex)),
                  fmt_int(len(v.raw.forecast)),
                  fmt_int(v.fc["date"].nunique() if len(v.fc) else 0)],
    })
    c1, c2 = st.columns([1, 1.2])
    with c1:
        show_df(cov, height=460)
    with c2:
        if not iss.empty:
            bycat = (iss.groupby(["category", "severity"]).size()
                        .rename("n").reset_index())
            fig = go.Figure()
            for sev, col in (("Critical", CLR["red"]), ("High", CLR["amber"]),
                             ("Medium", CLR["blue"]), ("Low", CLR["slate"])):
                g = bycat[bycat["severity"] == sev]
                if g.empty:
                    continue
                fig.add_trace(go.Bar(y=g["category"], x=g["n"], name=sev,
                                     orientation="h", marker_color=col))
            fig.update_layout(barmode="stack")
            show_fig(_fin(fig, "Findings by category", 460, ""), "dq_bar")

    section("03", "Findings detail")
    if iss.empty:
        st.success("Nothing to report.")
    else:
        sev = st.multiselect("Severity", ["Critical", "High", "Medium", "Low"],
                             default=["Critical", "High", "Medium"])
        cat = st.multiselect("Category", sorted(iss["category"].unique()))
        f = iss[iss["severity"].isin(sev)]
        if cat:
            f = f[f["category"].isin(cat)]
        show_df(f.rename(columns={"severity": "Severity", "category": "Category",
                                  "scope": "Event / scope", "ref": "Reference",
                                  "detail": "Detail"}), height=520)
        dl(f, "validation_findings.csv", "⬇ Findings (CSV)", "dl_iss")

    section("04", "Load log")
    with st.expander("Column resolution and file log", expanded=False):
        for n in v.raw.notes:
            st.markdown(f"- {n}")
        for name, miss in v.raw.missing.items():
            if miss:
                st.error(f"`{name}`: unresolved required column(s): {', '.join(miss)}")


def _exports(v: View) -> dict[str, pd.DataFrame]:
    p = v.port
    port_tbl = pd.DataFrame([
        dict(metric=m.label, key=m.key, aggregation=m.kind,
             before=p[f"before_{m.key}"], after=p[f"after_{m.key}"],
             delta=p[f"delta_{m.key}"]) for m in METRICS
    ] + [
        dict(metric="Wells", key="wells", aggregation="count",
             before=p["wells_before"], after=p["wells_after"], delta=p["wells_net"]),
    ])
    case_tbl = v.case_wide.reset_index() if len(v.case_wide) else pd.DataFrame()
    ent_ref = (v.mem.drop_duplicates(["entity_key", "side"])
                 .merge(v.ents, on="entity_key", how="left"))
    ent_life = (v.fc.groupby("entity_key")
                  .agg(months=("date", "count"), first=("date", "min"), last=("date", "max"),
                       life_revenue=("revenue", "sum"), life_opinc=("opinc", "sum"),
                       life_cf=("cf", "sum")).reset_index())
    ent_tbl = ent_ref.merge(ent_life, on="entity_key", how="left")
    return {
        "Events": v.events,
        "Portfolio Summary": port_tbl,
        "Case Summary": case_tbl,
        "Entity Summary": ent_tbl,
        "Forecast Monthly": v.panel,
        "Forecast Annual": (v.panel.groupby("year").sum(numeric_only=True).reset_index()
                            if len(v.panel) else pd.DataFrame()),
        "Membership": v.mem,
        "Validation": v.issues,
    }


def page_downloads(v: View):
    section("01", "Exports", "Raw values in base dollars — no display rounding")
    sheets = _exports(v)
    st.download_button(
        "⬇ Download full workbook (.xlsx, all sheets)",
        to_excel_bytes(sheets),
        file_name="inventory_redesign_dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_wb", type="primary")
    st.caption("One sheet per table below. Column names are canonical keys so the export "
               "is reload-safe.")

    section("02", "Individual tables")
    for i, (name, df) in enumerate(sheets.items()):
        with st.expander(f"{name} — {len(df):,} rows × {df.shape[1] if len(df) else 0} cols"):
            show_df(df.head(200), height=320)
            dl(df, f"{name.lower().replace(' ', '_')}.csv", f"⬇ {name} (CSV)", f"dlx_{i}")


PAGES: dict[str, Callable[[View], None]] = {
    "Executive Summary": page_exec,
    "Portfolio Economics": page_portfolio,
    "Capital Efficiency": page_capital,
    "Forecast Analysis": page_forecast,
    "Case Analysis": page_cases,
    "Event Explorer": page_event,
    "Entity Explorer": page_entity,
    "Data Quality": page_quality,
    "Downloads": page_downloads,
}


# =============================================================================
# SECTION 15 — SIDEBAR + MAIN
# =============================================================================

def sidebar(sig) -> dict:
    sb = st.sidebar
    sb.markdown(f"### {APP_ICON} Navigation")
    page = sb.radio("Page", list(PAGES), label_visibility="collapsed")

    sb.markdown("---")
    sb.markdown("### Scope")
    cases = sb.multiselect("Event cases", CASE_ORDER,
                           default=[c for c in CASE_ORDER if c != "Unclassified"])
    only_complete = sb.toggle(
        "Exclude events with unmatched entities", value=True,
        help="A populated source or enhanced entity that is missing from economics.xlsx "
             "makes the event unquantifiable. Excluding is the conservative default.")
    treat_zero = sb.toggle(
        "Treat unmatched entities as zero", value=False,
        help="Off (default): unmatched populated entities make the side NULL and flagged. "
             "On: they contribute zero — this understates the Before case. Use only for "
             "sensitivity checks.")

    fc, _ = build_forecast(sig, tuple(st.session_state["_scales"]))
    dr = None
    if len(fc):
        lo, hi = fc["date"].min().date(), fc["date"].max().date()
        if lo < hi:
            dr = sb.slider("Forecast window", min_value=lo, max_value=hi,
                           value=(lo, hi), format="MMM YYYY")

    sb.markdown("---")
    with sb.expander("⚙️ Units & joins", expanded=False):
        st.caption("Economics NPV/Investment are converted to base dollars. "
                   "Aries `M$` = thousands ⇒ 1000.")
        s_econ = st.number_input("economics.xlsx money × ", value=float(DEFAULT_SCALES["economics"]),
                                 step=1.0, format="%.4f")
        s_capex = st.number_input("capex.xlsx money × ", value=float(DEFAULT_SCALES["capex"]),
                                  step=1.0, format="%.4f")
        s_fc = st.number_input("forecast.xlsx money × ", value=float(DEFAULT_SCALES["forecast"]),
                               step=1.0, format="%.4f")
        dup = st.selectbox("Duplicate entities in economics",
                           ["first", "sum", "mean"], index=0,
                           help="'first' keeps the first row (safest). 'sum' assumes rows "
                                "are additive components of one entity.")

    sb.markdown("---")
    if sb.button("🔄 Reload workbooks", width="stretch" if "width" in
                 inspect.signature(st.button).parameters else None):
        st.cache_data.clear()
        st.rerun()
    for name, pth in st.session_state.get("_paths", {}).items():
        sb.caption(f"`{name}` ← {Path(pth).name}")

    return dict(page=page, cases=cases or CASE_ORDER, only_complete=only_complete,
                treat_zero=treat_zero, date_range=dr,
                scales=(s_econ, s_capex, s_fc), dup=dup)


def missing_files_screen(err: Exception):
    header()
    st.error(str(err))
    st.markdown("#### Expected inputs")
    st.markdown(
        "Place these four files next to `app.py` (or in `./data`). "
        "`.csv` is accepted as a fallback. Column names are matched "
        "case- and punctuation-insensitively, with fuzzy fallback.")
    for name, flds in SCHEMA.items():
        with st.expander(f"`{name}.xlsx`"):
            st.dataframe(pd.DataFrame([
                dict(Column=f.key, Expected=f.aliases[0],
                     Type={"txt": "text", "num": "number", "int": "integer"}[f.kind],
                     Required="yes" if f.required else "no") for f in flds]),
                hide_index=True)
    st.info("Run `python make_demo_data.py` to generate a synthetic set and see the "
            "dashboard working end-to-end.")


def main():
    st.session_state.setdefault("_scales", list(DEFAULT_SCALES.values()))
    st.session_state.setdefault("_dup", "first")
    st.session_state.setdefault("_tz", False)

    try:
        sig = file_signature()
        raw = load_raw(sig)
        st.session_state["_paths"] = raw.paths
    except Exception as e:                                   # noqa: BLE001
        missing_files_screen(e)
        return

    header()
    cfg = sidebar(sig)
    st.session_state["_scales"] = list(cfg["scales"])
    st.session_state["_dup"] = cfg["dup"]
    st.session_state["_tz"] = cfg["treat_zero"]

    hard = {k: m for k, m in raw.missing.items() if m}
    if hard:
        for k, m in hard.items():
            st.error(f"`{k}` is missing required column(s): **{', '.join(m)}**. "
                     "Downstream numbers will be null. See Data Quality → Load log.")

    try:
        v = build_view(sig, tuple(cfg["scales"]), cfg["dup"], cfg["treat_zero"],
                       cfg["cases"], cfg["only_complete"], cfg["date_range"])
    except Exception as e:                                   # noqa: BLE001
        st.error(f"Model build failed: {e}")
        with st.expander("Traceback"):
            st.exception(e)
        return

    if v.events.empty:
        st.warning("No events match the current scope. Widen the case filter or turn off "
                   "'Exclude events with unmatched entities'.")
        return
    if v.excluded:
        callout("warn", f"<b>{v.excluded}</b> event(s) excluded from all figures because a "
                        "populated entity did not join to economics.xlsx. See "
                        "<b>Data Quality</b> for the list.")
    if cfg["treat_zero"]:
        callout("warn", "<b>Sensitivity mode:</b> unmatched entities are being treated as "
                        "zero. The Before case is understated and value creation is "
                        "overstated. Do not present these numbers.")

    PAGES[cfg["page"]](v)

    st.markdown("---")
    st.caption(
        f"Inventory redesign model · {len(v.events):,} events · "
        f"{v.mem['entity_key'].nunique():,} entities · "
        f"{len(v.fc):,} forecast rows · money in base dollars "
        f"(economics ×{cfg['scales'][0]:g}, capex ×{cfg['scales'][1]:g}, "
        f"forecast ×{cfg['scales'][2]:g}) · ratios recomputed from aggregates.")


if __name__ == "__main__":
    main()