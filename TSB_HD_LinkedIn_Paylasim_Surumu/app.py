from __future__ import annotations

import json
import math
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from tsb_engine import (
    AMOUNT_METRICS,
    MAIN_BRANCHES,
    XlsxRaw,
    activity,
    append_update_log,
    archive_sources,
    bubble_sizes,
    classify_workbook,
    detect_period,
    get_record,
    history_to_json_bytes,
    load_history,
    merge_period,
    normalize_code,
    period_key,
    prepare_import,
    read_update_log,
    records_for_period_branch,
    rows_to_csv_bytes,
    safe_delta,
    safe_growth,
    safe_pp,
    save_history,
    sorted_periods,
)

APP_TITLE = "Sigorta Sektörü Kârlılık Analizi"
APP_VERSION = "v8.3 Final"
INFLATION_YOY_BY_PERIOD = {"2026H1": 0.3211}
INFLATION_NOTE = "Haziran 2026 TÜFE yıllık değişim: %32,11 (TÜİK)"
PERIOD_COLOR_PREVIOUS = "#0B6EBD"
PERIOD_COLOR_CURRENT = "#79BCE8"
NEUTRAL_SERIES_COLORS = ["#0B6EBD", "#79BCE8", "#7A8796"]
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
INITIAL_HISTORY = APP_DIR / "initial_history.json"
ACTIVE_HISTORY = DATA_DIR / "active_history.json"
UPDATE_LOG = DATA_DIR / "update_log.json"
IMPORT_ROOT = DATA_DIR / "imports"

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")

st.markdown(
    """
<style>
:root {--navy:#17365D; --blue:#1F4E78; --line:rgba(128,128,128,.24);}
.block-container {padding-top: 1.05rem; padding-bottom: 2.2rem; max-width: 1540px;}
[data-testid="stSidebar"] {border-right: 1px solid rgba(128,128,128,.18);}
[data-testid="stSidebar"] .block-container {padding-top: 1.05rem;}
[data-testid="stSidebar"] div[role="radiogroup"] > label {
    border: 1px solid rgba(128,128,128,.18); border-radius: 9px; padding: .42rem .55rem;
    margin-bottom: .28rem; background: rgba(128,128,128,.035);
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {border-color: rgba(31,78,120,.48);}
.app-kicker {font-size:.72rem; text-transform:uppercase; letter-spacing:.11em; opacity:.62; margin-bottom:.2rem;}
.app-title {font-size:1.45rem; font-weight:720; letter-spacing:-.025em; line-height:1.15; margin-bottom:.25rem;}
.app-subtitle {font-size:.80rem; opacity:.66; line-height:1.45; margin-bottom:.8rem;}
.kpi-card {border: 1px solid var(--line); border-radius: 12px; padding: 14px 16px; min-height: 116px; background: rgba(255,255,255,.018); box-shadow: 0 1px 2px rgba(0,0,0,.025);}
.kpi-title {font-size: .80rem; opacity: .70; margin-bottom: 9px; line-height:1.3;}
.kpi-value {font-size: 1.45rem; font-weight: 675; letter-spacing: -.025em;}
.kpi-sub {font-size: .77rem; opacity: .69; margin-top: 7px; line-height: 1.4;}
.net-ref-card {border:1px solid rgba(31,78,120,.24); border-radius:12px; padding:16px 18px; min-height:138px; background:rgba(31,78,120,.025);}
.net-ref-card.emphasis {border:1.5px solid rgba(31,78,120,.48); background:rgba(31,78,120,.055);}
.net-ref-title {font-size:.86rem; color:#4b5563; font-weight:600; margin-bottom:8px;}
.net-ref-value {font-size:1.72rem; color:#182235; font-weight:750; letter-spacing:-.03em; margin-bottom:9px;}
.net-ref-period {font-size:.84rem; color:#465365; line-height:1.5;}
.net-ref-delta {font-size:.92rem; color:#182235; font-weight:700; margin-top:5px;}
.net-ref-note {font-size:.82rem; color:#566273; line-height:1.5; margin-top:8px;}
.note-box {border-left: 4px solid rgba(31,78,120,.58); padding: 9px 12px; background: rgba(31,78,120,.05); border-radius: 4px;}
.neutral-box {border:1px solid rgba(31,78,120,.20); padding:10px 13px; border-radius:9px; background:rgba(31,78,120,.035); font-size:.80rem; opacity:.82; line-height:1.5;}
.small-muted {font-size:.78rem; opacity:.68;}
.section-caption {font-size:.81rem; opacity:.68; margin-top:-.35rem; margin-bottom:.85rem;}
.module-chip {display:inline-block; border:1px solid rgba(31,78,120,.22); border-radius:999px; padding:4px 9px; margin-right:5px; font-size:.72rem; opacity:.78;}
.status-ok {border:1px solid rgba(46,160,67,.35); border-radius:12px; padding:12px 14px; min-height:92px; background:rgba(46,160,67,.06);}
.status-bad {border:1px solid rgba(220,53,69,.35); border-radius:12px; padding:12px 14px; min-height:92px; background:rgba(220,53,69,.06);}
.status-title {font-size:.78rem; opacity:.72; margin-bottom:6px;}
.status-value {font-size:1.05rem; font-weight:650;}
[data-testid="stDataFrame"] {border-radius: 10px; overflow: hidden; border:1px solid rgba(128,128,128,.13);}
span[data-baseweb="tag"] {background-color: rgba(31,78,120,.14) !important; color: #17365D !important;}
span[data-baseweb="tag"] svg {fill: #17365D !important;}
[data-testid="stMetricDelta"] {color: inherit !important;}
div[data-testid="stPlotlyChart"] {border:1px solid rgba(128,128,128,.10); border-radius:12px; padding:4px 4px 0 4px;}
</style>
""",
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def cached_history() -> tuple[list[dict[str, Any]], str]:
    return load_history(INITIAL_HISTORY, ACTIVE_HISTORY)


@st.cache_data(show_spinner=False)
def cached_prepare_import(file_payload: tuple[tuple[str, bytes], ...]) -> dict[str, Any]:
    return prepare_import(file_payload, tolerance_tl=1.0)


def num(value: Any, default: float = 0.0) -> float:
    try:
        if value in (None, ""):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def tr_number(value: float, digits: int = 1) -> str:
    text = f"{value:,.{digits}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_tl(value: Any, compact: bool = True, signed: bool = False) -> str:
    x = num(value)
    prefix = "+" if signed and x > 0 else ""
    ax = abs(x)
    if compact:
        if ax >= 1_000_000_000:
            return f"{prefix}{tr_number(x / 1_000_000_000, 1)} mlr TL"
        if ax >= 1_000_000:
            return f"{prefix}{tr_number(x / 1_000_000, 1)} mn TL"
    return f"{prefix}{tr_number(x, 0)} TL"


def compact_money_table(
    rows: list[dict[str, Any]],
    *,
    money_cols: list[str] | None = None,
    signed_money_cols: list[str] | None = None,
    pct_cols: list[str] | None = None,
    signed_pct_cols: list[str] | None = None,
    point_cols: list[str] | None = None,
) -> Any:
    """Format table cells for display without converting numeric columns to text.

    Keeping the underlying dataframe numeric preserves correct interactive sorting.
    """
    df = pd.DataFrame(rows)
    formatters: dict[str, Any] = {}
    for col in money_cols or []:
        if col in df.columns:
            formatters[col] = lambda v: "—" if pd.isna(v) else fmt_tl(v)
    for col in signed_money_cols or []:
        if col in df.columns:
            formatters[col] = lambda v: "—" if pd.isna(v) else fmt_tl(v, signed=True)
    for col in pct_cols or []:
        if col in df.columns:
            formatters[col] = lambda v: "—" if pd.isna(v) else f"%{tr_number(float(v), 2)}"
    for col in signed_pct_cols or []:
        if col in df.columns:
            formatters[col] = lambda v: "—" if pd.isna(v) else f"{'+' if float(v) > 0 else ''}%{tr_number(float(v), 2)}"
    for col in point_cols or []:
        if col in df.columns:
            formatters[col] = lambda v: "—" if pd.isna(v) else f"{'+' if float(v) > 0 else ''}{tr_number(float(v), 2)} puan"
    return df.style.format(formatters, na_rep="—")


def fmt_pct(value: Any, digits: int = 1, signed: bool = False) -> str:
    x = num(value) * 100.0
    prefix = "+" if signed and x > 0 else ""
    return f"{prefix}%{tr_number(x, digits)}"


def fmt_growth(value: float | None) -> str:
    return "—" if value is None else fmt_pct(value, signed=True)


def fmt_pp(value: Any, digits: int = 1) -> str:
    x = num(value)
    prefix = "+" if x > 0 else ""
    return f"{prefix}{tr_number(x, digits)} puan"




def safe_ratio(numerator: Any, denominator: Any) -> float | None:
    try:
        n = float(numerator or 0.0)
        d = float(denominator or 0.0)
    except (TypeError, ValueError):
        return None
    if abs(d) < 1e-12:
        return None
    return n / d


def safe_real_growth(current_value: Any, previous_value: Any, current_period: str) -> float | None:
    nominal = safe_growth(current_value, previous_value)
    inflation = INFLATION_YOY_BY_PERIOD.get(str(current_period))
    if nominal is None or inflation is None:
        return None
    return (1.0 + nominal) / (1.0 + inflation) - 1.0


def market_share(row: dict[str, Any] | None, sector_row: dict[str, Any] | None) -> float | None:
    if not row or not sector_row:
        return None
    return safe_ratio(row.get("Brüt Yazılan Prim (TL)"), sector_row.get("Brüt Yazılan Prim (TL)"))


def tech_to_gwp(row: dict[str, Any] | None, include_transfer: bool) -> float | None:
    if not row:
        return None
    key = "Teknik Kâr/Zarar (TL)" if include_transfer else "Yatırım Hariç Teknik Sonuç (TL)"
    return safe_ratio(row.get(key), row.get("Brüt Yazılan Prim (TL)"))


def net_combined_components(row: dict[str, Any] | None) -> tuple[float | None, float | None, float | None]:
    """Return Net H/P (DERK dahil), expense/net earned, and their sum.

    These fields are sourced from the TSB Hasar-Prim workbook and the 614 expense
    value already reconciled in the app. They are currently surfaced only at the
    HD sector summary level.
    """
    if not row:
        return None, None, None
    net_hp = row.get("Net H/P (DERK Dahil)")
    net_exp = row.get("Net Masraf Oranı")
    net_combined = row.get("Net Bileşik Oran")
    return (None if net_hp is None else float(net_hp),
            None if net_exp is None else float(net_exp),
            None if net_combined is None else float(net_combined))


def fmt_optional_pct(value: float | None, signed: bool = False, digits: int = 1) -> str:
    return "—" if value is None else fmt_pct(value, digits=digits, signed=signed)

def pct_for_plot(value: float | None) -> float | None:
    """Convert 0-1 ratio to percentage points without turning undefined values into zero."""
    if value is None:
        return None
    try:
        x = float(value)
    except (TypeError, ValueError):
        return None
    return x * 100.0 if math.isfinite(x) else None

def _quantile(values: list[float], q: float) -> float:
    ordered = sorted(values)
    if not ordered:
        return 0.0
    if len(ordered) == 1:
        return ordered[0]
    pos = (len(ordered) - 1) * q
    low = int(pos)
    high = min(low + 1, len(ordered) - 1)
    weight = pos - low
    return ordered[low] * (1.0 - weight) + ordered[high] * weight

def robust_axis_range(values: list[float], reference: float | None = None, min_span: float = 20.0) -> tuple[float, float, int]:
    """Return an outlier-resistant focus range while keeping every raw point in the figure.

    The default branch-distribution view focuses the axis only; no source row is
    deleted. Users can always switch to 'Tüm değerleri göster'.
    """
    vals = [float(v) for v in values if v is not None and math.isfinite(float(v))]
    if not vals:
        base = 0.0 if reference is None else float(reference)
        return base - min_span / 2, base + min_span / 2, 0
    if len(vals) == 1:
        center = vals[0]
        span = min_span
        low, high = center - span / 2, center + span / 2
    elif len(vals) < 8:
        # Small samples: median absolute deviation is more stable than percentiles
        # when one company has a tiny denominator and a very large ratio.
        median = _quantile(vals, 0.5)
        deviations = [abs(v - median) for v in vals]
        mad = _quantile(deviations, 0.5)
        half = max(min_span / 2, 4.0 * mad)
        low, high = median - half, median + half
    else:
        q1, q3 = _quantile(vals, 0.25), _quantile(vals, 0.75)
        iqr = q3 - q1
        if abs(iqr) < 1e-12:
            median = _quantile(vals, 0.5)
            deviations = [abs(v - median) for v in vals]
            mad = _quantile(deviations, 0.5)
            half = max(min_span / 2, 4.0 * mad)
            low, high = median - half, median + half
        else:
            fence_low, fence_high = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            central = [v for v in vals if fence_low <= v <= fence_high]
            if not central:
                central = vals
            raw_low, raw_high = min(central), max(central)
            span = max(min_span, raw_high - raw_low)
            low, high = raw_low - span * 0.12, raw_high + span * 0.12
    if reference is not None and math.isfinite(float(reference)):
        ref = float(reference)
        low = min(low, ref - min_span * 0.10)
        high = max(high, ref + min_span * 0.10)
    if high <= low:
        high = low + min_span
    outliers = sum(1 for v in vals if v < low or v > high)
    return low, high, outliers

def net_ref_card(title: str, current_value: float | None, previous_value: float | None, previous: str, current: str, emphasis: bool = False) -> None:
    delta = None if current_value is None or previous_value is None else (current_value - previous_value) * 100
    cls = "net-ref-card emphasis" if emphasis else "net-ref-card"
    delta_html = "" if delta is None else f"<div class='net-ref-delta'>Değişim: {fmt_pp(delta)}</div>"
    st.markdown(
        f"<div class='{cls}'><div class='net-ref-title'>{title}</div>"
        f"<div class='net-ref-value'>{fmt_optional_pct(current_value)}</div>"
        f"<div class='net-ref-period'><b>{previous}</b>: {fmt_optional_pct(previous_value)} &nbsp;→&nbsp; <b>{current}</b>: {fmt_optional_pct(current_value)}</div>"
        f"{delta_html}</div>",
        unsafe_allow_html=True,
    )

def kpi(title: str, value: str, sub: str) -> None:
    st.markdown(
        f"<div class='kpi-card'><div class='kpi-title'>{title}</div>"
        f"<div class='kpi-value'>{value}</div><div class='kpi-sub'>{sub}</div></div>",
        unsafe_allow_html=True,
    )


def status_box(title: str, passed: bool, value: str, detail: str = "") -> None:
    cls = "status-ok" if passed else "status-bad"
    icon = "✓" if passed else "!"
    st.markdown(
        f"<div class='{cls}'><div class='status-title'>{title}</div>"
        f"<div class='status-value'>{icon} {value}</div>"
        f"<div class='small-muted'>{detail}</div></div>",
        unsafe_allow_html=True,
    )


def period_pair_rows(history: list[dict[str, Any]], previous: str, current: str, branch: str) -> list[dict[str, Any]]:
    previous_rows = {
        normalize_code(r["Şirket Kodu"]): r
        for r in records_for_period_branch(history, previous, branch, include_sector=False)
        if r.get("Şirket Tipi") == "HD"
    }
    current_rows = {
        normalize_code(r["Şirket Kodu"]): r
        for r in records_for_period_branch(history, current, branch, include_sector=False)
        if r.get("Şirket Tipi") == "HD"
    }
    codes = sorted(set(previous_rows) | set(current_rows), key=lambda x: int(x) if x.isdigit() else x)
    return [
        {"code": code, "previous": previous_rows.get(code), "current": current_rows.get(code)}
        for code in codes
    ]


def neutral_sector_summary(previous_row: dict[str, Any], current_row: dict[str, Any], previous: str, current: str) -> str:
    return (
        f"{current} brüt yazılan prim {fmt_tl(current_row['Brüt Yazılan Prim (TL)'])}; "
        f"{previous} dönemine göre değişim {fmt_growth(safe_growth(current_row['Brüt Yazılan Prim (TL)'], previous_row['Brüt Yazılan Prim (TL)']))}. "
        f"Brüt kazanılmış prim {fmt_tl(current_row['Brüt Kazanılmış Prim (TL)'])}; "
        f"brüt H/P (DERK dahil) {fmt_pct(current_row['Brüt H/P'])} ({fmt_pp(safe_pp(current_row['Brüt H/P'], previous_row['Brüt H/P']))}); "
        f"masraf oranı {fmt_pct(current_row['Masraf Oranı'])} ({fmt_pp(safe_pp(current_row['Masraf Oranı'], previous_row['Masraf Oranı']))}); "
        f"brüt bileşik oran {fmt_pct(current_row['Brüt Bileşik Oran'])} ({fmt_pp(safe_pp(current_row['Brüt Bileşik Oran'], previous_row['Brüt Bileşik Oran']))})."
    )


def sector_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    prev_total = get_record(history, previous, 9000, "HAYATDISI")
    curr_total = get_record(history, current, 9000, "HAYATDISI")
    if not prev_total or not curr_total:
        st.error("Seçili dönemlerden biri için HAYATDIŞI / 9000 T (HD) sektör toplamı bulunamadı.")
        return

    nominal_growth = safe_growth(curr_total["Brüt Yazılan Prim (TL)"], prev_total["Brüt Yazılan Prim (TL)"])
    real_growth = safe_real_growth(curr_total["Brüt Yazılan Prim (TL)"], prev_total["Brüt Yazılan Prim (TL)"], current)
    ex_ratio_prev = tech_to_gwp(prev_total, False)
    ex_ratio_curr = tech_to_gwp(curr_total, False)
    in_ratio_prev = tech_to_gwp(prev_total, True)
    in_ratio_curr = tech_to_gwp(curr_total, True)

    st.title("Sektör Özeti")
    st.caption(f"Hayat dışı şirketler • {previous} → {current} • 18 ana branş metodolojisi")

    st.markdown("### Brüt Teknik Görünüm")
    st.markdown("<div class='section-caption'>Prim, hasar ve faaliyet gideri temelli brüt teknik oranlar. Teknik sonuçlar bu bloktan ayrı gösterilir.</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    with cols[0]:
        kpi("Brüt Yazılan Prim", fmt_tl(curr_total["Brüt Yazılan Prim (TL)"]), f"{previous}: {fmt_tl(prev_total['Brüt Yazılan Prim (TL)'])}")
    with cols[1]:
        kpi("Brüt Kazanılmış Prim", fmt_tl(curr_total["Brüt Kazanılmış Prim (TL)"]), f"{previous}: {fmt_tl(prev_total['Brüt Kazanılmış Prim (TL)'])}")
    with cols[2]:
        kpi("Nominal Prim Büyümesi", fmt_growth(nominal_growth), f"{previous} → {current}")
    with cols[3]:
        kpi("Reel Prim Büyümesi", fmt_growth(real_growth), INFLATION_NOTE if current in INFLATION_YOY_BY_PERIOD else "Enflasyon girdisi tanımlı değil")

    cols = st.columns(3)
    with cols[0]:
        kpi("Brüt H/P (DERK Dahil)", fmt_pct(curr_total["Brüt H/P"]), f"{previous}: {fmt_pct(prev_total['Brüt H/P'])} · Δ {fmt_pp(safe_pp(curr_total['Brüt H/P'], prev_total['Brüt H/P']))}")
    with cols[1]:
        kpi("Masraf Oranı", fmt_pct(curr_total["Masraf Oranı"]), f"{previous}: {fmt_pct(prev_total['Masraf Oranı'])} · Δ {fmt_pp(safe_pp(curr_total['Masraf Oranı'], prev_total['Masraf Oranı']))}")
    with cols[2]:
        kpi("Brüt Bileşik Oran", fmt_pct(curr_total["Brüt Bileşik Oran"]), f"{previous}: {fmt_pct(prev_total['Brüt Bileşik Oran'])} · Δ {fmt_pp(safe_pp(curr_total['Brüt Bileşik Oran'], prev_total['Brüt Bileşik Oran']))}")

    left, right = st.columns(2)
    with left:
        fig = go.Figure()
        fig.add_bar(name=previous, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(prev_total["Brüt Yazılan Prim (TL)"]) / 1e9, num(prev_total["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.add_bar(name=current, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(curr_total["Brüt Yazılan Prim (TL)"]) / 1e9, num(curr_total["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.update_layout(barmode="group", title="Brüt prim görünümü", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = go.Figure()
        fig.add_bar(name=previous, x=["H/P (DERK Dahil)", "Masraf Oranı", "Brüt Bileşik"], y=[num(prev_total["Brüt H/P"]) * 100, num(prev_total["Masraf Oranı"]) * 100, num(prev_total["Brüt Bileşik Oran"]) * 100])
        fig.add_bar(name=current, x=["H/P (DERK Dahil)", "Masraf Oranı", "Brüt Bileşik"], y=[num(curr_total["Brüt H/P"]) * 100, num(curr_total["Masraf Oranı"]) * 100, num(curr_total["Brüt Bileşik Oran"]) * 100])
        fig.update_layout(barmode="group", title="Brüt teknik oranlar (%)", yaxis_title="%", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    net_hp_prev, net_exp_prev, net_comb_prev = net_combined_components(prev_total)
    net_hp_curr, net_exp_curr, net_comb_curr = net_combined_components(curr_total)
    if net_comb_curr is not None:
        st.markdown("#### Net Bileşik Referansı")
        st.markdown("<div class='section-caption'>Net H/P ve net kazanılmış prim bazlı gider oranının sektör toplamındaki birlikte görünümü.</div>", unsafe_allow_html=True)
        net_cols = st.columns([1, 1, 1.12])
        with net_cols[0]:
            net_ref_card("Net H/P (DERK Dahil)", net_hp_curr, net_hp_prev, previous, current)
        with net_cols[1]:
            net_ref_card("Faaliyet Gideri / Net Kazanılmış Prim", net_exp_curr, net_exp_prev, previous, current)
        with net_cols[2]:
            net_ref_card("Net Bileşik Oran (DERK Dahil)", net_comb_curr, net_comb_prev, previous, current, emphasis=True)
        st.markdown("<div class='net-ref-note'><b>Formül:</b> Net Bileşik Oran = Net H/P (DERK Dahil) + Faaliyet Giderleri / Net Kazanılmış Prim. Bu gösterim yalnız sektör toplamında kullanılır.</div>", unsafe_allow_html=True)

    st.divider()
    st.markdown("### Teknik Sonuç Görünümü")
    st.markdown("<div class='section-caption'>Mali gelir aktarımı hariç/dahil teknik sonuçlar ve bu sonuçların brüt yazılan prime oranı birlikte gösterilir. Bu bölüm brüt H/P ve bileşik oran bloklarından ayrı okunur.</div>", unsafe_allow_html=True)
    cols = st.columns(5)
    specs = [
        ("Mali Gelir Aktarımı Hariç Teknik Sonuç", fmt_tl(curr_total["Yatırım Hariç Teknik Sonuç (TL)"]), f"Δ {fmt_tl(safe_delta(curr_total['Yatırım Hariç Teknik Sonuç (TL)'], prev_total['Yatırım Hariç Teknik Sonuç (TL)']), signed=True)}"),
        ("Mali Gelir Aktarımı (603)", fmt_tl(curr_total["Yatırım Katkısı (TL)"]), f"Δ {fmt_tl(safe_delta(curr_total['Yatırım Katkısı (TL)'], prev_total['Yatırım Katkısı (TL)']), signed=True)}"),
        ("Aktarım Dahil Teknik Sonuç", fmt_tl(curr_total["Teknik Kâr/Zarar (TL)"]), f"Δ {fmt_tl(safe_delta(curr_total['Teknik Kâr/Zarar (TL)'], prev_total['Teknik Kâr/Zarar (TL)']), signed=True)}"),
        ("Aktarım Hariç Teknik Sonuç / Brüt Prim", fmt_optional_pct(ex_ratio_curr), f"{previous}: {fmt_optional_pct(ex_ratio_prev)} · Δ {'—' if ex_ratio_prev is None or ex_ratio_curr is None else fmt_pp((ex_ratio_curr-ex_ratio_prev)*100)}"),
        ("Aktarım Dahil Teknik Sonuç / Brüt Prim", fmt_optional_pct(in_ratio_curr), f"{previous}: {fmt_optional_pct(in_ratio_prev)} · Δ {'—' if in_ratio_prev is None or in_ratio_curr is None else fmt_pp((in_ratio_curr-in_ratio_prev)*100)}"),
    ]
    for col, (title, value, sub) in zip(cols, specs):
        with col: kpi(title, value, sub)

    left, right = st.columns(2)
    with left:
        fig = go.Figure()
        result_x = ["Aktarım Hariç Teknik Sonuç", "Mali Gelir Aktarımı (603)", "Aktarım Dahil Teknik Sonuç"]
        fig.add_bar(name=previous, x=result_x, y=[num(prev_total["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(prev_total["Yatırım Katkısı (TL)"]) / 1e9, num(prev_total["Teknik Kâr/Zarar (TL)"]) / 1e9])
        fig.add_bar(name=current, x=result_x, y=[num(curr_total["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(curr_total["Yatırım Katkısı (TL)"]) / 1e9, num(curr_total["Teknik Kâr/Zarar (TL)"]) / 1e9])
        fig.add_hline(y=0, line_width=1, line_dash="dot")
        fig.update_layout(barmode="group", title="Teknik sonuç bileşenleri", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = go.Figure()
        fig.add_bar(name=previous, x=["Aktarım Hariç", "Aktarım Dahil"], y=[pct_for_plot(ex_ratio_prev), pct_for_plot(in_ratio_prev)])
        fig.add_bar(name=current, x=["Aktarım Hariç", "Aktarım Dahil"], y=[pct_for_plot(ex_ratio_curr), pct_for_plot(in_ratio_curr)])
        fig.add_hline(y=0, line_width=1, line_dash="dot")
        fig.update_layout(barmode="group", title="Teknik sonuç / brüt yazılan prim", yaxis_title="%", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown("### Net Kâr Görünümü")
    prev_net = prev_total.get("Dönem Net Kâr/Zarar (TL)")
    curr_net = curr_total.get("Dönem Net Kâr/Zarar (TL)")
    left, right = st.columns([1, 2])
    with left:
        kpi("Dönem Net Kârı / Zararı", fmt_tl(curr_net), f"{previous}: {fmt_tl(prev_net)} · Δ {fmt_tl(safe_delta(curr_net, prev_net), signed=True)}")
        st.markdown("<div class='section-caption'>Kaynak: Gelir Tablosu dosyası → MALI sayfası → 69 Dönem Net Karı Veya Zararı.</div>", unsafe_allow_html=True)
    with right:
        fig = go.Figure(go.Bar(x=[previous, current], y=[num(prev_net)/1e9, num(curr_net)/1e9]))
        fig.add_hline(y=0, line_width=1, line_dash="dot")
        fig.update_layout(title="HD sektör dönem net kârı / zararı", yaxis_title="Milyar TL", showlegend=False, margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown("### Şirket Bazında Toplam Pazar Payı")
    pairs = period_pair_rows(history, previous, current, "HAYATDISI")
    rows = []
    for pair in pairs:
        p, c = pair["previous"], pair["current"]
        if not c: continue
        prev_share = market_share(p, prev_total) if p else None
        curr_share = market_share(c, curr_total)
        ex = tech_to_gwp(c, False)
        inc = tech_to_gwp(c, True)
        rows.append({
            "Şirket": c.get("Şirket Adı", ""),
            "Şirket Kodu": pair["code"],
            f"Brüt Prim {current}": fmt_tl(c["Brüt Yazılan Prim (TL)"]),
            "Nominal Prim Büyümesi": "—" if not p else fmt_growth(safe_growth(c["Brüt Yazılan Prim (TL)"], p["Brüt Yazılan Prim (TL)"])),
            "Reel Prim Büyümesi": "—" if not p else fmt_growth(safe_real_growth(c["Brüt Yazılan Prim (TL)"], p["Brüt Yazılan Prim (TL)"], current)),
            f"Toplam Pazar Payı {previous}": fmt_optional_pct(prev_share),
            f"Toplam Pazar Payı {current}": fmt_optional_pct(curr_share),
            "Pazar Payı Δ": "—" if prev_share is None or curr_share is None else fmt_pp((curr_share-prev_share)*100),
            f"Net Kâr/Zarar {current}": fmt_tl(c.get("Dönem Net Kâr/Zarar (TL)")),
            "Net Kâr/Zarar Δ": "—" if not p else fmt_tl(safe_delta(c.get("Dönem Net Kâr/Zarar (TL)"), p.get("Dönem Net Kâr/Zarar (TL)")), signed=True),
            "Aktarım Hariç Teknik Sonuç / Brüt Prim": fmt_optional_pct(ex),
            "Aktarım Dahil Teknik Sonuç / Brüt Prim": fmt_optional_pct(inc),
            "_share": curr_share or 0.0,
            "_premium": num(c["Brüt Yazılan Prim (TL)"]),
            "_real": safe_real_growth(c["Brüt Yazılan Prim (TL)"], p["Brüt Yazılan Prim (TL)"], current) if p else None,
            "_net": num(c.get("Dönem Net Kâr/Zarar (TL)")),
        })
    sort_choice = st.selectbox("Toplam pazar payı tablosu sıralaması", ["Pazar payı", "Brüt prim", "Reel prim büyümesi", "Net kâr/zarar", "Şirket adı"], key="sector_company_sort")
    def _sort(r):
        if sort_choice == "Pazar payı": return r["_share"]
        if sort_choice == "Brüt prim": return r["_premium"]
        if sort_choice == "Reel prim büyümesi": return -999 if r["_real"] is None else r["_real"]
        if sort_choice == "Net kâr/zarar": return r["_net"]
        return r["Şirket"].lower()
    rows.sort(key=_sort, reverse=sort_choice != "Şirket adı")
    display_rows = [{k:v for k,v in r.items() if not k.startswith("_")} for r in rows]
    st.dataframe(display_rows, hide_index=True, use_container_width=True, height=590, column_config={"Şirket": st.column_config.TextColumn("Şirket", pinned=True, width="medium")})

    st.markdown("### Branş Görünümü")
    branch_pairs = []
    for branch in MAIN_BRANCHES:
        p = get_record(history, previous, 9000, branch)
        c = get_record(history, current, 9000, branch)
        if p and c: branch_pairs.append((branch, p, c))
    thresholds = {"Tümü":0,"En az 0,5 mlr TL prim":500_000_000,"En az 1 mlr TL prim":1_000_000_000,"En az 5 mlr TL prim":5_000_000_000}
    threshold_name = st.selectbox(f"Grafiklerde minimum {current} brüt yazılan prim", list(thresholds.keys()), index=0)
    filtered=[x for x in branch_pairs if num(x[2]["Brüt Yazılan Prim (TL)"])>=thresholds[threshold_name]]
    gross_tab, result_tab = st.tabs(["Brüt Teknik Branş Görünümü", "Teknik Sonuç Branş Görünümü"])
    with gross_tab:
        ordered=sorted(filtered,key=lambda x:safe_pp(x[2]["Brüt Bileşik Oran"],x[1]["Brüt Bileşik Oran"]))
        fig=go.Figure(go.Bar(x=[safe_pp(c["Brüt Bileşik Oran"],p["Brüt Bileşik Oran"]) for _,p,c in ordered],y=[b for b,_,_ in ordered],orientation="h"))
        fig.add_vline(x=0,line_width=1,line_dash="dot")
        fig.update_layout(title="Brüt bileşik oran değişimi",xaxis_title="Puan",height=max(430,25*len(ordered)+130),margin=dict(t=55,l=20,r=20,b=20))
        st.plotly_chart(fig,use_container_width=True)
        gross_rows=[]
        for b,p,c in branch_pairs:
            gross_rows.append({"Branş":b,f"Brüt Prim {current}":fmt_tl(c["Brüt Yazılan Prim (TL)"]),"Nominal Büyüme":fmt_growth(safe_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"])),"Reel Büyüme":fmt_growth(safe_real_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"],current)),f"H/P (DERK Dahil) {current}":fmt_pct(c["Brüt H/P"]),f"Masraf {current}":fmt_pct(c["Masraf Oranı"]),f"Bileşik {current}":fmt_pct(c["Brüt Bileşik Oran"]),"Bileşik Δ":fmt_pp(safe_pp(c["Brüt Bileşik Oran"],p["Brüt Bileşik Oran"]))})
        st.dataframe(gross_rows,hide_index=True,use_container_width=True,height=590,column_config={"Branş":st.column_config.TextColumn("Branş",pinned=True,width="medium")})
    with result_tab:
        ordered=sorted(filtered,key=lambda x:num(x[2]["Teknik Kâr/Zarar (TL)"]))
        fig=go.Figure()
        fig.add_bar(name="Aktarım Hariç",y=[b for b,_,_ in ordered],x=[num(c["Yatırım Hariç Teknik Sonuç (TL)"])/1e9 for _,_,c in ordered],orientation="h")
        fig.add_bar(name="Mali Gelir Aktarımı",y=[b for b,_,_ in ordered],x=[num(c["Yatırım Katkısı (TL)"])/1e9 for _,_,c in ordered],orientation="h")
        fig.add_bar(name="Aktarım Dahil",y=[b for b,_,_ in ordered],x=[num(c["Teknik Kâr/Zarar (TL)"])/1e9 for _,_,c in ordered],orientation="h")
        fig.add_vline(x=0,line_width=1,line_dash="dot")
        fig.update_layout(barmode="group",title=f"{current} branş teknik sonuç bileşenleri",xaxis_title="Milyar TL",height=max(520,32*len(ordered)+150),legend_title_text="",margin=dict(t=55,l=20,r=20,b=20))
        st.plotly_chart(fig,use_container_width=True)
        result_rows=[]
        for b,p,c in branch_pairs:
            result_rows.append({"Branş":b,f"Brüt Prim {current}":fmt_tl(c["Brüt Yazılan Prim (TL)"]),f"Aktarım Hariç Sonuç {current}":fmt_tl(c["Yatırım Hariç Teknik Sonuç (TL)"]),"Aktarım Hariç / Brüt Prim":fmt_optional_pct(tech_to_gwp(c,False)),f"Mali Gelir Aktarımı {current}":fmt_tl(c["Yatırım Katkısı (TL)"]),f"Aktarım Dahil Sonuç {current}":fmt_tl(c["Teknik Kâr/Zarar (TL)"]),"Aktarım Dahil / Brüt Prim":fmt_optional_pct(tech_to_gwp(c,True))})
        st.dataframe(result_rows,hide_index=True,use_container_width=True,height=590,column_config={"Branş":st.column_config.TextColumn("Branş",pinned=True,width="medium")})


def branch_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    st.title("Branş Analizi")
    branch = st.selectbox("Ana branş", MAIN_BRANCHES)
    prev_sector = get_record(history, previous, 9000, branch)
    curr_sector = get_record(history, current, 9000, branch)
    if not prev_sector or not curr_sector:
        st.error("Seçili branş için sektör toplamı bulunamadı.")
        return

    nominal_growth=safe_growth(curr_sector["Brüt Yazılan Prim (TL)"],prev_sector["Brüt Yazılan Prim (TL)"])
    real_growth=safe_real_growth(curr_sector["Brüt Yazılan Prim (TL)"],prev_sector["Brüt Yazılan Prim (TL)"],current)
    ex_prev, ex_curr=tech_to_gwp(prev_sector,False),tech_to_gwp(curr_sector,False)
    in_prev, in_curr=tech_to_gwp(prev_sector,True),tech_to_gwp(curr_sector,True)
    st.caption(f"{branch} • {previous} → {current} • sektör ve şirket dağılımı")

    st.markdown("### Brüt Teknik Görünüm")
    cols=st.columns(4)
    with cols[0]: kpi("Brüt Yazılan Prim",fmt_tl(curr_sector["Brüt Yazılan Prim (TL)"]),f"{previous}: {fmt_tl(prev_sector['Brüt Yazılan Prim (TL)'])}")
    with cols[1]: kpi("Brüt Kazanılmış Prim",fmt_tl(curr_sector["Brüt Kazanılmış Prim (TL)"]),f"{previous}: {fmt_tl(prev_sector['Brüt Kazanılmış Prim (TL)'])}")
    with cols[2]: kpi("Nominal Prim Büyümesi",fmt_growth(nominal_growth),f"{previous} → {current}")
    with cols[3]: kpi("Reel Prim Büyümesi",fmt_growth(real_growth),INFLATION_NOTE if current in INFLATION_YOY_BY_PERIOD else "Enflasyon girdisi tanımlı değil")
    cols=st.columns(3)
    with cols[0]: kpi("Brüt H/P (DERK Dahil)",fmt_pct(curr_sector["Brüt H/P"]),f"Δ {fmt_pp(safe_pp(curr_sector['Brüt H/P'],prev_sector['Brüt H/P']))}")
    with cols[1]: kpi("Masraf Oranı",fmt_pct(curr_sector["Masraf Oranı"]),f"Δ {fmt_pp(safe_pp(curr_sector['Masraf Oranı'],prev_sector['Masraf Oranı']))}")
    with cols[2]: kpi("Brüt Bileşik Oran",fmt_pct(curr_sector["Brüt Bileşik Oran"]),f"Δ {fmt_pp(safe_pp(curr_sector['Brüt Bileşik Oran'],prev_sector['Brüt Bileşik Oran']))}")

    st.divider()
    st.markdown("### Teknik Sonuç Görünümü")
    cols=st.columns(5)
    specs=[
        ("Aktarım Hariç Teknik Sonuç",fmt_tl(curr_sector["Yatırım Hariç Teknik Sonuç (TL)"]),f"Δ {fmt_tl(safe_delta(curr_sector['Yatırım Hariç Teknik Sonuç (TL)'],prev_sector['Yatırım Hariç Teknik Sonuç (TL)']),signed=True)}"),
        ("Mali Gelir Aktarımı (603)",fmt_tl(curr_sector["Yatırım Katkısı (TL)"]),f"Δ {fmt_tl(safe_delta(curr_sector['Yatırım Katkısı (TL)'],prev_sector['Yatırım Katkısı (TL)']),signed=True)}"),
        ("Aktarım Dahil Teknik Sonuç",fmt_tl(curr_sector["Teknik Kâr/Zarar (TL)"]),f"Δ {fmt_tl(safe_delta(curr_sector['Teknik Kâr/Zarar (TL)'],prev_sector['Teknik Kâr/Zarar (TL)']),signed=True)}"),
        ("Aktarım Hariç Sonuç / Brüt Prim",fmt_optional_pct(ex_curr),f"{previous}: {fmt_optional_pct(ex_prev)}"),
        ("Aktarım Dahil Sonuç / Brüt Prim",fmt_optional_pct(in_curr),f"{previous}: {fmt_optional_pct(in_prev)}"),
    ]
    for col,(t,v,s) in zip(cols,specs):
        with col:kpi(t,v,s)

    left,right=st.columns(2)
    with left:
        fig=go.Figure()
        xs=["Aktarım Hariç","Mali Gelir Aktarımı","Aktarım Dahil"]
        fig.add_bar(name=previous,x=xs,y=[num(prev_sector["Yatırım Hariç Teknik Sonuç (TL)"])/1e9,num(prev_sector["Yatırım Katkısı (TL)"])/1e9,num(prev_sector["Teknik Kâr/Zarar (TL)"])/1e9])
        fig.add_bar(name=current,x=xs,y=[num(curr_sector["Yatırım Hariç Teknik Sonuç (TL)"])/1e9,num(curr_sector["Yatırım Katkısı (TL)"])/1e9,num(curr_sector["Teknik Kâr/Zarar (TL)"])/1e9])
        fig.add_hline(y=0,line_width=1,line_dash="dot")
        fig.update_layout(barmode="group",title="Branş teknik sonuç bileşenleri",yaxis_title="Milyar TL",legend_title_text="",margin=dict(t=55,l=20,r=20,b=20))
        st.plotly_chart(fig,use_container_width=True)
    with right:
        fig=go.Figure()
        fig.add_bar(name=previous,x=["Aktarım Hariç","Aktarım Dahil"],y=[pct_for_plot(ex_prev),pct_for_plot(in_prev)])
        fig.add_bar(name=current,x=["Aktarım Hariç","Aktarım Dahil"],y=[pct_for_plot(ex_curr),pct_for_plot(in_curr)])
        fig.add_hline(y=0,line_width=1,line_dash="dot")
        fig.update_layout(barmode="group",title="Teknik sonuç / brüt yazılan prim",yaxis_title="%",legend_title_text="",margin=dict(t=55,l=20,r=20,b=20))
        st.plotly_chart(fig,use_container_width=True)

    pairs=period_pair_rows(history,previous,current,branch)
    active_pairs=[pair for pair in pairs if activity(pair["previous"]) or activity(pair["current"])]
    current_active=[pair for pair in active_pairs if pair["current"]]

    st.markdown("### Şirket Dağılımı")
    gross_tab,result_tab=st.tabs(["Brüt Teknik Dağılım","Teknik Sonuç / Brüt Prim Dağılımı"])
    with gross_tab:
        # Brüt H/P ve masraf oranı yalnız brüt kazanılmış prim paydası bulunan
        # güncel şirket kayıtlarında grafiklenir. Kaynak satırlar aşağıdaki tabloda korunur.
        gross_points = [
            pair for pair in current_active
            if abs(num(pair["current"].get("Brüt Kazanılmış Prim (TL)"))) > 1.0
        ]
        scale_mode=st.radio(
            "Grafik ölçeği",
            ["Ana dağılıma odaklan", "Tüm değerleri göster"],
            horizontal=True,
            key=f"gross_distribution_scale_{branch}_{current}",
            help="Odak görünümü yalnız ekseni merkezler; hiçbir kaynak satırı silinmez. Tüm rakamlar aşağıdaki tabloda aynen korunur.",
        )
        if not gross_points:
            st.info("Bu branşta brüt H/P / masraf dağılımı için geçerli brüt kazanılmış prim paydası bulunan güncel şirket gözlemi yok.")
        else:
            sizes=bubble_sizes([pair["current"]["Brüt Yazılan Prim (TL)"] for pair in gross_points])
            hp_values=[num(pair["current"]["Brüt H/P"])*100 for pair in gross_points]
            expense_values=[num(pair["current"]["Masraf Oranı"])*100 for pair in gross_points]
            fig=go.Figure(go.Scatter(
                x=hp_values,y=expense_values,mode="markers",
                marker={"size":sizes,"opacity":0.72,"color":PERIOD_COLOR_PREVIOUS},
                text=[pair["current"]["Şirket Adı"] for pair in gross_points],
                customdata=[[fmt_tl(pair["current"]["Brüt Yazılan Prim (TL)"]),fmt_pct(pair["current"]["Brüt Bileşik Oran"]),fmt_optional_pct(market_share(pair["current"],curr_sector))] for pair in gross_points],
                hovertemplate=f"<b>%{{text}}</b><br>{current} Brüt H/P: %{{x:.1f}}%<br>{current} Masraf: %{{y:.1f}}%<br>Prim: %{{customdata[0]}}<br>Brüt Bileşik: %{{customdata[1]}}<br>Branş Pazar Payı: %{{customdata[2]}}<extra></extra>"
            ))
            hp_sector=num(curr_sector["Brüt H/P"])*100
            exp_sector=num(curr_sector["Masraf Oranı"])*100
            fig.add_vline(x=hp_sector,line_width=1,line_dash="dot")
            fig.add_hline(y=exp_sector,line_width=1,line_dash="dot")
            st.caption("X ekseni Brüt H/P (DERK Dahil) %, Y ekseni Masraf Oranı %'dır. Kesikli çizgiler seçili branşın sektör oranlarını gösterir.")
            if scale_mode == "Ana dağılıma odaklan":
                x_low,x_high,x_outliers=robust_axis_range(hp_values,hp_sector,min_span=20.0)
                y_low,y_high,y_outliers=robust_axis_range(expense_values,exp_sector,min_span=20.0)
                fig.update_xaxes(range=[x_low,x_high])
                fig.update_yaxes(range=[y_low,y_high])
                if x_outliers or y_outliers:
                    st.caption(f"Odak görünümünde H/P ekseninde {x_outliers}, masraf ekseninde {y_outliers} uç değer görünüm dışında kalır; 'Tüm değerleri göster' seçeneğinde tamamı görülebilir.")
            excluded = len(current_active)-len(gross_points)
            if excluded:
                st.caption(f"Brüt kazanılmış primi sıfır olan {excluded} gözlem oran grafiğine dahil edilmez; sayısal tabloda kaynak değeri korunur.")
            fig.update_layout(
                xaxis_title=f"{current} Brüt H/P (DERK Dahil) (%)",
                yaxis_title=f"{current} Masraf Oranı (%)",
                height=520,margin=dict(t=20,l=20,r=20,b=20),hovermode="closest"
            )
            st.plotly_chart(fig,use_container_width=True,key=f"branch_gross_scatter_{branch}_{current}")
    with result_tab:
        result_points=[]
        invalid_denominator=0
        for pair in current_active:
            ex_value=tech_to_gwp(pair["current"],False)
            in_value=tech_to_gwp(pair["current"],True)
            if ex_value is None or in_value is None:
                invalid_denominator += 1
                continue
            ex_pct=pct_for_plot(ex_value)
            in_pct=pct_for_plot(in_value)
            if ex_pct is None or in_pct is None:
                invalid_denominator += 1
                continue
            result_points.append((pair,ex_pct,in_pct))
        result_scale=st.radio(
            "Grafik ölçeği",
            ["Ana dağılıma odaklan", "Tüm değerleri göster"],
            horizontal=True,
            key=f"result_distribution_scale_{branch}_{current}",
            help="Küçük prim paydaları teknik sonuç / prim oranında çok yüksek değerler üretebilir. Odak görünümü yalnız ekseni merkezler; kaynak değerler değiştirilmez.",
        )
        if not result_points:
            st.info("Bu branşta teknik sonuç / brüt prim dağılımı için geçerli brüt prim paydası bulunan güncel şirket gözlemi yok.")
        else:
            sizes=bubble_sizes([item[0]["current"]["Brüt Yazılan Prim (TL)"] for item in result_points])
            ex_values=[item[1] for item in result_points]
            in_values=[item[2] for item in result_points]
            fig=go.Figure(go.Scatter(
                x=ex_values,y=in_values,mode="markers",marker={"size":sizes,"opacity":0.70,"color":PERIOD_COLOR_PREVIOUS},
                text=[item[0]["current"]["Şirket Adı"] for item in result_points],
                customdata=[[fmt_tl(item[0]["current"]["Brüt Yazılan Prim (TL)"]),fmt_optional_pct(market_share(item[0]["current"],curr_sector)),fmt_tl(item[0]["current"]["Yatırım Katkısı (TL)"])] for item in result_points],
                hovertemplate=f"<b>%{{text}}</b><br>Aktarım Hariç / Brüt Prim: %{{x:.1f}}%<br>Aktarım Dahil / Brüt Prim: %{{y:.1f}}%<br>Prim: %{{customdata[0]}}<br>Branş Pazar Payı: %{{customdata[1]}}<br>Mali Gelir Aktarımı: %{{customdata[2]}}<extra></extra>"
            ))
            ex_sector=pct_for_plot(ex_curr)
            in_sector=pct_for_plot(in_curr)
            if ex_sector is not None:
                fig.add_vline(x=ex_sector,line_width=1,line_dash="dot")
            if in_sector is not None:
                fig.add_hline(y=in_sector,line_width=1,line_dash="dot")
            if result_scale == "Ana dağılıma odaklan":
                x_low,x_high,x_outliers=robust_axis_range(ex_values,ex_sector,min_span=30.0)
                y_low,y_high,y_outliers=robust_axis_range(in_values,in_sector,min_span=30.0)
                fig.update_xaxes(range=[x_low,x_high])
                fig.update_yaxes(range=[y_low,y_high])
                if x_outliers or y_outliers:
                    st.caption(f"Odak görünümünde aktarım hariç ekseninde {x_outliers}, aktarım dahil ekseninde {y_outliers} uç değer görünüm dışında kalır; 'Tüm değerleri göster' seçeneğinde tamamı görülebilir.")
            if invalid_denominator:
                st.caption(f"Brüt yazılan primi sıfır olan {invalid_denominator} gözlem oran grafiğine dahil edilmez; sayısal tabloda kaynak değeri korunur.")
            fig.update_layout(
                xaxis_title="Aktarım Hariç Teknik Sonuç / Brüt Prim (%)",
                yaxis_title="Aktarım Dahil Teknik Sonuç / Brüt Prim (%)",
                height=520,margin=dict(t=20,l=20,r=20,b=20),hovermode="closest"
            )
            st.plotly_chart(fig,use_container_width=True,key=f"branch_result_scatter_{branch}_{current}")
            st.caption("Kesikli çizgiler seçili branşın sektör oranlarını; balon büyüklüğü şirket brüt prim hacmini gösterir.")

    st.markdown("### Şirket Bazında Sayısal Görünüm")
    gross_table_tab,result_table_tab=st.tabs(["Brüt Teknik + Pazar Payı","Teknik Sonuç + Kârlılık Oranları"])
    with gross_table_tab:
        rows=[]
        for pair in active_pairs:
            p,c=pair["previous"],pair["current"]
            rec=c or p or {}
            ps=market_share(p,prev_sector) if p else None; cs=market_share(c,curr_sector) if c else None
            rows.append({"Şirket":rec.get("Şirket Adı",""),"Şirket Kodu":pair["code"],f"Prim {current}":"—" if not c else fmt_tl(c["Brüt Yazılan Prim (TL)"]),"Nominal Prim Büyümesi":"—" if not p or not c else fmt_growth(safe_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"])),"Reel Prim Büyümesi":"—" if not p or not c else fmt_growth(safe_real_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"],current)),f"Branş Pazar Payı {current}":fmt_optional_pct(cs),"Pazar Payı Δ":"—" if ps is None or cs is None else fmt_pp((cs-ps)*100),f"H/P (DERK Dahil) {current}":"—" if not c else fmt_pct(c["Brüt H/P"]),f"Masraf {current}":"—" if not c else fmt_pct(c["Masraf Oranı"]),f"Bileşik {current}":"—" if not c else fmt_pct(c["Brüt Bileşik Oran"])})
        st.dataframe(rows,hide_index=True,use_container_width=True,height=590,column_config={"Şirket":st.column_config.TextColumn("Şirket",pinned=True,width="medium")})
        st.download_button("Brüt teknik/pazar payı tablosunu CSV indir",rows_to_csv_bytes(rows),file_name=f"{branch}_{previous}_{current}_brut_pazar_payi.csv",mime="text/csv")
    with result_table_tab:
        rows=[]
        for pair in active_pairs:
            p,c=pair["previous"],pair["current"]
            rec=c or p or {}
            ex=tech_to_gwp(c,False) if c else None; inc=tech_to_gwp(c,True) if c else None
            rows.append({"Şirket":rec.get("Şirket Adı",""),"Şirket Kodu":pair["code"],f"Brüt Prim {current}":"—" if not c else fmt_tl(c["Brüt Yazılan Prim (TL)"]),f"Aktarım Hariç Sonuç {current}":"—" if not c else fmt_tl(c["Yatırım Hariç Teknik Sonuç (TL)"]),"Aktarım Hariç / Brüt Prim":fmt_optional_pct(ex),"Sektör Farkı":"—" if ex is None or ex_curr is None else fmt_pp((ex-ex_curr)*100),f"Mali Gelir Aktarımı {current}":"—" if not c else fmt_tl(c["Yatırım Katkısı (TL)"]),f"Aktarım Dahil Sonuç {current}":"—" if not c else fmt_tl(c["Teknik Kâr/Zarar (TL)"]),"Aktarım Dahil / Brüt Prim":fmt_optional_pct(inc),"Sektör Farkı (Dahil)":"—" if inc is None or in_curr is None else fmt_pp((inc-in_curr)*100)})
        st.dataframe(rows,hide_index=True,use_container_width=True,height=590,column_config={"Şirket":st.column_config.TextColumn("Şirket",pinned=True,width="medium")})
        st.download_button("Teknik sonuç/kârlılık tablosunu CSV indir",rows_to_csv_bytes(rows),file_name=f"{branch}_{previous}_{current}_teknik_sonuc.csv",mime="text/csv")


def company_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    st.title("Şirket Detayı")
    current_companies=[r for r in records_for_period_branch(history,current,"HAYATDISI",include_sector=False) if r.get("Şirket Tipi")=="HD"]
    if not current_companies:
        st.warning(f"{current} döneminde HD şirket kaydı bulunamadı.");return
    current_companies.sort(key=lambda r:str(r.get("Şirket Adı","")).lower())
    name_to_code={r["Şirket Adı"]:normalize_code(r["Şirket Kodu"]) for r in current_companies}
    selected_name=st.selectbox(
        "Şirket",
        [None] + list(name_to_code.keys()),
        index=0,
        format_func=lambda value: "Şirket seçiniz" if value is None else value,
    )
    if selected_name is None:
        st.info("Detayları görüntülemek için bir şirket seç.")
        return
    code=name_to_code[selected_name]
    p=get_record(history,previous,code,"HAYATDISI");c=get_record(history,current,code,"HAYATDISI")
    if not c: st.error("Seçili şirketin güncel dönem kaydı bulunamadı.");return
    prev_sector=get_record(history,previous,9000,"HAYATDISI");sector=get_record(history,current,9000,"HAYATDISI")
    total_share_prev=market_share(p,prev_sector) if p and prev_sector else None;total_share_curr=market_share(c,sector) if sector else None
    nominal=safe_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"]) if p else None
    real=safe_real_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"],current) if p else None
    st.caption(f"{selected_name} • {previous} → {current} • yalnızca sayısal değerler, dönemsel değişimler ve benchmark farkları")

    st.markdown("### Şirket Özeti")
    cols=st.columns(5)
    with cols[0]: kpi("Brüt Yazılan Prim",fmt_tl(c["Brüt Yazılan Prim (TL)"]),f"{previous}: {'—' if not p else fmt_tl(p['Brüt Yazılan Prim (TL)'])}")
    with cols[1]: kpi("Toplam Pazar Payı",fmt_optional_pct(total_share_curr),f"{previous}: {fmt_optional_pct(total_share_prev)} · Δ {'—' if total_share_prev is None or total_share_curr is None else fmt_pp((total_share_curr-total_share_prev)*100)}")
    with cols[2]: kpi("Nominal Prim Büyümesi",fmt_growth(nominal),f"{previous} → {current}")
    with cols[3]: kpi("Reel Prim Büyümesi",fmt_growth(real),INFLATION_NOTE if current in INFLATION_YOY_BY_PERIOD else "Enflasyon girdisi tanımlı değil")
    with cols[4]: kpi("Dönem Net Kârı / Zararı",fmt_tl(c.get("Dönem Net Kâr/Zarar (TL)")),f"{previous}: {'—' if not p else fmt_tl(p.get('Dönem Net Kâr/Zarar (TL)'))} · Δ {'—' if not p else fmt_tl(safe_delta(c.get('Dönem Net Kâr/Zarar (TL)'),p.get('Dönem Net Kâr/Zarar (TL)')),signed=True)}")

    st.markdown("### Brüt Teknik Görünüm")
    cols=st.columns(4)
    metrics=[("Brüt Kazanılmış Prim","Brüt Kazanılmış Prim (TL)","amount"),("Brüt H/P (DERK Dahil)","Brüt H/P","ratio"),("Masraf Oranı","Masraf Oranı","ratio"),("Brüt Bileşik Oran","Brüt Bileşik Oran","ratio")]
    for col,(title,key,kind) in zip(cols,metrics):
        with col:
            if not p: sub=f"{previous}: kayıt yok"
            elif kind=="amount":sub=f"{previous}: {fmt_tl(p[key])} · Değişim {fmt_growth(safe_growth(c[key],p[key]))}"
            else:sub=f"{previous}: {fmt_pct(p[key])} · Δ {fmt_pp(safe_pp(c[key],p[key]))}"
            kpi(title,fmt_tl(c[key]) if kind=="amount" else fmt_pct(c[key]),sub)
    if sector:
        benchmark=[{"Metrik":"Brüt H/P (DERK Dahil)",f"Şirket {current}":fmt_pct(c["Brüt H/P"]),f"HD Sektör {current}":fmt_pct(sector["Brüt H/P"]),"Fark":fmt_pp((num(c["Brüt H/P"])-num(sector["Brüt H/P"]))*100)},{"Metrik":"Masraf Oranı",f"Şirket {current}":fmt_pct(c["Masraf Oranı"]),f"HD Sektör {current}":fmt_pct(sector["Masraf Oranı"]),"Fark":fmt_pp((num(c["Masraf Oranı"])-num(sector["Masraf Oranı"]))*100)},{"Metrik":"Brüt Bileşik",f"Şirket {current}":fmt_pct(c["Brüt Bileşik Oran"]),f"HD Sektör {current}":fmt_pct(sector["Brüt Bileşik Oran"]),"Fark":fmt_pp((num(c["Brüt Bileşik Oran"])-num(sector["Brüt Bileşik Oran"]))*100)}]
        st.dataframe(benchmark,hide_index=True,use_container_width=True)

    st.divider();st.markdown("### Teknik Sonuç Görünümü")
    ex_prev=tech_to_gwp(p,False) if p else None;ex_curr=tech_to_gwp(c,False);in_prev=tech_to_gwp(p,True) if p else None;in_curr=tech_to_gwp(c,True)
    cols=st.columns(5)
    specs=[("Aktarım Hariç Teknik Sonuç",fmt_tl(c["Yatırım Hariç Teknik Sonuç (TL)"]),"—" if not p else f"Δ {fmt_tl(safe_delta(c['Yatırım Hariç Teknik Sonuç (TL)'],p['Yatırım Hariç Teknik Sonuç (TL)']),signed=True)}"),("Mali Gelir Aktarımı (603)",fmt_tl(c["Yatırım Katkısı (TL)"]),"—" if not p else f"Δ {fmt_tl(safe_delta(c['Yatırım Katkısı (TL)'],p['Yatırım Katkısı (TL)']),signed=True)}"),("Aktarım Dahil Teknik Sonuç",fmt_tl(c["Teknik Kâr/Zarar (TL)"]),"—" if not p else f"Δ {fmt_tl(safe_delta(c['Teknik Kâr/Zarar (TL)'],p['Teknik Kâr/Zarar (TL)']),signed=True)}"),("Aktarım Hariç Sonuç / Brüt Prim",fmt_optional_pct(ex_curr),f"{previous}: {fmt_optional_pct(ex_prev)}"),("Aktarım Dahil Sonuç / Brüt Prim",fmt_optional_pct(in_curr),f"{previous}: {fmt_optional_pct(in_prev)}")]
    for col,(t,v,s) in zip(cols,specs):
        with col:kpi(t,v,s)
    if sector:
        sec_ex=tech_to_gwp(sector,False);sec_in=tech_to_gwp(sector,True)
        benchmark=[{"Metrik":"Aktarım Hariç Teknik Sonuç / Brüt Prim",f"Şirket {current}":fmt_optional_pct(ex_curr),f"HD Sektör {current}":fmt_optional_pct(sec_ex),"Fark":"—" if ex_curr is None or sec_ex is None else fmt_pp((ex_curr-sec_ex)*100)},{"Metrik":"Aktarım Dahil Teknik Sonuç / Brüt Prim",f"Şirket {current}":fmt_optional_pct(in_curr),f"HD Sektör {current}":fmt_optional_pct(sec_in),"Fark":"—" if in_curr is None or sec_in is None else fmt_pp((in_curr-sec_in)*100)}]
        st.dataframe(benchmark,hide_index=True,use_container_width=True)

    left,right=st.columns(2)
    with left:
        fig=go.Figure();xs=["Aktarım Hariç","Mali Gelir Aktarımı","Aktarım Dahil"]
        if p:fig.add_bar(name=previous,x=xs,y=[num(p["Yatırım Hariç Teknik Sonuç (TL)"])/1e9,num(p["Yatırım Katkısı (TL)"])/1e9,num(p["Teknik Kâr/Zarar (TL)"])/1e9])
        fig.add_bar(name=current,x=xs,y=[num(c["Yatırım Hariç Teknik Sonuç (TL)"])/1e9,num(c["Yatırım Katkısı (TL)"])/1e9,num(c["Teknik Kâr/Zarar (TL)"])/1e9]);fig.add_hline(y=0,line_width=1,line_dash="dot")
        fig.update_layout(barmode="group",title="Teknik sonuç bileşenleri",yaxis_title="Milyar TL",legend_title_text="",margin=dict(t=55,l=20,r=20,b=20));st.plotly_chart(fig,use_container_width=True)
    with right:
        fig=go.Figure();
        if p:fig.add_bar(name=previous,x=["Aktarım Hariç","Aktarım Dahil"],y=[pct_for_plot(ex_prev),pct_for_plot(in_prev)])
        fig.add_bar(name=current,x=["Aktarım Hariç","Aktarım Dahil"],y=[pct_for_plot(ex_curr),pct_for_plot(in_curr)]);fig.add_hline(y=0,line_width=1,line_dash="dot")
        fig.update_layout(barmode="group",title="Teknik sonuç / brüt yazılan prim",yaxis_title="%",legend_title_text="",margin=dict(t=55,l=20,r=20,b=20));st.plotly_chart(fig,use_container_width=True)

    st.markdown("### Branş Kırılımı")
    branch_pairs=[(branch,get_record(history,previous,code,branch),get_record(history,current,code,branch)) for branch in MAIN_BRANCHES]
    only_active=st.checkbox("Sadece hareket bulunan branşları göster",value=True)
    if only_active:branch_pairs=[x for x in branch_pairs if activity(x[1]) or activity(x[2])]
    gross_tab,result_tab=st.tabs(["Brüt Teknik + Branş Pazar Payı","Teknik Sonuç + Kârlılık Oranları"])
    with gross_tab:
        rows=[]
        for branch,pr,cr in branch_pairs:
            psec=get_record(history,previous,9000,branch);csec=get_record(history,current,9000,branch)
            ps=market_share(pr,psec) if pr and psec else None;cs=market_share(cr,csec) if cr and csec else None
            rows.append({"Branş":branch,f"Prim {current}":"—" if not cr else fmt_tl(cr["Brüt Yazılan Prim (TL)"]),"Nominal Prim Büyümesi":"—" if not pr or not cr else fmt_growth(safe_growth(cr["Brüt Yazılan Prim (TL)"],pr["Brüt Yazılan Prim (TL)"])),"Reel Prim Büyümesi":"—" if not pr or not cr else fmt_growth(safe_real_growth(cr["Brüt Yazılan Prim (TL)"],pr["Brüt Yazılan Prim (TL)"],current)),f"Branş Pazar Payı {current}":fmt_optional_pct(cs),"Pazar Payı Δ":"—" if ps is None or cs is None else fmt_pp((cs-ps)*100),f"H/P (DERK Dahil) {current}":"—" if not cr else fmt_pct(cr["Brüt H/P"]),f"Masraf {current}":"—" if not cr else fmt_pct(cr["Masraf Oranı"]),f"Bileşik {current}":"—" if not cr else fmt_pct(cr["Brüt Bileşik Oran"])})
        st.dataframe(rows,hide_index=True,use_container_width=True,height=590,column_config={"Branş":st.column_config.TextColumn("Branş",pinned=True,width="medium")})
    with result_tab:
        rows=[]
        for branch,pr,cr in branch_pairs:
            csec=get_record(history,current,9000,branch);ex=tech_to_gwp(cr,False) if cr else None;inc=tech_to_gwp(cr,True) if cr else None;sec_ex=tech_to_gwp(csec,False) if csec else None;sec_in=tech_to_gwp(csec,True) if csec else None
            rows.append({"Branş":branch,f"Brüt Prim {current}":"—" if not cr else fmt_tl(cr["Brüt Yazılan Prim (TL)"]),f"Aktarım Hariç Sonuç {current}":"—" if not cr else fmt_tl(cr["Yatırım Hariç Teknik Sonuç (TL)"]),"Aktarım Hariç / Brüt Prim":fmt_optional_pct(ex),"Branş Sektörü Farkı":"—" if ex is None or sec_ex is None else fmt_pp((ex-sec_ex)*100),f"Mali Gelir Aktarımı {current}":"—" if not cr else fmt_tl(cr["Yatırım Katkısı (TL)"]),f"Aktarım Dahil Sonuç {current}":"—" if not cr else fmt_tl(cr["Teknik Kâr/Zarar (TL)"]),"Aktarım Dahil / Brüt Prim":fmt_optional_pct(inc),"Branş Sektörü Farkı (Dahil)":"—" if inc is None or sec_in is None else fmt_pp((inc-sec_in)*100)})
        st.dataframe(rows,hide_index=True,use_container_width=True,height=590,column_config={"Branş":st.column_config.TextColumn("Branş",pinned=True,width="medium")})

    active_current=[(b,cr) for b,_,cr in branch_pairs if cr and activity(cr)]
    if active_current:
        st.markdown("### Seçili Branş Detayı")
        selected_branch=st.selectbox("Şirket içi branş",[b for b,_ in active_current]);cr=next(rec for b,rec in active_current if b==selected_branch);pr=get_record(history,previous,code,selected_branch);sector_branch=get_record(history,current,9000,selected_branch);prev_sector_branch=get_record(history,previous,9000,selected_branch)
        gross_detail_tab,result_detail_tab=st.tabs(["Brüt Teknik Detay","Teknik Sonuç Detayı"])
        with gross_detail_tab:
            cs=market_share(cr,sector_branch);ps=market_share(pr,prev_sector_branch) if pr else None
            cols=st.columns(5)
            with cols[0]:kpi("Branş Prim",fmt_tl(cr["Brüt Yazılan Prim (TL)"]),f"{previous}: {'—' if not pr else fmt_tl(pr['Brüt Yazılan Prim (TL)'])}")
            with cols[1]:kpi("Branş Pazar Payı",fmt_optional_pct(cs),f"Δ {'—' if ps is None or cs is None else fmt_pp((cs-ps)*100)}")
            with cols[2]:kpi("Branş H/P (DERK Dahil)",fmt_pct(cr["Brüt H/P"]),f"Δ {'—' if not pr else fmt_pp(safe_pp(cr['Brüt H/P'],pr['Brüt H/P']))}")
            with cols[3]:kpi("Branş Masraf",fmt_pct(cr["Masraf Oranı"]),f"Δ {'—' if not pr else fmt_pp(safe_pp(cr['Masraf Oranı'],pr['Masraf Oranı']))}")
            with cols[4]:kpi("Branş Bileşik",fmt_pct(cr["Brüt Bileşik Oran"]),f"Δ {'—' if not pr else fmt_pp(safe_pp(cr['Brüt Bileşik Oran'],pr['Brüt Bileşik Oran']))}")
        with result_detail_tab:
            ex=tech_to_gwp(cr,False);inc=tech_to_gwp(cr,True);sec_ex=tech_to_gwp(sector_branch,False) if sector_branch else None;sec_in=tech_to_gwp(sector_branch,True) if sector_branch else None
            cols=st.columns(5)
            with cols[0]:kpi("Aktarım Hariç Teknik Sonuç",fmt_tl(cr["Yatırım Hariç Teknik Sonuç (TL)"]),"—" if not pr else f"Δ {fmt_tl(safe_delta(cr['Yatırım Hariç Teknik Sonuç (TL)'],pr['Yatırım Hariç Teknik Sonuç (TL)']),signed=True)}")
            with cols[1]:kpi("Mali Gelir Aktarımı",fmt_tl(cr["Yatırım Katkısı (TL)"]),"—" if not pr else f"Δ {fmt_tl(safe_delta(cr['Yatırım Katkısı (TL)'],pr['Yatırım Katkısı (TL)']),signed=True)}")
            with cols[2]:kpi("Aktarım Dahil Teknik Sonuç",fmt_tl(cr["Teknik Kâr/Zarar (TL)"]),"—" if not pr else f"Δ {fmt_tl(safe_delta(cr['Teknik Kâr/Zarar (TL)'],pr['Teknik Kâr/Zarar (TL)']),signed=True)}")
            with cols[3]:kpi("Aktarım Hariç / Brüt Prim",fmt_optional_pct(ex),f"Branş sektörü: {fmt_optional_pct(sec_ex)}")
            with cols[4]:kpi("Aktarım Dahil / Brüt Prim",fmt_optional_pct(inc),f"Branş sektörü: {fmt_optional_pct(sec_in)}")

    export=[]
    for branch,pr,cr in branch_pairs:
        csec=get_record(history,current,9000,branch);export.append({"Branş":branch,f"Brüt Yazılan Prim {current}":"" if not cr else cr["Brüt Yazılan Prim (TL)"],f"Branş Pazar Payı {current}":"" if not cr or not csec else market_share(cr,csec),f"Brüt H/P (DERK Dahil) {current}":"" if not cr else cr["Brüt H/P"],f"Masraf Oranı {current}":"" if not cr else cr["Masraf Oranı"],f"Brüt Bileşik Oran {current}":"" if not cr else cr["Brüt Bileşik Oran"],f"Mali Gelir Aktarımı Hariç Teknik Sonuç {current}":"" if not cr else cr["Yatırım Hariç Teknik Sonuç (TL)"],f"Aktarım Hariç Teknik Sonuç / Brüt Prim {current}":"" if not cr else tech_to_gwp(cr,False),f"Mali Gelir Aktarımı (603) {current}":"" if not cr else cr["Yatırım Katkısı (TL)"],f"Aktarım Dahil Teknik Sonuç {current}":"" if not cr else cr["Teknik Kâr/Zarar (TL)"],f"Aktarım Dahil Teknik Sonuç / Brüt Prim {current}":"" if not cr else tech_to_gwp(cr,True)})
    st.download_button("Seçili şirket branş verisini CSV indir",rows_to_csv_bytes(export),file_name=f"{code}_{previous}_{current}_brans.csv",mime="text/csv")



def neutral_company_note() -> None:
    st.markdown(
        "<div class='neutral-box'><b>Şirket karşılaştırmaları</b> yalnızca seçilen finansal göstergelerin sayısal gösterimidir; niteliksel değerlendirme, puanlama veya performans sınıflaması yapılmaz.</div>",
        unsafe_allow_html=True,
    )


def metric_value(row: dict[str, Any] | None, metric: str, sector_row: dict[str, Any] | None = None) -> float | None:
    if not row:
        return None
    mapping = {
        "Brüt Yazılan Prim": "Brüt Yazılan Prim (TL)",
        "Brüt Kazanılmış Prim": "Brüt Kazanılmış Prim (TL)",
        "Brüt H/P (DERK Dahil)": "Brüt H/P",
        "Masraf Oranı": "Masraf Oranı",
        "Brüt Bileşik Oran": "Brüt Bileşik Oran",
        "Mali Gelir Aktarımı Hariç Teknik Sonuç": "Yatırım Hariç Teknik Sonuç (TL)",
        "Mali Gelir Aktarımı (603)": "Yatırım Katkısı (TL)",
        "Aktarım Dahil Teknik Sonuç": "Teknik Kâr/Zarar (TL)",
        "Dönem Net Kârı / Zararı": "Dönem Net Kâr/Zarar (TL)",
    }
    if metric == "Pazar Payı":
        return market_share(row, sector_row)
    if metric == "Aktarım Hariç Teknik Sonuç / Brüt Prim":
        return tech_to_gwp(row, False)
    if metric == "Aktarım Dahil Teknik Sonuç / Brüt Prim":
        return tech_to_gwp(row, True)
    key = mapping.get(metric)
    return None if not key else num(row.get(key))


def metric_kind(metric: str) -> str:
    if metric in {"Brüt Yazılan Prim", "Brüt Kazanılmış Prim", "Mali Gelir Aktarımı Hariç Teknik Sonuç", "Mali Gelir Aktarımı (603)", "Aktarım Dahil Teknik Sonuç", "Dönem Net Kârı / Zararı"}:
        return "amount"
    return "ratio"


def dynamic_table_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    st.title("Dinamik Tablo")
    st.caption("Şirket veya branş bazında seçilen metriğin iki dönem karşılaştırması • nötr veri görünümü")
    neutral_company_note()

    control1, control2, control3 = st.columns([1.1, 1.3, 1.5])
    with control1:
        level = st.selectbox("Tablo düzeyi", ["Şirket Toplamı", "Şirket × Branş"], key="dyn_level")
    branch = "HAYATDISI"
    with control2:
        if level == "Şirket × Branş":
            branch = st.selectbox("Branş", MAIN_BRANCHES, key="dyn_branch")
        else:
            st.selectbox("Branş", ["HAYATDISI"], disabled=True, key="dyn_branch_total")
    metrics = [
        "Brüt Yazılan Prim", "Brüt Kazanılmış Prim", "Pazar Payı",
        "Brüt H/P (DERK Dahil)", "Masraf Oranı", "Brüt Bileşik Oran",
        "Mali Gelir Aktarımı Hariç Teknik Sonuç", "Mali Gelir Aktarımı (603)",
        "Aktarım Dahil Teknik Sonuç", "Aktarım Hariç Teknik Sonuç / Brüt Prim",
        "Aktarım Dahil Teknik Sonuç / Brüt Prim",
    ]
    if level == "Şirket Toplamı":
        metrics.append("Dönem Net Kârı / Zararı")
    with control3:
        metric = st.selectbox("Metrik", metrics, key="dyn_metric")

    pairs = period_pair_rows(history, previous, current, branch)
    prev_sector = get_record(history, previous, 9000, branch)
    curr_sector = get_record(history, current, 9000, branch)
    rows = []
    numeric_rows = []
    for pair in pairs:
        p, c = pair["previous"], pair["current"]
        ref = c or p
        if not ref or (not activity(p) and not activity(c)):
            continue
        pv = metric_value(p, metric, prev_sector)
        cv = metric_value(c, metric, curr_sector)
        delta = None if pv is None or cv is None else cv - pv
        rows.append({
            "Şirket": ref.get("Şirket Adı", ""),
            "Şirket Kodu": pair["code"],
            previous: pv,
            current: cv,
            "Değişim": delta,
        })
        numeric_rows.append((ref.get("Şirket Adı", ""), pv, cv))
    rows.sort(key=lambda r: r["Şirket"].lower())

    kind = metric_kind(metric)
    raw_rows = [dict(row) for row in rows]
    if kind == "amount":
        # Sayıları string'e çevirmiyoruz. Alttaki DataFrame sayısal dtype'ı korur;
        # böylece kullanıcı sütun başlığından sıraladığında mn/mlr metnine göre değil
        # gerçek TL tutarına göre sıralama yapılır. Styler yalnızca hücre görünümünü değiştirir.
        display_df = pd.DataFrame(raw_rows, columns=["Şirket", "Şirket Kodu", previous, current, "Değişim"])

        def _fmt_amount(value: Any, signed: bool = False) -> str:
            if pd.isna(value):
                return "—"
            return fmt_tl(value, signed=signed)

        display_rows = display_df.style.format({
            previous: lambda value: _fmt_amount(value),
            current: lambda value: _fmt_amount(value),
            "Değişim": lambda value: _fmt_amount(value, signed=True),
        })
        config = {
            "Şirket": st.column_config.TextColumn("Şirket", pinned=True, width="medium"),
            "Şirket Kodu": st.column_config.TextColumn("Şirket Kodu"),
        }
    else:
        display_rows = [dict(row) for row in rows]
        config = {
            "Şirket": st.column_config.TextColumn("Şirket", pinned=True, width="medium"),
            previous: st.column_config.NumberColumn(previous, format="%.2f%%"),
            current: st.column_config.NumberColumn(current, format="%.2f%%"),
            "Değişim": st.column_config.NumberColumn("Değişim", format="%+.2f puan"),
        }
        # Streamlit yüzde gösterimi için 0-1 yerine yüzde puanı kullanıyoruz.
        for row in display_rows:
            for col in (previous, current):
                if row[col] is not None:
                    row[col] *= 100
            if row["Değişim"] is not None:
                row["Değişim"] *= 100

    st.markdown("### Sayısal Görünüm")
    st.dataframe(display_rows, hide_index=True, use_container_width=True, height=610, column_config=config)

    company_names = [r["Şirket"] for r in rows]
    selected = st.multiselect(
        "Grafikte yan yana görmek istediğin şirketler",
        company_names,
        default=[],
        help="Grafik yalnızca senin seçtiğin şirketleri gösterir; otomatik Top/Bottom sıralaması yapılmaz.",
    )
    if selected:
        selected = selected[:8]
        chart_rows = [x for x in numeric_rows if x[0] in selected]
        ratio_scale = None
        if kind == "ratio":
            ratio_scale = st.radio(
                "Grafik ölçeği",
                ["Ana dağılıma odaklan", "Tüm değerleri göster"],
                horizontal=True,
                key=f"dyn_chart_scale_{branch}_{metric}_{current}",
                help="Odak görünümü yalnız ekseni merkezler; tablodaki sayısal değerler değişmez.",
            )
        fig = go.Figure()
        if kind == "amount":
            fig.add_bar(name=previous, x=[x[0] for x in chart_rows], y=[num(x[1])/1e9 for x in chart_rows], marker_color=PERIOD_COLOR_PREVIOUS)
            fig.add_bar(name=current, x=[x[0] for x in chart_rows], y=[num(x[2])/1e9 for x in chart_rows], marker_color=PERIOD_COLOR_CURRENT)
            ytitle = "Milyar TL"
        else:
            prev_plot=[pct_for_plot(x[1]) for x in chart_rows]
            curr_plot=[pct_for_plot(x[2]) for x in chart_rows]
            fig.add_bar(name=previous, x=[x[0] for x in chart_rows], y=prev_plot, marker_color=PERIOD_COLOR_PREVIOUS)
            fig.add_bar(name=current, x=[x[0] for x in chart_rows], y=curr_plot, marker_color=PERIOD_COLOR_CURRENT)
            ytitle = "%"
            values=[v for v in prev_plot+curr_plot if v is not None]
            if ratio_scale == "Ana dağılıma odaklan" and values:
                y_low,y_high,y_outliers=robust_axis_range(values,None,min_span=30.0)
                fig.update_yaxes(range=[min(y_low,0.0),max(y_high,0.0)])
                if y_outliers:
                    st.caption(f"Odak görünümünde {y_outliers} uç oran eksen dışında kalır; 'Tüm değerleri göster' seçeneğinde ve tabloda tam değerler görülebilir.")
        fig.add_hline(y=0, line_width=1, line_dash="dot")
        fig.update_layout(barmode="group", title=f"{metric} • seçili şirketler", yaxis_title=ytitle, legend_title_text="", margin=dict(t=55,l=20,r=20,b=90))
        st.plotly_chart(fig, use_container_width=True,key=f"dynamic_chart_{branch}_{metric}_{current}")

    export_rows = []
    for r in raw_rows:
        export_rows.append({k: v for k, v in r.items()})
    st.download_button("Dinamik tabloyu CSV indir", rows_to_csv_bytes(export_rows), file_name=f"TSB_HD_{branch}_{metric.replace(' ','_')}_{previous}_{current}.csv", mime="text/csv")


def company_compare_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    st.title("Şirket Karşılaştırma")
    st.caption("Kullanıcının seçtiği şirketleri aynı veri seti ve sektör referansı üzerinde yan yana gösterir")
    neutral_company_note()

    all_current = records_for_period_branch(history, current, "HAYATDISI", include_sector=False)
    name_code = sorted([(r.get("Şirket Adı", ""), normalize_code(r.get("Şirket Kodu"))) for r in all_current if r.get("Şirket Tipi") == "HD"], key=lambda x: x[0].lower())
    names = [x[0] for x in name_code]
    code_by_name = {n:c for n,c in name_code}

    left, right = st.columns([2, 1])
    with left:
        selected_names = st.multiselect("Karşılaştırılacak şirketler", names, default=[], help="En okunaklı görünüm için 2–6 şirket seçmeni öneririm.")
    with right:
        branch = st.selectbox("Kapsam", ["HAYATDISI"] + list(MAIN_BRANCHES), key="compare_branch")
    if not selected_names:
        st.info("Karşılaştırma için en az bir şirket seç.")
        return
    if len(selected_names) > 8:
        st.warning("Grafik okunabilirliği için ilk 8 seçimi gösteriyorum.")
        selected_names = selected_names[:8]

    prev_sector = get_record(history, previous, 9000, branch)
    curr_sector = get_record(history, current, 9000, branch)
    recs = []
    for name in selected_names:
        code = code_by_name[name]
        p = get_record(history, previous, code, branch)
        c = get_record(history, current, code, branch)
        recs.append((name, code, p, c))

    tab1, tab2, tab3, tab4 = st.tabs(["Prim & Pazar Payı", "Brüt Teknik", "Teknik Sonuç", "Net Kâr"])

    with tab1:
        rows=[]
        fig=go.Figure()
        fig.add_bar(name=previous, x=[n for n,_,_,_ in recs], y=[num(p.get("Brüt Yazılan Prim (TL)"))/1e9 if p else 0 for _,_,p,_ in recs], marker_color=PERIOD_COLOR_PREVIOUS)
        fig.add_bar(name=current, x=[n for n,_,_,_ in recs], y=[num(c.get("Brüt Yazılan Prim (TL)"))/1e9 if c else 0 for _,_,_,c in recs], marker_color=PERIOD_COLOR_CURRENT)
        fig.update_layout(barmode="group", title="Brüt yazılan prim", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55,l=20,r=20,b=80))
        st.plotly_chart(fig,use_container_width=True)
        for n,code,p,c in recs:
            ps=market_share(p,prev_sector) if p and prev_sector else None; cs=market_share(c,curr_sector) if c and curr_sector else None
            rows.append({
                "Şirket":n,
                f"Prim {previous}":None if not p else p["Brüt Yazılan Prim (TL)"],
                f"Prim {current}":None if not c else c["Brüt Yazılan Prim (TL)"],
                "Nominal Büyüme":None if not p or not c else safe_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"]),
                "Reel Büyüme":None if not p or not c else safe_real_growth(c["Brüt Yazılan Prim (TL)"],p["Brüt Yazılan Prim (TL)"],current),
                f"Pazar Payı {previous}":ps,
                f"Pazar Payı {current}":cs,
                "Pazar Payı Δ":None if ps is None or cs is None else (cs-ps),
            })
        for r in rows:
            for col in ["Nominal Büyüme","Reel Büyüme",f"Pazar Payı {previous}",f"Pazar Payı {current}","Pazar Payı Δ"]:
                if r[col] is not None:r[col]*=100
        styled = compact_money_table(
            rows,
            money_cols=[f"Prim {previous}", f"Prim {current}"],
            signed_pct_cols=["Nominal Büyüme", "Reel Büyüme"],
            pct_cols=[f"Pazar Payı {previous}", f"Pazar Payı {current}"],
            point_cols=["Pazar Payı Δ"],
        )
        st.dataframe(
            styled,
            hide_index=True,
            use_container_width=True,
            column_config={"Şirket":st.column_config.TextColumn("Şirket",pinned=True,width="medium")},
        )

    with tab2:
        rows=[]
        compare_scale=st.radio(
            "Grafik ölçeği",
            ["Ana dağılıma odaklan", "Tüm değerleri göster"],
            horizontal=True,
            key=f"compare_gross_scale_{branch}_{current}",
            help="Odak görünümü yalnız grafik eksenini merkezler; aşağıdaki tablo bütün oranları tam değeriyle gösterir.",
        )
        fig=go.Figure()
        compare_values=[]
        for idx,(metric,key) in enumerate([("H/P (DERK Dahil)","Brüt H/P"),("Masraf Oranı","Masraf Oranı"),("Brüt Bileşik","Brüt Bileşik Oran")]):
            series=[num(c.get(key))*100 if c else None for _,_,_,c in recs]
            compare_values.extend([v for v in series if v is not None])
            fig.add_bar(name=metric,x=[n for n,_,_,_ in recs],y=series,marker_color=NEUTRAL_SERIES_COLORS[idx])
        if compare_scale == "Ana dağılıma odaklan" and compare_values:
            y_low,y_high,y_outliers=robust_axis_range(compare_values,None,min_span=30.0)
            fig.update_yaxes(range=[min(y_low,0.0),max(y_high,0.0)])
            if y_outliers:
                st.caption(f"Odak görünümünde {y_outliers} uç oran eksen dışında kalır; 'Tüm değerleri göster' seçeneğinde ve tabloda tam değerler görülebilir.")
        fig.update_layout(barmode="group",title=f"{current} brüt teknik oranlar",yaxis_title="%",legend_title_text="",margin=dict(t=55,l=20,r=20,b=80))
        st.plotly_chart(fig,use_container_width=True,key=f"compare_gross_chart_{branch}_{current}")
        for n,code,p,c in recs:
            if not c: continue
            rows.append({
                "Şirket":n,
                "H/P (DERK Dahil)":c["Brüt H/P"]*100,
                "HD Sektör H/P":None if not curr_sector else curr_sector["Brüt H/P"]*100,
                "H/P Farkı (puan)":None if not curr_sector else (c["Brüt H/P"]-curr_sector["Brüt H/P"])*100,
                "Masraf":c["Masraf Oranı"]*100,
                "HD Sektör Masraf":None if not curr_sector else curr_sector["Masraf Oranı"]*100,
                "Bileşik":c["Brüt Bileşik Oran"]*100,
                "HD Sektör Bileşik":None if not curr_sector else curr_sector["Brüt Bileşik Oran"]*100,
            })
        cfg={k:st.column_config.NumberColumn(format="%.2f%%") for k in ["H/P (DERK Dahil)","HD Sektör H/P","Masraf","HD Sektör Masraf","Bileşik","HD Sektör Bileşik"]}
        cfg["Şirket"]=st.column_config.TextColumn("Şirket",pinned=True,width="medium")
        cfg["H/P Farkı (puan)"]=st.column_config.NumberColumn(format="%+.2f puan")
        st.dataframe(rows,hide_index=True,use_container_width=True,column_config=cfg)

    with tab3:
        rows=[]
        fig=go.Figure()
        for idx,(label,key) in enumerate([("Aktarım Hariç","Yatırım Hariç Teknik Sonuç (TL)"),("Mali Gelir Aktarımı","Yatırım Katkısı (TL)"),("Aktarım Dahil","Teknik Kâr/Zarar (TL)")]):
            fig.add_bar(name=label,x=[n for n,_,_,_ in recs],y=[num(c.get(key))/1e9 if c else 0 for _,_,_,c in recs],marker_color=NEUTRAL_SERIES_COLORS[idx])
        fig.add_hline(y=0,line_width=1,line_dash="dot")
        fig.update_layout(barmode="group",title=f"{current} teknik sonuç bileşenleri",yaxis_title="Milyar TL",legend_title_text="",margin=dict(t=55,l=20,r=20,b=80))
        st.plotly_chart(fig,use_container_width=True)
        sec_ex=tech_to_gwp(curr_sector,False) if curr_sector else None; sec_in=tech_to_gwp(curr_sector,True) if curr_sector else None
        for n,code,p,c in recs:
            if not c: continue
            ex=tech_to_gwp(c,False); inc=tech_to_gwp(c,True)
            rows.append({
                "Şirket":n,
                "Aktarım Hariç Teknik Sonuç":c["Yatırım Hariç Teknik Sonuç (TL)"],
                "Mali Gelir Aktarımı (603)":c["Yatırım Katkısı (TL)"],
                "Aktarım Dahil Teknik Sonuç":c["Teknik Kâr/Zarar (TL)"],
                "Aktarım Hariç / Brüt Prim":None if ex is None else ex*100,
                "HD Sektör Referansı (Hariç)":None if sec_ex is None else sec_ex*100,
                "Fark (Hariç, puan)":None if ex is None or sec_ex is None else (ex-sec_ex)*100,
                "Aktarım Dahil / Brüt Prim":None if inc is None else inc*100,
                "HD Sektör Referansı (Dahil)":None if sec_in is None else sec_in*100,
                "Fark (Dahil, puan)":None if inc is None or sec_in is None else (inc-sec_in)*100,
            })
        styled = compact_money_table(
            rows,
            money_cols=["Aktarım Hariç Teknik Sonuç", "Mali Gelir Aktarımı (603)", "Aktarım Dahil Teknik Sonuç"],
            pct_cols=["Aktarım Hariç / Brüt Prim", "HD Sektör Referansı (Hariç)", "Aktarım Dahil / Brüt Prim", "HD Sektör Referansı (Dahil)"],
            point_cols=["Fark (Hariç, puan)", "Fark (Dahil, puan)"],
        )
        st.dataframe(
            styled,
            hide_index=True,
            use_container_width=True,
            column_config={"Şirket":st.column_config.TextColumn("Şirket",pinned=True,width="medium")},
        )

    with tab4:
        if branch != "HAYATDISI":
            st.info("Dönem net kârı / zararı branş bazında dağıtılmadığı için yalnız HAYATDIŞI kapsamındaki şirket toplamında gösterilir.")
        else:
            fig=go.Figure()
            fig.add_bar(name=previous,x=[n for n,_,_,_ in recs],y=[num(p.get("Dönem Net Kâr/Zarar (TL)"))/1e9 if p else 0 for _,_,p,_ in recs],marker_color=PERIOD_COLOR_PREVIOUS)
            fig.add_bar(name=current,x=[n for n,_,_,_ in recs],y=[num(c.get("Dönem Net Kâr/Zarar (TL)"))/1e9 if c else 0 for _,_,_,c in recs],marker_color=PERIOD_COLOR_CURRENT)
            fig.add_hline(y=0,line_width=1,line_dash="dot")
            fig.update_layout(barmode="group",title="Dönem net kârı / zararı",yaxis_title="Milyar TL",legend_title_text="",margin=dict(t=55,l=20,r=20,b=80))
            st.plotly_chart(fig,use_container_width=True)
            rows=[]
            for n,code,p,c in recs:
                rows.append({"Şirket":n,previous:None if not p else p.get("Dönem Net Kâr/Zarar (TL)"),current:None if not c else c.get("Dönem Net Kâr/Zarar (TL)"),"Değişim":None if not p or not c else safe_delta(c.get("Dönem Net Kâr/Zarar (TL)"),p.get("Dönem Net Kâr/Zarar (TL)"))})
            styled = compact_money_table(
                rows,
                money_cols=[previous, current],
                signed_money_cols=["Değişim"],
            )
            st.dataframe(
                styled,
                hide_index=True,
                use_container_width=True,
                column_config={"Şirket":st.column_config.TextColumn("Şirket",pinned=True,width="medium")},
            )

def methodology_page() -> None:
    st.title("Metodoloji")
    st.caption("Kaynak: Türkiye Sigorta Birliği (TSB) finansal tabloları • reel büyüme için TÜİK TÜFE")
    st.markdown("### Kapsam")
    st.write("Yalnızca hayat dışı (HD) şirketler. Sektör toplamı: 9000 / T (HD).")
    st.markdown("### Ana branş evreni")
    st.write(", ".join(MAIN_BRANCHES))
    st.markdown("### Net bileşik oran")
    st.write("Sektör özetinde Net Bileşik Oran = Net H/P (DERK Dahil) + Faaliyet Giderleri / Net Kazanılmış Prim olarak gösterilir. Net H/P ve net kazanılmış prim Hasar-Prim dosyasından; faaliyet gideri 614 hesabından alınır.")
    st.markdown("### Ana toplam dışında tutulan ek bilgi kırılımları")
    st.write("TRAFİK, KASKO, DEV. DEST. TARIM SİGORTALARI, MÜHENDİSLİK SİGORTALARI")

    st.markdown("### Brüt Teknik Görünüm")
    st.code("""Brüt Kazanılmış Prim = 60001 + 60003 + 60101 + 60103 + 60201
Brüt Gerçekleşen Hasar = -(61001 + 61101)
Brüt H/P (DERK Dahil) = Brüt Gerçekleşen Hasar / Brüt Kazanılmış Prim
Masraf Oranı = -614 / Brüt Kazanılmış Prim
Brüt Bileşik Oran = Brüt H/P (DERK Dahil) + Masraf Oranı""",language="text")

    st.markdown("### Teknik Sonuç Görünümü")
    st.code("""Mali Gelir Aktarımı (603) = Teknik Olmayan Bölümden Aktarılan Yatırım Gelirleri
Mali Gelir Aktarımı Hariç Teknik Sonuç = Raporlanan Teknik Kâr/Zarar - 603
Aktarım Dahil Teknik Sonuç = Raporlanan Teknik Kâr/Zarar
Aktarım Hariç Teknik Sonuç / Brüt Prim = Aktarım Hariç Teknik Sonuç / Brüt Yazılan Prim
Aktarım Dahil Teknik Sonuç / Brüt Prim = Raporlanan Teknik Kâr/Zarar / Brüt Yazılan Prim""",language="text")
    st.info("Ana uygulamada TSB raporuyla uyumlu Brüt H/P (DERK Dahil) kullanılır. TSB Hasar Prim Oranları dosyasındaki DERK Hariç H/P doğrulama çalışma kitabında ayrıca izlenebilir. Teknik sonuç / brüt prim oranları teknik sonuç bloğunda sunulur; brüt H/P, masraf ve bileşik oranların doğrudan muhasebesel karşılığı olarak yorumlanmaz.")

    st.markdown("### Dönem Net Kârı / Zararı")
    st.code("Gelir Tablosu dosyası → MALI sayfası → 69 Dönem Net Karı Veya Zararı",language="text")
    st.write("Net kâr/zarar şirket ve HD sektör toplamı seviyesinde gösterilir; branşlara dağıtılmaz.")

    st.markdown("### Pazar Payı")
    st.code("""Toplam Pazar Payı = Şirket HAYATDIŞI Brüt Yazılan Prim / HD Sektör HAYATDIŞI Brüt Yazılan Prim
Branş Pazar Payı = Şirket Branş Brüt Yazılan Prim / Aynı Branş HD Sektör Brüt Yazılan Prim""",language="text")
    st.write("Pazar payı değişimleri puan (pp) olarak gösterilir.")

    st.markdown("### Şirket Sunum İlkesi")
    st.write("Şirket bazında niteliksel iyi/kötü sınıflaması, heatmap, puanlama veya otomatik Top/Bottom performans listesi kullanılmaz. Şirket karşılaştırmaları kullanıcı seçimi ve sayısal sektör referansları üzerinden nötr biçimde gösterilir.")

    st.markdown("### Reel Prim Büyümesi")
    st.code("Reel Büyüme = (1 + Nominal Prim Büyümesi) / (1 + TÜFE) - 1",language="text")
    st.write("2025H1 → 2026H1 karşılaştırmasında Haziran 2026 yıllık TÜFE %32,11 kullanılır.")

    st.markdown("### Şirket Sayfası Sunum İlkesi")
    st.write("Şirket bazında yalnızca sayısal değerler, dönemsel değişimler ve sektör/branş benchmark farkları gösterilir. Niteliksel sınıflama veya hüküm üretilmez.")


def update_page(history: list[dict[str, Any]], history_source: str) -> None:
    st.title("Veri Güncelleme")
    st.caption("Yeni TSB dönemini üç kaynak dosyadan otomatik oluşturur, doğrular ve mevcut tarihçeye ekler.")
    st.info("Üç dosyayı aynı yükleme alanına bırakabilirsin. Uygulama Gelir Tablosu, Hasar-Prim ve Faaliyet Giderleri dosyalarını sayfa yapısından otomatik ayırt eder.")

    periods = sorted_periods(history)
    c1, c2, c3 = st.columns(3)
    with c1:
        kpi("Aktif Dönem Sayısı", str(len(periods)), ", ".join(periods))
    with c2:
        current_period = periods[-1] if periods else "—"
        current_companies = len(records_for_period_branch(history, current_period, "HAYATDISI", include_sector=False)) if periods else 0
        kpi("Son Dönem", current_period, f"HD şirket kaydı: {current_companies}")
    with c3:
        kpi("Veri Kaynağı", "Aktif veri" if history_source == "active" else "Paket başlangıcı", "Yeni dönem kaydedildiğinde yerel aktif veri dosyası oluşturulur.")

    uploaded = st.file_uploader("TSB dönem dosyaları (.xlsx) — tam olarak 3 dosya", type=["xlsx"], accept_multiple_files=True)
    payload: list[tuple[str, bytes]] = []
    if uploaded:
        st.markdown("### Dosya algılama")
        detected = []
        for file in uploaded:
            content = file.getvalue()
            payload.append((file.name, content))
            try:
                book = XlsxRaw(content)
                role = classify_workbook(book)
                period, raw = detect_period(book)
                role_text = {"income": "Gelir Tablosu", "claims": "Hasar-Prim", "expense": "Faaliyet Giderleri", "unknown": "Algılanamadı"}.get(role, role)
                detected.append({"Dosya": file.name, "Algılanan Tip": role_text, "Dönem": period, "Kaynak Dönem Metni": raw})
            except Exception as exc:
                detected.append({"Dosya": file.name, "Algılanan Tip": "Hata", "Dönem": "—", "Kaynak Dönem Metni": str(exc)})
        st.dataframe(detected, hide_index=True, use_container_width=True)

    result: dict[str, Any] | None = None
    if len(payload) == 3:
        try:
            with st.spinner("Dosyalar okunuyor ve kontroller çalıştırılıyor..."):
                result = cached_prepare_import(tuple(payload))
        except Exception as exc:
            st.error(f"Dosyalar işlenemedi: {exc}")
    elif len(payload) > 3:
        st.warning(f"Tam olarak 3 dosya seçmelisin. Şu anda {len(payload)} dosya seçili.")
    elif len(payload) > 0:
        st.warning(f"Devam etmek için toplam 3 dosya gerekli. Şu anda {len(payload)} dosya seçili.")

    if result:
        st.markdown("### Otomatik kontrol sonucu")
        cross = result["cross_validation"]
        recon = result["reconciliation"]
        period_checks = {info["period"] for info in result["period_info"].values()}
        cols = st.columns(4)
        with cols[0]:
            status_box("Dönem", len(period_checks) == 1, result["period"], "Üç kaynak dosyanın dönemi")
        with cols[1]:
            sheets_ok = all(v["passed"] for v in result["sheet_checks"].values())
            status_box("18 Ana Branş", sheets_ok, "Sayfa yapısı tam" if sheets_ok else "Eksik sayfa", "HAYATDIŞI + 18 ana branş")
        with cols[2]:
            status_box("Kaynaklar Arası Kontrol", cross["passed"], f"{cross['compared_records']} kayıt", f"Tolerans: {cross['tolerance_tl']:.0f} TL")
        with cols[3]:
            status_box("18 Branş ↔ HAYATDIŞI", recon["passed"], f"{recon['checked_codes']} toplam", f"Maks. fark: {fmt_tl(recon['max_abs_diff_tl'], compact=False)}")

        st.markdown("### Kontrol detayları")
        detail_rows = [
            {"Kontrol": "Brüt kazanılmış prim ↔ Hasar-Prim", "Maksimum Mutlak Fark (TL)": cross["max_earned_premium_diff_tl"], "Durum": "OK" if cross["max_earned_premium_diff_tl"] <= cross["tolerance_tl"] else "Kontrol"},
            {"Kontrol": "Brüt gerçekleşen hasar ↔ Hasar-Prim", "Maksimum Mutlak Fark (TL)": cross["max_gross_incurred_diff_tl"], "Durum": "OK" if cross["max_gross_incurred_diff_tl"] <= cross["tolerance_tl"] else "Kontrol"},
            {"Kontrol": "Faaliyet gideri ↔ Faaliyet Giderleri 614", "Maksimum Mutlak Fark (TL)": cross["max_expense_diff_tl"], "Durum": "OK" if cross["max_expense_diff_tl"] <= cross["tolerance_tl"] else "Kontrol"},
            {"Kontrol": "18 ana branş toplamı ↔ HAYATDIŞI", "Maksimum Mutlak Fark (TL)": recon["max_abs_diff_tl"], "Durum": "OK" if recon["passed"] else "Kontrol"},
        ]
        st.dataframe(detail_rows, hide_index=True, use_container_width=True)

        sector = result["sector_total"]
        st.markdown("### Dönem önizleme")
        cols = st.columns(5)
        with cols[0]: kpi("HD Şirket", str(result["company_count"]), f"Toplam normalize kayıt: {result['record_count']}")
        with cols[1]: kpi("Brüt Yazılan Prim", fmt_tl(sector["Brüt Yazılan Prim (TL)"]), "9000 / T (HD)")
        with cols[2]: kpi("Brüt Bileşik", fmt_pct(sector["Brüt Bileşik Oran"]), "9000 / T (HD)")
        with cols[3]: kpi("Mali Gelir Aktarımı (603)", fmt_tl(sector["Yatırım Katkısı (TL)"]), "9000 / T (HD)")
        with cols[4]: kpi("Aktarım Hariç Teknik Sonuç", fmt_tl(sector["Yatırım Hariç Teknik Sonuç (TL)"]), "9000 / T (HD)")

        st.download_button("İşlenen dönem verisini CSV indir", rows_to_csv_bytes(result["rows"]), file_name=f"TSB_HD_{result['period']}_normalize.csv", mime="text/csv")

        if result["period"] in periods:
            st.warning(f"{result['period']} aktif veri setinde zaten mevcut. Kaydetme işlemi bu dönemin mevcut kayıtlarını yeni dosyalardan üretilen kayıtlarla değiştirecek.")

        if result["passed"]:
            if st.button("✓ Dönemi aktif veri setine ekle / güncelle", type="primary", use_container_width=True):
                merged = merge_period(history, result["rows"])
                save_history(ACTIVE_HISTORY, merged, source="TSB HD Streamlit v5 otomatik güncelleme")
                archived = archive_sources(IMPORT_ROOT, result["period"], payload)
                append_update_log(
                    UPDATE_LOG,
                    {
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "period": result["period"],
                        "company_count": result["company_count"],
                        "record_count": result["record_count"],
                        "sources": result["source_roles"],
                        "archived": archived,
                        "validation": {"cross": result["cross_validation"], "reconciliation": {k: v for k, v in result["reconciliation"].items() if k != "details"}},
                    },
                )
                cached_history.clear()
                cached_prepare_import.clear()
                st.success(f"{result['period']} aktif veri setine kaydedildi.")
                st.rerun()
        else:
            st.error("Kontrollerden biri geçmediği için dönem aktif veri setine kaydedilemez. Kaynak dosyaları kontrol et.")

    st.divider()
    st.markdown("### Veri yönetimi")
    left, right = st.columns(2)
    with left:
        st.download_button("Aktif veri setini JSON yedekle", history_to_json_bytes(history), file_name="TSB_HD_aktif_veri_yedek.json", mime="application/json", use_container_width=True)
    with right:
        st.download_button("Aktif normalize veriyi CSV indir", rows_to_csv_bytes(history), file_name="TSB_HD_aktif_veri.csv", mime="text/csv", use_container_width=True)

    logs = read_update_log(UPDATE_LOG)
    if logs:
        with st.expander("İçe aktarma geçmişi", expanded=False):
            log_rows = []
            for entry in reversed(logs):
                sources = entry.get("sources", {})
                log_rows.append(
                    {
                        "Tarih": entry.get("timestamp", ""),
                        "Dönem": entry.get("period", ""),
                        "HD Şirket": entry.get("company_count", ""),
                        "Kayıt": entry.get("record_count", ""),
                        "Gelir Dosyası": sources.get("income", {}).get("file_name", ""),
                        "Hasar-Prim Dosyası": sources.get("claims", {}).get("file_name", ""),
                        "Faaliyet Giderleri Dosyası": sources.get("expense", {}).get("file_name", ""),
                    }
                )
            st.dataframe(log_rows, hide_index=True, use_container_width=True)

    if history_source == "active":
        with st.expander("Paket başlangıç verisine dön", expanded=False):
            st.write("Bu işlem aktif_history.json dosyasını kaldırır ve uygulamayı paketle gelen 2025H1–2026H1 başlangıç verisine döndürür. Arşivlenen kaynak dosyaları silmez.")
            confirm = st.checkbox("Başlangıç verisine dönmek istediğimi onaylıyorum")
            if st.button("Başlangıç verisine dön", disabled=not confirm):
                ACTIVE_HISTORY.unlink(missing_ok=True)
                cached_history.clear()
                st.rerun()




def main() -> None:
    if not INITIAL_HISTORY.exists():
        st.error(f"Başlangıç veri dosyası bulunamadı: {INITIAL_HISTORY}")
        st.stop()

    try:
        history, history_source = cached_history()
    except Exception as exc:
        st.exception(exc)
        st.stop()

    periods = sorted_periods(history)
    if not periods:
        st.error("Aktif veri setinde dönem bulunamadı.")
        st.stop()

    with st.sidebar:
        st.markdown("<div class='app-kicker'>HAYAT DIŞI • SİGORTA SEKTÖRÜ</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='app-title'>{APP_TITLE}</div>", unsafe_allow_html=True)
        st.markdown("<div class='app-subtitle'>Sunum görünümü + analist araçları<br>Yerel veri güncelleme destekli</div>", unsafe_allow_html=True)
        st.markdown("#### Analiz Modülleri")
        page = st.radio(
            "Görünüm",
            ["Sektör Özeti", "Branş Analizi", "Şirket Detayı", "Dinamik Tablo", "Şirket Karşılaştırma", "Veri Güncelleme", "Metodoloji"],
            label_visibility="collapsed",
        )
        st.divider()
        st.markdown("#### Dönem")
        if len(periods) >= 2:
            previous_default = max(0, len(periods) - 2)
            current_default = len(periods) - 1
            previous = st.selectbox("Başlangıç dönemi", periods, index=previous_default, key="previous_period")
            current = st.selectbox("Karşılaştırma dönemi", periods, index=current_default, key="current_period")
        else:
            previous = current = periods[0]
            st.caption(f"Tek dönem mevcut: {current}")
        st.divider()
        st.markdown(
            f"<div class='small-muted'><b>{APP_VERSION}</b> · DERK dahil H/P<br>Pazar payı · reel büyüme · net kâr<br>Teknik sonuç / brüt prim<br><br>Kaynak: TSB finansal tabloları.</div>",
            unsafe_allow_html=True,
        )

    if page == "Sektör Özeti":
        sector_page(history, previous, current)
    elif page == "Branş Analizi":
        branch_page(history, previous, current)
    elif page == "Şirket Detayı":
        company_page(history, previous, current)
    elif page == "Dinamik Tablo":
        dynamic_table_page(history, previous, current)
    elif page == "Şirket Karşılaştırma":
        company_compare_page(history, previous, current)
    elif page == "Veri Güncelleme":
        update_page(history, history_source)
    else:
        methodology_page()


if __name__ == "__main__":
    main()
