from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

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

APP_TITLE = "TSB HD Kârlılık Analizi"
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
.block-container {padding-top: 1.25rem; padding-bottom: 2rem; max-width: 1520px;}
.kpi-card {border: 1px solid rgba(128,128,128,.28); border-radius: 12px; padding: 14px 16px; min-height: 118px; background: rgba(255,255,255,.015);}
.kpi-title {font-size: .82rem; opacity: .72; margin-bottom: 9px;}
.kpi-value {font-size: 1.48rem; font-weight: 650; letter-spacing: -.02em;}
.kpi-sub {font-size: .78rem; opacity: .72; margin-top: 7px; line-height: 1.4;}
.note-box {border-left: 4px solid rgba(128,128,128,.6); padding: 9px 12px; background: rgba(128,128,128,.06); border-radius: 4px;}
.small-muted {font-size:.79rem; opacity:.70;}
.section-caption {font-size:.82rem; opacity:.70; margin-top:-.35rem; margin-bottom:.85rem;}
.status-ok {border:1px solid rgba(46,160,67,.35); border-radius:12px; padding:12px 14px; min-height:92px; background:rgba(46,160,67,.06);}
.status-bad {border:1px solid rgba(220,53,69,.35); border-radius:12px; padding:12px 14px; min-height:92px; background:rgba(220,53,69,.06);}
.status-title {font-size:.78rem; opacity:.72; margin-bottom:6px;}
.status-value {font-size:1.05rem; font-weight:650;}
[data-testid="stDataFrame"] {border-radius: 10px; overflow: hidden;}
[data-testid="stMetricDelta"] {color: inherit !important;}
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
        f"brüt H/P {fmt_pct(current_row['Brüt H/P'])} ({fmt_pp(safe_pp(current_row['Brüt H/P'], previous_row['Brüt H/P']))}); "
        f"masraf oranı {fmt_pct(current_row['Masraf Oranı'])} ({fmt_pp(safe_pp(current_row['Masraf Oranı'], previous_row['Masraf Oranı']))}); "
        f"brüt bileşik oran {fmt_pct(current_row['Brüt Bileşik Oran'])} ({fmt_pp(safe_pp(current_row['Brüt Bileşik Oran'], previous_row['Brüt Bileşik Oran']))})."
    )


def sector_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    prev_total = get_record(history, previous, 9000, "HAYATDISI")
    curr_total = get_record(history, current, 9000, "HAYATDISI")
    if not prev_total or not curr_total:
        st.error("Seçili dönemlerden biri için HAYATDIŞI / 9000 T (HD) sektör toplamı bulunamadı.")
        return

    st.title("Sektör Özeti")
    st.caption(f"Hayat dışı şirketler • {previous} → {current} • 18 ana branş metodolojisi")

    st.markdown("### Brüt Teknik Görünüm")
    st.markdown(
        "<div class='section-caption'>Prim, hasar ve faaliyet gideri temelli brüt teknik oranlar. Teknik kâr/zarar bu blokta gösterilmez.</div>",
        unsafe_allow_html=True,
    )
    cols = st.columns(5)
    gross_specs = [
        ("Brüt Yazılan Prim", "Brüt Yazılan Prim (TL)", "amount"),
        ("Brüt Kazanılmış Prim", "Brüt Kazanılmış Prim (TL)", "amount"),
        ("Brüt H/P", "Brüt H/P", "ratio"),
        ("Masraf Oranı", "Masraf Oranı", "ratio"),
        ("Brüt Bileşik Oran", "Brüt Bileşik Oran", "ratio"),
    ]
    for col, (title, key, kind) in zip(cols, gross_specs):
        with col:
            if kind == "amount":
                kpi(title, fmt_tl(curr_total[key]), f"{previous}: {fmt_tl(prev_total[key])} · Değişim {fmt_growth(safe_growth(curr_total[key], prev_total[key]))}")
            else:
                kpi(title, fmt_pct(curr_total[key]), f"{previous}: {fmt_pct(prev_total[key])} · Δ {fmt_pp(safe_pp(curr_total[key], prev_total[key]))}")

    st.markdown(
        f"<div class='note-box'>{neutral_sector_summary(prev_total, curr_total, previous, current)}</div>",
        unsafe_allow_html=True,
    )

    left, right = st.columns(2)
    with left:
        fig = go.Figure()
        fig.add_bar(name=previous, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(prev_total["Brüt Yazılan Prim (TL)"]) / 1e9, num(prev_total["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.add_bar(name=current, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(curr_total["Brüt Yazılan Prim (TL)"]) / 1e9, num(curr_total["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.update_layout(barmode="group", title="Brüt prim görünümü", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = go.Figure()
        fig.add_bar(name=previous, x=["Brüt H/P", "Masraf Oranı", "Brüt Bileşik"], y=[num(prev_total["Brüt H/P"]) * 100, num(prev_total["Masraf Oranı"]) * 100, num(prev_total["Brüt Bileşik Oran"]) * 100])
        fig.add_bar(name=current, x=["Brüt H/P", "Masraf Oranı", "Brüt Bileşik"], y=[num(curr_total["Brüt H/P"]) * 100, num(curr_total["Masraf Oranı"]) * 100, num(curr_total["Brüt Bileşik Oran"]) * 100])
        fig.update_layout(barmode="group", title="Brüt teknik oranlar (%)", yaxis_title="%", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown("### Teknik Sonuç Görünümü")
    st.markdown(
        "<div class='section-caption'>Bu bölüm brüt teknik oranlardan ayrı okunur. “Mali Gelir Aktarımı Hariç Teknik Sonuç”, raporlanan teknik kâr/zarardan 603 hesabının çıkarılmasıyla hesaplanır.</div>",
        unsafe_allow_html=True,
    )
    cols = st.columns(4)
    result_specs = [
        ("Brüt Yazılan Prim", "Brüt Yazılan Prim (TL)"),
        ("Mali Gelir Aktarımı Hariç Teknik Sonuç", "Yatırım Hariç Teknik Sonuç (TL)"),
        ("Mali Gelir Aktarımı (603)", "Yatırım Katkısı (TL)"),
        ("Aktarım Dahil Teknik Sonuç", "Teknik Kâr/Zarar (TL)"),
    ]
    for col, (title, key) in zip(cols, result_specs):
        with col:
            if key == "Brüt Yazılan Prim (TL)":
                sub = f"{previous}: {fmt_tl(prev_total[key])} · Değişim {fmt_growth(safe_growth(curr_total[key], prev_total[key]))}"
            else:
                sub = f"{previous}: {fmt_tl(prev_total[key])} · Δ {fmt_tl(safe_delta(curr_total[key], prev_total[key]), signed=True)}"
            kpi(title, fmt_tl(curr_total[key]), sub)

    fig = go.Figure()
    result_x = ["Aktarım Hariç Teknik Sonuç", "Mali Gelir Aktarımı (603)", "Aktarım Dahil Teknik Sonuç"]
    fig.add_bar(name=previous, x=result_x, y=[num(prev_total["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(prev_total["Yatırım Katkısı (TL)"]) / 1e9, num(prev_total["Teknik Kâr/Zarar (TL)"]) / 1e9])
    fig.add_bar(name=current, x=result_x, y=[num(curr_total["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(curr_total["Yatırım Katkısı (TL)"]) / 1e9, num(curr_total["Teknik Kâr/Zarar (TL)"]) / 1e9])
    fig.add_hline(y=0, line_width=1, line_dash="dot")
    fig.update_layout(barmode="group", title="Teknik sonuç bileşenleri", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
    st.plotly_chart(fig, use_container_width=True)

    branch_pairs = []
    for branch in MAIN_BRANCHES:
        p = get_record(history, previous, 9000, branch)
        c = get_record(history, current, 9000, branch)
        if p and c:
            branch_pairs.append((branch, p, c))

    st.markdown("### Branş görünümü")
    thresholds = {
        "Tümü": 0,
        "En az 0,5 mlr TL prim": 500_000_000,
        "En az 1 mlr TL prim": 1_000_000_000,
        "En az 5 mlr TL prim": 5_000_000_000,
    }
    threshold_name = st.selectbox(f"Grafiklerde minimum {current} brüt yazılan prim", list(thresholds.keys()), index=0)
    threshold = thresholds[threshold_name]
    filtered = [(b, p, c) for b, p, c in branch_pairs if num(c["Brüt Yazılan Prim (TL)"]) >= threshold]

    gross_tab, result_tab = st.tabs(["Brüt Teknik Branş Görünümü", "Teknik Sonuç Branş Görünümü"])
    with gross_tab:
        left, right = st.columns(2)
        with left:
            ordered = sorted(filtered, key=lambda x: safe_pp(x[2]["Brüt Bileşik Oran"], x[1]["Brüt Bileşik Oran"]))
            fig = go.Figure(go.Bar(
                x=[safe_pp(c["Brüt Bileşik Oran"], p["Brüt Bileşik Oran"]) for _, p, c in ordered],
                y=[b for b, _, _ in ordered],
                orientation="h",
                customdata=[[fmt_pct(p["Brüt Bileşik Oran"]), fmt_pct(c["Brüt Bileşik Oran"]), fmt_pp(safe_pp(c["Brüt Bileşik Oran"], p["Brüt Bileşik Oran"]))] for _, p, c in ordered],
                hovertemplate=f"<b>%{{y}}</b><br>{previous}: %{{customdata[0]}}<br>{current}: %{{customdata[1]}}<br>Δ: %{{customdata[2]}}<extra></extra>",
            ))
            fig.add_vline(x=0, line_width=1, line_dash="dot")
            fig.update_layout(title="Brüt bileşik oran değişimi", xaxis_title="Puan", yaxis_title="", height=max(430, 25 * len(ordered) + 130), margin=dict(t=55, l=20, r=20, b=20))
            st.plotly_chart(fig, use_container_width=True)
        with right:
            ordered = sorted(filtered, key=lambda x: safe_growth(x[2]["Brüt Yazılan Prim (TL)"], x[1]["Brüt Yazılan Prim (TL)"]) or 0.0)
            fig = go.Figure(go.Bar(
                x=[(safe_growth(c["Brüt Yazılan Prim (TL)"], p["Brüt Yazılan Prim (TL)"]) or 0.0) * 100 for _, p, c in ordered],
                y=[b for b, _, _ in ordered],
                orientation="h",
                customdata=[[fmt_tl(p["Brüt Yazılan Prim (TL)"]), fmt_tl(c["Brüt Yazılan Prim (TL)"]), fmt_growth(safe_growth(c["Brüt Yazılan Prim (TL)"], p["Brüt Yazılan Prim (TL)"]))] for _, p, c in ordered],
                hovertemplate=f"<b>%{{y}}</b><br>{previous}: %{{customdata[0]}}<br>{current}: %{{customdata[1]}}<br>Değişim: %{{customdata[2]}}<extra></extra>",
            ))
            fig.add_vline(x=0, line_width=1, line_dash="dot")
            fig.update_layout(title="Brüt yazılan prim değişimi", xaxis_title="%", yaxis_title="", height=max(430, 25 * len(ordered) + 130), margin=dict(t=55, l=20, r=20, b=20))
            st.plotly_chart(fig, use_container_width=True)

        gross_rows = []
        for b, p, c in branch_pairs:
            gross_rows.append({
                "Branş": b,
                f"Brüt Yazılan Prim {previous}": fmt_tl(p["Brüt Yazılan Prim (TL)"]),
                f"Brüt Yazılan Prim {current}": fmt_tl(c["Brüt Yazılan Prim (TL)"]),
                "Prim Değişim": fmt_growth(safe_growth(c["Brüt Yazılan Prim (TL)"], p["Brüt Yazılan Prim (TL)"])),
                f"Brüt Kazanılmış Prim {current}": fmt_tl(c["Brüt Kazanılmış Prim (TL)"]),
                f"Brüt H/P {current}": fmt_pct(c["Brüt H/P"]),
                "H/P Δ": fmt_pp(safe_pp(c["Brüt H/P"], p["Brüt H/P"])),
                f"Masraf Oranı {current}": fmt_pct(c["Masraf Oranı"]),
                "Masraf Δ": fmt_pp(safe_pp(c["Masraf Oranı"], p["Masraf Oranı"])),
                f"Brüt Bileşik {current}": fmt_pct(c["Brüt Bileşik Oran"]),
                "Bileşik Δ": fmt_pp(safe_pp(c["Brüt Bileşik Oran"], p["Brüt Bileşik Oran"])),
            })
        st.dataframe(gross_rows, hide_index=True, use_container_width=True, height=590)

    with result_tab:
        st.markdown("<div class='section-caption'>Brüt teknik oranlar bu tabloda yer almaz. Brüt yazılan prim yalnızca branş hacmini göstermek için tutulur.</div>", unsafe_allow_html=True)
        ordered = sorted(filtered, key=lambda x: num(x[2]["Teknik Kâr/Zarar (TL)"]))
        fig = go.Figure()
        fig.add_bar(name="Aktarım Hariç Teknik Sonuç", y=[b for b, _, _ in ordered], x=[num(c["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9 for _, _, c in ordered], orientation="h")
        fig.add_bar(name="Mali Gelir Aktarımı (603)", y=[b for b, _, _ in ordered], x=[num(c["Yatırım Katkısı (TL)"]) / 1e9 for _, _, c in ordered], orientation="h")
        fig.add_bar(name="Aktarım Dahil Teknik Sonuç", y=[b for b, _, _ in ordered], x=[num(c["Teknik Kâr/Zarar (TL)"]) / 1e9 for _, _, c in ordered], orientation="h")
        fig.add_vline(x=0, line_width=1, line_dash="dot")
        fig.update_layout(barmode="group", title=f"{current} branş teknik sonuç bileşenleri", xaxis_title="Milyar TL", yaxis_title="", height=max(520, 32 * len(ordered) + 150), legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

        result_rows = []
        for b, p, c in branch_pairs:
            result_rows.append({
                "Branş": b,
                f"Brüt Yazılan Prim {current}": fmt_tl(c["Brüt Yazılan Prim (TL)"]),
                f"Aktarım Hariç Teknik Sonuç {previous}": fmt_tl(p["Yatırım Hariç Teknik Sonuç (TL)"]),
                f"Aktarım Hariç Teknik Sonuç {current}": fmt_tl(c["Yatırım Hariç Teknik Sonuç (TL)"]),
                "Aktarım Hariç Δ": fmt_tl(safe_delta(c["Yatırım Hariç Teknik Sonuç (TL)"], p["Yatırım Hariç Teknik Sonuç (TL)"]), signed=True),
                f"Mali Gelir Aktarımı (603) {current}": fmt_tl(c["Yatırım Katkısı (TL)"]),
                "Mali Gelir Aktarımı Δ": fmt_tl(safe_delta(c["Yatırım Katkısı (TL)"], p["Yatırım Katkısı (TL)"]), signed=True),
                f"Aktarım Dahil Teknik Sonuç {current}": fmt_tl(c["Teknik Kâr/Zarar (TL)"]),
                "Aktarım Dahil Teknik Sonuç Δ": fmt_tl(safe_delta(c["Teknik Kâr/Zarar (TL)"], p["Teknik Kâr/Zarar (TL)"]), signed=True),
            })
        st.dataframe(result_rows, hide_index=True, use_container_width=True, height=590)


def branch_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    st.title("Branş Analizi")
    branch = st.selectbox("Ana branş", MAIN_BRANCHES)
    prev_sector = get_record(history, previous, 9000, branch)
    curr_sector = get_record(history, current, 9000, branch)
    if not prev_sector or not curr_sector:
        st.error("Seçili branş için sektör toplamı bulunamadı.")
        return

    st.caption(f"{branch} • {previous} → {current} • Sektör ve şirket dağılımı")

    st.markdown("### Brüt Teknik Görünüm")
    cols = st.columns(5)
    gross_specs = [
        ("Brüt Yazılan Prim", "Brüt Yazılan Prim (TL)", "amount"),
        ("Brüt Kazanılmış Prim", "Brüt Kazanılmış Prim (TL)", "amount"),
        ("Brüt H/P", "Brüt H/P", "ratio"),
        ("Masraf Oranı", "Masraf Oranı", "ratio"),
        ("Brüt Bileşik Oran", "Brüt Bileşik Oran", "ratio"),
    ]
    for col, (title, key, kind) in zip(cols, gross_specs):
        with col:
            if kind == "amount":
                kpi(title, fmt_tl(curr_sector[key]), f"{previous}: {fmt_tl(prev_sector[key])} · Değişim {fmt_growth(safe_growth(curr_sector[key], prev_sector[key]))}")
            else:
                kpi(title, fmt_pct(curr_sector[key]), f"{previous}: {fmt_pct(prev_sector[key])} · Δ {fmt_pp(safe_pp(curr_sector[key], prev_sector[key]))}")

    left, right = st.columns(2)
    with left:
        fig = go.Figure()
        fig.add_bar(name=previous, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(prev_sector["Brüt Yazılan Prim (TL)"]) / 1e9, num(prev_sector["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.add_bar(name=current, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(curr_sector["Brüt Yazılan Prim (TL)"]) / 1e9, num(curr_sector["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.update_layout(barmode="group", title="Branş brüt prim görünümü", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = go.Figure()
        fig.add_bar(name=previous, x=["H/P", "Masraf", "Bileşik"], y=[num(prev_sector["Brüt H/P"]) * 100, num(prev_sector["Masraf Oranı"]) * 100, num(prev_sector["Brüt Bileşik Oran"]) * 100])
        fig.add_bar(name=current, x=["H/P", "Masraf", "Bileşik"], y=[num(curr_sector["Brüt H/P"]) * 100, num(curr_sector["Masraf Oranı"]) * 100, num(curr_sector["Brüt Bileşik Oran"]) * 100])
        fig.update_layout(barmode="group", title="Branş brüt teknik oranları (%)", yaxis_title="%", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown("### Teknik Sonuç Görünümü")
    st.markdown("<div class='section-caption'>Brüt teknik oranlardan ayrı sunulur. Aktarım hariç teknik sonuç = raporlanan teknik kâr/zarar − 603 mali gelir aktarımı.</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    result_specs = [
        ("Brüt Yazılan Prim", "Brüt Yazılan Prim (TL)"),
        ("Mali Gelir Aktarımı Hariç Teknik Sonuç", "Yatırım Hariç Teknik Sonuç (TL)"),
        ("Mali Gelir Aktarımı (603)", "Yatırım Katkısı (TL)"),
        ("Aktarım Dahil Teknik Sonuç", "Teknik Kâr/Zarar (TL)"),
    ]
    for col, (title, key) in zip(cols, result_specs):
        with col:
            if key == "Brüt Yazılan Prim (TL)":
                sub = f"{previous}: {fmt_tl(prev_sector[key])} · Değişim {fmt_growth(safe_growth(curr_sector[key], prev_sector[key]))}"
            else:
                sub = f"{previous}: {fmt_tl(prev_sector[key])} · Δ {fmt_tl(safe_delta(curr_sector[key], prev_sector[key]), signed=True)}"
            kpi(title, fmt_tl(curr_sector[key]), sub)

    fig = go.Figure()
    result_x = ["Aktarım Hariç Teknik Sonuç", "Mali Gelir Aktarımı (603)", "Aktarım Dahil Teknik Sonuç"]
    fig.add_bar(name=previous, x=result_x, y=[num(prev_sector["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(prev_sector["Yatırım Katkısı (TL)"]) / 1e9, num(prev_sector["Teknik Kâr/Zarar (TL)"]) / 1e9])
    fig.add_bar(name=current, x=result_x, y=[num(curr_sector["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(curr_sector["Yatırım Katkısı (TL)"]) / 1e9, num(curr_sector["Teknik Kâr/Zarar (TL)"]) / 1e9])
    fig.add_hline(y=0, line_width=1, line_dash="dot")
    fig.update_layout(barmode="group", title="Branş teknik sonuç bileşenleri", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
    st.plotly_chart(fig, use_container_width=True)

    pairs = period_pair_rows(history, previous, current, branch)
    active_pairs = [pair for pair in pairs if activity(pair["previous"]) or activity(pair["current"])]
    current_active = [pair for pair in active_pairs if pair["current"]]

    st.markdown("### Şirket dağılımı")
    gross_tab, result_tab = st.tabs(["Brüt Teknik Dağılım", "Teknik Sonuç Dağılımı"])
    with gross_tab:
        sizes = bubble_sizes([pair["current"]["Brüt Yazılan Prim (TL)"] for pair in current_active])
        fig = go.Figure(go.Scatter(
            x=[num(pair["current"]["Brüt H/P"]) * 100 for pair in current_active],
            y=[num(pair["current"]["Masraf Oranı"]) * 100 for pair in current_active],
            mode="markers",
            marker={"size": sizes, "opacity": 0.7},
            text=[pair["current"]["Şirket Adı"] for pair in current_active],
            customdata=[[fmt_tl(pair["current"]["Brüt Yazılan Prim (TL)"]), fmt_pct(pair["current"]["Brüt Bileşik Oran"])] for pair in current_active],
            hovertemplate=f"<b>%{{text}}</b><br>{current} Brüt H/P: %{{x:.1f}}%<br>{current} Masraf: %{{y:.1f}}%<br>Prim: %{{customdata[0]}}<br>Brüt Bileşik: %{{customdata[1]}}<extra></extra>",
        ))
        fig.add_vline(x=num(curr_sector["Brüt H/P"]) * 100, line_width=1, line_dash="dot")
        fig.add_hline(y=num(curr_sector["Masraf Oranı"]) * 100, line_width=1, line_dash="dot")
        fig.update_layout(xaxis_title=f"{current} Brüt H/P (%)", yaxis_title=f"{current} Masraf Oranı (%)", height=520, margin=dict(t=20, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Kesikli çizgiler seçili branşın sektör brüt H/P ve masraf oranını gösterir. Balon büyüklüğü brüt yazılan prim hacmidir.")

    with result_tab:
        sizes = bubble_sizes([pair["current"]["Brüt Yazılan Prim (TL)"] for pair in current_active])
        fig = go.Figure(go.Scatter(
            x=[num(pair["current"]["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9 for pair in current_active],
            y=[num(pair["current"]["Yatırım Katkısı (TL)"]) / 1e9 for pair in current_active],
            mode="markers",
            marker={"size": sizes, "opacity": 0.7},
            text=[pair["current"]["Şirket Adı"] for pair in current_active],
            customdata=[[fmt_tl(pair["current"]["Brüt Yazılan Prim (TL)"]), fmt_tl(pair["current"]["Teknik Kâr/Zarar (TL)"])] for pair in current_active],
            hovertemplate=f"<b>%{{text}}</b><br>{current} Aktarım Hariç Teknik Sonuç: %{{x:.1f}} mlr TL<br>{current} Mali Gelir Aktarımı: %{{y:.1f}} mlr TL<br>Prim: %{{customdata[0]}}<br>Aktarım Dahil Teknik Sonuç: %{{customdata[1]}}<extra></extra>",
        ))
        fig.add_vline(x=0, line_width=1, line_dash="dot")
        fig.add_hline(y=0, line_width=1, line_dash="dot")
        fig.update_layout(xaxis_title=f"{current} Mali Gelir Aktarımı Hariç Teknik Sonuç (milyar TL)", yaxis_title=f"{current} Mali Gelir Aktarımı (milyar TL)", height=520, margin=dict(t=20, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Şirket adları yalnızca üzerine gelindiğinde görünür. Balon büyüklüğü brüt yazılan prim hacmidir.")

    st.markdown("### Şirket bazında sayısal görünüm")
    gross_table_tab, result_table_tab = st.tabs(["Brüt Teknik Tablo", "Teknik Sonuç Tablosu"])

    with gross_table_tab:
        sort_choice = st.selectbox("Brüt teknik tablo sıralaması", ["Şirket adı", "Prim değişimi", "Bileşik değişimi"], key="branch_gross_sort")
        reverse = st.checkbox("Sıralamayı ters çevir", value=False, key="branch_gross_reverse")

        def gross_sort_key(pair: dict[str, Any]) -> Any:
            p, c = pair["previous"], pair["current"]
            name = (c or p or {}).get("Şirket Adı", "")
            if sort_choice == "Şirket adı":
                return str(name).lower()
            if sort_choice == "Prim değişimi":
                return safe_growth((c or {}).get("Brüt Yazılan Prim (TL)"), (p or {}).get("Brüt Yazılan Prim (TL)")) or 0.0
            return safe_pp((c or {}).get("Brüt Bileşik Oran"), (p or {}).get("Brüt Bileşik Oran"))

        ordered = sorted(active_pairs, key=gross_sort_key, reverse=reverse)
        gross_table = []
        for pair in ordered:
            p, c = pair["previous"], pair["current"]
            rec = c or p or {}
            gross_table.append({
                "Şirket": rec.get("Şirket Adı", ""),
                "Şirket Kodu": pair["code"],
                f"Prim {previous}": "—" if not p else fmt_tl(p["Brüt Yazılan Prim (TL)"]),
                f"Prim {current}": "—" if not c else fmt_tl(c["Brüt Yazılan Prim (TL)"]),
                "Prim Değişim": "—" if not p or not c else fmt_growth(safe_growth(c["Brüt Yazılan Prim (TL)"], p["Brüt Yazılan Prim (TL)"])),
                f"H/P {current}": "—" if not c else fmt_pct(c["Brüt H/P"]),
                "H/P Δ": "—" if not p or not c else fmt_pp(safe_pp(c["Brüt H/P"], p["Brüt H/P"])),
                f"Masraf {current}": "—" if not c else fmt_pct(c["Masraf Oranı"]),
                "Masraf Δ": "—" if not p or not c else fmt_pp(safe_pp(c["Masraf Oranı"], p["Masraf Oranı"])),
                f"Bileşik {current}": "—" if not c else fmt_pct(c["Brüt Bileşik Oran"]),
                "Bileşik Δ": "—" if not p or not c else fmt_pp(safe_pp(c["Brüt Bileşik Oran"], p["Brüt Bileşik Oran"])),
            })
        st.dataframe(gross_table, hide_index=True, use_container_width=True, height=590)
        st.download_button("Brüt teknik tabloyu CSV indir", rows_to_csv_bytes(gross_table), file_name=f"{branch}_{previous}_{current}_brut_teknik.csv", mime="text/csv")

    with result_table_tab:
        sort_choice_result = st.selectbox("Teknik sonuç tablosu sıralaması", ["Şirket adı", "Aktarım hariç sonuç değişimi", "Aktarım dahil sonuç değişimi"], key="branch_result_sort")
        reverse_result = st.checkbox("Sıralamayı ters çevir", value=False, key="branch_result_reverse")

        def result_sort_key(pair: dict[str, Any]) -> Any:
            p, c = pair["previous"], pair["current"]
            name = (c or p or {}).get("Şirket Adı", "")
            if sort_choice_result == "Şirket adı":
                return str(name).lower()
            if sort_choice_result == "Aktarım hariç sonuç değişimi":
                return safe_delta((c or {}).get("Yatırım Hariç Teknik Sonuç (TL)"), (p or {}).get("Yatırım Hariç Teknik Sonuç (TL)"))
            return safe_delta((c or {}).get("Teknik Kâr/Zarar (TL)"), (p or {}).get("Teknik Kâr/Zarar (TL)"))

        ordered_result = sorted(active_pairs, key=result_sort_key, reverse=reverse_result)
        result_table = []
        for pair in ordered_result:
            p, c = pair["previous"], pair["current"]
            rec = c or p or {}
            result_table.append({
                "Şirket": rec.get("Şirket Adı", ""),
                "Şirket Kodu": pair["code"],
                f"Brüt Yazılan Prim {current}": "—" if not c else fmt_tl(c["Brüt Yazılan Prim (TL)"]),
                f"Aktarım Hariç Teknik Sonuç {current}": "—" if not c else fmt_tl(c["Yatırım Hariç Teknik Sonuç (TL)"]),
                "Aktarım Hariç Δ": "—" if not p or not c else fmt_tl(safe_delta(c["Yatırım Hariç Teknik Sonuç (TL)"], p["Yatırım Hariç Teknik Sonuç (TL)"]), signed=True),
                f"Mali Gelir Aktarımı (603) {current}": "—" if not c else fmt_tl(c["Yatırım Katkısı (TL)"]),
                "Mali Gelir Aktarımı Δ": "—" if not p or not c else fmt_tl(safe_delta(c["Yatırım Katkısı (TL)"], p["Yatırım Katkısı (TL)"]), signed=True),
                f"Aktarım Dahil Teknik Sonuç {current}": "—" if not c else fmt_tl(c["Teknik Kâr/Zarar (TL)"]),
                "Aktarım Dahil Δ": "—" if not p or not c else fmt_tl(safe_delta(c["Teknik Kâr/Zarar (TL)"], p["Teknik Kâr/Zarar (TL)"]), signed=True),
            })
        st.dataframe(result_table, hide_index=True, use_container_width=True, height=590)
        st.download_button("Teknik sonuç tablosunu CSV indir", rows_to_csv_bytes(result_table), file_name=f"{branch}_{previous}_{current}_teknik_sonuc.csv", mime="text/csv")


def company_page(history: list[dict[str, Any]], previous: str, current: str) -> None:
    st.title("Şirket Detayı")
    current_companies = [
        r for r in records_for_period_branch(history, current, "HAYATDISI", include_sector=False)
        if r.get("Şirket Tipi") == "HD"
    ]
    if not current_companies:
        st.warning(f"{current} döneminde HD şirket kaydı bulunamadı.")
        return
    current_companies.sort(key=lambda r: str(r.get("Şirket Adı", "")).lower())
    name_to_code = {r["Şirket Adı"]: normalize_code(r["Şirket Kodu"]) for r in current_companies}
    selected_name = st.selectbox("Şirket", list(name_to_code.keys()))
    code = name_to_code[selected_name]
    p = get_record(history, previous, code, "HAYATDISI")
    c = get_record(history, current, code, "HAYATDISI")
    if not c:
        st.error("Seçili şirketin güncel dönem kaydı bulunamadı.")
        return

    st.caption(f"{selected_name} • {previous} → {current} • yalnızca sayısal değerler, değişimler ve benchmark farkları")

    st.markdown("### Brüt Teknik Görünüm")
    cols = st.columns(5)
    gross_metrics = [
        ("Brüt Yazılan Prim", "Brüt Yazılan Prim (TL)", "amount"),
        ("Brüt Kazanılmış Prim", "Brüt Kazanılmış Prim (TL)", "amount"),
        ("Brüt H/P", "Brüt H/P", "ratio"),
        ("Masraf Oranı", "Masraf Oranı", "ratio"),
        ("Brüt Bileşik Oran", "Brüt Bileşik Oran", "ratio"),
    ]
    for col, (title, key, kind) in zip(cols, gross_metrics):
        with col:
            if not p:
                sub = f"{previous}: kayıt yok"
            elif kind == "amount":
                sub = f"{previous}: {fmt_tl(p[key])} · Değişim {fmt_growth(safe_growth(c[key], p[key]))}"
            else:
                sub = f"{previous}: {fmt_pct(p[key])} · Δ {fmt_pp(safe_pp(c[key], p[key]))}"
            kpi(title, fmt_tl(c[key]) if kind == "amount" else fmt_pct(c[key]), sub)

    left, right = st.columns(2)
    with left:
        fig = go.Figure()
        if p:
            fig.add_bar(name=previous, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(p["Brüt Yazılan Prim (TL)"]) / 1e9, num(p["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.add_bar(name=current, x=["Brüt Yazılan Prim", "Brüt Kazanılmış Prim"], y=[num(c["Brüt Yazılan Prim (TL)"]) / 1e9, num(c["Brüt Kazanılmış Prim (TL)"]) / 1e9])
        fig.update_layout(barmode="group", title="Brüt prim görünümü", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fig = go.Figure()
        if p:
            fig.add_bar(name=previous, x=["H/P", "Masraf", "Bileşik"], y=[num(p["Brüt H/P"]) * 100, num(p["Masraf Oranı"]) * 100, num(p["Brüt Bileşik Oran"]) * 100])
        fig.add_bar(name=current, x=["H/P", "Masraf", "Bileşik"], y=[num(c["Brüt H/P"]) * 100, num(c["Masraf Oranı"]) * 100, num(c["Brüt Bileşik Oran"]) * 100])
        fig.update_layout(barmode="group", title="Brüt teknik oranlar (%)", yaxis_title="%", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    sector = get_record(history, current, 9000, "HAYATDISI")
    if sector:
        st.markdown("#### Şirket / HD sektör brüt teknik karşılaştırması")
        benchmark = [
            {"Metrik": "Brüt H/P", f"Şirket {current}": fmt_pct(c["Brüt H/P"]), f"HD Sektör {current}": fmt_pct(sector["Brüt H/P"]), "Fark": fmt_pp((num(c["Brüt H/P"]) - num(sector["Brüt H/P"])) * 100)},
            {"Metrik": "Masraf Oranı", f"Şirket {current}": fmt_pct(c["Masraf Oranı"]), f"HD Sektör {current}": fmt_pct(sector["Masraf Oranı"]), "Fark": fmt_pp((num(c["Masraf Oranı"]) - num(sector["Masraf Oranı"])) * 100)},
            {"Metrik": "Brüt Bileşik", f"Şirket {current}": fmt_pct(c["Brüt Bileşik Oran"]), f"HD Sektör {current}": fmt_pct(sector["Brüt Bileşik Oran"]), "Fark": fmt_pp((num(c["Brüt Bileşik Oran"]) - num(sector["Brüt Bileşik Oran"])) * 100)},
        ]
        st.dataframe(benchmark, hide_index=True, use_container_width=True)

    st.divider()
    st.markdown("### Teknik Sonuç Görünümü")
    st.markdown("<div class='section-caption'>Brüt teknik oranlardan ayrı sunulur. Şirket bazında yalnızca sayısal değer ve dönemsel değişim gösterilir.</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    result_metrics = [
        ("Brüt Yazılan Prim", "Brüt Yazılan Prim (TL)"),
        ("Mali Gelir Aktarımı Hariç Teknik Sonuç", "Yatırım Hariç Teknik Sonuç (TL)"),
        ("Mali Gelir Aktarımı (603)", "Yatırım Katkısı (TL)"),
        ("Aktarım Dahil Teknik Sonuç", "Teknik Kâr/Zarar (TL)"),
    ]
    for col, (title, key) in zip(cols, result_metrics):
        with col:
            if not p:
                sub = f"{previous}: kayıt yok"
            elif key == "Brüt Yazılan Prim (TL)":
                sub = f"{previous}: {fmt_tl(p[key])} · Değişim {fmt_growth(safe_growth(c[key], p[key]))}"
            else:
                sub = f"{previous}: {fmt_tl(p[key])} · Δ {fmt_tl(safe_delta(c[key], p[key]), signed=True)}"
            kpi(title, fmt_tl(c[key]), sub)

    fig = go.Figure()
    result_x = ["Aktarım Hariç Teknik Sonuç", "Mali Gelir Aktarımı (603)", "Aktarım Dahil Teknik Sonuç"]
    if p:
        fig.add_bar(name=previous, x=result_x, y=[num(p["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(p["Yatırım Katkısı (TL)"]) / 1e9, num(p["Teknik Kâr/Zarar (TL)"]) / 1e9])
    fig.add_bar(name=current, x=result_x, y=[num(c["Yatırım Hariç Teknik Sonuç (TL)"]) / 1e9, num(c["Yatırım Katkısı (TL)"]) / 1e9, num(c["Teknik Kâr/Zarar (TL)"]) / 1e9])
    fig.add_hline(y=0, line_width=1, line_dash="dot")
    fig.update_layout(barmode="group", title="Teknik sonuç bileşenleri", yaxis_title="Milyar TL", legend_title_text="", margin=dict(t=55, l=20, r=20, b=20))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Branş kırılımı")
    branch_pairs = []
    for branch in MAIN_BRANCHES:
        pr = get_record(history, previous, code, branch)
        cr = get_record(history, current, code, branch)
        branch_pairs.append((branch, pr, cr))
    only_active = st.checkbox("Sadece hareket bulunan branşları göster", value=True)
    if only_active:
        branch_pairs = [(b, pr, cr) for b, pr, cr in branch_pairs if activity(pr) or activity(cr)]

    gross_tab, result_tab = st.tabs(["Brüt Teknik Branş Tablosu", "Teknik Sonuç Branş Tablosu"])
    with gross_tab:
        gross_table = []
        for branch, pr, cr in branch_pairs:
            gross_table.append({
                "Branş": branch,
                f"Prim {previous}": "—" if not pr else fmt_tl(pr["Brüt Yazılan Prim (TL)"]),
                f"Prim {current}": "—" if not cr else fmt_tl(cr["Brüt Yazılan Prim (TL)"]),
                "Prim Değişim": "—" if not pr or not cr else fmt_growth(safe_growth(cr["Brüt Yazılan Prim (TL)"], pr["Brüt Yazılan Prim (TL)"])),
                f"Brüt Kazanılmış Prim {current}": "—" if not cr else fmt_tl(cr["Brüt Kazanılmış Prim (TL)"]),
                f"H/P {current}": "—" if not cr else fmt_pct(cr["Brüt H/P"]),
                "H/P Δ": "—" if not pr or not cr else fmt_pp(safe_pp(cr["Brüt H/P"], pr["Brüt H/P"])),
                f"Masraf {current}": "—" if not cr else fmt_pct(cr["Masraf Oranı"]),
                "Masraf Δ": "—" if not pr or not cr else fmt_pp(safe_pp(cr["Masraf Oranı"], pr["Masraf Oranı"])),
                f"Bileşik {current}": "—" if not cr else fmt_pct(cr["Brüt Bileşik Oran"]),
                "Bileşik Δ": "—" if not pr or not cr else fmt_pp(safe_pp(cr["Brüt Bileşik Oran"], pr["Brüt Bileşik Oran"])),
            })
        st.dataframe(gross_table, hide_index=True, use_container_width=True, height=590)

    with result_tab:
        result_table = []
        for branch, pr, cr in branch_pairs:
            result_table.append({
                "Branş": branch,
                f"Brüt Yazılan Prim {current}": "—" if not cr else fmt_tl(cr["Brüt Yazılan Prim (TL)"]),
                f"Aktarım Hariç Teknik Sonuç {previous}": "—" if not pr else fmt_tl(pr["Yatırım Hariç Teknik Sonuç (TL)"]),
                f"Aktarım Hariç Teknik Sonuç {current}": "—" if not cr else fmt_tl(cr["Yatırım Hariç Teknik Sonuç (TL)"]),
                "Aktarım Hariç Δ": "—" if not pr or not cr else fmt_tl(safe_delta(cr["Yatırım Hariç Teknik Sonuç (TL)"], pr["Yatırım Hariç Teknik Sonuç (TL)"]), signed=True),
                f"Mali Gelir Aktarımı (603) {current}": "—" if not cr else fmt_tl(cr["Yatırım Katkısı (TL)"]),
                "Mali Gelir Aktarımı Δ": "—" if not pr or not cr else fmt_tl(safe_delta(cr["Yatırım Katkısı (TL)"], pr["Yatırım Katkısı (TL)"]), signed=True),
                f"Aktarım Dahil Teknik Sonuç {current}": "—" if not cr else fmt_tl(cr["Teknik Kâr/Zarar (TL)"]),
                "Aktarım Dahil Δ": "—" if not pr or not cr else fmt_tl(safe_delta(cr["Teknik Kâr/Zarar (TL)"], pr["Teknik Kâr/Zarar (TL)"]), signed=True),
            })
        st.dataframe(result_table, hide_index=True, use_container_width=True, height=590)

    active_current = [(b, cr) for b, _, cr in branch_pairs if cr and activity(cr)]
    if active_current:
        st.markdown("### Seçili branş detayı")
        selected_branch = st.selectbox("Şirket içi branş", [b for b, _ in active_current])
        cr = next(rec for b, rec in active_current if b == selected_branch)
        pr = get_record(history, previous, code, selected_branch)
        sector_branch = get_record(history, current, 9000, selected_branch)

        gross_detail_tab, result_detail_tab = st.tabs(["Brüt Teknik Detay", "Teknik Sonuç Detayı"])
        with gross_detail_tab:
            cols = st.columns(4)
            with cols[0]:
                kpi("Branş Prim", fmt_tl(cr["Brüt Yazılan Prim (TL)"]), f"{previous}: {'—' if not pr else fmt_tl(pr['Brüt Yazılan Prim (TL)'])}")
            with cols[1]:
                kpi("Branş H/P", fmt_pct(cr["Brüt H/P"]), f"Δ {'—' if not pr else fmt_pp(safe_pp(cr['Brüt H/P'], pr['Brüt H/P']))}")
            with cols[2]:
                kpi("Branş Masraf", fmt_pct(cr["Masraf Oranı"]), f"Δ {'—' if not pr else fmt_pp(safe_pp(cr['Masraf Oranı'], pr['Masraf Oranı']))}")
            with cols[3]:
                kpi("Branş Bileşik", fmt_pct(cr["Brüt Bileşik Oran"]), f"Δ {'—' if not pr else fmt_pp(safe_pp(cr['Brüt Bileşik Oran'], pr['Brüt Bileşik Oran']))}")
            if sector_branch:
                benchmark = [
                    {"Metrik": "Brüt H/P", f"Şirket {current}": fmt_pct(cr["Brüt H/P"]), f"Branş Sektörü {current}": fmt_pct(sector_branch["Brüt H/P"]), "Fark": fmt_pp((num(cr["Brüt H/P"]) - num(sector_branch["Brüt H/P"])) * 100)},
                    {"Metrik": "Masraf Oranı", f"Şirket {current}": fmt_pct(cr["Masraf Oranı"]), f"Branş Sektörü {current}": fmt_pct(sector_branch["Masraf Oranı"]), "Fark": fmt_pp((num(cr["Masraf Oranı"]) - num(sector_branch["Masraf Oranı"])) * 100)},
                    {"Metrik": "Brüt Bileşik", f"Şirket {current}": fmt_pct(cr["Brüt Bileşik Oran"]), f"Branş Sektörü {current}": fmt_pct(sector_branch["Brüt Bileşik Oran"]), "Fark": fmt_pp((num(cr["Brüt Bileşik Oran"]) - num(sector_branch["Brüt Bileşik Oran"])) * 100)},
                ]
                st.dataframe(benchmark, hide_index=True, use_container_width=True)

        with result_detail_tab:
            cols = st.columns(4)
            with cols[0]:
                kpi("Branş Prim", fmt_tl(cr["Brüt Yazılan Prim (TL)"]), f"{previous}: {'—' if not pr else fmt_tl(pr['Brüt Yazılan Prim (TL)'])}")
            with cols[1]:
                kpi("Aktarım Hariç Teknik Sonuç", fmt_tl(cr["Yatırım Hariç Teknik Sonuç (TL)"]), f"Δ {'—' if not pr else fmt_tl(safe_delta(cr['Yatırım Hariç Teknik Sonuç (TL)'], pr['Yatırım Hariç Teknik Sonuç (TL)']), signed=True)}")
            with cols[2]:
                kpi("Mali Gelir Aktarımı (603)", fmt_tl(cr["Yatırım Katkısı (TL)"]), f"Δ {'—' if not pr else fmt_tl(safe_delta(cr['Yatırım Katkısı (TL)'], pr['Yatırım Katkısı (TL)']), signed=True)}")
            with cols[3]:
                kpi("Aktarım Dahil Teknik Sonuç", fmt_tl(cr["Teknik Kâr/Zarar (TL)"]), f"Δ {'—' if not pr else fmt_tl(safe_delta(cr['Teknik Kâr/Zarar (TL)'], pr['Teknik Kâr/Zarar (TL)']), signed=True)}")

    combined_export = []
    for branch, pr, cr in branch_pairs:
        combined_export.append({
            "Branş": branch,
            f"Brüt Yazılan Prim {current}": "" if not cr else cr["Brüt Yazılan Prim (TL)"],
            f"Brüt H/P {current}": "" if not cr else cr["Brüt H/P"],
            f"Masraf Oranı {current}": "" if not cr else cr["Masraf Oranı"],
            f"Brüt Bileşik Oran {current}": "" if not cr else cr["Brüt Bileşik Oran"],
            f"Mali Gelir Aktarımı Hariç Teknik Sonuç {current}": "" if not cr else cr["Yatırım Hariç Teknik Sonuç (TL)"],
            f"Mali Gelir Aktarımı (603) {current}": "" if not cr else cr["Yatırım Katkısı (TL)"],
            f"Aktarım Dahil Teknik Sonuç {current}": "" if not cr else cr["Teknik Kâr/Zarar (TL)"],
        })
    st.download_button("Seçili şirket branş verisini CSV indir", rows_to_csv_bytes(combined_export), file_name=f"{code}_{previous}_{current}_brans.csv", mime="text/csv")


def methodology_page() -> None:
    st.title("Metodoloji")
    st.caption("Paylaşım sürümü • salt okunur demo • kaynak: Türkiye Sigorta Birliği (TSB)")
    st.markdown("### Kapsam")
    st.write("Yalnızca hayat dışı (HD) şirketler. Sektör toplamı: 9000 / T (HD).")
    st.markdown("### Ana branş evreni")
    st.write(", ".join(MAIN_BRANCHES))
    st.markdown("### Ana toplam dışında tutulan ek bilgi kırılımları")
    st.write("TRAFİK, KASKO, DEV. DEST. TARIM SİGORTALARI, MÜHENDİSLİK SİGORTALARI")

    st.markdown("### Brüt Teknik Görünüm")
    st.code(
        """Brüt Kazanılmış Prim = 60001 + 60003 + 60101 + 60103 + 60201
Brüt Gerçekleşen Hasar = -(61001 + 61101)
Brüt H/P = Brüt Gerçekleşen Hasar / Brüt Kazanılmış Prim
Masraf Oranı = -614 / Brüt Kazanılmış Prim
Brüt Bileşik Oran = Brüt H/P + Masraf Oranı""",
        language="text",
    )
    st.write("Bu blokta teknik kâr/zarar veya mali gelir aktarımı gösterilmez.")

    st.markdown("### Teknik Sonuç Görünümü")
    st.code(
        """Mali Gelir Aktarımı (603) = Teknik Olmayan Bölümden Aktarılan Yatırım Gelirleri
Mali Gelir Aktarımı Hariç Teknik Sonuç = Raporlanan Teknik Kâr/Zarar - 603
Aktarım Dahil Teknik Sonuç = Raporlanan Teknik Kâr/Zarar""",
        language="text",
    )
    st.info(
        "Mali Gelir Aktarımı Hariç Teknik Sonuç bu analizde kullanılan analitik göstergedir. "
        "Brüt H/P, masraf oranı veya brüt bileşik oranın doğrudan muhasebesel karşılığı olarak sunulmaz; "
        "teknik sonuç daha geniş teknik hesap ve karşılık kalemlerini içerir. Bu nedenle iki görünüm uygulamada ayrı bloklarda gösterilir."
    )

    st.markdown("### v5 sunum ilkesi")
    st.write(
        "Brüt Teknik Görünüm ve Teknik Sonuç Görünümü grafik, KPI ve tablolarda ayrı tutulur. "
        "Teknik sonuç tablolarında brüt yazılan prim yalnızca hacim referansı olarak yer alabilir."
    )

    st.markdown("### Paylaşım sürümü kapsamı")
    st.write(
        "Bu sürüm salt okunurdur ve 2025H1–2026H1 dönemlerini içerir. Veri yükleme/güncelleme ekranı paylaşım sürümünde yer almaz. "
        "Analitik kapsam 18 ana branş ve hayat dışı (HD) şirketlerle sınırlıdır."
    )

    st.markdown("### Şirket sayfası sunum ilkesi")
    st.write("Şirket bazında yalnızca sayısal değerler, dönemsel değişimler ve benchmark farkları gösterilir. Niteliksel sınıflama veya hüküm üretilmez.")


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
        st.markdown(f"## {APP_TITLE}")
        st.caption("Paylaşım sürümü · 2025H1–2026H1 · salt okunur demo")
        page = st.radio("Görünüm", ["Sektör Özeti", "Branş Analizi", "Şirket Detayı", "Metodoloji"])
        st.divider()

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
            f"<div class='small-muted'>Dönemler: {', '.join(periods)}<br>Salt okunur paylaşım sürümü<br>Kaynak: Türkiye Sigorta Birliği (TSB) finansal tabloları.</div>",
            unsafe_allow_html=True,
        )

    if page == "Sektör Özeti":
        sector_page(history, previous, current)
    elif page == "Branş Analizi":
        branch_page(history, previous, current)
    elif page == "Şirket Detayı":
        company_page(history, previous, current)
    else:
        methodology_page()


if __name__ == "__main__":
    main()
