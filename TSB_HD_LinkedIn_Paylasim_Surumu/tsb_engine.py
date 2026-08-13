from __future__ import annotations

import csv
import hashlib
import io
import json
import math
import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

MAIN_BRANCHES = [
    "KAZA",
    "HASTALIK-SAĞLIK",
    "KARA ARAÇLARI",
    "RAYLI ARAÇLAR",
    "HAVA ARAÇLARI",
    "SU ARAÇLARI",
    "NAKLİYAT",
    "YANGIN VE DOĞAL AFETLER",
    "GENEL ZARARLAR",
    "KARA ARAÇLARI SORUMLULUK",
    "HAVA ARAÇLARI SORUMLULUK",
    "SU ARAÇLARI SORUMLULUK",
    "GENEL SORUMLULUK",
    "KREDİ",
    "KEFALET",
    "FİNANSAL KAYIPLAR",
    "HUKUKSAL KORUMA",
    "DESTEK",
]

REQUIRED_SHEETS = ["HAYATDISI"] + MAIN_BRANCHES

AMOUNT_METRICS = [
    "Brüt Yazılan Prim (TL)",
    "Brüt Kazanılmış Prim (TL)",
    "Brüt Gerçekleşen Hasar (TL)",
    "Faaliyet Gideri (TL, işaretli)",
    "Teknik Kâr/Zarar (TL)",
    "Yatırım Katkısı (TL)",
    "Yatırım Hariç Teknik Sonuç (TL)",
]

RATIO_METRICS = [
    "Brüt H/P",
    "Masraf Oranı",
    "Brüt Bileşik Oran",
    "Yatırım Hariç Teknik Marj",
    "Yatırım Bağımlılığı",
]

RECON_METRICS = [
    "Brüt Yazılan Prim (TL)",
    "Brüt Kazanılmış Prim (TL)",
    "Brüt Gerçekleşen Hasar (TL)",
    "Faaliyet Gideri (TL, işaretli)",
    "Teknik Kâr/Zarar (TL)",
    "Yatırım Katkısı (TL)",
]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def normalize_code(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):
        return str(int(float(text)))
    return text


def _canonical_account(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    return str(value).strip()


def _col_to_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 0
    idx = 0
    for ch in match.group(1):
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _to_number(value: str) -> Any:
    if value is None or value == "":
        return None
    try:
        if re.fullmatch(r"[-+]?\d+", value):
            return int(value)
        return float(value)
    except (TypeError, ValueError):
        return value


class XlsxRaw:
    """Minimal XLSX reader using only the standard library.

    It reads cached values from workbook XML and preserves source signs.
    This keeps the desktop app lightweight and avoids changing the TSB files.
    """

    def __init__(self, data: bytes):
        self.data = data
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            self.shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in z.namelist():
                root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                self.shared_strings = [
                    "".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t"))
                    for si in root.findall(f"{{{NS_MAIN}}}si")
                ]

            wb = ET.fromstring(z.read("xl/workbook.xml"))
            rel_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            relmap = {
                rel.attrib["Id"]: rel.attrib["Target"]
                for rel in rel_root.findall(f"{{{NS_PKG_REL}}}Relationship")
            }
            self.sheet_paths: dict[str, str] = {}
            for sh in wb.find(f"{{{NS_MAIN}}}sheets"):
                target = relmap[sh.attrib[f"{{{NS_REL}}}id"]].lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                self.sheet_paths[sh.attrib["name"]] = target

    @property
    def sheets(self) -> list[str]:
        return list(self.sheet_paths.keys())

    def resolve_sheet(self, requested: str) -> str:
        if requested in self.sheet_paths:
            return requested
        trimmed = {name.strip(): name for name in self.sheet_paths}
        if requested.strip() in trimmed:
            return trimmed[requested.strip()]
        raise KeyError(f"Sayfa bulunamadı: {requested}")

    def rows(self, sheet_name: str) -> list[tuple[int, dict[int, Any]]]:
        actual = self.resolve_sheet(sheet_name)
        with zipfile.ZipFile(io.BytesIO(self.data)) as z:
            root = ET.fromstring(z.read(self.sheet_paths[actual]))
            result: list[tuple[int, dict[int, Any]]] = []
            for row in root.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
                cells: dict[int, Any] = {}
                for c in row.findall(f"{{{NS_MAIN}}}c"):
                    idx = _col_to_index(c.attrib["r"])
                    cell_type = c.attrib.get("t")
                    inline = c.find(f"{{{NS_MAIN}}}is")
                    v = c.find(f"{{{NS_MAIN}}}v")
                    value: Any = None
                    if inline is not None:
                        value = "".join(t.text or "" for t in inline.iter(f"{{{NS_MAIN}}}t"))
                    elif v is not None:
                        raw = v.text or ""
                        if cell_type == "s" and raw:
                            value = self.shared_strings[int(raw)]
                        elif cell_type == "b":
                            value = raw == "1"
                        else:
                            value = _to_number(raw)
                    cells[idx] = value
                result.append((int(row.attrib["r"]), cells))
            return result


def _period_from_text(text: str) -> tuple[str, str]:
    match = re.search(
        r"(\d{2})\.(\d{2})\.(\d{4})\s*-\s*(\d{2})\.(\d{2})\.(\d{4})",
        text,
    )
    if not match:
        raise ValueError(f"Dönem metni algılanamadı: {text!r}")
    end = datetime(int(match.group(6)), int(match.group(5)), int(match.group(4)))
    suffix = {3: "Q1", 6: "H1", 9: "9M", 12: "FY"}.get(end.month, f"M{end.month:02d}")
    return f"{end.year}{suffix}", text.strip()


def detect_period(book: XlsxRaw) -> tuple[str, str]:
    rows = dict(book.rows("HAYATDISI"))
    candidates = [
        str(rows.get(3, {}).get(0, "") or ""),
        str(rows.get(2, {}).get(0, "") or ""),
        str(rows.get(4, {}).get(0, "") or ""),
    ]
    for text in candidates:
        try:
            return _period_from_text(text)
        except ValueError:
            continue
    raise ValueError("HAYATDISI sayfasında dönem bilgisi bulunamadı.")


def classify_workbook(book: XlsxRaw) -> str:
    names = {name.strip() for name in book.sheets}
    if "Gelir Tablosu" in names or "Branş Kırılımlı Analiz" in names:
        return "income"
    if "OZET" in names:
        return "expense"

    try:
        rows = dict(book.rows("HAYATDISI"))
        first_text = " ".join(
            str(rows.get(r, {}).get(0, "") or "") for r in (1, 2, 3)
        ).lower()
        if "hasar" in first_text and "prim" in first_text:
            return "claims"
        if "faaliyet" in first_text and "gider" in first_text:
            return "expense"
    except Exception:
        pass

    if {"GENEL", "HAYATDISI", "KAZA"}.issubset(names):
        return "claims"
    return "unknown"


def _required_sheet_status(book: XlsxRaw) -> tuple[bool, list[str]]:
    available = {name.strip() for name in book.sheets}
    missing = [sheet for sheet in REQUIRED_SHEETS if sheet not in available]
    return len(missing) == 0, missing


def _income_sheet_records(book: XlsxRaw, sheet_name: str) -> dict[str, dict[str, Any]]:
    rows = dict(book.rows(sheet_name))
    account_row = rows.get(5, {})
    header_row = rows.get(6, {})
    account_index = {
        _canonical_account(value): idx
        for idx, value in account_row.items()
        if _canonical_account(value)
    }

    def idx_for(account: str) -> int:
        key = str(account)
        if key not in account_index:
            raise KeyError(f"{sheet_name}: {account} hesap kodu bulunamadı.")
        return account_index[key]

    technical_idx: int | None = None
    for idx, value in header_row.items():
        text = str(value or "").strip().lower()
        if "teknik" in text and "kar" in text and "zarar" in text:
            technical_idx = idx
            break
    if technical_idx is None:
        raise KeyError(f"{sheet_name}: Teknik Kar Zarar sütunu bulunamadı.")

    needed_accounts = ["60001", "60003", "60101", "60103", "60201", "61001", "61101", "614", "603"]
    indexes = {acc: idx_for(acc) for acc in needed_accounts}

    output: dict[str, dict[str, Any]] = {}
    for row_no, cells in rows.items():
        if row_no <= 6:
            continue
        code = cells.get(1)
        company_type = str(cells.get(2, "") or "").strip()
        if code in (None, "") or company_type not in ("HD", "T (HD)"):
            continue
        code_key = normalize_code(code)

        def val(account: str) -> float:
            raw = cells.get(indexes[account])
            try:
                return float(raw or 0.0)
            except (TypeError, ValueError):
                return 0.0

        gross_written = val("60001")
        gross_earned = val("60001") + val("60003") + val("60101") + val("60103") + val("60201")
        gross_incurred = -(val("61001") + val("61101"))
        expense = val("614")
        technical_result = float(cells.get(technical_idx) or 0.0)
        investment = val("603")
        loss_ratio = gross_incurred / gross_earned if abs(gross_earned) > 1e-12 else 0.0
        expense_ratio = -expense / gross_earned if abs(gross_earned) > 1e-12 else 0.0
        combined = loss_ratio + expense_ratio
        underlying = technical_result - investment
        underlying_margin = underlying / gross_earned if abs(gross_earned) > 1e-12 else 0.0
        investment_dependency = investment / technical_result if abs(technical_result) > 1e-12 else 0.0

        output[code_key] = {
            "Şirket Kodu": int(code) if isinstance(code, (int, float)) and float(code).is_integer() else code,
            "Şirket Adı": str(cells.get(0, "") or "").strip(),
            "Şirket Tipi": company_type,
            "Brüt Yazılan Prim (TL)": gross_written,
            "Brüt Kazanılmış Prim (TL)": gross_earned,
            "Brüt Gerçekleşen Hasar (TL)": gross_incurred,
            "Faaliyet Gideri (TL, işaretli)": expense,
            "Teknik Kâr/Zarar (TL)": technical_result,
            "Yatırım Katkısı (TL)": investment,
            "Brüt H/P": loss_ratio,
            "Masraf Oranı": expense_ratio,
            "Brüt Bileşik Oran": combined,
            "Yatırım Hariç Teknik Sonuç (TL)": underlying,
            "Yatırım Hariç Teknik Marj": underlying_margin,
            "Yatırım Bağımlılığı": investment_dependency,
        }
    return output


def build_period_from_income(income_bytes: bytes) -> tuple[str, str, list[dict[str, Any]]]:
    book = XlsxRaw(income_bytes)
    ok, missing = _required_sheet_status(book)
    if not ok:
        raise ValueError("Gelir tablosu dosyasında eksik sayfalar: " + ", ".join(missing))

    period, period_text = detect_period(book)
    by_sheet = {sheet: _income_sheet_records(book, sheet) for sheet in REQUIRED_SHEETS}
    companies = {
        code: rec
        for code, rec in by_sheet["HAYATDISI"].items()
        if rec.get("Şirket Tipi") == "HD"
    }

    zero_metrics = {
        "Brüt Yazılan Prim (TL)": 0.0,
        "Brüt Kazanılmış Prim (TL)": 0.0,
        "Brüt Gerçekleşen Hasar (TL)": 0.0,
        "Faaliyet Gideri (TL, işaretli)": 0.0,
        "Teknik Kâr/Zarar (TL)": 0.0,
        "Yatırım Katkısı (TL)": 0.0,
        "Brüt H/P": 0.0,
        "Masraf Oranı": 0.0,
        "Brüt Bileşik Oran": 0.0,
        "Yatırım Hariç Teknik Sonuç (TL)": 0.0,
        "Yatırım Hariç Teknik Marj": 0.0,
        "Yatırım Bağımlılığı": 0.0,
    }

    rows: list[dict[str, Any]] = []
    for code in sorted(companies, key=lambda x: int(x) if x.isdigit() else x):
        base = companies[code]
        for branch in REQUIRED_SHEETS:
            src = by_sheet[branch].get(code)
            if src is None:
                src = {
                    "Şirket Kodu": base["Şirket Kodu"],
                    "Şirket Adı": base["Şirket Adı"],
                    "Şirket Tipi": "HD",
                    **zero_metrics,
                }
                source_status = "Branş kaydı yok / sıfır"
            else:
                source_status = "Kaynakta mevcut"
            rows.append(
                {
                    "Dönem": period,
                    "Şirket Kodu": src["Şirket Kodu"],
                    "Şirket Adı": src["Şirket Adı"],
                    "Şirket Tipi": src["Şirket Tipi"],
                    "Branş": branch,
                    "Kaynak Durumu": source_status,
                    **{k: v for k, v in src.items() if k not in {"Şirket Kodu", "Şirket Adı", "Şirket Tipi"}},
                }
            )

    for branch in REQUIRED_SHEETS:
        sector = by_sheet[branch].get("9000")
        if not sector or sector.get("Şirket Tipi") != "T (HD)":
            raise ValueError(f"{branch}: 9000 / T (HD) sektör toplamı bulunamadı.")
        rows.append(
            {
                "Dönem": period,
                "Şirket Kodu": sector["Şirket Kodu"],
                "Şirket Adı": sector["Şirket Adı"],
                "Şirket Tipi": sector["Şirket Tipi"],
                "Branş": branch,
                "Kaynak Durumu": "Sektör toplamı",
                **{k: v for k, v in sector.items() if k not in {"Şirket Kodu", "Şirket Adı", "Şirket Tipi"}},
            }
        )

    return period, period_text, rows


def _external_records(book: XlsxRaw, sheet_name: str, kind: str) -> dict[str, Any]:
    rows = dict(book.rows(sheet_name))
    account_row = rows.get(5, {})
    if kind == "claims":
        earned_idx = next(
            (idx for idx, value in account_row.items() if str(value or "").strip() == "Kazanılmış Prim (Brüt)"),
            None,
        )
        incurred_idx = next(
            (idx for idx, value in account_row.items() if str(value or "").strip() == "Gerçekleşen Hasar(Brüt)"),
            None,
        )
        if earned_idx is None or incurred_idx is None:
            raise KeyError(f"{sheet_name}: hasar/prim kontrol sütunları bulunamadı.")
    elif kind == "expense":
        expense_idx = next(
            (idx for idx, value in account_row.items() if _canonical_account(value) == "614"),
            None,
        )
        if expense_idx is None:
            raise KeyError(f"{sheet_name}: 614 faaliyet gideri sütunu bulunamadı.")
    else:
        raise ValueError(kind)

    output: dict[str, Any] = {}
    for row_no, cells in rows.items():
        if row_no <= 6:
            continue
        code = cells.get(1)
        company_type = str(cells.get(2, "") or "").strip()
        if code in (None, "") or company_type not in ("HD", "T (HD)"):
            continue
        key = normalize_code(code)
        if kind == "claims":
            output[key] = (
                float(cells.get(earned_idx) or 0.0),
                -float(cells.get(incurred_idx) or 0.0),
            )
        else:
            output[key] = float(cells.get(expense_idx) or 0.0)
    return output


def validate_cross_sources(
    master_rows: list[dict[str, Any]],
    claims_bytes: bytes,
    expense_bytes: bytes,
    tolerance_tl: float = 1.0,
) -> dict[str, Any]:
    claims_book = XlsxRaw(claims_bytes)
    expense_book = XlsxRaw(expense_bytes)
    master_map = {(normalize_code(r["Şirket Kodu"]), r["Branş"]): r for r in master_rows}

    max_earned_diff = 0.0
    max_claim_diff = 0.0
    max_expense_diff = 0.0
    compared = 0
    missing_material = 0
    failed = 0

    for branch in REQUIRED_SHEETS:
        claim_records = _external_records(claims_book, branch, "claims")
        expense_records = _external_records(expense_book, branch, "expense")
        for (code, row_branch), row in master_map.items():
            if row_branch != branch:
                continue
            if code not in claim_records or code not in expense_records:
                material = any(
                    abs(float(row.get(metric) or 0.0)) > tolerance_tl
                    for metric in (
                        "Brüt Kazanılmış Prim (TL)",
                        "Brüt Gerçekleşen Hasar (TL)",
                        "Faaliyet Gideri (TL, işaretli)",
                    )
                )
                if material:
                    missing_material += 1
                continue

            earned_ext, claim_ext = claim_records[code]
            expense_ext = expense_records[code]
            earned_diff = abs(float(row["Brüt Kazanılmış Prim (TL)"]) - earned_ext)
            claim_diff = abs(float(row["Brüt Gerçekleşen Hasar (TL)"]) - claim_ext)
            expense_diff = abs(float(row["Faaliyet Gideri (TL, işaretli)"]) - expense_ext)
            max_earned_diff = max(max_earned_diff, earned_diff)
            max_claim_diff = max(max_claim_diff, claim_diff)
            max_expense_diff = max(max_expense_diff, expense_diff)
            compared += 1
            if max(earned_diff, claim_diff, expense_diff) > tolerance_tl:
                failed += 1

    passed = missing_material == 0 and failed == 0
    return {
        "passed": passed,
        "compared_records": compared,
        "failed_records": failed,
        "missing_material_records": missing_material,
        "max_earned_premium_diff_tl": max_earned_diff,
        "max_gross_incurred_diff_tl": max_claim_diff,
        "max_expense_diff_tl": max_expense_diff,
        "tolerance_tl": tolerance_tl,
    }


def validate_reconciliation(master_rows: list[dict[str, Any]], tolerance_tl: float = 1.0) -> dict[str, Any]:
    by_code: dict[str, list[dict[str, Any]]] = {}
    for row in master_rows:
        by_code.setdefault(normalize_code(row["Şirket Kodu"]), []).append(row)

    max_diff = 0.0
    failed_codes: list[str] = []
    details: list[dict[str, Any]] = []
    for code, rows in by_code.items():
        total = next((r for r in rows if r["Branş"] == "HAYATDISI"), None)
        branches = [r for r in rows if r["Branş"] in MAIN_BRANCHES]
        if total is None or len(branches) != len(MAIN_BRANCHES):
            failed_codes.append(code)
            continue

        metric_diffs = {}
        for metric in RECON_METRICS:
            branch_sum = sum(float(r.get(metric) or 0.0) for r in branches)
            total_value = float(total.get(metric) or 0.0)
            diff = branch_sum - total_value
            metric_diffs[metric] = diff
            max_diff = max(max_diff, abs(diff))
        code_max = max(abs(v) for v in metric_diffs.values()) if metric_diffs else 0.0
        if code_max > tolerance_tl:
            failed_codes.append(code)
        if code == "9000" or code_max > tolerance_tl:
            details.append({"Şirket Kodu": code, "Maksimum Mutlak Fark (TL)": code_max, **metric_diffs})

    return {
        "passed": len(failed_codes) == 0,
        "checked_codes": len(by_code),
        "failed_codes": failed_codes,
        "max_abs_diff_tl": max_diff,
        "tolerance_tl": tolerance_tl,
        "details": details,
    }


def prepare_import(files: Iterable[tuple[str, bytes]], tolerance_tl: float = 1.0) -> dict[str, Any]:
    classified: dict[str, tuple[str, bytes, XlsxRaw]] = {}
    unknown: list[str] = []
    duplicates: list[str] = []

    for name, content in files:
        book = XlsxRaw(content)
        role = classify_workbook(book)
        if role == "unknown":
            unknown.append(name)
            continue
        if role in classified:
            duplicates.append(role)
            continue
        classified[role] = (name, content, book)

    if unknown:
        raise ValueError("Dosya tipi algılanamadı: " + ", ".join(unknown))
    if duplicates:
        raise ValueError("Aynı tipten birden fazla dosya algılandı: " + ", ".join(sorted(set(duplicates))))
    missing_roles = [role for role in ("income", "claims", "expense") if role not in classified]
    if missing_roles:
        role_names = {"income": "Gelir Tablosu", "claims": "Hasar-Prim", "expense": "Faaliyet Giderleri"}
        raise ValueError("Eksik dosya: " + ", ".join(role_names[r] for r in missing_roles))

    role_sheet_checks = {}
    period_info = {}
    for role, (name, _, book) in classified.items():
        ok, missing_sheets = _required_sheet_status(book)
        role_sheet_checks[role] = {"passed": ok, "missing_sheets": missing_sheets, "file_name": name}
        period, raw_period = detect_period(book)
        period_info[role] = {"period": period, "raw": raw_period, "file_name": name}

    periods = {info["period"] for info in period_info.values()}
    period_consistent = len(periods) == 1
    if not period_consistent:
        raise ValueError("Üç dosyanın dönemleri aynı değil: " + ", ".join(sorted(periods)))
    if not all(check["passed"] for check in role_sheet_checks.values()):
        missing_text = []
        for role, check in role_sheet_checks.items():
            if not check["passed"]:
                missing_text.append(f"{role}: {', '.join(check['missing_sheets'])}")
        raise ValueError("Ana branş sayfaları eksik: " + " | ".join(missing_text))

    income_name, income_bytes, _ = classified["income"]
    _, claims_bytes, _ = classified["claims"]
    _, expense_bytes, _ = classified["expense"]
    period, raw_period, master_rows = build_period_from_income(income_bytes)
    cross = validate_cross_sources(master_rows, claims_bytes, expense_bytes, tolerance_tl=tolerance_tl)
    recon = validate_reconciliation(master_rows, tolerance_tl=tolerance_tl)

    company_count = len(
        {
            normalize_code(row["Şirket Kodu"])
            for row in master_rows
            if row.get("Şirket Tipi") == "HD" and row.get("Branş") == "HAYATDISI"
        }
    )
    sector_total = next(
        row
        for row in master_rows
        if normalize_code(row["Şirket Kodu"]) == "9000" and row["Branş"] == "HAYATDISI"
    )

    source_roles = {
        role: {
            "file_name": classified[role][0],
            "sha256": hashlib.sha256(classified[role][1]).hexdigest(),
        }
        for role in classified
    }

    return {
        "period": period,
        "raw_period": raw_period,
        "rows": master_rows,
        "company_count": company_count,
        "record_count": len(master_rows),
        "sector_total": sector_total,
        "source_roles": source_roles,
        "period_info": period_info,
        "sheet_checks": role_sheet_checks,
        "cross_validation": cross,
        "reconciliation": recon,
        "passed": cross["passed"] and recon["passed"] and period_consistent,
    }


def period_key(label: str) -> tuple[int, int]:
    match = re.match(r"^(\d{4})(Q1|H1|9M|FY|M\d{2})$", str(label))
    if not match:
        return (0, 0)
    year = int(match.group(1))
    suffix = match.group(2)
    rank = {"Q1": 3, "H1": 6, "9M": 9, "FY": 12}.get(suffix)
    if rank is None and suffix.startswith("M"):
        rank = int(suffix[1:])
    return year, int(rank or 0)


def sorted_periods(rows: Iterable[dict[str, Any]]) -> list[str]:
    return sorted({str(r.get("Dönem", "")) for r in rows if r.get("Dönem")}, key=period_key)


def load_history(initial_path: Path, active_path: Path) -> tuple[list[dict[str, Any]], str]:
    path = active_path if active_path.exists() else initial_path
    payload = json.loads(path.read_text(encoding="utf-8"))
    rows = payload["rows"] if isinstance(payload, dict) else payload
    return rows, ("active" if path == active_path else "initial")


def save_history(active_path: Path, rows: list[dict[str, Any]], source: str = "TSB HD Streamlit v4") -> None:
    active_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"schema_version": 1, "source": source, "updated_at": datetime.now().isoformat(timespec="seconds"), "rows": rows}
    _atomic_write_text(active_path, json.dumps(payload, ensure_ascii=False, separators=(",", ":")))


def merge_period(history_rows: list[dict[str, Any]], new_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not new_rows:
        return history_rows
    period = str(new_rows[0]["Dönem"])
    merged = [row for row in history_rows if str(row.get("Dönem")) != period]
    merged.extend(new_rows)
    merged.sort(key=lambda r: (period_key(str(r.get("Dönem", ""))), 1 if normalize_code(r.get("Şirket Kodu")) == "9000" else 0, normalize_code(r.get("Şirket Kodu")), REQUIRED_SHEETS.index(r.get("Branş")) if r.get("Branş") in REQUIRED_SHEETS else 999))
    return merged


def _atomic_write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=path.name + ".", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(text)
        os.replace(temp_name, path)
    finally:
        if os.path.exists(temp_name):
            os.remove(temp_name)


def append_update_log(log_path: Path, entry: dict[str, Any]) -> None:
    entries: list[dict[str, Any]] = []
    if log_path.exists():
        try:
            entries = json.loads(log_path.read_text(encoding="utf-8"))
        except Exception:
            entries = []
    entries.append(entry)
    _atomic_write_text(log_path, json.dumps(entries, ensure_ascii=False, indent=2))


def read_update_log(log_path: Path) -> list[dict[str, Any]]:
    if not log_path.exists():
        return []
    try:
        data = json.loads(log_path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def archive_sources(import_root: Path, period: str, files: Iterable[tuple[str, bytes]]) -> dict[str, str]:
    import_dir = import_root / period
    import_dir.mkdir(parents=True, exist_ok=True)
    saved: dict[str, str] = {}
    role_names = {"income": "Gelir_Tablosu.xlsx", "claims": "Hasar_Prim.xlsx", "expense": "Faaliyet_Giderleri.xlsx"}
    for original_name, content in files:
        role = classify_workbook(XlsxRaw(content))
        if role not in role_names:
            continue
        target = import_dir / role_names[role]
        target.write_bytes(content)
        saved[role] = str(target)
        (import_dir / f"{role}_original_name.txt").write_text(original_name, encoding="utf-8")
    return saved


def rows_to_csv_bytes(rows: list[dict[str, Any]]) -> bytes:
    if not rows:
        return b""
    fieldnames: list[str] = []
    seen = set()
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.add(key)
                fieldnames.append(key)
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue().encode("utf-8-sig")


def history_to_json_bytes(rows: list[dict[str, Any]]) -> bytes:
    payload = {"schema_version": 1, "exported_at": datetime.now().isoformat(timespec="seconds"), "rows": rows}
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def get_record(rows: Iterable[dict[str, Any]], period: str, company_code: Any, branch: str) -> dict[str, Any] | None:
    code = normalize_code(company_code)
    for row in rows:
        if str(row.get("Dönem")) == str(period) and normalize_code(row.get("Şirket Kodu")) == code and row.get("Branş") == branch:
            return row
    return None


def records_for_period_branch(rows: Iterable[dict[str, Any]], period: str, branch: str, include_sector: bool = False) -> list[dict[str, Any]]:
    result = []
    for row in rows:
        if str(row.get("Dönem")) != str(period) or row.get("Branş") != branch:
            continue
        is_sector = normalize_code(row.get("Şirket Kodu")) == "9000"
        if include_sector or not is_sector:
            result.append(row)
    return result


def safe_growth(current: Any, previous: Any) -> float | None:
    try:
        c = float(current or 0.0)
        p = float(previous or 0.0)
    except (TypeError, ValueError):
        return None
    if abs(p) < 1e-12:
        return None
    return c / p - 1.0


def safe_delta(current: Any, previous: Any) -> float:
    try:
        return float(current or 0.0) - float(previous or 0.0)
    except (TypeError, ValueError):
        return 0.0


def safe_pp(current: Any, previous: Any) -> float:
    return safe_delta(current, previous) * 100.0


def activity(row: dict[str, Any] | None) -> bool:
    if not row:
        return False
    return any(abs(float(row.get(metric) or 0.0)) > 0.5 for metric in AMOUNT_METRICS)


def bubble_sizes(values: Iterable[Any], min_size: float = 10.0, max_size: float = 42.0) -> list[float]:
    vals = [max(0.0, float(v or 0.0)) for v in values]
    if not vals:
        return []
    roots = [math.sqrt(v) for v in vals]
    lo, hi = min(roots), max(roots)
    if hi <= lo:
        return [(min_size + max_size) / 2.0 for _ in roots]
    return [min_size + (x - lo) / (hi - lo) * (max_size - min_size) for x in roots]
