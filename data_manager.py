"""
Data Management & Excel Import/Export Module for Weekly Automation Report App.

Responsibilities:
- Parse existing Excel (.xlsx) files into a structured JSON model.
- Persist weekly reports safely in `data/reports/<report_id>.json`.
- Automatically initialize from the local Excel file if no saved reports exist yet.
- Calculate formulas (Weekly Totals, Cumulative Totals with Medium-Low, Rates, Total Automated).
- Export structured report data back to an original-styled `.xlsx` workbook.
"""

from datetime import datetime, date
import io
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from security import is_editor_authenticated, safe_report_id, sanitize_excel_value, sanitize_text

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data" / "reports"

DEFAULT_SHEET1_NOTE = (
    "Total Maintained: Uygulama değişiklikleri sebebiyle bakımı yapılan otomasyon senaryo sayısıdır."
)
DEFAULT_SHEET1_WARNING = (
    "⚠️ Dikkat:\n"
    "1- Satır/sütun ölçülerini değiştirmeyelim.\n"
    "2- Sağ tablo, sol tablodan formülle hesaplanır.\n"
    "3- Göndermeden önce son bir kontrol yapmayı unutmayalım."
)
DEFAULT_SHEET2_WARNING = (
    '⚠️ Dikkat: Aşağıdaki sayılar "Total" satırına formülle ekleniyor, mail tablosuna eklemeyelim.'
)
DEFAULT_SHEET3_LEGEND = (
    "(X) Test otomasyon koşumunu direkt etkilemekte veya test yazım sürecini engellemektedir.\n"
    "(Y) Test otomasyon koşumunda veya test yazım sürecinde efor kaybına neden olmaktadır.\n"
    "(Z) Raporlama, analiz süreçlerinde soruna neden olmaktadır.\n"
    "(-) Minör bulgulara ait olup koşum ve senaryo yazımlarını etkilememektedir."
)

SHEET4_TYPE_ORDER = {
    "defect": 0,
    "bug": 1,
    "story": 2,
}


def sort_sheet4_items(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Guarantees Sheet 4 items are always ordered as:
    1. Defect -> 2. Bug -> 3. Story (preserving relative order within each group).
    """
    return sorted(
        items,
        key=lambda item: SHEET4_TYPE_ORDER.get(str(item.get("Type", "")).strip().lower(), 99),
    )


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def _to_int(val: Any, default: int = 0) -> int:
    if val is None or val == "":
        return default
    try:
        return int(round(float(val)))
    except Exception:
        return default


def _format_date_str(val: Any) -> str:
    if val is None or val == "":
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    text = str(val).strip()
    if " " in text and len(text) >= 10:
        text = text[:10]
    return sanitize_text(text, max_len=20)


def get_default_template_data() -> Dict[str, Any]:
    """Returns a clean template structure if no Excel file or saved JSON is present."""
    return {
        "meta": {
            "report_id": "w43_2026",
            "title": "w43 - Otomasyon Raporu",
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_by": "Sistem",
        },
        "sheet1": {
            "progress_time": "19 Oct - 23 Oct",
            "auto_calc_total_automated": True,
            "total_automated_manual": 66,
            "im_android": 10,
            "im_ios": 10,
            "voip_android": 7,
            "voip_ios": 0,
            "found_bugs": 2,
            "found_defects": 0,
            "note": DEFAULT_SHEET1_NOTE,
            "warning": DEFAULT_SHEET1_WARNING,
        },
        "sheet2": {
            "weekly_im": {
                "ios_gating": 3,
                "ios_high": 17,
                "android_gating": 28,
                "android_high": 11,
            },
            "weekly_voip": {
                "ios_gating": 0,
                "ios_high": 0,
                "android_gating": 2,
                "android_high": 5,
            },
            "automated_im": {
                "ios_gating": 503,
                "ios_high": 733,
                "ios_to_be": 11408,
                "android_gating": 532,
                "android_high": 705,
                "android_to_be": 11361,
            },
            "automated_voip": {
                "ios_gating": 0,
                "ios_high": 0,
                "ios_to_be": 0,
                "android_gating": 583,
                "android_high": 222,
                "android_to_be": 2284,
            },
            "medium_low": {
                "im_ios": 75,
                "im_android": 82,
                "voip_ios": 0,
                "voip_android": 32,
            },
            "warning": DEFAULT_SHEET2_WARNING,
        },
        "sheet3": {
            "title": "Problems & Requirements",
            "items": [],
            "legend": DEFAULT_SHEET3_LEGEND,
        },
        "sheet4": {
            "items": [],
            "maintenance_title": "Güncellemeler Sonrası Otomasyon Bakımları",
            "auto_sync_maintenance_count": True,
            "maintenance_items": [
                {
                    "Platform": "iOS & Android",
                    "Bakım Nedeni": "Release yükü kapsamında, arayüzde yapılan geliştirmeler doğrultusunda akış ve elementlerde güncellemeler yapılmıştır.",
                    "Adet": 27,
                }
            ],
        },
    }


def compute_metrics(report_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Computes all formula-based metrics across Sheet1, Sheet2, and Sheet4 so they are
    always consistent in UI previews, email tables, and Excel exports.
    """
    s1 = report_data.get("sheet1", {})
    s2 = report_data.get("sheet2", {})
    s4 = report_data.get("sheet4", {})

    w_im = s2.get("weekly_im", {})
    w_voip = s2.get("weekly_voip", {})
    a_im = s2.get("automated_im", {})
    a_voip = s2.get("automated_voip", {})
    ml = s2.get("medium_low", {})

    weekly_im_total = (
        _to_int(w_im.get("ios_gating"))
        + _to_int(w_im.get("ios_high"))
        + _to_int(w_im.get("android_gating"))
        + _to_int(w_im.get("android_high"))
    )
    weekly_voip_total = (
        _to_int(w_voip.get("ios_gating"))
        + _to_int(w_voip.get("ios_high"))
        + _to_int(w_voip.get("android_gating"))
        + _to_int(w_voip.get("android_high"))
    )

    im_ios_total = (
        _to_int(a_im.get("ios_gating"))
        + _to_int(a_im.get("ios_high"))
        + _to_int(ml.get("im_ios"))
    )
    im_android_total = (
        _to_int(a_im.get("android_gating"))
        + _to_int(a_im.get("android_high"))
        + _to_int(ml.get("im_android"))
    )
    voip_ios_total = (
        _to_int(a_voip.get("ios_gating"))
        + _to_int(a_voip.get("ios_high"))
        + _to_int(ml.get("voip_ios"))
    )
    voip_android_total = (
        _to_int(a_voip.get("android_gating"))
        + _to_int(a_voip.get("android_high"))
        + _to_int(ml.get("voip_android"))
    )

    im_ios_to_be = _to_int(a_im.get("ios_to_be"))
    im_android_to_be = _to_int(a_im.get("android_to_be"))
    voip_ios_to_be = _to_int(a_voip.get("ios_to_be"))
    voip_android_to_be = _to_int(a_voip.get("android_to_be"))

    im_ios_rate = (im_ios_total / im_ios_to_be) if im_ios_to_be > 0 else 0.0
    im_android_rate = (im_android_total / im_android_to_be) if im_android_to_be > 0 else 0.0
    voip_ios_rate = (voip_ios_total / voip_ios_to_be) if voip_ios_to_be > 0 else 0.0
    voip_android_rate = (voip_android_total / voip_android_to_be) if voip_android_to_be > 0 else 0.0

    total_automated_calc = weekly_im_total + weekly_voip_total
    if s1.get("auto_calc_total_automated", True):
        total_automated = total_automated_calc
    else:
        total_automated = _to_int(s1.get("total_automated_manual"), total_automated_calc)

    total_maintained_sum = (
        _to_int(s1.get("im_android"))
        + _to_int(s1.get("im_ios"))
        + _to_int(s1.get("voip_android"))
        + _to_int(s1.get("voip_ios"))
    )

    return {
        "weekly_im_total": weekly_im_total,
        "weekly_voip_total": weekly_voip_total,
        "total_automated_calc": total_automated_calc,
        "total_automated": total_automated,
        "total_maintained_sum": total_maintained_sum,
        "im_ios_total": im_ios_total,
        "im_android_total": im_android_total,
        "voip_ios_total": voip_ios_total,
        "voip_android_total": voip_android_total,
        "im_ios_rate": im_ios_rate,
        "im_android_rate": im_android_rate,
        "voip_ios_rate": voip_ios_rate,
        "voip_android_rate": voip_android_rate,
        "im_ios_rate_str": f"{round(im_ios_rate * 100)}%",
        "im_android_rate_str": f"{round(im_android_rate * 100)}%",
        "voip_ios_rate_str": f"{round(voip_ios_rate * 100)}%",
        "voip_android_rate_str": f"{round(voip_android_rate * 100)}%",
    }


def parse_excel_file(file_source: Any, report_id: str = "w43_2026", title: str = "w43 - Otomasyon Raporu") -> Dict[str, Any]:
    """
    Parses an Excel workbook (path or file-like object) matching the 4-sheet structure
    into a clean dictionary.
    """
    wb = openpyxl.load_workbook(file_source, data_only=True)
    data = get_default_template_data()
    data["meta"]["report_id"] = safe_report_id(report_id)
    data["meta"]["title"] = sanitize_text(title, max_len=80)
    data["meta"]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    sheet_names = wb.sheetnames

    # --- Sheet 1 ---
    if len(sheet_names) >= 1:
        ws1 = wb[sheet_names[0]]
        data["sheet1"]["progress_time"] = sanitize_text(ws1["B2"].value or "19 Oct - 23 Oct", max_len=60)
        data["sheet1"]["total_automated_manual"] = _to_int(ws1["B3"].value, 66)
        data["sheet1"]["im_android"] = _to_int(ws1["B6"].value, 0)
        data["sheet1"]["im_ios"] = _to_int(ws1["C6"].value, 0)
        data["sheet1"]["voip_android"] = _to_int(ws1["D6"].value, 0)
        data["sheet1"]["voip_ios"] = _to_int(ws1["E6"].value, 0)
        data["sheet1"]["found_bugs"] = _to_int(ws1["B7"].value, 0)
        data["sheet1"]["found_defects"] = _to_int(ws1["B8"].value, 0)
        if ws1["A9"].value:
            data["sheet1"]["note"] = sanitize_text(ws1["A9"].value, max_len=300)
        if ws1["A13"].value:
            data["sheet1"]["warning"] = sanitize_text(ws1["A13"].value, max_len=500)

    # --- Sheet 2 ---
    if len(sheet_names) >= 2:
        ws2 = wb[sheet_names[1]]
        data["sheet2"]["weekly_im"] = {
            "ios_gating": _to_int(ws2["C4"].value, 0),
            "ios_high": _to_int(ws2["C5"].value, 0),
            "android_gating": _to_int(ws2["E4"].value, 0),
            "android_high": _to_int(ws2["E5"].value, 0),
        }
        data["sheet2"]["weekly_voip"] = {
            "ios_gating": _to_int(ws2["H4"].value, 0),
            "ios_high": _to_int(ws2["H5"].value, 0),
            "android_gating": _to_int(ws2["J4"].value, 0),
            "android_high": _to_int(ws2["J5"].value, 0),
        }
        data["sheet2"]["automated_im"] = {
            "ios_gating": _to_int(ws2["C10"].value, 0),
            "ios_high": _to_int(ws2["C11"].value, 0),
            "ios_to_be": _to_int(ws2["C14"].value, 0),
            "android_gating": _to_int(ws2["E10"].value, 0),
            "android_high": _to_int(ws2["E11"].value, 0),
            "android_to_be": _to_int(ws2["E14"].value, 0),
        }
        data["sheet2"]["automated_voip"] = {
            "ios_gating": _to_int(ws2["H10"].value, 0),
            "ios_high": _to_int(ws2["H11"].value, 0),
            "ios_to_be": _to_int(ws2["H14"].value, 0),
            "android_gating": _to_int(ws2["J10"].value, 0),
            "android_high": _to_int(ws2["J11"].value, 0),
            "android_to_be": _to_int(ws2["J14"].value, 0),
        }
        data["sheet2"]["medium_low"] = {
            "im_ios": _to_int(ws2["C21"].value, 0),
            "im_android": _to_int(ws2["E21"].value, 0),
            "voip_ios": _to_int(ws2["H21"].value, 0),
            "voip_android": _to_int(ws2["J21"].value, 0),
        }
        if ws2["B18"].value:
            data["sheet2"]["warning"] = sanitize_text(ws2["B18"].value, max_len=300)

    # --- Sheet 3 (Problems & Requirements) ---
    if len(sheet_names) >= 3:
        ws3 = wb[sheet_names[2]]
        if ws3["B2"].value:
            data["sheet3"]["title"] = sanitize_text(ws3["B2"].value, max_len=120)
        items_s3: List[Dict[str, str]] = []
        for r in range(4, ws3.max_row + 1):
            b_val = ws3.cell(r, 2).value
            c_val = ws3.cell(r, 3).value
            d_val = ws3.cell(r, 4).value
            e_val = ws3.cell(r, 5).value
            f_val = ws3.cell(r, 6).value
            if b_val is None and c_val is None and d_val is None:
                continue
            b_str = str(b_val or "").strip()
            # Check if we hit the legend row "(X) Test otomasyon..."
            if b_str.startswith("(X)") or "Test otomasyon koşumunu" in b_str:
                data["sheet3"]["legend"] = sanitize_text(b_str, max_len=800)
                break
            items_s3.append(
                {
                    "Type": sanitize_text(b_val, max_len=40),
                    "Issue": sanitize_text(c_val, max_len=40),
                    "Summary": sanitize_text(str(d_val or "").strip(), max_len=400),
                    "Status": sanitize_text(e_val, max_len=40),
                    "Etki Analizi": sanitize_text(f_val, max_len=10),
                }
            )
        data["sheet3"]["items"] = items_s3

    # --- Sheet 4 (Defects/Bugs/Stories & Maintenance) ---
    if len(sheet_names) >= 4:
        ws4 = wb[sheet_names[3]]
        items_s4: List[Dict[str, str]] = []
        maint_items: List[Dict[str, Any]] = []
        in_maintenance_section = False

        for r in range(3, ws4.max_row + 1):
            b_val = ws4.cell(r, 2).value
            c_val = ws4.cell(r, 3).value
            d_val = ws4.cell(r, 4).value
            e_val = ws4.cell(r, 5).value
            f_val = ws4.cell(r, 6).value
            g_val = ws4.cell(r, 7).value
            h_val = ws4.cell(r, 8).value

            if b_val is None and c_val is None and d_val is None and h_val is None:
                continue

            b_str = str(b_val or "").strip()
            if "Güncellemeler Sonrası Otomasyon Bakımları" in b_str:
                data["sheet4"]["maintenance_title"] = sanitize_text(b_str, max_len=120)
                in_maintenance_section = True
                continue

            if not in_maintenance_section:
                items_s4.append(
                    {
                        "Type": sanitize_text(b_val, max_len=30),
                        "Issue": sanitize_text(c_val, max_len=40),
                        "Summary": sanitize_text(str(d_val or "").strip(), max_len=400),
                        "Status": sanitize_text(e_val, max_len=40),
                        "Platform": sanitize_text(f_val, max_len=30),
                        "Created": _format_date_str(g_val),
                        "Priority": sanitize_text(h_val, max_len=30),
                    }
                )
            else:
                if b_str.lower() == "platform" and str(c_val or "").strip().lower().startswith("bakım"):
                    continue
                maint_items.append(
                    {
                        "Platform": sanitize_text(b_val, max_len=50),
                        "Bakım Nedeni": sanitize_text(c_val, max_len=500),
                        "Adet": _to_int(h_val, 0),
                    }
                )

        data["sheet4"]["items"] = sort_sheet4_items(items_s4)
        if maint_items:
            data["sheet4"]["maintenance_items"] = maint_items

    return data


def save_report(
    report_id: str,
    report_data: Dict[str, Any],
    updated_by: str = "QA Kullanıcısı",
    system_init: bool = False,
) -> str:
    """
    Saves the given report dictionary to `data/reports/<safe_id>.json`.
    Enforces server-side RBAC check (`is_editor_authenticated()`) unless `system_init=True`.
    """
    if not system_init and not is_editor_authenticated():
        raise PermissionError("Bu işlem için Düzenleme Yetkisi (Editör Girişi) gereklidir.")

    ensure_data_dir()
    clean_id = safe_report_id(report_id)
    if "sheet4" in report_data and "items" in report_data["sheet4"]:
        report_data["sheet4"]["items"] = sort_sheet4_items(report_data["sheet4"]["items"])
    report_data.setdefault("meta", {})
    report_data["meta"]["report_id"] = clean_id
    report_data["meta"]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_data["meta"]["updated_by"] = sanitize_text(updated_by, max_len=50)

    target_path = DATA_DIR / f"{clean_id}.json"
    tmp_path = DATA_DIR / f"{clean_id}.json.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(report_data, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, target_path)
    return clean_id


def load_report(report_id: str) -> Optional[Dict[str, Any]]:
    """Loads a report dictionary by ID from `data/reports/<safe_id>.json`."""
    ensure_data_dir()
    clean_id = safe_report_id(report_id)
    target_path = DATA_DIR / f"{clean_id}.json"
    if not target_path.exists():
        return None
    with open(target_path, "r", encoding="utf-8") as f:
        return json.load(f)


def list_reports() -> List[Dict[str, str]]:
    """Returns a list of saved reports sorted newest first."""
    ensure_data_dir()
    reports: List[Dict[str, str]] = []
    for p in sorted(DATA_DIR.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
        try:
            with open(p, "r", encoding="utf-8") as f:
                raw = json.load(f)
            meta = raw.get("meta", {})
            progress_time = raw.get("sheet1", {}).get("progress_time", "")
            reports.append(
                {
                    "report_id": meta.get("report_id", p.stem),
                    "title": meta.get("title", p.stem),
                    "progress_time": progress_time,
                    "updated_at": meta.get("updated_at", ""),
                    "updated_by": meta.get("updated_by", ""),
                }
            )
        except Exception:
            continue
    return reports


def initialize_default_report_if_needed() -> str:
    """
    If no report exists in `data/reports/`, looks for the initial `.xlsx` file in the workspace
    and seeds `w43_2026.json`. Falls back to the clean template if no `.xlsx` is found.
    """
    ensure_data_dir()
    existing = list_reports()
    if existing:
        return existing[0]["report_id"]

    xlsx_files = sorted(BASE_DIR.glob("*.xlsx"))
    if xlsx_files:
        first_xlsx = xlsx_files[0]
        parsed = parse_excel_file(
            first_xlsx,
            report_id="w43_2026",
            title="w43 (19 Oct - 23 Oct) Otomasyon Raporu",
        )
        return save_report("w43_2026", parsed, updated_by="İlk Kurulum (Excel)", system_init=True)

    default_data = get_default_template_data()
    return save_report("w43_2026", default_data, updated_by="İlk Kurulum", system_init=True)


def create_new_week_report(
    base_report: Dict[str, Any],
    new_report_id: str,
    new_title: str,
    new_progress_time: str,
    rollover_weekly_to_cumulative: bool = True,
    updated_by: str = "QA Kullanıcısı",
) -> str:
    """
    Creates a new week's report cloned from `base_report`.
    If `rollover_weekly_to_cumulative` is True:
    - Adds current `weekly_im` Gating/High counts into `automated_im` Gating/High counts (if needed for next week baseline)
      and resets `weekly_im` / `weekly_voip` to 0 for the fresh week.
    """
    cloned = json.loads(json.dumps(base_report))
    clean_id = safe_report_id(new_report_id)
    cloned["meta"] = {
        "report_id": clean_id,
        "title": sanitize_text(new_title, max_len=80),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "updated_by": sanitize_text(updated_by, max_len=50),
    }
    cloned["sheet1"]["progress_time"] = sanitize_text(new_progress_time, max_len=60)

    if rollover_weekly_to_cumulative:
        # Reset weekly counters for the new week
        cloned["sheet2"]["weekly_im"] = {
            "ios_gating": 0,
            "ios_high": 0,
            "android_gating": 0,
            "android_high": 0,
        }
        cloned["sheet2"]["weekly_voip"] = {
            "ios_gating": 0,
            "ios_high": 0,
            "android_gating": 0,
            "android_high": 0,
        }
        cloned["sheet1"]["im_android"] = 0
        cloned["sheet1"]["im_ios"] = 0
        cloned["sheet1"]["voip_android"] = 0
        cloned["sheet1"]["voip_ios"] = 0
        cloned["sheet1"]["found_bugs"] = 0
        cloned["sheet1"]["found_defects"] = 0

    return save_report(clean_id, cloned, updated_by=updated_by)


def export_to_excel_bytes(report_data: Dict[str, Any]) -> bytes:
    """
    Exports `report_data` to an `.xlsx` byte stream with the exact layout,
    colors, formulas, and merged cells of the original workbook.
    """
    wb = openpyxl.Workbook()
    metrics = compute_metrics(report_data)

    # Common styles matching original Excel theme
    thin_side = Side(style="thin", color="000000")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    fill_purple = PatternFill("solid", fgColor="E5DFEC")   # IM header / Maintenance header
    fill_gray = PatternFill("solid", fgColor="F2F2F2")     # VoIP header / Note
    fill_yellow = PatternFill("solid", fgColor="FFFF99")   # Warning boxes
    fill_ios_1 = PatternFill("solid", fgColor="92CDDC")    # iOS Gating
    fill_ios_2 = PatternFill("solid", fgColor="B7DEE8")    # iOS High & Sheet3 Title/Legend
    fill_ios_3 = PatternFill("solid", fgColor="DAEEF3")    # iOS Med-Low & Sheet3 Headers
    fill_and_1 = PatternFill("solid", fgColor="C4D79B")    # Android Gating
    fill_and_2 = PatternFill("solid", fgColor="D8E4BC")    # Android High
    fill_and_3 = PatternFill("solid", fgColor="EBF1DE")    # Android Med-Low
    fill_org_1 = PatternFill("solid", fgColor="FABF8F")    # Total
    fill_org_2 = PatternFill("solid", fgColor="FCD5B4")    # Rate
    fill_org_3 = PatternFill("solid", fgColor="FDE9D9")    # To Be & Sheet4 Headers

    font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
    font_reg = Font(name="Calibri", size=11, bold=False, color="000000")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def style_range(ws, cell_range: str, border=border_all, fill=None, font=None, alignment=None):
        for row in ws[cell_range]:
            for cell in row:
                if border:
                    cell.border = border
                if fill:
                    cell.fill = fill
                if font:
                    cell.font = font
                if alignment:
                    cell.alignment = alignment

    # =========================================================================
    # SHEET 1
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Sheet1"
    s1 = report_data.get("sheet1", {})

    ws1["A2"] = "Progress Time"
    ws1["B2"] = sanitize_excel_value(s1.get("progress_time", ""))
    ws1.merge_cells("B2:E2")

    ws1["A3"] = "Total Automated"
    ws1["B3"] = metrics["total_automated"]
    ws1.merge_cells("B3:E3")

    ws1["A4"] = "Total Maintained"
    ws1.merge_cells("A4:A6")
    ws1["B4"] = "IM"
    ws1.merge_cells("B4:C4")
    ws1["D4"] = "VoIP"
    ws1.merge_cells("D4:E4")

    ws1["B5"] = "Android"
    ws1["C5"] = "iOS"
    ws1["D5"] = "Android"
    ws1["E5"] = "iOS"

    ws1["B6"] = _to_int(s1.get("im_android"))
    ws1["C6"] = _to_int(s1.get("im_ios"))
    ws1["D6"] = _to_int(s1.get("voip_android"))
    ws1["E6"] = _to_int(s1.get("voip_ios"))

    ws1["A7"] = "Found Bugs"
    ws1["B7"] = _to_int(s1.get("found_bugs"))
    ws1.merge_cells("B7:E7")

    ws1["A8"] = "Found Defects"
    ws1["B8"] = _to_int(s1.get("found_defects"))
    ws1.merge_cells("B8:E8")

    ws1["A9"] = sanitize_excel_value(s1.get("note", DEFAULT_SHEET1_NOTE))
    ws1.merge_cells("A9:E10")

    ws1["A13"] = sanitize_excel_value(s1.get("warning", DEFAULT_SHEET1_WARNING))
    ws1.merge_cells("A13:A16")

    style_range(ws1, "A2:E8", border=border_all, font=font_reg, alignment=align_center)
    for coord in ["A2", "A3", "A4", "B4", "D4", "B5", "C5", "D5", "E5", "A7", "A8"]:
        ws1[coord].font = font_bold
    style_range(ws1, "B4:C4", fill=fill_purple, font=font_bold, alignment=align_center)
    style_range(ws1, "D4:E4", fill=fill_gray, font=font_bold, alignment=align_center)
    style_range(ws1, "A9:E10", border=border_all, fill=fill_gray, font=font_reg, alignment=align_left)
    style_range(ws1, "A13:A16", border=border_all, fill=fill_yellow, font=font_reg, alignment=align_left)

    ws1.column_dimensions["A"].width = 24
    for col_let in ["B", "C", "D", "E"]:
        ws1.column_dimensions[col_let].width = 14

    # =========================================================================
    # SHEET 2
    # =========================================================================
    ws2 = wb.create_sheet("Sheet2")
    s2 = report_data.get("sheet2", {})
    w_im = s2.get("weekly_im", {})
    w_voip = s2.get("weekly_voip", {})
    a_im = s2.get("automated_im", {})
    a_voip = s2.get("automated_voip", {})
    ml = s2.get("medium_low", {})

    # Top table: Weekly Automated IM & VOIP
    ws2["B2"] = "Weekly Automated IM"
    ws2.merge_cells("B2:E2")
    ws2["G2"] = "Weekly Automated VOIP"
    ws2.merge_cells("G2:J2")

    ws2["B3"] = "iOS"
    ws2.merge_cells("B3:C3")
    ws2["D3"] = "Android"
    ws2.merge_cells("D3:E3")
    ws2["G3"] = "iOS"
    ws2.merge_cells("G3:H3")
    ws2["I3"] = "Android"
    ws2.merge_cells("I3:J3")

    ws2["B4"], ws2["C4"] = "Gating", _to_int(w_im.get("ios_gating"))
    ws2["D4"], ws2["E4"] = "Gating", _to_int(w_im.get("android_gating"))
    ws2["G4"], ws2["H4"] = "Gating", _to_int(w_voip.get("ios_gating"))
    ws2["I4"], ws2["J4"] = "Gating", _to_int(w_voip.get("android_gating"))

    ws2["B5"], ws2["C5"] = "High", _to_int(w_im.get("ios_high"))
    ws2["D5"], ws2["E5"] = "High", _to_int(w_im.get("android_high"))
    ws2["G5"], ws2["H5"] = "High", _to_int(w_voip.get("ios_high"))
    ws2["I5"], ws2["J5"] = "High", _to_int(w_voip.get("android_high"))

    ws2["B6"] = "TOTAL :"
    ws2.merge_cells("B6:C6")
    ws2["D6"] = "=SUM(C4:C5,E4:E5)"
    ws2.merge_cells("D6:E6")

    ws2["G6"] = "TOTAL :"
    ws2.merge_cells("G6:H6")
    ws2["I6"] = "=SUM(H4:H5,J4:J5)"
    ws2.merge_cells("I6:J6")

    style_range(ws2, "B2:E6", border=border_all, font=font_reg, alignment=align_center)
    style_range(ws2, "G2:J6", border=border_all, font=font_reg, alignment=align_center)
    style_range(ws2, "B2:E2", fill=fill_purple, font=font_bold)
    style_range(ws2, "G2:J2", fill=fill_gray, font=font_bold)
    ws2["B4"].fill = ws2["G4"].fill = fill_ios_1
    ws2["B5"].fill = ws2["G5"].fill = fill_ios_2
    ws2["D4"].fill = ws2["I4"].fill = fill_and_1
    ws2["D5"].fill = ws2["I5"].fill = fill_and_2
    ws2["B6"].font = ws2["D6"].font = ws2["G6"].font = ws2["I6"].font = font_bold

    # Middle table: Automated IM & VOIP
    ws2["B8"] = "Automated IM"
    ws2.merge_cells("B8:E8")
    ws2["G8"] = "Automated VOIP"
    ws2.merge_cells("G8:J8")

    ws2["B9"] = "iOS"
    ws2.merge_cells("B9:C9")
    ws2["D9"] = "Android"
    ws2.merge_cells("D9:E9")
    ws2["G9"] = "iOS"
    ws2.merge_cells("G9:H9")
    ws2["I9"] = "Android"
    ws2.merge_cells("I9:J9")

    ws2["B10"], ws2["C10"] = "Gating", _to_int(a_im.get("ios_gating"))
    ws2["D10"], ws2["E10"] = "Gating", _to_int(a_im.get("android_gating"))
    ws2["G10"], ws2["H10"] = "Gating", _to_int(a_voip.get("ios_gating"))
    ws2["I10"], ws2["J10"] = "Gating", _to_int(a_voip.get("android_gating"))

    ws2["B11"], ws2["C11"] = "High", _to_int(a_im.get("ios_high"))
    ws2["D11"], ws2["E11"] = "High", _to_int(a_im.get("android_high"))
    ws2["G11"], ws2["H11"] = "High", _to_int(a_voip.get("ios_high"))
    ws2["I11"], ws2["J11"] = "High", _to_int(a_voip.get("android_high"))

    ws2["B12"], ws2["C12"] = "Total", "=SUM(C10,C11,C21)"
    ws2["D12"], ws2["E12"] = "Total", "=SUM(E10,E11,E21)"
    ws2["G12"], ws2["H12"] = "Total", "=SUM(H10,H11,H21)"
    ws2["I12"], ws2["J12"] = "Total", "=SUM(J10,J11,J21)"

    ws2["B13"], ws2["C13"] = "Rate", "=IF(C14>0,C12/C14,0)"
    ws2["D13"], ws2["E13"] = "Rate", "=IF(E14>0,E12/E14,0)"
    ws2["G13"], ws2["H13"] = "Rate", "=IF(H14>0,H12/H14,0)"
    ws2["I13"], ws2["J13"] = "Rate", "=IF(J14>0,J12/J14,0)"
    ws2["C13"].number_format = "0%"
    ws2["E13"].number_format = "0%"
    ws2["H13"].number_format = "0%"
    ws2["J13"].number_format = "0%"

    ws2["B14"], ws2["C14"] = "To Be", _to_int(a_im.get("ios_to_be"))
    ws2["D14"], ws2["E14"] = "To Be", _to_int(a_im.get("android_to_be"))
    ws2["G14"], ws2["H14"] = "To Be", _to_int(a_voip.get("ios_to_be"))
    ws2["I14"], ws2["J14"] = "To Be", _to_int(a_voip.get("android_to_be"))

    style_range(ws2, "B8:E14", border=border_all, font=font_reg, alignment=align_center)
    style_range(ws2, "G8:J14", border=border_all, font=font_reg, alignment=align_center)
    style_range(ws2, "B8:E8", fill=fill_purple, font=font_bold)
    style_range(ws2, "G8:J8", fill=fill_gray, font=font_bold)
    ws2["B10"].fill = ws2["G10"].fill = fill_ios_1
    ws2["B11"].fill = ws2["G11"].fill = fill_ios_2
    ws2["D10"].fill = ws2["I10"].fill = fill_and_1
    ws2["D11"].fill = ws2["I11"].fill = fill_and_2
    for col_l in ["B", "D", "G", "I"]:
        ws2[f"{col_l}12"].fill = fill_org_1
        ws2[f"{col_l}13"].fill = fill_org_2
        ws2[f"{col_l}14"].fill = fill_org_3

    # Warning & Medium-Low helper table
    ws2["B18"] = sanitize_excel_value(s2.get("warning", DEFAULT_SHEET2_WARNING))
    ws2.merge_cells("B18:J19")
    style_range(ws2, "B18:J19", border=border_all, fill=fill_yellow, font=font_bold, alignment=align_center)

    ws2["B20"] = "IOS"
    ws2.merge_cells("B20:C20")
    ws2["D20"] = "Android"
    ws2.merge_cells("D20:E20")
    ws2["G20"] = "IOS"
    ws2.merge_cells("G20:H20")
    ws2["I20"] = "Android"
    ws2.merge_cells("I20:J20")

    ws2["B21"], ws2["C21"] = "Medium-Low", _to_int(ml.get("im_ios"))
    ws2["D21"], ws2["E21"] = "Medium-Low", _to_int(ml.get("im_android"))
    ws2["G21"], ws2["H21"] = "Medium-Low", _to_int(ml.get("voip_ios"))
    ws2["I21"], ws2["J21"] = "Medium-Low", _to_int(ml.get("voip_android"))

    style_range(ws2, "B20:E21", border=border_all, font=font_reg, alignment=align_center)
    style_range(ws2, "G20:J21", border=border_all, font=font_reg, alignment=align_center)
    style_range(ws2, "B20:C20", fill=fill_ios_3, font=font_bold)
    style_range(ws2, "G20:H20", fill=fill_ios_3, font=font_bold)
    style_range(ws2, "D20:E20", fill=fill_and_3, font=font_bold)
    style_range(ws2, "I20:J20", fill=fill_and_3, font=font_bold)
    ws2["B21"].fill = ws2["G21"].fill = fill_ios_3
    ws2["D21"].fill = ws2["I21"].fill = fill_and_3

    for col_let in ["B", "C", "D", "E", "G", "H", "I", "J"]:
        ws2.column_dimensions[col_let].width = 14

    # =========================================================================
    # SHEET 3
    # =========================================================================
    ws3 = wb.create_sheet("Sheet3")
    s3 = report_data.get("sheet3", {})
    ws3["B2"] = sanitize_excel_value(s3.get("title", "Problems & Requirements"))
    ws3.merge_cells("B2:F2")
    style_range(ws3, "B2:F2", border=border_all, fill=fill_ios_2, font=font_bold, alignment=align_center)

    headers_s3 = ["Type", "Issue", "Summary", "Status", "Etki Analizi"]
    for idx, h_text in enumerate(headers_s3, start=2):
        cell = ws3.cell(3, idx, h_text)
        cell.fill = fill_ios_3
        cell.font = font_bold
        cell.border = border_all
        cell.alignment = align_center

    row_cursor = 4
    for item in s3.get("items", []):
        ws3.cell(row_cursor, 2, sanitize_excel_value(item.get("Type", ""))).alignment = align_center
        ws3.cell(row_cursor, 3, sanitize_excel_value(item.get("Issue", ""))).alignment = align_center
        ws3.cell(row_cursor, 4, sanitize_excel_value(item.get("Summary", ""))).alignment = align_left
        ws3.cell(row_cursor, 5, sanitize_excel_value(item.get("Status", ""))).alignment = align_center
        ws3.cell(row_cursor, 6, sanitize_excel_value(item.get("Etki Analizi", ""))).alignment = align_center
        for c_i in range(2, 7):
            ws3.cell(row_cursor, c_i).border = border_all
            ws3.cell(row_cursor, c_i).font = font_reg
        row_cursor += 1

    legend_end = row_cursor + 4
    ws3.cell(row_cursor, 2, sanitize_excel_value(s3.get("legend", DEFAULT_SHEET3_LEGEND)))
    ws3.merge_cells(f"B{row_cursor}:F{legend_end}")
    style_range(
        ws3,
        f"B{row_cursor}:F{legend_end}",
        border=border_all,
        fill=fill_ios_2,
        font=font_reg,
        alignment=align_left,
    )

    ws3.column_dimensions["B"].width = 16
    ws3.column_dimensions["C"].width = 16
    ws3.column_dimensions["D"].width = 68
    ws3.column_dimensions["E"].width = 16
    ws3.column_dimensions["F"].width = 16

    # =========================================================================
    # SHEET 4
    # =========================================================================
    ws4 = wb.create_sheet("Sheet4")
    s4 = report_data.get("sheet4", {})
    headers_s4 = ["Type", "Issue", "Summary", "Status", "Platform", "Created", "Priority"]
    for idx, h_text in enumerate(headers_s4, start=2):
        cell = ws4.cell(2, idx, h_text)
        cell.fill = fill_org_3
        cell.font = font_bold
        cell.border = border_all
        cell.alignment = align_center

    r4 = 3
    for item in sort_sheet4_items(s4.get("items", [])):
        ws4.cell(r4, 2, sanitize_excel_value(item.get("Type", ""))).alignment = align_center
        ws4.cell(r4, 3, sanitize_excel_value(item.get("Issue", ""))).alignment = align_center
        ws4.cell(r4, 4, sanitize_excel_value(item.get("Summary", ""))).alignment = align_left
        ws4.cell(r4, 5, sanitize_excel_value(item.get("Status", ""))).alignment = align_center
        ws4.cell(r4, 6, sanitize_excel_value(item.get("Platform", ""))).alignment = align_center
        ws4.cell(r4, 7, sanitize_excel_value(item.get("Created", ""))).alignment = align_center
        ws4.cell(r4, 8, sanitize_excel_value(item.get("Priority", ""))).alignment = align_center
        for c_i in range(2, 9):
            ws4.cell(r4, c_i).border = border_all
            ws4.cell(r4, c_i).font = font_reg
        r4 += 1

    # Blank row before Maintenance table
    r4 += 1
    ws4.cell(r4, 2, sanitize_excel_value(s4.get("maintenance_title", "Güncellemeler Sonrası Otomasyon Bakımları")))
    ws4.merge_cells(f"B{r4}:H{r4}")
    style_range(ws4, f"B{r4}:H{r4}", border=border_all, fill=fill_purple, font=font_bold, alignment=align_center)

    r4 += 1
    ws4.cell(r4, 2, "Platform")
    ws4.cell(r4, 3, "Bakım Nedeni")
    ws4.merge_cells(f"C{r4}:G{r4}")
    ws4.cell(r4, 8, "Adet")
    style_range(ws4, f"B{r4}:H{r4}", border=border_all, font=font_bold, alignment=align_center)

    for m_item in s4.get("maintenance_items", []):
        r4 += 1
        ws4.cell(r4, 2, sanitize_excel_value(m_item.get("Platform", "iOS & Android")))
        ws4.cell(r4, 3, sanitize_excel_value(m_item.get("Bakım Nedeni", "")))
        ws4.merge_cells(f"C{r4}:G{r4}")
        adet_val = (
            metrics["total_maintained_sum"]
            if s4.get("auto_sync_maintenance_count", True)
            else _to_int(m_item.get("Adet", 0))
        )
        ws4.cell(r4, 8, adet_val)
        style_range(ws4, f"B{r4}:H{r4}", border=border_all, font=font_reg, alignment=align_center)
        ws4.cell(r4, 3).alignment = align_left

    ws4.column_dimensions["B"].width = 16
    ws4.column_dimensions["C"].width = 18
    ws4.column_dimensions["D"].width = 65
    ws4.column_dimensions["E"].width = 14
    ws4.column_dimensions["F"].width = 14
    ws4.column_dimensions["G"].width = 16
    ws4.column_dimensions["H"].width = 14

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

