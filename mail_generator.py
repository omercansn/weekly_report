"""
Email-Ready HTML Table Generator & One-Click Rich Clipboard Component.

Generates Outlook Desktop, Outlook Web, Gmail, and Apple Mail compatible HTML tables
matching the exact colors, borders, and merged cells of the original Excel workbook.
All user-supplied values are escaped via `security.escape_html` to prevent XSS.
"""

from datetime import datetime
import json
from typing import Any, Dict, List
import streamlit.components.v1 as components

from data_manager import DEFAULT_SHEET1_NOTE, DEFAULT_SHEET3_LEGEND, compute_metrics, sort_sheet4_items
from security import escape_html

# Exact Excel Theme Palette
COLOR_PURPLE = "#E5DFEC"     # IM Header & Maintenance Header
COLOR_GRAY = "#F2F2F2"       # VoIP Header & Sheet1 Note
COLOR_IOS_1 = "#92CDDC"      # iOS Gating (Sheet2)
COLOR_IOS_2 = "#B7DEE8"      # iOS High (Sheet2) & Sheet3 Title/Legend
COLOR_IOS_3 = "#DAEEF3"      # Sheet3 Column Headers
COLOR_AND_1 = "#C4D79B"      # Android Gating (Sheet2)
COLOR_AND_2 = "#D8E4BC"      # Android High (Sheet2)
COLOR_ORG_1 = "#FABF8F"      # Total (Sheet2)
COLOR_ORG_2 = "#FCD5B4"      # Rate (Sheet2)
COLOR_ORG_3 = "#FDE9D9"      # To Be (Sheet2) & Sheet4 Column Headers
COLOR_WHITE = "#FFFFFF"

TD_BASE = (
    "border: 1px solid #000000; padding: 6px 12px; "
    "font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000;"
)


def _format_created_display(raw_date: str) -> str:
    """Formats YYYY-MM-DD into MM-DD-YY to match Excel's `mm-dd-yy` display format."""
    if not raw_date:
        return ""
    try:
        dt = datetime.strptime(str(raw_date).strip()[:10], "%Y-%m-%d")
        return dt.strftime("%m-%d-%y")
    except Exception:
        return escape_html(raw_date)


def generate_sheet1_html(report_data: Dict[str, Any]) -> str:
    """
    Generates Sheet 1 HTML table (Progress Time, Total Automated, Total Maintained,
    Found Bugs, Found Defects, and the Total Maintained description note).
    Excludes the internal '⚠️ Dikkat' warning box.
    """
    s1 = report_data.get("sheet1", {})
    metrics = compute_metrics(report_data)

    progress_time = escape_html(s1.get("progress_time", ""))
    total_automated = metrics["total_automated"]
    im_android = int(s1.get("im_android", 0))
    im_ios = int(s1.get("im_ios", 0))
    voip_android = int(s1.get("voip_android", 0))
    voip_ios = int(s1.get("voip_ios", 0))
    found_bugs = int(s1.get("found_bugs", 0))
    found_defects = int(s1.get("found_defects", 0))
    note_text = escape_html(s1.get("note", DEFAULT_SHEET1_NOTE), preserve_newlines=True)

    return f"""
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; min-width: 560px;">
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold; width: 170px;">Progress Time</td>
    <td bgcolor="{COLOR_WHITE}" colspan="4" align="center" style="{TD_BASE}">{progress_time}</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold;">Total Automated</td>
    <td bgcolor="{COLOR_WHITE}" colspan="4" align="center" style="{TD_BASE}">{total_automated}</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_WHITE}" rowspan="3" align="center" valign="middle" style="{TD_BASE} font-weight: bold;">Total Maintained</td>
    <td bgcolor="{COLOR_PURPLE}" colspan="2" align="center" style="{TD_BASE} background-color: {COLOR_PURPLE}; font-weight: bold;">IM</td>
    <td bgcolor="{COLOR_GRAY}" colspan="2" align="center" style="{TD_BASE} background-color: {COLOR_GRAY}; font-weight: bold;">VoIP</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold; width: 95px;">Android</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold; width: 95px;">iOS</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold; width: 95px;">Android</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold; width: 95px;">iOS</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{im_android}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{im_ios}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{voip_android}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{voip_ios}</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold;">Found Bugs</td>
    <td bgcolor="{COLOR_WHITE}" colspan="4" align="center" style="{TD_BASE}">{found_bugs}</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold;">Found Defects</td>
    <td bgcolor="{COLOR_WHITE}" colspan="4" align="center" style="{TD_BASE}">{found_defects}</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_GRAY}" colspan="5" align="left" style="{TD_BASE} background-color: {COLOR_GRAY}; font-size: 10pt;">{note_text}</td>
  </tr>
</table>
""".strip()


def generate_sheet2_html(report_data: Dict[str, Any]) -> str:
    """
    Generates Sheet 2 HTML table (Weekly Automated IM & VOIP + Automated IM & VOIP).
    Strictly follows the rule:
    '⚠️ Dikkat: Aşağıdaki sayılar "Total" satırına formülle ekleniyor, mail tablosuna eklemeyelim.'
    So Medium-Low rows are included in Total/Rate calculation, but omitted from the mail output.
    """
    s2 = report_data.get("sheet2", {})
    w_im = s2.get("weekly_im", {})
    w_voip = s2.get("weekly_voip", {})
    a_im = s2.get("automated_im", {})
    a_voip = s2.get("automated_voip", {})
    m = compute_metrics(report_data)

    spacer_td = '<td style="border: none; background: transparent; width: 28px; padding: 0;">&nbsp;</td>'
    empty_row = '<tr><td colspan="9" style="border: none; height: 14px; padding: 0;">&nbsp;</td></tr>'

    return f"""
<table border="0" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000;">
  <!-- Weekly Automated Header -->
  <tr>
    <td bgcolor="{COLOR_PURPLE}" colspan="4" align="center" style="{TD_BASE} background-color: {COLOR_PURPLE}; font-weight: bold;">Weekly Automated IM</td>
    {spacer_td}
    <td bgcolor="{COLOR_GRAY}" colspan="4" align="center" style="{TD_BASE} background-color: {COLOR_GRAY}; font-weight: bold;">Weekly Automated VOIP</td>
  </tr>
  <!-- Weekly Platform Subheader -->
  <tr>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">iOS</td>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">Android</td>
    {spacer_td}
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">iOS</td>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">Android</td>
  </tr>
  <!-- Weekly Gating -->
  <tr>
    <td bgcolor="{COLOR_IOS_1}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_1}; width: 75px;">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} width: 65px;">{int(w_im.get("ios_gating", 0))}</td>
    <td bgcolor="{COLOR_AND_1}" align="center" style="{TD_BASE} background-color: {COLOR_AND_1}; width: 75px;">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} width: 65px;">{int(w_im.get("android_gating", 0))}</td>
    {spacer_td}
    <td bgcolor="{COLOR_IOS_1}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_1}; width: 75px;">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} width: 65px;">{int(w_voip.get("ios_gating", 0))}</td>
    <td bgcolor="{COLOR_AND_1}" align="center" style="{TD_BASE} background-color: {COLOR_AND_1}; width: 75px;">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} width: 65px;">{int(w_voip.get("android_gating", 0))}</td>
  </tr>
  <!-- Weekly High -->
  <tr>
    <td bgcolor="{COLOR_IOS_2}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(w_im.get("ios_high", 0))}</td>
    <td bgcolor="{COLOR_AND_2}" align="center" style="{TD_BASE} background-color: {COLOR_AND_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(w_im.get("android_high", 0))}</td>
    {spacer_td}
    <td bgcolor="{COLOR_IOS_2}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(w_voip.get("ios_high", 0))}</td>
    <td bgcolor="{COLOR_AND_2}" align="center" style="{TD_BASE} background-color: {COLOR_AND_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(w_voip.get("android_high", 0))}</td>
  </tr>
  <!-- Weekly TOTAL -->
  <tr>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE} font-weight: bold;">TOTAL :</td>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE} font-weight: bold;">{m["weekly_im_total"]}</td>
    {spacer_td}
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE} font-weight: bold;">TOTAL :</td>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE} font-weight: bold;">{m["weekly_voip_total"]}</td>
  </tr>
  {empty_row}
  <!-- Cumulative Automated Header -->
  <tr>
    <td bgcolor="{COLOR_PURPLE}" colspan="4" align="center" style="{TD_BASE} background-color: {COLOR_PURPLE}; font-weight: bold;">Automated IM</td>
    {spacer_td}
    <td bgcolor="{COLOR_GRAY}" colspan="4" align="center" style="{TD_BASE} background-color: {COLOR_GRAY}; font-weight: bold;">Automated VOIP</td>
  </tr>
  <!-- Cumulative Platform Subheader -->
  <tr>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">iOS</td>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">Android</td>
    {spacer_td}
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">iOS</td>
    <td bgcolor="{COLOR_WHITE}" colspan="2" align="center" style="{TD_BASE}">Android</td>
  </tr>
  <!-- Cumulative Gating -->
  <tr>
    <td bgcolor="{COLOR_IOS_1}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_1};">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_im.get("ios_gating", 0))}</td>
    <td bgcolor="{COLOR_AND_1}" align="center" style="{TD_BASE} background-color: {COLOR_AND_1};">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_im.get("android_gating", 0))}</td>
    {spacer_td}
    <td bgcolor="{COLOR_IOS_1}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_1};">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_voip.get("ios_gating", 0))}</td>
    <td bgcolor="{COLOR_AND_1}" align="center" style="{TD_BASE} background-color: {COLOR_AND_1};">Gating</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_voip.get("android_gating", 0))}</td>
  </tr>
  <!-- Cumulative High -->
  <tr>
    <td bgcolor="{COLOR_IOS_2}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_im.get("ios_high", 0))}</td>
    <td bgcolor="{COLOR_AND_2}" align="center" style="{TD_BASE} background-color: {COLOR_AND_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_im.get("android_high", 0))}</td>
    {spacer_td}
    <td bgcolor="{COLOR_IOS_2}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_voip.get("ios_high", 0))}</td>
    <td bgcolor="{COLOR_AND_2}" align="center" style="{TD_BASE} background-color: {COLOR_AND_2};">High</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_voip.get("android_high", 0))}</td>
  </tr>
  <!-- Cumulative Total (includes Medium-Low) -->
  <tr>
    <td bgcolor="{COLOR_ORG_1}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_1};">Total</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["im_ios_total"]}</td>
    <td bgcolor="{COLOR_ORG_1}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_1};">Total</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["im_android_total"]}</td>
    {spacer_td}
    <td bgcolor="{COLOR_ORG_1}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_1};">Total</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["voip_ios_total"]}</td>
    <td bgcolor="{COLOR_ORG_1}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_1};">Total</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["voip_android_total"]}</td>
  </tr>
  <!-- Cumulative Rate -->
  <tr>
    <td bgcolor="{COLOR_ORG_2}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_2};">Rate</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["im_ios_rate_str"]}</td>
    <td bgcolor="{COLOR_ORG_2}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_2};">Rate</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["im_android_rate_str"]}</td>
    {spacer_td}
    <td bgcolor="{COLOR_ORG_2}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_2};">Rate</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["voip_ios_rate_str"]}</td>
    <td bgcolor="{COLOR_ORG_2}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_2};">Rate</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m["voip_android_rate_str"]}</td>
  </tr>
  <!-- Cumulative To Be -->
  <tr>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3};">To Be</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_im.get("ios_to_be", 0))}</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3};">To Be</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_im.get("android_to_be", 0))}</td>
    {spacer_td}
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3};">To Be</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_voip.get("ios_to_be", 0))}</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3};">To Be</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{int(a_voip.get("android_to_be", 0))}</td>
  </tr>
</table>
""".strip()


def generate_sheet3_html(report_data: Dict[str, Any]) -> str:
    """Generates Sheet 3 HTML table (Problems & Requirements + Etki Analizi Legend)."""
    s3 = report_data.get("sheet3", {})
    title = escape_html(s3.get("title", "Problems & Requirements"))
    legend = escape_html(s3.get("legend", DEFAULT_SHEET3_LEGEND), preserve_newlines=True)
    items: List[Dict[str, Any]] = s3.get("items", [])

    rows_html: List[str] = []
    for item in items:
        t_val = escape_html(item.get("Type", ""))
        i_val = escape_html(item.get("Issue", ""))
        sum_val = escape_html(item.get("Summary", ""))
        st_val = escape_html(item.get("Status", ""))
        ea_val = escape_html(item.get("Etki Analizi", ""))
        rows_html.append(
            f"""
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{t_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{i_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="left" style="{TD_BASE}">{sum_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{st_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold;">{ea_val}</td>
  </tr>""".rstrip()
        )

    body_str = "\n".join(rows_html)
    return f"""
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; width: 100%; max-width: 960px;">
  <tr>
    <td bgcolor="{COLOR_IOS_2}" colspan="5" align="center" style="{TD_BASE} background-color: {COLOR_IOS_2}; font-weight: bold; font-size: 12pt;">{title}</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_IOS_3}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_3}; font-weight: bold; width: 110px;">Type</td>
    <td bgcolor="{COLOR_IOS_3}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_3}; font-weight: bold; width: 110px;">Issue</td>
    <td bgcolor="{COLOR_IOS_3}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_3}; font-weight: bold;">Summary</td>
    <td bgcolor="{COLOR_IOS_3}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_3}; font-weight: bold; width: 115px;">Status</td>
    <td bgcolor="{COLOR_IOS_3}" align="center" style="{TD_BASE} background-color: {COLOR_IOS_3}; font-weight: bold; width: 105px;">Etki Analizi</td>
  </tr>
{body_str}
  <tr>
    <td bgcolor="{COLOR_IOS_2}" colspan="5" align="left" style="{TD_BASE} background-color: {COLOR_IOS_2}; font-size: 10pt; line-height: 1.45;">{legend}</td>
  </tr>
</table>
""".strip()


def generate_sheet4_html(report_data: Dict[str, Any]) -> str:
    """Generates Sheet 4 HTML tables (Defects/Bugs/Stories + Güncellemeler Sonrası Otomasyon Bakımları)."""
    s4 = report_data.get("sheet4", {})
    metrics = compute_metrics(report_data)
    items: List[Dict[str, Any]] = sort_sheet4_items(s4.get("items", []))
    maint_title = escape_html(s4.get("maintenance_title", "Güncellemeler Sonrası Otomasyon Bakımları"))
    auto_sync = s4.get("auto_sync_maintenance_count", True)
    maint_items: List[Dict[str, Any]] = s4.get("maintenance_items", [])

    rows_html: List[str] = []
    for item in items:
        t_val = escape_html(item.get("Type", ""))
        i_val = escape_html(item.get("Issue", ""))
        sum_val = escape_html(item.get("Summary", ""))
        st_val = escape_html(item.get("Status", ""))
        pl_val = escape_html(item.get("Platform", ""))
        cr_val = _format_created_display(item.get("Created", ""))
        pr_val = escape_html(item.get("Priority", ""))
        rows_html.append(
            f"""
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{t_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{i_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="left" style="{TD_BASE}">{sum_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{st_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{pl_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{cr_val}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{pr_val}</td>
  </tr>""".rstrip()
        )

    maint_rows_html: List[str] = []
    for m_item in maint_items:
        m_plat = escape_html(m_item.get("Platform", "iOS & Android"))
        m_reason = escape_html(m_item.get("Bakım Nedeni", ""))
        m_count = metrics["total_maintained_sum"] if auto_sync else int(m_item.get("Adet", 0))
        maint_rows_html.append(
            f"""
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} white-space: nowrap;">{m_plat}</td>
    <td bgcolor="{COLOR_WHITE}" colspan="5" align="left" style="{TD_BASE}">{m_reason}</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">{m_count}</td>
  </tr>""".rstrip()
        )

    body_str = "\n".join(rows_html)
    maint_str = "\n".join(maint_rows_html)

    return f"""
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; width: 100%; max-width: 980px;">
  <tr>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3}; font-weight: bold; width: 85px;">Type</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3}; font-weight: bold; width: 115px;">Issue</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3}; font-weight: bold;">Summary</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3}; font-weight: bold; width: 85px;">Status</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3}; font-weight: bold; width: 90px;">Platform</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3}; font-weight: bold; width: 95px;">Created</td>
    <td bgcolor="{COLOR_ORG_3}" align="center" style="{TD_BASE} background-color: {COLOR_ORG_3}; font-weight: bold; width: 85px;">Priority</td>
  </tr>
{body_str}
  <tr>
    <td colspan="7" style="border: none; height: 16px; padding: 0;">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_PURPLE}" colspan="7" align="center" style="{TD_BASE} background-color: {COLOR_PURPLE}; font-weight: bold;">{maint_title}</td>
  </tr>
  <tr>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE} font-weight: bold;">Platform</td>
    <td bgcolor="{COLOR_WHITE}" colspan="5" align="center" style="{TD_BASE} font-weight: bold;">Bakım Nedeni</td>
    <td bgcolor="{COLOR_WHITE}" align="center" style="{TD_BASE}">Adet</td>
  </tr>
{maint_str}
</table>
""".strip()


def generate_combined_mail_html(
    report_data: Dict[str, Any],
    include_s1: bool = True,
    include_s2: bool = True,
    include_s3: bool = True,
    include_s4: bool = True,
    intro_text: str = "",
) -> str:
    """Combines selected sheet tables into a single ready-to-send email HTML body."""
    parts: List[str] = []
    if intro_text.strip():
        clean_intro = escape_html(intro_text.strip(), preserve_newlines=True)
        parts.append(
            f'<p style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; margin-bottom: 16px;">{clean_intro}</p>'
        )
    if include_s1:
        parts.append(generate_sheet1_html(report_data))
    if include_s2:
        parts.append(generate_sheet2_html(report_data))
    if include_s3:
        parts.append(generate_sheet3_html(report_data))
    if include_s4:
        parts.append(generate_sheet4_html(report_data))

    return '<br><br>\n'.join(parts)


def render_copy_card(html_table: str, label: str, key: str, height: int = 440) -> None:
    """
    Renders an interactive toolbar + live email preview card.
    Provides:
    1. '📋 Mail Tablosunu Kopyala (Zengin Görünüm)' -> Copies as `text/html` (and DOM selection fallback)
       so pasting into Outlook/Gmail preserves all colors, borders, and merged cells.
    2. '程式碼 HTML Kodunu Kopyala' -> Copies raw HTML source if needed.
    """
    escaped_js_html = json.dumps(html_table)
    safe_key = "".join(c for c in key if c.isalnum() or c == "_")

    widget_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body {{
          margin: 0;
          padding: 0;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          background: transparent;
        }}
        .mail-card {{
          border: 1px solid #d1d5db;
          border-radius: 10px;
          background: #ffffff;
          overflow: hidden;
          box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        }}
        .mail-toolbar {{
          display: flex;
          align-items: center;
          justify-content: space-between;
          flex-wrap: wrap;
          gap: 10px;
          padding: 10px 16px;
          background: #f8fafc;
          border-bottom: 1px solid #e2e8f0;
        }}
        .mail-title {{
          font-size: 14px;
          font-weight: 600;
          color: #1e293b;
        }}
        .btn-group {{
          display: flex;
          align-items: center;
          gap: 8px;
        }}
        .copy-btn {{
          cursor: pointer;
          border: none;
          border-radius: 7px;
          padding: 8px 14px;
          font-size: 13px;
          font-weight: 600;
          background: #2563eb;
          color: #ffffff;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
        }}
        .copy-btn:hover {{
          background: #1d4ed8;
        }}
        .copy-btn-secondary {{
          cursor: pointer;
          border: 1px solid #cbd5e1;
          border-radius: 7px;
          padding: 7px 12px;
          font-size: 12px;
          font-weight: 500;
          background: #ffffff;
          color: #334155;
          transition: all 0.15s ease;
        }}
        .copy-btn-secondary:hover {{
          background: #f1f5f9;
        }}
        .status-badge {{
          font-size: 12px;
          font-weight: 600;
          color: #15803d;
          background: #dcfce7;
          border: 1px solid #86efac;
          padding: 4px 10px;
          border-radius: 6px;
          display: none;
        }}
        .mail-preview-area {{
          padding: 16px;
          overflow-x: auto;
          background: #ffffff;
        }}
      </style>
    </head>
    <body>
      <div class="mail-card">
        <div class="mail-toolbar">
          <div class="mail-title">📧 {escape_html(label)} — Mail Önizleme & Kopyalama</div>
          <div class="btn-group">
            <span id="status_{safe_key}" class="status-badge">✅ Panoya Kopyalandı! Mail'e (Ctrl+V) yapıştırabilirsiniz.</span>
            <button class="copy-btn" onclick="copyRichTable_{safe_key}()">
              📋 Mail Tablosunu Kopyala
            </button>
            <button class="copy-btn-secondary" onclick="copyRawHtml_{safe_key}()">
              &lt;/&gt; HTML Kodu
            </button>
          </div>
        </div>
        <div id="preview_{safe_key}" class="mail-preview-area">
          {html_table}
        </div>
      </div>

      <script>
        const rawHtml_{safe_key} = {escaped_js_html};

        function showStatus_{safe_key}(msg) {{
          const badge = document.getElementById("status_{safe_key}");
          badge.textContent = msg;
          badge.style.display = "inline-block";
          setTimeout(() => {{
            badge.style.display = "none";
          }}, 4000);
        }}

        async function copyRichTable_{safe_key}() {{
          const previewEl = document.getElementById("preview_{safe_key}");
          try {{
            if (navigator.clipboard && window.ClipboardItem) {{
              const htmlBlob = new Blob([rawHtml_{safe_key}], {{ type: "text/html" }});
              const textBlob = new Blob([previewEl.innerText], {{ type: "text/plain" }});
              const item = new ClipboardItem({{
                "text/html": htmlBlob,
                "text/plain": textBlob
              }});
              await navigator.clipboard.write([item]);
              showStatus_{safe_key}("✅ Tablo kopyalandı! Outlook / Gmail'e doğrudan yapıştırabilirsiniz (Ctrl+V).");
              return;
            }}
          }} catch (err) {{
            // Fallback to DOM Range selection copy (works in all iframe environments)
          }}

          try {{
            const range = document.createRange();
            range.selectNodeContents(previewEl);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            document.execCommand("copy");
            selection.removeAllRanges();
            showStatus_{safe_key}("✅ Tablo kopyalandı! Outlook / Gmail'e doğrudan yapıştırabilirsiniz (Ctrl+V).");
          }} catch (fallbackErr) {{
            alert("Kopyalama sırasında hata oluştu. Tabloyu fareyle seçip Ctrl+C yapabilirsiniz.");
          }}
        }}

        async function copyRawHtml_{safe_key}() {{
          try {{
            await navigator.clipboard.writeText(rawHtml_{safe_key});
            showStatus_{safe_key}("✅ Ham HTML kodu panoya kopyalandı!");
          }} catch (e) {{
            const ta = document.createElement("textarea");
            ta.value = rawHtml_{safe_key};
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
            showStatus_{safe_key}("✅ Ham HTML kodu panoya kopyalandı!");
          }}
        }}
      </script>
    </body>
    </html>
    """
    components.html(widget_html, height=height, scrolling=True)

