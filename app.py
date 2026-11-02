"""
BiP / QA Haftalık Otomasyon Raporu — Streamlit Web Uygulaması

Yetki Modeli (RBAC):
- GÖRÜNTÜLEME MODU (Varsayılan — Linke giren herkes):
  Tüm haftalık raporları görüntüleyebilir, her bir sayfanın veya tüm raporun mail tablosunu
  tek tıkla kopyalayabilir (Copy-Paste) ve Excel (.xlsx) olarak indirebilir. Veri değiştiremez.
- DÜZENLEME MODU (Yetkili Parolası ile):
  Üst bardaki '🔓 Düzenleme Yetkisi Aç' butonundan parola girerek aktif edilir.
  Tabloları düzenleyebilir, satır ekleyip silebilir, yeni hafta oluşturabilir ve Excel yükleyebilir.
"""

import copy
import pandas as pd
import streamlit as st

from data_manager import (
    DEFAULT_SHEET1_NOTE,
    DEFAULT_SHEET3_LEGEND,
    compute_metrics,
    create_new_week_report,
    export_to_excel_bytes,
    initialize_default_report_if_needed,
    list_reports,
    load_report,
    parse_excel_file,
    save_report,
    sort_sheet4_items,
)
from mail_generator import (
    generate_combined_mail_html,
    generate_sheet1_html,
    generate_sheet2_html,
    generate_sheet3_html,
    generate_sheet4_html,
    render_copy_card,
)
from security import (
    attempt_editor_login,
    init_security_state,
    is_editor_authenticated,
    revoke_editor_access,
    sanitize_text,
    validate_xlsx_upload,
)

st.set_page_config(
    page_title="Haftalık Otomasyon Raporu",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        max-width: 1380px;
    }
    .mini-kpi {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 8px 12px;
        text-align: center;
    }
    .mini-kpi-label {
        font-size: 11px;
        color: #64748b;
        font-weight: 600;
        text-transform: uppercase;
    }
    .mini-kpi-val {
        font-size: 18px;
        font-weight: 700;
        color: #0f172a;
    }
    .step-badge {
        display: inline-block;
        background: #eff6ff;
        color: #1d4ed8;
        border: 1px solid #bfdbfe;
        font-size: 12px;
        font-weight: 600;
        padding: 3px 10px;
        border-radius: 999px;
        margin-bottom: 6px;
    }
    .role-badge-viewer {
        display: inline-block;
        background: #f1f5f9;
        color: #475569;
        border: 1px solid #cbd5e1;
        font-size: 12px;
        font-weight: 600;
        padding: 2px 9px;
        border-radius: 6px;
    }
    .role-badge-editor {
        display: inline-block;
        background: #dcfce7;
        color: #15803d;
        border: 1px solid #86efac;
        font-size: 12px;
        font-weight: 600;
        padding: 2px 9px;
        border-radius: 6px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def json_clone_with_optional_rollover(report: dict, add_weekly_to_cum: bool) -> dict:
    cloned = copy.deepcopy(report)
    if add_weekly_to_cum:
        w_im = cloned["sheet2"]["weekly_im"]
        w_voip = cloned["sheet2"]["weekly_voip"]
        a_im = cloned["sheet2"]["automated_im"]
        a_voip = cloned["sheet2"]["automated_voip"]

        a_im["ios_gating"] = int(a_im.get("ios_gating", 0)) + int(w_im.get("ios_gating", 0))
        a_im["ios_high"] = int(a_im.get("ios_high", 0)) + int(w_im.get("ios_high", 0))
        a_im["android_gating"] = int(a_im.get("android_gating", 0)) + int(w_im.get("android_gating", 0))
        a_im["android_high"] = int(a_im.get("android_high", 0)) + int(w_im.get("android_high", 0))

        a_voip["ios_gating"] = int(a_voip.get("ios_gating", 0)) + int(w_voip.get("ios_gating", 0))
        a_voip["ios_high"] = int(a_voip.get("ios_high", 0)) + int(w_voip.get("ios_high", 0))
        a_voip["android_gating"] = int(a_voip.get("android_gating", 0)) + int(w_voip.get("android_gating", 0))
        a_voip["android_high"] = int(a_voip.get("android_high", 0)) + int(w_voip.get("android_high", 0))
    return cloned


def main() -> None:
    init_security_state()
    is_editor = is_editor_authenticated()

    # Varsayılan Raporu Hazırla
    default_id = initialize_default_report_if_needed()
    if "active_report_id" not in st.session_state:
        st.session_state.active_report_id = default_id

    reports_meta = list_reports()
    report_ids = [r["report_id"] for r in reports_meta]
    if st.session_state.active_report_id not in report_ids and report_ids:
        st.session_state.active_report_id = report_ids[0]

    # İzleyici modunda her zaman diskteki en güncel veriyi oku (başka bir editör güncellemiş olabilir)
    if (
        not is_editor
        or "report_data" not in st.session_state
        or st.session_state.get("loaded_report_id") != st.session_state.active_report_id
    ):
        loaded = load_report(st.session_state.active_report_id)
        if loaded:
            st.session_state.report_data = loaded
            st.session_state.loaded_report_id = st.session_state.active_report_id

    report = st.session_state.report_data
    current_user = st.session_state.get("auth_user_name", "Misafir (Salt Okunur)")
    s1 = report["sheet1"]
    s2 = report["sheet2"]
    s3 = report["sheet3"]
    s4 = report["sheet4"]

    # =========================================================================
    # ÜST KONTROL BARI
    # =========================================================================
    top_c1, top_c2, top_c3, top_c4 = st.columns([2.4, 2.2, 1.2, 1.5], vertical_alignment="center")
    with top_c1:
        st.markdown("### 📊 Haftalık Otomasyon Raporu")
        if is_editor:
            st.markdown(
                f'<span class="role-badge-editor">✏️ Düzenleme Modu Aktif ({sanitize_text(current_user, 30)}) • Otomatik Kayıt</span>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<span class="role-badge-viewer">👁️ Görüntüleme Modu (Salt Okunur) — Mail Kopyalama & İndirme Açık</span>',
                unsafe_allow_html=True,
            )

    with top_c2:
        id_to_label = {
            r["report_id"]: f"📅 {r['title']} ({r['progress_time']})" for r in reports_meta
        }
        selected_id = st.selectbox(
            "Aktif Rapor",
            options=report_ids,
            index=report_ids.index(st.session_state.active_report_id)
            if st.session_state.active_report_id in report_ids
            else 0,
            format_func=lambda x: id_to_label.get(x, x),
            label_visibility="collapsed",
        )
        if selected_id != st.session_state.active_report_id:
            st.session_state.active_report_id = selected_id
            st.session_state.loaded_report_id = None
            st.rerun()

    with top_c3:
        excel_bytes = export_to_excel_bytes(report)
        st.download_button(
            label="📤 Excel (.xlsx) İndir",
            data=excel_bytes,
            file_name=f"Otomasyon_Raporu_{st.session_state.active_report_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with top_c4:
        if not is_editor:
            with st.popover("🔓 Düzenleme Yetkisi Aç", use_container_width=True):
                st.markdown("**🔒 Yetkili Editör Girişi**")
                st.caption("Rapor verilerini düzenlemek veya yeni hafta eklemek için yetkili parolasını girin.")
                with st.form("editor_login_form", clear_on_submit=False):
                    ed_name = st.text_input("Adınız / Rumuz", placeholder="Örn: QA Ekibi", max_chars=40)
                    ed_pwd = st.text_input("Düzenleme Parolası", type="password")
                    login_btn = st.form_submit_button("Yetkiyi Aç", type="primary", use_container_width=True)
                    if login_btn:
                        ok, msg = attempt_editor_login(ed_name, ed_pwd)
                        if ok:
                            st.toast(msg, icon="🔓")
                            st.rerun()
                        else:
                            st.error(msg)
        else:
            btn_col1, btn_col2 = st.columns([1.2, 1])
            with btn_col1:
                with st.popover("➕ Yeni Hafta / Yükle", use_container_width=True):
                    p_tab1, p_tab2 = st.tabs(["➕ Yeni Hafta", "📥 Excel Yükle"])
                    with p_tab1:
                        new_week_code = st.text_input("Yeni Hafta Kodu", placeholder="Örn: w44_2026")
                        new_progress = st.text_input("Tarih Aralığı", placeholder="Örn: 26 Oct - 30 Oct")
                        add_weekly_to_cum = st.checkbox(
                            "Bu haftanın Weekly sayılarını kümülatif tabloya ekle ve haftalık sayaçları sıfırla",
                            value=True,
                        )
                        if st.button("Oluştur", type="primary", use_container_width=True):
                            if not new_week_code.strip() or not new_progress.strip():
                                st.error("Lütfen Hafta Kodu ve Tarih Aralığı girin.")
                            else:
                                base_copy = json_clone_with_optional_rollover(report, add_weekly_to_cum)
                                created_id = create_new_week_report(
                                    base_report=base_copy,
                                    new_report_id=new_week_code.strip(),
                                    new_title=f"{new_week_code.strip()} Otomasyon Raporu",
                                    new_progress_time=new_progress.strip(),
                                    rollover_weekly_to_cumulative=add_weekly_to_cum,
                                    updated_by=current_user,
                                )
                                st.session_state.active_report_id = created_id
                                st.session_state.loaded_report_id = None
                                st.rerun()

                    with p_tab2:
                        uploaded_xlsx = st.file_uploader("Excel Seçin (.xlsx, Maks 5MB)", type=["xlsx"])
                        import_id = st.text_input("Rapor Kodu", value="w43_import")
                        if st.button("İçe Aktar", type="primary", use_container_width=True):
                            is_valid, err_msg = validate_xlsx_upload(uploaded_xlsx)
                            if not is_valid:
                                st.error(err_msg)
                            else:
                                parsed_data = parse_excel_file(
                                    uploaded_xlsx,
                                    report_id=import_id,
                                    title=f"{import_id} Otomasyon Raporu",
                                )
                                saved_id = save_report(import_id, parsed_data, updated_by=current_user)
                                st.session_state.active_report_id = saved_id
                                st.session_state.loaded_report_id = None
                                st.rerun()
            with btn_col2:
                if st.button("🔒 Kilitle", help="Düzenleme modunu kapat ve Salt Okunur moda geç", use_container_width=True):
                    revoke_editor_access()
                    st.toast("Düzenleme yetkisi kapatıldı (Salt Okunur Mod).", icon="🔒")
                    st.rerun()

    # =========================================================================
    # ANA SEKMELER
    # =========================================================================
    tab_all, tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📧 Tüm Raporu Kopyala (Hızlı Mail)",
            "1️⃣ Sayfa 1: Genel Özet",
            "2️⃣ Sayfa 2: IM & VoIP Metrikleri",
            "3️⃣ Sayfa 3: Problems & Requirements",
            "4️⃣ Sayfa 4: Bulgular & Bakımlar",
        ]
    )

    # -------------------------------------------------------------------------
    # TAB 1: SAYFA 1 (Genel Özet)
    # -------------------------------------------------------------------------
    with tab1:
        col_left1, col_right1 = st.columns([1, 1.3], gap="large")
        with col_left1:
            badge_txt = "Adım 1: Bilgileri Düzenleyin" if is_editor else "Özet Bilgiler (Düzenlemek için üstten yetki açın)"
            st.markdown(f'<span class="step-badge">{badge_txt}</span>', unsafe_allow_html=True)

            c_pt, c_fb, c_fd = st.columns([1.6, 1, 1])
            with c_pt:
                val_pt = st.text_input(
                    "Progress Time (Tarih Aralığı)",
                    value=s1.get("progress_time", "19 Oct - 23 Oct"),
                    disabled=not is_editor,
                )
            with c_fb:
                val_fb = st.number_input(
                    "Found Bugs", min_value=0, value=int(s1.get("found_bugs", 0)), disabled=not is_editor
                )
            with c_fd:
                val_fd = st.number_input(
                    "Found Defects", min_value=0, value=int(s1.get("found_defects", 0)), disabled=not is_editor
                )

            st.markdown("**Total Maintained (Bakımı Yapılan Senaryo Sayıları)**")
            df_maint_s1 = pd.DataFrame(
                [
                    {
                        "IM Android": int(s1.get("im_android", 0)),
                        "IM iOS": int(s1.get("im_ios", 0)),
                        "VoIP Android": int(s1.get("voip_android", 0)),
                        "VoIP iOS": int(s1.get("voip_ios", 0)),
                    }
                ]
            )
            edited_maint_s1 = st.data_editor(
                df_maint_s1,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                disabled=not is_editor,
                key="s1_maint_grid",
            )
            if is_editor:
                s1["progress_time"] = val_pt
                s1["found_bugs"] = int(val_fb)
                s1["found_defects"] = int(val_fd)
                row_m = edited_maint_s1.iloc[0]
                s1["im_android"] = int(row_m.get("IM Android", 0) or 0)
                s1["im_ios"] = int(row_m.get("IM iOS", 0) or 0)
                s1["voip_android"] = int(row_m.get("VoIP Android", 0) or 0)
                s1["voip_ios"] = int(row_m.get("VoIP iOS", 0) or 0)

            metrics = compute_metrics(report)
            st.caption(
                f"💡 **Total Automated ({metrics['total_automated']})** Sayfa 2'den otomatik hesaplanır. "
                f"**Total Maintained ({metrics['total_maintained_sum']})** Sayfa 4'e otomatik yansır."
            )

            if is_editor:
                with st.expander("⚙️ Gelişmiş Ayarlar (Manuel Total Automated / Not Düzenle)"):
                    s1["auto_calc_total_automated"] = st.checkbox(
                        "Total Automated değerini Sayfa 2'den otomatik hesapla",
                        value=s1.get("auto_calc_total_automated", True),
                    )
                    if not s1["auto_calc_total_automated"]:
                        s1["total_automated_manual"] = st.number_input(
                            "Total Automated (Manuel)",
                            min_value=0,
                            value=int(s1.get("total_automated_manual", 66)),
                        )
                    s1["note"] = st.text_area(
                        "Alt Bilgi Notu", value=s1.get("note", DEFAULT_SHEET1_NOTE), height=65
                    )

        with col_right1:
            st.markdown('<span class="step-badge">Tek Tıkla Mail İçin Kopyalayın</span>', unsafe_allow_html=True)
            render_copy_card(
                generate_sheet1_html(report),
                label="Sayfa 1 (Özet Tablosu)",
                key="s1_copy",
                height=340,
            )

    # -------------------------------------------------------------------------
    # TAB 2: SAYFA 2 (Haftalık & Kümülatif IM / VoIP)
    # -------------------------------------------------------------------------
    with tab2:
        if is_editor:
            st.markdown(
                '<span class="step-badge">Adım 1: Hücrelere Çift Tıklayarak Sayıları Güncelleyin (Toplam ve Yüzdeler Otomatik Hesaplanır)</span>',
                unsafe_allow_html=True,
            )

        w_im = s2["weekly_im"]
        w_voip = s2["weekly_voip"]
        a_im = s2["automated_im"]
        a_voip = s2["automated_voip"]
        ml = s2["medium_low"]

        grid_c1, grid_c2 = st.columns([1, 1.25], gap="large")

        with grid_c1:
            st.markdown("##### 📅 1. Haftalık Otomasyon (Weekly Automated)")
            df_weekly = pd.DataFrame(
                [
                    {
                        "Öncelik": "Gating",
                        "IM iOS": int(w_im.get("ios_gating", 0)),
                        "IM Android": int(w_im.get("android_gating", 0)),
                        "VOIP iOS": int(w_voip.get("ios_gating", 0)),
                        "VOIP Android": int(w_voip.get("android_gating", 0)),
                    },
                    {
                        "Öncelik": "High",
                        "IM iOS": int(w_im.get("ios_high", 0)),
                        "IM Android": int(w_im.get("android_high", 0)),
                        "VOIP iOS": int(w_voip.get("ios_high", 0)),
                        "VOIP Android": int(w_voip.get("android_high", 0)),
                    },
                ]
            )
            ed_weekly = st.data_editor(
                df_weekly,
                hide_index=True,
                use_container_width=True,
                disabled=True if not is_editor else ["Öncelik"],
                key="s2_weekly_grid",
            )
            if is_editor:
                w_im["ios_gating"] = int(ed_weekly.iloc[0]["IM iOS"] or 0)
                w_im["android_gating"] = int(ed_weekly.iloc[0]["IM Android"] or 0)
                w_voip["ios_gating"] = int(ed_weekly.iloc[0]["VOIP iOS"] or 0)
                w_voip["android_gating"] = int(ed_weekly.iloc[0]["VOIP Android"] or 0)

                w_im["ios_high"] = int(ed_weekly.iloc[1]["IM iOS"] or 0)
                w_im["android_high"] = int(ed_weekly.iloc[1]["IM Android"] or 0)
                w_voip["ios_high"] = int(ed_weekly.iloc[1]["VOIP iOS"] or 0)
                w_voip["android_high"] = int(ed_weekly.iloc[1]["VOIP Android"] or 0)

            m_now = compute_metrics(report)
            st.caption(
                f"✅ Otomatik Toplam — **Weekly IM TOTAL:** `{m_now['weekly_im_total']}` | "
                f"**Weekly VOIP TOTAL:** `{m_now['weekly_voip_total']}`"
            )

        with grid_c2:
            st.markdown("##### 📈 2. Kümülatif Otomasyon (Automated IM & VOIP)")
            df_cum = pd.DataFrame(
                [
                    {
                        "Metrik": "Gating",
                        "IM iOS": int(a_im.get("ios_gating", 0)),
                        "IM Android": int(a_im.get("android_gating", 0)),
                        "VOIP iOS": int(a_voip.get("ios_gating", 0)),
                        "VOIP Android": int(a_voip.get("android_gating", 0)),
                    },
                    {
                        "Metrik": "High",
                        "IM iOS": int(a_im.get("ios_high", 0)),
                        "IM Android": int(a_im.get("android_high", 0)),
                        "VOIP iOS": int(a_voip.get("ios_high", 0)),
                        "VOIP Android": int(a_voip.get("android_high", 0)),
                    },
                    {
                        "Metrik": "Medium-Low (Mailde gizli, Total'e dahil)",
                        "IM iOS": int(ml.get("im_ios", 0)),
                        "IM Android": int(ml.get("im_android", 0)),
                        "VOIP iOS": int(ml.get("voip_ios", 0)),
                        "VOIP Android": int(ml.get("voip_android", 0)),
                    },
                    {
                        "Metrik": "To Be (Hedef)",
                        "IM iOS": int(a_im.get("ios_to_be", 0)),
                        "IM Android": int(a_im.get("android_to_be", 0)),
                        "VOIP iOS": int(a_voip.get("ios_to_be", 0)),
                        "VOIP Android": int(a_voip.get("android_to_be", 0)),
                    },
                ]
            )
            ed_cum = st.data_editor(
                df_cum,
                hide_index=True,
                use_container_width=True,
                disabled=True if not is_editor else ["Metrik"],
                key="s2_cum_grid",
            )
            if is_editor:
                a_im["ios_gating"] = int(ed_cum.iloc[0]["IM iOS"] or 0)
                a_im["android_gating"] = int(ed_cum.iloc[0]["IM Android"] or 0)
                a_voip["ios_gating"] = int(ed_cum.iloc[0]["VOIP iOS"] or 0)
                a_voip["android_gating"] = int(ed_cum.iloc[0]["VOIP Android"] or 0)

                a_im["ios_high"] = int(ed_cum.iloc[1]["IM iOS"] or 0)
                a_im["android_high"] = int(ed_cum.iloc[1]["IM Android"] or 0)
                a_voip["ios_high"] = int(ed_cum.iloc[1]["VOIP iOS"] or 0)
                a_voip["android_high"] = int(ed_cum.iloc[1]["VOIP Android"] or 0)

                ml["im_ios"] = int(ed_cum.iloc[2]["IM iOS"] or 0)
                ml["im_android"] = int(ed_cum.iloc[2]["IM Android"] or 0)
                ml["voip_ios"] = int(ed_cum.iloc[2]["VOIP iOS"] or 0)
                ml["voip_android"] = int(ed_cum.iloc[2]["VOIP Android"] or 0)

                a_im["ios_to_be"] = int(ed_cum.iloc[3]["IM iOS"] or 0)
                a_im["android_to_be"] = int(ed_cum.iloc[3]["IM Android"] or 0)
                a_voip["ios_to_be"] = int(ed_cum.iloc[3]["VOIP iOS"] or 0)
                a_voip["android_to_be"] = int(ed_cum.iloc[3]["VOIP Android"] or 0)

            m_now = compute_metrics(report)
            st.caption(
                f"✅ Otomatik Total & Rate — **IM iOS:** `{m_now['im_ios_total']} ({m_now['im_ios_rate_str']})` | "
                f"**IM Android:** `{m_now['im_android_total']} ({m_now['im_android_rate_str']})` | "
                f"**VOIP Android:** `{m_now['voip_android_total']} ({m_now['voip_android_rate_str']})`"
            )

        st.markdown('<span class="step-badge">Tek Tıkla Mail İçin Kopyalayın</span>', unsafe_allow_html=True)
        render_copy_card(
            generate_sheet2_html(report),
            label="Sayfa 2 (Weekly & Automated IM/VOIP)",
            key="s2_copy",
            height=480,
        )

    # -------------------------------------------------------------------------
    # TAB 3: SAYFA 3 (Problems & Requirements)
    # -------------------------------------------------------------------------
    with tab3:
        if is_editor:
            st.markdown(
                '<span class="step-badge">Adım 1: Tabloyu Düzenleyin (Satır eklemek için en alt satıra tıklayın)</span>',
                unsafe_allow_html=True,
            )

        df_s3 = pd.DataFrame(s3.get("items", []))
        if df_s3.empty:
            df_s3 = pd.DataFrame(columns=["Type", "Issue", "Summary", "Status", "Etki Analizi"])

        edited_df_s3 = st.data_editor(
            df_s3,
            num_rows="dynamic" if is_editor else "fixed",
            disabled=not is_editor,
            use_container_width=True,
            hide_index=True,
            key="editor_s3",
            column_config={
                "Type": st.column_config.SelectboxColumn(
                    "Type", options=["Incident", "Improvement", "Task", "Requirement"], required=True, width="small"
                ),
                "Issue": st.column_config.TextColumn("Issue", required=True, width="small"),
                "Summary": st.column_config.TextColumn("Summary", required=True, width="large"),
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=["Analysis", "Pre-Analysis", "In Progress", "Vendor", "Open", "Resolved", "Closed"],
                    required=True,
                    width="small",
                ),
                "Etki Analizi": st.column_config.SelectboxColumn(
                    "Etki Analizi", options=["X", "Y", "Z", "-"], required=True, width="small"
                ),
            },
        )

        if is_editor:
            cleaned_s3_items = []
            for _, row in edited_df_s3.iterrows():
                issue_str = sanitize_text(row.get("Issue", "") if pd.notna(row.get("Issue")) else "", 40).strip()
                summary_str = sanitize_text(row.get("Summary", "") if pd.notna(row.get("Summary")) else "", 400).strip()
                if not issue_str and not summary_str:
                    continue
                cleaned_s3_items.append(
                    {
                        "Type": sanitize_text(
                            row.get("Type", "Incident") if pd.notna(row.get("Type")) else "Incident", 40
                        ),
                        "Issue": issue_str,
                        "Summary": summary_str,
                        "Status": sanitize_text(
                            row.get("Status", "Open") if pd.notna(row.get("Status")) else "Open", 40
                        ),
                        "Etki Analizi": sanitize_text(
                            row.get("Etki Analizi", "-") if pd.notna(row.get("Etki Analizi")) else "-", 10
                        ),
                    }
                )
            s3["items"] = cleaned_s3_items

            with st.expander("⚙️ Tablo Başlığı ve (X)/(Y)/(Z) Açıklama Metnini Düzenle"):
                s3["title"] = st.text_input("Tablo Başlığı", value=s3.get("title", "Problems & Requirements"))
                s3["legend"] = st.text_area(
                    "Etki Analizi Açıklaması", value=s3.get("legend", DEFAULT_SHEET3_LEGEND), height=95
                )

        st.markdown('<span class="step-badge">Tek Tıkla Mail İçin Kopyalayın</span>', unsafe_allow_html=True)
        render_copy_card(
            generate_sheet3_html(report),
            label="Sayfa 3 (Problems & Requirements)",
            key="s3_copy",
            height=580,
        )

    # -------------------------------------------------------------------------
    # TAB 4: SAYFA 4 (Bulgular & Bakımlar — Defect -> Bug -> Story Sıralı)
    # -------------------------------------------------------------------------
    with tab4:
        if is_editor:
            st.markdown(
                '<span class="step-badge">Adım 1: Defect / Bug / Story Listesini Düzenleyin (Sıralama: 1. Defect ➔ 2. Bug ➔ 3. Story)</span>',
                unsafe_allow_html=True,
            )
        st.caption(
            "📌 Kayıtlar mail formatına uygun olarak her zaman otomatik **1. Defect ➔ 2. Bug ➔ 3. Story** sırasıyla korunur."
        )

        df_s4 = pd.DataFrame(sort_sheet4_items(s4.get("items", [])))
        if df_s4.empty:
            df_s4 = pd.DataFrame(
                columns=["Type", "Issue", "Summary", "Status", "Platform", "Created", "Priority"]
            )

        edited_df_s4 = st.data_editor(
            df_s4,
            num_rows="dynamic" if is_editor else "fixed",
            disabled=not is_editor,
            use_container_width=True,
            hide_index=True,
            key="editor_s4",
            column_config={
                "Type": st.column_config.SelectboxColumn(
                    "Type", options=["Defect", "Bug", "Story"], required=True, width="small"
                ),
                "Issue": st.column_config.TextColumn("Issue", required=True, width="small"),
                "Summary": st.column_config.TextColumn("Summary", required=True, width="large"),
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=["Open", "Dev", "Analiz", "Analysis", "In Progress", "Test", "Closed"],
                    required=True,
                    width="small",
                ),
                "Platform": st.column_config.SelectboxColumn(
                    "Platform", options=["Android", "iOS", "İOS", "iOS & Android", "Web"], required=True, width="small"
                ),
                "Created": st.column_config.TextColumn("Created (YYYY-MM-DD)", width="small"),
                "Priority": st.column_config.SelectboxColumn(
                    "Priority", options=["Gating", "High", "Medium", "Low"], required=True, width="small"
                ),
            },
        )

        metrics_now = compute_metrics(report)

        if is_editor:
            cleaned_s4_items = []
            for _, row in edited_df_s4.iterrows():
                issue_str = sanitize_text(row.get("Issue", "") if pd.notna(row.get("Issue")) else "", 40).strip()
                summary_str = sanitize_text(row.get("Summary", "") if pd.notna(row.get("Summary")) else "", 400).strip()
                if not issue_str and not summary_str:
                    continue
                cleaned_s4_items.append(
                    {
                        "Type": sanitize_text(row.get("Type", "Defect") if pd.notna(row.get("Type")) else "Defect", 30),
                        "Issue": issue_str,
                        "Summary": summary_str,
                        "Status": sanitize_text(
                            row.get("Status", "Open") if pd.notna(row.get("Status")) else "Open", 40
                        ),
                        "Platform": sanitize_text(
                            row.get("Platform", "Android") if pd.notna(row.get("Platform")) else "Android", 30
                        ),
                        "Created": sanitize_text(row.get("Created", "") if pd.notna(row.get("Created")) else "", 20),
                        "Priority": sanitize_text(
                            row.get("Priority", "High") if pd.notna(row.get("Priority")) else "High", 30
                        ),
                    }
                )
            s4["items"] = sort_sheet4_items(cleaned_s4_items)

            with st.expander(
                f"⚙️ Güncellemeler Sonrası Otomasyon Bakımları Alt Tablosu (Adet: {metrics_now['total_maintained_sum']})"
            ):
                s4["auto_sync_maintenance_count"] = st.checkbox(
                    "Bakım 'Adet' sayısını Sayfa 1'deki Total Maintained toplamından otomatik al",
                    value=s4.get("auto_sync_maintenance_count", True),
                )
                df_maint = pd.DataFrame(s4.get("maintenance_items", []))
                if df_maint.empty:
                    df_maint = pd.DataFrame(
                        [
                            {
                                "Platform": "iOS & Android",
                                "Bakım Nedeni": "Release yükü kapsamında, arayüzde yapılan geliştirmeler doğrultusunda akış ve elementlerde güncellemeler yapılmıştır.",
                                "Adet": metrics_now["total_maintained_sum"],
                            }
                        ]
                    )
                elif s4["auto_sync_maintenance_count"] and len(df_maint) == 1:
                    df_maint.loc[df_maint.index[0], "Adet"] = metrics_now["total_maintained_sum"]

                edited_maint = st.data_editor(
                    df_maint,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key="editor_s4_maint",
                )
                cleaned_maint = []
                for _, row in edited_maint.iterrows():
                    plat_str = sanitize_text(
                        row.get("Platform", "") if pd.notna(row.get("Platform")) else "", 50
                    ).strip()
                    reason_str = sanitize_text(
                        row.get("Bakım Nedeni", "") if pd.notna(row.get("Bakım Nedeni")) else "", 500
                    ).strip()
                    if not plat_str and not reason_str:
                        continue
                    adet_val = (
                        metrics_now["total_maintained_sum"]
                        if s4["auto_sync_maintenance_count"] and len(edited_maint) == 1
                        else int(row.get("Adet", 0) if pd.notna(row.get("Adet")) else 0)
                    )
                    cleaned_maint.append(
                        {
                            "Platform": plat_str or "iOS & Android",
                            "Bakım Nedeni": reason_str,
                            "Adet": adet_val,
                        }
                    )
                s4["maintenance_items"] = cleaned_maint

        st.markdown('<span class="step-badge">Tek Tıkla Mail İçin Kopyalayın</span>', unsafe_allow_html=True)
        render_copy_card(
            generate_sheet4_html(report),
            label="Sayfa 4 (Bulgular & Otomasyon Bakımları)",
            key="s4_copy",
            height=600,
        )

    # -------------------------------------------------------------------------
    # TAB ALL: TÜM RAPORU TEK TIKLA KOPYALA (Varsayılan Açılış Sekmesi)
    # -------------------------------------------------------------------------
    with tab_all:
        m_all = compute_metrics(report)
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                f'<div class="mini-kpi"><div class="mini-kpi-label">Progress Time</div><div class="mini-kpi-val">{sanitize_text(s1.get("progress_time", "-"), 25)}</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f'<div class="mini-kpi"><div class="mini-kpi-label">Total Automated</div><div class="mini-kpi-val">{m_all["total_automated"]} (IM: {m_all["weekly_im_total"]} | VoIP: {m_all["weekly_voip_total"]})</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f'<div class="mini-kpi"><div class="mini-kpi-label">Total Maintained</div><div class="mini-kpi-val">{m_all["total_maintained_sum"]} Senaryo</div></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f'<div class="mini-kpi"><div class="mini-kpi-label">Kayıt Sayıları</div><div class="mini-kpi-val">{len(s3.get("items", []))} Problem • {len(s4.get("items", []))} Bulgu</div></div>',
                unsafe_allow_html=True,
            )

        st.write("")

        mail_cfg = report.setdefault(
            "mail_settings",
            {
                "include_s1": True,
                "include_s2": True,
                "include_s3": True,
                "include_s4": True,
                "intro_text": "",
            },
        )

        with st.container(border=True):
            header_note = (
                "✏️ Düzenleme Aktif"
                if is_editor
                else "🔒 Salt Okunur — Değiştirmek için sağ üstten 'Düzenleme Yetkisi Aç' butonunu kullanın"
            )
            st.markdown(
                f"#### ⚙️ Mail İçeriğini Özelleştir (Hangi Sayfalar Dahil Edilsin / Mail Giriş Yazısı) "
                f"<span style='font-size:12px; font-weight:500; color:#64748b; margin-left:8px;'>({header_note})</span>",
                unsafe_allow_html=True,
            )

            c_inc1, c_inc2, c_inc3, c_inc4 = st.columns(4)
            with c_inc1:
                inc_s1 = st.checkbox(
                    "📊 Sayfa 1: Genel Özet",
                    value=bool(mail_cfg.get("include_s1", True)),
                    disabled=not is_editor,
                )
            with c_inc2:
                inc_s2 = st.checkbox(
                    "📈 Sayfa 2: IM & VOIP Metrikleri",
                    value=bool(mail_cfg.get("include_s2", True)),
                    disabled=not is_editor,
                )
            with c_inc3:
                inc_s3 = st.checkbox(
                    "🛠️ Sayfa 3: Problems & Requirements",
                    value=bool(mail_cfg.get("include_s3", True)),
                    disabled=not is_editor,
                )
            with c_inc4:
                inc_s4 = st.checkbox(
                    "🐞 Sayfa 4: Bulgular & Bakımlar",
                    value=bool(mail_cfg.get("include_s4", True)),
                    disabled=not is_editor,
                )

            intro_msg = st.text_area(
                "📝 Mail Giriş Yazısı (Tabloların üstünde yer alacak metin)",
                value=mail_cfg.get("intro_text", ""),
                placeholder="Örn: Merhabalar,\n\nBu haftaki otomasyon raporumuz aşağıda bilgilerinize sunulmuştur.",
                height=135,
                disabled=not is_editor,
            )

            if is_editor:
                mail_cfg["include_s1"] = bool(inc_s1)
                mail_cfg["include_s2"] = bool(inc_s2)
                mail_cfg["include_s3"] = bool(inc_s3)
                mail_cfg["include_s4"] = bool(inc_s4)
                mail_cfg["intro_text"] = sanitize_text(intro_msg, max_len=2000)

        combined_html = generate_combined_mail_html(
            report,
            include_s1=mail_cfg.get("include_s1", True),
            include_s2=mail_cfg.get("include_s2", True),
            include_s3=mail_cfg.get("include_s3", True),
            include_s4=mail_cfg.get("include_s4", True),
            intro_text=mail_cfg.get("intro_text", ""),
        )

        render_copy_card(
            combined_html,
            label="Tüm Haftalık Otomasyon Raporu (Tek Tıkla Tüm Maili Kopyala)",
            key="all_copy",
            height=720,
        )

    # =========================================================================
    # SADECE YETKİLİ EDİTÖR MODUNDA OTOMATİK KAYIT YAPILIR
    # =========================================================================
    if is_editor:
        save_report(st.session_state.active_report_id, report, updated_by=current_user)


if __name__ == "__main__":
    main()
