"""
Ernährungsberatungs-Suite v74: Interaktives Web-Tool (Streamlit App für Hund & Katze)
Drei-Dateien-Modularisierung:
- Haupt-App: ernaehrungsberater_app_v74.py (Streamlit UI & Navigation)
- Datenbanken: futter_db.py (Fertigfutter- & Zutaten-Stammdaten)
- Rechen- & Export-Engine: export_engine.py (Rechenmodell & PDF/DOCX Exports)
"""

import math
import datetime
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

from futter_db import DEFAULT_FUTTER_DB_HUND, DEFAULT_FUTTER_DB_KATZE, INGREDIENTS_DB
from export_engine import (
    format_menge_str, parse_meal_supplements, get_bcs_info, get_alter_display_str,
    evaluiere_bedarfsdeckung, recalculate_ration_totals,
    berechne_bedarfsanalyse, analysiere_etikett, berechne_precision_menu,
    build_growth_curve_pdf_bytes, build_bedarfsanalyse_pdf_bytes,
    build_fertigfutter_pdf_bytes, build_eigenration_pdf_bytes,
    build_wochenplan_pdf_bytes, build_berater_dossier_pdf_bytes,
    build_bedarfsanalyse_docx_bytes, build_fertigfutter_docx_bytes,
    build_eigenration_docx_bytes, build_wochenplan_docx_bytes,
    build_berater_dossier_docx_bytes, build_growth_curve_docx_bytes
)

# Seite konfigurieren
st.set_page_config(
    page_title="Ernährungsberatung Hund & Katze v74",
    page_icon="🐾",
    layout="wide"
)

# -----------------------------------------------------------------------------
# DEFAULT PRESET FUTTER DATENBANK & INGREDIENTS DATABASE
# -----------------------------------------------------------------------------

# SESSION STATE INITIALIZATION
if "futter_db_hund" not in st.session_state:
    st.session_state["futter_db_hund"] = DEFAULT_FUTTER_DB_HUND.copy()

if "futter_db_katze" not in st.session_state:
    st.session_state["futter_db_katze"] = DEFAULT_FUTTER_DB_KATZE.copy()

if "ration_list_hund" not in st.session_state:
    st.session_state["ration_list_hund"] = []

if "ration_list_katze" not in st.session_state:
    st.session_state["ration_list_katze"] = []

if "pdf_header_text" not in st.session_state:
    st.session_state["pdf_header_text"] = "Praxis für Tierernährungsberatung | Individueller Fütterungsplan"

if "pdf_footer_text" not in st.session_state:
    st.session_state["pdf_footer_text"] = "Praxis-Handout Tierernährungsberatung | Evidenzbasierte FEDIAF/NRC Rationsberechnung"


# STREAMLIT USER INTERFACE (SPEZIES-AUSWAHL & NAVIGATION)
# -----------------------------------------------------------------------------
if "selected_species" not in st.session_state:
    st.session_state["selected_species"] = None

# TOP BANNER & TIERART-TOGGLE IN DER SIDEBAR
if st.session_state["selected_species"] is not None:
    current_sp = st.session_state["selected_species"]
    st.sidebar.markdown(f"### {'🐶' if current_sp == 'Hund' else '🐱'} Aktiver Modus: {current_sp}")
    if st.sidebar.button("🔄 Tierart wechseln / Zurück zur Startseite"):
        st.session_state["selected_species"] = None
        st.rerun()

# -----------------------------------------------------------------------------
# 1. WILLKOMMENS-STARTSEITE (LANDING PAGE)
# -----------------------------------------------------------------------------
if st.session_state["selected_species"] is None:
    st.title("🐾 Ernährungsberatungs-Suite (Hund & Katze) v74")
    st.caption("Evidenzbasierte FEDIAF & NRC Rationsberechnung, Precision Engine & Modulare PDF-Gutachten")

    st.markdown("---")
    st.subheader("🎯 Bitte wählen Sie die Spezies Ihres Patienten für das Beratungsmodul:")
    
    col_sp1, col_sp2 = st.columns(2)
    
    with col_sp1:
        st.markdown(
            """
            <div style="background-color: #f0f4f8; padding: 20px; border-radius: 10px; border: 2px solid #b4c8dc; text-align: center;">
                <h2>🐶 HUNDE-MODUL</h2>
                <p><b>Evidenzbasierte Ernährungsberatung für Hunde</b></p>
                <ul style="text-align: left; font-size: 0.9em;">
                    <li>Metabolisches Körpergewicht (KM<sup>0.75</sup>) & MER-Faktoren</li>
                    <li>Ca:P Precision Balancer (1.30 : 1 Ratio)</li>
                    <li>Allergie-Aware Protein-Filter & 33% / 100% Modus</li>
                    <li>5 Modulare Hunde-PDF-Exports & 4-Seiten Handout</li>
                </ul>
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.markdown("")
        if st.button("🐶 Hunde-Beratungs-Suite Starten", key="select_dog_btn", use_container_width=True, type="primary"):
            st.session_state["selected_species"] = "Hund"
            st.rerun()

    with col_sp2:
        st.markdown(
            """
            <div style="background-color: #fcf4f0; padding: 20px; border-radius: 10px; border: 2px solid #e8c8b4; text-align: center;">
                <h2>🐱 KATZEN-MODUL</h2>
                <p><b>Spezialisierte Karnivoren-Ernährungsberatung für Katzen</b></p>
                <ul style="text-align: left; font-size: 0.9em;">
                    <li>Felinales metabolisches Gewicht (KM<sup>0.67</sup>) & MER-Skalierung</li>
                    <li>Taurin Auto-Suggester & Preformes Vit-A Guarding (Leber)</li>
                    <li>Nierenfreundliche Ca:P Precision (1.00 : 1 Ratio)</li>
                    <li>Sanfter Abnahme-Zielkorridor (0.5 - 1.0%/Woche gegen Lipidose)</li>
                </ul>
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.markdown("")
        if st.button("🐱 Katzen-Beratungs-Suite Starten", key="select_cat_btn", use_container_width=True, type="primary"):
            st.session_state["selected_species"] = "Katze"
            st.rerun()

    st.markdown("---")
    st.info("💡 **Praxishinweis:** Sie können die Tierart jederzeit über die Sidebar umschalten. Alle Rationsdaten, Unverträglichkeiten und PDF-Kopfzeilen bleiben während der Sitzung gespeichert.")

# -----------------------------------------------------------------------------
# 2. HUNDE-BERATUNGS-SUITE (SPEZIES: HUND)
# -----------------------------------------------------------------------------
elif st.session_state["selected_species"] == "Hund":
    st.title("🐶 Ernährungsberatungs-Suite (Hund) v74")
    st.caption("Evidenzbasierte Hunde-Ernährungsberatung | FEDIAF/NRC Precision Engine")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📊 1. Bedarfsrechner", 
        "🏷️ 2. Futteranalysator & DB", 
        "🥩 3. Selbstgekocht / BARF-Baukasten", 
        "📅 4. Geführter Wochenplan & Menü-Assistent",
        "📑 5. Modulare PDF-Exports & Branding", 
        "📈 6. Nachkontrolle & Wiege-Protokoll",
        "➕ 7. Neues Fertigfutter erfassen & DB"
    ])

    # TAB 1: HUND BEDARFSRECHNER
    with tab1:
        st.header("Bedarfsanalyse für Hunde (Anamnese & Sollwerte)")
        col1, col2 = st.columns(2)
        with col1:
            col_h_id1, col_h_id2 = st.columns([1, 1])
            with col_h_id1:
                h_name = st.text_input("Patienten-Name (Hund)", value="Jaro", key="h_name")
            with col_h_id2:
                h_fall_id = st.text_input("Fall-ID / Aktenzeichen", value="FALL-2026-1001", key="h_fall_id")
            
            h_ist_gw = st.number_input("Ist-Gewicht (kg)", value=70.0, step=0.5, key="h_ist_gw")
            h_soll_gw = st.number_input("Soll- / Ziel-Gewicht (kg)", value=64.0, step=0.5, key="h_soll_gw")
            st.markdown("##### 📅 Alterseingabe (Hund)")
            col_h_a1, col_h_a2 = st.columns([1, 1])
            with col_h_a1:
                h_alter_mode = st.selectbox("Einheit", ["Jahre", "Monate", "Wochen", "Geburtsdatum"], index=0, key="h_alter_mode")
            with col_h_a2:
                if h_alter_mode == "Jahre":
                    h_alter = st.number_input("Alter (Jahre)", value=5.7, step=0.1, key="h_alter")
                elif h_alter_mode == "Monate":
                    h_mon = st.number_input("Alter (Monate)", value=68.0, step=1.0, key="h_alter_mon")
                    h_alter = round(h_mon / 12.0, 2)
                    st.caption(f"📅 Berechnet: **{round(h_mon * 4.33, 1)} Wochen** (~{h_alter} Jahre)")
                elif h_alter_mode == "Wochen":
                    h_wochen = st.number_input("Alter (Wochen)", value=33.0, step=1.0, key="h_alter_wochen")
                    h_alter = round(h_wochen / 52.0, 2)
                    st.caption(f"📅 Berechnet: **{h_wochen} Wochen** (~{round(h_wochen/4.33, 1)} Monate / {h_alter} Jahre)")
                else: # Geburtsdatum
                    import datetime
                    default_h_dob = datetime.date.today() - datetime.timedelta(days=int(5.7 * 365.25))
                    h_dob = st.date_input("Geburtsdatum", value=default_h_dob, key="h_dob")
                    days_old = max(1, (datetime.date.today() - h_dob).days)
                    h_alter = round(days_old / 365.25, 2)
                    st.caption(f"📅 Berechnet: **{round(days_old/7.0,1)} Wochen** (~{round(days_old/30.4,1)} Monate / {h_alter} Jahre)")
        with col2:
            h_lebensphase = st.selectbox("Lebensphase", ["Adult", "Senior", "Junior / Wachstum"], index=1, key="h_phase")
            h_kastriert = st.selectbox("Kastriert?", ["Ja", "Nein"], index=0, key="h_kast")
            h_ziel = st.selectbox("Klinisches Ziel", ["Gewichtsreduktion", "Gewichtserhalt", "Auffütterung / Zunahme"], index=0, key="h_ziel")
            if h_lebensphase == "Junior / Wachstum":
                st.info("💡 **Wachstums-Hinweis:** Bei Welpen im Wachstum ist der erhöhte Energie- & Baustoffbedarf automatisch in der Formel enthalten. Wählen Sie beim klinischen Ziel im Normalfall **'Gewichtserhalt'**, um Skeletterkrankungen durch Überfütterung zu vermeiden!")
            h_aktivitaet = st.selectbox("Aktivitätslevel", ["Inaktiv / Normal", "Aktiv", "Sport / Leistung"], index=0, key="h_akt")
            h_bcs = st.selectbox("Body Condition Score (BCS 1–9)", options=[1, 2, 3, 4, 5, 6, 7, 8, 9], index=6, help="1 = Kachektisch, 5 = Idealgewicht, 9 = Schwerst Adipös", key="h_bcs")
            h_bcs_title, h_bcs_desc = get_bcs_info(h_bcs, "Hund")
            st.info(f"⚖️ **{h_bcs_title}**\n{h_bcs_desc}")
            h_labor_txt = st.text_input("Klinische Laborwerte & Organparameter (optional)", value="Kreatinin: 1.1 mg/dl | SDMA: 12 µg/dl | Urin-pH: 6.6", key="h_labor")
            h_mahlzeiten_n = st.selectbox("Anzahl Mahlzeiten pro Tag", [1, 2, 3, 4, 5, 6], index=2, key="h_mahlzeiten_n")

        st.subheader("🛡️ Patientenschutz: Unverträglichkeiten & Medikation")
        c_unv1, c_unv2 = st.columns(2)
        with c_unv1:
            h_unvert = st.multiselect("Allergie- / Unverträglichkeits-Filter", ["Geflügel", "Rind", "Pferd", "Fisch", "Getreide", "Gemüse"], default=[], key="h_unvert")
        with c_unv2:
            h_med_txt = st.text_input("Medikation / Anamnese", value="Forthyron 400 µg (Schilddrüsenunterfunktion)", key="h_med")
            h_med_list = [m.strip() for m in h_med_txt.split(",") if m.strip()] if h_med_txt else []

        st.markdown("---")
        st.subheader("📐 Rechenmodell für Energiedichte (FEDIAF 4-Stufen vs. Atwater)")
        h_f_modell_t1 = st.radio(
            "Wählen Sie das mathematische Modell zur Energiedichte-Berechnung von Futtermitteln:",
            ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"],
            index=0 if st.session_state.get("h_f_modell", "FEDIAF 4-Stufen-Modell") == "FEDIAF 4-Stufen-Modell" else 1,
            help="🔘 FEDIAF 4-Stufen: Höchste Präzision bei Diäten & Rationsberechnungen (berücksichtigt organische Verdaulichkeit dE nach Rohfasergehalt).\n🔘 Modifiziertes Atwater: Standard-Schätzverfahren mit festen Faktoren (3.5 kcal/g Protein & NfE, 8.5 kcal/g Fett).",
            key="h_f_modell_t1"
        )
        st.session_state["h_f_modell"] = h_f_modell_t1

        h_end_gw = h_soll_gw
        if h_lebensphase == "Junior / Wachstum":
            st.markdown("##### 🦴 Rassespezifisches Wachstums-Setup (Groß- & Riesenrassen-Schutz)")
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                h_end_gw = st.number_input("Erwartetes Endgewicht (ausgewachsen in kg)", value=35.0 if h_ist_gw < 35.0 else h_ist_gw, step=1.0, key="h_end_gw")
                h_rasse_kat = st.selectbox("Rasse-Kategorie", ["Klein (< 15 kg)", "Mittel (15–30 kg)", "Groß (30–45 kg)", "Riesenrasse (> 45 kg, z. B. Leonberger, Dogge)"], index=3 if h_end_gw > 45 else (2 if h_end_gw > 30 else (1 if h_end_gw > 15 else 0)), key="h_rasse_kat")
            with col_g2:
                pct_g = round((h_ist_gw / h_end_gw) * 100.0, 1) if h_end_gw > 0 else 100.0
                if pct_g < 40.0:
                    phase_str = "Phase 1 (<40% Endgewicht - Hoher Energiebedarf 2.0x MER)"
                elif pct_g < 80.0:
                    phase_str = "Phase 2 (40-80% Endgewicht - Moderater Energiebedarf 1.6x MER)"
                else:
                    phase_str = "Phase 3 (80-100% Endgewicht - Ausreifungsphase 1.2x MER)"
                st.metric("Wachstumsfortschritt", f"{pct_g} % des Endgewichts")
                st.caption(f"📊 **Aktuelle Phase:** {phase_str}")

            if h_end_gw > 15.0:
                st.warning("🛡️ **Riesenrassen / Große Rassen-Schutz aktiv (>15 kg Endgewicht):** Strikte Calcium-Obergrenze max. 2,5 g Ca / 1.000 kcal ME (bzw. max. 1,8% i.TS. nach AAFCO/FEDIAF [cite: 4, 142, 269, 273, 328]) zur Prävention von Skeletterkrankungen (OCD, HD, Radius-Curvus-Syndrom) [cite: 4, 5, 21, 24]!")
            else:
                st.info("💡 **Wachstums-Hinweis:** Bei Welpen im Wachstum ist der erhöhte Energie- & Baustoffbedarf automatisch enthalten. Wählen Sie im Normalfall 'Gewichtserhalt'!")

        bedarf_hund = berechne_bedarfsanalyse("Hund", h_name, h_ist_gw, h_soll_gw, h_alter, h_lebensphase, h_kastriert, h_ziel, h_aktivitaet, h_mahlzeiten_n, h_unvert, h_med_list, h_end_gw, bcs=h_bcs, fall_id=h_fall_id, laborwerte=h_labor_txt)
        st.session_state["aktueller_bedarf"] = bedarf_hund
        st.session_state["aktueller_bedarf_hund"] = bedarf_hund

        st.markdown("---")
        st.subheader(f"📊 Berechnete Nährstoff-Sollwerte für {h_name} (FEDIAF / NRC)")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Metabolisches KGW", f"{bedarf_hund['km_meta']} kg^0.75")
        m2.metric("Tages-MER", f"{bedarf_hund['mer_kcal']} kcal", f"{bedarf_hund['mer_mj']} MJ")
        m3.metric("vRp (Eiweiß)", f"{bedarf_hund['vrp_g']} g", f"Brutto: {bedarf_hund['rp_g']} g")
        m4.metric("Calcium / Phosphor", f"{round(bedarf_hund['ca_mg'],0)}mg / {round(bedarf_hund['p_mg'],0)}mg", "Ca:P 1.30:1")
        m5.metric("Jod / Vit D3", f"{bedarf_hund['jod_ug']} µg / {bedarf_hund['vit_d_ie']} IE", f"Vit A: {bedarf_hund['vit_a_ie']} IE")

    # TAB 2: HUND FUTTERANALYSATOR
    with tab2:
        st.header("Fertigfutter-Analyse für Hunde (Weender-Deklaration & Datenbank)")
        st.subheader("📐 Rechenmodell für Energiedichte")
        h_f_modell_t2 = st.radio(
            "Wählen Sie das mathematische Modell zur Energiedichte-Berechnung:",
            ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"],
            index=0 if st.session_state.get("h_f_modell", "FEDIAF 4-Stufen-Modell") == "FEDIAF 4-Stufen-Modell" else 1,
            help="🔘 FEDIAF 4-Stufen: Höchste Präzision, berücksichtigt die Verdaulichkeit in Abhängigkeit vom Rohfasergehalt.\n🔘 Modifiziertes Atwater: Standard-Schätzverfahren (3.5 kcal/g Protein & NfE, 8.5 kcal/g Fett).",
            key="h_f_modell_t2"
        )
        st.session_state["h_f_modell"] = h_f_modell_t2
        h_f_modell = h_f_modell_t2
        st.markdown("---")

        st.subheader("🛒 Hinterlegtes Fertigfutter aus Datenbank wählen")
        db_hund = st.session_state.get("futter_db_hund", DEFAULT_FUTTER_DB_HUND)
        db_options_h = list(db_hund.keys()) + ["➕ Neues Futter eingeben / Eigenes Etikett"]
        selected_preset_h = st.selectbox(
            "Wählen Sie ein hinterlegtes Futter oder tragen Sie neue Werte ein:",
            db_options_h,
            index=0,
            key="h_f_preset_select"
        )

        if selected_preset_h in db_hund:
            preset_data = db_hund[selected_preset_h]
            def_name_h = selected_preset_h
            def_rp_h = float(preset_data.get("rp", 25.5))
            def_rfe_h = float(preset_data.get("rfe", 15.0))
            def_rfa_h = float(preset_data.get("rfa", 3.3))
            def_ra_h = float(preset_data.get("ra", 7.5))
            def_feuchte_h = float(preset_data.get("feuchte", 10.0))
            def_ca_h = float(preset_data.get("ca", 1.2))
            def_p_h = float(preset_data.get("p", 0.9))
            def_vit_a_h = float(preset_data.get("vit_a_add", 0))
            def_vit_d_h = float(preset_data.get("vit_d_add", 0))
            def_jod_h = float(preset_data.get("jod_add", 0))
            def_zn_h = float(preset_data.get("zn_add", 0))
            def_antiox_h = str(preset_data.get("konservierung", ""))
        else:
            def_name_h = "Eigenes Hundefutter"
            def_rp_h = 25.0
            def_rfe_h = 14.0
            def_rfa_h = 3.0
            def_ra_h = 7.0
            def_feuchte_h = 10.0
            def_ca_h = 1.2
            def_p_h = 0.9
            def_vit_a_h = 0.0
            def_vit_d_h = 0.0
            def_jod_h = 0.0
            def_zn_h = 0.0
            def_antiox_h = ""

        if selected_preset_h != st.session_state.get("last_selected_preset_h") or "h_f_name" not in st.session_state:
            st.session_state["last_selected_preset_h"] = selected_preset_h
            st.session_state["h_f_name"] = def_name_h
            st.session_state["h_f_rp"] = def_rp_h
            st.session_state["h_f_rfe"] = def_rfe_h
            st.session_state["h_f_rfa"] = def_rfa_h
            st.session_state["h_f_ra"] = def_ra_h
            st.session_state["h_f_feuchte"] = def_feuchte_h
            st.session_state["h_f_ca"] = def_ca_h
            st.session_state["h_f_p"] = def_p_h
            st.session_state["h_f_vit_a"] = def_vit_a_h
            st.session_state["h_f_vit_d"] = def_vit_d_h
            st.session_state["h_f_jod"] = def_jod_h
            st.session_state["h_f_zn"] = def_zn_h
            st.session_state["h_f_antiox"] = def_antiox_h

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            h_f_name = st.text_input("Futter-Bezeichnung", key="h_f_name")
            h_f_rp = st.number_input("Rohprotein (Rp %)", step=0.5, key="h_f_rp")
            h_f_rfe = st.number_input("Rohfett (Rfe %)", step=0.5, key="h_f_rfe")
            h_f_rfa = st.number_input("Rohfaser (Rfa %)", step=0.1, key="h_f_rfa")
        with col_f2:
            h_f_ra = st.number_input("Rohasche (Ra %)", step=0.5, key="h_f_ra")
            h_f_feuchte = st.number_input("Feuchtigkeit (%)", step=0.5, key="h_f_feuchte")
            h_f_ca = st.number_input("Calcium (Ca %)", step=0.1, key="h_f_ca")
            h_f_p = st.number_input("Phosphor (P %)", step=0.1, key="h_f_p")

        with st.expander("🧪 Deklarierte Zusatzstoffe & Konservierungsmittel (optional)", expanded=(st.session_state.get("h_f_vit_a", 0) > 0 or len(st.session_state.get("h_f_antiox", "")) > 0)):
            cz1, cz2 = st.columns(2)
            with cz1:
                h_f_vit_a = st.number_input("Vitamin A (IE/kg)", step=1000.0, key="h_f_vit_a")
                h_f_vit_d = st.number_input("Vitamin D3 (IE/kg)", step=100.0, key="h_f_vit_d")
            with cz2:
                h_f_jod = st.number_input("Jod (mg/kg)", step=0.1, key="h_f_jod")
                h_f_zn = st.number_input("Zink (mg/kg)", step=10.0, key="h_f_zn")
                h_f_antiox = st.text_input("Konservierung / Antioxidanzien", placeholder="z.B. Natürliche Tocopherole / BHA", key="h_f_antiox")

        if st.button("💾 Dieses Futter in der Hunde-Datenbank speichern", key="save_f_hund"):
            st.session_state["futter_db_hund"][h_f_name] = {
                "rp": h_f_rp, "rfe": h_f_rfe, "rfa": h_f_rfa, "ra": h_f_ra,
                "feuchte": h_f_feuchte, "ca": h_f_ca, "p": h_f_p,
                "vit_a_add": h_f_vit_a, "vit_d_add": h_f_vit_d, "jod_add": h_f_jod, "zn_add": h_f_zn, "konservierung": h_f_antiox
            }
            st.success(f"'{h_f_name}' wurde erfolgreich in deiner Futter-Datenbank gespeichert!")
            st.rerun()

        h_f_typ = "Nassfutter" if h_f_feuchte > 50.0 else "Trockenfutter"
        h_f_res = analysiere_etikett(h_f_name, h_f_rp, h_f_rfe, h_f_rfa, h_f_ra, h_f_feuchte, h_f_ca, h_f_p, "Hund", h_f_modell, h_f_typ, vit_a=h_f_vit_a, vit_d=h_f_vit_d, jod=h_f_jod, zn=h_f_zn, antiox=h_f_antiox)
        st.session_state["aktuelles_futter"] = h_f_res

        st.subheader(f"📊 Auswertung Weender-Analyse: {h_f_name}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Kohlenhydrate (NfE)", f"{h_f_res['nfe_ts']} % i.TS.")
        c2.metric("Energiedichte", f"{h_f_res['kcal_100g']} kcal / 100g ({h_f_res['modell']})")
        c3.metric("Trockensubstanz", f"{h_f_res['ts_pct']} %")
        c4.metric("Ca:P-Verhältnis", f"{h_f_res['ca_p_ratio']}:1" if h_f_res['ca_p_ratio'] else "n.a.")

        if h_f_res.get("vit_a", 0) > 0 or h_f_res.get("jod", 0) > 0 or h_f_res.get("antiox", ""):
            st.markdown("##### 🧪 Deklarierte Zusatzstoffe & Konservierung")
            cz1, cz2, cz3 = st.columns(3)
            cz1.metric("Vitamin A / D3", f"{int(h_f_res['vit_a'])} IE/kg / {int(h_f_res['vit_d'])} IE/kg" if h_f_res['vit_a'] > 0 else "n.a.")
            cz2.metric("Jod / Zink", f"{h_f_res['jod']} mg/kg / {h_f_res['zn']} mg/kg" if (h_f_res['jod'] or h_f_res['zn']) else "n.a.")
            cz3.metric("Konservierung", h_f_res['antiox'] if h_f_res['antiox'] else "Keine deklariert")

        b = st.session_state.get("aktueller_bedarf_hund", bedarf_hund)
        futtermenge_g = round((b['mer_kcal'] / h_f_res['kcal_100g']) * 100.0, 1) if h_f_res['kcal_100g'] > 0 else 0
        portion_g = round(futtermenge_g / b['mahlzeiten_n'], 1)
        st.info(f"💡 Für **{b['name']}** ({b['mer_kcal']} kcal/Tag) beträgt die Tagesfuttermenge: **{futtermenge_g} g pro Tag** (ca. **{portion_g} g pro Mahlzeit**).")

    # TAB 3: HUND BARF-BAUKASTEN
    with tab3:
        st.header("🥩 Selbstgekocht- & BARF-Baukasten für Hunde (Menü A-C)")
        col_m3_1, col_m3_2 = st.columns([3, 1])
        with col_m3_1:
            st.caption("📐 **Aktives Energiedichte-Modell:** " + st.session_state.get("h_f_modell", "FEDIAF 4-Stufen-Modell"))
        with col_m3_2:
            h_f_modell_t3 = st.selectbox("Rechenmodell anpassen", ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"], index=0 if st.session_state.get("h_f_modell", "FEDIAF 4-Stufen-Modell") == "FEDIAF 4-Stufen-Modell" else 1, key="h_f_modell_t3")
            st.session_state["h_f_modell"] = h_f_modell_t3
        b = st.session_state.get("aktueller_bedarf_hund", bedarf_hund)
        
        col_t3_mode1, col_t3_mode2 = st.columns(2)
        with col_t3_mode1:
            h_tab3_modus = st.selectbox(
                "🎯 Rations-Umfang wählen",
                [
                    "Frisch-Mahlzeit für Mischfütterung (33% ME - 1 Teil-Mahlzeit pro Tag)",
                    "Vollständige Tages-Frischration (100% ME - 100% BARF / Garspeise)"
                ],
                key="h_tab3_mode_select"
            )
        h_ratio_share = 0.3333333 if "33%" in h_tab3_modus else 1.0

        col_add1, col_add2, col_add3 = st.columns([3, 2, 2])
        with col_add1:
            selected_ing = st.selectbox("🛒 Zutat aus Katalog wählen", list(INGREDIENTS_DB.keys()), key="h_tab3_ing_select")
        with col_add2:
            ing_amount = st.number_input("Tagesmenge (g oder Stk)", value=100.0, step=5.0, key="h_tab3_ing_amount")
        with col_add3:
            st.write("")
            st.write("")
            if st.button("➕ Zutat hinzufügen", key="h_tab3_add_btn"):
                st.session_state["ration_list_hund"].append({"name": selected_ing, "menge": ing_amount})
                st.success(f"{selected_ing} ({ing_amount}g) hinzugefügt!")

        r_list = st.session_state.get("ration_list_hund", [])
        p_ing_tab3 = next((item["name"] for item in r_list if INGREDIENTS_DB[item["name"]]["kat"] in ["Fleisch", "Fisch"]), "Lachsfilet (fettreich)")
        s_ing_tab3 = next((item["name"] for item in r_list if INGREDIENTS_DB[item["name"]]["kat"] == "Stärke"), "Süßkartoffel (sehr weich gekocht)")
        v_ing_tab3 = next((item["name"] for item in r_list if INGREDIENTS_DB[item["name"]]["kat"] in ["Basis-Gemüse", "Blähgemüse"]), "Zucchini (gedämpft/püriert)")

        if r_list:
            st.subheader("📋 Eingegebene Roh-Werte des Kunden")
            col_list1, col_list2 = st.columns([3, 1])
            with col_list1:
                st.table([{"Zutat": item["name"], "Eingegebene Dosis": format_menge_str(item["name"], item["menge"], INGREDIENTS_DB[item["name"]]["kat"]), "Kategorie": INGREDIENTS_DB[item["name"]]["kat"]} for item in r_list])
            with col_list2:
                if st.button("🗑️ Neue Ration beginnen", key="h_tab3_reset_btn"):
                    st.session_state["ration_list_hund"] = []
                    st.rerun()

        m_opt_tab3 = berechne_precision_menu(p_ing_tab3, s_ing_tab3, v_ing_tab3, b, ratio_share=h_ratio_share)
        st.session_state["last_tab3_opt_hund"] = m_opt_tab3

        st.markdown("---")
        st.subheader(f"✨ Die optimierte & bedarfsdeckende Rezeptur ({'Teil-Mahlzeit 33% MER' if h_ratio_share < 0.5 else 'Vollration 100% MER'})")
        if m_opt_tab3["praxis_hinweis"]:
            st.info(m_opt_tab3["praxis_hinweis"])
        st.table(m_opt_tab3["ingredients"])

        st.subheader("📊 Nährstoff-Bilanz & Echte Bedarfsdeckung (Soll vs. Ist)")
        eval_dict_h = m_opt_tab3.get("eval", evaluiere_bedarfsdeckung(m_opt_tab3["totals"], b, ratio_share=h_ratio_share))
        if eval_dict_h["is_bedarfsdeckend"]:
            st.success("🟢 Diese Rezeptur ist zu 100 % bedarfsdeckend für alle essenziellen Nährstoffe.")
        else:
            st.warning(f"⚠️ Die Rezeptur weist Teil-Abweichungen auf bei: {', '.join(eval_dict_h['deficits'])}")

        import pandas as pd
        st.dataframe(pd.DataFrame(eval_dict_h["coverage_table"]), use_container_width=True)

    # TAB 4: HUND WOCHENPLAN (MODULARER 3-SCHRITT BAUSTEIN-ASSISTENT v74)
    with tab4:
        st.header("📅 Geführter Wochenplan & Menü-Assistent (Hund v74)")
        b = st.session_state.get("aktueller_bedarf_hund", bedarf_hund)
        h_model_active = st.session_state.get("h_f_modell", "FEDIAF 4-Stufen-Modell")
        
        st.subheader("⚙️ Schritt 1: Baustein-Anzahl & Rechenmodell festlegen")
        col_rm4_1, col_rm4_2 = st.columns(2)
        with col_rm4_1:
            h_f_modell_t4 = st.radio(
                "📐 Rechenmodell für Fertigfutter-Energiedichte:",
                ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"],
                index=0 if h_model_active == "FEDIAF 4-Stufen-Modell" else 1,
                key="h_f_modell_t4"
            )
            st.session_state["h_f_modell"] = h_f_modell_t4
            h_model_active = h_f_modell_t4
        with col_rm4_2:
            st.info(f"💡 **Hinweis:** Das gewählte Modell ({h_model_active}) wird direkt für die Präzisions-Berechnung der Futter-Mengen im Wochenplan verwendet.")

        col_q1, col_q2 = st.columns(2)
        with col_q1:
            h_n_fertig = st.slider("1️⃣ Anzahl Fertigfutter-Varianten (FF A-C)", min_value=1, max_value=3, value=2, key="h_n_fertig")
        with col_q2:
            h_n_gekocht = st.slider("2️⃣ Anzahl Selbstgekocht / BARF-Varianten (Menü A-C)", min_value=0, max_value=3, value=1, key="h_n_gekocht")

        h_ratio_share = 1.0
        if h_n_gekocht > 0:
            h_gekocht_modus = st.selectbox(
                "🎯 Rations-Modus für Selbstgekocht / BARF",
                [
                    "Mischfütterung (33% ME - 1 frische Teil-Mahlzeit z.B. mittags)",
                    "Vollständige Tages-Frischration (100% ME - 100% Garspeise/BARF)"
                ],
                key="h_gekocht_modus"
            )
            h_ratio_share = 0.3333333 if "33%" in h_gekocht_modus else 1.0

        col_supp1, col_supp2, col_supp3 = st.columns(3)
        with col_supp1:
            h_supp_m = st.text_input("🌅 Morgens (Zusätze/Med)", value="", placeholder="z.B. 1 Tab. Cosequin + 2 Tab. Forthyron", key="h_supp_m")
        with col_supp2:
            h_supp_mi = st.text_input("☀️ Mittags (Zusätze/Med)", value="", placeholder="z.B. 1 TL Lachsöl", key="h_supp_mi")
        with col_supp3:
            h_supp_a = st.text_input("🌙 Abends (Zusätze/Med)", value="", placeholder="z.B. 1 Tab. Cosequin + 2 Tab. Forthyron", key="h_supp_a")

        h_extra_supps = st.text_input("⚡ Schnelleingabe-Parser ('Morgens: ... / Abends: ...')", value="", key="h_extra_supps")
        h_fm_supp, h_fmi_supp, h_fa_supp = parse_meal_supplements(h_extra_supps, h_supp_m, h_supp_mi, h_supp_a)

        st.markdown("---")
        st.subheader("🥣 Schritt 2: Bausteine konfigurieren & berechnen")

        # 2A. FERTIGFUTTER BAUSTEINE (FF A-C)
        h_fertig_results = {}
        db_keys_h = list(st.session_state["futter_db_hund"].keys())
        cols_fh = st.columns(h_n_fertig)
        for i in range(h_n_fertig):
            f_lbl = f"FF {chr(65+i)}"
            with cols_fh[i]:
                st.markdown(f"**{f_lbl}**")
                sel_f = st.selectbox(f"Futter aus DB ({f_lbl})", db_keys_h, index=i % len(db_keys_h), key=f"h_f_sel_{i}")
                f_data = st.session_state["futter_db_hund"][sel_f]
                f_eval = analysiere_etikett(
                    sel_f, f_data.get("rp", 25.5), f_data.get("rfe", 15.0), f_data.get("rfa", 3.3),
                    f_data.get("ra", 7.5), f_data.get("feuchte", 10.0), f_data.get("ca", 1.2), f_data.get("p", 0.9),
                    "Hund", h_model_active
                )
                f_typ = f_data.get("typ", "Trockenfutter" if f_data.get("feuchte", 10) <= 50 else "Nassfutter")
                kcal_100g = f_eval["kcal_100g"]
                daily_g = round((b["mer_kcal"] / kcal_100g) * 100.0, 1) if kcal_100g > 0 else 0
                port_g = round(daily_g / b["mahlzeiten_n"], 1)

                st.caption(f"📦 **{f_typ}** | {kcal_100g} kcal/100g\n📊 **Tagesmenge:** {daily_g} g/Tag\n🥣 **Pro Mahlzeit ({b['mahlzeiten_n']}x):** {port_g} g")
                h_fertig_results[f_lbl] = {
                    "name": sel_f,
                    "typ": f_typ,
                    "kcal_100g": kcal_100g,
                    "daily_g": daily_g,
                    "portion_g": port_g
                }

        # 2B. BARF / GEKOCHT BAUSTEINE (Menü A-C)
        h_menu_results = {}
        if h_n_gekocht > 0:
            st.markdown(f"#### 🥩 Selbstgekocht / BARF Bausteine (Menü A-C) ({'33% ME Teil-Mahlzeit' if h_ratio_share < 0.5 else '100% ME Vollration'})")
            protein_opts = [ing for ing, d in INGREDIENTS_DB.items() if d.get("kat") in ["Fleisch", "Fisch"] and d.get("quelle") not in b.get("unvertraeglichkeiten", [])]
            staerke_opts = [ing for ing, d in INGREDIENTS_DB.items() if d.get("kat") == "Stärke" and d.get("quelle") not in b.get("unvertraeglichkeiten", [])]
            gemuese_opts = [ing for ing, d in INGREDIENTS_DB.items() if d.get("kat") in ["Basis-Gemüse", "Blähgemüse"]]

            cols_mh = st.columns(h_n_gekocht)
            for i in range(h_n_gekocht):
                m_label = f"Menü {chr(65+i)}"
                with cols_mh[i]:
                    st.markdown(f"**🍲 {m_label}**")
                    p_sel = st.selectbox(f"Protein ({m_label})", protein_opts, index=i % len(protein_opts), key=f"h_p_{i}")
                    s_sel = st.selectbox(f"Stärke ({m_label})", staerke_opts, index=i % len(staerke_opts), key=f"h_s_{i}")
                    v_sel = st.selectbox(f"Gemüse ({m_label})", gemuese_opts, index=i % len(gemuese_opts), key=f"h_v_{i}")

                    m_res = berechne_precision_menu(p_sel, s_sel, v_sel, b, ratio_share=h_ratio_share)
                    h_menu_results[m_label] = m_res

                    st.table(m_res["ingredients"])
                    st.caption(f"📊 {m_res['totals']['kcal']} kcal | Protein: {m_res['totals']['rp']}g | Ca:P {m_res['totals']['ca_p_ratio']}:1")

        st.markdown("---")
        st.subheader("📅 Schritt 3: 7-Tage Wochenplan-Verteilung (Montag–Sonntag)")
        st.caption("Wähle für jeden Wochentag den gewünschten Baustein aus deinen erstellten Fertigfutter- und BARF-Varianten!")

        h_plan_options = []
        for f_lbl, f_info in h_fertig_results.items():
            f_short = f_info["name"].split()[0]
            h_plan_options.append(f"{f_lbl} (Fertigfutter): {f_short} ({int(f_info['daily_g'])}g/Tag)")

        if h_n_gekocht > 0 and h_ratio_share < 0.5:
            for m_lbl, m_data in h_menu_results.items():
                tot_g = round(sum([float(x['Menge'].split()[0]) for x in m_data["ingredients"] if 'g' in x['Menge']]), 0)
                if h_n_fertig > 0:
                    for f_lbl, f_info in h_fertig_results.items():
                        ff_misch_tot_g = round(f_info['daily_g'] * (1.0 - h_ratio_share), 1)
                        h_plan_options.append(f"Misch: {f_lbl} ({ff_misch_tot_g}g Fertigfutter) + {m_lbl} ({int(tot_g)}g Selbstgekocht/BARF)")
                else:
                    h_plan_options.append(f"Misch: Fertigfutter + {m_lbl} ({int(tot_g)}g Selbstgekocht/BARF)")

        if h_n_gekocht > 0 and h_ratio_share >= 0.5:
            for m_lbl, m_data in h_menu_results.items():
                h_plan_options.append(f"Vollration: {m_lbl} (Selbstgekocht / BARF)")

        if not h_plan_options:
            h_plan_options = ["Standard-Fütterung"]

        tage_liste = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
        h_wochenplan_rows = []

        cols_days = st.columns(7)
        for idx, tag in enumerate(tage_liste):
            with cols_days[idx]:
                st.markdown(f"**{tag}**")
                day_choice = st.selectbox(f"{tag}", h_plan_options, index=idx % len(h_plan_options), key=f"h_day_sel_{idx}", label_visibility="collapsed")

                if "Misch:" in day_choice:
                    f_lbl = "FF A" if "FF A" in day_choice else ("FF B" if "FF B" in day_choice else "FF C")
                    m_lbl = "Menü A" if "Menü A" in day_choice else ("Menü B" if "Menü B" in day_choice else "Menü C")
                    
                    f_info = h_fertig_results.get(f_lbl, list(h_fertig_results.values())[0] if h_fertig_results else {"name": "Fertigfutter", "daily_g": 300.0})
                    m_data = h_menu_results.get(m_lbl, list(h_menu_results.values())[0] if h_menu_results else {"ingredients": []})
                    tot_m_g = round(sum([float(x['Menge'].split()[0]) for x in m_data["ingredients"] if 'g' in x['Menge']]), 0)

                    ff_misch_daily_g = round(f_info["daily_g"] * (1.0 - h_ratio_share), 1)
                    
                    if b["mahlzeiten_n"] >= 3:
                        ff_meal_g = round(ff_misch_daily_g / 2.0, 1)
                        morgens = f"🌅 {f_info['name']} ({ff_meal_g}g)"
                        mittags = f"☀️ 🍲 {m_lbl} ({int(tot_m_g)}g Selbstgekocht/BARF)"
                        abends = f"🌙 {f_info['name']} ({ff_meal_g}g)"
                    else:
                        morgens = f"🌅 {f_info['name']} ({ff_misch_daily_g}g)"
                        mittags = "Ruhephase"
                        abends = f"🌙 🍲 {m_lbl} ({int(tot_m_g)}g Selbstgekocht/BARF)"

                    if h_fm_supp: morgens += f" + 💊 {h_fm_supp}"
                    if h_fmi_supp and mittags != "Ruhephase": mittags += f" + 💊 {h_fmi_supp}"
                    if h_fa_supp: abends += f" + 💊 {h_fa_supp}"

                elif "Voll-BARF" in day_choice or "Menü" in day_choice:
                    m_lbl = "Menü A" if "Menü A" in day_choice else ("Menü B" if "Menü B" in day_choice else "Menü C")
                    m_data = h_menu_results.get(m_lbl, list(h_menu_results.values())[0] if h_menu_results else {"ingredients": []})
                    tot_m_g = round(sum([float(x['Menge'].split()[0]) for x in m_data["ingredients"] if 'g' in x['Menge']]), 0)
                    
                    if b["mahlzeiten_n"] >= 3:
                        p_g = round(tot_m_g / 3.0, 1)
                        morgens = f"🌅 🍲 {m_lbl} ({p_g}g)"
                        mittags = f"☀️ 🍲 {m_lbl} ({p_g}g)"
                        abends = f"🌙 🍲 {m_lbl} ({p_g}g)"
                    else:
                        p_g = round(tot_m_g / 2.0, 1)
                        morgens = f"🌅 🍲 {m_lbl} ({p_g}g)"
                        mittags = "Ruhephase"
                        abends = f"🌙 🍲 {m_lbl} ({p_g}g)"

                    if h_fm_supp: morgens += f" + 💊 {h_fm_supp}"
                    if h_fmi_supp and mittags != "Ruhephase": mittags += f" + 💊 {h_fmi_supp}"
                    if h_fa_supp: abends += f" + 💊 {h_fa_supp}"

                else:
                    f_lbl = "FF A" if "FF A" in day_choice else ("FF B" if "FF B" in day_choice else "FF C")
                    f_info = h_fertig_results.get(f_lbl, list(h_fertig_results.values())[0] if h_fertig_results else {"name": "Fertigfutter", "portion_g": 200.0, "daily_g": 300.0})
                    p_g = round(f_info["daily_g"] / float(b["mahlzeiten_n"]), 1)

                    morgens = f"🌅 {f_info['name']} ({p_g}g)"
                    if h_fm_supp: morgens += f" + 💊 {h_fm_supp}"

                    mittags = f"☀️ {f_info['name']} ({p_g}g)" if b["mahlzeiten_n"] >= 3 else "Ruhephase"
                    if h_fmi_supp and mittags != "Ruhephase": mittags += f" + 💊 {h_fmi_supp}"

                    abends = f"🌙 {f_info['name']} ({p_g}g)"
                    if h_fa_supp: abends += f" + 💊 {h_fa_supp}"

                h_wochenplan_rows.append({"Tag": tag, "Morgens": morgens, "Mittags": mittags, "Abends": abends})

        st.session_state["aktueller_wochenplan_hund"] = h_wochenplan_rows
        st.session_state["menu_results_pdf_hund"] = h_menu_results

        st.markdown("---")
        st.subheader("📋 Generierte 7-Tage Fütterungs-Übersicht für das Hunde-PDF")
        st.table(h_wochenplan_rows)

        treat_kcal = round(b["mer_kcal"] * 0.10, 0)
        st.info(f"🦴 **Die 10%-Leckerli- & Snack-Regelung für {b['name']}:** Das tägliche Belohnungs-Budget beträgt max. **{treat_kcal} kcal / Tag** (10% MER). Bei Gewichtsreduktion Leckerlis zwingend vom Hauptfutter abziehen (~20g Trockenfutter weniger). Empfohlene Snacks: Mageres Putenfeinblatt, Rinderlunge oder Gurken- und Möhrenstückchen.")

    with tab5:
        st.header("📑 Modulare PDF-Exports & Praxis-Branding (Hund)")
        st.write("Passen Sie Ihre Praxis-Kopf- und Fußzeilen an und laden Sie maßgeschneiderte PDF-Gutachten für Patienten und Berater herunter.")

        st.subheader("⚙️ Praxis-Branding & PDF-Header Konfiguration")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.session_state["pdf_header_text"] = st.text_input(
                "📋 Praxis-Kopfzeile (PDF Header)",
                value=st.session_state.get("pdf_header_text", "Praxis für Tierernährungsberatung | Individueller Fütterungsplan"),
                key="h_pdf_header_input"
            )
        with col_b2:
            st.session_state["pdf_footer_text"] = st.text_input(
                "👣 Praxis-Fußzeile (PDF Footer)",
                value=st.session_state.get("pdf_footer_text", "Praxis-Handout Tierernährungsberatung | Evidenzbasierte FEDIAF/NRC Rationsberechnung"),
                key="h_pdf_footer_input"
            )

        st.markdown("---")
        st.subheader("📥 5 Modulare PDF-Gutachten zum Download (Hund)")
        b = st.session_state.get("aktueller_bedarf_hund", bedarf_hund)
        h_model_active = st.session_state.get("h_f_modell", "FEDIAF 4-Stufen-Modell")
        f_raw = st.session_state.get("aktuelles_futter", {"name": "Belcando Getreidefrei Adult Ozean", "rp": 25.5, "rfe": 15.0, "rfa": 3.3, "ra": 7.5, "feuchte": 10.0, "ca": 1.2, "p": 0.9})
        f_res = analysiere_etikett(f_raw.get("name","Belcando"), f_raw.get("rp",25.5), f_raw.get("rfe",15.0), f_raw.get("rfa",3.3), f_raw.get("ra",7.5), f_raw.get("feuchte",10.0), f_raw.get("ca",1.2), f_raw.get("p",0.9), "Hund", h_model_active)
        w_rows = st.session_state.get("aktueller_wochenplan_hund", [])
        m_res_pdf = st.session_state.get("menu_results_pdf_hund", {})

        m_res_tmp = st.session_state.get("last_tab3_opt_hund", berechne_precision_menu("Lachsfilet (fettreich)", "Süßkartoffel (sehr weich gekocht)", "Zucchini (gedämpft/püriert)", b))
        
        pdf1_bytes = build_bedarfsanalyse_pdf_bytes(b)
        docx1_bytes = build_bedarfsanalyse_docx_bytes(b)
        
        pdf2_bytes = build_fertigfutter_pdf_bytes(b, f_res)
        docx2_bytes = build_fertigfutter_docx_bytes(b, f_res)
        
        pdf3_bytes = build_eigenration_pdf_bytes(b, st.session_state.get("ration_list_hund", []), m_res_tmp["ingredients"], m_res_tmp["totals"])
        docx3_bytes = build_eigenration_docx_bytes(b, st.session_state.get("ration_list_hund", []), m_res_tmp["ingredients"], m_res_tmp["totals"])
        
        if w_rows:
            pdf4_bytes = build_wochenplan_pdf_bytes(b, w_rows, m_res_pdf)
            docx4_bytes = build_wochenplan_docx_bytes(b, w_rows, m_res_pdf)
        else:
            pdf4_bytes = docx4_bytes = None
            
        pdf5_bytes = build_berater_dossier_pdf_bytes(b, b, f_res, m_res_pdf, w_rows)
        docx5_bytes = build_berater_dossier_docx_bytes(b, b, f_res, m_res_pdf, w_rows)

        st.markdown("##### 📄 Export-Option 1: Nährstoff-Bedarfsanalyse")
        c_ex1_a, c_ex1_b = st.columns(2)
        with c_ex1_a: st.download_button("⬇️ 1. Bedarfsanalyse PDF (Hund)", data=pdf1_bytes, file_name=f"bedarfsanalyse-{b['name'].lower()}.pdf", mime="application/pdf", key="h_pdf1")
        with c_ex1_b: st.download_button("📝 1. Bedarfsanalyse Word (.docx) (Hund)", data=docx1_bytes, file_name=f"bedarfsanalyse-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="h_docx1")

        st.markdown("##### 📄 Export-Option 2: Fertigfutter-Evaluierung")
        c_ex2_a, c_ex2_b = st.columns(2)
        with c_ex2_a: st.download_button("⬇️ 2. Fertigfutter-Analyse PDF (Hund)", data=pdf2_bytes, file_name=f"fertigfutter-analyse-{b['name'].lower()}.pdf", mime="application/pdf", key="h_pdf2")
        with c_ex2_b: st.download_button("📝 2. Fertigfutter-Analyse Word (.docx) (Hund)", data=docx2_bytes, file_name=f"fertigfutter-analyse-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="h_docx2")

        st.markdown("##### 📄 Export-Option 3: BARF & Eigenrations-Analyse")
        c_ex3_a, c_ex3_b = st.columns(2)
        with c_ex3_a: st.download_button("⬇️ 3. BARF- & Eigenrations-Analyse PDF (Hund)", data=pdf3_bytes, file_name=f"eigenration-analyse-{b['name'].lower()}.pdf", mime="application/pdf", key="h_pdf3")
        with c_ex3_b: st.download_button("📝 3. BARF- & Eigenrations-Analyse Word (.docx) (Hund)", data=docx3_bytes, file_name=f"eigenration-analyse-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="h_docx3")

        st.markdown("##### 📄 Export-Option 4: Kunden-Fütterungsplan Handout")
        c_ex4_a, c_ex4_b = st.columns(2)
        if w_rows:
            with c_ex4_a: st.download_button("⬇️ 4. Kunden-Fütterungsplan Handout PDF (Hund)", data=pdf4_bytes, file_name=f"futterplan-{b['name'].lower()}-handout.pdf", mime="application/pdf", key="h_pdf4")
            with c_ex4_b: st.download_button("📝 4. Kunden-Fütterungsplan Handout Word (.docx) (Hund)", data=docx4_bytes, file_name=f"futterplan-{b['name'].lower()}-handout.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="h_docx4")
        else:
            st.info("💡 Generieren Sie zuerst in Tab 4 einen Wochenplan.")

        st.markdown("##### 📄 Export-Option 5: Clinical Vet-Dossier")
        c_ex5_a, c_ex5_b = st.columns(2)
        with c_ex5_a: st.download_button("⬇️ 5. Clinical Vet-Dossier & Berater-Gutachten PDF (Hund)", data=pdf5_bytes, file_name=f"berater-dossier-{b['name'].lower()}.pdf", mime="application/pdf", key="h_pdf5")
        with c_ex5_b: st.download_button("📝 5. Clinical Vet-Dossier Word (.docx) (Hund)", data=docx5_bytes, file_name=f"berater-dossier-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="h_docx5")

    # TAB 7: HUND NEUES FUTTER ERFASSEN & DB-VERWALTUNG
    with tab7:
        st.header("➕ Neues Fertigfutter erfassen & DB-Verwaltung (Hund)")
        st.write("Erfassen Sie neue Hundefuttermittel für Ihre aktuelle Beratungssitzung (Weg A) oder generieren Sie einen vorgefertigten KI-Prompt zur dauerhaften Speicherung im App-Code (Weg B).")

        st.subheader("1️⃣ Produktdaten & Weender-Deklaration eingeben")
        col_add1, col_add2 = st.columns(2)
        with col_add1:
            st.markdown("##### 📦 Weender-Analytische Bestandteile")
            add_h_name = st.text_input("Produktname & Marke (Hund)", value="", placeholder="z.B. Josera Hypoallergenic", key="add_h_name")
            add_h_typ = st.selectbox("Futtermittel-Typ (Hund)", ["Trockenfutter", "Nassfutter"], index=0, key="add_h_typ")
            add_h_feuchte = st.number_input("Feuchtigkeit (%)", value=10.0 if add_h_typ == "Trockenfutter" else 78.0, step=0.5, key="add_h_feuchte")
            add_h_rp = st.number_input("Rohprotein Rp (%)", value=22.0, step=0.5, key="add_h_rp")
            add_h_rfe = st.number_input("Rohfett Rfe (%)", value=12.0, step=0.5, key="add_h_rfe")
            add_h_rfa = st.number_input("Rohfaser Rfa (%)", value=2.5, step=0.1, key="add_h_rfa")
            add_h_ra = st.number_input("Rohasche Ra (%)", value=6.5, step=0.1, key="add_h_ra")
            add_h_ca = st.number_input("Calcium Ca (%)", value=1.1, step=0.05, key="add_h_ca")
            add_h_p = st.number_input("Phosphor P (%)", value=0.8, step=0.05, key="add_h_p")

        with col_add2:
            st.markdown("##### 🧪 Deklarierte Zusatzstoffe & Konservierung (optional)")
            add_h_vita = st.number_input("Vitamin A (IE / kg)", value=15000.0, step=500.0, key="add_h_vita")
            add_h_vitd = st.number_input("Vitamin D3 (IE / kg)", value=1500.0, step=100.0, key="add_h_vitd")
            add_h_jod = st.number_input("Jod (mg / kg)", value=1.5, step=0.1, key="add_h_jod")
            add_h_zn = st.number_input("Zink (mg / kg)", value=120.0, step=5.0, key="add_h_zn")
            add_h_cu = st.number_input("Kupfer (mg / kg)", value=12.0, step=1.0, key="add_h_cu")
            add_h_antiox = st.selectbox("Antioxidanzien / Konservierung", ["Natürliche Tocopherole (Vit. E)", "BHA / BHT (E 320 / E 321)", "Keine Angaben / Unkonserviert"], key="add_h_antiox")

        st.markdown("---")
        st.subheader("2️⃣ Futter speichern oder KI-Prompt generieren")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown("##### 💾 Weg A: Futter für die aktuelle Sitzung speichern")
            st.write("Das Futter steht sofort in allen Dropdowns der aktuellen Beratung zur Verfügung.")
            if st.button("💾 Futter in Sitzung speichern (Hund)", key="add_h_save_btn"):
                if not add_h_name.strip():
                    st.error("⚠️ Bitte geben Sie eine Futter-Bezeichnung ein!")
                else:
                    st.session_state["futter_db_hund"][add_h_name.strip()] = {
                        "rp": add_h_rp, "rfe": add_h_rfe, "rfa": add_h_rfa, "ra": add_h_ra,
                        "feuchte": add_h_feuchte, "ca": add_h_ca, "p": add_h_p, "typ": add_h_typ,
                        "vit_a": add_h_vita, "vit_d": add_h_vitd, "jod": add_h_jod, "zn": add_h_zn, "cu": add_h_cu, "antiox": add_h_antiox
                    }
                    st.success(f"✅ '{add_h_name.strip()}' wurde in deiner Hunde-Datenbank gespeichert!")

        with col_b2:
            st.markdown("##### 🤖 Weg B: KI-Prompt für dauerhafte Code-Speicherung")
            st.write("Kopieren Sie diesen generierten Prompt und senden Sie ihn im Chat an Gemini Notebook, um das Futter fest im Quellcode zu verankern.")
            p_name = add_h_name.strip() if add_h_name.strip() else "Neues Hundefutter"
            prompt_h_text = f"""Lieber KI-Assistent, bitte füge folgendes Hundefutter dauerhaft in die App-Stamm-Datenbank ein:

• Produktname: {p_name}
• Spezies: Hund
• Typ: {add_h_typ}
• Analytische Bestandteile:
  - Rohprotein: {add_h_rp}%
  - Rohfett: {add_h_rfe}%
  - Rohfaser: {add_h_rfa}%
  - Rohasche: {add_h_ra}%
  - Feuchtigkeit: {add_h_feuchte}%
  - Calcium: {add_h_ca}%
  - Phosphor: {add_h_p}%
• Deklarierte Zusatzstoffe:
  - Vitamin A: {add_h_vita} IE/kg
  - Vitamin D3: {add_h_vitd} IE/kg
  - Jod: {add_h_jod} mg/kg
  - Zink: {add_h_zn} mg/kg
  - Kupfer: {add_h_cu} mg/kg
  - Konservierung: {add_h_antiox}"""
            st.text_area("📋 Vorlage-Prompt zum Kopieren für den Chat:", value=prompt_h_text, height=220, key="prompt_h_txt_area")

    # TAB 6: HUND NACHKONTROLLE
    with tab6:
        st.header("📈 14- & 30-Tage Nachkontrolle & Wiege-Protokoll (Hund)")
        b = st.session_state.get("aktueller_bedarf_hund", bedarf_hund)
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            w0 = st.number_input("Startgewicht (kg)", value=b["ist_gw"], step=0.1, key="h_w0")
            w30 = st.number_input("Tag 30 Gewicht (kg)", value=round(b["ist_gw"]-3.1, 1), step=0.1, key="h_w30")
        with col_w2:
            rate_per_week = round(((w0 - w30) / 30.0) * 7.0, 2)
            target_min = round(w0 * 0.010, 2)
            target_max = round(w0 * 0.015, 2)
            st.metric("Abnahmerate", f"-{rate_per_week} kg / Woche", f"Ziel: -{target_min} bis -{target_max} kg/Woche")


        if b.get("lebensphase") == "Junior / Wachstum" or st.checkbox("📈 Rassespezifische Wachstumskurve (Soll vs. Ist) anzeigen", key="h_show_growth"):
            st.markdown("---")
            st.subheader("🦴 Rassespezifische Hunde-Wachstumskurve (5-Klassen-Modell)")
            tab1_end_gw_h = float(b.get("end_gw", b["soll_gw"]))
            
            # Auto-determine class index
            if tab1_end_gw_h < 5.0:
                def_class_idx = 0
            elif tab1_end_gw_h < 10.0:
                def_class_idx = 1
            elif tab1_end_gw_h < 25.0:
                def_class_idx = 2
            elif tab1_end_gw_h <= 45.0:
                def_class_idx = 3
            else:
                def_class_idx = 4

            dog_classes = [
                "Hund · 1 · Toy (sehr klein, < 5 kg | 8–10 Monate)",
                "Hund · 2 · Small (klein, 5–10 kg | 10–12 Monate)",
                "Hund · 3 · Medium (mittel, 10–25 kg | 12–14 Monate)",
                "Hund · 4 · Large (groß, 25–45 kg | 14–18 Monate)",
                "Hund · 5 · Giant (Riesen, > 45 kg | 24 Monate)"
            ]
            
            col_dh1, col_dh2 = st.columns(2)
            with col_dh1:
                h_growth_class = st.selectbox(
                    "📐 Rasse- & Größenklasse für Entwicklungsmodell:",
                    dog_classes,
                    index=def_class_idx,
                    key="h_growth_class_sel"
                )

            if "last_h_growth_class" not in st.session_state or st.session_state["last_h_growth_class"] != h_growth_class:
                st.session_state["last_h_growth_class"] = h_growth_class
                if "Toy" in h_growth_class:
                    new_default_target_h = tab1_end_gw_h if (1.0 <= tab1_end_gw_h < 5.0) else 3.5
                elif "Small" in h_growth_class:
                    new_default_target_h = tab1_end_gw_h if (5.0 <= tab1_end_gw_h < 10.0) else 7.5
                elif "Medium" in h_growth_class:
                    new_default_target_h = tab1_end_gw_h if (10.0 <= tab1_end_gw_h < 25.0) else 18.0
                elif "Large" in h_growth_class:
                    new_default_target_h = tab1_end_gw_h if (25.0 <= tab1_end_gw_h <= 45.0) else 35.0
                else:
                    new_default_target_h = tab1_end_gw_h if (tab1_end_gw_h > 45.0) else 60.0
                st.session_state["h_end_gw_input"] = new_default_target_h

            with col_dh2:
                end_target = st.number_input(
                    "🎯 Erwartetes Endgewicht (ausgewachsen in kg):",
                    min_value=1.0,
                    max_value=120.0,
                    value=float(st.session_state.get("h_end_gw_input", tab1_end_gw_h)),
                    step=0.5,
                    key="h_end_gw_input"
                )

            months_h = [1, 2, 3, 4, 6, 8, 10, 12, 14, 18, 24]

            if "Toy" in h_growth_class:
                pcts_h = [0.25, 0.45, 0.65, 0.80, 0.95, 0.99, 1.00, 1.00, 1.00, 1.00, 1.00]
                dauer_str = "8–10 Monate"
                verlauf_str = "Sehr schneller früher Anstieg; frühes Abflachen der Gewichtskurve."
            elif "Small" in h_growth_class:
                pcts_h = [0.20, 0.38, 0.55, 0.70, 0.88, 0.95, 0.98, 1.00, 1.00, 1.00, 1.00]
                dauer_str = "10–12 Monate"
                verlauf_str = "Schneller, harmonischer Verlauf bis etwa zum ersten Geburtstag."
            elif "Medium" in h_growth_class:
                pcts_h = [0.15, 0.30, 0.45, 0.60, 0.80, 0.90, 0.95, 0.98, 1.00, 1.00, 1.00]
                dauer_str = "12–14 Monate"
                verlauf_str = "Mittlere Entwicklungsdauer; gleichmäßig abflachender Verlauf."
            elif "Large" in h_growth_class:
                pcts_h = [0.12, 0.25, 0.38, 0.50, 0.68, 0.80, 0.88, 0.92, 0.96, 1.00, 1.00]
                dauer_str = "14–18 Monate"
                verlauf_str = "Längere Entwicklung; spätere Zunahme an Substanz wird berücksichtigt."
            else: # Giant / Riesen
                pcts_h = [0.10, 0.20, 0.32, 0.44, 0.60, 0.72, 0.80, 0.86, 0.90, 0.96, 1.00]
                dauer_str = "24 Monate"
                verlauf_str = "Zwei überlagerte Modellphasen: früher Hauptzuwachs und spätere langsame Annäherung/Ausreifung."

            soll_weights_h = [round(end_target * p, 2) for p in pcts_h]
            soll_upper_h = [round(w * 1.10, 2) for w in soll_weights_h]
            soll_lower_h = [round(w * 0.90, 2) for w in soll_weights_h]

            st.info(f"💡 **Modell-Klasse:** {h_growth_class.split('(')[0].strip()}\n• **Erwartetes Endgewicht:** {end_target} kg\n• **Entwicklungsdauer:** {dauer_str}\n• **Modellverlauf:** {verlauf_str}")

            # Session State for Weighing History (Hund)
            patient_mon_h = round(float(b.get("alter", 0.5)) * 12.0, 1)
            patient_kg_h = float(b.get("ist_gw", 10.0))
            if "h_wiegehistorie" not in st.session_state:
                st.session_state["h_wiegehistorie"] = [
                    {"Monat": patient_mon_h, "Gewicht": patient_kg_h, "Notiz": "Erstaufnahme"}
                ]

            st.markdown("##### 📝 Fortlaufende Messpunkt-Erfassung (Wiegehistorie)")
            col_wh1, col_wh2, col_wh3, col_wh4 = st.columns([1.5, 1.5, 3, 1.5])
            with col_wh1:
                new_m_h = st.number_input("Alter (Monate)", value=patient_mon_h if patient_mon_h > 0 else 2.0, step=0.5, key="new_m_h")
            with col_wh2:
                new_kg_h = st.number_input("Gewicht (kg)", value=patient_kg_h if patient_kg_h > 0 else 5.0, step=0.1, key="new_kg_h")
            with col_wh3:
                new_note_h = st.text_input("Befund / Notiz", value="Routine-Check", key="new_note_h")
            with col_wh4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Hinzufügen", key="add_wiege_h"):
                    st.session_state["h_wiegehistorie"].append({
                        "Monat": float(new_m_h),
                        "Gewicht": float(new_kg_h),
                        "Notiz": str(new_note_h)
                    })
                    st.session_state["h_wiegehistorie"] = sorted(st.session_state["h_wiegehistorie"], key=lambda x: x["Monat"])
                    st.rerun()

            import pandas as pd
            if st.session_state["h_wiegehistorie"]:
                hist_df_h = pd.DataFrame(st.session_state["h_wiegehistorie"])
                st.dataframe(hist_df_h, use_container_width=True)
                
                col_del_h1, col_del_h2 = st.columns(2)
                with col_del_h1:
                    opts_del_h = [f"{i+1}. {pt['Monat']} Mon. - {pt['Gewicht']} kg ({pt['Notiz']})" for i, pt in enumerate(st.session_state["h_wiegehistorie"])]
                    selected_del_h = st.selectbox("Eintrag zum Löschen auswählen:", opts_del_h, key="sel_del_pt_h")
                    if st.button("❌ Ausgewählten Eintrag löschen", key="btn_del_single_h"):
                        idx_del = int(selected_del_h.split(".")[0]) - 1
                        st.session_state["h_wiegehistorie"].pop(idx_del)
                        st.rerun()
                with col_del_h2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.button("🔄 Auf Erstaufnahme zurücksetzen", key="reset_wiege_h"):
                            st.session_state["h_wiegehistorie"] = [{"Monat": patient_mon_h, "Gewicht": patient_kg_h, "Notiz": "Erstaufnahme"}]
                            st.rerun()
                    with col_b2:
                        if st.button("🧹 Historie komplett leeren", key="clear_wiege_h"):
                            st.session_state["h_wiegehistorie"] = []
                            st.rerun()

            import matplotlib.pyplot as plt

            fig_h, ax_h = plt.subplots(figsize=(9.5, 4.5))

            # Fill tolerance band
            ax_h.fill_between(months_h, soll_lower_h, soll_upper_h, color="#1f77b4", alpha=0.15, label="±10% Toleranzbereich (Gesunder Korridor)")
            ax_h.plot(months_h, soll_lower_h, linestyle="--", color="#1f77b4", alpha=0.4, linewidth=1)
            ax_h.plot(months_h, soll_upper_h, linestyle="--", color="#1f77b4", alpha=0.4, linewidth=1)
            ax_h.plot(months_h, soll_weights_h, marker="o", color="#1f77b4", linewidth=2.5, label="Soll-Gewichtskurve (kg)")

            # Plot historical weighing points
            if st.session_state["h_wiegehistorie"]:
                h_m_list = [pt["Monat"] for pt in st.session_state["h_wiegehistorie"]]
                h_kg_list = [pt["Gewicht"] for pt in st.session_state["h_wiegehistorie"]]
                ax_h.plot(h_m_list, h_kg_list, marker="s", color="#d62728", linewidth=2.2, linestyle="-", label=f"Gemessener Verlauf: {b['name']}")
                for pt in st.session_state["h_wiegehistorie"]:
                    ax_h.annotate(f"{pt['Gewicht']}kg", (pt["Monat"], pt["Gewicht"]), textcoords="offset points", xytext=(0,7), ha='center', fontsize=8, fontweight='bold', color="#d62728")

            ax_h.set_title(f"Soll-Wachstumskurve Hund mit ±10% Toleranzband ({h_growth_class.split('(')[0].strip()})", fontsize=11, fontweight="bold")
            ax_h.set_xlabel("Alter in Monaten", fontsize=10)
            ax_h.set_ylabel("Körpergewicht (kg)", fontsize=10)
            ax_h.grid(True, linestyle="--", alpha=0.6)
            ax_h.set_xticks(months_h)
            ax_h.legend(loc="upper left", fontsize=8.5)
            plt.tight_layout()
            st.pyplot(fig_h)
            plt.close(fig_h)

            st.markdown("##### 📊 Soll-Gewichtstabelle nach Alter (Monate)")
            table_h_df = pd.DataFrame({
                "Alter": [f"{m} Mon." for m in months_h],
                "Soll-Gewicht (kg)": [f"{w:.2f} kg" for w in soll_weights_h],
                "Prozent vom Endgewicht": [f"{int(p*100)} %" for p in pcts_h]
            })
            st.dataframe(table_h_df, use_container_width=True)

            pdf_growth_h_bytes = build_growth_curve_pdf_bytes(
                b, h_growth_class, end_target, months_h, soll_weights_h, soll_lower_h, soll_upper_h, st.session_state.get("h_wiegehistorie", []), is_katze=False
            )
            docx_growth_h_bytes = build_growth_curve_docx_bytes(
                b, h_growth_class, end_target, months_h, soll_weights_h, soll_lower_h, soll_upper_h, st.session_state.get("h_wiegehistorie", []), is_katze=False
            )
            c_gr_h1, c_gr_h2 = st.columns(2)
            with c_gr_h1:
                st.download_button(
                    "📄 Hunde-Wachstumskurve PDF Report herunterladen",
                    data=pdf_growth_h_bytes,
                    file_name=f"Wachstumskurve_Hund_{b['name']}.pdf",
                    mime="application/pdf",
                    key="dl_pdf_growth_h"
                )
            with c_gr_h2:
                st.download_button(
                    "📝 Hunde-Wachstumskurve Word (.docx) Report herunterladen",
                    data=docx_growth_h_bytes,
                    file_name=f"Wachstumskurve_Hund_{b['name']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_docx_growth_h"
                )
# -----------------------------------------------------------------------------
# 3. KATZEN-BERATUNGS-SUITE (SPEZIES: KATZE)
# -----------------------------------------------------------------------------
elif st.session_state["selected_species"] == "Katze":
    st.title("🐱 Katzen-Ernährungsberatungs-Suite v74")
    st.caption("Spezialisierte Karnivoren-Ernährungsberatung | Taurin, Preformes Vit-A & Feline Precision Engine")

    tab1_k, tab2_k, tab3_k, tab4_k, tab5_k, tab6_k, tab7_k = st.tabs([
        "📊 1. Katzen-Bedarfsrechner", 
        "🏷️ 2. Katzen-Futteranalysator & DB", 
        "🥩 3. Katzen-Selbstgekocht / BARF", 
        "📅 4. Geführter Katzen-Wochenplan",
        "📑 5. Modulare Katzen-PDF-Exports", 
        "📈 6. Katzen-Nachkontrolle & Wiege-Protokoll",
        "➕ 7. Neues Katzenfutter erfassen & DB"
    ])

    # TAB 1: KATZE BEDARFSRECHNER
    with tab1_k:
        st.header("Bedarfsanalyse für Katzen (Anamnese & Feline Sollwerte)")
        col1, col2 = st.columns(2)
        with col1:
            col_k_id1, col_k_id2 = st.columns([1, 1])
            with col_k_id1:
                k_name = st.text_input("Patienten-Name (Katze)", value="Mimi", key="k_name")
            with col_k_id2:
                k_fall_id = st.text_input("Fall-ID / Aktenzeichen", value="FALL-2026-1002", key="k_fall_id")
            
            k_ist_gw = st.number_input("Ist-Gewicht (kg)", value=5.5, step=0.1, key="k_ist_gw")
            k_soll_gw = st.number_input("Soll- / Ziel-Gewicht (kg)", value=4.5, step=0.1, key="k_soll_gw")
            st.markdown("##### 📅 Alterseingabe (Katze)")
            col_k_a1, col_k_a2 = st.columns([1, 1])
            with col_k_a1:
                k_alter_mode = st.selectbox("Einheit", ["Jahre", "Monate", "Wochen", "Geburtsdatum"], index=0, key="k_alter_mode")
            with col_k_a2:
                if k_alter_mode == "Jahre":
                    k_alter = st.number_input("Alter (Jahre)", value=6.0, step=0.1, key="k_alter")
                elif k_alter_mode == "Monate":
                    k_mon = st.number_input("Alter (Monate)", value=72.0, step=1.0, key="k_alter_mon")
                    k_alter = round(k_mon / 12.0, 2)
                    st.caption(f"📅 Berechnet: **{round(k_mon * 4.33, 1)} Wochen** (~{k_alter} Jahre)")
                elif k_alter_mode == "Wochen":
                    k_wochen = st.number_input("Alter (Wochen)", value=33.0, step=1.0, key="k_alter_wochen")
                    k_alter = round(k_wochen / 52.0, 2)
                    st.caption(f"📅 Berechnet: **{k_wochen} Wochen** (~{round(k_wochen/4.33, 1)} Monate / {k_alter} Jahre)")
                else: # Geburtsdatum
                    import datetime
                    default_k_dob = datetime.date.today() - datetime.timedelta(days=int(6.0 * 365.25))
                    k_dob = st.date_input("Geburtsdatum", value=default_k_dob, key="k_dob")
                    days_old = max(1, (datetime.date.today() - k_dob).days)
                    k_alter = round(days_old / 365.25, 2)
                    st.caption(f"📅 Berechnet: **{round(days_old/7.0,1)} Wochen** (~{round(days_old/30.4,1)} Monate / {k_alter} Jahre)")
        with col2:
            k_lebensphase = st.selectbox("Lebensphase", ["Adult", "Senior", "Kitten / Wachstum"], index=0, key="k_phase")
            k_kastriert = st.selectbox("Kastriert?", ["Ja", "Nein"], index=0, key="k_kast")
            k_ziel = st.selectbox("Klinisches Ziel", ["Gewichtsreduktion", "Gewichtserhalt", "Auffütterung / Zunahme"], index=0, key="k_ziel")
            if k_lebensphase in ["Junior / Wachstum", "Kitten / Wachstum", "Kitten"]:
                st.info("💡 **Wachstums-Hinweis:** Bei Kitten im Wachstum ist der erhöhte Energiebedarf in der Formel enthalten. Wählen Sie beim klinischen Ziel im Normalfall **'Gewichtserhalt'**, um hyperplastische Adipositas (übermäßige Fettzellenbildung) zu verhindern!")
            k_aktivitaet = st.selectbox("Aktivitätslevel", ["Indoor / Inaktiv", "Normal", "Freigänger / Aktiv"], index=0, key="k_akt")
            k_bcs = st.selectbox("Body Condition Score (BCS 1–9)", options=[1, 2, 3, 4, 5, 6, 7, 8, 9], index=6, help="1 = Kachektisch, 5 = Idealgewicht, 9 = Schwerst Adipös", key="k_bcs")
            k_bcs_title, k_bcs_desc = get_bcs_info(k_bcs, "Katze")
            st.info(f"⚖️ **{k_bcs_title}**\n{k_bcs_desc}")
            k_labor_txt = st.text_input("Klinische Laborwerte & Organparameter (optional)", value="Kreatinin: 1.4 mg/dl | SDMA: 11 µg/dl | Urin-pH: 6.5", key="k_labor")
            k_mahlzeiten_n = st.selectbox("Anzahl Mahlzeiten pro Tag", [1, 2, 3, 4, 5, 6], index=2, key="k_mahlzeiten_n")

        st.subheader("🛡️ Patientenschutz: Unverträglichkeiten & Medikation")
        c_unv1, c_unv2 = st.columns(2)
        with c_unv1:
            k_unvert = st.multiselect("Allergie- / Unverträglichkeits-Filter", ["Geflügel", "Rind", "Fisch", "Getreide"], default=[], key="k_unvert")
        with c_unv2:
            k_med_txt = st.text_input("Medikation / Anamnese", value="Keine", key="k_med")
            k_med_list = [m.strip() for m in k_med_txt.split(",") if m.strip()] if k_med_txt and k_med_txt != "Keine" else []

        k_end_gw = k_soll_gw
        if k_lebensphase == "Kitten / Wachstum":
            st.markdown("##### 🦴 Rassespezifisches Katzen-Wachstums-Setup (Maine Coon & Riesenrassen)")
            col_kg1, col_kg2 = st.columns(2)
            with col_kg1:
                k_end_gw = st.number_input("Erwartetes Endgewicht (ausgewachsen in kg)", value=4.5 if k_ist_gw < 4.5 else k_ist_gw, step=0.5, key="k_end_gw")
                k_rasse_kat = st.selectbox("Katzenrasse-Kategorie", ["Standard-Hauskatze / EKH (3,5–5,5 kg, ausgewachsen mit 12 Mon.)", "Große Rasse / Riesenrasse (Maine Coon, Norwegische Waldkatze 6–10 kg, Wachstum bis 3–4 Jahre)"], index=1 if k_end_gw > 5.5 else 0, key="k_rasse_kat")
            with col_kg2:
                pct_kg = round((k_ist_gw / k_end_gw) * 100.0, 1) if k_end_gw > 0 else 100.0
                st.metric("Kitten-Wachstumsfortschritt", f"{pct_kg} % des Endgewichts")
                if "Maine Coon" in k_rasse_kat or k_end_gw > 5.5:
                    st.caption("⏳ **Langzeit-Wachstum:** Riesenrassen wie Maine Coon wachsen kontinuierlich über **3 bis 4 Jahre**.")
                else:
                    st.caption("⏳ **Standard-Wachstum:** Ausgewachsen mit ca. 10 bis 12 Monaten.")

            st.info("💡 **Ad-libitum & Protein-Hinweis Kitten:** Katzenwelpen benötigen während des Wachstums reichlich tierisches Eiweiß (vRp ~ 55g / 1.000 kcal [cite: 26, 27, 33]) und dürfen ad-libitum gefüttert werden!")

        bedarf_katze = berechne_bedarfsanalyse("Katze", k_name, k_ist_gw, k_soll_gw, k_alter, k_lebensphase, k_kastriert, k_ziel, k_aktivitaet, k_mahlzeiten_n, k_unvert, k_med_list, k_end_gw, bcs=k_bcs, fall_id=k_fall_id, laborwerte=k_labor_txt)
        st.session_state["aktueller_bedarf"] = bedarf_katze
        st.session_state["aktueller_bedarf_katze"] = bedarf_katze

        st.markdown("---")
        st.subheader(f"📊 Berechnete Katzen-Sollwerte für {k_name} (FEDIAF / NRC)")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Metabolisches KGW", f"{bedarf_katze['km_meta']} kg^0.67")
        m2.metric("Tages-MER", f"{bedarf_katze['mer_kcal']} kcal", f"{bedarf_katze['mer_mj']} MJ")
        m3.metric("vRp (Eiweiß)", f"{bedarf_katze['vrp_g']} g", f"Brutto: {bedarf_katze['rp_g']} g")
        m4.metric("Calcium / Phosphor", f"{round(bedarf_katze['ca_mg'],0)}mg / {round(bedarf_katze['p_mg'],0)}mg", "Ca:P 1.00:1 (Ideal)")
        m5.metric("Taurin / Vit A", f"{bedarf_katze['taurin_mg']} mg Taurin", f"Vit A: {bedarf_katze['vit_a_ie']} IE")

        st.markdown("---")
        st.subheader("📐 Rechenmodell für Energiedichte (FEDIAF 4-Stufen vs. Atwater)")
        k_f_modell_t1 = st.radio(
            "Wählen Sie das mathematische Modell zur Energiedichte-Berechnung von Katzenfutter:",
            ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"],
            index=0 if st.session_state.get("k_f_modell", "FEDIAF 4-Stufen-Modell") == "FEDIAF 4-Stufen-Modell" else 1,
            help="🔘 FEDIAF 4-Stufen: Höchste Präzision bei Diäten & Rationsberechnungen.\n🔘 Modifiziertes Atwater: Standard-Schätzverfahren mit festen Faktoren.",
            key="k_f_modell_t1"
        )
        st.session_state["k_f_modell"] = k_f_modell_t1

        st.info("💡 **Stoffwechsel-Hinweis Katze:** Katzen sind obligate Karnivoren. Sie benötigen fertiges (preformes) Vitamin A aus Leber/Fleisch sowie essenzielles Taurin und Arachidonsäure im Futter.")

    # TAB 2: KATZE FUTTERANALYSATOR
    with tab2_k:
        st.header("Fertigfutter-Analyse für Katzen (Nass- & Trockenfutter & Datenbank)")
        st.subheader("📐 Rechenmodell für Energiedichte")
        k_f_modell_t2 = st.radio(
            "Wählen Sie das mathematische Modell zur Energiedichte-Berechnung:",
            ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"],
            index=0 if st.session_state.get("k_f_modell", "FEDIAF 4-Stufen-Modell") == "FEDIAF 4-Stufen-Modell" else 1,
            help="🔘 FEDIAF 4-Stufen: Höchste Präzision, berücksichtigt die Verdaulichkeit in Abhängigkeit vom Rohfasergehalt.\n🔘 Modifiziertes Atwater: Standard-Schätzverfahren (3.5 kcal/g Protein & NfE, 8.5 kcal/g Fett).",
            key="k_f_modell_t2"
        )
        st.session_state["k_f_modell"] = k_f_modell_t2
        k_f_modell = k_f_modell_t2
        st.markdown("---")

        st.subheader("🛒 Hinterlegtes Fertigfutter aus Datenbank wählen")
        db_katze = st.session_state.get("futter_db_katze", DEFAULT_FUTTER_DB_KATZE)
        db_options_k = list(db_katze.keys()) + ["➕ Neues Futter eingeben / Eigenes Etikett"]
        selected_preset_k = st.selectbox(
            "Wählen Sie ein hinterlegtes Katzenfutter oder tragen Sie neue Werte ein:",
            db_options_k,
            index=0,
            key="k_f_preset_select"
        )

        if selected_preset_k in db_katze:
            preset_data = db_katze[selected_preset_k]
            def_name_k = selected_preset_k
            def_rp_k = float(preset_data.get("rp", 10.2))
            def_rfe_k = float(preset_data.get("rfe", 5.2))
            def_rfa_k = float(preset_data.get("rfa", 0.4))
            def_ra_k = float(preset_data.get("ra", 2.0))
            def_feuchte_k = float(preset_data.get("feuchte", 80.0))
            def_ca_k = float(preset_data.get("ca", 0.25))
            def_p_k = float(preset_data.get("p", 0.20))
            def_taurin_k = float(preset_data.get("taurin_add", 0))
            def_vit_a_k = float(preset_data.get("vit_a_add", 0))
            def_vit_d_k = float(preset_data.get("vit_d_add", 0))
            def_jod_k = float(preset_data.get("jod_add", 0))
            def_zn_k = float(preset_data.get("zn_add", 0))
            def_antiox_k = str(preset_data.get("konservierung", ""))
        else:
            def_name_k = "Eigenes Katzenfutter"
            def_rp_k = 10.0
            def_rfe_k = 5.0
            def_rfa_k = 0.5
            def_ra_k = 2.0
            def_feuchte_k = 80.0
            def_ca_k = 0.25
            def_p_k = 0.20
            def_taurin_k = 0.0
            def_vit_a_k = 0.0
            def_vit_d_k = 0.0
            def_jod_k = 0.0
            def_zn_k = 0.0
            def_antiox_k = ""

        if selected_preset_k != st.session_state.get("last_selected_preset_k") or "k_f_name" not in st.session_state:
            st.session_state["last_selected_preset_k"] = selected_preset_k
            st.session_state["k_f_name"] = def_name_k
            st.session_state["k_f_rp"] = def_rp_k
            st.session_state["k_f_rfe"] = def_rfe_k
            st.session_state["k_f_rfa"] = def_rfa_k
            st.session_state["k_f_ra"] = def_ra_k
            st.session_state["k_f_feuchte"] = def_feuchte_k
            st.session_state["k_f_ca"] = def_ca_k
            st.session_state["k_f_p"] = def_p_k
            st.session_state["k_f_taurin"] = def_taurin_k
            st.session_state["k_f_vit_a"] = def_vit_a_k
            st.session_state["k_f_vit_d"] = def_vit_d_k
            st.session_state["k_f_jod"] = def_jod_k
            st.session_state["k_f_zn"] = def_zn_k
            st.session_state["k_f_antiox"] = def_antiox_k

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            k_f_name = st.text_input("Futter-Bezeichnung", key="k_f_name")
            k_f_rp = st.number_input("Rohprotein (Rp %)", step=0.5, key="k_f_rp")
            k_f_rfe = st.number_input("Rohfett (Rfe %)", step=0.5, key="k_f_rfe")
            k_f_rfa = st.number_input("Rohfaser (Rfa %)", step=0.1, key="k_f_rfa")
        with col_f2:
            k_f_ra = st.number_input("Rohasche (Ra %)", step=0.5, key="k_f_ra")
            k_f_feuchte = st.number_input("Feuchtigkeit (%)", step=0.5, key="k_f_feuchte")
            k_f_ca = st.number_input("Calcium (Ca %)", step=0.05, key="k_f_ca")
            k_f_p = st.number_input("Phosphor (P %)", step=0.05, key="k_f_p")

        with st.expander("🧪 Deklarierte Zusatzstoffe & Konservierungsmittel (optional)", expanded=(st.session_state.get("k_f_taurin", 0) > 0 or st.session_state.get("k_f_vit_a", 0) > 0)):
            cz1, cz2, cz3 = st.columns(3)
            with cz1:
                k_f_taurin = st.number_input("Essenzielles Taurin (mg/kg)", step=100.0, key="k_f_taurin")
                k_f_vit_a = st.number_input("Vitamin A (IE/kg)", step=1000.0, key="k_f_vit_a")
            with cz2:
                k_f_vit_d = st.number_input("Vitamin D3 (IE/kg)", step=100.0, key="k_f_vit_d")
                k_f_jod = st.number_input("Jod (mg/kg)", step=0.1, key="k_f_jod")
            with cz3:
                k_f_zn = st.number_input("Zink (mg/kg)", step=10.0, key="k_f_zn")
                k_f_antiox = st.text_input("Konservierung / Antioxidanzien", placeholder="z.B. Natürliche Tocopherole", key="k_f_antiox")

        if st.button("💾 Dieses Futter in der Katzen-Datenbank speichern", key="save_f_katze"):
            st.session_state["futter_db_katze"][k_f_name] = {
                "rp": k_f_rp, "rfe": k_f_rfe, "rfa": k_f_rfa, "ra": k_f_ra,
                "feuchte": k_f_feuchte, "ca": k_f_ca, "p": k_f_p,
                "taurin_add": k_f_taurin, "vit_a_add": k_f_vit_a, "vit_d_add": k_f_vit_d,
                "jod_add": k_f_jod, "zn_add": k_f_zn, "konservierung": k_f_antiox
            }
            st.success(f"'{k_f_name}' wurde erfolgreich in deiner Katzenfutter-Datenbank gespeichert!")
            st.rerun()

        k_f_typ = "Nassfutter" if k_f_feuchte > 50.0 else "Trockenfutter"
        k_f_res = analysiere_etikett(k_f_name, k_f_rp, k_f_rfe, k_f_rfa, k_f_ra, k_f_feuchte, k_f_ca, k_f_p, "Katze", k_f_modell, k_f_typ, taurin=k_f_taurin, vit_a=k_f_vit_a, vit_d=k_f_vit_d, jod=k_f_jod, zn=k_f_zn, antiox=k_f_antiox)
        st.session_state["aktuelles_futter_katze"] = k_f_res

        st.subheader(f"📊 Auswertung Weender-Analyse: {k_f_name}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Kohlenhydrate (NfE)", f"{k_f_res['nfe_ts']} % i.TS.")
        c2.metric("Energiedichte", f"{k_f_res['kcal_100g']} kcal / 100g ({k_f_res['modell']})")
        c3.metric("Trockensubstanz", f"{k_f_res['ts_pct']} %")
        c4.metric("Ca:P-Verhältnis", f"{k_f_res['ca_p_ratio']}:1" if k_f_res['ca_p_ratio'] else "n.a.")

        if k_f_res.get("taurin", 0) > 0 or k_f_res.get("vit_a", 0) > 0 or k_f_res.get("jod", 0) > 0 or k_f_res.get("antiox", ""):
            st.markdown("##### 🧪 Deklarierte Zusatzstoffe & Konservierung")
            cz1, cz2, cz3, cz4 = st.columns(4)
            cz1.metric("Essenzielles Taurin", f"{int(k_f_res['taurin'])} mg/kg" if k_f_res['taurin'] > 0 else "Kein Zusatz")
            cz2.metric("Vitamin A / D3", f"{int(k_f_res['vit_a'])} IE / {int(k_f_res['vit_d'])} IE" if k_f_res['vit_a'] > 0 else "n.a.")
            cz3.metric("Jod / Zink", f"{k_f_res['jod']} mg/kg / {k_f_res['zn']} mg/kg" if (k_f_res['jod'] or k_f_res['zn']) else "n.a.")
            cz4.metric("Konservierung", k_f_res['antiox'] if k_f_res['antiox'] else "Keine deklariert")

        for ze in k_f_res.get("zusatz_eignung", []):
            st.success(ze)
        for w in k_f_res["warnungen"]:
            st.warning(w)

        b = st.session_state.get("aktueller_bedarf_katze", bedarf_katze)
        futtermenge_g = round((b['mer_kcal'] / k_f_res['kcal_100g']) * 100.0, 1) if k_f_res['kcal_100g'] > 0 else 0
        portion_g = round(futtermenge_g / b['mahlzeiten_n'], 1)
        st.info(f"💡 Für **{b['name']}** ({b['mer_kcal']} kcal/Tag) beträgt die Tagesmenge dieses Katzenfutters: **{futtermenge_g} g pro Tag** (ca. **{portion_g} g pro Mahlzeit** bei {b['mahlzeiten_n']} Portionen).")

    # TAB 3: KATZE BARF-BAUKASTEN
    with tab3_k:
        st.header("🥩 Selbstgekocht- & BARF-Baukasten für Katzen (Menü A-C)")
        col_m3_k1, col_m3_k2 = st.columns([3, 1])
        with col_m3_k1:
            st.caption("📐 **Aktives Energiedichte-Modell:** " + st.session_state.get("k_f_modell", "FEDIAF 4-Stufen-Modell"))
        with col_m3_k2:
            k_f_modell_t3 = st.selectbox("Rechenmodell anpassen", ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"], index=0 if st.session_state.get("k_f_modell", "FEDIAF 4-Stufen-Modell") == "FEDIAF 4-Stufen-Modell" else 1, key="k_f_modell_t3")
            st.session_state["k_f_modell"] = k_f_modell_t3
        b = st.session_state.get("aktueller_bedarf_katze", bedarf_katze)
        
        col_t3_k1, col_t3_k2 = st.columns(2)
        with col_t3_k1:
            k_tab3_modus = st.selectbox(
                "🎯 Rations-Umfang wählen",
                [
                    "Frisch-Mahlzeit für Mischfütterung (33% ME - 1 Teil-Mahlzeit pro Tag)",
                    "Vollständige Tages-Frischration (100% ME - 100% BARF / Garspeise)"
                ],
                key="k_tab3_mode_select"
            )
        k_ratio_share = 0.3333333 if "33%" in k_tab3_modus else 1.0

        col_add1, col_add2, col_add3 = st.columns([3, 2, 2])
        with col_add1:
            selected_ing = st.selectbox("🛒 Zutat aus Katalog wählen", list(INGREDIENTS_DB.keys()), key="k_tab3_ing_select")
        with col_add2:
            ing_amount = st.number_input("Tagesmenge (g oder Stk)", value=50.0, step=5.0, key="k_tab3_ing_amount")
        with col_add3:
            st.write("")
            st.write("")
            if st.button("➕ Zutat hinzufügen", key="k_tab3_add_btn"):
                st.session_state["ration_list_katze"].append({"name": selected_ing, "menge": ing_amount})
                st.success(f"{selected_ing} ({ing_amount}g) hinzugefügt!")

        r_list = st.session_state.get("ration_list_katze", [])
        p_ing_tab3 = next((item["name"] for item in r_list if INGREDIENTS_DB[item["name"]]["kat"] in ["Fleisch", "Fisch"]), "Geflügelbrust (Huhn/Pute)")
        s_ing_tab3 = next((item["name"] for item in r_list if INGREDIENTS_DB[item["name"]]["kat"] == "Stärke"), "Reis (weich gekocht)")
        v_ing_tab3 = next((item["name"] for item in r_list if INGREDIENTS_DB[item["name"]]["kat"] in ["Basis-Gemüse", "Blähgemüse"]), "Zucchini (gedämpft/püriert)")

        if r_list:
            st.subheader("📋 Eingegebene Roh-Werte des Kunden")
            col_list1, col_list2 = st.columns([3, 1])
            with col_list1:
                st.table([{"Zutat": item["name"], "Eingegebene Dosis": format_menge_str(item["name"], item["menge"], INGREDIENTS_DB[item["name"]]["kat"]), "Kategorie": INGREDIENTS_DB[item["name"]]["kat"]} for item in r_list])
            with col_list2:
                if st.button("🗑️ Neue Ration beginnen", key="k_tab3_reset_btn"):
                    st.session_state["ration_list_katze"] = []
                    st.rerun()

        m_opt_tab3 = berechne_precision_menu(p_ing_tab3, s_ing_tab3, v_ing_tab3, b, ratio_share=k_ratio_share)
        st.session_state["last_tab3_opt_katze"] = m_opt_tab3

        st.markdown("---")
        st.subheader(f"✨ Feline Precision-Rezeptur ({'Teil-Mahlzeit 33% MER' if k_ratio_share < 0.5 else 'Vollration 100% MER'})")
        if m_opt_tab3["praxis_hinweis"]:
            st.info(m_opt_tab3["praxis_hinweis"])
        st.table(m_opt_tab3["ingredients"])

        st.subheader("📊 Nährstoff-Bilanz der OPTIMIERTEN Katzen-Rezeptur")
        c_o1, c_o2, c_o3, c_o4, c_o5 = st.columns(5)
        c_o1.metric("Energie (kcal)", f"{m_opt_tab3['totals']['kcal']} kcal", f"Soll: {round(b['mer_kcal']*k_ratio_share,1)} kcal")
        c_o2.metric("Rohprotein", f"{m_opt_tab3['totals']['rp']} g", f"Soll: {round(b['rp_g']*k_ratio_share,1)} g")
        c_o3.metric("Calcium / Phosphor", f"{m_opt_tab3['totals']['ca']}mg / {m_opt_tab3['totals']['p']}mg", f"Soll: {round(b['ca_mg']*k_ratio_share,0)} / {round(b['p_mg']*k_ratio_share,0)}")
        c_o4.metric("Ca:P Ratio", f"{m_opt_tab3['totals']['ca_p_ratio']} : 1", "🟢 Ideal 1.00 : 1")
        c_o5.metric("Taurin / Vit. A", f"{m_opt_tab3['totals'].get('taurin',0)} mg", f"Vit A: {m_opt_tab3['totals']['vit_a']} IE")

    # TAB 4: KATZE WOCHENPLAN (MODULARER 3-SCHRITT BAUSTEIN-ASSISTENT v74)
    with tab4_k:
        st.header("📅 Geführter Wochenplan & Baustein-Assistent für Katzen (v74 Precision Engine)")
        st.write("Erstelle maßgeschneiderte Fütterungspläne aus frei wählbaren Katzen-Fertigfuttern (Nass / Trocken), automatischer Ad-libitum-Fütterung und BARF-Menüs!")

        b = st.session_state.get("aktueller_bedarf_katze", bedarf_katze)

        # -------------------------------------------------------------------------
        # SCHRITT 1: BAUSTEIN-ANZAHL, AD-LIBITUM & RECHENMODELL FESTLEGEN
        # -------------------------------------------------------------------------
        st.subheader("⚙️ Schritt 1: Baustein-Anzahl, Ad-libitum & Rechenmodell festlegen")
        col_rm4_k1, col_rm4_k2 = st.columns(2)
        with col_rm4_k1:
            k_f_modell_t4 = st.radio(
                "📐 Rechenmodell für Katzenfutter-Energiedichte:",
                ["FEDIAF 4-Stufen-Modell", "Modifiziertes Atwater-Modell"],
                index=0 if st.session_state.get("k_f_modell", "FEDIAF 4-Stufen-Modell") == "FEDIAF 4-Stufen-Modell" else 1,
                key="k_f_modell_t4"
            )
            st.session_state["k_f_modell"] = k_f_modell_t4
            k_model_active = k_f_modell_t4
        with col_rm4_k2:
            st.info("💡 **Hinweis:** Das gewählte Modell wird direkt für die Präzisions-Berechnung der Katzenfutter-Mengen im Wochenplan verwendet.")

        st.markdown("#### 🐱 Ad-libitum Kombi-Fütterung (Nassfutter-Hauptmahlzeiten + TroFu-Depot)")
        k_use_adlib = st.checkbox(
            "🥖 Ad-libitum Kombi-Fütterung aktivieren (Automatisches TroFu-Depot im Puzzle-Feeder auf allen Wochentagen)",
            value=True,
            help="Aktiviert die synchrone Berechnung von Nass- und Trockenfutter. Das Trockenfutter-Depot wird AUTOMATISCH auf jeden Wochentag im Plan angewendet!",
            key="k_use_adlib_check"
        )
        
        if k_use_adlib:
            col_kombi1, col_kombi2 = st.columns(2)
            with col_kombi1:
                k_nafu_pct = st.slider("Anteil Nassfutter / Hauptmahlzeiten (% der Tages-Kalorien)", min_value=10, max_value=90, value=60, step=5, key="k_nafu_pct")
                k_trofu_pct = 100 - k_nafu_pct
                st.caption(f"📊 **Rest-Anteil Trockenfutter-Depot:** **{k_trofu_pct}% der Tages-Kalorien**")
            with col_kombi2:
                st.info(f"💡 **Ad-libitum & Häppchen-Logik:** Nass- ({k_nafu_pct}%) und Trockenfutter ({k_trofu_pct}%) ergeben zusammen exakt **100% MER ({b['mer_kcal']} kcal)**. Das TroFu-Depot wird AUTOMATISCH jeden Morgen zum Tagesplan hinzugefügt!")
        else:
            k_nafu_pct = 0
            k_trofu_pct = 0
            st.caption("ℹ️ Kombi-Fütterung deaktiviert. Wochenplan arbeitet mit getrennten 100%-Mahlzeiten oder Eigenrationen.")

        col_num_k1, col_num_k2 = st.columns(2)
        with col_num_k1:
            k_n_fertig = st.slider("1️⃣ Anzahl Fertigfutter-Varianten (FF A-C)", min_value=1, max_value=3, value=2, key="k_n_fertig")
        with col_num_k2:
            k_n_gekocht = st.slider("2️⃣ Anzahl Selbstgekocht / BARF-Varianten (Menü A-C)", min_value=0, max_value=3, value=1, key="k_n_gekocht")

        k_ratio_share = 1.0
        if k_n_gekocht > 0:
            k_gekocht_modus = st.selectbox(
                "🎯 Rations-Modus für Selbstgekocht / BARF",
                [
                    "Mischfütterung (33% ME - 1 frische Teil-Mahlzeit z.B. mittags)",
                    "Vollständige Tages-Frischration (100% ME - 100% Garspeise/BARF)"
                ],
                key="k_gekocht_modus"
            )
            k_ratio_share = 0.3333333 if "33%" in k_gekocht_modus else 1.0

        col_supp1, col_supp2, col_supp3 = st.columns(3)
        with col_supp1:
            k_supp_m = st.text_input("🌅 Morgens (Zusätze/Med)", value="", placeholder="z.B. 1 Prise Taurin", key="k_supp_m")
        with col_supp2:
            k_supp_mi = st.text_input("☀️ Mittags (Zusätze/Med)", value="", key="k_supp_mi")
        with col_supp3:
            k_supp_a = st.text_input("🌙 Abends (Zusätze/Med)", value="", key="k_supp_a")

        k_extra_supps = st.text_input("⚡ Schnelleingabe-Parser ('Morgens: ... / Abends: ...')", value="", key="k_extra_supps")
        k_fm_supp, k_fmi_supp, k_fa_supp = parse_meal_supplements(k_extra_supps, k_supp_m, k_supp_mi, k_supp_a)

        st.markdown("---")

        # -------------------------------------------------------------------------
        # SCHRITT 2: BAUSTEINE KONFIGURIEREN & LIVE-BERECHNUNG
        # -------------------------------------------------------------------------
        st.subheader("🥣 Schritt 2: Bausteine konfigurieren & berechnen")

        db_katze = st.session_state["futter_db_katze"]

        # 2A. AD-LIBITUM TROCKENFUTTER-DEPOT (falls aktiviert)
        k_adlib_depot_info = {}
        if k_use_adlib and k_trofu_pct > 0:
            st.markdown("#### 🥖 Ad-libitum Trockenfutter-Depot (Tägliche Häppchen-Fütterung)")
            trofu_keys = [k for k, v in db_katze.items() if v.get("typ") == "Trockenfutter" or v.get("feuchte", 0) <= 50]
            if not trofu_keys: trofu_keys = list(db_katze.keys())

            col_ad1, col_ad2 = st.columns(2)
            with col_ad1:
                k_trofu_sel = st.selectbox("🥖 Trockenfutter-Favorit für Ad-libitum Depot wählen", trofu_keys, index=0, key="k_trofu_sel")
                k_trofu_data = db_katze[k_trofu_sel]
                k_trofu_eval = analysiere_etikett(
                    k_trofu_sel, k_trofu_data.get("rp", 32.0), k_trofu_data.get("rfe", 20.0), k_trofu_data.get("rfa", 1.9),
                    k_trofu_data.get("ra", 7.5), k_trofu_data.get("feuchte", 10.0), k_trofu_data.get("ca", 1.0), k_trofu_data.get("p", 0.9),
                    "Katze", k_model_active
                )
                k_trofu_kcal_target = b["mer_kcal"] * (k_trofu_pct / 100.0)
                k_trofu_daily_g = round((k_trofu_kcal_target / k_trofu_eval["kcal_100g"]) * 100.0, 1) if k_trofu_eval["kcal_100g"] > 0 else 0

                st.caption(f"📊 **TroFu-Depot ({k_trofu_pct}% MER = {round(k_trofu_kcal_target,1)} kcal):** **{k_trofu_daily_g} g / Tag** ({k_trofu_sel})")
            with col_ad2:
                st.info("💡 **Anwendung:** Dieses Trockenfutter-Depot wird AUTOMATISCH jeden Tag im Wochenplan mit aufgeführt. Der Halter wiegt morgens die Grammzahl ab und füllt sie in einen **Puzzle-Feeder oder Snack-Ball**.")

            k_adlib_depot_info = {
                "name": k_trofu_sel,
                "daily_g": k_trofu_daily_g,
                "pct": k_trofu_pct
            }
            st.markdown("---")

        # 2B. FERTIGFUTTER BAUSTEINE (FF A-C)
        k_fertig_results = {}
        if k_n_fertig > 0:
            st.markdown("#### 罐 Katzen-Fertigfutter-Bausteine (Hauptmahlzeiten)")
            db_keys_k = list(db_katze.keys())
            cols_fk = st.columns(k_n_fertig)
            for i in range(k_n_fertig):
                f_lbl = f"FF {chr(65+i)}"
                with cols_fk[i]:
                    st.markdown(f"**{f_lbl}**")
                    sel_f = st.selectbox(f"Futter aus DB ({f_lbl})", db_keys_k, index=i % len(db_keys_k), key=f"k_f_sel_{i}")
                    f_data = db_katze[sel_f]
                    f_eval = analysiere_etikett(
                        sel_f, f_data.get("rp", 10.2), f_data.get("rfe", 5.2), f_data.get("rfa", 0.4),
                        f_data.get("ra", 2.0), f_data.get("feuchte", 80.0), f_data.get("ca", 0.25), f_data.get("p", 0.20),
                        "Katze", k_model_active
                    )

                    f_typ = f_data.get("typ", "Nassfutter" if f_data.get("feuchte", 80) > 50 else "Trockenfutter")
                    kcal_100g = f_eval["kcal_100g"]
                    
                    # Target MER % for FF
                    if k_use_adlib and f_typ == "Nassfutter":
                        target_kcal = b["mer_kcal"] * (k_nafu_pct / 100.0)
                        target_pct = k_nafu_pct
                    else:
                        target_kcal = b["mer_kcal"]
                        target_pct = 100

                    daily_g = round((target_kcal / kcal_100g) * 100.0, 1) if kcal_100g > 0 else 0
                    port_g = round(daily_g / b["mahlzeiten_n"], 1)

                    st.caption(f"📦 **{f_typ}** | {kcal_100g} kcal/100g\n📊 **Tagesmenge ({target_pct}% ME):** {daily_g} g/Tag\n🥣 **Pro Mahlzeit ({b['mahlzeiten_n']}x):** {port_g} g")

                    k_fertig_results[f_lbl] = {
                        "name": sel_f,
                        "typ": f_typ,
                        "kcal_100g": kcal_100g,
                        "daily_g": daily_g,
                        "portion_g": port_g
                    }

        # 2C. GEKOCHT / BARF BAUSTEINE (Menü A-C)
        k_menu_results = {}
        if k_n_gekocht > 0:
            st.markdown(f"#### 🥩 Selbstgekocht / BARF Bausteine (Menü A-C) ({'33% ME Teil-Mahlzeit' if k_ratio_share < 0.5 else '100% ME Vollration'})")
            p_opts = [ing for ing, d in INGREDIENTS_DB.items() if d.get("kat") in ["Fleisch", "Fisch"] and d.get("quelle") not in b.get("unvertraeglichkeiten", [])]
            s_opts = [ing for ing, d in INGREDIENTS_DB.items() if d.get("kat") == "Stärke" and d.get("quelle") not in b.get("unvertraeglichkeiten", [])]
            v_opts = [ing for ing, d in INGREDIENTS_DB.items() if d.get("kat") in ["Basis-Gemüse", "Blähgemüse"]]

            cols_mk = st.columns(k_n_gekocht)
            for i in range(k_n_gekocht):
                m_label = f"Menü {chr(65+i)}"
                with cols_mk[i]:
                    st.markdown(f"**🍲 {m_label}**")
                    p_sel = st.selectbox(f"Protein ({m_label})", p_opts, index=i % len(p_opts), key=f"k_p_sel_{i}")
                    s_sel = st.selectbox(f"Stärke ({m_label})", s_opts, index=i % len(s_opts), key=f"k_s_sel_{i}")
                    v_sel = st.selectbox(f"Gemüse ({m_label})", v_opts, index=i % len(v_opts), key=f"k_v_sel_{i}")

                    m_res = berechne_precision_menu(p_sel, s_sel, v_sel, b, ratio_share=k_ratio_share)
                    k_menu_results[m_label] = m_res

                    st.table(m_res["ingredients"])
                    st.caption(f"📊 {m_res['totals']['kcal']} kcal | Protein: {m_res['totals']['rp']}g | Taurin: {m_res['totals'].get('taurin',0)}mg | Ca:P {m_res['totals']['ca_p_ratio']}:1")

        st.markdown("---")

        # -------------------------------------------------------------------------
        # SCHRITT 3: 7-TAGE WOCHENPLAN-KALENDER (KATZE)
        # -------------------------------------------------------------------------
        st.subheader("📅 Schritt 3: 7-Tage Katzen-Wochenplan (Montag–Sonntag / Häppchenfütterung)")
        st.caption("Wähle für jeden Wochentag das Hauptfutter (FF A-C / Menü A-C). Das Ad-libitum Trockenfutter-Depot wird AUTOMATISCH auf alle Tage angewendet!")

        k_plan_options = []
        for f_lbl, f_info in k_fertig_results.items():
            f_short = f_info["name"].split()[0]
            k_plan_options.append(f"{f_lbl} (Fertigfutter): {f_short} ({int(f_info['daily_g'])}g/Tag)")

        if k_n_gekocht > 0 and k_ratio_share < 0.5:
            for m_lbl, m_data in k_menu_results.items():
                tot_g = round(sum([float(x['Menge'].split()[0]) for x in m_data["ingredients"] if 'g' in x['Menge']]), 0)
                if k_n_fertig > 0:
                    for f_lbl, f_info in k_fertig_results.items():
                        misch_f_g = int(f_info['daily_g'] * 0.6666667)
                        k_plan_options.append(f"Misch: {f_lbl} ({misch_f_g}g Fertigfutter) + {m_lbl} ({int(tot_g)}g Selbstgekocht/BARF)")
                else:
                    k_plan_options.append(f"Misch: Fertigfutter + {m_lbl} ({int(tot_g)}g Selbstgekocht/BARF)")

        if k_n_gekocht > 0 and k_ratio_share >= 0.5:
            for m_lbl, m_data in k_menu_results.items():
                k_plan_options.append(f"Vollration: {m_lbl} (Selbstgekocht / BARF)")

        if not k_plan_options:
            k_plan_options = ["Standard-Katzenfütterung"]

        tage_liste = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
        k_wochenplan_rows = []

        cols_days = st.columns(7)
        for idx, tag in enumerate(tage_liste):
            with cols_days[idx]:
                st.markdown(f"**{tag}**")
                day_choice = st.selectbox(f"{tag}", k_plan_options, index=idx % len(k_plan_options), key=f"k_day_sel_{idx}", label_visibility="collapsed")

                # Build day meal output
                if "Misch:" in day_choice:
                    parts = day_choice.split("+")
                    f_part = parts[0].replace("Misch:", "").strip()
                    m_part = parts[1].strip() if len(parts) > 1 else "Selbstgekocht / BARF"

                    main_morgens = f"🌅 {f_part}"
                    main_mittags = f"☀️ 🍲 {m_part}"
                    main_abends = f"🌙 {f_part}"

                elif "Voll-BARF" in day_choice or "Menü" in day_choice:
                    m_lbl = "Menü A" if "Menü A" in day_choice else ("Menü B" if "Menü B" in day_choice else "Menü C")
                    main_morgens = f"🌅 🍲 {m_lbl} (Portion 1)"
                    main_mittags = "Ruhephase" if b["mahlzeiten_n"] == 2 else f"☀️ 🍲 {m_lbl} (Portion 2)"
                    main_abends = f"🌙 🍲 {m_lbl} (Portion 3)"

                else:
                    f_lbl = "FF A" if "FF A" in day_choice else ("FF B" if "FF B" in day_choice else "FF C")
                    f_info = k_fertig_results.get(f_lbl, list(k_fertig_results.values())[0] if k_fertig_results else {"name": "Katzenfutter", "portion_g": 35.0})
                    p_g = int(f_info["portion_g"])

                    main_morgens = f"🌅 {f_info['name']} ({p_g}g)"
                    main_mittags = f"☀️ {f_info['name']} ({p_g}g)" if b["mahlzeiten_n"] >= 3 else "Ruhephase"
                    main_abends = f"🌙 {f_info['name']} ({p_g}g)"

                # AUTOMATIC AD-LIBITUM DEPOT ATTACHMENT TO MORNING
                morgens = main_morgens
                if k_use_adlib and k_adlib_depot_info:
                    morgens += f" + 🥖 TroFu-Depot ({int(k_adlib_depot_info['daily_g'])}g im Feeder)"
                if k_fm_supp:
                    morgens += f" + 💊 {k_fm_supp}"

                mittags = main_mittags
                if k_fmi_supp and mittags != "Ruhephase":
                    mittags += f" + 💊 {k_fmi_supp}"

                abends = main_abends
                if k_fa_supp:
                    abends += f" + 💊 {k_fa_supp}"

                k_wochenplan_rows.append({"Tag": tag, "Morgens": morgens, "Mittags": mittags, "Abends": abends})

        st.session_state["aktueller_wochenplan_katze"] = k_wochenplan_rows
        st.session_state["menu_results_pdf_katze"] = k_menu_results

        st.markdown("---")
        st.subheader("📋 Generierte 7-Tage Fütterungs-Übersicht für das Katzen-PDF")
        st.table(k_wochenplan_rows)

    with tab5_k:
        st.header("📑 Modulare PDF-Exports & Praxis-Branding (Katze)")
        st.write("Passen Sie Ihre Praxis-Kopf- und Fußzeilen an und laden Sie maßgeschneiderte PDF-Gutachten für Patienten und Berater herunter.")

        st.subheader("⚙️ Praxis-Branding & PDF-Header Konfiguration")
        col_bk1, col_bk2 = st.columns(2)
        with col_bk1:
            st.session_state["pdf_header_text"] = st.text_input(
                "📋 Praxis-Kopfzeile (PDF Header)",
                value=st.session_state.get("pdf_header_text", "Praxis für Tierernährungsberatung | Individueller Fütterungsplan"),
                key="k_pdf_header_input"
            )
        with col_bk2:
            st.session_state["pdf_footer_text"] = st.text_input(
                "👣 Praxis-Fußzeile (PDF Footer)",
                value=st.session_state.get("pdf_footer_text", "Praxis-Handout Tierernährungsberatung | Evidenzbasierte FEDIAF/NRC Rationsberechnung"),
                key="k_pdf_footer_input"
            )

        st.markdown("---")
        st.subheader("📥 5 Modulare PDF-Gutachten zum Download (Katze)")
        b = st.session_state.get("aktueller_bedarf_katze", bedarf_katze)
        f_res = st.session_state.get("aktuelles_futter_katze", {"name": "Catz Finefood Nassfutter", "kcal_100g": 102.0})
        w_rows = st.session_state.get("aktueller_wochenplan_katze", [])
        m_res_pdf = st.session_state.get("menu_results_pdf_katze", {})

        m_res_tmp = st.session_state.get("last_tab3_opt_katze", berechne_precision_menu("Geflügelbrust (Huhn/Pute)", "Reis (weich gekocht)", "Zucchini (gedämpft/püriert)", b))
        
        pdf1_bytes = build_bedarfsanalyse_pdf_bytes(b)
        docx1_bytes = build_bedarfsanalyse_docx_bytes(b)
        
        pdf2_bytes = build_fertigfutter_pdf_bytes(b, f_res)
        docx2_bytes = build_fertigfutter_docx_bytes(b, f_res)
        
        pdf3_bytes = build_eigenration_pdf_bytes(b, st.session_state.get("ration_list_katze", []), m_res_tmp["ingredients"], m_res_tmp["totals"])
        docx3_bytes = build_eigenration_docx_bytes(b, st.session_state.get("ration_list_katze", []), m_res_tmp["ingredients"], m_res_tmp["totals"])
        
        if w_rows:
            pdf4_bytes = build_wochenplan_pdf_bytes(b, w_rows, m_res_pdf)
            docx4_bytes = build_wochenplan_docx_bytes(b, w_rows, m_res_pdf)
        else:
            pdf4_bytes = docx4_bytes = None
            
        pdf5_bytes = build_berater_dossier_pdf_bytes(b, b, f_res, m_res_pdf, w_rows)
        docx5_bytes = build_berater_dossier_docx_bytes(b, b, f_res, m_res_pdf, w_rows)

        st.markdown("##### 📄 Export-Option 1: Katzen-Bedarfsanalyse")
        c_ex1_a, c_ex1_b = st.columns(2)
        with c_ex1_a: st.download_button("⬇️ 1. Katzen-Bedarfsanalyse PDF", data=pdf1_bytes, file_name=f"katze-bedarfsanalyse-{b['name'].lower()}.pdf", mime="application/pdf", key="k_pdf1")
        with c_ex1_b: st.download_button("📝 1. Katzen-Bedarfsanalyse Word (.docx)", data=docx1_bytes, file_name=f"katze-bedarfsanalyse-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="k_docx1")

        st.markdown("##### 📄 Export-Option 2: Katzen-Fertigfutter-Analyse")
        c_ex2_a, c_ex2_b = st.columns(2)
        with c_ex2_a: st.download_button("⬇️ 2. Katzen-Fertigfutter-Analyse PDF", data=pdf2_bytes, file_name=f"katze-fertigfutter-{b['name'].lower()}.pdf", mime="application/pdf", key="k_pdf2")
        with c_ex2_b: st.download_button("📝 2. Katzen-Fertigfutter-Analyse Word (.docx)", data=docx2_bytes, file_name=f"katze-fertigfutter-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="k_docx2")

        st.markdown("##### 📄 Export-Option 3: Katzen-Eigenrations-Analyse")
        c_ex3_a, c_ex3_b = st.columns(2)
        with c_ex3_a: st.download_button("⬇️ 3. Katzen-Eigenrations-Analyse PDF", data=pdf3_bytes, file_name=f"katze-eigenration-{b['name'].lower()}.pdf", mime="application/pdf", key="k_pdf3")
        with c_ex3_b: st.download_button("📝 3. Katzen-Eigenrations-Analyse Word (.docx)", data=docx3_bytes, file_name=f"katze-eigenration-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="k_docx3")

        st.markdown("##### 📄 Export-Option 4: Katzen-Fütterungsplan Handout")
        c_ex4_a, c_ex4_b = st.columns(2)
        if w_rows:
            with c_ex4_a: st.download_button("⬇️ 4. Katzen-Fütterungsplan Handout PDF", data=pdf4_bytes, file_name=f"katze-futterplan-{b['name'].lower()}-handout.pdf", mime="application/pdf", key="k_pdf4")
            with c_ex4_b: st.download_button("📝 4. Katzen-Fütterungsplan Handout Word (.docx)", data=docx4_bytes, file_name=f"katze-futterplan-{b['name'].lower()}-handout.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="k_docx4")
        else:
            st.info("💡 Generieren Sie zuerst in Tab 4 einen Wochenplan.")

        st.markdown("##### 📄 Export-Option 5: Clinical Vet-Dossier Katze")
        c_ex5_a, c_ex5_b = st.columns(2)
        with c_ex5_a: st.download_button("⬇️ 5. Clinical Vet-Dossier Katze PDF", data=pdf5_bytes, file_name=f"katze-berater-dossier-{b['name'].lower()}.pdf", mime="application/pdf", key="k_pdf5")
        with c_ex5_b: st.download_button("📝 5. Clinical Vet-Dossier Katze Word (.docx)", data=docx5_bytes, file_name=f"katze-berater-dossier-{b['name'].lower()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="k_docx5")

    # TAB 7: KATZE NEUES FUTTER ERFASSEN & DB-VERWALTUNG
    with tab7_k:
        st.header("➕ Neues Katzenfutter erfassen & DB-Verwaltung")
        st.write("Erfassen Sie neue Katzenfuttermittel für Ihre aktuelle Beratungssitzung (Weg A) oder generieren Sie einen vorgefertigten KI-Prompt zur dauerhaften Speicherung im App-Code (Weg B).")

        st.subheader("1️⃣ Produktdaten & Weender-Deklaration eingeben")
        col_addk1, col_addk2 = st.columns(2)
        with col_addk1:
            st.markdown("##### 📦 Weender-Analytische Bestandteile")
            add_k_name = st.text_input("Produktname & Marke (Katze)", value="", placeholder="z.B. GranataPet Symphonie Entenbrust", key="add_k_name")
            add_k_typ = st.selectbox("Futtermittel-Typ (Katze)", ["Nassfutter", "Trockenfutter"], index=0, key="add_k_typ")
            add_k_feuchte = st.number_input("Feuchtigkeit (%)", value=80.0 if add_k_typ == "Nassfutter" else 6.0, step=0.5, key="add_k_feuchte")
            add_k_rp = st.number_input("Rohprotein Rp (%)", value=10.5, step=0.5, key="add_k_rp")
            add_k_rfe = st.number_input("Rohfett Rfe (%)", value=5.5, step=0.5, key="add_k_rfe")
            add_k_rfa = st.number_input("Rohfaser Rfa (%)", value=0.4, step=0.1, key="add_k_rfa")
            add_k_ra = st.number_input("Rohasche Ra (%)", value=2.2, step=0.1, key="add_k_ra")
            add_k_ca = st.number_input("Calcium Ca (%)", value=0.25, step=0.05, key="add_k_ca")
            add_k_p = st.number_input("Phosphor P (%)", value=0.20, step=0.05, key="add_k_p")

        with col_addk2:
            st.markdown("##### 🧪 Deklarierte Feline Zusatzstoffe (besonders wichtig für Katzen)")
            add_k_taurin = st.number_input("Essenzielles Taurin (mg / kg)", value=1500.0, step=100.0, key="add_k_taurin")
            add_k_vita = st.number_input("Vitamin A (IE / kg)", value=20000.0, step=500.0, key="add_k_vita")
            add_k_vitd = st.number_input("Vitamin D3 (IE / kg)", value=1500.0, step=100.0, key="add_k_vitd")
            add_k_jod = st.number_input("Jod (mg / kg)", value=0.75, step=0.05, key="add_k_jod")
            add_k_zn = st.number_input("Zink (mg / kg)", value=120.0, step=5.0, key="add_k_zn")
            add_k_antiox = st.selectbox("Antioxidanzien / Konservierung (Katze)", ["Natürliche Tocopherole (Vit. E)", "BHA / BHT (E 320 / E 321)", "Keine Angaben / Unkonserviert"], key="add_k_antiox")

        st.markdown("---")
        st.subheader("2️⃣ Futter speichern oder KI-Prompt generieren")
        col_bk1, col_bk2 = st.columns(2)
        with col_bk1:
            st.markdown("##### 💾 Weg A: Katzenfutter für die aktuelle Sitzung speichern")
            st.write("Das Katzenfutter steht sofort in allen Dropdowns der aktuellen Beratung zur Verfügung.")
            if st.button("💾 Futter in Sitzung speichern (Katze)", key="add_k_save_btn"):
                if not add_k_name.strip():
                    st.error("⚠️ Bitte geben Sie eine Futter-Bezeichnung ein!")
                else:
                    st.session_state["futter_db_katze"][add_k_name.strip()] = {
                        "rp": add_k_rp, "rfe": add_k_rfe, "rfa": add_k_rfa, "ra": add_k_ra,
                        "feuchte": add_k_feuchte, "ca": add_k_ca, "p": add_k_p, "typ": add_k_typ,
                        "taurin": add_k_taurin, "vit_a": add_k_vita, "vit_d": add_k_vitd, "jod": add_k_jod, "zn": add_k_zn, "antiox": add_k_antiox
                    }
                    st.success(f"✅ '{add_k_name.strip()}' wurde in deiner Katzen-Datenbank gespeichert!")

        with col_bk2:
            st.markdown("##### 🤖 Weg B: KI-Prompt für dauerhafte Code-Speicherung")
            st.write("Kopieren Sie diesen generierten Prompt und senden Sie ihn im Chat an Gemini Notebook, um das Katzenfutter fest im Quellcode zu verankern.")
            pk_name = add_k_name.strip() if add_k_name.strip() else "Neues Katzenfutter"
            prompt_k_text = f"""Lieber KI-Assistent, bitte füge folgendes Katzenfutter dauerhaft in die App-Stamm-Datenbank ein:

• Produktname: {pk_name}
• Spezies: Katze
• Typ: {add_k_typ}
• Analytische Bestandteile:
  - Rohprotein: {add_k_rp}%
  - Rohfett: {add_k_rfe}%
  - Rohfaser: {add_k_rfa}%
  - Rohasche: {add_k_ra}%
  - Feuchtigkeit: {add_k_feuchte}%
  - Calcium: {add_k_ca}%
  - Phosphor: {add_k_p}%
• Deklarierte Zusatzstoffe:
  - Essenzielles Taurin: {add_k_taurin} mg/kg
  - Vitamin A: {add_k_vita} IE/kg
  - Vitamin D3: {add_k_vitd} IE/kg
  - Jod: {add_k_jod} mg/kg
  - Zink: {add_k_zn} mg/kg
  - Konservierung: {add_k_antiox}"""
            st.text_area("📋 Vorlage-Prompt zum Kopieren für den Chat:", value=prompt_k_text, height=220, key="prompt_k_txt_area")

    # TAB 6: KATZE NACHKONTROLLE
    with tab6_k:
        st.header("📈 14- & 30-Tage Nachkontrolle & Wiege-Protokoll (Katze)")
        b = st.session_state.get("aktueller_bedarf_katze", bedarf_katze)
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            w0 = st.number_input("Startgewicht (kg)", value=b["ist_gw"], step=0.1, key="k_w0")
            w30 = st.number_input("Tag 30 Gewicht (kg)", value=round(b["ist_gw"]-0.2, 1), step=0.1, key="k_w30")
        with col_w2:
            rate_per_week = round(((w0 - w30) / 30.0) * 7.0, 2)
            target_min = round(w0 * 0.005, 2)
            target_max = round(w0 * 0.010, 2)
            st.metric("Abnahmerate Katze", f"-{rate_per_week} kg / Woche", f"Sanftes Ziel: -{target_min} bis -{target_max} kg/Woche")
            st.caption("🚨 **Wichtig bei Katzen:** Gewichtsverlust muss sanft erfolgen (max. 1% pro Woche), um eine feline hepatische Lipidose (Fettleber) zu verhindern.")

        if b.get("lebensphase") == "Kitten / Wachstum" or st.checkbox("📈 Rassespezifische Katzen-Wachstumskurve anzeigen", key="k_show_growth"):
            st.markdown("---")
            st.subheader("🐱 Rassespezifische Katzen-Wachstumskurve (3-Klassen-Modell)")
            tab1_end_gw_k = float(b.get("end_gw", b["soll_gw"]))

            if tab1_end_gw_k < 3.5:
                def_cat_idx = 0
            elif tab1_end_gw_k <= 6.0:
                def_cat_idx = 1
            else:
                def_cat_idx = 2

            cat_classes = [
                "Katze · 1 · Zierliche Rassen (2.0–3.5 kg | 9–12 Mon. | z.B. Singapura, Siam)",
                "Katze · 2 · Standard / Hauskatzen (3.5–5.5 kg / Kater 7 kg | 12 Mon. | z.B. EKH, BKH)",
                "Katze · 3 · Spätentwickler / Riesen (6.0–10.0+ kg | 36–48 Mon. | z.B. Maine Coon, Norweger)"
            ]

            col_gc1, col_gc2 = st.columns(2)
            with col_gc1:
                k_growth_class = st.selectbox(
                    "📐 Rasse- & Entwicklungsmodell Katzenwelpen:",
                    cat_classes,
                    index=def_cat_idx,
                    key="k_growth_class_sel"
                )

            if "last_k_growth_class" not in st.session_state or st.session_state["last_k_growth_class"] != k_growth_class:
                st.session_state["last_k_growth_class"] = k_growth_class
                if "1 · Zierliche" in k_growth_class:
                    new_default_target_k = tab1_end_gw_k if (2.0 <= tab1_end_gw_k <= 3.5) else 3.0
                elif "2 · Standard" in k_growth_class:
                    new_default_target_k = tab1_end_gw_k if (3.5 <= tab1_end_gw_k <= 6.0) else 4.5
                else: # Katze 3
                    new_default_target_k = tab1_end_gw_k if (6.0 <= tab1_end_gw_k <= 12.0) else 8.0
                st.session_state["k_end_gw_input"] = new_default_target_k

            with col_gc2:
                end_target = st.number_input(
                    "🎯 Erwartetes Endgewicht (ausgewachsen in kg):",
                    min_value=1.0,
                    max_value=25.0,
                    value=float(st.session_state.get("k_end_gw_input", tab1_end_gw_k)),
                    step=0.5,
                    key="k_end_gw_input"
                )

            months_k = [1, 2, 4, 6, 8, 10, 12, 18, 24, 36, 48]

            if "1 · Zierliche" in k_growth_class:
                pcts_k = [0.20, 0.35, 0.55, 0.75, 0.90, 0.98, 1.00, 1.00, 1.00, 1.00, 1.00]
                dauer_str_k = "9–12 Monate"
                verlauf_str_k = "Sehr schneller früher Hauptzuwachs; frühe körperliche Reife (z.B. Singapura, Siam)."
            elif "2 · Standard" in k_growth_class:
                pcts_k = [0.15, 0.28, 0.45, 0.62, 0.78, 0.90, 1.00, 1.00, 1.00, 1.00, 1.00]
                dauer_str_k = "12 Monate"
                verlauf_str_k = "Harmonischer gleichmäßiger Anstieg bis zum 12. Monat (ca. 450 g/Monat Richtwert, z.B. EKH, BKH)."
            else: # Spätentwickler / Riesen (Maine Coon)
                pcts_k = [0.10, 0.18, 0.35, 0.52, 0.69, 0.76, 0.80, 0.90, 0.95, 0.98, 1.00]
                dauer_str_k = "36–48 Monate"
                verlauf_str_k = "Früher Hauptzuwachs im 1. Jahr (ca. 70% mit 8 Mon. / 80% mit 12 Mon.) & langsame Ausreifung über 3-4 Jahre."

            soll_k_weights = [round(end_target * p, 2) for p in pcts_k]
            soll_upper_k = [round(w * 1.10, 2) for w in soll_k_weights]
            soll_lower_k = [round(w * 0.90, 2) for w in soll_k_weights]

            st.info(f"💡 **Modell-Klasse Katze:** {k_growth_class.split('(')[0].strip()}\n• **Erwartetes Endgewicht:** {end_target} kg\n• **Entwicklungsdauer:** {dauer_str_k}\n• **Modellverlauf:** {verlauf_str_k}")

            # Session State for Weighing History (Katze)
            patient_mon_k = round(float(b.get("alter", 0.5)) * 12.0, 1)
            patient_kg_k = float(b.get("ist_gw", 4.0))
            if "k_wiegehistorie" not in st.session_state:
                st.session_state["k_wiegehistorie"] = [
                    {"Monat": patient_mon_k, "Gewicht": patient_kg_k, "Notiz": "Erstaufnahme"}
                ]

            st.markdown("##### 📝 Fortlaufende Messpunkt-Erfassung (Wiegehistorie)")
            col_wk1, col_wk2, col_wk3, col_wk4 = st.columns([1.5, 1.5, 3, 1.5])
            with col_wk1:
                new_m_k = st.number_input("Alter (Monate)", value=patient_mon_k if patient_mon_k > 0 else 2.0, step=0.5, key="new_m_k")
            with col_wk2:
                new_kg_k = st.number_input("Gewicht (kg)", value=patient_kg_k if patient_kg_k > 0 else 2.0, step=0.05, key="new_kg_k")
            with col_wk3:
                new_note_k = st.text_input("Befund / Notiz", value="Routine-Check", key="new_note_k")
            with col_wk4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Hinzufügen", key="add_wiege_k"):
                    st.session_state["k_wiegehistorie"].append({
                        "Monat": float(new_m_k),
                        "Gewicht": float(new_kg_k),
                        "Notiz": str(new_note_k)
                    })
                    st.session_state["k_wiegehistorie"] = sorted(st.session_state["k_wiegehistorie"], key=lambda x: x["Monat"])
                    st.rerun()

            import pandas as pd
            if st.session_state["k_wiegehistorie"]:
                hist_df_k = pd.DataFrame(st.session_state["k_wiegehistorie"])
                st.dataframe(hist_df_k, use_container_width=True)
                
                col_del_k1, col_del_k2 = st.columns(2)
                with col_del_k1:
                    opts_del_k = [f"{i+1}. {pt['Monat']} Mon. - {pt['Gewicht']} kg ({pt['Notiz']})" for i, pt in enumerate(st.session_state["k_wiegehistorie"])]
                    selected_del_k = st.selectbox("Eintrag zum Löschen auswählen:", opts_del_k, key="sel_del_pt_k")
                    if st.button("❌ Ausgewählten Eintrag löschen", key="btn_del_single_k"):
                        idx_del = int(selected_del_k.split(".")[0]) - 1
                        st.session_state["k_wiegehistorie"].pop(idx_del)
                        st.rerun()
                with col_del_k2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.button("🔄 Auf Erstaufnahme zurücksetzen", key="reset_wiege_k"):
                            st.session_state["k_wiegehistorie"] = [{"Monat": patient_mon_k, "Gewicht": patient_kg_k, "Notiz": "Erstaufnahme"}]
                            st.rerun()
                    with col_b2:
                        if st.button("🧹 Historie komplett leeren", key="clear_wiege_k"):
                            st.session_state["k_wiegehistorie"] = []
                            st.rerun()

            import matplotlib.pyplot as plt

            fig_k, ax_k = plt.subplots(figsize=(9.5, 4.5))

            # Fill tolerance band
            ax_k.fill_between(months_k, soll_lower_k, soll_upper_k, color="#2ca02c", alpha=0.15, label="±10% Toleranzbereich (Gesunder Korridor)")
            ax_k.plot(months_k, soll_lower_k, linestyle="--", color="#2ca02c", alpha=0.4, linewidth=1)
            ax_k.plot(months_k, soll_upper_k, linestyle="--", color="#2ca02c", alpha=0.4, linewidth=1)
            ax_k.plot(months_k, soll_k_weights, marker="o", color="#2ca02c", linewidth=2.5, label="Soll-Gewichtskurve Katze (kg)")

            # Plot historical weighing points
            if st.session_state["k_wiegehistorie"]:
                k_m_list = [pt["Monat"] for pt in st.session_state["k_wiegehistorie"]]
                k_kg_list = [pt["Gewicht"] for pt in st.session_state["k_wiegehistorie"]]
                ax_k.plot(k_m_list, k_kg_list, marker="s", color="#d62728", linewidth=2.2, linestyle="-", label=f"Gemessener Verlauf: {b['name']}")
                for pt in st.session_state["k_wiegehistorie"]:
                    ax_k.annotate(f"{pt['Gewicht']}kg", (pt["Monat"], pt["Gewicht"]), textcoords="offset points", xytext=(0,7), ha='center', fontsize=8, fontweight='bold', color="#d62728")

            ax_k.set_title(f"Soll-Wachstumskurve Katze mit ±10% Toleranzband ({k_growth_class.split('(')[0].strip()})", fontsize=11, fontweight="bold")
            ax_k.set_xlabel("Alter in Monaten", fontsize=10)
            ax_k.set_ylabel("Körpergewicht (kg)", fontsize=10)
            ax_k.grid(True, linestyle="--", alpha=0.6)
            ax_k.set_xticks(months_k)
            ax_k.legend(loc="upper left", fontsize=8.5)
            plt.tight_layout()
            st.pyplot(fig_k)
            plt.close(fig_k)

            st.markdown("##### 📊 Soll-Gewichtstabelle nach Alter (Monate)")
            table_k_df = pd.DataFrame({
                "Alter": [f"{m} Mon." for m in months_k],
                "Soll-Gewicht (kg)": [f"{w:.2f} kg" for w in soll_k_weights],
                "Prozent vom Endgewicht": [f"{int(p*100)} %" for p in pcts_k]
            })
            st.dataframe(table_k_df, use_container_width=True)

            pdf_growth_k_bytes = build_growth_curve_pdf_bytes(
                b, k_growth_class, end_target, months_k, soll_k_weights, soll_lower_k, soll_upper_k, st.session_state.get("k_wiegehistorie", []), is_katze=True
            )
            docx_growth_k_bytes = build_growth_curve_docx_bytes(
                b, k_growth_class, end_target, months_k, soll_k_weights, soll_lower_k, soll_upper_k, st.session_state.get("k_wiegehistorie", []), is_katze=True
            )
            c_gr_k1, c_gr_k2 = st.columns(2)
            with c_gr_k1:
                st.download_button(
                    "📄 Katzen-Wachstumskurve PDF Report herunterladen",
                    data=pdf_growth_k_bytes,
                    file_name=f"Wachstumskurve_Katze_{b['name']}.pdf",
                    mime="application/pdf",
                    key="dl_pdf_growth_k"
                )
            with c_gr_k2:
                st.download_button(
                    "📝 Katzen-Wachstumskurve Word (.docx) Report herunterladen",
                    data=docx_growth_k_bytes,
                    file_name=f"Wachstumskurve_Katze_{b['name']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_docx_growth_k"
                )
