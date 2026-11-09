"""
export_engine.py - Ernährungsphysiologische Rechen-Engine & PDF/Word Exports (v71)
"""

import io
import math
import re
import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from fpdf import FPDF, XPos, YPos
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from futter_db import DEFAULT_FUTTER_DB_HUND, DEFAULT_FUTTER_DB_KATZE, INGREDIENTS_DB

def get_alter_display_str(alter):
    if alter is None:
        return "n.a."
    try:
        alt_f = float(alter)
        if alt_f < 1.0:
            mon = round(alt_f * 12.0, 1)
            if mon.is_integer():
                mon = int(mon)
            return f"{mon} Mon."
        else:
            alt_r = round(alt_f, 1)
            if alt_r.is_integer():
                alt_r = int(alt_r)
            return f"{alt_r} J."
    except Exception:
        return f"{alter} J."


# HELPER FUNCTIONS FOR UNVERTRÄGLICHKEITEN & RATION RECALCULATION & EVALUATION
# -----------------------------------------------------------------------------
def is_excluded(ingredient_name, unvert_list):
    if not unvert_list:
        return False
    ing_lower = ingredient_name.lower()
    for u in unvert_list:
        u_str = u.strip().lower()
        if u_str and u_str in ing_lower:
            return True
    return False

def recalculate_ration_totals(ingredients_list):
    tot = {'kcal': 0.0, 'rp': 0.0, 'rfe': 0.0, 'ca': 0.0, 'p': 0.0, 'jod': 0.0, 'vit_a': 0.0, 'vit_d': 0.0, 'fe': 0.0, 'zn': 0.0, 'taurin': 0.0}
    for item in ingredients_list:
        ing_name = item['Zutat']
        m_str = item.get('Menge', '0')
        
        if ing_name == 'Zink-Chelat (Spurenelement)':
            match_mg = re.search(r'([\d\.]+)\s*mg', m_str)
            if match_mg:
                tot['zn'] += float(match_mg.group(1))
            continue
            
        try:
            m_g = float(re.findall(r'[-+]?\d*\.\d+|\d+', m_str)[0])
        except Exception:
            m_g = 0.0
            
        if ing_name in INGREDIENTS_DB:
            db_entry = INGREDIENTS_DB[ing_name]
            factor = m_g / 100.0
            for k in tot:
                tot[k] += db_entry.get(k, 0.0) * factor
                
    tot['ca_p_ratio'] = round(tot['ca'] / tot['p'], 2) if tot['p'] > 0 else 0.0
    for k in ['kcal', 'ca', 'p', 'jod', 'vit_a', 'vit_d', 'taurin']:
        tot[k] = round(tot[k], 0)
    for k in ['rp', 'rfe', 'fe', 'zn']:
        tot[k] = round(tot[k], 1)
    return tot

def evaluiere_bedarfsdeckung(totals, bedarf, ratio_share=1.0):
    spezies = bedarf.get('spezies', 'Hund')
    target_kcal = bedarf['mer_kcal'] * ratio_share
    target_rp = bedarf['rp_g'] * ratio_share
    target_ca = bedarf['ca_mg'] * ratio_share
    target_p = bedarf['p_mg'] * ratio_share
    target_jod = bedarf['jod_ug'] * ratio_share
    target_vit_a = bedarf['vit_a_ie'] * ratio_share
    target_vit_d = bedarf['vit_d_ie'] * ratio_share
    target_zn = bedarf['zn_mg'] * ratio_share
    target_taurin = bedarf.get('taurin_mg', 0.0) * ratio_share

    target_ca_p_str = '1.30 : 1' if spezies == 'Hund' else '1.00 : 1'

    items = [
        ('Energie', 'kcal', target_kcal, totals.get('kcal', 0.0)),
        ('Rohprotein', 'g', target_rp, totals.get('rp', 0.0)),
        ('Calcium', 'mg', target_ca, totals.get('ca', 0.0)),
        ('Phosphor', 'mg', target_p, totals.get('p', 0.0)),
        ('Jod', 'µg', target_jod, totals.get('jod', 0.0)),
        ('Vitamin A', 'IE', target_vit_a, totals.get('vit_a', 0.0)),
        ('Vitamin D3', 'IE', target_vit_d, totals.get('vit_d', 0.0)),
        ('Zink', 'mg', target_zn, totals.get('zn', 0.0)),
    ]
    if spezies == 'Katze':
        items.append(('Taurin', 'mg', target_taurin, totals.get('taurin', 0.0)))

    table_rows = []
    deficits = []
    is_fully_covered = True

    for label, unit, soll, ist in items:
        if soll > 0:
            deck_pct = round((ist / soll) * 100.0, 1)
        else:
            deck_pct = 100.0

        if deck_pct >= 95.0:
            status = '🟢 Bedarfsdeckend'
        elif deck_pct >= 80.0:
            status = '🟡 Ausreichend'
            is_fully_covered = False
            deficits.append(f'{label} ({deck_pct}%)')
        else:
            status = '🔴 Defizit'
            is_fully_covered = False
            deficits.append(f'{label} ({deck_pct}%)')

        table_rows.append({
            'Nährstoff': label,
            'Sollwert': f'{round(soll, 1)} {unit}',
            'Istwert': f'{round(ist, 1)} {unit}',
            'Deckung (%)': f'{deck_pct} %',
            'Status': status
        })

    ca_p_ist = totals.get('ca_p_ratio', 0.0)
    ca_p_ok = 1.10 <= ca_p_ist <= 1.80 if spezies == 'Hund' else 0.85 <= ca_p_ist <= 1.30

    ca_p_status = '🟢 Ideal' if ca_p_ok else '🟡 Abweichung'
    if not ca_p_ok:
        is_fully_covered = False
        deficits.append(f'Ca:P-Verhältnis ({ca_p_ist}:1)')

    table_rows.append({
        'Nährstoff': 'Ca:P-Verhältnis',
        'Sollwert': target_ca_p_str,
        'Istwert': f'{ca_p_ist} : 1',
        'Deckung (%)': 'n.a.',
        'Status': ca_p_status
    })

    return {
        'is_bedarfsdeckend': is_fully_covered,
        'coverage_table': table_rows,
        'deficits': deficits
    }


# -----------------------------------------------------------------------------
# HELPER FOR KITCHEN UNITS (TL / EL / ml for Oils)
# -----------------------------------------------------------------------------
def format_menge_str(name, menge_g, kat):
    if kat == "Öl":
        if menge_g <= 7.0:
            tl = round(menge_g / 4.0, 1)
            return f"{menge_g} g (~{tl} TL / ~{menge_g} ml)"
        else:
            el = round(menge_g / 10.0, 1)
            return f"{menge_g} g (~{el} EL / ~{menge_g} ml)"
    return f"{menge_g} g"

def parse_meal_supplements(text_input, morgens_in="", mittags_in="", abends_in=""):
    morgens = morgens_in.strip()
    mittags = mittags_in.strip()
    abends = abends_in.strip()
    
    if text_input and text_input.strip():
        txt = text_input.strip()
        txt_lower = txt.lower()
        if "morgens:" in txt_lower or "abends:" in txt_lower or "mittags:" in txt_lower:
            parts = re.split(r'[/|\n]', txt)
            for p in parts:
                p_clean = p.strip()
                p_l = p_clean.lower()
                if "morgens:" in p_l:
                    val = re.sub(r'(?i)morgens:', '', p_clean).strip()
                    morgens = (morgens + " + " + val).strip(" +") if morgens else val
                elif "mittags:" in p_l:
                    val = re.sub(r'(?i)mittags:', '', p_clean).strip()
                    mittags = (mittags + " + " + val).strip(" +") if mittags else val
                elif "abends:" in p_l:
                    val = re.sub(r'(?i)abends:', '', p_clean).strip()
                    abends = (abends + " + " + val).strip(" +") if abends else val
        else:
            abends = (abends + " + " + txt).strip(" +") if abends else txt
            
    return morgens, mittags, abends

# -----------------------------------------------------------------------------
# FPDF HELPER KLASSE & TEXT CLEANER (CONFIGURABLE HEADER & FOOTER)
# -----------------------------------------------------------------------------
class PDFBase(FPDF):
    def __init__(self, header_text=None, footer_text=None, orientation='P', unit='mm', format='A4'):
        super().__init__(orientation=orientation, unit=unit, format=format)
        import streamlit as st
        self.header_text = header_text or st.session_state.get("pdf_header_text", "Praxis für Tierernährungsberatung | Individueller Fütterungsplan")
        self.footer_text = footer_text or st.session_state.get("pdf_footer_text", "Praxis-Handout Tierernährungsberatung | Evidenzbasierte FEDIAF/NRC Rationsberechnung")
        self.set_margins(10, 16, 10)

    def sanitize(self, text):
        if not text:
            return ""
        replacements = {
            "—": "-", "–": "-", "“": '"', "”": '"', "’": "'", "‘": "'",
            "•": "*", "🟢": "[OK]", "🟡": "[!]", "🚨": "[WARN]", "⚠️": "[!]", "⚠": "[!]",
            "🐾": "", "📊": "", "🏷️": "", "🏷": "", "🍲": "", "📑": "", "📈": "", "🦴": "",
            "🥦": "", "🥩": "", "🛒": "", "💡": "", "🎯": "", "🐕": "", "🐶": "", "🐱": "", "📦": "", "⚙️": "", "⚙": "", "✨": "",
            "⚡": "", "💾": "", "📅": "", "📋": "", "📐": "", "📥": "", "🔄": "", "🔘": "", "🗑": "", "🛡": "", "➕": "+",
            "🌅": "Morgens:", "☀️": "Mittags:", "☀": "Mittags:", "🌙": "Abends:", "💊": "Med:", "🥚": "Ei:", "🐟": "Fisch:",
            "🥖": "", "🥣": "", "⚖️": "", "⚖": "", "🧪": "", "✍️": "", "✍": "", "🔍": "", "📝": ""
        }
        for k, v in replacements.items():
            text = text.replace(k, v)
        
        clean_chars = []
        for ch in text:
            try:
                ch.encode('latin-1')
                clean_chars.append(ch)
            except UnicodeEncodeError:
                pass
        return "".join(clean_chars)

    def header(self):
        self.set_y(8)
        self.set_font('Helvetica', 'B', 8)
        self.set_text_color(100, 100, 100)
        self.cell(135, 4, self.sanitize(self.header_text), new_x=XPos.RIGHT, new_y=YPos.TOP, align='L')
        self.cell(55, 4, self.sanitize(f'Seite {self.page_no()}'), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.set_draw_color(200, 200, 200)
        self.line(10, 13, 200, 13)
        self.set_y(17)

    def footer(self):
        self.set_y(-12)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 5, self.sanitize(self.footer_text), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')

# -----------------------------------------------------------------------------
# LOGIK BAUSTEIN 1: BEDARFSRECHNER (HUND & KATZE)
# -----------------------------------------------------------------------------

def get_bcs_info(bcs, spezies="Hund"):
    bcs = int(bcs)
    if bcs == 1:
        title = "BCS 1/9 - Kachektisch / Extremer Mangel (< -30% Idealgewicht)"
        desc = "Rippen, Lendenwirbel und Beckenknochen sind von weitem deutlich sichtbar. Kein greifbares Körperfett, schwerer Verlust an aktiver Muskelmasse. Dringender Gewichtsaufbau erforderlich."
    elif bcs == 2:
        title = "BCS 2/9 - Sehr Mager (-20% bis -30% Idealgewicht)"
        desc = "Rippen leicht tastbar und sichtbar. Dornfortsätze der Wirbel sichtbar. Sehr deutliche Taille und starker Bauchaufzug. Minimales Körperfett."
    elif bcs == 3:
        title = "BCS 3/9 - Mager / Schlank (-10% bis -20% Idealgewicht)"
        desc = "Rippen leicht tastbar mit minimaler Fettabdeckung. Akzentuierte Taille von oben und ausgeprägte Flankeneinziehung von der Seite."
    elif bcs == 4:
        title = "BCS 4/9 - Leicht Schlank (-5% bis -10% Idealgewicht)"
        desc = "Rippen leicht tastbar unter dünner Fettbedeckung. Taille von oben gut erkennbar, straffer Bauchaufzug. Sehr gute sportliche Kondition."
    elif bcs == 5:
        title = "BCS 5/9 - Idealgewicht / Optimaler Normalzustand"
        desc = "Rippen ohne ausgeprägten Druck leicht tastbar unter dünner Fettschicht. Proportionierte Taille hinter den Rippen erkennbar. Idealzustand laut FEDIAF/WSAVA."
    elif bcs == 6:
        title = "BCS 6/9 - Korpulent / Leichtes Übergewicht (+10% bis +15% Idealgewicht)"
        desc = "Rippen nur mit leichtem Druck tastbar unter mäßiger Fettschicht. Leichte Fettablagerung an Lende und Schwanzansatz. Taille schwach erkennbar."
    elif bcs == 7:
        title = "BCS 7/9 - Korpulent / Übergewicht (+15% bis +30% Idealgewicht)"
        desc = "Rippen nur mit deutlichem Druck tastbar. Deutliche Fettpolster über Lendenbereich und Schwanzansatz. Bauchumfang vermehrt, Taille kaum sichtbar."
    elif bcs == 8:
        title = "BCS 8/9 - Verfettet / Adipös (+30% bis +45% Idealgewicht)"
        desc = "Rippen unter dicker Fettschicht nicht mehr tastbar. Massive Fettansammlungen an Hals, Brust und Rücken. Keine Taille, Bauchumfangsvermehrung."
    else: # 9
        title = "BCS 9/9 - Schwerst Adipös (> +45% Idealgewicht)"
        desc = "Massive Fettpolster am gesamten Körper, Brustkorb und Nacken. Deutlicher Hängebauch, völlig fehlende Taille, stark eingeschränkte Bewegungs- und Herz-Kreislauf-Kapazität."
    return title, desc

def berechne_bedarfsanalyse(spezies, name, ist_gw, soll_gw, alter, lebensphase, kastriert, ziel, aktivitaet, mahlzeiten_n=3, unvertraeglichkeiten=[], medikation=[], end_gw=None, bcs=5, fall_id=None, laborwerte=""):
    gw = soll_gw if ziel == "Gewichtsreduktion" else ist_gw
    target_end_gw = end_gw if (end_gw is not None and end_gw > 0) else gw
    km_meta = math.pow(gw, 0.75) if spezies == "Hund" else math.pow(gw, 0.67)
    
    ist_gw_pct = round((ist_gw / target_end_gw) * 100.0, 1) if target_end_gw > 0 else 100.0
    ca_max_guard = False

    if spezies == "Hund":
        if ziel == "Gewichtsreduktion":
            mer_faktor = 70.0 # auf Sollgewicht
        elif ziel == "Auffütterung / Zunahme":
            mer_faktor = 130.0
        elif lebensphase in ["Junior / Wachstum", "Welpe"]:
            if ist_gw_pct < 40.0:
                mer_faktor = 220.0 # Phase 1: 220 kcal / km_meta
            elif ist_gw_pct < 80.0:
                mer_faktor = 176.0 # Phase 2: 176 kcal / km_meta
            else:
                mer_faktor = 132.0 # Phase 3: 132 kcal / km_meta
        elif aktivitaet == "Sport/Leistung":
            mer_faktor = 160.0
        elif aktivitaet == "Aktiv":
            mer_faktor = 130.0
        elif aktivitaet == "Inaktiv" or lebensphase == "Senior" or kastriert == "Ja":
            mer_faktor = 95.0
        else:
            mer_faktor = 110.0
        mer_kcal = mer_faktor * (math.pow(soll_gw, 0.75) if ziel == "Gewichtsreduktion" else km_meta)
    else: # Katze
        if ziel == "Gewichtsreduktion":
            mer_kcal = 53.0 * math.pow(soll_gw, 0.711)
        elif ziel == "Auffütterung / Zunahme":
            mer_kcal = 120.0 * km_meta
        elif lebensphase in ["Kitten / Wachstum", "Kitten"]:
            mer_kcal = 160.0 * km_meta
        elif aktivitaet in ["Aktiv / Freigänger", "Aktiv"]:
            mer_kcal = 120.0 * km_meta
        elif aktivitaet in ["Indoor/Inaktiv", "Inaktiv"] or kastriert == "Ja" or lebensphase == "Senior":
            mer_kcal = 75.0 * km_meta
        else:
            mer_kcal = 100.0 * km_meta

    mer_mj = (mer_kcal * 4.184) / 1000.0
    
    if spezies == "Hund":
        if ziel == "Gewichtsreduktion":
            vrp_g = max(10.0 * mer_mj, 3.7 * km_meta)
        elif lebensphase in ["Junior / Wachstum", "Welpe"]:
            vrp_g = max(12.0 * mer_mj, 6.0 * km_meta)
        else:
            vrp_g = max(10.0 * mer_mj, 4.5 * km_meta)
        rp_g = vrp_g / 0.80
        
        if lebensphase in ["Junior / Wachstum", "Welpe"]:
            if ist_gw_pct < 40.0:
                ca_mg = max(320.0 * km_meta, (mer_kcal / 1000.0) * 2500.0)
            elif ist_gw_pct < 80.0:
                ca_mg = max(240.0 * km_meta, (mer_kcal / 1000.0) * 2200.0)
            else:
                ca_mg = max(170.0 * km_meta, (mer_kcal / 1000.0) * 1800.0)
            
            # Calcium Guard for Large / Giant Breeds (>15 kg adult weight)
            if target_end_gw > 15.0:
                ca_max_mg = (mer_kcal / 1000.0) * 2500.0 # Strict AAFCO/FEDIAF limit: 2.5 g Ca / 1000 kcal ME
                if ca_mg > ca_max_mg:
                    ca_mg = ca_max_mg
                    ca_max_guard = True
        else:
            ca_mg = max(130.0 * km_meta, (mer_kcal / 1000.0) * 1700.0)

        p_mg = max(100.0 * km_meta, (mer_kcal / 1000.0) * 1250.0)
        jod_ug = max(15.0 * gw, (mer_kcal / 1000.0) * 400.0)
        vit_a_ie = max(150.0 * km_meta, (mer_kcal / 1000.0) * 1500.0)
        vit_d_ie = max(10.0 * km_meta, (mer_kcal / 1000.0) * 250.0)
        fe_mg = max(1.5 * km_meta, (mer_kcal / 1000.0) * 15.0)
        zn_mg = max(2.0 * km_meta, (mer_kcal / 1000.0) * 20.0)
        taurin_mg = 0.0
        ca_p_ratio = "1.30 : 1" if lebensphase not in ["Junior / Wachstum", "Welpe"] else "1.30 - 1.50 : 1"
    else: # Katze
        if lebensphase in ["Kitten / Wachstum", "Kitten"]:
            vrp_g = (mer_kcal / 1000.0) * 65.0
            ca_mg = (mer_kcal / 1000.0) * 2000.0
        else:
            vrp_g = (mer_kcal / 1000.0) * 55.0
            ca_mg = (mer_kcal / 1000.0) * 1500.0
        rp_g = (mer_kcal / 1000.0) * 65.0
        p_mg = (mer_kcal / 1000.0) * 1500.0 # Target 1.0 : 1
        jod_ug = (mer_kcal / 1000.0) * 350.0
        vit_a_ie = (mer_kcal / 1000.0) * 2000.0
        vit_d_ie = (mer_kcal / 1000.0) * 250.0
        fe_mg = (mer_kcal / 1000.0) * 15.0
        zn_mg = (mer_kcal / 1000.0) * 20.0
        taurin_mg = round(gw * 25.0, 0) # 25 mg/kg KGW Mindestbedarf
        ca_p_ratio = "0.9 - 1.1 : 1 (Ideal 1.00 : 1)"

    return {
        "spezies": spezies, "name": name, "ist_gw": round(ist_gw, 2), "soll_gw": round(soll_gw, 2), "end_gw": target_end_gw, "ist_gw_pct": ist_gw_pct,
        "alter": alter, "lebensphase": lebensphase, "kastriert": kastriert, "ziel": ziel, "aktivitaet": aktivitaet,
        "mahlzeiten_n": mahlzeiten_n, "ca_max_guard": ca_max_guard,
        "unvertraeglichkeiten": unvertraeglichkeiten, "medikation": medikation,
        "km_meta": round(km_meta, 2), "mer_kcal": round(mer_kcal, 1), "mer_mj": round(mer_mj, 2),
        "vrp_g": round(vrp_g, 1), "rp_g": round(rp_g, 1), "ca_mg": round(ca_mg, 1),
        "p_mg": round(p_mg, 1), "jod_ug": round(jod_ug, 1), "vit_a_ie": round(vit_a_ie, 1),
        "vit_d_ie": round(vit_d_ie, 1), "fe_mg": round(fe_mg, 1), "zn_mg": round(zn_mg, 1),
        "taurin_mg": taurin_mg, "ca_p_ratio": ca_p_ratio,
        "bcs": bcs, "fall_id": fall_id if (fall_id and str(fall_id).strip()) else f"FALL-2026-{abs(hash(name))%8999+1000}",
        "laborwerte": laborwerte
    }

def analysiere_etikett(name, rp, rfe, rfa, ra, feuchte, ca, p, spezies="Hund", modell="FEDIAF 4-Stufen-Modell", typ="Trockenfutter", taurin=0, vit_a=0, vit_d=0, jod=0, zn=0, cu=0, antiox=""):
    ts = max(0.1, 100.0 - feuchte)
    nfe = max(0.0, 100.0 - (feuchte + rp + rfe + rfa + ra))
    rohfaser_pct_ts = (rfa / ts) * 100.0
    
    if "FEDIAF" in modell:
        ge_kcal = (5.7 * rp) + (9.4 * rfe) + (4.1 * (nfe + rfa))
        if spezies == "Hund":
            verdaulichkeit_pct = 91.2 - (1.43 * rohfaser_pct_ts)
            de_kcal = ge_kcal * (verdaulichkeit_pct / 100.0)
            kcal_100g = de_kcal - (1.04 * rp)
        else: # Katze
            verdaulichkeit_pct = 87.9 - (0.88 * rohfaser_pct_ts)
            de_kcal = ge_kcal * (verdaulichkeit_pct / 100.0)
            kcal_100g = de_kcal - (0.77 * rp)
        kcal_100g = max(10.0, kcal_100g)
    else:
        kcal_100g = (3.5 * rp) + (8.5 * rfe) + (3.5 * nfe)

    rp_ts = (rp / ts) * 100.0
    rfe_ts = (rfe / ts) * 100.0
    nfe_ts = (nfe / ts) * 100.0
    
    ca_p = round(ca / p, 2) if (ca and p and p > 0) else None

    warnungen = []
    if spezies == "Hund":
        if nfe_ts > 45.0:
            warnungen.append("⚠️ Hoher Kohlenhydratgehalt in Trockensubstanz (> 45% TS).")
        if rp_ts < 20.0:
            warnungen.append("⚠️ Geringer Proteingehalt (< 20% TS).")
        if ca_p:
            if ca_p < 0.9:
                warnungen.append(f"🚨 Ca:P-Verhältnis zu niedrig ({ca_p}:1)! Risiko für Knochenentkalkung.")
            elif ca_p > 2.0:
                warnungen.append(f"⚠️ Ca:P-Verhältnis sehr hoch ({ca_p}:1).")
    else: # Katze
        if nfe_ts > 15.0:
            warnungen.append("⚠️ Erhöhter Kohlenhydratgehalt (> 15% TS). Katzen sind strikte Karnivoren.")
        if rp_ts < 30.0:
            warnungen.append("⚠️ Zu geringer Proteingehalt für Katzen (< 30% TS).")
        if ca_p:
            if ca_p < 0.8 or ca_p > 1.3:
                warnungen.append(f"⚠️ Ca:P-Verhältnis ({ca_p}:1) weicht vom felinen Idealbereich (0.9-1.1:1) ab.")

    zusatz_eignung = []
    if spezies == "Katze":
        if taurin and taurin > 0:
            if taurin < 1000:
                warnungen.append(f"⚠️ Deklarierter Taurin-Zusatz ({taurin} mg/kg) liegt unter feliner Empfehlung (mind. 1.000–2.000 mg/kg TS).")
            else:
                zusatz_eignung.append(f"🟢 Taurin-Zusatz ausreichend ({taurin} mg/kg) - schützt Herzkreislauf (DCM) & Netzhaut.")
        else:
            if typ in ["Nassfutter", "Trockenfutter"]:
                warnungen.append("⚠️ Kein Taurin-Zusatz deklariert! Katzen benötigen zwingend essenzielles Taurin im Futter.")
    
    if antiox and ("BHA" in antiox.upper() or "BHT" in antiox.upper() or "ETHOXYQUIN" in antiox.upper()):
        warnungen.append(f"🟡 Enthält synthetische Konservierungsstoffe / Antioxidanzien ({antiox}).")

    return {
        "name": name, "nfe_pct": round(nfe, 1), "ts_pct": round(ts, 1),
        "kcal_100g": round(kcal_100g, 1), "rp_ts": round(rp_ts, 1),
        "rfe_ts": round(rfe_ts, 1), "nfe_ts": round(nfe_ts, 1),
        "ca_p_ratio": ca_p, "warnungen": warnungen, "zusatz_eignung": zusatz_eignung, "modell": modell, "typ": typ,
        "rp": rp, "rfe": rfe, "rfa": rfa, "ra": ra, "feuchte": feuchte, "ca": ca, "p": p,
        "taurin": taurin, "vit_a": vit_a, "vit_d": vit_d, "jod": jod, "zn": zn, "cu": cu, "antiox": antiox
    }

# -----------------------------------------------------------------------------
# DYNAMISCHE PRECISION REZEPTUR-ENGINE (HUND & KATZE)
# -----------------------------------------------------------------------------
def berechne_precision_menu(p_ing, s_ing, v_ing, bedarf, ratio_share=1.0):
    spezies = bedarf.get("spezies", "Hund")
    target_kcal = bedarf["mer_kcal"] * ratio_share
    target_rp = bedarf["rp_g"] * ratio_share
    target_ca = bedarf["ca_mg"] * ratio_share
    target_p = bedarf["p_mg"] * ratio_share
    target_jod = bedarf["jod_ug"] * ratio_share
    target_vit_a = bedarf["vit_a_ie"] * ratio_share
    target_vit_d = bedarf["vit_d_ie"] * ratio_share
    target_zn = bedarf["zn_mg"] * ratio_share
    target_taurin = bedarf.get("taurin_mg", 0.0) * ratio_share
    kgw = bedarf["soll_gw"]
    unvert_list = bedarf.get("unvertraeglichkeiten", [])
    med_list = bedarf.get("medikation", [])

    db_p = INGREDIENTS_DB[p_ing]
    db_s = INGREDIENTS_DB[s_ing]
    db_v = INGREDIENTS_DB[v_ing]

    # Calculate base quantities purely mathematically from bedarf & nutrient targets
    p_rp_per_g = db_p["rp"] / 100.0
    raw_p_g = (target_rp / p_rp_per_g) if p_rp_per_g > 0 else 100.0
    p_g = round(raw_p_g, 1)

    is_bloat = (db_v.get("kat") == "Blähgemüse")
    max_veg_g = ((kgw * 0.6) if is_bloat else (kgw * 3.2)) * ratio_share if spezies == "Hund" else min(50.0 * ratio_share, (kgw * 1.5) * ratio_share)
    raw_v_g = min((150.0 * ratio_share) if spezies == "Hund" else (40.0 * ratio_share), max_veg_g)
    v_g = round(max(5.0, raw_v_g), 1)

    kcal_p = (db_p["kcal"] * p_g) / 100.0
    kcal_v = (db_v["kcal"] * v_g) / 100.0
    rem_kcal = max(10.0, target_kcal - (kcal_p + kcal_v))
    s_kcal_per_g = db_s["kcal"] / 100.0
    raw_s_g = (rem_kcal / s_kcal_per_g) if s_kcal_per_g > 0 else 50.0
    s_g = round(raw_s_g, 1)

    ingredients_list = [
        {"Zutat": p_ing, "Menge": f"{p_g} g", "Status / Hinweis": "Protein-Balanciert"},
        {"Zutat": s_ing, "Menge": f"{s_g} g", "Status / Hinweis": "Energie-Balanciert"},
        {"Zutat": v_ing, "Menge": f"{v_g} g", "Status / Hinweis": "Gemüse-Balanciert"}
    ]

    # Calculate initial complete ration totals
    tot = recalculate_ration_totals(ingredients_list)

    # Auto-Suggesters: Check Unverträglichkeiten and RECALCULATE complete ration after each addition
    target_ca_ratio = 1.30 if spezies == "Hund" else 1.00
    target_ca_for_p = max(target_ca, tot["p"] * target_ca_ratio)
    ca_def = target_ca_for_p - tot["ca"]

    if ca_def > 10.0:
        ca_supp = None
        ca_pct = 38.0
        if not is_excluded("Eierschalenpulver", unvert_list) and not is_excluded("Ei", unvert_list):
            ca_supp = "Eierschalenpulver (38% Ca)"
            ca_pct = 38.0
        elif not is_excluded("Algenkalk", unvert_list) and not is_excluded("Algen", unvert_list):
            ca_supp = "Algenkalk (30% Ca)"
            ca_pct = 30.0

        if ca_supp:
            add_ca_g = round(ca_def / (ca_pct * 10.0), 1)
            if add_ca_g >= 0.1:
                tl = round(add_ca_g / 4.0, 1)
                ingredients_list.append({"Zutat": ca_supp, "Menge": f"{add_ca_g} g (~{tl} TL)", "Status / Hinweis": f"✨ Auto-Suggester: Ca:P {target_ca_ratio}:1 Target"})
                tot = recalculate_ration_totals(ingredients_list) # RECALCULATE!

    if tot["vit_a"] < target_vit_a:
        vit_a_def = target_vit_a - tot["vit_a"]
        if vit_a_def > 10.0:
            if spezies == "Hund":
                supp_name = "Karotten / Möhren (püriert)"
                if not is_excluded(supp_name, unvert_list) and not is_excluded("Möhren", unvert_list) and not is_excluded("Karotten", unvert_list) and v_ing != supp_name:
                    add_g = round(vit_a_def / 17.0, 1)
                    add_g = min(60.0 * ratio_share, max(5.0, add_g))
                    if add_g >= 1.0:
                        ingredients_list.append({"Zutat": supp_name, "Menge": f"{add_g} g", "Status / Hinweis": "✨ Auto-Suggester: Provitamin A"})
                        tot = recalculate_ration_totals(ingredients_list) # RECALCULATE!
            else:
                supp_name = None
                if not is_excluded("Hühnerleber", unvert_list) and not is_excluded("Geflügel", unvert_list):
                    supp_name = "Hühnerleber"
                elif not is_excluded("Rinderleber", unvert_list) and not is_excluded("Rind", unvert_list):
                    supp_name = "Rinderleber"

                if supp_name:
                    vit_a_per_g = INGREDIENTS_DB[supp_name]["vit_a"] / 100.0
                    add_liver_g = round(vit_a_def / vit_a_per_g, 1)
                    add_liver_g = min(round(kgw * 1.0 * ratio_share, 1), max(0.5, add_liver_g))
                    if add_liver_g >= 0.5:
                        ingredients_list.append({"Zutat": supp_name, "Menge": f"{add_liver_g} g", "Status / Hinweis": "✨ Auto-Suggester (Katze): Preformes Vitamin A"})
                        tot = recalculate_ration_totals(ingredients_list) # RECALCULATE!

    if spezies == "Katze" and tot["taurin"] < target_taurin:
        taurin_def_mg = target_taurin - tot["taurin"]
        if taurin_def_mg > 10.0:
            supp_name = "Taurin (Reinsubstanz 100%)"
            if not is_excluded(supp_name, unvert_list) and not is_excluded("Taurin", unvert_list):
                add_taurin_g = round(taurin_def_mg / 1000.0, 2)
                ingredients_list.append({"Zutat": supp_name, "Menge": f"{add_taurin_g} g (~{int(taurin_def_mg)} mg)", "Status / Hinweis": "✨ Auto-Suggester (Katze): Essenzielles Taurin"})
                tot = recalculate_ration_totals(ingredients_list) # RECALCULATE!

    is_forthyron = any("forthyron" in str(m).lower() for m in med_list)
    if not is_forthyron and tot["jod"] < target_jod:
        jod_def = target_jod - tot["jod"]
        if jod_def > 5.0:
            supp_name = "Seealgenmehl (Jodspender)"
            if not is_excluded(supp_name, unvert_list) and not is_excluded("Seealgenmehl", unvert_list) and not is_excluded("Algen", unvert_list):
                add_seaweed_g = round(jod_def / 600.0, 2)
                if add_seaweed_g >= 0.05:
                    ingredients_list.append({"Zutat": supp_name, "Menge": f"{add_seaweed_g} g (Prise)", "Status / Hinweis": "✨ Auto-Suggester: Jod-Ausgleich"})
                    tot = recalculate_ration_totals(ingredients_list) # RECALCULATE!

    if tot["zn"] < target_zn:
        zn_def_mg = target_zn - tot["zn"]
        if zn_def_mg > 0.5:
            supp_name = "Zink-Chelat (Spurenelement)"
            if not is_excluded(supp_name, unvert_list) and not is_excluded("Zink", unvert_list):
                ingredients_list.append({"Zutat": supp_name, "Menge": f"{round(zn_def_mg, 1)} mg Zink (~0.1 g Pulver)", "Status / Hinweis": "✨ Auto-Suggester: Spurenelement Zink-Ausgleich"})
                tot = recalculate_ration_totals(ingredients_list) # RECALCULATE!

    praxis_hinweis = None
    if "Lachsfilet" in p_ing:
        praxis_hinweis = "💡 Praxishinweis: Lachs liefert wertvolle marine Omega-3-Fettsäuren (EPA/DHA)."

    # Evaluate complete coverage
    eval_res = evaluiere_bedarfsdeckung(tot, bedarf, ratio_share=ratio_share)

    return {
        "ingredients": ingredients_list,
        "totals": tot,
        "eval": eval_res,
        "praxis_hinweis": praxis_hinweis
    }


def build_growth_curve_pdf_bytes(patient, growth_class, end_target, months, soll_weights, soll_lower, soll_upper, wiegehistorie=[], is_katze=True):
    import io, matplotlib.pyplot as plt
    pdf = PDFBase()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(20, 40, 80)
    sp_title = "KATZE" if is_katze else "HUND"
    pdf.cell(0, 8, pdf.sanitize(f"RASSESPEZIFISCHE WACHSTUMSKURVE & GEWICHTSKONTROLLE ({sp_title})"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(100, 100, 100)
    fall_id = patient.get("fall_id", "FALL-2026-0001")
    pdf.cell(0, 5, pdf.sanitize(f"Wissenschaftliche Entwicklungsanalyse | Fall-ID: {fall_id} | Patient: {patient['name']}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.ln(2)

    # Patient Card
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 200, 220)
    y_start = pdf.get_y()
    pdf.rect(10, y_start, 190, 24, style="FD")
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(20, 40, 80)
    pdf.set_xy(12, y_start + 2)

    cls_short = growth_class.split("(")[0].strip() if "(" in growth_class else growth_class
    pdf.cell(90, 5, pdf.sanitize(f"Patient: {patient['name']} ({patient['spezies']})"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Erwartetes Endgewicht: {end_target} kg"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Aktuelles Gewicht: {round(patient['ist_gw'],2)} kg"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Modell-Klasse: {cls_short}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(180, 5, pdf.sanitize(f"Ernährungsstatus: {patient['lebensphase']} | Tagesenergie (MER): {patient['mer_kcal']} kcal/Tag"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(y_start + 27)

    # Render Matplotlib Figure
    fig, ax = plt.subplots(figsize=(9, 3.8), dpi=200)
    line_color = "#2ca02c" if is_katze else "#1f77b4"
    ax.fill_between(months, soll_lower, soll_upper, color=line_color, alpha=0.15, label="±10% Toleranzbereich (Gesunder Korridor)")
    ax.plot(months, soll_lower, linestyle="--", color=line_color, alpha=0.4, linewidth=1)
    ax.plot(months, soll_upper, linestyle="--", color=line_color, alpha=0.4, linewidth=1)
    ax.plot(months, soll_weights, marker="o", color=line_color, linewidth=2.5, label="Soll-Gewichtskurve (kg)")

    if wiegehistorie:
        m_list = [pt["Monat"] for pt in wiegehistorie]
        kg_list = [pt["Gewicht"] for pt in wiegehistorie]
        ax.plot(m_list, kg_list, marker="s", color="#d62728", linewidth=2.2, linestyle="-", label=f"Gemessener Verlauf: {patient['name']}")
        for pt in wiegehistorie:
            ax.annotate(f"{pt['Gewicht']}kg", (pt["Monat"], pt["Gewicht"]), textcoords="offset points", xytext=(0,6), ha='center', fontsize=7.5, fontweight='bold', color="#d62728")

    ax.set_title(f"Soll-Wachstumskurve mit ±10% Toleranzbereich ({cls_short})", fontsize=10, fontweight="bold")
    ax.set_xlabel("Alter in Monaten", fontsize=9)
    ax.set_ylabel("Körpergewicht (kg)", fontsize=9)
    ax.grid(True, linestyle="--", alpha=0.6)
    ax.set_xticks(months)
    ax.legend(loc="upper left", fontsize=8)
    plt.tight_layout()

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)

    pdf.image(buf, x=15, w=180)
    pdf.ln(3)

    # 1. Soll-Gewichtstabelle
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 5, pdf.sanitize("1. SOLL-GEWICHTSTABELLE NACH ALTER (MONATE)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font("Helvetica", "B", 8.0)
    pdf.set_fill_color(30, 60, 110)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(35, 5, pdf.sanitize("Alter (Monate)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
    pdf.cell(45, 5, pdf.sanitize("Soll-Gewicht (kg)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
    pdf.cell(60, 5, pdf.sanitize("±10% Korridor (kg)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
    pdf.cell(50, 5, pdf.sanitize("Prozent vom Endgewicht"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True, align="C")

    pdf.set_font("Helvetica", "", 8.0)
    pdf.set_text_color(30, 30, 30)
    for i, m in enumerate(months):
        sw = soll_weights[i]
        slo = soll_lower[i]
        sup = soll_upper[i]
        pct = round((sw / end_target) * 100) if end_target > 0 else 0
        bg = (245, 248, 252) if i % 2 == 0 else (255, 255, 255)
        pdf.set_fill_color(*bg)
        pdf.cell(35, 4.5, pdf.sanitize(f"{m} Mon."), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
        pdf.cell(45, 4.5, pdf.sanitize(f"{sw:.2f} kg"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
        pdf.cell(60, 4.5, pdf.sanitize(f"{slo:.2f} kg - {sup:.2f} kg"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
        pdf.cell(50, 4.5, pdf.sanitize(f"{pct} %"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True, align="C")

    pdf.ln(3)

    # 2. Wiegehistorie Table
    if wiegehistorie:
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(20, 40, 80)
        pdf.cell(0, 5, pdf.sanitize("2. ERFASSTE GEMESSENE WIEGEHISTORIE"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "B", 8.0)
        pdf.set_fill_color(30, 60, 110)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(30, 5, pdf.sanitize("Messpunkt"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
        pdf.cell(30, 5, pdf.sanitize("Alter (Mon.)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
        pdf.cell(35, 5, pdf.sanitize("Gemessen (kg)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
        pdf.cell(45, 5, pdf.sanitize("Abweichung / Status"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
        pdf.cell(50, 5, pdf.sanitize("Notiz / Befund"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True, align="C")

        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(30, 30, 30)
        for idx, pt in enumerate(wiegehistorie):
            m_val = pt["Monat"]
            g_val = pt["Gewicht"]
            notiz = pt.get("Notiz", "")
            closest_sw = soll_weights[min(range(len(months)), key=lambda i: abs(months[i]-m_val))]
            diff_pct = round(((g_val - closest_sw) / closest_sw) * 100.0, 1) if closest_sw > 0 else 0
            if abs(diff_pct) <= 10.0:
                st_str = f"Ideal ({diff_pct:+.1f}%)"
            elif diff_pct > 10.0:
                st_str = f"Über Korridor ({diff_pct:+.1f}%)"
            else:
                st_str = f"Unter Korridor ({diff_pct:+.1f}%)"

            bg = (245, 248, 252) if idx % 2 == 0 else (255, 255, 255)
            pdf.set_fill_color(*bg)
            pdf.cell(30, 4.5, pdf.sanitize(f"Nr. {idx+1}"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
            pdf.cell(30, 4.5, pdf.sanitize(f"{m_val} Mon."), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
            pdf.cell(35, 4.5, pdf.sanitize(f"{g_val:.2f} kg"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
            pdf.cell(45, 4.5, pdf.sanitize(st_str), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True, align="C")
            pdf.cell(50, 4.5, pdf.sanitize(notiz), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True, align="L")

    return bytes(pdf.output())


def build_bedarfsanalyse_pdf_bytes(patient):
    pdf = PDFBase()
    pdf.set_auto_page_break(auto=True, margin=15)
    is_katze = patient.get("spezies", "Hund") == "Katze"
    
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(20, 40, 80)
    title_str = f"VET-GUTACHTEN: FELINE BEDARFSANALYSE FÜR {patient['name'].upper()}" if is_katze else f"VET-GUTACHTEN: NÄHRSTOFF-BEDARFSANALYSE FÜR {patient['name'].upper()}"
    pdf.cell(0, 8, pdf.sanitize(title_str), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.ln(2)

    # Patient Card
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 200, 220)
    y_start = pdf.get_y()
    pdf.rect(10, y_start, 190, 24, style="FD")
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(20, 40, 80)
    pdf.set_xy(12, y_start + 2)

    unv_str = ", ".join(patient.get("unvertraeglichkeiten", [])) if patient.get("unvertraeglichkeiten") else "Keine"
    med_str = ", ".join(patient.get("medikation", [])) if patient.get("medikation") else "Keine"

    pdf.cell(90, 5, pdf.sanitize(f"Patient: {patient['name']} ({patient['spezies']})"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Alter & Status: {get_alter_display_str(patient['alter'])} ({patient['lebensphase']}) | {round(patient['ist_gw'],2)} kg (Soll: {round(patient['soll_gw'],2)} kg)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Energiebedarf (MER): {patient['mer_kcal']} kcal ME/Tag ({patient['mer_mj']} MJ)"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Mahlzeiten: {patient.get('mahlzeiten_n', 3)} Portionen pro Tag"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Unverträglichkeiten: {unv_str}"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Medikation / Anamnese: {med_str}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(y_start + 28)

    import streamlit as st
    active_model = patient.get("modell", st.session_state.get("k_f_modell" if is_katze else "h_f_modell", "FEDIAF 4-Stufen-Modell"))

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    sec_title = f"1. ENERGIE- & NÄHRSTOFF-SOLLWERTE ({active_model.upper()} / NRC NORM)"
    pdf.cell(0, 6, pdf.sanitize(sec_title), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 9.0)
    
    if is_katze:
        taurin_val = patient.get('taurin_mg', round(patient.get('ist_gw', 4.0) * 25.0, 0))
        rows = [
            ("Metabolisches Körpergewicht (feline)", f"{patient['km_meta']} kg^0.67"),
            ("Tages-Energiebedarf (MER)", f"{patient['mer_kcal']} kcal / Tag ({patient['mer_mj']} MJ)"),
            ("Verdauliches Rohprotein (vRp)", f"{patient['vrp_g']} g vRp (Brutto-Rp: {patient['rp_g']} g)"),
            ("Essenzielles Taurin Soll", f"{taurin_val} mg / Tag (mind. 25 mg/kg KGW)"),
            ("Calcium (Ca) Soll", f"{patient['ca_mg']} mg / Tag"),
            ("Phosphor (P) Soll", f"{patient['p_mg']} mg / Tag"),
            ("Ca:P-Verhältnis Ziel", f"{patient['ca_p_ratio']}"),
            ("Jod (I) Soll", f"{patient['jod_ug']} µg / Tag"),
            ("Preformes Vitamin A Soll", f"{patient['vit_a_ie']} IE / Tag (Tierische Herkunft)"),
            ("Vitamin D3 Soll", f"{patient['vit_d_ie']} IE / Tag"),
            ("Eisen (Fe) / Zink (Zn) Soll", f"{patient['fe_mg']} mg Fe / {patient['zn_mg']} mg Zn")
        ]
    else:
        rows = [
            ("Metabolisches Körpergewicht", f"{patient['km_meta']} kg^0.75"),
            ("Tages-Energiebedarf (MER)", f"{patient['mer_kcal']} kcal / Tag ({patient['mer_mj']} MJ)"),
            ("Verdauliches Rohprotein (vRp)", f"{patient['vrp_g']} g vRp (Brutto-Rp: {patient['rp_g']} g)"),
            ("Calcium (Ca) Soll", f"{patient['ca_mg']} mg / Tag"),
            ("Phosphor (P) Soll", f"{patient['p_mg']} mg / Tag"),
            ("Ca:P-Verhältnis Ziel", f"{patient['ca_p_ratio']}"),
            ("Jod (I) Soll", f"{patient['jod_ug']} µg / Tag"),
            ("Vitamin A Soll", f"{patient['vit_a_ie']} IE / Tag"),
            ("Vitamin D3 Soll", f"{patient['vit_d_ie']} IE / Tag"),
            ("Eisen (Fe) / Zink (Zn) Soll", f"{patient['fe_mg']} mg Fe / {patient['zn_mg']} mg Zn")
        ]

    for label, val in rows:
        pdf.cell(80, 5.5, pdf.sanitize(label), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L")
        pdf.cell(110, 5.5, pdf.sanitize(val), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")

    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("2. ERNÄHRUNGSPHYSIOLOGISCHE ERKLÄRUNGEN UND BEGRÜNDUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.ln(2)

    if is_katze:
        notes = [
            ("Metabolischer Energiebedarf (MER):", f"Der MER von {patient['mer_kcal']} kcal beruht auf der felinen Potenzauszahlung ({patient['km_meta']} kg^0.67) unter Einbezug von {active_model}. Für das Ziel '{patient['ziel']}' ist der Energiebedarf exakt auf das Zielgewicht von {patient['soll_gw']} kg eingestellt."),
            ("Essenzielles Taurin & Proteinstoffwechsel:", f"Katzen besitzen eine permanent hohe Aminosäure-Abbau-Rate und synthetisieren kein eigenes Taurin. Mit {patient['vrp_g']} g vRp und {patient.get('taurin_mg', round(patient.get('ist_gw',4.0)*25.0,0))} mg Taurin werden DCM (Herzerkrankung) und Netzhautatrophie zuverlässig vorgebeugt."),
            ("Feline Calcium-Phosphor-Balance:", f"Das Calcium-Phosphor-Verhältnis wird bei Katzen streng auf den nierenfreundlichen Idealbereich von {patient['ca_p_ratio']} eingestellt, um Harngries (Struvit) und Nierenüberlastung zu verhindern.")
        ]
        
        is_kitten = patient.get("lebensphase") in ["Kitten / Wachstum", "Kitten", "Junior / Wachstum"]
        if is_kitten:
            notes.append(("🐱 KITTEN-FÜTTERUNGSHINWEIS (WACHSTUMSFUTTER):", "WICHTIG FÜR DIE AUFZUCHT: Heranwachsende Kitten benötigen ein höheres Calcium-Phosphor-Verhältnis (1,2-1,5 : 1) und eine höhere Proteindichte für das Skelett- und Gewebewachstum. Bei Verfütterung von Fertigfutter wählen Sie zwingend ein deklariertes 'Kitten-Alleinfuttermittel'! Füttern Sie Kitten im Wachstum ad-libitum bzw. in vielen kleinen Tagesportionen, um eine gesunde Entwicklung zu gewährleisten."))
        else:
            notes.append(("Preformes Vitamin A & Wirkstoffversorgung:", f"Katzen können pflanzliches Beta-Carotin nicht in Vitamin A umwandeln. Vitamin A ({patient['vit_a_ie']} IE) muss als preformes Retinol aus tierischer Quelle (Leber/Fleisch) zugeführt werden."))
    else:
        notes = [
            ("Metabolischer Energiebedarf (MER):", f"Der MER von {patient['mer_kcal']} kcal beruht auf dem metabolischen Körpergewicht ({patient['km_meta']} kg^0.75) unter Einbezug von {active_model}. Für das Ziel '{patient['ziel']}' ist der Energiebedarf exakt auf das Zielgewicht von {patient['soll_gw']} kg eingestellt."),
            ("Proteingehalt & Muskelerhaltung:", f"Mit {patient['vrp_g']} g verdaulichem Rohprotein wird der Erhalt der aktiven Körper- und Muskelmasse sichergestellt, ohne die Nieren und Leber durch unnötigen Eiweißüberschuss zu belasten."),
            ("Calcium-Phosphor-Balance:", f"Das Calcium-Phosphor-Verhältnis wird nieren- und knochenfreundlich auf den Zielkorridor von {patient['ca_p_ratio']} eingestellt, um Skelettschäden und Nierenverkalkungen vorzubeugen."),
            ("Vitamine & Spurenelemente:", f"Jod ({patient['jod_ug']} µg) sichert die Schilddrüse, Vitamin A ({patient['vit_a_ie']} IE) Sehschärfe/Haut, Vitamin D3 ({patient['vit_d_ie']} IE) den Knochenbau und Zink ({patient['zn_mg']} mg) das Immunsystem.")
        ]

    for n_title, n_desc in notes:
        pdf.set_font("Helvetica", "B", 9.0)
        pdf.set_text_color(30, 60, 110)
        pdf.cell(0, 5, pdf.sanitize(f"• {n_title}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.set_font("Helvetica", "", 8.2)
        pdf.set_text_color(40, 40, 40)
        pdf.multi_cell(0, 4.2, pdf.sanitize(n_desc))
        pdf.ln(2)

    return bytes(pdf.output())

def build_fertigfutter_pdf_bytes(patient, futter_res):
    pdf = PDFBase()
    pdf.set_auto_page_break(auto=True, margin=15)
    is_katze = patient.get("spezies", "Hund") == "Katze"
    
    # PAGE 1: WEENDER-ANALYSE & DEKLARIERTE ZUSATZSTOFFE
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 8, pdf.sanitize(f"FERTIGFUTTER-EVALUIERUNG & KLINISCHE BEURTEILUNG FÜR {patient['name'].upper()}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.ln(2)

    # Patient Card Box
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 200, 220)
    y_start = pdf.get_y()
    pdf.rect(10, y_start, 190, 22, style="FD")
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(20, 40, 80)
    pdf.set_xy(12, y_start + 2)

    unv_str = ", ".join(patient.get("unvertraeglichkeiten", [])) if patient.get("unvertraeglichkeiten") else "Keine"
    med_str = ", ".join(patient.get("medikation", [])) if patient.get("medikation") else "Keine"

    pdf.cell(90, 5, pdf.sanitize(f"Patient: {patient['name']} ({patient['spezies']})"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Status: {patient['alter']} J. ({patient['lebensphase']}) | {patient['ist_gw']} kg (Soll: {patient['soll_gw']} kg)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Tagesenergie (MER): {patient['mer_kcal']} kcal ME/Tag"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Proteinziel: {patient['rp_g']} g | Ca:P-Ziel: {patient['ca_p_ratio']}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Unverträglichkeiten: {unv_str}"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Medikation: {med_str}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(y_start + 26)

    # Section 1: WEENDER ANALYSENDATEN
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize(f"1. WEENDER-ANALYSENDATEN: {futter_res.get('name','Fertigfutter').upper()}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)

    w_rows = [
        ("Futtertyp / Spezies", f"{futter_res.get('typ','Trockenfutter')} ({patient['spezies']})"),
        ("Rohprotein (Rp) / Rohfett (Rfe)", f"{futter_res.get('rp',0)} %  /  {futter_res.get('rfe',0)} %"),
        ("Rohfaser (Rfa) / Rohasche (Ra)", f"{futter_res.get('rfa',0)} %  /  {futter_res.get('ra',0)} %"),
        ("Feuchtigkeit / Trockensubstanz (TS)", f"{futter_res.get('feuchte',0)} %  /  TS: {futter_res.get('ts_pct',0)} %"),
        ("Kohlenhydrate (NfE in TS)", f"{futter_res.get('nfe_ts',0)} % i.TS"),
        ("Energiedichte (Modell)", f"{futter_res.get('kcal_100g',0)} kcal / 100g ({futter_res.get('modell','FEDIAF 4-Stufen')})"),
        ("Calcium (Ca) / Phosphor (P) Ratio", f"{futter_res.get('ca',0)} % Ca  /  {futter_res.get('p',0)} % P  (Ca:P {futter_res.get('ca_p_ratio','n.a.')} : 1)")
    ]
    for l1, l2 in w_rows:
        pdf.cell(75, 4.5, pdf.sanitize(l1), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L")
        pdf.cell(115, 4.5, pdf.sanitize(l2), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")

    pdf.ln(3)

    # Section 2: DEKLARIERTE ZUSATZSTOFFE & KONSERVIERUNG
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("2. DEKLARIERTE ERNÄHRUNGSPHYSIOLOGISCHE & TECHNOLOGISCHE ZUSATZSTOFFE"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)

    tau_val = f"{futter_res.get('taurin', 0)} mg / kg" if futter_res.get('taurin', 0) > 0 else ("Nicht separat deklariert" if is_katze else "Nicht erforderlich (Hund)")
    vit_a_val = f"{futter_res.get('vit_a', 0)} IE / kg" if futter_res.get('vit_a', 0) > 0 else "Keine Angabe"
    vit_d_val = f"{futter_res.get('vit_d', 0)} IE / kg" if futter_res.get('vit_d', 0) > 0 else "Keine Angabe"
    jod_val = f"{futter_res.get('jod', 0)} mg / kg" if futter_res.get('jod', 0) > 0 else "Keine Angabe"
    zn_val = f"{futter_res.get('zn', 0)} mg / kg" if futter_res.get('zn', 0) > 0 else "Keine Angabe"
    antiox_val = futter_res.get("antiox", "") or "Natürliche Antioxidanzien / Keine BHA/BHT Angabe"

    z_rows = [
        ("Essenzielles Taurin (Zusatz)", tau_val),
        ("Vitamin A / Vitamin D3", f"Vit A: {vit_a_val}  |  Vit D3: {vit_d_val}"),
        ("Spurenelemente (Jod / Zink)", f"Jod: {jod_val}  |  Zink: {zn_val}"),
        ("Konservierung / Antioxidanzien", antiox_val)
    ]
    for l1, l2 in z_rows:
        pdf.cell(75, 4.5, pdf.sanitize(l1), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="L")
        pdf.cell(115, 4.5, pdf.sanitize(l2), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")

    pdf.ln(3)

    # Section 3: FÜTTERUNGSEMPFEHLUNG & MAHLZEITEN-PORTIONIERUNG
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("3. PORTIONIERUNG & MAHLZEITEN-VERTEILUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")

    if futter_res.get('kcal_100g', 0) > 0:
        tages_g = round((patient['mer_kcal'] / futter_res['kcal_100g']) * 100.0, 1)
        port_g = round(tages_g / patient.get('mahlzeiten_n', 3), 1)
    else:
        tages_g, port_g = 0, 0

    pdf.set_font("Helvetica", "", 8.5)
    pdf.multi_cell(0, 4.2, pdf.sanitize(f"Für {patient['name']} (Tages-MER: {patient['mer_kcal']} kcal) beträgt die errechnete Futtermenge bei 100% Alleinfütterung exakt {tages_g} g pro Tag (ca. {port_g} g pro Mahlzeit bei {patient.get('mahlzeiten_n', 3)} Portionen täglich)."))

    # PAGE 2: KLINISCHE BEURTEILUNG & OPTIMIERUNGSEMPFEHLUNG
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("4. PATIENTENSPEZIFISCHE EVALUIERUNG & BEURTEILUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.ln(2)

    # Display Positives / Suitability Points
    zusatz_list = futter_res.get("zusatz_eignung", [])
    if zusatz_list:
        pdf.set_fill_color(240, 248, 240)
        pdf.set_draw_color(160, 210, 160)
        y_pos = pdf.get_y()
        pdf.rect(10, y_pos, 190, 6 + len(zusatz_list)*5, style="FD")
        pdf.set_xy(12, y_pos + 2)
        pdf.set_font("Helvetica", "B", 9.0)
        pdf.set_text_color(20, 100, 40)
        pdf.cell(0, 5, pdf.sanitize("🟢 POSITIVE ERNÄHRUNGSEIGENSCHAFTEN & EIGNUNGSPUNKTE:"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(40, 40, 40)
        for z in zusatz_list:
            pdf.set_x(12)
            pdf.cell(0, 4.5, pdf.sanitize(f"• {z}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_y(y_pos + 8 + len(zusatz_list)*5)
        pdf.ln(2)

    # Display Warnings / Critical Points
    warn_list = futter_res.get("warnungen", [])
    if warn_list:
        pdf.set_fill_color(254, 249, 231)
        pdf.set_draw_color(230, 200, 110)
        y_warn = pdf.get_y()
        pdf.rect(10, y_warn, 190, 6 + len(warn_list)*5, style="FD")
        pdf.set_xy(12, y_warn + 2)
        pdf.set_font("Helvetica", "B", 9.0)
        pdf.set_text_color(160, 80, 0)
        pdf.cell(0, 5, pdf.sanitize("⚠️ HINWEISE & KLINISCHE WARNPUNKTE:"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(40, 40, 40)
        for w in warn_list:
            pdf.set_x(12)
            pdf.cell(0, 4.5, pdf.sanitize(f"• {w}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_y(y_warn + 8 + len(warn_list)*5)
        pdf.ln(3)

    # Section 5: OPTIMIERUNG & PRAXIS-EMPFEHLUNG
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("5. OPTIMIERUNGS- & ERGÄNZUNGS-EMPFEHLUNGEN FÜR DEN HALTER"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(40, 40, 40)

    if is_katze:
        opt_text = f"• Taurin-Kontrolle: Katzen haben als obligate Karnivoren einen unersetzbaren Taurinbedarf (25 mg/kg KGW/Tag). Bei fehlender Taurin-Deklaration sollte die Ration täglich mit 1-2 Prisen Taurin-Pulver (~250-500 mg) auf Handwärme ergänzt werden.\n• Fütterungsart & Feuchte: Bei Trockenfütterung muss zwingend auf ausreichende Flüssigkeitszufuhr geachtet werden (Trinkbrunnen, Anfeuchten mit warmem Wasser) oder eine Kombi-Fütterung mit Nassfutter gewählt werden, um Struvitsteinen und Nierenbelastungen (FLUTD) vorzubeugen.\n• Portionierung: Das Futter idealerweise in {patient.get('mahlzeiten_n', 4)} Portionen verfüttern oder bei Trockenfutter als morgendliches Ad-libitum-Depot im Puzzle-Feeder bereitstellen."
    else:
        opt_text = f"• Rations-Anpassung: Bei einer Gesamttagesmenge von {tages_g} g sollte das Futter gleichmäßig auf {patient.get('mahlzeiten_n', 3)} Tagesportionen aufgeteilt werden.\n• Fettsäure-Ergänzung: Zur Optimierung der Hautbarriere und Gelenkgesundheit empfiehlt sich die Zugabe von hochwertigem Lachsöl (1 TL pro 10 kg KGW) für essenzielle Omega-3-Fettsäuren (EPA/DHA).\n• Snack-Budget: Belohnungen und Kauartikel dürfen max. 10% der Tagesenergie ({round(patient['mer_kcal']*0.1,0)} kcal) betragen und müssen von der Tagesfuttermenge abgezogen werden."

    pdf.multi_cell(0, 4.2, pdf.sanitize(opt_text))

    return bytes(pdf.output())

def build_eigenration_pdf_bytes(patient, raw_ration_list, optimized_list, totals, raw_totals=None):
    pdf = PDFBase()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    is_katze = patient.get("spezies", "Hund") == "Katze"
    target_taurin = patient.get("taurin_mg", round(patient['ist_gw'] * 25.0, 0)) if is_katze else 0

    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 8, pdf.sanitize(f"REZEPTUR-GUTACHTEN: EIGENRATION / BARF FÜR {patient['name'].upper()}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.ln(2)

    # Patient Card
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 200, 220)
    y_start = pdf.get_y()
    pdf.rect(10, y_start, 190, 24, style="FD")
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(20, 40, 80)
    pdf.set_xy(12, y_start + 2)

    unv_str = ", ".join(patient.get("unvertraeglichkeiten", [])) if patient.get("unvertraeglichkeiten") else "Keine"
    med_str = ", ".join(patient.get("medikation", [])) if patient.get("medikation") else "Keine"

    pdf.cell(90, 5, pdf.sanitize(f"Patient: {patient['name']} ({patient['spezies']})"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Alter & Status: {get_alter_display_str(patient['alter'])} ({patient['lebensphase']}) | {round(patient['ist_gw'],2)} kg (Soll: {round(patient['soll_gw'],2)} kg)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Tagesenergie (MER): {patient['mer_kcal']} kcal ME/Tag"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Proteinziel: {patient['rp_g']} g | Ca:P-Ziel: {patient['ca_p_ratio']}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Unverträglichkeiten: {unv_str}"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Medikation: {med_str}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(y_start + 28)

    # 1. AKTUELLE REZEPTUR-LISTE (EINGEGEBENE ROH-WERTE DES KUNDEN)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("1. AKTUELLE REZEPTUR-LISTE (EINGEGEBENE ROH-WERTE DES KUNDEN)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)

    if raw_ration_list:
        pdf.set_fill_color(30, 60, 110)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(80, 5.5, pdf.sanitize("Zutat"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
        pdf.cell(50, 5.5, pdf.sanitize("Eingegebene Tagesmenge"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
        pdf.cell(60, 5.5, pdf.sanitize("Kategorie"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)

        pdf.set_text_color(30, 30, 30)
        for item in raw_ration_list:
            ing_name = item.get('name', item.get('Zutat', 'Zutat'))
            ing_m = item.get('menge', item.get('Menge', '0 g'))
            m_str = f"{ing_m} g" if isinstance(ing_m, (int, float)) or (isinstance(ing_m, str) and not ing_m.endswith('g')) else str(ing_m)
            pdf.cell(80, 5, pdf.sanitize(ing_name), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(50, 5, pdf.sanitize(m_str), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(60, 5, pdf.sanitize(item.get('kat', item.get('Status / Hinweis', 'Zutat'))), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        pdf.multi_cell(0, 4.5, pdf.sanitize("Keine Roh-Rezeptur eingegeben (Standard-Baukasten)."))

    pdf.ln(4)

    # 2. IST vs. SOLL NÄHRSTOFF-ANALYSE & MÖGLICHE LANGZEIT-RISIKEN
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("2. IST VS. SOLL NÄHRSTOFF-ANALYSE & MÖGLICHE LANGZEIT-RISIKEN"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    
    if raw_totals:
        n_rows_raw = [
            ("Energie (kcal)", f"{raw_totals.get('kcal', 0)} kcal", f"Soll: {patient['mer_kcal']} kcal"),
            ("Rohprotein (g)", f"{raw_totals.get('rp', 0)} g", f"Soll: {patient['rp_g']} g"),
            ("Calcium (mg) / Phosphor (mg)", f"{raw_totals.get('ca', 0)} mg / {raw_totals.get('p', 0)} mg", f"Soll: {patient['ca_mg']} / {patient['p_mg']} mg"),
            (f"Ca:P-Verhältnis (Ziel {patient['ca_p_ratio']})", f"{raw_totals.get('ca_p_ratio', 0)} : 1", f"Ziel: {patient['ca_p_ratio']}"),
            ("Jod (µg) / Vitamin D3 (IE)", f"{raw_totals.get('jod', 0)} µg / {raw_totals.get('vit_d', 0)} IE", f"Soll: {patient['jod_ug']} µg / {patient['vit_d_ie']} IE")
        ]
        if is_katze:
            n_rows_raw.append(("Essenzielles Taurin (mg)", f"{int(raw_totals.get('taurin', 0))} mg", f"Soll: {int(target_taurin)} mg"))

        for l1, l2, l3 in n_rows_raw:
            pdf.cell(70, 5, pdf.sanitize(l1), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(60, 5, pdf.sanitize(l2), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(60, 5, pdf.sanitize(l3), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(2)
    # Risk Warning Box
    pdf.set_fill_color(254, 249, 231)
    pdf.set_draw_color(230, 200, 110)
    y_rk = pdf.get_y()
    pdf.rect(10, y_rk, 190, 32, style="FD")
    pdf.set_xy(12, y_rk + 2)
    pdf.set_font("Helvetica", "B", 9.0)
    pdf.set_text_color(160, 80, 0)
    pdf.cell(0, 5, pdf.sanitize("🚨 KLINISCHE RISIKO-BEURTEILUNG DER UNBALANCIERTEN ROH-FÜTTERUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 8.0)
    pdf.set_text_color(40, 40, 40)
    pdf.set_x(12)
    if is_katze:
        rk_desc = "Mögliche Langzeit-Risiken ohne Anpassung bei Katzen:\n• Taurin-Mangel: Führt zu kardiologischer Hypertrophie/DCM und irreversibler Netzhautdegeneration.\n• Calcium-Mangel & Ca:P-Imbalance: Renal belastend und Knochenentkalkung.\n• Preformes Vitamin A & Arginin: Katzen können Vitamin A nicht aus Pflanzennahrung bilden; Mangel führt zu Seh- und Immunstörungen.\n• Proteinmangel & Lipidose: Zu geringe Proteinzufuhr verursacht schnellen Muskelschwund und Fettleber."
    else:
        rk_desc = "Mögliche Langzeit-Risiken ohne Anpassung:\n• Calcium-Mangel & Ca:P-Imbalance: Führt zu Knochenentkalkung (renale/alimentäre Osteodystrophie).\n• Proteinexzess / Mangel: Proteinexzedenz belastet Leber und Nieren, während Proteinmangel Muskelschwund verursacht.\n• Spurenelement-Lücken (Jod/Zink): Unzureichendes Jod beeinträchtigt die Schilddrüse; Zinkmangel verursacht Haut- und Fellschäden.\n• Vitamin-Dysbalancen: Vitamin-D3-Exzess führt zu Organverkalkung, während Vitamin-A-Mangel Seh- und Hautstörungen bedingt."
    pdf.multi_cell(186, 3.8, pdf.sanitize(rk_desc))
    pdf.set_y(y_rk + 35)

    # 3. DIE OPTIMIERTE & BEDARFSDECKENDE TAGES-REZEPTUR
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("3. DIE OPTIMIERTE & BEDARFSDECKENDE TAGES-REZEPTUR (UNIVERSAL-GUARD)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.2)
    pdf.set_text_color(40, 40, 40)
    pdf.multi_cell(0, 4.2, pdf.sanitize("Das Universal-Guard System führt eine physiologische Live-Optimierung durch: Basis-Gemüse wird auf max. 3,2 g/kg KGW beschränkt, Leber gegen Vitamin-A-Hypervitaminose auf max. 0,8 g/kg KGW gedeckelt, Knochen zur Vorbeugung von Obstipation (Knochenkot) kontrolliert und Eierschalenpulver exakt auf das nierenfreundliche Calcium-Phosphor-Verhältnis dosiert."))
    pdf.ln(3)

    # 4. OPTIMIERTES & BEDARFSDECKENDES TAGES-REZEPT
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("4. OPTIMIERTES & BEDARFSDECKENDES TAGES-REZEPT"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)

    pdf.set_fill_color(30, 60, 110)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(75, 5.5, pdf.sanitize("Zutat"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
    pdf.cell(55, 5.5, pdf.sanitize("Optimierte Dosis (Küchenmaß)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
    pdf.cell(60, 5.5, pdf.sanitize("Status / Hinweis"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)

    pdf.set_text_color(30, 30, 30)
    for item in optimized_list:
        pdf.cell(75, 5, pdf.sanitize(item['Zutat']), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
        pdf.cell(55, 5, pdf.sanitize(str(item['Menge'])), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
        pdf.cell(60, 5, pdf.sanitize(item.get('Status / Hinweis', 'Balanced')), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(4)

    # 5. ENDGÜLTIGE NÄHRSTOFF-BILANZ DER REZEPTUR
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("5. ENDGÜLTIGE NÄHRSTOFF-BILANZ DER REZEPTUR"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    
    n_rows_opt = [
        ("End-Tages-Energie (kcal)", f"{totals.get('kcal', 0)} kcal", f"Soll: {patient['mer_kcal']} kcal"),
        ("End-Rohprotein (g)", f"{totals.get('rp', 0)} g", f"Soll: {patient['rp_g']} g"),
        ("Calcium (mg) / Phosphor (mg)", f"{totals.get('ca', 0)} mg / {totals.get('p', 0)} mg", f"Soll: {patient['ca_mg']} / {patient['p_mg']} mg"),
        (f"Ca:P-Verhältnis (Ziel {patient['ca_p_ratio']})", f"{totals.get('ca_p_ratio', 0)} : 1", f"Ziel: {patient['ca_p_ratio']}"),
        ("Jod (µg) / Vitamin D3 (IE)", f"{totals.get('jod', 0)} µg / {totals.get('vit_d', 0)} IE", f"Soll: {patient['jod_ug']} µg / {patient['vit_d_ie']} IE")
    ]
    if is_katze:
        n_rows_opt.append(("Essenzielles Taurin (mg)", f"{int(totals.get('taurin', 0))} mg", f"Soll: {int(target_taurin)} mg"))

    for l1, l2, l3 in n_rows_opt:
        pdf.cell(70, 5, pdf.sanitize(l1), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
        pdf.cell(60, 5, pdf.sanitize(l2), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
        pdf.cell(60, 5, pdf.sanitize(l3), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    return bytes(pdf.output())

def build_wochenplan_pdf_bytes(patient, wochen_rows, menu_results={}):
    pdf = PDFBase()
    pdf.set_auto_page_break(auto=True, margin=15)
    is_katze = patient.get("spezies", "Hund") == "Katze"
    is_kitten = is_katze and patient.get("lebensphase") in ["Kitten / Wachstum", "Kitten"]
    
    # PAGE 1: TITLE & PATIENT CARD + SECTION 1 (WÖCHENTLICHER KALENDER)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(20, 40, 80)
    
    if is_katze:
        pdf.cell(0, 8, pdf.sanitize("INDIVIDUELLER FELINER DIÄT- & FÜTTERUNGSPLAN"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(0, 5, pdf.sanitize(f"Spezialisierter Katzen-Ernährungsplan für {patient['name']}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    else:
        pdf.cell(0, 8, pdf.sanitize("INDIVIDUELLER DIÄT- & ERNÄHRUNGSPLAN (HUND)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(0, 5, pdf.sanitize(f"Altersgerechter Fütterungs- & Ernährungsplan für {patient['name']}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    
    pdf.ln(3)

    # Patient Card Box
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 200, 220)
    y_start = pdf.get_y()
    pdf.rect(10, y_start, 190, 24, style="FD")
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(20, 40, 80)
    pdf.set_xy(12, y_start + 2)

    unv_str = ", ".join(patient.get("unvertraeglichkeiten", [])) if patient.get("unvertraeglichkeiten") else "Keine"
    med_str = ", ".join(patient.get("medikation", [])) if patient.get("medikation") else "Keine"

    km_unit = "kg^0.67" if is_katze else "kg^0.75"
    pdf.cell(90, 5, pdf.sanitize(f"Patient: {patient['name']} ({patient['spezies']})"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Alter & Status: {get_alter_display_str(patient['alter'])} ({patient['lebensphase']}) | {round(patient['ist_gw'],2)} kg (Soll: {round(patient['soll_gw'],2)} kg)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Energiebedarf (MER): {patient['mer_kcal']} kcal ME/Tag ({patient['mer_mj']} MJ)"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Mahlzeiten: {patient.get('mahlzeiten_n', 3)} Portionen/Tag | KM: {patient['km_meta']} {km_unit}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Unverträglichkeiten: {unv_str}"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Medikation / Anamnese: {med_str}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(y_start + 28)

    # Section 1 Header
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    if is_katze:
        pdf.cell(0, 6, pdf.sanitize("1. WÖCHENTLICHER FELINER FÜTTERUNGS- & HÄPPCHENKALENDER"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(50, 50, 50)
        pdf.multi_cell(0, 4, pdf.sanitize("Dieses Schema strukturiert die täglichen Nassfutter-Hauptmahlzeiten, Ad-libitum-Trockenfutter-Depots im Puzzle-Feeder und frischen Eigenrationen über den Wochenzyklus hinweg. Bei der Kombi-Fütterung decken Nassfutter und Trockenfutter zusammen exakt 100% MER ab."))
    else:
        pdf.cell(0, 6, pdf.sanitize("1. WÖCHENTLICHER FÜTTERUNGS- & MEDIKATIONSKALENDER (HUND)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(50, 50, 50)
        pdf.multi_cell(0, 4, pdf.sanitize("Dieses Schema strukturiert die täglichen Trockenfutter- und Frischportionen über den Wochenzyklus hinweg. An Frisch- und Ei-Tagen wird die Hauptfuttermenge gezielt angepasst, um Kalorien- und Nährstoffüberschüsse zu vermeiden."))
    pdf.ln(3)

    # Table Header
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_fill_color(30, 60, 110)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(28, 6, pdf.sanitize("Wochentag"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(52, 6, pdf.sanitize("Morgens"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(58, 6, pdf.sanitize("Mittags"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(52, 6, pdf.sanitize("Abends"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", fill=True)

    pdf.set_font("Helvetica", "", 8.0)
    pdf.set_text_color(30, 30, 30)

    for i, row in enumerate(wochen_rows):
        bg = (245, 247, 250) if i % 2 == 0 else (255, 255, 255)
        pdf.set_fill_color(*bg)
        
        tag = row['Tag']
        morg = row['Morgens']
        mitt = row['Mittags']
        abnd = row['Abends']

        y0 = pdf.get_y()
        pdf.set_xy(10, y0)
        pdf.multi_cell(28, 5, pdf.sanitize(tag), border=1, align="C", fill=True)
        h1 = pdf.get_y() - y0
        
        pdf.set_xy(38, y0)
        pdf.multi_cell(52, 4, pdf.sanitize(morg), border=1, align="L", fill=True)
        h2 = pdf.get_y() - y0

        pdf.set_xy(90, y0)
        pdf.multi_cell(58, 4, pdf.sanitize(mitt), border=1, align="L", fill=True)
        h3 = pdf.get_y() - y0

        pdf.set_xy(148, y0)
        pdf.multi_cell(52, 4, pdf.sanitize(abnd), border=1, align="L", fill=True)
        h4 = pdf.get_y() - y0

        max_h = max(h1, h2, h3, h4)
        pdf.set_xy(10, y0 + max_h)

    # PAGE 2: CLINICAL RULES & WARNINGS
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(20, 40, 80)
    
    if is_katze:
        pdf.cell(0, 6, pdf.sanitize("2. FELINE KLINISCHE ERNÄHRUNGSREGELN & HÄPPCHEN-LOGIK"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.ln(3)

        # Rule Box 1 (Amber / Gold)
        pdf.set_fill_color(254, 249, 231)
        pdf.set_draw_color(230, 200, 110)
        y1 = pdf.get_y()
        pdf.rect(10, y1, 190, 32, style="FD")
        pdf.set_xy(12, y1 + 3)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(160, 80, 0)
        pdf.cell(0, 5, pdf.sanitize("FELINER STOFFWECHSEL & AD-LIBITUM HÄPPCHEN-LOGIK"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(40, 40, 40)
        pdf.set_x(12)
        pdf.multi_cell(186, 4, pdf.sanitize("Katzen sind obligate Karnivoren und natürliche Häppchenfresser (8-16 kleine Beutemahlzeiten/Tag). Bei der Kombi-Fütterung decken Nass- und Trockenfutter zusammen exakt 100% MER ab. Das Trockenfutter-Depot wird morgens abgewogen und in einem Puzzle-Feeder oder Snack-Ball bereitgestellt. Das verhindert Schlingen und erhält einen stabilen Urin-pH-Wert (Struvit-Prophylaxe)."))
        pdf.set_y(y1 + 35)

        # Rule Box 2 (Green)
        pdf.set_fill_color(240, 248, 240)
        pdf.set_draw_color(160, 210, 160)
        y2 = pdf.get_y()
        pdf.rect(10, y2, 190, 32, style="FD")
        pdf.set_xy(12, y2 + 3)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(20, 100, 40)
        pdf.cell(0, 5, pdf.sanitize("FELINE CA:P-PRECISION (1.00 : 1) & TAURIN-PROTECTION"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(40, 40, 40)
        pdf.set_x(12)
        pdf.multi_cell(186, 4, pdf.sanitize("Das Calcium-Phosphor-Verhältnis wird bei Katzen streng auf den nierenfreundlichen Idealwert von 1.00 : 1 eingestellt (Ca:P Target 0.9 - 1.1 : 1). Essentielles Taurin (mind. 25 mg/kg KGW) und preformes Vitamin A aus Leber/Fleisch sind zur Vorbeugung von Herzerkrankungen (DCM) und Netzhautdegeneration exakt abgedeckt."))
        pdf.set_y(y2 + 38)

        # Rule Box 3 (Rose / Lipidosis & Kitten Growth Protection)
        pdf.set_fill_color(253, 242, 242)
        pdf.set_draw_color(235, 180, 180)
        y3 = pdf.get_y()
        pdf.rect(10, y3, 190, 32, style="FD")
        pdf.set_xy(12, y3 + 3)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(180, 40, 40)
        
        if is_kitten:
            pdf.cell(0, 5, pdf.sanitize("🚨 ZWINGENDER KITTEN-WACHSTUMSHINWEIS (KEINE DIÄT IM WACHSTUM!)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("Helvetica", "", 8.0)
            pdf.set_text_color(40, 40, 40)
            pdf.set_x(12)
            pdf.multi_cell(186, 3.8, pdf.sanitize("Bei Kitten im Wachstum darf UNTER KEINEN UMSTÄNDEN die Futtermenge gekürzt oder eine Kalorienreduktionsdiät durchgeführt werden! Eine Futterrestriktion führt zu irreversiblen Wachstums- & Organschäden sowie zu lebensgefährlicher hepatischer Lipidose (Fettleber). Mopplige Kitten werden ad-libitum weitergefüttert – das Gewicht reguliert sich durch das anstehende Längenwachstum ('Herauswachsen') von selbst. Bewegung mit Spielzeug gezielt fördern!"))
        else:
            pdf.cell(0, 5, pdf.sanitize("SANFTER GEWICHTSVERLUST (FELINE LIPIDOSE-PROPHYLAXE)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("Helvetica", "", 8.5)
            pdf.set_text_color(40, 40, 40)
            pdf.set_x(12)
            pdf.multi_cell(186, 4, pdf.sanitize("Bei Gewichtsreduktion darf die Abnahmerate max. 0,5 - 1,0 % des Körpergewichts pro Woche betragen (~20-40 g/Woche). Eine zu schnelle Kalorienreduktion ist lebensgefährlich, da sie eine feline hepatische Lipidose (Fettleber) auslösen kann."))
        pdf.set_y(y3 + 35)
    else:
        pdf.cell(0, 6, pdf.sanitize("2. ZWINGENDE KLINISCHE REGELN ZU MEDIKATION & NÄHRSTOFFEN (HUND)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.ln(3)

        # Rule Box 1
        pdf.set_fill_color(254, 249, 231)
        pdf.set_draw_color(230, 200, 110)
        y1 = pdf.get_y()
        pdf.rect(10, y1, 190, 32, style="FD")
        pdf.set_xy(12, y1 + 3)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(160, 80, 0)
        pdf.cell(0, 5, pdf.sanitize("TIMING-REGEL (MEDIKATION & BALLASTSTOFFE)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(40, 40, 40)
        pdf.set_x(12)
        pdf.multi_cell(186, 4, pdf.sanitize("Morgens und abends sind Medikamente (wie Forthyron) nüchtern mit mindestens 2 Stunden Abstand zum Futter einzugeben. Die gekochte Mahlzeit mittags ist ideal, da hier keine Medikamentengabe erfolgt und Ballaststoffe die Wirkstoff-Absorption im Darm nicht beeinträchtigen."))
        pdf.set_y(y1 + 35)

        # Rule Box 2
        pdf.set_fill_color(240, 248, 240)
        pdf.set_draw_color(160, 210, 160)
        y2 = pdf.get_y()
        pdf.rect(10, y2, 190, 32, style="FD")
        pdf.set_xy(12, y2 + 3)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(20, 100, 40)
        pdf.cell(0, 5, pdf.sanitize("DIÄTETISCHE BEGRENZUNGEN (CALCIUM & 10%-LECKERLI-REGEL)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(40, 40, 40)
        pdf.set_x(12)
        pdf.multi_cell(186, 4, pdf.sanitize("Das Calcium-Phosphor-Verhältnis wird in allen Hunde-Rezepturen exakt auf den Zielwert von 1,30 : 1 eingestellt. Snacks und Belohnungen dürfen max. 10% des MER-Tagesbedarfs ausmachen und müssen bei Gewichtsreduktion zwingend von der Hauptration abgezogen werden."))
        pdf.set_y(y2 + 38)

    # PAGE 3: RECIPES & PREPARATION
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(20, 40, 80)
    if is_katze:
        pdf.cell(0, 6, pdf.sanitize("3. DIE OPTIMIERTEN FELINEN REZEPTUREN IM DETAIL"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(50, 50, 50)
        pdf.multi_cell(0, 4, pdf.sanitize("Diese Portionen sind diätetisch so zusammengestellt, dass sie leicht verdauliche, hochwertige tierische Proteine liefern und den felinen Stoffwechsel optimal unterstützen."))
    else:
        pdf.cell(0, 6, pdf.sanitize("3. DIE OPTIMIERTEN REZEPTUREN IM DETAIL (HUND)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(50, 50, 50)
        pdf.multi_cell(0, 4, pdf.sanitize("Diese Portionen sind diätetisch so zusammengestellt, dass sie leicht verdauliche, hochwertige Proteine liefern und die nieren- und gelenkschonenden Bedürfnisse des Patienten optimal unterstützen."))
    pdf.ln(3)

    if menu_results:
        for m_name, m_data in menu_results.items():
            pdf.set_font("Helvetica", "B", 9.5)
            pdf.set_fill_color(230, 236, 245)
            pdf.set_text_color(20, 40, 80)
            tau_str = f" | Taurin: {m_data['totals'].get('taurin',0)}mg" if is_katze else ""
            pdf.cell(0, 5, pdf.sanitize(f"  {m_name} (Nährstoffe: {m_data['totals']['kcal']} kcal | Rp: {m_data['totals']['rp']}g{tau_str} | Ca:P {m_data['totals']['ca_p_ratio']}:1)"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
            
            pdf.set_font("Helvetica", "", 8.0)
            pdf.set_text_color(30, 30, 30)
            for ing in m_data['ingredients']:
                pdf.cell(0, 4.5, pdf.sanitize(f"   * {ing['Zutat']} - {ing['Menge']} ({ing['Status / Hinweis']})"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(2)

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    
    if is_katze:
        pdf.cell(0, 6, pdf.sanitize("4. WISSENSCHAFTLICHE ZUBEREITUNGSRICHTLINIEN FÜR KATZEN"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        rules = [
            ("Taurin-Hitzeschutz:", "Taurin ist hitzeempfindlich! Nach dem Kochen/Dämpfen der Fleischration erst auf Handwärme abkühlen lassen und Taurinpulver frisch einrühren."),
            ("Thiaminase-Schutz:", "Fischstropfen immer schonend dämpfen (ca. 65°C), um das Vitamin-B1-zerstörende Enzym Thiaminase sicher zu inaktivieren."),
            ("Kohlenhydrat-Minimalismus:", "Da Katzen als Karnivoren eine sehr geringe Kohlenhydrattoleranz besitzen (~5% NfE), müssen Stärkequellen (Reis, Kartoffel) sehr weich gekocht und Gemüse fein püriert werden."),
            ("Serviertemperatur:", "Katzenfutter niemals eiskalt direkt aus dem Kühlschrank servieren! Futter immer auf Körpertemperatur (~38°C) anwärmen, um Geruchsaromen freizusetzen und Magenreizungen zu vermeiden.")
        ]
    else:
        pdf.cell(0, 6, pdf.sanitize("4. WISSENSCHAFTLICHE ZUBEREITUNGSRICHTLINIEN (HUND)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
        rules = [
            ("Thiaminase-Schutz:", "Fisch darf niemals roh verfüttert werden! Durch schonendes Garen (Dämpfen bei ca. 65°C) wird das Vitamin-B1-zerstörende Enzym Thiaminase inaktiviert, während seltene Fettsäuren erhalten bleiben."),
            ("Stärkekomplex:", "Süßkartoffeln, Kartoffeln und Reis müssen mindestens 20 Minuten sehr weich gekocht werden, bis sie geliert und für den Hundedarm verdaulich sind."),
            ("Gemüse-Pürierung:", "Da Hunde Pflanzencellulose nicht selbst spalten können, muss gedünstetes Gemüse püriert werden, um Vitamine und Ballaststoffe darmfreundlich aufzuschließen."),
            ("Eiklar-Koagulation:", "Hühnereier stets hartkochen! Dadurch werden Avidin (Biotin-Binder) und Trypsin-Hemmstoffe denaturiert.")
        ]
        
    for r_title, r_desc in rules:
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(30, 60, 110)
        pdf.cell(42, 4.5, pdf.sanitize(r_title), new_x=XPos.RIGHT, new_y=YPos.TOP)
        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(40, 40, 40)
        pdf.multi_cell(148, 4.5, pdf.sanitize(r_desc))
        pdf.ln(1)

    # PAGE 4: CLIENT MONITORING & KITTEN GROWTH GUIDELINES
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("5. HINWEISE FÜR DEN TIERHALTER ZUR NACHKONTROLLE & ERFOLGSMESSUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.ln(2)

    pdf.set_fill_color(245, 248, 252)
    pdf.set_draw_color(180, 200, 225)
    y_mon = pdf.get_y()
    pdf.rect(10, y_mon, 190, 46, style="FD")
    pdf.set_xy(12, y_mon + 3)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 5, pdf.sanitize("🔍 MONITORING-LEITFADEN FÜR DEN TIERHALTER"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 8.0)
    pdf.set_text_color(40, 40, 40)
    pdf.set_x(12)
    
    if is_katze:
        if is_kitten:
            mon_text = f"1. WACHSTUMS- & GEWICHTSKONTROLLE: Wiegen Sie {patient['name']} 1x pro Woche morgens vor der Fütterung. Bei Kitten ist Gewichtszunahme NORMAL und ERWÜNSCHT! Niemals das Futter kürzen.\n2. Akzeptanz & Fressverhalten: Ruhige Häppchenfütterung unterstützen (8-16 kleine Mahlzeiten/Tag im Puzzle-Feeder).\n3. Vitalität & Fell: Achten Sie auf Spieltrieb, glänzendes Fell und geformten Kot (Bristol 2-3).\n4. Fütterungs-Regel: Bei Verdacht auf Übergewicht das Futter NICHT kürzen, sondern Aktivität durch Spielzeug (Federangel, Kletterbaum) steigern. Kitten wachsen aus 'Babyspeck' von selbst heraus."
        else:
            mon_text = f"1. Wöchentliche Gewichtskontrolle: Wiegen Sie {patient['name']} 1x pro Woche morgens vor der Fütterung. Zielkorridor: 0,5% bis 1,0% Gewichtsverlust pro Woche (ca. {round(patient['ist_gw']*0.005,3)} bis {round(patient['ist_gw']*0.010,3)} kg/Woche).\n2. Akzeptanz & Fressverhalten: Ad-libitum Trockenfutter-Depot morgens abwiegen und im Puzzle-Feeder bereitstellen.\n3. Kotbeschaffenheit & Urinierung: Geformter Kot, normale Harnfrequenz zur Struvit-Prophylaxe.\n4. Vitalität: Beobachten Sie Aktivitätsfreude und glänzendes Fell."
    else:
        mon_text = f"1. Wöchentliche Gewichtskontrolle: Wiegen Sie {patient['name']} 1x pro Woche morgens vor der Fütterung. Zielkorridor: 1,0% bis 1,5% Gewichtsverlust pro Woche (ca. {round(patient['ist_gw']*0.010,2)} bis {round(patient['ist_gw']*0.015,2)} kg/Woche bei {patient['ist_gw']} kg).\n2. Verträglichkeit: Achten Sie auf ruhige Futteraufnahme und Ausbleiben von Grasfressen.\n3. Kotbeschaffenheit: Geformt, fest und dunkel (Bristol-Skala 2-3).\n4. Timing-Regel: Bei Medikamenteneingabe (Forthyron) zwingend 2 Std. Abstand zum Futter einhalten."
        
    pdf.multi_cell(186, 4.0, pdf.sanitize(mon_text))

    pdf.set_y(y_mon + 50)

    # Leckerli / Kitten Rule Box
    pdf.set_fill_color(254, 249, 231)
    pdf.set_draw_color(230, 200, 110)
    y_tr = pdf.get_y()
    pdf.rect(10, y_tr, 190, 36, style="FD")
    pdf.set_xy(12, y_tr + 3)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(160, 80, 0)
    
    if is_kitten:
        pdf.cell(0, 5, pdf.sanitize("🦴 KITTEN-FÜTTERUNGS- & LECKERLI-REGELUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(40, 40, 40)
        pdf.set_x(12)
        tr_text = f"• Ad-libitum Versorgungs-Garantie: Heranwachsende Katzenwelpen dürfen fressen, bis sie satt sind. Das Futter wird niemals auf Kalorien-Restriktion gesetzt.\n• Vorrang für Kitten-Alleinfutter: Verwenden Sie Futter, das explizit als Kitten-Nahrung deklariert ist, um den hohen Calcium- (1,4%) und Proteinbedarf (36% i.TS.) sicherzustellen.\n• Geeignete Belohnungen: Gefriergetrocknete reine Fleischhappen (100% Hühnerbrust/Entenbrust). Keine stärkehaltigen Hundesnacks verfüttern!"
    else:
        treat_budget_kcal = round(patient['mer_kcal'] * 0.10, 0)
        pdf.cell(0, 5, pdf.sanitize("🦴 DIE 10%-LECKERLI- & SNACK-REGELUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(40, 40, 40)
        pdf.set_x(12)
        tr_text = f"• Maximales Snack-Budget: Belohnungen dürfen maximal 10 % der Tages-Energie ausmachen (max. {treat_budget_kcal} kcal / Tag).\n• Diät-Anrechnung: Bei Gewichtsreduktion Leckerlis strikt vom Hauptfutter abziehen (z. B. ca. 20g Trockenfutter weniger).\n• Geeignete Leckerlis: Mageres Trockenfleisch (Putenbrust, Rinderlunge) oder kalorienneutrale Snacks."
        
    pdf.multi_cell(186, 4.0, pdf.sanitize(tr_text))

    return bytes(pdf.output())

def build_berater_dossier_pdf_bytes(patient, bedarf, futter_res={}, menu_results={}, wochen_rows=[]):
    import datetime
    pdf = PDFBase()
    pdf.set_auto_page_break(auto=True, margin=15)
    is_katze = patient.get("spezies", "Hund") == "Katze"
    fall_id = patient.get("fall_id", f"FALL-2026-{abs(hash(patient['name']))%8999+1000}")
    bcs_val = patient.get("bcs", 5)
    bcs_title, bcs_desc = get_bcs_info(bcs_val, patient.get("spezies", "Hund"))
    datum_str = datetime.date.today().strftime("%d.%m.%Y")
    
    # PAGE 1: CLINICAL CASE SUMMARY & ANAMNESE
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 8, pdf.sanitize("CLINICAL VET-DOSSIER & EXPERTEN-GUTACHTEN"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, pdf.sanitize(f"Aktenzeichen: {fall_id} | Datum: {datum_str} | Patient: {patient['name']} ({patient['spezies']})"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.ln(3)

    # Patient Card Box
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 200, 220)
    y_start = pdf.get_y()
    pdf.rect(10, y_start, 190, 28, style="FD")
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(20, 40, 80)
    pdf.set_xy(12, y_start + 2)

    unv_str = ", ".join(patient.get("unvertraeglichkeiten", [])) if patient.get("unvertraeglichkeiten") else "Keine"
    med_str = ", ".join(patient.get("medikation", [])) if patient.get("medikation") else "Keine"

    km_unit = "kg^0.67" if is_katze else "kg^0.75"
    pdf.cell(90, 5, pdf.sanitize(f"Patient: {patient['name']} ({patient['spezies']}) | Fall-ID: {fall_id}"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Alter & Status: {get_alter_display_str(patient['alter'])} ({patient['lebensphase']}) | {round(patient['ist_gw'],2)} kg (Soll: {round(patient['soll_gw'],2)} kg)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Körperkondition: BCS {bcs_val}/9 ({bcs_title.split('-')[1].strip() if '-' in bcs_title else 'Normal'})"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Energiebedarf (MER): {patient['mer_kcal']} kcal ME/Tag ({patient['mer_mj']} MJ)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(90, 5, pdf.sanitize(f"Mahlzeiten: {patient.get('mahlzeiten_n', 3)} Portionen / Tag"), new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.cell(90, 5, pdf.sanitize(f"Unverträglichkeiten: {unv_str}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.cell(180, 5, pdf.sanitize(f"Medikation / Anamnese: {med_str}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(y_start + 32)

    # 1.1 CLINICAL CASE SUMMARY
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("1. CLINICAL CASE SUMMARY & ANAMNESE-DIAGNOSTIK"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(40, 40, 40)
    
    if is_katze:
        case_summary = f"• Fall-Hintergrund: Patient {patient['name']} (Katze, {patient['alter']} Jahre, {patient['lebensphase']}) wurde zur Rationsüberprüfung vorgestellt. KGW: {patient['ist_gw']} kg (Ziel: {patient['soll_gw']} kg, '{patient['ziel']}'). Ernährungszustand: BCS {bcs_val}/9.\n• Feline Stoffwechsel-Besonderheiten & Allergien: Unverträglichkeiten: '{unv_str}'. Medikation: '{med_str}'. Als obligater Karnivor besitzt die Katze einen spezifischen Bedarf an essenziellem Taurin (mind. 25 mg/kg KGW), preformtem Vitamin A, Arachidonsäure und Arginin.\n• Klinische Zielsetzung: Erstellung einer 100% bedarfsdeckenden Katzenration zur Erreichung des Zielgewichts (Abnahmerate 0,5-1,0%/Woche zum Schutz vor feline hepatischer Lipidose)."
    else:
        has_forthyron = any("forthyron" in m.lower() for m in patient.get("medikation", [])) or "forthyron" in med_str.lower()
        med_note = "Aufgrund von Schilddrüsenmedikation (Forthyron) muss die Jodzufuhr streng kontrolliert und Seealgenmehl gesperrt werden." if has_forthyron else f"Medikamentöse Versorgung: '{med_str}'."
        case_summary = f"• Fall-Hintergrund: Patient {patient['name']} (Hund, {patient['alter']} Jahre, {patient['lebensphase']}) wurde zur Rationsüberprüfung vorgestellt. KGW: {patient['ist_gw']} kg (Ziel: {patient['soll_gw']} kg, '{patient['ziel']}'). Ernährungszustand: BCS {bcs_val}/9.\n• Anamnestische Risikofaktoren & Allergien: Unverträglichkeiten: '{unv_str}'. {med_note}\n• Klinische Zielsetzung: Erstellung einer 100% bedarfsdeckenden Ration zur sicheren Erreichung des Zielgewichts ohne Muskelverlust."
        
    pdf.multi_cell(0, 4.2, pdf.sanitize(case_summary))
    pdf.ln(3)

    # 1.2 BCS ERKLÄRUNGS-BOX (9-PUNKTE-SCHEMA)
    b_color = (254, 249, 231) if bcs_val in [6,7] else ((253, 242, 242) if bcs_val in [8,9,1,2] else (240, 248, 240))
    d_color = (230, 200, 110) if bcs_val in [6,7] else ((235, 180, 180) if bcs_val in [8,9,1,2] else (160, 210, 160))
    t_color = (160, 80, 0) if bcs_val in [6,7] else ((180, 40, 40) if bcs_val in [8,9,1,2] else (20, 100, 40))

    pdf.set_fill_color(*b_color)
    pdf.set_draw_color(*d_color)
    yb = pdf.get_y()
    pdf.rect(10, yb, 190, 25, style="FD")
    pdf.set_xy(12, yb + 2)
    pdf.set_font("Helvetica", "B", 9.0)
    pdf.set_text_color(*t_color)
    pdf.cell(0, 4.5, pdf.sanitize(f"⚖️ BODY CONDITION SCORE (BCS 1-9): {bcs_title}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 8.0)
    pdf.set_text_color(40, 40, 40)
    pdf.set_x(12)
    pdf.multi_cell(186, 3.8, pdf.sanitize(bcs_desc))
    pdf.set_y(yb + 28)

    # 1.3 KLINISCHE LABORWERTE & ORGANPARAMETER
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 5, pdf.sanitize("🧪 KLINISCHE LABORWERTE & DIAGNOSTISCHE ORGANPARAMETER"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.0)
    pdf.set_text_color(40, 40, 40)
    lab_txt = patient.get("laborwerte", "")
    if not lab_txt or not str(lab_txt).strip():
        lab_txt = "Keine auffälligen pathologischen Laborwerte in der Primäranamnese dokumentiert."
    pdf.multi_cell(0, 4.0, pdf.sanitize(f"• Dokumentierte Befunde: {lab_txt}"))
    pdf.ln(3)

    # PAGE 2: MATHEMATISCHE FORMELN & METHODIK
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("2. MATHEMATISCH-KLINISCHE BERECHNUNGSMETHODIK & ALGORITHMEN"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.0)

    if is_katze:
        f_rows = [
            ("Metabolisches Gewicht (KM^0.67):", f"{patient['km_meta']} kg^0.67 (Formel: KGW^0.67)"),
            ("Erhaltungsenergie (MER):", f"{patient['mer_kcal']} kcal ME / Tag ({patient['mer_mj']} MJ) | MER-Faktor: {round(patient['mer_kcal']/patient['km_meta'],1)} kcal / kg^0.67"),
            ("Verdauliches Rohprotein (vRp):", f"{patient['vrp_g']} g vRp (Formel: 55g vRp / 1000 kcal ME) | Brutto-Rp: {patient['rp_g']} g"),
            ("Calcium- & Phosphor-Soll:", f"Ca: {patient['ca_mg']} mg | P: {patient['p_mg']} mg (Ca:P Target Ratio 1.00 : 1)"),
            ("Essenzielles Taurin Soll:", f"{patient.get('taurin_mg', round(patient['ist_gw']*25.0,0))} mg Taurin / Tag (Formel: 25 mg / kg KGW)"),
            ("Preformes Vitamin A & D3:", f"Vit A: {patient['vit_a_ie']} IE | Vit D3: {patient['vit_d_ie']} IE (Preformes Vit. A aus Leber/Fleisch)")
        ]
    else:
        f_rows = [
            ("Metabolisches Gewicht (KM^0.75):", f"{patient['km_meta']} kg^0.75 (Formel: KGW^0.75)"),
            ("Erhaltungsenergie (MER):", f"{patient['mer_kcal']} kcal ME / Tag ({patient['mer_mj']} MJ) | MER-Faktor: {round(patient['mer_kcal']/patient['km_meta'],1)} kcal / kg^0.75"),
            ("Verdauliches Rohprotein (vRp):", f"{patient['vrp_g']} g vRp (Formel: max(10 * MJ, 4.5 * KM^0.75)) | Brutto-Rp: {patient['rp_g']} g"),
            ("Calcium- & Phosphor-Soll:", f"Ca: {patient['ca_mg']} mg | P: {patient['p_mg']} mg (Ca:P Target Ratio 1.30 : 1)"),
            ("Jod-Sollwert & Schranke:", f"{patient['jod_ug']} µg Jod / Tag (Formel: max(15 * KGW, 400 * MER/1000))"),
            ("Vitamin A & D3 Schranken:", f"Vit A: {patient['vit_a_ie']} IE | Vit D3: {patient['vit_d_ie']} IE | Hypervitaminose-D Schutz")
        ]
        
    for f_name, f_val in f_rows:
        pdf.cell(60, 5.5, pdf.sanitize(f_name), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
        pdf.cell(130, 5.5, pdf.sanitize(f_val), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 6, pdf.sanitize("3. WEENDER-ENERGIE-DIAGNOSTIK & FUTTERMITTEL-EVALUIERUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.2)
    pdf.set_text_color(40, 40, 40)
    
    if futter_res:
        f_info = f"• Analysiertes Fertigfutter: {futter_res.get('name', 'Preset')}\n• Weender-Deklaration: RP {futter_res.get('rp',0)}%, RFe {futter_res.get('rfe',0)}%, RFa {futter_res.get('rfa',0)}%, Ra {futter_res.get('ra',0)}%, Feuchte {futter_res.get('feuchte',0)}%\n• Berechnete Energiedichte ({futter_res.get('modell','FEDIAF 4-Stufen')}): {futter_res.get('kcal_100g',0)} kcal / 100g Futter\n• Trockensubstanz (TS): {futter_res.get('ts_pct',0)} % | NfE i.TS: {futter_res.get('nfe_ts',0)} % | Ca:P Ratio: {futter_res.get('ca_p_ratio','n.a.')} : 1"
        pdf.multi_cell(0, 4.2, pdf.sanitize(f_info))
    else:
        pdf.multi_cell(0, 4.2, pdf.sanitize("Keine Fertigfutter-Weender-Analyse im System hinterlegt."))

    # PAGE 3: NÄHRSTOFF-MATRIZEN ALLER MENÜS & RISIKO-ANALYSE
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("4. NÄHRSTOFF-MATRIZEN ALLER GENERIERTEN MENÜS"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.ln(2)

    if menu_results:
        pdf.set_font("Helvetica", "B", 8.0)
        pdf.set_fill_color(30, 60, 110)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(40, 5, pdf.sanitize("Menü-Variante"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
        pdf.cell(30, 5, pdf.sanitize("Energie (kcal)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
        pdf.cell(30, 5, pdf.sanitize("Rohprotein (g)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
        pdf.cell(45, 5, pdf.sanitize("Calcium / Phosphor (mg)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, fill=True)
        pdf.cell(45, 5, pdf.sanitize("Ca:P Ratio"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)

        pdf.set_font("Helvetica", "", 8.0)
        pdf.set_text_color(30, 30, 30)
        for m_name, m_data in menu_results.items():
            tot = m_data['totals']
            pdf.cell(40, 5, pdf.sanitize(m_name), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(30, 5, pdf.sanitize(f"{tot['kcal']} kcal"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(30, 5, pdf.sanitize(f"{tot['rp']} g"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(45, 5, pdf.sanitize(f"{tot['ca']} / {tot['p']} mg"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP)
            pdf.cell(45, 5, pdf.sanitize(f"{tot['ca_p_ratio']} : 1"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 6, pdf.sanitize("5. KLINISCHE BEURTEILUNG & RISIKO-EVALUIERUNG"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(40, 40, 40)
    
    if is_katze:
        eval_text = "• Befund-Synthese: Das vorliegende feline Ernährungskonzept erfüllt die strengen Nährstoff- und Energieempfehlungen der FEDIAF (2024/2025) für Katzen. Die Deckung von essenziellem Taurin, preformtem Vitamin A und Calcium/Phosphor (1.00 : 1) schützt vor Nephropathien, DCM und Lipidose.\n• Überwachungs-Empfehlungen: Wöchentliches Wiegen zur Sicherstellung der sanften Abnahmerate (0,5-1,0%/Woche). Akzeptanz- und Kotkontrolle. Bei Ad-libitum Kombi-Fütterung morgendliche Abwiegung des Trockenfutter-Depots im Puzzle-Feeder sicherstellen."
    else:
        eval_text = "• Befund-Synthese: Das vorliegende Ernährungskonzept erfüllt die strengen Nährstoff- und Energieempfehlungen der FEDIAF (2024/2025) für Hunde. Das Ca:P-Verhältnis ist mit 1,30 : 1 optimal im physiologischen Sollbereich.\n• Überwachungs-Empfehlungen: 14- und 30-Tage Wiege-Protokoll zur Kontrolle der Abnahmerate (1,0-1,5%/Woche). Kotkonsistency-Protokoll. Einhaltung der 2-Stunden-Timing-Regel bei Medikation."
        
    pdf.multi_cell(0, 4.2, pdf.sanitize(eval_text))

    # PAGE 4: STRUKTURIERTES VERLAUFS- & NACHKONTROLL-PROTOKOLL (RE-CHECK ARCHIV)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 7, pdf.sanitize("6. STRUKTURIERTES VERLAUFS- & NACHKONTROLL-PROTOKOLL (RE-CHECK ARCHIV)"), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="L")
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(60, 60, 60)
    pdf.multi_cell(0, 4.0, pdf.sanitize(f"Dieses Dokumentationsraster dient Ernährungsberater und Tierarzt zur lückenlosen Nachverfolgung des Patientengewichts, der Halter-Compliance und der Rationsanpassungen über einen Betreuungszeitraum von 90 Tagen | Fall-ID: {fall_id}"))
    pdf.ln(3)

    # Re-Check Protocol Table Header
    pdf.set_font("Helvetica", "B", 8.0)
    pdf.set_fill_color(30, 60, 110)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(28, 6, pdf.sanitize("Termin"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(18, 6, pdf.sanitize("KGW (kg)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(18, 6, pdf.sanitize("BCS (1-9)"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(22, 6, pdf.sanitize("Kot-Score"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(52, 6, pdf.sanitize("Halter-Compliance & Befinden"), border=1, new_x=XPos.RIGHT, new_y=YPos.TOP, align="C", fill=True)
    pdf.cell(52, 6, pdf.sanitize("Rations-Anpassung / Notizen"), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", fill=True)

    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(30, 30, 30)

    rechecks = [
        ("Tag 0 (Start)", f"{round(patient['ist_gw'],2)} kg", f"{bcs_val}/9", "Ideal (2-3)", f"Start Beratung | Ziel: {patient['ziel']}", "Initiale Ration freigegeben"),
        ("Tag 14 (Re-Check 1)", "", "", "", "", ""),
        ("Tag 30 (Re-Check 2)", "", "", "", "", ""),
        ("Tag 60 (Re-Check 3)", "", "", "", "", ""),
        ("Tag 90 (Abschluss)", "", "", "", "", "")
    ]

    for idx, (t_name, t_gw, t_bcs, t_kot, t_comp, t_notiz) in enumerate(rechecks):
        bg = (245, 248, 252) if idx == 0 else ((255, 255, 255) if idx % 2 == 1 else (250, 250, 250))
        pdf.set_fill_color(*bg)
        pdf.set_draw_color(180, 200, 220)

        y0 = pdf.get_y()
        h_row = 10.0 if idx > 0 else 7.5

        pdf.rect(10, y0, 190, h_row, style="FD")

        pdf.set_xy(10, y0 + (h_row - 4.0)/2.0)
        pdf.cell(28, 4.0, pdf.sanitize(t_name), align="C")
        pdf.cell(18, 4.0, pdf.sanitize(t_gw), align="C")
        pdf.cell(18, 4.0, pdf.sanitize(t_bcs), align="C")
        pdf.cell(22, 4.0, pdf.sanitize(t_kot), align="C")
        pdf.cell(52, 4.0, pdf.sanitize(t_comp), align="L")
        pdf.cell(52, 4.0, pdf.sanitize(t_notiz), align="L")

        # Vertical dividing lines
        pdf.line(38, y0, 38, y0 + h_row)
        pdf.line(56, y0, 56, y0 + h_row)
        pdf.line(74, y0, 74, y0 + h_row)
        pdf.line(96, y0, 96, y0 + h_row)
        pdf.line(148, y0, 148, y0 + h_row)

        pdf.set_y(y0 + h_row)

    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.set_text_color(20, 40, 80)
    pdf.cell(0, 5, pdf.sanitize("FREIE NOTIZEN DES TIERÄRZTLICHEN ERNÄHRUNGSBERATERS"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(1)
    
    pdf.set_fill_color(255, 255, 255)
    pdf.set_draw_color(200, 200, 200)
    y_notes = pdf.get_y()
    pdf.rect(10, y_notes, 190, 45, style="D")
    
    pdf.set_draw_color(230, 230, 230)
    for ly in range(int(y_notes) + 9, int(y_notes) + 42, 8):
        pdf.line(12, ly, 198, ly)

    return bytes(pdf.output())



def set_cell_background(cell, fill_hex):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill_hex)
    shd.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def style_table(table, col_widths=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Header row
    hdr_row = table.rows[0]
    hdr_row._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
    for cell in hdr_row.cells:
        set_cell_background(cell, "1E3C6E")
        set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for run in p.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.bold = True
                run.font.size = Pt(9.5)
                run.font.name = "Arial"
                
    # Data rows
    for r_idx, row in enumerate(table.rows[1:], start=1):
        row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        bg_color = "F5F8FC" if r_idx % 2 == 1 else "FFFFFF"
        for cell in row.cells:
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.color.rgb = RGBColor(34, 34, 34)
                    run.font.size = Pt(9.5)
                    run.font.name = "Arial"
                    
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                if i < len(row.cells):
                    row.cells[i].width = Inches(w)

def add_doc_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(20, 40, 80)
    return p

def add_heading_styled(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.bold = True
    if level == 1:
        run.font.size = Pt(12.5)
        run.font.color.rgb = RGBColor(20, 40, 80)
    else:
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(30, 60, 110)
    return p

def add_patient_box_docx(doc, patient):
    tbl = doc.add_table(rows=3, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    is_katze = patient.get("spezies", "Hund") == "Katze"
    km_unit = "kg^0.67" if is_katze else "kg^0.75"
    unv_str = ", ".join(patient.get("unvertraeglichkeiten", [])) if patient.get("unvertraeglichkeiten") else "Keine"
    med_str = ", ".join(patient.get("medikation", [])) if patient.get("medikation") else "Keine"
    
    cells_data = [
        [f"Patient: {patient['name']} ({patient['spezies']})", f"Alter & Status: {get_alter_display_str(patient['alter'])} ({patient['lebensphase']}) | {round(patient['ist_gw'],2)} kg (Soll: {round(patient['soll_gw'],2)} kg)"],
        [f"Energiebedarf (MER): {patient['mer_kcal']} kcal ME/Tag ({patient['mer_mj']} MJ)", f"Mahlzeiten: {patient.get('mahlzeiten_n', 3)} Portionen / Tag | KM: {patient['km_meta']} {km_unit}"],
        [f"Unverträglichkeiten: {unv_str}", f"Medikation: {med_str}"]
    ]
    
    col_widths = [3.5, 3.5]
    for r_idx, row in enumerate(tbl.rows):
        row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        for c_idx, cell in enumerate(row.cells):
            set_cell_background(cell, "F0F4F8")
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            cell.width = Inches(col_widths[c_idx])
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(cells_data[r_idx][c_idx])
            run.font.name = "Arial"
            run.font.size = Pt(9.0)
            run.font.bold = True if r_idx == 0 else False
            run.font.color.rgb = RGBColor(20, 40, 80)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

# 1. BUILD BEDARFSANALYSE DOCX BYTES
def build_bedarfsanalyse_docx_bytes(patient):
    doc = docx.Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(0.75)
        
    is_katze = patient.get("spezies", "Hund") == "Katze"
    title_str = f"VET-GUTACHTEN: FELINE BEDARFSANALYSE FÜR {patient['name'].upper()}" if is_katze else f"VET-GUTACHTEN: NÄHRSTOFF-BEDARFSANALYSE FÜR {patient['name'].upper()}"
    add_doc_title(doc, title_str)
    add_patient_box_docx(doc, patient)
    
    active_model = patient.get("modell", "FEDIAF 4-Stufen-Modell")
    add_heading_styled(doc, f"1. ENERGIE- & NÄHRSTOFF-SOLLWERTE ({active_model.upper()} / NRC NORM)", level=1)
    
    if is_katze:
        taurin_val = patient.get('taurin_mg', round(patient.get('ist_gw', 4.0) * 25.0, 0))
        rows = [
            ("Metabolisches Körpergewicht (feline)", f"{patient['km_meta']} kg^0.67"),
            ("Tages-Energiebedarf (MER)", f"{patient['mer_kcal']} kcal / Tag ({patient['mer_mj']} MJ)"),
            ("Verdauliches Rohprotein (vRp)", f"{patient['vrp_g']} g vRp (Brutto-Rp: {patient['rp_g']} g)"),
            ("Essenzielles Taurin Soll", f"{taurin_val} mg / Tag (mind. 25 mg/kg KGW)"),
            ("Calcium (Ca) Soll", f"{patient['ca_mg']} mg / Tag"),
            ("Phosphor (P) Soll", f"{patient['p_mg']} mg / Tag"),
            ("Ca:P-Verhältnis Ziel", f"{patient['ca_p_ratio']}"),
            ("Jod (I) Soll", f"{patient['jod_ug']} µg / Tag"),
            ("Preformes Vitamin A Soll", f"{patient['vit_a_ie']} IE / Tag (Tierische Herkunft)"),
            ("Vitamin D3 Soll", f"{patient['vit_d_ie']} IE / Tag"),
            ("Eisen (Fe) / Zink (Zn) Soll", f"{patient['fe_mg']} mg Fe / {patient['zn_mg']} mg Zn")
        ]
    else:
        rows = [
            ("Metabolisches Körpergewicht", f"{patient['km_meta']} kg^0.75"),
            ("Tages-Energiebedarf (MER)", f"{patient['mer_kcal']} kcal / Tag ({patient['mer_mj']} MJ)"),
            ("Verdauliches Rohprotein (vRp)", f"{patient['vrp_g']} g vRp (Brutto-Rp: {patient['rp_g']} g)"),
            ("Calcium (Ca) Soll", f"{patient['ca_mg']} mg / Tag"),
            ("Phosphor (P) Soll", f"{patient['p_mg']} mg / Tag"),
            ("Ca:P-Verhältnis Ziel", f"{patient['ca_p_ratio']}"),
            ("Jod (I) Soll", f"{patient['jod_ug']} µg / Tag"),
            ("Vitamin A Soll", f"{patient['vit_a_ie']} IE / Tag"),
            ("Vitamin D3 Soll", f"{patient['vit_d_ie']} IE / Tag"),
            ("Eisen (Fe) / Zink (Zn) Soll", f"{patient['fe_mg']} mg Fe / {patient['zn_mg']} mg Zn")
        ]
        
    tbl = doc.add_table(rows=len(rows)+1, cols=2)
    tbl.rows[0].cells[0].paragraphs[0].text = "Nährstoff / Parameter"
    tbl.rows[0].cells[1].paragraphs[0].text = "Berechneter Physiologischer Sollwert"
    for idx, (label, val) in enumerate(rows, start=1):
        tbl.rows[idx].cells[0].paragraphs[0].text = label
        tbl.rows[idx].cells[1].paragraphs[0].text = val
    style_table(tbl, col_widths=[3.0, 4.0])
    
    add_heading_styled(doc, "2. ERNÄHRUNGSPHYSIOLOGISCHE ERKLÄRUNGEN UND BEGRÜNDUNG", level=1)
    
    if is_katze:
        notes = [
            ("Metabolischer Energiebedarf (MER):", f"Der MER von {patient['mer_kcal']} kcal beruht auf der felinen Potenzauszahlung ({patient['km_meta']} kg^0.67) unter Einbezug von {active_model}. Für das Ziel '{patient['ziel']}' ist der Energiebedarf exakt auf das Zielgewicht von {patient['soll_gw']} kg eingestellt."),
            ("Essenzielles Taurin & Proteinstoffwechsel:", f"Katzen besitzen eine permanent hohe Aminosäure-Abbau-Rate und synthetisieren kein eigenes Taurin. Mit {patient['vrp_g']} g vRp und {patient.get('taurin_mg', round(patient.get('ist_gw',4.0)*25.0,0))} mg Taurin werden DCM (Herzerkrankung) und Netzhautatrophie zuverlässig vorgebeugt."),
            ("Feline Calcium-Phosphor-Balance:", f"Das Calcium-Phosphor-Verhältnis wird bei Katzen streng auf den nierenfreundlichen Idealbereich von {patient['ca_p_ratio']} eingestellt, um Harngries (Struvit) und Nierenüberlastung zu verhindern.")
        ]
    else:
        notes = [
            ("Metabolischer Energiebedarf (MER):", f"Der MER von {patient['mer_kcal']} kcal beruht auf dem metabolischen Körpergewicht ({patient['km_meta']} kg^0.75) unter Einbezug von {active_model}. Für das Ziel '{patient['ziel']}' ist der Energiebedarf exakt auf das Zielgewicht von {patient['soll_gw']} kg eingestellt."),
            ("Proteingehalt & Muskelerhaltung:", f"Mit {patient['vrp_g']} g verdaulichem Rohprotein wird der Erhalt der aktiven Körper- und Muskelmasse sichergestellt, ohne die Nieren und Leber durch unnötigen Eiweißüberschuss zu belasten."),
            ("Calcium-Phosphor-Balance:", f"Das Calcium-Phosphor-Verhältnis wird nieren- und knochenfreundlich auf den Zielkorridor von {patient['ca_p_ratio']} eingestellt, um Skelettschäden und Nierenverkalkungen vorzubeugen."),
            ("Vitamine & Spurenelemente:", f"Jod ({patient['jod_ug']} µg) sichert die Schilddrüse, Vitamin A ({patient['vit_a_ie']} IE) Sehschärfe/Haut, Vitamin D3 ({patient['vit_d_ie']} IE) den Knochenbau und Zink ({patient['zn_mg']} mg) das Immunsystem.")
        ]
        
    for n_title, n_desc in notes:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(f"• {n_title} ")
        r1.bold = True
        r1.font.color.rgb = RGBColor(30, 60, 110)
        r2 = p.add_run(n_desc)
        r2.font.color.rgb = RGBColor(40, 40, 40)
        
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# 2. BUILD FERTIGFUTTER DOCX BYTES
def build_fertigfutter_docx_bytes(patient, futter_res):
    doc = docx.Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(0.75)
        
    is_katze = patient.get("spezies", "Hund") == "Katze"
    add_doc_title(doc, f"FERTIGFUTTER-EVALUIERUNG & KLINISCHE BEURTEILUNG FÜR {patient['name'].upper()}")
    add_patient_box_docx(doc, patient)
    
    add_heading_styled(doc, f"1. WEENDER-ANALYSENDATEN: {futter_res.get('name','Fertigfutter').upper()}", level=1)
    
    w_rows = [
        ("Futtertyp / Spezies", f"{futter_res.get('typ','Trockenfutter')} ({patient['spezies']})"),
        ("Rohprotein (Rp) / Rohfett (Rfe)", f"{futter_res.get('rp',0)} %  /  {futter_res.get('rfe',0)} %"),
        ("Rohfaser (Rfa) / Rohasche (Ra)", f"{futter_res.get('rfa',0)} %  /  {futter_res.get('ra',0)} %"),
        ("Feuchtigkeit / Trockensubstanz (TS)", f"{futter_res.get('feuchte',0)} %  /  TS: {futter_res.get('ts_pct',0)} %"),
        ("Kohlenhydrate (NfE in TS)", f"{futter_res.get('nfe_ts',0)} % i.TS"),
        ("Energiedichte (Modell)", f"{futter_res.get('kcal_100g',0)} kcal / 100g ({futter_res.get('modell','FEDIAF 4-Stufen')})"),
        ("Calcium (Ca) / Phosphor (P) Ratio", f"{futter_res.get('ca',0)} % Ca  /  {futter_res.get('p',0)} % P  (Ca:P {futter_res.get('ca_p_ratio','n.a.')} : 1)")
    ]
    
    tbl = doc.add_table(rows=len(w_rows)+1, cols=2)
    tbl.rows[0].cells[0].paragraphs[0].text = "Parameter / Analysenwert"
    tbl.rows[0].cells[1].paragraphs[0].text = "Deklarierter Gehalt / Berechnung"
    for idx, (label, val) in enumerate(w_rows, start=1):
        tbl.rows[idx].cells[0].paragraphs[0].text = label
        tbl.rows[idx].cells[1].paragraphs[0].text = val
    style_table(tbl, col_widths=[3.0, 4.0])
    
    add_heading_styled(doc, "2. DEKLARIERTE ERNÄHRUNGSPHYSIOLOGISCHE & TECHNOLOGISCHE ZUSATZSTOFFE", level=1)
    
    tau_val = f"{futter_res.get('taurin', 0)} mg / kg" if futter_res.get('taurin', 0) > 0 else ("Nicht separat deklariert" if is_katze else "Nicht erforderlich (Hund)")
    vit_a_val = f"{futter_res.get('vit_a', 0)} IE / kg" if futter_res.get('vit_a', 0) > 0 else "Keine Angabe"
    vit_d_val = f"{futter_res.get('vit_d', 0)} IE / kg" if futter_res.get('vit_d', 0) > 0 else "Keine Angabe"
    jod_val = f"{futter_res.get('jod', 0)} mg / kg" if futter_res.get('jod', 0) > 0 else "Keine Angabe"
    zn_val = f"{futter_res.get('zn', 0)} mg / kg" if futter_res.get('zn', 0) > 0 else "Keine Angabe"
    antiox_val = futter_res.get('antiox', 'Keine synthetischen Antioxidanzien deklariert') if futter_res.get('antiox') else 'Keine Angabe'

    z_rows = [
        ("Essenzielles Taurin (Katze)", tau_val),
        ("Vitamin A / Vitamin D3 Zusatz", f"{vit_a_val}  /  {vit_d_val}"),
        ("Spurenelemente Jod / Zink", f"{jod_val}  /  {zn_val}"),
        ("Konservierungsmittel / Antioxidanzien", antiox_val)
    ]
    
    tbl_z = doc.add_table(rows=len(z_rows)+1, cols=2)
    tbl_z.rows[0].cells[0].paragraphs[0].text = "Deklarierter Zusatzstoff"
    tbl_z.rows[0].cells[1].paragraphs[0].text = "Konzentration / Angabe"
    for idx, (label, val) in enumerate(z_rows, start=1):
        tbl_z.rows[idx].cells[0].paragraphs[0].text = label
        tbl_z.rows[idx].cells[1].paragraphs[0].text = val
    style_table(tbl_z, col_widths=[3.0, 4.0])
    
    add_heading_styled(doc, "3. FÜTTERUNGSEMPFEHLUNG UND DOSIERUNG", level=1)
    if futter_res.get('kcal_100g', 0) > 0:
        tages_g = round((patient['mer_kcal'] / futter_res['kcal_100g']) * 100.0, 1)
        port_g = round(tages_g / patient.get('mahlzeiten_n', 3), 1)
    else:
        tages_g, port_g = 0, 0

    p_f = doc.add_paragraph()
    p_f.add_run(f"Für {patient['name']} (Tages-MER: {patient['mer_kcal']} kcal) beträgt die errechnete Futtermenge bei 100% Alleinfütterung exakt {tages_g} g pro Tag (ca. {port_g} g pro Mahlzeit bei {patient.get('mahlzeiten_n', 3)} Portionen täglich).")
    
    add_heading_styled(doc, "4. PATIENTENSPEZIFISCHE EVALUIERUNG & BEURTEILUNG", level=1)
    
    zusatz_list = futter_res.get("zusatz_eignung", [])
    if zusatz_list:
        p_z = doc.add_paragraph()
        r = p_z.add_run("🟢 POSITIVE ERNÄHRUNGSEIGENSCHAFTEN & EIGNUNGSPUNKTE:")
        r.bold = True
        r.font.color.rgb = RGBColor(20, 100, 40)
        for z in zusatz_list:
            p_item = doc.add_paragraph()
            p_item.add_run(f"• {z}")

    warn_list = futter_res.get("warnungen", [])
    if warn_list:
        p_w = doc.add_paragraph()
        r = p_w.add_run("⚠️ HINWEISE & KLINISCHE WARNPUNKTE:")
        r.bold = True
        r.font.color.rgb = RGBColor(160, 80, 0)
        for w in warn_list:
            p_item = doc.add_paragraph()
            p_item.add_run(f"• {w}")

    add_heading_styled(doc, "5. OPTIMIERUNGS- & ERGÄNZUNGS-EMPFEHLUNGEN FÜR DEN HALTER", level=1)
    if is_katze:
        opt_text = f"• Taurin-Kontrolle: Katzen haben als obligate Karnivoren einen unersetzbaren Taurinbedarf (25 mg/kg KGW/Tag). Bei fehlender Taurin-Deklaration sollte die Ration täglich mit 1-2 Prisen Taurin-Pulver (~250-500 mg) auf Handwärme ergänzt werden.\n• Fütterungsart & Feuchte: Bei Trockenfütterung muss zwingend auf ausreichende Flüssigkeitszufuhr geachtet werden (Trinkbrunnen, Anfeuchten mit warmem Wasser) oder eine Kombi-Fütterung mit Nassfutter gewählt werden, um Struvitsteinen und Nierenbelastungen (FLUTD) vorzubeugen.\n• Portionierung: Das Futter idealerweise in {patient.get('mahlzeiten_n', 4)} Portionen verfüttern oder bei Trockenfutter als morgendliches Ad-libitum-Depot im Puzzle-Feeder bereitstellen."
    else:
        opt_text = f"• Rations-Anpassung: Bei einer Gesamttagesmenge von {tages_g} g sollte das Futter gleichmäßig auf {patient.get('mahlzeiten_n', 3)} Tagesportionen aufgeteilt werden.\n• Fettsäure-Ergänzung: Zur Optimierung der Hautbarriere und Gelenkgesundheit empfiehlt sich die Zugabe von hochwertigem Lachsöl (1 TL pro 10 kg KGW) für essenzielle Omega-3-Fettsäuren (EPA/DHA).\n• Snack-Budget: Belohnungen und Kauartikel dürfen max. 10% der Tagesenergie ({round(patient['mer_kcal']*0.1,0)} kcal) betragen und müssen von der Tagesfuttermenge abgezogen werden."

    doc.add_paragraph(opt_text)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# 3. BUILD EIGENRATION DOCX BYTES
def build_eigenration_docx_bytes(patient, raw_ration_list, optimized_list, totals, raw_totals=None):
    doc = docx.Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(0.75)
        
    is_katze = patient.get("spezies", "Hund") == "Katze"
    target_taurin = patient.get("taurin_mg", round(patient['ist_gw'] * 25.0, 0)) if is_katze else 0
    
    add_doc_title(doc, f"REZEPTUR-GUTACHTEN: EIGENRATION / BARF FÜR {patient['name'].upper()}")
    add_patient_box_docx(doc, patient)
    
    add_heading_styled(doc, "1. AKTUELLE REZEPTUR-LISTE (EINGEGEBENE ROH-WERTE DES KUNDEN)", level=1)
    if raw_ration_list:
        tbl_r = doc.add_table(rows=len(raw_ration_list)+1, cols=3)
        tbl_r.rows[0].cells[0].paragraphs[0].text = "Zutat"
        tbl_r.rows[0].cells[1].paragraphs[0].text = "Eingegebene Tagesmenge"
        tbl_r.rows[0].cells[2].paragraphs[0].text = "Kategorie"
        for idx, item in enumerate(raw_ration_list, start=1):
            ing_name = item.get('name', item.get('Zutat', 'Zutat'))
            ing_m = item.get('menge', item.get('Menge', '0 g'))
            m_str = f"{ing_m} g" if isinstance(ing_m, (int, float)) or (isinstance(ing_m, str) and not ing_m.endswith('g')) else str(ing_m)
            tbl_r.rows[idx].cells[0].paragraphs[0].text = ing_name
            tbl_r.rows[idx].cells[1].paragraphs[0].text = m_str
            tbl_r.rows[idx].cells[2].paragraphs[0].text = str(item.get('kat', item.get('Status / Hinweis', 'Zutat')))
        style_table(tbl_r, col_widths=[3.0, 2.0, 2.0])
    else:
        doc.add_paragraph("Keine Roh-Rezeptur eingegeben (Standard-Baukasten).")

    add_heading_styled(doc, "2. IST VS. SOLL NÄHRSTOFF-ANALYSE & MÖGLICHE LANGZEIT-RISIKEN", level=1)
    if raw_totals:
        n_rows_raw = [
            ("Energie (kcal)", f"{raw_totals.get('kcal', 0)} kcal", f"Soll: {patient['mer_kcal']} kcal"),
            ("Rohprotein (g)", f"{raw_totals.get('rp', 0)} g", f"Soll: {patient['rp_g']} g"),
            ("Calcium (mg) / Phosphor (mg)", f"{raw_totals.get('ca', 0)} mg / {raw_totals.get('p', 0)} mg", f"Soll: {patient['ca_mg']} / {patient['p_mg']} mg"),
            ("Ca:P-Verhältnis", f"{raw_totals.get('ca_p_ratio', 0)} : 1", f"Ziel: {patient['ca_p_ratio']}"),
            ("Jod (µg) / Vitamin D3 (IE)", f"{raw_totals.get('jod', 0)} µg / {raw_totals.get('vit_d', 0)} IE", f"Soll: {patient['jod_ug']} µg / {patient['vit_d_ie']} IE")
        ]
        if is_katze:
            n_rows_raw.append(("Essenzielles Taurin (mg)", f"{int(raw_totals.get('taurin', 0))} mg", f"Soll: {int(target_taurin)} mg"))

        tbl_ana = doc.add_table(rows=len(n_rows_raw)+1, cols=3)
        tbl_ana.rows[0].cells[0].paragraphs[0].text = "Nährstoff / Parameter"
        tbl_ana.rows[0].cells[1].paragraphs[0].text = "Berechneter Ist-Wert"
        tbl_ana.rows[0].cells[2].paragraphs[0].text = "Physiologischer Sollwert"
        for idx, (l1, l2, l3) in enumerate(n_rows_raw, start=1):
            tbl_ana.rows[idx].cells[0].paragraphs[0].text = l1
            tbl_ana.rows[idx].cells[1].paragraphs[0].text = l2
            tbl_ana.rows[idx].cells[2].paragraphs[0].text = l3
        style_table(tbl_ana, col_widths=[2.8, 2.1, 2.1])

    add_heading_styled(doc, "3. OPTIMIERTES & BEDARFSDECKENDES TAGES-REZEPT (UNIVERSAL-GUARD)", level=1)
    if optimized_list:
        tbl_opt = doc.add_table(rows=len(optimized_list)+1, cols=3)
        tbl_opt.rows[0].cells[0].paragraphs[0].text = "Zutat"
        tbl_opt.rows[0].cells[1].paragraphs[0].text = "Optimierte Dosis (Küchenmaß)"
        tbl_opt.rows[0].cells[2].paragraphs[0].text = "Status / Hinweis"
        for idx, item in enumerate(optimized_list, start=1):
            tbl_opt.rows[idx].cells[0].paragraphs[0].text = item['Zutat']
            tbl_opt.rows[idx].cells[1].paragraphs[0].text = str(item['Menge'])
            tbl_opt.rows[idx].cells[2].paragraphs[0].text = str(item.get('Status / Hinweis', 'Balanced'))
        style_table(tbl_opt, col_widths=[2.8, 2.1, 2.1])

    add_heading_styled(doc, "4. ENDGÜLTIGE NÄHRSTOFF-BILANZ DER REZEPTUR", level=1)
    n_rows_opt = [
        ("End-Tages-Energie (kcal)", f"{totals.get('kcal', 0)} kcal", f"Soll: {patient['mer_kcal']} kcal"),
        ("End-Rohprotein (g)", f"{totals.get('rp', 0)} g", f"Soll: {patient['rp_g']} g"),
        ("Calcium (mg) / Phosphor (mg)", f"{totals.get('ca', 0)} mg / {totals.get('p', 0)} mg", f"Soll: {patient['ca_mg']} / {patient['p_mg']} mg"),
        (f"Ca:P-Verhältnis (Ziel {patient['ca_p_ratio']})", f"{totals.get('ca_p_ratio', 0)} : 1", f"Ziel: {patient['ca_p_ratio']}"),
        ("Jod (µg) / Vitamin D3 (IE)", f"{totals.get('jod', 0)} µg / {totals.get('vit_d', 0)} IE", f"Soll: {patient['jod_ug']} µg / {patient['vit_d_ie']} IE")
    ]
    if is_katze:
        n_rows_opt.append(("Essenzielles Taurin (mg)", f"{int(totals.get('taurin', 0))} mg", f"Soll: {int(target_taurin)} mg"))

    tbl_fin = doc.add_table(rows=len(n_rows_opt)+1, cols=3)
    tbl_fin.rows[0].cells[0].paragraphs[0].text = "Nährstoff"
    tbl_fin.rows[0].cells[1].paragraphs[0].text = "Finale Rezeptur-Summe"
    tbl_fin.rows[0].cells[2].paragraphs[0].text = "Sollwert Target"
    for idx, (l1, l2, l3) in enumerate(n_rows_opt, start=1):
        tbl_fin.rows[idx].cells[0].paragraphs[0].text = l1
        tbl_fin.rows[idx].cells[1].paragraphs[0].text = l2
        tbl_fin.rows[idx].cells[2].paragraphs[0].text = l3
    style_table(tbl_fin, col_widths=[2.8, 2.1, 2.1])

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# 4. BUILD WOCHENPLAN DOCX BYTES
def build_wochenplan_docx_bytes(patient, wochen_rows, menu_results={}):
    doc = docx.Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(0.75)
        
    is_katze = patient.get("spezies", "Hund") == "Katze"
    title_str = "INDIVIDUELLER FELINER DIÄT- & FÜTTERUNGSPLAN" if is_katze else "INDIVIDUELLER DIÄT- & ERNÄHRUNGSPLAN (HUND)"
    add_doc_title(doc, title_str)
    add_patient_box_docx(doc, patient)
    
    add_heading_styled(doc, "1. WÖCHENTLICHER FÜTTERUNGSKALENDER (MONTAG - SONNTAG)", level=1)
    
    if wochen_rows:
        tbl_w = doc.add_table(rows=len(wochen_rows)+1, cols=4)
        tbl_w.rows[0].cells[0].paragraphs[0].text = "Wochentag"
        tbl_w.rows[0].cells[1].paragraphs[0].text = "Morgens"
        tbl_w.rows[0].cells[2].paragraphs[0].text = "Mittags"
        tbl_w.rows[0].cells[3].paragraphs[0].text = "Abends"
        for idx, row in enumerate(wochen_rows, start=1):
            tbl_w.rows[idx].cells[0].paragraphs[0].text = str(row.get('Tag', ''))
            tbl_w.rows[idx].cells[1].paragraphs[0].text = str(row.get('Morgens', ''))
            tbl_w.rows[idx].cells[2].paragraphs[0].text = str(row.get('Mittags', ''))
            tbl_w.rows[idx].cells[3].paragraphs[0].text = str(row.get('Abends', ''))
        style_table(tbl_w, col_widths=[1.2, 1.9, 2.0, 1.9])

    add_heading_styled(doc, "2. DIE OPTIMIERTEN REZEPTUREN IM DETAIL", level=1)
    if menu_results:
        for m_name, m_data in menu_results.items():
            add_heading_styled(doc, f"🍲 {m_name}", level=2)
            ing_list = m_data.get('ingredients', [])
            if ing_list:
                tbl_m = doc.add_table(rows=len(ing_list)+1, cols=3)
                tbl_m.rows[0].cells[0].paragraphs[0].text = "Zutat"
                tbl_m.rows[0].cells[1].paragraphs[0].text = "Tagesmenge"
                tbl_m.rows[0].cells[2].paragraphs[0].text = "Hinweis"
                for idx, ing in enumerate(ing_list, start=1):
                    tbl_m.rows[idx].cells[0].paragraphs[0].text = str(ing.get('Zutat', ''))
                    tbl_m.rows[idx].cells[1].paragraphs[0].text = str(ing.get('Menge', ''))
                    tbl_m.rows[idx].cells[2].paragraphs[0].text = str(ing.get('Status / Hinweis', ''))
                style_table(tbl_m, col_widths=[2.8, 1.8, 2.4])

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# 5. BUILD BERATER DOSSIER DOCX BYTES
def build_berater_dossier_docx_bytes(patient, bedarf, futter_res={}, menu_results={}, wochen_rows=[]):
    doc = docx.Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(0.75)
        
    is_katze = patient.get("spezies", "Hund") == "Katze"
    fall_id = patient.get("fall_id", "FALL-2026-1001")
    bcs_val = patient.get("bcs", 5)
    
    add_doc_title(doc, "CLINICAL VET-DOSSIER & EXPERTEN-GUTACHTEN")
    add_patient_box_docx(doc, patient)
    
    add_heading_styled(doc, f"1. CLINICAL CASE SUMMARY & ANAMNESE | FALL-ID: {fall_id}", level=1)
    
    doc.add_paragraph(f"Patient {patient['name']} ({patient['spezies']}, {patient['alter']} Jahre) wurde zur Ernährungsberatung vorgestellt. Ist-Gewicht: {patient['ist_gw']} kg, Ziel-Gewicht: {patient['soll_gw']} kg. Body Condition Score: BCS {bcs_val}/9.")
    
    add_heading_styled(doc, "2. STRUKTURIERTES VERLAUFS- & NACHKONTROLL-PROTOKOLL (RE-CHECK ARCHIV)", level=1)
    
    rechecks = [
        ("Tag 0 (Start)", f"{patient['ist_gw']} kg", f"{bcs_val}/9", "Ideal (2-3)", f"Start Beratung | Ziel: {patient['ziel']}", "Initiale Ration freigegeben"),
        ("Tag 14 (Re-Check 1)", "", "", "", "", ""),
        ("Tag 30 (Re-Check 2)", "", "", "", "", ""),
        ("Tag 60 (Re-Check 3)", "", "", "", "", ""),
        ("Tag 90 (Abschluss)", "", "", "", "", "")
    ]
    
    tbl_rc = doc.add_table(rows=len(rechecks)+1, cols=6)
    headers = ["Termin", "KGW (kg)", "BCS", "Kot-Score", "Compliance & Befinden", "Rations-Anpassung"]
    for i, h in enumerate(headers):
        tbl_rc.rows[0].cells[i].paragraphs[0].text = h
        
    for idx, (t_name, t_gw, t_bcs, t_kot, t_comp, t_notiz) in enumerate(rechecks, start=1):
        tbl_rc.rows[idx].cells[0].paragraphs[0].text = t_name
        tbl_rc.rows[idx].cells[1].paragraphs[0].text = t_gw
        tbl_rc.rows[idx].cells[2].paragraphs[0].text = t_bcs
        tbl_rc.rows[idx].cells[3].paragraphs[0].text = t_kot
        tbl_rc.rows[idx].cells[4].paragraphs[0].text = t_comp
        tbl_rc.rows[idx].cells[5].paragraphs[0].text = t_notiz
        
    style_table(tbl_rc, col_widths=[1.2, 0.8, 0.6, 0.9, 1.8, 1.7])

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# 6. BUILD GROWTH CURVE DOCX BYTES
def build_growth_curve_docx_bytes(patient, growth_class, end_target, months, soll_weights, soll_lower, soll_upper, wiegehistorie, is_katze=False):
    doc = docx.Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(0.75)
        
    sp_str = "Katze" if is_katze else "Hund"
    add_doc_title(doc, f"RASSESPEZIPHISCHE {sp_str.upper()}-WACHSTUMSKURVE & ENTWICKLUNGSREPORT")
    add_patient_box_docx(doc, patient)
    
    add_heading_styled(doc, f"1. MODELLKLASSE & PARAMETER: {growth_class.split('(')[0].strip()}", level=1)
    doc.add_paragraph(f"Erwartetes Endgewicht: {end_target} kg | Patient: {patient['name']} ({get_alter_display_str(patient['alter'])}, {patient['ist_gw']} kg)")
    
    # Generate matplotlib figure for docx
    fig, ax = plt.subplots(figsize=(8.0, 3.8))
    color_main = "#2ca02c" if is_katze else "#1f77b4"
    
    ax.fill_between(months, soll_lower, soll_upper, color=color_main, alpha=0.15, label="±10% Toleranzbereich")
    ax.plot(months, soll_lower, linestyle="--", color=color_main, alpha=0.4, linewidth=1)
    ax.plot(months, soll_upper, linestyle="--", color=color_main, alpha=0.4, linewidth=1)
    ax.plot(months, soll_weights, marker="o", color=color_main, linewidth=2.5, label="Soll-Gewichtskurve (kg)")
    
    if wiegehistorie:
        m_list = [pt["Monat"] for pt in wiegehistorie]
        kg_list = [pt["Gewicht"] for pt in wiegehistorie]
        ax.plot(m_list, kg_list, marker="s", color="#d62728", linewidth=2.2, linestyle="-", label=f"Gemessen: {patient['name']}")
        for pt in wiegehistorie:
            ax.annotate(f"{pt['Gewicht']}kg", (pt["Monat"], pt["Gewicht"]), textcoords="offset points", xytext=(0,6), ha='center', fontsize=8, fontweight='bold', color="#d62728")

    ax.set_title(f"Soll-Wachstumskurve mit ±10% Toleranzband ({sp_str})", fontsize=10, fontweight="bold")
    ax.set_xlabel("Alter in Monaten", fontsize=9)
    ax.set_ylabel("Körpergewicht (kg)", fontsize=9)
    ax.grid(True, linestyle="--", alpha=0.5)
    ax.set_xticks(months)
    ax.legend(loc="upper left", fontsize=8)
    plt.tight_layout()
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format="png", dpi=200)
    plt.close(fig)
    img_buf.seek(0)
    
    add_heading_styled(doc, "2. DIAGRAMM WACHSTUMSKURVE & TOLERANZKORRIDOR", level=1)
    doc.add_picture(img_buf, width=Inches(6.2))
    
    add_heading_styled(doc, "3. SOLL-GEWICHTSTABELLE NACH ALTER", level=1)
    tbl_g = doc.add_table(rows=len(months)+1, cols=4)
    tbl_g.rows[0].cells[0].paragraphs[0].text = "Alter"
    tbl_g.rows[0].cells[1].paragraphs[0].text = "Soll-Gewicht (kg)"
    tbl_g.rows[0].cells[2].paragraphs[0].text = "Toleranzkorridor (±10%)"
    tbl_g.rows[0].cells[3].paragraphs[0].text = "Prozent Endgewicht"
    
    pcts = [round(w/end_target*100, 0) for w in soll_weights]
    for idx in range(len(months)):
        tbl_g.rows[idx+1].cells[0].paragraphs[0].text = f"{months[idx]} Mon."
        tbl_g.rows[idx+1].cells[1].paragraphs[0].text = f"{soll_weights[idx]:.2f} kg"
        tbl_g.rows[idx+1].cells[2].paragraphs[0].text = f"{soll_lower[idx]:.2f} - {soll_upper[idx]:.2f} kg"
        tbl_g.rows[idx+1].cells[3].paragraphs[0].text = f"{int(pcts[idx])} %"
        
    style_table(tbl_g, col_widths=[1.5, 1.8, 2.2, 1.5])
    
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# -----------------------------------------------------------------------------
