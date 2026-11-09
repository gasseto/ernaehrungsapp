"""
futter_db.py - Stamm-Datenbanken für Hundefutter, Katzenfutter und BARF-Zutaten (v71)
"""

DEFAULT_FUTTER_DB_HUND = {
    "Belcando Getreidefrei Adult Ozean": {
        "rp": 25.5, "rfe": 15.0, "rfa": 3.3, "ra": 7.5, "feuchte": 10.0, "ca": 1.2, "p": 0.9, "typ": "Trockenfutter"
    },
    "Eukanuba Grain Free Senior Large & Giant mit Seefisch": {
        "rp": 25.5, "rfe": 12.0, "rfa": 2.3, "ra": 7.0, "feuchte": 10.0, "ca": 1.1, "p": 0.9, "typ": "Trockenfutter"
    },
    "Royal Canin Medium Adult": {
        "rp": 25.0, "rfe": 14.0, "rfa": 1.3, "ra": 6.1, "feuchte": 9.5, "ca": 1.2, "p": 0.8, "typ": "Trockenfutter"
    },
    "Josera Festival": {
        "rp": 26.0, "rfe": 16.0, "rfa": 3.0, "ra": 6.5, "feuchte": 9.0, "ca": 1.4, "p": 1.0, "typ": "Trockenfutter"
    },
    "Wolf of Wilderness Green Fields (Lamm)": {
        "rp": 26.0, "rfe": 16.0, "rfa": 3.5, "ra": 7.5, "feuchte": 10.0, "ca": 1.3, "p": 1.0, "typ": "Trockenfutter"
    },
    "Purina BENEFUL Original Rind & Gemüse": {
        "rp": 25.0, "rfe": 12.0, "rfa": 3.0, "ra": 7.5, "feuchte": 10.0, "ca": 1.2, "p": 0.9, "typ": "Trockenfutter"
    },
    "Rinti Kennerfleisch Rind (Nassfutter)": {
        "rp": 10.0, "rfe": 9.0, "rfa": 0.6, "ra": 2.0, "feuchte": 78.0, "ca": 0.3, "p": 0.25, "typ": "Nassfutter"
    },
    "Terra Canis Pute mit Kürbis (Nassfutter)": {
        "rp": 9.2, "rfe": 6.1, "rfa": 0.8, "ra": 1.4, "feuchte": 79.2, "ca": 0.22, "p": 0.16, "typ": "Nassfutter"
    },
    "Premiere Best Meat Rind & Huhn (Nassfutter)": {
        "rp": 11.0, "rfe": 6.0, "rfa": 0.4, "ra": 2.0, "feuchte": 79.0, "ca": 0.25, "p": 0.20, "typ": "Nassfutter"
    },
    "Rocco Classic Rind (Nassfutter)": {
        "rp": 10.8, "rfe": 7.0, "rfa": 0.5, "ra": 2.0, "feuchte": 78.0, "ca": 0.26, "p": 0.21, "typ": "Nassfutter"
    },
    "Pedigree Adult Rind in Soße (Nassfutter)": {
        "rp": 7.5, "rfe": 4.5, "rfa": 0.5, "ra": 2.0, "feuchte": 82.0, "ca": 0.3, "p": 0.22, "typ": "Nassfutter"
    },
    "Anifit Gockel's Duett (Nassfutter)": {
        "rp": 10.8, "rfe": 6.2, "rfa": 0.5, "ra": 2.1, "feuchte": 78.5, "ca": 0.28, "p": 0.22, "typ": "Nassfutter"
    }
}

DEFAULT_FUTTER_DB_KATZE = {
    "Purina Vital Balance Junior Lachs & Karotten (Beutel)": {
        "rp": 12.5, "rfe": 4.0, "rfa": 0.5, "ra": 2.4, "feuchte": 78.0, "ca": 0.30, "p": 0.25, "taurin_add": 580, "vit_a_add": 910, "vit_d_add": 139, "jod_add": 0.39, "zn_add": 18, "konservierung": "Natürliche Tocopherole", "typ": "Nassfutter"
    },
    "Sheba Kitten Feine Vielfalt in Sauce": {
        "rp": 8.0, "rfe": 5.7, "rfa": 0.5, "ra": 1.8, "feuchte": 82.5, "ca": 0.32, "p": 0.25, "taurin_add": 1470, "vit_d_add": 245, "jod_add": 0.34, "zn_add": 15.9, "cu_add": 1.5, "konservierung": "Ohne künstliche Farbstoffe", "typ": "Nassfutter"
    },
    "Mio Junior Menu Filet in Sauce (Huhn & Truthahn)": {
        "rp": 8.2, "rfe": 5.0, "rfa": 0.4, "ra": 1.8, "feuchte": 81.0, "ca": 0.25, "p": 0.20, "taurin_add": 1000, "vit_d_add": 200, "zn_add": 20, "konservierung": "Ohne künstliche Konservierungsmittel", "typ": "Nassfutter"
    },
    "Concept for Life Maine Coon Kitten": {
        "rp": 36.0, "rfe": 23.0, "rfa": 2.8, "ra": 7.0, "feuchte": 10.0, "ca": 1.4, "p": 1.1, "taurin_add": 1500, "vit_a_add": 20000, "jod_add": 1.5, "zn_add": 100, "konservierung": "Natürliche Tocopherole", "typ": "Trockenfutter"
    },
    "Royal Canin Sterilised 37 Trockenfutter": {
        "rp": 37.0, "rfe": 12.0, "rfa": 6.2, "ra": 8.2, "feuchte": 5.5, "ca": 1.1, "p": 1.0, "taurin_add": 1500, "typ": "Trockenfutter"
    },
    "Purina ONE Sterilcat Huhn Trockenfutter": {
        "rp": 37.0, "rfe": 13.0, "rfa": 6.0, "ra": 7.5, "feuchte": 6.0, "ca": 1.2, "p": 1.0, "taurin_add": 1200, "typ": "Trockenfutter"
    },
    "Leonardo Adult Complete Grain-Free Poultry": {
        "rp": 32.0, "rfe": 20.0, "rfa": 1.9, "ra": 7.5, "feuchte": 10.0, "ca": 1.0, "p": 0.9, "taurin_add": 1400, "typ": "Trockenfutter"
    },
    "Smilla Adult Geflügel Trockenfutter": {
        "rp": 32.0, "rfe": 15.0, "rfa": 3.0, "ra": 6.5, "feuchte": 9.0, "ca": 1.1, "p": 0.95, "taurin_add": 1000, "typ": "Trockenfutter"
    },
    "Catz Finefood N°3 Geflügel Nassfutter": {
        "rp": 10.2, "rfe": 5.2, "rfa": 0.4, "ra": 2.0, "feuchte": 80.0, "ca": 0.25, "p": 0.20, "taurin_add": 1500, "typ": "Nassfutter"
    },
    "MAC's Cat Adult Lachs & Hühnchen Nassfutter": {
        "rp": 10.8, "rfe": 6.0, "rfa": 0.4, "ra": 2.4, "feuchte": 80.0, "ca": 0.26, "p": 0.21, "taurin_add": 1500, "typ": "Nassfutter"
    },
    "GranataPet Symphonie N°3 Entenbrust Nassfutter": {
        "rp": 10.5, "rfe": 5.8, "rfa": 0.4, "ra": 2.1, "feuchte": 80.0, "ca": 0.26, "p": 0.21, "taurin_add": 1500, "typ": "Nassfutter"
    },
    "Feringa Duett Geflügel & Kaninchen Nassfutter": {
        "rp": 10.3, "rfe": 5.6, "rfa": 0.4, "ra": 2.2, "feuchte": 80.0, "ca": 0.25, "p": 0.20, "taurin_add": 1500, "typ": "Nassfutter"
    },
    "Felix So gut wie es aussieht Rind in Gelee (Nassfutter)": {
        "rp": 11.5, "rfe": 2.5, "rfa": 0.05, "ra": 2.5, "feuchte": 82.0, "ca": 0.25, "p": 0.20, "taurin_add": 500, "typ": "Nassfutter"
    },
    "Whiskas 1+ Geflügel in Soße (Nassfutter)": {
        "rp": 9.0, "rfe": 4.0, "rfa": 0.3, "ra": 1.8, "feuchte": 83.0, "ca": 0.25, "p": 0.20, "taurin_add": 500, "typ": "Nassfutter"
    },
    "Gourmet Gold Feine Pastete Rind (Nassfutter)": {
        "rp": 10.5, "rfe": 7.0, "rfa": 0.05, "ra": 3.2, "feuchte": 77.5, "ca": 0.30, "p": 0.24, "taurin_add": 700, "typ": "Nassfutter"
    },
    "Sheba Sauce Spéciale Pute (Nassfutter)": {
        "rp": 8.5, "rfe": 4.5, "rfa": 0.3, "ra": 1.9, "feuchte": 82.0, "ca": 0.22, "p": 0.18, "taurin_add": 500, "typ": "Nassfutter"
    }
}


# -----------------------------------------------------------------------------

INGREDIENTS_DB = {
    # Fleisch & Fisch (pro 100g)
    "Rindermuskelfleisch (mager)": {"kcal": 120, "rp": 21.0, "rfe": 4.0, "ca": 10, "p": 190, "jod": 3, "vit_a": 0, "vit_d": 0, "fe": 2.2, "zn": 3.8, "taurin": 36, "kat": "Fleisch", "quelle": "Rind"},
    "Rindermuskelfleisch (fett)": {"kcal": 220, "rp": 17.0, "rfe": 17.0, "ca": 9, "p": 160, "jod": 3, "vit_a": 0, "vit_d": 0, "fe": 1.8, "zn": 3.2, "taurin": 30, "kat": "Fleisch", "quelle": "Rind"},
    "Geflügelbrust (Huhn/Pute)": {"kcal": 105, "rp": 23.0, "rfe": 1.5, "ca": 12, "p": 200, "jod": 4, "vit_a": 0, "vit_d": 0, "fe": 1.2, "zn": 1.0, "taurin": 34, "kat": "Fleisch", "quelle": "Geflügel"},
    "Pferdefleisch (mager)": {"kcal": 110, "rp": 21.5, "rfe": 2.5, "ca": 11, "p": 200, "jod": 2, "vit_a": 0, "vit_d": 0, "fe": 4.2, "zn": 3.5, "taurin": 30, "kat": "Fleisch", "quelle": "Pferd"},
    "Lachsfilet (fettreich)": {"kcal": 200, "rp": 20.0, "rfe": 13.0, "ca": 15, "p": 240, "jod": 30, "vit_a": 40, "vit_d": 900, "fe": 0.8, "zn": 0.9, "taurin": 130, "kat": "Fisch", "quelle": "Fisch"},
    "Kabeljau / Seelachs (mager)": {"kcal": 80, "rp": 18.0, "rfe": 0.8, "ca": 18, "p": 190, "jod": 120, "vit_a": 0, "vit_d": 40, "fe": 0.4, "zn": 0.5, "taurin": 40, "kat": "Fisch", "quelle": "Fisch"},
    # Innereien (pro 100g)
    "Rinderleber": {"kcal": 130, "rp": 19.5, "rfe": 4.5, "ca": 8, "p": 330, "jod": 14, "vit_a": 35000, "vit_d": 50, "fe": 6.5, "zn": 4.0, "taurin": 19, "kat": "Innereien", "quelle": "Rind"},
    "Hühnerleber": {"kcal": 116, "rp": 16.9, "rfe": 4.8, "ca": 11, "p": 240, "jod": 10, "vit_a": 12000, "vit_d": 40, "fe": 7.4, "zn": 3.1, "taurin": 25, "kat": "Innereien", "quelle": "Geflügel"},
    "Hühnerherzen": {"kcal": 155, "rp": 15.5, "rfe": 10.0, "ca": 12, "p": 190, "jod": 5, "vit_a": 30, "vit_d": 0, "fe": 4.5, "zn": 2.5, "taurin": 150, "kat": "Innereien", "quelle": "Geflügel"},
    # Gemüse & Stärke (pro 100g)
    "Zucchini (gedämpft/püriert)": {"kcal": 17, "rp": 1.2, "rfe": 0.2, "ca": 15, "p": 25, "jod": 1, "vit_a": 10, "vit_d": 0, "fe": 0.4, "zn": 0.2, "taurin": 0, "kat": "Basis-Gemüse", "quelle": "Gemüse"},
    "Karotten / Möhren (püriert)": {"kcal": 30, "rp": 1.0, "rfe": 0.2, "ca": 30, "p": 30, "jod": 2, "vit_a": 1700, "vit_d": 0, "fe": 0.5, "zn": 0.3, "taurin": 0, "kat": "Basis-Gemüse", "quelle": "Gemüse"},
    "Fenchel (gedünstet)": {"kcal": 19, "rp": 1.1, "rfe": 0.2, "ca": 40, "p": 35, "jod": 2, "vit_a": 50, "vit_d": 0, "fe": 0.7, "zn": 0.2, "taurin": 0, "kat": "Basis-Gemüse", "quelle": "Gemüse"},
    "Kürbis (gedämpft)": {"kcal": 26, "rp": 1.0, "rfe": 0.1, "ca": 20, "p": 44, "jod": 1, "vit_a": 400, "vit_d": 0, "fe": 0.8, "zn": 0.3, "taurin": 0, "kat": "Basis-Gemüse", "quelle": "Gemüse"},
    "Brokkoli (gekocht)": {"kcal": 35, "rp": 3.0, "rfe": 0.4, "ca": 48, "p": 65, "jod": 2, "vit_a": 80, "vit_d": 0, "fe": 0.8, "zn": 0.5, "taurin": 0, "kat": "Blähgemüse", "quelle": "Gemüse"},
    "Erbsen (gekocht)": {"kcal": 80, "rp": 5.4, "rfe": 0.4, "ca": 25, "p": 110, "jod": 2, "vit_a": 38, "vit_d": 0, "fe": 1.5, "zn": 1.0, "taurin": 0, "kat": "Blähgemüse", "quelle": "Gemüse"},
    "Süßkartoffel (sehr weich gekocht)": {"kcal": 86, "rp": 1.6, "rfe": 0.1, "ca": 30, "p": 47, "jod": 2, "vit_a": 700, "vit_d": 0, "fe": 0.6, "zn": 0.3, "taurin": 0, "kat": "Stärke", "quelle": "Gemüse"},
    "Reis (weich gekocht)": {"kcal": 110, "rp": 2.4, "rfe": 0.3, "ca": 10, "p": 35, "jod": 1, "vit_a": 0, "vit_d": 0, "fe": 0.4, "zn": 0.6, "taurin": 0, "kat": "Stärke", "quelle": "Getreide"},
    "Kartoffeln (gekocht)": {"kcal": 70, "rp": 1.7, "rfe": 0.1, "ca": 10, "p": 50, "jod": 1, "vit_a": 0, "vit_d": 0, "fe": 0.8, "zn": 0.3, "taurin": 0, "kat": "Stärke", "quelle": "Gemüse"},
    # Öle (pro 100g)
    "Lachsöl (Omega-3 Spender)": {"kcal": 900, "rp": 0.0, "rfe": 99.0, "ca": 0, "p": 0, "jod": 0, "vit_a": 0, "vit_d": 500, "fe": 0.0, "zn": 0.0, "taurin": 0, "kat": "Öl", "quelle": "Fisch"},
    "Dorschlebertran (Vit. A & D Spender)": {"kcal": 900, "rp": 0.0, "rfe": 99.0, "ca": 0, "p": 0, "jod": 0, "vit_a": 100000, "vit_d": 10000, "fe": 0.0, "zn": 0.0, "taurin": 0, "kat": "Öl", "quelle": "Fisch"},
    "Kürbiskernöl": {"kcal": 880, "rp": 0.0, "rfe": 99.0, "ca": 0, "p": 0, "jod": 0, "vit_a": 0, "vit_d": 0, "fe": 0.0, "zn": 0.0, "taurin": 0, "kat": "Öl", "quelle": "Pflanzlich"},
    "Weizenkeimöl (Vit E Spender)": {"kcal": 890, "rp": 0.0, "rfe": 99.0, "ca": 0, "p": 0, "jod": 0, "vit_a": 0, "vit_d": 0, "fe": 0.0, "zn": 0.0, "taurin": 0, "kat": "Öl", "quelle": "Getreide"},
    # Supplements (pro 100g / pro Einheit)
    "Eierschalenpulver (38% Ca)": {"kcal": 0, "rp": 0.0, "rfe": 0.0, "ca": 38000, "p": 800, "jod": 0, "vit_a": 0, "vit_d": 0, "fe": 0.0, "zn": 0.0, "taurin": 0, "kat": "Supplement"},
    "Algenkalk (30% Ca)": {"kcal": 0, "rp": 0.0, "rfe": 0.0, "ca": 30000, "p": 600, "jod": 400, "vit_a": 0, "vit_d": 0, "fe": 0.0, "zn": 0.0, "taurin": 0, "kat": "Supplement"},
    "Seealgenmehl (Jodspender)": {"kcal": 0, "rp": 5.0, "rfe": 1.0, "ca": 1000, "p": 200, "jod": 60000, "vit_a": 0, "vit_d": 0, "fe": 10.0, "zn": 2.0, "taurin": 0, "kat": "Supplement"},
    "Zink-Chelat (Spurenelement)": {"kcal": 0, "rp": 0.0, "rfe": 0.0, "ca": 0, "p": 0, "jod": 0, "vit_a": 0, "vit_d": 0, "fe": 0.0, "zn": 100000.0, "taurin": 0, "kat": "Supplement"},
    "Taurin (Reinsubstanz 100%)": {"kcal": 0, "rp": 100.0, "rfe": 0.0, "ca": 0, "p": 0, "jod": 0, "vit_a": 0, "vit_d": 0, "fe": 0.0, "zn": 0.0, "taurin": 100000.0, "kat": "Supplement"},
    "Cosequin DS (Knorpelschutz / Stk)": {"kcal": 5, "rp": 0.5, "rfe": 0.1, "ca": 50, "p": 30, "jod": 0, "vit_a": 0, "vit_d": 0, "fe": 0.0, "zn": 0.0, "taurin": 0, "kat": "Supplement"},
    "Vitamin-Optimix Senior (All-in-One)": {"kcal": 50, "rp": 10.0, "rfe": 1.0, "ca": 16000, "p": 4000, "jod": 15000, "vit_a": 250000, "vit_d": 25000, "fe": 500.0, "zn": 1200.0, "taurin": 0, "kat": "Supplement"}
}

