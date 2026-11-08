import io
import math
import os
import re
import pandas as pd
import streamlit as st

# Za čitanje Word (.docx) datoteka
try:
    import docx
except ImportError:
    os.system("pip install python-docx")
    import docx

# Importi za generiranje PDF-ova i registraciju fonta
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(
    page_title="Kontrola Englmayer Računa", page_icon="📦", layout="wide"
)

# Registracija fontova za PDF
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

try:
    if os.path.exists("DejaVuSans.ttf") and os.path.exists("DejaVuSans-Bold.ttf"):
        pdfmetrics.registerFont(TTFont("DejaVuSans", "DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", "DejaVuSans-Bold.ttf"))
        FONT_REGULAR = "DejaVuSans"
        FONT_BOLD = "DejaVuSans-Bold"
except Exception as e:
    pass

st.title("📦 Sustav za Kontrolu Englmayer Računa (Word & Excel Specifikacije)")
st.write(
    "Automatska kontrola troškova po paletama, poštanskim brojevima, van gabaritnim paletama (+50%) i dizel dodatku prema ponudi G. Englmayer (OF 002/2026)."
)

# Englmayer cjenik paleta po zonama (EUR bez PDV-a) iz ponude[cite: 6]
englmayer_cjenik = {
    "Zona 1": {300: 23.00, 400: 25.00, 500: 30.00, 600: 33.00, 700: 38.00},
    "Zona 2": {300: 26.00, 400: 29.00, 500: 35.00, 600: 39.00, 700: 43.00},
    "Zona 3": {300: 36.00, 400: 40.00, 500: 45.00, 600: 47.00, 700: 51.00},
    "Zona 4": {300: 42.00, 400: 47.00, 500: 50.00, 600: 55.00, 700: 65.00},
    "Zona 5": {300: 44.00, 400: 48.00, 500: 51.00, 600: 56.00, 700: 68.00},
    "Zona 6": {300: 55.00, 400: 59.00, 500: 63.00, 600: 65.00, 700: 79.00},
}

englmayer_otoci = {
    "južni otoci": {300: 89.00, 400: 100.00, 500: 110.00, 600: 121.00}, # Brač, Hvar, Korčula, Pelješac[cite: 7]
    "sjeverni otoci": {300: 79.00, 400: 89.00, 500: 100.00, 600: 110.00} # Cres, Lošinj, Krk, Rab, Pag[cite: 7]
}

# Strogi popis otočnih poštanskih brojeva i ključnih riječi isključivo za otoke
otocni_zipovi = [
    21400, 21403, 21404, 21405, 21410, 21412, 21413, 21420, 21423, 21424, 21425, 21426, 21450, 21454, 21460, # Brač, Hvar, Vis...
    20230, 20231, 20232, 20233, 20234, 20235, 20236, 20240, 20242, 20243, 20244, 20245, 20246, 20247, 20248, 20250, 20260, 20263, 20264, 20267, 20269, 20270, 20271, 20272, 20273, 20274, 20275, # Korčula, Mljet, Lastovo, Pelješac
    51500, 51511, 51512, 51513, 51514, 51515, 51516, 51517, 51521, 51522, 51523, 51550, 51551, 51554, 51555, 51556, 51557, 51559, 51564, # Krk, Cres, Lošinj, Rab
    23212, 23234, 23249, 23250, 23251, 23262, 23263, 23264, 23271, 23272, 23273, 23274, 23275 # Pag, Ugljan, Pašman, Dugi otok
]

def dohvati_cijenu_palete(zona, masa_palete):
    if "Otok - Južni" in zona:
        tablica = englmayer_otoci["južni otoci"]
        granice = [300, 400, 500, 600]
    elif "Otok - Sjeverni" in zona:
        tablica = englmayer_otoci["sjeverni otoci"]
        granice = [300, 400, 500, 600]
    else:
        tablica = englmayer_cjenik.get(zona, englmayer_cjenik["Zona 2"])
        granice = [300, 400, 500, 600, 700]

    for g in granice:
        if masa_palete <= g:
            return tablica[g]
            
    return tablica[granice[-1]]

def odredi_zonu_i_pbr(primatelj):
    p_str = str(primatelj)
    p_lc = p_str.lower()
    
    # 1. Izvlačenje poštanskog broja (5 znamenki)
    pbr_match = re.search(r'\b(\d{5})\b', p_str)
    pbr = int(pbr_match.group(1)) if pbr_match else 0
    pbr_prefix = int(str(pbr)[:2]) if pbr > 0 else 0

    # 2. Provjera je li poštanski broj izričito otočni
    if pbr in otocni_zipovi:
        # Razlikujemo sjeverne i južne otoke prema prefiksu ZIP-a
        if str(pbr).startswith("51") or str(pbr).startswith("23"):
            return "Otok - Sjeverni", pbr
        else:
            return "Otok - Južni", pbr

    # 3. Mjesta koja prema cjeniku gravitiraju Makarskoj, Imotskom i Pločama idu u Zonu 6[cite: 6]
    if any(g in p_lc for g in ["makarska", "imotski", "ploče", "metković"]):
        return "Zona 6", pbr

    # 4. Strogo određivanje zone prema rasponu poštanskog broja (Englmayer cjenik)[cite: 6]:
    if pbr_prefix == 10:
        return "Zona 1", pbr  # 10xxx - Zagreb i okolica[cite: 6]
    elif 40 <= pbr_prefix <= 49:
        return "Zona 2", pbr  # 40xxx - 49xxx[cite: 6]
    elif 34 <= pbr_prefix <= 35:
        return "Zona 3", pbr  # 34xxx - 35xxx (Slavonska Požega, Slavonski Brod...)[cite: 6]
    elif pbr_prefix == 51:
        return "Zona 4", pbr  # 51xxx - Rijeka i primorje[cite: 6]
    elif (31 <= pbr_prefix <= 33) or (pbr_prefix == 52):
        return "Zona 5", pbr  # 31-33xxx (Osijek, Vinkovci...), 52xxx (Pula, Istra)[cite: 6]
    elif pbr_prefix == 20:
        return "Zona 6", pbr  # 20xxx - Dubrovnik[cite: 6]
    elif 42 <= pbr_prefix <= 43 or 47 <= pbr_prefix <= 48:
        return "Zona 2", pbr  # 47xxx - Karlovac je Zona 2
        
    # Fallback po nazivima gradova ako ZIP nije pronađen
    if any(g in p_lc for g in ["zagreb", "velika gorica", "samobor", "zaprešić"]):
        return "Zona 1", pbr
    elif any(g in p_lc for g in ["osijek", "vukovar", "vinkovci", "slavonski brod", "županja", "požega", "đakovo"]):
        return "Zona 5", pbr
    elif any(g in p_lc for g in ["rijeka", "pula", "opatija", "poreč", "rovinj", "kukuljanovo"]):
        return "Zona 4", pbr
    elif any(g in p_lc for g in ["karlovac", "zadar", "split", "šibenik", "knin"]):
        return "Zona 2", pbr
    else:
        return "Zona 2", pbr

# Formula za dizel: do 1,46 EUR = 0%, na svakih daljnjih 0,07 EUR = 1%[cite: 3]
fn_dizel = lambda c: 0.0 if c <= 1.46 else max(0.0, math.ceil((c - 1.46) / 0.07))

def parsiraj_word_specifikaciju_englmayer(doc_file):
    doc = docx.Document(doc_file)
    tekst = "\n".join([p.text for p in doc.paragraphs])
    
    blokovi = tekst.split("Pošiljka ")
    parsed_data = []
    
    for blok in blokovi[1:]:
        linije = blok.strip().split("\n")
        if not linije:
            continue
            
        broj_posiljke = ""
        datum_naloga = ""
        datum_isporuke = ""
        primatelj = ""
        naplaceni_transport = 0.0
        naplaceni_dizel = 0.0
        
        palete_podaci = []
        
        for linija in linije:
            if "Broj pošiljke:" in linija:
                broj_posiljke = linija.split("Broj pošiljke:")[-1].strip()
            elif "Datum naloga:" in linija:
                datum_naloga = linija.split("Datum naloga:")[-1].strip()
            elif "Datum isporuke:" in linija:
                datum_isporuke = linija.split("Datum isporuke:")[-1].strip()
            elif "Primatelj:" in linija:
                primatelj = linija.split("Primatelj:")[-1].strip()
            elif "Cijena prijevoza:" in linija:
                c_match = re.search(r"([\d\.,]+)\s*EUR", linija)
                if c_match:
                    naplaceni_transport = float(c_match.group(1).replace(".", "").replace(",", "."))
            elif "Dizel dodatak:" in linija:
                d_match = re.search(r"([\d\.,]+)\s*EUR", linija)
                if d_match:
                    naplaceni_dizel = float(d_match.group(1).replace(".", "").replace(",", "."))
            
            elif "Masa:" in linija:
                je_vangabaritna = "EWP1" in linija.upper() or "VANGABARIT" in linija.upper()
                m_masa_match = re.search(r"Masa:\s*([\d\.,]+)\s*kg", linija, re.IGNORECASE)
                if m_masa_match:
                    val_str = m_masa_match.group(1).replace(".", "").replace(",", ".")
                    try:
                        val_f = float(val_str)
                        palete_podaci.append((val_f, je_vangabaritna))
                    except:
                        pass

        if not broj_posiljke or not primatelj:
            continue

        broj_paleta = len(palete_podaci)
        if broj_paleta == 0:
            broj_paleta = 1
            palete_podaci = [(0.0, False)]

        ukupna_masa = sum([p[0] for p in palete_podaci])
        
        zona, pbr = odredi_zonu_i_pbr(primatelj)

        ugovorena_osnova = 0.0
        for p_masa, je_vg in palete_podaci:
            osnovna_c = dohvati_cijenu_palete(zona, p_masa)
            if je_vg:
                osnovna_c *= 1.50  # +50% doplate za van gabaritne palete[cite: 6]
            ugovorena_osnova += osnovna_c

        parsed_data.append({
            "Broj pošiljke": broj_posiljke,
            "Datum naloga": datum_naloga,
            "Datum isporuke": datum_isporuke,
            "Primatelj": primatelj,
            "Poštanski broj (ZIP)": pbr if pbr > 0 else "",
            "Broj paleta": broj_paleta,
            "Masa (kg)": ukupna_masa,
            "Zona": zona,
            "Naplaćeni Transport (€)": naplaceni_transport,
            "Ugovorena Osnova (€)": round(ugovorena_osnova, 2),
            "Naplaćeni Dizel (€)": naplaceni_dizel,
            "Razlika Transport (€)": round(naplaceni_transport - ugovorena_osnova, 2)
        })
        
    return pd.DataFrame(parsed_data)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Izvjestaj")
    return output.getvalue()

st.sidebar.header("Parametri obračuna")
cijena_dizela_input = st.sidebar.number_input(
    "Prosječna cijena dizela (€ bez PDV-a):", min_value=1.00, max_value=3.00, value=2.02, step=0.01
)
posto_goriva = fn_dizel(cijena_dizela_input)
st.sidebar.info(f"Izračunati dizel dodatak po ugovoru: **{posto_goriva:.1f}%**")

uploaded_file = st.file_uploader(
    "Učitaj Englmayer specifikaciju (Podržava Word .docx i Excel/CSV)", type=["docx", "xlsx", "csv"]
)

if uploaded_file is not None:
    file_name = uploaded_file.name.lower()
    
    if file_name.endswith(".docx"):
        df = parsiraj_word_specifikaciju_englmayer(uploaded_file)
        st.success(f"Word specifikacija uspješno učitana! Pronađeno pošiljaka: {len(df)}")
    elif file_name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
        if "Broj pošiljke" in df.columns:
            df = df.dropna(subset=["Broj pošiljke"])
        st.success("Excel tablica uspješno učitana!")
    else:
        df = pd.read_csv(uploaded_file)
        if "Broj pošiljke" in df.columns:
            df = df.dropna(subset=["Broj pošiljke"])
        st.success("CSV tablica uspješno učitana!")

    if not df.empty:
        if "Ugovorena Osnova (€)" not in df.columns and "Masa (kg)" in df.columns:
            zone_list = []
            pbr_list = []
            ugovorene_liste = []
            for _, r in df.iterrows():
                z, pbr = odredi_zonu_i_pbr(r.get("Primatelj", ""))
                zone_list.append(z)
                pbr_list.append(pbr if pbr > 0 else "")
                masa = float(r.get("Masa (kg)", 0))
                br_p = int(r.get("Broj paleta", 1))
                if br_p < 1: br_p = 1
                ug_c = dohvati_cijenu_palete(z, masa / br_p) * br_p
                ugovorene_liste.append(round(ug_c, 2))
            df["Zona"] = zone_list
            df["Poštanski broj (ZIP)"] = pbr_list
            df["Ugovorena Osnova (€)"] = ugovorene_liste
            df["Ugovoreni Dizel (€)"] = df["Ugovorena Osnova (€)"] * (posto_goriva / 100.0)
            df["Razlika Transport (€)"] = df["Naplaćeni Transport (€)"] - df["Ugovorena Osnova (€)"]

        tab1, tab2, tab3 = st.tabs(["📊 Pregled i Usporedba", "🚨 Preplate / Odstupanja", "📈 Zbirne Sume"])
        
        with tab1:
            st.subheader("Detaljna tablica usporedbe naplaćenog i ugovorenog")
            st.dataframe(df, use_container_width=True)
            st.download_button(
                "📥 Preuzmi izvještaj (Excel)",
                to_excel(df),
                "englmayer_kontrola.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with tab2:
            st.subheader("Pošiljke s razlikom u cijeni (Naplaćeno > Ugovoreno)")
            if "Razlika Transport (€)" in df.columns:
                sumnjive = df[df["Razlika Transport (€)"] > 0.05]
                st.dataframe(sumnjive, use_container_width=True)
                st.download_button(
                    "📥 Preuzmi izvještaj preplate (Excel)",
                    to_excel(sumnjive),
                    "englmayer_preplate.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Nema izračunate razlike za prikaz.")
                
        with tab3:
            st.subheader("Zbirni financijski pregled")
            if "Naplaćeni Transport (€)" in df.columns:
                uk_naplaceno = df["Naplaćeni Transport (€)"].sum()
                uk_ugovoreno = df["Ugovorena Osnova (€)"].sum()
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Ukupno naplaćeni transport", f"{uk_naplaceno:,.2f} €")
                col2.metric("Ukupno ugovorena osnova", f"{uk_ugovoreno:,.2f} €")
                col3.metric("Ukupna razlika / Višak", f"{uk_naplaceno - uk_ugovoreno:,.2f} €")
