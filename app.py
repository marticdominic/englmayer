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
    "Automatska kontrola troškova po paletama i dizel dodatka prema ponudi G. Englmayer (OF 002/2026)."
)

# Englmayer cjenik paleta po zonama (EUR bez PDV-a) prema ponudi
englmayer_cjenik = {
    "Zona 1": {300: 23.00, 400: 25.00, 500: 30.00, 600: 33.00, 700: 38.00},
    "Zona 2": {300: 26.00, 400: 29.00, 500: 35.00, 600: 39.00, 700: 43.00},
    "Zona 3": {300: 36.00, 400: 40.00, 500: 45.00, 600: 47.00, 700: 51.00},
    "Zona 4": {300: 42.00, 400: 47.00, 500: 50.00, 600: 55.00, 700: 65.00},
    "Zona 5": {300: 44.00, 400: 48.00, 500: 51.00, 600: 56.00, 700: 68.00},
    "Zona 6": {300: 55.00, 400: 59.00, 500: 63.00, 600: 65.00, 700: 79.00},
}

englmayer_otoci = {
    "južni otoci": {300: 89.00, 400: 100.00, 500: 110.00, 600: 121.00},
    "sjeverni otoci": {300: 79.00, 400: 89.00, 500: 100.00, 600: 110.00}
}

juzni_otoci = ["brač", "hvar", "korčula", "pelješac", "supetar", "bol", "stari grad", "jelsa", "vela luka", "ugljan"]
sjeverni_otoci = ["cres", "lošinj", "krk", "rab", "pag", "mali lošinj", "baška", "malinska", "punat", "novalja"]

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

def odredi_zonu(primatelj):
    p_lc = str(primatelj).lower()
    
    is_otok_juzni = any(o in p_lc for o in juzni_otoci)
    is_otok_sjeverni = any(o in p_lc for o in sjeverni_otoci)
    
    if is_otok_juzni:
        return "Otok - Južni"
    elif is_otok_sjeverni:
        return "Otok - Sjeverni"
    else:
        if any(g in p_lc for g in ["zagreb", "velika gorica", "samobor", "zaprešić"]):
            return "Zona 1"
        elif any(g in p_lc for g in ["split", "zadar", "dubrovnik", "makarska", "ploče", "imotski", "šibenik", "knin"]):
            return "Zona 6" if any(x in p_lc for x in ["makarska", "ploče", "imotski"]) else "Zona 2"
        elif any(g in p_lc for g in ["rijeka", "pula", "opatija", "poreč", "rovinj", "kukuljanovo"]):
            return "Zona 2"
        elif any(g in p_lc for g in ["osijek", "vukovar", "vinkovci", "slavonski brod", "županja"]):
            return "Zona 3"
        else:
            return "Zona 2"

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
        ukupna_masa = 0.0
        naplaceni_transport = 0.0
        naplaceni_dizel = 0.0
        broj_paleta = 0
        
        palete_mase = []
        
        for linija in linije:
            if "Broj pošiljke:" in linija:
                broj_posiljke = linija.split("Broj pošiljke:")[-1].strip()
            elif "Datum naloga:" in linija:
                datum_naloga = linija.split("Datum naloga:")[-1].strip()
            elif "Datum isporuke:" in linija:
                datum_isporuke = linija.split("Datum isporuke:")[-1].strip()
            elif "Primatelj:" in linija:
                primatelj = linija.split("Primatelj:")[-1].strip()
            elif "Ukupna masa pošiljke:" in linija:
                m_match = re.search(r"([\d\.]+)\s*kg", linija)
                if m_match:
                    ukupna_masa = float(m_match.group(1).replace(".", "").replace(",", "."))
            elif "Cijena prijevoza:" in linija:
                c_match = re.search(r"([\d\.,]+)\s*EUR", linija)
                if c_match:
                    naplaceni_transport = float(c_match.group(1).replace(".", "").replace(",", "."))
            elif "Dizel dodatak:" in linija:
                d_match = re.search(r"([\d\.,]+)\s*EUR", linija)
                if d_match:
                    naplaceni_dizel = float(d_match.group(1).replace(".", "").replace(",", "."))
            elif "Paleta " in linija and " od " in linija:
                broj_paleta += 1
                match_m = re.search(r"Masa:\s*([\d\.,]+)\s*kg", linija)
                if match_m:
                    p_masa = float(match_m.group(1).replace(".", "").replace(",", "."))
                    palete_mase.append(p_masa)

        # Uklanjanje praznih redova (ako nema broja pošiljke ili primatelja)
        if not broj_posiljke or not primatelj:
            continue

        if broj_paleta == 0:
            broj_paleta = 1

        if palete_mase:
            zbroj_masa_paleta = sum(palete_mase)
            if zbroj_masa_paleta > 0:
                ukupna_masa = zbroj_masa_paleta

        # Određivanje zone
        zona = odredi_zonu(primatelj)

        # Izračun ugovorene osnove po cjeniku (zbrajanje cijena pojedinačnih paleta)
        ugovorena_osnova = 0.0
        if palete_mase:
            for p_masa in palete_mase:
                ugovorena_osnova += dohvati_cijenu_palete(zona, p_masa)
        else:
            prosjek_mase = ukupna_masa / broj_paleta if broj_paleta > 0 else ukupna_masa
            ugovorena_osnova = dohvati_cijenu_palete(zona, prosjek_mase) * broj_paleta

        parsed_data.append({
            "Broj pošiljke": broj_posiljke,
            "Datum naloga": datum_naloga,
            "Datum isporuke": datum_isporuke,
            "Primatelj": primatelj,
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
            ugovorene_liste = []
            for _, r in df.iterrows():
                z = odredi_zonu(r.get("Primatelj", ""))
                zone_list.append(z)
                masa = float(r.get("Masa (kg)", 0))
                br_p = int(r.get("Broj paleta", 1))
                if br_p < 1: br_p = 1
                ug_c = dohvati_cijenu_palete(z, masa / br_p) * br_p
                ugovorene_liste.append(round(ug_c, 2))
            df["Zona"] = zone_list
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
                    "📥 Preuzmi izvještaj preplata (Excel)",
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
