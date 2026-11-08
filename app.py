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
    page_title="Kontrola Englmayer Računa (Word & Excel)", page_icon="📦", layout="wide"
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

st.title("📦 Sustav za Kontrolu Englmayer Računa (Word Specifikacije & Cjenik)")
st.write(
    "Automatska kontrola troškova prijevoza i dizel dodatka prema ponudi G. Englmayer (OF 002/2026) s podrškom za Word (.docx) specifikacije."
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
    "južni otoci": {300: 89.00, 400: 100.00, 500: 110.00, 600: 121.00}, # Brač, Hvar, Korčula, Pelješac
    "sjeverni otoci": {300: 79.00, 400: 89.00, 500: 100.00, 600: 110.00} # Cres, Lošinj, Krk, Rab, Pag
}

juzni_otoci = ["brač", "hvar", "korčula", "pelješac", "supetar", "bol", "stari grad", "jelsa", "vela luka", "korčula", "ugljan"]
sjeverni_otoci = ["cres", "lošinj", "krk", "rab", "pag", "mali lošinj", "baška", "malinska", "punat", "novalja"]

def odredi_zonu_i_cijenu(primatelj, masa):
    p_lc = str(primatelj).lower()
    
    # Provjera otoka
    is_otok_juzni = any(o in p_lc for o in juzni_otoci)
    is_otok_sjeverni = any(o in p_lc for o in sjeverni_otoci)
    
    if is_otok_juzni:
        tablica = englmayer_otoci["južni otoci"]
        zona_naziv = "Otok - Južni"
    elif is_otok_sjeverni:
        tablica = englmayer_otoci["sjeverni otoci"]
        zona_naziv = "Otok - Sjeverni"
    else:
        # Automatska zona prema regiji u nazivu primatelja (npr. Split, Zadar, Osijek, Čakovec, Zagreb...)
        if any(g in p_lc for g in ["zagreb", "velika gorica", "samobor", "zaprešić"]):
            zona_naziv = "Zona 1"
        elif any(g in p_lc for g in ["split", "zadar", "dubrovnik", "makarska", "ploče", "imotski", "šibenik", "knin"]):
            zona_naziv = "Zona 6" if any(x in p_lc for x in ["makarska", "ploče", "imotski"]) else "Zona 2"
        elif any(g in p_lc for g in ["rijeka", "pula", "opatija", "poreč", "rovinj", "kukuljanovo"]):
            zona_naziv = "Zona 2"
        elif any(g in p_lc for g in ["osijek", "vukovar", "vinkovci", "slavonski brod", "županja"]):
            zona_naziv = "Zona 3"
        else:
            zona_naziv = "Zona 2" # Default
            
        tablica = englmayer_cjenik.get(zona_naziv, englmayer_cjenik["Zona 2"])

    granice = [300, 400, 500, 600, 700] if "Otok" not in zona_naziv else [300, 400, 500, 600]
    
    for g in granice:
        if masa <= g:
            return zona_naziv, tablica[g]
            
    max_g = granice[-1]
    return zona_naziv, tablica[max_g]

# Formula za dizel: do 1,46 EUR = 0%, na svakih daljnjih 0,07 EUR = 1%
fn_dizel = lambda c: 0.0 if c <= 1.46 else max(0.0, math.ceil((c - 1.46) / 0.07))

def parsiraj_word_specifikaciju(doc_file):
    doc = docx.Document(doc_file)
    tekst = "\n".join([p.text for p in doc.paragraphs])
    
    # Detekcija pošiljaka iz Word formata (prema obliku u vašem primjeru)
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
        masa_kg = 0.0
        naplaceni_transport = 0.0
        naplaceni_dizel = 0.0
        
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
                    masa_kg = float(m_match.group(1).replace(".", "").replace(",", "."))
            elif "Cijena prijevoza:" in linija:
                c_match = re.search(r"([\d\.,]+)\s*EUR", linija)
                if c_match:
                    naplaceni_transport = float(c_match.group(1).replace(".", "").replace(",", "."))
            elif "Dizel dodatak:" in linija:
                d_match = re.search(r"([\d\.,]+)\s*EUR", linija)
                if d_match:
                    naplaceni_dizel = float(d_match.group(1).replace(".", "").replace(",", "."))
                    
        zona, ugovorena_cijena = odredi_zonu_i_cijenu(primatelj, masa_kg)
        
        parsed_data.append({
            "Broj pošiljke": broj_posiljke,
            "Datum naloga": datum_naloga,
            "Datum isporuke": datum_isporuke,
            "Primatelj": primatelj,
            "Masa (kg)": masa_kg,
            "Zona": zona,
            "Naplaćeni Transport (€)": naplaceni_transport,
            "Ugovorena Osnova (€)": ugovorena_cijena,
            "Naplaćeni Dizel (€)": naplaceni_dizel,
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
        df = parsiraj_word_specifikaciju(uploaded_file)
        st.success(f"Word specifikacija uspješno učitana! Pronađeno pošiljaka: {len(df)}")
    elif file_name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
        st.success("Excel tablica uspješno učitana!")
    else:
        df = pd.read_csv(uploaded_file)
        st.success("CSV tablica uspješno učitana!")

    if not df.empty:
        if "Ugovorena Osnova (€)" not in df.columns and "Masa (kg)" in df.columns:
            # Izračun ugovorenih osnova ako već nisu tu
            zone_list = []
            ugovorene_liste = []
            for _, r in df.iterrows():
                z, uc = odredi_zonu_i_cijenu(r.get("Primatelj", ""), r.get("Masa (kg)", 0))
                zone_list.append(z)
                ugovorene_liste.append(uc)
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
