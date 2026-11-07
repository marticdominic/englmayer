import io
import math
import os
import re
import pandas as pd
import streamlit as st
from pypdf import PdfReader

# ReportLab importi za PDF izvještaje
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(
    page_title="Kontrola Englmayer Logističkih Računa", page_icon="🚛", layout="wide"
)

# Sigurna registracija fontova
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

try:
    if os.path.exists("DejaVuSans.ttf") and os.path.exists("DejaVuSans-Bold.ttf"):
        pdfmetrics.registerFont(TTFont("DejaVuSans", "DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", "DejaVuSans-Bold.ttf"))
        FONT_REGULAR = "DejaVuSans"
        FONT_BOLD = "DejaVuSans-Bold"
except Exception:
    pass

st.title("🚛 Sustav za Kontrolu i Analizu Englmayer Računa")
st.write(
    "Automatska ekstrakcija i kontrola troškova iz Englmayer PDF obračuna prema ugovornoj ponudi."
)

# --- UGOVORNI CJENIK ENGLMAYER (Zona, Težina -> Cijena u EUR) ---
CJENIK_ENGLMAYER = {
    "Zona 1": {300: 23.00, 400: 25.00, 500: 30.00, 600: 33.00, 700: 38.00},
    "Zona 2": {300: 26.00, 400: 29.00, 500: 35.00, 600: 39.00, 700: 43.00},
    "Zona 3": {300: 36.00, 400: 40.00, 500: 45.00, 600: 47.00, 700: 51.00},
    "Zona 4": {300: 42.00, 400: 47.00, 500: 50.00, 600: 55.00, 700: 65.00},
    "Zona 5": {300: 44.00, 400: 48.00, 500: 51.00, 600: 56.00, 700: 68.00},
    "Zona 6": {300: 55.00, 400: 59.00, 500: 63.00, 600: 65.00, 700: 79.00},
    "Otoci Kvarner": {300: 79.00, 400: 89.00, 500: 100.00, 600: 110.00},
    "Otoci Dalmacija": {300: 89.00, 400: 100.00, 500: 110.00, 600: 121.00},
}

def odredi_zonu_englmayer(grad):
    grad_clean = str(grad).upper().strip()
    otoci_kvarner = ["CRES", "LOŠINJ", "KRK", "RAB", "PAG"]
    otoci_dalmacija = ["BRAČ", "HVAR", "KORČULA", "PELJEŠAC", "UGLJAN", "NOVALJA"]
    zona_6_mjesta = ["MAKARSKA", "IMOTSKI", "PLOČE"]
    
    for otok in otoci_kvarner:
        if otok in grad_clean:
            return "Otoci Kvarner"
    for otok in otoci_dalmacija:
        if otok in grad_clean:
            return "Otoci Dalmacija"
    for z6 in zona_6_mjesta:
        if z6 in grad_clean:
            return "Zona 6"
            
    if grad_clean in ["ZAGREB", "VELIKA GORICA", "ZABOK", "KARLOVAC", "DUGO SELO"]:
        return "Zona 1"
    
    return "Zona 2"

def izracunaj_osnovnu_cijenu_palete(masa, zona):
    granice = [300, 400, 500, 600, 700]
    z_tablica = CJENIK_ENGLMAYER.get(zona, CJENIK_ENGLMAYER["Zona 2"])
    
    for g in granice:
        if masa <= g and g in z_tablica:
            return z_tablica[g]
    max_g = max([k for k in z_tablica.keys()])
    return z_tablica[max_g]

def izracunaj_dizel_dodatak_englmayer(cijena_goriva):
    if cijena_goriva <= 1.46:
        return 0.0
    razlika = cijena_goriva - 1.46
    postotak = (math.floor(razlika / 0.07) + 1) * 1.0
    return float(postotak)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Izvjestaj")
    return output.getvalue()


# --- PARSIRANJE ENGLMAYER PDF SPECIFIKACIJE SA SCANNED PDF FALLBACKOM ---
def parsiraj_englmayer_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    tekst_sveukupno = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            tekst_sveukupno += t + "\n"

    # Ako je PDF skeniran (prazan tekst)
    if not tekst_sveukupno.strip():
        st.warning("⚠️ Učitani PDF je skenirani dokument (slika) pa `pypdf` ne može izravno pročitati tekst.")
        st.info("💡 **Demo način:** Prikazujem probni set podataka pošiljaka kako bi mogao testirati sve tablice, izračune preplata i Excel izvoz!")
        
        demo_data = [
            {"LA-ID": "EP-2026-101", "Destinacija": "Split", "Masa (kg)": 450.0, "Naplaćeni Transport (€)": 50.00, "Naplaćeni Dizel Dodatak (€)": 5.00},
            {"LA-ID": "EP-2026-102", "Destinacija": "Zadar", "Masa (kg)": 280.0, "Naplaćeni Transport (€)": 29.00, "Naplaćeni Dizel Dodatak (€)": 3.00},
            {"LA-ID": "EP-2026-103", "Destinacija": "Rijeka", "Masa (kg)": 620.0, "Naplaćeni Transport (€)": 55.00, "Naplaćeni Dizel Dodatak (€)": 6.50},
            {"LA-ID": "EP-2026-104", "Destinacija": "Zagreb", "Masa (kg)": 150.0, "Naplaćeni Transport (€)": 23.00, "Naplaćeni Dizel Dodatak (€)": 2.50},
            {"LA-ID": "EP-2026-105", "Destinacija": "Osijek", "Masa (kg)": 510.0, "Naplaćeni Transport (€)": 51.00, "Naplaćeni Dizel Dodatak (€)": 6.00},
            {"LA-ID": "EP-2026-106", "Destinacija": "Makarska", "Masa (kg)": 380.0, "Naplaćeni Transport (€)": 65.00, "Naplaćeni Dizel Dodatak (€)": 7.50},
            {"LA-ID": "EP-2026-107", "Destinacija": "Krk", "Masa (kg)": 290.0, "Naplaćeni Transport (€)": 95.00, "Naplaćeni Dizel Dodatak (€)": 10.00},
        ]
        return pd.DataFrame(demo_data)

    posiljke = []  
    blenderi = re.split(r"(LA-ID:\s*EP-\d+|Broj:\s*\d+|EP-\d+)", tekst_sveukupno, flags=re.IGNORECASE)
    
    if len(blenderi) <= 1:
        st.warning("⚠️ PDF nema standardne 'LA-ID' oznake. Provjeri strukturu sirovog teksta ispod.")
        with st.expander("🔍 Debug: Sirovi tekst iz PDF-a (prvih 3000 znakova)", expanded=True):
            st.text(tekst_sveukupno[:3000])
        return pd.DataFrame()

    for i in range(1, len(blenderi), 2):
        la_id = blenderi[i].strip()
        sadrzaj_bloka = blenderi[i+1] if (i+1) < len(blenderi) else ""
        
        match_grad = re.search(r"(?:DPU|Mjesto istovara|Primatelj|Grad)\s*[-:]?\s*([^\n]+)", sadrzaj_bloka, re.IGNORECASE)
        grad = match_grad.group(1).strip() if match_grad else "Nepoznato"
        
        match_masa = re.findall(r"(\d+[\.,]\d{2})\s*(?:kg|KG)", sadrzaj_bloka)
        ukupna_masa = sum([float(m.replace(".", "").replace(",", ".")) for m in match_masa]) if match_masa else 200.0
        if ukupna_masa == 0:
            ukupna_masa = 200.0
            
        cijena_transporta = 0.0
        cijena_goriva_stavka = 0.0
        
        linije_bloka = sadrzaj_bloka.split("\n")
        for idx_l, linija in enumerate(linije_bloka):
            if any(kljuc in linija.lower() for kljuc in ["roba", "transport", "osnov", "cijena", "110"]):
                for naknadna in linije_bloka[idx_l:idx_l+3]:
                    m_izn = re.search(r"(\d+[\.,]\d{2})", naknadna)
                    if m_izn and cijena_transporta == 0.0:
                        val_pot = float(m_izn.group(1).replace(".", "").replace(",", "."))
                        if val_pot < 1000:
                            cijena_transporta = val_pot
            if any(kljuc in linija.lower() for kljuc in ["dizel", "goriv", "dodatak", "197", "199"]):
                for naknadna in linije_bloka[idx_l:idx_l+3]:
                    m_izn = re.search(r"(\d+[\.,]\d{2})", naknadna)
                    if m_izn and cijena_goriva_stavka == 0.0:
                        val_pot = float(m_izn.group(1).replace(".", "").replace(",", "."))
                        if val_pot < 200:
                            cijena_goriva_stavka = val_pot

        posiljke.append({
            "LA-ID": la_id,
            "Destinacija": grad,
            "Masa (kg)": ukupna_masa,
            "Naplaćeni Transport (€)": cijena_transporta if cijena_transporta > 0 else 45.00,
            "Naplaćeni Dizel Dodatak (€)": cijena_goriva_stavka
        })

    return pd.DataFrame(posiljke)


# --- SIDEBAR PARAMETRI ---
st.sidebar.header("Parametri obračuna Englmayer")
prosjecna_cijena_goriva = st.sidebar.number_input(
    "Prosječna cijena dizel goriva (€ bez PDV-a):",
    min_value=1.00,
    max_value=3.00,
    value=1.56,
    step=0.01,
)
posto_dizela = izracunaj_dizel_dodatak_englmayer(prosjecna_cijena_goriva)
st.sidebar.info(f"Izračunati dodatak za dizel gorivo: **{posto_dizela:.1f}%**")

uploaded_pdf = st.file_uploader("Učitaj Englmayer PDF Obračun / Specifikaciju", type=["pdf"])

if uploaded_pdf is not None:
    with st.spinner("Parsiram Englmayer PDF specifikaciju..."):
        res_df = parsiraj_englmayer_pdf(uploaded_pdf)

    if not res_df.empty:
        st.success(f"Uspješno učitano i obrađeno pošiljaka: **{len(res_df)}**")

        ugovorene_osnove = []
        ugovoreni_dizeli = []
        zone_lista = []

        for _, row in res_df.iterrows():
            zona = odredi_zonu_englmayer(row["Destinacija"])
            zona_lista.append(zona)
            
            osnova = izracunaj_osnovnu_cijenu_palete(row["Masa (kg)"], zona)
            ugovorene_osnove.append(osnova)
            
            dizel_ugovor = osnova * (posto_dizela / 100.0)
            ugovoreni_dizeli.append(round(dizel_ugovor, 2))

        res_df["Zona"] = zona_lista
        res_df["Ugovorena Osnova (€)"] = ugovorene_osnove
        res_df["Ugovoreni Dizel (€)"] = ugovoreni_dizeli
        
        res_df["Razlika Transport (€)"] = res_df["Naplaćeni Transport (€)"] - res_df["Ugovorena Osnova (€)"]
        res_df["Razlika Dizel (€)"] = res_df["Naplaćeni Dizel Dodatak (€)"] - res_df["Ugovoreni Dizel (€)"]
        res_df["Ukupna Razlika (€)"] = res_df["Razlika Transport (€)"] + res_df["Razlika Dizel (€)"]

        tab1, tab2, tab3, tab4 = st.tabs([
            "📋 1. Detaljna Analiza po Stavkama",
            "🚨 2. Pregled Preplata i Viškova",
            "📈 3. Zbirni Financijski Pregled",
            "📊 4. Vizualna Analitika"
        ])

        with tab1:
            st.subheader("Pregled stavki iz Englmayer PDF-a uspoređen s ugovornim cjenikom")
            st.dataframe(res_df, use_container_width=True)
            st.download_button(
                "📥 Preuzmi Analizu u Excel",
                to_excel(res_df),
                "englmayer_analiza_racuna.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab2:
            st.subheader("Izdvojene pošiljke s odstupanjima (Preplate)")
            preplate_df = res_df[res_df["Ukupna Razlika (€)"] > 0.05]
            if preplate_df.empty:
                st.success("Nema pronađenih preplata u ovom obračunu!")
            else:
                st.dataframe(preplate_df, use_container_width=True)
                st.download_button(
                    "📥 Preuzmi Izvještaj Preplata (Excel)",
                    to_excel(preplate_df),
                    "englmayer_preplate.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        with tab3:
            st.subheader("📈 Zbirni financijski sažetak Englmayer fakture")
            uk_naplaceno_t = res_df["Naplaćeni Transport (€)"].sum()
            uk_ugovoreno_t = res_df["Ugovorena Osnova (€)"].sum()
            uk_naplaceno_d = res_df["Naplaćeni Dizel Dodatak (€)"].sum()
            uk_ugovoreno_d = res_df["Ugovoreni Dizel (€)"].sum()
            
            sve_naplaceno = uk_naplaceno_t + uk_naplaceno_d
            sve_ugovoreno = uk_ugovoreno_t + uk_ugovoreno_d
            ukupna_preplata = sve_naplaceno - sve_ugovoreno

            c1, c2, c3 = st.columns(3)
            c1.metric("Ukupno naplaćeno (Faktura)", f"{sve_naplaceno:,.2f} €")
            c2.metric("Ukupno po ugovornom cjeniku", f"{sve_ugovoreno:,.2f} €")
            c3.metric("Ukupna preplata / Višak", f"{max(0, ukupna_preplata):,.2f} €", delta_color="inverse")

            st.markdown("---")
            zbirni_t = pd.DataFrame([
                {"Kategorija": "Osnovni transport", "Naplaćeno (€)": round(uk_naplaceno_t, 2), "Ugovoreno (€)": round(uk_ugovoreno_t, 2)},
                {"Kategorija": f"Dizel dodatak ({posto_dizela:.1f}%)", "Naplaćeno (€)": round(uk_naplaceno_d, 2), "Ugovoreno (€)": round(uk_ugovoreno_d, 2)},
                {"Kategorija": "SVEUKUPNO", "Naplaćeno (€)": round(sve_naplaceno, 2), "Ugovoreno (€)": round(sve_ugovoreno, 2)}
            ])
            st.dataframe(zbirni_t, use_container_width=True)

        with tab4:
            st.subheader("📊 Distribucija troškova po destinacijama (Top gradovi)")
            if "Destinacija" in res_df.columns:
                top_d = res_df.groupby("Destinacija")["Naplaćeni Transport (€)"].sum().reset_index()
                top_d = top_d.sort_values(by="Naplaćeni Transport (€)", ascending=False).head(10)
                st.bar_chart(top_d.set_index("Destinacija"))
