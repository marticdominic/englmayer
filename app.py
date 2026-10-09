import streamlit as st
import pandas as pd
import pdfplumber
import pdf2image
import pytesseract
from PIL import Image
import re
import math
import io

st.set_page_config(
    page_title="Revizija i Audit Invoica - G. Englmayer",
    page_icon="📊",
    layout="wide"
)

# --- UGOVORENE KONSTANTE I DICTIONARY ---
def odrediti_zonu(zip_str):
    """Određuje zonu dostave prema prvoj znamenki hrvatskog ZIP koda."""
    if not zip_str:
        return "Zona 3"
    z_clean = str(zip_str).strip()
    prva_znam = z_clean[0]
    if prva_znam in ['1', '4']:
        return "Zona 1"
    elif prva_znam == '4' and z_clean.startswith(('40', '42', '43', '44', '47', '48', '49')):
        return "Zona 2"
    elif prva_znam in ['5']:
        return "Zona 3"
    elif prva_znam in ['3']:
        return "Zona 4"
    elif prva_znam in ['2']:
        return "Zona 5"
    else:
        return "Zona 3"

def ugovorena_cijena_osnovna(zona, masa_kg, tip_palete):
    """Ugovorena osnovna cijena po cjeniku za OF 002/2026."""
    cijene = {
        "Zona 1": {100: 35.00, 300: 40.00, 600: 45.00, 99999: 50.00},
        "Zona 2": {100: 38.00, 300: 43.00, 600: 48.00, 99999: 53.00},
        "Zona 3": {100: 42.00, 300: 47.00, 600: 51.00, 99999: 55.00},
        "Zona 4": {100: 45.00, 300: 50.00, 600: 54.00, 99999: 58.00},
        "Zona 5": {100: 48.00, 300: 53.00, 600: 56.00, 99999: 61.00},
    }
    zona_tabela = cijene.get(zona, cijene["Zona 3"])
    osnova = 51.00
    for limit_kg, cijena in sorted(zona_tabela.items()):
        if masa_kg <= limit_kg:
            osnova = cijena
            break
            
    # +50% za OWP palete prema ugovoru
    if tip_palete == "OWP":
        osnova *= 1.5
    return round(osnova, 2)

def izracunaj_dizel_dodatak(osnovna_cijena, cijena_goriva_trenutna=1.65):
    """Izračun dizel dodatka primjenom math.ceil na baznih 1.46 EUR."""
    baseline = 1.46
    if cijena_goriva_trenutna <= baseline:
        return 0.0
    razlika = cijena_goriva_trenutna - baseline
    postotak = math.ceil(razlika / 0.05) * 0.015  # 1.5% po svakih 0.05 EUR
    return round(osnovna_cijena * postotak, 2)

# --- HIBRIDNI PARSER (PDF + OCR ZA SLIKE) ---
def ekstrahiraj_tekst_iz_datoteke(uploaded_file):
    """Prepoznaje format i vraća sav tekst (kroz pdfplumber ili Tesseract OCR za slike/skenove)."""
    file_extension = uploaded_file.name.split('.')[-1].lower()
    sav_tekst = ""
    
    if file_extension == 'pdf':
        # Prvo pokušaj čitati tekstualni sloj preko pdfplumber
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t and len(t.strip()) > 50:
                    sav_tekst += t + "\n--- STRANICA ---\n"
        
        # Ako je PDF skeniran (nema tekstualnog sloja), pokreni OCR preko pdf2image
        if not sav_tekst.strip():
            uploaded_file.seek(0)
            images = pdf2image.convert_from_bytes(uploaded_file.read())
            for img in images:
                ocr_t = pytesseract.image_to_string(img, lang='hrv+eng')
                sav_tekst += ocr_t + "\n--- STRANICA OCR ---\n"
    else:
        # Direktna slika (PNG, JPG, JPEG)
        img = Image.open(uploaded_file)
        sav_tekst = pytesseract.image_to_string(img, lang='hrv+eng')
        
    return sav_tekst

def parse_englmayer_tekst(tekst):
    redci_paleta = []
    blocks = re.split(r'Datum naloga:', tekst, flags=re.IGNORECASE)
    
    for block in blocks[1:]:
        p_nalog_tekst = "Datum naloga:" + block
        lines = [l.strip() for l in p_nalog_tekst.split('\n') if l.strip()]
        
        dt_naloga = ""
        dt_isporuke = ""
        shpt = ""
        ref = ""
        
        for l in lines:
            if l.upper().startswith("DATUM NALOGA:"):
                dt_naloga = l.split(":")[-1].strip()
            elif "DATUM ISPORUKE" in l.upper():
                dt_isporuke = l.split(":")[-1].strip()
            elif "POŠILJKA:" in l.upper() or "POSILJKA:" in l.upper():
                m_sh = re.search(r'(ZAG-[\d\-]+|EP-[\d\-]+)', l, re.IGNORECASE)
                if m_sh: shpt = m_sh.group(1)
                m_la = re.search(r'(EP-[\d]+)', l, re.IGNORECASE)
                if m_la: shpt = m_la.group(1)
            elif "REFERENCA:" in l.upper():
                ref = l.split(":")[-1].strip()

        # Ekstrakcija ZIP-a i Grada
        zip_kod = "10410"
        grad = "Zagreb"
        m_zip = re.search(r'(?:HR-)?(\d{5})\s+([A-Za-zČĆŠĐŽčćšđž\s,]+)', p_nalog_tekst)
        if m_zip:
            zip_kod = m_zip.group(1)
            grad = m_zip.group(2).strip().split(',')[0]

        zona = odrediti_zonu(zip_kod)

        # Financije iz PDF-a / OCR-a
        naplaceno_osnovna = 0.0
        naplaceno_gorivo = 0.0
        for l in lines:
            if re.match(r'^(110|100)\s+', l):
                m_izn = re.findall(r'([\d\.]*,\d{2})', l)
                if m_izn:
                    naplaceno_osnovna = float(m_izn[-1].replace('.', '').replace(',', '.'))
            elif any(d_kw in l.lower() for d_kw in ['dizel', 'dodatak', 'gorivo']):
                m_izn = re.findall(r'([\d\.]*,\d{2})', l)
                if m_izn:
                    naplaceno_gorivo = float(m_izn[-1].replace('.', '').replace(',', '.'))

        # Ekstrakcija redaka paleta (svaka paleta s pripadajućom kilažom zasebno)
        for idx_l, linija in enumerate(lines):
            linija_upper = linija.upper()
            if any(t in linija_upper for t in ['EWP', 'FP', 'OWP', 'CLL']) or 'OTP' in linija_upper:
                tip_palete = "FP"
                for t_tip in ['EWP', 'OWP', 'CLL', 'FP']:
                    if t_tip in linija_upper:
                        tip_palete = t_tip
                        break
                
                kolicina = 1
                m_kol = re.search(r'(\d+)\s+' + tip_palete, linija_upper)
                if m_kol:
                    kolicina = int(m_kol.group(1))

                masa_kg = 0.0
                for k in range(max(0, idx_l - 1), min(len(lines), idx_l + 3)):
                    m_masa = re.search(r'(\d+[\d\.]*,\d{2,3})', lines[k])
                    if m_masa:
                        val = float(m_masa.group(1).replace('.', '').replace(',', '.'))
                        if val > 2.0:
                            masa_kg = val
                            break

                if masa_kg > 0:
                    m_otp = re.search(r'(otp-?[\d\/]+)', linija, re.IGNORECASE)
                    oznaka = m_otp.group(1) if m_otp else "Standard"

                    redci_paleta.append({
                        'LA-ID': shpt if shpt else "EP-GUEST",
                        'Referenca': ref if ref else "N/A",
                        'Oznaka_Broj': oznaka,
                        'Datum_Naloga': dt_naloga,
                        'Datum_Isporuke': dt_isporuke,
                        'Grad': grad,
                        'ZIP': zip_kod,
                        'Zona': zona,
                        'Broj_Paleta': kolicina,
                        'Tip_Palete': tip_palete,
                        'Masa_Palete_KG': round(masa_kg, 2),
                        'Naplaceno_Osnovna_EUR': naplaceno_osnovna,
                        'Naplaceno_Gorivo_EUR': naplaceno_gorivo,
                        'Naplaceno_Ukupno_EUR': round(naplaceno_osnovna + naplaceno_gorivo, 2)
                    })
                    
    return pd.DataFrame(redci_paleta)

# --- STREAMLIT KORISNIČKO SUČELJE ---
st.title("📊 Revizija G. Englmayer Invoica (PDF & Slike OCR)")
st.markdown("Automatsko čitanje PDF dokumenata i slika (screenshotova), ekstrakcija paleta sa zasebnim kilažama, provjera ugovora i izračun preplata.")

uploaded_file = st.file_uploader("Učitajte PDF račun ili sliku specifikacije", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    with st.spinner("Obrada dokumenta i OCR analiza..."):
        tekst_dok = ekstrahiraj_tekst_iz_datoteke(uploaded_file)
        df_rezultat = parse_englmayer_tekst(tekst_dok)
        
    if df_rezultat.empty:
        st.warning("Nije moguće automatski detektirati stavke. Provjerite je li dokument čitljiv.")
        with st.expander("Prikaži sirovi ekstrahirani tekst / OCR"):
            st.text(tekst_dok)
    else:
        cijena_goriva_input = st.sidebar.number_input("Trenutna cijena dizela (€/L)", value=1.65, step=0.01)
        
        ugovorene_osnove = []
        ugovorena_goriva = []
        for index, row in df_rezultat.iterrows():
            ug_osn = ugovorena_cijena_osnovna(row['Zona'], row['Masa_Palete_KG'], row['Tip_Palete'])
            ug_gor = izracunaj_dizel_dodatak(ug_osn, cijena_goriva_input)
            ugovorene_osnove.append(ug_osn)
            ugovorena_goriva.append(ug_gor)
            
        df_rezultat['Ugovoreno_Osnovna_EUR'] = ugovorene_osnove
        df_rezultat['Ugovoreno_Gorivo_EUR'] = ugovorena_goriva
        df_rezultat['Ugovoreno_Ukupno_EUR'] = df_rezultat['Ugovoreno_Osnovna_EUR'] + df_rezultat['Ugovoreno_Gorivo_EUR']
        df_rezultat['Razlika_Preplata_EUR'] = round(df_rezultat['Naplaceno_Ukupno_EUR'] - df_rezultat['Ugovoreno_Ukupno_EUR'], 2)

        tab1, tab2, tab3 = st.tabs(["📋 Detaljni Pregled po Paletama", "💰 Financijska Usporedba i Preplate", "📈 Sažetak po Zonama"])
        
        with tab1:
            st.subheader("Ekstrahirane stavke i palete sa zasebnim kilažama")
            st.dataframe(df_rezultat, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rezultat.to_excel(writer, index=False, sheet_name='Revizija_Palete')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Preuzmi izvještaj u Excel formatu (.xlsx)",
                data=excel_data,
                file_name="Englmayer_Revizija_Palete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab2:
            st.subheader("Usporedba naplaćenog vs. ugovorenog iznosa")
            preplate_df = df_rezultat[df_rezultat['Razlika_Preplata_EUR'] > 0]
            total_preplata = preplate_df['Razlika_Preplata_EUR'].sum()
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Ukupno stavki", len(df_rezultat))
            col2.metric("Ukupno naplaćeno", f"{df_rezultat['Naplaceno_Ukupno_EUR'].sum():.2f} €")
            col3.metric("Ukupno utvrđene preplate", f"{total_preplata:.2f} €", delta_color="inverse")
            
            st.dataframe(preplate_df[['LA-ID', 'Referenca', 'Oznaka_Broj', 'Grad', 'Zona', 'Masa_Palete_KG', 'Naplaceno_Ukupno_EUR', 'Ugovoreno_Ukupno_EUR', 'Razlika_Preplata_EUR']], use_container_width=True)

        with tab3:
            st.subheader("Agregirani pregled po zonama dostave")
            zona_group = df_rezultat.groupby('Zona').agg(
                Broj_Stavki=('LA-ID', 'count'),
                Ukupna_Masa_KG=('Masa_Palete_KG', 'sum'),
                Naplaceno_EUR=('Naplaceno_Ukupno_EUR', 'sum'),
                Ugovoreno_EUR=('Ugovoreno_Ukupno_EUR', 'sum'),
                Razlika_EUR=('Razlika_Preplata_EUR', 'sum')
            ).reset_index()
            st.dataframe(zona_group, use_container_width=True)
