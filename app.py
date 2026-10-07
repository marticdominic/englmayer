import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import pypdf
import re
import io

# Konfiguracija stranice
st.set_page_config(
    page_title="Revizija Računa iz PDF-a - G. Englmayer",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Sustav za Reviziju Logističkih Računa (Isključivo PDF + Ugovorni Cjenik)")
st.markdown("Direktna analiza, razrada po paletama, provjera rokova isporuke i usporedba zbrojenih pošiljaka prema ugovoru br. OF 002/2026.")

# Sidebar - Parametri obračuna
st.sidebar.header("1. Ugovorni parametri")
cijena_goriva = st.sidebar.number_input("Prosječna cijena dizel goriva (€ bez PDV-a):", value=1.87, step=0.01)

def izracunaj_dodatak_gorivo(cijena):
    osnova = 1.46
    korak = 0.07
    if cijena <= osnova:
        return 0.0
    else:
        razlika = cijena - osnova
        return round((razlika / korak) * 1.0, 2)

dodatak_gorivo_pct = izracunaj_dodatak_gorivo(cijena_goriva)
st.sidebar.info(f"Izračunati dodatak za gorivo (baza 1.46 €): **{dodatak_gorivo_pct}%**")

st.sidebar.header("2. Učitavanje PDF-a")
uploaded_pdf = st.sidebar.file_uploader("Učitaj PDF specifikaciju računa", type=["pdf"])

if uploaded_pdf is not None:
    try:
        # Čitanje PDF-a
        reader = pypdf.PdfReader(uploaded_pdf)
        pdf_tekst = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                pdf_tekst += t + "\n"
        
        # 1. Ekstrakcija naplaćenih iznosa po pošiljkama (LA-ID) iz zbirnog dijela računa
        pdf_iznosi = {}
        for line in pdf_tekst.split('\n'):
            match = re.search(r'(EP-\d+)\s+(\d{2}\.\d{2}\.\d{4}\.)\s+(\d+)?\s*(\d+%\s+)?([\d\.,]+)', line)
            if match:
                shpt_id = match.group(1)
                amount_str = match.group(5).replace('.', '').replace(',', '.')
                try:
                    pdf_iznosi[shpt_id] = float(amount_str)
                except:
                    pass

        # 2. Parsiranje detaljne specifikacije paleta, primatelja, gradova, ZIP-ova i masa
        redci_paleta = []
        trenutni_shpt = None
        trenutni_datum_naloga = None
        trenutni_datum_isporuke = None
        trenutni_ref = None
        trenutni_primatelj = None
        trenutni_grad = "Zagreb"
        trenutni_zip = 10000

        for line in pdf_tekst.split('\n'):
            line_str = line.strip()
            
            m_nalog = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4}\.)\s*Pošiljka:\s*([^\s]+)\s*LA-ID:\s*(EP-\d+)', line_str)
            if m_nalog:
                trenutni_datum_naloga = m_nalog.group(1)
                trenutni_shpt = m_nalog.group(3)
                trenutni_primatelj = "N/A"
                trenutni_grad = "Zagreb"
                trenutni_zip = 10000
            
            if "Primatelj" in line_str:
                trenutni_primatelj = line_str
                m_zip_grad = re.search(r'HR-(\d{5})\s+([A-Za-zČĆŠĐŽčćšđž\s\-\.]+)', line_str)
                if m_zip_grad:
                    trenutni_zip = int(m_zip_grad.group(1))
                    trenutni_grad = m_zip_grad.group(2).strip()

            m_isporuka = re.search(r'Datum isporuke:\s*(\d{2}\.\d{2}\.\d{4})', line_str)
            if m_isporuka:
                trenutni_datum_isporuke = m_isporuka.group(1)
                
            m_ref = re.search(r'Referenca:\s*(\d+)', line_str)
            if m_ref:
                trenutni_ref = m_ref.group(1)

            m_paleta = re.search(r'([\d\.,]+)(\d)\s+(EWP|FP|OWP)', line_str)
            if m_paleta and trenutni_shpt:
                masa_str = m_paleta.group(1).replace('.', '').replace(',', '.')
                kolicina = int(m_paleta.group(2))
                tip_palete = m_paleta.group(3)
                try:
                    masa_kg = float(masa_str)
                    for _ in range(kolicina):
                        redci_paleta.append({
                            'LA-ID': trenutni_shpt,
                            'Referenca': trenutni_ref if trenutni_ref else "N/A",
                            'Datum_Naloga': trenutni_datum_naloga,
                            'Datum_Isporuke': trenutni_datum_isporuke,
                            'Primatelj': trenutni_primatelj,
                            'Grad': trenutni_grad,
                            'ZIP': trenutni_zip,
                            'Masa_Palete_KG': masa_kg,
                            'Tip_Palete': tip_palete
                        })
                except:
                    pass

        df_palete = pd.DataFrame(redci_paleta)
        st.success(f"PDF uspješno učitan! Pronađeno pojedinačnih paleta: {len(df_palete)}")

        if st.button("Pokreni reviziju i generiraj izvještaje"):
            
            # Određivanje zone
            def odredi_zonu(row):
                city = str(row.get('Grad', '')).strip().lower()
                zip_val = row.get('ZIP', 10000)
                
                if any(g in city for g in ['makarska', 'imotski', 'ploče', 'metković', 'dubrovnik', 'korčula', 'mokosica']):
                    return "Zona 6"
                
                try:
                    z = int(zip_val)
                    if 10000 <= z <= 10450:
                        return "Zona 1"
                    elif (20000 <= z <= 23999) or (50000 <= z <= 53999):
                        return "Zona 2"
                    elif 30000 <= z <= 35000:
                        return "Zona 3"
                    elif 40000 <= z <= 49000:
                        return "Zona 4"
                    elif 51000 <= z <= 51500:
                        return "Zona 5"
                    else:
                        return "Zona 2"
                except:
                    return "Zona 2"

            df_palete['Izracunata_Zona'] = df_palete.apply(odredi_zonu, axis=1)
            df_palete['Dopušteni_Rok_Radnih_Dana'] = df_palete['Izracunata_Zona'].apply(lambda z: 1 if z == "Zona 1" else (2 if z in ["Zona 2", "Zona 4", "Zona 5"] else 3))
            
            # Radni dani
            def izracunaj_radne_dane(row):
                try:
                    d_nalog = pd.to_datetime(row.get('Datum_Naloga'), format='%d.%m.%Y.', errors='coerce')
                    d_isporuka = pd.to_datetime(row.get('Datum_Isporuke'), format='%d.%m.%Y', errors='coerce')
                    if pd.isna(d_nalog) or pd.isna(d_isporuka):
                        return 0
                    dani = pd.bdate_range(start=d_nalog, end=d_isporuka)
                    return len(dani) - 1 if len(dani) > 0 else 0
                except:
                    return 0

            df_palete['Stvarni_Radni_Dani'] = df_palete.apply(izracunaj_radne_dane, axis=1)

            # Cijena palete
            def ugovorena_cijena_palete(row):
                zona = row['Izracunata_Zona']
                tezina = row['Masa_Palete_KG']
                paleta_tip = str(row.get('Tip_Palete', 'FP'))
                
                if zona == "Zona 1":
                    baza = 32.0 if tezina <= 300 else (38.0 if tezina <= 400 else (45.0 if tezina <= 500 else (52.0 if tezina <= 600 else 60.0)))
                elif zona == "Zona 6":
                    baza = 55.0 if tezina <= 300 else (65.0 if tezina <= 400 else (75.0 if tezina <= 500 else (85.0 if tezina <= 600 else 95.0)))
                else: 
                    baza = 38.0 if tezina <= 300 else (46.0 if tezina <= 500 else (54.0 if tezina <= 500 else (63.0 if tezina <= 600 else 72.0)))
                
                if paleta_tip.upper() == 'OWP':
                    baza = baza * 1.50
                return round(baza, 2)

            df_palete['Ugovorena_Osnovna_Cijena'] = df_palete.apply(ugovorena_cijena_palete, axis=1)
            
            faktor_goriva = 1.0 + (dodatak_gorivo_pct / 100.0)
            df_palete['Ugovoreno_Paleta_Sa_Gorivom'] = round(df_palete['Ugovorena_Osnovna_Cijena'] * faktor_goriva, 2)

            # Tabovi izvještaja
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "1. Pregled po Paletama", 
                "2. Provjera Rokova Isporuke", 
                "3. Zbirni Financijski Pregled",
                "4. Usporedba po Pošiljkama (Reference)",
                "5. Preplate po Pošiljkama"
            ])
            
            def konvertiraj_u_csv(data_frame):
                return data_frame.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            # 1. Pregled po paletama
            with tab1:
                st.subheader(f"Popis svih paleta izvađenih iz PDF-a ({len(df_palete)} stavki)")
                st.dataframe(df_palete)
                st.download_button("📥 Preuzmi palete (CSV)", konvertiraj_u_csv(df_palete), "palete_iz_pdf-a.csv", "text/csv")
            
            # 2. Rokovi isporuke
            with tab2:
                st.subheader("Provjera ugovorenih radnih dana dostave")
                cols_rok = ['LA-ID', 'Referenca', 'Datum_Naloga', 'Datum_Isporuke', 'Stvarni_Radni_Dani', 'Izracunata_Zona', 'Dopušteni_Rok_Radnih_Dana', 'Grad']
                st.dataframe(df_palete[cols_rok])
                st.download_button("📥 Preuzmi rokove (CSV)", konvertiraj_u_csv(df_palete[cols_rok]), "rokovi_isporuke_pdf.csv", "text/csv")
                
            # 3. Zbirni pregled
            with tab3:
                st.subheader("Zbirna rekapitulacija po ugovoru")
                ukupno_ugovor = df_palete['Ugovoreno_Paleta_Sa_Gorivom'].sum()
                c1, c2 = st.columns(2)
                c1.metric("Ukupno paleta u PDF-u", f"{len(df_palete)}")
                c2.metric("Ukupno po ugovornom cjeniku", f"{ukupno_ugovor:,.2f} €")
                
            # 4. Usporedba po pošiljkama (Reference)
            with tab4:
                st.subheader("Usporedba pošiljaka zbrojenih po referencama / LA-ID brojevima")
                
                # Agregacija po LA-ID i Referenci
                df_posiljke = df_palete.groupby(['LA-ID', 'Referenca', 'Grad', 'Izracunata_Zona', 'Datum_Naloga', 'Datum_Isporuke']).agg(
                    Broj_Paleta=('Masa_Palete_KG', 'count'),
                    Ukupna_Masa_KG=('Masa_Palete_KG', 'sum'),
                    Ugovoreno_Ukupno_EUR=('Ugovoreno_Paleta_Sa_Gorivom', 'sum')
                ).reset_index()
                
                # Dodavanje naplaćenog iznosa iz PDF-a
                df_posiljke['Naplaćeno_Po_PDF_EUR'] = df_posiljke['LA-ID'].map(pdf_iznosi).fillna(0.0)
                df_posiljke['Razlika (Naplaćeno - Ugovoreno)'] = round(df_posiljke['Naplaćeno_Po_PDF_EUR'] - df_posiljke['Ugovoreno_Ukupno_EUR'], 2)
                
                st.dataframe(df_posiljke)
                st.download_button("📥 Preuzmi usporedbu pošiljaka (CSV)", konvertiraj_u_csv(df_posiljke), "usporedba_po_posiljkama.csv", "text/csv")

            # 5. Preplate po pošiljkama
            with tab5:
                st.subheader("Izdvojene preplate (gdje je naplaćeni iznos veći od ugovornog)")
                if 'df_posiljke' in locals():
                    df_preplate = df_posiljke[df_posiljke['Razlika (Naplaćeno - Ugovoreno)'] > 0].sort_values(by='Razlika (Naplaćeno - Ugovoreno)', ascending=False)
                    st.dataframe(df_preplate)
                    st.download_button("📥 Preuzmi preplate po pošiljkama (CSV)", konvertiraj_u_csv(df_preplate), "preplate_po_posiljkama.csv", "text/csv")
                else:
                    st.info("Pregledajte prvo tab 4 za izračun.")

    except Exception as e:
        st.error(f"Greška kod obrade PDF-a: {e}")
else:
    st.info("Molimo učitajte PDF specifikaciju računa u bočnoj traci.")
