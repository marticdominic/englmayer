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

st.title("📄 Sustav za Reviziju Logističkih Računa (PDF + Službeni Ugovorni Cjenik OF 002/2026)")
st.markdown("Direktna analiza po paletama, analitika rokova isporuke i usporedba pošiljaka prema službenom ugovoru G. Englmayer.")

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

        if st.button("Pokreni reviziju prema službenom cjeniku"):
            
            # Određivanje zone prema službenom ugovoru (Prilog 1)
            def odredi_zonu(row):
                city = str(row.get('Grad', '')).strip().lower()
                zip_val = row.get('ZIP', 10000)
                
                # Zona 5 (mjesta koja gravitiraju Makarskoj, Imotskom i Pločama idu pod Zonu 6 prema ugovoru)
                if any(g in city for g in ['makarska', 'imotski', 'ploče', 'metković', 'dubrovnik', 'korčula', 'mokosica']):
                    return "Zona 6"
                
                try:
                    z = int(zip_val)
                    if (10000 <= z <= 10450) or (40000 <= z <= 49000): # Zona 1 i Zona 4 primjeri raspona
                        # Provjerimo točne raspone ZIP-ova iz ugovora (Prilog 1)
                        pass
                    
                    # Pojednostavljeni provjereni rasponi prema ugovoru:
                    if 10000 <= z <= 10450: return "Zona 1"
                    elif (20000 <= z <= 23999) or (50000 <= z <= 53999): return "Zona 2"
                    elif 30000 <= z <= 35000: return "Zona 3"
                    elif 40000 <= z <= 49000: return "Zona 4"
                    elif 51000 <= z <= 51500: return "Zona 5"
                    else: return "Zona 2"
                except:
                    return "Zona 2"

            df_palete['Izracunata_Zona'] = df_palete.apply(odredi_zonu, axis=1)
            df_palete['Dopušteni_Rok_Radnih_Dana'] = df_palete['Izracunata_Zona'].apply(lambda z: 1 if z == "Zona 1" else (2 if z in ["Zona 2", "Zona 4", "Zona 5"] else 3))
            
            # Radni dani i status roka
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
            df_palete['Status_Roka'] = df_palete.apply(
                lambda r: 'U roku' if r['Stvarni_Radni_Dani'] <= r['Dopušteni_Rok_Radnih_Dana'] else 'Izvan roka', 
                axis=1
            )

            # Službena ugovorna tablica cijena po paleti (Prilog 1)
            def ugovorena_cijena_palete(row):
                zona = row['Izracunata_Zona']
                tezina = row['Masa_Palete_KG']
                paleta_tip = str(row.get('Tip_Palete', 'FP'))
                
                # Matrica cijena [do 300, do 400, do 500, do 600, do 700]
                cjenik = {
                    "Zona 1": [23.0, 25.0, 30.0, 33.0, 38.0],
                    "Zona 2": [26.0, 29.0, 35.0, 39.0, 43.0],
                    "Zona 3": [36.0, 40.0, 45.0, 47.0, 51.0],
                    "Zona 4": [42.0, 47.0, 50.0, 55.0, 65.0],
                    "Zona 5": [44.0, 48.0, 51.0, 56.0, 68.0],
                    "Zona 6": [55.0, 59.0, 63.0, 65.0, 79.0]
                }
                
                zone_indeks = {"Zona 1": 0, "Zona 2": 1, "Zona 3": 2, "Zona 4": 3, "Zona 5": 4, "Zona 6": 5}
                z_idx = zone_indeks.get(zona, 1)
                
                if tezina <= 300: t_idx = 0
                elif tezina <= 400: t_idx = 1
                elif tezina <= 500: t_idx = 2
                elif tezina <= 600: t_idx = 3
                else: t_idx = 4
                
                baza = cjenik.get(zona, cjenik["Zona 2"])[t_idx]
                
                # +50% za OWP / van gabaritne palete
                if paleta_tip.upper() == 'OWP':
                    baza = baza * 1.50
                    
                return round(baza, 2)

            df_palete['Ugovorena_Osnovna_Cijena'] = df_palete.apply(ugovorena_cijena_palete, axis=1)
            
            faktor_goriva = 1.0 + (dodatak_gorivo_pct / 100.0)
            df_palete['Ugovoreno_Paleta_Sa_Gorivom'] = round(df_palete['Ugovorena_Osnovna_Cijena'] * faktor_goriva, 2)

            # Tabovi izvještaja
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "1. Pregled po Paletama", 
                "2. Provjera Rokova Isporuke (Analitika)", 
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
            
            # 2. Provjera rokova isporuke (Analitika)
            with tab2:
                st.subheader("Analitički izvještaj: Učinkovitost i točnost rokova dostave")
                
                ukupno_stavki = len(df_palete)
                broj_u_roku = len(df_palete[df_palete['Status_Roka'] == 'U roku'])
                broj_izvan_rok = len(df_palete[df_palete['Status_Roka'] == 'Izvan roka'])
                
                pct_u_roku = (broj_u_roku / ukupno_stavki) * 100 if ukupno_stavki > 0 else 0
                pct_izvan_rok = (broj_izvan_rok / ukupno_stavki) * 100 if ukupno_stavki > 0 else 0
                
                kpi1, kpi2, kpi3 = st.columns(3)
                kpi1.metric("U roku (Uspješnost)", f"{pct_u_roku:.1f}%", f"{broj_u_roku} paleta")
                kpi2.metric("Izvan roka (Kašnjenje)", f"{pct_izvan_rok:.1f}%", f"{broj_izvan_rok} paleta")
                kpi3.metric("Ukupno analizirano", f"{ukupno_stavki} paleta")
                
                st.markdown("---")
                
                cols_rok = ['LA-ID', 'Referenca', 'Datum_Naloga', 'Datum_Isporuke', 'Stvarni_Radni_Dani', 'Izracunata_Zona', 'Dopušteni_Rok_Radnih_Dana', 'Status_Roka', 'Grad']
                st.dataframe(df_palete[cols_rok])
                st.download_button("📥 Preuzmi analitiku rokova (CSV)", konvertiraj_u_csv(df_palete[cols_rok]), "analitika_rokova_isporuke.csv", "text/csv")
                
            # 3. Zbirni pregled
            with tab3:
                st.subheader("Zbirna rekapitulacija po ugovoru")
                ukupno_ugovor = df_palete['Ugovoreno_Paleta_Sa_Gorivom'].sum()
                c1, c2 = st.columns(2)
                c1.metric("Ukupno paleta u PDF-u", f"{len(df_palete)}")
                c2.metric("Ukupno po službenom ugovornom cjeniku", f"{ukupno_ugovor:,.2f} €")
                
            # 4. Usporedba po pošiljkama (Reference)
            with tab4:
                st.subheader("Usporedba pošiljaka zbrojenih po referencama / LA-ID brojevima")
                
                df_posiljke = df_palete.groupby(['LA-ID', 'Referenca', 'Grad', 'Izracunata_Zona', 'Datum_Naloga', 'Datum_Isporuke']).agg(
                    Broj_Paleta=('Masa_Palete_KG', 'count'),
                    Ukupna_Masa_KG=('Masa_Palete_KG', 'sum'),
                    Ugovoreno_Ukupno_EUR=('Ugovoreno_Paleta_Sa_Gorivom', 'sum')
                ).reset_index()
                
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
