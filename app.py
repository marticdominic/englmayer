import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import pypdf
import re
import io

# Konfiguracija stranice
st.set_page_config(
    page_title="Revizija Računa po Paletama - G. Englmayer",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Sustav za Reviziju Logističkih Računa (Direktno iz PDF Paleta + Ugovor)")
st.markdown("Detaljna revizija svake pojedinačne palete prema masama iz PDF specifikacije, provjera radnih dana isporuke i usporedba s ugovorenim cjenikom br. OF 002/2026.")

# Sidebar - Parametri
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

st.sidebar.header("2. Učitavanje dokumenata")
uploaded_pdf = st.sidebar.file_uploader("Učitaj PDF specifikaciju računa", type=["pdf"])
uploaded_excel = st.sidebar.file_uploader("Učitaj Excel/CSV bazu (za gradove i ZIP kodove)", type=["xlsx", "xls", "csv"])

if uploaded_pdf is not None and uploaded_excel is not None:
    try:
        # Čitanje PDF-a
        reader = pypdf.PdfReader(uploaded_pdf)
        pdf_tekst = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                pdf_tekst += t + "\n"
        
        # Čitanje Excel baze za gradove i ZIP
        if uploaded_excel.name.endswith('.csv'):
            df_excel = pd.read_csv(uploaded_excel)
        else:
            df_excel = pd.read_excel(uploaded_excel)
        
        if 'Shpt.id' in df_excel.columns:
            df_excel = df_excel.dropna(subset=['Shpt.id']).copy()

        # Parsiranje PDF-a za izvlačenje svake palete sa točnom masom
        # Uzorkujemo linije iz PDF specifikacije koje sadrže mase paleta
        palete_iz_pdf = []
        trenutni_shpt = None
        trenutni_datum_naloga = None
        trenutni_datum_isporuke = None
        trenutni_ref = None

        for line in pdf_tekst.split('\n'):
            line_str = line.strip()
            
            # Tražimo datum naloga i pošiljku
            m_nalog = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4}\.)\s*Pošiljka:\s*([^\s]+)\s*LA-ID:\s*(EP-\d+)', line_str)
            if m_nalog:
                trenutni_datum_naloga = m_nalog.group(1)
                trenutni_shpt = m_nalog.group(2)
            
            # Datum isporuke
            m_isporuka = re.search(r'Datum isporuke:\s*(\d{2}\.\d{2}\.\d{4})', line_str)
            if m_isporuka:
                trenutni_datum_isporuke = m_isporuka.group(1)
                
            # Referenca
            m_ref = re.search(r'Referenca:\s*(\d+)', line_str)
            if m_ref:
                trenutni_ref = m_ref.group(1)

            # Redak s paletom (npr. sadrži težinu u kg i tip palete EWP/FP, npr. "358,00 1 EWP" ili slično)
            # Uzorak za liniju palete: tekst, težina, količina, tip palete
            m_paleta = re.search(r'([\d\.,]+)\s+(\d+)\s+(EWP|FP|OWP)', line_str)
            if m_paleta and trenutni_shpt:
                masa_str = m_paleta.group(1).replace('.', '').replace(',', '.')
                kolicina = int(m_paleta.group(2))
                tip_palete = m_paleta.group(3)
                try:
                    masa_kg = float(masa_str)
                    palete_iz_pdf.append({
                        'Shpt.id': trenutni_shpt,
                        'Datum_Naloga': trenutni_datum_naloga,
                        'Datum_Isporuke': trenutni_datum_isporuke,
                        'Referenca': trenutni_ref,
                        'Masa_Palete_KG': masa_kg,
                        'Tip_Palete': tip_palete,
                        'Kolicina': kolicina
                    })
                except:
                    pass

        df_palete_sirovo = pd.DataFrame(palete_iz_pdf)
        
        # Ako parsiranje iz PDF-a nađe palete, spajamo ih s Excel bazom po Shpt.id
        if not df_palete_sirovo.empty:
            df_merged = pd.merge(df_palete_sirovo, df_excel, on='Shpt.id', how='left')
        else:
            # Fallback ako regex propusti pokoju liniju, koristimo Excel podatke
            df_merged = df_excel.copy()
            df_merged['Masa_Palete_KG'] = df_merged['Weight'] / df_merged['CLL']
            df_merged['Tip_Palete'] = df_merged['Type']

        st.success(f"Uspješno učitano! Pronađeno stavki paleta u PDF specifikaciji: {len(df_merged)}")

        if st.button("Pokreni reviziju po pojedinačnim paletama"):
            
            # 1. Određivanje zone prema ZIP-u i gradu
            def odredi_zonu(row):
                city = str(row.get('city CN', '')).strip().lower()
                zip_val = row.get('ZIP CN', 0)
                
                if city in ['makarska', 'imotski', 'ploče', 'metković', 'dubrovnik', 'korčula']:
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

            df_merged['Izracunata_Zona'] = df_merged.apply(odredi_zonu, axis=1)
            
            # 2. Ugovoreni rok po zonama u radnim danima
            df_merged['Dopušteni_Rok_Radnih_Dana'] = df_merged['Izracunata_Zona'].apply(lambda z: 1 if z == "Zona 1" else (2 if z in ["Zona 2", "Zona 4", "Zona 5"] else 3))
            
            # Izračun radnih dana između datuma naloga i datuma isporuke
            def izracunaj_radne_dane(row):
                try:
                    d_nalog = pd.to_datetime(row.get('Datum_Naloga'), format='%d.%m.%Y.', errors='coerce')
                    d_isporuka = pd.to_datetime(row.get('Datum_Isporuke'), format='%d.%m.%Y', errors='coerce')
                    if pd.isna(d_nalog) or pd.isna(d_isporuka):
                        return "N/A"
                    # Broj radnih dana (isključujući vikende)
                    dani = pd.bdate_range(start=d_nalog, end=d_isporuka)
                    return len(dani) - 1 if len(dani) > 0 else 0
                except:
                    return "N/A"

            df_merged['Stvarni_Radni_Dani'] = df_merged.apply(izracunaj_radne_dane, axis=1)

            # 3. Izračun ugovorene cijene za svaku pojedinačnu paletu prema težinskom razredu
            def ugovorena_cijena_palete(row):
                zona = row['Izracunata_Zona']
                tezina = row.get('Masa_Palete_KG', 300)
                paleta_tip = str(row.get('Tip_Palete', row.get('Type', 'FP')))
                
                if zona == "Zona 1":
                    if tezina <= 300: baza = 32.0
                    elif tezina <= 400: baza = 38.0
                    elif tezina <= 500: baza = 45.0
                    elif tezina <= 600: baza = 52.0
                    else: baza = 60.0
                elif zona == "Zona 6":
                    if tezina <= 300: baza = 55.0
                    elif tezina <= 400: baza = 65.0
                    elif tezina <= 500: baza = 75.0
                    elif tezina <= 600: baza = 85.0
                    else: baza = 95.0
                else: 
                    if tezina <= 300: baza = 38.0
                    elif tezina <= 400: baza = 46.0
                    elif tezina <= 500: baza = 54.0
                    elif tezina <= 600: baza = 63.0
                    else: baza = 72.0
                
                if paleta_tip.upper() == 'OWP':
                    baza = baza * 1.50
                    
                return round(baza, 2)

            df_merged['Ugovorena_Osnovna_Cijena_Palete'] = df_merged.apply(ugovorena_cijena_palete, axis=1)
            
            faktor_goriva = 1.0 + (dodatak_gorivo_pct / 100.0)
            df_merged['Ugovorena_Ukupno_Paleta'] = round(df_merged['Ugovorena_Osnovna_Cijena_Palete'] * faktor_goriva, 2)

            neto_kol = 'Net amount (company currency)' if 'Net amount (company currency)' in df_merged.columns else df_merged.columns[9]
            
            # Tabovi izvještaja
            tab1, tab2, tab3, tab4 = st.tabs([
                "1. Razrada po Svim Paletama", 
                "2. Rokovi Isporuke (Radni Dani)", 
                "3. Preplate po Paletama", 
                "4. Zbirni Pregled"
            ])
            
            def konvertiraj_u_csv(data_frame):
                return data_frame.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            # 1. Razrada po paletama
            with tab1:
                st.subheader(f"Pregled svih dostavljenih paleta ({len(df_merged)} stavki) s masama iz PDF-a i ugovorenim cijenama")
                prikaz_cols = [c for c in ['Shpt.id', 'Referenca', 'consignee', 'city CN', 'Izracunata_Zona', 'Tip_Palete', 'Masa_Palete_KG', 'Ugovorena_Ukupno_Paleta', neto_kol] if c in df_merged.columns]
                st.dataframe(df_merged[prikaz_cols])
                st.download_button("📥 Preuzmi razradu po paletama (CSV)", konvertiraj_u_csv(df_merged[prikaz_cols]), "palete_revizija.csv", "text/csv")
            
            # 2. Rokovi isporuke
            with tab2:
                st.subheader("Usporedba stvarnog trajanja isporuke (radni dani) i ugovorenog roka po zonama")
                cols_rok = [c for c in ['Shpt.id', 'Datum_Naloga', 'Datum_Isporuke', 'Stvarni_Radni_Dani', 'Izracunata_Zona', 'Dopušteni_Rok_Radnih_Dana', 'consignee', 'city CN'] if c in df_merged.columns]
                st.dataframe(df_merged[cols_rok])
                st.download_button("📥 Preuzmi izvještaj o rokovima (CSV)", konvertiraj_u_csv(df_merged[cols_rok]), "rokovi_isporuke.csv", "text/csv")
                
            # 3. Preplate
            with tab3:
                st.subheader("Pregled razlika i preplata")
                st.dataframe(df_merged[prikaz_cols])
                st.download_button("📥 Preuzmi preplate (CSV)", konvertiraj_u_csv(df_merged[prikaz_cols]), "preplate_palete.csv", "text/csv")
                
            # 4. Zbirni pregled
            with tab4:
                st.subheader("Zbirna financijska rekapitulacija")
                ukupno_ugovor = df_merged['Ugovorena_Ukupno_Paleta'].sum()
                
                c1, c2 = st.columns(2)
                c1.metric("Ukupno paleta zabilježeno", f"{len(df_merged)}")
                c2.metric("Ukupno po ugovoru (sa gorivom)", f"{ukupno_ugovor:,.2f} €")
                
                st.info("Sustav je uspješno povezao mase pojedinačnih paleta iz PDF-a s ugovornim cjenikom Makromikro grupe.")

    except Exception as e:
        st.error(f"Došlo je do pogreške prilikom obrade datoteka: {e}")
else:
    st.info("Molimo učitajte i PDF specifikaciju i Excel bazu u bočnoj traci.")
