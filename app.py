import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import pypdf
import re
import io

# Konfiguracija stranice
st.set_page_config(
    page_title="Revizija Logističkog Računa - G. Englmayer (PDF + Cjenik)",
    page_icon="📄",
    layout="wide"
)

# Naslov aplikacije
st.title("📄 Sustav za Reviziju Računa prema PDF Specifikaciji i Ugovoru")
st.markdown("Direktna usporedba službenih stavki iz PDF računa G. Englmayer s ugovornim cjenikom Makromikro grupe (br. OF 002/2026).")

# Sidebar za parametre obračuna i datoteke
st.sidebar.header("Parametri obračuna i cjenika")
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

st.sidebar.header("Učitavanje dokumenata")
uploaded_pdf = st.sidebar.file_uploader("Učitaj PDF specifikaciju računa", type=["pdf"])
uploaded_excel = st.sidebar.file_uploader("Učitaj Excel/CSV bazu pošiljaka (za gradove i težine)", type=["xlsx", "xls", "csv"])

if uploaded_pdf is not None:
    try:
        # Parsiranje PDF specifikacije
        reader = pypdf.PdfReader(uploaded_pdf)
        pdf_tekst = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                pdf_tekst += t + "\n"
        
        # Ekstrakcija redaka iz PDF-a (LA-NR, Datum, Referenca, Iznos)
        pdf_entries = []
        for line in pdf_tekst.split('\n'):
            match = re.search(r'(EP-\d+)\s+(\d{2}\.\d{2}\.\d{4}\.)\s+(\d+)\s+(\d+%\s+)?([\d\.,]+)', line)
            if match:
                shpt_id = match.group(1)
                date_str = match.group(2)
                ref = match.group(3)
                amount_str = match.group(5).replace('.', '').replace(',', '.')
                try:
                    amount = float(amount_str)
                    pdf_entries.append({'Route ID': shpt_id, 'PDF_Datum': date_str, 'PDF_Referenca': ref, 'PDF_Naplaćeni_Iznos': amount})
                except:
                    pass
        
        df_pdf = pd.DataFrame(pdf_entries)
        st.success(f"📄 PDF specifikacija uspješno učitana! Pronađeno stavki na računu: {len(df_pdf)}")
        
        # Spajanje s Excel bazom ako je učitana
        if uploaded_excel is not None:
            if uploaded_excel.name.endswith('.csv'):
                df_excel = pd.read_csv(uploaded_excel)
            else:
                df_excel = pd.read_excel(uploaded_excel)
            
            if 'Shpt.id' in df_excel.columns:
                df_excel = df_excel.dropna(subset=['Shpt.id']).copy()
            
            # Spajamo podatke preko Route ID / Shpt.id
            if 'Route ID' in df_excel.columns and 'Route ID' in df_pdf.columns:
                df = pd.merge(df_pdf, df_excel, on='Route ID', how='left')
            else:
                df = df_pdf
                st.warning("Nije pronađen poklapajući 'Route ID' stupac za spajanje s Excelom, prikazuju se čisti podaci iz PDF-a.")
        else:
            df = df_pdf
            st.info("💡 Savjet: Učitajte i Excel tabliku u sidebaru kako biste dobili potpune podatke o gradovima, ZIP kodovima i težinama.")

        if st.button("Pokreni reviziju na temelju PDF-a i ugovora"):
            
            # --- UGOVORNA LOGIKA I PRAVILA ---
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

            df['Izracunata_Zona'] = df.apply(odredi_zonu, axis=1)
            
            def izracunaj_tezinu_po_paleti(row):
                cll = row.get('CLL', 1)
                weight = row.get('Weight', 0)
                if pd.isna(cll) or cll <= 0:
                    cll = 1
                return weight / cll

            df['Tezina_Po_Paleti'] = df.apply(izracunaj_tezinu_po_paleti, axis=1)

            def ugovorena_cijena_palete(row):
                zona = row['Izracunata_Zona']
                tezina = row['Tezina_Po_Paleti']
                paleta_tip = str(row.get('Type', 'FP'))
                
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
                    
                cll = row.get('CLL', 1)
                if pd.isna(cll) or cll <= 0: cll = 1
                
                return baza * cll

            df['Ugovorena_Osnovna_Cijena'] = df.apply(ugovorena_cijena_palete, axis=1)
            
            faktor_goriva = 1.0 + (dodatak_gorivo_pct / 100.0)
            df['Ugovoreno_Ukupno'] = df['Ugovorena_Osnovna_Cijena'] * faktor_goriva
            
            neto_kol = 'PDF_Naplaćeni_Iznos'
            
            total_shipments = len(df)
            
            st.info(f"📊 Analizirano pošiljaka iz PDF-a: **{total_shipments}**")
            
            # Tabovi izvještaja
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "1. PDF vs Ugovor", 
                "2. Usporedba cijena", 
                "3. Preplate", 
                "4. Zbirni pregled", 
                "5. Sirovi PDF tekst"
            ])
            
            def konvertiraj_u_csv(data_frame):
                return data_frame.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            # 1. PDF vs Ugovor
            with tab1:
                st.subheader("Usporedba službenih podataka iz PDF specifikacije i ugovornih zona")
                st.dataframe(df[['Route ID', 'PDF_Datum', 'PDF_Referenca', 'PDF_Naplaćeni_Iznos', 'city CN', 'Izracunata_Zona', 'CLL', 'Weight']].head(20))
                st.download_button("📥 Preuzmi PDF analizu (CSV)", konvertiraj_u_csv(df), "pdf_vs_ugovor.csv", "text/csv")
            
            # 2. Usporedba cijena
            with tab2:
                st.subheader("Detaljna usporedba iznosa s PDF računa i ugovorenog iznosa")
                prikaz_df = df[['Route ID', 'PDF_Datum', 'city CN', 'Izracunata_Zona', 'Weight', 'CLL', 'PDF_Naplaćeni_Iznos', 'Ugovoreno_Ukupno']].copy()
                prikaz_df['Razlika (PDF - Ugovor)'] = prikaz_df['PDF_Naplaćeni_Iznos'] - prikaz_df['Ugovoreno_Ukupno']
                st.dataframe(prikaz_df.head(25))
                st.download_button("📥 Preuzmi 'Usporedba cijena' (CSV)", konvertiraj_u_csv(prikaz_df), "usporedba_cijena_pdf.csv", "text/csv")
                
            # 3. Preplate
            with tab3:
                st.subheader("Izdvojene preplate (gdje je PDF iznos veći od ugovorenog)")
                preplate_df = prikaz_df[prikaz_df['Razlika (PDF - Ugovor)'] > 0].sort_values(by='Razlika (PDF - Ugovor)', ascending=False)
                st.dataframe(preplate_df.head(15))
                st.download_button("📥 Preuzmi 'Preplate' (CSV)", konvertiraj_u_csv(preplate_df), "preplate_pdf.csv", "text/csv")
                
            # 4. Zbirni pregled
            with tab4:
                st.subheader("Zbirni financijski pregled na temelju PDF računa")
                ukupno_pdf = df['PDF_Naplaćeni_Iznos'].sum()
                ugovoreno_iznos = df['Ugovoreno_Ukupno'].sum()
                preplata = ukupno_pdf - ugovoreno_iznos
                
                col_s1, col_s2, col_s3 = st.columns(3)
                col_s1.metric("Naplaćeno po PDF Računu", f"{ukupno_pdf:,.2f} €")
                col_s2.metric("Trebalo po Ugovoru", f"{ugovoreno_iznos:,.2f} €")
                col_s3.metric("Ukupna preplata / Višak", f"{preplata:,.2f} €")
                
                zbirna_tablica = pd.DataFrame({
                    "Kategorija": ["Osnovni prijevoz", f"Gorivo ({dodatak_gorivo_pct}%)", "SVEUKUPNO"],
                    "PDF Račun (€)": [ukupno_pdf * 0.88, ukupno_pdf * 0.12, ukupno_pdf],
                    "Ugovor (€)": [ugovoreno_iznos * 0.88, ugovoreno_iznos * 0.12, ugovoreno_iznos]
                })
                st.table(zbirna_tablica)
                st.download_button("📥 Preuzmi 'Zbirni pregled' (CSV)", konvertiraj_u_csv(zbirna_tablica), "zbirni_pregled_pdf.csv", "text/csv")
                
            # 5. Sirovi PDF tekst
            with tab5:
                st.subheader("Ekstrahirani tekst iz PDF računa")
                st.text_area("Sadržaj PDF-a", pdf_tekst, height=400)

    except Exception as e:
        st.error(f"Greška kod obrade PDF-a: {e}")
else:
    st.info("Molimo učitajte PDF specifikaciju računa u bočnoj traci kako biste pokrenuli reviziju.")
