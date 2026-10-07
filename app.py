import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import pypdf
import re
import io

# Konfiguracija stranice
st.set_page_config(
    page_title="Detaljna Revizija Računa - G. Englmayer & Ugovor",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Sustav za Detaljnu Reviziju Logističkih Računa (PDF + Ugovorni Cjenik)")
st.markdown("Automatska usporedba stavki iz PDF specifikacije računa, ugovorenih zona, rokova isporuke i težinskih razreda paleta prema ugovoru br. OF 002/2026.")

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

st.sidebar.header("2. Učitavanje dokumenata")
uploaded_pdf = st.sidebar.file_uploader("Učitaj PDF specifikaciju računa", type=["pdf"])
uploaded_excel = st.sidebar.file_uploader("Učitaj Excel/CSV bazu (za gradove, ZIP i težine)", type=["xlsx", "xls", "csv"])

if uploaded_pdf is not None:
    try:
        # Parsiranje PDF specifikacije
        reader = pypdf.PdfReader(uploaded_pdf)
        pdf_tekst = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                pdf_tekst += t + "\n"
        
        # Ekstrakcija stavki iz PDF-a (LA-NR, Datum, Referenca, Iznos)
        pdf_entries = []
        for line in pdf_tekst.split('\n'):
            match = re.search(r'(EP-\d+)\s+(\d{2}\.\d{2}\.\d{4}\.)\s+(\d+)?\s*(\d+%\s+)?([\d\.,]+)', line)
            if match:
                shpt_id = match.group(1)
                date_str = match.group(2)
                ref = match.group(3) if match.group(3) else "N/A"
                amount_str = match.group(5).replace('.', '').replace(',', '.')
                try:
                    amount = float(amount_str)
                    pdf_entries.append({
                        'Route ID': shpt_id, 
                        'PDF_Datum_Isporuke': date_str, 
                        'PDF_Referenca': ref, 
                        'PDF_Naplaćeni_Iznos': amount
                    })
                except:
                    pass
        
        df_pdf = pd.DataFrame(pdf_entries)
        st.success(f"📄 PDF specifikacija uspješno učitana! Pronađeno stavki: {len(df_pdf)}")
        
        # Spajanje s Excel bazom pošiljaka (za detalje poput primatelja, grada, ZIP-a, težine, CLL, tipa palete)
        if uploaded_excel is not None:
            if uploaded_excel.name.endswith('.csv'):
                df_excel = pd.read_csv(uploaded_excel)
            else:
                df_excel = pd.read_excel(uploaded_excel)
            
            if 'Shpt.id' in df_excel.columns:
                df_excel = df_excel.dropna(subset=['Shpt.id']).copy()
            
            # Spajamo preko Route ID / Shpt.id
            if 'Route ID' in df_excel.columns and 'Route ID' in df_pdf.columns:
                df = pd.merge(df_pdf, df_excel, on='Route ID', how='left')
            else:
                df = df_pdf
                st.warning("Nije pronađen poklapajući stupac 'Route ID', prikazuju se osnovni podaci iz PDF-a.")
        else:
            df = df_pdf
            st.info("💡 Savjet: Učitajte i Excel/CSV tablicu u sidebaru kako bi sustav povukao primatelje, gradove, ZIP kodove, težine i broj paleta (CLL).")

        if st.button("Pokreni detaljnu reviziju (PDF + Ugovor)"):
            
            # --- UGOVORNA LOGIKA I PRAVILA ---
            
            # 1. Zona prema ZIP-u i gradu
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
            
            # 2. Rok isporuke (radni dani) i dopušteni ugovorni rok
            df['Dopušteni_Rok_Dana'] = df['Izracunata_Zona'].apply(lambda z: 1 if z == "Zona 1" else (2 if z in ["Zona 2", "Zona 4", "Zona 5"] else 3))
            
            # 3. Težina po paleti (Weight / CLL)
            def izracunaj_tezinu_po_paleti(row):
                cll = row.get('CLL', 1)
                weight = row.get('Weight', 0)
                if pd.isna(cll) or cll <= 0:
                    cll = 1
                return weight / cll

            df['Tezina_Po_Paleti'] = df.apply(izracunaj_tezinu_po_paleti, axis=1)

            # 4. Ugovorena cijena palete prema cjeniku
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
            df['Razlika_Cijene'] = df['PDF_Naplaćeni_Iznos'] - df['Ugovoreno_Ukupno']
            
            total_shipments = len(df)
            
            # Tabovi izvještaja
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "1. PDF i Ugovorne Cijene", 
                "2. Usporedba po Primateljima", 
                "3. Rokovi Isporuke i Zone", 
                "4. Preplate i Višak", 
                "5. Sirovi PDF Tekst"
            ])
            
            def konvertiraj_u_csv(data_frame):
                return data_frame.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            # 1. PDF i Ugovorne Cijene
            with tab1:
                cols_prikaz = [c for c in ['Route ID', 'PDF_Referenca', 'PDF_Datum_Isporuke', 'consignee', 'city CN', 'Izracunata_Zona', 'Weight', 'CLL', 'PDF_Naplaćeni_Iznos', 'Ugovoreno_Ukupno', 'Razlika_Cijene'] if c in df.columns]
                st.subheader("Usporedba naplaćenog iznosa s PDF računa i ugovornog izračuna")
                st.dataframe(df[cols_prikaz].head(25))
                st.download_button("📥 Preuzmi tablicu (CSV)", konvertiraj_u_csv(df[cols_prikaz]), "pdf_vs_ugovor_detaljno.csv", "text/csv")
            
            # 2. Usporedba po primateljima
            with tab2:
                st.subheader("Financijska analiza agregirana po primateljima")
                if 'consignee' in df.columns:
                    agregirano = df.groupby('consignee').agg({
                        'Route ID': 'count',
                        'PDF_Naplaćeni_Iznos': 'sum',
                        'Ugovoreno_Ukupno': 'sum',
                        'Razlika_Cijene': 'sum'
                    }).reset_index().rename(columns={'Route ID': 'Broj pošiljaka'})
                    st.dataframe(agregirano)
                    st.download_button("📥 Preuzmi izvještaj po primateljima (CSV)", konvertiraj_u_csv(agregirano), "primatelji_analiza.csv", "text/csv")
                else:
                    st.info("Za grupiranje po primateljima potrebno je učitati i Excel/CSV datoteku.")

            # 3. Rokovi isporuke i zone
            with tab3:
                st.subheader("Provjera ugovorenih rokova isporuke po zonama")
                cols_rokovi = [c for c in ['Route ID', 'PDF_Datum_Isporuke', 'consignee', 'city CN', 'Izracunata_Zona', 'Dopušteni_Rok_Dana'] if c in df.columns]
                st.dataframe(df[cols_rokovi].head(25))
                st.download_button("📥 Preuzmi izvještaj o rokovima (CSV)", konvertiraj_u_csv(df[cols_rokovi]), "rokovi_i_zone.csv", "text/csv")

            # 4. Preplate i višak
            with tab4:
                st.subheader("Izdvojene preplate (gdje je PDF iznos veći od ugovorenog)")
                preplate_df = df[df['Razlika_Cijene'] > 0].sort_values(by='Razlika_Cijene', ascending=False)
                st.dataframe(preplate_df[cols_prikaz].head(20))
                st.download_button("📥 Preuzmi preplate (CSV)", konvertiraj_u_csv(preplate_df[cols_prikaz]), "preplate_detaljno.csv", "text/csv")

            # 5. Sirovi PDF tekst
            with tab5:
                st.subheader("Izvorni tekst iz PDF specifikacije")
                st.text_area("Sadržaj PDF-a", pdf_tekst, height=400)

    except Exception as e:
        st.error(f"Došlo je do pogreške prilikom obrade PDF-a: {e}")
else:
    st.info("Molimo učitajte PDF specifikaciju računa u bočnoj traci kako biste pokrenuli reviziju.")
