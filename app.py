import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import pypdf
import re
import math

# Konfiguracija stranice
st.set_page_config(
    page_title="Revizija Računa iz PDF-a - G. Englmayer",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Sustav za Reviziju Logističkih Računa (PDF + Službeni Ugovorni Cjenik OF 002/2026)")
st.markdown("Direktna analiza s prisilnim čitanjem iz pariteta i pametnim mapiranjem gradova.")

# Prošireni službeni rječnik hrvatskih gradova i poštanskih brojeva
HR_GRADOVI_ZIP = {
    'zagreb': 10000, 'split': 21000, 'rijeka': 51000, 'osijek': 31000,
    'zadar': 23000, 'pula': 52100, 'slavonski brod': 35000, 'karlovac': 47000,
    'varaždin': 42000, 'šibenik': 22000, 'sisak': 44000, 'vinkovci': 32100,
    'velika gorica': 10410, 'dubrovnik': 20000, 'bjelovar': 43000,
    'koprivnica': 48000, 'vukovar': 32000, 'požega': 34000, 'đakovo': 31400,
    'samobor': 10430, 'čakovec': 40000, 'kutina': 44320, 'rovinj': 52210,
    'makarska': 21300, 'metković': 20350, 'imotski': 21260, 'ploče': 20340,
    'korčula': 20260, 'zaprešić': 10290, 'sveta nedelja': 10431, 'belišće': 31551,
    'valpovo': 31550, 'našice': 31500, 'crikvenica': 51260, 'poreč': 52440,
    'umag': 52470, 'labin': 52220, 'pazin': 52000, 'senj': 53270, 'gospić': 53000,
    'ogulin': 47300, 'dugo selo': 10370, 'vrbovec': 10340, 'jastrebarsko': 10450
}

# Sidebar - Parametri obračuna
st.sidebar.header("1. Ugovorni parametri")
cijena_goriva = st.sidebar.number_input("Prosječna cijena dizel goriva (€ bez PDV-a):", value=1.87, step=0.01)

def izracunaj_dodatak_gorivo(cijena):
    osnova = 1.46
    korak = 0.07
    if cijena <= osnova:
        return 0
    else:
        razlika = cijena - osnova
        tocan_iznos = (razlika / korak) * 1.0
        return math.ceil(tocan_iznos)

dodatak_gorivo_pct = izracunaj_dodatak_gorivo(cijena_goriva)
st.sidebar.info(f"Izračunati dodatak za gorivo (baza 1.46 €, zaokruženo naviše): **{dodatak_gorivo_pct}%**")

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
        
        # 1. Ekstrakcija ukupnih iznosa i stavki iz PDF-a
        pdf_iznosi = {}
        pdf_gorivo_iznosi = {}
        pdf_osnovna_iznosi = {}
        
        trenutni_ep = None
        for line in pdf_tekst.split('\n'):
            m_ep = re.search(r'(EP-\d+)', line)
            if m_ep:
                trenutni_ep = m_ep.group(1)
            
            if trenutni_ep:
                if any(kw in line for kw in ["Dizel", "dizel", "gorivo", "Gorivo"]):
                    m_amt = re.search(r'([\d\.,]+)\s*$', line)
                    if m_amt:
                        try:
                            val = float(m_amt.group(1).replace('.', '').replace(',', '.'))
                            pdf_gorivo_iznosi[trenutni_ep] = round(pdf_gorivo_iznosi.get(trenutni_ep, 0.0) + val, 2)
                        except:
                            pass
                elif any(kw in line for kw in ["Roba", "roba", "prijevoz", "Prijevoz", "zona", "Zona"]):
                    m_amt = re.search(r'([\d\.,]+)\s*$', line)
                    if m_amt:
                        try:
                            val = float(m_amt.group(1).replace('.', '').replace(',', '.'))
                            pdf_osnovna_iznosi[trenutni_ep] = round(pdf_osnovna_iznosi.get(trenutni_ep, 0.0) + val, 2)
                        except:
                            pass

            match = re.search(r'(EP-\d+)\s+(\d{2}\.\d{2}\.\d{4}\.)\s+(\d+)?\s*(\d+%\s+)?([\d\.,]+)', line)
            if match:
                shpt_id = match.group(1)
                amount_str = match.group(5).replace('.', '').replace(',', '.')
                try:
                    pdf_iznosi[shpt_id] = round(float(amount_str), 2)
                except:
                    pass

        # 2. PRISILNO ČITANJE IZ PARITETA S PAMETNIM PREPOZNAVANJEM
        redci_paleta = []
        blokovi_naloga = re.split(r'(Datum naloga:\s*\d{2}\.\d{2}\.\d{4}\.)', pdf_tekst)
        
        for b_idx in range(1, len(blokovi_naloga), 2):
            b_meta = blokovi_naloga[b_idx]
            b_sadrzaj = blokovi_naloga[b_idx + 1] if (b_idx + 1) < len(blokovi_naloga) else ""
            
            p_nalog_tekst = b_meta + b_sadrzaj
            
            m_meta_info = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4}\.)\s*Pošiljka:\s*([^\s]+)\s*LA-ID:\s*(EP-\d+)', p_nalog_tekst)
            if not m_meta_info:
                continue
                
            trenutni_datum_naloga = m_meta_info.group(1)
            trenutni_shpt = m_meta_info.group(3)
            
            m_ref = re.search(r'Referenca:\s*([^\s]+)', p_nalog_tekst)
            trenutni_ref = m_ref.group(1) if m_ref else "N/A"
            
            m_isporuka = re.search(r'Datum isporuke:\s*(\d{2}\.\d{2}\.\d{4})', p_nalog_tekst)
            trenutni_datum_isporuke = m_isporuka.group(1) if m_isporuka else None

            # PRISILNO IZVLAČENJE IZ RETKA S PARITETOM
            trenutni_grad = "Nepoznato"
            trenutni_zip = 0
            
            paritet_linija = ""
            for line in p_nalog_tekst.split('\n'):
                if "Paritet" in line and "istovareno" in line.lower():
                    paritet_linija = line.lower()
                    break
            
            if not paritet_linija:
                for line in p_nalog_tekst.split('\n'):
                    if "istovareno" in line.lower():
                        paritet_linija = line.lower()
                        break

            p_tekst_ciyi = paritet_linija if paritet_linija else p_nalog_tekst.lower()
            
            pronadeno = False
            for grad_naziv, z_broj in HR_GRADOVI_ZIP.items():
                if grad_naziv in p_tekst_ciyi:
                    trenutni_grad = grad_naziv.capitalize()
                    trenutni_zip = z_broj
                    pronadeno = True
                    break
            
            if not pronadeno and paritet_linija:
                dio_nakon = re.split(r'istovareno', paritet_linija, flags=re.IGNORECASE)
                if len(dio_nakon) > 1:
                    c_kandidat = dio_nakon[1].strip().split()
                    if c_kandidat:
                        c_ista = c_kandidat[0].strip(".,").capitalize()
                        if c_ista.lower() not in ["na", "mjesto", "dpu"]:
                            trenutni_grad = c_ista

            for line in p_nalog_tekst.split('\n'):
                m_paleta = re.search(r'^(.*?)\s+([\d\.,]+)(\d)\s+(EWP|FP|OWP)', line.strip())
                if m_paleta:
                    raw_oznaka = m_paleta.group(1).replace('Handelsware', '').strip()
                    oznaka_broj = raw_oznaka if raw_oznaka else "Standardna pošiljka"
                    
                    masa_str = m_paleta.group(2).replace('.', '').replace(',', '.')
                    kolicina = int(m_paleta.group(3))
                    tip_palete = m_paleta.group(4)
                    try:
                        masa_kg = float(masa_str)
                        for _ in range(kolicina):
                            redci_paleta.append({
                                'Oznaka_Broj': oznaka_broj,
                                'Referenca_Sustav': trenutni_ref,
                                'LA-ID': trenutni_shpt,
                                'Datum_Naloga': trenutni_datum_naloga,
                                'Datum_Isporuke': trenutni_datum_isporuke,
                                'Grad': trenutni_grad,
                                'ZIP': trenutni_zip,
                                'Masa_Palete_KG': round(masa_kg, 2),
                                'Tip_Palete': tip_palete
                            })
                    except:
                        pass

        df_palete = pd.DataFrame(redci_paleta)
        st.success(f"PDF uspješno učitan! Pronađeno pojedinačnih paleta: {len(df_palete)}")

        if st.button("Pokreni reviziju s točnim zoniranjem"):
            
            def odredi_zonu(row):
                city = str(row.get('Grad', '')).strip().lower()
                zip_val = str(row.get('ZIP', '00000')).zfill(5)
                prva_dva = int(zip_val[:2]) if zip_val[:2].isdigit() else 0
                
                if any(g in city for g in ['makarska', 'imotski', 'ploče', 'metković', 'dubrovnik', 'korčula', 'mokosica']) or prva_dva == 20:
                    return "Zona 6"
                
                if prva_dva == 10:
                    return "Zona 1"
                elif 40 <= prva_dva <= 49:
                    return "Zona 2"
                elif prva_dva in [34, 35, 51]:
                    return "Zona 3"
                elif (31 <= prva_dva <= 33) or prva_dva == 52:
                    return "Zona 4"
                elif (21 <= prva_dva <= 23) or prva_dva == 53:
                    return "Zona 5"
                else:
                    return "Zona 2"

            df_palete['Izracunata_Zona'] = df_palete.apply(odredi_zonu, axis=1)
            df_palete['Dopušteni_Rok_Radnih_Dana'] = df_palete['Izracunata_Zona'].apply(lambda z: 3 if z == "Zona 6" else (1 if z == "Zona 1" else 2))
            
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

            def ugovorena_cijena_palete(row):
                zona = row['Izracunata_Zona']
                tezina = row['Masa_Palete_KG']
                paleta_tip = str(row.get('Tip_Palete', 'FP'))
                
                cjenik = {
                    "Zona 1": [23.0, 25.0, 30.0, 33.0, 38.0],
                    "Zona 2": [26.0, 29.0, 35.0, 39.0, 43.0],
                    "Zona 3": [36.0, 40.0, 45.0, 47.0, 51.0],
                    "Zona 4": [42.0, 47.0, 50.0, 55.0, 65.0],
                    "Zona 5": [44.0, 48.0, 51.0, 56.0, 68.0],
                    "Zona 6": [55.0, 59.0, 63.0, 65.0, 79.0]
                }
                
                if tezina < 300.0: t_idx = 0
                elif tezina < 400.0: t_idx = 1
                elif tezina < 500.0: t_idx = 2
                elif tezina < 600.0: t_idx = 3
                else: t_idx = 4
                
                baza = cjenik.get(zona, cjenik["Zona 2"])[t_idx]
                if paleta_tip.upper() == 'OWP':
                    baza = baza * 1.50
                return round(baza, 2)

            df_palete['Ugovorena_Osnovna_Cijena'] = df_palete.apply(ugovorena_cijena_palete, axis=1)
            df_palete['Ugovorena_Osnovna_Ukupno'] = round(df_palete['Ugovorena_Osnovna_Cijena'], 2)
            df_palete['Ugovoreni_Iznos_Goriva'] = round(df_palete['Ugovorena_Osnovna_Cijena'] * (dodatak_gorivo_pct / 100.0), 2)
            df_palete['Ugovoreno_Paleta_Sa_Gorivom'] = round(df_palete['Ugovorena_Osnovna_Cijena'] + df_palete['Ugovoreni_Iznos_Goriva'], 2)

            # Tabovi izvještaja
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "1. Pregled po Paletama", 
                "2. Provjera Rokova Isporuke (Analitika)", 
                "3. Zbirni Financijski Pregled",
                "4. Usporedba po Pošiljkama (Reference)",
                "5. Preplate po Pošiljkama",
                "6. Skupna Raščlamba (Gorivo vs Dostava)"
            ])
            
            def konvertiraj_u_excel_csv(data_frame):
                df_export = data_frame.copy()
                for col in df_export.select_dtypes(include=['float64', 'float32', 'int64']):
                    df_export[col] = df_export[col].astype(str).str.replace('.', ',', regex=False)
                return df_export.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            with tab1:
                st.subheader(f"Popis svih paleta izvađenih iz PDF-a ({len(df_palete)} stavki)")
                st.dataframe(df_palete)
                st.download_button("📥 Preuzmi palete (CSV za Excel)", konvertiraj_u_excel_csv(df_palete), "palete_iz_pdf-a.csv", "text/csv")
            
            with tab2:
                st.subheader("Analitički izvještaj: Učinkovitost i točnost rokova dostave")
                ukupno_stavki = len(df_palete)
                broj_u_roku = len(df_palete[df_palete['Status_Roka'] == 'U roku'])
                broj_izvan_rok = len(df_palete[df_palete['Status_Roka'] == 'Izvan roka'])
                pct_u_roku = (broj_u_roku / ukupno_stavki) * 100 if ukupno_stavki > 0 else 0
                pct_izvan_rok = (broj_izvan_rok / ukupno_stavki) * 100 if ukupno_stavki > 0 else 0
                
                kpi1, kpi2, kpi3 = st.columns(3)
                kpi1.metric("U roku (Uspješnost)", f"{pct_u_roku:.2f}%", f"{broj_u_roku} paleta")
                kpi2.metric("Izvan roka (Kašnjenje)", f"{pct_izvan_rok:.2f}%", f"{broj_izvan_rok} paleta")
                kpi3.metric("Ukupno analizirano", f"{ukupno_stavki} paleta")
                st.markdown("---")
                cols_rok = ['Oznaka_Broj', 'LA-ID', 'Grad', 'ZIP', 'Izracunata_Zona', 'Datum_Naloga', 'Datum_Isporuke', 'Stvarni_Radni_Dani', 'Dopušteni_Rok_Radnih_Dana', 'Status_Roka']
                st.dataframe(df_palete[cols_rok])
                st.download_button("📥 Preuzmi analitiku rokova (CSV za Excel)", konvertiraj_u_excel_csv(df_palete[cols_rok]), "analitika_rokova_isporuke.csv", "text/csv")
                
            with tab3:
                st.subheader("Zbirna rekapitulacija po ugovoru")
                ukupno_ugovor = round(df_palete['Ugovoreno_Paleta_Sa_Gorivom'].sum(), 2)
                c1, c2 = st.columns(2)
                c1.metric("Ukupno paleta u PDF-u", f"{len(df_palete)}")
                c2.metric("Ukupno po službenom ugovornom cjeniku", f"{ukupno_ugovor:,.2f} €")
                
            with tab4:
                st.subheader("Usporedba pošiljaka zbrojenih po LA-ID brojevima")
                df_posiljke = df_palete.groupby(['LA-ID', 'Oznaka_Broj', 'Grad', 'ZIP', 'Izracunata_Zona']).agg(
                    Datum_Naloga=('Datum_Naloga', 'first'),
                    Datum_Isporuke=('Datum_Isporuke', 'max'),
                    Broj_Paleta=('Masa_Palete_KG', 'count'),
                    Ukupna_Masa_KG=('Masa_Palete_KG', 'sum'),
                    Ugovoreno_Osnovna_EUR=('Ugovorena_Osnovna_Ukupno', 'sum'),
                    Ugovoreno_Gorivo_EUR=('Ugovoreni_Iznos_Goriva', 'sum'),
                    Ugovoreno_Ukupno_EUR=('Ugovoreno_Paleta_Sa_Gorivom', 'sum')
                ).reset_index()
                
                df_posiljke['Ukupna_Masa_KG'] = df_posiljke['Ukupna_Masa_KG'].round(2)
                df_posiljke['Ugovoreno_Osnovna_EUR'] = df_posiljke['Ugovoreno_Osnovna_EUR'].round(2)
                df_posiljke['Ugovoreno_Gorivo_EUR'] = df_posiljke['Ugovoreno_Gorivo_EUR'].round(2)
                df_posiljke['Ugovoreno_Ukupno_EUR'] = df_posiljke['Ugovoreno_Ukupno_EUR'].round(2)
                
                df_posiljke['Naplaćeno_Po_PDF_EUR'] = df_posiljke['LA-ID'].map(pdf_iznosi).fillna(0.0).round(2)
                df_posiljke['Razlika (Naplaćeno - Ugovoreno)'] = round(df_posiljke['Naplaćeno_Po_PDF_EUR'] - df_posiljke['Ugovoreno_Ukupno_EUR'], 2)
                
                st.dataframe(df_posiljke)
                st.download_button("📥 Preuzmi usporedbu pošiljaka (CSV za Excel)", konvertiraj_u_excel_csv(df_posiljke), "usporedba_po_posiljkama.csv", "text/csv")

            with tab5:
                st.subheader("Izdvojene preplate (gdje je naplaćeni iznos veći od ugovornog)")
                if 'df_posiljke' in locals():
                    df_preplate = df_posiljke[df_posiljke['Razlika (Naplaćeno - Ugovoreno)'] > 0].sort_values(by='Razlika (Naplaćeno - Ugovoreno)', ascending=False)
                    st.dataframe(df_preplate)
                    st.download_button("📥 Preuzmi preplate po pošiljkama (CSV za Excel)", konvertiraj_u_excel_csv(df_preplate), "preplate_po_posiljkama.csv", "text/csv")
                else:
                    st.info("Pregledajte prvo tab 4 za izračun.")

            with tab6:
                st.subheader("Skupni financijski pregled komponenti (Cijeli račun)")
                if 'df_posiljke' in locals():
                    tot_naplaceno_osnovna = round(sum(pdf_osnovna_iznosi.values()), 2)
                    tot_ugovoreno_osnovna = round(df_posiljke['Ugovoreno_Osnovna_EUR'].sum(), 2)
                    
                    tot_naplaceno_gorivo = round(sum(pdf_gorivo_iznosi.values()), 2)
                    tot_ugovoreno_gorivo = round(df_posiljke['Ugovoreno_Gorivo_EUR'].sum(), 2)
                    
                    skupni_podaci = [{
                        'Komponenta': 'Osnovna cijena prijevoza (Dostava)',
                        'Naplaćeno ukupno (€)': tot_naplaceno_osnovna,
                        'Ugovoreno ukupno (€)': tot_ugovoreno_osnovna,
                        'Razlika (Naplaćeno - Ugovoreno) (€)': round(tot_naplaceno_osnovna - tot_ugovoreno_osnovna, 2)
                    }, {
                        'Komponenta': 'Dizel dodatak (Gorivo)',
                        'Naplaćeno ukupno (€)': tot_naplaceno_gorivo,
                        'Ugovoreno ukupno (€)': tot_ugovoreno_gorivo,
                        'Razlika (Naplaćeno - Ugovoreno) (€)': round(tot_naplaceno_gorivo - tot_ugovoreno_gorivo, 2)
                    }, {
                        'Komponenta': 'UKUPNO SVEUKUPNO',
                        'Naplaćeno ukupno (€)': round(tot_naplaceno_osnovna + tot_naplaceno_gorivo, 2),
                        'Ugovoreno ukupno (€)': round(tot_ugovoreno_osnovna + tot_ugovoreno_gorivo, 2),
                        'Razlika (Naplaćeno - Ugovoreno) (€)': round((tot_naplaceno_osnovna + tot_naplaceno_gorivo) - (tot_ugovoreno_osnovna + tot_ugovoreno_gorivo), 2)
                    }]
                    
                    df_skupno = pd.DataFrame(skupni_podaci)
                    st.table(df_skupno)
                    st.download_button("📥 Preuzmi skupnu rekapitulaciju (CSV za Excel)", konvertiraj_u_excel_csv(df_skupno), "skupna_rasclamba_gorivo_dostava.csv", "text/csv")
                else:
                    st.info("Pregledajte prvo tab 4 za izračun.")

    except Exception as e:
        st.error(f"Greška kod obrade PDF-a: {e}")
else:
    st.info("Molimo učitajte PDF specifikaciju računa u bočnoj traci.")
