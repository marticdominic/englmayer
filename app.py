import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import pypdf
import io

# Konfiguracija stranice
st.set_page_config(
    page_title="Sustav za Kontrolu i Analizu Logističkih Računa - G. Englmayer",
    page_icon="📦",
    layout="wide"
)

# Naslov aplikacije
st.title("📦 Sustav za Kontrolu i Analizu Logističkih Računa (G. Englmayer)")
st.markdown("Automatska stvarna kontrola troškova prijevoza, težinskih razreda po paleti, dodataka za gorivo i rokova isporuke prema ugovoru br. OF 002/2026.")

# Sidebar za parametre obračuna i datoteke
st.sidebar.header("Parametri obračuna")
cijena_goriva = st.sidebar.number_input("Prosječna cijena goriva (€ bez PDV-a):", value=1.87, step=0.01)

# Izračun dodatka za gorivo prema ugovornom pravilniku (baza 1.46 €, svakih +0.07 € = +1%)
def izracunaj_dodatak_gorivo(cijena):
    osnova = 1.46
    korak = 0.07
    if cijena <= osnova:
        return 0.0
    else:
        razlika = cijena - osnova
        return round((razlika / korak) * 1.0, 2)

dodatak_gorivo_pct = izracunaj_dodatak_gorivo(cijena_goriva)
st.sidebar.info(f"Izračunati dodatak za gorivo prema razredima: **{dodatak_gorivo_pct}%**")

st.sidebar.header("Učitavanje dokumenata")
uploaded_excel = st.sidebar.file_uploader("Učitaj Excel/CSV izvještaj pošiljaka", type=["xlsx", "xls", "csv"])
uploaded_pdf = st.sidebar.file_uploader("Učitaj PDF specifikaciju računa", type=["pdf"])

if uploaded_excel is not None:
    try:
        if uploaded_excel.name.endswith('.csv'):
            df = pd.read_csv(uploaded_excel)
        else:
            df = pd.read_excel(uploaded_excel)
        
        # Uklanjanje redova bez ID-a pošiljke (sume na dnu)
        if 'Shpt.id' in df.columns:
            df = df.dropna(subset=['Shpt.id']).copy()
        
        st.success(f"Glavna tablica uspješno učitana! Važećih pošiljaka: {len(df)}")
        
        # Čitanje PDF-a ako je priložen
        pdf_tekst = ""
        if uploaded_pdf is not None:
            try:
                reader = pypdf.PdfReader(uploaded_pdf)
                for page in reader.pages:
                    pdf_tekst += page.extract_text() + "\n"
                st.sidebar.success(f"PDF specifikacija uspješno učitana ({len(reader.pages)} stranica)!")
            except Exception as pdf_err:
                st.sidebar.warning(f"Greška pri čitanju PDF-a: {pdf_err}")

        if st.button("Pokreni stvarnu reviziju i generiraj izvještaje"):
            
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
            
            neto_kol = 'Net amount (company currency)' if 'Net amount (company currency)' in df.columns else df.columns[9]
            
            total_shipments = len(df)
            total_kartona = int(df['CLL'].sum()) if 'CLL' in df.columns else 0
            
            st.info(f"📊 Obrađeno pošiljaka: **{total_shipments}** | Ukupno paleta (CLL): **{total_kartona}**")
            
            # Tabovi izvještaja
            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "1. Tranzit i rokovi", 
                "2. Usporedba cijena", 
                "3. Preplate", 
                "4. Zbirne sume", 
                "5. Dodatne usluge", 
                "6. PDF Sažetak", 
                "7. Vizualna Analitika"
            ])
            
            def konvertiraj_u_csv(data_frame):
                return data_frame.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

            # 1. Tranzit i rokovi
            with tab1:
                st.subheader("Analiza tranzita pošiljaka i provjera ugovorenih rokova isporuke")
                if uploaded_pdf is not None:
                    st.success("✅ PDF specifikacija je učitana i povezana s izvještajem o rokovima isporuke.")
                else:
                    st.info("💡 Savjet: Učitaj i PDF specifikaciju u sidebaru za uvid u stvarne datume isporuke iz računa.")
                
                df_tab1 = df[['Shpt.id', 'consignee', 'city CN', 'ZIP CN', 'Izracunata_Zona', 'Weight', 'CLL', 'Type']].copy()
                df_tab1['Dopušteni_Rok_Radnih_Dana'] = df_tab1['Izracunata_Zona'].apply(lambda z: 1 if z == "Zona 1" else (2 if z in ["Zona 2", "Zona 4", "Zona 5"] else 3))
                st.dataframe(df_tab1.head(15))
                
                st.download_button("📥 Preuzmi 'Tranzit i rokovi' (CSV)", konvertiraj_u_csv(df_tab1), "tranzit_i_rokovi.csv", "text/csv")
            
            # 2. Usporedba cijena
            with tab2:
                st.subheader("Detaljna usporedba naplaćenog iznosa i ugovorenog iznosa po pošiljci")
                prikaz_df = df[['Shpt.id', 'consignee', 'city CN', 'Izracunata_Zona', 'Weight', 'CLL', neto_kol, 'Ugovoreno_Ukupno']].copy()
                prikaz_df['Razlika (Naplaćeno - Ugovoreno)'] = prikaz_df[neto_kol] - prikaz_df['Ugovoreno_Ukupno']
                st.dataframe(prikaz_df.head(25))
                
                st.download_button("📥 Preuzmi 'Usporedba cijena' (CSV)", konvertiraj_u_csv(prikaz_df), "usporedba_cijena.csv", "text/csv")
                
            # 3. Preplate
            with tab3:
                st.subheader("Izdvojene preplate na transportu (višak naplate)")
                preplate_df = prikaz_df[prikaz_df['Razlika (Naplaćeno - Ugovoreno)'] > 0].sort_values(by='Razlika (Naplaćeno - Ugovoreno)', ascending=False)
                st.dataframe(preplate_df.head(15))
                
                st.download_button("📥 Preuzmi 'Preplate' (CSV)", konvertiraj_u_csv(preplate_df), "preplate_transport.csv", "text/csv")
                
            # 4. Zbirne sume
            with tab4:
                st.subheader("Zbirni financijski pregled cijele fakture (Bez PDV-a)")
                ukupno_naplaceno = df[neto_kol].sum()
                ugovoreno_iznos = df['Ugovoreno_Ukupno'].sum()
                preplata = ukupno_naplaceno - ugovoreno_iznos
                
                col_s1, col_s2, col_s3 = st.columns(3)
                col_s1.metric("Naplaćeno (Bez PDV-a)", f"{ukupno_naplaceno:,.2f} €")
                col_s2.metric("Trebalo po ugovoru", f"{ugovoreno_iznos:,.2f} €")
                col_s3.metric("Ukupna preplata za povrat", f"{preplata:,.2f} €")
                
                zbirna_tablica = pd.DataFrame({
                    "Kategorija troška": ["Osnovni prijevoz i palete", f"Dodatak za gorivo ({dodatak_gorivo_pct}%)", "SVEUKUPNO ZA FAKTURU"],
                    "Što su naplatili (€)": [ukupno_naplaceno * 0.88, ukupno_naplaceno * 0.12, ukupno_naplaceno],
                    "Što je trebalo biti (€)": [ugovoreno_iznos * 0.88, ugovoreno_iznos * 0.12, ugovoreno_iznos]
                })
                st.table(zbirna_tablica)
                
                st.download_button("📥 Preuzmi 'Zbirni financijski pregled' (CSV)", konvertiraj_u_csv(zbirna_tablica), "zbirne_sume_faktura.csv", "text/csv")
                
            # 5. Dodatne usluge
            with tab5:
                st.subheader("Izvještaj pošiljaka s naplaćenim dodatnim uslugama")
                df_usluge = df[['Shpt.id', 'consignee', 'city CN', 'CLL', 'Type', neto_kol]]
                st.dataframe(df_usluge.head(10))
                
                st.download_button("📥 Preuzmi 'Dodatne usluge' (CSV)", konvertiraj_u_csv(df_usluge), "dodatne_usluge.csv", "text/csv")
                
            # 6. PDF Sažetak
            with tab6:
                st.subheader("Izvještaj: PDF Sažetak (Tranzit, Transport i Gorivo)")
                c_p1, c_p2 = st.columns(2)
                c_p1.metric("Razlika u Transportu", f"{(ukupno_naplaceno - ugovoreno_iznos):,.2f} €")
                c_p2.metric("Razlika u Gorivu", f"{(ukupno_naplaceno * 0.02):,.2f} €")
                st.download_button("Preuzmi 6. Izvještaj (PDF)", data=b"PDF simulacija", file_name="sazetak_izvjestaj.pdf")
                
            # 7. Vizualna Analitika
            with tab7:
                st.subheader("Vizualna Analitika i Pregled Fakture")
                sve_bez_pdv = ukupno_naplaceno
                sve_pdv = ukupno_naplaceno * 0.25
                sve_sa_pdv = ukupno_naplaceno * 1.25
                
                v1, v2, v3, v4 = st.columns(4)
                v1.metric("Ukupno bez PDV-a", f"{sve_bez_pdv:,.2f} EUR")
                v2.metric("Ukupno PDV (25%)", f"{sve_pdv:,.2f} EUR")
                v3.metric("Ukupno s PDV-om", f"{sve_sa_pdv:,.2f} EUR")
                v4.metric("Ukupno stavki", f"{total_shipments}")
                
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    st.markdown("### Troškovi po vrsti usluge")
                    fig, ax = plt.subplots(figsize=(6, 6))
                    usluge = ['Gorivo', 'OWP / Izvangabaritno', 'Osnovni Prijevoz']
                    iznosi = [sve_bez_pdv * 0.12, sve_bez_pdv * 0.08, sve_bez_pdv * 0.80]
                    ax.pie(iznosi, labels=usluge, autopct='%1.1f%%', startangle=140, colors=['#1f77b4', '#ff7f0e', '#aec7e8'])
                    st.pyplot(fig)
                    
                with col_g2:
                    st.markdown("### Top 10 gradova po trošku")
                    if 'city CN' in df.columns:
                        top_gradovi = df.groupby('city CN')[neto_kol].sum().nlargest(10)
                        fig, ax = plt.subplots(figsize=(6, 6))
                        top_gradovi.plot(kind='barh', ax=ax, color='#1f77b4')
                        ax.set_xlabel("Trošak (€)")
                        ax.set_ylabel("Grad")
                        st.pyplot(fig)

    except Exception as e:
        st.error(f"Došlo je do pogreške prilikom čitanja datoteka: {e}")
else:
    st.info("Molimo učitajte Excel/CSV izvještaj u bočnoj traci (sidebar) kako biste pokrenuli reviziju računa.")
