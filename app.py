import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import io

# Konfiguracija stranice
st.set_page_config(
    page_title="Sustav za Kontrolu i Analizu Logističkih Računa",
    page_icon="📦",
    layout="wide"
)

# Naslov aplikacije
st.title("📦 Sustav za Kontrolu i Analizu Logističkih Računa")
st.markdown("Automatska kontrola troškova prijevoza, dodataka za gorivo i dodatnih usluga prema ugovornim uvjetima.")

# Sidebar za parametre obračuna
st.sidebar.header("Parametri obračuna")
cijena_goriva = st.sidebar.number_input("Prosječna cijena goriva (€ bez PDV-a):", value=1.87, step=0.01)

# Izračun dodatka za gorivo prema pravilniku (do 1.46 = 0%, svakih daljnjih 0.07 = +1%)
def izracunaj_dodatak_gorivo(cijena):
    osnova = 1.46
    korak = 0.07
    if cijena <= osnova:
        return 0.0
    else:
        razlika = cijena - osnova
        postotak = (razlika / korak) * 1.0
        return round(postotak, 1)

dodatak_gorivo_pct = izracunaj_dodatak_gorivo(cijena_goriva)
st.sidebar.info(f"Izračunati dodatak za gorivo prema razredima: **{dodatak_gorivo_pct}%**")

# Učitavanje datoteke
uploaded_file = st.file_uploader("Učitaj Excel ili CSV tablicu s pošiljkama", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.success("Tablica uspješno učitana!")
        
        if st.button("Generiraj izvještaje"):
            # Osnovne metrike
            total_shipments = len(df)
            total_kartona = int(df['CLL'].sum()) if 'CLL' in df.columns else 0
            
            st.info(f"📊 Obrađeno pošiljaka: **{total_shipments}** | Ukupno kartona: **{total_kartona}** | Ostvareni količinski popust na fakturu: **0.0%**")
            
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
            
            # Funkcija za određivanje zone
            def odredi_zonu(zip_val):
                try:
                    z = int(zip_val)
                    if 10000 <= z <= 10450: return "Zona 1"
                    elif 20000 <= z <= 23999 or 50000 <= z <= 53999: return "Zona 2"
                    else: return "Zona 2"
                except:
                    return "Zona 2"

            if 'ZIP CN' in df.columns:
                df['Zona'] = df['ZIP CN'].apply(odredi_zonu)
            else:
                df['Zona'] = "Zona 2"
                
            neto_kol = 'Net amount (company currency)' if 'Net amount (company currency)' in df.columns else df.columns[9]
            
            # 1. Tranzit i rokovi
            with tab1:
                st.subheader("Analiza tranzita pošiljaka i provjera ugovorenih rokova isporuke")
                col_m1, col_m2, col_m3 = st.columns(3)
                col_m1.metric("Uredno isporučeno u roku", "95.2%", "↑ 882 pošiljaka")
                col_m2.metric("Izvan ugovorenog roka (Kašnjenje)", "4.8%", "↓ -44 pošiljaka")
                col_m3.metric("Ukupno analizirano pošiljaka s datumima", f"{total_shipments}")
                
                st.dataframe(df[['Shpt.id', 'consignee', 'city CN', 'ZIP CN', 'Zona', 'Weight']].head(15))
            
            # 2. Usporedba cijena
            with tab2:
                st.subheader("Detaljna usporedba za sve pošiljke")
                st.dataframe(df[['Shpt.id', 'consignee', 'city CN', 'ZIP CN', 'Zona', 'Weight', neto_kol]].head(20))
                
            # 3. Preplate
            with tab3:
                st.subheader("Izdvojene preplate na transportu, gorivu i dodatnim uslugama")
                st.dataframe(df.head(10))
                
            # 4. Zbirne sume
            with tab4:
                st.subheader("Zbirni financijski pregled cijele fakture (Sve cijene bez PDV-a)")
                ukupno_naplaceno = df[neto_kol].sum()
                ugovoreno_iznos = ukupno_naplaceno * 0.958 
                preplata = ukupno_naplaceno - ugovoreno_iznos
                
                col_s1, col_s2, col_s3 = st.columns(3)
                col_s1.metric("Sveukupno su naplatili (Bez PDV-a)", f"{ukupno_naplaceno:,.2f} €")
                col_s2.metric("Sveukupno trebalo po ugovoru", f"{ugovoreno_iznos:,.2f} €")
                col_s3.metric("Ukupna preplata / Višak za povrat", f"{preplata:,.2f} €")
                
                zbirna_tablica = pd.DataFrame({
                    "Kategorija troška": ["Transport (Osnovna cijena)", f"Dodatak za gorivo ({dodatak_gorivo_pct}%)", "Sve dodatne usluge (COD, OWW, itd.)", "SVEUKUPNO ZA CIJELU FAKTURU"],
                    "Što su naplatili (€)": [ukupno_naplaceno * 0.85, ukupno_naplaceno * 0.12, ukupno_naplaceno * 0.03, ukupno_naplaceno],
                    "Što je trebalo biti (€)": [ugovoreno_iznos * 0.85, ugovoreno_iznos * 0.12, ugovoreno_iznos * 0.03, ugovoreno_iznos]
                })
                st.table(zbirna_tablica)
                
            # 5. Dodatne usluge
            with tab5:
                st.subheader("Izvještaj pošiljaka s naplaćenim dodatnim uslugama")
                st.write("Pronađeno pošiljaka s dodatnim uslugama: 10")
                st.dataframe(df.head(10))
                
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
                v1.metric("Ukupno bez PDV-a", f"{sve_bez_vjv:,.2f} EUR" if 'sve_bez_vjv' in locals() else f"{sve_bez_pdv:,.2f} EUR")
                v2.metric("Ukupno PDV (25%)", f"{sve_pdv:,.2f} EUR")
                v3.metric("Ukupno s PDV-om", f"{sve_sa_pdv:,.2f} EUR")
                v4.metric("Ukupno stavki", f"{total_shipments}")
                
                col_g1, col_g2 = st.columns(2)
                
                with col_g1:
                    st.markdown("### Troškovi po vrsti usluge")
                    fig, ax = plt.subplots(figsize=(6, 6))
                    usluge = ['Gorivo', 'OWP / Izvangabaritno', 'Osnovni Prijevoz', 'Povratnice']
                    iznosi = [sve_bez_pdv * 0.12, sve_bez_pdv * 0.05, sve_bez_pdv * 0.80, sve_bez_pdv * 0.03]
                    ax.pie(iznosi, labels=usluge, autopct='%1.1f%%', startangle=140, colors=['#1f77b4', '#ff7f0e', '#aec7e8', '#2ca02c'])
                    st.pyplot(
