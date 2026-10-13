import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Specifikacija Parser", layout="wide")

st.title("📦 Englmayer Specifikacija Parser (Službeni Račun)")
st.write("Učitajte službenu Excel specifikaciju (npr. `3500131683-CSV.xlsx`) za automatsko čitanje svih stavaka, pošiljaka i financijskih iznosa.")

uploaded_file = st.file_uploader("Odaberite Excel specifikaciju (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada i spajanje podataka u tijeku..."):
        
        # Čitanje izvornog Excel izvoza specifikacije
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        processed_rows = []
        
        for idx, row in df_raw.iterrows():
            # Provjera osnovnih podataka pošiljke
            zag = row.get('Sendungs-ID', 'N/A')
            datum_izrade = row.get('Beleg-Datum', '')
            datum_dostave = row.get('Leistungsdatum', '')
            la_id = row.get('Beleg-Id', '') # ili odgovarajući ID
            referenca = row.get('Referenz', '')
            
            posiljatelj = row.get('Name Absender', 'Makromikro Grupa d.o.o.')
            primatelj = row.get('Name Empfänger', 'N/A')
            adresa_primatelja = row.get('Straße Empfänger', '')
            grad_primatelja = row.get('Ort Empfänger', '')
            
            otp_broj = row.get('Zeichen/Nr.', 'N/A')
            kolicina = row.get('Anz', 1)
            if pd.isna(kolicina) or kolicina <= 0:
                kolicina = 1
            else:
                kolicina = int(kolicina)
                
            pakiranje = row.get('Verp', row.get('Verpackung', 'EWP'))
            masa = row.get('Gewicht', row.get('Gew', 0.0))
            cbm = row.get('CBM', 0.0)
            
            # Financijski dio (franko, dizel dodatak, ukupno)
            fracht = row.get('Frachten (€)', row.get('Frachten', 0.0))
            diesel = row.get('Dieselzuschlag (€)', row.get('Dieselzuschlag', 0.0))
            ukupno_eur = row.get('Nettobetrag (€)', row.get('Nettobetrag', 0.0))
            
            if pd.isna(ukupno_eur) or ukupno_eur == 0:
                ukupno_eur = float(fracht or 0) + float(diesel or 0)

            # Ako pošiljka ima količinu > 1, možemo je raspisati ili zadržati zapis
            processed_rows.append({
                "Datum naloga": datum_izrade,
                "Datum dostave": datum_dostave,
                "Pošiljka": zag,
                "Referenca": referenca,
                "Pošiljatelj": posiljatelj,
                "Primatelj": primatelj,
                "Grad primatelja": grad_primatelja,
                "OTP / Broj naloga": otp_broj,
                "Tip palete": pakiranje,
                "Količina (Kom)": kolicina,
                "Masa (KG)": masa,
                "CBM": cbm,
                "Prijevoz (€)": fracht,
                "Dizel dodatak (€)": diesel,
                "Ukupno neto (€)": ukupno_eur
            })

        df_out = pd.DataFrame(processed_rows)

        if not df_out.empty:
            st.success(f"Uspješno obrađeno redaka specifikacije: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = uploaded_file.name.replace(".xlsx", "_uredna_specifikacija.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Specifikacija')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni uređeni Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke u priloženoj datoteci.")
