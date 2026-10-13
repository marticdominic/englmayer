import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Konačni Spojeni Sustav", layout="wide")

st.title("📦 Englmayer Spojeni Sustav (Sve pošiljke i reference)")
st.write("Učitajte službenu Excel specifikaciju (`3500131683-CSV.xlsx`) izvedenu iz PDF-a. Aplikacija detaljno raspisuje svaku pošiljku i svaku referencu/paletu.")

uploaded_file = st.file_uploader("Odaberite Excel specifikaciju (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Detaljna obrada i raspisivanje svih pošiljaka i referenci u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        detailed_rows = []
        
        # Prolazimo kroz svaki redak u izvozu specifikacije (gdje svaki redak ili grupa predstavlja stavku/referencu)
        for idx, row in df_raw.iterrows():
            zag = row.get('Sendungs-ID', 'N/A')
            if pd.isna(zag):
                continue
                
            datum_izrade = row.get('Beleg-Datum', '')
            datum_dostave = row.get('Leistungsdatum', '')
            referenca = row.get('Referenz', '')
            
            posiljatelj = row.get('Name Absender', 'Makromikro Grupa d.o.o.')
            primatelj = row.get('Name Empfänger', 'N/A')
            grad_primatelja = row.get('Ort Empfänger', '')
            
            # Podaci o paleti / stavci
            otp_broj = row.get('Zeichen/Nr.', 'N/A')
            kolicina = row.get('Anz', 1)
            if pd.isna(kolicina) or kolicina <= 0:
                kolicina = 1
            else:
                kolicina = int(kolicina)
                
            pakiranje = row.get('Verp', row.get('Verpackung', 'EWP'))
            masa = row.get('Gewicht', row.get('Gew', 0.0))
            cbm = row.get('CBM', 0.0)
            
            # Financije
            fracht = row.get('Frachten', 0.0)
            diesel = row.get('Dieselzuschlag', 0.0)
            ukupno_eur = row.get('Nettobetrag', 0.0)
            
            if pd.isna(ukupno_eur) or ukupno_eur == 0:
                ukupno_eur = float(fracht or 0) + float(diesel or 0)

            detailed_rows.append({
                "Datum izrade": datum_izrade,
                "Datum dostave": datum_dostave,
                "Pošiljka": zag,
                "Referenca": referenca,
                "Pošiljatelj": posiljatelj,
                "Primatelj": primatelj,
                "Grad primatelja": grad_primatelja,
                "OTP / Oznaka": otp_broj if pd.notna(otp_broj) else "N/A",
                "Tip palete": pakiranje if pd.notna(pakiranje) else "EWP",
                "Količina (Kom)": kolicina,
                "Masa (KG)": masa if pd.notna(masa) else 0.0,
                "CBM": cbm if pd.notna(cbm) else 0.0,
                "Prijevoz (€)": fracht if pd.notna(fracht) else 0.0,
                "Dizel dodatak (€)": diesel if pd.notna(diesel) else 0.0,
                "Ukupno neto (€)": ukupno_eur if pd.notna(ukupno_eur) else 0.0
            })

        df_out = pd.DataFrame(detailed_rows)

        if not df_out.empty:
            st.success(f"Uspješno obrađeno i raspisano ukupno stavki/referenci: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = "englmayer_sve_stavke_i_reference.xlsx"
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Sve_Reference')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni Excel sa svim stavkama (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke u priloženoj datoteci.")
