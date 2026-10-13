import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Spojeni Sustav (PDF + Excel)", layout="wide")

st.title("📦 Englmayer Spojeni Sustav (PDF Specifikacija + Excel)")
st.write("Učitajte službenu Excel specifikaciju (npr. `3500131683-CSV.xlsx`). Aplikacija automatski povezuje razdvojene retke, mase, CBM i cijene.")

uploaded_file = st.file_uploader("Odaberite Excel specifikaciju (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Pametno spajanje i popunjavanje podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        processed_rows = []
        
        # Grupiramo po pošiljkama (Sendungs-ID)
        for zag, group in df_raw.groupby('Sendungs-ID'):
            # Uzimamo osnovne podatke iz prve linije grupe
            first_row = group.iloc[0]
            
            datum_izrade = first_row.get('Beleg-Datum', '')
            datum_dostave = first_row.get('Leistungsdatum', '')
            referenca = first_row.get('Referenz', '')
            
            posiljatelj = first_row.get('Name Absender', 'Makromikro Grupa d.o.o.')
            primatelj = first_row.get('Name Empfänger', 'N/A')
            grad_primatelja = first_row.get('Ort Empfänger', '')
            
            fracht = first_row.get('Frachten', 0.0)
            diesel = first_row.get('Dieselzuschlag', 0.0)
            ukupno_eur = first_row.get('Nettobetrag', 0.0)
            
            if pd.isna(ukupno_eur) or ukupno_eur == 0:
                ukupno_eur = float(fracht or 0) + float(diesel or 0)

            # Tražimo vrijednosti koje su možda u kasnijim recima iste pošiljke
            valid_pak = group['Verp'].dropna().tolist() or group['Verpackung'].dropna().tolist()
            valid_gewicht = group['Gewicht'].dropna().tolist() or group['Gew'].dropna().tolist()
            valid_cbm = group['CBM'].dropna().tolist()
            
            default_pak = valid_pak[0] if valid_pak else "EWP"
            default_gewicht = valid_gewicht[0] if valid_gewicht else 0.0
            default_cbm = valid_cbm[0] if valid_cbm else 0.0

            # Prolazimo kroz sve stavke pošiljke
            for _, r in group.iterrows():
                otp_broj = r.get('Zeichen/Nr.', 'N/A')
                anz = r.get('Anz', 1)
                if pd.isna(anz) or anz <= 0:
                    anz = 1
                else:
                    anz = int(anz)
                    
                pak = r.get('Verp', r.get('Verpackung', default_pak))
                masa = r.get('Gewicht', r.get('Gew', default_gewicht))
                cbm = r.get('CBM', default_cbm)
                
                # Ako su i u redku i u defaultu prazni, stavljamo nulu
                if pd.isna(masa): masa = default_gewicht
                if pd.isna(cbm): cbm = default_cbm
                if pd.isna(pak): pak = default_pak

                processed_rows.append({
                    "Datum izrade": datum_izrade,
                    "Datum dostave": datum_dostave,
                    "Pošiljka": zag,
                    "Referenca": referenca,
                    "Pošiljatelj": posiljatelj,
                    "Primatelj": primatelj,
                    "Grad primatelja": grad_primatelja,
                    "OTP / Broj naloga": otp_broj if pd.notna(otp_broj) else "N/A",
                    "Tip palete": pak,
                    "Količina (Kom)": anz,
                    "Masa (KG)": masa,
                    "CBM": cbm,
                    "Prijevoz (€)": fracht,
                    "Dizel dodatak (€)": diesel,
                    "Ukupno neto (€)": ukupno_eur
                })

        df_out = pd.DataFrame(processed_rows)

        if not df_out.empty:
            st.success(f"Uspješno spojeno i obrađeno stavki: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = uploaded_file.name.replace(".xlsx", "_sve_spojeno_uredno.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Spojeno_Sve')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni uređeni Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke za spajanje.")
