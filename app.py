import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Spojeni Sustav (PDF + Excel)", layout="wide")

st.title("📦 Englmayer Spojeni Sustav (PDF Specifikacija + Excel)")
st.write("Učitajte službenu Excel specifikaciju (npr. `3500131683-CSV.xlsx`) koja automatski objedinjuje podatke iz PDF-a (palete, težine, CBM, cijene i primatelje).")

uploaded_file = st.file_uploader("Odaberite Excel specifikaciju (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Spajanje i obrada podataka iz PDF specifikacije i Excela u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        processed_rows = []
        
        # Spajanje podataka na razini svake pošiljke (Sendungs-ID)
        for zag, group in df_raw.groupby('Sendungs-ID'):
            first_row = group.iloc[0]
            
            datum_izrade = first_row.get('Beleg-Datum', '')
            datum_dostave = first_row.get('Leistungsdatum', '')
            referenca = first_row.get('Referenz', '')
            
            posiljatelj = first_row.get('Name Absender', 'Makromikro Grupa d.o.o.')
            primatelj = first_row.get('Name Empfänger', 'N/A')
            grad_primatelja = first_row.get('Ort Empfänger', '')
            
            # Financijski dio iz PDF/Excel specifikacije (Frachten + Dieselzuschlag)
            fracht = first_row.get('Frachten', 0.0)
            diesel = first_row.get('Dieselzuschlag', 0.0)
            ukupno_eur = first_row.get('Nettobetrag', 0.0)
            
            if pd.isna(ukupno_eur) or ukupno_eur == 0:
                ukupno_eur = float(fracht or 0) + float(diesel or 0)

            # Filtriranje i mapiranje svih stavki paleta (težine, CBM, OTP brojevi)
            valid_items = group[group['Gewicht'].notna() | group['Verp'].notna() | group['Verpackung'].notna()]
            if len(valid_items) == 0:
                valid_items = group

            for _, r in valid_items.iterrows():
                otp_broj = r.get('Zeichen/Nr.', 'N/A')
                anz = r.get('Anz', 1)
                if pd.isna(anz) or anz <= 0:
                    anz = 1
                else:
                    anz = int(anz)
                    
                pak = r.get('Verp', r.get('Verpackung', 'EWP'))
                masa = r.get('Gewicht', r.get('Gew', 0.0))
                cbm = r.get('CBM', 0.0)
                
                processed_rows.append({
                    "Datum izrade": datum_izrade,
                    "Datum dostave": datum_dostave,
                    "Pošiljka": zag,
                    "Referenca": referenca,
                    "Pošiljatelj": posiljatelj,
                    "Primatelj": primatelj,
                    "Grad primatelja": grad_primatelja,
                    "OTP / Broj naloga": otp_broj if pd.notna(otp_broj) else "N/A",
                    "Tip palete": pak if pd.notna(pak) else "EWP",
                    "Količina (Kom)": anz,
                    "Masa (KG)": masa if pd.notna(masa) else 0.0,
                    "CBM": cbm if pd.notna(cbm) else 0.0,
                    "Prijevoz (€)": fracht,
                    "Dizel dodatak (€)": diesel,
                    "Ukupno neto (€)": ukupno_eur
                })

        df_out = pd.DataFrame(processed_rows)

        if not df_out.empty:
            st.success(f"Uspješno spojeno stavki: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = uploaded_file.name.replace(".xlsx", "_spojeno_pdf_excel.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Spojeno_Sve')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni spojeni Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke za spajanje.")
