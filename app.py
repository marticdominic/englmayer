import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Master Parser", layout="wide")

st.title("📦 Englmayer Master Sustav (Sve pošiljke, palete i cijene)")
st.write("Učitajte službenu Excel specifikaciju (`3500131683-CSV.xlsx`). Aplikacija automatski čisti, spaja i strukturira sve podatke o pošiljkama.")

uploaded_file = st.file_uploader("Odaberite Excel specifikaciju (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada, čišćenje i spajanje svih podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        cleaned_records = []
        
        # Grupiramo po pošiljkama kako bismo osigurali da svaka stavka povuče točne financije i atribute
        for zag, group in df_raw.groupby('Sendungs-ID'):
            first_row = group.iloc[0]
            
            datum_izrade = first_row.get('Beleg-Datum', '')
            datum_dostave = first_row.get('Leistungsdatum', '')
            referenca = first_row.get('Referenz', '')
            
            posiljatelj = first_row.get('Name Absender', 'Makromikro Grupa d.o.o.')
            primatelj = first_row.get('Name Empfänger', 'N/A')
            adresa_primatelja = first_row.get('Straße Empfänger', '')
            grad_primatelja = first_row.get('Ort Empfänger', '')
            
            # Financije po pošiljci
            fracht = first_row.get('Frachten', 0.0)
            diesel = first_row.get('Dieselzuschlag', 0.0)
            netto = first_row.get('Nettobetrag', 0.0)
            if pd.isna(netto) or netto == 0:
                netto = float(fracht or 0) + float(diesel or 0)

            # Pronalazimo valjane vrijednosti težina i CBM-a unutar grupe ako su u nekom retku prazne
            valid_weights = group[group['Gewicht'].notna() & (group['Gewicht'] > 0)]['Gewicht']
            default_weight = valid_weights.iloc[0] if len(valid_weights) > 0 else 0.0

            valid_cbm = group[group['CBM'].notna() & (group['CBM'] > 0)]['CBM']
            default_cbm = valid_cbm.iloc[0] if len(valid_cbm) > 0 else 0.0

            for idx, r in group.iterrows():
                otp = r.get('Zeichen/Nr.', 'N/A')
                anz = r.get('Anz', 1)
                if pd.isna(anz) or anz <= 0:
                    anz = 1
                else:
                    anz = int(anz)
                    
                verp = r.get('Verp', r.get('Verpackung', 'EWP'))
                gewicht = r.get('Gewicht', r.get('Gew', default_weight))
                cbm = r.get('CBM', default_cbm)
                
                if pd.isna(gewicht) or gewicht == 0:
                    gewicht = default_weight
                if pd.isna(cbm) or cbm == 0:
                    cbm = default_cbm
                if pd.isna(otp) or str(otp).strip() == '':
                    otp = 'N/A'
                if pd.isna(verp) or str(verp).strip() == '':
                    verp = 'EWP'

                cleaned_records.append({
                    "Datum izrade": datum_izrade,
                    "Datum dostave": datum_dostave,
                    "Pošiljka (Sendungs-ID)": zag,
                    "Referenca": referenca,
                    "Pošiljatelj": posiljatelj,
                    "Primatelj": primatelj,
                    "Adresa primatelja": adresa_primatelja,
                    "Grad primatelja": grad_primatelja,
                    "OTP / Broj naloga": otp,
                    "Tip palete": verp,
                    "Količina (Kom)": anz,
                    "Masa (KG)": gewicht,
                    "CBM": cbm,
                    "Prijevoz (€)": fracht if pd.notna(fracht) else 0.0,
                    "Dizel dodatak (€)": diesel if pd.notna(diesel) else 0.0,
                    "Ukupno neto (€)": netto if pd.notna(netto) else 0.0
                })

        df_out = pd.DataFrame(cleaned_records)

        if not df_out.empty:
            st.success(f"Uspješno obrađeno i očišćeno ukupno stavki: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = uploaded_file.name.replace(".xlsx", "_MASTER_uredjeno.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Cisti_Podaci')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni uređeni Excel sa svim podacima (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke za obradu.")
else:
    st.info("Molimo vas da učitate Excel specifikaciju.")
