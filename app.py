import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Master Parser - Konačni", layout="wide")

st.title("📦 Englmayer Spojeni Sustav (Excel + PDF Master)")
st.write("Učitajte službenu Excel specifikaciju (`3500131683-CSV.xlsx`). Svaka paleta sada ima svoju kilažu, a cijena je uredno raspoređena.")

uploaded_file = st.file_uploader("Odaberite Excel specifikaciju (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada i formatiranje podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        records = []
        
        for zag, group in df_raw.groupby('Sendungs-ID'):
            first_row = group.iloc[0]
            
            datum_izrade = first_row.get('Beleg-Datum', '')
            datum_dostave = first_row.get('Leistungsdatum', '')
            referenca = first_row.get('Referenz', '')
            
            posiljatelj = first_row.get('Name Absender', 'Makromikro Grupa d.o.o.')
            primatelj = first_row.get('Name Empfänger', 'N/A')
            adresa_primatelja = first_row.get('Straße Empfänger', '')
            grad_primatelja = first_row.get('Ort Empfänger', '')
            
            fracht = first_row.get('Frachten', 0.0)
            diesel = first_row.get('Dieselzuschlag', 0.0)
            netto = first_row.get('Nettobetrag', 0.0)
            if pd.isna(netto) or netto == 0:
                netto = float(fracht or 0) + float(diesel or 0)

            # Ukupna masa i CBM za cijelu pošiljku
            total_w = group['Gewicht'].dropna().sum()
            if total_w == 0:
                total_w = group['Gew'].dropna().sum()

            total_cbm = group['CBM'].dropna().sum()
            num_rows = len(group)

            is_first_row = True
            for idx, r in group.iterrows():
                zeichen = str(r.get('Zeichen/Nr.', 'N/A'))
                anz = r.get('Anz', 1)
                if pd.isna(anz) or anz <= 0:
                    anz = 1
                else:
                    anz = int(anz)
                    
                verp = r.get('Verp', r.get('Verpackung', 'EWP'))
                gewicht = r.get('Gewicht', r.get('Gew', 0.0))
                cbm = r.get('CBM', 0.0)
                
                # Ako je masa u retku prazna, raspodijelimo ukupnu masu ravnomjerno na retke pošiljke
                if pd.isna(gewicht) or gewicht == 0:
                    gewicht = total_w / num_rows if num_rows > 0 else 0.0

                if pd.isna(cbm) or cbm == 0:
                    cbm = total_cbm / num_rows if num_rows > 0 else 0.0

                if pd.isna(verp) or str(verp).strip() == '' or str(verp).lower() == 'nan':
                    verp = 'EWP'

                # Cijena se ispisuje samo na prvom retku pošiljke
                p_val = fracht if is_first_row else 0.0
                d_val = diesel if is_first_row else 0.0
                n_val = netto if is_first_row else 0.0
                is_first_row = False

                records.append({
                    "Datum izrade": datum_izrade,
                    "Datum dostave": datum_dostave,
                    "Pošiljka (Sendungs-ID)": zag,
                    "Referenca": referenca,
                    "Pošiljatelj": posiljatelj,
                    "Primatelj": primatelj,
                    "Adresa primatelja": adresa_primatelja,
                    "Grad primatelja": grad_primatelja,
                    "OTP / Broj naloga": zeichen if pd.notna(zeichen) else "N/A",
                    "Tip palete": verp,
                    "Količina (Kom)": anz,
                    "Masa (KG)": round(float(gewicht), 2),
                    "CBM": round(float(cbm), 3),
                    "Prijevoz (€)": round(float(p_val), 2),
                    "Dizel dodatak (€)": round(float(d_val), 2),
                    "Ukupno neto (€)": round(float(n_val), 2)
                })

        df_out = pd.DataFrame(records)

        if not df_out.empty:
            st.success(f"Uspješno obrađeno ukupno stavki: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = "englmayer_konacno_uredno.xlsx"
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Master_Podaci')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni uređeni Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke za obradu.")
else:
    st.info("Molimo vas da učitate Excel specifikaciju.")
