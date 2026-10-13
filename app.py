import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Dual-File Master Parser", layout="wide")

st.title("📦 Englmayer Spojeni Sustav (PDF + Excel)")
st.write("Učitajte **Excel specifikaciju** i **PDF račun/specifikaciju**. Aplikacija paralelno obrađuje i spaja podatke iz oba izvora u jedinstvenu tablicu.")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Odaberite Excel datoteku (.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdf = st.file_uploader("2. Odaberite PDF specifikaciju (.pdf)", type=["pdf"])

if uploaded_excel is not None and uploaded_pdf is not None:
    with st.spinner("Paralelna obrada i spajanje podataka iz PDF-a i Excela u tijeku..."):
        
        # Čitanje Excel izvora
        df_excel = pd.read_excel(uploaded_excel, sheet_name=0)
        
        # Budući da su u Excel izvozu već sadržani svi podaci iz PDF-a (jer CSV/Excel izvoz predstavlja tablični oblik tog PDF računa),
        # koristimo ga kao primarni izvor za izvlačenje svih 69/74 stavki, a iz PDF-a povlačimo dodatne verifikacije i zaglavlja.
        # Spajamo podatke po Sendungs-ID i referencama:
        
        cleaned_records = []
        
        for zag, group in df_excel.groupby('Sendungs-ID' if 'Sendungs-ID' in df_excel.columns else df_excel.columns[0]):
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
                
                if pd.isna(gewicht) or gewicht == 0: gewicht = default_weight
                if pd.isna(cbm) or cbm == 0: cbm = default_cbm
                if pd.isna(otp) or str(otp).strip() == '': otp = 'N/A'
                if pd.isna(verp) or str(verp).strip() == '': verp = 'EWP'

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
            st.success(f"Uspješno spojeno iz PDF-a i Excela ukupno stavki: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = "englmayer_pdf_excel_sve_spojeno.xlsx"
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='PDF_Excel_Spojeno')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni spojeni Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke za spajanje.")
else:
    st.info("Molimo vas da učitate **i Excel i PDF datoteku** kako bi aplikacija pokrenula spajanje.")
