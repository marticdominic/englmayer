import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Master Parser (PDF + Excel)", layout="wide")

st.title("📦 Englmayer Sustav za obradu (PDF + Excel)")
st.write("Učitajte **i službenu PDF specifikaciju i Excel datoteku**. Aplikacija paralelno obrađuje oba izvora, osigurava točne kilaže za svaku paletu i uredno raspoređuje cijene.")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Odaberite Excel datoteku (.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdf = st.file_uploader("2. Odaberite PDF specifikaciju (.pdf)", type=["pdf"])

if uploaded_excel is not None and uploaded_pdf is not None:
    with st.spinner("Paralelna obrada PDF-a i Excela, izračun masa i formatiranje u tijeku..."):
        
        # Čitanje Excel izvora kao baze podataka
        df_raw = pd.read_excel(uploaded_excel, sheet_name=0)
        
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
                
                # Osiguravamo da svaka paleta/stavka ima svoju kilažu i CBM
                if pd.isna(gewicht) or gewicht == 0:
                    gewicht = total_w / num_rows if num_rows > 0 else 0.0

                if pd.isna(cbm) or cbm == 0:
                    cbm = total_cbm / num_rows if num_rows > 0 else 0.0

                if pd.isna(verp) or str(verp).strip() == '' or str(verp).lower() == 'nan':
                    verp = 'EWP'

                # Cijena se ispisuje samo na prvom retku pošiljke, dok svaka paleta ima svoju kilažu
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
            st.success(f"Uspješno spojeno iz PDF-a i Excela – ukupno obrađeno stavki: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = "englmayer_pdf_excel_master.xlsx"
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Master_PDF_Excel')
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
    st.info("Molimo vas da učitate **i Excel i PDF datoteku** kako bi aplikacija mogla pokrenuti obradu.")
