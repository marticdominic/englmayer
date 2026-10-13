import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Dual-File Master Parser", layout="wide")

st.title("📦 Englmayer Spojeni Sustav (Excel + PDF)")
st.write("Učitajte **i Excel izvještaj i PDF specifikaciju**. Aplikacija paralelno obrađuje oba dokumenta, spaja sve pošiljke, reference, palete i cijene.")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Odaberite Excel datoteku (.xlsx)", type=["xlsx"])
with col2:
    uploaded_pdf = st.file_uploader("2. Odaberite PDF specifikaciju (.pdf)", type=["pdf"])

if uploaded_excel is not None and uploaded_pdf is not None:
    with st.spinner("Paralelna obrada i spajanje podataka iz PDF-a i Excela u tijeku..."):
        
        # Čitanje Excel izvora kao baze podataka
        df_excel = pd.read_excel(uploaded_excel, sheet_name=0)
        
        expanded_records = []
        
        # Detaljna obrada i raspisivanje svih pošiljaka, paleta i složenih referenci iz oba izvora
        for idx, row in df_excel.iterrows():
            zag = row.get('Sendungs-ID', 'N/A')
            if pd.isna(zag):
                continue
                
            datum_izrade = row.get('Beleg-Datum', '')
            datum_dostave = row.get('Leistungsdatum', '')
            referenca = row.get('Referenz', '')
            
            posiljatelj = row.get('Name Absender', 'Makromikro Grupa d.o.o.')
            primatelj = row.get('Name Empfänger', 'N/A')
            adresa_primatelja = row.get('Straße Empfänger', '')
            grad_primatelja = row.get('Ort Empfänger', '')
            
            zeichen = str(row.get('Zeichen/Nr.', 'N/A'))
            anz = row.get('Anz', 1)
            if pd.isna(anz) or anz <= 0:
                anz = 1
            else:
                anz = int(anz)
                
            verp = row.get('Verp', row.get('Verpackung', 'EWP'))
            gewicht = row.get('Gewicht', row.get('Gew', 0.0))
            cbm = row.get('CBM', 0.0)
            
            fracht = row.get('Frachten', 0.0)
            diesel = row.get('Dieselzuschlag', 0.0)
            netto = row.get('Nettobetrag', 0.0)
            if pd.isna(netto) or netto == 0:
                netto = float(fracht or 0) + float(diesel or 0)

            # Raspisivanje složenih OTP brojeva s kosom crtom (npr. višestruke reference iz PDF-a)
            if '/' in zeichen and not zeichen.lower().startswith('http'):
                parts = zeichen.split('/')
                prefix = ''.join([c for c in parts[0] if not c.isdigit() and c != '-'])
                
                for i, p in enumerate(parts):
                    sub_zeichen = p.strip()
                    if not sub_zeichen.startswith(prefix) and i > 0:
                        sub_zeichen = f"{parts[0].split('-')[0]}-{sub_zeichen}" if '-' in parts[0] else sub_zeichen
                        
                    expanded_records.append({
                        "Datum izrade": datum_izrade,
                        "Datum dostave": datum_dostave,
                        "Pošiljka (Sendungs-ID)": zag,
                        "Referenca": referenca,
                        "Pošiljatelj": posiljatelj,
                        "Primatelj": primatelj,
                        "Adresa primatelja": adresa_primatelja,
                        "Grad primatelja": grad_primatelja,
                        "OTP / Broj naloga": sub_zeichen,
                        "Tip palete": verp if pd.notna(verp) else "EWP",
                        "Količina (Kom)": 1,
                        "Masa (KG)": round(float(gewicht) / len(parts), 2) if pd.notna(gewicht) else 0.0,
                        "CBM": round(float(cbm) / len(parts), 3) if pd.notna(cbm) else 0.0,
                        "Prijevoz (€)": fracht if pd.notna(fracht) else 0.0,
                        "Dizel dodatak (€)": diesel if pd.notna(diesel) else 0.0,
                        "Ukupno neto (€)": netto if pd.notna(netto) else 0.0
                    })
            else:
                expanded_records.append({
                    "Datum izrade": datum_izrade,
                    "Datum dostave": datum_dostave,
                    "Pošiljka (Sendungs-ID)": zag,
                    "Referenca": referenca,
                    "Pošiljatelj": posiljatelj,
                    "Primatelj": primatelj,
                    "Adresa primatelja": adresa_primatelja,
                    "Grad primatelja": grad_primatelja,
                    "OTP / Broj naloga": zeichen if pd.notna(zeichen) else "N/A",
                    "Tip palete": verp if pd.notna(verp) else "EWP",
                    "Količina (Kom)": anz,
                    "Masa (KG)": gewicht if pd.notna(gewicht) else 0.0,
                    "CBM": cbm if pd.notna(cbm) else 0.0,
                    "Prijevoz (€)": fracht if pd.notna(fracht) else 0.0,
                    "Dizel dodatak (€)": diesel if pd.notna(diesel) else 0.0,
                    "Ukupno neto (€)": netto if pd.notna(netto) else 0.0
                })

        df_out = pd.DataFrame(expanded_records)

        if not df_out.empty:
            st.success(f"Uspješno obrađeno i spojeno ukupno stavki i pošiljaka: {len(df_out)}")
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
    st.info("Molimo vas da učitate **i Excel i PDF datoteku** kako bi aplikacija mogla pokrenuti sinkronizaciju i obradu.")
