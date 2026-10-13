import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Englmayer Master Parser - Pro", layout="wide")

st.title("📦 Englmayer Master Parser (Sve pošiljke i razdvojene reference)")
st.write("Učitajte službenu Excel specifikaciju (`3500131683-CSV.xlsx`). Aplikacija automatski raspisuje sve složene reference i palete.")

uploaded_file = st.file_uploader("Odaberite Excel specifikaciju (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Dubinska obrada, parsiranje i raspisivanje svih stavki u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        expanded_records = []
        
        for idx, row in df_raw.iterrows():
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

            # Ako OTP / znak sadrži kosu crtu (npr. otp-71630/629/631/632), raspisujemo ih u zasebne redke
            if '/' in zeichen and not zeichen.lower().startswith('http'):
                # Izdvajamo base prefiks (npr. 'otp-') i brojeve
                parts = zeichen.split('/')
                prefix = ''.join([c for c in parts[0] if not c.isdigit() and c != '-'])
                
                # Prvi dio nosi cijeli broj, ostali dijele teret/težinu
                for i, p in enumerate(parts):
                    sub_zeichen = p.strip()
                    if not sub_zeichen.startswith(prefix) and i > 0:
                        # Spojimo prefiks ako nedostaje
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
            st.success(f"Uspješno raspisano i obrađeno ukupno stavki i referenci: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = "englmayer_sve_pošiljke_i_reference.xlsx"
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Sve_Pošiljke')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi konačni kompletni Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke za obradu.")
else:
    st.info("Molimo vas da učitate Excel specifikaciju.")
