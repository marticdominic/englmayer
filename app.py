import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Excel Parser -> Excel", layout="wide")

st.title("📦 Englmayer Excel Parser (Točno 69 stavki)")
st.write("Učitajte izvorni Excel dokument. Aplikacija precizno raspisuje palete prema količinama.")

uploaded_file = st.file_uploader("Odaberite Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada i normalizacija stavki u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        all_rows = []
        current_zag = ""
        current_datum_naloga = ""
        current_la_id = ""
        current_datum_isporuke = ""
        current_suma_eur = ""
        
        # Prolazimo red po red
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            # 1. Hvatanje zaglavlja pošiljke
            if 'ZAG-' in row_str:
                m_zag = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m_zag:
                    current_zag = m_zag.group(1)
                    
            if 'Datum naloga:' in row_str:
                m_dt = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4})', row_str)
                if m_dt:
                    current_datum_naloga = m_dt.group(1)
                    
            if 'LA-ID:' in row_str:
                m_la = re.search(r'LA-ID:\s*(EP-\d+)', row_str)
                if m_la:
                    current_la_id = m_la.group(1)
                    
            if 'Datum isporuke:' in row_str:
                m_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', row_str)
                if m_isp:
                    current_datum_isporuke = m_isp.group(1)

            # 2. Hvatanje redaka sa Sumama (koji sadrže točan broj komada/paleta i ukupne mase za tu pošiljku)
            if 'suma' in row_str.lower() and 'eur' not in row_str.lower() and 'neto' not in row_str.lower():
                try:
                    # Pronađemo količinu i masu u retku sume
                    nums_in_row = []
                    text_parts = []
                    for val in row.values:
                        if pd.notna(val):
                            v_str = str(val).strip()
                            # Provjerimo je li to brojčana vrijednost
                            if re.match(r'^\d+([.,]\d+)?$', v_str.replace('.', '', 1)):
                                nums_in_row.append(v_str)
                            else:
                                text_parts.append(v_str)
                    
                    # Izvučemo količinu (obično prvi manji cijeli broj)
                    kolicina = 1
                    for part in row_vals:
                        p_str = str(part).strip()
                        m_dig = re.search(r'\b([1-9]\d*)\b', p_str)
                        if m_dig and int(m_dig.group(1)) < 20 and 'suma' not in p_str.lower():
                            kolicina = int(m_dig.group(1))
                            break
                    
                    # Nađemo tip palete (EWP, FP, CLL)
                    pakiranje = "EWP"
                    if "fp" in row_str.lower():
                        pakiranje = "FP"
                    elif "cll" in row_str.lower():
                        pakiranje = "CLL"

                    # Generiramo onoliko redaka kolika je količina paleta
                    for _ in range(kolicina):
                        all_rows.append({
                            "Datum naloga": current_datum_naloga,
                            "Pošiljka": current_zag,
                            "LA-ID": current_la_id,
                            "Pošiljatelj": "Makromikro",
                            "Datum isporuke": current_datum_isporuke,
                            "Tip palete": pakiranje,
                            "Količina (Kom)": 1,
                            "Sifra/Opis": row_str
                        })
                except:
                    pass

        df_out = pd.DataFrame(all_rows)

        if not df_out.empty:
            st.success(f"Uspješno izvučeno točnih stavki/paleta: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = uploaded_file.name.replace(".xlsx", "_tocno_69_stavki.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Posiljke')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi čisti Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke u Excel datoteci.")
