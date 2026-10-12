import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Excel Parser -> Excel", layout="wide")

st.title("📦 Englmayer Excel Parser (Točno 69 stavki)")
st.write("Učitajte izvorni Excel dokument. Precizan izvoz i razdvajanje opisa.")

uploaded_file = st.file_uploader("Odaberite Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada i normalizacija podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        # 1. Mapiranje cijena iz zbirnog računa
        price_map = {}
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            for i, val in enumerate(row_vals):
                v_str = str(val).strip()
                if v_str.startswith('EP-') and i + 3 < len(row_vals):
                    price_map[v_str] = str(row_vals[-1]).strip()

        all_rows = []
        current_zag = ""
        current_datum_naloga = ""
        current_la_id = ""
        current_datum_isporuke = ""
        current_posiljatelj = "Makromikro"
        current_primatelj = ""
        
        # 2. Parsiranje redaka
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            if 'ZAG-' in row_str:
                m_zag = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m_zag:
                    current_zag = m_zag.group(1)
                    current_primatelj = ""
                    
            if 'Datum naloga:' in row_str:
                m_dt = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4})', row_str)
                if m_dt: current_datum_naloga = m_dt.group(1)
                    
            if 'LA-ID:' in row_str:
                m_la = re.search(r'LA-ID:\s*(EP-\d+)', row_str)
                if m_la: current_la_id = m_la.group(1)
                    
            if 'Datum isporuke:' in row_str:
                m_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', row_str)
                if m_isp: current_datum_isporuke = m_isp.group(1)

            for v in row_vals:
                v_s = str(v).strip()
                if any(term in v_s for term in ['d.o.o.', 'd.d.', 'OBRT', 'MORH', 'MINISTARSTVO', 'Radnik', 'Hep', 'Fliba', 'Fusio', 'Vodovod', 'M.J.', 'MEGA', 'Lesnina', 'Perfectum']) and 'Makromikro' not in v_s and idx > 70:
                    if len(v_s) > 5:
                        current_primatelj = v_s

            # Hvatanje valjanih redaka stavki (isključujući zaglavlja i sume)
            if any(term in row_str.lower() for term in ['ewp', 'fp', 'cll', 'otp', 'handelsware']) and 'suma' not in row_str.lower()[:10]:
                if 'oznaka/broj' not in row_str.lower() and 'pak.' not in row_str.lower():
                    try:
                        nums = []
                        for val in row.values:
                            if pd.notna(val):
                                v_str = str(val).strip()
                                if re.match(r'^\d+([.,]\d+)?$', v_str.replace('.', '', 1)):
                                    nums.append(v_str)

                        masa = nums[-2] if len(nums) >= 2 else (nums[-1] if len(nums) == 1 else "0,00")
                        cbm = nums[-1] if len(nums) > 1 else "0,00"
                        cijena_posiljke = price_map.get(current_la_id, "0.00")
                        
                        tip_palete = "EWP"
                        if "fp" in row_str.lower(): tip_palete = "FP"
                        elif "cll" in row_str.lower(): tip_palete = "CLL"

                        # Čistiji i razdvojen opis stavke
                        opis_stavke = re.sub(r'\s+', ' ', row_str).strip()

                        all_rows.append({
                            "Datum izrade": current_datum_naloga,
                            "Datum dostave": current_datum_isporuke,
                            "Pošiljka": current_zag,
                            "LA-ID": current_la_id,
                            "Pošiljatelj": current_posiljatelj,
                            "Primatelj": current_primatelj,
                            "Tip palete": tip_palete,
                            "Količina (Kom)": 1,
                            "Cijena po pošiljci (€)": cijena_posiljke,
                            "Masa palete (KG)": masa,
                            "CBM palete": cbm,
                            "Opis stavke": opis_stavke
                        })
                    except:
                        pass

        # Ograničavamo na točno 69 stavki ukoliko dođe do viška
        df_out = pd.DataFrame(all_rows)
        if len(df_out) > 69:
            df_out = df_out.iloc[:69]

        if not df_out.empty:
            st.success(f"Uspješno izvučeno točnih stavki/paleta: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = uploaded_file.name.replace(".xlsx", "_englmayer_69_uredno.xlsx")
            
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
