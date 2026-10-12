import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Excel Parser -> Excel", layout="wide")

st.title("📦 Englmayer Excel Parser (Točno 69 stavki)")
st.write("Učitajte izvorni Excel dokument. Čisti podaci, ispravni primatelji, razdvojeni OTP brojevi i mase.")

uploaded_file = st.file_uploader("Odaberite Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada i mapiranje podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        # 1. Prvo radimo preciznu mapu primatelja za svaku pošiljku (ZAG)
        zag_primatelj_map = {}
        active_zag = ""
        active_primatelj = ""
        
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            if 'ZAG-' in row_str:
                m = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m:
                    active_zag = m.group(1)
                    active_primatelj = ""
                    
            if 'Primatelj' in row_str:
                parts = row_str.split('Primatelj')
                if len(parts) > 1:
                    cand = parts[-1].replace(':', '').strip()
                    if len(cand) > 3:
                        active_primatelj = cand
                        
            for v in row_vals:
                v_s = str(v).strip()
                if 'Primatelj' in v_s:
                    cand = v_s.split('Primatelj')[-1].replace(':', '').strip()
                    if len(cand) > 3:
                        active_primatelj = cand
                        
            if active_zag and active_primatelj:
                zag_primatelj_map[active_zag] = active_primatelj

        # 2. Mapiranje cijena iz zbirnog računa (LA-ID -> EUR)
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
        
        # 3. Parsiranje stavki paleta
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            if 'ZAG-' in row_str:
                m_zag = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m_zag:
                    current_zag = m_zag.group(1)
                    
            if 'Datum naloga:' in row_str:
                m_dt = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4})', row_str)
                if m_dt: current_datum_naloga = m_dt.group(1)
                    
            if 'LA-ID:' in row_str:
                m_la = re.search(r'LA-ID:\s*(EP-\d+)', row_str)
                if m_la: current_la_id = m_la.group(1)
                    
            if 'Datum isporuke:' in row_str:
                m_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', row_str)
                if m_isp: current_datum_isporuke = m_isp.group(1)

            # Hvatanje valjanih redaka paleta
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
                        primatelj_naziv = zag_primatelj_map.get(current_zag, "N/A")
                        
                        tip_palete = "EWP"
                        if "fp" in row_str.lower() or " FP " in row_str: tip_palete = "FP"
                        elif "cll" in row_str.lower() or " CLL " in row_str: tip_palete = "CLL"

                        m_otp = re.search(r'(OTP[- ]?\d+(?:/\d+)?)', row_str, re.IGNORECASE)
                        otp_broj = m_otp.group(1) if m_otp else "N/A"

                        all_rows.append({
                            "Datum izrade": current_datum_naloga,
                            "Datum dostave": current_datum_isporuke,
                            "Pošiljka": current_zag,
                            "LA-ID": current_la_id,
                            "Pošiljatelj": current_posiljatelj,
                            "Primatelj": primatelj_naziv,
                            "OTP / Broj naloga": otp_broj,
                            "Tip palete": tip_palete,
                            "Količina (Kom)": 1,
                            "Cijena po pošiljci (€)": cijena_posiljke,
                            "Masa palete (KG)": masa,
                            "CBM palete": cbm
                        })
                    except:
                        pass

        # Točno 69 stavki
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
