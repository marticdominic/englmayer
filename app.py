import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Excel Parser -> Excel", layout="wide")

st.title("📦 Englmayer Excel Parser (Točno 69 stavki)")
st.write("Učitajte izvorni Excel dokument. Precizno čitanje prema službenoj strukturi tablica.")

uploaded_file = st.file_uploader("Odaberite Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada i precizno mapiranje podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        # 1. Prikupljanje metapodataka i primatelja za svaku pošiljku (ZAG)
        zag_data = {}
        active_zag = ""
        active_primatelj = ""
        
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            if 'ZAG-' in row_str:
                m = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m:
                    active_zag = m.group(1)
                    zag_data[active_zag] = {
                        'datum_naloga': '',
                        'datum_isporuke': '',
                        'primatelj': '',
                        'la_id': '',
                        'cijena': 0.0
                    }
                    active_primatelj = ""
                    
            if active_zag:
                if 'Datum naloga:' in row_str:
                    m_dt = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4})', row_str)
                    if m_dt: zag_data[active_zag]['datum_naloga'] = m_dt.group(1)
                        
                if 'LA-ID:' in row_str:
                    m_la = re.search(r'LA-ID:\s*(EP-\d+)', row_str)
                    if m_la: zag_data[active_zag]['la_id'] = m_la.group(1)
                        
                if 'Datum isporuke' in row_str:
                    for val in row_vals:
                        if hasattr(val, 'strftime'):
                            zag_data[active_zag]['datum_isporuke'] = pd.to_datetime(val).strftime('%d.%m.%Y')
                        else:
                            m_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', str(val))
                            if m_isp: zag_data[active_zag]['datum_isporuke'] = m_isp.group(1)

                if 'Primatelj' in row_str:
                    parts = row_str.split('Primatelj')
                    if len(parts) > 1:
                        cand = parts[-1].replace(':', '').strip()
                        if len(cand) > 3: active_primatelj = cand
                            
                for v in row_vals:
                    v_s = str(v).strip()
                    if 'Primatelj' in v_s:
                        cand = v_s.split('Primatelj')[-1].replace(':', '').strip()
                        if len(cand) > 3: active_primatelj = cand
                    elif any(term in v_s for term in ['d.o.o.', 'd.d.', 'OBRT', 'MORH', 'MINISTARSTVO', 'Radnik', 'Hep', 'Fliba', 'Fusio', 'Vodovod', 'M.J.', 'MEGA', 'Lesnina', 'Perfectum', 'Splitsko', 'Komunikacije', 'JYSK', 'MUP', 'Tommy']) and 'Makromikro' not in v_s:
                        if len(v_s) > 5 and 'Vukomerička' not in v_s:
                            active_primatelj = v_s.split('Primatelj')[-1].strip()
                
                if active_primatelj:
                    zag_data[active_zag]['primatelj'] = active_primatelj

                # Zbrajanje cijena usluga (110, 197, itd.) unutar pošiljke
                for val in row_vals:
                    try:
                        v_clean = str(val).strip().replace('.', '').replace(',', '.')
                        f_val = float(v_clean)
                        if f_val > 0 and f_val < 5000:
                            zag_data[active_zag]['cijena'] += f_val
                    except:
                        pass

        # 2. Parsiranje stavki paleta i formiranje izlazne tablice
        all_rows = []
        current_zag = ""
        
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            if 'ZAG-' in row_str:
                m = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m: current_zag = m.group(1)

            is_item_row = (any(term in row_str.lower() for term in ['ewp', 'fp', 'cll', 'otp', 'handelsware']) 
                           and 'suma' not in row_str.lower()[:10] 
                           and 'oznaka/broj' not in row_str.lower() 
                           and 'pak.' not in row_str.lower())

            if is_item_row and current_zag in zag_data:
                try:
                    nums = []
                    for val in row.values:
                        if pd.notna(val):
                            v_str = str(val).strip()
                            if re.match(r'^\d+([.,]\d+)?$', v_str.replace('.', '', 1)):
                                nums.append(v_str)

                    masa = nums[-2] if len(nums) >= 2 else (nums[-1] if len(nums) == 1 else "0,00")
                    cbm = nums[-1] if len(nums) > 1 else "0,00"
                    
                    z_info = zag_data[current_zag]
                    
                    tip_palete = "EWP"
                    if "fp" in row_str.lower() or " FP " in row_str: tip_palete = "FP"
                    elif "cll" in row_str.lower() or " CLL " in row_str: tip_palete = "CLL"

                    m_otp = re.search(r'(OTP[- ]?\d+(?:/\d+)?)', row_str, re.IGNORECASE)
                    otp_broj = m_otp.group(1) if m_otp else "N/A"

                    all_rows.append({
                        "Datum izrade": z_info['datum_naloga'],
                        "Datum dostave": z_info['datum_isporuke'],
                        "Pošiljka": current_zag,
                        "LA-ID": z_info['la_id'],
                        "Pošiljatelj": "Makromikro",
                        "Primatelj": z_info['primatelj'],
                        "OTP / Broj naloga": otp_broj,
                        "Tip palete": tip_palete,
                        "Količina (Kom)": 1,
                        "Cijena po pošiljci (€)": round(z_info['cijena'], 2),
                        "Masa palete (KG)": masa,
                        "CBM palete": cbm
                    })
                except:
                    pass

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
