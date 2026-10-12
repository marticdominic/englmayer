import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Excel Parser -> Excel", layout="wide")

st.title("📦 Englmayer Excel Parser (Točno 69 stavki)")
st.write("Učitajte izvorni Excel dokument. Precizan izvoz pošiljaka, masa, primatelja i cijena.")

uploaded_file = st.file_uploader("Odaberite Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada i mapiranje podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        # 1. Mapiranje cijena (LA-ID -> Iznos in EUR) iz tablice zbirnog računa
        price_map = {}
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            for i, val in enumerate(row_vals):
                v_str = str(val).strip()
                if v_str.startswith('EP-') and i + 3 < len(row_vals):
                    possible_price = row_vals[-1]
                    price_map[v_str] = str(possible_price).strip()

        all_rows = []
        
        # Varijable za praćenje tekuće pošiljke
        current_zag = ""
        current_datum_naloga = ""
        current_la_id = ""
        current_datum_isporuke = ""
        current_posiljatelj = "Makromikro"
        current_primatelj = ""
        
        # 2. Prolazimo red po red i parsiramo blokove pošiljaka
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            # Detekcija početka nove pošiljke ili zaglavlja
            if 'ZAG-' in row_str:
                m_zag = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m_zag:
                    current_zag = m_zag.group(1)
                    # Resetiramo primatelja za novu pošiljku
                    current_primatelj = ""
                    current_posiljatelj = "Makromikro"
                    
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

            # Precizno izvlačenje pošiljatelja i primatelja unutar bloka pošiljke (ignorirajući gornje zaglavlje računa)
            if 'Primatelj' in row_str and idx > 70:
                # Primatelj se obično nalazi u istoj ćeliji nakon naziva
                parts = row_str.split('Primatelj')
                if len(parts) > 1 and len(parts[-1].strip()) > 3:
                    current_primatelj = parts[-1].strip().replace(';', '').strip()
            
            for v in row_vals:
                v_s = str(v).strip()
                if ('d.o.o.' in v_s or 'd.d.' in v_s or 'OBRT' in v_s or 'MORH' in v_s or 'HEEP' in v_s or 'MINISTARSTVO' in v_s) and 'Makromikro Grupa d.o.o., Vukomerička' not in v_s:
                    if idx > 70: # Ignoriramo gornje zaglavlje računa
                        current_primatelj = v_s

            # Detekcija redaka paleta (EWP, FP, CLL, OTP, Handelsware, a nije Suma)
            if any(term in row_str.lower() for term in ['ewp', 'fp', 'cll', 'otp', 'handelsware']) and 'suma' not in row_str.lower()[:10]:
                if 'oznaka/broj' not in row_str.lower() and 'pak.' not in row_str.lower():
                    try:
                        # Pronalaženje mase palete
                        nums = []
                        for val in row.values:
                            if pd.notna(val):
                                v_str = str(val).strip()
                                if re.match(r'^\d+([.,]\d+)?$', v_str.replace('.', '', 1)):
                                    nums.append(v_str)

                        masa = nums[-2] if len(nums) >= 2 else (nums[-1] if len(nums) == 1 else "0,00")
                        cijena_posiljke = price_map.get(current_la_id, "0.00")
                        
                        # Određivanje tipa palete
                        tip_palete = "EWP"
                        if "fp" in row_str.lower():
                            tip_palete = "FP"
                        elif "cll" in row_str.lower():
                            tip_palete = "CLL"

                        all_rows.append({
                            "Datum izrade": current_datum_naloga,
                            "Datum dostave": current_datum_isporuke,
                            "Pošiljka": current_zag,
                            "LA-ID": current_la_id,
                            "Pošiljatelj": current_posiljatelj,
                            "Primatelj": current_primatelj,
                            "Tip palete": tip_palete,
                            "Cijena po pošiljci (€)": cijena_posiljke,
                            "Masa palete (KG)": masa,
                            "Oznaka/Opis stavke": row_str
                        })
                    except:
                        pass

        df_out = pd.DataFrame(all_rows)

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
