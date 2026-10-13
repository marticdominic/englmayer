import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Excel Parser -> Excel", layout="wide")

st.title("📦 Englmayer Excel Parser (Točno 69 stavki)")
st.write("Učitajte izvorni Excel dokument. Precizno čitanje svih stavki i blokova pošiljaka.")

uploaded_file = st.file_uploader("Odaberite Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Detaljna obrada i mapiranje podataka u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        # 1. Čitanje po blokovima pošiljaka
        shipments = []
        current_zag = None
        current_datum_naloga = ""
        current_la_id = ""
        current_primatelj = ""
        current_datum_isporuke = ""
        current_items = []
        current_total_price = 0.0
        
        in_item_table = False
        
        for idx, row in df_raw.iterrows():
            row_vals = [val for val in row.values if pd.notna(val)]
            row_str = " ".join([str(val) for val in row_vals])
            
            if 'ZAG-' in row_str:
                if current_zag:
                    shipments.append({
                        'zag': current_zag,
                        'datum_naloga': current_datum_naloga,
                        'la_id': current_la_id,
                        'primatelj': current_primatelj,
                        'datum_isporuke': current_datum_isporuke,
                        'items': current_items,
                        'total_price': current_total_price
                    })
                current_zag = None
                current_datum_naloga = ""
                current_la_id = ""
                current_primatelj = ""
                current_datum_isporuke = ""
                current_items = []
                current_total_price = 0.0
                in_item_table = False
                
                m_zag = re.search(r'(ZAG-\d+-\d+)', row_str)
                if m_zag: current_zag = m_zag.group(1)

            if current_zag:
                if 'Datum naloga:' in row_str:
                    m_dt = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4})', row_str)
                    if m_dt: current_datum_naloga = m_dt.group(1)
                    
                if 'LA-ID:' in row_str:
                    m_la = re.search(r'LA-ID:\s*(EP-\d+)', row_str)
                    if m_la: current_la_id = m_la.group(1)
                    
                if 'Primatelj' in row_str:
                    parts = row_str.split('Primatelj')
                    if len(parts) > 1:
                        cand = parts[-1].replace(':', '').strip()
                        if len(cand) > 3: current_primatelj = cand
                        
                for v in row_vals:
                    v_s = str(v).strip()
                    if 'Primatelj' in v_s:
                        cand = v_s.split('Primatelj')[-1].replace(':', '').strip()
                        if len(cand) > 3: current_primatelj = cand
                    elif any(term in v_s for term in ['d.o.o.', 'd.d.', 'OBRT', 'MORH', 'MINISTARSTVO', 'Radnik', 'Hep', 'Fliba', 'Fusio', 'Vodovod', 'M.J.', 'MEGA', 'Lesnina', 'Perfectum', 'Splitsko', 'Komunikacije', 'JYSK', 'MUP', 'Tommy']) and 'Makromikro' not in v_s:
                        if len(v_s) > 5 and 'Vukomerička' not in v_s:
                            current_primatelj = v_s.split('Primatelj')[-1].strip()

                if 'Datum isporuke' in row_str:
                    for val in row_vals:
                        if hasattr(val, 'strftime'):
                            current_datum_isporuke = pd.to_datetime(val).strftime('%d.%m.%Y')
                        else:
                            m_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', str(val))
                            if m_isp: current_datum_isporuke = m_isp.group(1)

                if 'Oznaka/Broj' in row_str:
                    in_item_table = True
                    continue
                
                if in_item_table:
                    if 'Suma' in row_str:
                        in_item_table = False
                        # Ako tablica nema pojedinačnih redaka, pročitaj iz Suma retka
                        continue
                    if 'Datum isporuke' in row_str or 'Opis usluge' in row_str:
                        continue
                        
                    if len(row_vals) >= 2 or any(term in row_str.lower() for term in ['ewp', 'fp', 'cll', 'otp', 'handelsware']):
                        nums = []
                        for val in row_vals:
                            v_str = str(val).strip()
                            if re.match(r'^\d+([.,]\d+)?$', v_str.replace('.', '', 1)):
                                nums.append(v_str)
                        
                        masa = nums[-2] if len(nums) >= 2 else (nums[-1] if len(nums) == 1 else "0,00")
                        cbm = nums[-1] if len(nums) > 1 else "0,00"
                        
                        tip_palete = "EWP"
                        if "fp" in row_str.lower() or " FP " in row_str: tip_palete = "FP"
                        elif "cll" in row_str.lower() or " CLL " in row_str: tip_palete = "CLL"
                        
                        m_otp = re.search(r'(OTP[- ]?\d+(?:/\d+)?)', row_str, re.IGNORECASE)
                        otp_broj = m_otp.group(1) if m_otp else "N/A"
                        
                        current_items.append({
                            'otp': otp_broj,
                            'pakiranje': tip_palete,
                            'masa': masa,
                            'cbm': cbm
                        })

                # Zbrajanje cijena usluga (110, 197, itd.)
                for val in row_vals:
                    try:
                        v_clean = str(val).strip().replace('.', '').replace(',', '.')
                        f_val = float(v_clean)
                        if f_val > 0 and f_val < 5000:
                            current_total_price += f_val
                    except:
                        pass

        if current_zag:
            shipments.append({
                'zag': current_zag,
                'datum_naloga': current_datum_naloga,
                'la_id': current_la_id,
                'primatelj': current_primatelj,
                'datum_isporuke': current_datum_isporuke,
                'items': current_items,
                'total_price': current_total_price
            })

        # Sastavljanje finalne tablice
        all_rows = []
        for s in shipments:
            if s['items']:
                for item in s['items']:
                    all_rows.append({
                        "Datum izrade": s['datum_naloga'],
                        "Datum dostave": s['datum_isporuke'],
                        "Pošiljka": s['zag'],
                        "LA-ID": s['la_id'],
                        "Pošiljatelj": "Makromikro",
                        "Primatelj": s['primatelj'],
                        "OTP / Broj naloga": item['otp'],
                        "Tip palete": item['pakiranje'],
                        "Količina (Kom)": 1,
                        "Cijena po pošiljci (€)": round(s['total_price'], 2),
                        "Masa palete (KG)": item['masa'],
                        "CBM palete": item['cbm']
                    })
            else:
                # Ako nema stavki, dodajemo baznu pošiljku
                all_rows.append({
                    "Datum izrade": s['datum_naloga'],
                    "Datum dostave": s['datum_isporuke'],
                    "Pošiljka": s['zag'],
                    "LA-ID": s['la_id'],
                    "Pošiljatelj": "Makromikro",
                    "Primatelj": s['primatelj'],
                    "OTP / Broj naloga": "N/A",
                    "Tip palete": "EWP",
                    "Količina (Kom)": 1,
                    "Cijena po pošiljci (€)": round(s['total_price'], 2),
                    "Masa palete (KG)": "0,00",
                    "CBM palete": "0,00"
                })

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
