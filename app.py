import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Excel Parser -> Excel", layout="wide")

st.title("📦 Englmayer Excel Parser")
st.write("Učitajte izvorni Excel dokument. Aplikacija precizno izvlači sve stavke i generira čist izlazni Excel.")

uploaded_file = st.file_uploader("Odaberite Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Obrada Excel datoteke u tijeku..."):
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=0)
        
        all_rows = []
        current_zag = ""
        current_datum_naloga = ""
        current_la_id = ""
        current_datum_isporuke = ""
        current_suma_eur = ""
        current_primatelj = ""
        current_grad = ""
        
        # Prolazimo red po red kroz Excel
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

            # 2. Provjera je li ovo redak s paletom/stavkom (sadrži EWP, FP, CLL, OTP ili Handelsware, a nije Suma)
            if any(term in row_str.lower() for term in ['ewp', 'fp', 'cll', 'otp', 'handelsware']) and 'suma' not in row_str.lower()[:10]:
                if 'oznaka/broj' not in row_str.lower() and 'pak.' not in row_str.lower():
                    try:
                        # Izvlačenje količine, mase i CBM-a iz ćelija
                        kol_val = "1"
                        masa_val = "0,00"
                        cbm_val = "0,00"
                        
                        # Tražimo brojčane vrijednosti u retku
                        nums = []
                        for val in row.values:
                            if pd.notna(val):
                                v_str = str(val).strip()
                                if v_str.replace('.', '', 1).isdigit() or ',' in v_str:
                                    nums.append(v_str)

                        all_rows.append({
                            "Datum naloga": current_datum_naloga,
                            "Pošiljka": current_zag,
                            "LA-ID": current_la_id,
                            "Pošiljatelj": "Makromikro",
                            "Primatelj": current_primatelj,
                            "Datum isporuke": current_datum_isporuke,
                            "Oznaka/Broj stavke": row_str,
                            "Masa palete (KG)": nums[-2] if len(nums) >= 2 else "0,00",
                            "CBM palete": nums[-1] if len(nums) >= 1 else "0,00"
                        })
                    except:
                        pass

        df_out = pd.DataFrame(all_rows)

        if not df_out.empty:
            st.success(f"Uspješno izvučeno stavki/paleta: {len(df_out)}")
            st.dataframe(df_out)

            file_name_xlsx = uploaded_file.name.replace(".xlsx", "_parsirano.xlsx")
            
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
            st.warning("Nije moguće pronaći stavke u Excel datoteci.")
