import streamlit as st
from pypdf import PdfReader
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Text PDF Parser -> Excel", layout="wide")

st.title("📦 Englmayer Parser (Direktno skeniranje svake palete)")
st.write("Učitajte tekstualni PDF. Skripta direktno pronalazi svaku paletu u dokumentu.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Čitanje i parsiranje dokumenta u tijeku..."):
        
        reader = PdfReader(uploaded_file)
        full_text = ""
        
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text += text + f"\n--STRANICA {i+1}--\n"

        # Dijelimo dokument po stranicama ili manjim cjelinama da ne gubimo kontekst
        pages_text = full_text.split("--STRANICA")
        all_rows = []
        
        for page_chunk in pages_text:
            if not page_chunk.strip():
                continue
                
            lines = [line.strip() for line in page_chunk.split("\n") if line.strip()]
            
            # Tražimo sve pošiljke i pripadajuće podatke na ovoj stranici
            # Budući da na jednoj stranici može biti više naloga, pratimo trenutni nalog
            trenutni_zag = ""
            trenutni_datum = ""
            trenutni_laid = ""
            trenutni_grad = ""
            trenutni_posiljatelj = "Makromikro"
            trenutni_primatelj = ""
            trenutni_pob = ""
            trenutni_datum_isp = ""
            trenutni_suma_eur = ""
            
            # Privremena lista paleta za trenutni nalog
            palete_u_nalogu = []
            
            for line in lines:
                if "Pošiljka:" in line:
                    m_zag = re.search(r'(ZAG-\d+-\d+)', line)
                    if m_zag:
                        # Ako smo već skupljali palete za prethodni nalog, spremamo ih prije prelaza na novi
                        if trenutni_zag and palete_u_nalogu:
                            for idx, paleta in enumerate(palete_u_nalogu):
                                redak = {
                                    "Datum naloga": trenutni_datum,
                                    "Pošiljka": trenutni_zag,
                                    "LA-ID": trenutni_laid,
                                    "Grad (Iz Pariteta)": trenutni_grad,
                                    "Pošiljatelj": trenutni_posiljatelj,
                                    "Primatelj": trenutni_primatelj,
                                    "Poštanski broj": trenutni_pob,
                                    "Datum isporuke": trenutni_datum_isp,
                                    "Količina (Kom)": paleta["kol"],
                                    "Oznaka/Broj stavke": paleta["linija"],
                                    "Masa palete (KG)": paleta["masa"],
                                    "CBM palete": paleta["cbm"],
                                    "Suma EUR": trenutni_suma_eur if idx == 0 else ""
                                }
                                all_rows.append(redak)
                            palete_u_nalogu = []
                            
                        trenutni_zag = m_zag.group(1)
                
                if "Datum naloga:" in line:
                    m_dat = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                    if m_dat:
                        trenutni_datum = m_dat.group(1)
                
                if "LA-ID:" in line:
                    m_laid = re.search(r'(EP-\d+)', line)
                    if m_laid:
                        trenutni_laid = m_laid.group(1)
                        
                if "Paritet:" in line and "istovareno" in line.lower():
                    trenutni_grad = line.split("istovareno")[-1].strip()
                    
                if "Pošiljatelj" in line or "Pošiljatelj/dobavljač" in line:
                    pos_txt = line.replace("Pošiljatelj/dobavljač", "").replace("Pošiljatelj", "").strip()
                    if pos_txt and "makromikro" not in pos_txt.lower():
                        trenutni_posiljatelj = pos_txt
                    else:
                        trenutni_posiljatelj = "Makromikro"
                        
                if "Primatelj" in line:
                    trenutni_primatelj = line.replace("Primatelj", "").strip()
                    m_pob = re.search(r'HR-\s*(\d{5})', trenutni_primatelj)
                    if m_pob:
                        trenutni_pob = m_pob.group(1)
                    else:
                        m_5 = re.search(r'\b\d{5}\b', trenutni_primatelj)
                        if m_5:
                            trenutni_pob = m_5.group(0)
                            
                if "Datum isporuke:" in line:
                    m_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                    if m_isp:
                        trenutni_datum_isp = m_isp.group(1)
                        
                if "Suma EUR" in line:
                    m_s = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if m_s:
                        trenutni_suma_eur = m_s.group(1)

                # Hvatanje redaka paleta (sadrže tipove ambalaže ili oznaku otp i mase)
                if any(term in line.lower() for term in ["ewp", "fp", "cll", "otp", "handelsware"]) and not line.startswith("Suma"):
                    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                    if len(nums) >= 1:
                        try:
                            parts = line.split()
                            kol_val = "1"
                            for p in parts:
                                if p.isdigit() and len(p) <= 2:
                                    kol_val = p
                                    break
                            masa_val = nums[0]
                            cbm_val = nums[-1] if len(nums) > 1 else "0,00"
                            
                            palete_u_nalogu.append({
                                "kol": kol_val,
                                "linija": line,
                                "masa": masa_val,
                                "cbm": cbm_val
                            })
                        except:
                            pass
            
            # Spremamo preostale palete s ove stranice
            if trenutni_zag and palete_u_nalogu:
                for idx, paleta in enumerate(palete_u_nalogu):
                    redak = {
                        "Datum naloga": trenutni_datum,
                        "Pošiljka": trenutni_zag,
                        "LA-ID": trenutni_laid,
                        "Grad (Iz Pariteta)": trenutni_grad,
                        "Pošiljatelj": trenutni_posiljatelj,
                        "Primatelj": trenutni_primatelj,
                        "Poštanski broj": trenutni_pob,
                        "Datum isporuke": trenutni_datum_isp,
                        "Količina (Kom)": paleta["kol"],
                        "Oznaka/Broj stavke": paleta["linija"],
                        "Masa palete (KG)": paleta["masa"],
                        "CBM palete": paleta["cbm"],
                        "Suma EUR": trenutni_suma_eur if idx == 0 else ""
                    }
                    all_rows.append(redak)

        df = pd.DataFrame(all_rows)

        if not df.empty:
            st.success(f"Uspješno izvučeno ukupno stavki/paleta: {len(df)}")
            st.dataframe(df)

            file_name_xlsx = uploaded_file.name.replace(".pdf", "_sve_palete_direktno.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Posiljke')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke u dokumentu.")
