import streamlit as st
from pypdf import PdfReader
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Text PDF Parser -> Excel", layout="wide")

st.title("📦 Englmayer Parser (Sa stupcem količine paleta)")
st.write("Učitajte tekstualni PDF. Aplikacija izvlači palete, količinu, razdvaja troškove i uređuje pošiljatelje.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Čitanje i precizno parsiranje dokumenta u tijeku..."):
        
        reader = PdfReader(uploaded_file)
        full_text = ""
        
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text += text + f"\n--STRANICA {i+1}--\n"

        with st.expander("🔍 Provjera izvučenog tekstualnog sadržaja"):
            st.text_area("Sirovi tekst iz PDF-a:", full_text, height=300)

        blocks = full_text.split("Datum naloga:")
        all_rows = []
        
        for block in blocks[1:]:
            lines = [line.strip() for line in block.split("\n") if line.strip()]
            if not lines:
                continue
                
            osnovni_podaci = {
                "Datum naloga": "",
                "Pošiljka": "",
                "LA-ID": "",
                "Grad (Iz Pariteta)": "",
                "Pošiljatelj": "",
                "Primatelj": "",
                "Poštanski broj": "",
                "Datum isporuke": "",
                "Suma EUR": ""
            }
            
            blok_tekst = " ".join(lines)
            
            # 1. Datum naloga
            match_datum = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[0])
            if match_datum:
                osnovni_podaci["Datum naloga"] = match_datum.group(1)

            # 2. Pošiljka (ZAG broj)
            match_posiljka = re.search(r'(ZAG-\d+-\d+)', blok_tekst)
            if match_posiljka:
                osnovni_podaci["Pošiljka"] = match_posiljka.group(1)

            # 3. LA-ID
            match_laid = re.search(r'LA-ID:\s*(EP-\d+)', blok_tekst)
            if match_laid:
                osnovni_podaci["LA-ID"] = match_laid.group(1)

            # 4. Paritet / Grad
            paritet_tekst = ""
            for line in lines:
                if "Paritet:" in line:
                    paritet_tekst = line
                elif "Datum isporuke:" in line:
                    match_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                    if match_isp:
                        osnovni_podaci["Datum isporuke"] = match_isp.group(1)

            if "istovareno" in paritet_tekst.lower():
                osnovni_podaci["Grad (Iz Pariteta)"] = paritet_tekst.split("istovareno")[-1].strip()

            # 5. Pošiljatelj i Primatelj
            posiljatelj_tekst = ""
            primatelj_tekst = ""
            
            for i, line in enumerate(lines):
                if "Pošiljatelj/dobavljač" in line or "Pošiljatelj" in line:
                    pos_rec = line.replace("Pošiljatelj/dobavljač", "").replace("Pošiljatelj", "").strip()
                    if not pos_rec and i + 1 < len(lines):
                        pos_rec = lines[i+1]
                    posiljatelj_tekst = pos_rec
                
                if "Primatelj" in line:
                    rec = line.replace("Primatelj", "").strip()
                    if not rec and i + 1 < len(lines):
                        rec = lines[i+1]
                    primatelj_tekst = rec

            if posiljatelj_tekst and "makromikro" not in posiljatelj_tekst.lower():
                osnovni_podaci["Pošiljka"] = osnovni_podaci["Pošiljka"] # drži original
                osnovni_podaci["Pošiljatelj"] = posiljatelj_tekst
            else:
                osnovni_podaci["Pošiljatelj"] = "Makromikro"

            osnovni_podaci["Primatelj"] = primatelj_tekst

            # Poštanski broj iz primatelja
            if primatelj_tekst:
                match_pob = re.search(r'HR-\s*(\d{5})', primatelj_tekst)
                if match_pob:
                    osnovni_podaci["Poštanski broj"] = match_pob.group(1)
                else:
                    match_5 = re.search(r'\b\d{5}\b', primatelj_tekst)
                    if match_5:
                        osnovni_podaci["Poštanski broj"] = match_5.group(0)

            # 6. Suma EUR
            suma_eur = ""
            match_suma = re.search(r'Suma EUR\D*(\d{1,3}(?:\.\d{3})*,\d{2})', blok_tekst)
            if match_suma:
                suma_eur = match_suma.group(1)
            osnovni_podaci["Suma EUR"] = suma_eur

            # 7. Čitanje isključivo tablice paleta uz izvlačenje količine
            palete_pronadjene = False
            in_palete_table = False
            
            for line in lines:
                if "Oznaka/Broj" in line:
                    in_palete_table = True
                    continue
                if "Datum isporuke:" in line or "Opis usluge" in line or "Broj Računa:" in line:
                    in_palete_table = False
                
                if in_palete_table:
                    if line.startswith("Suma") or "Oznaka/Broj" in line or "Pak." in line:
                        continue
                    
                    if any(term in line for term in ["EWP", "FP", "CLL", "otp", "OTP", "Handelsware"]):
                        nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                        if len(nums) >= 1:
                            try:
                                # Pokušavamo pronaći količinu (broj komada) u retku (npr. broj 1, 2, 3...)
                                kolicina_match = re.search(r'\b([1-9])\b', line)
                                kol_val = kolicina_match.group(1) if kolicina_match else "1"
                                
                                masa_val = nums[0]
                                cbm_val = nums[-1] if len(nums) > 1 else "0,00"
                                
                                redak_palete = osnovni_podaci.copy()
                                redak_palete["Količina (Kom)"] = kol_val
                                redak_palete["Oznaka/Broj stavke"] = line
                                redak_palete["Masa palete (KG)"] = masa_val
                                redak_palete["CBM palete"] = cbm_val
                                
                                all_rows.append(redak_palete)
                                palete_pronadjene = True
                            except:
                                pass
            
            if not palete_pronadjene:
                redak_palete = osnovni_podaci.copy()
                redak_palete["Količina (Kom)"] = "0"
                redak_palete["Oznaka/Broj stavke"] = "Nije pronađeno"
                redak_palete["Masa palete (KG)"] = "0,00"
                redak_palete["CBM palete"] = "0,00"
                all_rows.append(redak_palete)

        df = pd.DataFrame(all_rows)

        if not df.empty:
            st.success(f"Uspješno izvučeno stavki/paleta: {len(df)}")
            st.dataframe(df)

            file_name_xlsx = uploaded_file.name.replace(".pdf", "_sa_kolicinom.xlsx")
            
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
