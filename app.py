import streamlit as st
from pypdf import PdfReader
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Text PDF Parser -> Excel", layout="wide")

st.title("📦 Englmayer Parser (Direktno hvatanje svih paleta)")
st.write("Učitajte tekstualni PDF. Aplikacija automatski pronalazi sve palete prema tipu pakiranja i masi.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Čitanje i precizno parsiranje dokumenta u tijeku..."):
        
        reader = PdfReader(uploaded_file)
        full_text = ""
        
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text += text + f"\n--STRANICA {i+1}--\n"

        # Dijelimo tekst prema početku svake pošiljke (ZAG broj)
        blocks = re.split(r'(Pošiljka:\s*ZAG-\d+-\d+)', full_text)
        
        all_rows = []
        
        for i in range(1, len(blocks), 2):
            header = blocks[i]
            block_content = blocks[i+1] if i+1 < len(blocks) else ""
            lines = [line.strip() for line in (header + "\n" + block_content).split("\n") if line.strip()]
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
            match_datum = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4})', blok_tekst)
            if match_datum:
                osnovni_podaci["Datum naloga"] = match_datum.group(1)
            else:
                match_d2 = re.search(r'(\d{2}\.\d{2}\.\d{4})', blok_tekst)
                if match_d2:
                    osnovni_podaci["Datum naloga"] = match_d2.group(1)

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
            
            for line in lines:
                if "Pošiljatelj/dobavljač" in line or "Pošiljatelj" in line:
                    posiljatelj_tekst = line.replace("Pošiljatelj/dobavljač", "").replace("Pošiljatelj", "").strip()
                if "Primatelj" in line:
                    primatelj_tekst = line.replace("Primatelj", "").strip()

            if posiljatelj_tekst and "makromikro" not in posiljatelj_tekst.lower():
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

            # 7. Čitanje paleta: tražimo retke s paletama bez obzira na zaglavlje tablice
            palete_u_ovom_bloku = []
            
            for line in lines:
                # Preskačemo redove s troškovima usluga (110, 197) i općenite sume na dnu
                if line.startswith("110") or line.startswith("197") or "Suma EUR" in line or "Broj Računa:" in line:
                    continue
                
                # Uvjet za prepoznavanje retka palete: sadrži tip ambalaže ili oznaku, te brojčanu masu
                if any(term in line.lower() for term in ["ewp", "fp", "cll", "otp", "handelsware", "groupage"]):
                    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                    if len(nums) >= 1:
                        try:
                            # Izvlačenje količine komada (tražimo jednoznamenkasti broj s početka retka)
                            parts = line.split()
                            kol_val = "1"
                            for p in parts:
                                if p.isdigit() and len(p) <= 2:
                                    kol_val = p
                                    break

                            masa_val = nums[0]
                            cbm_val = nums[-1] if len(nums) > 1 else "0,00"
                            
                            palete_u_ovom_bloku.append({
                                "Količina (Kom)": kol_val,
                                "Oznaka/Broj stavke": line,
                                "Masa palete (KG)": masa_val,
                                "CBM palete": cbm_val
                            })
                        except:
                            pass

            # Dodavanje u konačnu listu i raspodjela sume EUR samo na prvu paletu
            if palete_u_ovom_bloku:
                for idx, paleta in enumerate(palete_u_ovom_bloku):
                    redak = osnovni_podaci.copy()
                    redak.update(paleta)
                    
                    if idx == 0:
                        redak["Suma EUR"] = suma_eur
                    else:
                        redak["Suma EUR"] = ""
                        
                    all_rows.append(redak)
            else:
                redak = osnovni_podaci.copy()
                redak["Količina (Kom)"] = "0"
                redak["Oznaka/Broj stavke"] = "Nije pronađeno"
                redak["Masa palete (KG)"] = "0,00"
                redak["CBM palete"] = "0,00"
                redak["Suma EUR"] = suma_eur
                all_rows.append(redak)

        df = pd.DataFrame(all_rows)

        if not df.empty:
            st.success(f"Uspješno izvučeno stavki/paleta: {len(df)}")
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
