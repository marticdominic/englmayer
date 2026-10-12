import streamlit as st
from pdf2image import convert_from_bytes
import pytesseract
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer OCR Parser -> Excel", layout="wide")

st.title("📦 Englmayer OCR Parser (Sve palete u Excel)")
st.write("Učitajte PDF specifikaciju. Aplikacija automatski čita slikovni PDF putem OCR-a, izvlači sve stavke (palete), gradove, poštanske brojeve i priprema Excel.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Obrada dokumenata, OCR i izvlačenje stavki u tijeku... Molimo pričekajte..."):
        
        pdf_bytes = uploaded_file.read()
        full_text = ""
        
        # Pretvorba stranica u slike i OCR čitanje
        images = convert_from_bytes(pdf_bytes)
        for i, img in enumerate(images):
            ocr_text = pytesseract.image_to_string(img, lang='hrv')
            full_text += ocr_text + f"\n--STRANICA {i+1}--\n"

        # Dijeljenje teksta po blokovima naloga
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
                "Primatelj": "",
                "Poštanski broj": "",
                "Datum isporuke": "",
                "Suma EUR": ""
            }
            
            blok_tekst = " ".join(lines)
            
            # 1. Precizno izvlačenje datuma naloga
            match_datum = re.search(r'(\d{2}\.\d{2}\.\d{4}\.?)', lines[0])
            if match_datum:
                osnovni_podaci["Datum naloga"] = match_datum.group(1)
            else:
                match_d2 = re.search(r'(\d{2}\.\d{2}\.\d{4})', blok_tekst)
                if match_d2:
                    osnovni_podaci["Datum naloga"] = match_d2.group(1)

            # 2. Precizno izvlačenje broja pošiljke (isključivo ZAG broj)
            match_posiljka = re.search(r'(ZAG-\d+-\d+)', blok_tekst)
            if match_posiljka:
                osnovni_podaci["Pošiljka"] = match_posiljka.group(1)

            # 3. LA-ID
            match_laid = re.search(r'LA-ID:\s*(EP-\d+)', blok_tekst)
            if match_laid:
                osnovni_podaci["LA-ID"] = match_laid.group(1)

            paritet_tekst = ""
            for line in lines:
                if "Paritet:" in line:
                    paritet_tekst = line
                elif "Datum isporuke:" in line:
                    match_isp = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                    if match_isp:
                        osnovni_podaci["Datum isporuke"] = match_isp.group(1)

            # Izvlačenje grada iz pariteta
            if "istovareno" in paritet_tekst.lower():
                osnovni_podaci["Grad (Iz Pariteta)"] = paritet_tekst.split("istovareno")[-1].strip()

            # Primatelj
            primatelj_tekst = ""
            for i, line in enumerate(lines):
                if "Primatelj" in line:
                    rec = line.replace("Primatelj", "").strip()
                    if not rec and i + 1 < len(lines):
                        rec = lines[i+1]
                    primatelj_tekst = rec
                    osnovni_podaci["Primatelj"] = rec

            # Poštanski broj iz primatelja
            if primatelj_tekst:
                match_pob = re.search(r'HR-\s*(\d{5})', primatelj_tekst)
                if match_pob:
                    osnovni_podaci["Poštanski broj"] = match_pob.group(1)
                else:
                    match_5 = re.search(r'\b\d{5}\b', primatelj_tekst)
                    if match_5:
                        osnovni_podaci["Poštanski broj"] = match_5.group(0)

            # Financijski iznos (Suma EUR)
            suma_eur = ""
            match_suma = re.search(r'Suma EUR\D*(\d{1,3}(?:\.\d{3})*,\d{2})', blok_tekst)
            if match_suma:
                suma_eur = match_suma.group(1)
            osnovni_podaci["Suma EUR"] = suma_eur

            # Robusnije prepoznavanje SVIH stavki/paleta unutar pošiljke
            palete_pronadjene = False
            for line in lines:
                if not "Suma" in line and not "Opis usluge" in line and not "Datum" in line:
                    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                    if len(nums) >= 1:
                        if any(term in line for term in ["EWP", "FP", "CLL", "Handelsware", "KG"]) or len(nums) >= 2:
                            try:
                                masa_val = nums[0]
                                cbm_val = nums[-1] if len(nums) > 1 else "0,00"
                                
                                redak_palete = osnovni_podaci.copy()
                                redak_palete["Masa palete (KG)"] = masa_val
                                redak_palete["CBM palete"] = cbm_val
                                redak_palete["Detalj paketa"] = line
                                
                                all_rows.append(redak_palete)
                                palete_pronadjene = True
                            except:
                                pass
            
            if not palete_pronadjene:
                redak_palete = osnovni_podaci.copy()
                redak_palete["Masa palete (KG)"] = "0,00"
                redak_palete["CBM palete"] = "0,00"
                redak_palete["Detalj paketa"] = "Nije pronađeno"
                all_rows.append(redak_palete)

        df = pd.DataFrame(all_rows)

        if not df.empty:
            st.success(f"Uspješno izvučeno stavki/paleta: {len(df)}")
            st.dataframe(df)

            # Generiranje Excel datoteke za preuzimanje
            file_name_xlsx = uploaded_file.name.replace(".pdf", "_sve_palete_izvuceno.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Posiljke')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi Excel s paletama (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke.")
