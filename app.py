import streamlit as st
from pdf2image import convert_from_bytes
import pytesseract
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer OCR Parser -> Excel", layout="wide")

st.title("📦 Englmayer OCR Parser (Razrada po paletama)")
st.write("Učitajte PDF specifikaciju. Aplikacija će za svaku paletu unutar pošiljke generirati zaseban redak u Excelu s pripadajućom kilažom.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Obrada dokumenata i razrada po paletama u tijeku..."):
        
        pdf_bytes = uploaded_file.read()
        full_text = ""
        
        images = convert_from_bytes(pdf_bytes)
        for i, img in enumerate(images):
            ocr_text = pytesseract.image_to_string(img, lang='hrv')
            full_text += ocr_text + f"\n--STRANICA {i+1}--\n"

        blocks = full_text.split("Datum naloga:")
        all_rows = []
        
        for block in blocks[1:]:
            lines = [line.strip() for line in block.split("\n") if line.strip()]
            if not lines:
                continue
                
            osnovni_podaci = {
                "Datum naloga": lines[0].replace(".", "").strip(),
                "Pošiljka": "",
                "LA-ID": "",
                "Grad (Iz Pariteta)": "",
                "Primatelj": "",
                "Poštanski broj": "",
                "Datum isporuke": "",
                "Suma EUR": ""
            }
            
            paritet_tekst = ""
            primatelj_tekst = ""

            for i, line in enumerate(lines):
                if "Pošiljka:" in line:
                    parts = line.split("LA-ID:")
                    osnovni_podaci["Pošiljka"] = parts[0].replace("Pošiljka:", "").strip()
                    if len(parts) > 1:
                        osnovni_podaci["LA-ID"] = parts[1].strip()
                elif line.startswith("LA-ID:") and not osnovni_podaci["LA-ID"]:
                    osnovni_podaci["LA-ID"] = line.replace("LA-ID:", "").strip()
                elif "Paritet:" in line:
                    paritet_tekst = line
                elif "Datum isporuke:" in line:
                    osnovni_podaci["Datum isporuke"] = line.replace("Datum isporuke:", "").strip()

            if "istovareno" in paritet_tekst.lower():
                osnovni_podaci["Grad (Iz Pariteta)"] = paritet_tekst.split("istovareno")[-1].strip()

            for i, line in enumerate(lines):
                if "Primatelj" in line:
                    rec = line.replace("Primatelj", "").strip()
                    if not rec and i + 1 < len(lines):
                        rec = lines[i+1]
                    primatelj_tekst = rec
                    osnovni_podaci["Primatelj"] = rec

            if primatelj_tekst:
                match_pob = re.search(r'HR-\s*(\d{5})', primatelj_tekst)
                if match_pob:
                    osnovni_podaci["Poštanski broj"] = match_pob.group(1)
                else:
                    match_5 = re.search(r'\b\d{5}\b', primatelj_tekst)
                    if match_5:
                        osnovni_podaci["Poštanski broj"] = match_5.group(0)

            # Financijski iznos
            suma_eur = ""
            for line in lines:
                if "Suma EUR" in line:
                    match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if match:
                        suma_eur = match.group(1)

            osnovni_podaci["Suma EUR"] = suma_eur

            # Pronalaženje svih pojedinačnih stavki (paleta) u pošiljci
            palete_pronadjene = False
            for line in lines:
                if any(pkg in line for pkg in ["EWP", "FP", "CLL"]) and not line.startswith("Suma"):
                    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                    if nums:
                        try:
                            masa_val = nums[0] # obično je masa prva
                            cbm_val = nums[-1] if len(nums) > 1 else "0,00"
                            
                            # Kreiramo kopiju osnovnih podataka za svaku paletu posebno
                            redak_palete = osnovni_podaci.copy()
                            redak_palete["Masa palete (KG)"] = masa_val
                            redak_palete["CBM palete"] = cbm_val
                            redak_palete["Detalj paketa"] = line
                            
                            all_rows.append(redak_palete)
                            palete_pronadjene = festgestellt = True
                        except:
                            pass
            
            # Ako iz nekog razloga ne nađe pojedinačne palete, doda osnovni zapis da se ne izgubi pošiljka
            if not palete_pronadjene:
                redak_palete = osnovni_podaci.copy()
                redak_palete["Masa palete (KG)"] = "0,00"
                redak_palete["CBM palete"] = "0,00"
                redak_palete["Detalj paketa"] = "Nije pronađeno"
                all_rows.append(redak_palete)

        df = pd.DataFrame(all_rows)

        if not df.empty:
            st.success(f"Uspješno obrađeno stavki/paleta: {len(df)}")
            st.dataframe(df)

            file_name_xlsx = uploaded_file.name.replace(".pdf", "_palete_izvuceno.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Pošiljke_Po_Paletama')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi podatke po paletama u Excel (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke.")
