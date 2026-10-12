import streamlit as st
from pdf2image import convert_from_bytes
import pytesseract
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer OCR Parser -> Excel", layout="wide")

st.title("📦 Englmayer OCR Parser (Grad i Poštanski broj)")
st.write("Učitajte PDF specifikaciju. Aplikacija automatski izvlači grad iz pariteta/reference, poštanski broj iz adrese primatelja te priprema Excel.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("OCR obrada slika i parsiranje podataka u tijeku... Molimo pričekajte..."):
        
        pdf_bytes = uploaded_file.read()
        full_text = ""
        
        # Pretvorba stranica u slike i OCR čitanje
        images = convert_from_bytes(pdf_bytes)
        
        for i, img in enumerate(images):
            ocr_text = pytesseract.image_to_string(img, lang='hrv')
            full_text += ocr_text + f"\n--STRANICA {i+1}--\n"

        # Parsiranje blokova pošiljaka
        blocks = full_text.split("Datum naloga:")
        shipments = []
        
        for block in blocks[1:]:
            lines = [line.strip() for line in block.split("\n") if line.strip()]
            if not lines:
                continue
                
            shipment_data = {
                "Datum naloga": "",
                "Pošiljka": "",
                "LA-ID": "",
                "Grad (Iz Pariteta)": "",
                "Primatelj": "",
                "Poštanski broj": "",
                "Ukupna Masa (KG)": 0.0,
                "Ukupni CBM": 0.0,
                "Datum isporuke": "",
                "Suma EUR": 0.0
            }
            
            # Datum naloga
            shipment_data["Datum naloga"] = lines[0].replace(".", "").strip()
            
            paritet_tekst = ""
            primatelj_tekst = ""

            for i, line in enumerate(lines):
                if "Pošiljka:" in line:
                    parts = line.split("LA-ID:")
                    shipment_data["Pošiljka"] = parts[0].replace("Pošiljka:", "").strip()
                    if len(parts) > 1:
                        shipment_data["LA-ID"] = parts[1].strip()
                elif line.startswith("LA-ID:") and not shipment_data["LA-ID"]:
                    shipment_data["LA-ID"] = line.replace("LA-ID:", "").strip()
                elif "Paritet:" in line:
                    paritet_tekst = line
                elif "Datum isporuke:" in line:
                    shipment_data["Datum isporuke"] = line.replace("Datum isporuke:", "").strip()

            # Izvlačenje grada iz pariteta (npr. "DPU - Isporučeno na mjesto istovareno Novalja" -> "Novalja")
            if paritet_tekst:
                # Tražimo riječ nakon zadnjeg pojavljivanja "istovareno" ili uzimamo zadnji dio
                if "istovareno" in paritet_tekst.lower():
                    grad_dio = paritet_tekst.split("istovareno")[-1].strip()
                    shipment_data["Grad (Iz Pariteta)"] = grad_dio

            # Primatelj
            for i, line in enumerate(lines):
                if "Primatelj" in line:
                    rec = line.replace("Primatelj", "").strip()
                    if not rec and i + 1 < len(lines):
                        rec = lines[i+1]
                    primatelj_tekst = rec
                    shipment_data["Primatelj"] = rec

            # Izvlačenje poštanskog broja iz adrese primatelja (npr. traži uzorak HR-xxxxx ili samo 5 znamenki)
            if primatelj_tekst:
                # Traži format HR-12345 ili slično
                match_pob = re.search(r'HR-\s*(\d{5})', primatelj_tekst)
                if match_pob:
                    shipment_data["Poštanski broj"] = match_pob.group(1)
                else:
                    # Alternativno traži bilo kojih 5 znamenki u nizu ako nema prefiksa HR-
                    match_5 = re.search(r'\b\d{5}\b', primatelj_tekst)
                    if match_5:
                        shipment_data["Poštanski broj"] = match_5.group(0)

            # Mase i CBM
            for line in lines:
                if line.startswith("Suma") and not "Suma EUR" in line:
                    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                    if len(nums) >= 2:
                        try:
                            shipment_data["Ukupna Masa (KG)"] = float(nums[-2].replace(".", "").replace(",", "."))
                            shipment_data["Ukupni CBM"] = float(nums[-1].replace(".", "").replace(",", "."))
                        except:
                            pass

            # Financijski iznos
            for line in lines:
                if "Suma EUR" in line:
                    match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if match:
                        shipment_data["Suma EUR"] = match.group(1)

            shipments.append(shipment_data)

        df = pd.DataFrame(shipments)

        if not df.empty:
            st.success(f"Uspješno izvučeno pošiljaka: {len(df)}")
            st.dataframe(df)

            # Excel izvoz
            file_name_xlsx = uploaded_file.name.replace(".pdf", "_grad_pob_izvuceno.xlsx")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Posiljke')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Preuzmi podatke u Excel formatu (.xlsx)",
                data=excel_data,
                file_name=file_name_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nije moguće pronaći podatke.")
