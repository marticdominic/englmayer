import streamlit as st
import pdfplumber
from pdf2image import convert_from_bytes
import pytesseract
from PIL import Image
import pandas as pd
import re
import io

st.set_page_config(page_title="Hibridni Englmayer PDF/OCR Parser", layout="wide")

st.title("📦 Hibridni Parser (Tekst + OCR Slike) za Englmayer PDF")
st.write("Učitajte PDF specifikaciju. Aplikacija će automatski pročitati i tekstualni sloj i skenirane dijelove (slike) pomoću OCR-a te izvući podatke o pošiljkama.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Čitanje PDF-a (tekst + OCR analiza slika)...jesite li strpljivi..."):
        
        pdf_bytes = uploaded_file.read()
        full_text = ""
        
        # 1. Pokušaj čitanja preko pdfplumber-a (tekstualni sloj)
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text and len(text.strip()) > 50:
                    full_text += text + "\n--STRANICA--\n"
                else:
                    # Ako je stranica prazna ili slika, koristi OCR preko pdf2image i pytesseract
                    images = convert_from_bytes(pdf_bytes, first_page=i+1, last_page=i+1)
                    for img in images:
                        ocr_text = pytesseract.image_to_string(img, lang='hrv') # hrvatski jezik za prepoznavanje znakova
                        full_text += ocr_text + "\n--STRANICA (OCR)--\n"

        # Dijeljenje teksta po blokovima naloga
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
                "Referenca": "",
                "Paritet": "",
                "Primatelj": "",
                "Ukupna Masa (KG)": 0.0,
                "Ukupni CBM": 0.0,
                "Datum isporuke": "",
                "Iznos Prijevoza (EUR)": 0.0,
                "Iznos Goriva/Dodataka (EUR)": 0.0,
                "Suma EUR": 0.0
            }
            
            # Datum naloga
            shipment_data["Datum naloga"] = lines[0].replace(".", "").strip()
            
            # Pošiljka i LA-ID
            for i, line in enumerate(lines):
                if "Pošiljka:" in line:
                    parts = line.split("LA-ID:")
                    shipment_data["Pošiljka"] = parts[0].replace("Pošiljka:", "").strip()
                    if len(parts) > 1:
                        shipment_data["LA-ID"] = parts[1].strip()
                elif line.startswith("LA-ID:") and not shipment_data["LA-ID"]:
                    shipment_data["LA-ID"] = line.replace("LA-ID:", "").strip()
                elif "Referenca:" in line and i + 1 < len(lines):
                    shipment_data["Referenca"] = lines[i+1].strip()
                elif "Paritet:" in line:
                    shipment_data["Paritet"] = line.replace("Paritet:", "").strip()
                elif "Datum isporuke:" in line:
                    shipment_data["Datum isporuke"] = line.replace("Datum isporuke:", "").strip()

            # Primatelj
            for i, line in enumerate(lines):
                if "Primatelj" in line:
                    rec = line.replace("Primatelj", "").strip()
                    if not rec and i + 1 < len(lines):
                        rec = lines[i+1]
                    shipment_data["Primatelj"] = rec

            # Ekstrakcija masa i CBM-a
            for line in lines:
                if line.startswith("Suma") and not "Suma EUR" in line:
                    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                    if len(nums) >= 2:
                        try:
                            shipment_data["Ukupna Masa (KG)"] = float(nums[-2].replace(".", "").replace(",", "."))
                            shipment_data["Ukupni CBM"] = float(nums[-1].replace(".", "").replace(",", "."))
                        except:
                            pass

            # Financije
            for line in lines:
                if "Suma EUR" in line:
                    match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if match:
                        shipment_data["Suma EUR"] = match.group(1)

            shipments.append(shipment_data)

        df = pd.DataFrame(shipments)

        st.success(f"Obrada završena! Pronađeno pošiljaka: {len(df)}")
        st.dataframe(df)

        # Excel preuzimanje
        file_name_xlsx = uploaded_file.name.replace(".pdf", "_hibridno_izvuceno.xlsx")
        
        @st.cache_data
        def convert_df_to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Posiljke')
            return output.getvalue()

        excel_data = convert_df_to_excel(df)

        st.download_button(
            label="📥 Preuzmi podatke u Excel formatu (.xlsx)",
            data=excel_data,
            file_name=file_name_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
