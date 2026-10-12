import streamlit as st
from pdf2image import convert_from_bytes
import pytesseract
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer OCR Parser -> Excel", layout="wide")

st.title("📦 Englmayer OCR Parser (Slikovni PDF u Excel)")
st.write("Učitajte slikovnu PDF specifikaciju. Aplikacija će pročitati tekst sa slika putem OCR-a, izvući podatke i pripremiti Excel.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("OCR obrada slika u tijeku... Molimo pričekajte (ovo može potrajati ovisno o broju stranica)..."):
        
        pdf_bytes = uploaded_file.read()
        full_text = ""
        
        # Pretvorba svake stranice PDF-a u sliku i OCR čitanje
        images = convert_from_bytes(pdf_bytes)
        
        for i, img in enumerate(images):
            # Prepoznavanje teksta sa slike (hrvatski jezik)
            ocr_text = pytesseract.image_to_string(img, lang='hrv')
            full_text += ocr_text + f"\n--STRANICA {i+1}--\n"

        # Parsiranje blokova pošiljaka iz OCR teksta
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
                "Suma EUR": 0.0
            }
            
            # Datum naloga
            shipment_data["Datum naloga"] = lines[0].replace(".", "").strip()
            
            # Pošiljka, LA-ID, Referenca, Paritet, Datum isporuke
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
            st.success(f"Uspješno izvučeno pošiljaka pomoću OCR-a: {len(df)}")
            st.dataframe(df)

            # Excel izvoz
            file_name_xlsx = uploaded_file.name.replace(".pdf", "_ocr_izvuceno.xlsx")
            
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
            st.warning("OCR nije uspio prepoznat tekst pošiljaka. Provjerite izgled dokumenta.")
