import streamlit as st
from pypdf import PdfReader
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Text PDF Parser -> Excel", layout="wide")

st.title("📦 Englmayer Parser (Direktno iz tekstualnog PDF-a)")
st.write("Učitajte tekstualni PDF. Aplikacija čita dokument bez OCR-a, precizno hvata svaku stavku palete i financije.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Čitanje i parsiranje tekstualnog PDF-a u tijeku..."):
        
        # Čitanje PDF-a preko pypdf (direktno iz teksta)
        reader = PdfReader(uploaded_file)
        full_text = ""
        
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text += text + f"\n--STRANICA {i+1}--\n"

        # Vizualni pregled sirovog teksta
        with st.expander("🔍 Provjera izvučenog tekstualnog sadržaja"):
            st.text_area("Sirovi tekst iz PDF-a:", full_text, height=300)

        # Dijeljenje po blokovima naloga (prema "Datum naloga:")
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
            
            # 1. Datum naloga
            match_datum = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[0])
            if match_datum:
                osnovni_podaci["Datum naloga"] = match_datum.group(1)

            # 2. Broj pošiljke (ZAG)
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

            # 5. Primatelj i Poštanski broj
            primatelj_tekst = ""
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

            # 6. Suma EUR (Ukupan iznos s dna računa)
            suma_eur = ""
            match_suma = re.search(r'Suma EUR\D*(\d{1,3}(?:\.\d{3})*,\d{2})', blok_tekst)
            if match_suma:
                suma_eur = match_suma.group(1)
            osnovni_podaci["Suma EUR"] = suma_eur

            # 7. Čitanje svake stavke palete iz tablice (točno kao na slici: Oznaka, Količi, Pak, Masa, CBM)
            palete_pronadjene = False
            in_table = False
            
            for line in lines:
                if "Oznaka/Broj" in line or "Količi" in line:
                    in_table = True
                    continue
                if "Datum isporuke:" in line or "Opis usluge" in line or "Suma EUR" in line:
                    in_table = False
                
                # Unutar tabljice, preskačemo redak "Suma"
                if in_table and not line.startswith("Suma") and not "Suma EUR" in line:
                    # Tražimo brojeve s decimalama (Masa i CBM)
                    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d+', line)
                    if len(nums) >= 1:
                        try:
                            # Izvlačimo oznaku/broj (npr. "otp 71014") s početka retka
                            oznaka_match = re.split(r'\s+[1-9]\s+', line)
                            oznaka_palete = oznaka_match[0] if len(oznaka_match) > 0 else line
                            
                            masa_val = nums[0]
                            cbm_val = nums[-1] if len(nums) > 1 else "0,00"
                            
                            redak_palete = osnovni_podaci.copy()
                            redak_palete["Oznaka/Broj stavke"] = oznaka_palete
                            redak_palete["Masa palete (KG)"] = masa_val
                            redak_palete["CBM palete"] = cbm_val
                            
                            all_rows.append(redak_palete)
                            palete_pronadjene = True
                        except:
                            pass
            
            if not palete_pronadjene:
                redak_palete = osnovni_podaci.copy()
                redak_palete["Oznaka/Broj stavke"] = "Nije pronađeno"
                redak_palete["Masa palete (KG)"] = "0,00"
                redak_palete["CBM palete"] = "0,00"
                all_rows.append(redak_palete)

        df = pd.DataFrame(all_rows)

        if not df.empty:
            st.success(f"Uspješno izvučeno stavki/paleta: {len(df)}")
            st.dataframe(df)

            # Preuzimanje Excela
            file_name_xlsx = uploaded_file.name.replace(".pdf", "_tekstualni_parser.xlsx")
            
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
