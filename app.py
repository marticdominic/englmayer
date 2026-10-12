import streamlit as st
from pypdf import PdfReader
import pandas as pd
import re
import io

st.set_page_config(page_title="Englmayer Parser -> Excel", layout="wide")

st.title("📦 Englmayer Parser (Fino podešavanje za 69 paleta)")
st.write("Učitajte tekstualni PDF. Prošireno prepoznavanje stavki i masa.")

uploaded_file = st.file_uploader("Odaberite PDF datoteku", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Obrada dokumenta u tijeku..."):
        
        reader = PdfReader(uploaded_file)
        full_text = ""
        
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                full_text += text + f"\n--STRANICA {i+1}--\n"

        blocks = re.split(r'(Pošiljka:\s*ZAG-\d+-\d+)', full_text)
        all_rows = []
        
        for i in range(1, len(blocks), 2):
            header = blocks[i]
            content = blocks[i+1] if i+1 < len(blocks) else ""
            
            if "Broj Računa:" in content:
                content = content.split("Broj Računa:")[0]

            lines = [line.strip() for line in (header + "\n" + content).split("\n") if line.strip()]
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
            
            match_posiljka = re.search(r'(ZAG-\d+-\d+)', header)
            if match_posiljka:
                osnovni_podaci["Pošiljka"] = match_posiljka.group(1)

            match_datum = re.search(r'Datum naloga:\s*(\d{2}\.\d{2}\.\d{4})', blok_tekst)
            if match_datum:
                osnovni_podaci["Datum naloga"] = match_datum.group(1)

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

            if "istovareno" in paritet_tekst.lower():
                osnovni_podaci["Grad (Iz Pariteta)"] = paritet_tekst.split("istovareno")[-1].strip()

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

            if primatelj_tekst:
                match_pob = re.search(r'HR-\s*(\d{5})', primatelj_tekst)
                if match_pob:
                    osnovni_podaci["Poštanski broj"] = match_pob.group(1)
                else:
                    match_5 = re.search(r'\b\d{5}\b', primatelj_tekst)
                    if match_5:
                        osnovni_podaci["Poštanski broj"] = match_5.group(0)

            suma_eur = ""
            match_suma = re.search(r'Suma EUR\D*(\d{1,3}(?:\.\d{3})*,\d{2})', blok_tekst)
            if match_suma:
                suma_eur = match_suma.group(1)

            palete_u_ovom_bloku = []
            
            for line in lines:
                line_lower = line.lower()
                
                # Izričito mičemo samo financije, zaglavlja i rekapitulacije
                if any(k in line_lower for k in ["suma", "neto iznos", "temeljni kapital", "ukupno", "pdv", "uplaćen", "paritet:", "datum isporuke:", "la-id:", "pošiljatelj", "primatelj", "opis usluge"]):
                    continue
                
                if line_lower.startswith("110") or line_lower.startswith("197") or line_lower.startswith("199"):
                    continue

                # Fleksibilnije traženje brojeva (masa s zarezom ili točkom)
                nums = re.findall(r'\d{1,3}(?:[.,]\d{3})*[.,]\d+', line)
                
                # Ako redak sadrži broj ili ambalažu, smatramo ga stavkom
                if nums or any(term in line_lower for term in ["ewp", "fp", "cll", "otp", "handelsware"]):
                    try:
                        parts = line.split()
                        kol_val = "1"
                        for p in parts:
                            if p.isdigit() and len(p) <= 2:
                                kol_val = p
                                break

                        masa_val = nums[0] if nums else "0,00"
                        cbm_val = nums[-1] if len(nums) > 1 else "0,00"
                        
                        palete_u_ovom_bloku.append({
                            "Količina (Kom)": kol_val,
                            "Oznaka/Broj stavke": line,
                            "Masa palete (KG)": masa_val,
                            "CBM palete": cbm_val
                        })
                    except:
                        pass

            if palete_u_ovom_bloku:
                for idx, paleta in enumerate(palete_u_ovom_bloku):
                    redak = osnovni_podaci.copy()
                    redak.update(paleta)
                    
                    if idx == 0:
                        redak["Suma EUR"] = suma_eur
                    else:
                        redak["Suma EUR"] = ""
                        
                    all_rows.append(redak)

        df = pd.DataFrame(all_rows)

        if not df.empty:
            st.success(f"Uspješno izvučeno stavki/paleta: {len(df)}")
            st.dataframe(df)

            file_name_xlsx = uploaded_file.name.replace(".pdf", "_englmayer_69.xlsx")
            
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
