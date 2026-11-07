import os
import json
import io
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(page_title="AI Čitač Transportnih Računa po Blokovima", layout="wide")

st.title("📦 AI Čitač Transportnih Računa (Metoda po Blokovima od 30 pošiljaka)")
st.write("Ova verzija dijeli dokument u manje dijelove kako bi izvukla 100% točne podatke za svaku paletu, masu i CBM bez preskakanja.")

api_key = st.text_input("Unesite svoj Google Gemini API ključ:", type="password")
uploaded_file = st.file_uploader("Odaberite PDF račun", type=["pdf"])

if uploaded_file is not None and api_key:
    pdf_bytes = uploaded_file.read()
    temp_path = "temp_invoice.pdf"
    with open(temp_path, "wb") as f:
        f.write(pdf_bytes)
    
    st.info("PDF je uspješno učitan. Možete pokrenuti detaljnu obradu po blokovima.")
    
    # Odabir veličine bloka
    velicina_bloka = st.slider("Koliko pošiljaka želite obraditi u jednom koraku?", min_value=10, max_value=50, value=30, step=10)
    
    if st.button("🚀 Pokreni pametnu ekstrakciju po blokovima"):
        try:
            client = genai.Client(api_key=api_key)
            
            with st.spinner("Učitavam dokument u Gemini AI..."):
                uploaded_pdf = client.files.upload(file=temp_path)
            
            all_data = []
            
            # Prvo tražimo od AI-ja listu svih pošiljaka u dokumentu da znamo ukupan broj
            with st.spinner("Analiziram strukturu računa i tražim popis svih pošiljaka..."):
                index_prompt = """
                Pregledaj ovaj PDF dokument i poprodi sve pošiljke redom. 
                Vrati mi isključivo JSON listu svih brojeva pošiljaka (npr. ["ZAG-000515836-1", "ZAG-000515837-1", ...]) 
                onim redoslijedom kako se pojavljuju u dokumentu.
                """
                index_response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=[uploaded_pdf, index_prompt],
                    config=types.GenerateContentConfig(
                        temperature=0.0,
                        response_mime_type="application/json"
                    )
                )
                popis_posiljaka = json.loads(index_response.text)
            
            total_posiljaka = len(popis_posiljaka)
            st.success( ukupno_p := f"Pronađeno ukupno {total_posiljaka} pošiljaka u dokumentu.")
            
            # Dijelimo pošiljke u manje blokove (npr. po 30)
            blokovi = [popis_posiljaka[i:i + velicina_bloka] for i in range(0, total_posiljaka, velicina_bloka)]
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, blok in enumerate(blokovi):
                status_text.text(f"Obrađujem blok {idx+1} od {len(blokovi)} (pošiljke {idx*velicina_bloka + 1} do {min((idx+1)*velicina_bloka, total_posiljaka)})...")
                
                blok_prompt = f"""
                Ti si sustav za ekstrakciju logističkih podataka. 
                Iz ovog PDF dokumenta izvuci detaljne podatke ISKLJUČIVO za sljedeće pošiljke: {json.dumps(blok)}.
                
                Za svaku od ovih pošiljaka vrati valjani JSON format sa strukturiranom listom:
                [
                  {{
                    "broj_posiljke": "...",
                    "la_id": "...",
                    "datum_naloga": "...",
                    "datum_isporuke": "...",
                    "primatelj": "...",
                    "palete": [
                      {{
                        "oznaka": "...",
                        "vrsta_pakiranja": "...",
                        "masa_kg": 0.0,
                        "cbm": 0.0
                      }}
                    ],
                    "ukupna_cijena_eur": 0.0
                  }}
                ]
                Nemoj preskočiti niti jednu stavku palete za ove pošiljke. Vrati samo traženi JSON.
                """
                
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=[uploaded_pdf, blok_prompt],
                    config=types.GenerateContentConfig(
                        temperature=0.0,
                        response_mime_type="application/json"
                    )
                )
                
                try:
                    blok_data = json.loads(response.text)
                    if isinstance(blok_data, list):
                        all_data.extend(blok_data)
                except Exception as e:
                    st.warning(f"Greška pri parsiranju bloka {idx+1}: {e}")
                
                progress_bar.progress((idx + 1) / len(blokovi))
            
            status_text.text("Obrada svih blokova je završena! Pripremam Excel tablicu...")
            
            # Pretvaranje u ravnu tablicu za Excel
            flat_rows = []
            for item in all_data:
                posiljka = item.get("broj_posiljke", "")
                la_id = item.get("la_id", "")
                d_naloga = item.get("datum_naloga", "")
                d_isporuke = item.get("datum_isporuke", "")
                primatelj = item.get("primatelj", "")
                ukupna_cijena = item.get("ukupna_cijena_eur", "")
                
                palete = item.get("palete", [])
                if not palete:
                    flat_rows.append({
                        "Broj pošiljke": posiljka,
                        "LA-ID": la_id,
                        "Datum naloga": d_naloga,
                        "Datum isporuke": d_isporuke,
                        "Primatelj": primatelj,
                        "Oznaka palete": "",
                        "Vrsta pakiranja": "",
                        "Masa (kg)": 0,
                        "CBM": 0,
                        "Ukupna cijena (EUR)": ukupna_cijena
                    })
                else:
                    for p in palete:
                        flat_rows.append({
                            "Broj pošiljke": posiljka,
                            "LA-ID": la_id,
                            "Datum naloga": d_naloga,
                            "Datum isporuke": d_isporuke,
                            "Primatelj": primatelj,
                            "Oznaka palete": p.get("oznaka", ""),
                            "Vrsta pakiranja": p.get("vrsta_pakiranja", ""),
                            "Masa (kg)": p.get("masa_kg", 0),
                            "CBM": p.get("cbm", 0),
                            "Ukupna cijena (EUR)": ukupna_cijena
                        })
            
            df = pd.DataFrame(flat_rows)
            st.success(f"Uspješno spojeno! Ukupno obrađeno redaka paleta: {len(df)}")
            
            # Prikaz tablice
            st.dataframe(df)
            
            # Generiranje Excel datoteke
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sve_Palete_Spojeno')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Preuzmi jedinstvenu Excel tablicu (.xlsx)",
                data=processed_data,
                file_name="sve_posiljke_palete_spojeno.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Došlo je do pogreške tijekom izvršavanja: {e}")

elif not api_key and uploaded_file:
    st.warning("Unesite Google Gemini API ključ iznad.")
