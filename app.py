import os
import json
import io
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(page_title="AI Čitač Transportnih Računa", layout="wide")

st.title("📦 AI Čitač Transportnih Računa i Paleta")
st.write("Učitajte PDF račun (poput Englmayer specifikacije), a AI će automatski izvući sve pošiljke, palete, mase, CBM i cijene u Excel tablicu.")

api_key = st.text_input("Unesite svoj Google Gemini API ključ:", type="password")

uploaded_file = st.file_uploader("Odaberite PDF račun", type=["pdf"])

if uploaded_file is not None and api_key:
    # Spremanje privremenog fajla u memoriju
    pdf_bytes = uploaded_file.read()
    temp_path = "temp_invoice.pdf"
    with open(temp_path, "wb") as f:
        f.write(pdf_bytes)
    
    st.info("PDF je učitan. Pokrećem AI obradu...")
    
    if st.button("Pokreni AI analizu dokumenta"):
        try:
            client = genai.Client(api_key=api_key)
            
            with st.spinner("Učitavam dokument u Gemini AI..."):
                uploaded_pdf = client.files.upload(file=temp_path)
            
            prompt = """
            Ti si stručni sustav za ekstrakciju podataka iz logističkih i transportnih računa (poput Englmayer računa).
            Pregledaj ovaj PDF dokument u cijelosti (sve stranice) i pronađi svaku pojedinačnu pošiljku.
            Za svaku pošiljku izvuci detalje i vrati isključivo VALJANI JSON format sa sljedećom strukturom:
            [
              {
                "broj_posiljke": "ZAG-...",
                "la_id": "EP-...",
                "datum_naloga": "DD.MM.YYYY.",
                "datum_isporuke": "DD.MM.YYYY.",
                "primatelj": "Naziv i grad",
                "palete": [
                  {
                    "oznaka": "OTP...",
                    "vrsta_pakiranja": "FP ili EWP ili CLL",
                    "masa_kg": 123.45,
                    "cbm": 1.234
                  }
                ],
                "ukupna_cijena_eur": 123.45
              }
            ]
            Pazi da obradiš SVE pošiljke i sve stavke paleta iz cijelog dokumenta bez preskakanja.
            """
            
            with st.spinner("AI analizira stavke, palete, mase i volumene... Molimo pričekajte."):
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=[uploaded_pdf, prompt],
                    config=types.GenerateContentConfig(
                        temperature=0.0,
                        response_mime_type="application/json"
                    )
                )
            
            data = json.loads(response.text)
            
            flat_rows = []
            for item in data:
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
            st.success(f"Uspješno obrađeno! Pronađeno pošiljaka: {len(data)}, ukupno redaka paleta: {len(df)}")
            
            st.dataframe(df)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Palete_Specifikacija')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Preuzmi Excel tablicu (.xlsx)",
                data=processed_data,
                file_name="analiza_racuna_paleta.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Došlo je do pogreške: {e}")

elif not api_key and uploaded_file:
    st.warning("Unesite Google Gemini API ključ iznad kako biste započeli obradu.")
