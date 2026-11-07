import os
import json
import io
import pandas as pd
import streamlit as st
import pypdf
from google import genai
from google.genai import types

st.set_page_config(page_title="AI Čitač PDF Računa po Stranicama", layout="wide")

st.title("📦 AI Čitač Transportnih Računa (Metoda po Stranicama)")
st.write("Ova verzija dijeli PDF na manje dijelove (po stranicama) kako ne bi preopteretila AI i kako bi izvući sve palete bez greške.")

api_key = st.text_input("Unesite svoj Google Gemini API ključ:", type="password")
uploaded_file = st.file_uploader("Odaberite PDF račun", type=["pdf"])

if uploaded_file is not None and api_key:
    pdf_bytes = uploaded_file.read()
    
    # Čitamo ukupan broj stranica koristeći pypdf
    pdf_reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    total_pages = len(pdf_reader.pages)
    
    st.info(f"PDF je učitan. Dokument ima ukupno {total_pages} stranica.")
    
    pages_per_chunk = st.slider("Koliko stranica obraditi u jednom koraku?", min_value=1, max_value=5, value=2, step=1)
    
    if st.button("🚀 Pokreni stabilnu obradu po stranicama"):
        try:
            client = genai.Client(api_key=api_key)
            all_data = []
            
            # Računamo raspone stranica (npr. 1-2, 3-4, 5-6...)
            chunks = []
            for i in range(0, total_pages, pages_per_chunk):
                chunks.append((i, min(i + pages_per_chunk, total_pages)))
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, (start_page, end_page) in enumerate(chunks):
                status_text.text(f"Obrađujem stranice {start_page + 1} do {end_page} od ukupno {total_pages}...")
                
                # Izvađamo samo određene stranice u novi privremeni PDF
                writer = pypdf.PdfWriter()
                for p in range(start_page, end_page):
                    writer.add_page(pdf_reader.pages[p])
                
                chunk_pdf_bytes = io.BytesIO()
                writer.write(chunk_pdf_bytes)
                chunk_pdf_bytes.seek(0)
                
                # Spremano privremeno za upload na Gemini
                temp_chunk_path = f"temp_chunk_{idx}.pdf"
                with open(temp_chunk_path, "wb") as f:
                    f.write(chunk_pdf_bytes.read())
                
                # Upload samo te 2-3 stranice na Gemini API
                uploaded_chunk = client.files.upload(file=temp_chunk_path)
                
                chunk_prompt = """
                Ti si sustav za ekstrakciju logističkih i transportnih podataka iz računa.
                Iz ovog dijela PDF-a (ovih nekoliko stranica) izvuci sve pošiljke i sve pojedinačne palete.
                
                Vrati isključivo čisti VALJANI JSON format kao listu objekata:
                [
                  {
                    "broj_posiljke": "npr. ZAG-000515836-1",
                    "la_id": "npr. EP-499113899",
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
                
                Ako na ovim stranicama nema pošiljki ili su samo opći podaci o računu, vrati praznu listu [].
                """
                
                try:
                    response = client.models.generate_content(
                        model='gemini-3.8-flash',
                        contents=[uploaded_chunk, chunk_prompt],
                        config=types.GenerateContentConfig(
                            temperature=0.0,
                            response_mime_type="application/json"
                        )
                    )
                    
                    page_data = json.loads(response.text)
                    if isinstance(page_data, list):
                        all_data.extend(page_data)
                except Exception as e:
                    st.warning(f"Upozorenje za stranice {start_page + 1}-{end_page}: {e}")
                
                # Brišemo privremenu datoteku
                if os.path.exists(temp_chunk_path):
                    os.remove(temp_chunk_path)
                
                progress_bar.progress((idx + 1) / len(chunks))
            
            status_text.text("Završena obrada svih stranica! Pripremam jedinstvenu Excel tablicu...")
            
            # Pretvaranje svih prikupljenih podataka u jednu Excel tablicu
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
            st.success(f"Uspješno obrađeno! Pronađeno ukupno {len(all_data)} pošiljaka s {len(df)} redaka paleta.")
            
            st.dataframe(df)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sve_Palete')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Preuzmi spojenu Excel tablicu (.xlsx)",
                data=processed_data,
                file_name="spojeni_racun_palete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Došlo je do greške: {e}")

elif not api_key and uploaded_file:
    st.warning("Molimo unesite API ključ.")
