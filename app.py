
import streamlit as st
import pandas as pd
import PyPDF2
from io import BytesIO
import openpyxl

st.set_page_config(page_title="Turni Notturni", layout="wide")

st.title("🗓️ Gestione Turni V5.2 (versione web)")
st.markdown("Carica il file PDF dei turni e il file Excel meta.xlsx per avviare l'elaborazione.")

# Upload dei file
pdf_file = st.file_uploader("📄 Carica il file PDF dei turni", type=["pdf"])
excel_file = st.file_uploader("📊 Carica il file meta.xlsx", type=["xlsx"])

# Avvio elaborazione
if st.button("🚀 Avvia Elaborazione"):
    if not pdf_file or not excel_file:
        st.warning("Carica sia il file PDF che il file Excel per procedere.")
    else:
        try:
            # Leggi PDF
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            # Carica Excel
            excel_data = pd.read_excel(excel_file, sheet_name=None)
            
            # Mostra contenuto PDF e fogli Excel
            with st.expander("📜 Testo estratto dal PDF"):
                st.text(text[:3000])  # Limitiamo preview

            with st.expander("📄 Fogli presenti in Excel"):
                for sheet_name, df in excel_data.items():
                    st.subheader(f"Foglio: {sheet_name}")
                    st.dataframe(df)

            st.success("✅ File elaborati correttamente! (mock)")

        except Exception as e:
            st.error(f"Errore durante l'elaborazione: {e}")
