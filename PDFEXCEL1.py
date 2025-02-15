import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

def extract_text_and_tables(pdf_file):
    text_data = []
    tables_data = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # Extraer texto con formato
            text = page.extract_text()
            text_data.append([f"Página {page_num}", text])
            
            # Extraer tablas si existen
            tables = page.extract_tables()
            for table in tables:
                df_table = pd.DataFrame(table)
                tables_data.append((f"Tabla_Página_{page_num}", df_table))
    
    return text_data, tables_data

st.title("Convertir PDF a Excel con Formato")

uploaded_file = st.file_uploader("Sube un archivo PDF", type=["pdf"])
if uploaded_file:
    text_data, tables_data = extract_text_and_tables(uploaded_file)
    
    if text_data or tables_data:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Guardar texto en una hoja
            df_text = pd.DataFrame(text_data, columns=["Página", "Texto"])
            df_text.to_excel(writer, sheet_name="Texto", index=False)
            
            # Guardar cada tabla en una hoja diferente
            for sheet_name, df_table in tables_data:
                df_table.to_excel(writer, sheet_name=sheet_name[:30], index=False)
        
        output.seek(0)
        
        st.download_button(
            label="Descargar Excel",
            data=output,
            file_name="pdf_formato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Archivo convertido con éxito.")
