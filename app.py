import streamlit as st
import pandas as pd
import io
import os

# Configuración básica de la página
st.set_page_config(page_title="Unificador de Excel", page_icon="📊")

st.title("📊 Unificador de Archivos Excel")
st.write("Sube múltiples archivos Excel. Se generará un solo archivo donde cada hoja tendrá el nombre del archivo original.")

# Subida de archivos (permite múltiples archivos xls y xlsx)
uploaded_files = st.file_uploader(
    "Sube tus archivos Excel aquí", 
    type=['xls', 'xlsx'], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Has subido {len(uploaded_files)} archivos.")
    
    if st.button("Unificar Archivos", type="primary"):
        # Crear un buffer en memoria para no guardar archivos temporales en el servidor
        output = io.BytesIO()
        
        try:
            # pd.ExcelWriter permite escribir múltiples hojas en un solo archivo
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for file in uploaded_files:
                    # Obtener nombre sin extensión
                    file_name = os.path.splitext(file.name)[0]
                    # Excel tiene un límite estricto de 31 caracteres para el nombre de las hojas
                    sheet_name = file_name[:31]
                    
                    # Leer el archivo subido (lee la primera hoja por defecto, que es tu 'Sheet')
                    df = pd.read_excel(file)
                    
                    # Escribir los datos en una nueva hoja del archivo final
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Obtener los datos binarios del buffer
            processed_data = output.getvalue()
            
            st.success("✅ ¡Archivos unificados con éxito!")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Archivo Unificado",
                data=processed_data,
                file_name="Archivos_Unificados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {str(e)}")
