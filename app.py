import streamlit as st
import pandas as pd
import io
import os

# Configuración básica de la página
st.set_page_config(page_title="Unificador de Excel", page_icon="📊")

st.title("📊 Unificador de Archivos Excel")
st.write("Sube múltiples archivos Excel (.xls o .xlsx). Se unificarán eliminando los encabezados 'Unnamed' y comenzando desde la fila 6.")

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
                    
                    # 1. header=None evita que Pandas genere las etiquetas 'Unnamed: X'
                    df = pd.read_excel(file, header=None)
                    
                    # 2. Corta las primeras 5 filas para empezar desde la fila 6 del archivo original
                    df = df.iloc[5:].reset_index(drop=True)
                    
                    # 3. Escribir los datos en la hoja sin incluir índices ni fila de encabezados extra
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            
            # Obtener los datos binarios del buffer
            processed_data = output.getvalue()
            
            st.success("✅ ¡Archivos unificados sin 'Unnamed' con éxito!")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Archivo Unificado",
                data=processed_data,
                file_name="Archivos_Unificados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {str(e)}")
