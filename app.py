import streamlit as st
import pandas as pd
import io
import os

# Configuración básica de la página
st.set_page_config(page_title="Unificador de Excel", page_icon="📊")

st.title("📊 Unificador de Archivos Excel")
st.write("Sube múltiples archivos Excel (.xls o .xlsx). Se unificarán copiando todo el contenido exactamente a partir de la fila 7 y sin etiquetas 'Unnamed'.")

# Subida de archivos (permite múltiples archivos xls y xlsx)
uploaded_files = st.file_uploader(
    "Sube tus archivos Excel aquí", 
    type=['xls', 'xlsx'], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Has subido {len(uploaded_files)} archivos.")
    
    if st.button("Unificar Archivos", type="primary"):
        output = io.BytesIO()
        
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for file in uploaded_files:
                    # Obtener nombre del archivo sin extensión (máximo 31 caracteres para la pestaña)
                    file_name = os.path.splitext(file.name)[0]
                    sheet_name = file_name[:31]
                    
                    # 1. header=None lee todo el archivo como datos puros, evitando crear etiquetas 'Unnamed'
                    df = pd.read_excel(file, header=None)
                    
                    # 2. startrow=6 coloca los datos a partir de la fila 7 en Excel (el índice de filas inicia en 0)
                    # 3. header=False e index=False evitan agregar nombres de columnas o numeración extra
                    df.to_excel(
                        writer, 
                        sheet_name=sheet_name, 
                        index=False, 
                        header=False, 
                        startrow=6
                    )
            
            processed_data = output.getvalue()
            
            st.success("✅ ¡Archivos unificados con éxito a partir de la fila 7!")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Archivo Unificado",
                data=processed_data,
                file_name="Archivos_Unificados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {str(e)}")
