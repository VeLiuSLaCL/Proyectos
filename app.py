import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Unificador de Excel", page_icon="📊")

st.title("📊 Unificador de Archivos Excel")
st.write("Sube tus archivos Excel (.xls o .xlsx). Se unificarán en un solo archivo con un diseño visual profesional automático.")

# Acepta tanto archivos antiguos (.xls) como nuevos (.xlsx)
uploaded_files = st.file_uploader(
    "Sube tus archivos Excel aquí", 
    type=['xls', 'xlsx'], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"Has subido {len(uploaded_files)} archivos.")
    
    if st.button("Unificar y Aplicar Formato", type="primary"):
        output = io.BytesIO()
        
        try:
            # Usamos xlsxwriter como motor para dar formato visual personalizado
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Definición de formatos visuales
                # 1. Encabezado con fondo azul oscuro y texto blanco en negrita
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'center',
                    'fg_color': '#1F4E78',
                    'font_color': '#FFFFFF',
                    'border': 1
                })
                
                # 2. Formato de celdas normales con borde delgado gris
                cell_format = workbook.add_format({
                    'border': 1,
                    'border_color': '#D9D9D9',
                    'valign': 'vcenter'
                })

                for file in uploaded_files:
                    # Obtener nombre sin extensión (máx 31 caracteres)
                    file_name = os.path.splitext(file.name)[0][:31]
                    
                    # Leer datos del archivo original
                    df = pd.read_excel(file)
                    
                    # Escribir los datos en la nueva hoja
                    df.to_excel(writer, sheet_name=file_name, index=False)
                    
                    # Obtener la hoja actual de xlsxwriter para darle formato
                    worksheet = writer.sheets[file_name]
                    
                    # Aplicar formato a los encabezados
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        
                        # Ajuste automático del ancho de la columna según el contenido
                        max_len = max(
                            df[value].astype(str).map(len).max() if not df.empty else 0,
                            len(str(value))
                        ) + 4
                        worksheet.set_column(col_num, col_num, max(max_len, 12), cell_format)

            processed_data = output.getvalue()
            
            st.success("✅ ¡Archivos unificados y formateados con éxito!")
            
            # Botón para descargar el resultado
            st.download_button(
                label="📥 Descargar Archivo Unificado",
                data=processed_data,
                file_name="Archivos_Unificados_Formateados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {str(e)}")
