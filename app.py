import streamlit as st
import io
import os
from copy import copy
import openpyxl

st.set_page_config(page_title="Unificador de Excel con Formato", page_icon="🎨")

st.title("🎨 Unificador de Archivos Excel")
st.write("Sube múltiples archivos **.xlsx**. Se generará un solo archivo conservando los colores y formatos originales de las celdas.")

# Restringimos a xlsx para garantizar que se conserve el formato
uploaded_files = st.file_uploader(
    "Sube tus archivos Excel (.xlsx) aquí", 
    type=['xlsx'], 
    accept_multiple_files=True
)

def copiar_estilo_y_datos(origen_sheet, destino_sheet):
    """Función para copiar datos y formatos de una hoja a otra"""
    
    # 1. Copiar dimensiones de las columnas (ancho)
    for col, dim in origen_sheet.column_dimensions.items():
        destino_sheet.column_dimensions[col].width = dim.width

    # 2. Copiar dimensiones de las filas (alto)
    for row, dim in origen_sheet.row_dimensions.items():
        destino_sheet.row_dimensions[row].height = dim.height

    # 3. Copiar valores y estilos de cada celda
    for row in origen_sheet.rows:
        for cell in row:
            new_cell = destino_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            
            # Si la celda original tiene estilo, lo copiamos
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)     # Esto copia el color de fondo
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    # 4. Copiar celdas combinadas (Merged cells)
    if origen_sheet.merged_cells:
        for merged_cell_range in origen_sheet.merged_cells.ranges:
            destino_sheet.merge_cells(str(merged_cell_range))

if uploaded_files:
    st.info(f"Has subido {len(uploaded_files)} archivos.")
    
    if st.button("Unificar Archivos", type="primary"):
        output = io.BytesIO()
        
        try:
            # Crear el libro de trabajo (Workbook) final
            wb_final = openpyxl.Workbook()
            # Eliminar la hoja por defecto que se crea vacía
            hoja_por_defecto = wb_final.active
            wb_final.remove(hoja_por_defecto)
            
            for file in uploaded_files:
                # Leer el archivo original con sus formatos
                wb_origen = openpyxl.load_workbook(file, data_only=False)
                
                # Seleccionar la primera hoja (o la que se llame 'Sheet')
                sheet_origen = wb_origen.active 
                
                # Obtener nombre sin extensión y limitarlo a 31 caracteres
                file_name = os.path.splitext(file.name)[0][:31]
                
                # Crear una nueva hoja en el archivo final
                sheet_destino = wb_final.create_sheet(title=file_name)
                
                # Ejecutar la función de copiado celda por celda
                copiar_estilo_y_datos(sheet_origen, sheet_destino)
            
            # Guardar el resultado en el buffer de memoria
            wb_final.save(output)
            processed_data = output.getvalue()
            
            st.success("✅ ¡Archivos unificados respetando colores y formatos con éxito!")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Archivo Final",
                data=processed_data,
                file_name="Archivos_Unificados_Formato.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {str(e)}")
