import streamlit as st
import io
import zipfile
from fpdf import FPDF

# (Mantén aquí la configuración de página y el bloque st.markdown con el CSS que ya tenías)

def convert_idx_to_pdf(file_bytes, filename):
    """Convierte el contenido de un archivo IDX en un PDF usando fpdf2."""
    # Decodificar texto
    try:
        text_content = file_bytes.decode('utf-8')
    except UnicodeDecodeError:
        text_content = file_bytes.decode('latin-1')

    # Crear PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Título
    pdf.set_font("helvetica", style="B", size=14)
    pdf.cell(0, 10, f"Exportación de: {filename}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # Cuerpo del texto
    pdf.set_font("helvetica", size=10)
    for line in text_content.split('\n'):
        clean_line = line.strip()
        if clean_line:
            # Forzamos compatibilidad de caracteres para evitar errores de codificación en PDF
            clean_line = clean_line.encode('latin-1', 'replace').decode('latin-1')
            pdf.multi_cell(0, 5, clean_line)
            
    # Retornar los bytes del PDF generado
    return bytes(pdf.output())

# (Mantén aquí el resto de tu código de interfaz a partir de "st.markdown("### 📥 1. Cargar Archivos")")
