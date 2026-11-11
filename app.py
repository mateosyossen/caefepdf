import streamlit as st
import io
import zipfile
import time
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# 1. Configuración de página
st.set_page_config(
    page_title="IDX to PDF Batch Converter",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Sistema visual Premium (Inspirado en el diseño anterior pero adaptado)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;800&display=swap');

.stApp {
    background: linear-gradient(180deg, #F8FAFC 0%, #EEF2F6 100%);
    font-family: 'DM Sans', sans-serif;
}
.hero {
    position: relative;
    overflow: hidden;
    padding: 30px 40px;
    margin-bottom: 30px;
    border-radius: 20px;
    background: linear-gradient(112deg, #0F172A 0%, #334155 68%, #3B82F6 100%);
    color: white;
    box-shadow: 0 14px 34px rgba(15, 23, 42, .15);
}
.hero h1 {
    font-family: 'Manrope', sans-serif;
    font-weight: 800;
    margin: 0;
    font-size: 32px;
}
.hero p {
    color: #93C5FD;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    font-size: 12px;
    margin-top: 5px;
}
.kpi-box {
    background: white;
    padding: 20px;
    border-radius: 15px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 4px 15px rgba(0,0,0,0.03);
    text-align: center;
}
.kpi-num {
    font-size: 32px;
    font-weight: 800;
    color: #0F172A;
    font-family: 'Manrope', sans-serif;
}
.kpi-label {
    font-size: 11px;
    color: #64748B;
    text-transform: uppercase;
    font-weight: 700;
    letter-spacing: 1px;
}
/* Estilizar el uploader de Streamlit */
[data-testid="stFileUploader"] {
    background: white;
    border: 2px dashed #3B82F6;
    border-radius: 15px;
    padding: 20px;
}
</style>
""", unsafe_allow_html=True)

# 3. Lógica de Conversión IDX a PDF
def convert_idx_to_pdf(file_bytes, filename):
    """Convierte el contenido de un archivo IDX (texto) en un PDF usando ReportLab."""
    # Intentar decodificar el texto
    try:
        text_content = file_bytes.decode('utf-8')
    except UnicodeDecodeError:
        text_content = file_bytes.decode('latin-1')

    # Preparar el PDF en memoria
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    
    styles = getSampleStyleSheet()
    style_normal = styles["Normal"]
    style_title = styles["Title"]
    
    story = []
    # Título del documento
    story.append(Paragraph(f"Exportación de: {filename}", style_title))
    story.append(Spacer(1, 20))
    
    # Procesar líneas (escapando caracteres especiales para ReportLab)
    for line in text_content.split('\n'):
        clean_line = line.strip().replace('<', '&lt;').replace('>', '&gt;')
        if clean_line:
            story.append(Paragraph(clean_line, style_normal))
            story.append(Spacer(1, 5))
            
    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer.read()

# 4. Interfaz de Usuario
st.markdown("""
    <div class="hero">
        <h1>⚡ IDX to PDF Batch Converter</h1>
        <p>Procesamiento masivo de archivos sin límite de peso</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Panel de Control")
    st.info("Sube todos los archivos `.idx` que necesites. El sistema los procesará en cola y generará un único archivo `.zip` con todos los PDFs listos para descargar.")
    st.markdown("---")
    st.markdown("Desarrollado con **Streamlit** & **ReportLab**")

st.markdown("### 📥 1. Cargar Archivos")
uploaded_files = st.file_uploader("Arrastra tus archivos .idx aquí (Puedes seleccionar múltiples archivos)", type=['idx'], accept_multiple_files=True)

if uploaded_files:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{len(uploaded_files)}</div><div class="kpi-label">Archivos Detectados</div></div>', unsafe_allow_html=True)
    with col2:
        total_size = sum([f.size for f in uploaded_files]) / (1024 * 1024)
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{total_size:.2f} MB</div><div class="kpi-label">Peso Total a Procesar</div></div>', unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 Iniciar Conversión Batch", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Buffer en memoria para el ZIP final
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for i, file in enumerate(uploaded_files):
                status_text.text(f"⏳ Procesando: {file.name} ({i+1}/{len(uploaded_files)})...")
                
                # Leer bytes y convertir a PDF
                file_bytes = file.read()
                pdf_data = convert_idx_to_pdf(file_bytes, file.name)
                
                # Crear nombre de salida (cambiar .idx por .pdf)
                pdf_filename = file.name.rsplit('.', 1)[0] + ".pdf"
                
                # Escribir el PDF dentro del archivo ZIP
                zip_file.writestr(pdf_filename, pdf_data)
                
                # Actualizar barra de progreso
                progress_bar.progress((i + 1) / len(uploaded_files))
                
        status_text.success("✅ ¡Conversión completada con éxito!")
        
        # Ofrecer descarga del ZIP
        zip_buffer.seek(0)
        st.markdown("### 📦 2. Descargar Resultados")
        st.download_button(
            label=f"⬇️ Descargar Lote Completo ({len(uploaded_files)} PDFs en .ZIP)",
            data=zip_buffer,
            file_name="Lote_Exportado_PDFs.zip",
            mime="application/zip",
            type="primary"
        )
