import streamlit as st
import io
import zipfile
import pandas as pd
from fpdf import FPDF

# 1. Configuración de página
st.set_page_config(
    page_title="Excel to PDF Batch Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Sistema visual Premium
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
    background: linear-gradient(112deg, #0F172A 0%, #1E293B 68%, #10B981 100%);
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
    color: #A7F3D0;
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
[data-testid="stFileUploader"] {
    background: white;
    border: 2px dashed #10B981;
    border-radius: 15px;
    padding: 20px;
}
</style>
""", unsafe_allow_html=True)

# 3. Lógica de Conversión Excel a PDF
def convert_excel_to_pdf(file_bytes, filename):
    """Lee el Excel con Pandas y dibuja una tabla en PDF usando fpdf2 (Formato Apaisado)."""
    # Leer el Excel en memoria
    df = pd.read_excel(io.BytesIO(file_bytes))
    df = df.fillna("") # Reemplazar valores nulos (NaN) con texto vacío
    
    # Crear PDF en formato Landscape (L) para que entren las columnas
    pdf = FPDF(orientation="L")
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Título
    pdf.set_font("helvetica", style="B", size=14)
    pdf.cell(0, 10, f"Reporte de datos: {filename}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    num_cols = len(df.columns)
    if num_cols == 0:
        pdf.set_font("helvetica", size=10)
        pdf.cell(0, 10, "El archivo Excel se encuentra vacío.", new_x="LMARGIN", new_y="NEXT")
        return bytes(pdf.output())

    # Calcular el ancho de las columnas (ancho disponible / cantidad de columnas)
    col_width = pdf.epw / num_cols
    line_height = 7
    
    # Imprimir Encabezados
    pdf.set_font("helvetica", style="B", size=8)
    for col in df.columns:
        # Limpiamos caracteres y truncamos a 25 caracteres para que no desborde la celda
        texto = str(col).encode('latin-1', 'replace').decode('latin-1')[:25]
        pdf.cell(col_width, line_height, texto, border=1)
    pdf.ln(line_height)
    
    # Imprimir Filas de datos
    pdf.set_font("helvetica", size=8)
    for _, row in df.iterrows():
        for item in row:
            texto = str(item).encode('latin-1', 'replace').decode('latin-1')[:25]
            pdf.cell(col_width, line_height, texto, border=1)
        pdf.ln(line_height)
            
    return bytes(pdf.output())

# 4. Interfaz de Usuario
st.markdown("""
    <div class="hero">
        <h1>📊 Excel to PDF Batch Converter</h1>
        <p>Procesamiento masivo de tablas sin límite de peso</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Panel de Control")
    st.info("Sube todos los archivos **.xlsx / .xls** que necesites. El sistema extraerá las tablas, las orientará de forma apaisada y generará un único archivo `.zip` con todos los PDFs listos.")

st.markdown("### 📥 1. Cargar Archivos Excel")
uploaded_files = st.file_uploader("Arrastra tus archivos aquí (Puedes seleccionar múltiples)", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{len(uploaded_files)}</div><div class="kpi-label">Archivos Detectados</div></div>', unsafe_allow_html=True)
    with col2:
        total_size = sum([f.size for f in uploaded_files]) / (1024 * 1024)
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{total_size:.2f} MB</div><div class="kpi-label">Peso Total a Procesar</div></div>', unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 Iniciar Conversión a PDF", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for i, file in enumerate(uploaded_files):
                status_text.text(f"⏳ Convirtiendo: {file.name} ({i+1}/{len(uploaded_files)})...")
                
                try:
                    file_bytes = file.read()
                    pdf_data = convert_excel_to_pdf(file_bytes, file.name)
                    
                    # Generar el nombre de salida .pdf
                    pdf_filename = file.name.rsplit('.', 1)[0] + ".pdf"
                    zip_file.writestr(pdf_filename, pdf_data)
                except Exception as e:
                    st.error(f"Error procesando {file.name}: {e}")
                
                progress_bar.progress((i + 1) / len(uploaded_files))
                
        status_text.success("✅ ¡Conversión completada con éxito!")
        
        zip_buffer.seek(0)
        st.markdown("### 📦 2. Descargar Resultados")
        st.download_button(
            label=f"⬇️ Descargar Lote Completo ({len(uploaded_files)} PDFs en .ZIP)",
            data=zip_buffer,
            file_name="Lote_Exportado_Tablas_PDF.zip",
            mime="application/zip",
            type="primary"
        )
