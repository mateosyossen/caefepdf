import streamlit as st
import io
import zipfile
import os
import subprocess
import tempfile
import openpyxl
from openpyxl.utils import get_column_letter

# 1. Configuración de página
st.set_page_config(
    page_title="Excel to PDF Master Converter",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Sistema visual Premium
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;800&display=swap');
.stApp { background: linear-gradient(180deg, #F8FAFC 0%, #EEF2F6 100%); font-family: 'DM Sans', sans-serif; }
.hero { position: relative; overflow: hidden; padding: 30px 40px; margin-bottom: 30px; border-radius: 20px; background: linear-gradient(112deg, #0F172A 0%, #1E293B 68%, #10B981 100%); color: white; box-shadow: 0 14px 34px rgba(15, 23, 42, .15); }
.hero h1 { font-family: 'Manrope', sans-serif; font-weight: 800; margin: 0; font-size: 32px; }
.hero p { color: #A7F3D0; font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; font-size: 12px; margin-top: 5px; }
.kpi-box { background: white; padding: 20px; border-radius: 15px; border: 1px solid #E2E8F0; box-shadow: 0 4px 15px rgba(0,0,0,0.03); text-align: center; }
.kpi-num { font-size: 32px; font-weight: 800; color: #0F172A; font-family: 'Manrope', sans-serif; }
.kpi-label { font-size: 11px; color: #64748B; text-transform: uppercase; font-weight: 700; letter-spacing: 1px; }
[data-testid="stFileUploader"] { background: white; border: 2px dashed #10B981; border-radius: 15px; padding: 20px; }
</style>
""", unsafe_allow_html=True)

# 3. Lógica de Conversión Perfecta y Automática
def prepare_and_convert_excel(uploaded_file):
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, uploaded_file.name)
        
        # 1. Guardar el archivo original temporalmente
        with open(input_path, "wb") as f:
            f.write(uploaded_file.read())
            
        # 2. Inyectar configuración de impresión automáticamente (Solo para .xlsx)
        if input_path.endswith('.xlsx'):
            try:
                wb = openpyxl.load_workbook(input_path)
                ws = wb.active # Toma la hoja principal
                
                # Detectar límites exactos del formulario
                col_letter = get_column_letter(ws.max_column)
                max_row = ws.max_row
                
                # Establecer Área de Impresión dinámica (Ej: A1:M50)
                ws.print_area = f'A1:{col_letter}{max_row}'
                
                # Forzar escala: Ajustar a 1 página de ancho
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0 # 0 permite que crezca hacia abajo si es muy largo
                
                wb.save(input_path)
            except Exception as e:
                print(f"No se pudo auto-configurar {uploaded_file.name}: {e}")

        # 3. Ejecutar LibreOffice Headless para exportar a PDF
        comando = [
            "libreoffice", "--headless", "--nologo", "--nofirststartwizard", 
            "--convert-to", "pdf", input_path, "--outdir", tmpdir
        ]
        
        proceso = subprocess.run(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        if proceso.returncode != 0:
            raise Exception(f"Fallo en motor PDF: {proceso.stderr.decode()}")
            
        # 4. Leer el PDF resultante
        pdf_filename = uploaded_file.name.rsplit('.', 1)[0] + ".pdf"
        pdf_path = os.path.join(tmpdir, pdf_filename)
        
        with open(pdf_path, "rb") as f:
            return f.read()

# 4. Interfaz de Usuario
st.markdown("""
    <div class="hero">
        <h1>📑 Excel to PDF Master Converter</h1>
        <p>Auto-escalado y conversión de alta fidelidad para Reportes</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Panel de Control")
    st.info("Sube reportes **.xlsx**. El sistema auto-detectará el área con datos, lo ajustará a 1 página de ancho y generará el PDF exacto.")

st.markdown("### 📥 1. Cargar Archivos Excel")
uploaded_files = st.file_uploader("Arrastra tus archivos aquí (.xlsx recomendado)", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{len(uploaded_files)}</div><div class="kpi-label">Archivos Detectados</div></div>', unsafe_allow_html=True)
    with col2:
        total_size = sum([f.size for f in uploaded_files]) / (1024 * 1024)
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{total_size:.2f} MB</div><div class="kpi-label">Peso Total a Procesar</div></div>', unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 Iniciar Conversión Automatizada", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for i, file in enumerate(uploaded_files):
                status_text.text(f"⏳ Auto-ajustando y convirtiendo: {file.name} ({i+1}/{len(uploaded_files)})...")
                
                try:
                    pdf_data = prepare_and_convert_excel(file)
                    
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
            file_name="Reportes_AutoAjustados_PDF.zip",
            mime="application/zip",
            type="primary"
        )
