import streamlit as st
import io
import zipfile
import os
import re
import subprocess
import tempfile
import openpyxl

# 1. Configuración de página
st.set_page_config(
    page_title="CAEFE - Convertidor Batch de Reportes a PDF",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Estilos visuales de la aplicación
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Manrope:wght@600;700;800&display=swap');

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
    background: linear-gradient(112deg, #0B172A 0%, #123451 68%, #075A78 100%); 
    color: white; 
    box-shadow: 0 14px 34px rgba(11, 23, 42, .15); 
}
.hero h1 { 
    font-family: 'Manrope', sans-serif; 
    font-weight: 800; 
    margin: 0; 
    font-size: 30px; 
    letter-spacing: -0.5px;
}
.hero p { 
    color: #9EDCF0; 
    font-weight: 700; 
    letter-spacing: 1.5px; 
    text-transform: uppercase; 
    font-size: 11px; 
    margin-top: 6px; 
}
.kpi-box { 
    background: white; 
    padding: 18px 20px; 
    border-radius: 15px; 
    border: 1px solid #E2E8F0; 
    box-shadow: 0 4px 15px rgba(0,0,0,0.03); 
    text-align: center; 
}
.kpi-num { 
    font-size: 28px; 
    font-weight: 800; 
    color: #0B172A; 
    font-family: 'Manrope', sans-serif; 
}
.kpi-label { 
    font-size: 11px; 
    color: #64748B; 
    text-transform: uppercase; 
    font-weight: 700; 
    letter-spacing: 1px; 
    margin-top: 4px;
}
[data-testid="stFileUploader"] { 
    background: white; 
    border: 2px dashed #08B7D8; 
    border-radius: 15px; 
    padding: 20px; 
}
.stButton > button, .stDownloadButton > button {
    border-radius: 11px;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

# 3. Lógica de procesamiento y conversión
def extraer_codigo_informe(texto):
    """Detecta y extrae códigos de informe tipo R-LP-XXXX-XXX."""
    if not texto:
        return None
    match = re.search(r'R-[A-Z]+-\d+(?:-\d+)?', str(texto), re.IGNORECASE)
    if match:
        return match.group(0).upper()
    return None

def prepare_and_convert_excel(uploaded_file):
    """
    Procesa un archivo Excel (.xls o .xlsx):
    - Convierte .xls a .xlsx si es necesario.
    - Localiza la hoja con datos reales mediante el código de informe.
    - Descarta hojas que sean plantillas sin llenar.
    - Configura el ajuste de página a 1 página de ancho.
    - Exporta a PDF con LibreOffice manteniendo imágenes y diseño.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        original_path = os.path.join(tmpdir, uploaded_file.name)
        with open(original_path, "wb") as f:
            f.write(uploaded_file.read())

        input_path = original_path

        # Si el archivo es .xls antiguo, se convierte primero a .xlsx
        if original_path.lower().endswith('.xls'):
            cmd_xlsx = [
                "libreoffice", "--headless", "--nologo", "--nofirststartwizard",
                "--convert-to", "xlsx", original_path, "--outdir", tmpdir
            ]
            subprocess.run(cmd_xlsx, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            posible_xlsx = os.path.splitext(original_path)[0] + ".xlsx"
            if os.path.exists(posible_xlsx):
                input_path = posible_xlsx

        numero_informe_detectado = None

        # Procesar con openpyxl si el archivo resultante es .xlsx
        if input_path.lower().endswith('.xlsx') and os.path.exists(input_path):
            try:
                wb = openpyxl.load_workbook(input_path)
                hojas_validas = []

                for sheetname in wb.sheetnames:
                    ws = wb[sheetname]
                    tiene_datos_reales = False

                    # Escanear el encabezado (primeras 20 filas y 10 columnas)
                    for r in range(1, min(ws.max_row + 1, 21)):
                        for c in range(1, min(ws.max_column + 1, 11)):
                            val = ws.cell(row=r, column=c).value
                            if val:
                                val_str = str(val).strip()
                                codigo = extraer_codigo_informe(val_str)
                                if codigo:
                                    tiene_datos_reales = True
                                    if not numero_informe_detectado:
                                        numero_informe_detectado = codigo
                                    break
                                # Detección si dice "INFORME N°" y contiene un número real
                                if "INFORME" in val_str.upper() and any(ch.isdigit() for ch in val_str):
                                    if "(NUMERO INTERNO)" not in val_str.upper():
                                        tiene_datos_reales = True
                                        limpio = val_str.replace("INFORME", "").replace("N°", "").replace(":", "").strip()
                                        if limpio and not numero_informe_detectado:
                                            numero_informe_detectado = limpio.split()[0]
                                        break
                        if tiene_datos_reales:
                            break

                    if tiene_datos_reales:
                        hojas_validas.append(sheetname)

                # Si detectamos hojas con datos reales y otras que son plantillas vacías,
                # removemos las vacías para que LibreOffice solo exporte el reporte lleno.
                if hojas_validas and len(hojas_validas) < len(wb.sheetnames):
                    for sheetname in list(wb.sheetnames):
                        if sheetname not in hojas_validas:
                            del wb[sheetname]

                # Ajustar escala a 1 página de ancho para todas las hojas restantes
                for ws in wb.worksheets:
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 0

                wb.save(input_path)
            except Exception as e:
                st.warning(f"Aviso al pre-procesar hojas de {uploaded_file.name}: {e}")

        # Conversión final a PDF mediante LibreOffice
        cmd_pdf = [
            "libreoffice", "--headless", "--nologo", "--nofirststartwizard",
            "--convert-to", "pdf", input_path, "--outdir", tmpdir
        ]
        proc = subprocess.run(cmd_pdf, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if proc.returncode != 0:
            raise Exception(f"Fallo en motor PDF: {proc.stderr.decode(errors='ignore')}")

        base_name = os.path.splitext(uploaded_file.name)[0]
        pdf_path = os.path.join(tmpdir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")

        if not os.path.exists(pdf_path):
            pdfs = [os.path.join(tmpdir, f) for f in os.listdir(tmpdir) if f.lower().endswith('.pdf')]
            if pdfs:
                pdf_path = pdfs[0]
            else:
                raise Exception("No se generó el archivo PDF correspondiente.")

        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        nombre_final = f"{numero_informe_detectado}.pdf" if numero_informe_detectado else f"{base_name}.pdf"
        return pdf_bytes, nombre_final

# 4. Interfaz de usuario
st.markdown("""
    <div class="hero">
        <h1>CAEFE S.A. · Exportador Batch de Informes a PDF</h1>
        <p>Conversión masiva de reportes técnicos con escala y formato original</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Instrucciones de uso")
    st.markdown(
        "- Sube uno o varios archivos de reporte (**`.xlsx`** o **`.xls`**).\n"
        "- El sistema identificará el número de informe, ajustará el ancho a 1 página y descartará pestañas de plantillas vacías.\n"
        "- Al finalizar la conversión, se descargará un único archivo **`.zip`** con todos los PDFs listos."
    )

st.markdown("### 📥 1. Cargar Archivos de Inspección")
uploaded_files = st.file_uploader(
    "Selecciona o arrastra los reportes de inspección:",
    type=['xlsx', 'xls'],
    accept_multiple_files=True
)

if uploaded_files:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{len(uploaded_files)}</div><div class="kpi-label">Reportes Cargados</div></div>', unsafe_allow_html=True)
    with c2:
        peso_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024)
        st.markdown(f'<div class="kpi-box"><div class="kpi-num">{peso_mb:.2f} MB</div><div class="kpi-label">Volumen Total</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 Iniciar Conversión por Lote", type="primary", use_container_width=True):
        barra_progreso = st.progress(0)
        estado = st.empty()
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, file in enumerate(uploaded_files):
                estado.text(f"⏳ Procesando: {file.name} ({idx + 1}/{len(uploaded_files)})...")
                try:
                    pdf_bytes, nombre_pdf = prepare_and_convert_excel(file)
                    zf.writestr(nombre_pdf, pdf_bytes)
                except Exception as ex:
                    st.error(f"Error procesando {file.name}: {ex}")

                barra_progreso.progress((idx + 1) / len(uploaded_files))

        estado.success("✅ ¡Lote completado con éxito!")
        zip_buffer.seek(0)

        st.markdown("### 📦 2. Descarga del Lote")
        st.download_button(
            label=f"📥 Descargar todos los PDFs ({len(uploaded_files)} archivos en .ZIP)",
            data=zip_buffer,
            file_name="Informes_CAEFE_PDF.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
