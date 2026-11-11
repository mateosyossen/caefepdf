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
    page_title="CAEFE S.A. - Suite Batch PDF Converter",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Detección de Logo Local
def obtener_ruta_logo():
    candidatos = [
        'pagina.png (1).jpg', 'pagina.png', 'logo.png', 'logo_caefe.png', 
        'caefe.png', 'logo_caefe_crop.png', '8.png'
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    return None

LOGO_PATH = obtener_ruta_logo()

# 3. CSS Pulido y Equilibrado
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,600;0,9..40,700;0,9..40,800&family=Manrope:wght@600;700;800&display=swap');

:root {
    --navy-bg: #091322;
    --navy-header: #0F1E36;
    --cyan: #00B4D8;
    --cyan-hover: #0096C7;
    --border-subtle: #E2E8F0;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background-color: #F8FAFC;
    color: #0F172A;
}

/* Espaciado del contenedor principal */
.block-container {
    max-width: 1350px !important;
    padding-top: 2rem !important;
    padding-bottom: 4rem !important;
    padding-left: 2.5rem !important;
    padding-right: 2.5rem !important;
}

/* Sidebar Limpio */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #091322 0%, #0F1E36 100%) !important;
    border-right: 1px solid rgba(255, 255, 255, 0.08);
}
[data-testid="stSidebar"] * {
    color: #CBD5E1 !important;
}
[data-testid="stSidebar"] code {
    background: rgba(255, 255, 255, 0.1) !important;
    color: #38BDF8 !important;
    border: none !important;
    padding: 2px 6px !important;
    border-radius: 4px !important;
}

/* Caja de Marca en Sidebar */
.brand-box {
    background: rgba(255, 255, 255, 0.04);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 14px;
    padding: 16px;
    text-align: center;
    margin-bottom: 24px;
}
.brand-title {
    color: #FFFFFF !important;
    font-weight: 800;
    font-size: 17px;
    letter-spacing: 0.8px;
    margin-top: 8px;
}
.brand-sub {
    color: #38BDF8 !important;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin-top: 2px;
}

/* Hero Header Integrado */
.hero-card {
    background: linear-gradient(135deg, #0B192E 0%, #102A4C 60%, #004B75 100%);
    border-radius: 18px;
    padding: 32px 36px;
    margin-bottom: 32px;
    color: #FFFFFF;
    box-shadow: 0 10px 30px rgba(11, 25, 46, 0.15);
    border: 1px solid rgba(255, 255, 255, 0.1);
}
.hero-tag {
    display: inline-block;
    background: rgba(0, 180, 216, 0.15);
    border: 1px solid rgba(0, 180, 216, 0.35);
    color: #7DD3FC !important;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    padding: 5px 12px;
    border-radius: 999px;
    margin-bottom: 14px;
}
.hero-card h1 {
    font-family: 'Manrope', sans-serif;
    font-size: 28px;
    font-weight: 800;
    color: #FFFFFF !important;
    margin: 0 0 8px 0;
    letter-spacing: -0.5px;
}
.hero-card p {
    color: #94A3B8 !important;
    font-size: 14px;
    margin: 0;
    font-weight: 400;
}

/* Tarjetas KPI de Estado */
.kpi-row {
    display: flex;
    gap: 16px;
    margin-bottom: 24px;
}
.kpi-card {
    flex: 1;
    background: #FFFFFF;
    border: 1px solid var(--border-subtle);
    border-radius: 14px;
    padding: 18px 22px;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03);
}
.kpi-card .kpi-label {
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.1px;
    color: #64748B;
}
.kpi-card .kpi-val {
    font-family: 'Manrope', sans-serif;
    font-size: 26px;
    font-weight: 800;
    color: #0F172A;
    margin-top: 4px;
}

/* Subida de archivos estilizada */
[data-testid="stFileUploader"] {
    background: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 14px;
    padding: 20px;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.02);
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--cyan);
}

/* Botones con estilo enterprise */
.stButton > button, .stDownloadButton > button {
    height: 46px;
    border: none !important;
    border-radius: 10px !important;
    background: linear-gradient(135deg, #0096C7 0%, #0077B6 100%) !important;
    color: #FFFFFF !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    box-shadow: 0 4px 14px rgba(0, 119, 182, 0.25) !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(0, 119, 182, 0.35) !important;
}
</style>
""", unsafe_allow_html=True)

# 4. Funciones de Procesamiento y Conversión
def extraer_codigo_informe(texto):
    if not texto:
        return None
    match = re.search(r'R-[A-Z]+-\d+(?:-\d+)?', str(texto), re.IGNORECASE)
    if match:
        return match.group(0).upper()
    return None

def prepare_and_convert_excel(uploaded_file):
    with tempfile.TemporaryDirectory() as tmpdir:
        original_path = os.path.join(tmpdir, uploaded_file.name)
        with open(original_path, "wb") as f:
            f.write(uploaded_file.read())

        input_path = original_path

        # Conversión previa de .xls a .xlsx conservando formato
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

        if input_path.lower().endswith('.xlsx') and os.path.exists(input_path):
            try:
                wb = openpyxl.load_workbook(input_path)
                hojas_validas = []

                for sheetname in wb.sheetnames:
                    ws = wb[sheetname]
                    tiene_datos_reales = False

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

                # Eliminar hojas plantilla vacías
                if hojas_validas and len(hojas_validas) < len(wb.sheetnames):
                    for sheetname in list(wb.sheetnames):
                        if sheetname not in hojas_validas:
                            del wb[sheetname]

                # Ajustar ancho a 1 página
                for ws in wb.worksheets:
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 0

                wb.save(input_path)
            except Exception:
                pass

        # Conversión a PDF mediante LibreOffice
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
                raise Exception("No se generó el PDF correspondiente.")

        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        nombre_final = f"{numero_informe_detectado}.pdf" if numero_informe_detectado else f"{base_name}.pdf"
        return pdf_bytes, nombre_final


# 5. Barra Lateral
with st.sidebar:
    if LOGO_PATH:
        st.markdown('<div class="brand-box">', unsafe_allow_html=True)
        st.image(LOGO_PATH, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="brand-box">
            <div class="brand-title">CAEFE S.A.</div>
            <div class="brand-sub">CONVERTIDOR BATCH</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("#### Configuración de Exportación")
    st.markdown(
        "- **Escala:** Ajuste a 1 hoja de ancho.\n"
        "- **Fidelidad:** Conserva sellos, firmas y fotos.\n"
        "- **Formatos:** Compatible con `.xlsx` y `.xls`."
    )
    st.markdown("---")
    st.caption("CAEFE S.A. · 2026")

# 6. Encabezado Principal
st.markdown("""
    <section class="hero-card">
        <span class="hero-tag">SISTEMA ACTIVO</span>
        <h1>Exportador de Informes a PDF</h1>
        <p>Conversión por lote con detección automática de número de reporte y ajuste de página.</p>
    </section>
""", unsafe_allow_html=True)

# 7. Área de Carga e Interacción
st.subheader("Cargar Archivos")
uploaded_files = st.file_uploader(
    "Selecciona o arrastra los archivos de reporte:",
    type=['xlsx', 'xls'],
    accept_multiple_files=True
)

if uploaded_files:
    total_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024)
    
    st.markdown(f"""
    <div class="kpi-row">
        <div class="kpi-card">
            <div class="kpi-label">Informes seleccionados</div>
            <div class="kpi-val">{len(uploaded_files)}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Tamaño acumulado</div>
            <div class="kpi-val">{total_mb:.2f} MB</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Formato de salida</div>
            <div class="kpi-val">PDF (A4)</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Procesar y Descargar Lote", type="primary", use_container_width=True):
        barra_progreso = st.progress(0)
        status_box = st.empty()
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, file in enumerate(uploaded_files):
                status_box.markdown(f"**Procesando ({idx + 1}/{len(uploaded_files)}):** `{file.name}`...")
                try:
                    pdf_bytes, nombre_pdf = prepare_and_convert_excel(file)
                    zf.writestr(nombre_pdf, pdf_bytes)
                except Exception as ex:
                    st.error(f"Error en {file.name}: {ex}")

                barra_progreso.progress((idx + 1) / len(uploaded_files))

        status_box.empty()
        st.success(f"¡Completado! Se generaron {len(uploaded_files)} archivos PDF listos para descarga.")
        zip_buffer.seek(0)

        st.download_button(
            label=f"Descargar archivo ZIP ({len(uploaded_files)} PDFs)",
            data=zip_buffer,
            file_name="Informes_CAEFE_Generados.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
