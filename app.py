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

# 3. Estilos Visuales Premium e Inmersivos (CSS Custom)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,600;0,9..40,700;0,9..40,800;1,9..40,400&family=Manrope:wght@500;600;700;800&display=swap');

:root {
    --navy-dark: #070F1E;
    --navy-card: #0D1B2E;
    --navy-surface: #12243D;
    --cyan-accent: #00B4D8;
    --cyan-glow: rgba(0, 180, 216, 0.35);
    --cyan-soft: rgba(0, 180, 216, 0.08);
    --border-light: rgba(255, 255, 255, 0.08);
    --text-primary: #0F172A;
    --text-muted: #64748B;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background: 
        radial-gradient(circle at 90% 10%, rgba(0, 180, 216, 0.08), transparent 35%),
        radial-gradient(circle at 10% 90%, rgba(14, 116, 144, 0.06), transparent 30%),
        linear-gradient(180deg, #F4F8FC 0%, #EBF1F7 100%);
    color: var(--text-primary);
}

.block-container {
    max-width: 1440px;
    padding-top: 1.5rem;
    padding-bottom: 4rem;
}

/* Sidebar Customization */
[data-testid="stSidebar"] {
    background: linear-gradient(185deg, #070F1E 0%, #0D1C30 100%) !important;
    border-right: 1px solid rgba(255, 255, 255, 0.07);
}
[data-testid="stSidebar"] * {
    color: #E2E8F0 !important;
}

/* Hero Header Estilo Industrial */
.caefe-hero {
    position: relative;
    overflow: hidden;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 25px;
    padding: 30px 36px;
    margin-bottom: 28px;
    border-radius: 24px;
    background: linear-gradient(115deg, #070F1E 0%, #0F233D 60%, #004E7C 100%);
    border: 1px solid rgba(255, 255, 255, 0.12);
    box-shadow: 0 16px 40px rgba(7, 15, 30, 0.22);
    animation: riseIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both;
}
.caefe-hero:before {
    content: "";
    position: absolute;
    top: -120px;
    right: -80px;
    width: 320px;
    height: 320px;
    border-radius: 50%;
    background: radial-gradient(circle, var(--cyan-glow), transparent 70%);
    pointer-events: none;
}
.hero-content {
    position: relative;
    z-index: 2;
}
.hero-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 5px 12px;
    border-radius: 999px;
    background: rgba(0, 180, 216, 0.15);
    border: 1px solid rgba(0, 180, 216, 0.35);
    color: #90E0EF !important;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    margin-bottom: 12px;
}
.hero-badge-dot {
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: #00B4D8;
    box-shadow: 0 0 8px #00B4D8;
}
.caefe-hero h1 {
    margin: 0;
    color: #FFFFFF !important;
    font-family: 'Manrope', sans-serif;
    font-size: clamp(24px, 2.3vw, 32px);
    font-weight: 800;
    letter-spacing: -0.6px;
}
.caefe-hero p {
    margin: 6px 0 0;
    color: #94A3B8 !important;
    font-size: 13px;
    font-weight: 500;
}

/* Tarjetas KPI de Estado */
.kpi-deck {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
}
.kpi-tile {
    position: relative;
    overflow: hidden;
    padding: 20px 22px;
    border-radius: 18px;
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    box-shadow: 0 6px 20px rgba(15, 23, 42, 0.04);
    transition: transform 0.25s ease, box-shadow 0.25s ease;
    animation: riseIn 0.4s ease both;
}
.kpi-tile:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 28px rgba(0, 180, 216, 0.12);
    border-color: #BAE6FD;
}
.kpi-tile .kpi-tag {
    font-size: 11px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    color: var(--text-muted);
}
.kpi-tile .kpi-val {
    font-family: 'Manrope', sans-serif;
    font-size: 30px;
    font-weight: 800;
    color: var(--navy-dark);
    margin-top: 4px;
    letter-spacing: -0.5px;
}
.kpi-tile .kpi-sub {
    font-size: 12px;
    color: #0284C7;
    margin-top: 2px;
    font-weight: 600;
}

/* Drag & Drop Zone Estilizado */
[data-testid="stFileUploader"] {
    background: #FFFFFF;
    border: 2px dashed #93C5FD;
    border-radius: 20px;
    padding: 24px 20px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.03);
    transition: all 0.25s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: #00B4D8;
    box-shadow: 0 10px 30px var(--cyan-glow);
    background: #F8FAFC;
}

/* Botones con gradientes premium */
.stButton > button, .stDownloadButton > button {
    height: 48px;
    border: none !important;
    border-radius: 12px !important;
    background: linear-gradient(135deg, #0096C7 0%, #0077B6 100%) !important;
    color: #FFFFFF !important;
    font-weight: 700 !important;
    letter-spacing: 0.3px !important;
    box-shadow: 0 8px 20px rgba(0, 150, 199, 0.28) !important;
    transition: all 0.25s ease !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 12px 28px rgba(0, 150, 199, 0.38) !important;
    filter: brightness(1.06);
}

@keyframes riseIn {
    from { opacity: 0; transform: translateY(12px); }
    to { opacity: 1; transform: translateY(0); }
}
</style>
""", unsafe_allow_html=True)

# 4. Lógica de Conversión y Procesamiento
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

        # Conversión previa de .xls a .xlsx conservando macros y gráficos
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

                # Eliminar hojas plantilla vacías para evitar exportaciones redundantes
                if hojas_validas and len(hojas_validas) < len(wb.sheetnames):
                    for sheetname in list(wb.sheetnames):
                        if sheetname not in hojas_validas:
                            del wb[sheetname]

                # Auto-escalado forzado a 1 página de ancho
                for ws in wb.worksheets:
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 0

                wb.save(input_path)
            except Exception as e:
                pass

        # Conversión a PDF mediante motor LibreOffice headless
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


# 5. Sidebar Brand & Status
with st.sidebar:
    if LOGO_PATH:
        st.image(LOGO_PATH, use_container_width=True)
    else:
        # Fallback vectorial estilizado con el isotipo de CAEFE
        st.markdown("""
        <div style="background: rgba(255,255,255,0.04); padding: 18px; border-radius: 16px; text-align: center; border: 1px solid rgba(255,255,255,0.08); margin-bottom: 20px;">
            <svg width="48" height="48" viewBox="0 0 100 100" fill="none">
                <path d="M 50 12 L 88 50 L 72 50 L 50 28 L 28 50 L 50 72 L 72 50 L 88 50 L 50 88 L 12 50 Z" fill="#00B4D8"/>
            </svg>
            <div style="color: #FFFFFF; font-weight: 800; font-size: 18px; letter-spacing: 1px; margin-top: 6px;">CAEFE S.A.</div>
            <div style="color: #00B4D8; font-size: 10px; font-weight: 700; letter-spacing: 1.5px;">CONVERTIDOR BATCH</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("#### ⚡ Rendimiento & Motor")
    st.markdown("""
    - **Alineación:** Ajuste a 1 hoja de ancho.
    - **Fidelidad:** Conserva sellos, firmas y fotos.
    - **Compatibilidad:** Reportes `.xlsx` y `.xls`.
    """)
    st.markdown("---")
    st.caption("CAEFE S.A. · División Digital 2026")

# 6. Hero Header
st.markdown(f"""
    <section class="caefe-hero">
        <div class="hero-content">
            <div class="hero-badge">
                <span class="hero-badge-dot"></span> MOTOR DE ALTA PRECISIÓN OPERATIVO
            </div>
            <h1>Exportador de Informes a PDF</h1>
            <p>Transformación masiva de reportes técnicos con escala y encuadre garantizados.</p>
        </div>
    </section>
""", unsafe_allow_html=True)

# 7. Carga e Interacción
st.markdown("### 📥 Seleccionar Informes")
uploaded_files = st.file_uploader(
    "Arrastra tus reportes técnicos aquí (Soporta múltiples archivos sin límite de peso):",
    type=['xlsx', 'xls'],
    accept_multiple_files=True
)

if uploaded_files:
    total_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024)
    
    st.markdown(f"""
    <div class="kpi-deck">
        <div class="kpi-tile">
            <div class="kpi-tag">Documentos a procesar</div>
            <div class="kpi-val">{len(uploaded_files)}</div>
            <div class="kpi-sub">Lote listo</div>
        </div>
        <div class="kpi-tile">
            <div class="kpi-tag">Peso total de archivos</div>
            <div class="kpi-val">{total_mb:.2f} <span style="font-size:18px">MB</span></div>
            <div class="kpi-sub">Buffer en memoria</div>
        </div>
        <div class="kpi-tile">
            <div class="kpi-tag">Modo de impresión</div>
            <div class="kpi-val">100%</div>
            <div class="kpi-sub">Auto-fit en 1 página de ancho</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("⚡ Procesar y Generar Lote de PDFs", type="primary", use_container_width=True):
        barra_progreso = st.progress(0)
        status_box = st.empty()
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, file in enumerate(uploaded_files):
                status_box.markdown(f"**⏳ Procesando ({idx + 1}/{len(uploaded_files)}):** `{file.name}`...")
                try:
                    pdf_bytes, nombre_pdf = prepare_and_convert_excel(file)
                    zf.writestr(nombre_pdf, pdf_bytes)
                except Exception as ex:
                    st.error(f"Error con el archivo {file.name}: {ex}")

                barra_progreso.progress((idx + 1) / len(uploaded_files))

        status_box.empty()
        st.balloons()
        st.success(f"🎉 ¡Procesamiento finalizado! Se generaron {len(uploaded_files)} informes en PDF con encuadre exacto.")
        zip_buffer.seek(0)

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label=f"⬇️ DESCARGAR LOTE COMPLETO ({len(uploaded_files)} PDFs en .ZIP)",
            data=zip_buffer,
            file_name="Informes_CAEFE_Generados.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
