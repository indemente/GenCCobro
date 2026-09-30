# -*- coding: utf-8 -*-
"""
Generador de Cuenta de Cobro — Streamlit
Formato institucional para profesionales, técnicos y tecnólogos.
Fuente: Arial (equivalente métrico a Helvetica).
"""

import os
import re
import io
import shutil
import subprocess
import tempfile

import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


# =====================================================================
#  CONFIGURACIÓN
# =====================================================================
st.set_page_config(
    page_title="Generador de Cuenta de Cobro",
    page_icon="📄",
    layout="wide",
)


# =====================================================================
#  NÚMERO → LETRAS
# =====================================================================
_UNI = ['','UNO','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE',
        'DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISEIS',
        'DIECISIETE','DIECIOCHO','DIECINUEVE','VEINTE']
_DEC = ['','','VEINTI','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA',
        'OCHENTA','NOVENTA']
_CEN = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
        'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS']

def _cen(n):
    if n == 0: return ''
    if n == 100: return 'CIEN'
    p = []
    c, r = divmod(n, 100)
    if c: p.append(_CEN[c])
    if r:
        if r <= 20: p.append(_UNI[r])
        else:
            d, u = divmod(r, 10)
            if d == 2: p.append('VEINTI' + _UNI[u] if u else 'VEINTE')
            else: p.append(_DEC[d] + (' Y ' + _UNI[u] if u else ''))
    return ' '.join(p)

def num_letras(n):
    n = int(n)
    if n == 0: return 'CERO'
    m, r = divmod(n, 1_000_000)
    mi, u = divmod(r, 1000)
    p = []
    if m: p.append('UN MILLON' if m == 1 else _cen(m) + ' MILLONES')
    if mi: p.append('MIL' if mi == 1 else _cen(mi) + ' MIL')
    if u: p.append(_cen(u))
    return ' '.join(p)

def horas_letras(s):
    s = str(s).strip().replace('.', ',')
    if ',' in s:
        e, d = s.split(',', 1)
        et = num_letras(e).capitalize()
        dt = ' '.join(num_letras(x).lower() for x in d if x.isdigit())
        return f"{et} punto {dt}"
    return num_letras(s).capitalize()

def unir_asig(lst):
    c = [f'"{a}"' for a in lst]
    if len(c) == 1: return c[0]
    if len(c) == 2: return f"{c[0]} y {c[1]}"
    return ", ".join(c[:-1]) + f" y {c[-1]}"


# =====================================================================
#  DECLARACIONES (texto genérico de ejemplo)
# =====================================================================
DECLARACIONES = [
    ("Certifico que no tengo personas contratadas para el cumplimiento de "
     "mi actividad. Lo anterior para que se de aplicación a la tabla 383 "
     "del E. T. en lo referente a la retención en la fuente a título de renta."),
    ("Estos ingresos no están gravados con impuesto de ICA por lo cual no "
     "se les practica retención de conformidad con el artículo IO del "
     "acuerdo Municipal 0434 de 3017 de Cali."),
    ("Declaro que no soy responsable del IVA, por lo tanto, no estoy obligado "
     "a expedir factura de ventas conforme al artículo 1.6.1.4.2 del decreto "
     "único reglamentario 1625 del 2016. (Artículo 42 del Decreto 3541 de "
     "1983) y artículo 511 del Estatuto Tributario."),
]


# =====================================================================
#  DATOS GENÉRICOS POR DEFECTO (solo para previsualización)
# =====================================================================
EMPRESA_DEFAULT      = "EMPRESA DE EJEMPLO S.A.S"
NIT_EMPRESA_DEFAULT  = "900.000.000 - 0"
NOMBRE_DEFAULT       = "JUAN PÉREZ GARCÍA"
CEDULA_DEFAULT       = "1.234.567.890"
CIUDAD_DEFAULT       = "BOGOTÁ"
VALOR_DEFAULT        = "1000000"
HORAS_DEFAULT        = "40"
ASIGNATURAS_DEFAULT  = [
    "Matemáticas básicas",
    "Física I",
]


# =====================================================================
#  GENERACIÓN DEL DOCX
# =====================================================================
FUENTE = 'Arial'


def _aplicar_fuente(run, size, bold=False):
    run.font.name = FUENTE
    run.font.size = Pt(size)
    run.bold = bold


def generar_docx(datos, ruta_docx, ruta_firma=None):
    doc = Document()

    for s in doc.sections:
        s.top_margin    = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin   = Cm(2.5)
        s.right_margin  = Cm(2.5)

    normal = doc.styles['Normal']
    normal.font.name = FUENTE
    normal.font.size = Pt(10)

    def par(size=10, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT,
            line_pt=14, before=0, after=0):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after  = Pt(after)
        pf.line_spacing = Pt(line_pt)
        return p

    # Fecha
    if datos.get('fecha'):
        p = par(size=10, align=WD_ALIGN_PARAGRAPH.LEFT, line_pt=14, after=36)
        _aplicar_fuente(p.add_run(datos['fecha']), 10, False)

    # Título
    p = par(size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
            line_pt=18, after=24)
    _aplicar_fuente(p.add_run(f"CUENTA DE COBRO N° {datos['numero']}"),
                    13, True)

    # Empresa
    p = par(size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
            line_pt=18, after=0)
    _aplicar_fuente(p.add_run(datos['empresa']), 11, True)

    # NIT
    p = par(size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
            line_pt=18, after=12)
    _aplicar_fuente(p.add_run(f"NIT: {datos['nit_empresa']}"), 11, True)

    # DEBE A
    p = par(size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
            line_pt=18, after=0)
    _aplicar_fuente(p.add_run("DEBE A:"), 12, True)

    p = par(size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
            line_pt=18, after=0)
    _aplicar_fuente(p.add_run(datos['nombre']), 12, True)

    p = par(size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
            line_pt=18, after=18)
    _aplicar_fuente(p.add_run(
        f"CC. {datos['cc']} de {datos['ciudad']}"), 12, True)

    # EL VALOR DE
    p = par(size=10, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
            line_pt=14, after=10)
    _aplicar_fuente(p.add_run(
        f"EL VALOR DE: {datos['valor_letras']} PESOS MCTE "
        f"(${datos['valor_fmt']})"), 10, True)

    # POR CONCEPTO DE
    p = par(size=10, align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_pt=14, after=9)
    _aplicar_fuente(p.add_run("POR CONCEPTO DE: "), 10, True)
    _aplicar_fuente(p.add_run(
        f'Por horas de docencia en {datos["palabras"]} '
        f'{datos["asignaturas"]} con una intensidad de '), 10, False)
    _aplicar_fuente(p.add_run(
        f"({datos['horas_letras']}) ({datos['horas_num']}) horas"),
        10, True)
    _aplicar_fuente(p.add_run(
        f", iniciadas el {datos['pd_ini']} de ({datos['pm_ini']}) "
        f"y terminadas el {datos['pd_fin']} de ({datos['pm_fin']}) "
        f"de {datos['anio']}."), 10, False)

    # Declaraciones
    for d in DECLARACIONES:
        p = par(size=10, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                line_pt=14, after=8)
        _aplicar_fuente(p.add_run(d), 10, False)

    # Autorizo
    p = par(size=10, bold=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            line_pt=14, after=30)
    _aplicar_fuente(p.add_run(
        "Autorizo se me realice transferencia o consignación a la cuenta "
        "bancaria entregada al inicio de mi contratación."), 10, True)

    # Imagen de la firma (line_spacing en múltiplo, NO en Pt)
    if ruta_firma and os.path.exists(ruta_firma):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_img.paragraph_format.line_spacing = 1.0
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after  = Pt(2)
        run = p_img.add_run()
        run.add_picture(ruta_firma, width=Cm(4.5))

    # Línea de firma
    p = par(size=10, align=WD_ALIGN_PARAGRAPH.LEFT, line_pt=13, after=0)
    _aplicar_fuente(p.add_run("_" * 45), 10, False)

    # FIRMA / datos
    p = par(size=10, align=WD_ALIGN_PARAGRAPH.LEFT, line_pt=13, after=0)
    _aplicar_fuente(p.add_run("FIRMA"), 10, False)

    if datos.get('direccion'):
        p = par(size=10, align=WD_ALIGN_PARAGRAPH.LEFT, line_pt=13, after=0)
        _aplicar_fuente(p.add_run(f"DIRECCION: {datos['direccion']}"),
                        10, False)
    if datos.get('telefono'):
        p = par(size=10, align=WD_ALIGN_PARAGRAPH.LEFT, line_pt=13, after=0)
        _aplicar_fuente(p.add_run(f"TELEFONO: {datos['telefono']}"),
                        10, False)
    if datos.get('correo'):
        p = par(size=10, align=WD_ALIGN_PARAGRAPH.LEFT, line_pt=13, after=0)
        _aplicar_fuente(p.add_run(f"EMAIL: {datos['correo']}"),
                        10, False)

    doc.save(ruta_docx)


# =====================================================================
#  CONVERSIÓN DOCX → PDF (LibreOffice headless)
# =====================================================================
def convertir_a_pdf(docx_path, pdf_dir):
    """Convierte un DOCX a PDF usando LibreOffice headless.
    Devuelve la ruta del PDF o None si falla."""
    for cmd in ("libreoffice", "soffice"):
        if shutil.which(cmd):
            try:
                subprocess.run(
                    [cmd, "--headless", "--convert-to", "pdf",
                     "--outdir", pdf_dir, docx_path],
                    check=True, timeout=90,
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                )
                pdf_path = os.path.join(
                    pdf_dir,
                    os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
                )
                if os.path.exists(pdf_path):
                    return pdf_path
            except (subprocess.CalledProcessError,
                    subprocess.TimeoutExpired) as e:
                st.warning(f"Error en {cmd}: {e}")
    return None


# =====================================================================
#  INTERFAZ STREAMLIT
# =====================================================================
st.title("📄 Generador de Cuenta de Cobro")
st.caption("Formato institucional — Profesionales, técnicos y tecnólogos")

# ---------- Barra lateral ----------
with st.sidebar:
    st.header("⚙️ Configuración")
    st.markdown(
        "Rellena el formulario y presiona **Generar**. "
        "Recibirás dos botones para descargar el DOCX y el PDF."
    )
    st.markdown("---")
    st.markdown("**Tips:**")
    st.markdown("- La firma debe ser `.png` con fondo transparente.")
    st.markdown("- Los valores se convierten a letras automáticamente.")
    st.markdown("- El PDF se genera con LibreOffice en el servidor.")
    st.markdown("---")
    st.caption("Los valores por defecto son solo de ejemplo. "
               "Cámbialos por tus datos reales.")

# ---------- Datos generales ----------
st.subheader("1. Datos generales")
c1, c2 = st.columns([1, 3])
with c1:
    numero = st.text_input("N° de cuenta", value="01")
with c2:
    fecha_opcional = st.text_input(
        "Fecha (opcional, déjala vacía si no quieres mostrarla)",
        value="",
        placeholder="Ej: Ciudad, 15 de enero de 2026",
    )

# ---------- Empresa ----------
st.subheader("2. Empresa que contrata")
c1, c2 = st.columns([2, 1])
with c1:
    empresa = st.text_input("Nombre de la empresa", value=EMPRESA_DEFAULT)
with c2:
    nit_empresa = st.text_input("NIT de la empresa", value=NIT_EMPRESA_DEFAULT)

# ---------- Docente ----------
st.subheader("3. Docente")
c1, c2, c3 = st.columns([2, 1, 1])
with c1:
    nombre = st.text_input("Nombre completo", value=NOMBRE_DEFAULT)
with c2:
    cc = st.text_input("Cédula", value=CEDULA_DEFAULT)
with c3:
    ciudad = st.text_input("Ciudad (CC)", value=CIUDAD_DEFAULT)

# ---------- Período y horas ----------
st.subheader("4. Período y horas")
c1, c2, c3 = st.columns(3)
with c1:
    pd_ini = st.text_input("Día inicio", value="01")
    pm_ini = st.text_input("Mes inicio", value="enero")
with c2:
    pd_fin = st.text_input("Día fin", value="31")
    pm_fin = st.text_input("Mes fin", value="enero")
with c3:
    anio = st.text_input("Año", value="2026")
    horas = st.text_input("Horas (ej: 40 o 40,5)", value=HORAS_DEFAULT)

if horas.strip():
    try:
        st.caption(f"→ En letras: *{horas_letras(horas.strip())}*")
    except Exception:
        pass

# ---------- Asignaturas ----------
st.subheader("5. Asignaturas")
if "asignaturas" not in st.session_state:
    st.session_state.asignaturas = list(ASIGNATURAS_DEFAULT)

c1, c2 = st.columns([4, 1])
with c1:
    nueva = st.text_input("Agregar asignatura",
                          key="nueva_asig", label_visibility="collapsed",
                          placeholder="Escribe una asignatura y presiona Enter…")
with c2:
    if st.button("➕ Agregar", use_container_width=True):
        t = nueva.strip().strip('"')
        if t:
            st.session_state.asignaturas.append(t)
            st.rerun()

# Lista con botones de eliminar
for i, asig in enumerate(list(st.session_state.asignaturas)):
    c1, c2 = st.columns([10, 1])
    with c1:
        st.markdown(f'**{i+1}.** "{asig}"')
    with c2:
        if st.button("🗑️", key=f"del_{i}", help="Eliminar"):
            st.session_state.asignaturas.pop(i)
            st.rerun()

# Detectar Enter en el campo de nueva asignatura
if nueva and nueva.endswith("\n"):
    t = nueva.strip().strip('"')
    if t and t not in st.session_state.asignaturas:
        st.session_state.asignaturas.append(t)
        st.rerun()

# ---------- Valor ----------
st.subheader("6. Valor")
valor_str = st.text_input("Valor en pesos (sin puntos ni comas)",
                          value=VALOR_DEFAULT)
valor_limpio = re.sub(r'\D', '', valor_str)
if valor_limpio:
    st.caption(f"→ En letras: *{num_letras(valor_limpio).capitalize()}*")

# ---------- Firma ----------
st.subheader("7. Firma")
firma_file = st.file_uploader(
    "Sube tu firma (.png, .jpg)", type=["png", "jpg", "jpeg"])

# ---------- Contacto ----------
st.subheader("8. Contacto (opcional)")
c1, c2, c3 = st.columns(3)
with c1:
    direccion = st.text_input("Dirección", value="")
with c2:
    telefono = st.text_input("Teléfono", value="")
with c3:
    correo = st.text_input("Correo", value="")


# =====================================================================
#  BOTÓN GENERAR
# =====================================================================
st.markdown("---")
generar = st.button("🚀 Generar Cuenta de Cobro", type="primary",
                    use_container_width=True)

if generar:
    # ---------- Validaciones ----------
    errores = []
    if not numero.strip():
        errores.append("El N° de cuenta es obligatorio.")
    if not st.session_state.asignaturas:
        errores.append("Agrega al menos una asignatura.")
    if not valor_limpio:
        errores.append("Ingresa el valor.")
    if not horas.strip():
        errores.append("Ingresa las horas.")
    if not empresa.strip():
        errores.append("Ingresa el nombre de la empresa.")
    if not nombre.strip():
        errores.append("Ingresa el nombre del docente.")

    if errores:
        for e in errores:
            st.error(e)
        st.stop()

    # ---------- Preparar datos ----------
    palabras = ("las asignaturas" if len(st.session_state.asignaturas) > 1
                else "la asignatura")

    datos = {
        'numero': numero.strip(),
        'fecha': fecha_opcional.strip(),
        'empresa': empresa.strip(),
        'nit_empresa': nit_empresa.strip(),
        'nombre': nombre.strip(),
        'cc': cc.strip(),
        'ciudad': ciudad.strip(),
        'valor_fmt': f"{int(valor_limpio):,}".replace(',', '.'),
        'valor_letras': num_letras(valor_limpio),
        'palabras': palabras,
        'asignaturas': unir_asig(st.session_state.asignaturas),
        'horas_letras': horas_letras(horas.strip()),
        'horas_num': horas.strip(),
        'pd_ini': pd_ini, 'pm_ini': pm_ini,
        'pd_fin': pd_fin, 'pm_fin': pm_fin,
        'anio': anio,
        'direccion': direccion.strip(),
        'telefono': telefono.strip(),
        'correo': correo.strip(),
    }

    # ---------- Generar en carpeta temporal ----------
    with st.spinner("Generando documento…"):
        tmp = tempfile.mkdtemp()
        rdx = os.path.join(tmp, f"cuenta_cobro_{datos['numero']}.docx")
        r_firma = None
        if firma_file is not None:
            r_firma = os.path.join(tmp, "firma.png")
            with open(r_firma, "wb") as f:
                f.write(firma_file.getbuffer())

        try:
            generar_docx(datos, rdx, ruta_firma=r_firma)
        except Exception as e:
            st.error(f"Error al generar el DOCX: {e}")
            st.stop()

        # ---------- Convertir a PDF ----------
        with st.spinner("Convirtiendo a PDF…"):
            rpdf = convertir_a_pdf(rdx, tmp)

    # ---------- Botones de descarga ----------
    st.success("✅ Cuenta de cobro generada")

    c1, c2 = st.columns(2)
    with c1:
        with open(rdx, "rb") as f:
            st.download_button(
                label="📝 Descargar DOCX",
                data=f.read(),
                file_name=os.path.basename(rdx),
                mime=("application/vnd.openxmlformats-officedocument"
                      ".wordprocessingml.document"),
                use_container_width=True,
            )
    with c2:
        if rpdf and os.path.exists(rpdf):
            with open(rpdf, "rb") as f:
                st.download_button(
                    label="📄 Descargar PDF",
                    data=f.read(),
                    file_name=os.path.basename(rpdf),
                    mime="application/pdf",
                    use_container_width=True,
                )
        else:
            st.warning("No se pudo generar el PDF. "
                       "Descarga el DOCX y conviértelo localmente.")

    # Vista previa (si hay PDF)
    if rpdf and os.path.exists(rpdf):
        with st.expander("👁️ Vista previa del PDF"):
            try:
                import pymupdf
                doc = pymupdf.open(rpdf)
                pagina = doc[0]
                pix = pagina.get_pixmap(dpi=110)
                st.image(pix.tobytes("png"), use_container_width=True)
            except ImportError:
                st.info("Instala `pymupdf` para la vista previa.")
