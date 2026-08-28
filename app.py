"""
Reporte Técnico Astara — App Streamlit
========================================
Genera 'Reporte de falla Astara.xlsx' a partir de:
  - Orden de Trabajo.pdf
  - Trabajo en Garantías.pdf
  - Plantilla 'Reporte Técnico Astara.xlsx'
  - Imágenes de respaldo

Todo se sube en UN SOLO campo de carga de archivos; la app detecta
automáticamente qué es cada documento (por nombre y, si hace falta,
por contenido).

La lógica de extracción de datos es la misma del script original
(regex, mapeos de concesionarios, reglas de filtrado del PDF de
Garantías, etc.). Lo único que cambia es el manejo de archivos:
en vez de rutas fijas en disco, se trabaja en memoria (BytesIO)
para que la app funcione igual en local (.bat) o en Streamlit
Community Cloud (GitHub).
"""

import io
import re
import logging

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage, ImageOps

try:
    import pypdf
    PDF_LIBRARY_AVAILABLE = True
except ImportError:
    try:
        import PyPDF2 as pypdf
        PDF_LIBRARY_AVAILABLE = True
    except ImportError:
        PDF_LIBRARY_AVAILABLE = False

# --------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("reporte_astara")

# --------------------------------------------------------------------------
# Constantes (idénticas al script original)
# --------------------------------------------------------------------------
CM_TO_PX_96DPI = 37.7952
TARGET_WIDTH = int(13.19 * CM_TO_PX_96DPI)
TARGET_HEIGHT = int(5.69 * CM_TO_PX_96DPI)

IMG_START_ROW = 28
IMG_ROW_INCREMENT = 11
IMG_COLUMNS = ("C", "G")

CELDAS_EXCEL = {
    "vin": "G5",
    "fecha_falla": "D7",
    "fecha_cierre_ot": "G7",
    "km": "I7",
    "concesionario": "D9",
    "asesor_responsable": "G9",
    "tecnico": "I9",
    "queja_cliente": "C12",
    "diagnostico_tecnico": "C16",
    "repuestos_usados": "C20",
    "trabajo_realizado": "C24",
}

CAMPOS_CRITICOS = (
    "vin",
    "fecha_falla",
    "fecha_cierre_ot",
    "queja_cliente",
    "diagnostico_tecnico",
)

ETIQUETAS_CAMPOS = {
    "vin": "VIN",
    "fecha_falla": "Fecha de falla",
    "fecha_cierre_ot": "Fecha cierre OT",
    "km": "Kilometraje",
    "concesionario": "Concesionario",
    "asesor_responsable": "Asesor responsable",
    "tecnico": "Técnico",
    "queja_cliente": "Queja del cliente",
    "diagnostico_tecnico": "Diagnóstico técnico",
    "repuestos_usados": "Repuestos usados",
    "trabajo_realizado": "Trabajo realizado",
}

MAPA_CONCESIONARIOS = {
    "OTA": "ASTARA RETAIL ALAMEDA",
    "OTE": "ASTARA RETAIL EL CORTIJO",
    "OTC": "ASTARA RETAIL CANTAGALLO",
    "OTN": "ASTARA RETAIL NOS",
    "OTG": "ASTARA RETAIL LA DEHESA",
    "OTM": "ASTARA RETAIL MACHALÍ",
}

ETIQUETAS_BLOQUE_DIAGNOSTICO = {
    "cód. avería",
    "tipo garantía",
    "% deferencia mo",
    "gf",
    "kmts",
}

EXTENSIONES_IMAGEN = {"jpg", "jpeg", "png"}


# --------------------------------------------------------------------------
# Extracción de texto de PDF (centralizada, trabaja sobre bytes en memoria)
# --------------------------------------------------------------------------
def extraer_texto_pdf(file_bytes: bytes, nombre: str = "") -> str:
    """Extrae y concatena el texto de todas las páginas de un PDF en memoria."""
    if not PDF_LIBRARY_AVAILABLE:
        logger.warning("No hay librería de PDF disponible (pypdf/PyPDF2).")
        return ""
    if not file_bytes:
        logger.warning(f"Archivo PDF vacío: {nombre}")
        return ""
    try:
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        logger.error(f"Error leyendo el PDF '{nombre}': {e}")
        return ""


# --------------------------------------------------------------------------
# Extracción de datos — Orden de Trabajo
# --------------------------------------------------------------------------
def extraer_datos_pdf(texto_completo: str) -> dict:
    datos = {
        "vin": "", "fecha_falla": "", "fecha_cierre_ot": "", "km": "",
        "concesionario": "", "asesor_responsable": "", "tecnico": "",
        "queja_cliente": "", "diagnostico_tecnico": "", "trabajo_realizado": "",
    }
    if not texto_completo:
        return datos

    try:
        vin_match = re.search(r"VIN\s*[:]?\s*([A-HJ-NPR-Z0-9]{17})", texto_completo, re.IGNORECASE | re.DOTALL)
        if vin_match:
            datos["vin"] = vin_match.group(1)
        else:
            logger.warning("No se pudo extraer el VIN del PDF de Orden de Trabajo.")

        f_apertura_match = re.search(r"Fecha\s+apertura\s+([0-9]{2}-[0-9]{2}-[0-9]{4})", texto_completo, re.IGNORECASE | re.DOTALL)
        if f_apertura_match:
            datos["fecha_falla"] = f_apertura_match.group(1)
        else:
            logger.warning("No se pudo extraer la fecha de apertura/falla.")

        f_cierre_match = re.search(r"F\.\s*Impresi[oó]n\s+([0-9]{2}-[0-9]{2}-[0-9]{2,4})", texto_completo, re.IGNORECASE | re.DOTALL)
        if f_cierre_match:
            f_str = f_cierre_match.group(1)
            if len(f_str.split("-")[-1]) == 2:
                partes = f_str.split("-")
                f_str = f"{partes[0]}-{partes[1]}-20{partes[2]}"
            datos["fecha_cierre_ot"] = f_str
        else:
            logger.warning("No se pudo extraer la fecha de cierre de la OT.")

        km_match = re.search(r"Kilómetros\s+([\d\.]+)", texto_completo, re.IGNORECASE | re.DOTALL)
        if km_match:
            datos["km"] = km_match.group(1)

        ot_match = re.search(r"Nº\s*O\.?T\.?\s*([A-Z]{3})(?:/[0-9]+/[0-9]{4})?", texto_completo, re.IGNORECASE | re.DOTALL)
        if ot_match:
            prefix = ot_match.group(1).upper()
            if prefix in MAPA_CONCESIONARIOS:
                datos["concesionario"] = MAPA_CONCESIONARIOS[prefix]
            else:
                logger.warning(f"Prefijo de OT '{prefix}' no está en el mapa de concesionarios.")

        asesor_regex = (
            r"VEHÍCULO\s*\n\s*([^\n]+?)(?:\s*\([\d\s-]+\)|\s*Tel[\.\s:]|"
            r"\s*E-mail|\s*Fax|\s*Móvil|\s*\d{9,})"
        )
        asesor_match = re.search(asesor_regex, texto_completo, re.DOTALL)
        if asesor_match:
            datos["asesor_responsable"] = _normalizar_espacios(asesor_match.group(1))
        else:
            asesor_match_fb = re.search(r"VEHÍCULO\s+([A-Za-z\s]+?)(?:\s*\(|\s*-)", texto_completo, re.DOTALL)
            if asesor_match_fb:
                datos["asesor_responsable"] = _normalizar_espacios(asesor_match_fb.group(1))

        tecnico_match = re.search(r"Técnico\s*:\s*(.*?)(?:\n|\r|$)", texto_completo, re.IGNORECASE | re.DOTALL)
        if tecnico_match:
            datos["tecnico"] = tecnico_match.group(1).strip()

        queja_regex = (
            r"(?:Taller Mecanica|QUEJA CLIENTE|Denominaci[oó]n de Falla)\s*\n*(.*?)"
            r"(?=\n{2,}|$|DIAGNÓSTICO TÉCNICO|TRABAJO REALIZADO|REPUESTOS|MEDIDAS)"
        )
        queja_match = re.search(queja_regex, texto_completo, re.IGNORECASE | re.DOTALL)
        if queja_match:
            datos["queja_cliente"] = queja_match.group(1).strip()

        diagnostico_match = re.search(
            r"DIAGNÓSTICO TÉCNICO\s*\n*(.*?)(?=\n{2,}|$|TRABAJO REALIZADO|REPUESTOS|MEDIDAS)",
            texto_completo, re.IGNORECASE | re.DOTALL,
        )
        if diagnostico_match:
            datos["diagnostico_tecnico"] = diagnostico_match.group(1).strip()

        trabajo_match = re.search(r"TRABAJO REALIZADO\s*\n*(.*?)(?=\n{2,}|$|REPUESTOS|MEDIDAS)", texto_completo, re.IGNORECASE | re.DOTALL)
        if trabajo_match:
            datos["trabajo_realizado"] = trabajo_match.group(1).strip()

    except Exception as e:
        logger.error(f"Error inesperado al parsear Orden de Trabajo: {e}")

    return datos


# --------------------------------------------------------------------------
# Extracción de datos — Trabajo en Garantías
# --------------------------------------------------------------------------
def _normalizar_espacios(texto: str) -> str:
    return re.sub(r"\s+", " ", texto).strip()


def _es_linea_numerica(linea: str) -> bool:
    return bool(re.fullmatch(r"[\d.,\s]+", linea))


def extraer_datos_from_garantias_pdf(texto_completo: str) -> dict:
    datos_garantias = {
        "queja_cliente_garantias": "",
        "diagnostico_tecnico_garantias": "",
        "repuestos_usados_garantias": "",
    }
    if not texto_completo:
        return datos_garantias

    try:
        queja_match = re.search(r"F\.\s*Cierre\s*(.*?)\s*\d+\s*Núm\.", texto_completo, re.DOTALL)
        if queja_match:
            datos_garantias["queja_cliente_garantias"] = _normalizar_espacios(queja_match.group(1))
        else:
            logger.warning("No se pudo extraer 'queja_cliente' del PDF de Garantías.")

        block_match = re.search(r"Kmts\s*([\s\S]*?)Total", texto_completo, re.DOTALL)
        if block_match:
            lineas_utiles = []
            for linea in block_match.group(1).splitlines():
                linea = linea.strip()
                if not linea:
                    continue
                if linea.lower() in ETIQUETAS_BLOQUE_DIAGNOSTICO:
                    continue
                if _es_linea_numerica(linea):
                    continue
                lineas_utiles.append(linea)
            if lineas_utiles:
                datos_garantias["diagnostico_tecnico_garantias"] = _normalizar_espacios(" ".join(lineas_utiles))
            else:
                logger.warning("Bloque de diagnóstico vacío tras filtrar etiquetas y números.")
        else:
            logger.warning("No se pudo extraer el bloque de diagnóstico técnico de Garantías.")

        tabla_match = re.search(r"Referencia\s*\n\s*Tipo\s*([\s\S]*?)Fecha Reparación", texto_completo, re.DOTALL)
        if tabla_match:
            lineas_tabla = [l.strip() for l in tabla_match.group(1).split("\n")]
            pares_repuestos = []
            for i, linea in enumerate(lineas_tabla):
                if linea.upper() != "MAT":
                    continue
                referencia = lineas_tabla[i - 1] if i - 1 >= 0 else ""
                denominacion = lineas_tabla[i - 2] if i - 2 >= 0 else ""
                if not referencia or not denominacion:
                    logger.warning(f"Fila de repuesto incompleta cerca de la línea {i}; se omite.")
                    continue
                codigo = referencia.split("-", 1)[1].strip() if "-" in referencia else referencia
                pares_repuestos.append(f"{codigo} - {denominacion}")
            if pares_repuestos:
                datos_garantias["repuestos_usados_garantias"] = "\n".join(pares_repuestos)
            else:
                logger.warning("No se encontraron filas de tipo 'MAT' (repuestos) en Garantías.")
        else:
            logger.warning("No se pudo extraer la tabla de repuestos de Garantías.")

    except Exception as e:
        logger.error(f"Error inesperado al parsear Trabajo en Garantías: {e}")

    return datos_garantias


def validar_campos_criticos(valores: dict) -> list:
    """Devuelve la lista de campos críticos que quedaron vacíos."""
    faltantes = [c for c in CAMPOS_CRITICOS if not valores.get(c)]
    for campo in faltantes:
        logger.warning(f"⚠ Campo crítico vacío tras la extracción: '{campo}'")
    return faltantes


# --------------------------------------------------------------------------
# Clasificación automática de archivos subidos en el campo único
# --------------------------------------------------------------------------
def clasificar_pdf(texto: str, nombre: str) -> str:
    """Determina si un PDF es 'orden', 'garantias' o 'desconocido',
    primero por nombre de archivo y, si no es concluyente, por contenido."""
    nombre_low = nombre.lower()
    if "garant" in nombre_low:
        return "garantias"
    if "orden" in nombre_low and "garant" not in nombre_low:
        return "orden"

    texto_low = texto.lower()
    marcadores_garantia = ["núm.", "kmts", "% deferencia", "tipo garantía", "fecha reparación"]
    marcadores_orden = ["vin", "kilómetros", "nº o.t", "queja cliente", "vehículo"]
    score_g = sum(1 for m in marcadores_garantia if m in texto_low)
    score_o = sum(1 for m in marcadores_orden if m in texto_low)
    if score_g > score_o:
        return "garantias"
    if score_o > score_g:
        return "orden"
    return "desconocido"


def clasificar_archivos_subidos(archivos_subidos: list) -> dict:
    """Recibe la lista de UploadedFile del campo único y los separa en:
    plantilla Excel, PDF Orden de Trabajo, PDF Garantías e imágenes,
    deduplicando imágenes por nombre (case-insensitive)."""
    resultado = {
        "excel_template": None,
        "pdf_orden": None,
        "pdf_orden_nombre": "",
        "pdf_garantias": None,
        "pdf_garantias_nombre": "",
        "imagenes": [],
        "sin_clasificar": [],
    }
    vistos_imagenes = set()
    pdfs_pendientes = []

    for archivo in archivos_subidos:
        nombre = archivo.name
        ext = nombre.lower().rsplit(".", 1)[-1] if "." in nombre else ""

        if ext == "xlsx":
            if resultado["excel_template"] is None:
                resultado["excel_template"] = archivo.getvalue()
            continue

        if ext == "pdf":
            data = archivo.getvalue()
            texto = extraer_texto_pdf(data, nombre)
            tipo = clasificar_pdf(texto, nombre)
            if tipo == "garantias" and resultado["pdf_garantias"] is None:
                resultado["pdf_garantias"] = texto
                resultado["pdf_garantias_nombre"] = nombre
            elif tipo == "orden" and resultado["pdf_orden"] is None:
                resultado["pdf_orden"] = texto
                resultado["pdf_orden_nombre"] = nombre
            else:
                pdfs_pendientes.append((nombre, texto))
            continue

        if ext in EXTENSIONES_IMAGEN:
            clave = nombre.lower()
            if clave not in vistos_imagenes:
                vistos_imagenes.add(clave)
                resultado["imagenes"].append(archivo)
            continue

        resultado["sin_clasificar"].append(nombre)

    # Los PDF que no se pudieron clasificar por contenido/nombre se asignan
    # a los huecos que queden, en el orden en que se subieron.
    for nombre, texto in pdfs_pendientes:
        if resultado["pdf_orden"] is None:
            resultado["pdf_orden"] = texto
            resultado["pdf_orden_nombre"] = nombre
        elif resultado["pdf_garantias"] is None:
            resultado["pdf_garantias"] = texto
            resultado["pdf_garantias_nombre"] = nombre
        else:
            resultado["sin_clasificar"].append(nombre)

    return resultado


# --------------------------------------------------------------------------
# Imágenes
# --------------------------------------------------------------------------
def calcular_celda_imagen(indice: int) -> str:
    fila = IMG_START_ROW + (indice // len(IMG_COLUMNS)) * IMG_ROW_INCREMENT
    columna = IMG_COLUMNS[indice % len(IMG_COLUMNS)]
    return f"{columna}{fila}"


def preparar_imagen(archivo_subido) -> io.BytesIO:
    """Aplica la misma transformación que el script original
    (orientación EXIF, RGBA->RGB, rotación si es vertical) y devuelve
    un buffer JPEG en memoria listo para insertar.

    Importante: openpyxl necesita reabrir la imagen desde un archivo
    o un buffer (usa 'fp' internamente al guardar), así que no basta
    con pasarle el objeto PIL ya transformado -eso rompe el guardado-;
    por eso el resultado se vuelca a un BytesIO en vez de devolver el
    objeto PIL directamente."""
    pil_img = PILImage.open(io.BytesIO(archivo_subido.getvalue()))
    pil_img = ImageOps.exif_transpose(pil_img)
    if pil_img.mode == "RGBA":
        pil_img = pil_img.convert("RGB")
    width, height = pil_img.size
    if height > width:
        pil_img = pil_img.rotate(-90, expand=True)

    buffer = io.BytesIO()
    pil_img.save(buffer, format="JPEG", quality=88)
    buffer.seek(0)
    return buffer


def insertar_imagenes(ws, archivos_imagenes: list) -> int:
    """Inserta las imágenes en la hoja. Devuelve cuántas se insertaron
    correctamente."""
    insertadas = 0
    for i, archivo in enumerate(archivos_imagenes):
        celda = calcular_celda_imagen(i)
        try:
            buffer = preparar_imagen(archivo)
            img = XLImage(buffer)
            img.width = TARGET_WIDTH
            img.height = TARGET_HEIGHT
            ws.add_image(img, celda)
            insertadas += 1
        except Exception as e:
            logger.error(f"Error al procesar/insertar la imagen '{archivo.name}': {e}")
    return insertadas


# --------------------------------------------------------------------------
# Generación del reporte
# --------------------------------------------------------------------------
def generar_reporte(clasificacion: dict) -> tuple:
    """Ejecuta todo el flujo y devuelve (bytes_excel, valores, faltantes,
    n_imagenes_insertadas)."""
    datos_ot = extraer_datos_pdf(clasificacion["pdf_orden"] or "")
    datos_garantias = extraer_datos_from_garantias_pdf(clasificacion["pdf_garantias"] or "")

    valores = {
        "vin": datos_ot["vin"],
        "fecha_falla": datos_ot["fecha_falla"],
        "fecha_cierre_ot": datos_ot["fecha_cierre_ot"],
        "km": datos_ot["km"],
        "concesionario": datos_ot["concesionario"],
        "asesor_responsable": datos_ot["asesor_responsable"],
        "tecnico": datos_ot.get("tecnico", ""),
        "trabajo_realizado": datos_ot.get("trabajo_realizado", ""),
        "queja_cliente": datos_garantias.get("queja_cliente_garantias", ""),
        "diagnostico_tecnico": datos_garantias.get("diagnostico_tecnico_garantias", ""),
        "repuestos_usados": datos_garantias.get("repuestos_usados_garantias", ""),
    }
    faltantes = validar_campos_criticos(valores)

    wb = openpyxl.load_workbook(io.BytesIO(clasificacion["excel_template"]))
    ws = wb.active

    for campo, celda in CELDAS_EXCEL.items():
        ws[celda] = valores.get(campo, "")

    n_insertadas = insertar_imagenes(ws, clasificacion["imagenes"])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer.getvalue(), valores, faltantes, n_insertadas


# --------------------------------------------------------------------------
# Interfaz Streamlit — tema automotriz
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Reporte Técnico Astara",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
:root {
    --astara-red: #E10600;
    --astara-dark: #10131A;
}
.stApp {
    background: radial-gradient(circle at top left, #171B24 0%, #0B0D12 60%);
}
[data-testid="stHeader"] { background: transparent; }
.astara-hero {
    background: linear-gradient(120deg, #E10600 0%, #8C0400 100%);
    padding: 1.6rem 2rem;
    border-radius: 14px;
    margin-bottom: 1.4rem;
    box-shadow: 0 8px 24px rgba(225, 6, 0, 0.25);
}
.astara-hero h1 {
    color: white;
    font-size: 2rem;
    margin: 0;
    letter-spacing: 0.5px;
}
.astara-hero p {
    color: #FFE3E1;
    margin: 0.3rem 0 0 0;
    font-size: 0.95rem;
}
.astara-card {
    background: #171B24;
    border: 1px solid #2A2F3A;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
}
.astara-badge {
    display: inline-block;
    background: #232833;
    color: #E10600;
    border: 1px solid #E10600;
    border-radius: 999px;
    padding: 2px 12px;
    font-size: 0.78rem;
    font-weight: 600;
    margin-right: 6px;
    margin-bottom: 6px;
}
.astara-badge.ok { color: #34D399; border-color: #34D399; }
.astara-badge.warn { color: #FBBF24; border-color: #FBBF24; }
.astara-badge.missing { color: #9CA3AF; border-color: #4B5563; }
div[data-testid="stFileUploaderDropzone"] {
    border: 1.5px dashed #E10600 !important;
    background: #14171F !important;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

st.markdown(
    """
    <div class="astara-hero">
        <h1>🚗 Reporte Técnico Astara</h1>
        <p>Generador automático del Reporte de Falla a partir de la Orden de Trabajo,
        el Trabajo en Garantías y las fotografías del vehículo.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### ℹ️ Cómo usar la app")
    st.markdown(
        """
        1. Arrastra **todos** tus archivos al campo de carga:
           - `Orden de Trabajo.pdf`
           - `Trabajo en Garantías.pdf`
           - `Reporte Técnico Astara.xlsx` (plantilla)
           - Fotos del vehículo (`.jpg`, `.jpeg`, `.png`)
        2. La app detecta sola qué es cada archivo.
        3. Revisa la vista previa de los datos importados.
        4. Descarga el reporte final ya rellenado.
        """
    )
    st.divider()
    st.caption("No es necesario un campo por tipo de archivo: todo se sube junto y la app los diferencia automáticamente.")

archivos_subidos = st.file_uploader(
    "Adjunta aquí la Orden de Trabajo, el Trabajo en Garantías, la plantilla del Reporte Técnico Astara y las fotos",
    type=["pdf", "xlsx", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if archivos_subidos:
    clasificacion = clasificar_archivos_subidos(archivos_subidos)

    st.markdown("#### 📎 Archivos detectados")
    badges = []
    badges.append(
        f'<span class="astara-badge {"ok" if clasificacion["excel_template"] else "missing"}">'
        f'Plantilla Excel: {"✔ detectada" if clasificacion["excel_template"] else "✘ falta"}</span>'
    )
    badges.append(
        f'<span class="astara-badge {"ok" if clasificacion["pdf_orden"] else "missing"}">'
        f'Orden de Trabajo: {clasificacion["pdf_orden_nombre"] or "✘ falta"}</span>'
    )
    badges.append(
        f'<span class="astara-badge {"ok" if clasificacion["pdf_garantias"] else "missing"}">'
        f'Trabajo en Garantías: {clasificacion["pdf_garantias_nombre"] or "✘ falta"}</span>'
    )
    badges.append(f'<span class="astara-badge ok">Fotos: {len(clasificacion["imagenes"])}</span>')
    st.markdown(" ".join(badges), unsafe_allow_html=True)

    if clasificacion["sin_clasificar"]:
        st.warning("Archivos que no se pudieron clasificar y fueron ignorados: " + ", ".join(clasificacion["sin_clasificar"]))

    faltan_obligatorios = not (clasificacion["excel_template"] and clasificacion["pdf_orden"] and clasificacion["pdf_garantias"])

    if faltan_obligatorios:
        st.info("Sube la plantilla Excel, la Orden de Trabajo y el Trabajo en Garantías para poder generar el reporte.")
    else:
        generar = st.button("⚙️ Generar Reporte Técnico Astara", type="primary", use_container_width=True)

        if generar:
            with st.spinner("Extrayendo datos y armando el reporte..."):
                try:
                    excel_bytes, valores, faltantes, n_insertadas = generar_reporte(clasificacion)
                except Exception as e:
                    st.error(f"No se pudo generar el reporte: {e}")
                    logger.error(f"Error generando el reporte: {e}")
                    excel_bytes = None

            if excel_bytes:
                if faltantes:
                    etiquetas_faltantes = ", ".join(ETIQUETAS_CAMPOS.get(c, c) for c in faltantes)
                    st.warning(f"⚠ Algunos campos críticos quedaron vacíos y conviene revisarlos a mano: {etiquetas_faltantes}")
                else:
                    st.success("Todos los campos críticos se completaron correctamente.")

                st.markdown("#### 👀 Vista previa de los datos importados al Excel")
                df_preview = pd.DataFrame(
                    [
                        {
                            "Campo": ETIQUETAS_CAMPOS.get(campo, campo),
                            "Celda": celda,
                            "Valor importado": valores.get(campo, "") or "—",
                        }
                        for campo, celda in CELDAS_EXCEL.items()
                    ]
                )
                st.dataframe(df_preview, use_container_width=True, hide_index=True)

                if clasificacion["imagenes"]:
                    st.markdown(f"#### 🖼️ Fotos insertadas en el reporte ({n_insertadas}/{len(clasificacion['imagenes'])})")
                    columnas = st.columns(4)
                    for i, archivo in enumerate(clasificacion["imagenes"]):
                        with columnas[i % 4]:
                            st.image(archivo.getvalue(), caption=f"{calcular_celda_imagen(i)} · {archivo.name}", use_container_width=True)

                st.markdown("#### ⬇️ Descargar reporte final")
                st.download_button(
                    label="Descargar 'Reporte de falla Astara.xlsx'",
                    data=excel_bytes,
                    file_name="Reporte de falla Astara.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
else:
    st.markdown(
        """
        <div class="astara-card">
        Aún no se han cargado archivos. Sube todos los documentos en el campo de arriba
        para comenzar.
        </div>
        """,
        unsafe_allow_html=True,
    )
