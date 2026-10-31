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

Control de peso del reporte (v2)
--------------------------------
El Excel final nunca supera los 25 MB, sin importar cuántas fotos se
adjunten (hasta ~25 o más):
  - Las 2 PRIMERAS fotos se guardan siempre con alta calidad.
  - Las demás se guardan con la misma calidad mientras el total quepa
    en el límite; solo si no cabe, se comprimen por niveles (resolución
    y calidad JPEG) hasta que el reporte quede bajo el límite.
  - Únicamente en un caso extremo (que ni con las secundarias al mínimo
    quepa) se reducen también las 2 primeras.
  - Al final se verifica el tamaño REAL del .xlsx guardado.
"""

import io
import re
import base64
import logging
import tempfile
import unicodedata
from pathlib import Path

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage, ImageOps, ImageDraw

# Soporte opcional para fotos de iPhone (HEIC/HEIF). Requiere 'pillow-heif'
# en requirements.txt; si no está instalado, esas fotos simplemente se ignoran.
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIF_DISPONIBLE = True
except ImportError:
    HEIF_DISPONIBLE = False

try:
    import pypdf
    PDF_LIBRARY_AVAILABLE = True
except ImportError:
    try:
        import PyPDF2 as pypdf
        PDF_LIBRARY_AVAILABLE = True
    except ImportError:
        PDF_LIBRARY_AVAILABLE = False

# Evita el aviso de "decompression bomb" con fotos de cámaras de alta resolución
PILImage.MAX_IMAGE_PIXELS = 200_000_000

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

# --------------------------------------------------------------------------
# Control de peso del reporte
# --------------------------------------------------------------------------
MB = 1024 * 1024
LIMITE_REPORTE_BYTES = 25 * MB          # límite duro solicitado
OBJETIVO_REPORTE_BYTES = int(24.0 * MB)  # se apunta un poco más abajo como margen de seguridad
N_FOTOS_PRIORITARIAS = 2                 # las 2 primeras fotos priorizan calidad

# (lado máximo en px, calidad JPEG). El nivel 0 equivale a la calidad
# original de la app (resolución completa, calidad 88).
NIVELES_PRIORITARIAS = [
    (4096, 92),
    (3200, 90),
    (2560, 88),
    (2048, 85),
]
NIVELES_SECUNDARIAS = [
    (4096, 88),   # sin ahorro: igual que antes
    (3000, 85),
    (2400, 83),
    (2000, 80),
    (1600, 78),
    (1366, 75),
    (1152, 72),
    (1024, 68),
    (900, 64),
    (800, 60),
]

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

# Formatos de foto aceptados. HEIC/HEIF (iPhone) solo si 'pillow-heif'
# está instalado.
EXTENSIONES_IMAGEN = {"jpg", "jpeg", "jfif", "png", "webp", "bmp", "gif", "tif", "tiff"}
if HEIF_DISPONIBLE:
    EXTENSIONES_IMAGEN |= {"heic", "heif"}

# Nombres EXACTOS (sin importar mayúsculas, tildes, guiones/guiones bajos ni
# copias tipo "(1)") que debe tener cada PDF para ser reconocido.
NOMBRE_PDF_ORDEN = "orden de trabajo"
NOMBRES_PDF_GARANTIAS = {"trabajo en garantias", "trabajo en garantia"}

# Plantilla: siempre la incluida en el repositorio, junto a este archivo.
TEMPLATE_FILENAME = "Reporte_de_falla_Astara.xlsx"
TEMPLATE_PATH = Path(__file__).resolve().parent / TEMPLATE_FILENAME


def cargar_plantilla_predeterminada() -> bytes:
    """Devuelve los bytes de la plantilla incluida en el repositorio,
    o None si no está presente."""
    if not TEMPLATE_PATH.exists():
        return None
    try:
        return TEMPLATE_PATH.read_bytes()
    except Exception as e:
        logger.error(f"No se pudo leer la plantilla predeterminada: {e}")
        return None


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
def normalizar_nombre_archivo(nombre: str) -> str:
    """Normaliza el nombre (sin extensión) para compararlo:
    minúsculas, sin tildes, '_' '-' '.' como espacios, sin marcas de copia
    del navegador/Windows como ' (1)' o ' - copia', y espacios colapsados.
      'Trabajo_en_Garantías (2).pdf' -> 'trabajo en garantias'"""
    base = nombre.rsplit(".", 1)[0] if "." in nombre else nombre
    base = unicodedata.normalize("NFKD", base)
    base = "".join(c for c in base if not unicodedata.combining(c)).lower()
    base = re.sub(r"[_\-.]+", " ", base)
    base = re.sub(r"\s*\(\d+\)\s*$", "", base)        # "(1)", "(2)"...
    base = re.sub(r"\s+(copia|copy)(\s*\d+)?\s*$", "", base)  # "- copia", "copy 2"
    return re.sub(r"\s+", " ", base).strip()


def clasificar_pdf(nombre: str) -> str:
    """Devuelve 'orden', 'garantias' o 'ignorar' SOLO según el nombre.
    Cualquier PDF con otro nombre se ignora (no se adivina por contenido)."""
    clave = normalizar_nombre_archivo(nombre)
    if clave == NOMBRE_PDF_ORDEN:
        return "orden"
    if clave in NOMBRES_PDF_GARANTIAS:
        return "garantias"
    return "ignorar"


def _es_imagen_valida(datos: bytes) -> bool:
    """Comprueba que el archivo realmente sea una imagen legible (lee solo
    la cabecera, es rápido)."""
    try:
        with PILImage.open(io.BytesIO(datos)) as img:
            img.verify()
        return True
    except Exception:
        return False


def clasificar_archivos_subidos(archivos_subidos: list) -> dict:
    """Recibe la lista de UploadedFile del campo único (puede venir vacía)
    y se queda SOLO con:
      - el PDF llamado 'Orden de Trabajo'
      - el PDF llamado 'Trabajo en Garantías'
      - las fotos (jpg, jpeg, png, webp, heic, ...), sin repetir nombre
    Todo lo demás (otros PDF, Excel, Word, etc.) se ignora y se informa.
    Si se sube más de un PDF con el mismo nombre, se usa el primero."""
    resultado = {
        "pdf_orden": None,
        "pdf_orden_nombre": "",
        "pdf_garantias": None,
        "pdf_garantias_nombre": "",
        "imagenes": [],
        "ignorados": [],
    }
    vistos_imagenes = set()

    for archivo in archivos_subidos:
        nombre = archivo.name
        ext = nombre.lower().rsplit(".", 1)[-1] if "." in nombre else ""

        if ext == "pdf":
            tipo = clasificar_pdf(nombre)
            if tipo == "orden" and resultado["pdf_orden"] is None:
                resultado["pdf_orden"] = extraer_texto_pdf(archivo.getvalue(), nombre)
                resultado["pdf_orden_nombre"] = nombre
            elif tipo == "garantias" and resultado["pdf_garantias"] is None:
                resultado["pdf_garantias"] = extraer_texto_pdf(archivo.getvalue(), nombre)
                resultado["pdf_garantias_nombre"] = nombre
            else:
                resultado["ignorados"].append(nombre)
            continue

        if ext in EXTENSIONES_IMAGEN:
            clave = nombre.lower()
            if clave in vistos_imagenes or not _es_imagen_valida(archivo.getvalue()):
                resultado["ignorados"].append(nombre)
                continue
            vistos_imagenes.add(clave)
            resultado["imagenes"].append(archivo)
            continue

        resultado["ignorados"].append(nombre)

    # Orden natural por nombre ("1", "2", ..., "10", "IMG_...", "WhatsApp..."):
    # así el orden no depende de cómo el navegador entregue los archivos, y
    # basta con nombrar "1" y "2" las fotos más importantes para que sean
    # las prioritarias (alta calidad) y queden primeras en el reporte.
    resultado["imagenes"].sort(key=lambda a: _clave_orden_natural(a.name))
    return resultado


def _clave_orden_natural(nombre: str) -> list:
    """'foto 10.jpg' va después de 'foto 2.jpg' (no antes, como en orden
    alfabético simple)."""
    base = nombre.rsplit(".", 1)[0].lower()
    return [(0, int(t), "") if t.isdigit() else (1, 0, t) for t in re.split(r"(\d+)", base) if t]


# --------------------------------------------------------------------------
# Imágenes — preparación y compresión adaptativa
# --------------------------------------------------------------------------
def calcular_celda_imagen(indice: int) -> str:
    fila = IMG_START_ROW + (indice // len(IMG_COLUMNS)) * IMG_ROW_INCREMENT
    columna = IMG_COLUMNS[indice % len(IMG_COLUMNS)]
    return f"{columna}{fila}"


def _abrir_normalizada(datos: bytes, lado_max: int, edicion: dict = None) -> PILImage.Image:
    """Abre la imagen y aplica la misma transformación que el script
    original (orientación EXIF, conversión a RGB, rotación si es vertical),
    luego las ediciones del usuario (rotación y recorte, ver aplicar_edicion)
    y la reduce a 'lado_max' px en su lado más largo si es más grande.

    Se usa draft() para que, en JPEG, el decodificador lea directamente a
    una escala menor cuando se va a reducir: ahorra mucha RAM y tiempo con
    fotos de celular de 12-50 MP (importante en Streamlit Cloud)."""
    img = PILImage.open(io.BytesIO(datos))
    if img.format == "JPEG":
        lado_decod = lado_max
        recorte = (edicion or {}).get("crop")
        if recorte:  # si se recorta, se decodifica más grande para no perder nitidez
            fraccion = max(min(recorte[2] - recorte[0], recorte[3] - recorte[1]), 0.05)
            lado_decod = int(lado_max / fraccion)
        img.draft("RGB", (lado_decod, lado_decod))
    img = ImageOps.exif_transpose(img)

    # Transparencias (PNG con alfa) sobre fondo blanco; resto de modos a RGB.
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        fondo = PILImage.new("RGB", img.size, (255, 255, 255))
        fondo.paste(img, mask=img.split()[-1])
        img = fondo
    elif img.mode != "RGB":
        img = img.convert("RGB")

    width, height = img.size
    if height > width:
        img = img.rotate(-90, expand=True)

    img = aplicar_edicion(img, edicion)

    if max(img.size) > lado_max:
        img.thumbnail((lado_max, lado_max), PILImage.LANCZOS)
    return img


def aplicar_edicion(img: PILImage.Image, edicion: dict) -> PILImage.Image:
    """Aplica la edición del usuario sobre la imagen ya normalizada:
      - 'rot': grados en sentido horario (0, 90, 180, 270)
      - 'crop': (izq, arriba, der, abajo) como FRACCIONES 0-1 de la imagen
        ya rotada, así sirve igual para la miniatura y para la foto completa.
    """
    if not edicion:
        return img
    rot = edicion.get("rot", 0) % 360
    if rot:
        img = img.rotate(-rot, expand=True)
    recorte = edicion.get("crop")
    if recorte:
        w, h = img.size
        izq, arr, der, aba = (max(0.0, min(1.0, float(v))) for v in recorte)
        caja = (int(izq * w), int(arr * h), int(der * w), int(aba * h))
        if caja[2] - caja[0] >= 10 and caja[3] - caja[1] >= 10:
            img = img.crop(caja)
    return img


def codificar_imagen(datos: bytes, lado_max: int, calidad: int, edicion: dict = None) -> bytes:
    """Devuelve la imagen normalizada (y editada) como JPEG con el tamaño/calidad dados."""
    img = _abrir_normalizada(datos, lado_max, edicion)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=calidad, optimize=True, progressive=True)
    img.close()
    return buffer.getvalue()


def comprimir_fotos_para_limite(fotos: list, presupuesto_bytes: int) -> tuple:
    """Decide la compresión de cada foto para que la suma quepa en
    'presupuesto_bytes'.

    fotos: lista de (nombre, bytes_originales, edicion) en el orden del reporte.

    Estrategia:
      1. Las N_FOTOS_PRIORITARIAS primeras van en su mejor nivel.
      2. El resto va en el nivel 0 (sin ahorro). Si el total no cabe,
         se baja un nivel a TODAS las secundarias y se vuelve a probar,
         hasta que quepa.
      3. Solo si ni con las secundarias al mínimo cabe, se bajan también
         las prioritarias, nivel a nivel.

    Devuelve (lista_jpeg_bytes, info) con info = {'nivel_prioritarias',
    'nivel_secundarias', 'total_bytes', 'cabe'}. Las fotos que fallen al
    procesarse quedan como None.
    """
    n = len(fotos)
    idx_prio = list(range(min(N_FOTOS_PRIORITARIAS, n)))
    idx_sec = list(range(len(idx_prio), n))
    resultado = [None] * n

    def _codificar(indices, nivel_tabla, nivel):
        lado, calidad = nivel_tabla[nivel]
        for i in indices:
            nombre, datos, edicion = fotos[i]
            try:
                resultado[i] = codificar_imagen(datos, lado, calidad, edicion)
            except Exception as e:
                logger.error(f"No se pudo procesar la imagen '{nombre}': {e}")
                resultado[i] = None

    def _total(indices):
        return sum(len(resultado[i]) for i in indices if resultado[i])

    nivel_p = 0
    nivel_s = 0
    _codificar(idx_prio, NIVELES_PRIORITARIAS, nivel_p)
    _codificar(idx_sec, NIVELES_SECUNDARIAS, nivel_s)

    while _total(range(n)) > presupuesto_bytes:
        if idx_sec and nivel_s < len(NIVELES_SECUNDARIAS) - 1:
            # Salto inteligente: estima cuánto hay que reducir para no
            # recodificar nivel por nivel cuando el exceso es grande.
            total_sec = _total(idx_sec)
            disponible_sec = presupuesto_bytes - _total(idx_prio)
            nivel_s += 1
            if disponible_sec > 0 and total_sec > 2.5 * disponible_sec:
                nivel_s = min(nivel_s + 1, len(NIVELES_SECUNDARIAS) - 1)
            _codificar(idx_sec, NIVELES_SECUNDARIAS, nivel_s)
        elif idx_prio and nivel_p < len(NIVELES_PRIORITARIAS) - 1:
            nivel_p += 1
            _codificar(idx_prio, NIVELES_PRIORITARIAS, nivel_p)
        else:
            break  # ya está todo al mínimo

    total = _total(range(n))
    info = {
        "nivel_prioritarias": nivel_p,
        "nivel_secundarias": nivel_s if idx_sec else None,
        "total_bytes": total,
        "cabe": total <= presupuesto_bytes,
    }
    logger.info(
        f"Fotos: {n} | nivel prioritarias={nivel_p} | nivel secundarias={info['nivel_secundarias']} "
        f"| total imágenes={total / MB:.2f} MB (presupuesto {presupuesto_bytes / MB:.2f} MB)"
    )
    return resultado, info


def insertar_imagenes(ws, imagenes_jpeg: list, nombres: list) -> int:
    """Inserta las imágenes (ya comprimidas) en la hoja. Devuelve cuántas
    se insertaron correctamente.

    openpyxl necesita reabrir la imagen desde un buffer al guardar, por eso
    a cada XLImage se le entrega su propio BytesIO."""
    insertadas = 0
    for i, jpeg in enumerate(imagenes_jpeg):
        if not jpeg:
            continue
        celda = calcular_celda_imagen(i)
        try:
            img = XLImage(io.BytesIO(jpeg))
            img.width = TARGET_WIDTH
            img.height = TARGET_HEIGHT
            ws.add_image(img, celda)
            insertadas += 1
        except Exception as e:
            logger.error(f"Error al insertar la imagen '{nombres[i]}': {e}")
    return insertadas


# --------------------------------------------------------------------------
# Generación del reporte
# --------------------------------------------------------------------------
def _escribir_celda(ws, coordenada: str, valor) -> None:
    """Escribe un valor aunque la celda forme parte de un rango combinado
    (en ese caso se escribe en la celda superior izquierda del rango, que
    es la única editable). Evita el error "'MergedCell' object attribute
    'value' is read-only"."""
    celda = ws[coordenada]
    if isinstance(celda, MergedCell):
        for rango in ws.merged_cells.ranges:
            if coordenada in rango:
                ws.cell(row=rango.min_row, column=rango.min_col).value = valor
                return
        logger.warning(f"No se pudo escribir en la celda combinada {coordenada}.")
        return
    celda.value = valor


def _ancho_columna(ws, letra: str) -> float:
    dim = ws.column_dimensions.get(letra)
    return dim.width if dim is not None and dim.width else 8.43


def _ajustar_alto_celda_texto(ws, coordenada: str, texto: str) -> None:
    """Si el texto no cabe en el recuadro combinado (ej. C16:I18), agranda
    el alto de sus filas para que se lea completo en el Excel. Las fotos
    están ancladas a celdas, así que se desplazan solas hacia abajo."""
    if not texto:
        return
    rango = next((r for r in ws.merged_cells.ranges if coordenada in r), None)
    if rango is None:
        return
    ancho_chars = sum(
        _ancho_columna(ws, openpyxl.utils.get_column_letter(c))
        for c in range(rango.min_col, rango.max_col + 1)
    ) * 1.05  # ~caracteres por línea con fuente de 10-11 pt
    lineas = 0
    for parrafo in str(texto).split("\n"):
        lineas += max(1, -(-len(parrafo) // max(int(ancho_chars), 1)))  # división hacia arriba
    alto_necesario = lineas * 15.0 + 4
    filas = list(range(rango.min_row, rango.max_row + 1))
    alto_actual = sum((ws.row_dimensions[f].height or 15.0) for f in filas)
    if alto_necesario <= alto_actual:
        return
    alto_por_fila = min(alto_necesario / len(filas), 409)  # 409 = máximo de Excel
    for f in filas:
        ws.row_dimensions[f].height = alto_por_fila

    celda = ws.cell(row=rango.min_row, column=rango.min_col)
    al = celda.alignment
    celda.alignment = openpyxl.styles.Alignment(
        horizontal=al.horizontal, vertical=al.vertical or "top", wrap_text=True,
        indent=al.indent, text_rotation=al.text_rotation, shrink_to_fit=False,
    )


# C20 (repuestos) mantiene el alto original de la plantilla aunque no se vea todo.
CAMPOS_TEXTO_LARGO = ("queja_cliente", "diagnostico_tecnico", "trabajo_realizado")


def _armar_excel(plantilla_bytes: bytes, valores: dict, imagenes_jpeg: list, nombres: list) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(plantilla_bytes))
    ws = wb.active
    for campo, celda in CELDAS_EXCEL.items():
        _escribir_celda(ws, celda, valores.get(campo, ""))
    for campo in CAMPOS_TEXTO_LARGO:
        _ajustar_alto_celda_texto(ws, CELDAS_EXCEL[campo], valores.get(campo, ""))
    n_insertadas = insertar_imagenes(ws, imagenes_jpeg, nombres)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), n_insertadas


def _peso_base_excel(plantilla_bytes: bytes, valores: dict) -> int:
    """Peso del reporte SIN fotos (plantilla + datos)."""
    excel, _ = _armar_excel(plantilla_bytes, valores, [], [])
    return len(excel)


def extraer_valores(clasificacion: dict) -> dict:
    """Lee los PDF y devuelve los valores sugeridos para cada celda."""
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
    return valores


def clave_foto(archivo) -> str:
    """Identificador de una foto subida (nombre + tamaño), para guardar sus
    ediciones aunque Streamlit vuelva a ejecutar la página."""
    return f"{archivo.name}|{len(archivo.getvalue())}"


def generar_reporte(valores: dict, imagenes: list, ediciones: dict = None) -> tuple:
    """Arma el Excel con los valores YA REVISADOS por el usuario y las fotos
    (con sus ediciones de rotación/recorte, si las hay).
    Devuelve (bytes_excel, faltantes, n_imagenes_insertadas, info_peso).

    El reporte resultante nunca supera LIMITE_REPORTE_BYTES (25 MB)."""
    faltantes = validar_campos_criticos(valores)

    plantilla_bytes = cargar_plantilla_predeterminada()
    if not plantilla_bytes:
        raise RuntimeError(
            f"No se encontró la plantilla '{TEMPLATE_FILENAME}' junto a la app. "
            f"Agrega el archivo '{TEMPLATE_FILENAME}' al repositorio."
        )

    ediciones = ediciones or {}
    fotos = [(a.name, a.getvalue(), ediciones.get(clave_foto(a))) for a in imagenes]
    nombres = [f[0] for f in fotos]

    # Presupuesto para fotos = objetivo - peso de la plantilla con datos
    # - un pequeño margen por la estructura XML/relaciones de cada imagen.
    peso_base = _peso_base_excel(plantilla_bytes, valores)
    overhead_fotos = 4 * 1024 * len(fotos)
    presupuesto = max(OBJETIVO_REPORTE_BYTES - peso_base - overhead_fotos, 1 * MB)

    imagenes_jpeg, info = comprimir_fotos_para_limite(fotos, presupuesto)
    excel_bytes, n_insertadas = _armar_excel(plantilla_bytes, valores, imagenes_jpeg, nombres)

    # Verificación final con el tamaño REAL del archivo. Si por algún motivo
    # se pasó, se reduce el presupuesto y se vuelve a comprimir.
    intentos = 0
    while len(excel_bytes) > LIMITE_REPORTE_BYTES and intentos < 3 and fotos:
        exceso = len(excel_bytes) - OBJETIVO_REPORTE_BYTES
        presupuesto = max(presupuesto - exceso - 512 * 1024, 1 * MB)
        imagenes_jpeg, info = comprimir_fotos_para_limite(fotos, presupuesto)
        excel_bytes, n_insertadas = _armar_excel(plantilla_bytes, valores, imagenes_jpeg, nombres)
        intentos += 1

    info.update({
        "peso_final_bytes": len(excel_bytes),
        "peso_original_fotos_bytes": sum(len(f[1]) for f in fotos),
        "bajo_limite": len(excel_bytes) <= LIMITE_REPORTE_BYTES,
        "n_fotos": len(fotos),
    })
    return excel_bytes, faltantes, n_insertadas, info


# --------------------------------------------------------------------------
# Interfaz Streamlit — colores institucionales Astara
#   Azul marino  #002060   ·   Amarillo  #FFC000
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
    --astara-navy: #002060;
    --astara-navy-dark: #001640;
    --astara-bg: #0D1A36;
    --astara-surface: #142449;
    --astara-border: #2A3F70;
    --astara-gold: #FFC000;
    --astara-gold-dark: #D9A300;
    --astara-text: #E6ECF7;
}

/* Fondo general oscuro (azul noche), con una franja azul institucional arriba */
.stApp {
    background:
        linear-gradient(180deg, var(--astara-navy) 0px, var(--astara-navy) 230px, var(--astara-bg) 230px, var(--astara-bg) 100%);
    color: var(--astara-text);
}
[data-testid="stHeader"] { background: transparent; }
.stApp p, .stApp li, .stApp label, .stApp span, .stApp h1, .stApp h2,
.stApp h3, .stApp h4, .stApp h5, [data-testid="stMarkdownContainer"] {
    color: var(--astara-text);
}

/* Barra lateral azul marino con texto claro y acentos amarillos */
[data-testid="stSidebar"] {
    background: var(--astara-navy);
    border-right: 4px solid var(--astara-gold);
}
[data-testid="stSidebar"] * { color: #FFFFFF !important; }
[data-testid="stSidebar"] h3 { color: var(--astara-gold) !important; }
[data-testid="stSidebar"] code {
    background: rgba(255, 192, 0, 0.15) !important;
    color: var(--astara-gold) !important;
}
[data-testid="stSidebar"] hr { border-color: rgba(255, 192, 0, 0.4); }

/* Encabezado: barra amarilla con texto azul (igual que los reportes) */
.astara-hero {
    background: var(--astara-gold);
    padding: 1.4rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.4rem;
    border-bottom: 6px solid var(--astara-gold-dark);
    box-shadow: 0 10px 28px rgba(0, 0, 0, 0.25);
}
.astara-hero h1 {
    color: var(--astara-navy) !important;
    font-size: 2rem;
    font-weight: 800;
    margin: 0;
    letter-spacing: 0.3px;
}
.astara-hero p {
    color: var(--astara-navy) !important;
    margin: 0.35rem 0 0 0;
    font-size: 0.97rem;
    opacity: 0.9;
}

/* Etiquetas de archivos detectados */
.astara-badge {
    display: inline-block;
    background: var(--astara-surface);
    color: var(--astara-text) !important;
    border: 1.5px solid var(--astara-border);
    border-radius: 999px;
    padding: 3px 12px;
    font-size: 0.8rem;
    font-weight: 600;
    margin-right: 6px;
    margin-bottom: 6px;
}
.astara-badge.ok { background: var(--astara-navy); border-color: #4C6FB8; color: #FFFFFF !important; }
.astara-badge.warn { background: rgba(255, 192, 0, 0.10); border-color: var(--astara-gold); color: var(--astara-gold) !important; }
.astara-badge.missing { background: transparent; border-color: #4B5563; color: #9CA3AF !important; }

/* Títulos de sección con acento amarillo */
.stApp h4 {
    color: #FFFFFF !important;
    border-left: 5px solid var(--astara-gold);
    padding-left: 0.6rem;
}

/* Textos secundarios (captions) un poco más tenues */
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * { color: #A9B6D3 !important; }

/* Zona de carga */
div[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed var(--astara-gold) !important;
    background: var(--astara-surface) !important;
}
div[data-testid="stFileUploaderDropzone"] * { color: var(--astara-text) !important; }

/* Formulario de revisión */
[data-testid="stForm"] {
    background: var(--astara-surface);
    border: 1px solid var(--astara-border) !important;
    border-top: 4px solid var(--astara-gold) !important;
    border-radius: 12px;
}
[data-testid="stForm"] textarea, [data-testid="stForm"] input {
    background: #0B1530 !important;
    color: var(--astara-text) !important;
}
[data-testid="stForm"] label p { font-weight: 600; }

/* Botones primarios: amarillo con texto azul */
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"],
.stFormSubmitButton > button,
button[data-testid="stBaseButton-primaryFormSubmit"],
button[data-testid="stBaseButton-primary"] {
    background: var(--astara-gold) !important;
    color: var(--astara-navy) !important;
    border: 2px solid var(--astara-gold-dark) !important;
    font-weight: 700 !important;
}
.stButton > button[kind="primary"] *,
.stDownloadButton > button[kind="primary"] *,
.stFormSubmitButton > button *,
button[data-testid="stBaseButton-primaryFormSubmit"] *,
button[data-testid="stBaseButton-primary"] * { color: var(--astara-navy) !important; }
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover,
.stFormSubmitButton > button:hover,
button[data-testid="stBaseButton-primaryFormSubmit"]:hover,
button[data-testid="stBaseButton-primary"]:hover {
    background: var(--astara-navy) !important;
    border-color: var(--astara-gold) !important;
}
.stButton > button[kind="primary"]:hover *,
.stDownloadButton > button[kind="primary"]:hover *,
.stFormSubmitButton > button:hover *,
button[data-testid="stBaseButton-primaryFormSubmit"]:hover *,
button[data-testid="stBaseButton-primary"]:hover * { color: var(--astara-gold) !important; }

/* Indicadores de peso */
[data-testid="stMetric"] {
    background: var(--astara-surface);
    border: 1px solid var(--astara-border);
    border-top: 4px solid var(--astara-gold);
    border-radius: 10px;
    padding: 0.7rem 1rem;
}
[data-testid="stMetricValue"], [data-testid="stMetricValue"] * { color: var(--astara-gold) !important; }
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

PLANTILLA_PREDETERMINADA_DISPONIBLE = cargar_plantilla_predeterminada() is not None

with st.sidebar:
    st.markdown("### ℹ️ Cómo usar la app")
    st.markdown(
        """
        1. Arrastra los archivos al campo de carga (ninguno es
           obligatorio, usa los que tengas a mano):
           - `Orden de Trabajo.pdf`
           - `Trabajo en Garantías.pdf`
           - Fotos del vehículo
        2. Los PDF **deben llamarse exactamente así**. Cualquier otro
           PDF, Excel, Word, etc. se ignora automáticamente.
        3. La plantilla del **Reporte Técnico Astara ya viene integrada**.
        4. Revisa el formulario: viene pre-llenado con lo leído de los PDF.
           Completa la **queja**, el **diagnóstico** y el **trabajo
           realizado** sin límite de caracteres.
        5. Si alguna foto quedó chueca o con sobrante, rótala (↺ ↻) o
           recórtala (✂️) en la sección de fotos.
        6. Genera el informe y descárgalo listo.
        """
    )
    st.caption("Formatos de foto aceptados: " + ", ".join(sorted(EXTENSIONES_IMAGEN)).upper())
    st.divider()
    st.markdown("### 📦 Peso del reporte")
    st.markdown(
        f"""
        El reporte **nunca supera {LIMITE_REPORTE_BYTES // MB} MB**.
        - Las fotos se ordenan **por nombre** (1, 2, 3… y luego el resto).
        - Las **{N_FOTOS_PRIORITARIAS} primeras** se guardan siempre en alta calidad:
          nombra **"1"** y **"2"** las más importantes (ej. VIN y tablero).
        - Las demás solo se comprimen si es necesario para no pasar el límite.
        """
    )
    if not PLANTILLA_PREDETERMINADA_DISPONIBLE:
        st.error(f"No se encontró '{TEMPLATE_FILENAME}' junto a la app. Agrégalo al repositorio.")

# Sin filtro de tipo en el cargador: así el usuario puede arrastrar una
# carpeta completa sin errores, y la app descarta lo que no corresponde.
archivos_subidos = st.file_uploader(
    "Adjunta aquí: 'Orden de Trabajo.pdf', 'Trabajo en Garantías.pdf' y/o fotos (todo opcional; lo demás se ignora)",
    accept_multiple_files=True,
)

clasificacion = clasificar_archivos_subidos(archivos_subidos or [])

st.markdown("#### 📎 Archivos detectados")
badges = []
if PLANTILLA_PREDETERMINADA_DISPONIBLE:
    badges.append('<span class="astara-badge ok">Plantilla: ✔ integrada</span>')
else:
    badges.append('<span class="astara-badge missing">Plantilla: ✘ no disponible</span>')
badges.append(
    f'<span class="astara-badge {"ok" if clasificacion["pdf_orden"] else "warn"}">'
    f'Orden de Trabajo: {clasificacion["pdf_orden_nombre"] or "no adjuntada"}</span>'
)
badges.append(
    f'<span class="astara-badge {"ok" if clasificacion["pdf_garantias"] else "warn"}">'
    f'Trabajo en Garantías: {clasificacion["pdf_garantias_nombre"] or "no adjuntada"}</span>'
)
badges.append(
    f'<span class="astara-badge {"ok" if clasificacion["imagenes"] else "warn"}">'
    f'Fotos: {len(clasificacion["imagenes"])}</span>'
)
st.markdown(" ".join(badges), unsafe_allow_html=True)
st.caption("Nada es obligatorio: lo que no adjuntes simplemente queda vacío en el reporte para completarlo a mano.")

if clasificacion["ignorados"]:
    n_ign = len(clasificacion["ignorados"])
    with st.expander(f"🚫 {n_ign} archivo(s) ignorado(s) por no corresponder (clic para ver)"):
        st.caption(
            "Solo se usan los PDF llamados exactamente 'Orden de Trabajo' y 'Trabajo en Garantías', "
            "y las fotos. Si un PDF que necesitas aparece aquí, cámbiale el nombre y vuelve a subirlo."
        )
        st.markdown("\n".join(f"- {n}" for n in clasificacion["ignorados"]))

# --------------------------------------------------------------------------
# Fotos: vista previa, rotar y recortar antes de generar
# --------------------------------------------------------------------------
if "ediciones" not in st.session_state:
    st.session_state["ediciones"] = {}   # clave_foto -> {"rot": int, "crop": tuple|None}
EDICIONES = st.session_state["ediciones"]
PROPORCION_RECUADRO = (TARGET_WIDTH, TARGET_HEIGHT)  # forma del recuadro de la foto en el informe


@st.cache_data(show_spinner=False, max_entries=300)
def _miniatura_jpeg(datos: bytes, rot: int, crop, lado: int) -> bytes:
    edicion = {"rot": rot, "crop": crop}
    img = _abrir_normalizada(datos, lado, edicion)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=80)
    return buffer.getvalue()


def _edicion_de(clave: str) -> dict:
    return EDICIONES.get(clave, {"rot": 0, "crop": None})


def rotar_foto(clave: str, grados: int) -> None:
    ed = dict(_edicion_de(clave))
    ed["rot"] = (ed.get("rot", 0) + grados) % 360
    ed["crop"] = None  # el recorte se definió sobre la orientación anterior
    ed["ajuste"] = None
    EDICIONES[clave] = ed


def abrir_recorte(clave: str) -> None:
    st.session_state["foto_en_recorte"] = clave


def cerrar_recorte() -> None:
    st.session_state.pop("foto_en_recorte", None)


def guardar_recorte(clave: str, crop, ajuste: dict = None) -> None:
    ed = dict(_edicion_de(clave))
    ed["crop"] = crop
    ed["ajuste"] = ajuste   # parámetros usados, para reabrir el editor igual
    EDICIONES[clave] = ed
    cerrar_recorte()


def restablecer_foto(clave: str) -> None:
    EDICIONES.pop(clave, None)
    cerrar_recorte()


@st.cache_data(show_spinner=False, max_entries=300)
def _base_editor(datos: bytes, rot: int) -> PILImage.Image:
    """Versión liviana (máx. 700 px) de la foto, rotada y sin recortar, para el
    editor y para calcular encuadres. Se guarda en caché: el editor responde rápido."""
    return _abrir_normalizada(datos, 700, {"rot": rot})


def calcular_encuadre(w: int, h: int, zoom: float = 1.0, px: float = 0.5, py: float = 0.5) -> tuple:
    """Recorte con la MISMA forma que el recuadro de la foto en el informe
    (así no sale estirada). zoom 1 = el recorte más grande que cabe en la foto;
    px/py = posición 0-1 (izquierda→derecha, arriba→abajo). Nunca se sale de la foto."""
    proporcion = PROPORCION_RECUADRO[0] / PROPORCION_RECUADRO[1]
    if w / h > proporcion:
        alto, ancho = h, h * proporcion
    else:
        ancho, alto = w, w / proporcion
    ancho, alto = ancho / zoom, alto / zoom
    izq, arr = px * (w - ancho), py * (h - alto)
    return (izq / w, arr / h, (izq + ancho) / w, (arr + alto) / h)


def _dibujar_marco(img: PILImage.Image, crop) -> PILImage.Image:
    """Oscurece lo que queda fuera del recorte y dibuja un marco amarillo."""
    w, h = img.size
    caja = (crop[0] * w, crop[1] * h, crop[2] * w - 1, crop[3] * h - 1)
    mascara = PILImage.new("L", img.size, 150)
    ImageDraw.Draw(mascara).rectangle(caja, fill=0)
    salida = PILImage.composite(PILImage.new("RGB", img.size, (0, 0, 0)), img, mascara)
    ImageDraw.Draw(salida).rectangle(caja, outline=(255, 192, 0), width=4)
    return salida


def encuadrar_todas(archivos: list) -> None:
    """Aplica a todas las fotos SIN recorte propio un encuadre centrado con la
    forma del recuadro del informe. Las que ya se recortaron a mano no se tocan."""
    for archivo in archivos:
        clave = clave_foto(archivo)
        ed = dict(_edicion_de(clave))
        if ed.get("crop"):
            continue
        try:
            w, h = _base_editor(archivo.getvalue(), ed.get("rot", 0)).size
        except Exception:
            continue
        ed["crop"] = calcular_encuadre(w, h)
        ed["ajuste"] = {"modo": "encuadre", "zoom": 1.0, "px": 0.5, "py": 0.5}
        EDICIONES[clave] = ed


def quitar_recortes_todas(archivos: list) -> None:
    for archivo in archivos:
        clave = clave_foto(archivo)
        if clave in EDICIONES:
            EDICIONES[clave] = {**EDICIONES[clave], "crop": None, "ajuste": None}


# ---- Recortador interactivo propio (sin librerías externas) --------------
# Componente "bidireccional" sin compilación: un index.html que habla con
# Streamlit por postMessage. El recuadro se mueve en el navegador (fluido) y
# solo avisa a la app al SOLTAR el mouse; nunca se sale de la foto.
_HTML_RECORTADOR = r"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  html, body { margin: 0; padding: 0; background: transparent; overflow: hidden;
               user-select: none; -webkit-user-select: none; }
  #zona { position: relative; margin: 0 auto; overflow: hidden; touch-action: none; }
  #zona img { display: block; width: 100%; height: 100%; pointer-events: none; }
  #caja { position: absolute; border: 2px solid #FFC000; box-sizing: border-box; cursor: move;
          box-shadow: 0 0 0 9999px rgba(0, 0, 0, .55); touch-action: none; }
  #caja::before, #caja::after { content: ""; position: absolute; border: 0 dashed rgba(255,192,0,.5); }
  #caja::before { left: 33.3%; right: 33.3%; top: 0; bottom: 0; border-left-width: 1px; border-right-width: 1px; }
  #caja::after  { top: 33.3%; bottom: 33.3%; left: 0; right: 0; border-top-width: 1px; border-bottom-width: 1px; }
  .esq { position: absolute; width: 18px; height: 18px; background: #FFC000; border: 2px solid #002060;
         border-radius: 3px; box-sizing: border-box; touch-action: none; z-index: 2; }
  .nw { left: -9px; top: -9px; cursor: nwse-resize; }   .ne { right: -9px; top: -9px; cursor: nesw-resize; }
  .sw { left: -9px; bottom: -9px; cursor: nesw-resize; } .se { right: -9px; bottom: -9px; cursor: nwse-resize; }
</style></head>
<body>
<div id="zona"><img id="img" alt=""><div id="caja">
  <div class="esq nw" data-c="nw"></div><div class="esq ne" data-c="ne"></div>
  <div class="esq sw" data-c="sw"></div><div class="esq se" data-c="se"></div>
</div></div>
<script>
  // Recortador propio, sin librerías externas.
  // Estado en píxeles de pantalla: x, y, w, h del recuadro dentro de la foto (W x H).
  function enviar(tipo, extra) {
    window.parent.postMessage(Object.assign({isStreamlitMessage: true, type: tipo}, extra || {}), "*");
  }
  var zona = document.getElementById("zona"), img = document.getElementById("img"),
      caja = document.getElementById("caja");
  var W = 0, H = 0, aspecto = null, ultimaClave = null, MIN = 24;
  var r = {x: 0, y: 0, w: 0, h: 0}, arrastre = null;

  function pintar() {
    caja.style.left = r.x + "px"; caja.style.top = r.y + "px";
    caja.style.width = r.w + "px"; caja.style.height = r.h + "px";
  }
  function enviarValor() {
    enviar("streamlit:setComponentValue", {dataType: "json", value: {
      l: r.x / W, t: r.y / H, r: (r.x + r.w) / W, b: (r.y + r.h) / H}});
  }
  function ajustarAspecto() {          // fuerza la forma y lo deja dentro de la foto
    if (aspecto) {
      if (r.w / r.h > aspecto) r.w = r.h * aspecto; else r.h = r.w / aspecto;
    }
    r.w = Math.min(r.w, W); r.h = Math.min(r.h, H);
    if (aspecto) { if (r.w / r.h > aspecto) r.w = r.h * aspecto; else r.h = r.w / aspecto; }
    r.x = Math.min(Math.max(r.x, 0), W - r.w); r.y = Math.min(Math.max(r.y, 0), H - r.h);
  }
  function punto(e) {
    var b = zona.getBoundingClientRect();
    return {x: e.clientX - b.left, y: e.clientY - b.top};
  }
  function iniciar(e) {
    e.preventDefault(); e.stopPropagation();
    var esquina = e.target.getAttribute("data-c");
    var p = punto(e);
    arrastre = {esquina: esquina, p0: p, r0: {x: r.x, y: r.y, w: r.w, h: r.h}};
    if (esquina) {                     // la esquina opuesta queda fija
      arrastre.ax = esquina.indexOf("w") >= 0 ? r.x + r.w : r.x;
      arrastre.ay = esquina.indexOf("n") >= 0 ? r.y + r.h : r.y;
      arrastre.izq = esquina.indexOf("w") >= 0; arrastre.arr = esquina.indexOf("n") >= 0;
    }
    try { e.target.setPointerCapture(e.pointerId); } catch (err) {}
  }
  function mover(e) {
    if (!arrastre) return;
    e.preventDefault();
    var p = punto(e), a = arrastre;
    if (!a.esquina) {                  // mover el recuadro completo
      r.x = Math.min(Math.max(a.r0.x + p.x - a.p0.x, 0), W - r.w);
      r.y = Math.min(Math.max(a.r0.y + p.y - a.p0.y, 0), H - r.h);
    } else {                           // redimensionar desde una esquina
      var maxW = a.izq ? a.ax : W - a.ax, maxH = a.arr ? a.ay : H - a.ay;
      var nw = Math.abs(p.x - a.ax), nh = Math.abs(p.y - a.ay);
      if (aspecto) {
        nw = Math.max(nw, nh * aspecto);
        nw = Math.min(nw, maxW, maxH * aspecto);
        nw = Math.max(nw, MIN * aspecto); nh = nw / aspecto;
      } else {
        nw = Math.max(MIN, Math.min(nw, maxW)); nh = Math.max(MIN, Math.min(nh, maxH));
      }
      r.w = nw; r.h = nh;
      r.x = a.izq ? a.ax - nw : a.ax; r.y = a.arr ? a.ay - nh : a.ay;
    }
    pintar();
  }
  function soltar() { if (arrastre) { arrastre = null; enviarValor(); } }

  caja.addEventListener("pointerdown", iniciar);
  window.addEventListener("pointermove", mover);
  window.addEventListener("pointerup", soltar);
  window.addEventListener("pointercancel", soltar);

  function render(args) {
    if (args.clave === ultimaClave) return;   // no reiniciar el recuadro en cada actualización
    ultimaClave = args.clave;
    aspecto = args.aspect || null;
    var disponible = document.body.clientWidth || 700;
    W = Math.min(disponible, args.alto_max * args.w / args.h);
    H = W * args.h / args.w;
    zona.style.width = W + "px"; zona.style.height = H + "px";
    enviar("streamlit:setFrameHeight", {height: Math.ceil(H) + 4});
    var c = args.inicial;
    if (c) {
      r = {x: c[0] * W, y: c[1] * H, w: (c[2] - c[0]) * W, h: (c[3] - c[1]) * H};
    } else {                           // por defecto: 90 % del máximo, centrado
      r = {w: W, h: H, x: 0, y: 0};
      if (aspecto) { if (W / H > aspecto) r.w = H * aspecto; else r.h = W / aspecto; }
      r.w *= 0.9; r.h *= 0.9; r.x = (W - r.w) / 2; r.y = (H - r.h) / 2;
    }
    ajustarAspecto();
    img.src = args.src;
    pintar();
    enviarValor();
  }
  window.addEventListener("message", function (e) {
    if (e.data && e.data.type === "streamlit:render") render(e.data.args);
  });
  enviar("streamlit:componentReady", {apiVersion: 1});
</script>
</body></html>
"""


@st.cache_resource(show_spinner=False)
def _declarar_recortador():
    """Escribe el index.html del componente en una carpeta temporal y lo
    registra en Streamlit (una sola vez por servidor)."""
    try:
        import streamlit.components.v1 as components
        carpeta = Path(tempfile.gettempdir()) / "astara_recortador_v1"
        carpeta.mkdir(parents=True, exist_ok=True)
        (carpeta / "index.html").write_text(_HTML_RECORTADOR, encoding="utf-8")
        return components.declare_component("astara_recortador", path=str(carpeta))
    except Exception as e:
        logger.error(f"No se pudo preparar el recortador interactivo: {e}")
        return None


@st.cache_data(show_spinner=False, max_entries=100)
def _data_url_editor(datos: bytes, rot: int) -> str:
    img = _base_editor(datos, rot)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode()


def _editor_recorte(archivo) -> None:
    """Editor de recorte de UNA foto (en ventana emergente): recuadro que se
    agranda/achica desde las esquinas y se mueve arrastrándolo, manteniendo
    la forma del recuadro del informe (salvo que se elija 'Proporción libre')."""
    clave = clave_foto(archivo)
    ed = _edicion_de(clave)
    rot = ed.get("rot", 0)
    base = _base_editor(archivo.getvalue(), rot)
    w, h = base.size

    libre = st.checkbox(
        "Proporción libre", value=(ed.get("ajuste") or {}).get("modo") == "libre_arrastre",
        key=f"libre_{clave}",
        help="Por defecto el recuadro mantiene la forma de la foto en el informe, "
             "así no sale estirada. Marca esto para recortar con cualquier forma.",
    )
    st.caption("Arrastra las **esquinas** del recuadro amarillo para agrandarlo o achicarlo y "
               "arrástralo desde el **centro** para moverlo. Luego pulsa *Aplicar recorte*.")

    recortador = _declarar_recortador()
    crop = None
    if recortador is not None:
        valor = recortador(
            src=_data_url_editor(archivo.getvalue(), rot), w=w, h=h,
            aspect=None if libre else PROPORCION_RECUADRO[0] / PROPORCION_RECUADRO[1],
            inicial=list(ed["crop"]) if ed.get("crop") else None,
            clave=f"{clave}|{rot}|{libre}", alto_max=520,
            key=f"recortador_{clave}_{rot}_{libre}", default=None,
        )
        if isinstance(valor, dict) and all(k in valor for k in ("l", "t", "r", "b")):
            crop = (valor["l"], valor["t"], valor["r"], valor["b"])
    modo_guardado = "libre_arrastre" if libre else "arrastre"

    # Plan B por si el recortador no se muestra en algún navegador
    with st.expander("Alternativa: ajustar con deslizadores"):
        usar_sl = st.checkbox("Usar deslizadores en vez del recuadro", key=f"usar_sl_{clave}")
        zoom = st.slider("🔍 Acercar", 100, 300, 100, step=5, format="%d%%", key=f"zoom_{clave}")
        px = st.slider("↔️ Mover (izquierda – derecha)", 0, 100, 50, key=f"px_{clave}")
        py = st.slider("↕️ Mover (arriba – abajo)", 0, 100, 50, key=f"py_{clave}")
        if usar_sl or crop is None:
            crop = calcular_encuadre(w, h, zoom / 100, px / 100, py / 100)
            modo_guardado = "encuadre"
            st.image(_dibujar_marco(base, crop), width=420)

    if crop:
        c1, _ = st.columns([1, 2])
        c1.caption("Así quedará en el informe:")
        c1.image(aplicar_edicion(base, {"crop": crop}), use_container_width=True)

    a1, a2, a3 = st.columns(3)
    if a1.button("✅ Aplicar recorte", type="primary", use_container_width=True,
                 key=f"aplicar_{clave}", disabled=crop is None):
        guardar_recorte(clave, crop, {"modo": modo_guardado})
        st.rerun()
    if a2.button("↩️ Deshacer cambios", use_container_width=True, key=f"deshacer_{clave}",
                 help="Quita la rotación y el recorte de esta foto"):
        restablecer_foto(clave)
        st.rerun()
    if a3.button("Cancelar", use_container_width=True, key=f"cancelar_{clave}"):
        cerrar_recorte()
        st.rerun()


# Ventana emergente si la versión de Streamlit la tiene (1.34+); si no, el
# editor se muestra arriba de las fotos.
_dialogo_api = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)
if _dialogo_api is not None:
    try:
        _dialogo_recorte = _dialogo_api("✂️ Recortar foto", width="large")(_editor_recorte)
    except TypeError:
        _dialogo_recorte = _dialogo_api("✂️ Recortar foto")(_editor_recorte)
else:
    _dialogo_recorte = None


if clasificacion["imagenes"]:
    st.markdown("#### 🖼️ Fotos del informe")
    st.caption(
        "Así quedarán en el informe y en este orden. ⭐ = alta calidad. "
        "Usa ↺ ↻ para rotar y ✂️ para recortar; los cambios se aplican al generar el informe."
    )
    g1, g2, _ = st.columns([1.4, 1.4, 3])
    g1.button("📐 Encuadrar todas", on_click=encuadrar_todas, args=(clasificacion["imagenes"],),
              use_container_width=True,
              help="Recorta de una vez todas las fotos a la forma del recuadro del informe (centradas), "
                   "para que no salgan estiradas. Las que ya recortaste a mano no se modifican. "
                   "Después puedes ajustar cualquiera con ✂️.")
    g2.button("🧹 Quitar todos los recortes", on_click=quitar_recortes_todas,
              args=(clasificacion["imagenes"],), use_container_width=True)

    # ---- Editor de recorte (una foto a la vez) ----
    clave_abierta = st.session_state.get("foto_en_recorte")
    archivo_en_recorte = next(
        (a for a in clasificacion["imagenes"] if clave_foto(a) == clave_abierta), None
    )
    if archivo_en_recorte is not None:
        if _dialogo_recorte is not None:
            # La ventana se mantiene abierta sola mientras se usa; se limpia la
            # marca para que no reaparezca si el usuario la cierra con la X.
            st.session_state.pop("foto_en_recorte", None)
            _dialogo_recorte(archivo_en_recorte)
        else:
            with st.container(border=True):
                st.markdown("##### ✂️ Recortar foto")
                _editor_recorte(archivo_en_recorte)

    columnas = st.columns(4)
    for i, archivo in enumerate(clasificacion["imagenes"]):
        clave = clave_foto(archivo)
        ed = _edicion_de(clave)
        marcas = []
        if ed.get("rot"):
            marcas.append(f"rotada {ed['rot']}°")
        if ed.get("crop"):
            marcas.append("recortada")
        estrella = "⭐ " if i < N_FOTOS_PRIORITARIAS else ""
        with columnas[i % 4]:
            try:
                st.image(_miniatura_jpeg(archivo.getvalue(), ed.get("rot", 0), ed.get("crop"), 480),
                         use_container_width=True)
            except Exception:
                st.warning("No se pudo mostrar esta foto.")
            st.caption(f"{estrella}{i + 1}. {calcular_celda_imagen(i)} · {archivo.name}"
                       + (f" · ✏️ {', '.join(marcas)}" if marcas else ""))
            b1, b2, b3 = st.columns(3)
            b1.button("↺", key=f"rotizq_{i}", on_click=rotar_foto, args=(clave, -90),
                      help="Rotar 90° a la izquierda", use_container_width=True)
            b2.button("↻", key=f"rotder_{i}", on_click=rotar_foto, args=(clave, 90),
                      help="Rotar 90° a la derecha", use_container_width=True)
            b3.button("✂️", key=f"recortar_{i}", on_click=abrir_recorte, args=(clave,),
                      help="Recortar", use_container_width=True)

# --------------------------------------------------------------------------
# Paso 2: revisar / completar los textos antes de generar
# --------------------------------------------------------------------------
valores_pdf = extraer_valores(clasificacion)
CAMPOS_DATOS_GENERALES = ("vin", "fecha_falla", "fecha_cierre_ot", "km",
                          "concesionario", "asesor_responsable", "tecnico")


def cargar_en_formulario(vals: dict) -> None:
    """Rellena los campos del formulario con lo leído de los PDF."""
    for campo in CELDAS_EXCEL:
        st.session_state[f"campo_{campo}"] = vals.get(campo, "") or ""


# Solo se recargan los campos cuando cambian los PDF; así lo que el usuario
# escribió no se pierde al agregar o quitar fotos.
firma_pdfs = hash((clasificacion["pdf_orden"] or "", clasificacion["pdf_garantias"] or ""))
if st.session_state.get("firma_pdfs") != firma_pdfs:
    cargar_en_formulario(valores_pdf)
    st.session_state["firma_pdfs"] = firma_pdfs

col_tit, col_btn = st.columns([4, 1])
with col_tit:
    st.markdown("#### ✍️ Revisa y completa el informe")
with col_btn:
    st.button("↺ Volver a lo leído de los PDF", on_click=cargar_en_formulario,
              args=(valores_pdf,), use_container_width=True)
st.caption(
    "Los campos vienen pre-llenados con lo que se pudo leer de los PDF. "
    "Completa o corrige aquí lo que haga falta (sin límite de caracteres) y luego genera el informe."
)

with st.form("form_reporte", border=True):
    st.text_area(
        "1) ¿Cuál fue la queja del cliente?",
        key="campo_queja_cliente", height=110,
        placeholder="Ej.: Cliente ingresa con el motor bloqueado, este no gira.",
    )

    st.text_area(
        "2) ¿Cuál es el resultado de su investigación? (Diagnóstico técnico)",
        key="campo_diagnostico_tecnico", height=180,
        placeholder="Ej.: Se inspecciona visualmente el motor y presenta una biela desprendida por un costado del block...",
    )
    st.text_area("3) Repuestos utilizados", key="campo_repuestos_usados", height=140)
    st.text_area(
        "4) ¿Cómo se reparó / qué se realizó para solucionar el problema? (Trabajo técnico realizado)",
        key="campo_trabajo_realizado", height=140,
        placeholder="Ej.: Se reemplaza ensamble motor y válvula EGR, se realizan pruebas de funcionamiento...",
    )

    with st.expander("🚗 Datos del vehículo y la OT (VIN, fechas, km, concesionario, asesor, técnico)"):
        cols = st.columns(3)
        for i, campo in enumerate(CAMPOS_DATOS_GENERALES):
            cols[i % 3].text_input(ETIQUETAS_CAMPOS[campo], key=f"campo_{campo}")

    generar = st.form_submit_button("⚙️ Generar Reporte Técnico Astara", type="primary",
                                    use_container_width=True)

if generar:
    valores = {campo: (st.session_state.get(f"campo_{campo}", "") or "").strip() for campo in CELDAS_EXCEL}
    faltantes = validar_campos_criticos(valores)

    n_fotos = len(clasificacion["imagenes"])
    mensaje_spinner = "Armando el reporte..."
    if n_fotos > 10:
        mensaje_spinner = f"Armando el reporte y optimizando {n_fotos} fotos para no superar {LIMITE_REPORTE_BYTES // MB} MB..."
    with st.spinner(mensaje_spinner):
        try:
            excel_bytes, _, n_insertadas, info_peso = generar_reporte(valores, clasificacion["imagenes"], EDICIONES)
        except Exception as e:
            st.error(f"No se pudo generar el reporte: {e}")
            logger.error(f"Error generando el reporte: {e}")
            excel_bytes = None

    if excel_bytes:
        if faltantes:
            etiquetas_faltantes = ", ".join(ETIQUETAS_CAMPOS.get(c, c) for c in faltantes)
            st.warning(f"⚠ Estos campos importantes quedaron vacíos; puedes completarlos arriba y volver a generar: {etiquetas_faltantes}")
        else:
            st.success("✅ Informe generado con todos los campos importantes completos.")

        # ---- Peso del reporte ----
        st.markdown("#### 📦 Peso del reporte")
        c1, c2, c3 = st.columns(3)
        c1.metric("Peso final del Excel", f"{info_peso['peso_final_bytes'] / MB:.2f} MB",
                  help=f"Límite: {LIMITE_REPORTE_BYTES // MB} MB")
        c2.metric("Peso original de las fotos", f"{info_peso['peso_original_fotos_bytes'] / MB:.2f} MB")
        if info_peso["n_fotos"] == 0:
            estado = "Sin fotos"
        elif info_peso["nivel_prioritarias"] == 0 and not info_peso["nivel_secundarias"]:
            estado = "Sin compresión extra"
        elif info_peso["nivel_prioritarias"] == 0:
            estado = f"Fotos 3+ comprimidas (nivel {info_peso['nivel_secundarias']})"
        else:
            estado = "Todas comprimidas"
        c3.metric("Optimización aplicada", estado)

        if not info_peso["bajo_limite"]:
            st.error(
                f"⚠ Aun con la compresión máxima el reporte pesa {info_peso['peso_final_bytes'] / MB:.2f} MB "
                f"(> {LIMITE_REPORTE_BYTES // MB} MB). Reduce la cantidad de fotos."
            )
        elif info_peso["nivel_prioritarias"] > 0:
            st.info("Por la gran cantidad de fotos también se redujeron levemente las 2 primeras para respetar el límite.")

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
            n_editadas = sum(1 for a in clasificacion["imagenes"] if EDICIONES.get(clave_foto(a)))
            st.caption(f"🖼️ Fotos insertadas: {n_insertadas}/{len(clasificacion['imagenes'])}"
                       + (f" · {n_editadas} con rotación/recorte aplicado" if n_editadas else ""))
        else:
            st.info("No se adjuntaron fotos; el reporte queda sin imágenes insertadas.")

        st.markdown("#### ⬇️ Descargar reporte final")
        st.download_button(
            label="Descargar 'Reporte de falla Astara.xlsx'",
            data=excel_bytes,
            file_name="Reporte de falla Astara.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
