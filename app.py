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
import logging
from pathlib import Path

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

EXTENSIONES_IMAGEN = {"jpg", "jpeg", "png"}

# Plantilla predeterminada: se busca junto a este archivo (app.py) en el
# repositorio. Así no hace falta adjuntarla en cada uso; solo se necesita
# subir una plantilla distinta si en algún caso puntual se quiere usar otra.
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
    """Recibe la lista de UploadedFile del campo único (puede venir vacía)
    y los separa en: plantilla Excel (opcional; si no se sube ninguna se
    usa la plantilla predeterminada incluida en el repositorio), PDF Orden
    de Trabajo, PDF Garantías e imágenes, deduplicando imágenes por nombre
    (case-insensitive). Ningún archivo es obligatorio: lo que no se
    encuentre queda simplemente sin aportar datos."""
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
# Imágenes — preparación y compresión adaptativa
# --------------------------------------------------------------------------
def calcular_celda_imagen(indice: int) -> str:
    fila = IMG_START_ROW + (indice // len(IMG_COLUMNS)) * IMG_ROW_INCREMENT
    columna = IMG_COLUMNS[indice % len(IMG_COLUMNS)]
    return f"{columna}{fila}"


def _abrir_normalizada(datos: bytes, lado_max: int) -> PILImage.Image:
    """Abre la imagen y aplica la misma transformación que el script
    original (orientación EXIF, conversión a RGB, rotación si es vertical),
    reduciéndola a 'lado_max' px en su lado más largo si es más grande.

    Se usa draft() para que, en JPEG, el decodificador lea directamente a
    una escala menor cuando se va a reducir: ahorra mucha RAM y tiempo con
    fotos de celular de 12-50 MP (importante en Streamlit Cloud)."""
    img = PILImage.open(io.BytesIO(datos))
    if img.format == "JPEG":
        img.draft("RGB", (lado_max, lado_max))
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

    if max(img.size) > lado_max:
        img.thumbnail((lado_max, lado_max), PILImage.LANCZOS)
    return img


def codificar_imagen(datos: bytes, lado_max: int, calidad: int) -> bytes:
    """Devuelve la imagen normalizada como JPEG con el tamaño/calidad dados."""
    img = _abrir_normalizada(datos, lado_max)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=calidad, optimize=True, progressive=True)
    img.close()
    return buffer.getvalue()


def comprimir_fotos_para_limite(fotos: list, presupuesto_bytes: int) -> tuple:
    """Decide la compresión de cada foto para que la suma quepa en
    'presupuesto_bytes'.

    fotos: lista de (nombre, bytes_originales) en el orden del reporte.

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
            nombre, datos = fotos[i]
            try:
                resultado[i] = codificar_imagen(datos, lado, calidad)
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
def _armar_excel(plantilla_bytes: bytes, valores: dict, imagenes_jpeg: list, nombres: list) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(plantilla_bytes))
    ws = wb.active
    for campo, celda in CELDAS_EXCEL.items():
        ws[celda] = valores.get(campo, "")
    n_insertadas = insertar_imagenes(ws, imagenes_jpeg, nombres)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), n_insertadas


def _peso_base_excel(plantilla_bytes: bytes, valores: dict) -> int:
    """Peso del reporte SIN fotos (plantilla + datos)."""
    excel, _ = _armar_excel(plantilla_bytes, valores, [], [])
    return len(excel)


def generar_reporte(clasificacion: dict) -> tuple:
    """Ejecuta todo el flujo y devuelve (bytes_excel, valores, faltantes,
    n_imagenes_insertadas, info_peso). Ningún archivo es obligatorio: si
    falta la Orden de Trabajo, el de Garantías, fotos o incluso la plantilla
    (se usa la predeterminada del repositorio en ese caso), el reporte se
    genera igual dejando en blanco lo que no se pudo obtener.

    El reporte resultante nunca supera LIMITE_REPORTE_BYTES (25 MB)."""
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

    plantilla_bytes = clasificacion["excel_template"] or cargar_plantilla_predeterminada()
    if not plantilla_bytes:
        raise RuntimeError(
            f"No se encontró la plantilla predeterminada ('{TEMPLATE_FILENAME}') junto a la app "
            "y tampoco se adjuntó una plantilla .xlsx. Sube una plantilla o agrega el archivo "
            f"'{TEMPLATE_FILENAME}' al repositorio."
        )

    fotos = [(a.name, a.getvalue()) for a in clasificacion["imagenes"]]
    nombres = [n for n, _ in fotos]

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
        "peso_original_fotos_bytes": sum(len(d) for _, d in fotos),
        "bajo_limite": len(excel_bytes) <= LIMITE_REPORTE_BYTES,
        "n_fotos": len(fotos),
    })
    return excel_bytes, valores, faltantes, n_insertadas, info


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

/* Botones primarios: amarillo con texto azul */
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"],
button[data-testid="stBaseButton-primary"] {
    background: var(--astara-gold) !important;
    color: var(--astara-navy) !important;
    border: 2px solid var(--astara-gold-dark) !important;
    font-weight: 700 !important;
}
.stButton > button[kind="primary"] *,
.stDownloadButton > button[kind="primary"] *,
button[data-testid="stBaseButton-primary"] * { color: var(--astara-navy) !important; }
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover,
button[data-testid="stBaseButton-primary"]:hover {
    background: var(--astara-navy) !important;
    border-color: var(--astara-gold) !important;
}
.stButton > button[kind="primary"]:hover *,
.stDownloadButton > button[kind="primary"]:hover *,
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
        1. Arrastra los archivos que tengas al campo de carga (ninguno es
           obligatorio, usa los que tengas a mano):
           - `Orden de Trabajo.pdf`
           - `Trabajo en Garantías.pdf`
           - Fotos del vehículo (`.jpg`, `.jpeg`, `.png`)
        2. La plantilla del **Reporte Técnico Astara ya viene integrada**,
           no hace falta adjuntarla (solo si quieres usar una distinta a
           la de siempre).
        3. La app detecta sola qué es cada archivo que subas.
        4. Genera el reporte: lo que no se haya adjuntado queda en blanco
           para completar a mano.
        5. Revisa la vista previa y descarga el reporte final.
        """
    )
    st.divider()
    st.markdown("### 📦 Peso del reporte")
    st.markdown(
        f"""
        El reporte **nunca supera {LIMITE_REPORTE_BYTES // MB} MB**.
        - Las **{N_FOTOS_PRIORITARIAS} primeras fotos** se guardan siempre en alta calidad
          (súbelas primero si son las más importantes).
        - Las demás solo se comprimen si es necesario para no pasar el límite.
        """
    )
    st.divider()
    st.caption("No es necesario un campo por tipo de archivo: todo se sube junto y la app los diferencia automáticamente.")
    if not PLANTILLA_PREDETERMINADA_DISPONIBLE:
        st.error(f"No se encontró '{TEMPLATE_FILENAME}' junto a la app. Agrégalo al repositorio o súbelo manualmente en el campo de carga.")

archivos_subidos = st.file_uploader(
    "Adjunta aquí lo que tengas: Orden de Trabajo, Trabajo en Garantías y/o fotos (todo opcional)",
    type=["pdf", "xlsx", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

clasificacion = clasificar_archivos_subidos(archivos_subidos or [])

st.markdown("#### 📎 Archivos detectados")
badges = []
if clasificacion["excel_template"]:
    badges.append('<span class="astara-badge ok">Plantilla: ✔ personalizada cargada</span>')
elif PLANTILLA_PREDETERMINADA_DISPONIBLE:
    badges.append('<span class="astara-badge ok">Plantilla: ✔ predeterminada integrada</span>')
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

if clasificacion["sin_clasificar"]:
    st.warning("Archivos que no se pudieron clasificar y fueron ignorados: " + ", ".join(clasificacion["sin_clasificar"]))

generar = st.button("⚙️ Generar Reporte Técnico Astara", type="primary", use_container_width=True)

if generar:
    n_fotos = len(clasificacion["imagenes"])
    mensaje_spinner = "Extrayendo datos y armando el reporte..."
    if n_fotos > 10:
        mensaje_spinner = f"Extrayendo datos y optimizando {n_fotos} fotos para no superar {LIMITE_REPORTE_BYTES // MB} MB..."
    with st.spinner(mensaje_spinner):
        try:
            excel_bytes, valores, faltantes, n_insertadas, info_peso = generar_reporte(clasificacion)
        except Exception as e:
            st.error(f"No se pudo generar el reporte: {e}")
            logger.error(f"Error generando el reporte: {e}")
            excel_bytes = None

    if excel_bytes:
        if faltantes:
            etiquetas_faltantes = ", ".join(ETIQUETAS_CAMPOS.get(c, c) for c in faltantes)
            st.warning(f"⚠ Algunos campos críticos quedaron vacíos (por falta de archivos o de datos legibles) y conviene revisarlos a mano: {etiquetas_faltantes}")
        else:
            st.success("Todos los campos críticos se completaron correctamente.")

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
                f"(> {LIMITE_REPORTE_BYTES // MB} MB). Revisa si la plantilla adjunta es muy pesada o reduce la cantidad de fotos."
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
            st.markdown(f"#### 🖼️ Fotos insertadas en el reporte ({n_insertadas}/{len(clasificacion['imagenes'])})")
            columnas = st.columns(4)
            for i, archivo in enumerate(clasificacion["imagenes"]):
                etiqueta = "⭐ " if i < N_FOTOS_PRIORITARIAS else ""
                with columnas[i % 4]:
                    st.image(archivo.getvalue(), caption=f"{etiqueta}{calcular_celda_imagen(i)} · {archivo.name}", use_container_width=True)
            st.caption(f"⭐ = foto prioritaria (se conserva en alta calidad).")
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
