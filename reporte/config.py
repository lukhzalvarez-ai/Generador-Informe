"""Constantes y configuración compartida: celdas del Excel, mapas,
niveles de compresión, formatos aceptados y ubicación de la plantilla."""

import logging
from pathlib import Path

from PIL import Image as PILImage

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("reporte_astara")

# Evita el aviso de "decompression bomb" con fotos de cámaras de alta resolución
PILImage.MAX_IMAGE_PIXELS = 200_000_000

# Soporte opcional para fotos de iPhone (HEIC/HEIF). Requiere 'pillow-heif'
# en requirements.txt; si no está instalado, esas fotos simplemente se ignoran.
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIF_DISPONIBLE = True
except ImportError:
    HEIF_DISPONIBLE = False

# --------------------------------------------------------------------------
# Fotos en el Excel (idénticas al script original)
# --------------------------------------------------------------------------
CM_TO_PX_96DPI = 37.7952

TARGET_WIDTH = int(13.19 * CM_TO_PX_96DPI)

TARGET_HEIGHT = int(5.69 * CM_TO_PX_96DPI)

IMG_START_ROW = 28

IMG_ROW_INCREMENT = 11

IMG_COLUMNS = ("C", "G")

# Forma del recuadro de cada foto en el informe (ancho, alto)
PROPORCION_RECUADRO = (TARGET_WIDTH, TARGET_HEIGHT)

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

# --------------------------------------------------------------------------
# Celdas y textos
# --------------------------------------------------------------------------
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

CAMPOS_TEXTO_LARGO = ("queja_cliente", "diagnostico_tecnico", "trabajo_realizado")

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

# --------------------------------------------------------------------------
# Archivos aceptados
# --------------------------------------------------------------------------
# Formatos de foto aceptados. HEIC/HEIF (iPhone) solo si 'pillow-heif' está instalado.
EXTENSIONES_IMAGEN = {"jpg", "jpeg", "jfif", "png", "webp", "bmp", "gif", "tif", "tiff"}
if HEIF_DISPONIBLE:
    EXTENSIONES_IMAGEN |= {"heic", "heif"}

NOMBRE_PDF_ORDEN = "orden de trabajo"

NOMBRES_PDF_GARANTIAS = {"trabajo en garantias", "trabajo en garantia"}

# Archivos ZIP: se abren y se usa lo que traigan adentro (fotos y los 2 PDF).
# Límites de seguridad para que un ZIP dañado o malicioso no tumbe el servidor.
ZIP_MAX_DESCOMPRIMIDO = 800 * MB     # suma de lo extraído de un ZIP
ZIP_MAX_ARCHIVO = 150 * MB           # cada archivo dentro del ZIP
ZIP_MAX_ELEMENTOS = 500              # cantidad de archivos dentro del ZIP

# Plantilla: siempre la incluida en el repositorio, junto a app.py.
TEMPLATE_FILENAME = "Reporte_de_falla_Astara.xlsx"
TEMPLATE_PATH = Path(__file__).resolve().parent.parent / TEMPLATE_FILENAME
