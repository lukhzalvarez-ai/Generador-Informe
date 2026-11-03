"""Configuración del Service Report BYD: celdas de la plantilla, tablas,
recuadros de fotos, límites de peso y formatos aceptados.

Si BYD cambia la plantilla (filas o columnas), basta con ajustar este archivo."""

import logging
from pathlib import Path

from PIL import Image as PILImage

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("reporte_byd")

PILImage.MAX_IMAGE_PIXELS = 200_000_000

# Fotos de iPhone (HEIC/HEIF) si 'pillow-heif' está instalado
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIF_DISPONIBLE = True
except ImportError:
    HEIF_DISPONIBLE = False

# --------------------------------------------------------------------------
# Plantilla (va dentro de la carpeta "Reporte tecnico BYD")
# --------------------------------------------------------------------------
TEMPLATE_FILENAME = "Reporte_tecnico_BYD.xlsx"
TEMPLATE_PATH = Path(__file__).resolve().parent.parent / TEMPLATE_FILENAME

# --------------------------------------------------------------------------
# Celdas de texto (fila de valores bajo cada título de la plantilla)
# --------------------------------------------------------------------------
CELDAS = {
    "report_number": "B5",
    "country": "E5",
    "vin": "B8",
    "model": "C8",
    "sales_date": "D8",
    "delivery_date": "E8",
    "kilometrage": "F8",
    "service_date": "G8",
    "fault_description": "B14",
    "troubleshooting": "B16",
    "repair_measure": "B18",
}

ETIQUETAS = {
    "report_number": "Report Number",
    "country": "Country",
    "vin": "VIN",
    "model": "Model",
    "sales_date": "Sales Date",
    "delivery_date": "Delivery date",
    "kilometrage": "Kilometrage",
    "service_date": "Service Date",
    "fault_description": "Fault Description",
    "troubleshooting": "Troubleshooting Ideas",
    "repair_measure": "Repair Measure",
}

PAIS_POR_DEFECTO = "Chile"

# Campos que NO vienen en los PDF y se llenan a mano en la app
CAMPOS_MANUALES = ("report_number", "country", "sales_date")
# Fechas: se escriben como fecha real en el Excel
CAMPOS_FECHA = ("sales_date", "delivery_date", "service_date")
# Se escriben como número si lo que se ingresó es un número (ej. "22.969" -> 22969)
CAMPOS_NUMERICOS = ("kilometrage",)

# Campos que conviene no dejar vacíos (se avisa, pero igual se genera)
CAMPOS_CRITICOS = ("vin", "model", "sales_date", "fault_description", "repair_measure")

# Recuadros de texto largo: se agrandan si el texto no cabe
CAMPOS_TEXTO_LARGO = ("fault_description", "troubleshooting", "repair_measure")

# --------------------------------------------------------------------------
# Tablas
# --------------------------------------------------------------------------
# Malfunction Part No. (B:D) / Name (E:G): 3 filas fijas
FILAS_PIEZA_FALLA = (10, 11, 12)
# Labor Hours Code (B:D) / Name Of Labor (E:G): 3 filas fijas
FILAS_MANO_OBRA = (20, 21, 22)
# Spare parts: B código, C nombre, D cantidad, E precio unitario, F:G total.
# La plantilla trae 3 filas; si hay más repuestos se insertan filas nuevas.
FILA_REPUESTOS = 24
FILAS_REPUESTOS_PLANTILLA = 3

# --------------------------------------------------------------------------
# Fotos
# --------------------------------------------------------------------------
# Foto 1 → B28:D28 (punto de falla) · Foto 2 → E28:G28 (etiqueta/código de pieza)
# Foto 3 → B30:G30 (factura) · Fotos 4+ → filas nuevas bajo la 30, mismo tamaño que B30
FILA_FOTOS_SUPERIOR = 28
FILA_FOTOS_ANCHA = 30
RECUADROS_FOTO = [
    {"fila": FILA_FOTOS_SUPERIOR, "col_ini": "B", "col_fin": "D",
     "nombre": "Photo Of Malfunction Point", "ayuda": "punto de falla"},
    {"fila": FILA_FOTOS_SUPERIOR, "col_ini": "E", "col_fin": "G",
     "nombre": "Photo Of Parts Brand (Code)", "ayuda": "etiqueta / código de la pieza"},
    {"fila": FILA_FOTOS_ANCHA, "col_ini": "B", "col_fin": "G",
     "nombre": "Photo Of Parts Procurement Invoice", "ayuda": "factura de compra"},
]
FOTO_ADICIONAL = {"col_ini": "B", "col_fin": "G", "nombre": "Additional photo", "ayuda": "foto adicional"}
# Desde la 3ª foto: ancho completo B:G y alto = ancho / 1.5 (formato 3:2, la mitad de
# alto que una foto vertical completa). La fila del Excel toma ese alto.
PROPORCION_FOTO_ANCHA = 1.5
MARGEN_FOTO_PX = 6             # aire entre la foto y el borde del recuadro

# Forma por defecto (ancho, alto) usada por utilidades genéricas de imagen
PROPORCION_RECUADRO = (4, 3)

# --------------------------------------------------------------------------
# Control de peso (igual criterio que Astara)
# --------------------------------------------------------------------------
MB = 1024 * 1024
LIMITE_REPORTE_BYTES = 25 * MB
OBJETIVO_REPORTE_BYTES = int(24.0 * MB)
N_FOTOS_PRIORITARIAS = 2       # las 2 fotos de arriba (B28 y E28) en alta calidad

NIVELES_PRIORITARIAS = [(4096, 92), (3200, 90), (2560, 88), (2048, 85)]
NIVELES_SECUNDARIAS = [
    (4096, 88), (3000, 85), (2400, 83), (2000, 80), (1600, 78),
    (1366, 75), (1152, 72), (1024, 68), (900, 64), (800, 60),
]

# --------------------------------------------------------------------------
# Archivos aceptados (mismos nombres de PDF que Astara)
# --------------------------------------------------------------------------
EXTENSIONES_IMAGEN = {"jpg", "jpeg", "jfif", "png", "webp", "bmp", "gif", "tif", "tiff"}
if HEIF_DISPONIBLE:
    EXTENSIONES_IMAGEN |= {"heic", "heif"}

NOMBRE_PDF_ORDEN = "orden de trabajo"
NOMBRES_PDF_GARANTIAS = {"trabajo en garantias", "trabajo en garantia"}
# Factura de compra: un PDF llamado "Factura" (también "Factura 12345", "Factura (1)"...).
# Se convierte en imagen y va en "Photo Of Parts Procurement Invoice", a todo el ancho y
# con su forma real (sin recortar). Si no se adjunta, el informe sale igual sin ella.
NOMBRE_PDF_FACTURA = "factura"
FACTURA_ANCHO_PX = 1650        # ≈ 200 dpi en hoja A4/carta: se lee bien al imprimir
FACTURA_CALIDAD_JPEG = 85
FACTURA_MAX_PAGINAS = 3

ZIP_MAX_DESCOMPRIMIDO = 800 * MB
ZIP_MAX_ARCHIVO = 150 * MB
ZIP_MAX_ELEMENTOS = 500

# --------------------------------------------------------------------------
# Lectura de PDF (copias propias para poder ajustarlas sin tocar Astara)
# --------------------------------------------------------------------------
MAPA_CONCESIONARIOS_BYD = {
    "OTA": "ASTARA RETAIL ALAMEDA",
    "OTE": "ASTARA RETAIL EL CORTIJO",
    "OTC": "ASTARA RETAIL CANTAGALLO",
    "OTN": "ASTARA RETAIL NOS",
    "OTG": "ASTARA RETAIL LA DEHESA",
    "OTM": "ASTARA RETAIL MACHALÍ",
}

ETIQUETAS_BLOQUE_DIAGNOSTICO_BYD = {
    "cód. avería",
    "tipo garantía",
    "% deferencia mo",
    "gf",
    "kmts",
}
