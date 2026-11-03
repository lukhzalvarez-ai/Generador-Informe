"""Clasificación de los archivos subidos: qué es la Orden de Trabajo, qué es
Garantías, qué son fotos y qué se ignora. También identifica cada archivo
de forma barata (sin recorrer los megas completos)."""

import io
import re
import zipfile
import hashlib
import unicodedata
from pathlib import PurePosixPath

from PIL import Image as PILImage

from .config import (logger, MB, EXTENSIONES_IMAGEN, NOMBRE_PDF_ORDEN, NOMBRES_PDF_GARANTIAS, NOMBRE_PDF_FACTURA,
                     ZIP_MAX_DESCOMPRIMIDO, ZIP_MAX_ARCHIVO, ZIP_MAX_ELEMENTOS)
from .pdf import extraer_texto_pdf


class ArchivoEnMemoria(io.BytesIO):
    """Archivo extraído de un ZIP que se comporta igual que un archivo subido
    (tiene .name, .size y .getvalue()), así el resto de la app no distingue."""

    def __init__(self, nombre: str, datos: bytes, origen: str = ""):
        super().__init__(datos)
        self.name = nombre
        self.size = len(datos)
        self.origen = origen     # nombre del ZIP del que salió


def _nombre_en_zip(info: zipfile.ZipInfo) -> str:
    """Nombre real del archivo dentro del ZIP. El compresor de Windows guarda
    los nombres con tildes en la codificación antigua del sistema (cp850) sin
    avisarlo; Python los lee como cp437 y 'Garantías' saldría 'Garant¡as'.
    Se corrige para que los PDF se reconozcan igual."""
    nombre = info.filename
    if not info.flag_bits & 0x800:          # el ZIP no declara UTF-8
        try:
            nombre = nombre.encode("cp437").decode("cp850")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return nombre


def es_zip(archivo) -> bool:
    return archivo.name.lower().endswith(".zip")


def expandir_zip(nombre_zip: str, datos: bytes) -> tuple:
    """Abre un ZIP y devuelve (archivos_utiles, ignorados, avisos):
      - archivos_utiles: fotos y PDF como ArchivoEnMemoria (de cualquier subcarpeta)
      - ignorados: nombres de lo que no corresponde (Excel, Word, ZIP dentro de ZIP...)
      - avisos: problemas para mostrar al usuario (dañado, con clave, muy grande)."""
    utiles, ignorados, avisos = [], [], []
    try:
        zf = zipfile.ZipFile(io.BytesIO(datos))
    except zipfile.BadZipFile:
        return [], [], [f"'{nombre_zip}' no es un ZIP válido o está dañado."]

    with zf:
        elementos = [i for i in zf.infolist() if not i.is_dir()]
        if len(elementos) > ZIP_MAX_ELEMENTOS:
            avisos.append(f"'{nombre_zip}' trae {len(elementos)} archivos; solo se revisaron los primeros "
                          f"{ZIP_MAX_ELEMENTOS}.")
            elementos = elementos[:ZIP_MAX_ELEMENTOS]
        total = 0
        for info in elementos:
            nombre = PurePosixPath(_nombre_en_zip(info).replace("\\", "/")).name
            # basura de macOS/Windows que no son archivos reales
            if not nombre or nombre.startswith("._") or "__MACOSX" in info.filename \
                    or nombre.lower() in ("thumbs.db", "desktop.ini", ".ds_store"):
                continue
            ext = nombre.lower().rsplit(".", 1)[-1] if "." in nombre else ""
            if ext != "pdf" and ext not in EXTENSIONES_IMAGEN:
                ignorados.append(f"{nombre} (en {nombre_zip})")
                continue
            if info.flag_bits & 0x1:
                avisos.append(f"'{nombre}' dentro de '{nombre_zip}' está protegido con clave y no se pudo abrir.")
                continue
            if info.file_size > ZIP_MAX_ARCHIVO:
                avisos.append(f"'{nombre}' dentro de '{nombre_zip}' es demasiado grande "
                              f"({info.file_size / MB:.0f} MB) y se omitió.")
                continue
            if total + info.file_size > ZIP_MAX_DESCOMPRIMIDO:
                avisos.append(f"'{nombre_zip}' es demasiado grande al descomprimir; se usaron solo los "
                              f"primeros {ZIP_MAX_DESCOMPRIMIDO // MB} MB.")
                break
            try:
                contenido = zf.read(info)
            except Exception as e:
                logger.error(f"No se pudo extraer '{nombre}' de '{nombre_zip}': {e}")
                avisos.append(f"'{nombre}' dentro de '{nombre_zip}' está dañado y se omitió.")
                continue
            total += len(contenido)
            utiles.append(ArchivoEnMemoria(nombre, contenido, origen=nombre_zip))
    return utiles, ignorados, avisos


def tamano(archivo) -> int:
    """Tamaño en bytes sin copiar el contenido."""
    t = getattr(archivo, "size", None)
    return t if isinstance(t, int) else len(archivo.getvalue())


def huella(archivo) -> str:
    """Identificador rápido de un archivo: nombre + tamaño + md5 del primer y
    último bloque de 64 KB. Sirve de clave de caché sin procesar los 5-10 MB
    de cada foto en cada clic (eso era lo que hacía lenta la página)."""
    datos = archivo.getvalue()
    md5 = hashlib.md5(datos[:65536])
    md5.update(datos[-65536:])
    return f"{archivo.name}|{len(datos)}|{md5.hexdigest()}"


def clave_foto(archivo) -> str:
    """Identificador de una foto subida (nombre + tamaño), para guardar sus
    ediciones aunque Streamlit vuelva a ejecutar la página."""
    return f"{archivo.name}|{tamano(archivo)}"


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
    if clave == NOMBRE_PDF_FACTURA or clave.startswith(NOMBRE_PDF_FACTURA + " "):
        return "factura"
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


def clasificar_archivos_subidos(archivos_subidos: list, leer_pdf=None, es_imagen=None) -> dict:
    """Recibe la lista de UploadedFile del campo único (puede venir vacía)
    y se queda SOLO con:
      - el PDF llamado 'Orden de Trabajo'
      - el PDF llamado 'Trabajo en Garantías'
      - las fotos (jpg, jpeg, png, webp, heic, ...), sin repetir nombre
    Todo lo demás (otros PDF, Excel, Word, etc.) se ignora y se informa.
    Si se sube más de un PDF con el mismo nombre, se usa el primero.

    leer_pdf(archivo) -> texto y es_imagen(archivo) -> bool se pueden pasar
    para usar versiones con caché (la interfaz lo hace)."""
    leer_pdf = leer_pdf or (lambda a: extraer_texto_pdf(a.getvalue(), a.name))
    es_imagen = es_imagen or (lambda a: _es_imagen_valida(a.getvalue()))
    resultado = {
        "pdf_orden": None,
        "pdf_orden_nombre": "",
        "pdf_garantias": None,
        "pdf_garantias_nombre": "",
        "pdf_factura": None,          # el archivo (se convierte en imagen, no se lee el texto)
        "pdf_factura_nombre": "",
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
                resultado["pdf_orden"] = leer_pdf(archivo)
                resultado["pdf_orden_nombre"] = nombre
            elif tipo == "garantias" and resultado["pdf_garantias"] is None:
                resultado["pdf_garantias"] = leer_pdf(archivo)
                resultado["pdf_garantias_nombre"] = nombre
            elif tipo == "factura" and resultado["pdf_factura"] is None:
                resultado["pdf_factura"] = archivo
                resultado["pdf_factura_nombre"] = nombre
            else:
                resultado["ignorados"].append(nombre)
            continue

        if ext in EXTENSIONES_IMAGEN:
            clave = nombre.lower()
            if clave in vistos_imagenes or not es_imagen(archivo):
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
