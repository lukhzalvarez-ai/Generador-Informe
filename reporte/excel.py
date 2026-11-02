"""Armado del Excel final a partir de la plantilla: escribe los textos,
ajusta el alto de los recuadros de texto largo e inserta las fotos sin
superar el límite de peso."""

import io

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage

from .config import (logger, MB, TARGET_WIDTH, TARGET_HEIGHT, CELDAS_EXCEL, CAMPOS_TEXTO_LARGO,
                     LIMITE_REPORTE_BYTES, OBJETIVO_REPORTE_BYTES, TEMPLATE_FILENAME, TEMPLATE_PATH)
from .pdf import validar_campos_criticos
from .imagenes import calcular_celda_imagen, comprimir_fotos_para_limite
from .archivos import clave_foto


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
