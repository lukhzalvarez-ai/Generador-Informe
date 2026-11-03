"""Armado del Service Report BYD a partir de su plantilla:
  - escribe los datos del vehículo y los textos (tal cual vienen),
  - llena las tablas (pieza en falla, mano de obra, repuestos) insertando filas
    si hay más repuestos que los 3 de la plantilla,
  - pone las fotos en sus recuadros SIN deformarlas (ajustadas y centradas):
      foto 1 → B28:D28 · foto 2 → E28:G28 · foto 3 → B30:G30 ·
      fotos 4+ → filas nuevas bajo la anterior, del mismo tamaño que B30,
  - y no supera los 25 MB."""

import io
import re
import copy
import datetime
from functools import lru_cache

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage

from .config import (logger, MB, CELDAS, CAMPOS_TEXTO_LARGO, CAMPOS_CRITICOS, TEMPLATE_PATH,
                     CAMPOS_FECHA, CAMPOS_NUMERICOS,
                     TEMPLATE_FILENAME, FILAS_PIEZA_FALLA, FILAS_MANO_OBRA, FILA_REPUESTOS,
                     FILAS_REPUESTOS_PLANTILLA, PROPORCION_FOTO_ANCHA, RECUADROS_FOTO, FOTO_ADICIONAL, MARGEN_FOTO_PX,
                     LIMITE_REPORTE_BYTES, OBJETIVO_REPORTE_BYTES)
from .imagenes import comprimir_fotos_para_limite, encuadre_relleno, tamano_orientado
from .archivos import clave_foto


# --------------------------------------------------------------------------
# Plantilla y medidas
# --------------------------------------------------------------------------
def cargar_plantilla() -> bytes:
    if not TEMPLATE_PATH.exists():
        return None
    try:
        return TEMPLATE_PATH.read_bytes()
    except Exception as e:
        logger.error(f"No se pudo leer la plantilla BYD: {e}")
        return None


def _px_columna(ws, letra: str) -> int:
    dim = ws.column_dimensions.get(letra)
    ancho = dim.width if dim is not None and dim.width else 8.43
    return int(round(ancho * 7))


def _px_fila(ws, fila: int) -> int:
    alto = ws.row_dimensions[fila].height or 15.0
    return int(round(alto * 96 / 72))


def _tamano_recuadro(ws, fila: int, col_ini: str, col_fin: str) -> tuple:
    ancho = sum(_px_columna(ws, get_column_letter(c))
                for c in range(column_index_from_string(col_ini), column_index_from_string(col_fin) + 1))
    return ancho, _px_fila(ws, fila)


@lru_cache(maxsize=1)
def proporciones_recuadros() -> tuple:
    """(ancho/alto) de: recuadro 1, recuadro 2, recuadro ancho (3 y siguientes).
    Se usa en la pantalla para que el recortador tenga la forma correcta."""
    datos = cargar_plantilla()
    if not datos:
        return (1.25, 1.25, PROPORCION_FOTO_ANCHA)
    ws = openpyxl.load_workbook(io.BytesIO(datos)).active
    res = []
    for r in RECUADROS_FOTO[:-1]:
        w, h = _tamano_recuadro(ws, r["fila"], r["col_ini"], r["col_fin"])
        res.append(w / h)
    res.append(PROPORCION_FOTO_ANCHA)      # 3ª en adelante: el alto de la fila se adapta a 3:2
    return tuple(res)


def proporcion_para_indice(indice: int) -> float:
    p = proporciones_recuadros()
    return p[min(indice, len(p) - 1)]


def alto_libre(indice: int) -> bool:
    """(Desactivado) Antes la 3ª foto en adelante iba completa con el alto que necesitara;
    salían demasiado largas. Ahora todas usan un recuadro fijo: la 3ª en adelante con
    ancho B:G y formato 3:2 (PROPORCION_FOTO_ANCHA)."""
    return False


def nombre_recuadro(indice: int) -> tuple:
    """(nombre en el informe, ayuda en español) del recuadro de la foto N (0 = primera)."""
    if indice < len(RECUADROS_FOTO):
        r = RECUADROS_FOTO[indice]
    else:
        r = FOTO_ADICIONAL
    return r["nombre"], r["ayuda"]


# --------------------------------------------------------------------------
# Números y fechas: se escriben como valores reales (no texto), así la plantilla
# les aplica su formato (ej. precios con "$ 295,344").
# --------------------------------------------------------------------------
FORMATO_FECHA_RESPALDO = "dd-mm-yyyy"   # si la plantilla no define formato de fecha
# Formato de moneda usado en los reportes BYD rellenados (ej. "$ 295,344")
FORMATO_PRECIO = '_ "$"* #,##0_ ;_ "$"* \\-#,##0_ ;_ "$"* "-"_ ;_ @_ '


def a_numero(valor):
    """'22.969' -> 22969 · '295344' -> 295344 · '0,4' / '0.4' -> 0.4 · '$ 2.426' -> 2426.
    Si no es un número, se devuelve el texto tal cual."""
    if isinstance(valor, (int, float)) or valor in (None, ""):
        return valor
    t = str(valor).strip().replace("$", "").replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", t):          # miles con punto (formato chileno)
        n = float(t.replace(".", "").replace(",", "."))
    elif re.fullmatch(r"-?\d+(,\d+)?", t):                       # decimal con coma
        n = float(t.replace(",", "."))
    elif re.fullmatch(r"-?\d+\.\d+", t):                         # decimal con punto
        n = float(t)
    else:
        return str(valor).strip()
    return int(n) if n.is_integer() else n


def a_fecha(valor):
    """date/datetime o texto 'dd-mm-aaaa', 'dd-mm-aa', 'dd/mm/aaaa' -> datetime. Si no, texto."""
    if isinstance(valor, datetime.datetime):
        return valor
    if isinstance(valor, datetime.date):
        return datetime.datetime(valor.year, valor.month, valor.day)
    t = str(valor or "").strip()
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})", t)
    if not m:
        return t
    d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
    a = a + 2000 if a < 100 else a
    try:
        return datetime.datetime(a, mes, d)
    except ValueError:
        return t


# --------------------------------------------------------------------------
# Escritura de celdas
# --------------------------------------------------------------------------
def _escribir(ws, coordenada: str, valor) -> None:
    """Escribe aunque la celda sea parte de un rango combinado (usa la de arriba a la izquierda)."""
    valor = valor if valor not in ("", None) else None
    celda = ws[coordenada]
    if isinstance(celda, MergedCell):
        for rango in ws.merged_cells.ranges:
            if coordenada in rango:
                ws.cell(row=rango.min_row, column=rango.min_col).value = valor
                return
        return
    celda.value = valor


def _ajustar_alto_texto(ws, coordenada: str, texto: str) -> None:
    """Agranda el alto de las filas del recuadro si el texto no cabe."""
    if not texto:
        return
    rango = next((r for r in ws.merged_cells.ranges if coordenada in r), None)
    filas = list(range(rango.min_row, rango.max_row + 1)) if rango else [ws[coordenada].row]
    cols = range(rango.min_col, rango.max_col + 1) if rango else [ws[coordenada].column]
    ancho_chars = sum(_px_columna(ws, get_column_letter(c)) for c in cols) / 6.2
    lineas = sum(max(1, -(-len(p) // max(int(ancho_chars), 1))) for p in str(texto).split("\n"))
    necesario = lineas * 15.0 + 6
    actual = sum((ws.row_dimensions[f].height or 15.0) for f in filas)
    if necesario > actual:
        for f in filas:
            ws.row_dimensions[f].height = min(necesario / len(filas), 409)
    celda = ws.cell(row=filas[0], column=cols[0])
    al = copy.copy(celda.alignment)
    celda.alignment = openpyxl.styles.Alignment(horizontal=al.horizontal, vertical="top",
                                                wrap_text=True, indent=al.indent)


def _lineas_texto(texto: str, ancho_px: int, px_por_letra: float = 8.6) -> int:
    """Cuántas líneas ocupa un texto con ajuste de línea POR PALABRAS (como Excel),
    estimando el ancho de cada letra en Arial 11."""
    max_letras = max(int(ancho_px / px_por_letra), 1)
    lineas, actual = 1, 0
    for palabra in texto.split():
        largo = len(palabra)
        while largo > max_letras:          # palabra más larga que la celda: se corta
            if actual:
                lineas += 1
            largo -= max_letras
            actual = 0
            lineas += 1 if largo else 0
        if actual == 0:
            actual = largo
        elif actual + 1 + largo <= max_letras:
            actual += 1 + largo
        else:
            lineas += 1
            actual = largo
    return lineas


def _fijar_alto_fila(ws, fila: int, minimo: float = 15.0) -> None:
    """Fija el alto de una fila según el texto más largo que tenga (considerando
    celdas combinadas). Si se deja en automático, Excel/LibreOffice agrandan la
    fila al abrir el archivo DESPUÉS de ubicar las fotos, y estas quedan
    desplazadas encima de la tabla."""
    lineas_max = 1
    for c in range(column_index_from_string("B"), column_index_from_string("G") + 1):
        celda = ws.cell(row=fila, column=c)
        if isinstance(celda, MergedCell) or celda.value in (None, ""):
            continue
        rango = next((r for r in ws.merged_cells.ranges if celda.coordinate in r), None)
        cols = range(rango.min_col, rango.max_col + 1) if rango else [c]
        ancho_px = sum(_px_columna(ws, get_column_letter(x)) for x in cols) - 6
        lineas = sum(_lineas_texto(p, ancho_px) for p in str(celda.value).split("\n"))
        lineas_max = max(lineas_max, lineas)
    alto = max(minimo, ws.row_dimensions[fila].height or 0, lineas_max * 14.0 + 3)
    ws.row_dimensions[fila].height = alto


def _ajustar_impresion(ws) -> None:
    """Imprime/exporta a PDF con todo el ancho del informe en una página
    (la plantilla original corta la columna 'Total Price')."""
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def _insertar_filas(ws, desde: int, cantidad: int) -> None:
    """Inserta 'cantidad' filas antes de 'desde' moviendo también lo que openpyxl
    no mueve solo: rangos combinados y altos de fila."""
    if cantidad <= 0:
        return
    afectados = [r for r in list(ws.merged_cells.ranges) if r.min_row >= desde]
    for r in afectados:
        ws.merged_cells.remove(r)
    alturas = {f: ws.row_dimensions[f].height for f in range(desde, ws.max_row + 1)}
    ws.insert_rows(desde, cantidad)
    for r in afectados:
        r.shift(0, cantidad)
        ws.merged_cells.add(r)
    for f in range(desde, ws.max_row + 1):
        ws.row_dimensions[f].height = None
    for f, alto in alturas.items():
        ws.row_dimensions[f + cantidad].height = alto


def _copiar_estilo_fila(ws, fila_origen: int, fila_destino: int, col_ini="B", col_fin="G") -> None:
    for c in range(column_index_from_string(col_ini), column_index_from_string(col_fin) + 1):
        origen = ws.cell(row=fila_origen, column=c)
        destino = ws.cell(row=fila_destino, column=c)
        if origen.has_style:
            destino._style = copy.copy(origen._style)
    ws.row_dimensions[fila_destino].height = ws.row_dimensions[fila_origen].height


def _llenar_tabla_dos_columnas(ws, filas: tuple, datos: list) -> None:
    """Tablas 'código (B:D) / nombre (E:G)' de 3 filas fijas."""
    for i, fila in enumerate(filas):
        d = datos[i] if i < len(datos) else {}
        _escribir(ws, f"B{fila}", d.get("codigo", ""))
        _escribir(ws, f"E{fila}", a_numero(d.get("nombre", "")))


def _llenar_repuestos(ws, repuestos: list) -> int:
    """Llena la tabla de repuestos; si hay más de 3 inserta filas. Devuelve
    cuántas filas se insertaron (para correr lo que está debajo)."""
    extra = max(0, len(repuestos) - FILAS_REPUESTOS_PLANTILLA)
    ultima = FILA_REPUESTOS + FILAS_REPUESTOS_PLANTILLA - 1
    if extra:
        _insertar_filas(ws, ultima + 1, extra)
        for k in range(extra):
            fila = ultima + 1 + k
            _copiar_estilo_fila(ws, ultima, fila)
            ws.merge_cells(f"F{fila}:G{fila}")
    for i in range(max(len(repuestos), FILAS_REPUESTOS_PLANTILLA)):
        fila = FILA_REPUESTOS + i
        r = repuestos[i] if i < len(repuestos) else {}
        _escribir(ws, f"B{fila}", r.get("codigo", ""))
        _escribir(ws, f"C{fila}", r.get("nombre", ""))
        _escribir(ws, f"D{fila}", a_numero(r.get("cantidad", "")))
        _escribir(ws, f"E{fila}", a_numero(r.get("precio", "")))
        _escribir(ws, f"F{fila}", a_numero(r.get("total", "")))
        for col in ("E", "F"):
            if isinstance(ws[f"{col}{fila}"].value, (int, float)):
                ws[f"{col}{fila}"].number_format = FORMATO_PRECIO
        if r.get("nombre"):
            celda = ws[f"C{fila}"]
            celda.alignment = openpyxl.styles.Alignment(
                horizontal=celda.alignment.horizontal, vertical="center", wrap_text=True)
    return extra


# --------------------------------------------------------------------------
# Fotos
# --------------------------------------------------------------------------
ALTO_MAX_FILA_PX = 545          # Excel no permite filas de más de 409 pt (= 545 px)


def _colocar_imagen(ws, jpeg: bytes, fila: int, col_ini: str, off_x: int, off_y: int, w: int, h: int) -> None:
    col = column_index_from_string(col_ini)
    while off_x >= _px_columna(ws, get_column_letter(col)):
        off_x -= _px_columna(ws, get_column_letter(col))
        col += 1
    img = XLImage(io.BytesIO(jpeg))
    img.width, img.height = w, h
    marcador = AnchorMarker(col=col - 1, colOff=pixels_to_EMU(off_x), row=fila - 1, rowOff=pixels_to_EMU(off_y))
    img.anchor = OneCellAnchor(_from=marcador, ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
    ws.add_image(img)


def _colocar_fotos_anchas(ws, jpegs: list, fila_inicio: int) -> int:
    """Fotos 3 en adelante: cada una ocupa TODO el ancho B:G y la fila toma el alto que
    necesita la foto (sin recortarla). Si pasa el máximo de Excel por fila, la foto
    ocupa 2 o más filas combinadas. Devuelve cuántas se insertaron."""
    ancho_caja = sum(_px_columna(ws, get_column_letter(c)) for c in
                     range(column_index_from_string(FOTO_ADICIONAL["col_ini"]),
                           column_index_from_string(FOTO_ADICIONAL["col_fin"]) + 1))
    ancho_foto = ancho_caja - 2 * MARGEN_FOTO_PX
    if not jpegs:
        return 0
    for r in [r for r in ws.merged_cells.ranges if r.min_row == fila_inicio and r.min_col == 2]:
        ws.merged_cells.remove(r)                      # B30:G30 se vuelve a combinar a la medida
    fila, insertadas = fila_inicio, 0
    for jpeg in jpegs:
        try:
            with PILImage.open(io.BytesIO(jpeg)) as im:
                iw, ih = im.size
            alto_foto = max(1, round(ancho_foto * ih / iw))
            total = alto_foto + 2 * MARGEN_FOTO_PX
            n_filas = -(-total // ALTO_MAX_FILA_PX)          # división hacia arriba
            alto_fila_pt = round(total / n_filas * 72 / 96 + 0.75, 2)
            for f in range(fila, fila + n_filas):
                if f != fila_inicio:
                    _copiar_estilo_fila(ws, fila_inicio, f)
                ws.row_dimensions[f].height = alto_fila_pt
            ws.merge_cells(f"{FOTO_ADICIONAL['col_ini']}{fila}:{FOTO_ADICIONAL['col_fin']}{fila + n_filas - 1}")
            _colocar_imagen(ws, jpeg, fila, FOTO_ADICIONAL["col_ini"], MARGEN_FOTO_PX, MARGEN_FOTO_PX,
                            ancho_foto, alto_foto)
            fila += n_filas
            insertadas += 1
        except Exception as e:
            logger.error(f"No se pudo insertar una foto en la fila {fila}: {e}")
    return insertadas


def _colocar_foto(ws, jpeg: bytes, fila: int, col_ini: str, col_fin: str) -> None:
    """Inserta la foto llenando su recuadro (viene recortada con su forma) sin deformarla."""
    caja_w, caja_h = _tamano_recuadro(ws, fila, col_ini, col_fin)
    caja_w -= 2 * MARGEN_FOTO_PX
    caja_h -= 2 * MARGEN_FOTO_PX
    with PILImage.open(io.BytesIO(jpeg)) as im:
        iw, ih = im.size
    if abs((iw / ih) / (caja_w / caja_h) - 1) < 0.03:
        w, h = caja_w, caja_h            # ya viene con la forma del recuadro: lo llena completo
    else:                                # (solo con "proporción libre") se ajusta sin deformar
        escala = min(caja_w / iw, caja_h / ih)
        w, h = max(1, int(iw * escala)), max(1, int(ih * escala))
    off_x = MARGEN_FOTO_PX + (caja_w - w) // 2
    off_y = MARGEN_FOTO_PX + (caja_h - h) // 2

    _colocar_imagen(ws, jpeg, fila, col_ini, off_x, off_y, w, h)


# --------------------------------------------------------------------------
# Armado completo
# --------------------------------------------------------------------------
def _armar_excel(plantilla: bytes, valores: dict, imagenes_jpeg: list) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(plantilla))
    ws = wb.active

    # Las fechas usan el mismo formato que trae la plantilla en "Sales Date"
    formato_fecha = ws[CELDAS["sales_date"]].number_format
    if formato_fecha in (None, "", "General"):
        formato_fecha = FORMATO_FECHA_RESPALDO
    for campo, celda in CELDAS.items():
        valor = valores.get(campo, "")
        if campo in CAMPOS_FECHA:
            valor = a_fecha(valor)
            if isinstance(valor, datetime.datetime):
                ws[celda].number_format = formato_fecha
        elif campo in CAMPOS_NUMERICOS:
            valor = a_numero(valor)
        _escribir(ws, celda, valor)
    for campo in CAMPOS_TEXTO_LARGO:
        _ajustar_alto_texto(ws, CELDAS[campo], valores.get(campo, ""))
    _llenar_tabla_dos_columnas(ws, FILAS_PIEZA_FALLA, valores.get("piezas_falla", []))
    _llenar_tabla_dos_columnas(ws, FILAS_MANO_OBRA, valores.get("mano_obra", []))
    desplazamiento = _llenar_repuestos(ws, valores.get("repuestos", []))
    # Altos fijos en las filas con datos (ver _fijar_alto_fila)
    filas_datos = [8, *FILAS_PIEZA_FALLA, *FILAS_MANO_OBRA]
    filas_datos += list(range(FILA_REPUESTOS, FILA_REPUESTOS + max(len(valores.get("repuestos", [])),
                                                                   FILAS_REPUESTOS_PLANTILLA)))
    for f in filas_datos:
        _fijar_alto_fila(ws, f)
    # El resto de las filas sin alto definido también se fija (la plantilla usa un
    # alto base menor al que necesita su letra y cada programa lo corrige distinto).
    for f in range(1, ws.max_row + 1):
        if ws.row_dimensions[f].height is None:
            _fijar_alto_fila(ws, f)
    _ajustar_impresion(ws)

    validas = [j for j in imagenes_jpeg if j]
    n_arriba = len(RECUADROS_FOTO) - 1                 # fotos 1 y 2 (B28:D28 y E28:G28)
    insertadas = 0
    for jpeg, r in zip(validas[:n_arriba], RECUADROS_FOTO[:n_arriba]):
        try:
            _colocar_foto(ws, jpeg, r["fila"] + desplazamiento, r["col_ini"], r["col_fin"])
            insertadas += 1
        except Exception as e:
            logger.error(f"No se pudo insertar una foto en la fila {r['fila']}: {e}")
    insertadas += _colocar_fotos_anchas(ws, validas[n_arriba:], RECUADROS_FOTO[-1]["fila"] + desplazamiento)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), insertadas


def validar_campos_criticos(valores: dict) -> list:
    return [c for c in CAMPOS_CRITICOS if not valores.get(c)]


def generar_reporte_byd(valores: dict, imagenes: list, ediciones: dict = None) -> tuple:
    """Devuelve (bytes_excel, faltantes, n_fotos_insertadas, info_peso). Nunca supera 25 MB."""
    faltantes = validar_campos_criticos(valores)
    plantilla = cargar_plantilla()
    if not plantilla:
        raise RuntimeError(f"No se encontró la plantilla '{TEMPLATE_FILENAME}' en la carpeta "
                           "'Reporte tecnico BYD'. Agrégala al repositorio.")
    ediciones = ediciones or {}
    fotos = []
    for i, a in enumerate(imagenes):
        datos = a.getvalue()
        ed = dict(ediciones.get(clave_foto(a)) or {})
        try:  # cada foto llena su recuadro completo (recorte con la forma del recuadro)
            w, h = tamano_orientado(datos, ed.get("rot", 0))
            if alto_libre(i):          # 3ª en adelante: foto completa (o el recorte que eligió el usuario)
                ed["crop"] = ed.get("crop")
            else:
                libre = (ed.get("ajuste") or {}).get("modo") == "libre_arrastre"
                ed["crop"] = encuadre_relleno(w, h, ed.get("crop"), proporcion_para_indice(i), libre)
        except Exception as e:
            logger.warning(f"No se pudo calcular el encuadre de {a.name}: {e}")
        fotos.append((a.name, datos, ed))

    peso_base = len(_armar_excel(plantilla, valores, [])[0])
    presupuesto = max(OBJETIVO_REPORTE_BYTES - peso_base - 4 * 1024 * len(fotos), 1 * MB)
    jpegs, info = comprimir_fotos_para_limite(fotos, presupuesto)
    excel, insertadas = _armar_excel(plantilla, valores, jpegs)

    intentos = 0
    while len(excel) > LIMITE_REPORTE_BYTES and intentos < 3 and fotos:
        presupuesto = max(presupuesto - (len(excel) - OBJETIVO_REPORTE_BYTES) - 512 * 1024, 1 * MB)
        jpegs, info = comprimir_fotos_para_limite(fotos, presupuesto)
        excel, insertadas = _armar_excel(plantilla, valores, jpegs)
        intentos += 1

    info.update({
        "peso_final_bytes": len(excel),
        "peso_original_fotos_bytes": sum(len(f[1]) for f in fotos),
        "bajo_limite": len(excel) <= LIMITE_REPORTE_BYTES,
        "n_fotos": len(fotos),
    })
    return excel, faltantes, insertadas, info
