"""Lectura de un Service Report BYD ya generado (a medio completar).

Flujo de uso: para la solicitud de garantía se genera un informe a medio llenar;
para cobrar el caso se adjunta ese mismo Excel y la app lo vuelve a cargar
(textos, fechas, tablas, fotos y factura) para completarlo y emitir el definitivo.

Se reconoce por su CONTENIDO (títulos de la plantilla BYD), no por el nombre del
archivo, y los campos se buscan por sus títulos, no por celdas fijas, para tolerar
filas agregadas (repuestos de más, fotos adicionales o cambios hechos a mano).
"""

import io
import json
import datetime

import openpyxl

from .config import logger
from .archivos import ArchivoEnMemoria

EXTENSIONES_EXCEL = {"xlsx", "xlsm"}
PROPIEDAD_GENERADOR = "GeneradorReporteBYD"      # marca interna que deja la app al generar

# Títulos de la plantilla (en minúsculas, sin ":" ni espacios sobrantes)
T_REPORT = "report number"
T_VIN = "vin"
T_PIEZA = "malfunction part no."
T_FALLA = "fault description"
T_DIAG = "troubleshooting ideas"
T_REPARACION = "repair measure"
T_MO = "labor hours code"
T_REP = "spare part code"
T_FOTO = "photo of malfunction point"
T_FACTURA = "photo of parts procurement invoice"


def _norm(v) -> str:
    if not isinstance(v, str):
        return ""
    return " ".join(v.replace("\n", " ").split()).strip().rstrip(":").strip().lower()


def _filas_por_titulo(ws) -> dict:
    """{titulo: fila} buscando en las columnas B..G (primera aparición)."""
    filas = {}
    for fila in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400), min_col=2, max_col=7):
        for celda in fila:
            t = _norm(celda.value)
            if not t:
                continue
            for titulo in (T_REPORT, T_VIN, T_PIEZA, T_FALLA, T_DIAG, T_REPARACION, T_MO, T_REP, T_FOTO, T_FACTURA):
                if titulo not in filas and (t == titulo or t.startswith(titulo + " ")):
                    filas[titulo] = celda.row
    return filas


def es_informe_byd(datos: bytes) -> bool:
    """True si el Excel es un Service Report BYD (tiene los títulos clave de la plantilla)."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(datos), read_only=True)
        ws = wb.worksheets[0]
        encontrados = set()
        for fila in ws.iter_rows(min_row=1, max_row=200, min_col=2, max_col=7, values_only=True):
            for v in fila:
                t = _norm(v)
                for titulo in (T_VIN, T_PIEZA, T_FALLA, T_REP, T_FOTO):
                    if t == titulo or t.startswith(titulo + " "):
                        encontrados.add(titulo)
        wb.close()
        return len(encontrados) >= 4
    except Exception:
        return False


def _texto(v) -> str:
    """Valor de celda -> texto para los campos/tablas de la página (números al estilo chileno)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d-%m-%Y")
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return f"{v:.4f}".rstrip("0").rstrip(".").replace(".", ",")
    if isinstance(v, int):
        return str(v)
    return str(v).strip()


def _fecha(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y", "%m-%d-%Y", "%m/%d/%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _filas_tabla(ws, desde: int, hasta: int, columnas: list) -> list:
    """Filas desde..hasta-1 como listas de texto (columnas = letras)."""
    res = []
    for f in range(desde, hasta):
        fila = [_texto(ws[f"{c}{f}"].value) for c in columnas]
        res.append(fila)
    return res


def _metadatos(wb) -> dict:
    try:
        for p in wb.custom_doc_props.props:
            if p.name == PROPIEDAD_GENERADOR:
                return json.loads(p.value or "{}")
    except Exception:
        pass
    return {}


def leer_informe_byd(datos: bytes, nombre: str = "informe.xlsx") -> dict:
    """Devuelve:
      campos   : {campo: valor} (fechas como date; el resto texto)
      piezas   : [[no, nombre], ...]          (3 filas)
      mano_obra: [[codigo, nombre], ...]      (3 filas)
      repuestos: [[cod, nombre, cant, precio, total], ...]
      fotos    : [ArchivoEnMemoria, ...] en el orden del informe (sin la factura)
      facturas : [bytes, ...] imágenes de la factura que traía el informe
      factura_por_forma : True si la factura se reconoció solo por su forma (informes antiguos)
      avisos   : [texto, ...] campos que no se pudieron leer
    """
    wb = openpyxl.load_workbook(io.BytesIO(datos))
    ws = wb.worksheets[0]
    filas = _filas_por_titulo(ws)
    avisos = []
    campos = {}

    def valor(fila_titulo, col, desplazamiento=1):
        f = filas.get(fila_titulo)
        return ws[f"{col}{f + desplazamiento}"].value if f else None

    if T_REPORT in filas:
        campos["report_number"] = _texto(valor(T_REPORT, "B"))
        campos["country"] = _texto(valor(T_REPORT, "E"))
    else:
        avisos.append("Report Number / Country")
    if T_VIN in filas:
        campos["vin"] = _texto(valor(T_VIN, "B"))
        campos["model"] = _texto(valor(T_VIN, "C"))
        campos["sales_date"] = _fecha(valor(T_VIN, "D"))
        campos["delivery_date"] = _fecha(valor(T_VIN, "E"))
        campos["kilometrage"] = _texto(valor(T_VIN, "F"))
        campos["service_date"] = _fecha(valor(T_VIN, "G"))
    else:
        avisos.append("Vehicle's Information (VIN, Model, fechas, km)")
    for campo, titulo in (("fault_description", T_FALLA), ("troubleshooting", T_DIAG),
                          ("repair_measure", T_REPARACION)):
        if titulo in filas:
            campos[campo] = _texto(valor(titulo, "B"))
        else:
            avisos.append(titulo.capitalize())

    def tabla(titulo, siguiente, columnas, n_min):
        if titulo not in filas:
            avisos.append(titulo.capitalize())
            return []
        desde = filas[titulo] + 1
        hasta = filas.get(siguiente, desde + n_min)
        hasta = max(desde, min(hasta, desde + 60))
        return _filas_tabla(ws, desde, hasta, columnas)

    piezas = [[r[0], r[1]] for r in tabla(T_PIEZA, T_FALLA, ["B", "E"], 3)][:3]
    mano_obra = [[r[0], r[1]] for r in tabla(T_MO, T_REP, ["B", "E"], 3)][:3]
    repuestos = tabla(T_REP, T_FOTO, ["B", "C", "D", "E", "F"], 3)
    repuestos = [r for r in repuestos if any(x.strip() for x in r)] or repuestos[:3]

    # ---------------- Fotos y factura ----------------
    fila_fotos = (filas.get(T_FOTO) or 0) + 1
    fila_factura = (filas.get(T_FACTURA) or fila_fotos + 1) + 1
    imagenes = []
    for img in getattr(ws, "_images", []):
        try:
            desde = img.anchor._from
            fila, col = desde.row + 1, desde.col
            if fila < fila_fotos:               # logo de la plantilla u otras imágenes de arriba
                continue
            imagenes.append((fila, col, img._data()))
        except Exception as e:
            logger.warning(f"Imagen del informe no leída: {e}")
    imagenes.sort(key=lambda x: (x[0], x[1]))
    arriba = [x for x in imagenes if x[0] < fila_factura]
    anchas = [x for x in imagenes if x[0] >= fila_factura]

    meta = _metadatos(wb)
    facturas, por_forma = [], False
    if "facturas" in meta:
        n = int(meta.get("facturas") or 0)
        facturas, anchas = [x[2] for x in anchas[:n]], anchas[n:]
    else:
        # Informes generados antes de la marca interna: la factura es una hoja vertical
        # al comienzo del espacio "Procurement Invoice".
        while anchas and _es_hoja_vertical(anchas[0][2]):
            facturas.append(anchas.pop(0)[2])
            por_forma = True

    fotos = []
    for i, (_, _, b) in enumerate(arriba + anchas, start=1):
        ext = "png" if b[:4] == b"\x89PNG" else "jpg"
        a = ArchivoEnMemoria(f"Informe anterior · foto {i:02d}.{ext}", b, origen=nombre)
        a.del_informe = True
        fotos.append(a)

    return {"campos": campos, "piezas": piezas, "mano_obra": mano_obra, "repuestos": repuestos,
            "fotos": fotos, "facturas": facturas, "factura_por_forma": por_forma, "avisos": avisos}


def _es_hoja_vertical(datos: bytes) -> bool:
    from PIL import Image as PILImage
    try:
        with PILImage.open(io.BytesIO(datos)) as im:
            w, h = im.size
        return 1.2 <= h / w <= 1.6
    except Exception:
        return False
