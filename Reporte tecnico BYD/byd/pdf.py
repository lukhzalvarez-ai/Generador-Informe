"""Lectura de los PDF (Orden de Trabajo y Trabajo en Garantías) para el
Service Report BYD. Las funciones de lectura base son una copia de las de
Astara; al final están las específicas de BYD (modelo, fechas y textos de los
trabajos de la OT). Las tablas de piezas, mano de obra y repuestos se llenan
a mano en la app.
Los textos se copian TAL CUAL vienen en los PDF (no se traducen)."""

import io
import re

from .config import (logger, ETIQUETAS_BLOQUE_DIAGNOSTICO_BYD as ETIQUETAS_BLOQUE_DIAGNOSTICO,
                     MAPA_CONCESIONARIOS_BYD as MAPA_CONCESIONARIOS, CAMPOS_CRITICOS)

try:
    import pypdf
    PDF_LIBRARY_AVAILABLE = True
except ImportError:
    try:
        import PyPDF2 as pypdf
        PDF_LIBRARY_AVAILABLE = True
    except ImportError:
        PDF_LIBRARY_AVAILABLE = False


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
                logger.info("Bloque de diagnóstico vacío tras filtrar etiquetas y números.")
        else:
            logger.warning("No se pudo extraer el bloque de diagnóstico técnico de Garantías.")

        # (BYD) Los repuestos se ingresan a mano: no se leen de Garantías.

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
# Específico BYD
# --------------------------------------------------------------------------
def _anio_completo(fecha: str) -> str:
    """'28-02-25' -> '28-02-2025' (las fechas de 4 dígitos quedan igual)."""
    partes = fecha.split("-")
    if len(partes) == 3 and len(partes[2]) == 2:
        return f"{partes[0]}-{partes[1]}-20{partes[2]}"
    return fecha


def _buscar(patron: str, texto: str, grupo: int = 1, flags=re.IGNORECASE) -> str:
    m = re.search(patron, texto or "", flags)
    return m.group(grupo).strip() if m else ""


def extraer_trabajos_ot(texto_ot: str) -> list:
    """Textos de los trabajos de la Orden de Trabajo: lo que viene después de
    'Taller Mecánica' hasta el número de línea. Ej.:
      'luz trasera de portalón se humedece por dentro y no prende/...'
      'Garantia por FOCO TRASERO CENTRAL, FUSIBLE ESPECIAL 7.5 AMP, ...'"""
    bloques = re.findall(r"Taller\s+Mec[aá]nica\s*\n(.*?)\n\s*\d+\s*(?:\n|$)", texto_ot or "", re.DOTALL)
    return [_normalizar_espacios(b) for b in bloques if _normalizar_espacios(b)]


def extraer_valores_byd(clasificacion: dict) -> dict:
    """Valores que SÍ se pueden leer de los PDF para el Service Report BYD.

    Se leen: VIN, modelo, fecha de entrega (inicio de garantía, si viene),
    kilometraje, fecha de servicio (apertura de la OT) y los textos de falla /
    diagnóstico / reparación (tal cual vienen).

    Se llenan A MANO en la app (no vienen en los PDF): Report Number, Country,
    Sales Date, Malfunction Part No./Name, Labor Hours Code / Name Of Labor y
    la tabla de repuestos (código, nombre, cantidad, precio unitario, total)."""
    texto_ot = clasificacion.get("pdf_orden") or ""
    texto_ga = clasificacion.get("pdf_garantias") or ""
    ot = extraer_datos_pdf(texto_ot)
    ga = extraer_datos_from_garantias_pdf(texto_ga)

    vin = ot["vin"] or _buscar(r"VIN\s*[:]?\s*([A-HJ-NPR-Z0-9]{17})", texto_ga)
    modelo = _buscar(r"Asesor de\s*\n\s*Serv\.\s*\n\s*([^\n]+)", texto_ot)
    inicio_garantia = _anio_completo(_buscar(r"F\.\s*Inicio\s*Garant[ií]a\s*\n\s*(\d{2}-\d{2}-\d{2,4})", texto_ot))
    km = ot["km"] or _buscar(r"Kmts\s*[\s\S]*?\n\s*([\d]{1,3}(?:\.\d{3})+|\d+)\s*\n", texto_ga)

    # Trabajos de la OT: los que empiezan con "Garantía por ..." describen lo que se
    # reparó; el resto describe la falla / lo que pidió el cliente.
    trabajos = extraer_trabajos_ot(texto_ot)
    de_garantia = [t for t in trabajos if re.match(r"garant[ií]a\b", t, re.IGNORECASE)]
    de_falla = [t for t in trabajos if t not in de_garantia]
    queja_ga = ga.get("queja_cliente_garantias", "")

    falla = "\n".join(de_falla) or ot.get("queja_cliente", "") or queja_ga
    reparacion = ot.get("trabajo_realizado", "") or "\n".join(de_garantia) or \
        (queja_ga if re.match(r"garant[ií]a\b", queja_ga, re.IGNORECASE) else "")

    return {
        "vin": vin,
        "model": modelo,
        "delivery_date": inicio_garantia,
        "kilometrage": km,
        "service_date": ot["fecha_falla"],
        "fault_description": falla,
        "troubleshooting": ga.get("diagnostico_tecnico_garantias", "") or ot.get("diagnostico_tecnico", ""),
        "repair_measure": reparacion,
    }
