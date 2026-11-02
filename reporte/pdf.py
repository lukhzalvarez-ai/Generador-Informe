"""Lectura de los PDF (Orden de Trabajo y Trabajo en Garantías) y armado de
los valores sugeridos para cada celda. Misma lógica de extracción original."""

import io
import re

from .config import (logger, MAPA_CONCESIONARIOS, ETIQUETAS_BLOQUE_DIAGNOSTICO, CAMPOS_CRITICOS)

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


def extraer_valores(clasificacion: dict) -> dict:
    """Devuelve los valores sugeridos para cada celda a partir del texto de los PDF
    (clasificacion["pdf_orden"] y clasificacion["pdf_garantias"])."""
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
