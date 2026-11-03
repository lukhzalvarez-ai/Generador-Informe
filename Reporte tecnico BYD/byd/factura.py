"""Factura de compra (PDF) -> imagen JPEG para el recuadro
"Photo Of Parts Procurement Invoice" del Service Report BYD."""

import io

from PIL import Image as PILImage

from .config import logger, FACTURA_ANCHO_PX, FACTURA_CALIDAD_JPEG, FACTURA_MAX_PAGINAS


def factura_a_imagenes(datos_pdf: bytes) -> list:
    """Cada página de la factura (máx. FACTURA_MAX_PAGINAS) como JPEG, con su forma real.
    Devuelve [] si el PDF no se puede leer (el informe se genera igual, sin factura)."""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        logger.error("Falta 'pypdfium2' (requirements.txt): no se puede convertir la factura.")
        return []
    imagenes = []
    try:
        pdf = pdfium.PdfDocument(datos_pdf)
        for i in range(min(len(pdf), FACTURA_MAX_PAGINAS)):
            pagina = pdf[i]
            ancho_pt = pagina.get_width() or 595
            img = pagina.render(scale=FACTURA_ANCHO_PX / ancho_pt).to_pil().convert("RGB")
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=FACTURA_CALIDAD_JPEG, optimize=True)
            imagenes.append(buffer.getvalue())
            pagina.close()
        pdf.close()
    except Exception as e:
        logger.error(f"No se pudo convertir la factura en imagen: {e}")
        return []
    return imagenes


def miniatura(jpeg: bytes, lado: int = 500) -> bytes:
    with PILImage.open(io.BytesIO(jpeg)) as im:
        im = im.convert("RGB")
        im.thumbnail((lado, lado))
        buffer = io.BytesIO()
        im.save(buffer, format="JPEG", quality=80)
        return buffer.getvalue()
