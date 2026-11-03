"""Procesamiento de fotos: orientación, rotación/recorte del usuario,
compresión adaptativa para no superar los 25 MB y utilidades de encuadre."""

import io

from PIL import Image as PILImage, ImageOps, ImageDraw

from .config import (logger, MB, N_FOTOS_PRIORITARIAS, NIVELES_PRIORITARIAS,
                     NIVELES_SECUNDARIAS, PROPORCION_RECUADRO)



def abrir_normalizada(datos: bytes, lado_max: int, edicion: dict = None) -> PILImage.Image:
    """Abre la imagen y aplica la misma transformación que el script
    original (orientación EXIF, conversión a RGB, rotación si es vertical),
    luego las ediciones del usuario (rotación y recorte, ver aplicar_edicion)
    y la reduce a 'lado_max' px en su lado más largo si es más grande.

    Se usa draft() para que, en JPEG, el decodificador lea directamente a
    una escala menor cuando se va a reducir: ahorra mucha RAM y tiempo con
    fotos de celular de 12-50 MP (importante en Streamlit Cloud)."""
    img = PILImage.open(io.BytesIO(datos))
    if img.format == "JPEG":
        lado_decod = lado_max
        recorte = (edicion or {}).get("crop")
        if recorte:  # si se recorta, se decodifica más grande para no perder nitidez
            fraccion = max(min(recorte[2] - recorte[0], recorte[3] - recorte[1]), 0.05)
            lado_decod = int(lado_max / fraccion)
        img.draft("RGB", (lado_decod, lado_decod))
    img = ImageOps.exif_transpose(img)

    # Transparencias (PNG con alfa) sobre fondo blanco; resto de modos a RGB.
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        fondo = PILImage.new("RGB", img.size, (255, 255, 255))
        fondo.paste(img, mask=img.split()[-1])
        img = fondo
    elif img.mode != "RGB":
        img = img.convert("RGB")

    # BYD: NO se gira automáticamente (las fotos verticales quedan verticales y se
    # centran en su recuadro). Si alguna viene mal girada se usa el botón ↻ de la página.

    img = aplicar_edicion(img, edicion)

    if max(img.size) > lado_max:
        img.thumbnail((lado_max, lado_max), PILImage.LANCZOS)
    return img

def aplicar_edicion(img: PILImage.Image, edicion: dict) -> PILImage.Image:
    """Aplica la edición del usuario sobre la imagen ya normalizada:
      - 'rot': grados en sentido horario (0, 90, 180, 270)
      - 'crop': (izq, arriba, der, abajo) como FRACCIONES 0-1 de la imagen
        ya rotada, así sirve igual para la miniatura y para la foto completa.
    """
    if not edicion:
        return img
    rot = edicion.get("rot", 0) % 360
    if rot:
        img = img.rotate(-rot, expand=True)
    recorte = edicion.get("crop")
    if recorte:
        w, h = img.size
        izq, arr, der, aba = (max(0.0, min(1.0, float(v))) for v in recorte)
        caja = (int(izq * w), int(arr * h), int(der * w), int(aba * h))
        if caja[2] - caja[0] >= 10 and caja[3] - caja[1] >= 10:
            img = img.crop(caja)
    return img

def codificar_imagen(datos: bytes, lado_max: int, calidad: int, edicion: dict = None) -> bytes:
    """Devuelve la imagen normalizada (y editada) como JPEG con el tamaño/calidad dados."""
    img = abrir_normalizada(datos, lado_max, edicion)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=calidad, optimize=True, progressive=True)
    img.close()
    return buffer.getvalue()

def comprimir_fotos_para_limite(fotos: list, presupuesto_bytes: int) -> tuple:
    """Decide la compresión de cada foto para que la suma quepa en
    'presupuesto_bytes'.

    fotos: lista de (nombre, bytes_originales, edicion) en el orden del reporte.

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
            nombre, datos, edicion = fotos[i]
            try:
                resultado[i] = codificar_imagen(datos, lado, calidad, edicion)
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

def calcular_encuadre(w: int, h: int, zoom: float = 1.0, px: float = 0.5, py: float = 0.5,
                      proporcion: float = None) -> tuple:
    """Recorte con la MISMA forma que el recuadro de la foto en el informe
    (así no sale estirada). zoom 1 = el recorte más grande que cabe en la foto;
    px/py = posición 0-1 (izquierda→derecha, arriba→abajo). Nunca se sale de la foto.
    En BYD cada recuadro tiene su propia forma: se pasa en 'proporcion' (ancho/alto)."""
    proporcion = proporcion or PROPORCION_RECUADRO[0] / PROPORCION_RECUADRO[1]
    if w / h > proporcion:
        alto, ancho = h, h * proporcion
    else:
        ancho, alto = w, w / proporcion
    ancho, alto = ancho / zoom, alto / zoom
    izq, arr = px * (w - ancho), py * (h - alto)
    return (izq / w, arr / h, (izq + ancho) / w, (arr + alto) / h)

def encuadre_relleno(w: int, h: int, crop, proporcion: float, libre: bool = False) -> tuple:
    """Recorte final para que la foto LLENE su recuadro completo sin deformarse
    (como en el informe Astara, pero sin estirar):
      - sin recorte del usuario  -> recorte centrado con la forma del recuadro;
      - con recorte del usuario  -> se respeta, ajustándolo a la forma del recuadro
        alrededor de su centro (p. ej. si la foto cambió de posición);
      - 'libre' (proporción libre elegida a mano) -> se respeta tal cual."""
    if not crop:
        return calcular_encuadre(w, h, proporcion=proporcion)
    l, t, r, b = (max(0.0, min(1.0, float(v))) for v in crop)
    if libre:
        return (l, t, r, b)
    cw, ch = (r - l) * w, (b - t) * h
    if cw < 2 or ch < 2:
        return calcular_encuadre(w, h, proporcion=proporcion)
    if abs(cw / ch - proporcion) / proporcion < 0.01:
        return (l, t, r, b)
    if cw / ch > proporcion:
        cw = ch * proporcion
    else:
        ch = cw / proporcion
    cx, cy = (l + r) / 2 * w, (t + b) / 2 * h
    x0 = min(max(cx - cw / 2, 0), w - cw)
    y0 = min(max(cy - ch / 2, 0), h - ch)
    return (x0 / w, y0 / h, (x0 + cw) / w, (y0 + ch) / h)


def tamano_orientado(datos: bytes, rot: int = 0) -> tuple:
    """Ancho/alto de la foto tal como se ve (EXIF + rotación del usuario), sin decodificarla."""
    with PILImage.open(io.BytesIO(datos)) as im:
        w, h = im.size
        try:
            if im.getexif().get(274) in (5, 6, 7, 8):
                w, h = h, w
        except Exception:
            pass
    if rot % 180:
        w, h = h, w
    return w, h


def dibujar_marco(img: PILImage.Image, crop) -> PILImage.Image:
    """Oscurece lo que queda fuera del recorte y dibuja un marco amarillo."""
    w, h = img.size
    caja = (crop[0] * w, crop[1] * h, crop[2] * w - 1, crop[3] * h - 1)
    mascara = PILImage.new("L", img.size, 150)
    ImageDraw.Draw(mascara).rectangle(caja, fill=0)
    salida = PILImage.composite(PILImage.new("RGB", img.size, (0, 0, 0)), img, mascara)
    ImageDraw.Draw(salida).rectangle(caja, outline=(255, 192, 0), width=4)
    return salida
