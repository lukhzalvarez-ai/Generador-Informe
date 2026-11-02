"""
Reporte Técnico Astara — App Streamlit (interfaz)
==================================================
Genera 'Reporte de falla Astara.xlsx' a partir de la Orden de Trabajo, el
Trabajo en Garantías y las fotos, todo subido en un solo campo.

Estructura del proyecto:
  app.py                 ← esta pantalla (Streamlit): carga, fotos, formulario, descarga
  reporte/config.py      ← constantes: celdas, mapas, límites de peso, formatos
  reporte/pdf.py         ← lectura de los PDF y valores sugeridos
  reporte/archivos.py    ← clasificación de archivos subidos e identificadores
  reporte/imagenes.py    ← rotar/recortar/comprimir fotos (máx. 25 MB)
  reporte/excel.py       ← armado del Excel desde la plantilla
  reporte/componentes.py ← recortador y ordenador de fotos (HTML propio)

Rendimiento: lo costoso (leer PDF, miniaturas) se guarda en caché usando una
"huella" rápida de cada archivo, así cada clic no vuelve a procesar los
megas de todas las fotos.
"""

import io
import base64

import streamlit as st
import pandas as pd

from reporte.config import (
    logger, MB, LIMITE_REPORTE_BYTES, N_FOTOS_PRIORITARIAS, CELDAS_EXCEL, ETIQUETAS_CAMPOS,
    EXTENSIONES_IMAGEN, TEMPLATE_FILENAME, PROPORCION_RECUADRO,
)
from reporte.pdf import extraer_texto_pdf, extraer_valores, validar_campos_criticos
from reporte.archivos import (
    clasificar_archivos_subidos, clave_foto, huella, _es_imagen_valida,
)
from reporte.imagenes import (
    abrir_normalizada, aplicar_edicion, calcular_celda_imagen, calcular_encuadre, dibujar_marco,
)
from reporte.excel import cargar_plantilla_predeterminada, generar_reporte
from reporte.componentes import recortador as componente_recortador, ordenador as componente_ordenador


# --------------------------------------------------------------------------
# Caché: cada archivo se identifica por su huella (nombre + tamaño + muestra
# de su contenido). Los bytes van en parámetros con "_" para que Streamlit
# NO los recorra completos al buscar en la caché.
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=50)
def _texto_pdf(huella_archivo: str, _datos: bytes, nombre: str) -> str:
    return extraer_texto_pdf(_datos, nombre)


def leer_pdf_cache(archivo) -> str:
    return _texto_pdf(huella(archivo), archivo.getvalue(), archivo.name)


@st.cache_data(show_spinner=False, max_entries=500)
def _imagen_valida(huella_archivo: str, _datos: bytes) -> bool:
    return _es_imagen_valida(_datos)


def imagen_valida_cache(archivo) -> bool:
    return _imagen_valida(huella(archivo), archivo.getvalue())


@st.cache_data(show_spinner=False, max_entries=600)
def _miniatura(huella_archivo: str, rot: int, crop, lado: int, _datos: bytes) -> bytes:
    img = abrir_normalizada(_datos, lado, {"rot": rot, "crop": crop})
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=80)
    return buffer.getvalue()


def miniatura_jpeg(archivo, rot: int, crop, lado: int) -> bytes:
    return _miniatura(huella(archivo), rot, tuple(crop) if crop else None, lado, archivo.getvalue())


@st.cache_data(show_spinner=False, max_entries=300)
def _base_editor(huella_archivo: str, rot: int, _datos: bytes):
    """Versión liviana (máx. 700 px) de la foto, rotada y sin recortar, para el
    editor y para calcular encuadres."""
    return abrir_normalizada(_datos, 700, {"rot": rot})


def base_editor(archivo, rot: int):
    return _base_editor(huella(archivo), rot, archivo.getvalue())


@st.cache_data(show_spinner=False, max_entries=100)
def _data_url(huella_archivo: str, rot: int, _datos: bytes) -> str:
    img = _base_editor(huella_archivo, rot, _datos)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode()


def data_url_editor(archivo, rot: int) -> str:
    return _data_url(huella(archivo), rot, archivo.getvalue())


# --------------------------------------------------------------------------
# Interfaz Streamlit — colores institucionales Astara
#   Azul marino  #002060   ·   Amarillo  #FFC000
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Reporte Técnico Astara",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
:root {
    --astara-navy: #002060;
    --astara-navy-dark: #001640;
    --astara-bg: #0D1A36;
    --astara-surface: #142449;
    --astara-border: #2A3F70;
    --astara-gold: #FFC000;
    --astara-gold-dark: #D9A300;
    --astara-text: #E6ECF7;
}

/* Fondo general oscuro (azul noche), con una franja azul institucional arriba */
.stApp {
    background:
        linear-gradient(180deg, var(--astara-navy) 0px, var(--astara-navy) 230px, var(--astara-bg) 230px, var(--astara-bg) 100%);
    color: var(--astara-text);
}
[data-testid="stHeader"] { background: transparent; }
.stApp p, .stApp li, .stApp label, .stApp span, .stApp h1, .stApp h2,
.stApp h3, .stApp h4, .stApp h5, [data-testid="stMarkdownContainer"] {
    color: var(--astara-text);
}

/* Barra lateral azul marino con texto claro y acentos amarillos */
[data-testid="stSidebar"] {
    background: var(--astara-navy);
    border-right: 4px solid var(--astara-gold);
}
[data-testid="stSidebar"] * { color: #FFFFFF !important; }
[data-testid="stSidebar"] h3 { color: var(--astara-gold) !important; }
[data-testid="stSidebar"] code {
    background: rgba(255, 192, 0, 0.15) !important;
    color: var(--astara-gold) !important;
}
[data-testid="stSidebar"] hr { border-color: rgba(255, 192, 0, 0.4); }

/* Encabezado: barra amarilla con texto azul (igual que los reportes) */
.astara-hero {
    background: var(--astara-gold);
    padding: 1.4rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.4rem;
    border-bottom: 6px solid var(--astara-gold-dark);
    box-shadow: 0 10px 28px rgba(0, 0, 0, 0.25);
}
.astara-hero h1 {
    color: var(--astara-navy) !important;
    font-size: 2rem;
    font-weight: 800;
    margin: 0;
    letter-spacing: 0.3px;
}
.astara-hero p {
    color: var(--astara-navy) !important;
    margin: 0.35rem 0 0 0;
    font-size: 0.97rem;
    opacity: 0.9;
}

/* Etiquetas de archivos detectados */
.astara-badge {
    display: inline-block;
    background: var(--astara-surface);
    color: var(--astara-text) !important;
    border: 1.5px solid var(--astara-border);
    border-radius: 999px;
    padding: 3px 12px;
    font-size: 0.8rem;
    font-weight: 600;
    margin-right: 6px;
    margin-bottom: 6px;
}
.astara-badge.ok { background: var(--astara-navy); border-color: #4C6FB8; color: #FFFFFF !important; }
.astara-badge.warn { background: rgba(255, 192, 0, 0.10); border-color: var(--astara-gold); color: var(--astara-gold) !important; }
.astara-badge.missing { background: transparent; border-color: #4B5563; color: #9CA3AF !important; }

/* Títulos de sección con acento amarillo */
.stApp h4 {
    color: #FFFFFF !important;
    border-left: 5px solid var(--astara-gold);
    padding-left: 0.6rem;
}

/* Textos secundarios (captions) un poco más tenues */
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * { color: #A9B6D3 !important; }

/* Zona de carga */
div[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed var(--astara-gold) !important;
    background: var(--astara-surface) !important;
}
div[data-testid="stFileUploaderDropzone"] * { color: var(--astara-text) !important; }

/* Formulario de revisión */
[data-testid="stForm"] {
    background: var(--astara-surface);
    border: 1px solid var(--astara-border) !important;
    border-top: 4px solid var(--astara-gold) !important;
    border-radius: 12px;
}
[data-testid="stForm"] textarea, [data-testid="stForm"] input {
    background: #0B1530 !important;
    color: var(--astara-text) !important;
}
[data-testid="stForm"] label p { font-weight: 600; }

/* Botones primarios: amarillo con texto azul */
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"],
.stFormSubmitButton > button,
button[data-testid="stBaseButton-primaryFormSubmit"],
button[data-testid="stBaseButton-primary"] {
    background: var(--astara-gold) !important;
    color: var(--astara-navy) !important;
    border: 2px solid var(--astara-gold-dark) !important;
    font-weight: 700 !important;
}
.stButton > button[kind="primary"] *,
.stDownloadButton > button[kind="primary"] *,
.stFormSubmitButton > button *,
button[data-testid="stBaseButton-primaryFormSubmit"] *,
button[data-testid="stBaseButton-primary"] * { color: var(--astara-navy) !important; }
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover,
.stFormSubmitButton > button:hover,
button[data-testid="stBaseButton-primaryFormSubmit"]:hover,
button[data-testid="stBaseButton-primary"]:hover {
    background: var(--astara-navy) !important;
    border-color: var(--astara-gold) !important;
}
.stButton > button[kind="primary"]:hover *,
.stDownloadButton > button[kind="primary"]:hover *,
.stFormSubmitButton > button:hover *,
button[data-testid="stBaseButton-primaryFormSubmit"]:hover *,
button[data-testid="stBaseButton-primary"]:hover * { color: var(--astara-gold) !important; }

/* Indicadores de peso */
[data-testid="stMetric"] {
    background: var(--astara-surface);
    border: 1px solid var(--astara-border);
    border-top: 4px solid var(--astara-gold);
    border-radius: 10px;
    padding: 0.7rem 1rem;
}
[data-testid="stMetricValue"], [data-testid="stMetricValue"] * { color: var(--astara-gold) !important; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

st.markdown(
    """
    <div class="astara-hero">
        <h1>🚗 Reporte Técnico Astara</h1>
        <p>Generador automático del Reporte de Falla a partir de la Orden de Trabajo,
        el Trabajo en Garantías y las fotografías del vehículo.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

PLANTILLA_PREDETERMINADA_DISPONIBLE = cargar_plantilla_predeterminada() is not None

with st.sidebar:
    st.markdown("### ℹ️ Cómo usar la app")
    st.markdown(
        """
        1. Arrastra los archivos al campo de carga (ninguno es
           obligatorio, usa los que tengas a mano):
           - `Orden de Trabajo.pdf`
           - `Trabajo en Garantías.pdf`
           - Fotos del vehículo
        2. Los PDF **deben llamarse exactamente así**. Cualquier otro
           PDF, Excel, Word, etc. se ignora automáticamente.
        3. La plantilla del **Reporte Técnico Astara ya viene integrada**.
        4. Revisa el formulario: viene pre-llenado con lo leído de los PDF.
           Completa la **queja**, el **diagnóstico** y el **trabajo
           realizado** sin límite de caracteres.
        5. En la sección de fotos puedes ordenarlas arrastrándolas, dejar
           fuera las que sobren o estén repetidas (🗑️), rotarlas (↺ ↻)
           y recortarlas (✂️). Las fotos idénticas se excluyen solas.
        6. Genera el informe y descárgalo listo.
        7. Para hacer otro, pulsa **🔄 Nuevo informe** (arriba a la derecha)
           y la página queda vacía.
        """
    )
    st.caption("Formatos de foto aceptados: " + ", ".join(sorted(EXTENSIONES_IMAGEN)).upper())
    st.divider()
    st.markdown("### 📦 Peso del reporte")
    st.markdown(
        f"""
        El reporte **nunca supera {LIMITE_REPORTE_BYTES // MB} MB**.
        - Las fotos se ordenan **por nombre** (1, 2, 3… y luego el resto).
        - Las **{N_FOTOS_PRIORITARIAS} primeras** se guardan siempre en alta calidad:
          nombra **"1"** y **"2"** las más importantes (ej. VIN y tablero).
        - Las demás solo se comprimen si es necesario para no pasar el límite.
        """
    )
    if not PLANTILLA_PREDETERMINADA_DISPONIBLE:
        st.error(f"No se encontró '{TEMPLATE_FILENAME}' junto a la app. Agrégalo al repositorio.")

def nuevo_informe() -> None:
    """Deja la página vacía para empezar otro informe: borra todo lo guardado en
    la sesión (textos, recortes, orden, exclusiones...). El cargador de archivos
    no se puede vaciar directamente, así que se le cambia la clave para que
    Streamlit cree uno nuevo, vacío."""
    siguiente = st.session_state.get("n_cargador", 0) + 1
    st.session_state.clear()
    st.session_state["n_cargador"] = siguiente


_col_carga, _col_reset = st.columns([5, 1])
with _col_reset:
    _popover = getattr(st, "popover", None)
    if _popover is not None:
        with _popover("🔄 Nuevo informe", use_container_width=True):
            st.markdown("¿Empezar un informe nuevo? Se quitarán **todos los archivos, textos y "
                        "ajustes de fotos** de la página.")
            st.button("Sí, dejar todo vacío", type="primary", on_click=nuevo_informe,
                      use_container_width=True, key="confirmar_nuevo_informe")
    else:
        st.button("🔄 Nuevo informe", on_click=nuevo_informe, use_container_width=True,
                  key="confirmar_nuevo_informe", help="Quita todos los archivos, textos y ajustes")

# Sin filtro de tipo en el cargador: así el usuario puede arrastrar una
# carpeta completa sin errores, y la app descarta lo que no corresponde.
with _col_carga:
    archivos_subidos = st.file_uploader(
        "Adjunta aquí: 'Orden de Trabajo.pdf', 'Trabajo en Garantías.pdf' y/o fotos (todo opcional; lo demás se ignora)",
        accept_multiple_files=True,
        key=f"cargador_{st.session_state.get('n_cargador', 0)}",
    )

clasificacion = clasificar_archivos_subidos(archivos_subidos or [], leer_pdf=leer_pdf_cache,
                                             es_imagen=imagen_valida_cache)


# --------------------------------------------------------------------------
# Orden y selección de fotos (mover / quitar del informe)
# --------------------------------------------------------------------------
def _sincronizar_fotos(imagenes: list) -> tuple:
    """Mantiene en la sesión el orden elegido y las fotos excluidas.
    - Fotos nuevas se agregan al final (en orden por nombre).
    - Fotos idénticas (mismo contenido, otro nombre) se excluyen solas UNA vez;
      si el usuario la vuelve a incluir, se respeta.
    Devuelve (fotos_del_informe_en_orden, fotos_excluidas, por_clave)."""
    import hashlib
    por_clave = {clave_foto(a): a for a in imagenes}
    claves = list(por_clave)
    orden = [c for c in st.session_state.get("orden_fotos", []) if c in por_clave]
    orden += [c for c in claves if c not in orden]
    st.session_state["orden_fotos"] = orden

    excluidas = st.session_state.setdefault("fotos_excluidas", {})   # clave -> motivo
    for c in list(excluidas):
        if c not in por_clave:
            del excluidas[c]
    hashes = st.session_state.setdefault("hash_fotos", {})            # clave -> md5
    revisadas = st.session_state.setdefault("fotos_revisadas_dup", set())
    vistos = {}
    for c in orden:
        if c not in hashes:
            hashes[c] = hashlib.md5(por_clave[c].getvalue()).hexdigest()
        h = hashes[c]
        if h in vistos and c not in revisadas and c not in excluidas:
            excluidas[c] = f"duplicada: idéntica a {por_clave[vistos[h]].name}"
        vistos.setdefault(h, c)
        revisadas.add(c)

    incluidas = [por_clave[c] for c in orden if c not in excluidas]
    fuera = [por_clave[c] for c in orden if c in excluidas]
    return incluidas, fuera, por_clave


def _claves_incluidas() -> list:
    excl = st.session_state.get("fotos_excluidas", {})
    return [c for c in st.session_state.get("orden_fotos", []) if c not in excl]


def _guardar_orden(incluidas: list) -> None:
    excl = st.session_state.get("fotos_excluidas", {})
    resto = [c for c in st.session_state.get("orden_fotos", []) if c in excl]
    st.session_state["orden_fotos"] = incluidas + resto


def mover_foto(clave: str, delta: int) -> None:
    inc = _claves_incluidas()
    if clave not in inc:
        return
    i = inc.index(clave)
    j = max(0, min(len(inc) - 1, i + delta))
    inc.insert(j, inc.pop(i))
    _guardar_orden(inc)


def foto_al_inicio(clave: str) -> None:
    inc = _claves_incluidas()
    if clave in inc:
        inc.remove(clave)
        _guardar_orden([clave] + inc)


def excluir_foto(clave: str) -> None:
    st.session_state.setdefault("fotos_excluidas", {})[clave] = "quitada a mano"
    if st.session_state.get("foto_en_recorte") == clave:
        st.session_state.pop("foto_en_recorte", None)


def incluir_foto(clave: str) -> None:
    st.session_state.setdefault("fotos_excluidas", {}).pop(clave, None)
    inc = _claves_incluidas()            # vuelve al final del informe
    if clave in inc:
        inc.remove(clave)
        _guardar_orden(inc + [clave])


def ordenar_por_nombre(imagenes: list) -> None:
    st.session_state["orden_fotos"] = [clave_foto(a) for a in imagenes]   # ya vienen en orden natural


FOTOS_INFORME, FOTOS_FUERA, _ = _sincronizar_fotos(clasificacion["imagenes"])

st.markdown("#### 📎 Archivos detectados")
badges = []
if PLANTILLA_PREDETERMINADA_DISPONIBLE:
    badges.append('<span class="astara-badge ok">Plantilla: ✔ integrada</span>')
else:
    badges.append('<span class="astara-badge missing">Plantilla: ✘ no disponible</span>')
badges.append(
    f'<span class="astara-badge {"ok" if clasificacion["pdf_orden"] else "warn"}">'
    f'Orden de Trabajo: {clasificacion["pdf_orden_nombre"] or "no adjuntada"}</span>'
)
badges.append(
    f'<span class="astara-badge {"ok" if clasificacion["pdf_garantias"] else "warn"}">'
    f'Trabajo en Garantías: {clasificacion["pdf_garantias_nombre"] or "no adjuntada"}</span>'
)
badges.append(
    f'<span class="astara-badge {"ok" if FOTOS_INFORME else "warn"}">'
    f'Fotos: {len(FOTOS_INFORME)}'
    + (f' (+{len(FOTOS_FUERA)} fuera del informe)' if FOTOS_FUERA else '') + '</span>'
)
st.markdown(" ".join(badges), unsafe_allow_html=True)
st.caption("Nada es obligatorio: lo que no adjuntes simplemente queda vacío en el reporte para completarlo a mano.")

if clasificacion["ignorados"]:
    n_ign = len(clasificacion["ignorados"])
    with st.expander(f"🚫 {n_ign} archivo(s) ignorado(s) por no corresponder (clic para ver)"):
        st.caption(
            "Solo se usan los PDF llamados exactamente 'Orden de Trabajo' y 'Trabajo en Garantías', "
            "y las fotos. Si un PDF que necesitas aparece aquí, cámbiale el nombre y vuelve a subirlo."
        )
        st.markdown("\n".join(f"- {n}" for n in clasificacion["ignorados"]))

# --------------------------------------------------------------------------
# Fotos: vista previa, rotar y recortar antes de generar
# --------------------------------------------------------------------------
if "ediciones" not in st.session_state:
    st.session_state["ediciones"] = {}   # clave_foto -> {"rot": int, "crop": tuple|None}
EDICIONES = st.session_state["ediciones"]


def _edicion_de(clave: str) -> dict:
    return EDICIONES.get(clave, {"rot": 0, "crop": None})


def rotar_foto(clave: str, grados: int) -> None:
    ed = dict(_edicion_de(clave))
    ed["rot"] = (ed.get("rot", 0) + grados) % 360
    ed["crop"] = None  # el recorte se definió sobre la orientación anterior
    ed["ajuste"] = None
    EDICIONES[clave] = ed


def abrir_recorte(clave: str) -> None:
    st.session_state["foto_en_recorte"] = clave


def cerrar_recorte() -> None:
    st.session_state.pop("foto_en_recorte", None)


def guardar_recorte(clave: str, crop, ajuste: dict = None) -> None:
    ed = dict(_edicion_de(clave))
    ed["crop"] = crop
    ed["ajuste"] = ajuste   # parámetros usados, para reabrir el editor igual
    EDICIONES[clave] = ed
    cerrar_recorte()


def restablecer_foto(clave: str) -> None:
    EDICIONES.pop(clave, None)
    cerrar_recorte()


def encuadrar_todas(archivos: list) -> None:
    """Aplica a todas las fotos SIN recorte propio un encuadre centrado con la
    forma del recuadro del informe. Las que ya se recortaron a mano no se tocan."""
    for archivo in archivos:
        clave = clave_foto(archivo)
        ed = dict(_edicion_de(clave))
        if ed.get("crop"):
            continue
        try:
            w, h = base_editor(archivo, ed.get("rot", 0)).size
        except Exception:
            continue
        ed["crop"] = calcular_encuadre(w, h)
        ed["ajuste"] = {"modo": "encuadre", "zoom": 1.0, "px": 0.5, "py": 0.5}
        EDICIONES[clave] = ed


def quitar_recortes_todas(archivos: list) -> None:
    for archivo in archivos:
        clave = clave_foto(archivo)
        if clave in EDICIONES:
            EDICIONES[clave] = {**EDICIONES[clave], "crop": None, "ajuste": None}


# ---- Editor de recorte (usa el recortador propio de reporte/componentes.py) ----
def _editor_recorte(archivo) -> None:
    """Editor de recorte de UNA foto (en ventana emergente): recuadro que se
    agranda/achica desde las esquinas y se mueve arrastrándolo, manteniendo
    la forma del recuadro del informe (salvo que se elija 'Proporción libre')."""
    clave = clave_foto(archivo)
    ed = _edicion_de(clave)
    rot = ed.get("rot", 0)
    base = base_editor(archivo, rot)
    w, h = base.size

    libre = st.checkbox(
        "Proporción libre", value=(ed.get("ajuste") or {}).get("modo") == "libre_arrastre",
        key=f"libre_{clave}",
        help="Por defecto el recuadro mantiene la forma de la foto en el informe, "
             "así no sale estirada. Marca esto para recortar con cualquier forma.",
    )
    st.caption("Arrastra las **esquinas** del recuadro amarillo para agrandarlo o achicarlo y "
               "arrástralo desde el **centro** para moverlo. Luego pulsa *Aplicar recorte*.")

    recortador = componente_recortador()
    crop = None
    if recortador is not None:
        valor = recortador(
            src=data_url_editor(archivo, rot), w=w, h=h,
            aspect=None if libre else PROPORCION_RECUADRO[0] / PROPORCION_RECUADRO[1],
            inicial=list(ed["crop"]) if ed.get("crop") else None,
            clave=f"{clave}|{rot}|{libre}", alto_max=520,
            key=f"recortador_{clave}_{rot}_{libre}", default=None,
        )
        if isinstance(valor, dict) and all(k in valor for k in ("l", "t", "r", "b")):
            crop = (valor["l"], valor["t"], valor["r"], valor["b"])
    modo_guardado = "libre_arrastre" if libre else "arrastre"

    # Plan B por si el recortador no se muestra en algún navegador
    with st.expander("Alternativa: ajustar con deslizadores"):
        usar_sl = st.checkbox("Usar deslizadores en vez del recuadro", key=f"usar_sl_{clave}")
        zoom = st.slider("🔍 Acercar", 100, 300, 100, step=5, format="%d%%", key=f"zoom_{clave}")
        px = st.slider("↔️ Mover (izquierda – derecha)", 0, 100, 50, key=f"px_{clave}")
        py = st.slider("↕️ Mover (arriba – abajo)", 0, 100, 50, key=f"py_{clave}")
        if usar_sl or crop is None:
            crop = calcular_encuadre(w, h, zoom / 100, px / 100, py / 100)
            modo_guardado = "encuadre"
            st.image(dibujar_marco(base, crop), width=420)

    if crop:
        c1, _ = st.columns([1, 2])
        c1.caption("Así quedará en el informe:")
        c1.image(aplicar_edicion(base, {"crop": crop}), use_container_width=True)

    a1, a2, a3 = st.columns(3)
    if a1.button("✅ Aplicar recorte", type="primary", use_container_width=True,
                 key=f"aplicar_{clave}", disabled=crop is None):
        guardar_recorte(clave, crop, {"modo": modo_guardado})
        st.rerun()
    if a2.button("↩️ Deshacer cambios", use_container_width=True, key=f"deshacer_{clave}",
                 help="Quita la rotación y el recorte de esta foto"):
        restablecer_foto(clave)
        st.rerun()
    if a3.button("Cancelar", use_container_width=True, key=f"cancelar_{clave}"):
        cerrar_recorte()
        st.rerun()


# Ventana emergente si la versión de Streamlit la tiene (1.34+); si no, el
# editor se muestra arriba de las fotos.
_dialogo_api = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)
if _dialogo_api is not None:
    try:
        _dialogo_recorte = _dialogo_api("✂️ Recortar foto", width="large")(_editor_recorte)
    except TypeError:
        _dialogo_recorte = _dialogo_api("✂️ Recortar foto")(_editor_recorte)
else:
    _dialogo_recorte = None


if clasificacion["imagenes"]:
    st.markdown("#### 🖼️ Fotos del informe")
    st.caption(
        "Así quedarán en el informe y en este orden. ⭐ = alta calidad (las 2 primeras). "
        "↺ ↻ para rotar · ✂️ para recortar · 🗑️ para dejarla fuera. "
        "Los cambios se aplican al generar el informe."
    )
    g1, g2, g3, _ = st.columns([1.3, 1.3, 1.3, 2])
    g1.button("📐 Encuadrar todas", on_click=encuadrar_todas, args=(FOTOS_INFORME,),
              use_container_width=True,
              help="Recorta de una vez todas las fotos a la forma del recuadro del informe (centradas), "
                   "para que no salgan estiradas. Las que ya recortaste a mano no se modifican. "
                   "Después puedes ajustar cualquiera con ✂️.")
    g2.button("🧹 Quitar todos los recortes", on_click=quitar_recortes_todas,
              args=(FOTOS_INFORME,), use_container_width=True)
    g3.button("🔤 Ordenar por nombre", on_click=ordenar_por_nombre,
              args=(clasificacion["imagenes"],), use_container_width=True,
              help="Vuelve al orden original (1, 2, 3… y luego el resto por nombre).")

    # ---- Editor de recorte (una foto a la vez) ----
    clave_abierta = st.session_state.get("foto_en_recorte")
    archivo_en_recorte = next((a for a in FOTOS_INFORME if clave_foto(a) == clave_abierta), None)
    if archivo_en_recorte is not None:
        if _dialogo_recorte is not None:
            # La ventana se mantiene abierta sola mientras se usa; se limpia la
            # marca para que no reaparezca si el usuario la cierra con la X.
            st.session_state.pop("foto_en_recorte", None)
            _dialogo_recorte(archivo_en_recorte)
        else:
            with st.container(border=True):
                st.markdown("##### ✂️ Recortar foto")
                _editor_recorte(archivo_en_recorte)

    # ---- Ordenar arrastrando ----
    ordenador = componente_ordenador()
    if ordenador is not None and len(FOTOS_INFORME) > 1:
        st.markdown("**🔀 Arrastra las fotos para cambiar el orden** "
                    "· las 2 primeras (borde amarillo) van en alta calidad")
        items_orden, firma = [], []
        for a in FOTOS_INFORME:
            ed = _edicion_de(clave_foto(a))
            mini = miniatura_jpeg(a, ed.get("rot", 0), ed.get("crop"), 220)
            items_orden.append({"k": clave_foto(a), "nombre": a.name,
                                "src": "data:image/jpeg;base64," + base64.b64encode(mini).decode()})
            firma.append(f"{clave_foto(a)}:{ed.get('rot', 0)}:{ed.get('crop')}")
        claves_actuales = [it["k"] for it in items_orden]
        valor_orden = ordenador(items=items_orden, prioritarias=N_FOTOS_PRIORITARIAS,
                                clave=str(hash(tuple(firma))), key="ordenador_fotos", default=None)
        # Se aplica solo un orden NUEVO (cada arrastre trae un id distinto); así no
        # se pisa lo que se cambie después con otros botones.
        if isinstance(valor_orden, dict) and valor_orden.get("id") != st.session_state.get("orden_id_aplicado"):
            st.session_state["orden_id_aplicado"] = valor_orden.get("id")
            nuevo = [k for k in valor_orden.get("orden", []) if k in claves_actuales]
            if len(nuevo) == len(claves_actuales) and nuevo != claves_actuales:
                _guardar_orden(nuevo)
                st.rerun()

    if not FOTOS_INFORME:
        st.info("Todas las fotos quedaron fuera del informe. Puedes volver a incluirlas abajo.")
    total = len(FOTOS_INFORME)
    for inicio in range(0, total, 4):          # filas de 4 para que el orden se lea claro
        columnas = st.columns(4)
        for j, archivo in enumerate(FOTOS_INFORME[inicio:inicio + 4]):
            i = inicio + j
            clave = clave_foto(archivo)
            ed = _edicion_de(clave)
            marcas = []
            if ed.get("rot"):
                marcas.append(f"rotada {ed['rot']}°")
            if ed.get("crop"):
                marcas.append("recortada")
            estrella = "⭐ " if i < N_FOTOS_PRIORITARIAS else ""
            with columnas[j]:
                try:
                    st.image(miniatura_jpeg(archivo, ed.get("rot", 0), ed.get("crop"), 480),
                             use_container_width=True)
                except Exception:
                    st.warning("No se pudo mostrar esta foto.")
                st.caption(f"{estrella}**{i + 1}.** {calcular_celda_imagen(i)} · {archivo.name}"
                           + (f" · ✏️ {', '.join(marcas)}" if marcas else ""))
                b1, b2, b3, b4 = st.columns(4)
                b1.button("↺", key=f"rotizq_{clave}", on_click=rotar_foto, args=(clave, -90),
                          help="Rotar 90° a la izquierda", use_container_width=True)
                b2.button("↻", key=f"rotder_{clave}", on_click=rotar_foto, args=(clave, 90),
                          help="Rotar 90° a la derecha", use_container_width=True)
                b3.button("✂️", key=f"recortar_{clave}", on_click=abrir_recorte, args=(clave,),
                          help="Recortar", use_container_width=True)
                b4.button("🗑️", key=f"excluir_{clave}", on_click=excluir_foto, args=(clave,),
                          help="Dejar fuera del informe (se puede volver a incluir)",
                          use_container_width=True)

    # ---- Fotos que quedaron fuera ----
    if FOTOS_FUERA:
        motivos = st.session_state.get("fotos_excluidas", {})
        with st.expander(f"🗑️ Fotos fuera del informe ({len(FOTOS_FUERA)}) — clic para ver o volver a incluir"):
            columnas_fuera = st.columns(6)
            for k, archivo in enumerate(FOTOS_FUERA):
                clave = clave_foto(archivo)
                ed = _edicion_de(clave)
                with columnas_fuera[k % 6]:
                    try:
                        st.image(miniatura_jpeg(archivo, ed.get("rot", 0), ed.get("crop"), 300),
                                 use_container_width=True)
                    except Exception:
                        pass
                    st.caption(f"{archivo.name} · {motivos.get(clave, '')}")
                    st.button("↩️ Incluir", key=f"incluir_{clave}", on_click=incluir_foto,
                              args=(clave,), use_container_width=True,
                              help="Vuelve a agregarla al final del informe")

# --------------------------------------------------------------------------
# Paso 2: revisar / completar los textos antes de generar
# --------------------------------------------------------------------------
valores_pdf = extraer_valores(clasificacion)
CAMPOS_DATOS_GENERALES = ("vin", "fecha_falla", "fecha_cierre_ot", "km",
                          "concesionario", "asesor_responsable", "tecnico")


def cargar_en_formulario(vals: dict) -> None:
    """Rellena los campos del formulario con lo leído de los PDF."""
    for campo in CELDAS_EXCEL:
        st.session_state[f"campo_{campo}"] = vals.get(campo, "") or ""


# Solo se recargan los campos cuando cambian los PDF; así lo que el usuario
# escribió no se pierde al agregar o quitar fotos.
firma_pdfs = hash((clasificacion["pdf_orden"] or "", clasificacion["pdf_garantias"] or ""))
if st.session_state.get("firma_pdfs") != firma_pdfs:
    cargar_en_formulario(valores_pdf)
    st.session_state["firma_pdfs"] = firma_pdfs

col_tit, col_btn = st.columns([4, 1])
with col_tit:
    st.markdown("#### ✍️ Revisa y completa el informe")
with col_btn:
    st.button("↺ Volver a lo leído de los PDF", on_click=cargar_en_formulario,
              args=(valores_pdf,), use_container_width=True)
st.caption(
    "Los campos vienen pre-llenados con lo que se pudo leer de los PDF. "
    "Completa o corrige aquí lo que haga falta (sin límite de caracteres) y luego genera el informe."
)

with st.form("form_reporte", border=True):
    st.text_area(
        "1) ¿Cuál fue la queja del cliente?",
        key="campo_queja_cliente", height=110,
        placeholder="Ej.: Cliente ingresa con el motor bloqueado, este no gira.",
    )

    st.text_area(
        "2) ¿Cuál es el resultado de su investigación? (Diagnóstico técnico)",
        key="campo_diagnostico_tecnico", height=180,
        placeholder="Ej.: Se inspecciona visualmente el motor y presenta una biela desprendida por un costado del block...",
    )
    st.text_area("3) Repuestos utilizados", key="campo_repuestos_usados", height=140)
    st.text_area(
        "4) ¿Cómo se reparó / qué se realizó para solucionar el problema? (Trabajo técnico realizado)",
        key="campo_trabajo_realizado", height=140,
        placeholder="Ej.: Se reemplaza ensamble motor y válvula EGR, se realizan pruebas de funcionamiento...",
    )

    with st.expander("🚗 Datos del vehículo y la OT (VIN, fechas, km, concesionario, asesor, técnico)"):
        cols = st.columns(3)
        for i, campo in enumerate(CAMPOS_DATOS_GENERALES):
            cols[i % 3].text_input(ETIQUETAS_CAMPOS[campo], key=f"campo_{campo}")

    generar = st.form_submit_button("⚙️ Generar Reporte Técnico Astara", type="primary",
                                    use_container_width=True)

if generar:
    valores = {campo: (st.session_state.get(f"campo_{campo}", "") or "").strip() for campo in CELDAS_EXCEL}
    faltantes = validar_campos_criticos(valores)

    n_fotos = len(FOTOS_INFORME)
    mensaje_spinner = "Armando el reporte..."
    if n_fotos > 10:
        mensaje_spinner = f"Armando el reporte y optimizando {n_fotos} fotos para no superar {LIMITE_REPORTE_BYTES // MB} MB..."
    with st.spinner(mensaje_spinner):
        try:
            excel_bytes, _, n_insertadas, info_peso = generar_reporte(valores, FOTOS_INFORME, EDICIONES)
        except Exception as e:
            st.error(f"No se pudo generar el reporte: {e}")
            logger.error(f"Error generando el reporte: {e}")
            excel_bytes = None

    if excel_bytes:
        if faltantes:
            etiquetas_faltantes = ", ".join(ETIQUETAS_CAMPOS.get(c, c) for c in faltantes)
            st.warning(f"⚠ Estos campos importantes quedaron vacíos; puedes completarlos arriba y volver a generar: {etiquetas_faltantes}")
        else:
            st.success("✅ Informe generado con todos los campos importantes completos.")

        # ---- Peso del reporte ----
        st.markdown("#### 📦 Peso del reporte")
        c1, c2, c3 = st.columns(3)
        c1.metric("Peso final del Excel", f"{info_peso['peso_final_bytes'] / MB:.2f} MB",
                  help=f"Límite: {LIMITE_REPORTE_BYTES // MB} MB")
        c2.metric("Peso original de las fotos", f"{info_peso['peso_original_fotos_bytes'] / MB:.2f} MB")
        if info_peso["n_fotos"] == 0:
            estado = "Sin fotos"
        elif info_peso["nivel_prioritarias"] == 0 and not info_peso["nivel_secundarias"]:
            estado = "Sin compresión extra"
        elif info_peso["nivel_prioritarias"] == 0:
            estado = f"Fotos 3+ comprimidas (nivel {info_peso['nivel_secundarias']})"
        else:
            estado = "Todas comprimidas"
        c3.metric("Optimización aplicada", estado)

        if not info_peso["bajo_limite"]:
            st.error(
                f"⚠ Aun con la compresión máxima el reporte pesa {info_peso['peso_final_bytes'] / MB:.2f} MB "
                f"(> {LIMITE_REPORTE_BYTES // MB} MB). Reduce la cantidad de fotos."
            )
        elif info_peso["nivel_prioritarias"] > 0:
            st.info("Por la gran cantidad de fotos también se redujeron levemente las 2 primeras para respetar el límite.")

        st.markdown("#### 👀 Vista previa de los datos importados al Excel")
        df_preview = pd.DataFrame(
            [
                {
                    "Campo": ETIQUETAS_CAMPOS.get(campo, campo),
                    "Celda": celda,
                    "Valor importado": valores.get(campo, "") or "—",
                }
                for campo, celda in CELDAS_EXCEL.items()
            ]
        )
        st.dataframe(df_preview, use_container_width=True, hide_index=True)

        if FOTOS_INFORME:
            n_editadas = sum(1 for a in FOTOS_INFORME if EDICIONES.get(clave_foto(a)))
            st.caption(f"🖼️ Fotos insertadas: {n_insertadas}/{len(FOTOS_INFORME)}"
                       + (f" · {len(FOTOS_FUERA)} dejadas fuera" if FOTOS_FUERA else "")
                       + (f" · {n_editadas} con rotación/recorte aplicado" if n_editadas else ""))
        else:
            st.info("El reporte queda sin imágenes insertadas.")

        st.markdown("#### ⬇️ Descargar reporte final")
        st.download_button(
            label="Descargar 'Reporte de falla Astara.xlsx'",
            data=excel_bytes,
            file_name="Reporte de falla Astara.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.button("🔄 Empezar otro informe (dejar la página vacía)", on_click=nuevo_informe,
                  use_container_width=True, key="nuevo_informe_final",
                  help="Descarga primero el informe: esto quita todos los archivos, textos y ajustes.")
