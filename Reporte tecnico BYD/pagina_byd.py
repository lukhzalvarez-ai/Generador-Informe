"""
Service Report BYD — página del informe BYD
============================================
Se abre desde la pantalla de inicio (app.py). Todo lo de BYD vive en esta
carpeta y NO comparte código ni datos con el informe Astara:

  Reporte tecnico BYD/
    pagina_byd.py            ← esta pantalla
    Reporte_tecnico_BYD.xlsx ← plantilla integrada
    byd/config.py            ← celdas, tablas, recuadros de fotos, límites
    byd/pdf.py               ← lectura de Orden de Trabajo y Trabajo en Garantías
    byd/archivos.py          ← clasificación de archivos y ZIP
    byd/imagenes.py          ← rotar / recortar / comprimir fotos
    byd/excel.py             ← armado del Service Report
    byd/componentes.py       ← recortador y ordenador de fotos

Todos los datos de esta página en la sesión usan el prefijo "byd_", así no se
mezclan con los del informe Astara.
"""

import io
import datetime
import sys
import base64
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))   # para importar el paquete 'byd'

from byd.config import (  # noqa: E402
    logger, MB, LIMITE_REPORTE_BYTES, N_FOTOS_PRIORITARIAS, CELDAS, ETIQUETAS, EXTENSIONES_IMAGEN,
    TEMPLATE_FILENAME, PAIS_POR_DEFECTO, RECUADROS_FOTO, CAMPOS_MANUALES,
)
from byd.pdf import extraer_texto_pdf, extraer_valores_byd  # noqa: E402
from byd.archivos import (  # noqa: E402
    clasificar_archivos_subidos, clave_foto, huella, _es_imagen_valida, es_zip, expandir_zip,
)
from byd.imagenes import (abrir_normalizada, aplicar_edicion, calcular_encuadre, dibujar_marco,  # noqa: E402
                          encuadre_relleno)
from byd.excel import (  # noqa: E402
    cargar_plantilla, generar_reporte_byd, proporcion_para_indice, nombre_recuadro,
    validar_campos_criticos, a_numero, a_fecha,
)
from byd.componentes import recortador as componente_recortador, ordenador as componente_ordenador  # noqa: E402


def K(nombre: str) -> str:
    """Clave de sesión propia de BYD (no choca con las del informe Astara)."""
    return f"byd_{nombre}"


# --------------------------------------------------------------------------
# Caché (mismo criterio que Astara: huella rápida, bytes sin recorrer)
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=50)
def _byd_texto_pdf(huella_archivo: str, _datos: bytes, nombre: str) -> str:
    return extraer_texto_pdf(_datos, nombre)


def leer_pdf_cache(archivo) -> str:
    return _byd_texto_pdf(huella(archivo), archivo.getvalue(), archivo.name)


@st.cache_data(show_spinner=False, max_entries=500)
def _byd_imagen_valida(huella_archivo: str, _datos: bytes) -> bool:
    return _es_imagen_valida(_datos)


def imagen_valida_cache(archivo) -> bool:
    return _byd_imagen_valida(huella(archivo), archivo.getvalue())


@st.cache_data(show_spinner=False, max_entries=600)
def _byd_miniatura(huella_archivo: str, rot: int, crop, lado: int, _datos: bytes) -> bytes:
    img = abrir_normalizada(_datos, lado, {"rot": rot, "crop": crop})
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=80)
    return buffer.getvalue()


def miniatura_jpeg(archivo, rot: int, crop, lado: int) -> bytes:
    return _byd_miniatura(huella(archivo), rot, tuple(crop) if crop else None, lado, archivo.getvalue())


@st.cache_data(show_spinner=False, max_entries=300)
def _byd_base_editor(huella_archivo: str, rot: int, _datos: bytes):
    return abrir_normalizada(_datos, 700, {"rot": rot})


def base_editor(archivo, rot: int):
    return _byd_base_editor(huella(archivo), rot, archivo.getvalue())


@st.cache_data(show_spinner=False, max_entries=100)
def _byd_data_url(huella_archivo: str, rot: int, _datos: bytes) -> str:
    img = _byd_base_editor(huella_archivo, rot, _datos)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode()


def data_url_editor(archivo, rot: int) -> str:
    return _byd_data_url(huella(archivo), rot, archivo.getvalue())


@st.cache_resource(show_spinner="Abriendo ZIP…", max_entries=6, ttl=3600)
def _byd_contenido_zip(huella_archivo: str, _datos: bytes, nombre: str) -> tuple:
    return expandir_zip(nombre, _datos)


def expandir_subidos(archivos: list) -> tuple:
    salida, ignorados, avisos, resumen = [], [], [], []
    for a in archivos:
        if es_zip(a):
            utiles, ign, av = _byd_contenido_zip(huella(a), a.getvalue(), a.name)
            salida += utiles
            ignorados += ign
            avisos += av
            resumen.append((a.name, len(utiles)))
        else:
            salida.append(a)
    return salida, ignorados, avisos, resumen


# --------------------------------------------------------------------------
# Estilo BYD (gris oscuro + azules suaves y gris #C1CAC9)
# --------------------------------------------------------------------------
st.markdown(
    """
<style>
:root { --byd-line: #C1CAC9; --byd-red: #3A62A6; --byd-red-dark: #2C4D86; --byd-bg: #16181C; --byd-surface: #23262B;
        --byd-border: #3C4048; --byd-text: #F1F1F1; }
.stApp { background: linear-gradient(180deg, #2B2E33 0px, #2B2E33 210px, var(--byd-bg) 210px); color: var(--byd-text); }
[data-testid="stHeader"] { background: transparent; }
.stApp p, .stApp li, .stApp label, .stApp span, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
[data-testid="stMarkdownContainer"] { color: var(--byd-text); }
[data-testid="stSidebar"] { background: #1E2024; border-right: 4px solid var(--byd-line); }
[data-testid="stSidebar"] * { color: #FFFFFF !important; }
[data-testid="stSidebar"] h3 { color: #FF4D6A !important; }
[data-testid="stSidebar"] code { background: rgba(127,168,224,.18) !important; color: #A9C6F0 !important; }
.byd-hero { background: #3A62A6; padding: 1.3rem 2rem; border-radius: 12px; margin-bottom: 1.3rem;
            border-bottom: 6px solid var(--byd-line); box-shadow: 0 10px 28px rgba(0,0,0,.35); }
.stApp .byd-hero h1, .stApp .byd-hero h1 * { color: #FFFFFF !important; margin: 0; font-size: 2rem; font-weight: 800; }
.stApp .byd-hero p { color: #DCE6F7 !important; margin: .35rem 0 0 0; }
.byd-badge { display: inline-block; background: var(--byd-surface); color: var(--byd-text) !important;
             border: 1.5px solid var(--byd-border); border-radius: 999px; padding: 3px 12px;
             font-size: .8rem; font-weight: 600; margin: 0 6px 6px 0; }
.byd-badge.ok { background: #2F3338; border-color: #8A9099; }
.byd-badge.warn { background: rgba(193,202,201,.10); border-color: #C1CAC9; color: #E6ECF7 !important; }
.stApp h4 { color: #FFFFFF !important; border-left: 5px solid var(--byd-line); padding-left: .6rem; }
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * { color: #B8BDC6 !important; }
div[data-testid="stFileUploaderDropzone"] { border: 2px dashed var(--byd-line) !important; background: var(--byd-surface) !important; }
div[data-testid="stFileUploaderDropzone"] * { color: var(--byd-text) !important; }
[data-testid="stForm"] { background: var(--byd-surface); border: 1px solid var(--byd-border) !important;
                         border-top: 4px solid var(--byd-line) !important; border-radius: 12px; }
[data-testid="stForm"] textarea, [data-testid="stForm"] input { background: #17191D !important; color: var(--byd-text) !important; }
[data-testid="stForm"] label p { font-weight: 600; }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"], .stFormSubmitButton > button,
button[data-testid="stBaseButton-primaryFormSubmit"], button[data-testid="stBaseButton-primary"] {
    background: var(--byd-red) !important; color: #FFFFFF !important; border: 2px solid var(--byd-red-dark) !important; font-weight: 700 !important; }
.stButton > button[kind="primary"] *, .stDownloadButton > button[kind="primary"] *, .stFormSubmitButton > button *,
button[data-testid="stBaseButton-primaryFormSubmit"] *, button[data-testid="stBaseButton-primary"] * { color: #FFFFFF !important; }
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover, .stFormSubmitButton > button:hover,
button[data-testid="stBaseButton-primaryFormSubmit"]:hover, button[data-testid="stBaseButton-primary"]:hover {
    background: var(--byd-red-dark) !important; border-color: #FFFFFF !important; }
[data-testid="stMetric"] { background: var(--byd-surface); border: 1px solid var(--byd-border);
                           border-top: 4px solid var(--byd-line); border-radius: 10px; padding: .7rem 1rem; }
[data-testid="stMetricValue"], [data-testid="stMetricValue"] * { color: #FF4D6A !important; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """<div class="byd-hero"><h1>🔋 Service Report BYD</h1>
    <p>Rellena el Service Report de BYD a partir de la Orden de Trabajo, el Trabajo en Garantías
    y las fotos. La plantilla en inglés ya viene integrada; los textos se copian tal cual vienen.</p></div>""",
    unsafe_allow_html=True,
)

PLANTILLA_DISPONIBLE = cargar_plantilla() is not None

with st.sidebar:
    st.markdown("### ℹ️ Cómo usar (BYD)")
    st.markdown(
        """
        1. Sube `Orden de Trabajo.pdf`, `Trabajo en Garantías.pdf` y las fotos
           (sueltos o en un **ZIP**). Nada es obligatorio.
        2. Ordena las fotos arrastrándolas. Van así en el informe:
           - **1** → *Malfunction Point* (arriba izquierda)
           - **2** → *Parts Brand (Code)* (arriba derecha)
           - **3** → *Procurement Invoice* (ancha)
           - **4 en adelante** → filas nuevas, mismo tamaño que la 3
        3. Revisa lo leído de los PDF y completa **a mano** (✍️):
           Report Number, Country, Sales Date, pieza en falla,
           mano de obra y repuestos.
        4. Genera y descarga el Service Report.
        """
    )
    st.caption("Formatos de foto: " + ", ".join(sorted(EXTENSIONES_IMAGEN)).upper())
    st.caption(f"El informe nunca supera {LIMITE_REPORTE_BYTES // MB} MB; las 2 primeras fotos van en alta calidad.")
    if not PLANTILLA_DISPONIBLE:
        st.error(f"No se encontró '{TEMPLATE_FILENAME}' en la carpeta 'Reporte tecnico BYD'.")


# --------------------------------------------------------------------------
# Carga de archivos + "Nuevo informe"
# --------------------------------------------------------------------------
def nuevo_informe() -> None:
    """Vacía SOLO lo del informe BYD (no toca el informe Astara)."""
    siguiente = st.session_state.get(K("n_cargador"), 0) + 1
    for k in list(st.session_state.keys()):
        if str(k).startswith("byd_"):
            del st.session_state[k]
    st.session_state[K("n_cargador")] = siguiente


_col_carga, _col_reset = st.columns([5, 1])
with _col_reset:
    _popover = getattr(st, "popover", None)
    if _popover is not None:
        with _popover("🔄 Nuevo informe", use_container_width=True):
            st.markdown("¿Empezar un Service Report nuevo? Se quitarán **todos los archivos, textos, "
                        "tablas y ajustes de fotos** de esta página.")
            st.button("Sí, dejar todo vacío", type="primary", on_click=nuevo_informe,
                      use_container_width=True, key=K("confirmar_nuevo"))
    else:
        st.button("🔄 Nuevo informe", on_click=nuevo_informe, use_container_width=True, key=K("confirmar_nuevo"))

with _col_carga:
    archivos_subidos = st.file_uploader(
        "Adjunta aquí: 'Orden de Trabajo.pdf', 'Trabajo en Garantías.pdf' y/o fotos, sueltos o en un ZIP "
        "(todo opcional; lo demás se ignora)",
        accept_multiple_files=True,
        key=K(f"cargador_{st.session_state.get(K('n_cargador'), 0)}"),
    )

archivos_expandidos, IGNORADOS_EN_ZIP, AVISOS_ZIP, RESUMEN_ZIPS = expandir_subidos(archivos_subidos or [])
clasificacion = clasificar_archivos_subidos(archivos_expandidos, leer_pdf=leer_pdf_cache,
                                             es_imagen=imagen_valida_cache)
clasificacion["ignorados"] += IGNORADOS_EN_ZIP


# --------------------------------------------------------------------------
# Orden y selección de fotos
# --------------------------------------------------------------------------
def _sincronizar_fotos(imagenes: list) -> tuple:
    import hashlib
    por_clave = {clave_foto(a): a for a in imagenes}
    orden = [c for c in st.session_state.get(K("orden_fotos"), []) if c in por_clave]
    orden += [c for c in por_clave if c not in orden]
    st.session_state[K("orden_fotos")] = orden
    excluidas = st.session_state.setdefault(K("fotos_excluidas"), {})
    for c in list(excluidas):
        if c not in por_clave:
            del excluidas[c]
    hashes = st.session_state.setdefault(K("hash_fotos"), {})
    revisadas = st.session_state.setdefault(K("fotos_revisadas_dup"), set())
    vistos = {}
    for c in orden:
        if c not in hashes:
            hashes[c] = hashlib.md5(por_clave[c].getvalue()).hexdigest()
        h = hashes[c]
        if h in vistos and c not in revisadas and c not in excluidas:
            excluidas[c] = f"duplicada: idéntica a {por_clave[vistos[h]].name}"
        vistos.setdefault(h, c)
        revisadas.add(c)
    return ([por_clave[c] for c in orden if c not in excluidas],
            [por_clave[c] for c in orden if c in excluidas])


def _claves_incluidas() -> list:
    excl = st.session_state.get(K("fotos_excluidas"), {})
    return [c for c in st.session_state.get(K("orden_fotos"), []) if c not in excl]


def _guardar_orden(incluidas: list) -> None:
    excl = st.session_state.get(K("fotos_excluidas"), {})
    resto = [c for c in st.session_state.get(K("orden_fotos"), []) if c in excl]
    st.session_state[K("orden_fotos")] = incluidas + resto


def excluir_foto(clave: str) -> None:
    st.session_state.setdefault(K("fotos_excluidas"), {})[clave] = "quitada a mano"
    if st.session_state.get(K("foto_en_recorte")) == clave:
        st.session_state.pop(K("foto_en_recorte"), None)


def incluir_foto(clave: str) -> None:
    st.session_state.setdefault(K("fotos_excluidas"), {}).pop(clave, None)
    inc = _claves_incluidas()
    if clave in inc:
        inc.remove(clave)
        _guardar_orden(inc + [clave])


def ordenar_por_nombre(imagenes: list) -> None:
    st.session_state[K("orden_fotos")] = [clave_foto(a) for a in imagenes]


FOTOS_INFORME, FOTOS_FUERA = _sincronizar_fotos(clasificacion["imagenes"])

st.markdown("#### 📎 Archivos detectados")
badges = [
    f'<span class="byd-badge {"ok" if PLANTILLA_DISPONIBLE else "warn"}">Plantilla BYD: '
    f'{"✔ integrada" if PLANTILLA_DISPONIBLE else "✘ no disponible"}</span>',
    f'<span class="byd-badge {"ok" if clasificacion["pdf_orden"] else "warn"}">Orden de Trabajo: '
    f'{clasificacion["pdf_orden_nombre"] or "no adjuntada"}</span>',
    f'<span class="byd-badge {"ok" if clasificacion["pdf_garantias"] else "warn"}">Trabajo en Garantías: '
    f'{clasificacion["pdf_garantias_nombre"] or "no adjuntada"}</span>',
    f'<span class="byd-badge {"ok" if FOTOS_INFORME else "warn"}">Fotos: {len(FOTOS_INFORME)}'
    + (f' (+{len(FOTOS_FUERA)} fuera del informe)' if FOTOS_FUERA else '') + '</span>',
]
st.markdown(" ".join(badges), unsafe_allow_html=True)
if RESUMEN_ZIPS:
    st.caption("📦 " + " · ".join(f"{n}: {k} archivo(s) útil(es) extraído(s)" for n, k in RESUMEN_ZIPS))
for aviso in AVISOS_ZIP:
    st.warning(f"📦 {aviso}")
if clasificacion["ignorados"]:
    with st.expander(f"🚫 {len(clasificacion['ignorados'])} archivo(s) ignorado(s) por no corresponder (clic para ver)"):
        st.caption("Solo se usan los PDF llamados exactamente 'Orden de Trabajo' y 'Trabajo en Garantías', "
                   "y las fotos. Si un PDF que necesitas aparece aquí, cámbiale el nombre y vuelve a subirlo.")
        st.markdown("\n".join(f"- {n}" for n in clasificacion["ignorados"]))


# --------------------------------------------------------------------------
# Fotos: rotar / recortar / ordenar
# --------------------------------------------------------------------------
if K("ediciones") not in st.session_state:
    st.session_state[K("ediciones")] = {}
EDICIONES = st.session_state[K("ediciones")]


def _edicion_de(clave: str) -> dict:
    return EDICIONES.get(clave, {"rot": 0, "crop": None})


def _indice_de(clave: str) -> int:
    inc = _claves_incluidas()
    return inc.index(clave) if clave in inc else 0


def rotar_foto(clave: str, grados: int) -> None:
    ed = dict(_edicion_de(clave))
    ed["rot"] = (ed.get("rot", 0) + grados) % 360
    ed["crop"] = None
    ed["ajuste"] = None
    EDICIONES[clave] = ed


def abrir_recorte(clave: str) -> None:
    st.session_state[K("foto_en_recorte")] = clave


def cerrar_recorte() -> None:
    st.session_state.pop(K("foto_en_recorte"), None)


def guardar_recorte(clave: str, crop, ajuste: dict = None) -> None:
    ed = dict(_edicion_de(clave))
    ed["crop"] = crop
    ed["ajuste"] = ajuste
    EDICIONES[clave] = ed
    cerrar_recorte()


def restablecer_foto(clave: str) -> None:
    EDICIONES.pop(clave, None)
    cerrar_recorte()


def encuadrar_todas(archivos: list) -> None:
    """Recorta cada foto (sin recorte propio) a la forma del recuadro donde va."""
    for i, archivo in enumerate(archivos):
        clave = clave_foto(archivo)
        ed = dict(_edicion_de(clave))
        if ed.get("crop"):
            continue
        try:
            w, h = base_editor(archivo, ed.get("rot", 0)).size
        except Exception:
            continue
        ed["crop"] = calcular_encuadre(w, h, proporcion=proporcion_para_indice(i))
        ed["ajuste"] = {"modo": "encuadre"}
        EDICIONES[clave] = ed


def crop_en_informe(archivo, indice: int, ed: dict):
    """Recorte con el que la foto quedará en el informe (llena su recuadro completo)."""
    try:
        w, h = base_editor(archivo, ed.get("rot", 0)).size
    except Exception:
        return ed.get("crop")
    libre = (ed.get("ajuste") or {}).get("modo") == "libre_arrastre"
    return encuadre_relleno(w, h, ed.get("crop"), proporcion_para_indice(indice), libre)


def quitar_recortes_todas(archivos: list) -> None:
    for archivo in archivos:
        clave = clave_foto(archivo)
        if clave in EDICIONES:
            EDICIONES[clave] = {**EDICIONES[clave], "crop": None, "ajuste": None}


def _editor_recorte(archivo) -> None:
    clave = clave_foto(archivo)
    ed = _edicion_de(clave)
    rot = ed.get("rot", 0)
    base = base_editor(archivo, rot)
    w, h = base.size
    indice = _indice_de(clave)
    nombre, ayuda = nombre_recuadro(indice)
    proporcion = proporcion_para_indice(indice)

    st.caption(f"Foto {indice + 1} → **{nombre}** ({ayuda}). Arrastra las **esquinas** del recuadro "
               "para agrandarlo o achicarlo y arrástralo desde el **centro** para moverlo.")
    libre = st.checkbox("Proporción libre", value=(ed.get("ajuste") or {}).get("modo") == "libre_arrastre",
                        key=K(f"libre_{clave}"),
                        help="Por defecto el recuadro tiene la forma del espacio de esta foto en el informe, "
                             "así la foto lo llena completo. Con proporción libre la foto no se deforma: "
                             "si no calza, queda centrada con bordes blancos.")
    recortador = componente_recortador()
    crop = None
    if recortador is not None:
        valor = recortador(
            src=data_url_editor(archivo, rot), w=w, h=h, aspect=None if libre else proporcion,
            inicial=list(crop_en_informe(archivo, indice, ed) or []) or None,
            clave=f"{clave}|{rot}|{libre}|{round(proporcion, 3)}", alto_max=520,
            key=K(f"recortador_{clave}_{rot}_{libre}_{round(proporcion, 3)}"), default=None,
        )
        if isinstance(valor, dict) and all(k in valor for k in ("l", "t", "r", "b")):
            crop = (valor["l"], valor["t"], valor["r"], valor["b"])
    modo = "libre_arrastre" if libre else "arrastre"
    with st.expander("Alternativa: ajustar con deslizadores"):
        usar_sl = st.checkbox("Usar deslizadores en vez del recuadro", key=K(f"usar_sl_{clave}"))
        zoom = st.slider("🔍 Acercar", 100, 300, 100, step=5, format="%d%%", key=K(f"zoom_{clave}"))
        px = st.slider("↔️ Mover (izquierda – derecha)", 0, 100, 50, key=K(f"px_{clave}"))
        py = st.slider("↕️ Mover (arriba – abajo)", 0, 100, 50, key=K(f"py_{clave}"))
        if usar_sl or crop is None:
            crop = calcular_encuadre(w, h, zoom / 100, px / 100, py / 100, proporcion=proporcion)
            modo = "encuadre"
            st.image(dibujar_marco(base, crop), width=420)
    if crop:
        c1, _ = st.columns([1, 2])
        c1.caption("Así quedará:")
        c1.image(aplicar_edicion(base, {"crop": crop}), use_container_width=True)
    a1, a2, a3 = st.columns(3)
    if a1.button("✅ Aplicar recorte", type="primary", use_container_width=True, key=K(f"aplicar_{clave}"),
                 disabled=crop is None):
        guardar_recorte(clave, crop, {"modo": modo})
        st.rerun()
    if a2.button("↩️ Deshacer cambios", use_container_width=True, key=K(f"deshacer_{clave}")):
        restablecer_foto(clave)
        st.rerun()
    if a3.button("Cancelar", use_container_width=True, key=K(f"cancelar_{clave}")):
        cerrar_recorte()
        st.rerun()


_dialogo_api = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)
if _dialogo_api is not None:
    try:
        _dialogo_recorte = _dialogo_api("✂️ Recortar foto", width="large")(_editor_recorte)
    except TypeError:
        _dialogo_recorte = _dialogo_api("✂️ Recortar foto")(_editor_recorte)
else:
    _dialogo_recorte = None


if clasificacion["imagenes"]:
    st.markdown("#### 🖼️ Fotos del Service Report")
    st.caption("Orden en el informe: **1** Malfunction Point · **2** Parts Brand (Code) · "
               "**3** Procurement Invoice · **4+** filas adicionales. Cada foto **llena su espacio completo**: "
               "se recorta centrada con la forma del recuadro (la vista previa muestra exactamente lo que "
               "saldrá). Si quieres otro encuadre usa ✂️. ↺ ↻ rotar · 🗑️ dejar fuera.")
    g2, g3, _ = st.columns([1.3, 1.3, 3.3])
    g2.button("🧹 Quitar todos los recortes", on_click=quitar_recortes_todas, args=(FOTOS_INFORME,),
              use_container_width=True, key=K("quitar_recortes"),
              help="Vuelve todas las fotos al encuadre automático (centrado).")
    g3.button("🔤 Ordenar por nombre", on_click=ordenar_por_nombre, args=(clasificacion["imagenes"],),
              use_container_width=True, key=K("ordenar_nombre"))

    clave_abierta = st.session_state.get(K("foto_en_recorte"))
    archivo_en_recorte = next((a for a in FOTOS_INFORME if clave_foto(a) == clave_abierta), None)
    if archivo_en_recorte is not None:
        if _dialogo_recorte is not None:
            st.session_state.pop(K("foto_en_recorte"), None)
            _dialogo_recorte(archivo_en_recorte)
        else:
            with st.container(border=True):
                st.markdown("##### ✂️ Recortar foto")
                _editor_recorte(archivo_en_recorte)

    ordenador = componente_ordenador()
    if ordenador is not None and len(FOTOS_INFORME) > 1:
        st.markdown("**🔀 Arrastra las fotos para cambiar el orden** · la 1 y la 2 (borde azul) van arriba en "
                    "alta calidad; la 3 en el espacio de la factura")
        items, firma = [], []
        for pos, a in enumerate(FOTOS_INFORME):
            ed = _edicion_de(clave_foto(a))
            mini = miniatura_jpeg(a, ed.get("rot", 0), crop_en_informe(a, pos, ed), 220)
            items.append({"k": clave_foto(a), "nombre": a.name,
                          "src": "data:image/jpeg;base64," + base64.b64encode(mini).decode()})
            firma.append(f"{clave_foto(a)}:{ed.get('rot', 0)}:{ed.get('crop')}")
        claves_actuales = [it["k"] for it in items]
        valor_orden = ordenador(items=items, prioritarias=N_FOTOS_PRIORITARIAS, clave=str(hash(tuple(firma))),
                                key=K("ordenador_fotos"), default=None)
        if isinstance(valor_orden, dict) and valor_orden.get("id") != st.session_state.get(K("orden_id_aplicado")):
            st.session_state[K("orden_id_aplicado")] = valor_orden.get("id")
            nuevo = [k for k in valor_orden.get("orden", []) if k in claves_actuales]
            if len(nuevo) == len(claves_actuales) and nuevo != claves_actuales:
                _guardar_orden(nuevo)
                st.rerun()

    if not FOTOS_INFORME:
        st.info("Todas las fotos quedaron fuera del informe. Puedes volver a incluirlas abajo.")
    for inicio in range(0, len(FOTOS_INFORME), 4):
        columnas = st.columns(4)
        for j, archivo in enumerate(FOTOS_INFORME[inicio:inicio + 4]):
            i = inicio + j
            clave = clave_foto(archivo)
            ed = _edicion_de(clave)
            marcas = ([f"rotada {ed['rot']}°"] if ed.get("rot") else []) + (["recortada"] if ed.get("crop") else [])
            nombre, ayuda = nombre_recuadro(i)
            with columnas[j]:
                try:
                    st.image(miniatura_jpeg(archivo, ed.get("rot", 0), crop_en_informe(archivo, i, ed), 480),
                             use_container_width=True)
                except Exception:
                    st.warning("No se pudo mostrar esta foto.")
                st.caption(f"{'⭐ ' if i < N_FOTOS_PRIORITARIAS else ''}**{i + 1}. {ayuda.capitalize()}** · "
                           f"{archivo.name}" + (f" · ✏️ {', '.join(marcas)}" if marcas else ""))
                b1, b2, b3, b4 = st.columns(4)
                b1.button("↺", key=K(f"rotizq_{clave}"), on_click=rotar_foto, args=(clave, -90),
                          help="Rotar 90° a la izquierda", use_container_width=True)
                b2.button("↻", key=K(f"rotder_{clave}"), on_click=rotar_foto, args=(clave, 90),
                          help="Rotar 90° a la derecha", use_container_width=True)
                b3.button("✂️", key=K(f"recortar_{clave}"), on_click=abrir_recorte, args=(clave,),
                          help="Recortar", use_container_width=True)
                b4.button("🗑️", key=K(f"excluir_{clave}"), on_click=excluir_foto, args=(clave,),
                          help="Dejar fuera del informe", use_container_width=True)

    if FOTOS_FUERA:
        motivos = st.session_state.get(K("fotos_excluidas"), {})
        with st.expander(f"🗑️ Fotos fuera del informe ({len(FOTOS_FUERA)}) — clic para ver o volver a incluir"):
            cols = st.columns(6)
            for k, archivo in enumerate(FOTOS_FUERA):
                clave = clave_foto(archivo)
                ed = _edicion_de(clave)
                with cols[k % 6]:
                    try:
                        st.image(miniatura_jpeg(archivo, ed.get("rot", 0), ed.get("crop"), 300), use_container_width=True)
                    except Exception:
                        pass
                    st.caption(f"{archivo.name} · {motivos.get(clave, '')}")
                    st.button("↩️ Incluir", key=K(f"incluir_{clave}"), on_click=incluir_foto, args=(clave,),
                              use_container_width=True)


# --------------------------------------------------------------------------
# Formulario: datos, textos y tablas
# --------------------------------------------------------------------------
COLS_PIEZA = {"codigo": "Malfunction Part No.", "nombre": "Malfunction Part Name"}
COLS_MO = {"codigo": "Labor Hours Code", "nombre": "Name Of Labor"}
COLS_REP = {"codigo": "Spare Part Code", "nombre": "Spare Part Name", "cantidad": "Quantity",
            "precio": "Unit Price", "total": "Total Price"}
CAMPOS_AUTO = ("vin", "model", "delivery_date", "kilometrage", "service_date")


def _tabla_vacia(columnas: dict, filas: int) -> pd.DataFrame:
    return pd.DataFrame([{t: "" for t in columnas.values()} for _ in range(filas)],
                        columns=list(columnas.values())).astype(str)


def _df_a_lista(df, columnas: dict) -> list:
    """Tabla editada -> lista de dicts; descarta filas totalmente vacías."""
    if df is None:
        return []
    inverso = {titulo: campo for campo, titulo in columnas.items()}
    res = []
    for _, fila in pd.DataFrame(df).fillna("").iterrows():
        d = {inverso[c]: str(fila[c]).strip() for c in fila.index if c in inverso}
        d = {k: ("" if v in ("None", "nan") else v) for k, v in d.items()}
        if any(d.values()):
            res.append(d)
    return res


def _completar_totales(repuestos: list) -> list:
    """Si 'Total Price' quedó vacío y hay cantidad y precio numéricos, lo calcula."""
    for r in repuestos:
        if not r.get("total"):
            cant, precio = a_numero(r.get("cantidad", "")), a_numero(r.get("precio", ""))
            if isinstance(cant, (int, float)) and isinstance(precio, (int, float)):
                total = cant * precio
                r["total"] = int(total) if float(total).is_integer() else round(total, 2)
    return repuestos


VALORES_PDF = extraer_valores_byd(clasificacion)


CAMPOS_FECHA_AUTO = ("delivery_date", "service_date")      # se leen de los PDF pero con calendario
CAMPOS_LEIDOS = CAMPOS_AUTO + ("fault_description", "troubleshooting", "repair_measure")


def _a_date(valor):
    """Texto 'dd-mm-aaaa' de los PDF -> fecha para el calendario (None si no se puede)."""
    f = a_fecha(valor)
    return f.date() if isinstance(f, datetime.datetime) else None


def _valor_leido(vals: dict, campo: str):
    v = vals.get(campo, "") or ""
    return _a_date(v) if campo in CAMPOS_FECHA_AUTO else v


def cargar_lo_leido(vals: dict) -> None:
    """Rellena SOLO los campos que vienen de los PDF (no toca lo escrito a mano)."""
    for campo in CAMPOS_LEIDOS:
        st.session_state[K(f"campo_{campo}")] = _valor_leido(vals, campo)


# Campos manuales y tablas: se crean UNA vez y se conservan aunque cambien los PDF.
# Se revisan uno por uno porque Streamlit borra el valor de un campo cuando deja de
# mostrarse (p. ej. al ir a otra página); así Country vuelve siempre a "Chile".
for _campo, _defecto in (("report_number", ""), ("country", PAIS_POR_DEFECTO), ("sales_date", None)):
    if K(f"campo_{_campo}") not in st.session_state:
        st.session_state[K(f"campo_{_campo}")] = _defecto
for _campo in CAMPOS_LEIDOS:           # si Streamlit los borró al cambiar de página, se recargan
    if K(f"campo_{_campo}") not in st.session_state:
        st.session_state[K(f"campo_{_campo}")] = _valor_leido(VALORES_PDF, _campo)
if K("df_piezas") not in st.session_state:
    st.session_state[K("df_piezas")] = _tabla_vacia(COLS_PIEZA, 3)
    st.session_state[K("df_mo")] = _tabla_vacia(COLS_MO, 3)
    st.session_state[K("df_rep")] = _tabla_vacia(COLS_REP, 3)

firma_pdfs = hash((clasificacion["pdf_orden"] or "", clasificacion["pdf_garantias"] or ""))
if st.session_state.get(K("firma_pdfs")) != firma_pdfs:
    cargar_lo_leido(VALORES_PDF)
    st.session_state[K("firma_pdfs")] = firma_pdfs

col_tit, col_btn = st.columns([4, 1])
with col_tit:
    st.markdown("#### ✍️ Revisa y completa el Service Report")
with col_btn:
    st.button("↺ Volver a lo leído de los PDF", on_click=cargar_lo_leido, args=(VALORES_PDF,),
              use_container_width=True, key=K("recargar_pdf"),
              help="Vuelve a poner lo que se leyó de los PDF. No borra lo que escribiste a mano.")
st.caption("📄 = se lee de los PDF (texto tal cual, puedes corregirlo) · ✍️ = se llena a mano "
           "(no viene en los PDF). Lo escrito a mano se conserva aunque cambies los PDF.")

with st.form(K("form"), border=True):
    st.markdown("**✍️ Datos a mano**")
    m1, m2, m3 = st.columns(3)
    m1.text_input(f"✍️ {ETIQUETAS['report_number']}", key=K("campo_report_number"))
    m2.text_input(f"✍️ {ETIQUETAS['country']}", key=K("campo_country"))
    m3.date_input(f"✍️ {ETIQUETAS['sales_date']}", key=K("campo_sales_date"), format="DD-MM-YYYY")

    st.markdown("**📄 Vehicle's Information (de los PDF)**")
    cols = st.columns(5)
    for i, campo in enumerate(CAMPOS_AUTO):
        if campo in CAMPOS_FECHA_AUTO:
            cols[i].date_input(f"📄 {ETIQUETAS[campo]}", key=K(f"campo_{campo}"), format="DD-MM-YYYY",
                               help="Se llena sola desde los PDF si viene; si no, elígela en el calendario.")
        else:
            cols[i].text_input(f"📄 {ETIQUETAS[campo]}", key=K(f"campo_{campo}"))

    st.markdown("**✍️ Malfunction Part** (máx. 3 filas)")
    ed_piezas = st.data_editor(st.session_state[K("df_piezas")], num_rows="fixed", hide_index=True,
                               use_container_width=True, key=K("ed_piezas"))

    st.text_area("📄 Fault Description (falla / queja del cliente — de la OT)",
                 key=K("campo_fault_description"), height=100)
    st.text_area("📄 Troubleshooting Ideas (diagnóstico)", key=K("campo_troubleshooting"), height=110)
    st.text_area("📄 Repair Measure (reparación — línea de garantía de la OT)",
                 key=K("campo_repair_measure"), height=100)

    st.markdown("**✍️ Labor** (máx. 3 filas)")
    ed_mo = st.data_editor(st.session_state[K("df_mo")], num_rows="fixed", hide_index=True,
                           use_container_width=True, key=K("ed_mo"))

    st.markdown("**✍️ Spare Parts** · agrega filas con ➕ (si hay más de 3, el informe agrega filas). "
                "Si dejas *Total Price* vacío se calcula como cantidad × precio unitario.")
    ed_rep = st.data_editor(st.session_state[K("df_rep")], num_rows="dynamic", hide_index=True,
                            use_container_width=True, key=K("ed_rep"))

    generar = st.form_submit_button("⚙️ Generar Service Report BYD", type="primary", use_container_width=True)

if generar:
    valores = {}
    for campo in CELDAS:
        v = st.session_state.get(K(f"campo_{campo}"), "")
        valores[campo] = v.strip() if isinstance(v, str) else v       # Sales Date es una fecha
    valores["country"] = valores.get("country") or PAIS_POR_DEFECTO      # nunca queda vacío
    valores["piezas_falla"] = _df_a_lista(ed_piezas, COLS_PIEZA)[:3]
    valores["mano_obra"] = _df_a_lista(ed_mo, COLS_MO)[:3]
    valores["repuestos"] = _completar_totales(_df_a_lista(ed_rep, COLS_REP))
    faltantes = validar_campos_criticos(valores)

    n_fotos = len(FOTOS_INFORME)
    with st.spinner("Armando el Service Report…" if n_fotos <= 10 else
                    f"Armando el Service Report y optimizando {n_fotos} fotos…"):
        try:
            excel_bytes, _, n_insertadas, info_peso = generar_reporte_byd(valores, FOTOS_INFORME, EDICIONES)
        except Exception as e:
            st.error(f"No se pudo generar el Service Report: {e}")
            logger.error(f"Error generando el Service Report BYD: {e}")
            excel_bytes = None

    if excel_bytes:
        if faltantes:
            st.warning("⚠ Estos campos importantes quedaron vacíos; puedes completarlos arriba y volver a generar: "
                       + ", ".join(ETIQUETAS.get(c, c) for c in faltantes))
        else:
            st.success("✅ Service Report generado con todos los campos importantes completos.")
        if 0 < n_fotos < len(RECUADROS_FOTO):
            st.info(f"Hay {n_fotos} foto(s): quedaron espacios de foto vacíos en el informe "
                    f"(tiene {len(RECUADROS_FOTO)}).")

        st.markdown("#### 📦 Peso del informe")
        c1, c2, c3 = st.columns(3)
        c1.metric("Peso final del Excel", f"{info_peso['peso_final_bytes'] / MB:.2f} MB",
                  help=f"Límite: {LIMITE_REPORTE_BYTES // MB} MB")
        c2.metric("Peso original de las fotos", f"{info_peso['peso_original_fotos_bytes'] / MB:.2f} MB")
        c3.metric("Fotos insertadas", f"{n_insertadas}/{n_fotos}")
        if not info_peso["bajo_limite"]:
            st.error(f"⚠ Aun con la compresión máxima pesa {info_peso['peso_final_bytes'] / MB:.2f} MB. "
                     "Reduce la cantidad de fotos.")

        st.markdown("#### 👀 Vista previa de los datos")
        st.dataframe(pd.DataFrame([{"Campo": ("✍️ " if c in CAMPOS_MANUALES else "📄 ") + ETIQUETAS[c],
                                    "Valor": str(valores.get(c) or "—")} for c in CELDAS]),
                     use_container_width=True, hide_index=True)
        st.caption(f"Tablas: {len(valores['piezas_falla'])} pieza(s) en falla · {len(valores['mano_obra'])} "
                   f"línea(s) de mano de obra · {len(valores['repuestos'])} repuesto(s)")

        numero = valores.get("report_number", "").replace("/", "-").strip()
        nombre_archivo = f"Service Report BYD{' - ' + numero if numero else ''}.xlsx"
        st.markdown("#### ⬇️ Descargar")
        st.download_button(f"Descargar '{nombre_archivo}'", data=excel_bytes, file_name=nombre_archivo,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", use_container_width=True, key=K("descargar"))
        st.button("🔄 Empezar otro Service Report (dejar la página vacía)", on_click=nuevo_informe,
                  use_container_width=True, key=K("nuevo_final"))
