"""
Generador de Reportes Técnicos — pantalla de inicio
====================================================
Al abrir la app se elige qué informe rellenar; cada uno vive en su propia
página y no comparte código ni datos con el otro:

  app.py                                  ← esta pantalla (inicio + menú)
  informe_astara.py  + reporte/           ← Reporte de Falla Astara
  Reporte tecnico BYD/pagina_byd.py + byd/ ← Service Report BYD
  assets/                                 ← logos de la pantalla de inicio

Requiere Streamlit 1.40 o superior (st.navigation).
"""

import base64
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="Reportes Técnicos",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

RAIZ = Path(__file__).resolve().parent


def _logo_base64(nombre: str) -> str:
    """Logo de la carpeta assets/ como data-URL (si falta, se muestra solo el texto)."""
    ruta = RAIZ / "assets" / nombre
    try:
        return "data:image/png;base64," + base64.b64encode(ruta.read_bytes()).decode()
    except OSError:
        return ""


def _tarjeta(clase: str, logo: str, titulo: str, detalle: str) -> str:
    img = f'<img src="{logo}" alt="{titulo}">' if logo else f"<h2>{titulo}</h2>"
    return (f'<div class="tarjeta {clase}"><div class="logo">{img}</div>'
            f'<div class="pie"><b>{titulo}</b><span>{detalle}</span></div></div>')


def pagina_inicio() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: linear-gradient(180deg, #002060 0px, #002060 210px, #0D1A36 210px); }
        .inicio-titulo h1 { color: #FFFFFF !important; margin-bottom: 0; }
        .inicio-titulo p  { color: #C9D3EA !important; margin-top: .3rem; }
        .tarjeta { border-radius: 16px; overflow: hidden; background: #FFFFFF;
                   box-shadow: 0 10px 28px rgba(0,0,0,.35); margin-bottom: .7rem;
                   transition: transform .15s ease, box-shadow .15s ease; }
        .tarjeta:hover { transform: translateY(-3px); box-shadow: 0 16px 34px rgba(0,0,0,.45); }
        .tarjeta .logo { height: 190px; display: flex; align-items: center; justify-content: center;
                         padding: 1.5rem 2.5rem; background: #FFFFFF; }
        .tarjeta .logo img { max-width: 78%; max-height: 120px; object-fit: contain; }
        .tarjeta .pie { padding: .8rem 1.2rem; display: flex; flex-direction: column; gap: .15rem; }
        .tarjeta .pie b { font-size: 1.15rem; }
        .tarjeta .pie span { font-size: .9rem; opacity: .9; }
        .tarjeta.astara { border-bottom: 6px solid #FFC000; }
        .tarjeta.astara .pie { background: #FFC000; }
        .tarjeta.astara .pie * { color: #002060 !important; }
        .tarjeta.byd { border-bottom: 6px solid #D50C1B; }
        .tarjeta.byd .pie { background: #D50C1B; }
        .tarjeta.byd .pie * { color: #FFFFFF !important; }
        div.st-key-ir_astara button { background: #FFC000 !important; border-color: #FFC000 !important; }
        div.st-key-ir_astara button p { color: #002060 !important; font-weight: 700; }
        div.st-key-ir_byd button { background: #D50C1B !important; border-color: #D50C1B !important; }
        div.st-key-ir_byd button p { color: #FFFFFF !important; font-weight: 700; }
        </style>
        <div class="inicio-titulo">
          <h1>🛠️ Generador de Reportes Técnicos</h1>
          <p>Elige el informe que quieres rellenar. Cada uno trae su plantilla integrada:
          solo sube la Orden de Trabajo, el Trabajo en Garantías y las fotos.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.write("")
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown(_tarjeta("astara", _logo_base64("logo_astara.png"), "Reporte de Falla Astara",
                             "Plantilla en español · fotos en grilla · máx. 25 MB"),
                    unsafe_allow_html=True)
        if st.button("Rellenar informe Astara  →", type="primary", use_container_width=True,
                     key="ir_astara"):
            st.switch_page(PAGINA_ASTARA)
    with c2:
        st.markdown(_tarjeta("byd", _logo_base64("logo_byd.png"), "Service Report BYD",
                             "Plantilla en inglés · piezas, mano de obra y repuestos"),
                    unsafe_allow_html=True)
        if st.button("Rellenar informe BYD  →", type="primary", use_container_width=True,
                     key="ir_byd"):
            st.switch_page(PAGINA_BYD)
    st.caption("Puedes cambiar de informe en cualquier momento desde el menú de la izquierda. "
               "Al cambiar, la página del otro informe empieza vacía.")


PAGINA_INICIO = st.Page(pagina_inicio, title="Inicio", icon="🏠", default=True)
PAGINA_ASTARA = st.Page(str(RAIZ / "informe_astara.py"), title="Reporte Astara", icon="🚗",
                        url_path="astara")
PAGINA_BYD = st.Page(str(RAIZ / "Reporte tecnico BYD" / "pagina_byd.py"), title="Reporte BYD",
                     icon="🔋", url_path="byd")

st.navigation([PAGINA_INICIO, PAGINA_ASTARA, PAGINA_BYD]).run()
