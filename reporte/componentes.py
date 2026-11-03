"""Componentes propios de la interfaz (sin librerías externas): recortador
de fotos con esquinas y ordenador de fotos arrastrando. Cada uno es un
index.html que conversa con Streamlit por postMessage."""

import tempfile
from pathlib import Path

import streamlit as st

from .config import logger


HTML_RECORTADOR = r"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  html, body { margin: 0; padding: 0; background: transparent; overflow: hidden;
               user-select: none; -webkit-user-select: none; }
  #zona { position: relative; margin: 0 auto; overflow: hidden; touch-action: none; }
  #zona img { display: block; width: 100%; height: 100%; pointer-events: none; }
  #caja { position: absolute; border: 2px solid #FFC000; box-sizing: border-box; cursor: move;
          box-shadow: 0 0 0 9999px rgba(0, 0, 0, .55); touch-action: none; }
  #caja::before, #caja::after { content: ""; position: absolute; border: 0 dashed rgba(255,192,0,.5); }
  #caja::before { left: 33.3%; right: 33.3%; top: 0; bottom: 0; border-left-width: 1px; border-right-width: 1px; }
  #caja::after  { top: 33.3%; bottom: 33.3%; left: 0; right: 0; border-top-width: 1px; border-bottom-width: 1px; }
  .esq { position: absolute; width: 18px; height: 18px; background: #FFC000; border: 2px solid #002060;
         border-radius: 3px; box-sizing: border-box; touch-action: none; z-index: 2; }
  .nw { left: -9px; top: -9px; cursor: nwse-resize; }   .ne { right: -9px; top: -9px; cursor: nesw-resize; }
  .sw { left: -9px; bottom: -9px; cursor: nesw-resize; } .se { right: -9px; bottom: -9px; cursor: nwse-resize; }
</style></head>
<body>
<div id="zona"><img id="img" alt=""><div id="caja">
  <div class="esq nw" data-c="nw"></div><div class="esq ne" data-c="ne"></div>
  <div class="esq sw" data-c="sw"></div><div class="esq se" data-c="se"></div>
</div></div>
<script>
  // Recortador propio, sin librerías externas.
  // Estado en píxeles de pantalla: x, y, w, h del recuadro dentro de la foto (W x H).
  function enviar(tipo, extra) {
    window.parent.postMessage(Object.assign({isStreamlitMessage: true, type: tipo}, extra || {}), "*");
  }
  var zona = document.getElementById("zona"), img = document.getElementById("img"),
      caja = document.getElementById("caja");
  var W = 0, H = 0, aspecto = null, ultimaClave = null, MIN = 24;
  var r = {x: 0, y: 0, w: 0, h: 0}, arrastre = null;

  function pintar() {
    caja.style.left = r.x + "px"; caja.style.top = r.y + "px";
    caja.style.width = r.w + "px"; caja.style.height = r.h + "px";
  }
  function enviarValor() {
    enviar("streamlit:setComponentValue", {dataType: "json", value: {
      l: r.x / W, t: r.y / H, r: (r.x + r.w) / W, b: (r.y + r.h) / H}});
  }
  function ajustarAspecto() {          // fuerza la forma y lo deja dentro de la foto
    if (aspecto) {
      if (r.w / r.h > aspecto) r.w = r.h * aspecto; else r.h = r.w / aspecto;
    }
    r.w = Math.min(r.w, W); r.h = Math.min(r.h, H);
    if (aspecto) { if (r.w / r.h > aspecto) r.w = r.h * aspecto; else r.h = r.w / aspecto; }
    r.x = Math.min(Math.max(r.x, 0), W - r.w); r.y = Math.min(Math.max(r.y, 0), H - r.h);
  }
  function punto(e) {
    var b = zona.getBoundingClientRect();
    return {x: e.clientX - b.left, y: e.clientY - b.top};
  }
  function iniciar(e) {
    e.preventDefault(); e.stopPropagation();
    var esquina = e.target.getAttribute("data-c");
    var p = punto(e);
    arrastre = {esquina: esquina, p0: p, r0: {x: r.x, y: r.y, w: r.w, h: r.h}};
    if (esquina) {                     // la esquina opuesta queda fija
      arrastre.ax = esquina.indexOf("w") >= 0 ? r.x + r.w : r.x;
      arrastre.ay = esquina.indexOf("n") >= 0 ? r.y + r.h : r.y;
      arrastre.izq = esquina.indexOf("w") >= 0; arrastre.arr = esquina.indexOf("n") >= 0;
    }
    try { e.target.setPointerCapture(e.pointerId); } catch (err) {}
  }
  function mover(e) {
    if (!arrastre) return;
    e.preventDefault();
    var p = punto(e), a = arrastre;
    if (!a.esquina) {                  // mover el recuadro completo
      r.x = Math.min(Math.max(a.r0.x + p.x - a.p0.x, 0), W - r.w);
      r.y = Math.min(Math.max(a.r0.y + p.y - a.p0.y, 0), H - r.h);
    } else {                           // redimensionar desde una esquina
      var maxW = a.izq ? a.ax : W - a.ax, maxH = a.arr ? a.ay : H - a.ay;
      var nw = Math.abs(p.x - a.ax), nh = Math.abs(p.y - a.ay);
      if (aspecto) {
        nw = Math.max(nw, nh * aspecto);
        nw = Math.min(nw, maxW, maxH * aspecto);
        nw = Math.max(nw, MIN * aspecto); nh = nw / aspecto;
      } else {
        nw = Math.max(MIN, Math.min(nw, maxW)); nh = Math.max(MIN, Math.min(nh, maxH));
      }
      r.w = nw; r.h = nh;
      r.x = a.izq ? a.ax - nw : a.ax; r.y = a.arr ? a.ay - nh : a.ay;
    }
    pintar();
  }
  function soltar() { if (arrastre) { arrastre = null; enviarValor(); } }

  caja.addEventListener("pointerdown", iniciar);
  window.addEventListener("pointermove", mover);
  window.addEventListener("pointerup", soltar);
  window.addEventListener("pointercancel", soltar);

  function render(args) {
    if (args.clave === ultimaClave) return;   // no reiniciar el recuadro en cada actualización
    ultimaClave = args.clave;
    aspecto = args.aspect || null;
    var disponible = document.body.clientWidth || 700;
    W = Math.min(disponible, args.alto_max * args.w / args.h);
    H = W * args.h / args.w;
    zona.style.width = W + "px"; zona.style.height = H + "px";
    enviar("streamlit:setFrameHeight", {height: Math.ceil(H) + 4});
    var c = args.inicial;
    if (c) {
      r = {x: c[0] * W, y: c[1] * H, w: (c[2] - c[0]) * W, h: (c[3] - c[1]) * H};
    } else {                           // por defecto: 90 % del máximo, centrado
      r = {w: W, h: H, x: 0, y: 0};
      if (aspecto) { if (W / H > aspecto) r.w = H * aspecto; else r.h = W / aspecto; }
      r.w *= 0.9; r.h *= 0.9; r.x = (W - r.w) / 2; r.y = (H - r.h) / 2;
    }
    ajustarAspecto();
    img.src = args.src;
    pintar();
    enviarValor();
  }
  window.addEventListener("message", function (e) {
    if (e.data && e.data.type === "streamlit:render") render(e.data.args);
  });
  enviar("streamlit:componentReady", {apiVersion: 1});
</script>
</body></html>
"""

HTML_ORDENADOR = r"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  html, body { margin: 0; padding: 0; background: transparent; overflow: hidden;
               font-family: "Source Sans Pro", "Segoe UI", sans-serif; }
  #grid { display: flex; flex-wrap: wrap; gap: 8px; padding: 6px 2px; }
  .item { width: 128px; background: #142449; border: 2px solid #2A3F70; border-radius: 8px;
          padding: 4px; cursor: grab; position: relative; box-sizing: border-box; }
  .item:active { cursor: grabbing; }
  .item.prio { border-color: #FFC000; }
  .item img { width: 100%; height: 70px; object-fit: cover; border-radius: 4px; display: block;
              pointer-events: none; }
  .item .n { position: absolute; top: 7px; left: 7px; background: #002060; color: #FFC000;
             font-weight: 700; border-radius: 4px; padding: 0 6px; font-size: 13px; }
  .item .nom { color: #E6ECF7; font-size: 11px; white-space: nowrap; overflow: hidden;
               text-overflow: ellipsis; margin-top: 3px; }
  .item.arrastrando { opacity: .3; }
  .item.antes { box-shadow: -6px 0 0 -1px #FFC000; }
  .item.despues { box-shadow: 6px 0 0 -1px #FFC000; }
</style></head>
<body><div id="grid"></div>
<script>
  function enviar(tipo, extra) {
    window.parent.postMessage(Object.assign({isStreamlitMessage: true, type: tipo}, extra || {}), "*");
  }
  var grid = document.getElementById("grid");
  var items = {}, orden = [], prioritarias = 2, ultimaClave = null, arrastrada = null;

  function ajustarAlto() { enviar("streamlit:setFrameHeight", {height: grid.scrollHeight + 4}); }
  function limpiarMarcas() {
    grid.querySelectorAll(".antes, .despues").forEach(function (el) { el.classList.remove("antes", "despues"); });
  }
  function dibujar() {
    grid.innerHTML = "";
    orden.forEach(function (k, i) {
      var it = items[k], d = document.createElement("div");
      d.className = "item" + (i < prioritarias ? " prio" : "");
      d.draggable = true; d.dataset.k = k;
      d.title = it.nombre + " — arrastra para mover";
      d.innerHTML = '<img src="' + it.src + '"><span class="n">' + (i < prioritarias ? "⭐ " : "") +
                    (i + 1) + '</span><div class="nom"></div>';
      d.querySelector(".nom").textContent = it.nombre;
      d.addEventListener("dragstart", function (e) {
        arrastrada = k; d.classList.add("arrastrando");
        e.dataTransfer.effectAllowed = "move"; e.dataTransfer.setData("text/plain", k);
      });
      d.addEventListener("dragend", function () { d.classList.remove("arrastrando"); limpiarMarcas(); arrastrada = null; });
      d.addEventListener("dragover", function (e) {
        e.preventDefault(); e.dataTransfer.dropEffect = "move";
        if (!arrastrada || arrastrada === k) return;
        var r = d.getBoundingClientRect(), antes = e.clientX < r.left + r.width / 2;
        limpiarMarcas(); d.classList.add(antes ? "antes" : "despues");
      });
      d.addEventListener("drop", function (e) {
        e.preventDefault();
        var origen = arrastrada || e.dataTransfer.getData("text/plain");
        limpiarMarcas();
        if (!origen || origen === k) return;
        var r = d.getBoundingClientRect(), antes = e.clientX < r.left + r.width / 2;
        orden.splice(orden.indexOf(origen), 1);
        var destino = orden.indexOf(k) + (antes ? 0 : 1);
        orden.splice(destino, 0, origen);
        dibujar();
        enviar("streamlit:setComponentValue", {dataType: "json", value: {orden: orden.slice(), id: Date.now()}});
      });
      grid.appendChild(d);
    });
    setTimeout(ajustarAlto, 0);
  }
  function render(args) {
    if (args.clave === ultimaClave) return;
    ultimaClave = args.clave;
    items = {}; orden = [];
    prioritarias = args.prioritarias || 2;
    args.items.forEach(function (it) { items[it.k] = it; orden.push(it.k); });
    dibujar();
  }
  window.addEventListener("message", function (e) {
    if (e.data && e.data.type === "streamlit:render") render(e.data.args);
  });
  window.addEventListener("resize", ajustarAlto);
  enviar("streamlit:componentReady", {apiVersion: 1});
</script></body></html>
"""


HTML_BORRADOR = r"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="margin:0">
<script>
  // Borrador del Reporte de Falla Astara guardado en ESTE navegador (localStorage).
  // Solo se guarda la última sesión: cada guardado reemplaza al anterior.
  var CLAVE = "astara_reporte_falla_borrador_v1";
  var leido = false, ultimoId = null;
  function enviar(tipo, extra) {
    window.parent.postMessage(Object.assign({isStreamlitMessage: true, type: tipo}, extra || {}), "*");
  }
  function leer() { try { return window.localStorage.getItem(CLAVE); } catch (e) { return null; } }
  window.addEventListener("message", function (ev) {
    if (!ev.data || ev.data.type !== "streamlit:render") return;
    var a = ev.data.args || {};
    if (!leido) {                       // primero se lee lo que había (antes de escribir nada)
      leido = true;
      enviar("streamlit:setComponentValue", {dataType: "json", value: {leido: leer() || ""}});
    }
    if (a.id && a.id !== ultimoId) {
      ultimoId = a.id;
      try {
        if (a.accion === "guardar" && a.datos) {
          window.localStorage.setItem(CLAVE, JSON.stringify(
            {t: new Date().toLocaleString("es-CL"), datos: a.datos}));
        } else if (a.accion === "borrar") {
          window.localStorage.removeItem(CLAVE);
        }
      } catch (e) {}
    }
  });
  enviar("streamlit:componentReady", {apiVersion: 1});
  enviar("streamlit:setFrameHeight", {height: 0});
</script></body></html>
"""


@st.cache_resource(show_spinner=False)
def declarar_componente(nombre: str, html: str):
    """Escribe el index.html de un componente propio en una carpeta temporal y
    lo registra en Streamlit (una sola vez por servidor). None si falla."""
    try:
        import streamlit.components.v1 as components
        carpeta = Path(tempfile.gettempdir()) / f"{nombre}_v1"
        carpeta.mkdir(parents=True, exist_ok=True)
        (carpeta / "index.html").write_text(html, encoding="utf-8")
        return components.declare_component(nombre, path=str(carpeta))
    except Exception as e:
        logger.error(f"No se pudo preparar el componente '{nombre}': {e}")
        return None


def recortador():
    return declarar_componente("astara_recortador", HTML_RECORTADOR)


def ordenador():
    return declarar_componente("astara_ordenador", HTML_ORDENADOR)


def borrador():
    return declarar_componente("astara_borrador", HTML_BORRADOR)
