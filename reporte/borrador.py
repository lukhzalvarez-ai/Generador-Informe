"""Borrador de la última sesión, guardado en el NAVEGADOR del usuario (no en el servidor).

  - Textos, fechas, tablas y ajustes de fotos -> localStorage (unos pocos KB).
  - PDF y fotos (archivos originales)         -> IndexedDB del navegador.

Si la página se recarga, se corta la conexión o la app se reinicia, al volver se
ofrece "Recuperar borrador" y los archivos vuelven solos, sin adjuntarlos de nuevo.
Solo se guarda la última sesión: cada guardado reemplaza al anterior.

Este archivo es igual en reporte/ (Astara) y en byd/ (BYD); cada informe usa su
propio almacén en el navegador, así que no se mezclan.
"""

import base64
import hashlib
import json
import time

import streamlit as st

from .archivos import ArchivoEnMemoria, clave_foto
from .componentes import declarar_componente

LOTE_MAX_BYTES = 8 * 1024 * 1024      # archivos enviados al navegador por vuelta

HTML_BORRADOR = r"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="margin:0">
<script>
  var CLAVE = "__CLAVE__", DBN = CLAVE + "_archivos";
  var leido = false, ultimoId = null, ultimoPid = null, cont = 0, cola = Promise.resolve();

  function enviar(valor) {
    valor.seq = Date.now() * 1000 + (++cont);
    window.parent.postMessage({isStreamlitMessage: true, type: "streamlit:setComponentValue",
                               dataType: "json", value: valor}, "*");
  }
  function leerLS() { try { return window.localStorage.getItem(CLAVE); } catch (e) { return null; } }
  function abrir() {
    return new Promise(function (ok, mal) {
      var r = indexedDB.open(DBN, 1);
      r.onupgradeneeded = function () { r.result.createObjectStore("a"); };
      r.onsuccess = function () { ok(r.result); };
      r.onerror = function () { mal(r.error); };
    });
  }
  function tx(modo, fn) {
    return abrir().then(function (d) {
      return new Promise(function (ok, mal) {
        var t = d.transaction("a", modo), req = fn(t.objectStore("a"));
        t.oncomplete = function () { d.close(); ok(req ? req.result : undefined); };
        t.onerror = t.onabort = function () { d.close(); mal(t.error); };
      });
    });
  }
  function claves() { return tx("readonly", function (s) { return s.getAllKeys(); }).catch(function () { return []; }); }
  function aBlob(b64) {
    var bin = atob(b64), n = bin.length, u = new Uint8Array(n);
    for (var i = 0; i < n; i++) u[i] = bin.charCodeAt(i);
    return new Blob([u]);
  }
  function aB64(blob) {
    return new Promise(function (ok) {
      var fr = new FileReader();
      fr.onload = function () { var s = String(fr.result); ok(s.substring(s.indexOf(",") + 1)); };
      fr.onerror = function () { ok(null); };
      fr.readAsDataURL(blob);
    });
  }
  function procesar(A) {
    var cambios = false, p = Promise.resolve();
    try {
      if (A.accion === "borrar") {
        window.localStorage.removeItem(CLAVE);
        p = p.then(function () { return tx("readwrite", function (s) { return s.clear(); }); })
             .then(function () { cambios = true; }, function () {});
      } else if (A.accion === "guardar" && A.datos) {
        window.localStorage.setItem(CLAVE, JSON.stringify({t: new Date().toLocaleString("es-CL"), datos: A.datos}));
      }
    } catch (e) {}
    (A.nuevos || []).forEach(function (n) {          // uno por uno: menos memoria
      p = p.then(function () {
        return tx("readwrite", function (s) { return s.put({nombre: n.nombre, datos: aBlob(n.b64)}, n.k); })
          .then(function () {}, function () {});
      }).then(function () { cambios = true; });
    });
    if (A.conservar) {
      p = p.then(claves).then(function (ks) {
        var q = Promise.resolve();
        ks.filter(function (k) { return A.conservar.indexOf(k) < 0; }).forEach(function (k) {
          q = q.then(function () { return tx("readwrite", function (s) { return s.delete(k); }).catch(function () {}); });
        });
        return q;
      }).then(function () { cambios = true; });
    }
    return p.then(function () { if (cambios) return claves().then(function (ks) { enviar({guardados: ks}); }); })
            .catch(function () {});
  }
  function pedir(k) {
    return tx("readonly", function (s) { return s.get(k); }).then(function (x) {
      if (!x || !x.datos) return {k: k, faltante: true};
      return aB64(x.datos).then(function (b64) { return b64 ? {k: k, nombre: x.nombre, b64: b64} : {k: k, faltante: true}; });
    }, function () { return {k: k, faltante: true}; }).then(function (r) { enviar({archivo: r}); });
  }
  window.addEventListener("message", function (ev) {
    if (!ev.data || ev.data.type !== "streamlit:render") return;
    var a = ev.data.args || {};
    if (!leido) {
      leido = true;
      cola = cola.then(claves).then(function (ks) { enviar({leido: leerLS() || "", guardados: ks}); });
    }
    if (a.id && a.id !== ultimoId) { ultimoId = a.id; var A = a; cola = cola.then(function () { return procesar(A); }); }
    if (a.pedir && a.pid !== ultimoPid) { ultimoPid = a.pid; var P = a.pedir; cola = cola.then(function () { return pedir(P); }); }
  });
  window.parent.postMessage({isStreamlitMessage: true, type: "streamlit:componentReady", apiVersion: 1}, "*");
  window.parent.postMessage({isStreamlitMessage: true, type: "streamlit:setFrameHeight", height: 0}, "*");
</script></body></html>
"""


def componente(almacen: str):
    """almacen: nombre del borrador en el navegador (uno por informe)."""
    return declarar_componente(f"{almacen}_borrador_v2", HTML_BORRADOR.replace("__CLAVE__", almacen))


# --------------------------------------------------------------------------
# Estado (todas las claves pasan por K para que BYD use su prefijo "byd_")
# --------------------------------------------------------------------------
def archivos_recuperados(K) -> list:
    return st.session_state.get(K("bor_recuperados"), [])


def restaurando(K) -> tuple:
    """(recuperados, total) mientras vuelven los archivos del navegador; None si no."""
    pendientes = st.session_state.get(K("bor_por_recuperar")) or []
    if not pendientes:
        return None
    hechos = len(archivos_recuperados(K))
    return hechos, hechos + len(pendientes)


def iniciar_recuperacion(K, d: dict) -> list:
    """Al recuperar un borrador: pide al navegador sus archivos. Devuelve los nombres
    de los que no estaban guardados (hay que adjuntarlos de nuevo)."""
    lista = [x for x in (d.get("archivos_idb") or []) if isinstance(x, list) and len(x) == 2]
    st.session_state[K("bor_por_recuperar")] = lista
    st.session_state[K("bor_recuperados")] = []
    # Borradores antiguos (sin archivos guardados): hay que adjuntarlos de nuevo
    return [] if lista else list(d.get("archivos") or [])


def para_procesar(K, subidos: list) -> list:
    """Archivos recuperados del borrador + los subidos ahora (sin repetir)."""
    recuperados = archivos_recuperados(K)
    ya = {clave_foto(a) for a in recuperados}
    return recuperados + [a for a in (subidos or []) if clave_foto(a) not in ya]


def quitar_recuperados(K) -> None:
    st.session_state.pop(K("bor_recuperados"), None)
    st.session_state.pop(K("bor_por_recuperar"), None)


# --------------------------------------------------------------------------
# Sincronización con el navegador (se llama una vez, al final de la página)
# --------------------------------------------------------------------------
def sincronizar(comp, K, armar_datos, archivos: list, key: str) -> None:
    """armar_datos() -> dict del borrador, o None si aún no hay nada que guardar.
    archivos: PDF y fotos que deben quedar guardados (objetos con .name y .getvalue())."""
    ss = st.session_state
    accion, datos, nuevos, conservar, pedir = None, None, [], None, None
    en_recuperacion = bool(ss.get(K("bor_por_recuperar")))

    if ss.get(K("borrar_borrador")):
        accion = "borrar"
    elif ss.get(K("borrador_leido")) and not ss.get(K("borrador_encontrado")) and not en_recuperacion:
        d = armar_datos()
        if d is not None:
            por_clave = {clave_foto(a): a for a in archivos}
            d["archivos_idb"] = [[k, a.name] for k, a in por_clave.items()]
            accion, datos = "guardar", json.dumps(d, ensure_ascii=False, default=str)
            idb = set(ss.get(K("bor_idb"), []))
            omitir = idb | set(ss.get(K("bor_pend"), [])) | set(ss.get(K("bor_fallidos"), []))
            total = 0
            for k, a in por_clave.items():
                if k in omitir:
                    continue
                b = a.getvalue()
                if nuevos and total + len(b) > LOTE_MAX_BYTES:
                    break
                nuevos.append({"k": k, "nombre": a.name, "b64": base64.b64encode(b).decode("ascii")})
                total += len(b)
            if nuevos:
                ss[K("bor_pend")] = list(set(ss.get(K("bor_pend"), [])) | {n["k"] for n in nuevos})
            if idb - set(por_clave):
                conservar = list(por_clave)
    if en_recuperacion:
        pedir = ss[K("bor_por_recuperar")][0][0]

    firma = f"{accion}|{datos}|{[n['k'] for n in nuevos]}|{conservar}|{ss.get(K('n_cargador'), 0)}"
    ident = hashlib.md5(firma.encode("utf-8")).hexdigest() if (accion or nuevos or conservar) else None
    pid = hashlib.md5(f"{pedir}|{len(archivos_recuperados(K))}|{time.time() // 30}".encode()).hexdigest() \
        if pedir else None
    valor = comp(accion=accion, datos=datos, id=ident, nuevos=nuevos, conservar=conservar,
                 pedir=pedir, pid=pid, key=key, default=None)
    if accion == "borrar":
        ss.pop(K("borrar_borrador"), None)
        for k in ("bor_idb", "bor_pend", "bor_fallidos"):
            ss[K(k)] = []

    if not isinstance(valor, dict) or valor.get("seq") == ss.get(K("bor_seq")):
        return
    ss[K("bor_seq")] = valor.get("seq")
    seguir = False

    if "guardados" in valor:
        guardados = set(valor.get("guardados") or [])
        pend = set(ss.get(K("bor_pend"), []))
        if "leido" not in valor:      # lo enviado y no guardado (ej. sin espacio) no se reintenta
            ss[K("bor_fallidos")] = list(set(ss.get(K("bor_fallidos"), [])) | (pend - guardados))
        ss[K("bor_idb")] = list(guardados)
        ss[K("bor_pend")] = []
        actuales = {clave_foto(a) for a in archivos}
        seguir = bool(actuales - guardados - set(ss.get(K("bor_fallidos"), [])))

    if "leido" in valor and not ss.get(K("borrador_leido")):
        ss[K("borrador_leido")] = True
        try:
            guardado = json.loads(valor["leido"]) if valor["leido"] else None
        except Exception:
            guardado = None
        if guardado and guardado.get("datos") and not ss.get(K("borrador_decidido")):
            ss[K("borrador_encontrado")] = guardado
        seguir = True

    if isinstance(valor.get("archivo"), dict):
        r = valor["archivo"]
        pendientes = ss.get(K("bor_por_recuperar")) or []
        if pendientes and pendientes[0][0] == r.get("k"):
            k, nombre = pendientes.pop(0)
            if r.get("b64"):
                try:
                    ss.setdefault(K("bor_recuperados"), []).append(
                        ArchivoEnMemoria(r.get("nombre") or nombre, base64.b64decode(r["b64"]), origen="borrador"))
                except Exception:
                    ss.setdefault(K("archivos_por_adjuntar"), []).append(nombre)
            else:
                ss.setdefault(K("archivos_por_adjuntar"), []).append(nombre)
            ss[K("bor_por_recuperar")] = pendientes
            if not pendientes:   # ya están todos: lo recuperado cuenta como guardado
                ss[K("bor_idb")] = [clave_foto(a) for a in archivos_recuperados(K)]
            seguir = True

    if seguir:
        st.rerun()
