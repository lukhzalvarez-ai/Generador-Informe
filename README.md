# Reporte Técnico Astara

App en Streamlit que genera el `Reporte de falla Astara.xlsx` a partir de:

- `Orden de Trabajo.pdf`
- `Trabajo en Garantías.pdf`
- La plantilla `Reporte Técnico Astara.xlsx`
- Fotografías del vehículo (`.jpg`, `.jpeg`, `.png`)

Todo se sube en **un solo campo de carga**; la app detecta automáticamente
qué archivo es cada uno (por nombre y, si hace falta, por su contenido).

## Estructura del proyecto

```
reporte-astara/
├── app.py                        # App Streamlit (toda la lógica)
├── requirements.txt              # Dependencias
├── .streamlit/config.toml        # Tema visual (rojo/negro, estilo automotriz)
├── Iniciar_Reporte_Astara.bat    # Lanzador local para Windows
└── README.md
```

## Opción A — Publicar en GitHub + Streamlit Community Cloud (link público)

1. Crea un repositorio en GitHub (puede ser privado) y sube estos archivos
   tal como están (`app.py`, `requirements.txt`, `.streamlit/config.toml`).
2. Entra a https://share.streamlit.io/ con tu cuenta (puedes iniciar sesión
   con GitHub).
3. Click en **"New app"**, elige el repositorio, la rama y `app.py` como
   archivo principal.
4. Deploy. Streamlit te entrega un link público del tipo
   `https://tu-app.streamlit.app` que puedes compartir o guardar como acceso
   directo.

Cada vez que hagas `git push` con cambios, la app se actualiza sola.

## Opción B — Ejecutar en tu PC con el `.bat`

1. Asegúrate de tener Python instalado (https://www.python.org/downloads/).
2. Copia toda la carpeta `reporte-astara` a tu PC.
3. Haz doble clic en **`Iniciar_Reporte_Astara.bat`**.
   - La primera vez crea un entorno virtual e instala las dependencias
     (puede tardar uno o dos minutos).
   - Luego abre automáticamente tu navegador en `http://localhost:8501`.
4. Deja la ventana negra (consola) abierta mientras usas la app; al
   cerrarla, la app deja de estar disponible.

> Tip: crea un acceso directo del `.bat` en el escritorio para abrirlo con
> un doble clic sin buscar la carpeta cada vez.

## Notas técnicas

- La lógica de extracción de datos (regex, mapeo de concesionarios, reglas
  del PDF de Garantías, inserción de imágenes) es la misma que la del script
  original; solo cambió el manejo de archivos, que ahora trabaja en memoria
  (sin rutas fijas como `/content/...`) para que funcione igual en local o
  en la nube.
- La vista previa muestra, campo por campo, qué celda del Excel recibió qué
  valor, además de una galería con las fotos que se insertarán y en qué
  posición.
- Si un campo crítico (VIN, fechas, queja del cliente, diagnóstico técnico)
  queda vacío tras la extracción, la app lo avisa explícitamente para que se
  revise a mano antes de enviar el reporte.
