# Aria Escalón · Generador de Reporte de Disponibilidad

App web **sin backend** que convierte un Excel de inventario en un **PDF comercial**
con la identidad de Aria Escalón, listo para que la fuerza de ventas descargue siempre
la última versión.

Todo el procesamiento ocurre **en el navegador** (lectura del Excel, generación del PDF).
No hay servidor, base de datos ni datos que salgan del equipo del usuario. Se publica como
sitio estático (ideal para **GitHub Pages**).

---

## ✨ Qué hace

1. Subes (o arrastras) un Excel con el inventario de apartamentos.
2. La app detecta las columnas, agrupa por modelo y calcula el resumen
   (total de unidades, modelos activos, niveles, precio mínimo y máximo).
3. Muestra una **vista previa** del reporte y permite **descargar el PDF**
   con tipografías de marca (Albra Book Light + PP Mori) y el logo incrustados.

---

## ▶️ Probarlo en local

Lo más sencillo: **abre `index.html`** con doble clic. Funciona directamente desde el
archivo (`file://`), sin instalar nada.

Si tu navegador bloqueara la carga de las fuentes vía `file://`, levanta un servidor
local de un solo comando (cualquiera de estos):

```bash
# Opción Python (viene preinstalado en Mac/Linux)
python3 -m http.server 8080

# Opción Node
npx serve .
```

Luego abre <http://localhost:8080>.

> Pulsa **"Ver ejemplo"** para cargar datos de muestra y ver el PDF al instante,
> sin necesidad de subir un archivo.

---

## ☁️ Publicar en GitHub Pages

1. Crea un repositorio y sube **todo el contenido de esta carpeta** (incluyendo
   `index.html` en la raíz).
2. En GitHub: **Settings → Pages → Build and deployment**.
3. En *Source* elige **Deploy from a branch**, rama `main` y carpeta `/ (root)`.
4. Guarda. En 1–2 minutos tu reporte estará disponible en
   `https://<usuario>.github.io/<repo>/`.

El archivo `.nojekyll` ya está incluido para que GitHub Pages sirva las fuentes y demás
recursos sin procesarlos con Jekyll.

> ⚠️ **Importante:** `.nojekyll` es un archivo *oculto*. Si subes el proyecto
> **arrastrando los archivos** a la web de GitHub, el navegador suele **omitirlo**.
> Verifica que `.nojekyll` aparezca en la raíz del repositorio. Si no está, créalo:
> **Add file → Create new file → nombre `.nojekyll` (vacío) → Commit**.

### 🛠️ Si la página carga pero no deja subir el Excel

Casi siempre es porque las librerías de `js/lib/` no se publicaron (Jekyll las omitió).
Solución:

1. Confirma que existe `.nojekyll` en la raíz del repo (ver arriba).
2. Recarga con **Ctrl/Cmd + Shift + R** (limpia caché).
3. Para diagnosticar: abre la página, pulsa **F12 → Console**. Si ves errores `404`
   apuntando a `js/lib/pdfmake.min.js` o `xlsx.full.min.js`, es exactamente este caso.

Esta versión además **avisa en pantalla** qué archivo faltó, en vez de quedarse sin responder.

---

## 📋 Formato del Excel

Solo importa que existan estas **cinco columnas** (en cualquier orden, con o sin acentos):

| Columna             | Ejemplo        | Notas                                  |
|---------------------|----------------|----------------------------------------|
| **Modelo**          | `A`, `C+`      | Letra del modelo. "MODELO A" también sirve. |
| **Unidad**          | `1`            | Número de unidad dentro del nivel.     |
| **Nivel**           | `2`            | Piso / nivel.                          |
| **Mts2 Totales**    | `136.38`       | Área total en m².                      |
| **Precio con Bodega** | `406603.90`  | Precio. Acepta `$` y separadores de miles. |

La etiqueta de cada unidad en el PDF se arma como `APTO {nivel}-{unidad}`
(por ejemplo, nivel 2 + unidad 1 → **APTO 2-1**).

Tienes una **plantilla lista** en [`data/ejemplo.xlsx`](data/ejemplo.xlsx) y también se
puede descargar desde el botón "Descargar plantilla" de la app.

---

## 🗂️ Estructura

```
aria-dashboard/
├── index.html               App principal (UI + carga + descarga)
├── js/
│   ├── report.js            Diseño del PDF (pdfmake) + normalización de datos
│   ├── app.js               Lógica: leer Excel, vista previa, descarga
│   └── lib/
│       ├── xlsx.full.min.js Lectura de Excel (SheetJS)
│       ├── pdfmake.min.js   Generación de PDF
│       └── aria-assets.js   Fuentes de marca + logo en base64 (para el PDF)
├── assets/
│   ├── logo.png             Logo (UI y portada del PDF)
│   └── fonts/               Albra Book Light, PP Mori Regular/SemiBold (.otf)
├── data/
│   └── ejemplo.xlsx         Plantilla / datos de muestra
├── .nojekyll
└── README.md
```

---

## 🔤 Nota sobre las tipografías

Las fuentes incluidas son las versiones **TRIAL** que se proporcionaron
(`AlbraBookTRIAL-Light`). La versión de prueba de Albra **no incluye algunos glifos**
(por ejemplo el signo `+`); la app ya lo resuelve dibujando esos caracteres con PP Mori.

Para producción, reemplaza los archivos en `assets/fonts/` por las versiones con
**licencia completa** (manteniendo los mismos nombres de archivo) y vuelve a generar
`js/lib/aria-assets.js` con el script de la sección siguiente. Así el PDF quedará con
el set de caracteres completo.

---

## 🔁 Regenerar `aria-assets.js` (al cambiar fuentes o logo)

Ese archivo contiene las fuentes y el logo codificados en base64 para incrustarlos en el
PDF. Si actualizas algún archivo de `assets/`, regéneralo con:

```bash
python3 - <<'PY'
import base64, json, os
files = {
  'Albra-Light.otf':    'assets/fonts/Albra-Light.otf',
  'PPMori-Regular.otf': 'assets/fonts/PPMori-Regular.otf',
  'PPMori-SemiBold.otf':'assets/fonts/PPMori-SemiBold.otf',
  'aria-logo.png':      'assets/logo.png',
}
out = ['(function(){','  if (typeof window === "undefined") return;',
       '  window.ARIA_VFS = window.ARIA_VFS || {};']
for name, path in files.items():
    b = base64.b64encode(open(path,'rb').read()).decode()
    out.append('  window.ARIA_VFS[%s] = %s;' % (json.dumps(name), json.dumps(b)))
out.append('})();')
open('js/lib/aria-assets.js','w').write('\n'.join(out))
print('OK')
PY
```

> Si cambias el logo, conviene reducirlo a ~760 px de ancho antes de codificarlo,
> para mantener liviano el archivo.

---

## ⚙️ Detalles técnicos

- **Sin backend / sin estado:** nada se guarda; cada descarga refleja el Excel cargado.
- **PDF vectorial:** texto seleccionable y nítido a cualquier zoom; fuentes incrustadas.
- **Librerías vendorizadas** (no se cargan desde CDN): funciona offline.
- **Tamaño de página:** Carta (Letter), vertical. Paginación automática por modelos.
