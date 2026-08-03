# Dashboard de Asistencias — CCRM

Progressive Web App (PWA) de asistencia para una organización ministerial/iglesia. Permite cargar un reporte Excel, visualizarlo en KPIs, gráficos y tablas filtrables por Grupo Ministerial, y sincronizarlo con GitHub / Google Sheets. Construida en **JavaScript vanilla (ES6+)**, con una arquitectura modular de objetos literales (motores/"engines"), Bootstrap 5 y Chart.js.

---

## Tabla de contenido

- [Stack técnico](#stack-técnico)
- [Estructura de archivos](#estructura-de-archivos)
- [Arquitectura general](#arquitectura-general)
- [Seguridad y control de acceso (RBAC)](#seguridad-y-control-de-acceso-rbac)
- [Sesión y usuarios](#sesión-y-usuarios)
- [Carga y parseo de Excel](#carga-y-parseo-de-excel)
- [Filtros, KPIs, gráficos y tablas](#filtros-kpis-gráficos-y-tablas)
- [Feedback visual "Recalculando"](#feedback-visual-recalculando)
- [Tema claro/oscuro](#tema-claroscuro)
- [Módulos de UX independientes](#módulos-de-ux-independientes)
- [Integraciones externas](#integraciones-externas)
- [Rendimiento](#rendimiento)
- [Cómo agregar cosas nuevas](#cómo-agregar-cosas-nuevas)

---

## Stack técnico

- **JavaScript ES6+** puro (sin frameworks de UI), organizado en objetos literales tipo "engine" (`UIController`, `ChartEngine`, `TableEngine`, `AccessManager`, etc.).
- **Bootstrap 5** — modales, offcanvas, tabs, componentes base.
- **Chart.js** — gráficos (dona, embudo, barras, barras apiladas, ranking horizontal, sparklines).
- **SheetJS (XLSX)** — parseo de archivos Excel, ejecutado dentro de un **Web Worker** (`excelWorker.js`) para no bloquear el hilo principal.
- **GitHub API v3** — hosting (GitHub Pages) y persistencia de reportes/config.
- **Bot de Telegram** — notificaciones de auditoría (login/logout, cargas, descargas, eliminaciones, cambios de configuración).
- **PWA** — instalable, pensado para uso en campo/móvil.

---

## Estructura de archivos

```
index.html                index.html principal (estructura, <template> de modales pesados)
style.css                 Todo el CSS del dashboard (temas, componentes, animaciones)
app.js                    Núcleo de la aplicación: motores de datos, filtros, KPIs, gráficos, tablas, sesión, tema, sync
excelWorker.js             Web Worker: XLSX.read() del archivo binario fuera del hilo principal

SecurityConfig.js          Diccionario RBAC (usuario → grupos ministeriales permitidos / 'ALL')
AccessManager.js           Motor que filtra registros según SecurityConfig.js (fail-closed)
USUARIOS.JS                 Lista blanca de usuarios autorizados para iniciar sesión

TelegramEngine.js           Notificaciones de auditoría vía Bot de Telegram
SwitchSessionEngine.js       Modal "Cambiar Sesión" (valida contra USUARIOS.JS, reinicia sesión)

InteractiveLife.js           Micro-interacciones visuales: ripple, tilt 3D, conteo animado de KPIs,
                              animación de entrada de Chart.js, cascada de tablas, splash de login,
                              partículas de login, barrido de cambio de tema
FullscreenReconnect.js       Botón flotante para reanudar pantalla completa si el navegador la cierra sola
LazyModals.js                 Inyección diferida (lazy) de los modales pesados de Historial y Google Sheets
```

---

## Arquitectura general

El proyecto **no usa módulos ES6 (`import`/`export`)**. Todos los scripts se cargan como `<script>` clásicos en `index.html`, en un orden pensado para que las variables globales (`ACCESS_RULES`, `USUARIOS_REGISTRADOS`, `AccessManager`, `TelegramEngine`, etc.) existan antes de que `app.js` las use. Esta decisión es deliberada: los módulos ES6 dependen de resolución de red/CORS que puede fallar si el sitio se abre directamente desde disco o en ciertos contextos de hosting, y eso dejaba el sistema en *fail-closed* por error.

`app.js` concentra los siguientes "motores" (objetos literales, no clases):

| Motor | Responsabilidad |
|---|---|
| `ExcelParser` | Parsea el workbook (ya leído por el Worker) a arreglos de registros |
| `DataStore` | Estado central: datos crudos, filtros activos, buffer del archivo |
| `FilterEngine` | Lee/puebla/resetea los `<select>` de filtros (grupo, estado, célula, servicio, nuevo) |
| `KPIEngine` | Calcula los KPIs (asistencia, ausentismo, nuevos, por grupo) sobre el set filtrado |
| `ChartEngine` | Crea/actualiza/destruye las instancias de Chart.js |
| `TableEngine` | Renderiza las tablas (Personas, Excluidos, Nuevos, Histórico) |
| `AbsenceEngine` | Monitor de ausencias |
| `TrendEngine` | Flechas/porcentaje/sparklines de comparación contra una línea base |
| `UIController` | Coordina todo: eventos, carga de archivo, `refresh()` |
| `ThemeEngine` | Tema claro/oscuro, persistencia en `localStorage`, sincroniza Chart.js |
| `SessionEngine` / `AuthEngine` | Sesión activa, expiración de token, logout |
| `AuditEngine` | Envoltorio sobre `TelegramEngine` para notificar acciones sensibles |
| `CloudEngine` / `SaveEngine` / `GSheetsEngine` / `HistoryEngine` | Sincronización con GitHub / Google Sheets / historial de reportes |

### Flujo de datos (alto nivel)

```
Archivo .xlsx
   │  (FileReader → ArrayBuffer)
   ▼
excelWorker.js (Web Worker, hilo aparte)
   │  XLSX.read() + serializa hojas a objeto plano
   ▼
ExcelParser.parse(workbook)   (hilo principal)
   │  genera DataStore.rawMain / rawExcluidos / etc.
   ▼
AccessManager.applyFilter(datos, usuario)   ← filtro de PERMISOS (por grupo ministerial)
   │
   ▼
FilterEngine (filtros de UI: grupo, estado, célula, servicio, nuevo)
   │
   ▼
KPIEngine.compute(filtered) → kpis
   │
   ├─→ UIController.updateKPICards(kpis)
   ├─→ TrendEngine.render(kpis)
   ├─→ ChartEngine.renderAll(kpis)
   ├─→ TableEngine.renderAll(filtered)
   └─→ AbsenceEngine.render(filtered)
```

---

## Seguridad y control de acceso (RBAC)

### `USUARIOS.JS`
Lista blanca (`USUARIOS_REGISTRADOS`) de todos los nombres de usuario que pueden iniciar sesión. Se compara sin distinguir mayúsculas/espacios.

### `SecurityConfig.js`
Diccionario `ACCESS_RULES`: mapea cada usuario a lo que puede ver:
- `'ALL'` → acceso total (liderazgo pastoral / administración: `PASTOR`, `PASTORA`, `MASTER`).
- `Array<string>` → lista blanca de valores permitidos para el campo "Grupo Ministerial" de cada registro (comparación insensible a mayúsculas/espacios), correspondiente al texto de la celda combinada de cada grupo en el Excel.

**Cualquier usuario que NO aparezca en `ACCESS_RULES` recibe fail-closed** (`[]`, no ve nada). Agregar acceso a un usuario nuevo requiere agregarlo explícitamente aquí.

### `AccessManager.js`
Motor que aplica esas reglas sobre los datos **ya parseados**, antes de que lleguen a `FilterEngine`/`UIController`/`ChartEngine`. Política **estrictamente fail-closed**: ante cualquier duda, dato ausente o formato inesperado, la respuesta por defecto es no mostrar nada. Expone `AccessManager.applyFilter(datos, usuario)`.

> Todos estos tres archivos se cargan como `<script>` clásico (no `type="module"`) por la misma razón de robustez explicada arriba.

---

## Sesión y usuarios

- `SessionEngine`/`AuthEngine` (en `app.js`) manejan login, logout, expiración del token y el nombre de usuario activo en `sessionStorage` (llave `ccrm_dashboard_user`).
- **`SwitchSessionEngine.js`** — módulo independiente para el botón "Cambiar Sesión" del menú lateral: valida el nuevo nombre contra `USUARIOS_REGISTRADOS`, reemplaza la sesión activa en el mismo `sessionStorage` y recarga la página, para que `AccessManager`, `KPIEngine`, `ChartEngine` y demás se reinicialicen limpios con el nuevo usuario.
- Todo intento de login (exitoso, fallido, o solicitud de soporte) se notifica vía Telegram (`TelegramEngine.notifySession` / `notifyFailedLogin` / `notifySupport`).

---

## Carga y parseo de Excel

1. El usuario selecciona un archivo → `UIController.loadFile()`.
2. El `ArrayBuffer` se envía a **`excelWorker.js`**, que corre en un **Web Worker** separado del hilo principal. Su única responsabilidad es la parte cara en CPU (`XLSX.read()`); serializa cada hoja a un objeto plano de celdas y lo devuelve por `postMessage`.
3. De vuelta en el hilo principal, `ExcelParser.parse(workbook)` reconstruye los registros exactamente igual que si `XLSX.read()` hubiera corrido ahí mismo — ninguna lógica de negocio cambia, solo se mueve el trabajo pesado fuera del hilo de UI.
4. `AccessManager` filtra los registros según el usuario en sesión.
5. `FilterEngine.populate()` puebla los `<select>` de filtros con los grupos/estados detectados.
6. `UIController.refresh()` calcula y pinta todo.

---

## Filtros, KPIs, gráficos y tablas

Filtros disponibles: **Grupo Ministerial, Estado, Célula, Servicio, Nuevo**. Se leen con `FilterEngine.read()` y se aplican con `DataStore.applyFilters()`.

### `UIController.refresh(opts)`

Punto de entrada central que dispara todo el recálculo. Diseñado en **dos fases** para que la UI nunca se sienta "trabada":

- **Fase 1 (síncrona, instantánea):** marca los valores KPI (y, si el cambio es de grupo, también las tarjetas de gráficos) como **"Recalculando…"**.
- **Fase 2 (diferida):** ejecuta el cómputo real (`applyFilters` → `KPIEngine.compute` → `updateKPICards` → `TrendEngine` → `ChartEngine.renderAll` → `TableEngine.renderAll` → `AbsenceEngine.render`).

```js
UIController.refresh();                        // filtros rápidos (estado, célula, servicio, nuevo)
UIController.refresh({ groupChange: true });    // cambio de Grupo Ministerial / reset de filtros
```

Cuando `groupChange` es `true` (cambio de Grupo Ministerial o reset de filtros, el único filtro que puede implicar un salto grande de volumen de datos), la Fase 2 se demora deliberadamente ~1.8s tras un doble `requestAnimationFrame` — le da colchón al cómputo más pesado y, detrás del overlay "Recalculando", disimula cualquier micro-lag real. El resto de los filtros usa solo un doble `requestAnimationFrame` (prácticamente instantáneo).

### `ChartEngine`

Gráficos activos: **Donut** (asistencia general), **Embudo** (nuevos), **Barras por grupo**, **Barras apiladas** (SI/NO/NUEVO por grupo), **Ranking combinado** (asistencia vs. ausentismo, todas las barras horizontales). Cada cambio de filtro hace `destroy()` + `new Chart()` de las instancias afectadas.

Puntos clave de robustez:
- **Animación uniforme:** `animation` (duración/easing) **y** `animations.numbers` (propiedades: `x, y, width, height, circumference, endAngle, innerRadius, outerRadius`) están declarados explícitamente en `baseOptions()` para que **todos** los tipos de gráfico (dona, barras, ranking) animen el mismo conjunto de propiedades de forma uniforme — sin esto, cada controlador de Chart.js trae su propio set de propiedades animadas por defecto y algunas gráficas aparecían "de golpe" mientras otras sí mostraban relleno.
- **Reflow forzado en el ranking:** `_renderCombinedRankChart` cambia el alto del canvas dinámicamente según cuántos grupos hay; se fuerza un reflow (`void canvasEl.offsetHeight`) justo después de cambiar la altura y antes de crear el chart, para que Chart.js mida el tamaño correcto desde el primer frame (evita el "salto"/desaparición al pasar a "Todos los grupos").
- **Interactividad explícita:** `events` e `interaction` se declaran por instancia en `baseOptions()`, para que clic-en-leyenda y tooltips nunca dependan implícitamente de que ningún otro script deje intactos los defaults globales de Chart.js.
- **`ChartEngine.updateTheme(theme)`:** al cambiar de tema claro/oscuro, en vez de destruir y recrear todas las gráficas, actualiza **in-place** los colores (ticks, grid, leyenda, tooltip) de las instancias ya existentes y llama a `chart.update('none')` (sin animación) — cambio instantáneo, sin parpadeo ni "gráficas que desaparecen".

### `TableEngine`

Renderiza las tablas de **Personas, Excluidos, Nuevos e Histórico**, con búsqueda en vivo (`bindTableSearch`). El `<thead>` de `.data-table` usa `position: sticky` con `z-index: 10` y fondo sólido reforzado (`background-color` + `background-clip: padding-box`) para que no se transparente ni se superponga con las filas al hacer scroll.

---

## Feedback visual "Recalculando"

Dos piezas trabajando juntas (ver `style.css` y `UIController` en `app.js`):

- **KPIs:** `.kpi-value.kpi-recalculando` / `.kpi-pct.kpi-recalculando` — texto reemplazado por "Recalculando…", color `var(--gold)` (el ámbar de la paleta del dashboard), con un pulso sutil (`@keyframes kpiRecalcPulse`, respeta `prefers-reduced-motion`).
- **Gráficos (solo en cambio de grupo):** `.chart-card.chart-recalculando::after` — overlay opaco (fondo sólido del tema) que cubre toda la tarjeta del gráfico sin desmontar el canvas ni la instancia de Chart.js debajo, con el mismo texto/color/pulso.

`UIController._setKpisRecalculando(active)` y `UIController._setChartsRecalculando(active)` activan/desactivan estas clases; `_performRefresh()` las apaga al terminar el recálculo real.

---

## Tema claro/oscuro

`ThemeEngine` aplica `data-theme="light"|"dark"` sobre `<html>`, persiste la preferencia en `localStorage` y actualiza los `Chart.defaults` globales de color. Si ya hay gráficos creados, usa `ChartEngine.updateTheme()` (in-place, sin parpadeo); si el dashboard está visible pero aún no hay ninguna instancia, cae al `renderAll()` completo como respaldo.

---

## Módulos de UX independientes

Estos archivos están diseñados para **no depender de nada más** (y viceversa: nada más depende de ellos) — se agregan a `index.html` con un simple `<script src="...">`, en cualquier orden respecto al resto:

- **`InteractiveLife.js`** — ripple en botones, tilt 3D en tarjetas, conteo animado de números KPI (duración corta, ≤400ms, vía `requestAnimationFrame`, sin bloquear el hilo principal), configuración global de animación de entrada de Chart.js (fusionada con `Object.assign`, nunca reemplazada por completo — un reemplazo total rompía la interactividad de los gráficos), cascada de filas en tablas, splash de marca, partículas de login, barrido circular al cambiar de tema, anillo de progreso del loading overlay, revelado teatral del dashboard tras el login. Respeta `prefers-reduced-motion` en todos sus efectos.
- **`FullscreenReconnect.js`** — si el usuario estaba en pantalla completa y el navegador lo saca automáticamente (cambio de pestaña/app), muestra un botón flotante para reanudarla.
- **`LazyModals.js`** — los modales pesados de Historial de Reportes y Google Sheets viven en `<template>` en `index.html` (inertes, sin costo de layout inicial) y se inyectan al DOM real recién en el primer clic del botón que los abre.

---

## Integraciones externas

- **`TelegramEngine.js`** — notifica al chat privado configurado (vía Bot API) cada acción sensible: eliminar, cargar, descargar, guardar, login/logout, login fallido, uso de funciones de análisis (Comparar/Tendencia), cambio de archivo predeterminado, solicitud de soporte. Todo "fire-and-forget": nunca bloquea ni rompe el flujo principal si falla.
- **GitHub API v3 / GitHub Pages** — hosting del sitio y persistencia de reportes/configuración (`CloudEngine` / `SaveEngine` en `app.js`).
- **Google Sheets** (`GSheetsEngine`) — sincronización opcional de reportes.

---

## Rendimiento

- Parseo de Excel en **Web Worker** (`excelWorker.js`) — el hilo principal nunca se congela leyendo el archivo binario.
- `UIController.refresh()` en dos fases (ver arriba) para que el feedback visual llegue antes que el cómputo pesado.
- Animaciones acotadas (≤400ms en KPIs, 350ms en gráficos) — vivas pero rápidas, vía `requestAnimationFrame`.
- Modales pesados diferidos (`LazyModals.js`) — no se parsean ni calculan layout hasta que se usan.
- Todos los efectos decorativos de `InteractiveLife.js` respetan `prefers-reduced-motion`.

---

## Cómo agregar cosas nuevas

- **Nuevo usuario con acceso:** agregarlo a `USUARIOS_REGISTRADOS` (`USUARIOS.JS`) **y** a `ACCESS_RULES` (`SecurityConfig.js`) con `'ALL'` o su lista de grupos permitidos. Si se omite en `SecurityConfig.js`, queda fail-closed (no ve nada) aunque pueda iniciar sesión.
- **Nuevo gráfico:** agregar el método `render*` en `ChartEngine`, seguir el patrón `destroy(id)` → `new Chart(ctx, { ...this.baseOptions(), ... })`, y registrarlo en `ChartEngine.renderAll()`.
- **Nuevo KPI:** agregarlo en `KPIEngine.compute()`, pintarlo en `UIController.updateKPICards()`, y (si aplica) declararlo en `KPI_DETAIL_MAP` para que la tarjeta sea clickeable.
- **Nuevo módulo de UX puramente visual:** seguir el patrón de `FullscreenReconnect.js`/`LazyModals.js` — IIFE autocontenida, sin tocar `AccessManager`, `SecurityConfig`, `USUARIOS.JS` ni la lógica de datos de `app.js`.
