/* ================================================================
   LazyModals.js — Modularidad del DOM: inyección diferida de
   modales pesados (Historial de Reportes y Google Sheets)
   ────────────────────────────────────────────────────────────
   PROBLEMA QUE RESUELVE:
   #modalHistorial y #gsheetsModal son dos de los bloques de HTML
   más pesados del dashboard (listas de archivos, formularios de
   sync, banners, spinners...), pero la inmensa mayoría de las
   sesiones jamás los abre. Antes se parseaban, se les calculaba
   layout/estilo y quedaban vivos en el DOM desde el primer momento,
   aunque el usuario nunca hiciera clic en "Historial" ni en
   "Conectar Google Sheets".

   SOLUCIÓN:
   En index.html, ambos modales ahora están envueltos en
   <template id="tpl-modalHistorial"> / <template id="tpl-gsheetsModal">.
   El contenido de un <template> es INERTE: el navegador lo parsea
   pero no lo renderiza, no le aplica CSS ni lo cuenta en el layout
   inicial. Este script los clona e inserta en el DOM real la
   PRIMERA vez que el usuario hace clic en el botón que abre cada
   modal — nunca antes.

   POR QUÉ SE HACE EN FASE DE CAPTURA (capture: true):
   Tanto Bootstrap (para #modalHistorial, vía su propio
   data-bs-toggle="modal") como GSheetsEngine (para #gsheetsModal,
   vía su propio listener en app.js) esperan que el modal YA EXISTA
   en el DOM en el momento en que procesan el clic. Un listener en
   fase de CAPTURA sobre `document` se ejecuta SIEMPRE antes que
   cualquier listener en fase de burbuja (la fase por defecto, que
   usan tanto el data-api de Bootstrap como los listeners normales
   de app.js) — sin importar el orden en que se cargaron los
   scripts. Así garantizamos que, para cuando Bootstrap/app.js
   procesen el clic, el modal ya esté en el DOM.

   QUÉ NO SE TOCA:
   Este archivo no modifica AccessManager.js, SecurityConfig.js,
   USUARIOS.JS, ni ninguna lógica de filtros/roles. Solo mueve EL
   MOMENTO en que dos bloques de HTML se insertan en el DOM.
================================================================ */

const LazyModals = {

  _injected: new Set(),

  /* Clona el <template> indicado y lo inserta al final de <body>.
     Idempotente: si ya se inyectó, no vuelve a hacer nada.
     Devuelve el elemento del modal ya presente en el DOM real
     (o null si la plantilla no existe, p. ej. si index.html cambió). */
  inject(templateId, modalId) {
    if (this._injected.has(modalId)) return document.getElementById(modalId);

    const tpl = document.getElementById(templateId);
    if (!tpl || !tpl.content || !tpl.content.firstElementChild) {
      console.error(`[LazyModals] No se encontró el <template id="${templateId}">.`);
      return null;
    }

    const node = tpl.content.firstElementChild.cloneNode(true);
    document.body.appendChild(node);
    this._injected.add(modalId);

    return node;
  },

  /* Engancha la inyección diferida a uno o varios selectores
     disparadores. La PRIMERA vez que se hace clic en alguno de
     ellos, inyecta el modal (una sola vez) y ejecuta `onFirstOpen`
     (si se indica) ANTES de que el clic continúe su curso normal
     hacia Bootstrap/app.js. Clics posteriores ya no pasan por aquí
     (el modal ya existe, así que Bootstrap/app.js lo manejan solos,
     exactamente igual que si nunca hubiera sido diferido).

     `takeOverFirstClick` (opcional): si es `true`, ese primer clic
     NUNCA llega al data-api de Bootstrap (se cancela con
     preventDefault/stopPropagation) y `onFirstOpen(modalEl)` es
     responsable de mostrar el modal por su cuenta. Es necesario
     para triggers que usan data-bs-toggle="modal" en el propio
     botón (p. ej. #btnAbrirHistorial): si se deja que Bootstrap
     procese ese mismo clic síncronamente, intenta mostrar un modal
     que el navegador acaba de insertar y aún no tiene layout
     calculado (no hubo reflow entre la inserción y el show()), así
     que la transición fade→show no se aplica y el modal no aparece
     — recién en el SEGUNDO clic, con el nodo ya existente y con
     layout ya calculado, Bootstrap logra mostrarlo. Forzando el
     show() nosotros mismos (tras un reflow explícito) evita ese
     problema desde el primer clic. */
  bindTriggers(selectors, templateId, modalId, onFirstOpen, takeOverFirstClick = false) {
    document.addEventListener('click', (ev) => {
      if (this._injected.has(modalId)) return; // ya inyectado: no interferir

      const trigger = selectors
        .map(sel => ev.target.closest(sel))
        .find(Boolean);
      if (!trigger) return;

      if (takeOverFirstClick) {
        ev.preventDefault();
        ev.stopPropagation(); // evita que el data-api de Bootstrap procese este mismo clic
      }

      const modalEl = this.inject(templateId, modalId);
      if (!modalEl) return;

      if (typeof onFirstOpen === 'function') onFirstOpen(modalEl);
    }, true /* fase de captura — ver comentario arriba */);
  },
};


document.addEventListener('DOMContentLoaded', () => {

  /* ── Historial de Reportes ──
     Se abre vía data-bs-toggle="modal" data-bs-target="#modalHistorial"
     (Bootstrap nativo, no requiere JS propio para abrirse). Tras
     inyectarlo, hace falta ejecutar la inicialización que antes
     corría en el DOMContentLoaded global (ver app.js) y que depende
     de que el modal ya exista: HistoryEngine.init() (carga la lista
     de archivos, botón reintentar, botón recargar) y el enlace
     "Renovar ahora" / banner de caducidad del token, que viven
     dentro de este mismo modal (AuthEngine.bindExpiryLink /
     AuthEngine.checkExpiry). */
  LazyModals.bindTriggers(['#btnAbrirHistorial'], 'tpl-modalHistorial', 'modalHistorial', (modalEl) => {
    if (window.HistoryEngine) HistoryEngine.init();
    if (window.AuthEngine) {
      AuthEngine.bindExpiryLink();
      AuthEngine.checkExpiry();
    }
    /* El contenedor de la lista (#listaReportesContainer) recién
       existe en el DOM en este punto — es el momento correcto para
       que UsuarioRules empiece a vigilarlo y oculte "Descargar" a
       LECTOR en cada archivo que se renderice (ver comentario en
       Usuario_Rules.js: attachToHistorialModal()). */
    if (window.UsuarioRules) UsuarioRules.attachToHistorialModal();

    /* Fuerza un reflow del nodo recién insertado antes de mostrarlo:
       sin esto, Bootstrap calcula el layout del modal como si aún
       no estuviera en el DOM y la transición fade→show no se aplica
       en el primer clic (ver comentario en bindTriggers). */
    void modalEl.offsetHeight;
    bootstrap.Modal.getOrCreateInstance(modalEl).show();
  }, true /* takeOverFirstClick: este botón usa data-bs-toggle, hay que interceptar su primer clic */);

  /* ── Google Sheets ──
     A diferencia de Historial, este modal SÍ se abre por JS propio
     (GSheetsEngine.state.modalRef.show()), no por data-api de
     Bootstrap. GSheetsEngine.initModal() es lo que crea esa
     instancia de bootstrap.Modal y registra todos los listeners
     internos del formulario de sync — antes corría siempre al
     cargar la página; ahora se dispara aquí, la primera vez que el
     usuario abre el modal, y después llamamos nosotros mismos a
     `.show()` para no depender de que el listener que initModal()
     acaba de registrar llegue a ejecutarse para este mismo clic. */
  const openGsheets = (modalEl) => {
    if (window.GSheetsEngine && !GSheetsEngine.state.modalRef) {
      GSheetsEngine.initModal();
    }
    GSheetsEngine.state.modalRef?.show();
  };
  LazyModals.bindTriggers(['#btnGsheets', '#btnGsheetsEmpty'], 'tpl-gsheetsModal', 'gsheetsModal', openGsheets);
});
