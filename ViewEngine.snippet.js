/* ────────────────────────────────────────────────────────────
   VIEW ENGINE — Carga dinámica de vistas (partials) + gestión
   del ciclo de vida de Event Listeners.

   PEGAR ESTO EN app.js, ANTES de cualquier código que asuma
   que el DOM de login/dashboard ya existe (es decir, ANTES de
   cualquier `document.getElementById('btnIniciarSesion')`
   que hoy corras "al vuelo" al cargar el script).

   IMPORTANTE: DataStore, TelegramEngine y el resto de tus
   motores NO se tocan. Solo cambia CUÁNDO se llama a sus
   inicializadores de listeners.
──────────────────────────────────────────────────────────── */
const ViewEngine = {

  /* Vista actualmente montada en #app-container */
  currentView: null,

  /* Caché en memoria de HTML ya descargado (evita refetch en
     navegaciones repetidas — opcional, quita esto si tus
     vistas cambian dinámicamente en el servidor) */
  _cache: {},

  /* Contenedor raíz donde se inyecta cada vista */
  get container() {
    return document.getElementById('app-container');
  },

  /* ── Carga y monta una vista dentro de #app-container ── */
  async loadView(viewName, { showLoading = true } = {}) {
    const overlay = document.getElementById('loadingOverlay');

    try {
      if (showLoading && overlay) overlay.classList.remove('d-none');

      const html = await this._fetchView(viewName);

      // 1. Limpieza de la vista anterior (listeners, timers, charts)
      this._onViewUnload(this.currentView);

      // 2. Inyección del nuevo HTML
      this.container.innerHTML = html;
      this.currentView = viewName;

      // 3. Inicialización — SOLO después de que el HTML ya está
      //    en el DOM real. Este es el punto crítico que evita
      //    los "Cannot read properties of null" en botones que
      //    aún no existían.
      this._onViewLoaded(viewName);

    } catch (err) {
      console.error(`[ViewEngine] Error al cargar la vista "${viewName}":`, err);
      this.container.innerHTML = `
        <div class="p-5 text-center text-danger">
          No se pudo cargar la vista "${viewName}". Revisa la consola.
        </div>`;
    } finally {
      if (showLoading && overlay) overlay.classList.add('d-none');
    }
  },

  /* ── Descarga el HTML del partial (con caché simple) ── */
  async _fetchView(viewName) {
    if (this._cache[viewName]) return this._cache[viewName];

    const res = await fetch(`vistas/${viewName}.html`);
    if (!res.ok) throw new Error(`HTTP ${res.status} al pedir vistas/${viewName}.html`);

    const html = await res.text();
    this._cache[viewName] = html;
    return html;
  },

  /* ── Hook de limpieza: aquí matas listeners/intervalos que
       pudieran seguir "vivos" apuntando a nodos ya destruidos.
       Con delegación de eventos (ver abajo) esto normalmente
       no hace falta para clicks, pero SÍ es el lugar correcto
       para, por ejemplo, destruir instancias de Chart.js
       (chart.destroy()) antes de tirar el canvas viejo. ── */
  _onViewUnload(viewName) {
    if (viewName === 'dashboard') {
      // Ejemplo: evitar leaks de Chart.js al salir del dashboard
      if (window.ChartEngine?.destroyAll) {
        ChartEngine.destroyAll();
      }
    }
  },

  /* ── Hook de inicialización: se ejecuta UNA vez por cada
       carga de vista, después de la inyección del HTML.
       Aquí es donde "conectas" tus motores existentes con
       el DOM recién insertado. ── */
  _onViewLoaded(viewName) {
    switch (viewName) {
      case 'login':
        // AuthEngine / SessionEngine ya existen como objetos;
        // solo les pedimos que (re)enganchen sus listeners.
        SessionEngine.initLoginListeners();
        break;

      case 'dashboard':
        UIController.init();          // engancha tabs, tema, sidebar…
        AuthEngine.initListeners();
        CloudEngine.initListeners();
        HistoryEngine.initListeners();
        AbsenceEngine.initListeners();
        FilterEngine.initListeners();
        // Si el usuario ya tenía datos en memoria (DataStore),
        // repintamos sin tener que re-parsear el Excel:
        if (DataStore.hasData?.()) {
          KPIEngine.render();
          ChartEngine.renderAll();
          TableEngine.renderAll();
        }
        break;
    }
  },
};
