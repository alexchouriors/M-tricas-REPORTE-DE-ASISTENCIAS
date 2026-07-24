/* ════════════════════════════════════════════════════════════
   Usuario Rules.js
   ────────────────────────────────────────────────────────────
   Motor de Permisos de INTERFAZ (UI) por rol de usuario.

   IMPORTANTE — Qué NO hace este archivo:
   Este módulo NO filtra registros ni decide qué filas del reporte
   puede ver cada usuario; esa responsabilidad es EXCLUSIVA de
   AccessManager.js + SecurityConfig.js y no se toca aquí.

   Qué SÍ hace este archivo:
   Oculta (`d-none`), deshabilita (`disabled`) o habilita botones y
   controles del DOM (index.html) según la ETIQUETA (rol) del
   usuario en sesión: MAESTRO, EDITOR o LECTOR.

   REGLA DE ORO: Iniciar Sesión, Cerrar Sesión y Soporte están
   SIEMPRE disponibles para cualquier usuario/etiqueta — son la
   puerta de entrada/salida del sistema y jamás se gatean aquí.
   Lo mismo aplica a la interfaz "de cortesía" que el proyecto ya
   considera universal: `fileInput`, `btnTheme`, `btnDesktopView`.

   POLÍTICA DE SEGURIDAD VISUAL: fail-safe.
   - Si un usuario no está mapeado a ninguna etiqueta, se asume la
     etiqueta más restrictiva (LECTOR), nunca MAESTRO ni EDITOR.
   - Si un elemento del DOM no existe (vista distinta, HTML aún no
     cargado, etc.), se ignora sin lanzar error (`if (elemento)`).
   - Esto es control de INTERFAZ únicamente: no reemplaza ninguna
     validación de seguridad real en el backend/API (GitHub, etc.).

   NOTA DE ARQUITECTURA: script clásico (NO type="module"), igual
   que USUARIOS.JS / SecurityConfig.js / AccessManager.js. Debe
   cargarse en index.html DESPUÉS de esos tres y ANTES de app.js,
   que es quien invoca `UsuarioRules.applyUIPermissions(usuario)`
   al iniciar sesión y al restaurar una sesión existente.
   ════════════════════════════════════════════════════════════ */

/* ── Normaliza texto para comparar sin sensibilidad a mayúsculas
   ni espacios extremos/múltiples (mismo criterio que AccessManager) ── */
function _urNormalize(str) {
  return (str || '').toString().trim().toUpperCase().replace(/\s+/g, ' ');
}

/**
 * Extrae de forma segura el nombre de usuario (string) a partir de
 * `usuario`, que puede llegar como string directo o como objeto de
 * sesión ({ nombre, username, user, name }). Mismo criterio que
 * AccessManager.extractUserName(), duplicado aquí a propósito para
 * que este archivo sea 100% independiente y no dependa del orden
 * de carga de AccessManager.js.
 *
 * @param {string|Object|null|undefined} usuario
 * @returns {string|null}
 */
function _urExtractUserName(usuario) {
  if (typeof usuario === 'string') {
    const trimmed = usuario.trim();
    return trimmed !== '' ? trimmed : null;
  }
  if (usuario && typeof usuario === 'object') {
    const candidate = usuario.nombre ?? usuario.username ?? usuario.user ?? usuario.name ?? null;
    return (typeof candidate === 'string' && candidate.trim() !== '') ? candidate.trim() : null;
  }
  return null;
}

/* ════════════════════════════════════════════════════════════
   1) MAPEO DE USUARIOS → ETIQUETA (ROL)
   ────────────────────────────────────────────────────────────
   Claves comparadas de forma insensible a mayúsculas/espacios.
   Usuarios de USUARIOS.JS que NO aparecen aquí caen fail-safe en
   'LECTOR' (el rol más restrictivo), nunca en 'MAESTRO' ni 'EDITOR'.
   ════════════════════════════════════════════════════════════ */
const USER_ROLE_MAP = {
  MASTER:  'MAESTRO',
  PASTOR:  'MAESTRO',
  PASTORA: 'MAESTRO',
  RENNY:   'LECTOR',

  AIRAM:            'LECTOR',
  MAYERLIN:         'LECTOR',
  JONATHAN:         'LECTOR',
  MAGALIS:          'LECTOR',
  EDITH:           'LECTOR',
  ALEXANDRO:        'LECTOR',
  GARDYS:           'LECTOR',
  YENSI:            'LECTOR',
  SORENNYS:         'LECTOR',
  JOSEFA:           'LECTOR',
  FRANGLYS:         'LECTOR',
  DAYANA:           'LECTOR',
  YELSSY:           'LECTOR',
  'PASTOR CARLOS D':  'LECTOR',
  'PASTORA GÉNESIS':  'LECTOR',
  ANAIS:            'LECTOR',
  ATHAIS:           'LECTOR',
  JACKSON:          'LECTOR',
};

/* Rol asignado por defecto a cualquier usuario válido que no esté
   en USER_ROLE_MAP (fail-safe: el más restrictivo, nunca 'MAESTRO'). */
const DEFAULT_ROLE = 'LECTOR';

/* ── Excepción puntual: estos usuarios SÍ conservan el permiso
   CARGA_PREDETERMINADA (lo necesitan implícitamente porque su rol es
   LECTOR), pero el botón "Base de Datos" (btnAbrirDbDefault) se les
   OCULTA de la interfaz por medida de seguridad/orden, ya que no lo
   necesitan operar directamente. El permiso queda intacto para no
   romper nada que dependa de él; solo se fuerza la visibilidad del
   botón al final de applyUIPermissions(). ── */
const USUARIOS_OCULTAR_BTN_DB_DEFAULT = new Set([
  'RENNY', 'AIRAM', 'MAYERLIN', 'JONATHAN', 'MAGALIS', 'EDITH', 'ALEXANDRO',
  'GARDYS', 'YENSI', 'SORENNYS', 'JOSEFA', 'FRANGLYS', 'DAYANA',
  'YELSSY', 'PASTOR CARLOS D', 'PASTORA GÉNESIS', 'ANAIS', 'ATHAIS',
  'JACKSON',
]);

/* ════════════════════════════════════════════════════════════
   2) CLAVES DE PERMISO (una por cada función de la matriz)
   ════════════════════════════════════════════════════════════ */
const PERMISOS = {
  CARGAR_EXCEL_LOCAL:     'verCargarExcelLocal',   // fileInput (siempre permitido — ver nota abajo)
  CARGAR_GITHUB:          'cargarReporteGitHub',   // btnAbrirHistorial + botón "Cargar al Dashboard" por archivo
  DESCARGAR_GITHUB:       'descargarReporteGitHub',// enlace "Descargar" por archivo (dentro del Historial)
  SYNC_GSHEETS:           'sincronizarGSheets',    // btnConnectSheet, btnGsheetsEmpty
  DESCONECTAR_GSHEETS:    'desconectarGSheet',     // btnDisconnect
  GUARDAR_MENU:           'guardarDesdeMenu',      // btnMenuGuardar
  GUARDAR_NUBE:           'guardarEnNube',         // btnCloudSave
  CONFIGURAR_TOKEN:       'configurarTokenGitHub', // btnAuthConfig
  CARGA_PREDETERMINADA:   'cargaPredeterminada',   // btnAbrirDbDefault
  FILTRAR_VISTAS:         'filtrarVistas',         // barra de filtros + botón "Limpiar filtros"
  CAMBIAR_TEMA:           'cambiarTemaVisual',     // btnTheme (siempre permitido — ver nota abajo)
  MODO_ESCRITORIO:        'forzarModoEscritorio',  // btnDesktopView (siempre permitido — ver nota abajo)
  ELIMINAR:               'eliminarArchivo',       // btnAbrirEliminar
};

/* ════════════════════════════════════════════════════════════
   3) MATRIZ DE PERMISOS POR ETIQUETA
   ────────────────────────────────────────────────────────────
   'ALL'  → acceso total (todo lo que exista o se agregue a futuro).
   Array  → lista blanca explícita de PERMISOS.* permitidos.
   ════════════════════════════════════════════════════════════ */
const ROLE_PERMISSIONS = {

  /* MAESTRO: acceso total — incluye cualquier permiso futuro sin
     necesidad de tocar este archivo de nuevo. */
  MAESTRO: 'ALL',

  /* EDITOR: todo excepto Eliminar. */
  EDITOR: [
    PERMISOS.CARGAR_GITHUB,
    PERMISOS.DESCARGAR_GITHUB,
    PERMISOS.SYNC_GSHEETS,
    PERMISOS.DESCONECTAR_GSHEETS,
    PERMISOS.GUARDAR_MENU,
    PERMISOS.GUARDAR_NUBE,
    PERMISOS.CONFIGURAR_TOKEN,
    PERMISOS.CARGA_PREDETERMINADA,
    PERMISOS.FILTRAR_VISTAS,
  ],

  /* LECTOR: acceso restringido — sin nube, sin GSheets, sin eliminar. */
  LECTOR: [
    PERMISOS.CARGAR_GITHUB,
    PERMISOS.CONFIGURAR_TOKEN,
    PERMISOS.CARGA_PREDETERMINADA,
    PERMISOS.FILTRAR_VISTAS,
  ],
};

/* ════════════════════════════════════════════════════════════
   4) ELEMENTOS DEL DOM POR PERMISO (IDs estáticos)
   ────────────────────────────────────────────────────────────
   NOTA: `fileInput`, `btnTheme` y `btnDesktopView` (y por supuesto
   `btnCerrarSesion`/login/soporte) NO aparecen en este mapa a
   propósito: son universales y este módulo nunca los toca.
   ════════════════════════════════════════════════════════════ */
const PERMISO_A_IDS = {
  [PERMISOS.CARGAR_GITHUB]:        ['btnAbrirHistorial'],
  [PERMISOS.SYNC_GSHEETS]:         ['btnConnectSheet', 'btnGsheetsEmpty'],
  [PERMISOS.DESCONECTAR_GSHEETS]:  ['btnDisconnect'],
  [PERMISOS.GUARDAR_MENU]:         ['btnMenuGuardar'],
  [PERMISOS.GUARDAR_NUBE]:         ['btnCloudSave'],
  [PERMISOS.CONFIGURAR_TOKEN]:     ['btnAuthConfig'],
  [PERMISOS.CARGA_PREDETERMINADA]: ['btnAbrirDbDefault'],
  [PERMISOS.FILTRAR_VISTAS]:       ['filterGroup', 'filterEstado', 'filterCelula', 'filterServicio', 'filterNuevo', 'btnResetFilters'],
  [PERMISOS.ELIMINAR]:             ['btnAbrirEliminar'],
};

/* IDs de elementos de tipo <select>/<button> de filtros que, al no
   tener permiso, se DESHABILITAN en vez de ocultarse (para no romper
   el layout de la barra de filtros). El resto de PERMISO_A_IDS se
   oculta con 'd-none' además de deshabilitarse. */
const IDS_SOLO_DESHABILITAR = new Set([
  'filterGroup', 'filterEstado', 'filterCelula', 'filterServicio', 'filterNuevo',
]);

/* Contenedor donde HistoryEngine (app.js) inyecta dinámicamente, por
   cada archivo, un botón "Cargar al Dashboard" (.btn-cargar-reporte)
   y un enlace "Descargar" (a.btn-outline-success). Como se regeneran
   en cada render, se gatean vía MutationObserver (ver más abajo). */
const HISTORIAL_LIST_CONTAINER_ID = 'listaReportesContainer';

const UsuarioRules = {

  /* Guarda el último permiso de descarga resuelto, para poder
     re-aplicarlo cada vez que el Historial vuelve a renderizar su
     lista (ver _observeHistorialList()). */
  _ultimoPuedeDescargar: false,
  _historialObserverIniciado: false,

  /** Determina el rol (etiqueta) del usuario, fail-safe a LECTOR. */
  _resolveRole(usuario) {
    const userName = _urExtractUserName(usuario);
    if (!userName) {
      console.warn('[UsuarioRules] Usuario inválido/no reconocible — se aplica rol restrictivo por defecto:', DEFAULT_ROLE);
      return DEFAULT_ROLE;
    }

    const target = _urNormalize(userName);
    const matchKey = Object.keys(USER_ROLE_MAP).find(k => _urNormalize(k) === target);

    if (!matchKey) {
      console.warn(`[UsuarioRules] "${userName}" no tiene rol de UI asignado — se aplica rol restrictivo por defecto: ${DEFAULT_ROLE}`);
      return DEFAULT_ROLE;
    }

    return USER_ROLE_MAP[matchKey];
  },

  /** true si el rol resuelto tiene el permiso indicado. */
  _hasPermission(role, permiso) {
    const rule = ROLE_PERMISSIONS[role];
    if (rule === 'ALL') return true;
    if (Array.isArray(rule)) return rule.includes(permiso);
    return false; // Rol desconocido en ROLE_PERMISSIONS → fail-safe sin permisos
  },

  /** Habilita/deshabilita y muestra/oculta un elemento por su id. */
  _setElementState(id, allowed) {
    const el = document.getElementById(id);
    if (!el) return; // Fail-safe visual: nunca truena si el elemento no existe

    if (allowed) {
      el.disabled = false;
      el.removeAttribute('aria-disabled');
      if (!IDS_SOLO_DESHABILITAR.has(id)) el.classList.remove('d-none');
    } else {
      el.disabled = true;
      el.setAttribute('aria-disabled', 'true');
      if (!IDS_SOLO_DESHABILITAR.has(id)) el.classList.add('d-none');
    }
  },

  /** Habilita/deshabilita visualmente el enlace "Descargar" (<a>) de
      una fila del Historial, que no soporta `disabled` de forma nativa. */
  _setDownloadLinkState(anchorEl, allowed) {
    if (!anchorEl) return;
    if (allowed) {
      anchorEl.classList.remove('d-none');
      anchorEl.removeAttribute('aria-disabled');
      anchorEl.style.pointerEvents = '';
      anchorEl.tabIndex = 0;
    } else {
      anchorEl.classList.add('d-none');
      anchorEl.setAttribute('aria-disabled', 'true');
      anchorEl.style.pointerEvents = 'none';
      anchorEl.tabIndex = -1;
    }
  },

  /** Aplica el permiso de descarga vigente a todas las filas ya
      renderizadas del Historial (se llama tras cada re-render). */
  _applyDownloadPermissionToList() {
    const container = document.getElementById(HISTORIAL_LIST_CONTAINER_ID);
    if (!container) return;
    container.querySelectorAll('a.btn-outline-success').forEach(a => {
      this._setDownloadLinkState(a, this._ultimoPuedeDescargar);
    });
  },

  /** Observa el contenedor del Historial: cada vez que HistoryEngine
      regenera la lista de archivos, vuelve a aplicar el permiso de
      descarga a los enlaces recién creados. Se registra una sola vez. */
  _observeHistorialList() {
    if (this._historialObserverIniciado) return;
    const container = document.getElementById(HISTORIAL_LIST_CONTAINER_ID);
    if (!container || typeof MutationObserver === 'undefined') return;

    const observer = new MutationObserver(() => this._applyDownloadPermissionToList());
    observer.observe(container, { childList: true, subtree: true });
    this._historialObserverIniciado = true;
  },

  /** Punto de entrada público para cuando el modal de Historial recién
      se inyecta en el DOM (ver LazyModals.js). Como #modalHistorial y
      su #listaReportesContainer viven dentro de un <template> hasta
      el primer clic en "Historial", intentar `_observeHistorialList()`
      en el login (cuando se llama a applyUIPermissions()) siempre
      fallaba en silencio: el contenedor todavía no existía. Sin el
      observer activo, cuando HistoryEngine luego renderizaba la lista
      de archivos de forma asíncrona, el enlace "Descargar" de cada
      fila nunca se ocultaba para LECTOR — quedaba visible y funcional
      para cualquiera. Este método debe llamarse justo después de
      inyectar el modal (antes o después de HistoryEngine.init(), da
      igual el orden) para que el observer quede activo a tiempo de
      capturar esa primera renderización y todas las siguientes. */
  attachToHistorialModal() {
    this._observeHistorialList();
    this._applyDownloadPermissionToList();
  },

  /**
   * Función principal: aplica los permisos de INTERFAZ correspondientes
   * al usuario indicado. Debe llamarse:
   *   - Justo después de un inicio de sesión exitoso.
   *   - Al restaurar una sesión existente (recarga de página).
   *
   * No modifica sessionStorage, datos ni filtros de registros — solo
   * el DOM (visibilidad/estado de botones y controles).
   *
   * @param {string|Object} userName - Usuario en sesión (string o { nombre, ... })
   */
  applyUIPermissions(userName) {
    const role = this._resolveRole(userName);

    Object.entries(PERMISO_A_IDS).forEach(([permiso, ids]) => {
      const allowed = this._hasPermission(role, permiso);
      ids.forEach(id => this._setElementState(id, allowed));
    });

    /* Descarga de archivos del Historial: se gatea aparte porque los
       enlaces se generan dinámicamente por cada archivo listado. */
    this._ultimoPuedeDescargar = this._hasPermission(role, PERMISOS.DESCARGAR_GITHUB);
    this._applyDownloadPermissionToList();
    this._observeHistorialList();

    /* Excepción puntual: fuerza la ocultación de "Base de Datos" para
       ciertos usuarios aunque su rol (LECTOR) sí tenga el permiso
       CARGA_PREDETERMINADA — ver USUARIOS_OCULTAR_BTN_DB_DEFAULT. */
    const nombreNormalizado = _urNormalize(_urExtractUserName(userName) || '');
    if (USUARIOS_OCULTAR_BTN_DB_DEFAULT.has(nombreNormalizado)) {
      const btnDb = document.getElementById('btnAbrirDbDefault');
      if (btnDb) {
        btnDb.classList.add('d-none');
        btnDb.disabled = true;
        btnDb.setAttribute('aria-disabled', 'true');
      }
    }
  },
};

/* Exposición explícita en window (una `const` de script clásico no se
   agrega sola a window), para consistencia con AccessManager.js y
   fácil consumo desde app.js como `window.UsuarioRules`. */
window.UsuarioRules = UsuarioRules;
