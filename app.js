/* ================================================================
   DASHBOARD IGLESIA — app.js
   Arquitectura modular ES6:
     - DataStore: estado centralizado
     - ExcelParser: lectura y normalización del Excel
     - KPIEngine: cálculo de indicadores (extensible)
     - ChartEngine: renderizado y actualización de gráficos
     - TableEngine: renderizado y filtros de tablas
     - FilterEngine: gestión de filtros globales
     - UIController: coordinación general de la UI
================================================================ */

'use strict';

/* ────────────────────────────────────────────────────────────
   0. WORKER ENGINE — Delega el parseo pesado de SheetJS
   (XLSX.read del archivo .xlsx/.xlsm) a un Web Worker para NO
   bloquear el hilo principal ni congelar la pantalla de carga.

   El worker (excelWorker.js) hace únicamente la parte cara en CPU:
   XLSX.read() + serializar cada hoja a un objeto plano. Devuelve un
   objeto { SheetNames, Sheets } 100% compatible con lo que espera
   ExcelParser.parse(workbook) (misma forma que produce XLSX.read en
   el hilo principal), así que TODA la lógica de negocio existente
   (ExcelParser.parseMainSheet, parseExcluidosSheet, RBAC, etc.)
   queda intacta y sin cambios — el worker solo mueve el trabajo
   pesado fuera del hilo principal, no cambia qué se calcula.

   Si el navegador no soporta Web Workers, o el worker falla al
   crearse/cargar SheetJS (por ejemplo, red bloqueada), se degrada
   automáticamente a XLSX.read() síncrono en el hilo principal para
   no romper la carga de archivos. */
const WorkerEngine = {
  _worker: null,
  _seq: 0,
  _pending: new Map(),
  _unavailable: false, // true si el worker falló y debemos usar el fallback síncrono

  _getWorker() {
    if (this._worker || this._unavailable) return this._worker;
    try {
      this._worker = new Worker('excelWorker.js');
      this._worker.onmessage = (e) => {
        const { id, ok, sheetNames, sheets, error } = e.data || {};
        const pending = this._pending.get(id);
        if (!pending) return;
        this._pending.delete(id);
        if (ok) pending.resolve({ SheetNames: sheetNames, Sheets: sheets });
        else pending.reject(new Error(error || 'Error desconocido en el Web Worker de Excel.'));
      };
      this._worker.onerror = (err) => {
        console.error('[WorkerEngine] Error fatal en excelWorker.js, se usará el hilo principal como respaldo:', err);
        this._unavailable = true;
        this._pending.forEach(p => p.reject(err));
        this._pending.clear();
      };
    } catch (err) {
      console.error('[WorkerEngine] No se pudo crear el Web Worker, se usará el hilo principal como respaldo:', err);
      this._unavailable = true;
      this._worker = null;
    }
    return this._worker;
  },

  /* Fallback 100% síncrono en el hilo principal (comportamiento
     idéntico al que tenía la app antes de esta optimización). */
  _parseWorkbookSync(arrayBuffer) {
    const data = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
    return XLSX.read(data, { type: 'array', cellDates: false });
  },

  /* Punto de entrada: recibe un ArrayBuffer del .xlsx/.xlsm y
     devuelve una Promise con el "workbook" ya parseado, procesado
     en un Web Worker cuando es posible. */
  async parseWorkbookAsync(arrayBuffer) {
    if (this._unavailable || typeof Worker === 'undefined') {
      return this._parseWorkbookSync(arrayBuffer);
    }

    try {
      const worker = this._getWorker();
      if (!worker) return this._parseWorkbookSync(arrayBuffer);

      const id = ++this._seq;
      // Copiamos el buffer porque se transfiere (Transferable) al worker
      // y el original queda inutilizable en el hilo principal tras eso.
      const bufferCopy = (arrayBuffer instanceof Uint8Array ? arrayBuffer.slice() : arrayBuffer.slice(0));
      const raw = bufferCopy instanceof Uint8Array ? bufferCopy.buffer : bufferCopy;

      const result = await new Promise((resolve, reject) => {
        this._pending.set(id, { resolve, reject });
        worker.postMessage({ id, arrayBuffer: raw }, [raw]);
      });

      return result;
    } catch (err) {
      console.warn('[WorkerEngine] Falló el parseo en Web Worker, reintentando en el hilo principal:', err);
      return this._parseWorkbookSync(arrayBuffer);
    }
  },
};

/* ────────────────────────────────────────────────────────────
   1. DATA STORE — fuente única de verdad
──────────────────────────────────────────────────────────── */
const DataStore = {
  /* Datos crudos de cada hoja */
  rawMain:      [],   // Registros hoja principal (reporte activo)
  rawExcluidos: [],   // Hoja "Excluidos"
  rawNuevoEx:   [],   // Hoja "NUEVO EX"
  rawAntiguoEx: [],   // Hoja "ANTIGUO EX"

  /* Metadatos del archivo */
  reportTitle:  '',
  fileName:     '',
  rawBuffer:    null,   // ArrayBuffer del último archivo cargado localmente (para subida a GitHub)

  /* Estado de la UI */
  includeExcluidos: false,

  /* true = la vista principal está proyectando DataStore.comparisonBaseline
     (ver TrendViewEngine) en vez de rawMain/rawExcluidos. NO destruye el
     Excel original: rawMain/rawExcluidos siguen intactos, solo cambia
     qué devuelve getActiveMain() mientras esta bandera esté activa. */
  viewingTrend: false,

  /* Filtros activos */
  filters: {
    group:    '',
    estado:   '',
    celula:   '',
    servicio: '',
    nuevo:    '',
  },

  /* Devuelve los registros activos aplicando la regla de excluidos */
  getActiveMain() {
    /* Vista de Tendencia: reemplaza TEMPORALMENTE lo que ven
       TableEngine/ChartEngine/KPIEngine (todos consumen esta misma
       función), sin tocar rawMain/rawExcluidos. Se desactiva con
       TrendViewEngine.exitTrendView(), que vuelve a mostrar esta
       misma rama de código su comportamiento normal de siempre. */
    if (this.viewingTrend && this.comparisonBaseline) {
      return this.comparisonBaseline.rawMain;
    }
    if (this.includeExcluidos) {
      // Mezcla registros principales con excluidos
      return [...this.rawMain, ...this.rawExcluidos];
    }
    return this.rawMain;
  },

  /* Aplica todos los filtros sobre un array de registros */
  applyFilters(records) {
    const f = this.filters;
    return records.filter(r => {
      if (f.group    && r.grupo    !== f.group)    return false;
      if (f.estado   && r.estado.toUpperCase()  !== f.estado.toUpperCase())   return false;
      if (f.celula   && r.celula.toUpperCase()  !== f.celula.toUpperCase())   return false;
      if (f.servicio && r.servicio.toUpperCase() !== f.servicio.toUpperCase()) return false;
      if (f.nuevo === 'si'  && !r.esNuevo) return false;
      if (f.nuevo === 'no'  && r.esNuevo)  return false;
      return true;
    });
  },

  /* Filtra la tabla de "Excluidos" únicamente por el grupo ministerial
     seleccionado en el <select> de filtros (this.filters.group), SIN
     importar el rol/permiso del usuario en sesión.

     Corrige el bug donde un usuario con acceso 'ALL' (RBAC) veía
     TODOS los excluidos de TODOS los grupos aunque hubiera elegido un
     grupo específico en el dropdown, porque renderExcluidos() se
     pintaba directamente con DataStore.rawExcluidos "crudo" en vez de
     respetar el valor actual del selector.

     NOTA: esto es INDEPENDIENTE del filtrado por rol (AccessManager),
     que ya recorta DataStore.rawExcluidos al cargar el archivo para
     usuarios sin acceso 'ALL'. Este método solo añade el filtro del
     dropdown por encima de eso — para un usuario no-admin, filtrar de
     nuevo por su propio grupo no cambia nada (ya viene restringido),
     pero para un admin sí respeta ahora el grupo elegido en pantalla.
     Si el select está en "Todos los grupos" (valor vacío), devuelve
     el arreglo sin cambios. */
  /* Normaliza el nombre de un grupo ministerial para poder comparar
     el mismo grupo aunque esté escrito de forma distinta entre hojas
     del Excel (Main vs Excluidos), por ejemplo:
       "Jonathan y Mayerling"          (hoja principal)
       "MINISTRO Jonathan y Mayerling" (hoja Excluidos)
     Quita prefijos de rol (Ministro/Ministra/Líder/Líderes/etc.),
     pasa a minúsculas y colapsa espacios — NO afecta el texto que
     se muestra en pantalla, solo se usa para comparar. */
  _normalizeGrupoLabel(str) {
    return (str || '')
      .toString()
      .trim()
      .toLowerCase()
      .replace(/^(ministro|ministros|ministra|ministras|l[ií]der|l[ií]deres|pastor|pastora)\s+/i, '')
      .replace(/\s+/g, ' ')
      .trim();
  },

  /* Filtra la tabla de "Excluidos" únicamente por el grupo ministerial
     seleccionado en el <select> de filtros (this.filters.group), SIN
     importar el rol/permiso del usuario en sesión.

     Corrige el bug donde un usuario con acceso 'ALL' (RBAC) veía
     TODOS los excluidos de TODOS los grupos aunque hubiera elegido un
     grupo específico en el dropdown, porque renderExcluidos() se
     pintaba directamente con DataStore.rawExcluidos "crudo" en vez de
     respetar el valor actual del selector.

     Usa comparación NORMALIZADA (ver _normalizeGrupoLabel) en vez de
     igualdad exacta de string, porque el <select> se puebla solo con
     los encabezados de la hoja principal (rawMain), mientras que la
     hoja "Excluidos" puede tener su propio encabezado de grupo con
     una redacción ligeramente distinta para el mismo grupo (con o sin
     prefijo "MINISTRO/MINISTRA/LÍDER", mayúsculas distintas, etc.).
     Con igualdad exacta, esos registros nunca coincidían y la tabla
     de Excluidos quedaba vacía al elegir un grupo específico.

     NOTA: esto es INDEPENDIENTE del filtrado por rol (AccessManager),
     que ya recorta DataStore.rawExcluidos al cargar el archivo para
     usuarios sin acceso 'ALL'. Este método solo añade el filtro del
     dropdown por encima de eso. Si el select está en "Todos los
     grupos" (valor vacío), devuelve el arreglo sin cambios. */
  filterExcluidosByGroup(records) {
    const group = this.filters.group;
    if (!group) return records; // "Todos los grupos" → sin filtrar

    const targetNormalizado = this._normalizeGrupoLabel(group);
    return records.filter(r => r && this._normalizeGrupoLabel(r.grupo) === targetNormalizado);
  },
};


/* ────────────────────────────────────────────────────────────
   2. EXCEL PARSER — convierte la hoja a registros normalizados
──────────────────────────────────────────────────────────── */
const ExcelParser = {

  /* Normaliza un valor de celda a string seguro */
  str(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  },

  /* Lee una celda directamente del worksheet por fila/columna (0-indexed),
     prefiriendo el TEXTO TAL COMO SE VE en Excel (cell.w) sobre el valor
     crudo (cell.v). Necesario para columnas como el teléfono: si la
     celda es numérica, sheet_to_json({header:1}) devuelve cell.v (un
     número JS), que puede perder ceros a la izquierda o convertirse a
     notación exponencial con números largos. cell.w conserva el
     formato de despliegue real de Excel. */
  cellText(ws, rowIndex, colIndex) {
    const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = ws[addr];
    if (!cell) return '';
    if (typeof cell.w === 'string' && cell.w.trim() !== '') return cell.w.trim();
    return this.str(cell.v);
  },

  /* Devuelve true si el valor de celda debe considerarse "vacío"
     (null, undefined, string vacío, 0 numérico, booleano false, etc.)
     Necesario porque SheetJS puede devolver 0 o false en celdas en blanco
     dependiendo de cómo fue generado el Excel. */
  isEmpty(v) {
    if (v === null || v === undefined) return true;
    const s = String(v).trim();
    return s === '' || s === '0' || s === 'false';
  },

  /* Convierte número de serie Excel a fecha legible */
  excelDate(v) {
    if (!v) return '';
    if (typeof v === 'string' && v.includes('-')) return v.substring(0,10);
    if (typeof v === 'number') {
      const d = XLSX.SSF.parse_date_code(v);
      if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
    }
    return this.str(v);
  },

  /* ── Parsea la hoja principal del reporte ── */
  parseMainSheet(ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const records = [];

    /* Offset real de la hoja: sheet_to_json({header:1}) genera el
       array `rows` empezando en la PRIMERA fila del rango usado por
       la hoja (ws['!ref']), que no siempre es la fila/columna 1 (A).
       Si la hoja no arranca en A1, direccionar celdas "a mano" (como
       hace cellText() para el teléfono) con el índice `i` del array
       sin corregir apuntaría a la celda equivocada. Se calcula una
       sola vez y se suma a cualquier lectura directa por celda. */
    const range     = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const rowOffset = range.s.r;
    const colOffset = range.s.c;

    // Detectar título del reporte (fila 1, columna A)
    const title = this.str(rows[1]?.[0]) || this.str(rows[0]?.[0]) || '';
    DataStore.reportTitle = title;

    let currentGroup = 'Sin Grupo';

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;

      const col0 = this.str(row[0]);
      const col1 = this.str(row[1]);
      const col2 = this.str(row[2]);
      const col3 = this.str(row[3]);
      const col4 = this.str(row[4]);
      const col5 = this.str(row[5]);
      const colTelefono = this.cellText(ws, rowOffset + i, colOffset + 6);   // N°Telefónico (columna G, verificado contra el Excel real: fila 22 = encabezados, G es la 7ª letra = índice 0-based 6)

      /* Detectar encabezado de grupo ministerial:
         - La fila contiene texto en col0
         - col1 (Nombre) DEBE estar vacía — si hay nombre, es fila de persona
         - col2 (Célula) y col3 (Servicio) deben estar vacías (sin datos de asistencia)
         - NO empieza con número ni con "N°" ni con "TOTAL"
      */
      const isGroupHeader = (
        col0.length > 3 &&
        col1 === '' &&
        this.isEmpty(row[2]) &&
        this.isEmpty(row[3]) &&
        !/^\d/.test(col0) &&
        !col0.startsWith('N°') &&
        !col0.startsWith('TOTAL') &&
        !col0.startsWith('REPOR') &&
        !col0.startsWith('Tema') &&
        !col0.startsWith('Fecha')
      );

      if (isGroupHeader) {
        currentGroup = col0.trim();
        continue;
      }

      /* ── Detectar fila de datos de persona ──
         Columnas del Excel (0-indexed):
           col0 (A) = N°
           col1 (B) = Nombre
           col2 (C) = Célula    → SI | NO | NUEVO*
           col3 (D) = Servicio  → SI | NO
           col4 (E) = Estado    → tipo de miembro; NUEVO* aquí = nuevo en célula
           col5 (F) = Fecha última falta

         REGLA NUEVO (confirmada contra los KPIs del Excel):
           • Nuevo en CÉLULA   → col4 (Estado) === 'NUEVO'   (8 personas)
           • Nuevo en SERVICIO → col2 (Célula)  === 'NUEVO'   (1 persona)
           * El campo "Célula" con NUEVO indica que llegó nuevo al servicio
             y fue derivado a célula por primera vez.
      */
      const num = parseFloat(col0);
      if (!isNaN(num) && num > 0 && col1 && col1.length > 1) {
        const celVal2  = col2.toUpperCase();   // Campo Célula   (col C)
        const serVal3  = col3.toUpperCase();   // Campo Servicio (col D)
        const estadoE  = col4.toUpperCase();   // Campo Estado   (col E)

        // Nuevo en CÉLULA: su Estado (col E) dice 'NUEVO'
        const esNuevoCelula   = (estadoE === 'NUEVO');

        // Nuevo en SERVICIO: el campo Célula (col C) dice 'NUEVO'
        const esNuevoServicio = (celVal2 === 'NUEVO');

        const esNuevo = esNuevoCelula || esNuevoServicio;

        // Fecha de última falta (col F = índice 5)
        const fechaRaw = row[5];
        const fecha = this.excelDate(fechaRaw);

        records.push({
          num:             num,
          nombre:          col1,
          celula:          celVal2  || 'NO',    // valor real del campo Célula (C)
          servicio:        serVal3  || 'NO',    // valor real del campo Servicio (D)
          estado:          col4     || '',      // valor real del campo Estado (E)
          grupo:           currentGroup.trim(),
          esNuevo:         esNuevo,
          esNuevoCelula:   esNuevoCelula,       // NUEVO en célula (Estado=NUEVO)
          esNuevoServicio: esNuevoServicio,     // NUEVO en servicio (Célula=NUEVO)
          fecha:           fecha,               // fecha de última ausencia registrada
          telefono:        colTelefono,         // N°Telefónico (columna G)
          fuente:          'principal',
        });
      }
    }

    return records;
  },

  /* ── Parsea la hoja "Excluidos" ── */
  parseExcluidosSheet(ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const records = [];
    let currentGroup = 'Excluidos';

    /* Mismo criterio que parseMainSheet(): offset real de la hoja para
       poder direccionar la celda del teléfono (columna G) de forma
       robusta, incluso si esta hoja no arranca en A1. */
    const range     = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const rowOffset = range.s.r;
    const colOffset = range.s.c;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const col0 = this.str(row[0]);
      const col1 = this.str(row[1]);
      const col2 = this.str(row[2]);
      const col3 = this.str(row[3]);
      const col4 = this.str(row[4]);
      const colTelefono = this.cellText(ws, rowOffset + i, colOffset + 6); // N°Telefónico (columna G)

      // Detectar encabezado de grupo
      if (
        col0.length > 3 &&
        col1 === '' &&
        this.isEmpty(row[2]) &&
        !/^\d/.test(col0) &&
        !col0.startsWith('TOTAL')
      ) {
        currentGroup = col0.trim();
        continue;
      }

      const num = parseFloat(col0);
      if (!isNaN(num) && num > 0 && col1 && col1.length > 1) {
        const fechaRaw = row[5];
        records.push({
          num:      num,
          nombre:   col1,
          celula:   col2.toUpperCase() || 'NO',
          servicio: col3.toUpperCase() || 'NO',
          estado:   col4 || '',
          grupo:    currentGroup.trim(),
          esNuevo:  false,
          fecha:    this.excelDate(fechaRaw || row[5]),
          telefono: colTelefono,          // N°Telefónico (columna G)
          fuente:   'excluidos',
        });
      }
    }
    return records;
  },

  /* ── Parsea la hoja "NUEVO EX" ── */
  parseNuevoExSheet(ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const records = [];
    let currentGroup = 'NUEVO EX';

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const col0 = this.str(row[0]);
      const col1 = this.str(row[1]);
      const col2 = this.str(row[2]);
      const col3 = this.str(row[3]);
      const col4 = this.str(row[4]);

      // Encabezado grupo
      if (
        col0.length > 3 &&
        col1 === '' &&
        this.isEmpty(row[2]) &&
        !/^\d/.test(col0) &&
        !col0.startsWith('TOTAL')
      ) {
        currentGroup = col0.trim();
        continue;
      }

      const num = parseFloat(col0);
      if (!isNaN(num) && num > 0 && col1 && col1.length > 1) {
        records.push({
          num:      num,
          nombre:   col1,
          celula:   col2.toUpperCase() || 'NO',
          servicio: col3.toUpperCase() || 'NO',
          estado:   col4 || '',
          grupo:    currentGroup.trim(),
          esNuevo:  true,
          fecha:    this.excelDate(row[5]),
          fuente:   'nuevo_ex',
        });
      }
    }
    return records;
  },

  /* ── Parsea la hoja "ANTIGUO EX" ── */
  parseAntiguoExSheet(ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const records = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const col0 = this.str(row[0]);
      const col1 = this.str(row[1]);
      const col2 = this.str(row[2]);

      // Solo filas con número + nombre + estado
      const num = parseFloat(col0);
      if (!isNaN(num) && num > 0 && col1 && col1.length > 1) {
        records.push({
          num:    num,
          nombre: col1,
          estado: col2 || 'NO',
          fecha:  this.excelDate(row[3]),
          fuente: 'antiguo_ex',
        });
      }
    }
    return records;
  },

  /* ── Punto de entrada principal ── */
  parse(workbook) {
    const sheetNames = workbook.SheetNames;

    /* Busca una hoja por nombre parcial (insensible a mayúsculas) */
    const find = (keyword) => {
      const name = sheetNames.find(n =>
        n.toLowerCase().includes(keyword.toLowerCase())
      );
      return name ? workbook.Sheets[name] : null;
    };

    /* La hoja principal es la primera que NO sea de las especiales */
    const specialNames = ['excluidos','nuevo ex','antiguo ex','hoja','curri'];
    const mainSheetName = sheetNames.find(n =>
      !specialNames.some(s => n.toLowerCase().includes(s))
    ) || sheetNames[0];

    const wsMain      = workbook.Sheets[mainSheetName];
    const wsExcl      = find('excluido');
    const wsNuevoEx   = find('nuevo ex');
    const wsAntiguoEx = find('antiguo ex');

    DataStore.rawMain      = wsMain      ? this.parseMainSheet(wsMain)         : [];
    DataStore.rawExcluidos = wsExcl      ? this.parseExcluidosSheet(wsExcl)    : [];
    DataStore.rawNuevoEx   = wsNuevoEx   ? this.parseNuevoExSheet(wsNuevoEx)   : [];
    DataStore.rawAntiguoEx = wsAntiguoEx ? this.parseAntiguoExSheet(wsAntiguoEx) : [];

    /* ────────────────────────────────────────────────────────────
       CONTROL DE ACCESO (RBAC) — Filtra los registros según el
       "Grupo Ministerial" permitido para el usuario en sesión,
       ANTES de que lleguen a FilterEngine/UIController/ChartEngine.

       Se aplica aquí, en el único punto de entrada de TODO parseo
       de Excel, para cubrir automáticamente:
         • Carga manual/local (input de archivo / botón "Cargar Excel")
         • Historial de GitHub (HistoryEngine)
         • Auto-carga del archivo predeterminado (AutoLoadEngine)
       sin tener que duplicar la llamada en cada uno de esos flujos.

       window.AccessManager lo expone AccessManager.js (módulo ES6,
       ver index.html). rawAntiguoEx no tiene campo `grupo` (esa hoja
       no distingue por grupo ministerial), así que no se filtra: no
       hay información suficiente para clasificarla con seguridad.
    ──────────────────────────────────────────────────────────── */
    if (window.AccessManager) {
      const usuarioActivo = AuditEngine.getUser();
      DataStore.rawMain      = window.AccessManager.applyFilter(DataStore.rawMain,      usuarioActivo);
      DataStore.rawExcluidos = window.AccessManager.applyFilter(DataStore.rawExcluidos, usuarioActivo);
      DataStore.rawNuevoEx   = window.AccessManager.applyFilter(DataStore.rawNuevoEx,   usuarioActivo);
    } else {
      /* FAIL-CLOSED: si el módulo RBAC no cargó (fallo de red, bloqueo
         del <script type="module">, etc.), NO se debe mostrar el reporte
         sin filtrar — eso sería otra vía de fail-open. Se vacían los
         datos y se avisa por consola; el usuario verá un dashboard sin
         registros en vez de datos que no le corresponden. */
      console.error('[ExcelParser] AccessManager no disponible — filtro RBAC no aplicado. Se bloquean los datos por seguridad (fail-closed).');
      DataStore.rawMain      = [];
      DataStore.rawExcluidos = [];
      DataStore.rawNuevoEx   = [];
    }
  },

  /**
   * Variante de solo-lectura de parse(): parsea un workbook y devuelve
   * un objeto NUEVO { rawMain, rawExcluidos, rawNuevoEx } sin tocar
   * DataStore en ningún momento — el reporte actualmente cargado en
   * el dashboard queda intacto. Usada por ComparativaEngine para leer
   * un archivo histórico sin sustituir el reporte activo.
   *
   * Aplica AccessManager.applyFilter() de forma OBLIGATORIA sobre
   * rawMain, con el mismo criterio fail-closed que parse(): sin
   * AccessManager disponible, rawMain se vacía por seguridad.
   *
   * @param {Object} workbook - Workbook ya leído por XLSX.read()
   * @returns {{rawMain: Array, rawExcluidos: Array, rawNuevoEx: Array}}
   */
  parseStandalone(workbook) {
    const sheetNames = workbook.SheetNames;

    const find = (keyword) => {
      const name = sheetNames.find(n =>
        n.toLowerCase().includes(keyword.toLowerCase())
      );
      return name ? workbook.Sheets[name] : null;
    };

    const specialNames = ['excluidos', 'nuevo ex', 'antiguo ex', 'hoja', 'curri'];
    const mainSheetName = sheetNames.find(n =>
      !specialNames.some(s => n.toLowerCase().includes(s))
    ) || sheetNames[0];

    const wsMain    = workbook.Sheets[mainSheetName];
    const wsExcl    = find('excluido');
    const wsNuevoEx = find('nuevo ex');

    let rawMain        = wsMain    ? this.parseMainSheet(wsMain)       : [];
    const rawExcluidos = wsExcl    ? this.parseExcluidosSheet(wsExcl)  : [];
    const rawNuevoEx   = wsNuevoEx ? this.parseNuevoExSheet(wsNuevoEx) : [];

    /* Control de acceso OBLIGATORIO — mismo criterio fail-closed que
       parse(): sin AccessManager disponible, no se muestra nada. */
    if (window.AccessManager) {
      const usuarioActivo = AuditEngine.getUser();
      rawMain = window.AccessManager.applyFilter(rawMain, usuarioActivo);
    } else {
      console.error('[ExcelParser] AccessManager no disponible — comparativa histórica bloqueada por seguridad (fail-closed).');
      rawMain = [];
    }

    return { rawMain, rawExcluidos, rawNuevoEx };
  },
};


/* ────────────────────────────────────────────────────────────
   3. KPI ENGINE — calcula todos los indicadores
   Para añadir nuevos KPIs en el futuro: agregar métodos aquí
   y llamarlos desde compute().
──────────────────────────────────────────────────────────── */
const KPIEngine = {

  /* ── Predicados puros por métrica ──
     Única fuente de verdad: compute() los usa para CONTAR y
     getRecordsByMetric() los usa para LISTAR. Así el número que se ve
     en la tarjeta y los nombres que se ven en el modal de detalle
     nunca pueden desincronizarse entre sí. */
  metricPredicates: {
    total:          () => true,
    celulasSI:      r => r.celula === 'SI',
    celulasNO:      r => r.celula === 'NO',
    servicioSI:     r => r.servicio === 'SI',
    servicioNO:     r => r.servicio === 'NO',
    ambosSI:        r => r.celula === 'SI' && r.servicio === 'SI',
    ambosNO:        r => r.celula === 'NO' && r.servicio === 'NO',
    nuevosCelula:   r => !!r.esNuevoCelula,
    nuevosServicio: r => !!r.esNuevoServicio,
  },

  /**
   * Devuelve el subconjunto de `records` que conforma la métrica
   * indicada (mismo criterio exacto que compute()). Es una función
   * de solo lectura: no muta `records` ni ningún estado de DataStore.
   *
   * @param {Array<Object>} records - Normalmente el mismo array ya
   *   filtrado que se le pasa a compute() (post AccessManager + filtros UI)
   * @param {string} metricKey - Una de las claves de `metricPredicates`
   * @returns {Array<Object>} Registros que cumplen la métrica
   */
  getRecordsByMetric(records, metricKey) {
    if (!Array.isArray(records)) return [];
    const predicate = this.metricPredicates[metricKey];
    if (typeof predicate !== 'function') {
      console.warn(`[KPIEngine] Métrica desconocida: "${metricKey}"`);
      return [];
    }
    return records.filter(predicate);
  },

  /* Calcula todos los KPIs sobre los registros filtrados */
  compute(records) {
    const total = records.length;

    // Asistencia célula
    const celulasSI   = records.filter(this.metricPredicates.celulasSI).length;
    const celulasNO   = records.filter(this.metricPredicates.celulasNO).length;
    const celulasNUEVO = records.filter(r => r.celula === 'NUEVO').length;

    // Asistencia servicio
    const servicioSI   = records.filter(this.metricPredicates.servicioSI).length;
    const servicioNO   = records.filter(this.metricPredicates.servicioNO).length;
    const servicioNUEVO = records.filter(r => r.servicio === 'NUEVO').length;

    // Asistencia ambos — criterio EXACTO del Excel:
    // célula = 'SI' estricto  AND  servicio = 'SI' estricto
    // NUEVO *no* cuenta: un nuevo en servicio no asistió a célula y viceversa
    const ambosSI = records.filter(this.metricPredicates.ambosSI).length;

    // Inasistencia total — ausente en ambos (NUEVO tampoco cuenta aquí)
    const ambosNO = records.filter(this.metricPredicates.ambosNO).length;

    // ── Nuevos (usando los flags corregidos del parser) ──
    // esNuevoCelula   → Estado (col E) = 'NUEVO'  (nuevos integrados a célula)
    // esNuevoServicio → Célula (col C) = 'NUEVO'  (nuevos que llegaron al servicio)
    const nuevosCelula   = records.filter(this.metricPredicates.nuevosCelula).length;
    const nuevosServicio = records.filter(this.metricPredicates.nuevosServicio).length;
    const totalNuevos    = records.filter(r => r.esNuevo).length;

    // Porcentajes (seguros ante división por cero)
    const pct = (a, b) => b === 0 ? 0 : Math.round((a / b) * 100 * 10) / 10;

    return {
      total,
      celulasSI,
      celulasNO,
      celulasSIPct:  pct(celulasSI, total),
      celulasNOPct:  pct(celulasNO, total),
      servicioSI,
      servicioNO,
      servicioSIPct: pct(servicioSI, total),
      servicioNOPct: pct(servicioNO, total),
      ambosSI,
      ambosNO,
      ambosSIPct:    pct(ambosSI, total),
      ambosNOPct:    pct(ambosNO, total),
      nuevosCelula,
      nuevosServicio,
      totalNuevos,
      pctGeneral:    pct(celulasSI + servicioSI, total * 2),
      pctCelula:     pct(celulasSI, total),
      pctServicio:   pct(servicioSI, total),

      // --- Datos para gráficos ---
      byGroup: this.byGroup(records),
    };
  },

  /* Agrega datos por grupo ministerial */
  byGroup(records) {
    const groups = {};
    records.forEach(r => {
      if (!groups[r.grupo]) {
        groups[r.grupo] = {
          si: 0,          // célula = SI
          no: 0,          // célula = NO (y no son nuevos)
          nuevosCel: 0,   // Estado (col E) = NUEVO → nuevo en célula
          nuevosSrv: 0,   // Célula (col C) = NUEVO → nuevo en servicio
          siSrv: 0,       // servicio = SI
          noSrv: 0,       // servicio = NO
          total: 0,
        };
      }
      const g = groups[r.grupo];
      g.total++;

      // Asistencia a célula: campo Célula (col C)
      if (r.celula === 'SI')  g.si++;
      if (r.celula === 'NO')  g.no++;

      // Asistencia a servicio: campo Servicio (col D)
      if (r.servicio === 'SI') g.siSrv++;
      if (r.servicio === 'NO') g.noSrv++;

      // Nuevos (por sus flags específicos)
      if (r.esNuevoCelula)   g.nuevosCel++;
      if (r.esNuevoServicio) g.nuevosSrv++;
    });
    return groups;
  },

  /**
   * NUEVO — método aditivo, no reemplaza ni modifica compute().
   * Calcula el delta (variación) entre dos snapshots de KPIs ya
   * calculados por compute(): el "actual" (reporte cargado ahora
   * mismo) y el "anterior" (línea base de tendencia elegida desde
   * el Historial — ver TrendEngine). Usa la misma fórmula que la
   * Comparativa Histórica: ((Actual - Anterior) / Anterior) * 100.
   *
   * @param {Object} kpisActual   - Resultado de compute() sobre el reporte activo
   * @param {Object} kpisAnterior - Resultado de compute() sobre la línea base
   * @returns {Object} { [metric]: { value, pct, direction, text } }
   *   direction: 'up' | 'down' | 'flat'
   *   text: cadena lista para mostrar, p. ej. "+12.5%", "-4%", "N/A"
   */
  computeDelta(kpisActual, kpisAnterior) {
    /* Mismo set de métricas "de conteo" que ya usan las tarjetas
       clickeables (KPI_DETAIL_MAP) — las de porcentaje (pctGeneral,
       etc.) no aplican aquí porque no representan una cantidad de
       personas 1-a-1. */
    const metrics = Object.values(KPI_DETAIL_MAP).map(m => m.metric);
    const deltas = {};

    metrics.forEach(metric => {
      const actual   = kpisActual?.[metric]   ?? 0;
      const anterior = kpisAnterior?.[metric] ?? 0;

      let pct = null;
      if (anterior !== 0) {
        pct = Math.round(((actual - anterior) / anterior) * 100 * 10) / 10;
      } else if (actual === 0) {
        pct = 0;
      }
      // anterior === 0 && actual !== 0 → pct queda en null ("N/A": crecimiento indefinido)

      const direction = pct === null ? 'flat' : pct > 0 ? 'up' : pct < 0 ? 'down' : 'flat';
      const text = pct === null ? 'N/A' : `${pct > 0 ? '+' : ''}${pct}%`;

      deltas[metric] = { value: actual - anterior, pct, direction, text };
    });

    return deltas;
  },
};


/* ────────────────────────────────────────────────────────────
   MAPA DE TARJETAS KPI CLICKEABLES → MÉTRICA + TÍTULO DEL MODAL
   ────────────────────────────────────────────────────────────
   Cada entrada vincula el id del <div class="kpi-value" id="..."> ya
   existente en index.html con la clave de KPIEngine.metricPredicates
   y el título legible que se muestra en el encabezado del modal.
   Las tarjetas de porcentaje (kpiPctGeneral/Celula/Servicio) NO están
   aquí a propósito: no representan una lista de personas 1-a-1.
──────────────────────────────────────────────────────────── */
const KPI_DETAIL_MAP = {
  kpiTotal:          { metric: 'total',          title: 'Total Registrados' },
  kpiCelulasSI:      { metric: 'celulasSI',       title: 'Asistencia a Célula' },
  kpiCelulasNO:      { metric: 'celulasNO',       title: 'Inasistencia a Célula' },
  kpiServicioSI:     { metric: 'servicioSI',      title: 'Asistencia a Servicio' },
  kpiServicioNO:     { metric: 'servicioNO',      title: 'Inasistencia a Servicio' },
  kpiAmbosSI:        { metric: 'ambosSI',         title: 'Asistió a Ambos' },
  kpiAmbosNO:        { metric: 'ambosNO',         title: 'Ausentes en Ambos' },
  kpiNuevosCelula:   { metric: 'nuevosCelula',    title: 'Nuevos en Célula' },
  kpiNuevosServicio: { metric: 'nuevosServicio',  title: 'Nuevos en Servicio' },
};


/* ────────────────────────────────────────────────────────────
   MODAL ENGINE — popup de detalle de personas por tarjeta KPI
   ────────────────────────────────────────────────────────────
   Módulo independiente y de solo-DOM: no calcula nada por su cuenta,
   solo recibe (título, lista de registros) y los pinta. La lista de
   registros SIEMPRE llega ya calculada por KPIEngine.getRecordsByMetric()
   sobre el mismo array filtrado que usa el resto del dashboard, así que
   nunca puede desincronizarse ni exponer datos fuera del alcance del
   usuario en sesión (AccessManager ya se aplicó antes, en DataStore).

   Se inyecta en el DOM una sola vez (lazy init) y se reutiliza en
   cada apertura, igual que el resto de overlays del proyecto.
──────────────────────────────────────────────────────────── */
const ModalEngine = {

  _initialized: false,
  _closeTimeout: null,
  _flashTimeout: null,
  _jsPdfPromise: null,

  /* Estado de la apertura actual (se resetea en cada open()) */
  _currentTitle: '',
  _currentPersonas: [],   // [{ nombre, grupo }], siempre ordenado por nombre
  _searchQuery: '',
  _groupMode: false,      // true = agrupado por "grupo", false = lista plana

  /* Crea el markup del modal una sola vez y lo agrega a <body> */
  _ensureBuilt() {
    if (this._initialized) return;

    const backdrop = document.createElement('div');
    backdrop.id = 'kpiDetailModal';
    backdrop.className = 'kpi-modal-backdrop d-none';
    backdrop.innerHTML = `
      <div class="kpi-modal" role="dialog" aria-modal="true" aria-labelledby="kpiModalTitle">
        <div class="kpi-modal-header">
          <div class="kpi-modal-heading">
            <div class="kpi-modal-title-row">
              <h3 class="kpi-modal-title" id="kpiModalTitle"></h3>
              <div class="kpi-export-dropdown" id="kpiExportDropdown">
                <button type="button" class="kpi-export-btn" id="kpiExportBtn"
                        aria-haspopup="true" aria-expanded="false" title="Exportar la lista visible">
                  <i class="bi bi-download"></i>
                  <span id="kpiExportLabel">EXPORTAR LISTA</span>
                  <i class="bi bi-chevron-down kpi-export-caret"></i>
                </button>
                <ul class="kpi-export-menu d-none" id="kpiExportMenu" role="menu">
                  <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="pdf"><i class="bi bi-file-earmark-pdf"></i> Exportar a PDF</button></li>
                  <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="word"><i class="bi bi-file-earmark-word"></i> Exportar a Word (.doc)</button></li>
                  <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="txt"><i class="bi bi-file-earmark-text"></i> Exportar a Archivo de Texto (.txt)</button></li>
                  <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="copy"><i class="bi bi-clipboard"></i> Copiar al Portapapeles</button></li>
                </ul>
              </div>
            </div>
            <span class="kpi-modal-count" id="kpiModalCount"></span>
          </div>
          <button type="button" class="kpi-modal-close" id="kpiModalClose" aria-label="Cerrar">
            <i class="bi bi-x-lg"></i>
          </button>
        </div>
        <div class="kpi-modal-toolbar">
          <div class="kpi-modal-search">
            <i class="bi bi-search"></i>
            <input type="text" id="kpiModalSearch" placeholder="Buscar por nombre o grupo..." autocomplete="off" />
          </div>
          <button type="button" class="kpi-modal-group-toggle" id="kpiModalGroupToggle"
                  title="Organizar por grupo" aria-pressed="false">
            <i class="bi bi-diagram-3"></i>
            <span>Agrupar</span>
          </button>
        </div>
        <div class="kpi-modal-body">
          <ul class="kpi-modal-list" id="kpiModalList"></ul>
        </div>
      </div>`;
    document.body.appendChild(backdrop);

    // Cierra al hacer clic fuera del cuadro (sobre el backdrop)
    backdrop.addEventListener('click', (e) => {
      if (e.target === backdrop) this.close();
    });

    // Cierra con el botón "X"
    backdrop.querySelector('#kpiModalClose')?.addEventListener('click', () => this.close());

    // Cierra con la tecla Escape, solo si el modal está visible
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && backdrop.classList.contains('kpi-modal-visible')) {
        const menu = document.getElementById('kpiExportMenu');
        if (menu && !menu.classList.contains('d-none')) { this._toggleExportMenu(false); return; }
        this.close();
      }
    });

    // Dropdown "EXPORTAR LISTA"
    const exportBtn  = backdrop.querySelector('#kpiExportBtn');
    const exportMenu = backdrop.querySelector('#kpiExportMenu');
    exportBtn?.addEventListener('click', (e) => {
      e.stopPropagation();
      this._toggleExportMenu(exportMenu.classList.contains('d-none'));
    });
    exportMenu?.addEventListener('click', (e) => {
      const item = e.target.closest('[data-export]');
      if (!item) return;
      e.stopPropagation();
      this._toggleExportMenu(false);
      this._handleExport(item.getAttribute('data-export'));
    });
    // Cierra el menú al hacer clic en cualquier otra parte del modal
    backdrop.addEventListener('click', (e) => {
      if (!e.target.closest('#kpiExportDropdown')) this._toggleExportMenu(false);
    });

    // Buscador: filtra en vivo por nombre o grupo (sin distinguir acentos/mayúsculas)
    backdrop.querySelector('#kpiModalSearch')?.addEventListener('input', (e) => {
      this._searchQuery = e.target.value;
      this._render();
    });

    // Toggle "Agrupar": reorganiza la lista por grupo ministerial
    const groupToggle = backdrop.querySelector('#kpiModalGroupToggle');
    groupToggle?.addEventListener('click', () => {
      this._groupMode = !this._groupMode;
      groupToggle.classList.toggle('kpi-modal-toggle-active', this._groupMode);
      groupToggle.setAttribute('aria-pressed', String(this._groupMode));
      this._render();
    });

    this._initialized = true;
  },

  /* Normaliza texto para buscar sin sensibilidad a acentos/mayúsculas */
  _normalizeSearch(str) {
    return (str || '')
      .toString()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .trim();
  },

  /**
   * Abre el modal mostrando la lista de nombres de `records`.
   * @param {string} title - Título legible (ej. "Asistencia a Célula")
   * @param {Array<Object>} records - Registros a listar (deben tener `.nombre`)
   */
  open(title, records) {
    this._ensureBuilt();
    clearTimeout(this._closeTimeout);

    const backdrop = document.getElementById('kpiDetailModal');
    const titleEl  = document.getElementById('kpiModalTitle');
    if (!backdrop || !titleEl) return; // fail-safe visual

    this._currentTitle = title;
    this._currentPersonas = (Array.isArray(records) ? records : [])
      .map(r => ({
        nombre: (r && r.nombre) ? String(r.nombre).trim() : '',
        grupo:  (r && r.grupo)  ? String(r.grupo).trim()  : '',
      }))
      .filter(p => p.nombre !== '')
      .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));

    // Reinicia buscador y modo de agrupación en cada apertura
    this._searchQuery = '';
    this._groupMode = false;
    const searchInput = document.getElementById('kpiModalSearch');
    if (searchInput) searchInput.value = '';
    const groupToggle = document.getElementById('kpiModalGroupToggle');
    if (groupToggle) {
      groupToggle.classList.remove('kpi-modal-toggle-active');
      groupToggle.setAttribute('aria-pressed', 'false');
    }

    titleEl.textContent = title;
    this._toggleExportMenu(false);
    this._render();

    backdrop.classList.remove('d-none');
    // Fuerza reflow antes de agregar la clase de transición (fade + scale-in)
    void backdrop.offsetWidth;
    backdrop.classList.add('kpi-modal-visible');
    document.body.classList.add('kpi-modal-open'); // bloquea el scroll de fondo
  },

  /**
   * Vuelve a pintar la lista a partir de `_currentPersonas`, aplicando el
   * texto de búsqueda actual y el modo de organización (plano o agrupado
   * por "grupo"). Se llama en open() y cada vez que cambia la búsqueda
   * o el toggle de agrupar, sin volver a tocar KPIEngine/DataStore.
   */
  _render() {
    const listEl  = document.getElementById('kpiModalList');
    const countEl = document.getElementById('kpiModalCount');
    if (!listEl || !countEl) return;

    const personas = this._getVisiblePersonas();

    const total = this._currentPersonas.length;
    const shown = personas.length;
    countEl.textContent = (shown === total)
      ? `${total} persona${total === 1 ? '' : 's'}`
      : `${shown} de ${total} persona${total === 1 ? '' : 's'}`;

    listEl.innerHTML = '';

    if (personas.length === 0) {
      const li = document.createElement('li');
      li.className = 'kpi-modal-empty';
      li.textContent = this._searchQuery.trim() !== ''
        ? 'Ninguna persona coincide con la búsqueda.'
        : 'No hay personas registradas en esta categoría.';
      listEl.appendChild(li);
      return;
    }

    const frag = document.createDocumentFragment();

    if (this._groupMode) {
      // Agrupa por "grupo" (sin duplicar el nombre del grupo en cada fila),
      // ordenando los grupos alfabéticamente y, dentro de cada uno, por nombre
      const grupos = new Map();
      personas.forEach(p => {
        const key = p.grupo !== '' ? p.grupo : 'Sin grupo asignado';
        if (!grupos.has(key)) grupos.set(key, []);
        grupos.get(key).push(p);
      });

      const clavesOrdenadas = Array.from(grupos.keys())
        .sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));

      let contador = 1;
      clavesOrdenadas.forEach(key => {
        const personasDelGrupo = grupos.get(key);

        const header = document.createElement('li');
        header.className = 'kpi-modal-group-header';
        header.textContent = `${key} (${personasDelGrupo.length})`;
        frag.appendChild(header);

        personasDelGrupo.forEach(p => {
          frag.appendChild(this._buildItem(p, contador++, { mostrarGrupo: false }));
        });
      });
    } else {
      // Lista plana ordenada por nombre, mostrando el grupo junto al nombre
      personas.forEach((p, i) => {
        frag.appendChild(this._buildItem(p, i + 1, { mostrarGrupo: true }));
      });
    }

    listEl.appendChild(frag);
  },

  /* Construye un <li> de persona; `mostrarGrupo` oculta el grupo cuando
     ya se muestra como encabezado (modo agrupado). */
  _buildItem(persona, index, { mostrarGrupo }) {
    const li = document.createElement('li');
    li.className = 'kpi-modal-item';

    const badge = document.createElement('span');
    badge.className = 'kpi-modal-item-index';
    badge.textContent = String(index);
    li.appendChild(badge);

    const name = document.createElement('span');
    name.className = 'kpi-modal-item-name';
    name.textContent = persona.nombre; // textContent: nunca interpreta HTML
    li.appendChild(name);

    if (mostrarGrupo && persona.grupo !== '') {
      const group = document.createElement('span');
      group.className = 'kpi-modal-item-group';
      group.textContent = `— ${persona.grupo}`;
      li.appendChild(group);
    }

    return li;
  },

  /* Devuelve las personas EXACTAMENTE como se ven ahora en el modal
     (aplicando la búsqueda activa). Lo usan _render() y la exportación,
     así que lista visible y lista exportada nunca se desincronizan. */
  _getVisiblePersonas() {
    const query = this._normalizeSearch(this._searchQuery);
    return query === ''
      ? this._currentPersonas
      : this._currentPersonas.filter(p =>
          this._normalizeSearch(p.nombre).includes(query) ||
          this._normalizeSearch(p.grupo).includes(query)
        );
  },

  /* Estructura de exportación: respeta búsqueda activa y modo agrupado.
     Devuelve { grouped, rows, sections }:
       - plano:    rows = [{ n, nombre, grupo }]
       - agrupado: sections = [{ grupo, items: [{ n, nombre }] }] */
  _getExportData() {
    const personas = this._getVisiblePersonas();

    if (!this._groupMode) {
      return {
        grouped: false,
        total: personas.length,
        rows: personas.map((p, i) => ({ n: i + 1, nombre: p.nombre, grupo: p.grupo })),
        sections: [],
      };
    }

    const grupos = new Map();
    personas.forEach(p => {
      const key = p.grupo !== '' ? p.grupo : 'Sin grupo asignado';
      if (!grupos.has(key)) grupos.set(key, []);
      grupos.get(key).push(p);
    });

    let contador = 1;
    const sections = Array.from(grupos.keys())
      .sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }))
      .map(key => ({
        grupo: key,
        items: grupos.get(key).map(p => ({ n: contador++, nombre: p.nombre })),
      }));

    return { grouped: true, total: personas.length, rows: [], sections };
  },

  /* Texto apilado: un registro por línea (con grupo si aplica) */
  _buildPlainText() {
    const d = this._getExportData();
    const lines = [];
    lines.push(this._currentTitle.toUpperCase());
    lines.push(`Total: ${d.total} persona${d.total === 1 ? '' : 's'}`);
    if (this._searchQuery.trim() !== '') lines.push(`Filtro de búsqueda: "${this._searchQuery.trim()}"`);
    lines.push('');

    if (d.grouped) {
      d.sections.forEach((sec, idx) => {
        if (idx > 0) lines.push('');
        lines.push(`${sec.grupo.toUpperCase()} (${sec.items.length})`);
        sec.items.forEach(it => lines.push(`${it.n}. ${it.nombre}`));
      });
    } else {
      d.rows.forEach(r => lines.push(r.grupo !== '' ? `${r.n}. ${r.nombre} — ${r.grupo}` : `${r.n}. ${r.nombre}`));
    }
    return lines.join('\n');
  },

  _escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  },

  /* HTML de tabla compatible con Word (.doc) */
  _buildWordHtml() {
    const d = this._getExportData();
    const esc = (v) => this._escapeHtml(v);
    const cell = 'border:1px solid #999;padding:4px 8px;';
    let body = '';

    if (d.grouped) {
      body += `<tr><th style="${cell}background:#e8eefc;width:40px;">#</th><th style="${cell}background:#e8eefc;">Nombre</th></tr>`;
      d.sections.forEach(sec => {
        body += `<tr><td colspan="2" style="${cell}background:#d9d9d9;font-weight:bold;">${esc(sec.grupo)} (${sec.items.length})</td></tr>`;
        sec.items.forEach(it => {
          body += `<tr><td style="${cell}">${it.n}</td><td style="${cell}">${esc(it.nombre)}</td></tr>`;
        });
      });
    } else {
      body += `<tr><th style="${cell}background:#e8eefc;width:40px;">#</th><th style="${cell}background:#e8eefc;">Nombre</th><th style="${cell}background:#e8eefc;">Grupo</th></tr>`;
      d.rows.forEach(r => {
        body += `<tr><td style="${cell}">${r.n}</td><td style="${cell}">${esc(r.nombre)}</td><td style="${cell}">${esc(r.grupo)}</td></tr>`;
      });
    }

    const filtro = this._searchQuery.trim() !== ''
      ? `<p style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#555;">Filtro de búsqueda: &quot;${esc(this._searchQuery.trim())}&quot;</p>`
      : '';

    return `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>${esc(this._currentTitle)}</title></head>
<body style="font-family:Calibri,Arial,sans-serif;">
<h2 style="font-family:Calibri,Arial,sans-serif;">${esc(this._currentTitle)}</h2>
<p style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#555;">Total: ${d.total} persona${d.total === 1 ? '' : 's'}</p>
${filtro}
<table style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;">${body}</table>
</body></html>`;
  },

  /* Nombre de archivo seguro: titulo_AAAA-MM-DD.ext */
  _buildFileName(ext) {
    const base = this._normalizeSearch(this._currentTitle)
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '') || 'lista';
    const hoy = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${base}_${hoy.getFullYear()}-${pad(hoy.getMonth() + 1)}-${pad(hoy.getDate())}.${ext}`;
  },

  /* Descarga un Blob mediante un <a download> temporal */
  _downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 1000);
  },

  /* Carga jsPDF una sola vez, de forma dinámica, desde el mismo CDN
     que usa el resto del proyecto. */
  _loadJsPdf() {
    if (window.jspdf && window.jspdf.jsPDF) return Promise.resolve(window.jspdf.jsPDF);
    if (this._jsPdfPromise) return this._jsPdfPromise;

    this._jsPdfPromise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
      script.async = true;
      script.onload = () => {
        if (window.jspdf && window.jspdf.jsPDF) resolve(window.jspdf.jsPDF);
        else { this._jsPdfPromise = null; reject(new Error('jsPDF no disponible')); }
      };
      script.onerror = () => {
        this._jsPdfPromise = null;
        script.remove();
        reject(new Error('No se pudo cargar jsPDF'));
      };
      document.head.appendChild(script);
    });
    return this._jsPdfPromise;
  },

  async _exportPdf() {
    const JsPDF = await this._loadJsPdf();
    const d = this._getExportData();
    const doc = new JsPDF({ unit: 'pt', format: 'letter' });

    const marginX = 48;
    const marginTop = 56;
    const marginBottom = 48;
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const maxW = pageW - marginX * 2;
    const lineH = 15;
    let y = marginTop;

    const ensureSpace = (needed) => {
      if (y + needed > pageH - marginBottom) {
        doc.addPage();
        y = marginTop;
      }
    };

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.splitTextToSize(this._currentTitle, maxW).forEach(l => {
      ensureSpace(22);
      doc.text(l, marginX, y);
      y += 22;
    });

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Total: ${d.total} persona${d.total === 1 ? '' : 's'}`, marginX, y);
    y += 14;
    if (this._searchQuery.trim() !== '') {
      doc.text(`Filtro de búsqueda: "${this._searchQuery.trim()}"`, marginX, y);
      y += 14;
    }
    y += 8;
    doc.setTextColor(0);
    doc.setFontSize(11);

    const writeLine = (text) => {
      doc.splitTextToSize(text, maxW).forEach(l => {
        ensureSpace(lineH);
        doc.text(l, marginX, y);
        y += lineH;
      });
    };

    if (d.grouped) {
      d.sections.forEach((sec, idx) => {
        if (idx > 0) y += 8;
        ensureSpace(lineH + 6);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(11);
        writeLine(`${sec.grupo.toUpperCase()} (${sec.items.length})`);
        doc.setFont('helvetica', 'normal');
        sec.items.forEach(it => writeLine(`${it.n}. ${it.nombre}`));
      });
    } else {
      d.rows.forEach(r => writeLine(r.grupo !== '' ? `${r.n}. ${r.nombre} — ${r.grupo}` : `${r.n}. ${r.nombre}`));
    }

    doc.save(this._buildFileName('pdf'));
  },

  async _copyToClipboard(text) {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function' && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return;
    }
    // Respaldo para contextos sin Clipboard API (p. ej. HTTP o WebViews antiguos)
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.top = '-1000px';
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    if (!ok) throw new Error('No se pudo copiar al portapapeles');
  },

  _toggleExportMenu(show) {
    const menu = document.getElementById('kpiExportMenu');
    const btn  = document.getElementById('kpiExportBtn');
    if (!menu || !btn) return;
    menu.classList.toggle('d-none', !show);
    btn.setAttribute('aria-expanded', String(!!show));
    btn.classList.toggle('kpi-export-btn-open', !!show);
  },

  /* Mensaje temporal en la etiqueta del botón (feedback sin librerías) */
  _flashExportLabel(text, isError = false) {
    const label = document.getElementById('kpiExportLabel');
    const btn   = document.getElementById('kpiExportBtn');
    if (!label || !btn) return;
    clearTimeout(this._flashTimeout);
    label.textContent = text;
    btn.classList.toggle('kpi-export-btn-error', isError);
    btn.classList.toggle('kpi-export-btn-ok', !isError);
    this._flashTimeout = setTimeout(() => {
      label.textContent = 'EXPORTAR LISTA';
      btn.classList.remove('kpi-export-btn-error', 'kpi-export-btn-ok');
    }, 2200);
  },

  async _handleExport(kind) {
    if (this._getVisiblePersonas().length === 0) {
      this._flashExportLabel('SIN DATOS PARA EXPORTAR', true);
      return;
    }

    try {
      if (kind === 'txt') {
        const blob = new Blob(['\ufeff' + this._buildPlainText()], { type: 'text/plain;charset=utf-8' });
        this._downloadBlob(blob, this._buildFileName('txt'));
        this._flashExportLabel('¡TXT DESCARGADO!');
      } else if (kind === 'word') {
        const blob = new Blob(['\ufeff' + this._buildWordHtml()], { type: 'application/msword;charset=utf-8' });
        this._downloadBlob(blob, this._buildFileName('doc'));
        this._flashExportLabel('¡WORD DESCARGADO!');
      } else if (kind === 'copy') {
        await this._copyToClipboard(this._buildPlainText());
        this._flashExportLabel('¡COPIADO!');
      } else if (kind === 'pdf') {
        this._flashExportLabel('GENERANDO PDF...');
        await this._exportPdf();
        this._flashExportLabel('¡PDF DESCARGADO!');
      }
    } catch (err) {
      console.error('[ModalEngine] Error al exportar:', err);
      this._flashExportLabel('ERROR AL EXPORTAR', true);
    }
  },

  /* Cierra el modal con una pequeña transición de salida */
  close() {
    const backdrop = document.getElementById('kpiDetailModal');
    if (!backdrop) return;

    this._toggleExportMenu(false);
    backdrop.classList.remove('kpi-modal-visible');
    document.body.classList.remove('kpi-modal-open');
    clearTimeout(this._closeTimeout);
    this._closeTimeout = setTimeout(() => backdrop.classList.add('d-none'), 250);
  },
};
window.ModalEngine = ModalEngine;


/* ────────────────────────────────────────────────────────────
   LIST EXPORT ENGINE — botón "EXPORTAR LISTA" para las pestañas
   "Excluidos" y "Monitor de Ausencias"
   ────────────────────────────────────────────────────────────
   Inyecta el mismo botón desplegable del modal de KPIs (PDF, Word
   .doc, TXT y Copiar) en la barra de búsqueda de cada pestaña.
   Exporta EXACTAMENTE lo que se ve en pantalla en ese momento:
     - Excluidos: grupo seleccionado + búsqueda activa de la tabla.
     - Ausencias: filtro de nivel + búsqueda activa del monitor.
   Reutiliza los helpers de ModalEngine (descarga de Blob, jsPDF
   dinámico, portapapeles y escape de HTML).
──────────────────────────────────────────────────────────── */
const ListExportEngine = {

  _flashTimeouts: new WeakMap(),

  /* Configuración por sección: de dónde salen los registros visibles,
     qué columnas llevan Word/PDF y cómo se escribe cada línea apilada. */
  SECTIONS: {
    excluidos: {
      barSelector: '#tabExcluidos .table-search-bar',
      searchId: 'searchExcluidos',
      title: 'Excluidos',
      columns: [
        { label: 'Nombre',     get: r => r.nombre || '' },
        { label: 'Grupo',      get: r => r.grupo || '' },
        { label: 'Teléfono',   get: r => r.telefono || '' },
        { label: 'Célula',     get: r => r.celula || '' },
        { label: 'Servicio',   get: r => r.servicio || '' },
        { label: 'Estado',     get: r => r.estado || '' },
        { label: 'Fecha',      get: r => r.fecha || '' },
      ],
      line: r => (r.grupo ? `${r.nombre} — ${r.grupo}` : `${r.nombre}`),
      getRecords() {
        const all = (typeof TableEngine !== 'undefined' && TableEngine._excluidosRecords) || [];
        const rows = document.querySelectorAll('#tableExcluidos tbody tr');
        const visibles = [];
        rows.forEach((row, i) => {
          if (row.style.display !== 'none' && all[i]) visibles.push(all[i]);
        });
        return visibles;
      },
    },
    ausencias: {
      barSelector: '#tabAusencias .table-search-bar',
      searchId: 'searchAusencias',
      title: 'Monitor de Ausencias',
      columns: [
        { label: 'Nombre',            get: r => r.nombre || '' },
        { label: 'Grupo',             get: r => r.grupo || '' },
        { label: 'Teléfono',          get: r => r.telefono || '' },
        { label: 'Última Falta',      get: r => r.fechaFormatted || '' },
        { label: 'Tiempo sin asistir',get: r => (r.timeFmt ? r.timeFmt.main + (r.timeFmt.detail ? ` (${r.timeFmt.detail})` : '') : '') },
        { label: 'Célula',            get: r => r.celula || '' },
        { label: 'Servicio',          get: r => r.servicio || '' },
        { label: 'Estado',            get: r => r.estado || '' },
        { label: 'Nivel de Alerta',   get: r => (r.levelObj ? r.levelObj.label : '') },
      ],
      line: r => {
        const base = r.grupo ? `${r.nombre} — ${r.grupo}` : `${r.nombre}`;
        const tiempo = r.timeFmt ? r.timeFmt.main : '';
        const nivel = r.levelObj ? r.levelObj.label : '';
        return `${base} — Sin asistir: ${tiempo} (${nivel})`;
      },
      getRecords() {
        return (typeof AbsenceEngine !== 'undefined' && Array.isArray(AbsenceEngine._visibleData))
          ? AbsenceEngine._visibleData
          : [];
      },
    },
  },

  /* Subtítulo con los filtros activos (para encabezar el archivo) */
  _getFilterNote(key) {
    const cfg = this.SECTIONS[key];
    const notes = [];
    const q = (document.getElementById(cfg.searchId)?.value || '').trim();
    if (q !== '') notes.push(`Búsqueda: "${q}"`);
    if (key === 'ausencias') {
      const activeBtn = document.querySelector('.aus-filter-btn.active');
      const lvl = activeBtn ? (activeBtn.dataset.level || '') : '';
      if (lvl !== '') notes.push(`Nivel: ${activeBtn.textContent.trim()}`);
    }
    return notes.join(' · ');
  },

  _buildPlainText(key) {
    const cfg = this.SECTIONS[key];
    const records = cfg.getRecords();
    const note = this._getFilterNote(key);
    const lines = [cfg.title.toUpperCase(), `Total: ${records.length} persona${records.length === 1 ? '' : 's'}`];
    if (note) lines.push(`Filtros: ${note}`);
    lines.push('');
    records.forEach((r, i) => lines.push(`${i + 1}. ${cfg.line(r)}`));
    return lines.join('\n');
  },

  _buildWordHtml(key) {
    const cfg = this.SECTIONS[key];
    const records = cfg.getRecords();
    const note = this._getFilterNote(key);
    const esc = (v) => ModalEngine._escapeHtml(v);
    const cell = 'border:1px solid #999;padding:4px 8px;';
    const head = `<tr><th style="${cell}background:#e8eefc;width:40px;">#</th>` +
      cfg.columns.map(c => `<th style="${cell}background:#e8eefc;">${esc(c.label)}</th>`).join('') + '</tr>';
    const body = records.map((r, i) =>
      `<tr><td style="${cell}">${i + 1}</td>` +
      cfg.columns.map(c => `<td style="${cell}">${esc(c.get(r))}</td>`).join('') + '</tr>'
    ).join('');

    return `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>${esc(cfg.title)}</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><![endif]-->
<style>@page { size: landscape; }</style></head>
<body style="font-family:Calibri,Arial,sans-serif;">
<h2 style="font-family:Calibri,Arial,sans-serif;">${esc(cfg.title)}</h2>
<p style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#555;">Total: ${records.length} persona${records.length === 1 ? '' : 's'}</p>
${note ? `<p style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#555;">Filtros: ${esc(note)}</p>` : ''}
<table style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:10pt;">${head}${body}</table>
</body></html>`;
  },

  _buildFileName(key, ext) {
    const base = key === 'ausencias' ? 'monitor_de_ausencias' : 'excluidos';
    const hoy = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${base}_${hoy.getFullYear()}-${pad(hoy.getMonth() + 1)}-${pad(hoy.getDate())}.${ext}`;
  },

  async _exportPdf(key) {
    const JsPDF = await ModalEngine._loadJsPdf();
    const cfg = this.SECTIONS[key];
    const records = cfg.getRecords();
    const note = this._getFilterNote(key);
    const doc = new JsPDF({ unit: 'pt', format: 'letter' });

    const marginX = 48, marginTop = 56, marginBottom = 48, lineH = 15;
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const maxW = pageW - marginX * 2;
    let y = marginTop;

    const ensureSpace = (needed) => {
      if (y + needed > pageH - marginBottom) { doc.addPage(); y = marginTop; }
    };

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(cfg.title, marginX, y);
    y += 22;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Total: ${records.length} persona${records.length === 1 ? '' : 's'}`, marginX, y);
    y += 14;
    if (note) {
      doc.splitTextToSize(`Filtros: ${note}`, maxW).forEach(l => { doc.text(l, marginX, y); y += 14; });
    }
    y += 8;
    doc.setTextColor(0);
    doc.setFontSize(11);

    records.forEach((r, i) => {
      doc.splitTextToSize(`${i + 1}. ${cfg.line(r)}`, maxW).forEach(l => {
        ensureSpace(lineH);
        doc.text(l, marginX, y);
        y += lineH;
      });
    });

    doc.save(this._buildFileName(key, 'pdf'));
  },

  /* Posiciona el menú (fixed) bajo el botón sin salirse de la pantalla */
  _toggleMenu(wrapper, show) {
    const menu = wrapper.querySelector('.kpi-export-menu');
    const btn  = wrapper.querySelector('.kpi-export-btn');
    if (!menu || !btn) return;

    if (show) {
      menu.classList.remove('d-none');
      const r = btn.getBoundingClientRect();
      const mw = menu.offsetWidth || 250;
      const left = Math.max(8, Math.min(r.left, window.innerWidth - mw - 8));
      menu.style.position = 'fixed';
      menu.style.top = `${r.bottom + 6}px`;
      menu.style.left = `${left}px`;
    } else {
      menu.classList.add('d-none');
    }
    btn.setAttribute('aria-expanded', String(!!show));
    btn.classList.toggle('kpi-export-btn-open', !!show);
  },

  _closeAllMenus() {
    document.querySelectorAll('.list-export-dropdown').forEach(w => this._toggleMenu(w, false));
  },

  _flash(wrapper, text, isError = false) {
    const label = wrapper.querySelector('.kpi-export-label');
    const btn   = wrapper.querySelector('.kpi-export-btn');
    if (!label || !btn) return;
    clearTimeout(this._flashTimeouts.get(wrapper));
    label.textContent = text;
    btn.classList.toggle('kpi-export-btn-error', isError);
    btn.classList.toggle('kpi-export-btn-ok', !isError);
    this._flashTimeouts.set(wrapper, setTimeout(() => {
      label.textContent = 'EXPORTAR LISTA';
      btn.classList.remove('kpi-export-btn-error', 'kpi-export-btn-ok');
    }, 2200));
  },

  async _handle(key, kind, wrapper) {
    const cfg = this.SECTIONS[key];
    if (cfg.getRecords().length === 0) {
      this._flash(wrapper, 'SIN DATOS PARA EXPORTAR', true);
      return;
    }

    try {
      if (kind === 'txt') {
        const blob = new Blob(['\ufeff' + this._buildPlainText(key)], { type: 'text/plain;charset=utf-8' });
        ModalEngine._downloadBlob(blob, this._buildFileName(key, 'txt'));
        this._flash(wrapper, '¡TXT DESCARGADO!');
      } else if (kind === 'word') {
        const blob = new Blob(['\ufeff' + this._buildWordHtml(key)], { type: 'application/msword;charset=utf-8' });
        ModalEngine._downloadBlob(blob, this._buildFileName(key, 'doc'));
        this._flash(wrapper, '¡WORD DESCARGADO!');
      } else if (kind === 'copy') {
        await ModalEngine._copyToClipboard(this._buildPlainText(key));
        this._flash(wrapper, '¡COPIADO!');
      } else if (kind === 'pdf') {
        this._flash(wrapper, 'GENERANDO PDF...');
        await this._exportPdf(key);
        this._flash(wrapper, '¡PDF DESCARGADO!');
      }
    } catch (err) {
      console.error('[ListExportEngine] Error al exportar:', err);
      this._flash(wrapper, 'ERROR AL EXPORTAR', true);
    }
  },

  /* Crea el botón desplegable y lo inserta antes del contador de registros */
  _mount(key) {
    const cfg = this.SECTIONS[key];
    const bar = document.querySelector(cfg.barSelector);
    if (!bar || bar.querySelector('.list-export-dropdown')) return;

    const wrapper = document.createElement('div');
    wrapper.className = 'kpi-export-dropdown list-export-dropdown';
    wrapper.innerHTML = `
      <button type="button" class="kpi-export-btn" aria-haspopup="true" aria-expanded="false" title="Exportar la lista visible">
        <i class="bi bi-download"></i>
        <span class="kpi-export-label">EXPORTAR LISTA</span>
        <i class="bi bi-chevron-down kpi-export-caret"></i>
      </button>
      <ul class="kpi-export-menu d-none" role="menu">
        <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="pdf"><i class="bi bi-file-earmark-pdf"></i> Exportar a PDF</button></li>
        <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="word"><i class="bi bi-file-earmark-word"></i> Exportar a Word (.doc)</button></li>
        <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="txt"><i class="bi bi-file-earmark-text"></i> Exportar a Archivo de Texto (.txt)</button></li>
        <li role="none"><button type="button" role="menuitem" class="kpi-export-item" data-export="copy"><i class="bi bi-clipboard"></i> Copiar al Portapapeles</button></li>
      </ul>`;

    const count = bar.querySelector('.table-count');
    if (count) bar.insertBefore(wrapper, count); else bar.appendChild(wrapper);

    const menu = wrapper.querySelector('.kpi-export-menu');
    wrapper.querySelector('.kpi-export-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      const abrir = menu.classList.contains('d-none');
      this._closeAllMenus();
      this._toggleMenu(wrapper, abrir);
    });
    menu.addEventListener('click', (e) => {
      const item = e.target.closest('[data-export]');
      if (!item) return;
      e.stopPropagation();
      this._toggleMenu(wrapper, false);
      this._handle(key, item.getAttribute('data-export'), wrapper);
    });
  },

  init() {
    if (this._initialized) return;
    this._initialized = true;

    Object.keys(this.SECTIONS).forEach(key => this._mount(key));

    // Cierres globales del menú (clic fuera, Escape, scroll, resize)
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.list-export-dropdown')) this._closeAllMenus();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') this._closeAllMenus();
    });
    window.addEventListener('resize', () => this._closeAllMenus());
    window.addEventListener('scroll', () => this._closeAllMenus(), true);
  },
};
window.ListExportEngine = ListExportEngine;

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => ListExportEngine.init());
} else {
  ListExportEngine.init();
}




/* ────────────────────────────────────────────────────────────
   4. CHART ENGINE — crea y actualiza todos los gráficos
   Para añadir nuevos gráficos: agregar instancia aquí e
   inicializarla en init().
──────────────────────────────────────────────────────────── */
const ChartEngine = {
  instances: {},  // Almacena instancias de Chart.js para actualizarlas sin recrear

  /* Paleta de colores consistente */
  palette: [
    '#f0b429','#22c55e','#38bdf8','#a78bfa',
    '#fb923c','#f472b6','#34d399','#60a5fa',
    '#fbbf24','#4ade80','#818cf8','#2dd4bf',
  ],

  /* Opciones base compartidas por todos los gráficos */
  baseOptions() {
    return {
      responsive: true,
      maintainAspectRatio: false,
      /* Interactividad explícita: clic en la leyenda (mostrar/ocultar
         series) y hover con tooltip SIEMPRE activos, sin depender de
         que ningún override global de Chart.defaults (p. ej. la
         animación de entrada de InteractiveLife.js) los deje intactos
         por accidente. No se define plugins.legend.onClick a propósito:
         cada tipo de gráfico (dona vs. barras) tiene su propio
         comportamiento de clic por defecto en Chart.js (alternar arco
         vs. alternar dataset) y fijar uno genérico aquí rompería el
         otro — dejamos que Chart.js resuelva el correcto según el tipo.*/
      events: ['mousemove', 'mouseout', 'click', 'touchstart', 'touchmove'],
      interaction: {
        mode: 'nearest',
        intersect: true,
      },
      /* Animación corta y consistente: evita que, al cambiar de filtro
         de grupo, el gráfico recién recreado quede "vacío" varios
         cientos de ms mientras anima desde cero (percibido como que
         "desaparece"). No se desactiva del todo para conservar la
         transición suave del diseño, solo se acota su duración. */
      animation: {
        duration: 350,
        easing: 'easeOutQuart',
      },
      /* Config explícita (y genérica, sin callbacks dependientes de
         escalas — eso fue lo que rompía los gráficos antes) de qué
         propiedades numéricas anima Chart.js en cada redibujado. Sin
         esto, cada tipo de gráfico usa el set de propiedades animadas
         que trae por defecto SU PROPIO controlador interno (el de
         dona no es igual al de barras), lo que hacía que algunas
         gráficas aparecieran "de golpe" mientras otras sí mostraban
         el efecto de relleno. Al declararlo aquí, TODAS (dona,
         embudo, barras, apiladas, ranking) animan el mismo conjunto
         de propiedades con la misma duración/easing. */
      animations: {
        numbers: {
          type: 'number',
          properties: ['x', 'y', 'width', 'height', 'circumference', 'endAngle', 'innerRadius', 'outerRadius'],
          duration: 350,
          easing: 'easeOutQuart',
        },
      },
      plugins: {
        legend: {
          labels: {
            color: Chart.defaults.color,
            font: { family: 'Outfit', size: 12 },
            boxWidth: 12,
            padding: 16,
          },
        },
        tooltip: {
          backgroundColor: '#1c2333',
          borderColor: 'rgba(255,255,255,.07)',
          borderWidth: 1,
          titleColor: '#e2e8f0',
          bodyColor: '#94a3b8',  // dark default; ThemeEngine overrides Chart.defaults
          padding: 10,
          cornerRadius: 8,
          titleFont: { family: 'Outfit', weight: '600' },
          bodyFont:  { family: 'Outfit' },
        },
      },
    };
  },

  /* Destruye una instancia si existe */
  destroy(id) {
    if (this.instances[id]) {
      this.instances[id].destroy();
      delete this.instances[id];
    }
  },

  /* Actualiza SOLO los colores de tema (ticks, grid, leyenda, tooltip)
     de los gráficos YA CREADOS, sin destruirlos ni recrearlos.

     Por qué: ThemeEngine.apply() antes llamaba a renderAll(kpis) en
     cada cambio de tema, lo que hace destroy() + new Chart() de las
     5 gráficas de golpe — cada una arranca su animación de entrada
     desde cero, y durante ese instante el canvas queda en blanco.
     Al ser un clic directo del usuario sobre el botón de tema, ese
     parpadeo se nota mucho más que en un cambio de filtro. Como los
     DATOS no cambian al cambiar de tema (solo los colores), no hace
     falta recrear nada: basta con mutar las opciones de color de cada
     instancia existente y pedirle un update sin animación
     (chart.update('none')) para que el redibujado sea instantáneo y
     nunca pase por un estado vacío.

     No toca datasets, labels ni ningún dato — solo propiedades de
     color dentro de options. Si todavía no hay instancias (dashboard
     vacío), no hace nada: el próximo renderAll() ya usará los nuevos
     Chart.defaults. */
  updateTheme(theme) {
    const isLight = theme === 'light';
    const textColor     = isLight ? '#0f172a'         : '#94a3b8';
    const gridColor      = isLight ? 'rgba(0,0,0,.08)' : 'rgba(255,255,255,.04)';
    const tooltipBg      = isLight ? '#ffffff'          : '#1c2333';
    const tooltipBorder  = isLight ? 'rgba(0,0,0,.12)'  : 'rgba(255,255,255,.07)';
    const tooltipTitle   = isLight ? '#0f172a'          : '#e2e8f0';
    const tooltipBody    = isLight ? '#1e293b'          : '#94a3b8';

    Object.values(this.instances).forEach(chart => {
      if (!chart || !chart.options) return;

      if (chart.options.scales) {
        Object.values(chart.options.scales).forEach(scale => {
          if (scale.ticks) scale.ticks.color = textColor;
          if (scale.grid)  scale.grid.color  = gridColor;
        });
      }

      if (chart.options.plugins?.legend?.labels) {
        chart.options.plugins.legend.labels.color = textColor;
      }
      if (chart.options.plugins?.tooltip) {
        chart.options.plugins.tooltip.backgroundColor = tooltipBg;
        chart.options.plugins.tooltip.borderColor     = tooltipBorder;
        chart.options.plugins.tooltip.titleColor      = tooltipTitle;
        chart.options.plugins.tooltip.bodyColor       = tooltipBody;
      }

      chart.update('none'); // 'none' = sin animación: cambio de color instantáneo, sin parpadeo
    });
  },

  /* ── Donut: Asistencia vs Inasistencia general ── */
  renderDonut(kpis) {
    this.destroy('donut');
    const ctx = document.getElementById('chartDonut');
    if (!ctx) return;
    this.instances.donut = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['Asistió Célula', 'Asistió Servicio', 'Ambos', 'Ausente en Ambos', 'Nuevos'],
        datasets: [{
          data: [
            kpis.celulasSI,
            kpis.servicioSI,
            kpis.ambosSI,
            kpis.ambosNO,
            kpis.totalNuevos,
          ],
          backgroundColor: ['#22c55e','#38bdf8','#a78bfa','#ef4444','#f0b429'],
          borderColor: '#161b24',
          borderWidth: 3,
          hoverOffset: 8,
        }],
      },
      options: {
        ...this.baseOptions(),
        cutout: '65%',
        plugins: {
          ...this.baseOptions().plugins,
          legend: { position: 'bottom', ...this.baseOptions().plugins.legend },
        },
      },
    });
  },

  /* ── Embudo: Nuevos ── */
  renderFunnel(kpis) {
    this.destroy('funnel');
    const ctx = document.getElementById('chartFunnel');
    if (!ctx) return;

    // Datos embudo: Total registrados → Nuevos en célula → Nuevos en servicio
    const steps = [
      { label: 'Total Registrados', value: kpis.total },
      { label: 'Nuevos en Célula',  value: kpis.nuevosCelula },
      { label: 'Nuevos en Servicio', value: kpis.nuevosServicio },
    ];

    this.instances.funnel = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: steps.map(s => s.label),
        datasets: [{
          label: 'Personas',
          data: steps.map(s => s.value),
          backgroundColor: ['#38bdf8cc','#22c55ecc','#f0b429cc'],
          borderRadius: 6,
          borderSkipped: false,
        }],
      },
      options: {
        ...this.baseOptions(),
        indexAxis: 'y',
        plugins: {
          ...this.baseOptions().plugins,
          legend: { display: false },
        },
        scales: {
          x: {
            grid: { color: 'rgba(255,255,255,.04)' },
            ticks: { color: Chart.defaults.color, font: { family: 'Outfit', size: 11 } },
          },
          y: {
            grid: { display: false },
            ticks: { color: Chart.defaults.color, font: { family: 'Outfit', size: 11 } },
          },
        },
      },
    });
  },

  /* ── Barras: Asistencia por grupo ministerial ── */
  renderBarGroup(kpis) {
    this.destroy('barGroup');
    const ctx = document.getElementById('chartBarGroup');
    if (!ctx) return;

    const groups = Object.keys(kpis.byGroup);
    const siData = groups.map(g => kpis.byGroup[g].si + (kpis.byGroup[g].nuevosCel || 0) + (kpis.byGroup[g].nuevosSrv || 0));
    const noData = groups.map(g => kpis.byGroup[g].no);

    // Nombres de grupo acortados para el eje
    const shortLabels = groups.map(g => g.replace(/Ministr[ao]s?\s*/i, '').substring(0, 22));

    this.instances.barGroup = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: shortLabels,
        datasets: [
          {
            label: 'Asistió Célula',
            data: siData,
            backgroundColor: '#22c55ecc',
            borderRadius: 4,
          },
          {
            label: 'No Asistió',
            data: noData,
            backgroundColor: '#ef4444cc',
            borderRadius: 4,
          },
        ],
      },
      options: {
        ...this.baseOptions(),
        plugins: {
          ...this.baseOptions().plugins,
          legend: { position: 'top', ...this.baseOptions().plugins.legend },
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { color: Chart.defaults.color, font: { family: 'Outfit', size: 11 }, maxRotation: 30 },
          },
          y: {
            grid: { color: 'rgba(255,255,255,.04)' },
            ticks: { color: Chart.defaults.color, font: { family: 'Outfit', size: 11 } },
          },
        },
      },
    });
  },

  /* ── Barras apiladas: SI / NO / NUEVO por grupo ── */
  renderStacked(kpis) {
    this.destroy('stacked');
    const ctx = document.getElementById('chartStacked');
    if (!ctx) return;

    const groups = Object.keys(kpis.byGroup);
    const shortLabels = groups.map(g => g.replace(/Ministr[ao]s?\s*/i,'').substring(0,22));

    this.instances.stacked = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: shortLabels,
        datasets: [
          {
            label: 'SI (Asistió)',
            data: groups.map(g => kpis.byGroup[g].si),
            backgroundColor: '#22c55e99',
            stack: 'celula',
            borderRadius: 3,
          },
          {
            label: 'NO (Ausente)',
            data: groups.map(g => kpis.byGroup[g].no),
            backgroundColor: '#ef444499',
            stack: 'celula',
            borderRadius: 3,
          },
          {
            label: 'NUEVO',
            data: groups.map(g => (kpis.byGroup[g].nuevosCel || 0) + (kpis.byGroup[g].nuevosSrv || 0)),
            backgroundColor: '#f0b42999',
            stack: 'celula',
            borderRadius: 3,
          },
        ],
      },
      options: {
        ...this.baseOptions(),
        plugins: {
          ...this.baseOptions().plugins,
          legend: { position: 'top', ...this.baseOptions().plugins.legend },
        },
        scales: {
          x: {
            stacked: true,
            grid: { display: false },
            ticks: { color: Chart.defaults.color, font: { family: 'Outfit', size: 11 }, maxRotation: 30 },
          },
          y: {
            stacked: true,
            grid: { color: 'rgba(255,255,255,.04)' },
            ticks: { color: Chart.defaults.color, font: { family: 'Outfit', size: 11 } },
          },
        },
      },
    });
  },

  /* ── Ranking integral de grupos — barras horizontales agrupadas ──
     Un solo gráfico que reemplaza los antiguos "Top Asistencia" y
     "Mayor Ausentismo": muestra TODOS los grupos ordenados de mayor
     a menor asistencia, con dos barras horizontales por grupo
     (asistencia % en verde, ausentismo % en rojo), para ver el
     panorama comparativo completo de un solo vistazo. */
  renderRankings(kpis) {
    const allGroups = Object.entries(kpis.byGroup).map(([name, data]) => {
      // Asistentes reales a célula = SI + nuevos en célula + nuevos en servicio
      const asistentes = data.si + (data.nuevosCel || 0) + (data.nuevosSrv || 0);
      const ausentes   = data.no || 0;
      const total      = data.total || 0;
      const pctAsist   = total > 0 ? Math.round((asistentes / total) * 100) : 0;
      const pctAus     = total > 0 ? Math.round((ausentes   / total) * 100) : 0;

      const short = name
        .replace(/Ministr[ao]s?\s*/i, '')
        .replace(/Lider[a]?s?\s*/i, '')
        .trim()
        .substring(0, 26) || name.substring(0, 26);

      return { name, short, asistentes, ausentes, total, pctAsist, pctAus };
    });

    // Ordenado de mayor a menor asistencia — un único ranking integral
    const ranked = [...allGroups].sort((a, b) => b.pctAsist - a.pctAsist);

    this._renderCombinedRankChart('chartRankCombined', ranked);
  },

  /* Renderiza el ranking integral como barras horizontales agrupadas
     (asistencia vs ausentismo, lado a lado, por grupo). */
  _renderCombinedRankChart(canvasId, items) {
    this.destroy(canvasId);
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    const colorAsist   = '#22c55e';
    const bgAsist       = 'rgba(34,197,94,.15)';
    const colorAus      = '#ef4444';
    const bgAus          = 'rgba(239,68,68,.15)';

    // Altura dinámica: más grupos → canvas más alto (evita apretujar barras)
    const canvasEl = document.getElementById(canvasId);
    if (canvasEl) {
      canvasEl.style.height = Math.max(320, items.length * 38) + 'px';
      // Fuerza un reflow ANTES de crear el chart: sin esto, Chart.js mide
      // el canvas con el alto anterior (el del grupo previamente
      // seleccionado) y el ResizeObserver interno lo corrige recién en
      // un frame posterior, lo que se percibe como que el gráfico
      // "desaparece" un instante al pasar a "Todos los grupos" (o
      // viceversa) por el salto de tamaño.
      void canvasEl.offsetHeight;
    }

    this.instances[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: items.map(i => i.short),
        datasets: [
          {
            label: 'Asistencia %',
            data: items.map(i => i.pctAsist),
            backgroundColor: bgAsist,
            borderColor: colorAsist,
            borderWidth: 2,
            borderRadius: 5,
            borderSkipped: false,
            hoverBackgroundColor: colorAsist + '88',
          },
          {
            label: 'Ausentismo %',
            data: items.map(i => i.pctAus),
            backgroundColor: bgAus,
            borderColor: colorAus,
            borderWidth: 2,
            borderRadius: 5,
            borderSkipped: false,
            hoverBackgroundColor: colorAus + '88',
          },
        ],
      },
      options: {
        ...this.baseOptions(),
        indexAxis: 'y',
        plugins: {
          ...this.baseOptions().plugins,
          legend: {
            display: true,
            position: 'top',
            labels: { color: Chart.defaults.color, font: { family: 'Outfit', size: 11 }, boxWidth: 14 },
          },
          tooltip: {
            ...this.baseOptions().plugins.tooltip,
            callbacks: {
              label: (ctx) => {
                const item = items[ctx.dataIndex];
                return ctx.dataset.label === 'Asistencia %'
                  ? ` ${item.pctAsist}% asistencia (${item.asistentes}/${item.total})`
                  : ` ${item.pctAus}% ausencia (${item.ausentes}/${item.total})`;
              },
            },
          },
        },
        scales: {
          x: {
            beginAtZero: true,
            max: 100,
            grid: { color: 'rgba(255,255,255,.04)' },
            ticks: {
              color: Chart.defaults.color,
              font: { family: 'Outfit', size: 10 },
              callback: v => v + '%',
            },
          },
          y: {
            grid: { display: false },
            ticks: {
              color: Chart.defaults.color,
              font: { family: 'Outfit', size: 10 },
            },
          },
        },
      },
    });
  },

  /* Punto de entrada: renderiza/actualiza todos los gráficos */
  renderAll(kpis) {
    this.renderDonut(kpis);
    this.renderFunnel(kpis);
    this.renderBarGroup(kpis);
    this.renderStacked(kpis);
    this.renderRankings(kpis);  // uses chartRankCombined canvas
  },
};


/* ────────────────────────────────────────────────────────────
   5. TABLE ENGINE — renderiza tablas con búsqueda
──────────────────────────────────────────────────────────── */
const TableEngine = {

  /* Devuelve badge HTML según valor de asistencia */
  badge(val) {
    const v = (val || '').toUpperCase();
    if (v === 'SI')    return `<span class="badge-si">SI</span>`;
    if (v === 'NO')    return `<span class="badge-no">NO</span>`;
    if (v === 'NUEVO') return `<span class="badge-nuevo">NUEVO</span>`;
    return `<span style="color:var(--text-dim)">${val || '—'}</span>`;
  },

  /* Limpia un número telefónico y arma el enlace de WhatsApp.
     Conserva el '+' inicial si existe; descarta cualquier otro
     carácter no numérico (espacios, guiones, paréntesis, etc.).
     Si no hay número registrado, devuelve un guion silenciado. */
  waLink(telefono) {
    const raw = (telefono || '').toString().trim();
    if (!raw) return '<span style="color:var(--text-dim)">—</span>';

    const tienePlus = raw.startsWith('+');
    const soloDigitos = raw.replace(/[^\d]/g, '');
    if (!soloDigitos) return '<span style="color:var(--text-dim)">—</span>';

    const numeroLimpio = (tienePlus ? '+' : '') + soloDigitos;
    return `<a href="https://wa.me/${numeroLimpio}" target="_blank" rel="noopener noreferrer" class="tel-whatsapp-link" title="Abrir chat de WhatsApp">
      <i class="bi bi-whatsapp me-1"></i>${raw}
    </a>`;
  },

  /* Render tabla de personas */
  renderPersonas(records) {
    const tbody = document.querySelector('#tablePersonas tbody');
    if (!tbody) return;

    tbody.innerHTML = records.map((r, i) => `
      <tr>
        <td>${i+1}</td>
        <td>${r.nombre}</td>
        <td style="color:var(--text-dim);font-size:11px">${r.grupo.replace(/Ministr[ao]s?\s*/i,'').substring(0,30)}</td>
        <td>${this.waLink(r.telefono)}</td>
        <td>${this.badge(r.celula)}</td>
        <td>${this.badge(r.servicio)}</td>
        <td style="color:var(--text-dim);font-size:11px">${r.estado || '—'}</td>
        <td>${r.esNuevo ? '<span class="badge-nuevo-tag">NUEVO</span>' : ''}</td>
      </tr>
    `).join('');

    document.getElementById('countPersonas').textContent = `${records.length} registros`;
  },

  /* Render tabla de excluidos */
  renderExcluidos(records) {
    const tbody = document.querySelector('#tableExcluidos tbody');
    if (!tbody) return;

    this._excluidosRecords = records;  // usado por ListExportEngine

    tbody.innerHTML = records.map((r, i) => `
      <tr>
        <td>${i+1}</td>
        <td>${r.nombre}</td>
        <td style="color:var(--text-dim);font-size:11px">${r.grupo.substring(0,30)}</td>
        <td>${this.waLink(r.telefono)}</td>
        <td>${this.badge(r.celula)}</td>
        <td>${this.badge(r.servicio)}</td>
        <td style="color:var(--text-dim);font-size:11px">${r.estado || '—'}</td>
        <td style="color:var(--text-dim);font-size:11px">${r.fecha || '—'}</td>
      </tr>
    `).join('');

    document.getElementById('countExcluidos').textContent = `${records.length} registros`;
  },

  /* Render tabla de nuevos — diferencia célula vs servicio correctamente */
  renderNuevos(records) {
    const tbody = document.querySelector('#tableNuevos tbody');
    if (!tbody) return;

    // Nuevos en célula: Estado (col E) = 'NUEVO'
    // Nuevos en servicio: Célula (col C) = 'NUEVO'
    const nuevos = records.filter(r => r.esNuevo);

    tbody.innerHTML = nuevos.map((r, i) => {
      // Etiqueta de tipo
      let tipoTag = '';
      if (r.esNuevoCelula && r.esNuevoServicio) {
        tipoTag = `<span class="badge-nuevo-cel">Célula</span> <span class="badge-nuevo-srv">Servicio</span>`;
      } else if (r.esNuevoCelula) {
        tipoTag = `<span class="badge-nuevo-cel">Célula</span>`;
      } else if (r.esNuevoServicio) {
        tipoTag = `<span class="badge-nuevo-srv">Servicio</span>`;
      }

      // Para nuevo en célula: su Célula puede ser SI/NO, su Estado es NUEVO
      // Para nuevo en servicio: su campo Célula dice NUEVO (llegó por primera vez)
      const celulaDisplay = r.esNuevoServicio ? '<span class="badge-nuevo">NUEVO</span>' : this.badge(r.celula);
      const servicioDisplay = this.badge(r.servicio);

      return `<tr>
        <td>${i+1}</td>
        <td>${r.nombre}</td>
        <td style="color:var(--text-dim);font-size:11px">${r.grupo.replace(/Ministr[ao]s?\s*/i,'').substring(0,28)}</td>
        <td>${this.waLink(r.telefono)}</td>
        <td>${celulaDisplay}</td>
        <td>${servicioDisplay}</td>
        <td>${tipoTag}</td>
      </tr>`;
    }).join('');

    document.getElementById('countNuevos').textContent = `${nuevos.length} nuevos (${records.filter(r=>r.esNuevoCelula).length} célula · ${records.filter(r=>r.esNuevoServicio).length} servicio)`;
  },

  /* Render tabla histórico (ANTIGUO EX) */
  renderHistorico(records) {
    const tbody = document.querySelector('#tableHistorico tbody');
    if (!tbody) return;

    tbody.innerHTML = records.map((r, i) => `
      <tr>
        <td>${i+1}</td>
        <td>${r.nombre}</td>
        <td>${this.badge(r.estado)}</td>
        <td style="color:var(--text-dim);font-size:11px">${r.fecha || '—'}</td>
      </tr>
    `).join('');

    document.getElementById('countHistorico').textContent = `${records.length} registros`;
  },

  /* Filtra una tabla por texto */
  filterTable(tableId, searchText) {
    const rows = document.querySelectorAll(`#${tableId} tbody tr`);
    const q = searchText.toLowerCase();
    let visible = 0;
    rows.forEach(row => {
      const match = row.textContent.toLowerCase().includes(q);
      row.style.display = match ? '' : 'none';
      if (match) visible++;
    });
    return visible;
  },

  /* Renderiza todas las tablas */
  renderAll(filteredMain) {
    this.renderPersonas(filteredMain);
    /* FIX: antes se pasaba DataStore.rawExcluidos "crudo", ignorando
       el <select> de grupo para usuarios con acceso 'ALL'. Ahora
       siempre se respeta el valor actual del dropdown (ver
       DataStore.filterExcluidosByGroup), sin importar el rol. */
    this.renderExcluidos(DataStore.filterExcluidosByGroup(DataStore.rawExcluidos));
    this.renderNuevos(filteredMain);
    this.renderHistorico(DataStore.rawAntiguoEx);
  },
};


/* ────────────────────────────────────────────────────────────
   6. FILTER ENGINE — gestiona filtros y opciones dinámicas
──────────────────────────────────────────────────────────── */
const FilterEngine = {

  /* Puebla los selects de filtros con valores únicos del dataset */
  populate(records) {
    // Grupos únicos
    const groups = [...new Set(records.map(r => r.grupo))].sort();
    const selGroup = document.getElementById('filterGroup');
    if (selGroup) {
      selGroup.innerHTML = '<option value="">Todos los grupos</option>' +
        groups.map(g => `<option value="${g}">${g.replace(/Ministr[ao]s?\s*/i,'') || g}</option>`).join('');
    }

    // Estados únicos
    const estados = [...new Set(records.map(r => r.estado).filter(Boolean))].sort();
    const selEstado = document.getElementById('filterEstado');
    if (selEstado) {
      selEstado.innerHTML = '<option value="">Todos</option>' +
        estados.map(e => `<option value="${e}">${e}</option>`).join('');
    }
  },

  /* Lee los filtros actuales de los selects */
  read() {
    DataStore.filters.group    = document.getElementById('filterGroup')?.value    || '';
    DataStore.filters.estado   = document.getElementById('filterEstado')?.value   || '';
    DataStore.filters.celula   = document.getElementById('filterCelula')?.value   || '';
    DataStore.filters.servicio = document.getElementById('filterServicio')?.value || '';
    DataStore.filters.nuevo    = document.getElementById('filterNuevo')?.value    || '';
  },

  /* Resetea todos los filtros */
  reset() {
    ['filterGroup','filterEstado','filterCelula','filterServicio','filterNuevo']
      .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
    DataStore.filters = { group:'', estado:'', celula:'', servicio:'', nuevo:'' };
  },
};


/* ────────────────────────────────────────────────────────────
   7. UI CONTROLLER — coordina todo
──────────────────────────────────────────────────────────── */
const UIController = {

  /* Inicializa el controlador y vincula eventos */
  init() {
    // Vincula el clic en las tarjetas KPI al modal de detalle de personas
    this.bindKPICardClicks();

    // Cargar archivo
    document.getElementById('fileInput')?.addEventListener('change', e => {
      const file = e.target.files[0];
      if (file) this.loadFile(file);
      e.target.value = ''; // Permite recargar el mismo archivo
    });

    // Pantalla Completa (solo visible/relevante en móviles — ver style.css).
    // Fullscreen API estándar de HTML5: si el documento NO está en pantalla
    // completa, la solicita (oculta la barra de navegación del teléfono);
    // si ya lo está, sale. No toca ningún motor de datos, sesión ni
    // notificaciones — es un toggle puramente de presentación del navegador.
    document.getElementById('btnDesktopView')?.addEventListener('click', function () {
      const btn = this;
      const elFullscreen =
        document.fullscreenElement ||
        document.webkitFullscreenElement ||   // Safari/iOS
        document.msFullscreenElement;         // IE/Edge viejo

      if (!elFullscreen) {
        const el = document.documentElement;
        const request =
          el.requestFullscreen ||
          el.webkitRequestFullscreen ||
          el.msRequestFullscreen;

        if (request) {
          request.call(el).catch(err => {
            console.error('[UIController] No se pudo entrar en pantalla completa:', err);
          });
        }
      } else {
        const exit =
          document.exitFullscreen ||
          document.webkitExitFullscreen ||
          document.msExitFullscreen;

        if (exit) {
          exit.call(document).catch(err => {
            console.error('[UIController] No se pudo salir de pantalla completa:', err);
          });
        }
      }

      btn.classList.toggle('active');
    });

    // Mantiene sincronizado el estado visual (.active) del botón si el
    // usuario sale de pantalla completa por otra vía (ej. gesto del
    // sistema operativo o tecla Esc), no solo con el propio clic.
    document.addEventListener('fullscreenchange', () => {
      const btn = document.getElementById('btnDesktopView');
      if (!btn) return;
      const enFullscreen = !!(document.fullscreenElement || document.webkitFullscreenElement || document.msFullscreenElement);
      btn.classList.toggle('active', enFullscreen);
    });

    // Toggle excluidos — botón oculto visualmente en topbar (ver index.html).
    // Listener comentado para evitar referencias muertas; la lógica de
    // DataStore.includeExcluidos / this.refresh() permanece intacta para uso futuro.
    // document.getElementById('toggleExcluidos')?.addEventListener('change', e => {
    //   DataStore.includeExcluidos = e.target.checked;
    //   this.refresh();
    // });

    // Filtros: actualizar en cambio. El filtro de Grupo Ministerial es el
    // único que puede implicar un salto grande de volumen de datos (de un
    // grupo específico a "Todos los grupos" o viceversa), así que es el
    // único que dispara el overlay extendido "Recalculando…" en las
    // gráficas (ver refresh({ groupChange })).
    ['filterGroup','filterEstado','filterCelula','filterServicio','filterNuevo']
      .forEach(id => {
        document.getElementById(id)?.addEventListener('change', () => {
          FilterEngine.read();
          this.refresh({ groupChange: id === 'filterGroup' });
        });
      });

    // Reset filtros — también puede saltar de/hacia "Todos los grupos"
    document.getElementById('btnResetFilters')?.addEventListener('click', () => {
      FilterEngine.reset();
      this.refresh({ groupChange: true });
    });

    // Búsqueda en tablas
    this.bindTableSearch('searchPersonas',  'tablePersonas',  'countPersonas');
    this.bindTableSearch('searchExcluidos', 'tableExcluidos', 'countExcluidos');
    this.bindTableSearch('searchNuevos',    'tableNuevos',    'countNuevos');
    this.bindTableSearch('searchHistorico', 'tableHistorico', 'countHistorico');
  },

  /* Vincula el evento de búsqueda a una tabla */
  bindTableSearch(inputId, tableId, countId) {
    document.getElementById(inputId)?.addEventListener('input', e => {
      const visible = TableEngine.filterTable(tableId, e.target.value);
      document.getElementById(countId).textContent = `${visible} registros`;
    });
  },

  /* Carga y procesa el archivo Excel */
  loadFile(file) {
    this.showLoading(true);
    DataStore.fileName = file.name;

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = e.target.result; // ArrayBuffer
        const workbook = await WorkerEngine.parseWorkbookAsync(data); // procesado en Web Worker, no bloquea la UI
        ExcelParser.parse(workbook);
        FilterEngine.populate(DataStore.rawMain);
        this.refresh();
        this.showDashboard();
        document.getElementById('footerFile').textContent = file.name;
        document.getElementById('reportTitle').textContent = DataStore.reportTitle || file.name;

        /* Guarda el buffer en DataStore y habilita el botón de subida */
        DataStore.rawBuffer = e.target.result;
        CloudEngine.enableUploadBtn(file.name);
        SaveEngine.enable(file.name);

        /* Notificación de auditoría: carga de archivo local (fire-and-forget).
           Cubre tanto el input principal de la pantalla de carga como el
           botón "Cargar Excel" de la barra superior, ya que ambos apuntan
           al mismo #fileInput y disparan este mismo flujo. */
        const auditUser = AuditEngine.getUser();
        if (auditUser) {
          AuditEngine.notify({ action: 'cargar_local', user: auditUser, fileName: file.name });
        }
      } catch (err) {
        console.error('Error al procesar el archivo:', err);
        alert(`Error al leer el archivo:\n${err.message}`);
      } finally {
        this.showLoading(false);
      }
    };
    reader.onerror = () => {
      this.showLoading(false);
      alert('No se pudo leer el archivo.');
    };
    reader.readAsArrayBuffer(file);
  },

  /* Recalcula KPIs, actualiza gráficos y tablas con filtros aplicados.

     FASE 1 (síncrona, instantánea): marca los valores de las tarjetas
     KPI como "Recalculando…" — feedback visual inmediato de que el
     cambio de filtro sí se registró, antes de arrancar el trabajo
     pesado (compute + render de tablas/gráficos). Si el cambio es de
     Grupo Ministerial (groupChange:true), también cubre cada tarjeta
     de gráfico con el mismo overlay "Recalculando…".

     FASE 2 (diferida): ejecuta el mismo trabajo que antes hacía
     refresh() de forma síncrona — ni el orden ni la lógica de
     cómputo/filtrado cambian, solo el momento en que corre. Para
     cambios de grupo se añade un margen deliberado (~1.8s, dentro del
     tope de 2-3s) antes de recalcular: le da colchón al cómputo más
     pesado (salto grande de volumen de datos, p. ej. a "Todos los
     grupos") y, detrás del overlay, cualquier micro-lag real queda
     disimulado en vez de sentirse como que el dashboard "se traba".
     Para el resto de filtros (estado, célula, servicio, nuevo) el
     comportamiento es igual de instantáneo que antes (doble rAF). */
  refresh(opts = {}) {
    const groupChange = !!opts.groupChange;

    this._setKpisRecalculando(true);
    if (groupChange) this._setChartsRecalculando(true);

    const run = () => this._performRefresh({ groupChange });

    if (groupChange) {
      const GROUP_CHANGE_DELAY_MS = 1800; // margen deliberado — tope pedido: 2-3s
      setTimeout(() => requestAnimationFrame(run), GROUP_CHANGE_DELAY_MS);
    } else {
      requestAnimationFrame(() => requestAnimationFrame(run));
    }
  },

  /* Activa/desactiva el estado visual "Recalculando…" (amarillo, fiel
     al tema — var(--gold)) sobre los números de las tarjetas KPI.
     No toca TrendEngine ni ningún dato: updateKPICards() sobrescribe
     estos mismos nodos con el valor real en cuanto termina el cómputo. */
  _setKpisRecalculando(active) {
    document.querySelectorAll('.kpi-value, .kpi-pct').forEach(el => {
      el.classList.toggle('kpi-recalculando', active);
      if (active) el.textContent = 'Recalculando…';
    });
  },

  /* Activa/desactiva el overlay "Recalculando…" sobre cada tarjeta de
     gráfico (.chart-card). Es puramente visual (CSS ::after, ver
     style.css): no desmonta el canvas ni ninguna instancia de
     Chart.js, así que ChartEngine.renderAll() de abajo sigue
     funcionando exactamente igual. */
  _setChartsRecalculando(active) {
    document.querySelectorAll('.chart-card').forEach(el => {
      el.classList.toggle('chart-recalculando', active);
    });
  },

  /* Trabajo real de refresco — idéntico, en el mismo orden, al que
     antes vivía directamente en refresh(). */
  _performRefresh(opts = {}) {
    const active   = DataStore.getActiveMain();
    const filtered = DataStore.applyFilters(active);
    const kpis     = KPIEngine.compute(filtered);

    /* Guarda el último conjunto filtrado para que el modal de detalle
       de las tarjetas KPI (bindKPICardClicks) siempre liste exactamente
       las mismas personas que están detrás de los números mostrados,
       sin tener que recalcular filtros por su cuenta. */
    this._lastFilteredRecords = filtered;

    this.updateKPICards(kpis);
    TrendEngine.render(kpis); // NUEVO — pinta flechas/% y sparklines si hay línea base activa
    ChartEngine.renderAll(kpis);
    TableEngine.renderAll(filtered);
    AbsenceEngine.render(filtered);  // Monitor de ausencias

    this._setKpisRecalculando(false);
    if (opts.groupChange) this._setChartsRecalculando(false);
  },

  /**
   * Hace clickeables las tarjetas de KPI declaradas en KPI_DETAIL_MAP:
   * al hacer clic, abre ModalEngine con el título de la métrica y la
   * lista de personas que la conforman (KPIEngine.getRecordsByMetric
   * sobre el último array ya filtrado — respeta AccessManager y los
   * filtros activos de la UI). Se llama una sola vez desde init().
   */
  bindKPICardClicks() {
    Object.entries(KPI_DETAIL_MAP).forEach(([valueId, { metric, title }]) => {
      const valueEl = document.getElementById(valueId);
      const cardEl  = valueEl?.closest('.kpi-card');
      if (!cardEl) return; // Fail-safe visual: la tarjeta no existe en esta vista

      cardEl.classList.add('kpi-card-clickable');
      cardEl.setAttribute('role', 'button');
      cardEl.setAttribute('tabindex', '0');
      cardEl.setAttribute('aria-label', `Ver personas: ${title}`);

      const abrirDetalle = () => {
        const records = this._lastFilteredRecords || [];
        const subset  = KPIEngine.getRecordsByMetric(records, metric);
        if (window.ModalEngine) window.ModalEngine.open(title, subset);
      };

      cardEl.addEventListener('click', abrirDetalle);
      // Accesibilidad: también abre con Enter/Espacio si la tarjeta tiene foco
      cardEl.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          abrirDetalle();
        }
      });
    });
  },

  /* Actualiza los valores en los cards de KPI */
  updateKPICards(kpis) {
    const set = (id, val) => {
      const el = document.getElementById(id);
      if (el) el.textContent = val;
    };

    set('kpiTotal',        kpis.total);
    set('kpiCelulasSI',    kpis.celulasSI);
    set('kpiCelulasSIPct', `${kpis.celulasSIPct}%`);
    set('kpiCelulasNO',    kpis.celulasNO);
    set('kpiCelulasNOPct', `${kpis.celulasNOPct}%`);
    set('kpiServicioSI',   kpis.servicioSI);
    set('kpiServicioSIPct',`${kpis.servicioSIPct}%`);
    set('kpiServicioNO',   kpis.servicioNO);
    set('kpiServicioNOPct',`${kpis.servicioNOPct}%`);
    set('kpiAmbosSI',      kpis.ambosSI);
    set('kpiAmbosSIPct',   `${kpis.ambosSIPct}%`);
    set('kpiAmbosNO',      kpis.ambosNO);
    set('kpiAmbosNOPct',   `${kpis.ambosNOPct}%`);
    set('kpiNuevosCelula',  kpis.nuevosCelula);
    set('kpiNuevosServicio',kpis.nuevosServicio);
    set('kpiPctGeneral',   `${kpis.pctGeneral}%`);
    set('kpiPctCelula',    `${kpis.pctCelula}%`);
    set('kpiPctServicio',  `${kpis.pctServicio}%`);
  },

  /* Muestra u oculta el overlay de carga */
  showLoading(show) {
    const el = document.getElementById('loadingOverlay');
    if (!el) return;
    el.classList.toggle('d-none', !show);
  },

  /* Muestra el dashboard y oculta el estado vacío */
  showDashboard() {
    document.getElementById('emptyState')?.classList.add('d-none');
    document.getElementById('dashboardContent')?.classList.remove('d-none');
  },
};


/* ────────────────────────────────────────────────────────────
   7A. INITIAL LOAD OVERLAY — "Cargando su experiencia personalizada"
   Overlay independiente de UIController.showLoading()/#loadingOverlay
   (ese sigue igual, se usa para procesar archivos manuales). Este
   módulo solo controla #initialLoadOverlay, acoplado ÚNICAMENTE al
   flujo de auto-carga del reporte predeterminado + tendencia guardada
   que corre justo después del login (ver SessionEngine más abajo:
   AutoLoadEngine.loadDefaultFile().then(() => DbDefaultEngine...)).
   No modifica la lógica de AutoLoadEngine/DbDefaultEngine: solo
   envuelve esa misma llamada con show()/setProgress()/complete(). */
const InitialLoadOverlay = {
  _el: null,
  _fillEl: null,
  _hideTimer: null,

  _ensureEl() {
    if (!this._el) {
      this._el = document.getElementById('initialLoadOverlay');
      this._fillEl = document.getElementById('initialLoadFill');
    }
    return this._el;
  },

  /* Aparece de inmediato y arranca la barra en un pequeño % (feedback
     instantáneo de que "ya empezó a pasar algo"), en vez de en 0. */
  show() {
    const el = this._ensureEl();
    if (!el) return;
    clearTimeout(this._hideTimer);
    el.classList.remove('initial-load-hidden');
    el.classList.remove('d-none');
    this.setProgress(8);
  },

  /* Actualiza el ancho de la barra (0–100). Se llama en cada paso
     real del proceso (reporte cargado, tendencia aplicada, etc.),
     así que refleja avance real y no una animación falsa/indefinida. */
  setProgress(pct) {
    if (!this._ensureEl() || !this._fillEl) return;
    const clamped = Math.max(0, Math.min(100, pct));
    this._fillEl.style.width = clamped + '%';
  },

  /* Completa la barra al 100% y desvanece el overlay con fade-out
     suave (transición CSS de opacity). SIEMPRE debe llamarse al
     terminar el proceso, tanto en éxito como en error, para no dejar
     al usuario con el overlay trabado en pantalla. */
  complete() {
    const el = this._ensureEl();
    if (!el) return;
    this.setProgress(100);

    clearTimeout(this._hideTimer);
    this._hideTimer = setTimeout(() => {
      el.classList.add('initial-load-hidden'); // dispara la transición de opacity en CSS
      setTimeout(() => {
        el.classList.add('d-none');
        this.setProgress(0); // listo por si se vuelve a usar (p. ej. otro login en la misma pestaña)
      }, 500); // debe coincidir con la duración de la transición definida en style.css
    }, 250); // breve pausa para que se alcance a ver la barra llena al 100%
  },
};


/* ────────────────────────────────────────────────────────────
   7B. ABSENCE ENGINE — Monitor de ausencias y alertas
   Calcula días sin asistir y clasifica por nivel de alerta
──────────────────────────────────────────────────────────── */
const AbsenceEngine = {

  /*
    NIVELES DE ALERTA:
    ─────────────────────────────────────────────
    normal   →  0–6 días   (< 1 semana)
    watch    →  7–13 días  (1–2 semanas)  "Seguimiento"
    warn     → 14–27 días  (2–4 semanas)  "Advertencia"
    critical → 28+ días    (> 4 semanas)  "Crítico"
    ─────────────────────────────────────────────
    Solo se procesan personas con fecha de última falta registrada
    Y cuya asistencia actual sea NO en ambas (célula y servicio).
  */

  LEVELS: [
    { key: 'normal',   label: 'Normal',       maxDays: 6,  cls: 'alert-normal', icon: '●' },
    { key: 'watch',    label: 'Seguimiento',   maxDays: 13, cls: 'alert-watch',  icon: '◉' },
    { key: 'warn',     label: 'Advertencia',   maxDays: 27, cls: 'alert-warn',   icon: '▲' },
    { key: 'critical', label: 'Crítico',       maxDays: Infinity, cls: 'alert-crit', icon: '⚠' },
  ],

  /* Calcula el nivel de alerta según días de ausencia */
  getLevel(days) {
    return this.LEVELS.find(l => days <= l.maxDays) || this.LEVELS[3];
  },

  /* Formatea el tiempo transcurrido en texto legible */
  formatTime(days) {
    if (days < 0)    return { main: 'Hoy', detail: '' };
    if (days === 0)  return { main: 'Hoy', detail: '' };
    if (days === 1)  return { main: '1 día', detail: '' };
    if (days < 7)    return { main: `${days} días`, detail: '' };
    if (days < 14)   return { main: '1 semana', detail: `${days} días` };
    if (days < 30) {
      const w = Math.floor(days / 7);
      const d = days % 7;
      return { main: `${w} sem${w > 1 ? 's' : ''}`, detail: d ? `${days} días` : `${days} días` };
    }
    if (days < 365) {
      const m = Math.floor(days / 30.44);
      const d = days - Math.round(m * 30.44);
      return {
        main:   `${m} mes${m > 1 ? 'es' : ''}`,
        detail: `${days} días totales`,
      };
    }
    const y = Math.floor(days / 365);
    const m = Math.floor((days % 365) / 30.44);
    return {
      main:   `${y} año${y > 1 ? 's' : ''}${m ? ` ${m} mes${m > 1 ? 'es' : ''}` : ''}`,
      detail: `${days} días totales`,
    };
  },

  /* Procesa los registros y devuelve los datos de ausencia */
  process(records) {
    const today = new Date();
    today.setHours(0,0,0,0);

    const result = [];

    records.forEach(r => {
      if (!r.fecha) return;  // Sin fecha registrada, no aplica

      // Parsear fecha
      const parts = r.fecha.split('-');
      if (parts.length < 3) return;
      const fechaDate = new Date(
        parseInt(parts[0]),
        parseInt(parts[1]) - 1,
        parseInt(parts[2])
      );
      if (isNaN(fechaDate.getTime())) return;

      const diffMs   = today.getTime() - fechaDate.getTime();
      const days     = Math.round(diffMs / (1000 * 60 * 60 * 24));
      const level    = this.getLevel(days);
      const timeFmt  = this.formatTime(days);

      result.push({
        ...r,
        diasAusente:   days,
        nivel:         level.key,
        levelObj:      level,
        timeFmt,
        fechaFormatted: fechaDate.toLocaleDateString('es-VE', {
          day: '2-digit', month: 'short', year: 'numeric'
        }),
      });
    });

    // Ordenar: críticos primero, luego por días descendente
    result.sort((a, b) => {
      const levelOrder = { critical: 0, warn: 1, watch: 2, normal: 3 };
      const lo = (levelOrder[a.nivel] ?? 4) - (levelOrder[b.nivel] ?? 4);
      if (lo !== 0) return lo;
      return b.diasAusente - a.diasAusente;
    });

    return result;
  },

  /* Renderiza la tabla completa de ausencias */
  render(records) {
    const data = this.process(records);

    // Actualizar summary cards
    const counts = { normal: 0, watch: 0, warn: 0, critical: 0 };
    data.forEach(r => { if (counts[r.nivel] !== undefined) counts[r.nivel]++; });

    const setEl = (id, val) => { const e = document.getElementById(id); if (e) e.textContent = val; };
    setEl('ausNormalCount', counts.normal);
    setEl('ausWatchCount',  counts.watch);
    setEl('ausWarnCount',   counts.warn);
    setEl('ausCritCount',   counts.critical);

    // Actualizar contador
    const counter = document.getElementById('countAusencias');
    if (counter) counter.textContent = `${data.length} con fecha registrada`;

    // Guardar datos para filtro por nivel
    this._currentData = data;
    this.renderRows(data);
  },

  /* Renderiza las filas según el nivel activo */
  renderRows(data) {
    const tbody = document.getElementById('tableAusenciasBody');
    if (!tbody) return;

    this._visibleData = data;  // usado por ListExportEngine (respeta nivel + búsqueda)

    if (data.length === 0) {
      tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;color:var(--text-dim);padding:32px">
        No hay registros con fecha de falta disponibles
      </td></tr>`;
      return;
    }

    tbody.innerHTML = data.map((r, i) => {
      const lvl = r.levelObj;
      const dot = `<span class="alert-dot"></span>`;
      const alertPill = `<span class="alert-pill ${lvl.cls}">${dot}${lvl.icon} ${lvl.label}</span>`;

      const timeHtml = `
        <div class="time-badge">${r.timeFmt.main}</div>
        ${r.timeFmt.detail ? `<div class="time-detail">${r.timeFmt.detail}</div>` : ''}
      `;

      const grpShort = (r.grupo || '')
        .replace(/Ministr[ao]s?\s*/i,'')
        .replace(/Lider[a]?\s*/i,'')
        .trim().substring(0,26);

      return `<tr data-level="${r.nivel}">
        <td>${i+1}</td>
        <td style="font-weight:500">${r.nombre}</td>
        <td style="color:var(--text-dim);font-size:11px">${grpShort}</td>
        <td>${TableEngine.waLink(r.telefono)}</td>
        <td style="font-size:12px;color:var(--text-dim)">${r.fechaFormatted}</td>
        <td>${timeHtml}</td>
        <td>$$BADGE_C$$</td>
        <td>$$BADGE_S$$</td>
        <td style="color:var(--text-dim);font-size:11px">${r.estado || '—'}</td>
        <td>${alertPill}</td>
      </tr>`.replace('$$BADGE_C$$', TableEngine.badge(r.celula))
             .replace('$$BADGE_S$$', TableEngine.badge(r.servicio));
    }).join('');
  },

  /* Filtro por nivel y texto */
  filterRows(levelKey, searchText) {
    let data = this._currentData || [];
    if (levelKey) data = data.filter(r => r.nivel === levelKey);
    if (searchText) {
      const q = searchText.toLowerCase();
      data = data.filter(r =>
        r.nombre.toLowerCase().includes(q) ||
        r.grupo.toLowerCase().includes(q)
      );
    }
    this.renderRows(data);
    const counter = document.getElementById('countAusencias');
    if (counter) counter.textContent = `${data.length} registros`;
  },

  _currentData: [],
  _visibleData: [],
  _activeLevel: '',
  _activeSearch: '',

  /* Inicializa los eventos de filtro */
  initEvents() {
    // Botones de nivel
    document.querySelectorAll('.aus-filter-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.aus-filter-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        this._activeLevel = btn.dataset.level || '';
        this.filterRows(this._activeLevel, this._activeSearch);
      });
    });

    // Búsqueda
    document.getElementById('searchAusencias')?.addEventListener('input', e => {
      this._activeSearch = e.target.value;
      this.filterRows(this._activeLevel, this._activeSearch);
    });
  },
};


/* ────────────────────────────────────────────────────────────
   8. GOOGLE SHEETS SYNC ENGINE
   Convierte cualquier URL de Google Sheets al endpoint CSV
   y sincroniza automáticamente según el intervalo elegido.
──────────────────────────────────────────────────────────── */
const GSheetsEngine = {

  /* Estado interno */
  state: {
    url:        '',       // URL CSV activa
    timer:      null,     // ID del setInterval de auto-sync
    interval:   60,       // segundos entre sincronizaciones
    connected:  false,
    syncing:    false,
    lastSync:   null,     // Date del último sync exitoso
    modalRef:   null,     // Instancia Bootstrap modal
  },

  /* ── Convierte cualquier URL de Google Sheets a CSV export ── */
  toCsvUrl(raw) {
    raw = raw.trim();

    // Ya es un CSV publicado correcto
    if (raw.includes('pub?') && raw.includes('output=csv')) return raw;
    if (raw.includes('/pub?') || raw.includes('&output=csv'))  return raw;

    // URL normal: https://docs.google.com/spreadsheets/d/ID/edit#gid=GID
    const matchId  = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    const matchGid = raw.match(/[#&]gid=(\d+)/);

    if (matchId) {
      const id  = matchId[1];
      const gid = matchGid ? matchGid[1] : '0';
      return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
    }

    // URL de publicación sin output=csv
    if (raw.includes('/pub')) {
      return raw.includes('?') ? raw + '&output=csv' : raw + '?output=csv';
    }

    return raw; // Devuelve tal cual, intentamos igual
  },

  /* ── Descarga el CSV y lo convierte a workbook SheetJS ── */
  async fetchCsv(csvUrl) {
    // Usamos un proxy CORS gratuito para evitar bloqueos del navegador
    // cuando la hoja se descarga directamente (CORS policy de Google)
    const proxyUrl = `https://corsproxy.io/?${encodeURIComponent(csvUrl)}`;

    const resp = await fetch(proxyUrl, {
      cache: 'no-store',
      headers: { 'Accept': 'text/csv,*/*' },
    });

    if (!resp.ok) throw new Error(`HTTP ${resp.status} — ${resp.statusText}`);

    const text = await resp.text();
    if (text.length < 10) throw new Error('La hoja está vacía o no es pública');

    // Verifica que no sea una página de error HTML de Google
    if (text.trimStart().startsWith('<!DOCTYPE') || text.trimStart().startsWith('<html')) {
      throw new Error('La hoja no es pública o el enlace es incorrecto. Verifica "Publicar en la web".');
    }

    // Convierte CSV → workbook SheetJS (como si fuera un Excel de una hoja)
    const wb = XLSX.read(text, { type: 'string', raw: false });
    return wb;
  },

  /* ── Sincroniza: descarga, parsea y actualiza el dashboard ── */
  async sync(silent = false) {
    if (this.state.syncing) return;
    this.state.syncing = true;
    this._setDotState('syncing');

    try {
      const wb = await this.fetchCsv(this.state.url);
      ExcelParser.parse(wb);
      FilterEngine.populate(DataStore.rawMain);
      UIController.refresh();
      UIController.showDashboard();

      this.state.lastSync = new Date();
      this.state.connected = true;
      this._setDotState('live');
      this._updateConnStatus();

      // Actualiza nombre del reporte en topbar
      document.getElementById('reportTitle').textContent =
        DataStore.reportTitle || 'Google Sheets — En vivo';
      document.getElementById('footerFile').textContent =
        '🟢 Google Sheets · Última sync: ' + this._timeStr(this.state.lastSync);

      if (!silent) this._toast('Sincronizado correctamente', 'success');
      this._hideError();

    } catch (err) {
      console.error('GSheets sync error:', err);
      this._setDotState('error');
      this._showError(err.message);
      if (!silent) this._toast('Error al sincronizar: ' + err.message, 'error');
    } finally {
      this.state.syncing = false;
    }
  },

  /* ── Inicia la conexión y el temporizador ── */
  connect(rawUrl, intervalSec) {
    this.disconnect(); // Limpia timer anterior
    this.state.url      = this.toCsvUrl(rawUrl);
    this.state.interval = parseInt(intervalSec, 10);

    // Primera sincronización inmediata
    this.sync(false);

    // Configura auto-sync si el intervalo es > 0
    if (this.state.interval > 0) {
      this.state.timer = setInterval(
        () => this.sync(true),
        this.state.interval * 1000
      );
    }

    // Actualiza UI del modal
    this._setConnectedUI(true);
  },

  /* ── Desconecta y limpia ── */
  disconnect() {
    if (this.state.timer) {
      clearInterval(this.state.timer);
      this.state.timer = null;
    }
    this.state.connected = false;
    this.state.url       = '';
    this._setDotState('');
    this._setConnectedUI(false);
    document.getElementById('footerFile').textContent = 'Sin archivo';
    document.getElementById('reportTitle').textContent = 'Cargue un archivo para comenzar';
  },

  /* ── Helpers de UI ── */

  _setDotState(state) {
    // Dot en topbar
    const dot = document.getElementById('gsheetStatus');
    if (dot) { dot.className = 'gsheet-status'; if (state) dot.classList.add(state); }
    // Dot en modal
    const connDot = document.getElementById('connDot');
    if (connDot) { connDot.className = 'conn-dot'; if (state) connDot.classList.add(state); }
  },

  _updateConnStatus() {
    const label    = document.getElementById('connLabel');
    const lastSync = document.getElementById('connLastSync');
    const status   = document.getElementById('gsheetConnStatus');

    if (status) status.classList.remove('d-none');
    if (label) label.textContent = this.state.connected ? '🟢 Conectado' : 'Desconectado';
    if (lastSync && this.state.lastSync) {
      const intText = this.state.interval > 0
        ? ` · Próxima sync en ~${this.state.interval}s`
        : ' · Modo manual';
      lastSync.textContent = 'Última sync: ' + this._timeStr(this.state.lastSync) + intText;
    }
  },

  _setConnectedUI(connected) {
    const btnConn   = document.getElementById('btnConnectSheet');
    const btnDisc   = document.getElementById('btnDisconnect');
    const connStatus = document.getElementById('gsheetConnStatus');

    if (connected) {
      if (btnDisc)    btnDisc.classList.remove('d-none');
      if (connStatus) connStatus.classList.remove('d-none');
      if (btnConn)    btnConn.innerHTML = '<i class="bi bi-arrow-repeat me-2"></i>Re-sincronizar';
    } else {
      if (btnDisc)    btnDisc.classList.add('d-none');
      if (connStatus) connStatus.classList.add('d-none');
      if (btnConn)    btnConn.innerHTML = '<i class="bi bi-link-45deg me-2"></i>Conectar y Sincronizar';
    }
  },

  _showError(msg) {
    const box = document.getElementById('gsheetError');
    const txt = document.getElementById('gsheetErrorMsg');
    if (box) box.classList.remove('d-none');
    if (txt) txt.textContent = msg;
  },

  _hideError() {
    const box = document.getElementById('gsheetError');
    if (box) box.classList.add('d-none');
  },

  _timeStr(date) {
    if (!date) return '—';
    return date.toLocaleTimeString('es-VE', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  },

  /* Muestra una notificación temporal */
  _toast(msg, type = 'info') {
    const existing = document.querySelector('.sync-toast');
    if (existing) existing.remove();

    const icons = { success: '✅', error: '❌', info: '🔄' };
    const toast = document.createElement('div');
    toast.className = `sync-toast ${type}`;
    toast.innerHTML = `<span>${icons[type] || '•'}</span><span>${msg}</span>`;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 4000);
  },

  /* ── Lee la URL según el modo activo (csv / url) ── */
  _getActiveUrl() {
    const mode = document.querySelector('.gsheet-mode-tab.active')?.dataset.mode || 'csv';
    if (mode === 'csv') {
      return document.getElementById('gsheetCsvUrl')?.value.trim() || '';
    } else {
      return document.getElementById('gsheetNormalUrl')?.value.trim() || '';
    }
  },

  /* ── Lee el intervalo seleccionado ── */
  _getInterval() {
    const checked = document.querySelector('input[name="syncInterval"]:checked');
    return checked ? parseInt(checked.value, 10) : 60;
  },

  /* ── Inicializa todos los eventos del modal ── */
  initModal() {
    const modalEl = document.getElementById('gsheetsModal');
    if (!modalEl) return;
    this.state.modalRef = new bootstrap.Modal(modalEl);

    // Abrir modal desde topbar — botón oculto visualmente (ver index.html).
    // Listener comentado; this.state.modalRef.show() y toda la lógica de
    // sincronización con Google Sheets permanecen intactas para uso futuro.
    // document.getElementById('btnGsheets')?.addEventListener('click', () => {
    //   this.state.modalRef.show();
    // });

    // Abrir modal desde empty state
    document.getElementById('btnGsheetsEmpty')?.addEventListener('click', () => {
      this.state.modalRef.show();
    });

    // Tabs modo CSV / URL normal
    document.querySelectorAll('.gsheet-mode-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.gsheet-mode-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const mode = tab.dataset.mode;
        document.getElementById('gsheetInputCsv').classList.toggle('d-none', mode !== 'csv');
        document.getElementById('gsheetInputUrl').classList.toggle('d-none', mode !== 'url');
      });
    });

    // Botón Conectar
    document.getElementById('btnConnectSheet')?.addEventListener('click', () => {
      const url = this._getActiveUrl();
      if (!url) { this._showError('Por favor ingresa un enlace de Google Sheets'); return; }
      this._hideError();
      const interval = this._getInterval();
      this.connect(url, interval);
      this._updateConnStatus();
    });

    // Botón Sincronizar ahora (dentro del modal)
    document.getElementById('btnSyncNow')?.addEventListener('click', () => {
      if (this.state.url) this.sync(false);
    });

    // Botón Desconectar
    document.getElementById('btnDisconnect')?.addEventListener('click', () => {
      this.disconnect();
      this._toast('Desconectado de Google Sheets', 'info');
    });

    // Al abrir modal, rellenar URL si ya hay una activa
    modalEl.addEventListener('show.bs.modal', () => {
      this._hideError();
      if (this.state.url) {
        document.getElementById('gsheetCsvUrl').value = this.state.url;
      }
      this._updateConnStatus();
    });
  },
};

/* Expone GSheetsEngine en window: LazyModals.js comprueba
   `window.GSheetsEngine` antes de llamar a initModal() la primera
   vez que se abre el modal de Google Sheets. Una declaración `const`
   de nivel superior en un script clásico NO se agrega automáticamente
   como propiedad de `window` (a diferencia de `var`/`function`), así
   que sin esta línea esa comprobación siempre era `false` y
   GSheetsEngine.initModal() nunca llegaba a ejecutarse. */
window.GSheetsEngine = GSheetsEngine;


/* ────────────────────────────────────────────────────────────
   9. BOOTSTRAP — arranque cuando el DOM esté listo
──────────────────────────────────────────────────────────── */
/* ────────────────────────────────────────────────────────────
   THEME ENGINE — Modo claro / oscuro
   Aplica data-theme="light" | "dark" al <html>
   Actualiza Chart.js defaults para colores de ejes/grid
──────────────────────────────────────────────────────────── */
const ThemeEngine = {

  STORAGE_KEY: 'iglesia_dash_theme',

  /* Devuelve el tema activo */
  current() {
    return document.documentElement.getAttribute('data-theme') || 'dark';
  },

  /* Aplica el tema y sincroniza todo */
  apply(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem(this.STORAGE_KEY, theme);
    this._updateIcon(theme);
    this._updateChartDefaults(theme);

    // Si ya hay gráficos creados, se actualizan sus colores in-place
    // (sin destruir/recrear) para que el cambio de tema sea instantáneo
    // y no produzca el parpadeo de "las gráficas desaparecen".
    const hayGraficosCreados = Object.keys(ChartEngine.instances).length > 0;

    if (hayGraficosCreados) {
      ChartEngine.updateTheme(theme);
    } else if (!document.getElementById('dashboardContent').classList.contains('d-none')) {
      // Respaldo: dashboard visible pero sin instancias todavía
      // (mismo comportamiento que existía antes de este cambio).
      const active   = DataStore.getActiveMain();
      const filtered = DataStore.applyFilters(active);
      const kpis     = KPIEngine.compute(filtered);
      ChartEngine.renderAll(kpis);
    }
  },

  /* Alterna entre claro y oscuro */
  toggle() {
    this.apply(this.current() === 'dark' ? 'light' : 'dark');
  },

  /* Actualiza el icono del botón */
  _updateIcon(theme) {
    const icon = document.getElementById('themeIcon');
    if (!icon) return;
    // Oscuro → mostrar sol (para cambiar a claro)
    // Claro  → mostrar luna (para cambiar a oscuro)
    icon.className = theme === 'dark' ? 'bi bi-sun-fill' : 'bi bi-moon-fill';
  },

  /* Actualiza los defaults globales de Chart.js */
  _updateChartDefaults(theme) {
    const isLight = theme === 'light';
    // Light mode: use near-black so axis labels, ticks and legends are clearly readable
    const textColor   = isLight ? '#0f172a'          : '#94a3b8';
    const gridColor   = isLight ? 'rgba(0,0,0,.08)'  : 'rgba(255,255,255,.04)';
    const tooltipBg   = isLight ? '#ffffff'           : '#1c2333';
    const tooltipBorder = isLight ? 'rgba(0,0,0,.12)' : 'rgba(255,255,255,.07)';
    const tooltipTitle  = isLight ? '#0f172a'          : '#e2e8f0';
    const tooltipBody   = isLight ? '#1e293b'          : '#94a3b8';

    // Scale defaults
    Chart.defaults.color = textColor;
    Chart.defaults.borderColor = gridColor;

    // Plugin defaults
    Chart.defaults.plugins.tooltip.backgroundColor = tooltipBg;
    Chart.defaults.plugins.tooltip.borderColor     = tooltipBorder;
    Chart.defaults.plugins.tooltip.titleColor      = tooltipTitle;
    Chart.defaults.plugins.tooltip.bodyColor       = tooltipBody;

    // Legend
    Chart.defaults.plugins.legend.labels.color = textColor;
  },

  /* Inicializa: carga preferencia guardada o usa oscuro por defecto */
  init() {
    const saved = localStorage.getItem(this.STORAGE_KEY) || 'dark';
    this.apply(saved);

    document.getElementById('btnTheme')?.addEventListener('click', () => {
      this.toggle();
    });
  },
};


/* ────────────────────────────────────────────────────────────
   9.5 AUDIT ENGINE — Capa de seguridad/auditoría compartida
       Obtiene el nombre del usuario con sesión iniciada
       (guardado por SessionEngine en sessionStorage) y notifica
       por Telegram las acciones sensibles (Eliminar / Cargar /
       Descargar / Guardar).
──────────────────────────────────────────────────────────── */
const AuditEngine = {

  /**
   * Devuelve el nombre del usuario actualmente en sesión.
   * Ya no se pregunta con prompt(): el nombre se registró una
   * única vez al iniciar sesión y se reutiliza para todas las
   * acciones de auditoría.
   *
   * @returns {string|null} Nombre del usuario o null si no hay sesión activa
   */
  getUser() {
    return SessionEngine.getUser();
  },

  /**
   * Dispara la notificación de Telegram sin bloquear la UI.
   * Los errores se registran en consola pero nunca interrumpen
   * el flujo de la acción principal.
   */
  notify({ action, user, fileName, extra }) {
    if (typeof TelegramEngine === 'undefined') {
      console.warn('[AuditEngine] TelegramEngine no está disponible; se omite la notificación.');
      return;
    }
    /* Fire-and-forget: no se usa await para no bloquear la interfaz */
    TelegramEngine.notify({ action, user, fileName, extra })
      .catch(err => console.error('[AuditEngine] Error al notificar por Telegram:', err));
  },
};


/* ────────────────────────────────────────────────────────────
   9.6 SESSION ENGINE — Pantalla de inicio de sesión + auditoría
       Controla el overlay de login/logout, persiste el nombre
       del usuario en sessionStorage y notifica por Telegram
       cada inicio/cierre de sesión.
──────────────────────────────────────────────────────────── */
/* ────────────────────────────────────────────────────────────
   9.5 CLOCK ENGINE — reloj en tiempo real de la topbar
   Módulo pequeño e independiente: solo pinta #sessionClock cada
   segundo. Arranca únicamente desde SessionEngine._updateSessionUI()
   (después de un login exitoso o al restaurar sesión), nunca antes.
──────────────────────────────────────────────────────────── */
const ClockEngine = {
  _intervalId: null,

  _format(date) {
    const dia = date.toLocaleDateString('es-PE', { weekday: 'long' });
    const diaCap = dia.charAt(0).toUpperCase() + dia.slice(1);
    const fecha = date.toLocaleDateString('es-PE', { day: '2-digit', month: 'short', year: 'numeric' });
    const hora = date.toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
    return `${diaCap}, ${fecha} · ${hora}`;
  },

  _tick() {
    const el = document.getElementById('sessionClock');
    if (!el) return; // el bloque puede estar oculto/ausente antes del login
    el.textContent = this._format(new Date());
  },

  /* Idempotente: si ya hay un interval corriendo, no crea otro
     (evita duplicar timers si se llama más de una vez por error). */
  start() {
    if (this._intervalId) return;
    this._tick();
    this._intervalId = setInterval(() => this._tick(), 1000);
  },

  stop() {
    clearInterval(this._intervalId);
    this._intervalId = null;
  },
};


const SessionEngine = {

  STORAGE_KEY: 'ccrm_dashboard_user',

  _el(id) { return document.getElementById(id); },

  /** Devuelve el nombre de usuario en sesión, o null si no hay sesión activa */
  getUser() {
    const name = sessionStorage.getItem(this.STORAGE_KEY);
    return (name && name.trim() !== '') ? name : null;
  },

  /** true si hay una sesión activa */
  isLoggedIn() {
    return this.getUser() !== null;
  },

  /* ── Muestra el overlay de login (sin animación, estado inicial) ── */
  showOverlay() {
    const overlay = this._el('loginOverlay');
    if (!overlay) return;
    overlay.classList.remove('login-overlay-hidden', 'login-overlay-fadeout');
    overlay.style.display = 'flex';
    /* Reinicia al estado "landing" (logo a la izquierda + botón) */
    overlay.classList.remove('login-state-active');
    const nameStep = this._el('loginNameStep');
    if (nameStep) nameStep.classList.remove('login-name-step-visible');
    const input = this._el('loginNameInput');
    if (input) {
      input.value = '';
      input.type = 'password'; // Siempre vuelve a mostrarse enmascarado
    }
    const toggleIcon = this._el('loginNameToggleIcon');
    if (toggleIcon) toggleIcon.className = 'bi bi-eye';
    this._el('loginNameError')?.classList.add('d-none');
    this._el('loginNameNotFoundError')?.classList.add('d-none');
    this._el('btnSolicitarSoporte')?.classList.add('d-none');
    /* Reinicia también el paso de soporte por si quedó abierto */
    this._el('loginSupportStep')?.classList.remove('login-support-step-visible');
    this._el('loginSupportForm')?.classList.remove('d-none');
    this._el('loginSupportSent')?.classList.add('d-none');
    this._el('loginSupportError')?.classList.add('d-none');
    const supportPhoneInput = this._el('supportPhoneInput');
    if (supportPhoneInput) supportPhoneInput.value = '';
  },

  /* ── Oculta el overlay instantáneamente (sin fade), usado al cargar con sesión ya activa ── */
  hideOverlayInstant() {
    const overlay = this._el('loginOverlay');
    if (!overlay) return;
    overlay.style.display = 'none';
    overlay.classList.add('login-overlay-hidden');
  },

  /* ── Paso 1 → 2: centra el logo y revela el input de nombre ── */
  _revealNameStep() {
    const overlay = this._el('loginOverlay');
    if (!overlay) return;
    overlay.classList.add('login-state-active');
    const nameStep = this._el('loginNameStep');
    setTimeout(() => {
      if (nameStep) nameStep.classList.add('login-name-step-visible');
      this._el('loginNameInput')?.focus();
    }, 450); // Coincide con la duración de la transición del logo (ver CSS)
  },

  /**
   * Devuelve el array de usuarios autorizados definido en USUARIOS.JS
   * (variable global `USUARIOS_REGISTRADOS`, cargada vía <script src="USUARIOS.JS">
   * antes que app.js — ver index.html). Se usa un global en vez de fetch()
   * porque USUARIOS.JS no es un JSON válido (es un archivo .js con `const`),
   * y además evita problemas de CORS si el dashboard se abre con file://
   * en vez de servirse desde un servidor HTTP.
   *
   * Devuelve `null` (en vez de []) cuando la variable no está disponible,
   * para poder distinguir "archivo cargado pero vacío" de "no se pudo cargar".
   */
  async _fetchUsuariosAutorizados() {
    try {
      if (typeof USUARIOS_REGISTRADOS === 'undefined') {
        throw new Error('USUARIOS.JS no se cargó (variable USUARIOS_REGISTRADOS no definida).');
      }
      return Array.isArray(USUARIOS_REGISTRADOS) ? USUARIOS_REGISTRADOS : [];
    } catch (err) {
      console.error('[SessionEngine] Error al cargar USUARIOS.JS:', err);
      return null;
    }
  },

  /* ── Confirma el nombre: valida contra USUARIOS.JS antes de abrir sesión ── */
  async _confirmLogin() {
    const input = this._el('loginNameInput');
    const name = (input?.value || '').trim();

    const errEl = this._el('loginNameError');
    const notFoundEl = this._el('loginNameNotFoundError');
    const supportBtn = this._el('btnSolicitarSoporte');
    const confirmBtn = this._el('btnConfirmarNombre');

    if (name === '') {
      if (errEl) errEl.classList.remove('d-none');
      notFoundEl?.classList.add('d-none');
      supportBtn?.classList.add('d-none');
      input?.focus();
      return;
    }
    if (errEl) errEl.classList.add('d-none');
    notFoundEl?.classList.add('d-none');

    /* Deshabilita el botón mientras se verifica contra USUARIOS.JS */
    if (confirmBtn) confirmBtn.disabled = true;
    const usuarios = await this._fetchUsuariosAutorizados();
    if (confirmBtn) confirmBtn.disabled = false;

    if (usuarios === null) {
      /* Error de red/lectura del archivo: por seguridad no se permite el
         acceso, pero se ofrece la vía de soporte igualmente */
      if (notFoundEl) {
        notFoundEl.textContent = 'No se pudo verificar tu usuario (error de conexión). Intenta de nuevo o solicita soporte.';
        notFoundEl.classList.remove('d-none');
      }
      supportBtn?.classList.remove('d-none');
      input?.focus();
      return;
    }

    const nameUpper = name.toUpperCase();
    const isAuthorized = usuarios.some(u => String(u).trim().toUpperCase() === nameUpper);

    if (!isAuthorized) {
      if (notFoundEl) {
        notFoundEl.textContent = 'Usuario no encontrado. Verifica el nombre ingresado.';
        notFoundEl.classList.remove('d-none');
      }
      supportBtn?.classList.remove('d-none');

      /* Notificación de auditoría: intento de login fallido (fire-and-forget),
         disparada antes de que el usuario vea la opción de "Solicitar soporte" */
      if (typeof TelegramEngine !== 'undefined') {
        TelegramEngine.notifyFailedLogin(name)
          .catch(err => console.error('[SessionEngine] Error al notificar login fallido:', err));
      }

      input?.focus();
      return;
    }

    supportBtn?.classList.add('d-none');
    sessionStorage.setItem(this.STORAGE_KEY, name);

    /* Aplica los permisos de INTERFAZ (Usuario Rules.js) para el usuario
       recién autenticado — no toca datos ni filtros RBAC (AccessManager) */
    if (window.UsuarioRules) window.UsuarioRules.applyUIPermissions(name);

    /* Notificación de auditoría por Telegram (fire-and-forget) */
    if (typeof TelegramEngine !== 'undefined') {
      TelegramEngine.notifySession('login', name)
        .catch(err => console.error('[SessionEngine] Error al notificar inicio de sesión:', err));
    }

    /* Auto-carga del archivo predeterminado + tendencia guardada.
       Extraído a _autoLoadDefaultExperience() (ver más abajo) para
       poder reutilizarlo también en init() cuando la sesión ya
       estaba activa (p. ej. tras SwitchSessionEngine.js), caso que
       antes se quedaba sin auto-cargar nada. */
    this._autoLoadDefaultExperience();

    /* Desvanece el overlay y revela el dashboard */
    const overlay = this._el('loginOverlay');
    if (overlay) {
      overlay.classList.add('login-overlay-fadeout');
      setTimeout(() => {
        overlay.style.display = 'none';
        overlay.classList.add('login-overlay-hidden');
      }, 500); // Coincide con la duración del fade-out (ver CSS)
    }

    /* Refleja el usuario en sesión donde corresponda en la UI */
    this._updateSessionUI(name);
  },

  /* ── Paso 2 → 3: oculta el formulario de nombre y revela el de soporte,
       pre-llenando el nombre que el usuario intentó ingresar ── */
  _showSupportStep() {
    const nameStep = this._el('loginNameStep');
    const supportStep = this._el('loginSupportStep');
    const attemptedName = (this._el('loginNameInput')?.value || '').trim();

    nameStep?.classList.remove('login-name-step-visible');
    setTimeout(() => {
      const supportUserInput = this._el('supportUserInput');
      if (supportUserInput) supportUserInput.value = attemptedName;
      supportStep?.classList.add('login-support-step-visible');
      this._el('supportPhoneInput')?.focus();
    }, 300);
  },

  /* ── Regresa del paso de soporte (o de la confirmación) al formulario de nombre,
       reiniciando el subformulario de soporte a su estado inicial ── */
  _backToLoginFromSupport() {
    const nameStep = this._el('loginNameStep');
    const supportStep = this._el('loginSupportStep');

    supportStep?.classList.remove('login-support-step-visible');

    setTimeout(() => {
      this._el('loginSupportForm')?.classList.remove('d-none');
      this._el('loginSupportSent')?.classList.add('d-none');
      const phoneInput = this._el('supportPhoneInput');
      if (phoneInput) phoneInput.value = '';
      this._el('loginSupportError')?.classList.add('d-none');
      nameStep?.classList.add('login-name-step-visible');
    }, 300);
  },

  /* ── Valida y envía la solicitud de soporte (nombre + teléfono) por Telegram ── */
  async _submitSupportRequest() {
    const userInput = this._el('supportUserInput');
    const phoneInput = this._el('supportPhoneInput');
    const errEl = this._el('loginSupportError');

    const user = (userInput?.value || '').trim();
    const phone = (phoneInput?.value || '').trim();

    /* Debe iniciar con '+', seguido de dígitos y espacios opcionales,
       con un mínimo de 10 caracteres en total */
    const phoneRegex = /^\+[\d\s]+$/;
    const digitCount = (phone.match(/\d/g) || []).length;
    const isValidPhone = phone.length >= 10 && phoneRegex.test(phone) && digitCount >= 9;

    if (user === '' || !isValidPhone) {
      if (errEl) errEl.classList.remove('d-none');
      phoneInput?.focus();
      return;
    }
    if (errEl) errEl.classList.add('d-none');

    const btn = this._el('btnEnviarSoporte');
    if (btn) btn.disabled = true;

    if (typeof TelegramEngine !== 'undefined') {
      try {
        await TelegramEngine.notifySupport(user, phone);
      } catch (err) {
        console.error('[SessionEngine] Error al notificar solicitud de soporte:', err);
      }
    }

    if (btn) btn.disabled = false;

    this._el('loginSupportForm')?.classList.add('d-none');
    this._el('loginSupportSent')?.classList.remove('d-none');
  },

  /* ── Cierra la sesión: notifica, limpia storage, y recarga la app
       para vaciar por completo las métricas/gráficos cargados
       (así la siguiente sesión siempre arranca desde cero) ── */
  async logout() {
    const name = this.getUser();
    sessionStorage.removeItem(this.STORAGE_KEY);

    if (name && typeof TelegramEngine !== 'undefined') {
      try {
        /* Se espera el envío (con límite de 1.5s) para no perder la
           notificación al recargar la página inmediatamente después */
        await Promise.race([
          TelegramEngine.notifySession('logout', name),
          new Promise(resolve => setTimeout(resolve, 1500)),
        ]);
      } catch (err) {
        console.error('[SessionEngine] Error al notificar cierre de sesión:', err);
      }
    }

    /* Recarga completa: limpia DataStore, gráficos, tablas y filtros
       en memoria, dejando el dashboard listo para el próximo usuario */
    window.location.reload();
  },

  /* Icono por rol (LECTOR ⭐ / EDITOR ⭐⭐ / MAESTRO 👑) */
  _ROLE_ICONS: { LECTOR: '⭐', EDITOR: '⭐⭐', MAESTRO: '👑' },
  _ROLE_CLASS: { LECTOR: 'role-lector', EDITOR: 'role-editor', MAESTRO: 'role-maestro' },

  /* ── Actualiza referencias visuales del usuario activo (si existieran) ── */
  _updateSessionUI(name) {
    const label = this._el('sessionUserLabel');
    if (label) label.textContent = name;

    /* Icono de rol: reutiliza UsuarioRules._resolveRole(), el mismo
       método que ya usa app.js directamente en otro punto (ver
       DbDefaultEngine más abajo) — no se modifica Usuario_Rules.js. */
    const roleIcon = this._el('sessionRoleIcon');
    if (roleIcon && window.UsuarioRules) {
      const role = window.UsuarioRules._resolveRole(name);
      roleIcon.textContent = this._ROLE_ICONS[role] || '';
      roleIcon.className = `session-role-icon ${this._ROLE_CLASS[role] || ''}`;
      roleIcon.title = `Rol: ${role}`;
    }

    /* Revela el bloque (usuario + rol + reloj) — permanece oculto
       hasta este punto, que solo se alcanza tras un login exitoso o
       al restaurar una sesión ya activa (ver init() más abajo). */
    this._el('sessionInfoBlock')?.classList.remove('d-none');

    /* Arranca el reloj en tiempo real de la topbar (idempotente) */
    ClockEngine.start();
  },

  /* ── Auto-carga del archivo predeterminado (config.json → REPORTES/<archivo>)
       + tendencia guardada. Fire-and-forget: no bloquea el fade-out del
       overlay de login ni el flujo que la llame. Inmediatamente DESPUÉS de
       que termine de cargar (encadenado con .then, no en paralelo), se
       revisa si hay una tendencia predeterminada guardada en localStorage
       (ver DbDefaultEngine) y se aplica automáticamente.

       InitialLoadOverlay envuelve esta misma llamada para mostrar la
       pantalla "Cargando su experiencia personalizada" con una barra de
       progreso real: 8% al empezar, 65% cuando el reporte ya cargó, 100%
       cuando la tendencia (si había una guardada) también terminó — y se
       desvanece con fade-out. Se usa .finally() para garantizar que el
       overlay SIEMPRE desaparezca, incluso si algo falla, sin cambiar en
       nada el comportamiento de AutoLoadEngine/DbDefaultEngine.

       Se llama tanto desde _confirmLogin() (login normal) como desde
       init() cuando la sesión YA estaba activa al cargar la página (p. ej.
       tras un reload disparado por SwitchSessionEngine.js) — antes solo
       corría en el primer caso, por lo que cambiar de sesión dejaba el
       dashboard sin el Excel/tendencia predeterminados. ── */
  _autoLoadDefaultExperience() {
    if (typeof AutoLoadEngine === 'undefined') return;

    InitialLoadOverlay.show();

    AutoLoadEngine.loadDefaultFile()
      .then(() => {
        InitialLoadOverlay.setProgress(65);
        if (typeof DbDefaultEngine !== 'undefined') {
          return DbDefaultEngine.applyStoredTrendIfAny();
        }
      })
      .catch(err => {
        console.error('[SessionEngine] Error durante la auto-carga de la experiencia predeterminada:', err);
      })
      .finally(() => {
        InitialLoadOverlay.complete();
      });
  },

  init() {
    /* Si ya existe una sesión (misma pestaña), no se muestra el login */
    if (this.isLoggedIn()) {
      this.hideOverlayInstant();
      this._updateSessionUI(this.getUser());
      /* Restaura los permisos de INTERFAZ (Usuario Rules.js) para el
         usuario ya autenticado — no afecta datos ni filtros RBAC */
      if (window.UsuarioRules) window.UsuarioRules.applyUIPermissions(this.getUser());
      /* Misma auto-carga que ocurre en un login normal: sin esto, un
         reload con sesión ya activa (p. ej. tras Cambiar Sesión) dejaba
         el dashboard vacío hasta que el usuario cargara un Excel a mano. */
      this._autoLoadDefaultExperience();
    } else {
      this.showOverlay();
    }

    /* Botón "Iniciar Sesión" (paso 1 → 2) */
    this._el('btnIniciarSesion')?.addEventListener('click', () => this._revealNameStep());

    /* Confirmar nombre (botón + Enter) */
    this._el('btnConfirmarNombre')?.addEventListener('click', () => this._confirmLogin());
    this._el('loginNameInput')?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') this._confirmLogin();
    });

    /* Alterna mostrar/ocultar el nombre escrito (type password <-> text).
       Puramente visual: this._el('loginNameInput').value sigue
       capturándose igual sin importar el `type` del input. */
    this._el('btnToggleLoginName')?.addEventListener('click', () => {
      const input = this._el('loginNameInput');
      const icon  = this._el('loginNameToggleIcon');
      if (!input) return;
      const isPassword = input.type === 'password';
      input.type = isPassword ? 'text' : 'password';
      if (icon) icon.className = isPassword ? 'bi bi-eye-slash' : 'bi bi-eye';
    });

    /* Flujo de soporte: usuario no encontrado en USUARIOS.JS */
    this._el('btnSolicitarSoporte')?.addEventListener('click', () => this._showSupportStep());
    this._el('btnEnviarSoporte')?.addEventListener('click', () => this._submitSupportRequest());
    this._el('supportPhoneInput')?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') this._submitSupportRequest();
    });
    this._el('btnVolverLoginDesdeSoporte')?.addEventListener('click', () => this._backToLoginFromSupport());
    this._el('btnVolverLoginFinal')?.addEventListener('click', () => this._backToLoginFromSupport());

    /* Botón "Cerrar sesión" en la barra lateral del Menú */
    this._el('btnCerrarSesion')?.addEventListener('click', () => {
      /* Cierra el offcanvas del menú si estuviera abierto */
      const offcanvasEl = this._el('sidebarMenu');
      if (offcanvasEl && window.bootstrap) {
        const instance = bootstrap.Offcanvas.getInstance(offcanvasEl);
        instance?.hide();
      }
      this.logout();
    });
  },
};


/* ────────────────────────────────────────────────────────────
   10. HISTORY ENGINE — Panel lateral con repositorio GitHub
       Consulta la API de GitHub para listar y cargar archivos
       Excel (.xlsx, .xlsm, .xls) desde la carpeta REPORTES.
──────────────────────────────────────────────────────────── */
const HistoryEngine = {

  /* ── Configuración ── */
  GITHUB_API: 'https://api.github.com/repos/alexchouriors/M-tricas-REPORTE-DE-ASISTENCIAS-NUEVA/contents/REPORTES',
  VALID_EXTS: ['.xlsx', '.xlsm', '.xls'],

  /* ── Estado interno ── */
  _files:       [],   // Lista de archivos obtenidos de la API
  _loadingFile: false, // Previene cargas simultáneas

  /* ── Utilidades de DOM ── */
  _el(id) { return document.getElementById(id); },

  /* Muestra solo uno de los estados del panel */
  _setState(state) {
    const states = { loading: 'historyLoading', error: 'historyError', empty: 'historyEmpty' };
    Object.entries(states).forEach(([key, id]) => {
      const el = this._el(id);
      if (!el) return;
      el.classList.toggle('d-none', key !== state);
    });

    const listEl = this._el('listaReportesContainer');
    if (listEl) listEl.classList.toggle('d-none', state !== 'list');
  },

  /* Formatea tamaño de bytes */
  _fmtSize(bytes) {
    if (!bytes || bytes === 0) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  },

  /* Obtiene la extensión del nombre de archivo */
  _ext(name) {
    const m = name.toLowerCase().match(/\.(xlsx|xlsm|xls)$/);
    return m ? '.' + m[1] : '';
  },

  /* Clase de icono según extensión */
  _iconClass(ext) {
    const map = { '.xlsx': 'bi-file-earmark-spreadsheet history-file-ext-xlsx',
                  '.xlsm': 'bi-file-earmark-spreadsheet history-file-ext-xlsm',
                  '.xls':  'bi-file-earmark-spreadsheet history-file-ext-xls'  };
    return map[ext] || 'bi-file-earmark';
  },

  /* Clase de badge según extensión */
  _badgeClass(ext) {
    const map = { '.xlsx': 'badge-xlsx', '.xlsm': 'badge-xlsm', '.xls': 'badge-xls' };
    return map[ext] || '';
  },

  /* ── Consulta la API de GitHub ── */
  async fetchFileList() {
    this._setState('loading');

    try {
      const headers = { 'Accept': 'application/vnd.github.v3+json' };
      const token = AuthEngine.getToken();
      if (token) headers['Authorization'] = `Bearer ${token}`;

      const res = await fetch(this.GITHUB_API, { headers });

      if (!res.ok) {
        const msg = res.status === 404
          ? 'Repositorio o carpeta no encontrada (404).'
          : res.status === 403
            ? 'Límite de peticiones a la API de GitHub excedido. Intenta en unos minutos.'
            : `Error ${res.status}: ${res.statusText}`;
        throw new Error(msg);
      }

      const items = await res.json();

      /* Filtrar solo archivos con extensión Excel válida */
      this._files = items.filter(item =>
        item.type === 'file' && this.VALID_EXTS.includes(this._ext(item.name))
      );

      if (this._files.length === 0) {
        this._setState('empty');
        return;
      }

      this._renderList();
      this._setState('list');

    } catch (err) {
      this._showError(err.message || 'Error desconocido al contactar la API de GitHub.');
    }
  },

  /* ── Muestra el estado de error con mensaje ── */
  _showError(msg) {
    const msgEl = this._el('historyErrorMsg');
    if (msgEl) msgEl.textContent = msg;
    this._setState('error');
  },

  /* ── Renderiza la lista de archivos ── */
  _renderList() {
    const listEl = this._el('listaReportesContainer');
    if (!listEl) return;

    listEl.innerHTML = '';
    this._files.forEach((file, idx) => {
      const ext  = this._ext(file.name);
      const item = document.createElement('div');
      item.className = 'list-group-item d-flex align-items-center justify-content-between flex-wrap gap-2';
      item.dataset.idx = idx;

      item.innerHTML = `
        <div class="d-flex align-items-center gap-2 text-truncate">
          <i class="bi ${this._iconClass(ext)} text-success fs-5"></i>
          <span class="text-truncate" title="${file.name}">${file.name}</span>
        </div>
        <div class="d-flex align-items-center gap-2 ms-auto">
          <button type="button" class="btn btn-sm btn-outline-primary btn-cargar-reporte" data-idx="${idx}">
            <i class="bi bi-cloud-arrow-down me-1"></i>Cargar al Dashboard
          </button>
          <button type="button" class="btn btn-sm btn-outline-warning btn-comparar-reporte" data-idx="${idx}">
            <i class="bi bi-bar-chart-line me-1"></i>Comparar
          </button>
          <button type="button" class="btn btn-sm btn-outline-info btn-tendencia-reporte" data-idx="${idx}">
            <i class="bi bi-graph-up-arrow me-1"></i>Tendencia
          </button>
          <a class="btn btn-sm btn-outline-success" href="${file.download_url}" target="_blank" rel="noopener noreferrer">
            <i class="bi bi-download me-1"></i>Descargar
          </a>
        </div>
      `;

      /* Botón "Cargar al Dashboard": valida usuario (auditoría) antes de ejecutar la lógica existente */
      item.querySelector('.btn-cargar-reporte').addEventListener('click', () => {
        const user = AuditEngine.getUser();
        if (!user) return; // Sin sesión activa: se aborta la acción

        AuditEngine.notify({ action: 'cargar', user, fileName: file.name });
        this._loadFile(file, item);
      });

      /* Botón "Comparar": genera la Comparativa Histórica SIN tocar el
         reporte actualmente cargado en el dashboard (ver ComparativaEngine) */
      item.querySelector('.btn-comparar-reporte').addEventListener('click', () => {
        const user = AuditEngine.getUser();
        if (!user) return; // Sin sesión activa: se aborta la acción

        ComparativaEngine.compare(file);

        /* Notificación de auditoría por Telegram (fire-and-forget) */
        if (typeof TelegramEngine !== 'undefined') {
          TelegramEngine.notifyFeatureUsed(user, 'Comparó sus datos.')
            .catch(err => console.error('[HistoryEngine] Error al notificar uso de Comparar:', err));
        }
      });

      /* Botón "Tendencia": fija este archivo como línea base de
         tendencia para las flechas/% y sparklines de las tarjetas KPI
         del dashboard principal (ver TrendEngine) */
      item.querySelector('.btn-tendencia-reporte').addEventListener('click', () => {
        const user = AuditEngine.getUser();
        if (!user) return; // Sin sesión activa: se aborta la acción

        TrendEngine.setBaseline(file);

        /* Notificación de auditoría por Telegram (fire-and-forget) */
        if (typeof TelegramEngine !== 'undefined') {
          TelegramEngine.notifyFeatureUsed(user, 'Revisó su tendencia.')
            .catch(err => console.error('[HistoryEngine] Error al notificar uso de Tendencia:', err));
        }
      });

      /* Botón "Descargar": registra al usuario en sesión antes de permitir la descarga */
      const btnDescargar = item.querySelector('a.btn-outline-success');
      if (btnDescargar) {
        btnDescargar.addEventListener('click', (e) => {
          const user = AuditEngine.getUser();
          if (!user) { e.preventDefault(); return; } // Sin sesión activa: se aborta la acción

          AuditEngine.notify({ action: 'descargar', user, fileName: file.name });
          /* No se hace preventDefault: el <a href> sigue su curso normal de descarga */
        });
      }

      listEl.appendChild(item);
    });
  },

  /* ── Descarga y procesa el archivo seleccionado ── */
  async _loadFile(file, itemEl) {
    if (this._loadingFile) return;
    this._loadingFile = true;

    /* UI: marcar item activo y mostrar spinner global */
    itemEl.classList.add('loading');
    const loadingOverlay = this._el('historyFileLoading');
    const loadingName    = this._el('historyFileLoadingName');
    if (loadingOverlay) loadingOverlay.classList.remove('d-none');
    if (loadingName)    loadingName.textContent = file.name;

    try {
      /* Usa la download_url que provee la API de GitHub */
      const url = file.download_url;
      if (!url) throw new Error('El archivo no tiene URL de descarga disponible.');

      const res = await fetch(url);
      if (!res.ok) throw new Error(`No se pudo descargar el archivo (${res.status}).`);

      const buffer   = await res.arrayBuffer();
      const workbook = await WorkerEngine.parseWorkbookAsync(buffer); // offload a Web Worker (no bloquea la UI)

      /* Guarda el nombre en DataStore y parsea con ExcelParser existente */
      DataStore.fileName = file.name;
      ExcelParser.parse(workbook);

      /* Actualiza filtros, gráficos y dashboard (mismo flujo que carga local) */
      FilterEngine.populate(DataStore.rawMain);
      UIController.refresh();
      UIController.showDashboard();
      const footerEl = document.getElementById('footerFile');
      if (footerEl) footerEl.textContent = file.name;
      const titleEl = document.getElementById('reportTitle');
      if (titleEl) titleEl.textContent = DataStore.reportTitle || file.name;

      /* Cierra el Modal tras cargar exitosamente */
      const modalEl = this._el('modalHistorial');
      if (modalEl) {
        const bsModal = bootstrap.Modal.getInstance(modalEl);
        if (bsModal) bsModal.hide();
      }

    } catch (err) {
      /* Muestra el error dentro del panel para no interrumpir el dashboard */
      this._showError(`Error al cargar "${file.name}": ${err.message}`);
    } finally {
      this._loadingFile = false;
      itemEl.classList.remove('loading');
      if (loadingOverlay) loadingOverlay.classList.add('d-none');
    }
  },

  /* ── Inicialización: eventos y primera carga ── */
  init() {
    /* Al abrirse el Modal, carga la lista si aún no hay archivos */
    const modalEl = document.getElementById('modalHistorial');
    if (!modalEl) return;

    modalEl.addEventListener('show.bs.modal', () => {
      /* Solo hace fetch si la lista está vacía o en estado de error/inicial */
      const listEl = this._el('listaReportesContainer');
      const hasItems = listEl && listEl.children.length > 0;
      if (!hasItems) this.fetchFileList();
    });

    /* Botón "Reintentar" en estado de error */
    this._el('btnHistoryRetry')?.addEventListener('click', () => this.fetchFileList());

    /* Botón ícono recargar dentro del info-box del modal */
    this._el('btnRecargarRepo')?.addEventListener('click', () => {
      const btn  = this._el('btnRecargarRepo');
      const icon = this._el('reloadRepoIcon');
      if (btn) btn.disabled = true;
      if (btn)  btn.classList.add('spinning');
      // reinicia animación CSS
      if (icon) { icon.style.animation = 'none'; void icon.offsetWidth; icon.style.animation = ''; }
      this._files = [];
      const listEl = this._el('listaReportesContainer');
      if (listEl) listEl.innerHTML = '';
      this.fetchFileList().finally(() => {
        if (btn) { btn.disabled = false; btn.classList.remove('spinning'); }
      });
    });
  },
};

/* Expone HistoryEngine en window: LazyModals.js comprueba
   `window.HistoryEngine` como callback `onFirstOpen` al inyectar
   #modalHistorial, para llamar a HistoryEngine.init() (que registra
   el listener 'show.bs.modal' encargado de disparar fetchFileList()).
   Al ser `HistoryEngine` una declaración `const` de nivel superior en
   un script clásico, NO se agrega automáticamente como propiedad de
   `window` (a diferencia de `var`/`function`) — por eso esa
   comprobación siempre resultaba `false`, init() nunca se ejecutaba,
   el listener de apertura del modal jamás quedaba enganchado y el
   botón "Historial" se quedaba mostrando el spinner de carga para
   siempre, sin llegar a pedir la lista de archivos a GitHub. */
window.HistoryEngine = HistoryEngine;


/* ────────────────────────────────────────────────────────────
   10.5 COMPARATIVA ENGINE — Comparativa Histórica de KPIs
       Descarga y parsea un archivo del historial de forma AISLADA
       (ExcelParser.parseStandalone — ver arriba) sin sobrescribir
       DataStore.rawMain ni ningún dato del reporte que el usuario
       está viendo actualmente. El resultado siempre pasa por
       AccessManager.applyFilter() antes de calcular ningún KPI.
──────────────────────────────────────────────────────────── */
const ComparativaEngine = {

  _loading: false,
  _modalRef: null,

  _el(id) { return document.getElementById(id); },

  /* ── Punto de entrada: botón "Comparar" de un item del historial ── */
  async compare(file) {
    if (this._loading) return;

    if (!Array.isArray(DataStore.rawMain) || DataStore.rawMain.length === 0) {
      alert('Primero carga un reporte en el dashboard antes de generar una comparativa histórica.');
      return;
    }

    this._loading = true;
    this._showLoading(file.name);

    try {
      const url = file.download_url;
      if (!url) throw new Error('El archivo no tiene URL de descarga disponible.');

      const res = await fetch(url);
      if (!res.ok) throw new Error(`No se pudo descargar el archivo (${res.status}).`);

      const buffer   = await res.arrayBuffer();
      const workbook = await WorkerEngine.parseWorkbookAsync(buffer); // offload a Web Worker (no bloquea la UI)

      /* Parseo AISLADO — no toca el reporte actualmente cargado.
         Ya viene filtrado por AccessManager.applyFilter() (obligatorio
         y fail-closed, ver ExcelParser.parseStandalone). */
      const historico = ExcelParser.parseStandalone(workbook);

      /* KPIs del reporte ACTUAL — mismos datos que ven las tarjetas en
         pantalla ahora mismo (ya filtrados por RBAC + filtros de UI activos) */
      const activeNow    = DataStore.applyFilters(DataStore.getActiveMain());
      const kpisActual   = KPIEngine.compute(activeNow);

      /* KPIs del reporte ANTERIOR — se le aplican los MISMOS filtros de
         UI activos ahora mismo (grupo, estado, célula, servicio, nuevo),
         para que la comparación sea simétrica: "[Grupo X] Actual" vs.
         "[Grupo X] Anterior", nunca "[Grupo X] Actual" vs. "Todos los
         grupos Anterior". Sin esto, con un grupo filtrado el histórico
         se comparaba contra el total general sin filtrar. */
      const historicoFiltrado = DataStore.applyFilters(historico.rawMain);
      const kpisAnterior = KPIEngine.compute(historicoFiltrado);

      this._render(file.name, kpisActual, kpisAnterior);

      const modalEl = this._el('modalComparativa');
      if (modalEl) {
        this._modalRef = this._modalRef || new bootstrap.Modal(modalEl);
        this._modalRef.show();
      }

    } catch (err) {
      alert(`No se pudo generar la comparativa histórica:\n${err.message}`);
    } finally {
      this._loading = false;
      this._hideLoading();
    }
  },

  /* ── Reutiliza el overlay de carga que ya usa "Cargar al Dashboard" ── */
  _showLoading(fileName) {
    const overlay = this._el('historyFileLoading');
    const label   = this._el('historyFileLoadingName');
    if (overlay) overlay.classList.remove('d-none');
    if (label)   label.textContent = `Comparando contra "${fileName}"...`;
  },

  _hideLoading() {
    const overlay = this._el('historyFileLoading');
    if (overlay) overlay.classList.add('d-none');
  },

  /* ── Calcula la variación % con la fórmula solicitada, a salvo de
       división por cero ── */
  _calcVariacion(actual, anterior) {
    if (anterior === 0) {
      if (actual === 0) return { text: '0%', cls: 'comparativa-pct-neutral' };
      return { text: 'N/A', cls: 'comparativa-pct-neutral' };
    }
    const variacion   = ((actual - anterior) / anterior) * 100;
    const redondeado  = Math.round(variacion * 10) / 10;
    const cls  = redondeado > 0 ? 'comparativa-pct-pos'
               : redondeado < 0 ? 'comparativa-pct-neg'
               : 'comparativa-pct-neutral';
    const signo = redondeado > 0 ? '+' : '';
    return { text: `${signo}${redondeado}%`, cls };
  },

  /* ── Pinta la tabla del modal ── */
  _render(fileName, kpisActual, kpisAnterior) {
    const subtitleEl = this._el('modalComparativaSubtitle');
    if (subtitleEl) subtitleEl.textContent = `Reporte actual vs. "${fileName}"`;

    const tbody = this._el('comparativaTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    /* Reutiliza KPI_DETAIL_MAP (misma fuente de verdad que el resto
       del dashboard) para no duplicar nombres de métricas ni títulos */
    Object.values(KPI_DETAIL_MAP).forEach(({ metric, title }) => {
      const actual   = kpisActual[metric]   ?? 0;
      const anterior = kpisAnterior[metric] ?? 0;
      const { text, cls } = this._calcVariacion(actual, anterior);

      const row = document.createElement('tr');
      row.innerHTML = `
        <td>${title}</td>
        <td class="text-center">${actual}</td>
        <td class="text-center">${anterior}</td>
        <td class="text-center"><span class="comparativa-pct-badge ${cls}">${text}</span></td>
      `;
      tbody.appendChild(row);
    });
  },
};


/* ────────────────────────────────────────────────────────────
   10.6 SPARKLINE ENGINE — mini-gráficos de 2 puntos (antes/ahora)
       en las tarjetas KPI. Usa Chart.js (ya cargado por ChartEngine)
       pero mantiene su PROPIO registro de instancias (`instances`),
       separado de ChartEngine.instances, sobre canvases con id
       distinto (`${valueId}Spark`) — cero colisión con los gráficos
       grandes que ya gestiona ChartEngine.
──────────────────────────────────────────────────────────── */
const SparklineEngine = {

  instances: {},

  /* Dibuja o actualiza el sparkline de una tarjeta KPI */
  render(valueId, anterior, actual, direction) {
    const canvas = document.getElementById(`${valueId}Spark`);
    if (!canvas || typeof Chart === 'undefined') return;

    const color = direction === 'up'   ? '#22c55e'
                : direction === 'down' ? '#ef4444'
                :                        '#94a3b8';

    if (this.instances[valueId]) {
      const inst = this.instances[valueId];
      inst.data.datasets[0].data = [anterior, actual];
      inst.data.datasets[0].borderColor = color;
      inst.data.datasets[0].pointBackgroundColor = color;
      inst.update();
      return;
    }

    this.instances[valueId] = new Chart(canvas.getContext('2d'), {
      type: 'line',
      data: {
        labels: ['Antes', 'Ahora'],
        datasets: [{
          data: [anterior, actual],
          borderColor: color,
          backgroundColor: 'transparent',
          borderWidth: 2,
          pointRadius: 2,
          pointHoverRadius: 3,
          pointBackgroundColor: color,
          tension: 0.35,
        }],
      },
      options: {
        responsive: false,
        maintainAspectRatio: false,
        animation: false,
        layout: { padding: 2 },
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        scales: {
          x: { display: false },
          y: { display: false },
        },
      },
    });
  },

  /* Destruye una instancia puntual (la visibilidad la controla el
     contenedor .kpi-trend-row, gestionado por TrendEngine) */
  clear(valueId) {
    const inst = this.instances[valueId];
    if (inst) {
      inst.destroy();
      delete this.instances[valueId];
    }
  },

  /* Oculta/destruye todas las instancias (al quitar la línea base) */
  clearAll() {
    Object.keys(this.instances).forEach(id => this.clear(id));
  },
};


/* ────────────────────────────────────────────────────────────
   10.7 TREND ENGINE — Tendencia (▲/▼) en las tarjetas KPI del
       Dashboard principal, EXTENDIENDO KPIEngine/DataStore/
       HistoryEngine sin modificar su lógica existente:
         • DataStore: solo se le agrega la propiedad
           `comparisonBaseline` (no se toca ninguna función).
         • KPIEngine: usa el nuevo método aditivo computeDelta().
         • HistoryEngine: reutiliza su lista ya renderizada,
           solo se le agregó el botón "Tendencia".
       Reutiliza ExcelParser.parseStandalone() (ya existente,
       creado para la Comparativa Histórica) para leer el archivo
       elegido de forma AISLADA — nunca sustituye DataStore.rawMain,
       el reporte activo del dashboard queda intacto.
──────────────────────────────────────────────────────────── */
/* Tarjetas donde un aumento (+) es una MALA noticia y una disminución (-)
   es BUENA: aquí se invierte solo el COLOR (verde/rojo) de la tendencia,
   nunca el cálculo matemático ni la flecha ▲/▼, que siguen reflejando la
   dirección real del dato. Usa las claves de KPI_DETAIL_MAP. */
const KPI_INVERTED_POLARITY = ['kpiCelulasNO', 'kpiServicioNO', 'kpiAmbosNO'];

const TrendEngine = {

  _loading: false,

  _el(id) { return document.getElementById(id); },

  /* ── Fija un archivo del historial como línea base de tendencia ── */
  async setBaseline(file) {
    if (this._loading) return;
    this._loading = true;

    const overlay = this._el('historyFileLoading');
    const label   = this._el('historyFileLoadingName');
    if (overlay) overlay.classList.remove('d-none');
    if (label)   label.textContent = `Estableciendo "${file.name}" como línea base de tendencia...`;

    try {
      const url = file.download_url;
      if (!url) throw new Error('El archivo no tiene URL de descarga disponible.');

      const res = await fetch(url);
      if (!res.ok) throw new Error(`No se pudo descargar el archivo (${res.status}).`);

      const buffer   = await res.arrayBuffer();
      const workbook = await WorkerEngine.parseWorkbookAsync(buffer); // offload a Web Worker (no bloquea la UI)

      /* Parseo AISLADO — no toca el reporte actualmente cargado.
         Ya viene filtrado por AccessManager.applyFilter() (obligatorio
         y fail-closed, ver ExcelParser.parseStandalone). */
      const historico = ExcelParser.parseStandalone(workbook);

      /* Única propiedad nueva en DataStore — no se modifica ninguna
         función existente del motor, solo se le agrega este dato. */
      DataStore.comparisonBaseline = { fileName: file.name, rawMain: historico.rawMain };

      AuthEngine._toast(`Tendencia activa: comparando contra "${file.name}" ✓`, 'success');

      /* Recalcula con el flujo normal — refresh() ya llama a
         TrendEngine.render() como parte de su secuencia habitual */
      UIController.refresh();

      const modalEl = this._el('modalHistorial');
      if (modalEl) {
        const bsModal = bootstrap.Modal.getInstance(modalEl);
        if (bsModal) bsModal.hide();
      }

    } catch (err) {
      alert(`No se pudo establecer la línea base de tendencia:\n${err.message}`);
    } finally {
      this._loading = false;
      if (overlay) overlay.classList.add('d-none');
    }
  },

  /* ── Desactiva la comparativa de tendencia ── */
  clearBaseline() {
    DataStore.comparisonBaseline = null;
    SparklineEngine.clearAll();
    UIController.refresh();
    AuthEngine._toast('Comparativa de tendencia desactivada', 'info');
  },

  /**
   * Pinta flecha + % en cada tarjeta declarada en KPI_DETAIL_MAP y,
   * si Chart.js está disponible, su mini-sparkline de 2 puntos.
   * Se invoca desde UIController.refresh() DESPUÉS de updateKPICards(),
   * así nunca compite por el mismo DOM ni altera los valores absolutos
   * que ya pinta ese método.
   *
   * @param {Object} kpisActual - Resultado de KPIEngine.compute() sobre
   *   el reporte/filtros activos ahora mismo (mismo objeto que ya usa
   *   updateKPICards() y ChartEngine.renderAll()).
   */
  render(kpisActual) {
    const baseline = DataStore.comparisonBaseline;

    if (!baseline) {
      Object.keys(KPI_DETAIL_MAP).forEach(valueId => {
        this._el(`${valueId}TrendRow`)?.classList.add('d-none');
      });
      SparklineEngine.clearAll();
      this._toggleBanner(false);
      return;
    }

    /* Se le aplican los MISMOS filtros de UI activos ahora mismo
       (grupo, estado, célula, servicio, nuevo) para que la comparación
       sea simétrica: "[Grupo X] Actual" vs. "[Grupo X] Anterior".
       Sin esto, con un grupo filtrado la línea base se comparaba contra
       el total general sin filtrar. */
    const baselineFiltrado = DataStore.applyFilters(baseline.rawMain);
    const kpisAnterior = KPIEngine.compute(baselineFiltrado);
    const deltas = KPIEngine.computeDelta(kpisActual, kpisAnterior);

    Object.entries(KPI_DETAIL_MAP).forEach(([valueId, { metric }]) => {
      const rowEl   = this._el(`${valueId}TrendRow`);
      const trendEl = this._el(`${valueId}Trend`);
      const d = deltas[metric];
      if (!rowEl || !trendEl || !d) return;

      rowEl.classList.remove('d-none');
      trendEl.classList.remove('kpi-trend-up', 'kpi-trend-down', 'kpi-trend-flat');

      const icon = d.direction === 'up' ? '▲' : d.direction === 'down' ? '▼' : '►';
      const anteriorVal = kpisAnterior[metric] ?? 0;

      /* Polaridad invertida SOLO para color: el icono y el texto del
         porcentaje (d.text) siguen mostrando la dirección matemática
         real; únicamente cambia qué color (verde/rojo) se le asigna. */
      const isInverted = KPI_INVERTED_POLARITY.includes(valueId);
      const colorDirection = isInverted
        ? (d.direction === 'up' ? 'down' : d.direction === 'down' ? 'up' : 'flat')
        : d.direction;

      trendEl.classList.add(`kpi-trend-${colorDirection}`);
      trendEl.textContent = `${icon} ${d.text}`;
      trendEl.title = `Anterior: ${anteriorVal} — Línea base: ${baseline.fileName}`;

      const prevEl = this._el(`${valueId}TrendPrev`);
      if (prevEl) prevEl.textContent = `antes: ${anteriorVal}`;

      SparklineEngine.render(valueId, anteriorVal, kpisActual[metric] ?? 0, colorDirection);
    });

    this._toggleBanner(true, baseline.fileName);
  },

  _toggleBanner(show, fileName) {
    const banner = this._el('trendBaselineBanner');
    if (!banner) return;
    banner.classList.toggle('d-none', !show);
    if (show) {
      const nameEl = this._el('trendBaselineFileName');
      if (nameEl) nameEl.textContent = fileName;
    }
  },

  init() {
    this._el('btnClearTrendBaseline')?.addEventListener('click', () => this.clearBaseline());
  },
};


/* ────────────────────────────────────────────────────────────
   10.5 TREND VIEW ENGINE — botón "Cargar Tendencia" de la Top Bar
   ────────────────────────────────────────────────────────────
   Diferencia clave con TrendEngine.setBaseline(): ese método trae un
   archivo del HISTORIAL (descarga remota desde GitHub) y lo usa solo
   para calcular flechas/% de comparación (DataStore.comparisonBaseline
   + TrendEngine.render()), sin cambiar qué se ve en las tablas/gráficos
   principales. TrendViewEngine, en cambio, deja elegir un Excel LOCAL
   (mismo flujo de carga que "Cargar Excel": un <input type="file">) y
   PROYECTA esos datos como si fueran el reporte activo en TableEngine,
   ChartEngine y KPIEngine — sin tocar rawMain/rawExcluidos.

   Cómo queda no-destructivo:
   - Reutiliza ExcelParser.parseStandalone() (el mismo parseo aislado
     y con RBAC obligatorio que ya usa TrendEngine.setBaseline()), NO
     ExcelParser.parse() (ese sí sobrescribiría DataStore.rawMain).
   - Guarda el resultado en DataStore.comparisonBaseline (la MISMA
     propiedad que usa TrendEngine) y activa DataStore.viewingTrend.
   - DataStore.getActiveMain() — el único punto que leen
     TableEngine/ChartEngine/KPIEngine — devuelve ese dataset mientras
     viewingTrend esté activo, y vuelve a rawMain/rawExcluidos en
     cuanto exitTrendView() lo desactiva. Ningún motor de UI necesita
     saber que existe una "vista de tendencia": para ellos es un
     array de registros como cualquier otro.
──────────────────────────────────────────────────────────── */
const TrendViewEngine = {

  _loading: false,

  /* Carga un Excel local y lo proyecta en la vista principal como
     Tendencia. No modifica DataStore.rawMain/rawExcluidos. */
  async loadLocalFile(file) {
    if (this._loading) return;
    this._loading = true;
    UIController.showLoading(true);

    try {
      const buffer   = await file.arrayBuffer();
      const workbook = await WorkerEngine.parseWorkbookAsync(buffer); // mismo Web Worker que "Cargar Excel"
      const parsed   = ExcelParser.parseStandalone(workbook); // parseo aislado, RBAC obligatorio y fail-closed

      DataStore.comparisonBaseline = { fileName: file.name, rawMain: parsed.rawMain };
      DataStore.viewingTrend = true;

      FilterEngine.reset();     // limpia filtros de la vista anterior (grupo/estado/etc.)
      UIController.refresh();
      UIController.showDashboard();

      const titleEl = document.getElementById('reportTitle');
      if (titleEl) titleEl.textContent = `Tendencia — ${file.name}`;
      this._toggleBanner(true, file.name);

      AuthEngine._toast(`Viendo tendencia de "${file.name}" ✓`, 'success');

      /* Notificación de auditoría por Telegram (fire-and-forget) */
      const auditUser = AuditEngine.getUser();
      if (auditUser) {
        AuditEngine.notify({ action: 'cargar_tendencia_local', user: auditUser, fileName: file.name });
      }
    } catch (err) {
      console.error('[TrendViewEngine] Error al cargar la tendencia:', err);
      alert(`No se pudo cargar la tendencia:\n${err.message}`);
    } finally {
      this._loading = false;
      UIController.showLoading(false);
    }
  },

  /* Vuelve a la vista normal (el Excel general ya cargado en
     DataStore.rawMain). No borra comparisonBaseline: si el usuario
     tenía además una comparativa de flechas/% activa desde el
     Historial, sigue disponible tal cual estaba. */
  exitTrendView() {
    DataStore.viewingTrend = false;
    UIController.refresh();

    const titleEl = document.getElementById('reportTitle');
    if (titleEl) titleEl.textContent = DataStore.reportTitle || DataStore.fileName;
    this._toggleBanner(false);

    AuthEngine._toast('Volviste a la vista normal', 'info');
  },

  _toggleBanner(show, fileName) {
    const banner = document.getElementById('trendViewBanner');
    if (!banner) return;
    banner.classList.toggle('d-none', !show);
    if (show) {
      const nameEl = document.getElementById('trendViewFileName');
      if (nameEl) nameEl.textContent = fileName;
    }
  },

  init() {
    document.getElementById('fileInputTrend')?.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) this.loadLocalFile(file);
      e.target.value = ''; // permite volver a elegir el mismo archivo después
    });
    document.getElementById('btnExitTrendView')?.addEventListener('click', () => this.exitTrendView());
  },
};


/* ────────────────────────────────────────────────────────────
   11. AUTH ENGINE — Gestión del Personal Access Token (PAT)
       Almacena el token en localStorage.
       Alerta de caducidad a los 350 días (tokens duran 1 año).
──────────────────────────────────────────────────────────── */
const AuthEngine = {

  STORAGE_KEY_TOKEN:    'iglesia_gh_token',
  STORAGE_KEY_SAVED_AT: 'iglesia_gh_token_saved_at',
  EXPIRY_WARN_DAYS:     350,   // Aviso cuando restan ~15 días para expirar

  /* ── Getter / Setter ── */
  getToken()  { return localStorage.getItem(this.STORAGE_KEY_TOKEN) || ''; },
  getSavedAt(){ return parseInt(localStorage.getItem(this.STORAGE_KEY_SAVED_AT) || '0', 10); },

  saveToken(token) {
    localStorage.setItem(this.STORAGE_KEY_TOKEN,    token.trim());
    localStorage.setItem(this.STORAGE_KEY_SAVED_AT, Date.now().toString());
  },

  clearToken() {
    localStorage.removeItem(this.STORAGE_KEY_TOKEN);
    localStorage.removeItem(this.STORAGE_KEY_SAVED_AT);
  },

  /* Días transcurridos desde que se guardó el token */
  daysSinceSaved() {
    const saved = this.getSavedAt();
    if (!saved) return 0;
    return Math.floor((Date.now() - saved) / (1000 * 60 * 60 * 24));
  },

  /* Comprueba si el token está próximo a caducar */
  isNearExpiry() {
    return this.getToken() && this.daysSinceSaved() >= this.EXPIRY_WARN_DAYS;
  },

  /* ── Alerta de caducidad en el banner del offcanvas ── */
  checkExpiry() {
    const banner = document.getElementById('authExpiryBanner');
    if (!banner) return;
    if (this.isNearExpiry()) {
      const days = this.daysSinceSaved();
      const remaining = 365 - days;
      document.getElementById('authExpiryDays').textContent =
        remaining <= 0 ? 'ya ha caducado' : `caduca en ~${remaining} día${remaining !== 1 ? 's' : ''}`;
      banner.classList.remove('d-none');
    } else {
      banner.classList.add('d-none');
    }
  },

  /* ── Inicializa el modal de configuración ── */
  initModal() {
    /* Poblar input al abrir */
    const modalEl = document.getElementById('authModal');
    if (!modalEl) return;
    this._modalRef = new bootstrap.Modal(modalEl);

    /* Botón topbar */
    document.getElementById('btnAuthConfig')?.addEventListener('click', () => {
      document.getElementById('authTokenInput').value = this.getToken();
      this._updateModalStatus();
      this._modalRef.show();
    });

    /* Guardar */
    document.getElementById('btnAuthSave')?.addEventListener('click', () => {
      const val = document.getElementById('authTokenInput')?.value.trim() || '';
      if (!val) { this._setModalError('El token no puede estar vacío.'); return; }
      this.saveToken(val);
      this._setModalError('');
      this._updateModalStatus();
      this.checkExpiry();
      /* Muestra confirmación y cierra */
      this._toast('Token guardado correctamente ✓', 'success');
      setTimeout(() => this._modalRef.hide(), 800);
    });

    /* Borrar */
    document.getElementById('btnAuthClear')?.addEventListener('click', () => {
      this.clearToken();
      document.getElementById('authTokenInput').value = '';
      this._updateModalStatus();
      this.checkExpiry();
      CloudEngine.disableUploadBtn();
    });

    /* Toggle visibilidad del campo */
    document.getElementById('btnAuthToggle')?.addEventListener('click', () => {
      const input = document.getElementById('authTokenInput');
      const icon  = document.getElementById('authToggleIcon');
      if (!input) return;
      const isPass = input.type === 'password';
      input.type = isPass ? 'text' : 'password';
      icon.className = isPass ? 'bi bi-eye-slash' : 'bi bi-eye';
    });

    /* Comprobar caducidad en cada apertura */
    modalEl.addEventListener('show.bs.modal', () => this._updateModalStatus());
  },

  _updateModalStatus() {
    const token   = this.getToken();
    const days    = this.daysSinceSaved();
    const statusEl = document.getElementById('authStatus');
    if (!statusEl) return;
    if (!token) {
      statusEl.className = 'auth-status auth-status-none';
      statusEl.innerHTML = '<i class="bi bi-shield-x me-1"></i>Sin token configurado';
    } else if (this.isNearExpiry()) {
      statusEl.className = 'auth-status auth-status-warn';
      statusEl.innerHTML = `<i class="bi bi-exclamation-triangle me-1"></i>Token guardado — caduca pronto (${days} días)`;
    } else {
      statusEl.className = 'auth-status auth-status-ok';
      statusEl.innerHTML = `<i class="bi bi-shield-check me-1"></i>Token activo — ${days} día${days !== 1 ? 's' : ''} guardado`;
    }
  },

  _setModalError(msg) {
    const el = document.getElementById('authModalError');
    if (!el) return;
    el.textContent = msg;
    el.classList.toggle('d-none', !msg);
  },

  _toast(msg, type = 'info') {
    const icons = { success: '✅', error: '❌', info: 'ℹ️' };
    const t = document.createElement('div');
    t.className = `sync-toast ${type}`;
    t.innerHTML = `<span>${icons[type]||'•'}</span><span>${msg}</span>`;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 3500);
  },

  init() {
    this.initModal();
    this.checkExpiry();
    this.bindExpiryLink();
  },

  /* Enlace "Renovar ahora" dentro del banner de caducidad (vive dentro
     de #modalHistorial, que ahora se inyecta dinámicamente al DOM —
     ver LazyModals.js). Se llama desde init() (no-op si el modal aún
     no fue inyectado, gracias al optional chaining) Y desde
     LazyModals justo después de inyectar #modalHistorial, para que
     el enlace quede funcionando la primera vez que el usuario abre
     el Historial. Es idempotente: si ya se enlazó, no vuelve a hacerlo. */
  _expiryLinkBound: false,
  bindExpiryLink() {
    if (this._expiryLinkBound) return;
    const link = document.getElementById('btnExpiryOpenAuth');
    if (!link) return;

    link.addEventListener('click', e => {
      e.preventDefault();

      /* Cierra el modal de Historial antes de abrir el de configuración del token */
      const historialEl = document.getElementById('modalHistorial');
      if (historialEl) {
        const bsHistorial = bootstrap.Modal.getInstance(historialEl);
        if (bsHistorial) bsHistorial.hide();
      }

      document.getElementById('authTokenInput').value = this.getToken();
      this._updateModalStatus();
      this._modalRef?.show();
    });
    this._expiryLinkBound = true;
  },
};

/* Expone AuthEngine en window: LazyModals.js comprueba
   `window.AuthEngine` al inyectar #modalHistorial por primera vez,
   para enlazar el link "Renovar ahora" y revisar la caducidad del
   token (AuthEngine.bindExpiryLink / checkExpiry). Sin esta línea
   esa comprobación también era siempre `false` por el mismo motivo
   que HistoryEngine (ver nota arriba). */
window.AuthEngine = AuthEngine;


/* ────────────────────────────────────────────────────────────
   12. CLOUD ENGINE — Subida de archivos a GitHub via API PUT
──────────────────────────────────────────────────────────── */
const CloudEngine = {

  GITHUB_UPLOAD_BASE: 'https://api.github.com/repos/alexchouriors/M-tricas-REPORTE-DE-ASISTENCIAS-NUEVA/contents/REPORTES/',

  /* Convierte ArrayBuffer a string Base64 */
  _bufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
  },

  /* Habilita el botón de guardar en la nube con el nombre del archivo */
  enableUploadBtn(fileName) {
    const btn = document.getElementById('btnCloudSave');
    if (!btn) return;
    btn.classList.remove('d-none');
    btn.disabled = false;
    btn.dataset.fileName = fileName;
    btn.title = `Guardar "${fileName}" en GitHub`;
  },

  disableUploadBtn() {
    const btn = document.getElementById('btnCloudSave');
    if (!btn) return;
    btn.classList.add('d-none');
    btn.disabled = true;
    btn.dataset.fileName = '';
  },

  /* ── Modal de confirmación de nombre ── */
  _openUploadModal(suggestedName) {
    const input = document.getElementById('cloudFileNameInput');
    if (input) input.value = suggestedName;
    const modalEl = document.getElementById('cloudModal');
    if (modalEl) {
      this._cloudModalRef = this._cloudModalRef || new bootstrap.Modal(modalEl);
      document.getElementById('cloudModalError')?.classList.add('d-none');
      this._cloudModalRef.show();
    }
  },

  /* ── Estado de loading en el botón del modal ── */
  _setUploading(uploading) {
    const btn = document.getElementById('btnCloudConfirm');
    if (!btn) return;
    btn.disabled = uploading;
    btn.innerHTML = uploading
      ? '<span class="spinner-border spinner-border-sm me-2"></span>Subiendo…'
      : '<i class="bi bi-cloud-arrow-up me-2"></i>Subir';
  },

  _setModalError(msg) {
    const el = document.getElementById('cloudModalError');
    if (!el) return;
    el.textContent = msg;
    el.classList.toggle('d-none', !msg);
  },

  /* ── Petición PUT a la API de GitHub ── */
  async uploadFile(fileName) {
    const token = AuthEngine.getToken();
    if (!token) {
      this._setModalError('No hay token configurado. Ve a Configuración → Token GitHub.');
      return;
    }

    const buffer = DataStore.rawBuffer;
    if (!buffer) {
      this._setModalError('No hay archivo cargado en el dashboard.');
      return;
    }

    /* Asegura extensión válida */
    const safeName = fileName.trim() || DataStore.fileName;
    if (!safeName) { this._setModalError('El nombre del archivo es obligatorio.'); return; }

    this._setUploading(true);
    this._setModalError('');

    try {
      const base64Content = this._bufferToBase64(buffer);
      const apiUrl = this.GITHUB_UPLOAD_BASE + encodeURIComponent(safeName);

      /* Primero comprobamos si el archivo ya existe (para obtener su SHA y hacer update) */
      let sha = null;
      const checkRes = await fetch(apiUrl, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/vnd.github.v3+json',
        },
      });
      if (checkRes.ok) {
        const existing = await checkRes.json();
        sha = existing.sha;
      }

      const body = {
        message: `Dashboard: ${sha ? 'Actualiza' : 'Sube'} reporte ${safeName}`,
        content: base64Content,
      };
      if (sha) body.sha = sha;   // Requerido para actualizar un archivo existente

      const putRes = await fetch(apiUrl, {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept':        'application/vnd.github.v3+json',
          'Content-Type':  'application/json',
        },
        body: JSON.stringify(body),
      });

      if (!putRes.ok) {
        const errData = await putRes.json().catch(() => ({}));
        const detail  = errData.message || putRes.statusText;
        if (putRes.status === 401) throw new Error('Token inválido o sin permisos (401). Verifica tu PAT.');
        if (putRes.status === 422) throw new Error('Error de validación (422): ' + detail);
        throw new Error(`Error ${putRes.status}: ${detail}`);
      }

      /* Éxito */
      this._cloudModalRef?.hide();
      AuthEngine._toast(`"${safeName}" subido exitosamente a GitHub ✓`, 'success');

      /* Refresca la lista del Historial */
      HistoryEngine._files = [];
      const listEl = document.getElementById('listaReportesContainer');
      if (listEl) listEl.innerHTML = '';
      /* Si el modal de Historial está abierto, re-fetcha; si no, en la próxima apertura lo hará */
      const modalHistorialEl = document.getElementById('modalHistorial');
      if (modalHistorialEl && modalHistorialEl.classList.contains('show')) HistoryEngine.fetchFileList();

    } catch (err) {
      this._setModalError(err.message);
    } finally {
      this._setUploading(false);
    }
  },

  /* ── Inicializa eventos ── */
  init() {
    /* Botón topbar "Guardar en la Nube" → abre modal */
    document.getElementById('btnCloudSave')?.addEventListener('click', () => {
      this._openUploadModal(DataStore.fileName || 'reporte.xlsx');
    });

    /* Confirmar subida desde el modal */
    document.getElementById('btnCloudConfirm')?.addEventListener('click', () => {
      const name = document.getElementById('cloudFileNameInput')?.value.trim();
      if (!name) { this._setModalError('Ingresa un nombre para el archivo.'); return; }
      this.uploadFile(name);
    });
  },
};


/* ────────────────────────────────────────────────────────────
   12.5 SAVE ENGINE — Botón "Guardar" del menú lateral (offcanvas)
       Se habilita solo cuando hay un Excel cargado localmente.
       Flujo: pide nuevo nombre → pide responsable → sube a
       GitHub (REPORTES/) → notifica por Telegram.
──────────────────────────────────────────────────────────── */
/* ────────────────────────────────────────────────────────────
   11.5 GITHUB QUEUE — Prevención de Rate Limits
   Cola de peticiones con espaciado mínimo entre llamadas a la API
   de GitHub. No modifica NINGUNA petición existente por sí sola:
   los módulos que quieran protegerse de ráfagas (guardado manual,
   futuros guardados automáticos, etc.) simplemente envuelven su
   función de red con GitHubQueue.enqueue(fn).

   Por qué hace falta: la API REST de GitHub aplica rate-limit por
   token/IP. Si el guardado manual coincide con un ciclo de
   auto-sync de Google Sheets (u otra llamada concurrente a GitHub),
   varias peticiones podrían dispararse casi al mismo tiempo. Esta
   cola las serializa (una a la vez) y garantiza un espaciado mínimo
   entre ellas, sin bloquear la UI: el llamador sigue recibiendo una
   Promise normal, solo que puede tardar unos milisegundos más en
   iniciarse si hay otra petición reciente en curso. */
const GitHubQueue = {
  MIN_SPACING_MS: 800, // espaciado mínimo entre peticiones consecutivas

  _queue: [],
  _processing: false,
  _lastRequestAt: 0,

  /* Encola `taskFn` (una función que devuelve una Promise, típicamente
     una llamada fetch) y devuelve una Promise que se resuelve/rechaza
     con el resultado de esa tarea cuando finalmente se ejecuta. */
  enqueue(taskFn) {
    return new Promise((resolve, reject) => {
      this._queue.push({ taskFn, resolve, reject });
      this._processNext();
    });
  },

  async _processNext() {
    if (this._processing) return; // ya hay un ciclo de procesamiento activo
    this._processing = true;

    while (this._queue.length > 0) {
      const { taskFn, resolve, reject } = this._queue.shift();

      // Respeta el espaciado mínimo desde la última petición disparada
      const wait = this.MIN_SPACING_MS - (Date.now() - this._lastRequestAt);
      if (wait > 0) await new Promise(r => setTimeout(r, wait));

      this._lastRequestAt = Date.now();
      try {
        const result = await taskFn();
        resolve(result);
      } catch (err) {
        reject(err);
      }
    }

    this._processing = false;
  },
};


const SaveEngine = {

  GITHUB_UPLOAD_BASE: 'https://api.github.com/repos/alexchouriors/M-tricas-REPORTE-DE-ASISTENCIAS-NUEVA/contents/REPORTES/',

  _el(id) { return document.getElementById(id); },

  /* ── Habilita el botón al cargar un Excel local ── */
  enable(fileName) {
    const btn  = this._el('btnMenuGuardar');
    const icon = this._el('btnMenuGuardarIcon');
    const text = this._el('btnMenuGuardarText');
    if (!btn) return;

    btn.disabled = false;
    btn.classList.add('is-enabled');
    btn.dataset.fileName = fileName || '';
    btn.title = `Guardar "${fileName || ''}" en GitHub`;

    if (icon) icon.className = 'bi bi-save2-fill';
    if (text) text.textContent = 'Guardar';
  },

  /* ── Vuelve al estado bloqueado por defecto ── */
  disable() {
    const btn  = this._el('btnMenuGuardar');
    const icon = this._el('btnMenuGuardarIcon');
    const text = this._el('btnMenuGuardarText');
    if (!btn) return;

    btn.disabled = true;
    btn.classList.remove('is-enabled');
    btn.dataset.fileName = '';
    btn.title = '';

    if (icon) icon.className = 'bi bi-lock-fill';
    if (text) text.textContent = 'Guardar';
  },

  /* Convierte ArrayBuffer a string Base64 (mismo criterio que CloudEngine) */
  _bufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
  },

  /* ── Sube el archivo en memoria a GitHub dentro de REPORTES/ ── */
  async _uploadToGitHub(fileName) {
    const token = AuthEngine.getToken();
    if (!token) { alert('No hay token de GitHub configurado. Ve a Configuración → Token GitHub.'); return false; }

    const buffer = DataStore.rawBuffer;
    if (!buffer) { alert('No hay un archivo Excel cargado en el Dashboard.'); return false; }

    try {
      const base64Content = this._bufferToBase64(buffer);
      const apiUrl = this.GITHUB_UPLOAD_BASE + encodeURIComponent(fileName);

      /* Verifica si el archivo ya existe (para update con SHA) */
      let sha = null;
      const checkRes = await fetch(apiUrl, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/vnd.github.v3+json',
        },
      });
      if (checkRes.ok) {
        const existing = await checkRes.json();
        sha = existing.sha;
      }

      const body = {
        message: `Dashboard: ${sha ? 'Actualiza' : 'Guarda'} reporte ${fileName}`,
        content: base64Content,
      };
      if (sha) body.sha = sha;

      const putRes = await fetch(apiUrl, {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept':        'application/vnd.github.v3+json',
          'Content-Type':  'application/json',
        },
        body: JSON.stringify(body),
      });

      if (!putRes.ok) {
        const errData = await putRes.json().catch(() => ({}));
        const detail  = errData.message || putRes.statusText;
        if (putRes.status === 401) throw new Error('Token inválido o sin permisos (401). Verifica tu PAT.');
        if (putRes.status === 422) throw new Error('Error de validación (422): ' + detail);
        throw new Error(`Error ${putRes.status}: ${detail}`);
      }

      return true;
    } catch (err) {
      alert(`No se pudo guardar el archivo:\n${err.message}`);
      return false;
    }
  },

  /* ── Flujo completo del botón "Guardar" ── */
  async handleClick() {
    /* 1) Nuevo nombre para guardar el archivo */
    const nameInput = window.prompt('Ingrese el nuevo nombre para guardar el archivo');
    if (nameInput === null) return;
    const newName = nameInput.trim();
    if (newName === '') return;

    /* 2) Responsable de la acción: usuario con sesión activa */
    const user = AuditEngine.getUser();
    if (!user) return;

    /* 3) Asegura extensión válida reutilizando la del archivo original si falta */
    let safeName = newName;
    if (!/\.(xlsx|xlsm|xls)$/i.test(safeName)) {
      const origExt = (DataStore.fileName.match(/\.(xlsx|xlsm|xls)$/i) || [])[0] || '.xlsx';
      safeName += origExt;
    }

    const btn = this._el('btnMenuGuardar');
    const originalHTML = btn ? btn.innerHTML : '';

    /* ── FEEDBACK OPTIMISTA (Optimistic UI) ──
       Mostramos éxito de inmediato: el usuario no espera a la ida y
       vuelta con la API de GitHub para sentir que "ya guardó". La
       petición real viaja en segundo plano, encolada a través de
       GitHubQueue (ver arriba) para no chocar con otras llamadas a
       la API de GitHub que puedan estar en curso (p. ej. un ciclo de
       auto-sync de Google Sheets que dispare otra petición casi al
       mismo tiempo). Si la petición real termina fallando, se revierte
       el estado del botón y se avisa al usuario — no se le miente
       silenciosamente sobre el resultado final. */
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="bi bi-check2-circle me-2"></i>Guardado ✓';
    }
    AuthEngine._toast(`"${safeName}" guardado ✓ (sincronizando con GitHub…)`, 'success');

    GitHubQueue.enqueue(() => this._uploadToGitHub(safeName))
      .then((ok) => {
        if (btn) { btn.disabled = false; btn.innerHTML = originalHTML; }
        if (!ok) return; // _uploadToGitHub ya mostró el detalle del error

        /* 4) Notificación de auditoría por Telegram */
        AuditEngine.notify({ action: 'guardar', user, fileName: safeName });

        /* Invalida caché del Historial para reflejar el nuevo/actualizado archivo */
        HistoryEngine._files = [];
        const listEl = document.getElementById('listaReportesContainer');
        if (listEl) listEl.innerHTML = '';
        const modalHistorialEl = document.getElementById('modalHistorial');
        if (modalHistorialEl && modalHistorialEl.classList.contains('show')) HistoryEngine.fetchFileList();
      })
      .catch((err) => {
        console.error('[SaveEngine] Falló la confirmación real del guardado:', err);
        if (btn) { btn.disabled = false; btn.innerHTML = originalHTML; }
        AuthEngine._toast(`⚠️ No se pudo confirmar "${safeName}" en GitHub. Intenta de nuevo.`, 'error');
      });
  },

  init() {
    this._el('btnMenuGuardar')?.addEventListener('click', () => this.handleClick());
  },
};


/* ────────────────────────────────────────────────────────────
   13. DELETE ENGINE — Eliminación de archivos en GitHub via API DELETE
       Lista los archivos de /REPORTES y los elimina usando su SHA.
──────────────────────────────────────────────────────────── */
const DeleteEngine = {

  GITHUB_API: 'https://api.github.com/repos/alexchouriors/M-tricas-REPORTE-DE-ASISTENCIAS-NUEVA/contents/REPORTES',
  VALID_EXTS: ['.xlsx', '.xlsm', '.xls'],

  _files: [],

  _el(id) { return document.getElementById(id); },

  _ext(name) {
    const m = name.toLowerCase().match(/\.(xlsx|xlsm|xls)$/);
    return m ? '.' + m[1] : '';
  },

  _setState(state) {
    const map = { loading: 'deleteLoading', error: 'deleteError', empty: 'deleteEmpty' };
    Object.entries(map).forEach(([key, id]) => {
      const el = this._el(id);
      if (el) el.classList.toggle('d-none', key !== state);
    });
    const listEl = this._el('listaEliminarContainer');
    if (listEl) listEl.classList.toggle('d-none', state !== 'list');
  },

  _showError(msg) {
    const msgEl = this._el('deleteErrorMsg');
    if (msgEl) msgEl.textContent = msg;
    this._setState('error');
  },

  async fetchFileList() {
    this._setState('loading');
    try {
      const headers = { 'Accept': 'application/vnd.github.v3+json' };
      const token = AuthEngine.getToken();
      if (token) headers['Authorization'] = `Bearer ${token}`;

      const res = await fetch(this.GITHUB_API, { headers });
      if (!res.ok) {
        const msg = res.status === 404
          ? 'Repositorio o carpeta no encontrada (404).'
          : res.status === 403
            ? 'Límite de peticiones a la API de GitHub excedido. Intenta en unos minutos.'
            : `Error ${res.status}: ${res.statusText}`;
        throw new Error(msg);
      }

      const items = await res.json();
      this._files = items.filter(i => i.type === 'file' && this.VALID_EXTS.includes(this._ext(i.name)));

      if (this._files.length === 0) { this._setState('empty'); return; }
      this._renderList();
      this._setState('list');
    } catch (err) {
      this._showError(err.message || 'Error desconocido al contactar la API de GitHub.');
    }
  },

  _renderList() {
    const listEl = this._el('listaEliminarContainer');
    if (!listEl) return;
    listEl.innerHTML = '';
    this._files.forEach((file, idx) => {
      const item = document.createElement('div');
      item.className = 'list-group-item d-flex align-items-center justify-content-between flex-wrap gap-2';
      item.dataset.idx = idx;
      item.innerHTML = `
        <div class="d-flex align-items-center gap-2 text-truncate">
          <i class="bi bi-file-earmark-spreadsheet text-success fs-5"></i>
          <span class="text-truncate" title="${file.name}">${file.name}</span>
        </div>
        <button type="button" class="btn-delete-file" data-idx="${idx}" title="Eliminar ${file.name}">
          <i class="bi bi-trash-fill"></i>Eliminar
        </button>`;
      item.querySelector('.btn-delete-file').addEventListener('click', () => this._confirmDelete(file, item));
      listEl.appendChild(item);
    });
  },

  async _confirmDelete(file, itemEl) {
    const confirmed = window.confirm(`¿Estás seguro que quieres eliminar "${file.name}"?\n\nEsta acción es irreversible.`);
    if (!confirmed) return;

    /* Capa de seguridad/auditoría: usa el nombre del usuario con sesión activa */
    const user = AuditEngine.getUser();
    if (!user) return; // Sin sesión activa: se aborta la acción

    const token = AuthEngine.getToken();
    if (!token) { alert('No hay token de GitHub configurado. Ve a Configuración → Token GitHub.'); return; }

    const btn = itemEl.querySelector('.btn-delete-file');
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span>'; }

    try {
      const apiUrl = `${this.GITHUB_API}/${encodeURIComponent(file.name)}`;
      const res = await fetch(apiUrl, {
        method: 'DELETE',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept':        'application/vnd.github.v3+json',
          'Content-Type':  'application/json',
        },
        body: JSON.stringify({ message: `Dashboard: Elimina reporte ${file.name}`, sha: file.sha }),
      });

      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        const detail  = errData.message || res.statusText;
        if (res.status === 401) throw new Error('Token inválido o sin permisos (401).');
        if (res.status === 422) throw new Error('Error de validación (422): ' + detail);
        throw new Error(`Error ${res.status}: ${detail}`);
      }

      itemEl.style.transition = 'opacity .3s';
      itemEl.style.opacity = '0';
      setTimeout(() => itemEl.remove(), 300);
      this._files = this._files.filter(f => f.sha !== file.sha);
      if (this._files.length === 0) this._setState('empty');

      AuthEngine._toast(`"${file.name}" eliminado correctamente ✓`, 'success');

      /* Notificación de auditoría por Telegram (no bloquea la interfaz) */
      AuditEngine.notify({ action: 'eliminar', user, fileName: file.name });

      /* Invalida caché del HistoryEngine */
      HistoryEngine._files = [];
      const histListEl = document.getElementById('listaReportesContainer');
      if (histListEl) histListEl.innerHTML = '';

    } catch (err) {
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="bi bi-trash-fill"></i>Eliminar'; }
      alert(`No se pudo eliminar "${file.name}":\n${err.message}`);
    }
  },

  init() {
    const modalEl = document.getElementById('modalEliminar');
    if (!modalEl) return;
    modalEl.addEventListener('show.bs.modal', () => {
      this._files = [];
      const listEl = this._el('listaEliminarContainer');
      if (listEl) listEl.innerHTML = '';
      this.fetchFileList();
    });
    this._el('btnDeleteRetry')?.addEventListener('click', () => this.fetchFileList());
  },
};


/* ────────────────────────────────────────────────────────────
   13.5 DB DEFAULT ENGINE — Botón "Base de Datos (Beta)"
       Flujo: pide token de GitHub por sesión (nunca leído/guardado
       en caché ni localStorage) → abre un modal clon de "Eliminar"
       que lista los .xlsx del repositorio → "Establecer como
       predeterminado" hace PUT a config.json en la raíz del repo
       con { archivo_predeterminado: "<nombre>.xlsx" }.
──────────────────────────────────────────────────────────── */
const DbDefaultEngine = {

  GITHUB_CONTENTS_BASE: 'https://api.github.com/repos/alexchouriors/M-tricas-REPORTE-DE-ASISTENCIAS-NUEVA/contents/',
  VALID_EXTS: ['.xlsx', '.xlsm', '.xls'],

  _files: [],
  _sessionToken: '',   // Solo en memoria durante la sesión del modal; nunca persistido
  _currentDefault: '', // Nombre del archivo actualmente marcado como predeterminado (config.json)
  _currentTrend: '',   // Nombre del archivo actualmente marcado como tendencia (config.json)

  _el(id) { return document.getElementById(id); },

  _ext(name) {
    const m = name.toLowerCase().match(/\.(xlsx|xlsm|xls)$/);
    return m ? '.' + m[1] : '';
  },

  /* ── Paso 1: modal de token efímero ── */
  _openTokenModal() {
    this._sessionToken = '';
    const input = this._el('dbTokenInput');
    if (input) input.value = '';
    this._setTokenError('');
    const modalEl = this._el('modalDbToken');
    if (!modalEl) return;
    this._tokenModalRef = this._tokenModalRef || new bootstrap.Modal(modalEl);
    this._tokenModalRef.show();
  },

  _setTokenError(msg) {
    const el = this._el('dbTokenModalError');
    if (!el) return;
    el.textContent = msg;
    el.classList.toggle('d-none', !msg);
  },

  _confirmToken() {
    const val = this._el('dbTokenInput')?.value.trim() || '';
    if (!val) { this._setTokenError('El token no puede estar vacío.'); return; }

    /* Guardado SOLO en memoria (variable de instancia); jamás en
       localStorage/sessionStorage/caché, y se pide de nuevo en
       cada apertura del botón "Base de Datos (Beta)". */
    this._sessionToken = val;
    this._setTokenError('');
    this._tokenModalRef?.hide();

    /* Abre el modal de configuración tras el cierre del de token */
    setTimeout(() => this._openConfigModal(), 300);
  },

  /* ── Paso 2: modal clon de "Eliminar" con la lista de archivos ── */
  _openConfigModal() {
    const modalEl = this._el('modalDbDefault');
    if (!modalEl) return;
    this._configModalRef = this._configModalRef || new bootstrap.Modal(modalEl);
    this._files = [];
    const listEl = this._el('listaDbDefaultContainer');
    if (listEl) listEl.innerHTML = '';
    this._configModalRef.show();
    this.fetchFileList();
  },

  _setState(state) {
    const map = { loading: 'dbDefaultLoading', error: 'dbDefaultError', empty: 'dbDefaultEmpty' };
    Object.entries(map).forEach(([key, id]) => {
      const el = this._el(id);
      if (el) el.classList.toggle('d-none', key !== state);
    });
    const listEl = this._el('listaDbDefaultContainer');
    if (listEl) listEl.classList.toggle('d-none', state !== 'list');
  },

  _showError(msg) {
    const msgEl = this._el('dbDefaultErrorMsg');
    if (msgEl) msgEl.textContent = msg;
    this._setState('error');
  },

  /**
   * Verifica que el usuario en sesión tenga rol MAESTRO antes de
   * permitir cualquier escritura de tendencia/predeterminado. No basta
   * con ocultar el botón en la UI: esta es la validación real,
   * fail-closed, a nivel de lógica — igual que exige AccessManager
   * para los datos.
   *
   * @returns {boolean}
   */
  _isMaster() {
    if (typeof UsuarioRules === 'undefined') {
      console.error('[DbDefaultEngine] UsuarioRules no está cargado — se deniega por seguridad (fail-closed).');
      return false;
    }
    const usuarioActual = AuditEngine.getUser();
    return UsuarioRules._resolveRole(usuarioActual) === 'MAESTRO';
  },

  /**
   * Se llama al iniciar sesión (ver SessionEngine._confirmLogin), DESPUÉS
   * de que el archivo predeterminado principal ya terminó de cargar.
   * Lee `archivo_tendencia` desde config.json en GitHub (misma fuente
   * global que usa AutoLoadEngine para el predeterminado — YA NO
   * localStorage, así que funciona igual en cualquier dispositivo para
   * cualquier usuario, sin necesidad de configurarlo por su cuenta).
   * Si existe, resuelve su download_url y dispara TrendEngine.setBaseline()
   * (el mismo motor que ya usa el botón "Tendencia" del Historial) para
   * que el dashboard aparezca ya cruzado contra esa línea base, sin
   * clics adicionales.
   */
  async applyStoredTrendIfAny() {
    try {
      /* Igual que HistoryEngine.fetchFileList(): se agrega el token si
         está disponible para usar el límite de 5000 peticiones/hora de
         la API autenticada en vez del límite de 60/hora sin autenticar
         (compartido por IP entre todos los que usan el dashboard) — sin
         esto, estas lecturas podían fallar en silencio con 403 "rate
         limit exceeded" apenas hubiera uso simultáneo, sin mostrar
         ningún error visible al usuario. */
      const token = AuthEngine.getToken();
      const authHeader = token ? { 'Authorization': `Bearer ${token}` } : {};

      const cfgRes = await fetch(this.GITHUB_CONTENTS_BASE + 'config.json', {
        cache: 'no-store',
        headers: { 'Accept': 'application/vnd.github.v3.raw', ...authHeader },
      });
      if (!cfgRes.ok) return;

      const config   = await cfgRes.json().catch(() => null);
      const fileName = config?.archivo_tendencia;
      if (!fileName) return;
      if (typeof TrendEngine === 'undefined') return;

      /* Necesitamos el download_url del archivo — config.json solo
         guarda el nombre, igual que hace con archivo_predeterminado. */
      const fileRes = await fetch(
        this.GITHUB_CONTENTS_BASE + 'REPORTES/' + encodeURIComponent(fileName),
        { cache: 'no-store', headers: { 'Accept': 'application/vnd.github.v3+json', ...authHeader } }
      );
      if (!fileRes.ok) return;
      const fileMeta = await fileRes.json();
      if (!fileMeta.download_url) return;

      await TrendEngine.setBaseline({ name: fileName, download_url: fileMeta.download_url });
    } catch (err) {
      console.error('[DbDefaultEngine] No se pudo aplicar la tendencia predeterminada guardada:', err);
    }
  },

  /* ── Consulta config.json y muestra cuál es el archivo predeterminado
       Y la tendencia actuales en los banners informativos del modal ── */
  async _fetchCurrentDefault() {
    const label      = this._el('dbDefaultCurrentLabel');
    const trendLabel = this._el('dbDefaultCurrentTrendLabel');
    try {
      /* Usa el token de sesión del modal si ya se ingresó, y si no,
         cae al token persistido de AuthEngine — mismo motivo que en
         applyStoredTrendIfAny(): evitar el límite de 60 peticiones/hora
         sin autenticar. */
      const token = this._sessionToken || AuthEngine.getToken();
      const authHeader = token ? { 'Authorization': `Bearer ${token}` } : {};

      const res = await fetch(this.GITHUB_CONTENTS_BASE + 'config.json', {
        cache: 'no-store',
        headers: { 'Accept': 'application/vnd.github.v3.raw', ...authHeader },
      });
      if (!res.ok) {
        this._currentDefault = '';
        this._currentTrend   = '';
        if (label)      label.innerHTML      = '<i class="bi bi-star me-1"></i>Aún no hay ningún archivo predeterminado configurado.';
        if (trendLabel) trendLabel.innerHTML = '<i class="bi bi-graph-up me-1"></i>Aún no hay ninguna tendencia configurada.';
        return;
      }
      const config = await res.json().catch(() => null);
      this._currentDefault = config?.archivo_predeterminado || '';
      this._currentTrend   = config?.archivo_tendencia || '';

      if (label) {
        label.innerHTML = this._currentDefault
          ? `<i class="bi bi-star-fill me-1"></i>Predeterminado actual: <strong>${this._currentDefault}</strong>`
          : '<i class="bi bi-star me-1"></i>Aún no hay ningún archivo predeterminado configurado.';
      }
      if (trendLabel) {
        trendLabel.innerHTML = this._currentTrend
          ? `<i class="bi bi-graph-up-arrow me-1"></i>Tendencia actual: <strong>${this._currentTrend}</strong>`
          : '<i class="bi bi-graph-up me-1"></i>Aún no hay ninguna tendencia configurada.';
      }
    } catch (err) {
      this._currentDefault = '';
      this._currentTrend   = '';
      if (label)      label.innerHTML      = '<i class="bi bi-star me-1"></i>No se pudo consultar el predeterminado actual.';
      if (trendLabel) trendLabel.innerHTML = '<i class="bi bi-graph-up me-1"></i>No se pudo consultar la tendencia actual.';
    }
  },

  async fetchFileList() {
    this._setState('loading');
    /* Consulta en paralelo cuál es el predeterminado actual, para
       reflejarlo en el banner y marcar el item correspondiente */
    this._fetchCurrentDefault();
    try {
      const headers = { 'Accept': 'application/vnd.github.v3+json' };
      if (this._sessionToken) headers['Authorization'] = `Bearer ${this._sessionToken}`;

      const res = await fetch(this.GITHUB_CONTENTS_BASE + 'REPORTES', { headers });
      if (!res.ok) {
        const msg = res.status === 401
          ? 'Token inválido o sin permisos (401).'
          : res.status === 404
            ? 'Repositorio o carpeta no encontrada (404).'
            : res.status === 403
              ? 'Límite de peticiones a la API de GitHub excedido. Intenta en unos minutos.'
              : `Error ${res.status}: ${res.statusText}`;
        throw new Error(msg);
      }

      const items = await res.json();
      this._files = items.filter(i => i.type === 'file' && this.VALID_EXTS.includes(this._ext(i.name)));

      if (this._files.length === 0) { this._setState('empty'); return; }
      this._renderList();
      this._setState('list');
    } catch (err) {
      this._showError(err.message || 'Error desconocido al contactar la API de GitHub.');
    }
  },

  _renderList() {
    const listEl = this._el('listaDbDefaultContainer');
    if (!listEl) return;
    listEl.innerHTML = '';

    this._files.forEach((file, idx) => {
      const isCurrent = !!this._currentDefault && file.name === this._currentDefault;
      const isTrend   = !!this._currentTrend && file.name === this._currentTrend;

      const item = document.createElement('div');
      item.className = 'list-group-item d-flex align-items-center justify-content-between flex-wrap gap-2'
        + (isCurrent ? ' list-group-item-current-default' : '');
      item.dataset.idx = idx;
      item.innerHTML = `
        <div class="d-flex align-items-center gap-2 text-truncate">
          <i class="bi bi-file-earmark-spreadsheet text-success fs-5"></i>
          <span class="text-truncate" title="${file.name}">${file.name}</span>
          ${isCurrent ? '<span class="badge-current-default ms-1"><i class="bi bi-star-fill me-1"></i>Predeterminado</span>' : ''}
          ${isTrend ? '<span class="badge-current-trend ms-1"><i class="bi bi-graph-up-arrow me-1"></i>Tendencia</span>' : ''}
        </div>
        <div class="d-flex align-items-center gap-2 ms-auto db-default-actions">
          <button type="button" class="btn-set-default${isCurrent ? ' is-current' : ''}" data-idx="${idx}"
                  title="${isCurrent ? `"${file.name}" ya es el predeterminado` : `Establecer ${file.name} como predeterminado`}"
                  ${isCurrent ? 'disabled' : ''}>
            <i class="bi bi-star-fill"></i>${isCurrent ? 'Ya es el predeterminado' : 'Establecer como predeterminado'}
          </button>
          <button type="button" class="btn-set-trend${isTrend ? ' is-current' : ''}" data-idx="${idx}"
                  title="${isTrend ? `"${file.name}" ya es la tendencia activa` : `Establecer ${file.name} como tendencia`}"
                  ${isTrend ? 'disabled' : ''}>
            <i class="bi bi-graph-up-arrow"></i>${isTrend ? 'Tendencia activa' : 'Establecer tendencia'}
          </button>
        </div>`;

      if (!isCurrent) {
        item.querySelector('.btn-set-default').addEventListener('click', () => this._setDefault(file, item));
      }
      if (!isTrend) {
        item.querySelector('.btn-set-trend').addEventListener('click', () => this._setTrend(file, item));
      }
      listEl.appendChild(item);
    });
  },

  /* ── Guarda la tendencia predeterminada en config.json (GitHub) —
       misma fuente global que usa el archivo predeterminado — y la
       aplica de inmediato al dashboard, reutilizando
       TrendEngine.setBaseline() (el mismo motor que ya usa el botón
       "Tendencia" del Historial). SOLO usuarios con rol MAESTRO
       pueden ejecutar esta escritura (ver _isMaster()). ── */
  async _setTrend(file, itemEl) {
    /* Validación de permisos a nivel de lógica — fail-closed, no
       depende solo de que el botón esté oculto en la UI. */
    if (!this._isMaster()) {
      alert('Solo un usuario con rol MAESTRO puede establecer la tendencia predeterminada.');
      return;
    }

    const token = this._sessionToken;
    if (!token) { alert('Sesión de token expirada. Vuelve a abrir "Base de Datos (Beta)".'); return; }

    const btn = itemEl.querySelector('.btn-set-trend');
    const originalHTML = btn ? btn.innerHTML : '';
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span>'; }

    try {
      const apiUrl = this.GITHUB_CONTENTS_BASE + 'config.json';

      /* Lee config.json actual para preservar archivo_predeterminado
         (y cualquier otra clave futura) y obtener el sha para el PUT.
         cache:'no-store' evita servir una respuesta 404 cacheada de
         cuando el archivo aún no existía (ver nota en _setDefault). */
      let sha = null;
      let existingConfig = {};
      const checkRes = await fetch(apiUrl, {
        cache: 'no-store',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/vnd.github.v3+json',
        },
      });
      if (checkRes.ok) {
        const existing = await checkRes.json();
        sha = existing.sha;
        try {
          existingConfig = JSON.parse(decodeURIComponent(escape(atob(existing.content.replace(/\n/g, '')))));
        } catch {
          existingConfig = {};
        }
      }

      const newConfig = { ...existingConfig, archivo_tendencia: file.name };
      const content = JSON.stringify(newConfig, null, 2);
      const base64Content = btoa(unescape(encodeURIComponent(content)));

      const body = {
        message: `Dashboard: Establece "${file.name}" como tendencia predeterminada`,
        content: base64Content,
      };
      if (sha) body.sha = sha;

      let putRes = await fetch(apiUrl, {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept':        'application/vnd.github.v3+json',
          'Content-Type':  'application/json',
        },
        body: JSON.stringify(body),
      });

      /* Salvaguarda ante 422 "sha wasn't supplied" — mismo patrón que
         _setDefault(): refresca el sha (y el contenido, para no pisar
         un archivo_predeterminado guardado justo entre medio) y reintenta. */
      if (!putRes.ok && putRes.status === 422) {
        const retryCheck = await fetch(apiUrl, {
          cache: 'no-store',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/vnd.github.v3+json',
          },
        });
        if (retryCheck.ok) {
          const existing = await retryCheck.json();
          if (existing.sha) {
            let retryConfig = {};
            try {
              retryConfig = JSON.parse(decodeURIComponent(escape(atob(existing.content.replace(/\n/g, '')))));
            } catch {
              retryConfig = {};
            }
            const mergedRetry = { ...retryConfig, archivo_tendencia: file.name };
            body.sha = existing.sha;
            body.content = btoa(unescape(encodeURIComponent(JSON.stringify(mergedRetry, null, 2))));
            putRes = await fetch(apiUrl, {
              method: 'PUT',
              headers: {
                'Authorization': `Bearer ${token}`,
                'Accept':        'application/vnd.github.v3+json',
                'Content-Type':  'application/json',
              },
              body: JSON.stringify(body),
            });
          }
        }
      }

      if (!putRes.ok) {
        const errData = await putRes.json().catch(() => ({}));
        const detail  = errData.message || putRes.statusText;
        if (putRes.status === 401) throw new Error('Token inválido o sin permisos (401). Verifica tu PAT.');
        if (putRes.status === 422) throw new Error('Error de validación (422): ' + detail);
        throw new Error(`Error ${putRes.status}: ${detail}`);
      }

      AuthEngine._toast(`"${file.name}" establecido como tendencia predeterminada ✓`, 'success');

      /* Notificación de auditoría por Telegram (fire-and-forget) */
      if (typeof TelegramEngine !== 'undefined') {
        const usuarioActual = AuditEngine.getUser() || 'Desconocido';
        TelegramEngine.notifyFeatureUsed(usuarioActual, `Estableció "${file.name}" como tendencia predeterminada.`)
          .catch(err => console.error('[DbDefaultEngine] Error al notificar cambio de tendencia:', err));
      }

      /* Refleja de inmediato el nuevo estado en la UI del modal
         (banner + badge en la lista), sin esperar a la próxima apertura */
      this._currentTrend = file.name;
      this._renderList();
      const trendLabel = this._el('dbDefaultCurrentTrendLabel');
      if (trendLabel) trendLabel.innerHTML = `<i class="bi bi-graph-up-arrow me-1"></i>Tendencia actual: <strong>${file.name}</strong>`;

      /* Aplica la tendencia al dashboard ahora mismo, sin esperar al
         próximo login (reutiliza TrendEngine.setBaseline sin tocarlo) */
      await TrendEngine.setBaseline(file);

    } catch (err) {
      if (btn) { btn.disabled = false; btn.innerHTML = originalHTML; }
      alert(`No se pudo establecer "${file.name}" como tendencia predeterminada:\n${err.message}`);
    }
  },

  /* ── PUT a config.json en la raíz del repo con el nombre elegido ── */
  async _setDefault(file, itemEl) {
    if (!this._isMaster()) {
      alert('Solo un usuario con rol MAESTRO puede establecer el archivo predeterminado.');
      return;
    }

    const token = this._sessionToken;
    if (!token) { alert('Sesión de token expirada. Vuelve a abrir "Base de Datos (Beta)".'); return; }

    const btn = itemEl.querySelector('.btn-set-default');
    const originalHTML = btn ? btn.innerHTML : '';
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span>'; }

    try {
      const apiUrl = this.GITHUB_CONTENTS_BASE + 'config.json';

      /* Comprueba si config.json ya existe (para actualizar con su SHA).
         cache:'no-store' evita que el navegador reutilice una respuesta
         404 cacheada de la primera vez que el archivo aún no existía
         (causa del error 422 "sha wasn't supplied" en el segundo intento).
         NOTA: no se agrega el header 'Cache-Control' porque no es un
         header "simple" para CORS — GitHub rechaza el preflight que
         dispara y el fetch entero falla con "Failed to fetch". La opción
         cache:'no-store' del propio fetch() ya evita la caché sin
         necesidad de headers adicionales. */
      let sha = null;
      let existingConfig = {};
      const checkRes = await fetch(apiUrl, {
        cache: 'no-store',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/vnd.github.v3+json',
        },
      });
      if (checkRes.ok) {
        const existing = await checkRes.json();
        sha = existing.sha;
        try {
          existingConfig = JSON.parse(decodeURIComponent(escape(atob(existing.content.replace(/\n/g, '')))));
        } catch {
          existingConfig = {};
        }
      }

      const newConfig = { ...existingConfig, archivo_predeterminado: file.name };
      const content = JSON.stringify(newConfig, null, 2);
      const base64Content = btoa(unescape(encodeURIComponent(content)));

      const body = {
        message: `Dashboard: Establece "${file.name}" como archivo predeterminado`,
        content: base64Content,
      };
      if (sha) body.sha = sha;

      let putRes = await fetch(apiUrl, {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept':        'application/vnd.github.v3+json',
          'Content-Type':  'application/json',
        },
        body: JSON.stringify(body),
      });

      /* Salvaguarda: si el archivo ya existía pero el sha no llegó a
         tiempo (422 "sha wasn't supplied"), refresca el sha una vez
         más y reintenta el PUT antes de reportar error. */
      if (!putRes.ok && putRes.status === 422) {
        const retryCheck = await fetch(apiUrl, {
          cache: 'no-store',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/vnd.github.v3+json',
          },
        });
        if (retryCheck.ok) {
          const existing = await retryCheck.json();
          if (existing.sha) {
            let retryConfig = {};
            try {
              retryConfig = JSON.parse(decodeURIComponent(escape(atob(existing.content.replace(/\n/g, '')))));
            } catch {
              retryConfig = {};
            }
            const mergedRetry = { ...retryConfig, archivo_predeterminado: file.name };
            body.sha = existing.sha;
            body.content = btoa(unescape(encodeURIComponent(JSON.stringify(mergedRetry, null, 2))));
            putRes = await fetch(apiUrl, {
              method: 'PUT',
              headers: {
                'Authorization': `Bearer ${token}`,
                'Accept':        'application/vnd.github.v3+json',
                'Content-Type':  'application/json',
              },
              body: JSON.stringify(body),
            });
          }
        }
      }

      if (!putRes.ok) {
        const errData = await putRes.json().catch(() => ({}));
        const detail  = errData.message || putRes.statusText;
        if (putRes.status === 401) throw new Error('Token inválido o sin permisos (401). Verifica tu PAT.');
        if (putRes.status === 422) throw new Error('Error de validación (422): ' + detail);
        throw new Error(`Error ${putRes.status}: ${detail}`);
      }

      AuthEngine._toast(`"${file.name}" establecido como predeterminado ✓`, 'success');

      /* Notificación de auditoría por Telegram (fire-and-forget) */
      if (typeof TelegramEngine !== 'undefined') {
        const usuarioActual = AuditEngine.getUser() || 'Desconocido';
        TelegramEngine.notifyDefaultFileChanged(usuarioActual, file.name)
          .catch(err => console.error('[DbDefaultEngine] Error al notificar cambio de predeterminado:', err));
      }

      /* Refleja de inmediato el nuevo predeterminado en la UI del modal
         (banner + badge en la lista), sin esperar a la próxima apertura */
      this._currentDefault = file.name;
      this._renderList();
      const label = this._el('dbDefaultCurrentLabel');
      if (label) label.innerHTML = `<i class="bi bi-star-fill me-1"></i>Predeterminado actual: <strong>${file.name}</strong>`;

      /* Carga el archivo en el dashboard de inmediato, sin necesidad de
         pasarlo antes por el Historial (fire-and-forget: no bloquea ni
         condiciona el resultado de haberlo marcado como predeterminado) */
      AutoLoadEngine.loadFileByName(file.name).then(ok => {
        if (ok) AuthEngine._toast(`"${file.name}" cargado en el dashboard ✓`, 'success');
      });

    } catch (err) {
      if (btn) { btn.disabled = false; btn.innerHTML = originalHTML; }
      alert(`No se pudo establecer "${file.name}" como predeterminado:\n${err.message}`);
    }
  },

  init() {
    this._el('btnAbrirDbDefault')?.addEventListener('click', () => this._openTokenModal());
    this._el('btnDbTokenConfirm')?.addEventListener('click', () => this._confirmToken());

    this._el('btnDbTokenToggle')?.addEventListener('click', () => {
      const input = this._el('dbTokenInput');
      const icon  = this._el('dbTokenToggleIcon');
      if (!input) return;
      const isPass = input.type === 'password';
      input.type = isPass ? 'text' : 'password';
      icon.className = isPass ? 'bi bi-eye-slash' : 'bi bi-eye';
    });

    /* Permite confirmar con Enter dentro del input del token */
    this._el('dbTokenInput')?.addEventListener('keydown', e => {
      if (e.key === 'Enter') this._confirmToken();
    });

    /* Por seguridad: limpia el token de memoria y el campo al cerrar
       cualquiera de los dos modales (nunca queda residuo en la app) */
    this._el('modalDbToken')?.addEventListener('hidden.bs.modal', () => {
      const input = this._el('dbTokenInput');
      if (input) input.value = '';
    });
    this._el('modalDbDefault')?.addEventListener('hidden.bs.modal', () => {
      this._sessionToken = '';
    });

    this._el('btnDbDefaultRetry')?.addEventListener('click', () => this.fetchFileList());
  },
};


/* ────────────────────────────────────────────────────────────
   13.6 AUTO LOAD ENGINE — Carga automática del archivo predeterminado
       Se dispara justo después de un login exitoso (ver SessionEngine).
       Hace un fetch a config.json en la raíz del repo y, si existe,
       descarga y carga ese archivo con el mismo flujo que usa
       HistoryEngine._loadFile(). Agrega el token de AuthEngine cuando
       está disponible (igual que HistoryEngine.fetchFileList): la API
       de contenidos de GitHub es de lectura pública, pero sin token
       comparte un límite de solo 60 peticiones/hora POR IP entre todos
       los usuarios del dashboard — con varias personas usándolo desde
       la misma red, ese límite se agotaba fácilmente y estas lecturas
       empezaban a fallar en silencio (403), dejando de autocargar el
       archivo/tendencia predeterminados sin ningún error visible.
──────────────────────────────────────────────────────────── */
const AutoLoadEngine = {

  GITHUB_CONTENTS_BASE: 'https://api.github.com/repos/alexchouriors/M-tricas-REPORTE-DE-ASISTENCIAS-NUEVA/contents/',

  async loadDefaultFile() {
    try {
      const token = AuthEngine.getToken();
      const authHeader = token ? { 'Authorization': `Bearer ${token}` } : {};

      /* Lee config.json (si no existe, no hace nada) */
      const cfgRes = await fetch(this.GITHUB_CONTENTS_BASE + 'config.json', {
        cache: 'no-store',
        headers: { 'Accept': 'application/vnd.github.v3.raw', ...authHeader },
      });
      if (!cfgRes.ok) return;

      const config   = await cfgRes.json().catch(() => null);
      const fileName = config?.archivo_predeterminado;
      if (!fileName) return;

      await this.loadFileByName(fileName);
    } catch (err) {
      /* Nunca debe romper el flujo de login: solo se registra en consola */
      console.error('[AutoLoadEngine] No se pudo autocargar el archivo predeterminado:', err);
    }
  },

  /**
   * Descarga y carga un archivo puntual de REPORTES/ en el dashboard
   * (mismo flujo que HistoryEngine._loadFile). Reutilizable tanto desde
   * loadDefaultFile() (al iniciar sesión) como desde DbDefaultEngine
   * (al establecer un archivo como predeterminado, para reflejarlo de
   * inmediato sin tener que pasar por el Historial).
   *
   * @param {string} fileName
   * @returns {Promise<boolean>} true si se cargó correctamente
   */
  async loadFileByName(fileName) {
    try {
      const token = AuthEngine.getToken();
      const authHeader = token ? { 'Authorization': `Bearer ${token}` } : {};

      /* 1) Obtiene la metadata del archivo (necesitamos su download_url) */
      const fileRes = await fetch(
        this.GITHUB_CONTENTS_BASE + 'REPORTES/' + encodeURIComponent(fileName),
        { cache: 'no-store', headers: { 'Accept': 'application/vnd.github.v3+json', ...authHeader } }
      );
      if (!fileRes.ok) return false;
      const fileMeta = await fileRes.json();
      if (!fileMeta.download_url) return false;

      /* 2) Descarga y parsea el Excel */
      const bufRes = await fetch(fileMeta.download_url);
      if (!bufRes.ok) return false;
      const buffer   = await bufRes.arrayBuffer();
      const workbook = await WorkerEngine.parseWorkbookAsync(buffer); // offload a Web Worker (no bloquea la UI)

      DataStore.fileName = fileName;
      ExcelParser.parse(workbook);

      FilterEngine.populate(DataStore.rawMain);
      UIController.refresh();
      UIController.showDashboard();

      const footerEl = document.getElementById('footerFile');
      if (footerEl) footerEl.textContent = fileName;
      const titleEl = document.getElementById('reportTitle');
      if (titleEl) titleEl.textContent = DataStore.reportTitle || fileName;

      return true;
    } catch (err) {
      console.error('[AutoLoadEngine] No se pudo cargar el archivo:', fileName, err);
      return false;
    }
  },
};



document.addEventListener('DOMContentLoaded', () => {
  SessionEngine.init();
  ThemeEngine.init();
  UIController.init();
  /* GSheetsEngine.initModal() y HistoryEngine.init() YA NO se llaman
     aquí: ambos requieren que su modal (#gsheetsModal / #modalHistorial)
     ya exista en el DOM, y esos modales ahora se inyectan dinámicamente
     recién al abrirlos (ver LazyModals.js, que los invoca justo después
     de inyectar cada template). Esto ahorra el costo de construir esos
     dos modales pesados en la carga inicial para el 100% de las
     sesiones que nunca los abren. */
  AbsenceEngine.initEvents();
  AuthEngine.init();
  CloudEngine.init();
  SaveEngine.init();
  DeleteEngine.init();
  DbDefaultEngine.init();
  TrendEngine.init();
  TrendViewEngine.init();

  /*
    EXTENSIBILIDAD FUTURA:
    ─────────────────────
    Para añadir un nuevo KPI:
      1. Calcular el valor en KPIEngine.compute()
      2. Añadir el card HTML en index.html
      3. Actualizarlo en UIController.updateKPICards()

    Para añadir un nuevo gráfico:
      1. Añadir el canvas en index.html
      2. Crear el método ChartEngine.renderMyChart(kpis)
      3. Llamarlo dentro de ChartEngine.renderAll(kpis)

    Para añadir una nueva hoja de Excel:
      1. Añadir ExcelParser.parseMySheet(ws) con la lógica de lectura
      2. Agregarlo en ExcelParser.parse(workbook)
      3. Almacenar en DataStore.rawMySheet
      4. Renderizarlo en TableEngine.renderMyTable()

    DataStore.filters puede extenderse con nuevas claves sin
    romper el código existente (applyFilters() usa solo las claves
    definidas en el objeto).
  */
});

/* ================================================================
   MÓDULOS FUSIONADOS: SebastianAI + KpiHelpEngine + ReporteEjecutivoEngine
   ────────────────────────────────────────────────────────────
   Se conservan las 3 IIFE originales (function(){...})() TAL CUAL.
   Motivo (no son redundantes, son obligatorias): los 3 archivos
   declaran identificadores de nivel superior que CHOCAN entre sí
   -- p. ej. `init`, `STYLE_ID`/`injectStyles`/`inyectarEstilos`,
   `obtenerRecomendaciones`, `el`, `t`, `chip`, `wrap`, `menuEl`,
   `critico`, `alerta`, `normal`, `vigilar`, `usuario`, `opciones`,
   entre otros -- tanto entre sí como, en varios casos, con nombres
   ya usados en este mismo app.js. Sin sus IIFE, cualquiera de esas
   colisiones sería un `SyntaxError: Identifier '...' has already
   been declared` que rompería la carga COMPLETA del script (no solo
   del módulo en cuestión). Cada IIFE es, por diseño, el mecanismo
   que evita justamente ese problema -- se conservan intactas.

   Cada módulo se auto-inicializa exactamente igual que en su
   archivo original (mismo patrón `if (document.readyState ===
   'loading') ... else init()` al final de cada IIFE), así que no
   hace falta enganchar nada manualmente aquí.

   window.SebastianAI, window.KpiHelpEngine y window.
   ReporteEjecutivoEngine (si aplica) se siguen exponiendo
   exactamente igual que antes -- ninguna API pública cambia.
================================================================ */


/* ---- SebastianAI.js ---- */
/* ════════════════════════════════════════════════════════════
   SebastianAI.js
   ────────────────────────────────────────────────────────────
   Asistente Virtual (Bot de Comandos) — módulo 100% INDEPENDIENTE.

   Qué hace:
   - Botón flotante (esquina inferior) con animación de "respiración".
   - Ventana de chat propia, con sus estilos inyectados vía <style>
     y consumiendo las MISMAS variables CSS nativas del proyecto
     (var(--bg-card), var(--text-main), var(--gold), etc.), así que
     hereda Modo Claro/Oscuro automáticamente sin JS propio de tema.
   - Flujo de bienvenida secuencial la primera vez que se abre el
     chat en la sesión (sessionStorage, con llave propia).
   - NO usa ninguna API de IA externa: es un parser de comandos
     (switch/regex) que LEE directamente el DOM ya renderizado
     (valores de las tarjetas KPI, `<select>` de filtros y el
     usuario en sesión).

   Por qué esto es seguro para el control de acceso (RBAC):
   Este módulo NUNCA toca DataStore, AccessManager ni ningún dato
   crudo. Solo lee `textContent`/`value` de elementos que YA están
   en el DOM — y esos elementos ya fueron pintados por UIController
   después de que AccessManager filtró los datos según el usuario en
   sesión. En otras palabras: si el usuario no puede ver un dato en
   pantalla, Sebastián tampoco puede leerlo, porque nunca llegó a
   existir en el DOM. No hay bypass posible de los permisos.

   Qué NO hace (a propósito):
   - No importa, modifica ni referencia AccessManager.js,
     SecurityConfig.js, USUARIOS.JS, ChartEngine, TableEngine,
     DataStore ni KPIEngine. Solo LEE nodos del DOM ya existentes.
   - No agrega listeners en fase de captura ni intercepta clics de
     otros módulos (LazyModals, SwitchSessionEngine, etc.).
   - No llama a ninguna API externa de IA — 100% parser local.

   Cómo usarlo:
   Agrega este script en index.html, en cualquier orden respecto a
   los demás módulos independientes (no depende de ninguno, salvo
   que si TelegramEngine.js ya está cargado, opcionalmente podría
   auditar el uso — no lo hace por defecto, ver nota al final):
       <script src="SebastianAI.js"></script>
   Coloca el ícono del botón junto a index.html con el nombre exacto
   configurado en ICON_SRC más abajo (por defecto: "icon_1.png").
   ════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  /* ── Configuración ── */
  const ICON_SRC        = 'icon_1.png'; // ruta del ícono del botón flotante
  const BOT_NAME         = 'Sebastián';
  const WELCOME_FLAG_BASE = 'sebastianai_welcomed'; // sessionStorage — una vez por USUARIO, por sesión de pestaña
  const MENU_INFORMED_BASE = 'sebastianai_menu_informed'; // sessionStorage — aviso proactivo del Menú, una vez por USUARIO
  const STYLE_ID          = 'sebastianAIStyles';
  const FAB_ID             = 'sebastianAIFab';
  const PANEL_ID            = 'sebastianAIPanel';

  const FALLBACK_MSG =
    'NO TENGO EL ADIESTRAMIENTO NECESARIO PARA RESPONDER A ELLO PERO LO TENDRE EN CUENTA PARA UNA PROXIMA CONSULTA';

  /* Catálogo ÚNICO de comandos — se usa tanto en el flujo de
     bienvenida (ahora muestra TODOS) como en el comando "ayuda".
     Cubre TODAS las secciones del dashboard: KPIs, filtros, tablas,
     monitor de ausencias, reporte cargado, tema, sesión, el botón
     "Menú" del sidebar, y un resumen que las junta todas. */
  const ALL_COMMANDS = [
    'Reporte Ejecutivo',
    '¿Cuál es el total de asistencia?',
    '¿Cómo van las células?',
    '¿Cómo va el servicio?',
    '¿Cuántos nuevos hay?',
    '¿Qué grupo ministerial estoy viendo?',
    '¿Cuáles son mis filtros activos?',
    '¿Qué reporte está cargado?',
    '¿Cómo está el monitor de ausencias?',
    '¿Cuántas personas hay en la tabla?',
    '¿Cómo va mi tendencia?',
    '¿En qué tema estoy?',
    '¿Qué usuario tengo activo?',
    '¿Qué hace cada botón del Menú?',
    'Dame un resumen completo',
  ];

  /* ══════════════════════════════════════════════════════════
     ESTILOS — inyectados aislados, consumen las variables CSS
     nativas del proyecto (definidas en :root / [data-theme="light"]
     en style.css) para que el chat cambie de tema automáticamente
     junto con el resto del dashboard, sin ningún JS propio de tema.
     ══════════════════════════════════════════════════════════ */
  function injectStyles() {
    /* CSS ahora vive en style.css — ya no se inyecta por JS. */
    return;
  }

  /* ══════════════════════════════════════════════════════════
     LECTOR DE PANTALLA — extrae SOLO lo que ya está en el DOM.
     Nunca consulta datos crudos ni AccessManager directamente.
     ══════════════════════════════════════════════════════════ */
  const Reader = {
    text(id, fallback = 'No disponible en este momento') {
      const el = document.getElementById(id);
      if (!el) return fallback;
      const t = (el.textContent || '').trim();
      return t !== '' ? t : fallback;
    },

    /* Si el valor está en medio del overlay "Recalculando…"
       (ver UIController._setKpisRecalculando en app.js), lo detecta
       para no reportar ese texto como si fuera un dato real. */
    isRecalculando(id) {
      const el = document.getElementById(id);
      return !!el && el.classList.contains('kpi-recalculando');
    },

    selectLabel(id, fallback = 'Todos') {
      const el = document.getElementById(id);
      if (!el || typeof el.selectedIndex !== 'number' || el.selectedIndex < 0) return fallback;
      const opt = el.options[el.selectedIndex];
      return opt ? opt.textContent.trim() : fallback;
    },

    sesionActiva() {
      try {
        return sessionStorage.getItem('ccrm_dashboard_user') || null;
      } catch (e) {
        return null;
      }
    },

    temaActivo() {
      const t = document.documentElement.getAttribute('data-theme');
      return t === 'light' ? 'Claro' : 'Oscuro';
    },

    /* Nombre del reporte/archivo cargado — mismo nodo que pinta
       UIController al cargar el Excel (#reportTitle) y su respaldo
       en el pie de página (#footerFile). */
    reporteActivo() {
      const titulo = this.text('reportTitle', '');
      if (titulo && titulo !== 'Cargue un archivo para comenzar') return titulo;
      const footer = this.text('footerFile', '');
      return footer || null;
    },
  };

  /* ══════════════════════════════════════════════════════════
     PARSER DE COMANDOS — switch/regex, sin ninguna API externa.
     Cada handler lee el DOM vía Reader y arma la respuesta.
     ══════════════════════════════════════════════════════════ */
  function normalizar(str) {
    return (str || '')
      .toString()
      .trim()
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // quita acentos
  }

  function avisoSiRecalculando(ids) {
    if (ids.some(id => Reader.isRecalculando(id))) {
      return 'Justo estoy recalculando esos números — dame un segundo y vuelve a preguntarme 🙂';
    }
    return null;
  }

  /* Trae las recomendaciones del MISMO motor (decisionEngine) que usa
     el botón "?" de cada tarjeta KPI, vía la API pública que expone
     KpiHelpEngine.js (window.KpiHelpEngine.getRecomendaciones). Acepta
     uno o varios kpiId (p. ej. célula SI + célula NO en una sola
     respuesta) y fusiona el resultado sin duplicar por `comando`. Si
     KpiHelpEngine.js no está cargado, devuelve [] silenciosamente —
     el chat sigue funcionando igual, solo sin los chips sugeridos. */
  /* Umbral de inasistencia (%) a partir del cual un frente se
     considera "para vigilar"/"crítico". Un solo número, fácil de
     ajustar si el criterio cambia más adelante. */
  const UMBRAL_VIGILAR_PCT  = 15;
  const UMBRAL_CRITICO_PCT  = 30;

  /* Extrae el número de un texto tipo "37%" / "37,5 %" ya pintado en
     el DOM. Devuelve `null` si el texto no trae un porcentaje
     reconocible (dato ausente, "No disponible en este momento", una
     tarjeta que no existe en la vista restringida del usuario, etc.)
     — SIN inventar ni recalcular el número: solo lee lo que
     Reader.text() ya trajo del DOM. */
  function parsePorcentaje(texto) {
    if (!texto) return null;
    const m = texto.match(/(-?\d+(?:[.,]\d+)?)\s*%/);
    if (!m) return null;
    return parseFloat(m[1].replace(',', '.'));
  }

  /* Extrae el número entero de un texto tipo "128" / "128 registros"
     ya pintado en el DOM. Devuelve 0 si no hay número reconocible
     (mismo criterio fail-safe que parsePorcentaje). */
  function parseEntero(texto) {
    if (!texto) return 0;
    const m = texto.match(/-?\d+/);
    return m ? parseInt(m[0], 10) : 0;
  }

  /* Clasifica un % de INASISTENCIA (más alto = peor) en tres niveles.
     Fail-safe: sin dato → 'sin dato', tratado como digno de mención
     pero no como crítico (no se puede diagnosticar lo que no se
     puede leer). */
  function clasificarInasistencia(pct) {
    if (pct === null) return { icon: '❔', nivel: 'sin dato' };
    if (pct >= UMBRAL_CRITICO_PCT) return { icon: '🔴', nivel: 'crítico' };
    if (pct >= UMBRAL_VIGILAR_PCT) return { icon: '⚠️', nivel: 'para vigilar' };
    return { icon: '✅', nivel: 'saludable' };
  }

  /**
   * Analiza UN frente de asistencia (Célula / Servicio / Ambos): lee
   * SI, NO y % de inasistencia ya renderizados (Reader.text — mismo
   * dato que ve el usuario en su tarjeta KPI, ya filtrado por
   * AccessManager/vista restringida), lo clasifica y arma su propia
   * línea de diagnóstico + una recomendación puntual según el nivel.
   *
   * @param {string} nombre        - Nombre legible del frente (ej. 'Célula').
   * @param {string} siId          - id del DOM con el valor "SI".
   * @param {string} noId          - id del DOM con el valor "NO".
   * @param {string} pctNoId       - id del DOM con el % de inasistencia.
   * @param {string} kpiIdNo       - kpi-id (KpiHelpEngine) para las recomendaciones si no está saludable.
   * @param {string} recomendacion - Texto de acción sugerida si el frente NO está saludable.
   * @returns {{texto:string, nivel:string, kpiIdSiCritico:string|null}}
   */
  function analizarFrenteAsistencia(nombre, siId, noId, pctNoId, kpiIdNo, recomendacion) {
    const si = Reader.text(siId);
    const no = Reader.text(noId);
    const pct = parsePorcentaje(Reader.text(pctNoId, ''));
    const clase = clasificarInasistencia(pct);
    const pctTexto = pct === null ? 'sin dato' : `${pct}%`;

    const accion = clase.nivel === 'saludable'
      ? 'Sin acción requerida por ahora.'
      : clase.nivel === 'sin dato'
        ? 'No se pudo leer el porcentaje — revisa manualmente esta tarjeta.'
        : recomendacion;

    return {
      texto: `${clase.icon} ${nombre} — SI: ${si} · NO: ${no} (${pctTexto} de inasistencia, ${clase.nivel})\n   → ${accion}`,
      nivel: clase.nivel,
      kpiIdSiCritico: (clase.nivel === 'para vigilar' || clase.nivel === 'crítico') ? kpiIdNo : null,
    };
  }

  /**
   * Analiza UN frente de "Nuevos" (Célula / Servicio): a diferencia de
   * los frentes de asistencia, aquí más alto siempre es mejor (no hay
   * "% de inasistencia" que evaluar) — el diagnóstico es simplemente
   * si hubo o no ingresos nuevos en el periodo del reporte.
   */
  function analizarFrenteNuevos(nombre, valorId, kpiIdSiCero, recomendacionSiCero) {
    const valorTexto = Reader.text(valorId, '0');
    const valor = parseEntero(valorTexto);
    const icon = valor > 0 ? '✅' : '⚠️';
    const nivel = valor > 0 ? 'con ingresos nuevos' : 'sin ingresos nuevos';
    const accion = valor > 0
      ? 'Da seguimiento cercano a los recién llegados para que no se pierdan en las próximas semanas.'
      : recomendacionSiCero;

    return {
      texto: `${icon} ${nombre}: ${valorTexto} (${nivel})\n   → ${accion}`,
      kpiIdSiCritico: valor === 0 ? kpiIdSiCero : null,
    };
  }

  /* Analiza el Monitor de Ausencias completo (las 4 categorías) y
     prioriza la recomendación por el nivel más urgente presente. */
  function analizarMonitorAusencias() {
    const normal  = parseEntero(Reader.text('ausNormalCount', '0'));
    const vigilar = parseEntero(Reader.text('ausWatchCount', '0'));
    const alerta  = parseEntero(Reader.text('ausWarnCount', '0'));
    const critico = parseEntero(Reader.text('ausCritCount', '0'));

    let icon = '✅', nivel = 'bajo control', accion = 'No hay casos urgentes en el monitor de ausencias.';
    if (critico > 0) {
      icon = '🔴'; nivel = 'requiere atención inmediata';
      accion = `Prioriza contacto directo esta semana con las ${critico} persona(s) en estado Crítico.`;
    } else if (alerta > 0) {
      icon = '⚠️'; nivel = 'para vigilar de cerca';
      accion = `Da seguimiento a las ${alerta} persona(s) en Alerta antes de que pasen a Crítico.`;
    }

    return `${icon} Monitor de Ausencias — Normal: ${normal} · A vigilar: ${vigilar} · Alerta: ${alerta} · Crítico: ${critico} (${nivel})\n   → ${accion}`;
  }

  /**
   * Motor de análisis PROFUNDO para el "Análisis con IA"
   * (AiAnalysisButton.js): recorre KPI por KPI, cada uno con su propio
   * diagnóstico y recomendación puntual, y luego las secciones de
   * abajo (tablas + Monitor de Ausencias) — todo a partir de
   * Reader.text() sobre el DOM YA renderizado (KPIEngine + AccessManager
   * en app.js), así que respeta exactamente la misma vista
   * filtrada/restringida que el usuario tiene en pantalla en ese
   * momento: este motor nunca lee ni calcula datos fuera de su alcance.
   *
   * @returns {{texto:string, kpiIds:string[]}}
   */
  function analizarDashboardCompleto() {
    const nombre = Reader.reporteActivo();

    const frenteCelula = analizarFrenteAsistencia(
      'Célula', 'kpiCelulasSI', 'kpiCelulasNO', 'kpiCelulasNOPct', 'celula-no',
      'Filtra por Grupo Ministerial para ver si la inasistencia es general o de un líder/célula puntual, y prioriza contacto con quienes llevan más tiempo sin asistir.'
    );
    const frenteServicio = analizarFrenteAsistencia(
      'Servicio', 'kpiServicioSI', 'kpiServicioNO', 'kpiServicioNOPct', 'servicio-no',
      'Revisa si la inasistencia a servicio coincide con la de célula (posible desconexión general) o es un frente aparte.'
    );
    const frenteAmbos = analizarFrenteAsistencia(
      'Ambos (Célula + Servicio)', 'kpiAmbosSI', 'kpiAmbosNO', 'kpiAmbosNOPct', 'ambos-no',
      'Estas personas no asistieron a NADA en el periodo — son la prioridad más alta de seguimiento pastoral.'
    );
    const frenteNuevosCelula = analizarFrenteNuevos(
      'Nuevos en Célula', 'kpiNuevosCelula', 'nuevos-celula',
      'No hubo ingresos nuevos a célula en este periodo — evalúa estrategias de invitación con los líderes de cada grupo.'
    );
    const frenteNuevosServicio = analizarFrenteNuevos(
      'Nuevos en Servicio', 'kpiNuevosServicio', 'nuevos-servicio',
      'No hubo visitas nuevas al servicio en este periodo — evalúa reforzar la difusión/invitación general.'
    );
    const monitorAusencias = analizarMonitorAusencias();

    // Recomendaciones: siempre 'pct-general' + cada frente que NO salió "saludable"
    const kpiIds = ['pct-general'];
    [frenteCelula, frenteServicio, frenteAmbos, frenteNuevosCelula, frenteNuevosServicio]
      .forEach(f => { if (f.kpiIdSiCritico) kpiIds.push(f.kpiIdSiCritico); });

    const totalFrentesAtencion = kpiIds.length - 1;
    const cierre = totalFrentesAtencion > 0
      ? `${totalFrentesAtencion} frente(s) necesitan atención — revisa las recomendaciones que te dejo abajo.`
      : `Todos los frentes están en niveles saludables. Buen trabajo — sigue así.`;

    const texto = [
      `📊 REPORTE EJECUTIVO — ${nombre || 'sin reporte cargado'}`,
      `Vista: ${Reader.selectLabel('filterGroup', 'Todos los grupos')} · Usuario: ${Reader.sesionActiva() || 'sin sesión'}`,
      ``,
      `Total registrados: ${Reader.text('kpiTotal')} · Asistencia general: ${Reader.text('kpiPctGeneral')}`,
      ``,
      frenteCelula.texto,
      ``,
      frenteServicio.texto,
      ``,
      frenteAmbos.texto,
      ``,
      frenteNuevosCelula.texto,
      ``,
      frenteNuevosServicio.texto,
      ``,
      `📋 Personas: ${Reader.text('countPersonas', '—')} · Excluidos: ${Reader.text('countExcluidos', '—')} · Nuevos: ${Reader.text('countNuevos', '—')} · Histórico: ${Reader.text('countHistorico', '—')}`,
      ``,
      monitorAusencias,
      ``,
      `Diagnóstico general: ${cierre}`,
    ].join('\n');

    return { texto, kpiIds };
  }

  /* Frase EXACTA que envía el botón "Reporte Ejecutivo" (#btnReporteEjecutivo,
     ver init() más abajo) vía window.SebastianAI.preguntar('Reporte Ejecutivo'),
     y también el `comando` que agrega KpiHelpEngine.decisionEngine() como
     opción sugerida permanente ('REPORTE EJECUTIVO'). Se compara normalizada
     (sin acentos/mayúsculas) para tolerar cualquier diferencia menor de
     capitalización — nunca al revés: este archivo no decide el texto del
     botón, solo lo reconoce. */
  const FRASE_ANALISIS_IA = normalizar('Reporte Ejecutivo');

  function obtenerRecomendaciones(kpiIds) {
    if (typeof window.KpiHelpEngine === 'undefined' || typeof window.KpiHelpEngine.getRecomendaciones !== 'function') {
      return [];
    }
    const ids = Array.isArray(kpiIds) ? kpiIds : [kpiIds];
    const vistos = new Set();
    const combinadas = [];
    ids.forEach(id => {
      (window.KpiHelpEngine.getRecomendaciones(id) || []).forEach(opt => {
        if (!opt || !opt.comando || vistos.has(opt.comando)) return;
        vistos.add(opt.comando);
        combinadas.push(opt);
      });
    });
    return combinadas;
  }

  const COMMANDS = [
    /* ── Comandos que disparan los CHIPS del motor de recomendaciones
       (ver OPCIONES_FIJAS_POR_KPI / decisionEngine en KpiHelpEngine.js)
       — antes caían todos en el fallback porque el parser no los
       reconocía. Van primero en la lista para que no los intercepte
       ningún regex más genérico de más abajo (p. ej. "detalle célula"
       contiene "célula", que si no fuera primero lo capturaría el
       comando genérico de Célula). No abren el modal de personas —
       eso sigue siendo EXCLUSIVO de hacer clic en la tarjeta — solo
       dan más contexto en números y, cuando aplica, invitan a hacerlo. */
    {
      test: t => /\bcompar(ar|a)\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiCelulasSI', 'kpiServicioSI']);
        if (aviso) return { text: aviso };
        return {
          text: [
            `Comparando Célula vs. Servicio con tus filtros actuales:`,
            `• Célula — SI: ${Reader.text('kpiCelulasSI')} (${Reader.text('kpiCelulasSIPct')}) · NO: ${Reader.text('kpiCelulasNO')} (${Reader.text('kpiCelulasNOPct')})`,
            `• Servicio — SI: ${Reader.text('kpiServicioSI')} (${Reader.text('kpiServicioSIPct')}) · NO: ${Reader.text('kpiServicioNO')} (${Reader.text('kpiServicioNOPct')})`,
          ].join('\n'),
          kpiIds: ['pct-celula', 'pct-servicio'],
        };
      },
    },
    {
      test: t => /\bdetalle\s+celula\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiCelulasNO']);
        if (aviso) return { text: aviso };
        return {
          text: `Célula — no asistieron: ${Reader.text('kpiCelulasNO')} (${Reader.text('kpiCelulasNOPct')}). Para ver los nombres, haz clic directamente en la tarjeta de Célula.`,
          kpiIds: ['celula-no'],
        };
      },
    },
    {
      test: t => /\bdetalle\s+servicio\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiServicioNO']);
        if (aviso) return { text: aviso };
        return {
          text: `Servicio — no asistieron: ${Reader.text('kpiServicioNO')} (${Reader.text('kpiServicioNOPct')}). Para ver los nombres, haz clic directamente en la tarjeta de Servicio.`,
          kpiIds: ['servicio-no'],
        };
      },
    },
    {
      test: t => /\bdetalle\s+ambos\s+si\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiAmbosSI']);
        if (aviso) return { text: aviso };
        return {
          text: `Núcleo comprometido (asistieron a ambos): ${Reader.text('kpiAmbosSI')} (${Reader.text('kpiAmbosSIPct')}). Para ver los nombres, haz clic directamente en la tarjeta de Ambos SI.`,
          kpiIds: ['ambos-si'],
        };
      },
    },
    {
      test: t => /\bdetalle\s+ambos\s+no\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiAmbosNO']);
        if (aviso) return { text: aviso };
        return {
          text: `En riesgo (no asistieron a ninguno): ${Reader.text('kpiAmbosNO')} (${Reader.text('kpiAmbosNOPct')}). Para ver los nombres, haz clic directamente en la tarjeta de Ambos NO.`,
          kpiIds: ['ambos-no'],
        };
      },
    },
    {
      test: t => /\bdetalle\s+nuevos\s+celula\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiNuevosCelula']);
        if (aviso) return { text: aviso };
        return {
          text: `Nuevos en Célula: ${Reader.text('kpiNuevosCelula')}. Para ver los nombres, haz clic directamente en la tarjeta de Nuevos en Célula.`,
          kpiIds: ['nuevos-celula'],
        };
      },
    },
    {
      test: t => /\bdetalle\s+nuevos\s+servicio\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiNuevosServicio']);
        if (aviso) return { text: aviso };
        return {
          text: `Nuevos en Servicio: ${Reader.text('kpiNuevosServicio')}. Para ver los nombres, haz clic directamente en la tarjeta de Nuevos en Servicio.`,
          kpiIds: ['nuevos-servicio'],
        };
      },
    },
    {
      test: t => /\bplan de seguimiento\b/.test(t),
      run: () => ({
        text: [
          `Un plan de seguimiento simple para quienes están en riesgo:`,
          `1. Filtra por Grupo Ministerial para acotar a un líder/célula específico.`,
          `2. Haz clic en la tarjeta "Ausentes en Ambos" para ver la lista de personas.`,
          `3. Prioriza contacto directo con quienes llevan más tiempo sin asistir (revisa el Monitor de Ausencias — pregúntame "ausencias").`,
        ].join('\n'),
        kpiIds: ['ambos-no'],
      }),
    },
    {
      test: t => /\banalizar caida\b/.test(t),
      run: () => ({
        text: [
          `Para analizar una caída, te sugiero en este orden:`,
          `1. Activa una línea base de tendencia (Menú → Historial) y pregúntame "tendencia" para ver el cambio real.`,
          `2. Filtra por Grupo Ministerial para ver si la caída es general o de un grupo puntual.`,
          `3. Revisa el Monitor de Ausencias (pregúntame "ausencias") para ver si hay personas en estado Crítico o de Alerta.`,
        ].join('\n'),
      }),
    },

    {
      test: t => /\b(usuario|sesion|quien soy|mi nombre)\b/.test(t),
      run: () => {
        const user = Reader.sesionActiva();
        return user
          ? `Tu sesión activa es: ${user}.`
          : 'No detecto ninguna sesión activa en este momento.';
      },
    },
    {
      test: t => /\bgrupo\b/.test(t),
      run: () => `Grupo Ministerial seleccionado: ${Reader.selectLabel('filterGroup', 'Todos los grupos')}.`,
    },
    {
      test: t => /\b(filtros?|estado de los filtros)\b/.test(t) && !/\bgrupo\b/.test(t),
      run: () => {
        return [
          `Estos son tus filtros activos ahora mismo:`,
          `• Grupo Ministerial: ${Reader.selectLabel('filterGroup', 'Todos los grupos')}`,
          `• Estado: ${Reader.selectLabel('filterEstado', 'Todos')}`,
          `• Célula: ${Reader.selectLabel('filterCelula', 'Todas')}`,
          `• Servicio: ${Reader.selectLabel('filterServicio', 'Todos')}`,
          `• Nuevo: ${Reader.selectLabel('filterNuevo', 'Todos')}`,
        ].join('\n');
      },
    },
    {
      test: t => /\b(total|asistencia total|kpi total)\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiTotal', 'kpiPctGeneral']);
        if (aviso) return aviso;
        return {
          text: `El total de asistencia con tus filtros actuales es ${Reader.text('kpiTotal')} (${Reader.text('kpiPctGeneral')} de asistencia general).`,
          kpiIds: ['total', 'pct-general'],
        };
      },
    },
    {
      test: t => /\bcelula|c[ée]lulas?\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiCelulasSI', 'kpiCelulasNO']);
        if (aviso) return aviso;
        return {
          text: [
            `Célula — con tus filtros actuales:`,
            `• Asistieron (SI): ${Reader.text('kpiCelulasSI')} (${Reader.text('kpiCelulasSIPct')})`,
            `• No asistieron (NO): ${Reader.text('kpiCelulasNO')} (${Reader.text('kpiCelulasNOPct')})`,
          ].join('\n'),
          kpiIds: ['celula-si', 'celula-no'],
        };
      },
    },
    {
      test: t => /\bservicio\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiServicioSI', 'kpiServicioNO']);
        if (aviso) return aviso;
        return {
          text: [
            `Servicio — con tus filtros actuales:`,
            `• Asistieron (SI): ${Reader.text('kpiServicioSI')} (${Reader.text('kpiServicioSIPct')})`,
            `• No asistieron (NO): ${Reader.text('kpiServicioNO')} (${Reader.text('kpiServicioNOPct')})`,
          ].join('\n'),
          kpiIds: ['servicio-si', 'servicio-no'],
        };
      },
    },
    {
      test: t => /\b(ambos|celula y servicio)\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiAmbosSI', 'kpiAmbosNO']);
        if (aviso) return aviso;
        return {
          text: [
            `Célula y Servicio (ambos) — con tus filtros actuales:`,
            `• Asistieron a ambos (SI): ${Reader.text('kpiAmbosSI')} (${Reader.text('kpiAmbosSIPct')})`,
            `• No asistieron a ninguno (NO): ${Reader.text('kpiAmbosNO')} (${Reader.text('kpiAmbosNOPct')})`,
          ].join('\n'),
          kpiIds: ['ambos-si', 'ambos-no'],
        };
      },
    },
    {
      test: t => /\bnuevos?\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiNuevosCelula', 'kpiNuevosServicio']);
        if (aviso) return aviso;
        return {
          text: [
            `Personas nuevas — con tus filtros actuales:`,
            `• En Célula: ${Reader.text('kpiNuevosCelula')}`,
            `• En Servicio: ${Reader.text('kpiNuevosServicio')}`,
          ].join('\n'),
          kpiIds: ['nuevos-celula', 'nuevos-servicio'],
        };
      },
    },
    {
      test: t => /\b(reporte|archivo( cargado)?|excel)\b/.test(t),
      run: () => {
        const nombre = Reader.reporteActivo();
        return nombre
          ? `El reporte cargado actualmente es: "${nombre}".`
          : 'Todavía no hay ningún reporte cargado en el dashboard.';
      },
    },
    {
      test: t => /\b(personas|tabla de personas)\b/.test(t),
      run: () => `Tabla de Personas: ${Reader.text('countPersonas', 'sin datos')}.`,
    },
    {
      test: t => /\bexcluidos?\b/.test(t),
      run: () => `Tabla de Excluidos: ${Reader.text('countExcluidos', 'sin datos')}.`,
    },
    {
      test: t => /\b(historico|histórico)\b/.test(t),
      run: () => `Tabla de Histórico: ${Reader.text('countHistorico', 'sin datos')}.`,
    },
    {
      test: t => /\b(ausencia|ausente|monitor de ausencias)\b/.test(t),
      run: () => {
        return [
          `Monitor de Ausencias — con tus filtros actuales:`,
          `• Normal: ${Reader.text('ausNormalCount', '0')}`,
          `• A vigilar: ${Reader.text('ausWatchCount', '0')}`,
          `• Alerta: ${Reader.text('ausWarnCount', '0')}`,
          `• Crítico: ${Reader.text('ausCritCount', '0')}`,
          `(${Reader.text('countAusencias', 'sin datos')})`,
        ].join('\n');
      },
    },
    {
      test: t => /\b(resumen|todo|resumen completo|dashboard completo|panorama general)\b/.test(t),
      run: () => {
        const aviso = avisoSiRecalculando(['kpiTotal', 'kpiPctGeneral']);
        const nombre = Reader.reporteActivo();

        return {
          text: [
            `Resumen completo del dashboard ahora mismo:`,
            ``,
            `📄 Reporte: ${nombre || 'ninguno cargado'}`,
            `👤 Usuario: ${Reader.sesionActiva() || 'sin sesión'}`,
            `🎨 Tema: ${Reader.temaActivo()}`,
            ``,
            `🔎 Filtros: Grupo "${Reader.selectLabel('filterGroup', 'Todos los grupos')}" · Estado "${Reader.selectLabel('filterEstado', 'Todos')}" · Célula "${Reader.selectLabel('filterCelula', 'Todas')}" · Servicio "${Reader.selectLabel('filterServicio', 'Todos')}" · Nuevo "${Reader.selectLabel('filterNuevo', 'Todos')}"`,
            ``,
            aviso || [
              `📊 KPIs — Total: ${Reader.text('kpiTotal')} (${Reader.text('kpiPctGeneral')})`,
              `   Célula SI/NO: ${Reader.text('kpiCelulasSI')} / ${Reader.text('kpiCelulasNO')}`,
              `   Servicio SI/NO: ${Reader.text('kpiServicioSI')} / ${Reader.text('kpiServicioNO')}`,
              `   Ambos SI/NO: ${Reader.text('kpiAmbosSI')} / ${Reader.text('kpiAmbosNO')}`,
              `   Nuevos (Célula/Servicio): ${Reader.text('kpiNuevosCelula')} / ${Reader.text('kpiNuevosServicio')}`,
            ].join('\n'),
            ``,
            `📋 Tablas — Personas: ${Reader.text('countPersonas', '—')} · Excluidos: ${Reader.text('countExcluidos', '—')} · Nuevos: ${Reader.text('countNuevos', '—')} · Histórico: ${Reader.text('countHistorico', '—')}`,
            ``,
            `⏱️ Ausencias — Normal: ${Reader.text('ausNormalCount', '0')} · A vigilar: ${Reader.text('ausWatchCount', '0')} · Alerta: ${Reader.text('ausWarnCount', '0')} · Crítico: ${Reader.text('ausCritCount', '0')}`,
          ].join('\n'),
          kpiIds: ['pct-general'],
        };
      },
    },
    {
      test: t => /\btendencia\b/.test(t),
      run: () => {
        const banner = document.getElementById('trendBaselineBanner');
        const activa = !!banner && !banner.classList.contains('d-none');

        if (!activa) {
          return 'Todavía no tienes una línea base de tendencia activa. Para activarla: abre el Menú → Historial → y elige un reporte anterior como línea base de comparación.';
        }

        const archivo = Reader.text('trendBaselineFileName', 'línea base');

        // Mismos valueId que declara KPI_DETAIL_MAP en app.js — TrendEngine
        // pinta `${valueId}TrendRow` (fila, oculta con d-none si ese KPI no
        // tiene delta), `${valueId}Trend` (ícono + %) y `${valueId}TrendPrev`
        // ("antes: N") para cada uno.
        const kpisConTendencia = [
          { id: 'kpiTotal',          label: 'Total' },
          { id: 'kpiCelulasSI',      label: 'Célula SI' },
          { id: 'kpiCelulasNO',      label: 'Célula NO' },
          { id: 'kpiServicioSI',     label: 'Servicio SI' },
          { id: 'kpiServicioNO',     label: 'Servicio NO' },
          { id: 'kpiAmbosSI',        label: 'Ambos SI' },
          { id: 'kpiAmbosNO',        label: 'Ambos NO' },
          { id: 'kpiNuevosCelula',   label: 'Nuevos Célula' },
          { id: 'kpiNuevosServicio', label: 'Nuevos Servicio' },
        ];

        const lineas = kpisConTendencia
          .map(({ id, label }) => {
            const rowEl = document.getElementById(`${id}TrendRow`);
            if (!rowEl || rowEl.classList.contains('d-none')) return null; // sin delta para este KPI
            const cambio = Reader.text(`${id}Trend`, '');
            const antes  = Reader.text(`${id}TrendPrev`, '');
            if (!cambio) return null;
            return `• ${label}: ${cambio}${antes ? ` (${antes})` : ''}`;
          })
          .filter(Boolean);

        if (lineas.length === 0) {
          return `Tienes activa la línea base "${archivo}", pero todavía no hay tendencia calculada para mostrar con tus filtros actuales.`;
        }

        return [`📈 Tendencia — comparando contra "${archivo}":`, ...lineas].join('\n');
      },
    },
    {
      test: t => /\b(menu|menú)\b/.test(t),
      run: () => explicarMenu(),
    },
    {
      test: t => /\b(tema|modo (claro|oscuro)|theme)\b/.test(t),
      run: () => `El dashboard está en Modo ${Reader.temaActivo()} ahora mismo.`,
    },
    {
      test: t => /\b(ayuda|comandos|help|opciones)\b/.test(t),
      run: () => ({ text: 'Estos son todos los comandos que puedes usarme:', showCommandList: ALL_COMMANDS }),
    },
  ];

  /* Texto explicativo de cada botón dentro del offcanvas "Menú" del
     sidebar (#sidebarMenu en index.html). Es texto fijo (no lee datos
     dinámicos) porque describe FUNCIONALIDAD de la interfaz, no datos
     del reporte — se reutiliza tanto para el comando "menú" como para
     el aviso proactivo cuando el usuario abre esa opción. */
  function explicarMenu() {
    return [
      `El botón "Menú" abre el panel lateral con estas opciones:`,
      `• Guardar — sube el reporte actual a GitHub (se activa solo cuando cargaste un Excel local).`,
      `• Eliminar — borra un reporte guardado en el historial.`,
      `• Base de Datos (Beta) — define el archivo predeterminado que se autocarga al iniciar sesión (pide tu token de GitHub cada vez, nunca lo guarda).`,
      `• Historial — abre la lista de reportes guardados en GitHub para cargarlos o descargarlos.`,
      `• Cambiar Sesión — cambia de usuario sin cerrar del todo el dashboard.`,
      `• Cerrar sesión — termina tu sesión actual.`,
      `• Versión Ligera — abre una versión alterna del dashboard, más liviana.`,
    ].join('\n');
  }

  function procesarComando(rawText) {
    const t = normalizar(rawText);

    /* Bypass deliberado: la frase que envía AiAnalysisButton.js es
       conocida y fija (no texto libre del usuario), pero comparte
       palabras con comandos más genéricos — en particular "reporte",
       que también dispara el comando de "¿qué archivo está cargado?"
       (test: /\b(reporte|archivo( cargado)?|excel)\b/). Como ese
       comando aparece ANTES en el array COMMANDS, sin este bypass
       intercepta la frase antes de llegar al comando "resumen" y el
       botón nunca dispara el análisis profundo. Resolverlo así (en
       vez de reordenar COMMANDS o cambiar el texto del botón) evita
       que cualquier futura palabra en común entre comandos vuelva a
       romper esta conexión. */
    if (t.includes(FRASE_ANALISIS_IA)) {
      const aviso = avisoSiRecalculando(['kpiTotal', 'kpiPctGeneral']);
      const { texto, kpiIds } = analizarDashboardCompleto();
      return { text: aviso || texto, kpiIds };
    }

    for (const cmd of COMMANDS) {
      if (cmd.test(t)) {
        const result = cmd.run(t); // `t` = texto normalizado (minúsculas, sin acentos) — opcional, los run() existentes que no lo declaran simplemente lo ignoran
        return typeof result === 'string' ? { text: result } : result;
      }
    }
    return { text: FALLBACK_MSG };
  }

  /* ══════════════════════════════════════════════════════════
     UI — construcción del botón flotante y el panel de chat.
     ══════════════════════════════════════════════════════════ */
  let panelEl, bodyEl, inputEl, sendBtn, opened = false;

  function crearFab() {
    const fab = document.createElement('button');
    fab.id = FAB_ID;
    fab.type = 'button';
    fab.setAttribute('aria-label', `Abrir chat con ${BOT_NAME}`);
    fab.innerHTML = `<img src="${ICON_SRC}" alt="${BOT_NAME}"><span class="sebIA-badge" id="sebIA-badge">0</span>`;
    fab.addEventListener('click', togglePanel);
    document.body.appendChild(fab);
    return fab;
  }

  function crearPanel() {
    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    panel.innerHTML = `
      <div class="sebIA-header">
        <img src="${ICON_SRC}" alt="${BOT_NAME}">
        <div class="sebIA-header-title">
          <strong>${BOT_NAME}</strong>
          <span>● En línea</span>
        </div>
        <div class="sebIA-header-actions">
          <button type="button" class="sebIA-icon-btn" id="sebIA-clear" aria-label="Borrar chat y recargar comandos" title="Borrar chat">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M3 6h18"></path>
              <path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
              <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path>
              <path d="M10 11v6"></path>
              <path d="M14 11v6"></path>
            </svg>
          </button>
          <button type="button" class="sebIA-icon-btn" id="sebIA-voice" aria-label="Activar lectura en voz alta" title="Leer respuestas en voz alta" aria-pressed="false">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M11 5 6 9H2v6h4l5 4V5Z"></path>
              <path d="M15.5 8.5a5 5 0 0 1 0 7"></path>
              <path d="M18.5 5.5a9 9 0 0 1 0 13"></path>
            </svg>
          </button>
          <button type="button" class="sebIA-close" aria-label="Cerrar chat">&times;</button>
        </div>
      </div>
      <div class="sebIA-body" id="sebIA-body"></div>
      <div class="sebIA-footer">
        <input type="text" class="sebIA-input" id="sebIA-input" placeholder="Escribe tu consulta..." autocomplete="off">
        <button type="button" class="sebIA-send" id="sebIA-send">Enviar</button>
      </div>
    `;
    document.body.appendChild(panel);

    bodyEl  = panel.querySelector('#sebIA-body');
    inputEl = panel.querySelector('#sebIA-input');
    sendBtn = panel.querySelector('#sebIA-send');

    panel.querySelector('.sebIA-close').addEventListener('click', togglePanel);
    panel.querySelector('#sebIA-clear').addEventListener('click', limpiarChat);
    panel.querySelector('#sebIA-voice').addEventListener('click', toggleVoz);
    sendBtn.addEventListener('click', onEnviar);
    inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') onEnviar();
    });

    return panel;
  }

  /* ── Ícono de papelera: borra la conversación y solo vuelve a
     mostrar la lista de comandos disponibles (NO repite el mensaje de
     presentación completo — ese es exclusivo de la primera vez que se
     abre el chat en la sesión, ver iniciarBienvenida). ── */
  function limpiarChat() {
    if (!bodyEl) return;
    bodyEl.innerHTML = '';
    agregarMensaje('Comandos disponibles:', 'bot');
    agregarListaComandos(ALL_COMMANDS);
  }

  /* ── Ícono de altavoz: activa/desactiva que Sebastián LEA en voz
     alta cada respuesta suya (Web Speech API — SpeechSynthesis, nativa
     del navegador, sin ninguna librería ni servicio externo). Si el
     navegador no soporta síntesis de voz, el botón simplemente no hace
     nada perceptible — nunca rompe el chat. ── */
  let vozActiva = false;
  let vozSeleccionada = null; // voz masculina en español ya resuelta (o null si no se encontró ninguna)

  function vozDisponible() {
    return typeof window !== 'undefined' && 'speechSynthesis' in window;
  }

  /* Nombres comunes de voces masculinas/femeninas que exponen los
     navegadores/SO más usados (Google, Microsoft/Edge, Apple) para
     español. No hay forma estándar de pedir "voz de hombre" en la Web
     Speech API — cada motor decide qué voces instala — así que se
     reconoce por nombre. Si ninguna coincide, cae a la primera voz en
     español que NO tenga nombre reconociblemente femenino; si tampoco
     hay, usa la voz por defecto del navegador (nunca deja de hablar
     por esto). */
  const NOMBRES_VOZ_MASCULINA = /jorge|pablo|diego|carlos|miguel|enrique|juan|raul|raúl|male|hombre|var[oó]n/i;
  const NOMBRES_VOZ_FEMENINA  = /m[oó]nica|helena|paulina|female|mujer|laura|mar[ií]a|sabina|elvira|conchita|luc[ií]a|paloma/i;

  function elegirVozMasculina(voces) {
    if (!Array.isArray(voces) || voces.length === 0) return null;
    const esVoces = voces.filter(v => /^es(-|_|$)/i.test(v.lang));
    const candidatas = esVoces.length > 0 ? esVoces : voces;

    return (
      candidatas.find(v => NOMBRES_VOZ_MASCULINA.test(v.name)) ||
      candidatas.find(v => !NOMBRES_VOZ_FEMENINA.test(v.name)) ||
      candidatas[0]
    );
  }

  /* Las voces se cargan de forma asíncrona en varios navegadores
     (Chrome en particular devuelve [] la primera vez) — se intenta de
     inmediato y también se re-intenta cuando el navegador avisa que ya
     están listas. Sin dependencias externas, 100% Web Speech API. */
  function cargarVozMasculina() {
    if (!vozDisponible()) return;
    const voces = window.speechSynthesis.getVoices();
    if (voces && voces.length > 0) vozSeleccionada = elegirVozMasculina(voces);
  }

  function hablar(texto) {
    if (!vozActiva || !vozDisponible() || !texto) return;
    try {
      window.speechSynthesis.cancel(); // no encimar lecturas si llegan mensajes seguidos
      const utter = new SpeechSynthesisUtterance(texto);
      utter.lang = 'es-ES';
      utter.rate = 1;
      // Tono más grave como respaldo SIEMPRE (además de elegir una voz
      // masculina cuando existe): así suena más varonil incluso en
      // navegadores que solo traen una voz en español.
      utter.pitch = 0.8;
      if (vozSeleccionada) utter.voice = vozSeleccionada;
      window.speechSynthesis.speak(utter);
    } catch (e) {
      console.error('[SebastianAI] No se pudo leer el mensaje en voz alta:', e);
    }
  }

  function toggleVoz() {
    if (!vozDisponible()) return; // botón inerte si el navegador no soporta síntesis de voz
    vozActiva = !vozActiva;
    const btn = document.getElementById('sebIA-voice');
    if (btn) {
      btn.classList.toggle('sebIA-voice-on', vozActiva);
      btn.setAttribute('aria-pressed', String(vozActiva));
      btn.title = vozActiva ? 'Silenciar respuestas' : 'Leer respuestas en voz alta';
    }
    if (!vozActiva && vozDisponible()) window.speechSynthesis.cancel();
  }

  function scrollAbajo() {
    bodyEl.scrollTop = bodyEl.scrollHeight;
  }

  function agregarMensaje(texto, tipo) {
    const msg = document.createElement('div');
    msg.className = `sebIA-msg ${tipo}`;
    msg.textContent = texto;
    bodyEl.appendChild(msg);
    scrollAbajo();
    if (tipo === 'bot') {
      if (!opened) incrementarNoLeidos();
      hablar(texto);
    }
    return msg;
  }

  function agregarListaComandos(lista = ALL_COMMANDS) {
    const wrap = document.createElement('div');
    wrap.className = 'sebIA-msg bot';
    const list = document.createElement('div');
    list.className = 'sebIA-cmd-list';
    lista.forEach(cmd => {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'sebIA-cmd-chip';
      chip.textContent = cmd;
      chip.addEventListener('click', () => enviarTexto(cmd));
      list.appendChild(chip);
    });
    wrap.appendChild(list);
    bodyEl.appendChild(wrap);
    scrollAbajo();
  }

  function mostrarTyping() {
    const t = document.createElement('div');
    t.className = 'sebIA-typing';
    t.id = 'sebIA-typing-indicator';
    t.innerHTML = '<span></span><span></span><span></span>';
    bodyEl.appendChild(t);
    scrollAbajo();
    return t;
  }

  function quitarTyping(el) {
    el?.remove();
  }

  /* Simula "escribiendo…" antes de mostrar la respuesta — puramente
     estético, no afecta el rendimiento del resto del dashboard (solo
     un setTimeout local a este módulo). */
  function responderConDelay(texto, delayMs, callback) {
    const typingEl = mostrarTyping();
    setTimeout(() => {
      quitarTyping(typingEl);
      if (callback) callback();
      else agregarMensaje(texto, 'bot');
    }, delayMs);
  }

  function enviarTexto(texto) {
    const limpio = (texto || '').trim();
    if (limpio === '') return;

    agregarMensaje(limpio, 'user');
    inputEl.value = '';

    const resultado = procesarComando(limpio);
    responderConDelay(null, 500 + Math.random() * 400, () => {
      agregarMensaje(resultado.text, 'bot');
      if (resultado.showCommandList) {
        agregarListaComandos(Array.isArray(resultado.showCommandList) ? resultado.showCommandList : ALL_COMMANDS);
      }
      if (resultado.kpiIds) {
        const recomendaciones = obtenerRecomendaciones(resultado.kpiIds);
        if (recomendaciones.length > 0) agregarOpcionesSugeridas(recomendaciones);
      }
    });
  }

  function onEnviar() {
    enviarTexto(inputEl.value);
  }

  /* Construye la llave real de sessionStorage añadiendo el usuario
     ACTUAL como sufijo a una llave "base" (WELCOME_FLAG_BASE,
     MENU_INFORMED_BASE, etc.).

     POR QUÉ ES NECESARIO: tanto "Cambiar Sesión" (SwitchSessionEngine.js)
     como el login/logout normal (SessionEngine en app.js) reemplazan
     al usuario activo y hacen un window.location.reload() — pero un
     reload NO vacía sessionStorage (solo se vacía al cerrar la
     pestaña). Si estas llaves fueran globales (sin el usuario), el
     flag de "ya se dio la bienvenida" quedaba en '1' de la sesión
     anterior y el nuevo usuario nunca recibía su bienvenida
     personalizada — el cambio de sesión no se "reconocía" como tal.
     Al incluir el usuario en la llave, cada usuario tiene su propio
     estado dentro de la misma pestaña, y un cambio de usuario
     siempre se comporta como una sesión nueva para Sebastián. */
  function _claveConSesion(base) {
    const usuario = (typeof Reader !== 'undefined' && Reader.sesionActiva) ? Reader.sesionActiva() : null;
    const sufijo = usuario ? String(usuario).trim().toUpperCase() : 'ANONIMO';
    return `${base}::${sufijo}`;
  }

  function yaSeDioBienvenida() {
    try {
      return sessionStorage.getItem(_claveConSesion(WELCOME_FLAG_BASE)) === '1';
    } catch (e) {
      return false;
    }
  }

  function marcarBienvenidaDada() {
    try {
      sessionStorage.setItem(_claveConSesion(WELCOME_FLAG_BASE), '1');
    } catch (e) { /* sessionStorage no disponible — no es crítico */ }
  }

  /* Flujo de bienvenida: un único mensaje de presentación (IGUAL para
     todos los usuarios — sin saludo personalizado por nombre) seguido
     inmediatamente por TODOS los comandos disponibles, y luego el
     recordatorio de "escribe AYUDA". Solo una vez por sesión de
     pestaña/usuario (ver _claveConSesion). */
  function iniciarBienvenida() {
    if (yaSeDioBienvenida()) return;
    marcarBienvenidaDada();

    responderConDelay(null, 500, () => {
      agregarMensaje(
        'SOY SEBASTIAN, TU ASISTENTE DE IA PERSONALIZADO, AQUI TE PROPORCIONO UNA SERIE DE COMANDO PARA QUE ME PIDAS LO QUE NECESITAS!',
        'bot'
      );
      agregarListaComandos(ALL_COMMANDS);

      responderConDelay(null, 700, () => {
        agregarMensaje('SI NECESITAS VER LOS COMANDOS NUEVAMENTE ESCRIBE AYUDA', 'bot');
      });
    });
  }

  let unreadCount = 0;

  function actualizarBadge() {
    const badge = document.getElementById('sebIA-badge');
    if (!badge) return;
    badge.textContent = unreadCount > 99 ? '99+' : String(unreadCount);
    badge.classList.toggle('sebIA-visible', unreadCount > 0);
  }

  function incrementarNoLeidos(cantidad = 1) {
    unreadCount += cantidad;
    actualizarBadge();
  }

  function limpiarNoLeidos() {
    unreadCount = 0;
    actualizarBadge();
  }

  function yaSeInformoMenu() {
    try {
      return sessionStorage.getItem(_claveConSesion(MENU_INFORMED_BASE)) === '1';
    } catch (e) {
      return false;
    }
  }

  function marcarMenuInformado() {
    try {
      sessionStorage.setItem(_claveConSesion(MENU_INFORMED_BASE), '1');
    } catch (e) { /* sessionStorage no disponible — no es crítico */ }
  }

  /* Se dispara cuando el usuario abre el offcanvas "Menú" del sidebar
     (#sidebarMenu, ver index.html). Solo la primera vez por sesión:
     si el chat ya está abierto, Sebastián explica ahí mismo qué hace
     cada botón; si está cerrado, deja el mensaje ya escrito en el
     historial (para cuando lo abra) — el badge numérico del botón
     botón flotante como aviso de que tiene algo nuevo que contarle.
     Nunca abre el panel por su cuenta — eso sería intrusivo. */
  function onMenuAbierto() {
    if (yaSeInformoMenu()) return;
    marcarMenuInformado();

    // Si el panel de chat todavía no existe (script cargándose fuera
    // de orden) no hay nada que hacer — se perdería el aviso, pero no
    // rompe nada del resto del dashboard.
    if (!bodyEl) return;

    agregarMensaje('Vi que abriste el Menú — te cuento rápido qué hace cada botón:', 'bot');
    agregarMensaje(explicarMenu(), 'bot');
    // El badge se actualiza solo (ver agregarMensaje) si el chat está cerrado.
  }

  /* Bootstrap mantiene un "focus trap" activo mientras el offcanvas
     #sidebarMenu sigue abierto: cada vez que el foco entra a un
     elemento que NO es descendiente del offcanvas, un listener interno
     de Bootstrap ('focusin' sobre `document`) lo redirige de vuelta
     adentro. Como nuestro botón flotante y el panel de chat viven
     fuera del offcanvas (son hijos directos de <body>), si el usuario
     abre el Menú y luego, SIN cerrarlo, hace clic en el chat, ese
     focus trap le robaba el foco al input y la ventana quedaba
     "bloqueada" (no se podía escribir ni hacer clic dentro).

     ANTES esto se resolvía cerrando el Menú automáticamente al abrir
     el chat — pero eso le quita al usuario la decisión de cuándo
     cerrar cada cosa, así que ya NO se hace.

     LA SOLUCIÓN, en su lugar: interceptamos el evento 'focusin' en
     fase de CAPTURA (que siempre se ejecuta ANTES que los listeners
     normales en fase de burbuja, sin importar el orden de carga de
     los scripts). Si el foco entra a un elemento dentro de nuestro
     propio botón/panel, detenemos la propagación inmediata del
     evento con stopImmediatePropagation(): así el listener de
     Bootstrap (registrado en fase de burbuja) NUNCA llega a
     ejecutarse para ese foco puntual, y el focus trap deja nuestro
     chat tranquilo sin necesidad de cerrar el offcanvas ni tocar
     nada de app.js/Bootstrap. El Menú permanece abierto (o cerrado)
     exactamente como el usuario lo dejó. */
  function neutralizarFocusTrapDelMenu() {
    document.addEventListener('focusin', (ev) => {
      const dentroDelChat =
        (panelEl && panelEl.contains(ev.target)) ||
        ev.target?.closest?.(`#${FAB_ID}`);
      if (!dentroDelChat) return;

      const menuEl = document.getElementById('sidebarMenu');
      if (!menuEl || !menuEl.classList.contains('show')) return; // Menú cerrado: nada que neutralizar

      ev.stopImmediatePropagation();
    }, true /* fase de captura — ver comentario arriba */);
  }

  function togglePanel() {
    opened = !opened;
    panelEl.classList.toggle('sebIA-open', opened);
    if (opened) {
      inputEl?.focus();
      limpiarNoLeidos();
      iniciarBienvenida();
    }
  }

  function init() {
    /* Aborta la inicialización si todavía no hay sesión activa (Nuevo).
       IMPORTANTE — por qué es un `setTimeout(init, 500)` y no un `return`
       seco: este init() se ejecuta UNA sola vez en DOMContentLoaded, es
       decir ANTES del login (el overlay de login vive en la misma página).
       El login normal (_confirmLogin en app.js) NO recarga la página —
       solo oculta el overlay — así que un `return` simple dejaría a
       Sebastián permanentemente desactivado incluso después de iniciar
       sesión con éxito. Reintentar cada 500ms es inofensivo (no crea
       ningún nodo ni listener mientras tanto) y, apenas
       Reader.sesionActiva() deja de ser null, el resto de init() corre
       UNA sola vez con total normalidad. */
    if (!Reader.sesionActiva()) {
      setTimeout(init, 500);
      return;
    }

    injectStyles();
    crearFab();
    panelEl = crearPanel();
    neutralizarFocusTrapDelMenu();

    // Resuelve la voz masculina para lectura en voz alta. En varios
    // navegadores (Chrome en particular) la lista de voces se carga de
    // forma asíncrona, por eso se intenta ya mismo y también se
    // re-intenta cuando el navegador avisa que están listas.
    cargarVozMasculina();
    if (vozDisponible()) {
      window.speechSynthesis.addEventListener('voiceschanged', cargarVozMasculina);
    }

    // Consciencia del botón "Menú" del sidebar (#sidebarMenu, offcanvas
    // de Bootstrap en index.html). Solo escucha — no interfiere con
    // Bootstrap ni con ningún otro módulo (LazyModals, etc.).
    document.getElementById('sidebarMenu')?.addEventListener('shown.bs.offcanvas', onMenuAbierto);

    // Apenas entra al dashboard, si todavía no se dio la bienvenida en
    // esta sesión, el botón flotante ya muestra "1" — como un mensaje
    // de WhatsApp esperando — para que el usuario sepa que Sebastián
    // tiene algo que decirle sin necesidad de abrir el chat primero.
    if (!yaSeDioBienvenida()) incrementarNoLeidos();

    inicializarAPIPublica();
  }

  /* ══════════════════════════════════════════════════════════
     API PÚBLICA (window.SebastianAI)
     ────────────────────────────────────────────────────────────
     Todo lo demás en este archivo es privado a propósito (closure
     de la IIFE). Esta es la ÚNICA puerta de entrada que se expone
     hacia afuera, pensada para módulos 100% independientes como
     KpiHelpEngine.js (ícono de ayuda "?" en las tarjetas KPI): les
     permite abrir el chat e inyectar contenido SIN que necesiten
     conocer ni tocar el DOM interno del panel (bodyEl, panelEl,
     etc.), exactamente el mismo patrón de aislamiento que ya usan
     AccessManager/TelegramEngine/LazyModals entre sí.
     ══════════════════════════════════════════════════════════ */
  function inicializarAPIPublica() {
    window.SebastianAI = {
      /**
       * Abre (si estaba cerrado) el panel de chat y, opcionalmente,
       * agrega un mensaje del bot con una lista de opciones sugeridas
       * (botones) debajo — pensado para el motor de recomendaciones
       * de KpiHelpEngine.js, pero utilizable por cualquier módulo.
       *
       * @param {string} mensaje - Texto que dirá Sebastián (tipo 'bot').
       * @param {Array<{label:string, comando:string}>} [opciones] -
       *   Sugerencias interactivas. Al hacer clic en una, su `comando`
       *   se envía al chat como si el usuario lo hubiera escrito (reutiliza
       *   enviarTexto(), así el motor de respuestas de Sebastián de
       *   siempre responde igual que si fuera texto libre).
       */
      abrirConMensaje(mensaje, opciones) {
        if (!opened) togglePanel(); // reutiliza la apertura normal (respeta badge, bienvenida, etc.)

        if (typeof mensaje === 'string' && mensaje.trim() !== '') {
          agregarMensaje(mensaje, 'bot');
        }

        if (Array.isArray(opciones) && opciones.length > 0) {
          agregarOpcionesSugeridas(opciones);
        }
      },

      /**
       * Abre (si estaba cerrado) el panel de chat y envía `texto` EXACTAMENTE
       * como si el usuario lo hubiera escrito y presionado Enter: reutiliza
       * enviarTexto() por completo, así la respuesta pasa por el mismo
       * procesarComando() de siempre (mismos datos ya filtrados por
       * AccessManager/vista restringida del usuario — ver Reader.text()) y,
       * si el comando trae kpiIds, dispara el motor de recomendaciones
       * (obtenerRecomendaciones) exactamente igual que si el usuario lo
       * hubiera pedido por texto libre.
       *
       * Pensado para botones externos que quieren "hacerle una pregunta"
       * a Sebastián sin duplicar ninguna lógica de respuesta (p. ej. el
       * botón "Análisis con IA" de la barra de filtros).
       *
       * @param {string} texto - Frase a enviar (debe calzar con alguna
       *   regex de COMMANDS para obtener una respuesta útil).
       */
      preguntar(texto) {
        if (!opened) togglePanel();
        enviarTexto(texto);
      },
    };
  }

  /* Renderiza una fila de botones-sugerencia dentro del chat, con el
     mismo lenguaje visual que agregarListaComandos() (misma clase
     .sebIA-cmd-chip) para no introducir un estilo nuevo. Cada botón,
     al pulsarse, envía su `comando` como si el usuario lo hubiera
     tecleado — así el motor de recomendaciones externo (KpiHelpEngine)
     no necesita conocer la lógica interna de respuestas de Sebastián,
     solo qué texto "preguntar". */
  function agregarOpcionesSugeridas(opciones) {
    const wrap = document.createElement('div');
    wrap.className = 'sebIA-msg bot';
    const list = document.createElement('div');
    list.className = 'sebIA-cmd-list';
    opciones.forEach(({ label, comando }) => {
      if (!label || !comando) return;
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'sebIA-cmd-chip';
      chip.textContent = label;
      chip.addEventListener('click', () => enviarTexto(comando));
      list.appendChild(chip);
    });
    wrap.appendChild(list);
    bodyEl.appendChild(wrap);
    scrollAbajo();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();


/* ---- KpiHelpEngine.js ---- */
/* ════════════════════════════════════════════════════════════
   KpiHelpEngine.js
   ────────────────────────────────────────────────────────────
   Módulo 100% INDEPENDIENTE para los íconos "?" que invocan a
   Sebastián IA:
     • En la esquina inferior derecha de cada tarjeta KPI
       (.kpi-help-btn, data-kpi-id — ver index.html/style.css).
     • Junto a la barra de búsqueda de cada sección de tabla:
       Personas, Excluidos, Nuevos y Monitor de Ausencias
       (.table-help-btn, data-section-id — ver index.html/style.css).

   Al hacer clic:
     1. Lee el contexto del botón desde sus atributos data-*
        (data-kpi-id + data-user-view, o data-section-id).
     2. Arma un mensaje: para KPIs, una explicación (qué es, para
        qué sirve); para secciones de tabla, un ANÁLISIS con los
        números reales ya renderizados en el DOM (conteos de
        TableEngine/AbsenceEngine).
     3. Genera un array de "Nuevas Opciones" (chips) — por
        decisionEngine para KPIs, o una lista fija por sección.
     4. Invoca la API pública de Sebastián IA (window.SebastianAI,
        expuesta en SebastianAI.js) para abrir el chat con ese
        mensaje + esas opciones.

   Qué NO hace (a propósito):
   - No modifica SebastianAI.js, AccessManager.js, SecurityConfig.js,
     USUARIOS.JS ni ninguna lógica de negocio de app.js. Solo LEE el
     usuario en sesión (misma llave de sessionStorage que ya usa el
     resto del proyecto) y los contadores ya pintados en el DOM para
     personalizar/armar el mensaje.
   - No calcula KPIs reales: los "datos" que usa el motor de
     recomendaciones de las tarjetas KPI son simulados (ver
     SIMULATED_DATA_STATE más abajo) — están marcados con comentarios
     "🔌 CONECTAR AQUÍ" en los puntos exactos donde deberías
     sustituirlos por tus valores reales. Las secciones de tabla, en
     cambio, SÍ leen datos reales (los contadores del DOM), porque
     esos números ya existen — no hace falta simularlos.
   - No abre el modal de personas ni ninguna tabla — eso sigue siendo
     EXCLUSIVO de hacer clic directamente en la tarjeta/sección.

   Cómo usarlo:
   Solo agrega este script en index.html, DESPUÉS de SebastianAI.js
   (necesita `window.SebastianAI` ya expuesto):
       <script src="KpiHelpEngine.js"></script>
   ════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  /* Debe coincidir con SessionEngine.STORAGE_KEY (app.js) — se lee
     directamente de sessionStorage (no de SessionEngine) para que
     este módulo siga siendo 100% independiente, igual que hace
     SwitchSessionEngine.js. */
  const SESSION_STORAGE_KEY = 'ccrm_dashboard_user';

  function usuarioActual() {
    try {
      return sessionStorage.getItem(SESSION_STORAGE_KEY) || null;
    } catch (e) {
      return null;
    }
  }

  /* ══════════════════════════════════════════════════════════
     1) DICCIONARIO DE EXPLICACIONES POR KPI
     ────────────────────────────────────────────────────────────
     Clave = data-kpi-id de la tarjeta (ver index.html). Cada
     entrada describe QUÉ ES el KPI y PARA QUÉ SIRVE — esto es lo
     que arma el primer mensaje que Sebastián le muestra al usuario.
     ══════════════════════════════════════════════════════════ */
  const KPI_EXPLICACIONES = {
    'total':           { titulo: 'Total Registrados',          que: 'la cantidad total de personas registradas en el reporte cargado.', paraQue: 'darte una foto general del tamaño de tu grupo/iglesia antes de mirar el detalle.' },
    'celula-si':       { titulo: 'Asistencia a Célula',        que: 'cuántas personas SÍ asistieron a célula en el periodo del reporte.', paraQue: 'medir el compromiso con la reunión de célula, semana a semana.' },
    'celula-no':       { titulo: 'Inasistencia a Célula',      que: 'cuántas personas NO asistieron a célula en el periodo del reporte.', paraQue: 'detectar a tiempo a quienes se están alejando de la célula.' },
    'servicio-si':     { titulo: 'Asistencia a Servicio',      que: 'cuántas personas SÍ asistieron al servicio/culto en el periodo.', paraQue: 'medir el compromiso con el servicio general, aparte de la célula.' },
    'servicio-no':     { titulo: 'Inasistencia a Servicio',    que: 'cuántas personas NO asistieron al servicio/culto en el periodo.', paraQue: 'identificar a quienes podrían necesitar seguimiento pastoral.' },
    'ambos-si':        { titulo: 'Asistió a Ambos',            que: 'cuántas personas asistieron TANTO a célula COMO a servicio.', paraQue: 'reconocer al núcleo más comprometido del grupo.' },
    'ambos-no':        { titulo: 'Ausente en Ambos',           que: 'cuántas personas NO asistieron ni a célula ni a servicio.', paraQue: 'priorizar el seguimiento — son quienes más riesgo de alejarse tienen.' },
    'nuevos-celula':   { titulo: 'Nuevos en Célula',           que: 'cuántas personas asistieron a célula por primera vez en este periodo.', paraQue: 'medir qué tan bien está creciendo tu célula con gente nueva.' },
    'nuevos-servicio': { titulo: 'Nuevos en Servicio',         que: 'cuántas personas asistieron al servicio por primera vez en este periodo.', paraQue: 'medir el alcance de nuevas visitas al servicio general.' },
    'pct-general':     { titulo: '% Asistencia General',       que: 'el porcentaje de asistencia general (célula + servicio) sobre el total de registrados.', paraQue: 'tener un solo número que resuma la salud general del grupo.' },
    'pct-celula':      { titulo: '% Asistencia Célula',        que: 'el porcentaje de asistencia SOLO a célula sobre el total de registrados.', paraQue: 'comparar qué tan fuerte está la célula frente al servicio general.' },
    'pct-servicio':    { titulo: '% Asistencia Servicio',      que: 'el porcentaje de asistencia SOLO a servicio sobre el total de registrados.', paraQue: 'comparar qué tan fuerte está el servicio general frente a la célula.' },
  };

  /* ══════════════════════════════════════════════════════════
     1-B) SECCIONES DE TABLA (Personas / Excluidos / Nuevos /
     Monitor de Ausencias) — ver .table-help-btn en index.html,
     dentro de cada .table-search-bar.

     A diferencia de los KPI (que solo describen QUÉ ES el
     indicador), aquí cada `analizar(usuario)` LEE en vivo los
     contadores que TableEngine/AbsenceEngine ya pintaron en el DOM
     (los mismos #countXxx / #ausXxxCount que usa SebastianAI.js) y
     arma una respuesta con el dato real — nunca abre la tabla ni el
     modal de personas, eso sigue siendo exclusivo de hacer clic en
     la sección misma. Como solo lee texto ya renderizado (que a su
     vez ya pasó por AccessManager), respeta el mismo RBAC que el
     resto del proyecto sin tener que reimplementarlo aquí.
     ══════════════════════════════════════════════════════════ */
  function _texto(id, fallback) {
    const el = document.getElementById(id);
    if (!el) return fallback;
    const t = (el.textContent || '').trim();
    return t !== '' ? t : fallback;
  }

  const SECCIONES_TABLA = {
    personas: {
      titulo: 'Personas',
      analizar: () => `esta sección lista a las personas activas que cumplen tus filtros actuales: ${_texto('countPersonas', 'sin datos todavía')}. Es la fuente completa detrás de casi todos los KPIs del dashboard.`,
    },
    excluidos: {
      titulo: 'Excluidos',
      analizar: () => `esta sección lista a las personas marcadas como excluidas del conteo regular: ${_texto('countExcluidos', 'sin datos todavía')}. Útil para revisar quiénes están fuera del reporte activo y por qué.`,
    },
    nuevos: {
      titulo: 'Nuevos',
      analizar: () => `esta sección lista a las personas nuevas del periodo: ${_texto('countNuevos', 'sin datos todavía')}. El desglose entre célula y servicio ya viene incluido en ese conteo.`,
    },
    ausencias: {
      titulo: 'Monitor de Ausencias',
      analizar: () => {
        const normal = _texto('ausNormalCount', '0');
        const vigilar = _texto('ausWatchCount', '0');
        const alerta = _texto('ausWarnCount', '0');
        const critico = _texto('ausCritCount', '0');
        return `este monitor clasifica a cada persona por su racha de inasistencia: ${_texto('countAusencias', 'sin datos todavía')} — Normal: ${normal} · A vigilar: ${vigilar} · Alerta: ${alerta} · Crítico: ${critico}. Prioriza revisar primero a quienes están en Crítico.`;
      },
    },
  };

  const OPCIONES_FIJAS_POR_SECCION = {
    personas:  [{ label: '📊 Ver resumen completo', comando: 'RESUMEN' }],
    excluidos: [{ label: '📋 Ver histórico', comando: 'HISTORICO' }],
    nuevos:    [
      { label: '🙌 Detalle nuevos en Célula', comando: 'DETALLE NUEVOS CELULA' },
      { label: '🙌 Detalle nuevos en Servicio', comando: 'DETALLE NUEVOS SERVICIO' },
    ],
    ausencias: [
      { label: '⚠️ Ver en riesgo (Ambos NO)', comando: 'DETALLE AMBOS NO' },
      { label: '🚨 Plan de seguimiento', comando: 'PLAN DE SEGUIMIENTO' },
    ],
  };

  function armarMensajeSeccion(sectionId, usuario) {
    const info = SECCIONES_TABLA[sectionId];
    const saludo = usuario ? `${usuario}, ` : '';

    if (!info) {
      return `${saludo}esa sección todavía no tiene un análisis configurado. Avísale al equipo técnico para agregarla.`;
    }

    return `${saludo}${info.analizar()}`;
  }

  /* ══════════════════════════════════════════════════════════
     2) ESTADO DE DATOS (SIMULADO)
     ────────────────────────────────────────────────────────────
     🔌 CONECTAR AQUÍ: en producción, reemplaza esta función por tu
     lógica real — por ejemplo, leer el texto ya calculado en el DOM
     (los mismos <div id="kpiXxx"> que pinta KPIEngine en app.js) o,
     mejor aún, consultar directamente tu DataStore/KPIEngine si
     expone esos números en memoria. La firma (qué devuelve) debe
     mantenerse igual para que el motor de recomendaciones de abajo
     siga funcionando sin cambios.
     ══════════════════════════════════════════════════════════ */
  function leerEstadoSimuladoDelKPI(kpiId) {
    // 🔌 CONECTAR AQUÍ (opción real, ejemplo):
    //   const el = document.getElementById(idMap[kpiId]);
    //   const valor = el ? parseInt(el.textContent, 10) : null;
    //
    // Por ahora, simulamos una tendencia aleatoria estable por
    // sesión de página (no cambia en cada clic) solo para poder
    // mostrar cómo el motor de recomendaciones reacciona distinto
    // según el estado de los datos:
    if (!leerEstadoSimuladoDelKPI._cache) leerEstadoSimuladoDelKPI._cache = {};
    if (!(kpiId in leerEstadoSimuladoDelKPI._cache)) {
      const tendencias = ['subiendo', 'bajando', 'estable'];
      leerEstadoSimuladoDelKPI._cache[kpiId] = tendencias[Math.floor(Math.random() * tendencias.length)];
    }
    return { tendencia: leerEstadoSimuladoDelKPI._cache[kpiId] };
  }

  /* ══════════════════════════════════════════════════════════
     3) MOTOR DE RECOMENDACIONES (decisionEngine)
     ────────────────────────────────────────────────────────────
     Devuelve un array de "Nuevas Opciones" — cada una es un botón
     que se inyecta en el chat de Sebastián. `label` es lo que ve el
     usuario; `comando` es el texto que se envía al chat al hacer
     clic (Sebastián lo procesa como si el usuario lo hubiera
     escrito — ver window.SebastianAI.abrirConMensaje en
     SebastianAI.js).

     Combina dos cosas:
       a) Opciones FIJAS por kpi-id (siempre aplican a ese KPI).
       b) Opciones CONDICIONALES según el estado simulado de los
          datos (ej: si la tendencia es "bajando", sugiere opciones
          de alerta/seguimiento en vez de solo informativas).
     ══════════════════════════════════════════════════════════ */
  const OPCIONES_FIJAS_POR_KPI = {
    'total':           [{ label: '📈 Ver tendencia', comando: 'TENDENCIA' }],
    'celula-si':       [{ label: '📊 Comparar con servicio', comando: 'COMPARAR' }],
    'celula-no':       [{ label: '📉 Ver quiénes faltaron', comando: 'DETALLE CELULA' }],
    'servicio-si':     [{ label: '📊 Comparar con célula', comando: 'COMPARAR' }],
    'servicio-no':     [{ label: '📉 Ver quiénes faltaron', comando: 'DETALLE SERVICIO' }],
    'ambos-si':        [{ label: '⭐ Ver core comprometido', comando: 'DETALLE AMBOS SI' }],
    'ambos-no':        [{ label: '⚠️ Ver en riesgo', comando: 'DETALLE AMBOS NO' }],
    'nuevos-celula':   [{ label: '🙌 Ver nuevos de célula', comando: 'DETALLE NUEVOS CELULA' }],
    'nuevos-servicio': [{ label: '🙌 Ver nuevos de servicio', comando: 'DETALLE NUEVOS SERVICIO' }],
    'pct-general':     [{ label: '📈 Ver tendencia general', comando: 'TENDENCIA' }],
    'pct-celula':      [{ label: '📈 Ver tendencia célula', comando: 'TENDENCIA' }],
    'pct-servicio':    [{ label: '📈 Ver tendencia servicio', comando: 'TENDENCIA' }],
  };

  /**
   * Motor de decisiones: combina reglas fijas + condicionales según
   * `estado` (ver leerEstadoSimuladoDelKPI) para el kpi-id dado.
   *
   * @param {string} kpiId
   * @param {{tendencia: 'subiendo'|'bajando'|'estable'}} estado
   * @returns {Array<{label:string, comando:string}>}
   */
  function decisionEngine(kpiId, estado) {
    const opciones = [...(OPCIONES_FIJAS_POR_KPI[kpiId] || [])];

    // Regla condicional genérica: KPIs de tipo "inasistencia"/"ausente"
    // que además vienen "subiendo" → priorizar acción de seguimiento.
    const esKpiDeAlerta = kpiId.includes('no') || kpiId === 'ambos-no';

    if (estado.tendencia === 'subiendo' && esKpiDeAlerta) {
      opciones.unshift({ label: '🚨 Sugerir plan de seguimiento', comando: 'PLAN DE SEGUIMIENTO' });
    } else if (estado.tendencia === 'bajando' && !esKpiDeAlerta) {
      opciones.unshift({ label: '🔎 ¿Por qué está bajando?', comando: 'ANALIZAR CAIDA' });
    } else if (estado.tendencia === 'estable') {
      opciones.push({ label: '✅ Todo estable, ¿algo más?', comando: 'AYUDA' });
    }

    // 🔌 CONECTAR AQUÍ: si tu proyecto ya tiene comandos reales de
    // Sebastián para "TENDENCIA", "COMPARAR", etc. (ver ALL_COMMANDS
    // dentro de SebastianAI.js), estos `comando` deben coincidir con
    // esos textos exactos para que el bot los reconozca al enviarse.

    // Opción permanente (Nuevo): siempre disponible al final de la
    // lista, sin importar el estado de la tendencia — dispara el
    // "Reporte Ejecutivo" completo vía Sebastián IA (ver
    // FRASE_ANALISIS_IA en SebastianAI.js, que reconoce este mismo
    // texto de comando y responde con analizarDashboardCompleto()).
    opciones.push({ label: '📝 Generar Reporte Ejecutivo', comando: 'REPORTE EJECUTIVO' });

    return opciones;
  }

  /* ══════════════════════════════════════════════════════════
     4) ARMADO DEL MENSAJE + INVOCACIÓN A SEBASTIÁN
     ══════════════════════════════════════════════════════════ */
  function armarMensajeExplicativo(kpiId, userView, usuario) {
    const info = KPI_EXPLICACIONES[kpiId];
    const saludo = usuario ? `${usuario}, ` : '';

    if (!info) {
      // Fail-safe: kpi-id desconocido (p. ej. si el HTML cambia y
      // este archivo no se actualizó) — igual respondemos algo útil
      // en vez de romper el flujo del chat.
      return `${saludo}este indicador (vista: ${userView || 'general'}) todavía no tiene una explicación configurada. Avísale al equipo técnico para agregarla.`;
    }

    return `${saludo}este KPI es "${info.titulo}": muestra ${info.que} Sirve para ${info.paraQue}`;
  }

  function onClickAyuda(ev) {
    const btnKpi = ev.target.closest('.kpi-help-btn');
    const btnSeccion = ev.target.closest('.table-help-btn');
    const btn = btnKpi || btnSeccion;
    if (!btn) return;

    // Evita que el clic burbujee hacia comportamientos del propio
    // .kpi-card (abrir el modal de personas) — ES EL MOTIVO por el
    // que este listener corre en fase de CAPTURA (ver init() más
    // abajo): el listener que abre ese modal está pegado directamente
    // sobre cada .kpi-card, más cerca del botón "?" en el árbol del
    // DOM que `document`, así que en fase de burbuja normal siempre
    // se dispara primero. En captura, `document` se visita ANTES de
    // que el clic descienda hasta la tarjeta, así que interceptamos
    // aquí y el modal nunca llega a abrirse.
    ev.preventDefault();
    ev.stopPropagation();

    if (typeof window.SebastianAI === 'undefined' || typeof window.SebastianAI.abrirConMensaje !== 'function') {
      console.error('[KpiHelpEngine] window.SebastianAI no está disponible — ¿se cargó SebastianAI.js antes que este script?');
      return;
    }

    const usuario = usuarioActual();

    if (btnSeccion) {
      // ── Botón de una SECCIÓN DE TABLA (Personas/Excluidos/Nuevos/
      //    Monitor de Ausencias) — ver .table-help-btn en index.html.
      const sectionId = btnSeccion.dataset.sectionId;
      if (!sectionId) {
        console.error('[KpiHelpEngine] Botón de ayuda de tabla sin data-section-id — no se puede continuar.');
        return;
      }
      const mensaje = armarMensajeSeccion(sectionId, usuario);
      const opciones = OPCIONES_FIJAS_POR_SECCION[sectionId] || [];
      window.SebastianAI.abrirConMensaje(mensaje, opciones);
      return;
    }

    // ── Botón de una tarjeta KPI (comportamiento original, sin cambios) ──
    const kpiId = btnKpi.dataset.kpiId;
    const userView = btnKpi.dataset.userView;

    if (!kpiId) {
      console.error('[KpiHelpEngine] Botón de ayuda sin data-kpi-id — no se puede continuar.');
      return;
    }

    const estado = leerEstadoSimuladoDelKPI(kpiId); // 🔌 CONECTAR AQUÍ (ver función arriba)
    const mensaje = armarMensajeExplicativo(kpiId, userView, usuario);
    const opciones = decisionEngine(kpiId, estado);

    window.SebastianAI.abrirConMensaje(mensaje, opciones);
  }

  function init() {
    // Fase de CAPTURA (true) — MISMA técnica que ya usa LazyModals.js en
    // este proyecto. Es imprescindible aquí: el listener que abre el
    // modal de personas está pegado directamente sobre cada .kpi-card
    // (ver UIController.bindKPICardClicks en app.js), no sobre
    // `document`. En fase de burbuja (el comportamiento por defecto),
    // ese listener de la tarjeta SIEMPRE se dispara primero porque está
    // más cerca del botón "?" en el árbol del DOM que `document` — para
    // cuando nuestro preventDefault()/stopPropagation() de más abajo se
    // ejecutaba, el modal de personas ya se había abierto. Escuchando en
    // fase de captura, `document` se visita ANTES de que el clic
    // descienda hasta la tarjeta, así que interceptamos y detenemos el
    // evento ahí mismo — el clic nunca llega a bindKPICardClicks() y el
    // modal de personas no se abre. El resto de la tarjeta (fuera del
    // botón "?") sigue abriendo la lista de personas exactamente igual
    // que antes, porque onClickAyuda() solo actúa cuando el clic
    // realmente vino de `.kpi-help-btn` o `.table-help-btn` (ver
    // closest() más arriba).
    document.addEventListener('click', onClickAyuda, true);

    inicializarAPIPublica();
  }

  /* ══════════════════════════════════════════════════════════
     API PÚBLICA (window.KpiHelpEngine)
     ────────────────────────────────────────────────────────────
     Mismo patrón de aislamiento que window.SebastianAI: la ÚNICA
     puerta de entrada hacia afuera de esta IIFE. Permite que
     SebastianAI.js reutilice el MISMO decisionEngine (reglas fijas +
     condicionales por kpi-id) desde CUALQUIER respuesta del chat —
     no solo desde el clic en el botón "?" — para que el motor de
     recomendaciones esté activo también cuando el usuario pregunta
     por texto libre ("¿cómo van las células?", etc.). No se duplica
     lógica: es la misma función decisionEngine() de siempre. */
  function inicializarAPIPublica() {
    window.KpiHelpEngine = {
      /**
       * @param {string} kpiId - Una de las claves de KPI_EXPLICACIONES
       *   ('total', 'celula-si', 'celula-no', 'servicio-si', ...).
       * @returns {Array<{label:string, comando:string}>} Recomendaciones
       *   generadas por decisionEngine para ese KPI en su estado actual.
       *   Devuelve [] si el kpiId no existe — nunca lanza error.
       */
      getRecomendaciones(kpiId) {
        if (!kpiId || !KPI_EXPLICACIONES[kpiId]) return [];
        const estado = leerEstadoSimuladoDelKPI(kpiId);
        return decisionEngine(kpiId, estado);
      },
    };
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();


/* ---- ReporteEjecutivoEngine.js ---- */
/* ════════════════════════════════════════════════════════════
   ReporteEjecutivoEngine.js
   ────────────────────────────────────────────────────────────
   Módulo 100% INDEPENDIENTE. El botón "Reporte Ejecutivo"
   (#btnReporteEjecutivo, menú lateral) abre una ventana superpuesta
   propia — un "Reporte Ejecutivo Detallado" — construida 100% por
   JavaScript (sin tocar index.html). CERO dependencia de
   SebastianAI en todo el flujo, y CERO navegación/comandos: el
   Plan de Seguimiento es contenido de informe (texto de lectura),
   no botones que saltan a otra parte del dashboard.

   RBAC / SEGURIDAD:
   Este módulo NUNCA toca DataStore ni ExcelParser directamente, y
   NUNCA aplica sus propias reglas de permisos. Todo lo que lee — los
   textos de KPIs en el DOM, UIController._lastFilteredRecords y
   AbsenceEngine._currentData — es siempre el MISMO dataset que ya
   pasó por AccessManager.applyFilter() para pintar el resto del
   dashboard. Si el usuario no tiene permiso para ver un dato, ese
   dato no está en esas fuentes, así que este motor tampoco puede
   mostrarlo.

   ESTILO:
   Sus estilos viven en style.css (bloque "ReporteEjecutivoEngine"),
   consumen EXCLUSIVAMENTE las variables CSS nativas del proyecto
   (--bg-card, --gold, --danger, etc.), así que respetan el tema
   oscuro/claro automáticamente.

   ANÁLISIS Y PLAN DE SEGUIMIENTO:
   Ya NO usa window.KpiHelpEngine.getRecomendaciones() — esa función
   depende de leerEstadoSimuladoDelKPI() (tendencia ALEATORIA, ver
   "🔌 CONECTAR AQUÍ" en ese archivo), por eso antes solo devolvía
   frases genéricas. En su lugar, este módulo analiza directamente
   los registros reales ya filtrados (nombres, grupo, teléfono,
   días sin asistir) para armar un informe con: personas concretas
   a las que hacerle seguimiento (con enlace de WhatsApp si hay
   teléfono), un diagnóstico de causa probable (¿concentrada en un
   grupo o repartida? ¿solo célula, solo servicio, o ambos?) y un
   plan de acción sugerido. Todo texto de informe — nada de botones
   que ejecuten comandos ni que naveguen a otra parte del dashboard.
   ════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  const STYLE_ID    = 'reporteEjecutivoStyles';
  const OVERLAY_ID  = 'reporteEjecutivoOverlay';
  const BTN_ID      = 'btnReporteEjecutivo';

  /* ── Umbrales de estado (mismos criterios en todo el reporte) ──
     % de INASISTENCIA (NO) sobre el total de esa categoría:
       < 15%  → ✅ Saludable
       15–30% → ⚠️ Atención
       > 30%  → 🔴 Crítico                                          */
  const UMBRAL_ALERTA   = 15;
  const UMBRAL_CRITICO  = 30;

  /* ══════════════════════════════════════════════════════════
     LECTOR DE DOM (RBAC vía DOM) — mismo patrón que Reader en
     SebastianAI.js: nunca lanza error, siempre hay un fallback,
     y detecta el estado "recalculando" para no mostrar basura.
     ══════════════════════════════════════════════════════════ */
  const ReaderRE = {
    /** Texto crudo del nodo, o `fallback` si no existe / está vacío. */
    text(id, fallback = '—') {
      const el = document.getElementById(id);
      if (!el) return fallback;
      const t = (el.textContent || '').trim();
      return t !== '' ? t : fallback;
    },

    /** true si el nodo está en medio del overlay "Recalculando…"
        (ver UIController._setKpisRecalculando en app.js). */
    isRecalculando(id) {
      const el = document.getElementById(id);
      return !!el && el.classList.contains('kpi-recalculando');
    },

    /** Extrae un número (entero o decimal) del texto de un nodo,
        tolerando '%', espacios, comas de miles, etc. Devuelve `null`
        si no se pudo interpretar como número (nunca NaN silencioso). */
    numero(id) {
      const raw = this.text(id, '');
      if (raw === '' || raw === '—') return null;
      const limpio = raw.replace(/\./g, '').replace(',', '.').match(/-?\d+(\.\d+)?/);
      if (!limpio) return null;
      const n = parseFloat(limpio[0]);
      return Number.isFinite(n) ? n : null;
    },
  };

  /** true si CUALQUIERA de los ids dados está recalculando —
      usado para mostrar un aviso en vez de datos a medio pintar. */
  function algunoRecalculando(ids) {
    return ids.some(id => ReaderRE.isRecalculando(id));
  }

  /** Clasifica un % de inasistencia en 'ok' | 'alerta' | 'critico'. */
  function clasificar(pctNo) {
    if (pctNo === null) return 'ok';
    if (pctNo > UMBRAL_CRITICO) return 'critico';
    if (pctNo >= UMBRAL_ALERTA) return 'alerta';
    return 'ok';
  }

  const ICONO_ESTADO = { ok: '✅', alerta: '⚠️', critico: '🔴' };
  const LABEL_ESTADO = { ok: 'Saludable', alerta: 'Atención', critico: 'Crítico' };

  /* ══════════════════════════════════════════════════════════
     ESTILOS — inyectados una sola vez, 100% variables nativas.
     ══════════════════════════════════════════════════════════ */
  function inyectarEstilos() {
    /* CSS ahora vive en style.css — ya no se inyecta por JS. */
    return;
  }

  /* ══════════════════════════════════════════════════════════
     RECOLECCIÓN DE DATOS (100% desde el DOM)
     ══════════════════════════════════════════════════════════ */
  function recolectarDatos() {
    return {
      total:        ReaderRE.text('kpiTotal'),
      pctGeneral:   ReaderRE.text('kpiPctGeneral'),

      celulaSI:     ReaderRE.text('kpiCelulasSI'),
      celulaSIPct:  ReaderRE.text('kpiCelulasSIPct'),
      celulaNO:     ReaderRE.text('kpiCelulasNO'),
      celulaNOPct:  ReaderRE.numero('kpiCelulasNOPct'),
      celulaNOPctTexto: ReaderRE.text('kpiCelulasNOPct'),

      servicioSI:     ReaderRE.text('kpiServicioSI'),
      servicioSIPct:  ReaderRE.text('kpiServicioSIPct'),
      servicioNO:     ReaderRE.text('kpiServicioNO'),
      servicioNOPct:  ReaderRE.numero('kpiServicioNOPct'),
      servicioNOPctTexto: ReaderRE.text('kpiServicioNOPct'),

      ambosSI:     ReaderRE.text('kpiAmbosSI'),
      ambosSIPct:  ReaderRE.text('kpiAmbosSIPct'),
      ambosNO:     ReaderRE.text('kpiAmbosNO'),
      ambosNOPct:  ReaderRE.numero('kpiAmbosNOPct'),
      ambosNOPctTexto: ReaderRE.text('kpiAmbosNOPct'),

      nuevosCelula:    ReaderRE.text('kpiNuevosCelula'),
      nuevosServicio:  ReaderRE.text('kpiNuevosServicio'),

      ausNormal: ReaderRE.text('ausNormalCount', '0'),
      ausWatch:  ReaderRE.text('ausWatchCount', '0'),
      ausWarn:   ReaderRE.text('ausWarnCount', '0'),
      ausCrit:   ReaderRE.text('ausCritCount', '0'),
      ausCritNum: ReaderRE.numero('ausCritCount') || 0,

      recalculando: algunoRecalculando([
        'kpiTotal', 'kpiPctGeneral',
        'kpiCelulasSI', 'kpiCelulasNO', 'kpiServicioSI', 'kpiServicioNO',
        'kpiAmbosSI', 'kpiAmbosNO', 'kpiNuevosCelula', 'kpiNuevosServicio',
      ]),
    };
  }

  /* ══════════════════════════════════════════════════════════
     MOTOR DE ANÁLISIS REAL — nombres, causas y plan de acción
     ────────────────────────────────────────────────────────────
     A propósito, esto YA NO usa window.KpiHelpEngine.
     getRecomendaciones(): esa función basa sus recomendaciones en
     leerEstadoSimuladoDelKPI(), que devuelve una TENDENCIA
     ALEATORIA (Math.random(), ver comentario "🔌 CONECTAR AQUÍ" en
     ese archivo) — por eso el reporte solo mostraba frases
     genéricas. Todo lo de aquí abajo se calcula sobre los
     registros REALES ya filtrados por RBAC:
       • UIController._lastFilteredRecords — mismo dataset que
         alimenta las tarjetas KPI (bindKPICardClicks).
       • AbsenceEngine._currentData — mismo dataset, ya procesado
         y ordenado, que alimenta el Monitor de Ausencias.
     No se inventa ni se simula nada: si el Excel no trae esos
     datos, las listas simplemente salen vacías.
     ══════════════════════════════════════════════════════════ */

  function registrosReales() {
    return (typeof UIController !== 'undefined' && Array.isArray(UIController._lastFilteredRecords))
      ? UIController._lastFilteredRecords
      : [];
  }

  function ausenciasReales() {
    return (typeof AbsenceEngine !== 'undefined' && Array.isArray(AbsenceEngine._currentData))
      ? AbsenceEngine._currentData
      : [];
  }

  /* Nombres reales que requieren seguimiento YA — nivel Crítico
     (> 4 semanas sin asistir), en el mismo orden que ya usa
     AbsenceEngine.process() (crítico primero, luego por días desc). */
  function personasPrioritarias(limite = 8) {
    return ausenciasReales().filter(r => r.nivel === 'critical').slice(0, limite);
  }

  /* Nombres reales en riesgo de volverse crítico (2–4 semanas). */
  function personasEnAdvertencia(limite = 6) {
    return ausenciasReales().filter(r => r.nivel === 'warn').slice(0, limite);
  }

  /* CAUSA #1 — ¿la inasistencia está concentrada en un grupo/
     liderazgo puntual, o repartida entre varios? Concentrada ⇒
     probable causa local (horario, actividad paralela, conflicto
     con ese líder/célula). Repartida ⇒ probable causa general
     (fecha, feriado, clima) que afectó a varios grupos por igual. */
  function analizarConcentracionPorGrupo(metricKey) {
    const subset = (typeof KPIEngine !== 'undefined')
      ? KPIEngine.getRecordsByMetric(registrosReales(), metricKey)
      : [];
    if (subset.length === 0) return null;

    const porGrupo = {};
    subset.forEach(r => {
      const g = (r.grupo || '').trim() || 'Sin grupo asignado';
      porGrupo[g] = (porGrupo[g] || 0) + 1;
    });
    const distribucion = Object.entries(porGrupo).sort((a, b) => b[1] - a[1]);
    const [grupoTop, cantidadTop] = distribucion[0];
    const pct = Math.round((cantidadTop / subset.length) * 100);

    return { total: subset.length, grupoTop, cantidadTop, pct, gruposDistintos: distribucion.length };
  }

  /* CAUSA #2 — ¿faltan solo a célula, solo a servicio, o a ambos?
     Solo servicio ⇒ posible conflicto de horario con el culto.
     Solo célula ⇒ posible fricción con esa célula/grupo puntual.
     Ambos ⇒ riesgo real de desconexión/abandono, requiere contacto
     pastoral directo, no solo administrativo. */
  function analizarPatronFaltas() {
    let soloCelula = 0, soloServicio = 0, ambos = 0;
    registrosReales().forEach(r => {
      const faltaCelula   = r.celula === 'NO';
      const faltaServicio = r.servicio === 'NO';
      if (faltaCelula && faltaServicio) ambos++;
      else if (faltaCelula) soloCelula++;
      else if (faltaServicio) soloServicio++;
    });
    return { soloCelula, soloServicio, ambos };
  }

  /* Bloque de texto de informe: título + párrafo explicativo. */
  function crearBloqueTexto(titulo, texto) {
    const div = document.createElement('div');
    div.className = 're-analisis-bloque';
    div.innerHTML = `<div class="re-analisis-titulo">${titulo}</div><p class="re-analisis-texto">${texto}</p>`;
    return div;
  }

  /* Lista de PERSONAS REALES (no comandos, no navegación): nombre,
     grupo, tiempo exacto sin asistir y, si hay teléfono, el mismo
     enlace directo de WhatsApp que ya usa el resto del dashboard
     (TableEngine.waLink) — para que el seguimiento sea accionable
     sin depender de ningún chat ni IA. */
  function crearListaPersonas(personas) {
    const list = document.createElement('ul');
    list.className = 're-reco-list';
    personas.forEach(r => {
      const li = document.createElement('li');
      li.className = 're-reco-item';
      const grupoTxt = (r.grupo || '').trim();
      const contacto = (typeof TableEngine !== 'undefined' && r.telefono) ? TableEngine.waLink(r.telefono) : '';
      li.innerHTML = `<strong>${r.nombre || 'Sin nombre'}</strong>` +
        (grupoTxt ? ` · ${grupoTxt}` : '') +
        ` · sin asistir hace ${r.timeFmt ? r.timeFmt.main : (r.diasAusente + ' días')}` +
        (contacto ? ` · ${contacto}` : '');
      list.appendChild(li);
    });
    return list;
  }

  /* ══════════════════════════════════════════════════════════
     RENDER — construcción del contenido del modal
     ══════════════════════════════════════════════════════════ */
  function crearCard(label, value, sub) {
    const div = document.createElement('div');
    div.className = 're-card';
    div.innerHTML = `
      <div class="re-card-label">${label}</div>
      <div class="re-card-value">${value}</div>
      ${sub ? `<div class="re-card-sub">${sub}</div>` : ''}
    `;
    return div;
  }

  function crearCardEstado(label, valor, pctTexto, estado) {
    const div = document.createElement('div');
    div.className = `re-card re-estado-${estado}`;
    div.innerHTML = `
      <div class="re-card-label">${label}</div>
      <div class="re-card-value">${valor}</div>
      <div class="re-card-sub">${pctTexto}</div>
      <div class="re-estado-badge">${ICONO_ESTADO[estado]} ${LABEL_ESTADO[estado]}</div>
    `;
    return div;
  }

  function construirCuerpo(datos) {
    const body = document.createElement('div');
    body.className = 're-body';

    /* Aviso de "recalculando" — se muestra arriba de todo y NO
       bloquea el resto del reporte (algunas secciones pueden estar
       listas mientras otras aún recalculan). */
    if (datos.recalculando) {
      const aviso = document.createElement('div');
      aviso.className = 're-recalc-notice';
      aviso.innerHTML = `<i class="bi bi-arrow-repeat"></i> Algunos números se están recalculando — este reporte podría no reflejar aún el estado más reciente.`;
      body.appendChild(aviso);
    }

    /* ── 1) Resumen General ── */
    const secResumen = document.createElement('div');
    secResumen.className = 're-section';
    secResumen.innerHTML = `<div class="re-section-title"><i class="bi bi-clipboard-data"></i>Resumen General</div>`;
    const heroGrid = document.createElement('div');
    heroGrid.className = 're-hero-grid';
    heroGrid.appendChild(crearCard('Total Registrados', datos.total));
    heroGrid.appendChild(crearCard('% Asistencia General', datos.pctGeneral));
    heroGrid.firstChild.className = 're-hero-card';
    heroGrid.children[1].className = 're-hero-card';
    secResumen.appendChild(heroGrid);
    body.appendChild(secResumen);

    /* ── 2) Desglose de Asistencia ── */
    const estadoCelula   = clasificar(datos.celulaNOPct);
    const estadoServicio = clasificar(datos.servicioNOPct);
    const estadoAmbos    = clasificar(datos.ambosNOPct);

    const secDesglose = document.createElement('div');
    secDesglose.className = 're-section';
    secDesglose.innerHTML = `<div class="re-section-title"><i class="bi bi-bar-chart-line"></i>Desglose de Asistencia</div>`;
    const gridDesglose = document.createElement('div');
    gridDesglose.className = 're-grid';
    gridDesglose.appendChild(crearCard('Célula — Asistió', datos.celulaSI, datos.celulaSIPct));
    gridDesglose.appendChild(crearCardEstado('Célula — Faltó', datos.celulaNO, datos.celulaNOPctTexto, estadoCelula));
    gridDesglose.appendChild(crearCard('Servicio — Asistió', datos.servicioSI, datos.servicioSIPct));
    gridDesglose.appendChild(crearCardEstado('Servicio — Faltó', datos.servicioNO, datos.servicioNOPctTexto, estadoServicio));
    gridDesglose.appendChild(crearCard('Ambos — Asistió', datos.ambosSI, datos.ambosSIPct));
    gridDesglose.appendChild(crearCardEstado('Ambos — Faltó', datos.ambosNO, datos.ambosNOPctTexto, estadoAmbos));
    secDesglose.appendChild(gridDesglose);
    body.appendChild(secDesglose);

    /* ── 3) Nuevos Ingresos ── */
    const secNuevos = document.createElement('div');
    secNuevos.className = 're-section';
    secNuevos.innerHTML = `<div class="re-section-title"><i class="bi bi-person-plus"></i>Nuevos Ingresos</div>`;
    const gridNuevos = document.createElement('div');
    gridNuevos.className = 're-grid';
    gridNuevos.appendChild(crearCard('Nuevos en Célula', datos.nuevosCelula));
    gridNuevos.appendChild(crearCard('Nuevos en Servicio', datos.nuevosServicio));
    secNuevos.appendChild(gridNuevos);
    body.appendChild(secNuevos);

    /* ── 4) Monitor de Ausencias ── */
    const secAus = document.createElement('div');
    secAus.className = 're-section';
    secAus.innerHTML = `<div class="re-section-title"><i class="bi bi-exclamation-diamond"></i>Monitor de Ausencias</div>`;
    const gridAus = document.createElement('div');
    gridAus.className = 're-aus-grid';
    const cNormal = crearCard('Normal', datos.ausNormal); cNormal.className = 're-card re-aus-normal';
    const cWatch  = crearCard('Seguimiento', datos.ausWatch); cWatch.className = 're-card re-aus-watch';
    const cWarn   = crearCard('⚠️ Advertencia', datos.ausWarn); cWarn.className = 're-card re-aus-warn';
    const cCrit   = crearCard('🔴 Crítico', datos.ausCrit); cCrit.className = 're-card re-aus-crit';
    gridAus.append(cNormal, cWatch, cWarn, cCrit);
    secAus.appendChild(gridAus);
    body.appendChild(secAus);

    /* ── 5) Plan de Seguimiento ── */
    const secPlan = document.createElement('div');
    secPlan.className = 're-section';
    secPlan.innerHTML = `<div class="re-section-title"><i class="bi bi-clipboard-check"></i>Plan de Seguimiento</div>`;

    let huboContenido = false;

    /* -- 5.1) Personas: seguimiento inmediato (nivel Crítico) -- */
    const prioritarios = personasPrioritarias(8);
    if (prioritarios.length > 0) {
      huboContenido = true;
      secPlan.appendChild(crearBloqueTexto(
        '🔴 Seguimiento inmediato',
        `${prioritarios.length} persona(s) llevan más de 4 semanas sin asistir ni a célula ni a servicio. Son la prioridad #1 de contacto pastoral esta semana — cuanto más tiempo pasa sin contacto, más difícil es reconectarlas.`
      ));
      secPlan.appendChild(crearListaPersonas(prioritarios));
    }

    /* -- 5.2) Personas: en riesgo de volverse crítico -- */
    const advertencia = personasEnAdvertencia(6);
    if (advertencia.length > 0) {
      huboContenido = true;
      secPlan.appendChild(crearBloqueTexto(
        '⚠️ En riesgo — contactar antes de que pase a crítico',
        `${advertencia.length} persona(s) llevan entre 2 y 4 semanas sin asistir. Todavía están a tiempo de un contacto preventivo (llamada o visita) antes de entrar en la categoría anterior.`
      ));
      secPlan.appendChild(crearListaPersonas(advertencia));
    }

    /* -- 5.3) Causa probable #1: ¿concentrada en un grupo o repartida? -- */
    const focosCausa = [
      { metric: 'ambosNO',    nombre: 'la inasistencia a ambos (célula y servicio)' },
      { metric: 'celulasNO',  nombre: 'la inasistencia a célula' },
      { metric: 'servicioNO', nombre: 'la inasistencia a servicio' },
    ];
    for (const foco of focosCausa) {
      const c = analizarConcentracionPorGrupo(foco.metric);
      if (!c || c.total < 3) continue; // muestra insuficiente para sacar una conclusión útil
      huboContenido = true;
      const texto = c.pct >= 40
        ? `${foco.nombre[0].toUpperCase()}${foco.nombre.slice(1)} está concentrada principalmente en <strong>${c.grupoTop}</strong> (${c.cantidadTop} de ${c.total} casos, ${c.pct}%). Esto sugiere una causa puntual de ese grupo — vale la pena conversar con su líder sobre horarios, actividades paralelas o algún conflicto reciente.`
        : `${foco.nombre[0].toUpperCase()}${foco.nombre.slice(1)} está repartida entre ${c.gruposDistintos} grupos distintos, sin que ninguno concentre más del ${c.pct}% de los casos. Esto sugiere una causa más general (fecha especial, feriado, clima) que un problema de un grupo puntual.`;
      secPlan.appendChild(crearBloqueTexto('🔎 Causa probable', texto));
      break; // un solo diagnóstico de concentración es suficiente para no saturar el informe
    }

    /* -- 5.4) Causa probable #2: patrón célula / servicio / ambos -- */
    const patron = analizarPatronFaltas();
    const totalConFalta = patron.soloCelula + patron.soloServicio + patron.ambos;
    if (totalConFalta > 0) {
      huboContenido = true;
      const partes = [];
      if (patron.ambos > 0)        partes.push(`<strong>${patron.ambos}</strong> faltan a ambos (mayor riesgo de desconexión — requieren contacto pastoral directo, no solo un recordatorio)`);
      if (patron.soloServicio > 0) partes.push(`<strong>${patron.soloServicio}</strong> faltan solo al servicio (revisar si el horario del culto choca con algo — trabajo, transporte, otro compromiso)`);
      if (patron.soloCelula > 0)   partes.push(`<strong>${patron.soloCelula}</strong> faltan solo a célula (revisar si hay fricción puntual con ese grupo o su horario)`);
      secPlan.appendChild(crearBloqueTexto('🧭 Patrón de inasistencia', partes.join('; ') + '.'));
    }

    /* -- 5.5) Plan de acción — recomendaciones accionables fijas,
       basadas en lo anterior (no dependen de ningún motor externo
       ni de IA: son buenas prácticas de seguimiento pastoral). -- */
    if (huboContenido) {
      const acciones = document.createElement('ul');
      acciones.className = 're-reco-list';
      const items = [
        'Prioriza el contacto (llamada o visita) con la lista de "Seguimiento inmediato" esta misma semana.',
        'Para el grupo con mayor concentración de faltas (si lo hay), coordina con su líder una conversación breve sobre posibles causas.',
        'A quienes faltan solo a servicio, ofréceles alternativas (otro horario de culto, transmisión en vivo) si el problema es de horario.',
        'A quienes faltan a ambos, prioriza una visita personal antes que un mensaje — es la señal de mayor riesgo de abandono.',
        'Vuelve a generar este reporte tras el contacto para verificar si las personas listadas se movieron a un nivel de menor riesgo.',
      ];
      items.forEach(txt => {
        const li = document.createElement('li');
        li.className = 're-reco-item';
        li.textContent = txt;
        acciones.appendChild(li);
      });
      const tituloAcciones = document.createElement('div');
      tituloAcciones.className = 're-analisis-titulo';
      tituloAcciones.textContent = '✅ Plan de acción sugerido';
      secPlan.appendChild(tituloAcciones);
      secPlan.appendChild(acciones);
    } else {
      const ok = document.createElement('div');
      ok.className = 're-plan-empty';
      ok.innerHTML = `✅ <div>No se detectaron personas en riesgo ni patrones de inasistencia relevantes en los datos filtrados actuales. El panorama se ve saludable en general.</div>`;
      secPlan.appendChild(ok);
    }

    body.appendChild(secPlan);

    return body;
  }

  /* ══════════════════════════════════════════════════════════
     OVERLAY — se crea UNA sola vez y se reutiliza (idempotente).
     Cada apertura RE-RENDERIZA el body con datos frescos del DOM.
     ══════════════════════════════════════════════════════════ */
  let overlayEl = null;

  function cerrarModal() {
    if (overlayEl) overlayEl.classList.remove('re-visible');
    document.removeEventListener('keydown', onEscape);
  }

  function onEscape(ev) {
    if (ev.key === 'Escape') cerrarModal();
  }

  function crearOverlaySiHaceFalta() {
    if (overlayEl) return overlayEl;

    overlayEl = document.createElement('div');
    overlayEl.id = OVERLAY_ID;

    // Clic fuera del modal (en el fondo oscuro) también cierra
    overlayEl.addEventListener('click', (ev) => {
      if (ev.target === overlayEl) cerrarModal();
    });

    document.body.appendChild(overlayEl);
    return overlayEl;
  }

  function abrirReporte() {
    inyectarEstilos();
    const overlay = crearOverlaySiHaceFalta();

    // Re-construye el modal completo con datos frescos cada vez que
    // se abre — así el reporte nunca queda "viejo" de una apertura
    // anterior si el usuario cambió filtros o cargó otro archivo.
    overlay.innerHTML = '';

    const modal = document.createElement('div');
    modal.className = 're-modal';
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    modal.setAttribute('aria-labelledby', 're-titulo');

    const header = document.createElement('div');
    header.className = 're-header';
    header.innerHTML = `
      <div class="re-header-icon"><i class="bi bi-file-earmark-bar-graph"></i></div>
      <div class="re-header-text">
        <h2 id="re-titulo">Reporte Ejecutivo Detallado</h2>
        <span>${new Date().toLocaleString('es-PE', { dateStyle: 'medium', timeStyle: 'short' })}</span>
      </div>
      <button type="button" class="re-close" aria-label="Cerrar">&times;</button>
    `;
    header.querySelector('.re-close').addEventListener('click', cerrarModal);

    const datos = recolectarDatos();
    const body = construirCuerpo(datos);

    const footer = document.createElement('div');
    footer.className = 're-footer';
    const btnCerrar = document.createElement('button');
    btnCerrar.type = 'button';
    btnCerrar.className = 're-footer-btn';
    btnCerrar.textContent = 'Cerrar';
    btnCerrar.addEventListener('click', cerrarModal);
    footer.appendChild(btnCerrar);

    modal.append(header, body, footer);
    overlay.appendChild(modal);

    overlay.classList.add('re-visible');
    document.addEventListener('keydown', onEscape);
  }

  /* ══════════════════════════════════════════════════════════
     INIT — engancha el botón del menú lateral
     ══════════════════════════════════════════════════════════ */
  function init() {
    document.getElementById(BTN_ID)?.addEventListener('click', () => {
      // Cierra el offcanvas del menú de forma suave, igual que hacía
      // la conexión anterior con Sebastián (mismo patrón que el resto
      // de botones del menú lateral).
      const menuEl = document.getElementById('sidebarMenu');
      if (menuEl && window.bootstrap) {
        window.bootstrap.Offcanvas.getInstance(menuEl)?.hide();
      }
      abrirReporte();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
