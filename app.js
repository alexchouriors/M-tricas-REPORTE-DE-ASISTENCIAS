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
            <h3 class="kpi-modal-title" id="kpiModalTitle"></h3>
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
      if (e.key === 'Escape' && backdrop.classList.contains('kpi-modal-visible')) this.close();
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

    const query = this._normalizeSearch(this._searchQuery);
    const personas = query === ''
      ? this._currentPersonas
      : this._currentPersonas.filter(p =>
          this._normalizeSearch(p.nombre).includes(query) ||
          this._normalizeSearch(p.grupo).includes(query)
        );

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

  /* Cierra el modal con una pequeña transición de salida */
  close() {
    const backdrop = document.getElementById('kpiDetailModal');
    if (!backdrop) return;

    backdrop.classList.remove('kpi-modal-visible');
    document.body.classList.remove('kpi-modal-open');
    clearTimeout(this._closeTimeout);
    this._closeTimeout = setTimeout(() => backdrop.classList.add('d-none'), 250);
  },
};
window.ModalEngine = ModalEngine;




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

  /* ── Rankings de grupos — barras verticales Chart.js ── */
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
        .substring(0, 22) || name.substring(0, 22);

      return { name, short, asistentes, ausentes, total, pctAsist, pctAus };
    });

    // Top asistencia: mayor % de asistencia primero
    const top    = [...allGroups].sort((a,b) => b.pctAsist - a.pctAsist).slice(0, 5);
    // Mayor ausentismo: mayor % de ausencia primero
    const bottom = [...allGroups].sort((a,b) => b.pctAus   - a.pctAus  ).slice(0, 5);

    this._renderRankChart('chartRankTop',    top,    'pctAsist', '#22c55e', 'rgba(34,197,94,.15)');
    this._renderRankChart('chartRankBottom', bottom, 'pctAus',   '#ef4444', 'rgba(239,68,68,.15)');
  },

  /* Renderiza un ranking como gráfico de barras verticales */
  _renderRankChart(canvasId, items, pctField, color, bgColor) {
    this.destroy(canvasId);
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    const isAttendance = (pctField === 'pctAsist');
    const label = isAttendance ? 'Asistencia %' : 'Ausentismo %';

    this.instances[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: items.map(i => i.short),
        datasets: [{
          label,
          data:  items.map(i => i[pctField]),
          backgroundColor: items.map(() => bgColor),
          borderColor:     items.map(() => color),
          borderWidth: 2,
          borderRadius: 6,
          borderSkipped: false,
          hoverBackgroundColor: color + '88',
        }],
      },
      options: {
        ...this.baseOptions(),
        plugins: {
          ...this.baseOptions().plugins,
          legend: { display: false },
          tooltip: {
            ...this.baseOptions().plugins.tooltip,
            callbacks: {
              label: (ctx) => {
                const item = items[ctx.dataIndex];
                return isAttendance
                  ? ` ${item.pctAsist}% asistencia (${item.asistentes}/${item.total})`
                  : ` ${item.pctAus}% ausencia (${item.ausentes}/${item.total})`;
              },
            },
          },
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: {
              color: Chart.defaults.color,
              font: { family: 'Outfit', size: 10 },
              maxRotation: 30,
              minRotation: 15,
            },
          },
          y: {
            beginAtZero: true,
            max: 100,
            grid: { color: 'rgba(255,255,255,.04)' },
            ticks: {
              color: Chart.defaults.color,
              font: { family: 'Outfit', size: 10 },
              callback: v => v + '%',
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
    this.renderRankings(kpis);  // uses chartRankTop / chartRankBottom canvases
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

    // Filtros: actualizar en cambio
    ['filterGroup','filterEstado','filterCelula','filterServicio','filterNuevo']
      .forEach(id => {
        document.getElementById(id)?.addEventListener('change', () => {
          FilterEngine.read();
          this.refresh();
        });
      });

    // Reset filtros
    document.getElementById('btnResetFilters')?.addEventListener('click', () => {
      FilterEngine.reset();
      this.refresh();
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

  /* Recalcula KPIs, actualiza gráficos y tablas con filtros aplicados */
  refresh() {
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

    // Re-renderiza gráficos si hay datos cargados
    if (!document.getElementById('dashboardContent').classList.contains('d-none')) {
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

    /* Auto-carga del archivo predeterminado (config.json → REPORTES/<archivo>),
       fire-and-forget: no bloquea el fade-out del overlay ni el login en sí.
       Inmediatamente DESPUÉS de que termine de cargar (encadenado con .then,
       no en paralelo), se revisa si hay una tendencia predeterminada guardada
       en localStorage (ver DbDefaultEngine) y se aplica automáticamente. */
    if (typeof AutoLoadEngine !== 'undefined') {
      AutoLoadEngine.loadDefaultFile().then(() => {
        if (typeof DbDefaultEngine !== 'undefined') {
          DbDefaultEngine.applyStoredTrendIfAny();
        }
      });
    }

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

  /* ── Actualiza referencias visuales del usuario activo (si existieran) ── */
  _updateSessionUI(name) {
    const label = this._el('sessionUserLabel');
    if (label) label.textContent = name;
  },

  init() {
    /* Si ya existe una sesión (misma pestaña), no se muestra el login */
    if (this.isLoggedIn()) {
      this.hideOverlayInstant();
      this._updateSessionUI(this.getUser());
      /* Restaura los permisos de INTERFAZ (Usuario Rules.js) para el
         usuario ya autenticado — no afecta datos ni filtros RBAC */
      if (window.UsuarioRules) window.UsuarioRules.applyUIPermissions(this.getUser());
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
      const cfgRes = await fetch(this.GITHUB_CONTENTS_BASE + 'config.json', {
        cache: 'no-store',
        headers: { 'Accept': 'application/vnd.github.v3.raw' },
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
        { cache: 'no-store', headers: { 'Accept': 'application/vnd.github.v3+json' } }
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
      const res = await fetch(this.GITHUB_CONTENTS_BASE + 'config.json', {
        cache: 'no-store',
        headers: { 'Accept': 'application/vnd.github.v3.raw' },
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
       Hace un fetch silencioso a config.json en la raíz del repo y,
       si existe, descarga y carga ese archivo con el mismo flujo que
       usa HistoryEngine._loadFile(), sin requerir token (lectura
       pública de la API de contenidos de GitHub).
──────────────────────────────────────────────────────────── */
const AutoLoadEngine = {

  GITHUB_CONTENTS_BASE: 'https://api.github.com/repos/alexchouriors/M-tricas-REPORTE-DE-ASISTENCIAS-NUEVA/contents/',

  async loadDefaultFile() {
    try {
      /* Lee config.json de forma silenciosa (si no existe, no hace nada) */
      const cfgRes = await fetch(this.GITHUB_CONTENTS_BASE + 'config.json', {
        cache: 'no-store',
        headers: { 'Accept': 'application/vnd.github.v3.raw' },
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
      /* 1) Obtiene la metadata del archivo (necesitamos su download_url) */
      const fileRes = await fetch(
        this.GITHUB_CONTENTS_BASE + 'REPORTES/' + encodeURIComponent(fileName),
        { cache: 'no-store', headers: { 'Accept': 'application/vnd.github.v3+json' } }
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
