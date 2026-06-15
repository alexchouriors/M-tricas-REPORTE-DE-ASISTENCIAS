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

      /* Detectar encabezado de grupo ministerial:
         - La fila contiene texto en col0
         - NO empieza con número ni con "N°" ni con "TOTAL"
         - Col1 está vacía (no es una fila de datos con nombre)
         - Col2 está vacía o vacía de datos de asistencia
      */
      const isGroupHeader = (
        col0.length > 3 &&
        col2 === '' &&
        col3 === '' &&
        !/^\d/.test(col0) &&
        !col0.startsWith('N°') &&
        !col0.startsWith('TOTAL') &&
        !col0.startsWith('REPOR') &&
        !col0.startsWith('Tema') &&
        !col0.startsWith('Fecha')
      );

      if (isGroupHeader) {
        currentGroup = col0;
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
          grupo:           currentGroup,
          esNuevo:         esNuevo,
          esNuevoCelula:   esNuevoCelula,       // NUEVO en célula (Estado=NUEVO)
          esNuevoServicio: esNuevoServicio,     // NUEVO en servicio (Célula=NUEVO)
          fecha:           fecha,               // fecha de última ausencia registrada
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

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const col0 = this.str(row[0]);
      const col1 = this.str(row[1]);
      const col2 = this.str(row[2]);
      const col3 = this.str(row[3]);
      const col4 = this.str(row[4]);

      // Detectar encabezado de grupo
      if (col0.length > 3 && col2 === '' && !/^\d/.test(col0) && !col0.startsWith('TOTAL')) {
        currentGroup = col0;
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
          grupo:    currentGroup,
          esNuevo:  false,
          fecha:    this.excelDate(fechaRaw || row[5]),
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
      if (col0.length > 3 && !/^\d/.test(col0) && !col0.startsWith('TOTAL')) {
        currentGroup = col0;
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
          grupo:    currentGroup,
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
  },
};


/* ────────────────────────────────────────────────────────────
   3. KPI ENGINE — calcula todos los indicadores
   Para añadir nuevos KPIs en el futuro: agregar métodos aquí
   y llamarlos desde compute().
──────────────────────────────────────────────────────────── */
const KPIEngine = {

  /* Calcula todos los KPIs sobre los registros filtrados */
  compute(records) {
    const total = records.length;

    // Asistencia célula
    const celulasSI   = records.filter(r => r.celula  === 'SI').length;
    const celulasNO   = records.filter(r => r.celula  === 'NO').length;
    const celulasNUEVO = records.filter(r => r.celula === 'NUEVO').length;

    // Asistencia servicio
    const servicioSI   = records.filter(r => r.servicio === 'SI').length;
    const servicioNO   = records.filter(r => r.servicio === 'NO').length;
    const servicioNUEVO = records.filter(r => r.servicio === 'NUEVO').length;

    // Asistencia ambos
    const ambosSI = records.filter(r =>
      (r.celula === 'SI' || r.celula === 'NUEVO') &&
      (r.servicio === 'SI' || r.servicio === 'NUEVO')
    ).length;

    // Inasistencia total (no fue a ninguno)
    const ambosNO = records.filter(r =>
      r.celula === 'NO' && r.servicio === 'NO'
    ).length;

    // ── Nuevos (usando los flags corregidos del parser) ──
    // esNuevoCelula   → Estado (col E) = 'NUEVO'  (nuevos integrados a célula)
    // esNuevoServicio → Célula (col C) = 'NUEVO'  (nuevos que llegaron al servicio)
    const nuevosCelula   = records.filter(r => r.esNuevoCelula).length;
    const nuevosServicio = records.filter(r => r.esNuevoServicio).length;
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
};


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
            color: '#94a3b8',
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
          bodyColor: '#94a3b8',
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
            ticks: { color: '#64748b', font: { family: 'Outfit', size: 11 } },
          },
          y: {
            grid: { display: false },
            ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 11 } },
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
            ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 11 }, maxRotation: 30 },
          },
          y: {
            grid: { color: 'rgba(255,255,255,.04)' },
            ticks: { color: '#64748b', font: { family: 'Outfit', size: 11 } },
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
            ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 11 }, maxRotation: 30 },
          },
          y: {
            stacked: true,
            grid: { color: 'rgba(255,255,255,.04)' },
            ticks: { color: '#64748b', font: { family: 'Outfit', size: 11 } },
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
              color: '#94a3b8',
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
              color: '#64748b',
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
    return `<span style="color:#64748b">${val || '—'}</span>`;
  },

  /* Render tabla de personas */
  renderPersonas(records) {
    const tbody = document.querySelector('#tablePersonas tbody');
    if (!tbody) return;

    tbody.innerHTML = records.map((r, i) => `
      <tr>
        <td>${i+1}</td>
        <td>${r.nombre}</td>
        <td style="color:#94a3b8;font-size:11px">${r.grupo.replace(/Ministr[ao]s?\s*/i,'').substring(0,30)}</td>
        <td>${this.badge(r.celula)}</td>
        <td>${this.badge(r.servicio)}</td>
        <td style="color:#64748b;font-size:11px">${r.estado || '—'}</td>
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
        <td style="color:#94a3b8;font-size:11px">${r.grupo.substring(0,30)}</td>
        <td>${this.badge(r.celula)}</td>
        <td>${this.badge(r.servicio)}</td>
        <td style="color:#64748b;font-size:11px">${r.estado || '—'}</td>
        <td style="color:#64748b;font-size:11px">${r.fecha || '—'}</td>
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
        <td style="color:#94a3b8;font-size:11px">${r.grupo.replace(/Ministr[ao]s?\s*/i,'').substring(0,28)}</td>
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
        <td style="color:#64748b;font-size:11px">${r.fecha || '—'}</td>
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
    this.renderExcluidos(DataStore.rawExcluidos);
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
    // Cargar archivo
    document.getElementById('fileInput')?.addEventListener('change', e => {
      const file = e.target.files[0];
      if (file) this.loadFile(file);
      e.target.value = ''; // Permite recargar el mismo archivo
    });

    // Toggle excluidos
    document.getElementById('toggleExcluidos')?.addEventListener('change', e => {
      DataStore.includeExcluidos = e.target.checked;
      this.refresh();
    });

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
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        ExcelParser.parse(workbook);
        FilterEngine.populate(DataStore.rawMain);
        this.refresh();
        this.showDashboard();
        document.getElementById('footerFile').textContent = file.name;
        document.getElementById('reportTitle').textContent = DataStore.reportTitle || file.name;
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

    this.updateKPICards(kpis);
    ChartEngine.renderAll(kpis);
    TableEngine.renderAll(filtered);
    AbsenceEngine.render(filtered);  // Monitor de ausencias
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
      tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;color:var(--text-muted);padding:32px">
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
        <td style="color:#94a3b8;font-size:11px">${grpShort}</td>
        <td style="font-size:12px;color:var(--text-dim)">${r.fechaFormatted}</td>
        <td>${timeHtml}</td>
        <td>$$BADGE_C$$</td>
        <td>$$BADGE_S$$</td>
        <td style="color:#64748b;font-size:11px">${r.estado || '—'}</td>
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

    // Abrir modal desde topbar
    document.getElementById('btnGsheets')?.addEventListener('click', () => {
      this.state.modalRef.show();
    });

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


/* ────────────────────────────────────────────────────────────
   9. BOOTSTRAP — arranque cuando el DOM esté listo
──────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  UIController.init();
  GSheetsEngine.initModal();
  AbsenceEngine.initEvents();

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
