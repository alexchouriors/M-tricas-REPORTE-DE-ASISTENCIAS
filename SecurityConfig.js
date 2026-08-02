/* ════════════════════════════════════════════════════════════
   SecurityConfig.js
   ────────────────────────────────────────────────────────────
   Diccionario de permisos de acceso (RBAC) por usuario.

   Cada clave es el nombre de usuario TAL COMO aparece en
   USUARIOS.JS / se ingresa en el login (se compara sin importar
   mayúsculas/minúsculas ni espacios extremos — ver AccessManager).

   El valor puede ser:
     • "ALL"            → acceso total, ve todas las filas del reporte.
     • Array de strings  → lista blanca de valores permitidos para el
                           campo "Grupo Ministerial" de cada registro
                           (comparación insensible a mayúsculas/espacios).
                           Estos valores corresponden al texto de la
                           celda combinada/centrada (A:E) que ExcelParser
                           detecta como encabezado de grupo en la hoja
                           principal del reporte (p. ej. la fila con la
                           celda combinada "MINISTRO Jonathan y Mayerling").

   IMPORTANTE — usuarios NO listados aquí:
   Cualquier usuario que NO aparezca explícitamente en este diccionario
   recibe FAIL-CLOSED (no ve ningún dato — [] en vez de acceso total).
   Si un usuario nuevo debe tener acceso, hay que agregarlo aquí a
   propósito, ya sea con 'ALL' o con su lista de grupos permitidos.

   NOTA DE ARQUITECTURA: este archivo se carga como <script> clásico
   (NO como type="module") para que quede disponible como variable
   global ANTES de que se ejecute app.js, igual que USUARIOS.JS. Los
   módulos ES6 (import/export) dependen de que el navegador pueda
   resolverlos como recursos de red/CORS, lo cual puede fallar en
   ciertos contextos (p. ej. al abrir index.html directamente sin
   servidor) y dejaba el sistema entero en fail-closed por error.
   ════════════════════════════════════════════════════════════ */

const ACCESS_RULES = {

  /* ── Grupo "Jonathan y Mayerling" ──
     Acepta ambas variantes de texto que aparecen en la celda
     combinada y centrada (A:E) del encabezado de grupo en el Excel:
       - "Jonathan y Mayerling"
       - "MINISTRO Jonathan y Mayerling"                            */
  JONATHAN: [
    'Jonathan y Mayerling',
    'MINISTRO Jonathan y Mayerling',
  ],
  MAYERLIN: [
    'Jonathan y Mayerling',
    'MINISTRO Jonathan y Mayerling',
  ],

  /* ── Grupo "Renny y Airam" ──
     Celda combinada/centrada A21:E21 del Excel:
       - "Renny y Airam"
       - "MINISTROS Renny y Airam"                                  */
  RENNY: [
    'Renny y Airam',
    'MINISTROS Renny y Airam',
  ],
  AIRAM: [
    'Renny y Airam',
    'MINISTROS Renny y Airam',
  ],

  /* ── Grupo "Magalis, Edith y Alexandro" ──
     Celda combinada/centrada A118:E118 del Excel:
       - "Magalis, Edith y Alexandro"
       - "Ministra Magalis, Edith y Alexandro"                     */
  MAGALIS: [
    'Magalis, Edith y Alexandro',
    'Ministra Magalis, Edith y Alexandro',
  ],
  EDITH: [
    'Magalis, Edith y Alexandro',
    'Ministra Magalis, Edith y Alexandro',
  ],
  ALEXANDRO: [
    'Magalis, Edith y Alexandro',
    'Ministra Magalis, Edith y Alexandro',
  ],
  GARDYS: [
    'Magalis, Edith y Alexandro',
    'Ministra Magalis, Edith y Alexandro',
  ],

  /* ── Grupo "LIDERES YENSI, SORENNYS Y YELSSY" ──
     Celda combinada/centrada A578:E578 del Excel:
       - "LIDERES YENSI, SORENNYS Y YELSSY"
       - "LIDERES YENSI, SORENNYS Y YELSSY"
     NOTA: los usuarios reales en USUARIOS.JS son "SORENNIS" y "YELSSY"
     (no "SORENNYS"); se usan esos nombres exactos para que el login
     y la regla de acceso coincidan.                                 */
  YENSI: [
    'LIDERES YENSI, SORENNYS Y YELSSY',
    'LIDERES YENSI, SORENNYS Y YELSSY',
  ],
  SORENNYS: [
    'LIDERES YENSI, SORENNYS Y YELSSY',
    'LIDERES YENSI, SORENNYS Y YELSSY',
  ],
  YELSSY: [
    'LIDERES YENSI, SORENNYS Y YELSSY',
    'LIDERES YENSI, SORENNYS Y YELSSY',
  ],

  /* ── Grupo "LIDER Josefa y Dayana, FRANGLIS" ──
     Celda combinada/centrada A327:E327 del Excel (mismo texto tanto
     en el nombre "oficial" del grupo como en el Excel):
       - "LIDER Josefa y Dayana, FRANGLIS"                          */
  JOSEFA: [
    'LIDER Josefa y Dayana, FRANGLIS',
  ],
  DAYANA: [
    'LIDER Josefa y Dayana, FRANGLIS',
  ],
  FRANGLYS: [
    'LIDER Josefa y Dayana, FRANGLIS',
  ],

  /* ── Acceso total (liderazgo pastoral / administración) ── */
  PASTOR:  'ALL',
  PASTORA: 'ALL',
  MASTER:  'ALL',

  /* ── Grupo "Lideres Jackson, Athais y Anais" ── */
  JACKSON: [
    'Lideres Jackson, Athais y Anais',
  ],
  ANAIS: [
    'Lideres Jackson, Athais y Anais',
  ],
  ATHAIS: [
    'Lideres Jackson, Athais y Anais',
  ],

  /* ── Grupo "Pastores Carlos Duarte y Génesis" ── */
  'PASTOR CARLOS D': [
    'Pastores Carlos Duarte y Génesis',
  ],
  'PASTORA GÉNESIS': [
    'Pastores Carlos Duarte y Génesis',
  ],
};
