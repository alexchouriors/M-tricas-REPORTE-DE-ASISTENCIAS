/* ════════════════════════════════════════════════════════════
   AccessManager.js
   ────────────────────────────────────────────────────────────
   Motor de Control de Acceso (RBAC / filtro de vistas).

   Intercepta arreglos de registros YA PARSEADOS por ExcelParser
   y devuelve solo las filas que el usuario en sesión tiene
   permitido ver, según las reglas declaradas en SecurityConfig.js.

   Independiente de index.html: no toca el DOM ni conoce nada de
   UIController/ChartEngine/TableEngine. Solo transforma datos.

   POLÍTICA DE SEGURIDAD: este módulo es estrictamente FAIL-CLOSED.
   Ante cualquier duda, dato ausente o formato inesperado, la
   respuesta por defecto es NO MOSTRAR NADA ([]), nunca 'ALL'.

   NOTA DE ARQUITECTURA: script clásico (NO type="module"). Se carga
   con <script src="AccessManager.js"> igual que el resto del
   proyecto (app.js, TelegramEngine.js, etc.), justo después de
   SecurityConfig.js, así ACCESS_RULES ya existe como global cuando
   este archivo se ejecuta — sin depender de resolución de módulos
   ES6, que puede fallar según el contexto en que se sirva el sitio.
   ════════════════════════════════════════════════════════════ */

/* ── Llave del registro que contiene el "Grupo Ministerial" ──
   Configurable en un solo lugar: si ExcelParser cambia el nombre
   de la propiedad (p. ej. a 'grupoMinisterial'), solo hay que
   actualizar esta constante — no hace falta tocar el resto del
   archivo. Valor actual acorde a ExcelParser.parseMainSheet(),
   que hoy genera cada registro como { grupo: currentGroup, ... }. */
const EXCEL_GROUP_KEY = 'grupo';

/* Normaliza texto para comparar sin sensibilidad a mayúsculas,
   espacios extremos ni espacios múltiples internos. */
function normalize(str) {
  return (str || '').toString().trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Extrae de forma segura el nombre de usuario (string) a partir del
 * parámetro `usuario`, que puede llegar en distintas formas según el
 * origen de la sesión:
 *   - string directo               → "JONATHAN"
 *   - objeto de sesión              → { nombre: "JONATHAN", rol: "..." }
 *   - variantes de nomenclatura     → { username, user, name }
 *
 * Si no puede extraer un string válido y no vacío, devuelve `null`
 * (nunca un fallback silencioso como "[object Object]"), para que
 * el llamador pueda aplicar fail-closed de inmediato.
 *
 * @param {string|Object|null|undefined} usuario
 * @returns {string|null}
 */
function extractUserName(usuario) {
  if (typeof usuario === 'string') {
    const trimmed = usuario.trim();
    return trimmed !== '' ? trimmed : null;
  }

  if (usuario && typeof usuario === 'object') {
    const candidate =
      usuario.nombre ??
      usuario.username ??
      usuario.user ??
      usuario.name ??
      null;

    if (typeof candidate === 'string' && candidate.trim() !== '') {
      return candidate.trim();
    }
    return null; // Objeto sin ninguna propiedad de nombre reconocible
  }

  return null; // null, undefined, número, booleano, etc.
}

const AccessManager = {

  /**
   * Devuelve la regla de acceso configurada para un usuario:
   *   - 'ALL'          → sin restricciones
   *   - Array<string>   → lista blanca de grupos ministeriales permitidos
   *   - null            → usuario inválido/no reconocido (fail-closed)
   *
   * Los usuarios que NO aparecen en ACCESS_RULES se tratan como sin
   * permisos (fail-closed), para evitar exponer datos por error de
   * configuración.
   *
   * @param {string|Object} usuario
   * @returns {'ALL'|string[]|null}
   */
  _getRule(usuario) {
    if (typeof ACCESS_RULES === 'undefined') {
      console.error('[AccessManager] SecurityConfig.js no está cargado — no se puede resolver ninguna regla (fail-closed).');
      return null;
    }

    const userName = extractUserName(usuario);
    if (!userName) return null; // No se pudo determinar el usuario → fail-closed

    const target = normalize(userName);
    const matchKey = Object.keys(ACCESS_RULES).find(k => normalize(k) === target);

    if (!matchKey) return null; // Usuario no registrado en ACCESS_RULES → fail-closed

    return ACCESS_RULES[matchKey];
  },

  /**
   * Filtra un arreglo de registros según las reglas de acceso del
   * usuario indicado, comparando contra `record[EXCEL_GROUP_KEY]`
   * (el "Grupo Ministerial" detectado por ExcelParser).
   *
   * Comportamiento SIEMPRE fail-closed:
   *   - `usuario` ausente, inválido, o de tipo inesperado (p. ej. un
   *     objeto sin propiedad de nombre reconocible)  → []
   *   - `usuario` no encontrado en ACCESS_RULES        → []
   *   - Regla de grupos específicos + registro sin
   *     `record[EXCEL_GROUP_KEY]` válido               → ese registro se excluye
   *
   * Solo se devuelven TODOS los datos cuando la regla resuelta es
   * explícitamente 'ALL'.
   *
   * @param {Array<Object>} datos   - Registros parseados (DataStore.rawMain, etc.)
   * @param {string|Object} usuario - Usuario en sesión (string o { nombre, ... })
   * @returns {Array<Object>} Subconjunto de `datos` permitido para `usuario`
   */
  applyFilter(datos, usuario) {
    if (!Array.isArray(datos)) return [];

    const rule = this._getRule(usuario);

    // Fail-closed: usuario inválido, no reconocido, o sin regla resoluble
    if (rule === null) return [];

    // Acceso total explícito: única forma de devolver todos los datos
    if (rule === 'ALL') return datos;

    // Regla de grupos específicos
    const allowedGroups = Array.isArray(rule) ? rule.map(normalize) : [];

    return datos.filter(record => {
      if (!record || typeof record[EXCEL_GROUP_KEY] !== 'string') return false;
      return allowedGroups.includes(normalize(record[EXCEL_GROUP_KEY]));
    });
  },
};

/* NOTA: una declaración `const` a nivel superior de un script clásico
   crea una variable global (accesible como `AccessManager` a secas),
   pero NO se agrega automáticamente como propiedad de `window` (eso
   solo ocurre con `var` o `function`). app.js consulta explícitamente
   `window.AccessManager`, así que se expone aquí también por esa vía
   para máxima compatibilidad y robustez. */
window.AccessManager = AccessManager;
