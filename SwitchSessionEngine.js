/* ════════════════════════════════════════════════════════════
   SwitchSessionEngine.js
   ────────────────────────────────────────────────────────────
   Módulo 100% INDEPENDIENTE para el botón "Cambiar Sesión" del
   menú lateral: abre el modal #modalCambiarSesion, valida el
   nuevo nombre contra USUARIOS.JS (misma lista que usa el login
   normal) y, si es válido, reemplaza la sesión activa y recarga
   la página — EXACTAMENTE el mismo patrón que ya usa
   SessionEngine.logout() en app.js, para garantizar que
   AccessManager, KPIEngine, ChartEngine y UsuarioRules se
   reinicialicen limpios con el nuevo usuario, sin arrastrar
   datos/filtros de la sesión anterior.

   Qué NO hace (a propósito):
   - No modifica SessionEngine, AccessManager, SecurityConfig.js
     ni USUARIOS.JS. Solo LEE `USUARIOS_REGISTRADOS` (global de
     USUARIOS.JS) y escribe en la misma llave de sessionStorage
     que ya usa SessionEngine, para que ambos permanezcan en
     sincronía sin que un archivo necesite conocer al otro.
   - No decide permisos ni roles: eso lo sigue resolviendo
     AccessManager/SecurityConfig.js de siempre, al recargar.

   Cómo usarlo:
   Solo agrega este script en index.html, DESPUÉS de USUARIOS.JS
   (necesita `USUARIOS_REGISTRADOS` ya cargado) y en cualquier
   orden respecto a los demás módulos independientes:
       <script src="SwitchSessionEngine.js"></script>
   ════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  /* Debe coincidir EXACTAMENTE con SessionEngine.STORAGE_KEY en
     app.js — es la llave de sessionStorage donde vive el usuario
     activo. Se duplica aquí (en vez de leerla desde SessionEngine)
     para que este módulo siga siendo cargable de forma
     independiente incluso si algún día se sirve sin app.js. */
  const SESSION_STORAGE_KEY = 'ccrm_dashboard_user';

  function el(id) { return document.getElementById(id); }

  function showError(msg) {
    const errEl = el('switchSessionError');
    if (!errEl) return;
    errEl.textContent = msg;
    errEl.classList.remove('d-none');
  }

  function clearError() {
    el('switchSessionError')?.classList.add('d-none');
  }

  function resetModal() {
    const input = el('switchSessionInput');
    if (input) input.value = '';
    clearError();
    const btn = el('btnSwitchSessionAccept');
    if (btn) btn.disabled = false;
  }

  /** Lee USUARIOS_REGISTRADOS (mismo global que usa SessionEngine
      en app.js) de forma segura. Devuelve `null` si no está disponible. */
  function getUsuariosAutorizados() {
    if (typeof USUARIOS_REGISTRADOS === 'undefined' || !Array.isArray(USUARIOS_REGISTRADOS)) {
      console.error('[SwitchSessionEngine] USUARIOS.JS no está cargado — no se puede validar (fail-closed).');
      return null;
    }
    return USUARIOS_REGISTRADOS;
  }

  function confirmSwitch() {
    const input = el('switchSessionInput');
    const btn = el('btnSwitchSessionAccept');
    const name = (input?.value || '').trim();

    if (name === '') {
      showError('Ingresa un nombre de usuario.');
      input?.focus();
      return;
    }

    const usuarios = getUsuariosAutorizados();
    if (usuarios === null) {
      showError('No se pudo verificar el usuario (error de carga). Intenta de nuevo.');
      return;
    }

    const nameUpper = name.toUpperCase();
    const isAuthorized = usuarios.some(u => String(u).trim().toUpperCase() === nameUpper);

    if (!isAuthorized) {
      showError('Usuario no encontrado. Verifica el nombre ingresado.');
      input?.focus();
      return;
    }

    clearError();
    if (btn) btn.disabled = true; // evita doble clic mientras se recarga

    /* Notificación de auditoría (fire-and-forget), igual que hace
       SessionEngine en login/logout — no bloquea el cambio de sesión
       si falla o tarda. */
    if (typeof TelegramEngine !== 'undefined') {
      TelegramEngine.notifySession('login', name)
        .catch(err => console.error('[SwitchSessionEngine] Error al notificar cambio de sesión:', err));
    }

    /* Reemplaza la sesión activa y recarga: mismo patrón que
       SessionEngine.logout(), así todo el dashboard (DataStore,
       AccessManager, gráficos, tablas, permisos de UsuarioRules)
       se reinicializa limpio con el nuevo usuario. */
    sessionStorage.setItem(SESSION_STORAGE_KEY, name);
    window.location.reload();
  }

  function init() {
    const modalEl = el('modalCambiarSesion');
    if (!modalEl) return;

    // Reinicia el formulario cada vez que el modal se abre
    modalEl.addEventListener('shown.bs.modal', () => {
      resetModal();
      el('switchSessionInput')?.focus();
    });

    el('btnSwitchSessionAccept')?.addEventListener('click', confirmSwitch);
    el('switchSessionInput')?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') confirmSwitch();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
