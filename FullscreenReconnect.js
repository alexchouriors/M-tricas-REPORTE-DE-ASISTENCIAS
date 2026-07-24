/* ════════════════════════════════════════════════════════════
   FullscreenReconnect.js
   ────────────────────────────────────────────────────────────
   Módulo 100% INDEPENDIENTE de UX para pantalla completa.

   Qué hace:
   - Escucha 'fullscreenchange' (con prefijos vendor para Safari/
     iOS y navegadores antiguos).
   - Si el usuario ESTABA en pantalla completa y el navegador lo
     saca automáticamente (cambio de pestaña, cambio de app, etc.),
     muestra un botón flotante discreto: "Reanudar pantalla completa".
   - Al pulsarlo, vuelve a pedir fullscreen y el botón se oculta.

   Qué NO hace (a propósito):
   - No importa, referencia, ni modifica AccessManager.js,
     SecurityConfig.js, USUARIOS.JS, TelegramEngine.js ni app.js.
   - No toca sesión, roles, ni ningún dato del dashboard.
   - No se auto-invoca para ENTRAR a pantalla completa por su
     cuenta; solo reacciona cuando el propio usuario ya había
     activado el modo pantalla completa y el navegador lo cerró.

   Cómo usarlo:
   Solo agrega este script en index.html, en cualquier orden
   respecto a los demás (no depende de ninguno):
       <script src="FullscreenReconnect.js"></script>

   Es completamente autónomo: crea su propio botón, sus propios
   estilos (inyectados vía <style> con un id único) y su propia
   lógica. No requiere ningún cambio en style.css ni en index.html
   más allá de incluir el <script>.
   ════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  const STYLE_ID = 'fsReconnectStyles';
  const BTN_ID   = 'fsReconnectBtn';

  /* Bandera interna: solo mostramos el botón si el usuario había
     entrado a pantalla completa por su propia acción (evita mostrar
     el botón en cargas normales donde nunca se usó fullscreen). */
  let fueActivadoPorUsuario = false;

  /* ── Helpers cross-browser ── */
  function elementoEnFullscreen() {
    return !!(
      document.fullscreenElement ||
      document.webkitFullscreenElement ||
      document.mozFullScreenElement ||
      document.msFullscreenElement
    );
  }

  function solicitarFullscreen(el) {
    const target = el || document.documentElement;
    const req =
      target.requestFullscreen ||
      target.webkitRequestFullscreen ||
      target.mozRequestFullScreen ||
      target.msRequestFullscreen;

    if (req) {
      return req.call(target).catch(() => {
        /* El navegador puede rechazar si no hay gesto de usuario
           reciente; en ese caso simplemente dejamos el botón visible
           para que el usuario lo intente de nuevo con un clic. */
      });
    }
    return Promise.resolve();
  }

  /* ── Inyección de estilos (aislados, sin tocar style.css) ── */
  function inyectarEstilos() {
    if (document.getElementById(STYLE_ID)) return;

    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${BTN_ID} {
        position: fixed;
        bottom: 22px;
        right: 22px;
        z-index: 2147483647;
        display: none;
        align-items: center;
        gap: 8px;
        padding: 10px 16px;
        border: none;
        border-radius: 999px;
        background: rgba(22, 27, 36, 0.92);
        color: #f1f5f9;
        font-family: 'Outfit', system-ui, sans-serif;
        font-size: 14px;
        font-weight: 500;
        letter-spacing: 0.2px;
        box-shadow: 0 6px 20px rgba(0,0,0,0.35), 0 0 0 1px rgba(255,255,255,0.06);
        backdrop-filter: blur(6px);
        cursor: pointer;
        transition: transform 0.18s ease, opacity 0.18s ease, background 0.18s ease;
        opacity: 0;
        transform: translateY(8px);
      }
      #${BTN_ID}.fs-visible {
        display: inline-flex;
        opacity: 1;
        transform: translateY(0);
      }
      #${BTN_ID}:hover {
        background: rgba(34, 197, 94, 0.95);
        color: #0d1117;
      }
      #${BTN_ID} svg {
        width: 16px;
        height: 16px;
        flex-shrink: 0;
      }
      @media (max-width: 480px) {
        #${BTN_ID} {
          bottom: 16px;
          right: 16px;
          left: 16px;
          justify-content: center;
        }
      }
    `;
    document.head.appendChild(style);
  }

  /* ── Creación del botón flotante ── */
  function crearBoton() {
    let btn = document.getElementById(BTN_ID);
    if (btn) return btn;

    btn = document.createElement('button');
    btn.id = BTN_ID;
    btn.type = 'button';
    btn.setAttribute('aria-label', 'Reanudar pantalla completa');
    btn.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
           stroke-linecap="round" stroke-linejoin="round">
        <path d="M8 3H5a2 2 0 0 0-2 2v3"></path>
        <path d="M21 8V5a2 2 0 0 0-2-2h-3"></path>
        <path d="M3 16v3a2 2 0 0 0 2 2h3"></path>
        <path d="M16 21h3a2 2 0 0 0 2-2v-3"></path>
      </svg>
      <span>Reanudar pantalla completa</span>
    `;

    btn.addEventListener('click', () => {
      solicitarFullscreen(document.documentElement).finally(() => {
        ocultarBoton();
      });
    });

    document.body.appendChild(btn);
    return btn;
  }

  function mostrarBoton() {
    const btn = document.getElementById(BTN_ID) || crearBoton();
    btn.classList.add('fs-visible');
  }

  function ocultarBoton() {
    const btn = document.getElementById(BTN_ID);
    if (btn) btn.classList.remove('fs-visible');
  }

  /* ── Listener principal ── */
  function manejarCambioFullscreen() {
    if (elementoEnFullscreen()) {
      /* Entró (o volvió a entrar) a pantalla completa */
      fueActivadoPorUsuario = true;
      ocultarBoton();
    } else if (fueActivadoPorUsuario) {
      /* Salió de pantalla completa habiendo estado antes en ese modo
         → probablemente por cambio de pestaña/app, no por decisión
         explícita de salir. Mostramos el botón de reconexión rápida. */
      mostrarBoton();
    }
  }

  function init() {
    inyectarEstilos();

    document.addEventListener('fullscreenchange', manejarCambioFullscreen);
    document.addEventListener('webkitfullscreenchange', manejarCambioFullscreen);
    document.addEventListener('mozfullscreenchange', manejarCambioFullscreen);
    document.addEventListener('MSFullscreenChange', manejarCambioFullscreen);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
