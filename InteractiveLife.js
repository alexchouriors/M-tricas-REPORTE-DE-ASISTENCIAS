/* ════════════════════════════════════════════════════════════
   InteractiveLife.js
   ────────────────────────────────────────────────────────────
   Módulo 100% INDEPENDIENTE de micro-interacciones y "vida"
   visual para todo el dashboard — incluida la pantalla de login.

   Qué hace:
   - Login: orbes flotantes de fondo, logo con respiración/tilt,
     "shake" en el formulario cuando aparece un error, ripple en
     todos los botones.
   - Dashboard: ripple en botones, conteo animado de números KPI
     cuando cambian de valor, tilt 3D suave en tarjetas KPI y de
     gráficos al mover el mouse.

   Qué NO hace (a propósito):
   - No importa, referencia, ni modifica AccessManager.js,
     SecurityConfig.js, USUARIOS.JS, TelegramEngine.js, app.js ni
     ninguna lógica de negocio, datos, sesión o permisos. Solo
     observa el DOM (MutationObserver) y agrega clases/estilos
     puramente visuales.

   Cómo usarlo:
   Solo agrega este script en index.html, en cualquier orden
   respecto a los demás (no depende de ninguno):
       <script src="InteractiveLife.js"></script>
   ════════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  const prefersReducedMotion =
    window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ────────────────────────────────────────────────────────────
     1) RIPPLE — efecto de onda al hacer clic en cualquier botón
  ──────────────────────────────────────────────────────────── */
  function initRipple() {
    const SELECTOR = 'button, .btn, [role="button"]';

    document.addEventListener('click', (ev) => {
      const target = ev.target.closest(SELECTOR);
      if (!target || target.disabled) return;
      if (prefersReducedMotion) return;

      const rect = target.getBoundingClientRect();
      const size = Math.max(rect.width, rect.height) * 1.4;
      const ripple = document.createElement('span');
      ripple.className = 'il-ripple';
      ripple.style.width = ripple.style.height = size + 'px';
      ripple.style.left = (ev.clientX - rect.left - size / 2) + 'px';
      ripple.style.top = (ev.clientY - rect.top - size / 2) + 'px';

      const computed = getComputedStyle(target);
      if (computed.position === 'static') target.classList.add('il-ripple-host');

      target.appendChild(ripple);
      ripple.addEventListener('animationend', () => ripple.remove());
      setTimeout(() => ripple.remove(), 700); // fallback de seguridad
    }, true);
  }

  /* ────────────────────────────────────────────────────────────
     2) TILT 3D — inclinación sutil de tarjetas al mover el mouse
  ──────────────────────────────────────────────────────────── */
  function initTilt() {
    if (prefersReducedMotion) return;
    const SELECTOR = '.kpi-card, .chart-card, .login-brand-card';

    document.addEventListener('mousemove', (ev) => {
      const card = ev.target.closest(SELECTOR);
      if (!card) return;
      const rect = card.getBoundingClientRect();
      const px = (ev.clientX - rect.left) / rect.width;   // 0..1
      const py = (ev.clientY - rect.top) / rect.height;    // 0..1
      const rotY = (px - 0.5) * 8;   // grados
      const rotX = (0.5 - py) * 8;
      card.style.transform = `perspective(700px) rotateX(${rotX}deg) rotateY(${rotY}deg)`;
      card.classList.add('il-tilting');
    }, true);

    document.addEventListener('mouseleave', (ev) => {
      const card = ev.target.closest && ev.target.closest(SELECTOR);
      if (!card) return;
      card.style.transform = '';
      card.classList.remove('il-tilting');
    }, true);
  }

  /* ────────────────────────────────────────────────────────────
     3) CONTEO ANIMADO — anima los números de .kpi-value cuando
        su texto cambia (de "—" a un valor, o de un valor a otro),
        sin tocar la lógica que los calcula (KPIEngine en app.js).
  ──────────────────────────────────────────────────────────── */
  function animateNumber(el, from, to, suffix, duration) {
    if (prefersReducedMotion || from === to || isNaN(from) || isNaN(to)) {
      el.textContent = to + suffix;
      return;
    }
    const start = performance.now();
    function tick(now) {
      const p = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - p, 3); // easeOutCubic
      const value = Math.round(from + (to - from) * eased);
      el.textContent = value.toLocaleString('es') + suffix;
      if (p < 1) requestAnimationFrame(tick);
      else el.textContent = to.toLocaleString('es') + suffix;
    }
    requestAnimationFrame(tick);
  }

  function initKpiCountUp() {
    const values = document.querySelectorAll('.kpi-value, .kpi-pct');
    values.forEach((el) => {
      let lastNumeric = parseNumeric(el.textContent);
      let selfUpdating = false;

      const observer = new MutationObserver(() => {
        if (selfUpdating) return;
        const newText = el.textContent;
        const newNumeric = parseNumeric(newText);
        const suffixMatch = newText.match(/[%]$/);
        const suffix = suffixMatch ? '%' : '';

        if (newNumeric === null) { lastNumeric = null; return; }
        if (lastNumeric === null) { lastNumeric = newNumeric; return; } // primer valor real, sin animar desde "—"

        selfUpdating = true;
        observer.disconnect();
        animateNumber(el, lastNumeric, newNumeric, suffix, 650);
        lastNumeric = newNumeric;
        setTimeout(() => {
          selfUpdating = false;
          observer.observe(el, { childList: true, characterData: true, subtree: true });
        }, 700);
      });

      observer.observe(el, { childList: true, characterData: true, subtree: true });
    });
  }

  function parseNumeric(text) {
    const cleaned = (text || '').replace(/[.,\s]/g, '').match(/-?\d+/);
    return cleaned ? parseInt(cleaned[0], 10) : null;
  }

  /* ────────────────────────────────────────────────────────────
     4) LOGIN — orbes flotantes de fondo, logo con vida propia
        y "shake" cuando aparece un error de validación.
  ──────────────────────────────────────────────────────────── */
  function initLoginLife() {
    const overlay = document.getElementById('loginOverlay');
    if (!overlay) return;

    // ── Orbes flotantes de fondo (puramente decorativos) ──
    if (!prefersReducedMotion && !overlay.querySelector('.il-orb')) {
      const orbCount = 4;
      for (let i = 0; i < orbCount; i++) {
        const orb = document.createElement('div');
        orb.className = `il-orb il-orb-${i + 1}`;
        overlay.prepend(orb);
      }
    }

    // ── El logo respira suavemente y reacciona al hover ──
    const logo = document.getElementById('loginLogo');
    if (logo && !prefersReducedMotion) {
      logo.classList.add('il-logo-live');
    }

    // ── Shake en el paso de nombre cuando aparece un error ──
    const nameRow = document.querySelector('.login-name-input-row');
    ['loginNameError', 'loginNameNotFoundError'].forEach((id) => {
      const errEl = document.getElementById(id);
      if (!errEl || !nameRow) return;
      const obs = new MutationObserver(() => {
        if (!errEl.classList.contains('d-none') && !prefersReducedMotion) {
          nameRow.classList.remove('il-shake');
          void nameRow.offsetWidth; // fuerza reflow para reiniciar la animación
          nameRow.classList.add('il-shake');
        }
      });
      obs.observe(errEl, { attributes: true, attributeFilter: ['class'] });
    });

    // ── Shake en el formulario de soporte ──
    const supportForm = document.getElementById('loginSupportForm');
    const supportErr = document.getElementById('loginSupportError');
    if (supportForm && supportErr) {
      const obs2 = new MutationObserver(() => {
        if (!supportErr.classList.contains('d-none') && !prefersReducedMotion) {
          supportForm.classList.remove('il-shake');
          void supportForm.offsetWidth;
          supportForm.classList.add('il-shake');
        }
      });
      obs2.observe(supportErr, { attributes: true, attributeFilter: ['class'] });
    }
  }

  /* ── Init ── */
  function init() {
    initRipple();
    initTilt();
    initKpiCountUp();
    initLoginLife();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
