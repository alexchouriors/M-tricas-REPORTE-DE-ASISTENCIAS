/* ════════════════════════════════════════════════════════════
   InteractiveLife.js
   ────────────────────────────────────────────────────────────
   Módulo 100% INDEPENDIENTE de micro-interacciones y "vida"
   visual para todo el dashboard — incluida la pantalla de login.

   Qué hace:
   - Splash de marca al abrir la app (antes del botón "Iniciar
     Sesión"), saltable con un clic.
   - Login: campo de partículas tipo constelación en canvas
     (reactivo al mouse), logo con respiración/hover, "shake" en
     el formulario cuando aparece un error, ripple en botones.
   - Dashboard: ripple en botones, conteo animado de números KPI,
     tilt 3D en tarjetas, gráficos Chart.js con animación de
     entrada más expresiva, tablas en cascada, anillo de progreso
     en el loading overlay, revelado teatral tras el login, y un
     barrido circular al cambiar de tema claro/oscuro.

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
  function animateNumber(el, from, to, suffix, duration, decimals) {
    if (prefersReducedMotion || from === to || isNaN(from) || isNaN(to)) {
      el.textContent = formatNumber(to, decimals) + suffix;
      return;
    }
    const start = performance.now();
    function tick(now) {
      const p = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - p, 3); // easeOutCubic
      const value = from + (to - from) * eased;
      el.textContent = formatNumber(value, decimals) + suffix;
      if (p < 1) requestAnimationFrame(tick);
      else el.textContent = formatNumber(to, decimals) + suffix;
    }
    requestAnimationFrame(tick);
  }

  function formatNumber(value, decimals) {
    return decimals > 0
      ? value.toFixed(decimals)
      : Math.round(value).toLocaleString('es');
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
        const suffixMatch = newText.match(/%\s*$/);
        const suffix = suffixMatch ? '%' : '';
        // Conserva la cantidad de decimales tal como los emite app.js
        // (ej. "36.5%" → 1 decimal), para no perder precisión al animar.
        const decMatch = newText.match(/\.(\d+)/);
        const decimals = decMatch ? decMatch[1].length : 0;

        if (newNumeric === null) { lastNumeric = null; return; }
        if (lastNumeric === null) { lastNumeric = newNumeric; return; } // primer valor real, sin animar desde "—"

        // Duración reducida (antes 650ms) para eliminar la sensación de
        // retardo al cambiar de filtro. requestAnimationFrame (dentro de
        // animateNumber) ya evita bloquear el hilo principal, así que
        // gráficos y tablas se renderizan de inmediato sin esperar a que
        // termine este conteo.
        const COUNT_DURATION = 350; // ms — tope solicitado: máx. 400ms

        selfUpdating = true;
        observer.disconnect();
        animateNumber(el, lastNumeric, newNumeric, suffix, COUNT_DURATION, decimals);
        lastNumeric = newNumeric;
        setTimeout(() => {
          selfUpdating = false;
          observer.observe(el, { childList: true, characterData: true, subtree: true });
        }, COUNT_DURATION + 50); // pequeño margen sobre la duración real
      });

      observer.observe(el, { childList: true, characterData: true, subtree: true });
    });
  }

  function parseNumeric(text) {
    // Solo tolera "." como separador decimal (no como separador de miles),
    // ya que los KPI de app.js nunca emiten miles con punto en este proyecto.
    const cleaned = (text || '').replace(/,/g, '').match(/-?\d+(\.\d+)?/);
    return cleaned ? parseFloat(cleaned[0]) : null;
  }

  /* ────────────────────────────────────────────────────────────
     4) LOGIN — canvas de partículas tipo constelación (reactivo
        al mouse), logo con vida propia y "shake" en errores.
  ──────────────────────────────────────────────────────────── */
  function initParticleField(overlay) {
    if (prefersReducedMotion) return;

    const canvas = document.createElement('canvas');
    canvas.className = 'il-particle-canvas';
    overlay.prepend(canvas);
    const ctx = canvas.getContext('2d');

    let w, h, particles, mouse = { x: -9999, y: -9999 };
    const COUNT = 60;
    const LINK_DIST = 130;
    const colors = ['#4f7cff', '#f0b429', '#22c55e', '#ef4444'];

    function resize() {
      w = canvas.width = overlay.clientWidth;
      h = canvas.height = overlay.clientHeight;
    }

    function makeParticles() {
      particles = Array.from({ length: COUNT }, () => ({
        x: Math.random() * w,
        y: Math.random() * h,
        vx: (Math.random() - 0.5) * 0.35,
        vy: (Math.random() - 0.5) * 0.35,
        r: Math.random() * 1.8 + 0.6,
        c: colors[Math.floor(Math.random() * colors.length)],
      }));
    }

    function step() {
      ctx.clearRect(0, 0, w, h);

      // Actualiza y dibuja partículas
      particles.forEach((p) => {
        p.x += p.vx; p.y += p.vy;
        if (p.x < 0 || p.x > w) p.vx *= -1;
        if (p.y < 0 || p.y > h) p.vy *= -1;

        // Leve atracción hacia el mouse (efecto vivo, no invasivo)
        const dx = mouse.x - p.x, dy = mouse.y - p.y;
        const dist = Math.hypot(dx, dy);
        if (dist < 160) {
          p.x -= dx * 0.0018;
          p.y -= dy * 0.0018;
        }

        ctx.beginPath();
        ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
        ctx.fillStyle = p.c;
        ctx.globalAlpha = 0.75;
        ctx.fill();
      });
      ctx.globalAlpha = 1;

      // Líneas entre partículas cercanas (efecto constelación)
      for (let i = 0; i < particles.length; i++) {
        for (let j = i + 1; j < particles.length; j++) {
          const a = particles[i], b = particles[j];
          const d = Math.hypot(a.x - b.x, a.y - b.y);
          if (d < LINK_DIST) {
            ctx.beginPath();
            ctx.moveTo(a.x, a.y);
            ctx.lineTo(b.x, b.y);
            ctx.strokeStyle = `rgba(120,150,255,${(1 - d / LINK_DIST) * 0.22})`;
            ctx.lineWidth = 1;
            ctx.stroke();
          }
        }
      }

      requestAnimationFrame(step);
    }

    resize();
    makeParticles();
    requestAnimationFrame(step);

    window.addEventListener('resize', () => { resize(); });
    overlay.addEventListener('mousemove', (ev) => {
      const rect = overlay.getBoundingClientRect();
      mouse.x = ev.clientX - rect.left;
      mouse.y = ev.clientY - rect.top;
    });
    overlay.addEventListener('mouseleave', () => { mouse.x = -9999; mouse.y = -9999; });
  }

  function initLoginLife() {
    const overlay = document.getElementById('loginOverlay');
    if (!overlay) return;

    // ── Campo de partículas tipo constelación de fondo ──
    if (!overlay.querySelector('.il-particle-canvas')) {
      initParticleField(overlay);
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

  /* ────────────────────────────────────────────────────────────
     10) SPLASH DE MARCA — pantalla de bienvenida breve al abrir
         la app, ANTES incluso del botón "Iniciar Sesión". Se
         inyecta por encima de #loginOverlay (que sigue montado
         normalmente debajo) y se retira sola. Se puede saltar
         con un clic. No depende de SessionEngine ni lo bloquea:
         el login overlay ya está listo debajo desde el inicio.
  ──────────────────────────────────────────────────────────── */
  function initSplashScreen() {
    const loginOverlay = document.getElementById('loginOverlay');
    if (!loginOverlay) return;

    const logoSrc = document.getElementById('loginLogo')?.getAttribute('src') || 'LOGO_IGLEISA.png';

    const splash = document.createElement('div');
    splash.className = 'il-splash';
    splash.innerHTML = `
      <div class="il-splash-glow"></div>
      <img src="${logoSrc}" alt="C.C.R.M" class="il-splash-logo" />
      <div class="il-splash-title">Comunidad Cristiana Restaurando los Muros</div>
      <div class="il-splash-sub">Reporte de métricas</div>
      <div class="il-splash-dots"><span></span><span></span><span></span></div>
    `;
    document.body.appendChild(splash);

    let dismissed = false;
    const dismiss = () => {
      if (dismissed) return;
      dismissed = true;
      splash.classList.add('il-splash-out');
      setTimeout(() => splash.remove(), 650);
    };

    // Se retira sola tras la secuencia de entrada; un clic la salta antes.
    const autoTimer = setTimeout(dismiss, prefersReducedMotion ? 200 : 2200);
    splash.addEventListener('click', () => { clearTimeout(autoTimer); dismiss(); });
  }

  /* ────────────────────────────────────────────────────────────
     5) GRÁFICOS CON VIDA — animación de entrada más expresiva
        para todos los Chart.js del dashboard (crecen/giran al
        aparecer), aplicada vía Chart.defaults ANTES de que
        ChartEngine cree las instancias. No modifica app.js.
  ──────────────────────────────────────────────────────────── */
  function initChartLife() {
    if (typeof Chart === 'undefined') {
      // Chart.js aún no cargó (carga con defer) — reintenta pronto
      setTimeout(initChartLife, 150);
      return;
    }
    if (prefersReducedMotion) return;

    /* NOTA: se usa solo `animation` (duración/easing globales), NO un
       override de `animations.y` con `from` dependiente de las escalas.
       Ese enfoque se probó antes pero rompía los gráficos al cambiar de
       tema: ThemeEngine destruye y vuelve a crear cada Chart, y en ese
       primer frame las escalas aún no están calculadas, por lo que el
       callback `from` devolvía un valor inválido y el gráfico quedaba
       invisible. La animación estándar de Chart.js ya anima barras y
       arcos de forma robusta sin depender de las escalas.

       IMPORTANTE: se FUSIONA con Object.assign en vez de reemplazar
       `Chart.defaults.animation` por un objeto nuevo. Sobrescribirlo
       por completo (`Chart.defaults.animation = {...}`) borra
       cualquier propiedad interna que Chart.js espere ahí además de
       duration/easing, lo que puede dejar el motor de animación en un
       estado inconsistente y, con él, la re-pintura que dispara el
       hover/click sobre leyenda y tooltips — los gráficos quedan
       "pegados" como una imagen estática en vez de reaccionar al
       mouse. Con Object.assign solo se tocan duration/easing y todo
       lo demás que Chart.js trae por defecto queda intacto. */
    Object.assign(Chart.defaults.animation, {
      duration: 500,
      easing: 'easeOutQuart',
    });
    Chart.defaults.transitions.active.animation.duration = 300;
  }

  /* ────────────────────────────────────────────────────────────
     6) CAMBIO DE TEMA — barrido circular expandiéndose desde el
        botón, puramente decorativo por encima del cambio real
        (que sigue manejando ThemeEngine en app.js).
  ──────────────────────────────────────────────────────────── */
  function initThemeWipe() {
    const btn = document.getElementById('btnTheme');
    if (!btn || prefersReducedMotion) return;

    btn.addEventListener('click', (ev) => {
      const goingLight = document.documentElement.getAttribute('data-theme') !== 'light';
      const rect = btn.getBoundingClientRect();
      const cx = rect.left + rect.width / 2;
      const cy = rect.top + rect.height / 2;
      const maxR = Math.hypot(Math.max(cx, window.innerWidth - cx), Math.max(cy, window.innerHeight - cy));

      const wipe = document.createElement('div');
      wipe.className = 'il-theme-wipe';
      wipe.style.left = cx + 'px';
      wipe.style.top = cy + 'px';
      wipe.style.background = goingLight ? '#f0f4f8' : '#0d1117';
      document.body.appendChild(wipe);

      requestAnimationFrame(() => {
        wipe.style.width = wipe.style.height = (maxR * 2.2) + 'px';
        wipe.style.opacity = '0';
      });
      setTimeout(() => wipe.remove(), 650);
    }, true); // captura: se dispara junto con (no reemplaza) el listener real de ThemeEngine
  }

  /* ────────────────────────────────────────────────────────────
     7) TABLAS EN CASCADA — cuando TableEngine reemplaza las filas
        (tbody.innerHTML = ...), cada <tr> nueva entra en cascada
        con un pequeño fade + slide, en vez de aparecer de golpe.
        Se engancha por MutationObserver a cada <tbody>; no toca
        TableEngine ni cómo se generan las filas.
  ──────────────────────────────────────────────────────────── */
  function initTableCascade() {
    if (prefersReducedMotion) return;
    const MAX_STAGGERED = 25; // más allá de esto, delay fijo (evita esperas largas en tablas grandes)
    const STEP_MS = 22;

    document.querySelectorAll('tbody').forEach((tbody) => {
      const observer = new MutationObserver((mutations) => {
        const addedRows = [];
        mutations.forEach((m) => {
          m.addedNodes.forEach((n) => {
            if (n.nodeType === 1 && n.tagName === 'TR') addedRows.push(n);
          });
        });
        if (!addedRows.length) return;

        // Si TableEngine reemplazó TODO el contenido (innerHTML =), todas
        // las filas llegan en la misma tanda de mutaciones → animamos todas.
        const rows = tbody.querySelectorAll('tr');
        rows.forEach((tr, i) => {
          tr.classList.remove('il-row-in');
          void tr.offsetWidth;
          tr.style.animationDelay = (Math.min(i, MAX_STAGGERED) * STEP_MS) + 'ms';
          tr.classList.add('il-row-in');
        });
      });
      observer.observe(tbody, { childList: true });
    });
  }

  /* ────────────────────────────────────────────────────────────
     8) LOADING OVERLAY — anillo de progreso giratorio detrás del
        icono, inyectado una sola vez. Puramente decorativo.
  ──────────────────────────────────────────────────────────── */
  function initLoadingRing() {
    const overlay = document.getElementById('loadingOverlay');
    const content = overlay?.querySelector('.loading-content');
    if (!overlay || !content || content.querySelector('.il-ring')) return;

    const ring = document.createElement('div');
    ring.className = 'il-ring';
    ring.innerHTML = `<svg viewBox="0 0 80 80"><circle class="il-ring-track" cx="40" cy="40" r="34"/><circle class="il-ring-progress" cx="40" cy="40" r="34"/></svg>`;
    content.insertBefore(ring, content.firstChild);
    content.classList.add('il-loading-content-live');
  }

  /* ────────────────────────────────────────────────────────────
     9) REVELADO DEL DASHBOARD — cuando el overlay de login se
        desvanece (login-overlay-fadeout), el topbar, KPIs y
        gráficos entran con un barrido escalonado más teatral.
  ──────────────────────────────────────────────────────────── */
  function initDashboardReveal() {
    const overlay = document.getElementById('loginOverlay');
    if (!overlay || prefersReducedMotion) return;

    const obs = new MutationObserver(() => {
      if (overlay.classList.contains('login-overlay-fadeout')) {
        const topbar = document.querySelector('.topbar');
        if (topbar) {
          topbar.classList.remove('il-reveal');
          void topbar.offsetWidth;
          topbar.classList.add('il-reveal');
        }
        // Destello breve de celebración al entrar (sutil, no invasivo)
        const flash = document.createElement('div');
        flash.className = 'il-success-flash';
        document.body.appendChild(flash);
        setTimeout(() => flash.remove(), 900);
      }
    });
    obs.observe(overlay, { attributes: true, attributeFilter: ['class'] });
  }

  /* ── Init ── */
  function init() {
    initSplashScreen();
    initRipple();
    initTilt();
    initKpiCountUp();
    initLoginLife();
    initChartLife();
    initThemeWipe();
    initTableCascade();
    initLoadingRing();
    initDashboardReveal();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
