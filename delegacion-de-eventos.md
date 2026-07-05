# Por qué "delegar" y no solo "reenganchar"

Hay dos estrategias válidas y se pueden mezclar:

## Estrategia A — Reenganche directo en `_onViewLoaded` (la que usé arriba)

Cada motor expone un método `initListeners()` que hace `getElementById` +
`addEventListener`. Se llama una vez por vista cargada. Es simple y explícita,
pero si el usuario navega login → dashboard → login → dashboard varias veces
**sin limpiar**, puedes duplicar listeners (el mismo click dispara 2 o 3 veces).

Como en tu app el HTML se reemplaza por completo (`container.innerHTML = html`),
los nodos viejos —y sus listeners— se destruyen solos con el garbage collector.
Por eso la Estrategia A es segura en tu caso concreto: no hay riesgo de
duplicar, porque el botón viejo ya no existe.

## Estrategia B — Delegación en un ancestro estable (recomendada si el
`container` NO se reemplaza completo, por ejemplo si solo cambias un
`<section>` interno)

En vez de escuchar en el botón, escuchas en `#app-container` (que nunca se
destruye) y filtras por `closest()`:

```js
// Se registra UNA sola vez, en el arranque de la app — nunca se re-liga.
document.getElementById('app-container').addEventListener('click', (e) => {
  const btn = e.target.closest('[data-action]');
  if (!btn) return;

  switch (btn.dataset.action) {
    case 'login-start':
      SessionEngine.mostrarInputNombre();
      break;
    case 'login-confirm':
      SessionEngine.confirmarNombre();
      break;
    case 'menu-eliminar':
      // el modal lo abre Bootstrap vía data-bs-toggle, esto es solo
      // para lógica de negocio adicional si la necesitas
      AuditEngine.registrarIntento('eliminar');
      break;
  }
});
```

Y en el HTML, en vez de `id="btnIniciarSesion"` usarías:

```html
<button type="button" class="login-btn-start" data-action="login-start">
  Iniciar Sesión
</button>
```

### ¿Cuál usar en tu proyecto?

Con tu patrón actual (`container.innerHTML = html` — reemplazo total en cada
`loadView`), **la Estrategia A es suficiente y más fácil de mantener**: no
tienes que tocar tus IDs existentes ni tu `AuditEngine`/`TelegramEngine`, que
ya buscan por `id`. Usa la Estrategia B únicamente si en el futuro divides el
dashboard en sub-vistas que se intercambian dentro de sí mismo (por ejemplo,
tabs que se cargan por fetch en vez de con Bootstrap `.tab-pane`).

## Regla de oro para `TelegramEngine` y `AuditEngine`

Estos dos son objetos globales cargados por `<script>`, no dependen del DOM
para existir — solo sus *listeners de UI* dependen del DOM. Es decir:
`TelegramEngine.notify(...)` puedes llamarlo desde cualquier `_onViewLoaded`
sin problema, porque el objeto en sí vive en memoria desde el primer
`<script src="TelegramEngine.js">`, independientemente de qué vista esté
montada.
