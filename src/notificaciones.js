// Notificaciones tipo "toast" y ventana de confirmación, en lugar de alert/confirm del navegador

const ICONOS = {
  success: '<path d="M20 6 9 17l-5-5"/>',
  error: '<circle cx="12" cy="12" r="9"/><path d="m15 9-6 6M9 9l6 6"/>',
  warning: '<path d="M12 3 2 20h20L12 3z"/><path d="M12 10v4M12 17.5v.01"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 7.5v.01"/>'
};

function icono(tipo) {
  return `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONOS[tipo] || ICONOS.info}</svg>`;
}

function contenedor() {
  let c = document.getElementById("toasts");
  if (!c) {
    c = document.createElement("div");
    c.id = "toasts";
    c.className = "toasts";
    c.setAttribute("aria-live", "polite");
    document.body.appendChild(c);
  }
  return c;
}

/**
 * Muestra una notificación.
 * @param {"success"|"error"|"warning"|"info"} tipo
 * @param {string} titulo
 * @param {string} [mensaje]
 * @param {number} [duracion] ms; 0 = no se cierra sola
 */
export function notificar(tipo, titulo, mensaje = "", duracion) {
  const ms = duracion ?? (tipo === "error" ? 8000 : 4500);
  const t = document.createElement("div");
  t.className = `toast toast-${tipo}`;
  t.setAttribute("role", tipo === "error" ? "alert" : "status");
  t.innerHTML = `
    <span class="toast-icon">${icono(tipo)}</span>
    <div class="toast-body"><strong></strong><p></p></div>
    <button type="button" class="toast-close" aria-label="Cerrar">×</button>
    ${ms ? `<span class="toast-bar" style="animation-duration:${ms}ms"></span>` : ""}`;
  t.querySelector("strong").textContent = titulo;
  const p = t.querySelector("p");
  if (mensaje) p.textContent = mensaje; else p.remove();

  const cerrar = () => {
    t.classList.add("saliendo");
    setTimeout(() => t.remove(), 200);
  };
  t.querySelector(".toast-close").addEventListener("click", cerrar);
  contenedor().appendChild(t);
  if (ms) setTimeout(cerrar, ms);
  return cerrar;
}

/**
 * Ventana de confirmación. Devuelve una promesa con true/false.
 */
export function confirmar({ titulo, mensaje = "", aceptar = "Aceptar", cancelar = "Cancelar", peligro = false }) {
  return new Promise((resolve) => {
    const fondo = document.createElement("div");
    fondo.className = "modal-fondo";
    fondo.innerHTML = `
      <div class="modal" role="alertdialog" aria-modal="true" aria-labelledby="modalTitulo">
        <span class="modal-icon ${peligro ? "peligro" : ""}">${icono(peligro ? "warning" : "info")}</span>
        <h3 id="modalTitulo"></h3>
        <p></p>
        <div class="modal-acciones">
          <button type="button" class="btn-outline" data-r="0"></button>
          <button type="button" class="${peligro ? "btn-peligro" : "btn-confirmar"}" data-r="1"></button>
        </div>
      </div>`;
    fondo.querySelector("h3").textContent = titulo;
    const p = fondo.querySelector("p");
    if (mensaje) p.textContent = mensaje; else p.remove();
    fondo.querySelector('[data-r="0"]').textContent = cancelar;
    const btnOk = fondo.querySelector('[data-r="1"]');
    btnOk.textContent = aceptar;

    const cerrar = (r) => {
      document.removeEventListener("keydown", teclas);
      fondo.remove();
      resolve(r);
    };
    const teclas = (e) => { if (e.key === "Escape") cerrar(false); };
    fondo.addEventListener("click", (e) => {
      if (e.target === fondo) cerrar(false);
      const r = e.target.closest("[data-r]");
      if (r) cerrar(r.dataset.r === "1");
    });
    document.addEventListener("keydown", teclas);
    document.body.appendChild(fondo);
    btnOk.focus();
  });
}
