// Carrusel automático del panel de marca del login

const INTERVALO = 5000;

export function iniciarCarrusel(raiz) {
  if (!raiz) return;
  const slides = [...raiz.querySelectorAll(".slide")];
  const dots = [...raiz.querySelectorAll(".dot")];
  if (slides.length < 2) return;

  const sinMovimiento = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  let actual = 0;
  let timer = null;

  function mostrar(i) {
    actual = (i + slides.length) % slides.length;
    slides.forEach((s, n) => {
      s.classList.toggle("activa", n === actual);
      s.setAttribute("aria-hidden", n === actual ? "false" : "true");
    });
    dots.forEach((d, n) => {
      d.classList.toggle("activa", n === actual);
      d.setAttribute("aria-current", n === actual ? "true" : "false");
    });
  }

  function iniciar() {
    if (sinMovimiento || timer) return;
    timer = setInterval(() => mostrar(actual + 1), INTERVALO);
  }

  function detener() {
    clearInterval(timer);
    timer = null;
  }

  dots.forEach((d, n) => d.addEventListener("click", () => { mostrar(n); detener(); iniciar(); }));
  raiz.addEventListener("mouseenter", detener);
  raiz.addEventListener("mouseleave", iniciar);
  document.addEventListener("visibilitychange", () => (document.hidden ? detener() : iniciar()));

  mostrar(0);
  iniciar();
}
