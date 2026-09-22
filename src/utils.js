// Utilidades compartidas por los módulos de Personal e Inventario

export function escapeHtml(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

const pad2 = (n) => String(n).padStart(2, "0");

// Convierte la fecha del Excel (número de serie, texto dd/mm/aaaa o aaaa-mm-dd) a "aaaa-mm-dd".
// Devuelve null si no se reconoce, para no detener la importación.
export function normalizarFecha(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    return d && d.y > 1900 ? `${d.y}-${pad2(d.m)}-${pad2(d.d)}` : null;
  }
  const s = String(v).trim();
  if (!s) return "";
  let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/);
  if (m) {
    const dia = +m[1], mes = +m[2];
    const anio = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    if (mes >= 1 && mes <= 12 && dia >= 1 && dia <= 31) return `${anio}-${pad2(mes)}-${pad2(dia)}`;
    return null;
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  return null;
}

// DNI como texto; si Excel lo guardó como número se recuperan los ceros a la izquierda (8 dígitos)
export function textoDni(v) {
  if (typeof v === "number") return String(v).padStart(8, "0");
  return String(v ?? "").trim();
}

// Clave para comparar DNIs sin importar ceros a la izquierda ni espacios
export function claveDni(v) {
  return textoDni(v).replace(/\s+/g, "").replace(/^0+/, "").toUpperCase();
}

// Quita espacios de los nombres de columna (p. ej. "DNI " -> "DNI")
export function limpiarClaves(row) {
  const r = {};
  for (const [k, v] of Object.entries(row)) r[String(k).trim()] = v;
  return r;
}

export const ICONO_EDITAR = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>';

// Números de página con "…": 1 … 4 5 6 … 12
function rangoPaginas(pagina, paginas) {
  const set = new Set([1, paginas, pagina - 1, pagina, pagina + 1]);
  const nums = [...set].filter(n => n >= 1 && n <= paginas).sort((a, b) => a - b);
  const res = [];
  nums.forEach((n, i) => {
    if (i && n - nums[i - 1] > 1) res.push("…");
    res.push(n);
  });
  return res;
}

// HTML de la barra de paginación ("Mostrando X–Y de N" + botones con data-pag)
export function htmlPaginacion(total, pagina, paginas, porPagina) {
  const desde = (pagina - 1) * porPagina;
  const hasta = Math.min(desde + porPagina, total);
  const info = total ? `Mostrando <b>${desde + 1}–${hasta}</b> de <b>${total}</b> registros` : "0 registros";
  const botones = rangoPaginas(pagina, paginas).map(n => n === "…"
    ? `<span class="pag-sep">…</span>`
    : `<button type="button" class="pag-btn${n === pagina ? " activa" : ""}" data-pag="${n}"${n === pagina ? ' aria-current="page"' : ""}>${n}</button>`
  ).join("");
  return `
    <span class="pag-info">${info}</span>
    <div class="pag-controles">
      <button type="button" class="pag-btn" data-pag="${pagina - 1}"${pagina === 1 ? " disabled" : ""} aria-label="Página anterior">‹ Anterior</button>
      ${botones}
      <button type="button" class="pag-btn" data-pag="${pagina + 1}"${pagina === paginas ? " disabled" : ""} aria-label="Página siguiente">Siguiente ›</button>
    </div>`;
}

// Conecta una zona de arrastre + botón + input de archivo a un manejador
export function conectarImportador({ zona, boton, input, alSoltar }) {
  boton.addEventListener("click", () => input.click());
  zona.addEventListener("dragover", (e) => { e.preventDefault(); zona.classList.add("dragover"); });
  zona.addEventListener("dragleave", () => zona.classList.remove("dragover"));
  zona.addEventListener("drop", (e) => {
    e.preventDefault();
    zona.classList.remove("dragover");
    if (e.dataTransfer.files[0]) alSoltar(e.dataTransfer.files[0]);
  });
  input.addEventListener("change", (e) => { if (e.target.files[0]) alSoltar(e.target.files[0]); });
}
