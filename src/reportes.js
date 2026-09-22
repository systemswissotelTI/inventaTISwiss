// Reportes dinámicos: reproducen las tablas dinámicas del Excel (Dinamica, ImpresoasDinamica, Micros)
// y muestran los catálogos de la hoja "Datos". Todo se calcula en el navegador a partir del inventario.

import { escapeHtml } from "./utils.js";
import { notificar } from "./notificaciones.js";
import { exportarExcel } from "./exportar.js";

const EN_BLANCO = "(en blanco)";

const REPORTES = [
  {
    id: "computo", nombre: "Equipos de cómputo", hoja: "Dinamica",
    equipos: ["PC", "Laptop"], estados: ["Operativo"],
    grupos: ["Propiedad", "Unidad", "Centro de Costo"],
    columnas: ["Usuario", "Cargo", "Equipo", "Nuevo Nombre Host", "Nombre Host", "IP", "Tipo", "Marca", "S/N", "Estado", "Ram", "Almacenamiento", "Sistema Operativo", "Office"]
  },
  {
    id: "impresoras", nombre: "Impresoras", hoja: "ImpresoasDinamica",
    equipos: ["Impresora"], estados: [],
    grupos: ["Propiedad", "Marca", "Modelo"],
    columnas: ["S/N", "IP", "Observación", "Estado"]
  },
  {
    id: "micros", nombre: "Micros", hoja: "Micros",
    equipos: ["Micros"], estados: ["Operativo", "Custodia - Operativo"],
    grupos: ["Propiedad", "Unidad", "Centro de Costo"],
    columnas: ["Cargo", "Usuario", "Marca", "Modelo", "S/N", "IP", "Estado"]
  },
  {
    id: "celulares", nombre: "Celulares y tablets",
    equipos: ["Celular", "Tablet"], estados: [],
    grupos: ["Unidad", "Centro de Costo"],
    columnas: ["Usuario", "Cargo", "Equipo", "Marca", "Modelo", "Operador", "Linea", "IMEI1", "Estado"]
  },
  {
    id: "todo", nombre: "Todo el inventario",
    equipos: [], estados: [],
    grupos: ["Equipo", "Estado"],
    columnas: ["Item", "Usuario", "Unidad", "Marca", "Modelo", "S/N", "IP", "Nombre Host"]
  },
  { id: "catalogos", nombre: "Catálogos (Datos)", catalogos: true }
];

const AGRUPACIONES = [
  ["Propiedad", "Unidad", "Centro de Costo"],
  ["Unidad", "Usuario"],
  ["Equipo", "Marca", "Modelo"],
  ["Propiedad", "Marca", "Modelo"],
  ["Estado", "Equipo"],
  ["Equipo", "Estado"],
  ["Unidad", "Centro de Costo"],
  []
];

let ctx = { obtenerFilas: () => [], obtenerCatalogos: () => ({}) };
let el = {};
let reporte = REPORTES[0];
let filtro = {};
let plegados = new Set();

const texto = (v) => {
  const s = String(v ?? "").trim();
  return s && s !== "#N/A" ? s : EN_BLANCO;
};

export function iniciarReportes(opciones) {
  ctx = { ...ctx, ...opciones };
  el = {
    presets: document.getElementById("repPresets"),
    cuerpo: document.getElementById("repCuerpo"),
    catalogos: document.getElementById("repCatalogos"),
    kpis: document.getElementById("repKpis"),
    fEquipo: document.getElementById("repFEquipo"),
    fEstado: document.getElementById("repFEstado"),
    fPropiedad: document.getElementById("repFPropiedad"),
    fUnidad: document.getElementById("repFUnidad"),
    agrupar: document.getElementById("repAgrupar"),
    buscar: document.getElementById("repBuscar"),
    thead: document.getElementById("repThead"),
    tbody: document.getElementById("repTbody"),
    exportar: document.getElementById("repExportar"),
    plegar: document.getElementById("repPlegar"),
    desplegar: document.getElementById("repDesplegar")
  };

  el.presets.innerHTML = REPORTES.map(r =>
    `<button type="button" class="preset" role="tab" data-rep="${r.id}"${r.hoja ? ` title="Equivale a la hoja ${escapeHtml(r.hoja)} del Excel"` : ""}>${escapeHtml(r.nombre)}</button>`).join("");
  el.agrupar.innerHTML = AGRUPACIONES.map(g =>
    `<option value="${escapeHtml(g.join("|"))}">${g.length ? escapeHtml(g.join(" › ")) : "Sin agrupar"}</option>`).join("");

  el.presets.addEventListener("click", (e) => {
    const b = e.target.closest("[data-rep]");
    if (b) elegirReporte(b.dataset.rep);
  });
  const alternar = (conjunto) => (e) => {
    const b = e.target.closest("[data-v]");
    if (!b) return;
    const v = b.dataset.v;
    if (v === "") conjunto.clear();
    else if (conjunto.has(v)) conjunto.delete(v);
    else conjunto.add(v);
    pintar();
  };
  el.fEquipo.addEventListener("click", (e) => alternar(filtro.equipos)(e));
  el.fEstado.addEventListener("click", (e) => alternar(filtro.estados)(e));
  el.fPropiedad.addEventListener("change", () => { filtro.propiedad = el.fPropiedad.value; pintar(); });
  el.fUnidad.addEventListener("change", () => { filtro.unidad = el.fUnidad.value; pintar(); });
  el.agrupar.addEventListener("change", () => { filtro.grupos = el.agrupar.value ? el.agrupar.value.split("|") : []; plegados.clear(); pintar(); });
  el.buscar.addEventListener("input", () => { filtro.q = el.buscar.value.trim().toLowerCase(); pintar(); });
  el.tbody.addEventListener("click", (e) => {
    const fila = e.target.closest("tr[data-ruta]");
    if (!fila) return;
    const ruta = fila.dataset.ruta;
    if (plegados.has(ruta)) plegados.delete(ruta); else plegados.add(ruta);
    pintar();
  });
  el.plegar.addEventListener("click", () => { plegados = new Set(rutasDeGrupo()); pintar(); });
  el.desplegar.addEventListener("click", () => { plegados.clear(); pintar(); });
  el.exportar.addEventListener("click", exportar);

  let guardado = null;
  try { guardado = localStorage.getItem("reporte"); } catch {}
  elegirReporte(REPORTES.some(r => r.id === guardado) ? guardado : REPORTES[0].id);
}

function elegirReporte(id) {
  reporte = REPORTES.find(r => r.id === id) || REPORTES[0];
  try { localStorage.setItem("reporte", reporte.id); } catch {}
  el.presets.querySelectorAll(".preset").forEach(b => {
    const activo = b.dataset.rep === reporte.id;
    b.classList.toggle("activa", activo);
    b.setAttribute("aria-selected", activo ? "true" : "false");
  });
  el.cuerpo.hidden = Boolean(reporte.catalogos);
  el.catalogos.hidden = !reporte.catalogos;
  el.exportar.hidden = Boolean(reporte.catalogos);
  if (reporte.catalogos) { pintarCatalogos(); return; }

  filtro = {
    equipos: new Set(reporte.equipos),
    estados: new Set(reporte.estados),
    propiedad: "",
    unidad: "",
    grupos: [...reporte.grupos],
    q: ""
  };
  plegados.clear();
  el.buscar.value = "";
  el.agrupar.value = filtro.grupos.join("|");
  pintar();
}

export function refrescarReportes() {
  if (!el.tbody) return;
  if (reporte.catalogos) pintarCatalogos(); else pintar();
}

// ---------- Datos filtrados ----------

function filasFiltradas() {
  const q = filtro.q;
  return ctx.obtenerFilas().filter(r => {
    if (filtro.equipos.size && !filtro.equipos.has(texto(r.Equipo))) return false;
    if (filtro.estados.size && !filtro.estados.has(texto(r.Estado))) return false;
    if (filtro.propiedad && texto(r.Propiedad) !== filtro.propiedad) return false;
    if (filtro.unidad && texto(r.Unidad) !== filtro.unidad) return false;
    if (q && !Object.entries(r).some(([k, v]) => k !== "Clave Gmail" && String(v ?? "").toLowerCase().includes(q))) return false;
    return true;
  });
}

function contar(filas, k) {
  const m = new Map();
  filas.forEach(r => { const v = texto(r[k]); m.set(v, (m.get(v) || 0) + 1); });
  return [...m.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "es"));
}

function chipsToggle(cont, conteo, seleccion) {
  // Los valores seleccionados se muestran aunque no haya equipos con ellos
  const valores = new Map(conteo);
  seleccion.forEach(v => { if (!valores.has(v)) valores.set(v, 0); });
  cont.innerHTML = `<button type="button" class="chip-toggle${seleccion.size ? "" : " activa"}" data-v="">Todos</button>`
    + [...valores.entries()].map(([v, n]) =>
      `<button type="button" class="chip-toggle${seleccion.has(v) ? " activa" : ""}" data-v="${escapeHtml(v)}" aria-pressed="${seleccion.has(v)}">${escapeHtml(v)} <b>${n}</b></button>`).join("");
}

function opciones(select, conteo, actual, todos) {
  select.innerHTML = `<option value="">${todos}</option>` + conteo.map(([v, n]) => `<option value="${escapeHtml(v)}">${escapeHtml(v)} (${n})</option>`).join("");
  select.value = conteo.some(([v]) => v === actual) ? actual : "";
}

// ---------- Tabla agrupada ----------

function arbol(filas, grupos) {
  if (!grupos.length) return filas;
  const [k, ...resto] = grupos;
  const m = new Map();
  filas.forEach(r => { const v = texto(r[k]); if (!m.has(v)) m.set(v, []); m.get(v).push(r); });
  return [...m.entries()]
    .sort((a, b) => (a[0] === EN_BLANCO) - (b[0] === EN_BLANCO) || a[0].localeCompare(b[0], "es"))
    .map(([v, hijos]) => ({ campo: k, valor: v, total: hijos.length, hijos: arbol(hijos, resto) }));
}

let ultimasRutas = [];
function rutasDeGrupo() { return ultimasRutas; }

function pintar() {
  const todas = ctx.obtenerFilas();
  const base = todas.filter(r =>
    (!filtro.propiedad || texto(r.Propiedad) === filtro.propiedad) &&
    (!filtro.unidad || texto(r.Unidad) === filtro.unidad));
  chipsToggle(el.fEquipo, contar(base, "Equipo"), filtro.equipos);
  chipsToggle(el.fEstado, contar(base.filter(r => !filtro.equipos.size || filtro.equipos.has(texto(r.Equipo))), "Estado"), filtro.estados);
  opciones(el.fPropiedad, contar(todas, "Propiedad"), filtro.propiedad, "Todas las propiedades");
  opciones(el.fUnidad, contar(todas, "Unidad"), filtro.unidad, "Todas las unidades");

  const filas = filasFiltradas();
  pintarKpis(filas);

  const columnas = reporte.columnas.filter(c => !filtro.grupos.includes(c));
  const nCols = columnas.length + 1;
  el.thead.innerHTML = `<tr><th class="th-grupo">${filtro.grupos.length ? escapeHtml(filtro.grupos.join(" › ")) : "#"}</th>${columnas.map(c => `<th>${escapeHtml(c)}</th>`).join("")}</tr>`;

  if (!todas.length) {
    el.tbody.innerHTML = `<tr><td colspan="${nCols}" class="vacio">No hay equipos. Importa el Excel en la pestaña Inventario TI.</td></tr>`;
    return;
  }
  if (!filas.length) {
    el.tbody.innerHTML = `<tr><td colspan="${nCols}" class="vacio">Ningún equipo cumple los filtros.</td></tr>`;
    return;
  }

  const html = [];
  const rutas = [];
  let n = 0;
  const hoja = (r, nivel) => {
    n++;
    html.push(`<tr class="rep-fila"><td class="rep-num" style="padding-left:${12 + nivel * 18}px">${n}</td>${columnas.map(c => {
      const v = String(r[c] ?? "").trim();
      if (c === "Estado" && v) {
        const tono = /inoperativo|obsoleto/i.test(v) ? "rojo" : /revisar|reparar|repotenciar|custodia/i.test(v) ? "ambar" : "verde";
        return `<td><span class="estado estado-${tono}">${escapeHtml(v)}</span></td>`;
      }
      return `<td${v === "#N/A" ? ' class="na"' : ""}>${escapeHtml(v)}</td>`;
    }).join("")}</tr>`);
  };
  const recorrer = (nodos, nivel, ruta) => {
    nodos.forEach(nodo => {
      if (!nodo.hijos) { hoja(nodo, nivel); return; }
      const r = `${ruta}${nodo.valor}\u0001`;
      rutas.push(r);
      const plegado = plegados.has(r);
      html.push(`<tr class="rep-grupo nivel-${Math.min(nivel, 3)}" data-ruta="${escapeHtml(r)}" aria-expanded="${!plegado}"><td colspan="${nCols}" style="padding-left:${12 + nivel * 18}px"><span class="flecha">${plegado ? "▸" : "▾"}</span><span class="grupo-campo">${escapeHtml(nodo.campo)}:</span> <b>${escapeHtml(nodo.valor)}</b> <span class="grupo-cuenta">${nodo.total}</span></td></tr>`);
      if (plegado) { n += nodo.total; return; }
      recorrer(nodo.hijos, nivel + 1, r);
    });
  };
  recorrer(arbol(filas, filtro.grupos), 0, "");
  html.push(`<tr class="rep-total"><td colspan="${nCols}">Total general <span class="grupo-cuenta">${filas.length}</span></td></tr>`);
  el.tbody.innerHTML = html.join("");
  ultimasRutas = rutas;
}

function pintarKpis(filas) {
  const operativos = filas.filter(r => /^operativo$/i.test(String(r.Estado ?? "").trim())).length;
  const custodia = filas.filter(r => /custodia/i.test(String(r.Estado ?? ""))).length;
  const usuarios = new Set(filas.map(r => texto(r.Usuario)).filter(v => v !== EN_BLANCO)).size;
  const sinAsignar = filas.filter(r => texto(r.Usuario) === EN_BLANCO).length;
  const pct = filas.length ? Math.round(operativos / filas.length * 100) : 0;
  const kpi = (titulo, valor, detalle, tono = "") =>
    `<div class="kpi ${tono}"><span class="kpi-titulo">${titulo}</span><span class="kpi-valor">${valor}</span><span class="kpi-detalle">${detalle}</span></div>`;
  el.kpis.innerHTML =
    kpi("Equipos", filas.length, "según los filtros") +
    kpi("Operativos", operativos, `${pct}% del total`, "verde") +
    kpi("En custodia", custodia, "guardados por TI", "ambar") +
    kpi("Usuarios / áreas", usuarios, `${sinAsignar} sin asignar`, sinAsignar ? "rojo" : "");
}

// ---------- Catálogos (hoja Datos) ----------

function pintarCatalogos() {
  const cat = ctx.obtenerCatalogos();
  const claves = Object.keys(cat);
  el.catalogos.innerHTML = claves.length
    ? `<p class="ayuda">Listas de valores de la hoja <b>Datos</b> del Excel. Se usan como sugerencias en el formulario de equipos y se actualizan al importar el inventario.</p>
       <div class="catalogos">${claves.map(k => `<div class="catalogo"><h3>${escapeHtml(k)} <span class="grupo-cuenta">${cat[k].length}</span></h3><div class="catalogo-valores">${cat[k].map(v => `<span class="valor">${escapeHtml(v)}</span>`).join("")}</div></div>`).join("")}</div>`
    : `<p class="vacio">Aún no hay catálogos. Importa el Excel en la pestaña <b>Inventario TI</b>: la hoja <b>Datos</b> se carga automáticamente.</p>`;
}

// ---------- Exportar ----------

async function exportar() {
  const filas = filasFiltradas();
  if (!filas.length) { notificar("info", "No hay datos que exportar"); return; }
  const columnas = reporte.columnas.filter(c => !filtro.grupos.includes(c));

  // Mismo árbol que en pantalla: filas de grupo con su recuento y debajo los equipos
  const salida = [];
  const recorrer = (nodos, nivel) => nodos.forEach(nodo => {
    if (!nodo.hijos) { salida.push(nodo); return; }
    salida.push({ grupo: { nivel, texto: `${nodo.campo}: ${nodo.valor}  (${nodo.total})` } });
    recorrer(nodo.hijos, nivel + 1);
  });
  recorrer(arbol(filas, filtro.grupos), 0);

  const partes = [];
  if (filtro.equipos.size) partes.push(`Equipo: ${[...filtro.equipos].join(", ")}`);
  if (filtro.estados.size) partes.push(`Estado: ${[...filtro.estados].join(", ")}`);
  if (filtro.propiedad) partes.push(`Propiedad: ${filtro.propiedad}`);
  if (filtro.unidad) partes.push(`Unidad: ${filtro.unidad}`);
  if (filtro.q) partes.push(`Búsqueda: "${filtro.q}"`);
  if (filtro.grupos.length) partes.push(`Agrupado por ${filtro.grupos.join(" › ")}`);

  el.exportar.disabled = true;
  try {
    await exportarExcel({
      archivo: `Reporte ${reporte.nombre}`,
      hoja: reporte.nombre,
      titulo: `Reporte · ${reporte.nombre}`,
      detalle: partes.join(" · "),
      columnas,
      filas: salida,
      total: `Total general: ${filas.length} equipos`
    });
    notificar("success", "Reporte exportado", `${filas.length} equipos en Excel.`);
  } catch (error) {
    notificar("error", "No se pudo exportar", error.message);
  } finally {
    el.exportar.disabled = false;
  }
}
