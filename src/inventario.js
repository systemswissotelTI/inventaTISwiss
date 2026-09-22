// Inventario de equipos TI (hoja "InventarioSwiss" del Excel)
// Código, Usuario, Cargo, Centro de Costo y Unidad no se guardan: se obtienen de Personal por DNI,
// igual que los VLOOKUP del Excel.

import { db } from "./firebase-config.js";
import { collection, addDoc, getDocs, getDoc, setDoc, deleteDoc, updateDoc, doc, writeBatch } from "firebase/firestore";
import { notificar, confirmar } from "./notificaciones.js";
import { exportarExcel } from "./exportar.js";
import {
  escapeHtml, normalizarFecha, textoDni, claveDni, limpiarClaves,
  ICONO_EDITAR, htmlPaginacion, conectarImportador
} from "./utils.js";

const COLECCION = "inventario";
const HOJA = "InventarioSwiss";
const HOJA_DATOS = "Datos";
const DOC_CATALOGOS = ["catalogos", "datos"];
const POR_PAGINA = 50;

// Columnas en el mismo orden que la hoja. "de" = campo de Personal del que se obtiene (VLOOKUP por DNI)
const CAMPOS = [
  { k: "Item", grupo: "Adquisición", numero: true },
  { k: "Fecha", grupo: "Adquisición", fecha: true },
  { k: "Pedido", grupo: "Adquisición" },
  { k: "Orden de Compra", grupo: "Adquisición" },
  { k: "GR - Movistar", grupo: "Adquisición" },
  { k: "Equipo", grupo: "Equipo", requerido: true },
  { k: "Tipo", grupo: "Equipo" },
  { k: "Marca", grupo: "Equipo" },
  { k: "Modelo", grupo: "Equipo" },
  { k: "S/N", grupo: "Equipo" },
  { k: "P/N", grupo: "Equipo" },
  { k: "Operador", grupo: "Línea móvil" },
  { k: "Linea", grupo: "Línea móvil" },
  { k: "IMEI1", grupo: "Línea móvil" },
  { k: "IMEI2", grupo: "Línea móvil" },
  { k: "Cuenta Gmail", grupo: "Línea móvil" },
  { k: "Clave Gmail", grupo: "Línea móvil", secreto: true },
  { k: "Mac Address", grupo: "Red" },
  { k: "Mac Address - Wifi", grupo: "Red" },
  { k: "IP", grupo: "Red" },
  { k: "Nombre Host", grupo: "Red" },
  { k: "Nuevo Nombre Host", grupo: "Red" },
  { k: "Procesador", grupo: "Hardware" },
  { k: "Generación Procesador", grupo: "Hardware" },
  { k: "Chip Video", grupo: "Hardware" },
  { k: "Ram", grupo: "Hardware" },
  { k: "Ram Actual", grupo: "Hardware" },
  { k: "Almacenamiento", grupo: "Hardware" },
  { k: "Pantalla", grupo: "Hardware" },
  { k: "Sistema Operativo", grupo: "Software" },
  { k: "Antivirus", grupo: "Software" },
  { k: "Office", grupo: "Software" },
  { k: "Fech. Ini Garantía", grupo: "Garantía y estado", fecha: true },
  { k: "Fech. Fin Garantía", grupo: "Garantía y estado", fecha: true },
  { k: "Propiedad", grupo: "Garantía y estado" },
  { k: "Estado", grupo: "Garantía y estado" },
  { k: "DNI", grupo: "Asignación" },
  { k: "Código", grupo: "Asignación", de: "Código" },
  { k: "Usuario", grupo: "Asignación", de: "Nombre" },
  { k: "Cargo", grupo: "Asignación", de: "Puesto Real" },
  { k: "Centro de Costo", grupo: "Asignación", de: "CC Real" },
  { k: "Unidad", grupo: "Asignación", de: "Unidad Real" },
  { k: "Observación", grupo: "Asignación", largo: true }
];
const GUARDADOS = CAMPOS.filter(c => !c.de);

let col = null;
let equipos = [];
let catalogos = {};
let lista = [];
let pagina = 1;
let ctx = { requiereConfig: () => false, obtenerPersonal: () => [], alCambiar: () => {} };
let el = {};

// ---------- Relación con Personal ----------

function mapaPersonal() {
  const m = new Map();
  ctx.obtenerPersonal().forEach(p => { const k = claveDni(p.DNI); if (k && !m.has(k)) m.set(k, p); });
  return m;
}

// Valor mostrado de una columna; para las derivadas devuelve { na: true } si el DNI no está en Personal
function valor(e, c, personal) {
  if (!c.de) return e[c.k] ?? "";
  const k = claveDni(e.DNI);
  if (!k) return "";
  const p = personal.get(k);
  return p ? (p[c.de] ?? "") : { na: true };
}

// Equipos con las columnas de Personal ya resueltas (para los reportes)
export function equiposResueltos() {
  const personal = mapaPersonal();
  return equipos.map(e => {
    const r = { ...e };
    CAMPOS.filter(c => c.de).forEach(c => {
      const v = valor(e, c, personal);
      r[c.k] = typeof v === "object" ? "#N/A" : v;
    });
    return r;
  });
}

// Listas de la hoja "Datos" del Excel (valores permitidos por columna)
export function obtenerCatalogos() {
  return catalogos;
}

export function contarEquiposPorDni() {
  const m = new Map();
  equipos.forEach(e => { const k = claveDni(e.DNI); if (k) m.set(k, (m.get(k) || 0) + 1); });
  return m;
}

// ---------- Carga ----------

export function iniciarInventario(opciones) {
  ctx = { ...ctx, ...opciones };
  col = db ? collection(db, COLECCION) : null;
  el = {
    thead: document.getElementById("invThead"),
    tbody: document.getElementById("invTbody"),
    wrap: document.getElementById("invTablaWrap"),
    paginacion: document.getElementById("invPaginacion"),
    buscar: document.getElementById("invBuscar"),
    fEquipo: document.getElementById("invFiltroEquipo"),
    fEstado: document.getElementById("invFiltroEstado"),
    resumen: document.getElementById("invResumen"),
    status: document.getElementById("invStatus"),
    statusText: document.getElementById("invStatusText"),
    progreso: document.getElementById("invProgress"),
    input: document.getElementById("invInput")
  };

  el.thead.innerHTML = `<tr><th class="th-acciones">Acciones</th>${CAMPOS.map(c => `<th${c.de ? ' class="th-derivada" title="Se obtiene de Personal por DNI"' : ""}>${escapeHtml(c.k)}</th>`).join("")}</tr>`;

  conectarImportador({
    zona: document.getElementById("invSeccion"),
    boton: document.getElementById("invSelectBtn"),
    input: el.input,
    alSoltar: importar
  });
  el.buscar.addEventListener("input", () => filtrar(true));
  el.fEquipo.addEventListener("change", () => filtrar(true));
  el.fEstado.addEventListener("change", () => filtrar(true));
  el.paginacion.addEventListener("click", (e) => {
    const b = e.target.closest("[data-pag]");
    if (!b || b.disabled) return;
    pagina = Number(b.dataset.pag);
    pintar();
  });
  el.resumen.addEventListener("click", (e) => {
    const b = e.target.closest("[data-equipo]");
    if (!b) return;
    el.fEquipo.value = b.dataset.equipo;
    filtrar(true);
  });
  document.getElementById("invNuevo").addEventListener("click", () => abrirFormulario(null));
  document.getElementById("invBorrarTodo").addEventListener("click", borrarTodo);
  document.getElementById("invExportar").addEventListener("click", (e) => exportarLista(e.currentTarget));
}

// Descarga en Excel (con formato) los equipos que se ven con los filtros actuales
async function exportarLista(boton) {
  if (!lista.length) { notificar("info", "No hay equipos que exportar"); return; }
  const personal = mapaPersonal();
  const columnas = CAMPOS.filter(c => !c.secreto);
  const filas = lista.map(e => Object.fromEntries(columnas.map(c => {
    const v = valor(e, c, personal);
    return [c.k, typeof v === "object" ? "#N/A" : v];
  })));
  const filtros = [el.fEquipo.value && `Equipo: ${el.fEquipo.value}`, el.fEstado.value && `Estado: ${el.fEstado.value}`,
    el.buscar.value.trim() && `Búsqueda: "${el.buscar.value.trim()}"`].filter(Boolean).join(" · ");
  boton.disabled = true;
  try {
    await exportarExcel({
      archivo: "Inventario TI",
      hoja: "Inventario TI",
      titulo: "Inventario de equipos TI",
      detalle: filtros || "Todos los equipos",
      columnas: columnas.map(c => c.k),
      filas,
      total: `Total: ${filas.length} equipos`
    });
    notificar("success", "Inventario exportado", `${filas.length} equipos en Excel.`);
  } catch (error) {
    notificar("error", "No se pudo exportar", error.message);
  } finally {
    boton.disabled = false;
  }
}

export async function cargarInventario() {
  if (!ctx.requiereConfig() || !col) return;
  try {
    const snap = await getDocs(col);
    equipos = [];
    snap.forEach(d => equipos.push({ id: d.id, ...d.data() }));
    equipos.sort((a, b) => (Number(a.Item) || Infinity) - (Number(b.Item) || Infinity));
    try {
      const cat = await getDoc(doc(db, ...DOC_CATALOGOS));
      catalogos = cat.exists() ? (cat.data().listas || {}) : {};
    } catch { catalogos = {}; }
    actualizarFiltros();
    filtrar(false);
    ctx.alCambiar(equipos.length);
  } catch (error) {
    notificar("error", "No se pudo cargar el inventario", error.message);
  }
}

export function limpiarInventario() {
  equipos = [];
  lista = [];
  if (el.tbody) { filtrar(true); ctx.alCambiar(0); }
}

// Vuelve a pintar (p. ej. cuando cambia Personal y hay que actualizar Usuario, Cargo…)
export function refrescarInventario() {
  if (el.tbody) pintar();
}

// Muestra solo los equipos de un DNI (desde la columna "Equipos" de Personal)
export function verEquiposDe(dni) {
  el.fEquipo.value = "";
  el.fEstado.value = "";
  el.buscar.value = textoDni(dni);
  filtrar(true);
}

// ---------- Filtros y tabla ----------

function distintos(k) {
  return [...new Set(equipos.map(e => String(e[k] ?? "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
}

function opcionesSelect(select, valores, todos) {
  const actual = select.value;
  select.innerHTML = `<option value="">${todos}</option>` + valores.map(v => `<option>${escapeHtml(v)}</option>`).join("");
  if (valores.includes(actual)) select.value = actual;
}

function actualizarFiltros() {
  opcionesSelect(el.fEquipo, distintos("Equipo"), "Todos los equipos");
  opcionesSelect(el.fEstado, distintos("Estado"), "Todos los estados");

  const conteo = new Map();
  equipos.forEach(e => { const t = String(e.Equipo || "Sin tipo").trim(); conteo.set(t, (conteo.get(t) || 0) + 1); });
  const chips = [...conteo.entries()].sort((a, b) => b[1] - a[1])
    .map(([t, n]) => `<button type="button" class="chip" data-equipo="${escapeHtml(t === "Sin tipo" ? "" : t)}"><span>${escapeHtml(t)}</span><b>${n}</b></button>`).join("");
  el.resumen.innerHTML = equipos.length
    ? `<button type="button" class="chip chip-total" data-equipo=""><span>Total</span><b>${equipos.length}</b></button>${chips}`
    : "";
}

function filtrar(reiniciar) {
  const q = el.buscar.value.trim().toLowerCase();
  const qDni = claveDni(q);
  const fe = el.fEquipo.value;
  const fs = el.fEstado.value;
  const personal = mapaPersonal();
  lista = equipos.filter(e => {
    if (fe && String(e.Equipo ?? "").trim() !== fe) return false;
    if (fs && String(e.Estado ?? "").trim() !== fs) return false;
    if (!q) return true;
    if (qDni && /^\d+$/.test(q) && claveDni(e.DNI) === qDni) return true;
    return CAMPOS.some(c => {
      if (c.secreto) return false;
      const v = valor(e, c, personal);
      return typeof v !== "object" && String(v).toLowerCase().includes(q);
    });
  });
  if (reiniciar) pagina = 1;
  pintar();
}

function celda(e, c, personal) {
  const v = valor(e, c, personal);
  if (typeof v === "object") return `<td class="na" title="El DNI no está en Personal">#N/A</td>`;
  if (c.secreto) return `<td>${v ? "••••••" : ""}</td>`;
  if (c.k === "Estado" && v) {
    const tono = /inoperativo|obsoleto/i.test(v) ? "rojo" : /revisar|custodia/i.test(v) ? "ambar" : "verde";
    return `<td><span class="estado estado-${tono}">${escapeHtml(v)}</span></td>`;
  }
  return `<td${c.largo ? ' class="largo"' : ""}>${escapeHtml(v)}</td>`;
}

function pintar() {
  const total = lista.length;
  const paginas = Math.max(1, Math.ceil(total / POR_PAGINA));
  pagina = Math.min(Math.max(1, pagina), paginas);
  const desde = (pagina - 1) * POR_PAGINA;
  const personal = mapaPersonal();

  el.tbody.innerHTML = total ? "" : `<tr><td colspan="${CAMPOS.length + 1}" class="vacio">No hay equipos para mostrar.</td></tr>`;
  lista.slice(desde, desde + POR_PAGINA).forEach(e => {
    const tr = document.createElement("tr");
    const nombre = `${e.Equipo || "equipo"} ${e["S/N"] || e.Item || ""}`.trim();
    tr.innerHTML = `<td class="acciones"><button type="button" class="accion editar" title="Editar" aria-label="Editar ${escapeHtml(nombre)}">${ICONO_EDITAR}</button><button type="button" class="accion eliminar" title="Eliminar" aria-label="Eliminar ${escapeHtml(nombre)}">✕</button></td>`
      + CAMPOS.map(c => celda(e, c, personal)).join("");
    tr.querySelector(".editar").addEventListener("click", () => abrirFormulario(e));
    tr.querySelector(".eliminar").addEventListener("click", () => eliminar(e, nombre));
    el.tbody.appendChild(tr);
  });
  el.wrap.scrollTop = 0;
  el.paginacion.innerHTML = htmlPaginacion(total, pagina, paginas, POR_PAGINA);
}

// ---------- Alta y edición ----------

function abrirFormulario(e) {
  if (!ctx.requiereConfig()) return;
  const nuevo = !e;
  const personal = mapaPersonal();
  const datos = e ? { ...e } : { Item: Math.max(0, ...equipos.map(x => Number(x.Item) || 0)) + 1 };

  // Sugerencias con los valores ya usados en cada columna
  const listas = GUARDADOS.filter(c => !c.fecha && !c.numero && !c.secreto && !c.largo && c.k !== "DNI")
    .map((c, i) => ({ c, id: `dl-inv-${i}`, valores: [...new Set([...(catalogos[c.k] || []), ...distintos(c.k)])] }));
  const idLista = new Map(listas.map(l => [l.c.k, l.id]));
  const dlPersonal = [...personal.values()].map(p => `<option value="${escapeHtml(textoDni(p.DNI))}">${escapeHtml(p.Nombre)}</option>`).join("");

  const grupos = [...new Set(CAMPOS.map(c => c.grupo))];
  const html = grupos.map(g => {
    const campos = CAMPOS.filter(c => c.grupo === g).map(c => {
      const i = CAMPOS.indexOf(c);
      const id = `inv-${i}`;
      const label = `<label for="${id}">${escapeHtml(c.k)}${c.requerido ? ' <span class="req">*</span>' : ""}</label>`;
      if (c.de) return `<div class="campo"><label>${escapeHtml(c.k)}</label><div class="derivado" data-de="${escapeHtml(c.de)}"></div></div>`;
      const v = escapeHtml(datos[c.k]);
      if (c.largo) return `<div class="campo ancho">${label}<textarea id="${id}" name="c${i}" rows="2">${v}</textarea></div>`;
      if (c.k === "DNI") return `<div class="campo">${label}<input id="${id}" name="c${i}" value="${v}" list="dl-inv-personal" autocomplete="off" placeholder="Buscar DNI o nombre"></div>`;
      const tipo = c.fecha && (!datos[c.k] || /^\d{4}-\d{2}-\d{2}$/.test(datos[c.k])) ? "date"
        : c.numero ? "number" : c.secreto ? "password" : "text";
      const lista = idLista.get(c.k);
      return `<div class="campo">${label}<input id="${id}" name="c${i}" type="${tipo}" value="${v}"${lista ? ` list="${lista}"` : ""} autocomplete="off"></div>`;
    }).join("");
    return `<fieldset class="grupo"><legend>${g}</legend><div class="form-grid tres">${campos}</div></fieldset>`;
  }).join("");

  const fondo = document.createElement("div");
  fondo.className = "modal-fondo";
  fondo.innerHTML = `
    <form class="modal modal-form modal-grande" role="dialog" aria-modal="true" aria-labelledby="invTitulo" novalidate>
      <div class="modal-cabecera">
        <span class="modal-icon">${ICONO_EDITAR}</span>
        <div><h3 id="invTitulo">${nuevo ? "Nuevo equipo" : "Editar equipo"}</h3><p></p></div>
      </div>
      ${html}
      <datalist id="dl-inv-personal">${dlPersonal}</datalist>
      ${listas.map(l => `<datalist id="${l.id}">${l.valores.map(v => `<option value="${escapeHtml(v)}">`).join("")}</datalist>`).join("")}
      <div class="modal-acciones">
        <button type="button" class="btn-outline" data-cancelar>Cancelar</button>
        <button type="submit" class="btn-confirmar">${nuevo ? "Agregar equipo" : "Guardar cambios"}</button>
      </div>
    </form>`;
  fondo.querySelector(".modal-cabecera p").textContent = nuevo ? "Completa los datos del equipo" : `${e.Equipo || ""} ${e.Marca || ""} ${e.Modelo || ""}`.trim();
  const form = fondo.querySelector("form");
  const inputDni = form.elements[`c${CAMPOS.findIndex(c => c.k === "DNI")}`];

  // Vista previa en vivo de los datos de Personal (como el VLOOKUP)
  const vistaPrevia = () => {
    const k = claveDni(inputDni.value);
    const p = k ? personal.get(k) : null;
    form.querySelectorAll(".derivado").forEach(d => {
      d.textContent = !k ? "—" : p ? (p[d.dataset.de] || "—") : "#N/A";
      d.classList.toggle("na", Boolean(k && !p));
    });
  };
  inputDni.addEventListener("input", vistaPrevia);
  vistaPrevia();

  const cerrar = () => { document.removeEventListener("keydown", teclas); fondo.remove(); };
  const teclas = (ev) => { if (ev.key === "Escape") cerrar(); };
  document.addEventListener("keydown", teclas);
  fondo.addEventListener("mousedown", (ev) => { if (ev.target === fondo) cerrar(); });
  fondo.querySelector("[data-cancelar]").addEventListener("click", cerrar);

  form.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const nuevosDatos = {};
    GUARDADOS.forEach(c => {
      let v = form.elements[`c${CAMPOS.indexOf(c)}`].value.trim();
      if (c.numero && v !== "") v = Number(v);
      nuevosDatos[c.k] = v;
    });
    const faltan = CAMPOS.filter(c => c.requerido && !nuevosDatos[c.k]).map(c => c.k);
    if (faltan.length) { notificar("warning", "Faltan datos obligatorios", faltan.join(", ")); return; }

    const sn = String(nuevosDatos["S/N"]).toUpperCase();
    const repetido = sn && equipos.find(x => x.id !== e?.id && String(x["S/N"] ?? "").toUpperCase() === sn);
    if (repetido) {
      const seguir = await confirmar({
        titulo: `Ya existe un equipo con S/N ${nuevosDatos["S/N"]}`,
        mensaje: `${repetido.Equipo || ""} ${repetido.Marca || ""} ${repetido.Modelo || ""} (Item ${repetido.Item ?? "—"}). ¿Guardar de todos modos?`,
        aceptar: "Guardar igual"
      });
      if (!seguir) return;
    }

    const btn = form.querySelector('[type="submit"]');
    btn.disabled = true;
    btn.textContent = "Guardando…";
    try {
      if (nuevo) await addDoc(col, nuevosDatos);
      else await updateDoc(doc(db, COLECCION, e.id), nuevosDatos);
      cerrar();
      notificar("success", nuevo ? "Equipo agregado" : "Cambios guardados", `${nuevosDatos.Equipo} ${nuevosDatos["S/N"] || ""}`.trim());
      cargarInventario();
    } catch (error) {
      notificar("error", "No se pudo guardar", error.message);
      btn.disabled = false;
      btn.textContent = nuevo ? "Agregar equipo" : "Guardar cambios";
    }
  });

  document.body.appendChild(fondo);
  form.elements[`c${CAMPOS.findIndex(c => c.k === (nuevo ? "Equipo" : "Item"))}`].focus();
}

async function eliminar(e, nombre) {
  if (!ctx.requiereConfig()) return;
  const ok = await confirmar({ titulo: "¿Eliminar este equipo?", mensaje: nombre, aceptar: "Eliminar", peligro: true });
  if (!ok) return;
  try {
    await deleteDoc(doc(db, COLECCION, e.id));
    notificar("success", "Equipo eliminado", nombre);
    cargarInventario();
  } catch (error) {
    notificar("error", "No se pudo eliminar", error.message);
  }
}

async function borrarTodo() {
  if (!ctx.requiereConfig()) return;
  if (!equipos.length) { notificar("info", "No hay equipos que borrar"); return; }
  const ok = await confirmar({
    titulo: `¿Eliminar los ${equipos.length} equipos?`,
    mensaje: "Esta acción no se puede deshacer.",
    aceptar: "Eliminar todo",
    peligro: true
  });
  if (!ok) return;
  try {
    for (let i = 0; i < equipos.length; i += 100) {
      const batch = writeBatch(db);
      equipos.slice(i, i + 100).forEach(x => batch.delete(doc(db, COLECCION, x.id)));
      await batch.commit();
    }
    notificar("success", "Inventario vaciado", `${equipos.length} equipos eliminados.`);
    cargarInventario();
  } catch (error) {
    notificar("error", "No se pudo vaciar el inventario", error.message);
  }
}

// ---------- Importación ----------

// Identifica el mismo equipo entre importaciones: Item del Excel; si no hay, S/N, IMEI o MAC
function claveEquipo(r) {
  const n = (v) => String(v ?? "").replace(/[\s:-]/g, "").toUpperCase();
  if (r.Item !== "" && r.Item !== undefined && r.Item !== null) return `item:${Number(r.Item) || n(r.Item)}`;
  if (n(r["S/N"])) return `sn:${n(r["S/N"])}`;
  if (n(r.IMEI1)) return `imei:${n(r.IMEI1)}`;
  if (n(r["Mac Address"])) return `mac:${n(r["Mac Address"])}`;
  return "";
}

function registroDesdeFila(r, fechasMal) {
  const o = {};
  GUARDADOS.forEach(c => {
    let v = r[c.k];
    if (c.fecha) {
      const f = normalizarFecha(v);
      if (f === null) { fechasMal.n++; v = String(v).trim(); } else v = f;
    } else if (c.k === "DNI") {
      v = v === "" || v === undefined ? "" : textoDni(v);
    } else if (c.numero) {
      v = v === "" || v === undefined ? "" : (Number(v) || String(v).trim());
    } else {
      v = String(v ?? "").trim();
    }
    if (v === "#N/A") v = "";
    o[c.k] = v;
  });
  return o;
}

// Lee la hoja del inventario y devuelve los equipos listos para guardar (sin tocar la base de datos)
export function leerHojaInventario(wb) {
  const nombreHoja = wb.SheetNames.find(n => n.trim().toLowerCase() === HOJA.toLowerCase())
    || wb.SheetNames.find(n => /inventario/i.test(n));
  if (!nombreHoja) {
    return { error: ["Hoja no encontrada", `El archivo debe tener una hoja llamada "${HOJA}". Hojas encontradas: ${wb.SheetNames.join(", ")}.`] };
  }
  const ws = wb.Sheets[nombreHoja];

  // La cabecera no está en la primera fila: se busca la fila que tiene "Equipo" y "S/N"
  const matriz = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const filaCab = matriz.findIndex(r => r.some(c => String(c).trim() === "Equipo") && r.some(c => String(c).trim() === "S/N"));
  if (filaCab < 0) {
    return { error: ["Cabecera no encontrada", `No se encontró la fila de títulos (Equipo, S/N…) en la hoja "${nombreHoja}".`] };
  }
  // "range" es una fila absoluta; la hoja puede no empezar en la fila 1 (p. ej. A2:AR705)
  const inicio = XLSX.utils.decode_range(ws["!ref"]).s.r;
  const filas = XLSX.utils.sheet_to_json(ws, { range: inicio + filaCab, defval: "" }).map(limpiarClaves);

  // Solo filas con datos del equipo (las demás son la plantilla con fórmulas)
  const conDatos = filas.filter(r => ["Equipo", "Marca", "Modelo", "S/N", "IMEI1", "Mac Address"].some(k => String(r[k] ?? "").trim()));
  if (!conDatos.length) return { error: ["Sin equipos", `La hoja "${nombreHoja}" no tiene equipos con datos.`] };

  const fechasMal = { n: 0 };
  const porClave = new Map();
  conDatos.forEach(r => { const reg = registroDesdeFila(r, fechasMal); porClave.set(claveEquipo(reg) || Symbol(), reg); });
  return { nombreHoja, registros: [...porClave.values()], fechasMal: fechasMal.n };
}

// Lee la hoja "Datos": cada columna es una lista de valores permitidos (catálogo)
export function leerHojaDatos(wb) {
  const nombre = wb.SheetNames.find(n => n.trim().toLowerCase() === HOJA_DATOS.toLowerCase());
  if (!nombre) return null;
  const ws = wb.Sheets[nombre];
  const matriz = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const filaCab = matriz.findIndex(r => r.some(c => String(c).trim() === "Equipo") && r.some(c => String(c).trim() === "Marca"));
  if (filaCab < 0) return null;
  const cab = matriz[filaCab].map(c => String(c).trim());
  const listas = {};
  cab.forEach((h, i) => {
    if (!h) return;
    const valores = [...new Set(matriz.slice(filaCab + 1).map(r => String(r[i] ?? "").trim()).filter(Boolean))];
    if (valores.length) listas[h] = valores;
  });
  return Object.keys(listas).length ? listas : null;
}

async function importar(file) {
  if (!ctx.requiereConfig()) return;
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const leido = leerHojaInventario(wb);
    if (leido.error) { notificar("error", ...leido.error); return; }
    const { registros } = leido;
    const listasDatos = leerHojaDatos(wb);

    await cargarInventario();
    const existentes = new Map();
    equipos.forEach(x => { const k = claveEquipo(x); if (k && !existentes.has(k)) existentes.set(k, x.id); });
    const aActualizar = registros.filter(r => existentes.has(claveEquipo(r))).length;
    const nuevos = registros.length - aActualizar;

    const personal = mapaPersonal();
    const sinPersona = registros.filter(r => claveDni(r.DNI) && !personal.has(claveDni(r.DNI))).length;

    const ok = await confirmar({
      titulo: `¿Importar ${registros.length} equipos?`,
      mensaje: `${nuevos} nuevos y ${aActualizar} ya existentes (mismo Item) que se actualizarán.`
        + (sinPersona ? ` ${sinPersona} tienen un DNI que no está en Personal (se verá #N/A).` : "")
        + (listasDatos ? ` También se actualizan los catálogos de la hoja Datos (${Object.keys(listasDatos).length} listas).` : ""),
      aceptar: "Importar"
    });
    if (!ok) return;

    el.status.hidden = false;
    for (let i = 0; i < registros.length; i += 100) {
      const batch = writeBatch(db);
      const lote = registros.slice(i, i + 100);
      lote.forEach(r => {
        const id = existentes.get(claveEquipo(r));
        batch.set(id ? doc(db, COLECCION, id) : doc(col), r);
      });
      await batch.commit();
      const hechos = i + lote.length;
      el.progreso.style.width = `${Math.round(hechos / registros.length * 100)}%`;
      el.statusText.textContent = `Importando ${hechos}/${registros.length}…`;
    }
    if (listasDatos) await setDoc(doc(db, ...DOC_CATALOGOS), { listas: listasDatos, actualizado: new Date().toISOString() });
    notificar("success", "Inventario importado", `${nuevos} equipos nuevos y ${aActualizar} actualizados${listasDatos ? "; catálogos de Datos actualizados" : ""}.`);
    if (leido.fechasMal) notificar("warning", "Fechas sin reconocer", `${leido.fechasMal} fechas se guardaron tal cual.`);
    setTimeout(() => { el.status.hidden = true; el.progreso.style.width = "0%"; }, 1500);
    cargarInventario();
  } catch (error) {
    notificar("error", "No se pudo importar", error.message);
    el.status.hidden = true;
  } finally {
    el.input.value = "";
  }
}
