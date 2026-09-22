import { db, auth, configOk } from "./firebase-config.js";
import { collection, addDoc, getDocs, deleteDoc, updateDoc, doc, writeBatch } from "firebase/firestore";
import {
  onAuthStateChanged, signInWithEmailAndPassword, signOut, sendPasswordResetEmail,
  setPersistence, browserLocalPersistence, browserSessionPersistence
} from "firebase/auth";
import { notificar, confirmar } from "./notificaciones.js";
import { iniciarCarrusel } from "./carrusel.js";
import {
  escapeHtml, normalizarFecha, textoDni, claveDni, limpiarClaves,
  ICONO_EDITAR, htmlPaginacion, conectarImportador
} from "./utils.js";
import { iniciarInventario, cargarInventario, limpiarInventario, refrescarInventario, contarEquiposPorDni, verEquiposDe } from "./inventario.js";

const COLLECTION = "personal";
const col = configOk ? collection(db, COLLECTION) : null;
let allRecords = [];

// AVISO DE ESTADO
const estado = document.getElementById("estado");
function mostrarAviso(msg) {
  estado.textContent = msg;
  estado.hidden = false;
}
function ocultarAviso() {
  estado.hidden = true;
  estado.textContent = "";
}
function requiereConfig() {
  if (configOk) return true;
  mostrarAviso("Falta la configuración de Firebase (VITE_FIREBASE_API_KEY / VITE_FIREBASE_PROJECT_ID). La app no puede leer ni guardar datos.");
  return false;
}

// IMPORTADOR EXCEL
const dropZone = document.getElementById("seccionPersonal");
const excelInput = document.getElementById("excelInput");
const selectFileBtn = document.getElementById("selectFileBtn");

conectarImportador({ zona: dropZone, boton: selectFileBtn, input: excelInput, alSoltar: procesarExcel });

async function procesarExcel(file) {
  if (!requiereConfig()) return;
  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });
    const nombreHoja = wb.SheetNames.find(n => n.trim().toLowerCase() === "personal");
    if (!nombreHoja) {
      notificar("error", "Hoja no encontrada", `El archivo debe tener una hoja llamada "Personal". Hojas encontradas: ${wb.SheetNames.join(", ")}.`);
      excelInput.value = "";
      return;
    }

    const jsonData = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja], { defval: "" }).map(limpiarClaves);
    if (jsonData.length === 0) { notificar("warning", "Hoja vacía", "La hoja \"Personal\" no tiene filas."); excelInput.value = ""; return; }

    // El Código es opcional: hay filas de áreas o puestos compartidos (p. ej. "Sistemas") a las que se asignan equipos
    const registrosValidos = jsonData.filter(r => String(r.DNI).trim() && String(r.Nombre).trim());
    const omitidos = jsonData.length - registrosValidos.length;
    if (registrosValidos.length === 0) {
      notificar("warning", "Sin registros válidos", "Ninguna fila tiene DNI y Nombre a la vez.");
      excelInput.value = "";
      return;
    }

    // Una fila por DNI dentro del archivo (gana la última)
    const porDni = new Map();
    registrosValidos.forEach(r => porDni.set(claveDni(r.DNI), r));
    const filas = [...porDni.values()];
    const repetidosEnArchivo = registrosValidos.length - filas.length;

    // Si el DNI ya existe se actualiza ese registro en vez de crear otro
    await cargar();
    const existentes = new Map();
    allRecords.forEach(r => { const k = claveDni(r.DNI); if (k && !existentes.has(k)) existentes.set(k, r.id); });
    const aActualizar = filas.filter(r => existentes.has(claveDni(r.DNI))).length;
    const nuevos = filas.length - aActualizar;

    const detalles = [];
    if (aActualizar) detalles.push(`${aActualizar} ya existen (mismo DNI) y se actualizarán.`);
    if (repetidosEnArchivo) detalles.push(`${repetidosEnArchivo} filas repetidas en el archivo se unifican.`);
    if (omitidos) detalles.push(`${omitidos} filas sin DNI o Nombre se omiten.`);
    const ok = await confirmar({
      titulo: `¿Importar ${filas.length} registros?`,
      mensaje: `${nuevos} nuevos. ${detalles.join(" ")}`,
      aceptar: "Importar"
    });
    if (!ok) { excelInput.value = ""; return; }
    let fechasNoReconocidas = 0;
    
    const statusDiv = document.getElementById("importStatus");
    const statusText = document.getElementById("statusText");
    const progressFill = document.getElementById("progressFill");
    statusDiv.hidden = false;
    
    const batchSize = 20;
    for (let i = 0; i < filas.length; i += batchSize) {
      const batch = writeBatch(db);
      const lote = filas.slice(i, i + batchSize);
      
      lote.forEach(row => {
        const idExistente = existentes.get(claveDni(row.DNI));
        batch.set(idExistente ? doc(db, COLLECTION, idExistente) : doc(col), {
          DNI: textoDni(row.DNI),
          Código: (row.Código || "").toString().trim(),
          Nombre: (row.Nombre || "").toString().trim().toUpperCase(),
          Foto: (row.Foto || "").toString().trim(),
          "Fecha de Ingreso": (() => {
            const f = normalizarFecha(row["Fecha de Ingreso"]);
            if (f === null) { fechasNoReconocidas++; return String(row["Fecha de Ingreso"]).trim(); }
            return f;
          })(),
          "Puesto Real": (row["Puesto Real"] || "").toString().trim(),
          "CC Real": (row["CC Real"] || "").toString().trim(),
          "Unidad Real": (row["Unidad Real"] || "").toString().trim(),
          "Nombre Host": (row["Nombre Host"] || "").toString().trim(),
          "Nueva Tajeta Micros": (row["Nueva Tajeta Micros"] || "").toString().trim()
        });
      });
      
      await batch.commit();
      const pct = Math.round(((i + lote.length) / filas.length) * 100);
      progressFill.style.width = pct + "%";
      statusText.textContent = `Importando ${i + lote.length}/${filas.length}...`;
    }
    
    statusText.textContent = `${filas.length} registros importados`;
    notificar("success", "Importación completada", `${nuevos} nuevos y ${aActualizar} actualizados${omitidos ? `; ${omitidos} filas omitidas` : ""}.`);
    if (fechasNoReconocidas) notificar("warning", "Fechas sin reconocer", `${fechasNoReconocidas} fechas se guardaron tal cual porque no tienen un formato de fecha válido.`);
    setTimeout(() => { statusDiv.hidden = true; progressFill.style.width = "0%"; excelInput.value = ""; cargar(); }, 1500);
  } catch (error) {
    notificar("error", "No se pudo importar", error.message);
    document.getElementById("importStatus").hidden = true;
    excelInput.value = "";
  }
}

// CRUD
async function eliminar(id, nombre) {
  if (!requiereConfig()) return;
  const ok = await confirmar({ titulo: "¿Eliminar este registro?", mensaje: nombre || "", aceptar: "Eliminar", peligro: true });
  if (!ok) return;
  try {
    await deleteDoc(doc(db, COLLECTION, id));
    notificar("success", "Registro eliminado", nombre || "");
    cargar();
  } catch (error) {
    notificar("error", "No se pudo eliminar", error.message);
  }
}

async function cargar() {
  if (!requiereConfig()) return;
  try {
    const snap = await getDocs(col);
    allRecords = [];
    snap.forEach(d => allRecords.push({ id: d.id, ...d.data() }));
    aplicarFiltro(false);
    actualizarBotonDuplicados();
    actualizarContadores();
    refrescarInventario();
    ocultarAviso();
  } catch (error) {
    console.error("Error:", error);
    mostrarAviso(`Error al cargar los datos: ${error.message}`);
  }
}

// EDICIÓN
const CAMPOS = [
  { k: "DNI", label: "DNI", requerido: true },
  { k: "Código", label: "Código" },
  { k: "Nombre", label: "Nombre", requerido: true, ancho: true },
  { k: "Foto", label: "Foto", opciones: ["", "SI", "NO"] },
  { k: "Fecha de Ingreso", label: "Fecha de ingreso", fecha: true },
  { k: "Puesto Real", label: "Puesto" },
  { k: "CC Real", label: "Centro de costo" },
  { k: "Unidad Real", label: "Unidad" },
  { k: "Nombre Host", label: "Host" },
  { k: "Nueva Tajeta Micros", label: "Tarjeta Micros" }
];

// Ventana de alta (p = null) o edición de un registro de Personal
function abrirEdicion(p) {
  if (!requiereConfig()) return;
  const nuevo = !p;
  p = p || {};
  const fondo = document.createElement("div");
  fondo.className = "modal-fondo";
  const campos = CAMPOS.map((c, i) => {
    const v = escapeHtml(p[c.k]);
    const id = `ed-${i}`;
    let control;
    if (c.opciones) {
      const actual = String(p[c.k] ?? "");
      const ops = c.opciones.includes(actual) ? c.opciones : [...c.opciones, actual];
      control = `<select id="${id}" name="f${i}">${ops.map(o => `<option${o === actual ? " selected" : ""}>${escapeHtml(o)}</option>`).join("")}</select>`;
    } else {
      // Si la fecha guardada no es aaaa-mm-dd se edita como texto para no perderla
      const tipo = c.fecha && (!p[c.k] || /^\d{4}-\d{2}-\d{2}$/.test(p[c.k])) ? "date" : "text";
      control = `<input id="${id}" name="f${i}" type="${tipo}" value="${v}"${c.requerido ? " required" : ""}>`;
    }
    return `<div class="campo${c.ancho ? " ancho" : ""}"><label for="${id}">${c.label}${c.requerido ? ' <span class="req">*</span>' : ""}</label>${control}</div>`;
  }).join("");

  fondo.innerHTML = `
    <form class="modal modal-form" role="dialog" aria-modal="true" aria-labelledby="edTitulo" novalidate>
      <div class="modal-cabecera">
        <span class="modal-icon">${ICONO_EDITAR}</span>
        <div><h3 id="edTitulo">${nuevo ? "Nuevo registro" : "Editar registro"}</h3><p></p></div>
      </div>
      <div class="form-grid">${campos}</div>
      <div class="modal-acciones">
        <button type="button" class="btn-outline" data-cancelar>Cancelar</button>
        <button type="submit" class="btn-confirmar">${nuevo ? "Agregar" : "Guardar cambios"}</button>
      </div>
    </form>`;
  fondo.querySelector(".modal-cabecera p").textContent = nuevo ? "Completa los datos de la persona o área" : (p.Nombre || "");
  const form = fondo.querySelector("form");

  const cerrar = () => { document.removeEventListener("keydown", teclas); fondo.remove(); };
  const teclas = (e) => { if (e.key === "Escape") cerrar(); };
  document.addEventListener("keydown", teclas);
  fondo.addEventListener("mousedown", (e) => { if (e.target === fondo) cerrar(); });
  fondo.querySelector("[data-cancelar]").addEventListener("click", cerrar);

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const datos = {};
    CAMPOS.forEach((c, i) => {
      let v = form.elements[`f${i}`].value.trim();
      if (c.k === "Nombre") v = v.toUpperCase();
      datos[c.k] = v;
    });
    const faltan = CAMPOS.filter(c => c.requerido && !datos[c.k]).map(c => c.label);
    if (faltan.length) {
      notificar("warning", "Faltan datos obligatorios", faltan.join(", "));
      return;
    }
    // Evitar duplicados: el DNI no puede repetirse en otro registro
    const existente = allRecords.find(r => r.id !== p.id && claveDni(r.DNI) === claveDni(datos.DNI));
    if (existente) {
      const editar = await confirmar({
        titulo: `Ya existe un registro con DNI ${datos.DNI}`,
        mensaje: `${existente.Nombre || ""}. Para evitar duplicados, edita el registro existente.`,
        aceptar: "Editar existente",
        cancelar: "Volver"
      });
      if (editar) { cerrar(); abrirEdicion(existente); }
      return;
    }
    const btn = form.querySelector('[type="submit"]');
    btn.disabled = true;
    btn.textContent = "Guardando…";
    try {
      if (nuevo) await addDoc(col, datos);
      else await updateDoc(doc(db, COLLECTION, p.id), datos);
      cerrar();
      notificar("success", nuevo ? "Registro agregado" : "Cambios guardados", datos.Nombre);
      cargar();
    } catch (error) {
      notificar("error", nuevo ? "No se pudo agregar" : "No se pudieron guardar los cambios", error.message);
      btn.disabled = false;
      btn.textContent = nuevo ? "Agregar" : "Guardar cambios";
    }
  });

  document.body.appendChild(fondo);
  form.elements.f0.focus();
}

// TABLA PAGINADA
const POR_PAGINA = 50;
const tablaWrap = document.getElementById("tablaWrap");
const paginacion = document.getElementById("paginacion");
let listaActual = [];
let pagina = 1;

function mostrarTabla(registros, reiniciar = true) {
  listaActual = registros;
  if (reiniciar) pagina = 1;
  pintarPagina();
}

function pintarPagina() {
  const tbody = document.getElementById("tbody");
  const total = listaActual.length;
  const paginas = Math.max(1, Math.ceil(total / POR_PAGINA));
  pagina = Math.min(Math.max(1, pagina), paginas);
  const desde = (pagina - 1) * POR_PAGINA;

  tbody.innerHTML = "";
  if (total === 0) {
    tbody.innerHTML = `<tr><td colspan="12" class="vacio">No hay registros para mostrar.</td></tr>`;
  }
  const equipos = contarEquiposPorDni();
  listaActual.slice(desde, desde + POR_PAGINA).forEach(p => {
    const tr = document.createElement("tr");
    const n = equipos.get(claveDni(p.DNI)) || 0;
    tr.innerHTML = `<td>${escapeHtml(p.DNI)}</td><td>${escapeHtml(p.Código)}</td><td>${escapeHtml(p.Nombre)}</td><td>${escapeHtml(p.Foto)}</td><td class="nowrap">${escapeHtml(p["Fecha de Ingreso"])}</td><td>${escapeHtml(p["Puesto Real"])}</td><td>${escapeHtml(p["CC Real"])}</td><td>${escapeHtml(p["Unidad Real"])}</td><td>${escapeHtml(p["Nombre Host"])}</td><td>${escapeHtml(p["Nueva Tajeta Micros"])}</td><td class="centro">${n ? `<button type="button" class="chip-equipos" title="Ver equipos asignados">${n}</button>` : '<span class="sin-dato">0</span>'}</td><td class="acciones"><button type="button" class="accion editar" title="Editar" aria-label="Editar ${escapeHtml(p.Nombre)}">${ICONO_EDITAR}</button><button type="button" class="accion eliminar" title="Eliminar" aria-label="Eliminar ${escapeHtml(p.Nombre)}">✕</button></td>`;
    tr.querySelector(".editar").addEventListener("click", () => abrirEdicion(p));
    tr.querySelector(".eliminar").addEventListener("click", () => eliminar(p.id, p.Nombre));
    tr.querySelector(".chip-equipos")?.addEventListener("click", () => { mostrarPestana("inventario"); verEquiposDe(p.DNI); });
    tbody.appendChild(tr);
  });
  tablaWrap.scrollTop = 0;
  paginacion.innerHTML = htmlPaginacion(total, pagina, paginas, POR_PAGINA);
}

paginacion.addEventListener("click", (e) => {
  const b = e.target.closest("[data-pag]");
  if (!b || b.disabled) return;
  pagina = Number(b.dataset.pag);
  pintarPagina();
});

// BÚSQUEDA
const searchInput = document.getElementById("searchInput");
function aplicarFiltro(reiniciar = true) {
  const q = searchInput.value.trim().toLowerCase();
  const filtrados = q
    ? allRecords.filter(p => (p.Nombre || "").toLowerCase().includes(q) || String(p.DNI || "").toLowerCase().includes(q))
    : allRecords;
  mostrarTabla(filtrados, reiniciar);
}
searchInput.addEventListener("input", () => aplicarFiltro(true));

// DUPLICADOS
const dupBtn = document.getElementById("dupBtn");

// Puntúa qué copia conservar: más campos con datos y sin la fecha errónea 1970-01-01
function puntuar(r) {
  let s = CAMPOS.filter(c => String(r[c.k] ?? "").trim()).length;
  if (r["Fecha de Ingreso"] === "1970-01-01") s -= 5;
  return s;
}

function sobrantesDuplicados() {
  const grupos = new Map();
  allRecords.forEach(r => {
    const k = claveDni(r.DNI);
    if (!k) return;
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(r);
  });
  const sobran = [];
  grupos.forEach(g => {
    if (g.length < 2) return;
    g.sort((a, b) => puntuar(b) - puntuar(a));
    sobran.push(...g.slice(1));
  });
  return sobran;
}

function actualizarBotonDuplicados() {
  const n = sobrantesDuplicados().length;
  dupBtn.hidden = n === 0;
  dupBtn.textContent = `Quitar duplicados (${n})`;
}

dupBtn.addEventListener("click", async () => {
  if (!requiereConfig()) return;
  const sobran = sobrantesDuplicados();
  if (!sobran.length) { notificar("info", "No hay duplicados"); return; }
  const dnis = new Set(sobran.map(r => claveDni(r.DNI))).size;
  const ok = await confirmar({
    titulo: `¿Quitar ${sobran.length} registros duplicados?`,
    mensaje: `Hay ${dnis} ${dnis === 1 ? "DNI repetido" : "DNI repetidos"}. Se conserva una copia de cada uno (la más completa, evitando fechas 1970-01-01) y se eliminan las demás.`,
    aceptar: "Quitar duplicados",
    peligro: true
  });
  if (!ok) return;
  try {
    for (let i = 0; i < sobran.length; i += 20) {
      const batch = writeBatch(db);
      sobran.slice(i, i + 20).forEach(r => batch.delete(doc(db, COLLECTION, r.id)));
      await batch.commit();
    }
    notificar("success", "Duplicados eliminados", `${sobran.length} registros eliminados; queda uno por DNI.`);
    cargar();
  } catch (error) {
    notificar("error", "No se pudieron quitar los duplicados", error.message);
  }
});

// LIMPIAR TODO
document.getElementById("clearAllBtn").addEventListener("click", async () => {
  if (!requiereConfig()) return;
  if (allRecords.length === 0) { notificar("info", "No hay registros que borrar"); return; }
  const ok = await confirmar({
    titulo: `¿Eliminar los ${allRecords.length} registros?`,
    mensaje: "Esta acción no se puede deshacer.",
    aceptar: "Eliminar todo",
    peligro: true
  });
  if (!ok) return;
  
  try {
    const batchSize = 20;
    for (let i = 0; i < allRecords.length; i += batchSize) {
      const batch = writeBatch(db);
      allRecords.slice(i, i + batchSize).forEach(r => batch.delete(doc(db, COLLECTION, r.id)));
      await batch.commit();
    }
    notificar("success", "Registros eliminados", `${allRecords.length} registros eliminados.`);
    cargar();
  } catch (error) {
    notificar("error", "No se pudieron eliminar", error.message);
  }
});

// PESTAÑAS
const pestanas = document.querySelectorAll(".pestana");
function mostrarPestana(nombre) {
  pestanas.forEach(b => {
    const activa = b.dataset.vista === nombre;
    b.classList.toggle("activa", activa);
    b.setAttribute("aria-selected", activa ? "true" : "false");
    document.getElementById(b.getAttribute("aria-controls")).hidden = !activa;
  });
  document.getElementById("subtitulo").textContent = nombre === "inventario" ? "Inventario de equipos TI" : "Gestión de Personal";
  try { localStorage.setItem("pestana", nombre); } catch {}
}
pestanas.forEach(b => b.addEventListener("click", () => mostrarPestana(b.dataset.vista)));
try { mostrarPestana(localStorage.getItem("pestana") === "inventario" ? "inventario" : "personal"); } catch { mostrarPestana("personal"); }

function actualizarContadores(totalInventario) {
  document.getElementById("cuentaPersonal").textContent = allRecords.length;
  if (totalInventario !== undefined) document.getElementById("cuentaInventario").textContent = totalInventario;
}

iniciarInventario({
  requiereConfig,
  obtenerPersonal: () => allRecords,
  alCambiar: (total) => { actualizarContadores(total); pintarPagina(); }
});

// PIE DE PÁGINA
document.querySelectorAll(".anio").forEach(e => { e.textContent = new Date().getFullYear(); });

// AUTENTICACIÓN
iniciarCarrusel(document.getElementById("carrusel"));

const splash = document.getElementById("splash");
const loginView = document.getElementById("login");
const appView = document.getElementById("app");
const loginForm = document.getElementById("loginForm");
const loginEmail = document.getElementById("loginEmail");
const loginPassword = document.getElementById("loginPassword");
const loginBtn = document.getElementById("loginBtn");
const loginAlert = document.getElementById("loginAlert");
const rememberMe = document.getElementById("rememberMe");
const togglePassword = document.getElementById("togglePassword");
const capsWarning = document.getElementById("capsWarning");

const MENSAJES_AUTH = {
  "auth/invalid-credential": "Correo o contraseña incorrectos.",
  "auth/invalid-login-credentials": "Correo o contraseña incorrectos.",
  "auth/wrong-password": "Correo o contraseña incorrectos.",
  "auth/user-not-found": "Correo o contraseña incorrectos.",
  "auth/invalid-email": "El correo electrónico no es válido.",
  "auth/user-disabled": "Esta cuenta está deshabilitada. Contacta con el administrador.",
  "auth/too-many-requests": "Demasiados intentos. Espera unos minutos e inténtalo de nuevo.",
  "auth/network-request-failed": "Sin conexión. Revisa tu red e inténtalo de nuevo.",
  "auth/operation-not-allowed": "El acceso con correo y contraseña no está habilitado en Firebase."
};

function mostrarLoginAlerta(msg, tipo = "error") {
  loginAlert.textContent = msg;
  loginAlert.className = `login-alert ${tipo}`;
  loginAlert.hidden = false;
}

function setCargando(cargando) {
  loginBtn.disabled = cargando;
  loginBtn.classList.toggle("loading", cargando);
}

function mostrarVista(user) {
  splash.hidden = true;
  loginView.hidden = Boolean(user);
  appView.hidden = !user;
}

if (!configOk) {
  mostrarVista(null);
  mostrarLoginAlerta("Falta la configuración de Firebase (VITE_FIREBASE_API_KEY / VITE_FIREBASE_PROJECT_ID). No es posible iniciar sesión.");
  loginBtn.disabled = true;
} else {
  onAuthStateChanged(auth, (user) => {
    mostrarVista(user);
    if (user) {
      document.getElementById("userEmail").textContent = user.email || "";
      loginForm.reset();
      rememberMe.checked = true;
      loginAlert.hidden = true;
      cargar();
      cargarInventario();
    } else {
      allRecords = [];
      mostrarTabla([]);
      limpiarInventario();
      ocultarAviso();
      loginEmail.focus();
    }
  });
}

loginForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const email = loginEmail.value.trim();
  const password = loginPassword.value;
  if (!email || !password) {
    mostrarLoginAlerta("Introduce tu correo y tu contraseña.");
    return;
  }
  loginAlert.hidden = true;
  setCargando(true);
  try {
    await setPersistence(auth, rememberMe.checked ? browserLocalPersistence : browserSessionPersistence);
    await signInWithEmailAndPassword(auth, email, password);
  } catch (error) {
    mostrarLoginAlerta(MENSAJES_AUTH[error.code] || `No se pudo iniciar sesión: ${error.message}`);
    loginPassword.select();
  } finally {
    setCargando(false);
  }
});

document.getElementById("forgotBtn").addEventListener("click", async () => {
  const email = loginEmail.value.trim();
  if (!email) {
    mostrarLoginAlerta("Escribe tu correo arriba y vuelve a pulsar «¿Olvidaste tu contraseña?».", "info");
    loginEmail.focus();
    return;
  }
  try {
    await sendPasswordResetEmail(auth, email);
    mostrarLoginAlerta(`Si ${email} tiene una cuenta, recibirás un correo para restablecer la contraseña.`, "success");
  } catch (error) {
    mostrarLoginAlerta(MENSAJES_AUTH[error.code] || `No se pudo enviar el correo: ${error.message}`);
  }
});

togglePassword.addEventListener("click", () => {
  const visible = loginPassword.type === "text";
  loginPassword.type = visible ? "password" : "text";
  togglePassword.textContent = visible ? "Mostrar" : "Ocultar";
  togglePassword.setAttribute("aria-label", visible ? "Mostrar contraseña" : "Ocultar contraseña");
});

["keydown", "keyup"].forEach(evt => loginPassword.addEventListener(evt, (e) => {
  if (e.getModifierState) capsWarning.hidden = !e.getModifierState("CapsLock");
}));

document.getElementById("logoutBtn").addEventListener("click", () => signOut(auth));

// INICIALIZAR
document.getElementById("personalNuevo").addEventListener("click", () => abrirEdicion(null));
