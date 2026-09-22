import { db, auth, configOk } from "./firebase-config.js";
import { collection, addDoc, getDocs, deleteDoc, doc, writeBatch } from "firebase/firestore";
import {
  onAuthStateChanged, signInWithEmailAndPassword, signOut, sendPasswordResetEmail,
  setPersistence, browserLocalPersistence, browserSessionPersistence
} from "firebase/auth";
import { notificar, confirmar } from "./notificaciones.js";

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

function escapeHtml(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// IMPORTADOR EXCEL
const dropZone = document.getElementById("dropZone");
const excelInput = document.getElementById("excelInput");
const selectFileBtn = document.getElementById("selectFileBtn");

selectFileBtn.addEventListener("click", () => excelInput.click());
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragover"); });
dropZone.addEventListener("dragleave", () => { dropZone.classList.remove("dragover"); });
dropZone.addEventListener("drop", (e) => { e.preventDefault(); dropZone.classList.remove("dragover"); procesarExcel(e.dataTransfer.files[0]); });
excelInput.addEventListener("change", (e) => { if(e.target.files[0]) procesarExcel(e.target.files[0]); });

const pad2 = (n) => String(n).padStart(2, "0");

// Convierte la fecha del Excel (número de serie, texto dd/mm/aaaa o aaaa-mm-dd) a "aaaa-mm-dd".
// Devuelve null si no se reconoce, para no detener la importación.
function normalizarFecha(v) {
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

// Quita espacios de los nombres de columna (p. ej. "DNI " -> "DNI")
function limpiarClaves(row) {
  const r = {};
  for (const [k, v] of Object.entries(row)) r[String(k).trim()] = v;
  return r;
}

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

    const registrosValidos = jsonData.filter(r => r.DNI && r.Código && r.Nombre);
    const omitidos = jsonData.length - registrosValidos.length;
    if (registrosValidos.length === 0) {
      notificar("warning", "Sin registros válidos", "Ninguna fila tiene DNI, Código y Nombre a la vez.");
      excelInput.value = "";
      return;
    }

    const ok = await confirmar({
      titulo: `¿Importar ${registrosValidos.length} registros?`,
      mensaje: omitidos ? `Se omitirán ${omitidos} filas sin DNI, Código o Nombre.` : `Desde la hoja "${nombreHoja}" de ${file.name}.`,
      aceptar: "Importar"
    });
    if (!ok) { excelInput.value = ""; return; }
    let fechasNoReconocidas = 0;
    
    const statusDiv = document.getElementById("importStatus");
    const statusText = document.getElementById("statusText");
    const progressFill = document.getElementById("progressFill");
    statusDiv.style.display = "block";
    
    const batchSize = 20;
    for (let i = 0; i < registrosValidos.length; i += batchSize) {
      const batch = writeBatch(db);
      const lote = registrosValidos.slice(i, i + batchSize);
      
      lote.forEach(row => {
        batch.set(doc(col), {
          DNI: (row.DNI || "").toString().trim(),
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
      const pct = Math.round(((i + lote.length) / registrosValidos.length) * 100);
      progressFill.style.width = pct + "%";
      statusText.textContent = `Importando ${i + lote.length}/${registrosValidos.length}...`;
    }
    
    statusText.textContent = `${registrosValidos.length} registros importados`;
    notificar("success", "Importación completada", `${registrosValidos.length} registros importados${omitidos ? `, ${omitidos} filas omitidas` : ""}.`);
    if (fechasNoReconocidas) notificar("warning", "Fechas sin reconocer", `${fechasNoReconocidas} fechas se guardaron tal cual porque no tienen un formato de fecha válido.`);
    setTimeout(() => { statusDiv.style.display = "none"; excelInput.value = ""; cargar(); }, 1500);
  } catch (error) {
    notificar("error", "No se pudo importar", error.message);
    document.getElementById("importStatus").style.display = "none";
    excelInput.value = "";
  }
}

// CRUD MANUAL
async function guardar(e) {
  e.preventDefault();
  if (!requiereConfig()) return;
  const f = e.target;
  const data = {
    DNI: f.DNI.value.trim(),
    Código: f.Código.value.trim(),
    Nombre: f.Nombre.value.trim().toUpperCase(),
    Foto: f.Foto.value,
    "Fecha de Ingreso": f["Fecha de Ingreso"].value,
    "Puesto Real": f["Puesto Real"].value.trim(),
    "CC Real": f["CC Real"].value.trim(),
    "Unidad Real": f["Unidad Real"].value.trim(),
    "Nombre Host": f["Nombre Host"].value.trim(),
    "Nueva Tajeta Micros": f["Nueva Tajeta Micros"].value.trim()
  };
  try {
    await addDoc(col, data);
    f.reset();
    notificar("success", "Registro guardado", data.Nombre);
    cargar();
  } catch (error) {
    notificar("error", "No se pudo guardar", error.message);
  }
}

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
    mostrarTabla(allRecords);
    ocultarAviso();
  } catch (error) {
    console.error("Error:", error);
    mostrarAviso(`Error al cargar los datos: ${error.message}`);
  }
}

function mostrarTabla(registros) {
  const tbody = document.getElementById("tbody");
  tbody.innerHTML = "";
  
  registros.forEach(p => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(p.DNI)}</td><td>${escapeHtml(p.Código)}</td><td>${escapeHtml(p.Nombre)}</td><td>${escapeHtml(p.Foto)}</td><td>${escapeHtml(p["Fecha de Ingreso"])}</td><td>${escapeHtml(p["Puesto Real"])}</td><td>${escapeHtml(p["CC Real"])}</td><td>${escapeHtml(p["Unidad Real"])}</td><td>${escapeHtml(p["Nombre Host"])}</td><td>${escapeHtml(p["Nueva Tajeta Micros"])}</td><td class="del">✕</td>`;
    tr.querySelector(".del").addEventListener("click", () => eliminar(p.id, p.Nombre));
    tbody.appendChild(tr);
  });
}

// BÚSQUEDA
document.getElementById("searchInput").addEventListener("input", (e) => {
  const q = e.target.value.toLowerCase();
  const filtrados = allRecords.filter(p => (p.Nombre || "").toLowerCase().includes(q) || String(p.DNI || "").toLowerCase().includes(q));
  mostrarTabla(filtrados);
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

// AUTENTICACIÓN
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
    } else {
      allRecords = [];
      mostrarTabla([]);
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
document.getElementById("form-personal").addEventListener("submit", guardar);
