import { db } from "./firebase-config.js";
import { collection, addDoc, getDocs, deleteDoc, doc, writeBatch } from "firebase/firestore";

const COLLECTION = "personal";
const col = collection(db, COLLECTION);
let allRecords = [];

// IMPORTADOR EXCEL
const dropZone = document.getElementById("dropZone");
const excelInput = document.getElementById("excelInput");
const selectFileBtn = document.getElementById("selectFileBtn");

selectFileBtn.addEventListener("click", () => excelInput.click());
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.style.background = "#e3f2fd"; });
dropZone.addEventListener("dragleave", () => { dropZone.style.background = "#f0f7ff"; });
dropZone.addEventListener("drop", (e) => { e.preventDefault(); dropZone.style.background = "#f0f7ff"; procesarExcel(e.dataTransfer.files[0]); });
excelInput.addEventListener("change", (e) => { if(e.target.files[0]) procesarExcel(e.target.files[0]); });

async function procesarExcel(file) {
  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });
    const ws = wb.Sheets["Personal"];
    if (!ws) { alert("❌ Hoja 'Personal' no encontrada"); return; }
    
    const jsonData = XLSX.utils.sheet_to_json(ws);
    if (jsonData.length === 0) { alert("❌ Hoja vacía"); return; }
    
    const registrosValidos = jsonData.filter(r => r.DNI && r.Código && r.Nombre);
    if (registrosValidos.length === 0) { alert("❌ Sin registros válidos"); return; }
    
    if (!window.confirm(`¿Importar ${registrosValidos.length} registros?`)) return;
    
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
          "Fecha de Ingreso": row["Fecha de Ingreso"] ? new Date(row["Fecha de Ingreso"]).toISOString().split("T")[0] : "",
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
      statusText.textContent = `📤 ${i + lote.length}/${registrosValidos.length}...`;
    }
    
    statusText.textContent = `✅ ${registrosValidos.length} importados`;
    setTimeout(() => { statusDiv.style.display = "none"; excelInput.value = ""; cargar(); }, 2000);
  } catch (error) {
    alert(`❌ ${error.message}`);
    document.getElementById("importStatus").style.display = "none";
  }
}

// CRUD MANUAL
async function guardar(e) {
  e.preventDefault();
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
    cargar();
  } catch (error) {
    alert(`❌ ${error.message}`);
  }
}

async function eliminar(id) {
  if (window.confirm("¿Eliminar?")) {
    try {
      await deleteDoc(doc(db, COLLECTION, id));
      cargar();
    } catch (error) {
      alert(`❌ ${error.message}`);
    }
  }
}

async function cargar() {
  try {
    const snap = await getDocs(col);
    allRecords = [];
    snap.forEach(d => allRecords.push({ id: d.id, ...d.data() }));
    mostrarTabla(allRecords);
  } catch (error) {
    console.error("Error:", error);
  }
}

function mostrarTabla(registros) {
  const tbody = document.getElementById("tbody");
  tbody.innerHTML = "";
  
  registros.forEach(p => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${p.DNI || ""}</td><td>${p.Código || ""}</td><td>${p.Nombre || ""}</td><td>${p.Foto || ""}</td><td>${p["Fecha de Ingreso"] || ""}</td><td>${p["Puesto Real"] || ""}</td><td>${p["CC Real"] || ""}</td><td>${p["Unidad Real"] || ""}</td><td>${p["Nombre Host"] || ""}</td><td>${p["Nueva Tajeta Micros"] || ""}</td><td class="del">✕</td>`;
    tr.querySelector(".del").addEventListener("click", () => eliminar(p.id));
    tbody.appendChild(tr);
  });
}

// BÚSQUEDA
document.getElementById("searchInput").addEventListener("input", (e) => {
  const q = e.target.value.toLowerCase();
  const filtrados = allRecords.filter(p => (p.Nombre || "").toLowerCase().includes(q) || (p.DNI || "").toLowerCase().includes(q));
  mostrarTabla(filtrados);
});

// LIMPIAR TODO
document.getElementById("clearAllBtn").addEventListener("click", async () => {
  if (allRecords.length === 0) { alert("Sin registros"); return; }
  if (!window.confirm(`⚠️ ¿Eliminar ${allRecords.length} registros?`)) return;
  
  try {
    const batchSize = 20;
    for (let i = 0; i < allRecords.length; i += batchSize) {
      const batch = writeBatch(db);
      allRecords.slice(i, i + batchSize).forEach(r => batch.delete(doc(db, COLLECTION, r.id)));
      await batch.commit();
    }
    alert(`✅ ${allRecords.length} eliminados`);
    cargar();
  } catch (error) {
    alert(`❌ ${error.message}`);
  }
});

// INICIALIZAR
document.getElementById("form-personal").addEventListener("submit", guardar);
cargar();
