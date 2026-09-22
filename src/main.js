import { db } from "./firebase-config.js";
import { collection, addDoc, getDocs, deleteDoc, doc } from "firebase/firestore";

// Nombre de la colección: cambia esto (y los campos del form) para crear más sistemas
// a partir de otras hojas, ej. "inventarioSwiss", "dinamica", "pcLaptop", etc.
const COLLECTION = "personal";
const col = collection(db, COLLECTION);

async function guardar(e) {
  e.preventDefault();
  const f = document.forms.f;
  const data = {
    dni: f.dni.value,
    codigo: f.codigo.value,
    nombre: f.nombre.value,
    foto: f.foto.value,
    fechaIngreso: f.fechaIngreso.value,
    puestoReal: f.puestoReal.value,
    ccReal: f.ccReal.value,
    unidadReal: f.unidadReal.value,
    nombreHost: f.nombreHost.value,
    nuevaTarjetaMicros: f.nuevaTarjetaMicros.value
  };
  await addDoc(col, data);
  f.reset();
  cargar();
}

async function eliminar(id) {
  await deleteDoc(doc(db, COLLECTION, id));
  cargar();
}

async function cargar() {
  const snap = await getDocs(col);
  const tbody = document.getElementById("tbody");
  tbody.innerHTML = "";
  snap.forEach(d => {
    const p = d.data();
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${p.dni || ""}</td><td>${p.codigo || ""}</td><td>${p.nombre || ""}</td>
      <td>${p.foto || ""}</td><td>${p.fechaIngreso || ""}</td><td>${p.puestoReal || ""}</td>
      <td>${p.ccReal || ""}</td><td>${p.unidadReal || ""}</td><td>${p.nombreHost || ""}</td>
      <td>${p.nuevaTarjetaMicros || ""}</td>
      <td class="del">✕</td>`;
    tr.querySelector(".del").addEventListener("click", () => eliminar(d.id));
    tbody.appendChild(tr);
  });
}

document.getElementById("form-personal").addEventListener("submit", guardar);
cargar();
