import { db } from "./firebase-config.js";
import { collection, addDoc, getDocs, deleteDoc, doc } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";

// Nombre de la colección: cambia esto (y los campos del form) para crear más sistemas
// a partir de otras hojas, ej. "inventarioSwiss", "dinamica", "pcLaptop", etc.
const COLLECTION = "personal";
const col = collection(db, COLLECTION);

window.guardar = async function () {
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
};

window.eliminar = async function (id) {
  await deleteDoc(doc(db, COLLECTION, id));
  cargar();
};

window.cargar = async function () {
  const snap = await getDocs(col);
  const tbody = document.getElementById("tbody");
  tbody.innerHTML = "";
  snap.forEach(d => {
    const p = d.data();
    tbody.innerHTML += `<tr>
      <td>${p.dni || ""}</td><td>${p.codigo || ""}</td><td>${p.nombre || ""}</td>
      <td>${p.foto || ""}</td><td>${p.fechaIngreso || ""}</td><td>${p.puestoReal || ""}</td>
      <td>${p.ccReal || ""}</td><td>${p.unidadReal || ""}</td><td>${p.nombreHost || ""}</td>
      <td>${p.nuevaTarjetaMicros || ""}</td>
      <td class="del" onclick="eliminar('${d.id}')">✕</td></tr>`;
  });
};

cargar();
