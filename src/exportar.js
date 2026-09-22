// Exportación a Excel con formato (colores de la app, sin cuadrícula) usando ExcelJS.
// ExcelJS se carga solo al exportar para no aumentar el tamaño inicial de la app.

const COLOR = {
  primario: "FF1F3A5F",
  primarioOscuro: "FF152A45",
  blanco: "FFFFFFFF",
  texto: "FF1F2933",
  gris: "FF5B6776",
  borde: "FFD9DEE5",
  zebra: "FFF8F9FB",
  grupos: ["FFDCE6F2", "FFEAF0F8", "FFF3F6FB", "FFF8FAFD"],
  estados: {
    verde: { fondo: "FFDCFCE7", texto: "FF166534" },
    ambar: { fondo: "FFFEF3C7", texto: "FF92400E" },
    rojo: { fondo: "FFFEE2E2", texto: "FF991B1B" }
  }
};

const relleno = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const bordeInferior = { bottom: { style: "thin", color: { argb: COLOR.borde } } };

function tonoEstado(v) {
  if (!v) return null;
  if (/inoperativo|obsoleto/i.test(v)) return "rojo";
  if (/revisar|reparar|repotenciar|custodia/i.test(v)) return "ambar";
  if (/operativo/i.test(v)) return "verde";
  return null;
}

/**
 * Descarga un Excel con formato.
 * @param {object} o
 * @param {string} o.archivo   nombre del archivo sin extensión
 * @param {string} o.hoja      nombre de la hoja
 * @param {string} o.titulo    título grande de la primera fila
 * @param {string} [o.detalle] línea descriptiva (filtros aplicados, etc.)
 * @param {string[]} o.columnas
 * @param {Array} o.filas      objetos { columna: valor } o, en reportes agrupados,
 *                             { grupo: { nivel, texto } } para las filas de grupo
 * @param {string} [o.total]   texto de la fila final (p. ej. "Total general: 120")
 */
export async function exportarExcel({ archivo, hoja, titulo, detalle = "", columnas, filas, total }) {
  const { default: ExcelJS } = await import("exceljs");
  const libro = new ExcelJS.Workbook();
  libro.creator = "inventaTISwiss";
  libro.created = new Date();

  const ws = libro.addWorksheet(hoja.slice(0, 31), {
    views: [{ showGridLines: false, state: "frozen", ySplit: 4 }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });
  const nCols = Math.max(1, columnas.length);

  // Fila 1: título
  ws.mergeCells(1, 1, 1, nCols);
  const t = ws.getCell(1, 1);
  t.value = titulo;
  t.font = { name: "Calibri", size: 16, bold: true, color: { argb: COLOR.blanco } };
  t.fill = relleno(COLOR.primario);
  t.alignment = { vertical: "middle", indent: 1 };
  ws.getRow(1).height = 30;

  // Fila 2: detalle y fecha
  ws.mergeCells(2, 1, 2, nCols);
  const d = ws.getCell(2, 1);
  const fecha = new Date().toLocaleString("es-PE", { dateStyle: "long", timeStyle: "short" });
  d.value = `${detalle ? `${detalle} · ` : ""}Generado el ${fecha} · Swissôtel Lima`;
  d.font = { name: "Calibri", size: 10, italic: true, color: { argb: COLOR.gris } };
  d.alignment = { indent: 1 };
  ws.getRow(2).height = 18;

  // Fila 4: cabecera
  const cab = ws.getRow(4);
  columnas.forEach((c, i) => {
    const celda = cab.getCell(i + 1);
    celda.value = c;
    celda.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLOR.blanco } };
    celda.fill = relleno(COLOR.primarioOscuro);
    celda.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  });
  cab.height = 24;

  // Datos
  let fila = 5;
  let zebra = false;
  filas.forEach(r => {
    const row = ws.getRow(fila);
    if (r.grupo) {
      ws.mergeCells(fila, 1, fila, nCols);
      const c = row.getCell(1);
      c.value = r.grupo.texto;
      c.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLOR.primario } };
      c.fill = relleno(COLOR.grupos[Math.min(r.grupo.nivel, COLOR.grupos.length - 1)]);
      c.alignment = { indent: 1 + r.grupo.nivel * 2 };
      row.height = 20;
      zebra = false;
    } else {
      columnas.forEach((col, i) => {
        const celda = row.getCell(i + 1);
        const v = r[col] ?? "";
        celda.value = typeof v === "number" ? v : String(v);
        celda.font = { name: "Calibri", size: 10, color: { argb: COLOR.texto } };
        celda.border = bordeInferior;
        celda.alignment = { vertical: "middle", indent: 1 };
        if (zebra) celda.fill = relleno(COLOR.zebra);
        const tono = col === "Estado" ? tonoEstado(String(v)) : null;
        if (tono) {
          celda.fill = relleno(COLOR.estados[tono].fondo);
          celda.font = { name: "Calibri", size: 10, bold: true, color: { argb: COLOR.estados[tono].texto } };
        }
        if (v === "#N/A") celda.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FFB45309" } };
      });
      zebra = !zebra;
    }
    fila++;
  });

  // Fila de total
  if (total) {
    ws.mergeCells(fila, 1, fila, nCols);
    const c = ws.getCell(fila, 1);
    c.value = total;
    c.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLOR.blanco } };
    c.fill = relleno(COLOR.primario);
    c.alignment = { indent: 1 };
    ws.getRow(fila).height = 22;
  }

  // Filtros de Excel en la cabecera cuando no hay filas de grupo
  if (!filas.some(r => r.grupo) && filas.length) {
    ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + filas.length, column: nCols } };
  }

  // Ancho de columnas según el contenido
  columnas.forEach((col, i) => {
    const largo = Math.max(col.length, ...filas.filter(r => !r.grupo).map(r => String(r[col] ?? "").length));
    ws.getColumn(i + 1).width = Math.min(45, Math.max(10, largo + 4));
  });

  const buffer = await libro.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${archivo} ${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
