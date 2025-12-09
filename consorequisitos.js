/* ================================
   📌 consorequisitos.gs
   Consolidar y comparar requisitos
   ================================ */

function consolidarRequisitos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaReq = ss.getSheetByName("Requ por procesos");
  const hojaNormas = ss.getSheetByName("Normas");

  if (!hojaReq || !hojaNormas) {
    SpreadsheetApp.getUi().alert("❌ No se encontró alguna de las hojas requeridas.");
    return;
  }

  // Mapeo columnas: Requ por procesos → Normas
  const mapeo = [
    { req: 5, normas: 1 },  // E → A
    { req: 6, normas: 2 },  // F → B
    { req: 7, normas: 3 },  // G → C
    { req: 8, normas: 4 },  // H → D
    { req: 9, normas: 5 },  // I → E
    { req: 10, normas: 6 }, // J → F
    { req: 11, normas: 7 }, // K → G
    { req: 12, normas: 8 }, // L → H
    { req: 13, normas: 9 }  // M → I
  ];

  const lastRowReq = hojaReq.getLastRow();
  const lastRowNormas = hojaNormas.getLastRow();

  if (lastRowReq < 2 || lastRowNormas < 2) {
    SpreadsheetApp.getUi().alert("⚠️ No hay datos suficientes para comparar.");
    return;
  }

  // Colores
  const COLOR_OK = "#ccffcc";   // verde claro
  const COLOR_NO = "#ffcccc";   // rojo claro

  // Limpiar colores previos en Normas (solo celdas relevantes)
  hojaNormas.getRange(2, 1, lastRowNormas - 1, 9).setBackground(null);

  // Recorremos cada par de columnas a comparar
  mapeo.forEach(par => {
    const colReq = par.req;
    const colNorm = par.normas;

    // Extraer datos columna REQ (fila 2 hacia abajo)
    const valoresReq = hojaReq
      .getRange(2, colReq, lastRowReq - 1, 1)
      .getValues()
      .map(r => normalizar(r[0]))
      .filter(v => v !== "");

    // Convertir a Set para búsquedas rápidas
    const setReq = new Set(valoresReq);

    // Datos de Normas (fila 2 hacia abajo)
    const rangoNormas = hojaNormas.getRange(2, colNorm, lastRowNormas - 1, 1);
    const valoresNormas = rangoNormas.getValues();

    // Determinar colores fila por fila
    const colores = valoresNormas.map(fila => {
      const valor = normalizar(fila[0]);

      if (valor === "") return [null]; // celda vacía → sin color

      if (setReq.has(valor)) {
        return [COLOR_OK]; // encontrado → verde
      } else {
        return [COLOR_NO]; // no encontrado → rojo
      }
    });

    // Aplicar colores en bloque
    rangoNormas.setBackgrounds(colores);
  });

  SpreadsheetApp.getUi().alert("✅ Consolidación completada correctamente.");
}

/*
 * Normalizar texto para comparación:
 * - Pasar a minúsculas
 * - Eliminar espacios al inicio y final
 * - Reemplazar múltiples espacios internos por 1
 */
function normalizar(v) {
  if (typeof v !== "string") v = String(v || "");
  return v
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

/*
 * Agregar el botón al menú "Auditoría"
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Auditoría")
    .addItem("Abrir visor", "mostrarModal")
    .addItem("🌐 Abrir web Plan de Auditoría", "abrirWebAuditoria")
    .addSeparator()
    .addItem("Consolida requisitos", "consolidarRequisitos")
    .addToUi();
}

