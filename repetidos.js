function eliminarDuplicadosNuevoLooker_containerBound() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // usa el libro donde está el script
  const hoja = ss.getSheetByName("Auditados");
  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Auditados' en el archivo activo.");
    return;
  }

  const lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No hay datos (fila 2 en adelante) en la hoja.");
    return;
  }

  // Leer todo desde fila 2 hasta la última fila (todas las columnas usadas)
  const lastCol = hoja.getLastColumn();
  const dataRange = hoja.getRange(2, 1, lastRow - 1, lastCol);
  const datos = dataRange.getValues();

  const vistos = new Set();
  const datosFiltrados = [];

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    const nombreRaw = fila[0]; // Columna A
    const nombre = (nombreRaw || "").toString().trim();

    // Ignorar filas vacías (si quieres eliminar filas vacías, descomenta la línea siguiente)
    // if (!nombre) continue;

    const clave = nombre.toUpperCase(); // comparación case-insensitive
    if (!vistos.has(clave)) {
      vistos.add(clave);
      datosFiltrados.push(fila); // conservar primera aparición completa de la fila
    } else {
      // duplicado -> se omite (no se añade a datosFiltrados)
    }
  }

  // Borrar las filas originales y escribir las filtradas
  // Para no romper la hoja: mejor sobrescribir el rango a partir de la fila 2
  // 1) Limpiar el rango original (desde fila 2)
  hoja.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // 2) Escribir datos filtrados (si hay)
  if (datosFiltrados.length > 0) {
    hoja.getRange(2, 1, datosFiltrados.length, lastCol).setValues(datosFiltrados);
  }

  SpreadsheetApp.getUi().alert("Duplicados eliminados: quedaron " + datosFiltrados.length + " filas únicas (desde A2).");
}
