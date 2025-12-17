/**
 * ============================================================================
 * 📁 ARCHIVO: 04_Visor_Backend.gs
 * 🎯 OBJETIVO: Backend para el Visor Modal (Relaciones Procesos <-> Normas).
 * ============================================================================
 */

/**
 * VISTA 1: Obtiene estructura jerárquica: Proceso -> Subproceso -> Responsable -> Norma -> Requisito
 */
function obtenerDatos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requ por procesos");
    if (!hoja) return {};

    const data = hoja.getDataRange().getValues();
    if (data.length < 2) return {};

    const encabezados = data[0];
    const procesos = {};
    
    // Índices de columnas donde puede haber normas (ajustar según tu hoja real)
    const columnasNormas = [4, 5, 6, 7, 8, 9, 11, 12, 13, 14];

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      
      // Construir jerarquía
      const proc = fila[0] ? fila[0].toString().trim() : "";
      const codSub = fila[1] ? fila[1].toString().trim() : "";
      const nomSub = fila[2] ? fila[2].toString().trim() : "";
      const sub = (codSub + " " + nomSub).trim();
      const resp = fila[3] ? fila[3].toString().trim() : "";

      if (!proc || !sub || !resp) continue;

      if (!procesos[proc]) procesos[proc] = {};
      if (!procesos[proc][sub]) procesos[proc][sub] = {};
      if (!procesos[proc][sub][resp]) procesos[proc][sub][resp] = {};

      // Buscar requisitos marcados en las columnas de normas
      columnasNormas.forEach(idx => {
        const norma = encabezados[idx] ? encabezados[idx].toString().trim() : "";
        const requisito = fila[idx] ? fila[idx].toString().trim() : "";
        
        if (norma && requisito) {
          if (!procesos[proc][sub][resp][norma]) {
            procesos[proc][sub][resp][norma] = [];
          }
          const lista = procesos[proc][sub][resp][norma];
          if (!lista.includes(requisito)) lista.push(requisito);
        }
      });
    }
    return procesos;
  } catch (e) {
    Logger.log("❌ Error Visor Datos: " + e.toString());
    return {};
  }
}

/**
 * Obtiene la descripción detallada de un requisito específico desde la hoja "Requisitos".
 */
function obtenerDescripcion(norma, numero) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisitos");
    if (!hoja) return null;

    const data = hoja.getDataRange().getValues();
    // Búsqueda simple: columna A = norma, columna B = número requisito
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == norma && data[i][1] == numero) {
        return {
          titulo: data[i][2],      // Título del requisito
          subtitulo: data[i][3],   // Subtítulo
          descripcion: data[i][4], // Texto largo
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

/**
 * VISTA 2 (Inversa): Norma -> Requisito -> Proceso -> Subproceso -> Responsable
 */
function obtenerDatosReversa() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requ por procesos");
    if (!hoja) return {};

    const data = hoja.getDataRange().getValues();
    if (data.length < 2) return {};

    const encabezados = data[0];
    const datosInvertidos = {};
    const columnasNormasIdx = [4, 5, 6, 7, 8, 10, 11, 12, 13];

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];

      const proc = fila[0] ? fila[0].toString().trim() : "";
      const sub = ((fila[1]||"") + " " + (fila[2]||"")).trim();
      const resp = fila[3] ? fila[3].toString().trim() : "";

      if (!proc || !sub || !resp) continue;

      columnasNormasIdx.forEach((idx) => {
        if (idx >= encabezados.length) return;
        const norma = encabezados[idx] ? encabezados[idx].toString().trim() : "";
        if (idx >= fila.length) return;
        const requisito = fila[idx] ? fila[idx].toString().trim() : "";

        if (norma && requisito) {
          // Construcción del árbol invertido
          if (!datosInvertidos[norma]) datosInvertidos[norma] = {};
          if (!datosInvertidos[norma][requisito]) datosInvertidos[norma][requisito] = {};
          if (!datosInvertidos[norma][requisito][proc]) datosInvertidos[norma][requisito][proc] = {};
          if (!datosInvertidos[norma][requisito][proc][sub]) datosInvertidos[norma][requisito][proc][sub] = [];
          
          const listaResp = datosInvertidos[norma][requisito][proc][sub];
          if (!listaResp.includes(resp)) listaResp.push(resp);
        }
      });
    }
    return datosInvertidos;
  } catch (e) {
    Logger.log("❌ Error Visor Reversa: " + e.toString());
    return {};
  }
}
