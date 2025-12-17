/**
 * ============================================================================
 * 📁 ARCHIVO: 05_Administracion.gs
 * 🎯 OBJETIVO: Importación y limpieza de Data Maestra externa.
 * ============================================================================
 */

function procesarImportacionMaestra(url) {
  try {
    // 1. Abrir libro externo
    let ssExterna;
    try {
      ssExterna = SpreadsheetApp.openByUrl(url);
    } catch (e) {
      return { success: false, error: "No se pudo acceder al enlace. Verifique permisos o URL." };
    }

    const hojaDetalle = ssExterna.getSheetByName("Detalle");
    if (!hojaDetalle) {
      return { success: false, error: "No se encontró la hoja 'Detalle' en el archivo." };
    }

    const lastRow = hojaDetalle.getLastRow();
    if (lastRow < 8) {
      return { success: false, error: "La hoja 'Detalle' parece estar vacía (menos de 8 filas)." };
    }

    // 2. Leer datos desde fila 8
    // Necesitamos H(8), I(9), J(10) para nombres y Y(25) para correo.
    // Leemos todo el rango hasta la columna Y para optimizar llamadas
    const numFilas = lastRow - 7; // Desde fila 8
    const datosCrudos = hojaDetalle.getRange(8, 1, numFilas, 25).getValues();

    const dataProcesada = [];

    // 3. Procesar fila por fila
    for (let i = 0; i < datosCrudos.length; i++) {
      const fila = datosCrudos[i];
      
      // Índices (base 0): H=7, I=8, J=9, Y=24
      const correo = String(fila[24] || "").trim();
      
      // FILTRO: Solo dominio @noel.com.co
      if (correo.toLowerCase().endsWith("@noel.com.co")) {
        const nombre = String(fila[7] || "").trim();
        const apellido1 = String(fila[8] || "").trim();
        const apellido2 = String(fila[9] || "").trim();
        
        // Unificación de nombre completo
        const nombreCompleto = `${nombre} ${apellido1} ${apellido2}`.replace(/\s+/g, " ").trim();
        
        if (nombreCompleto) {
          dataProcesada.push([nombreCompleto, correo]);
        }
      }
    }

    if (dataProcesada.length === 0) {
      return { success: false, error: "No se encontraron registros con correo @noel.com.co" };
    }

    // 4. Escribir en hoja local "Auditados"
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAuditados = ssLocal.getSheetByName("Auditados");
    
    if (!hojaAuditados) {
      return { success: false, error: "No existe la hoja 'Auditados' en este archivo." };
    }

    // Limpiar contenido anterior (desde fila 2)
    const lastRowLocal = hojaAuditados.getLastRow();
    if (lastRowLocal >= 2) {
      hojaAuditados.getRange(2, 1, lastRowLocal - 1, 2).clearContent();
    }

    // Pegar nuevos datos
    hojaAuditados.getRange(2, 1, dataProcesada.length, 2).setValues(dataProcesada);

    return { 
      success: true, 
      mensaje: `✅ Importación exitosa.\nRegistros cargados: ${dataProcesada.length}` 
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
function forzarPermisos() {
  SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1-kH-Fjy7FQA4k5A_YJIqPeUXQ5goGfEW/edit?gid=568542471#gid=568542471");
}
