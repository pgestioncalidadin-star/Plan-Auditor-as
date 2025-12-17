/**
 * ============================================================================
 * 📁 ARCHIVO: 02_Gestion_Datos.gs
 * 🎯 OBJETIVO: CRUD y lógica de negocio (Fechas, Edición en sitio, Calendario).
 * ============================================================================
 */

function obtenerAuditores() { return obtenerPersonasGenerico("Auditores"); }
function obtenerAuditados() { return obtenerPersonasGenerico("Auditados"); }

function obtenerPersonasGenerico(nombreHoja) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return [];
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return [];
    const data = hoja.getRange(2, 1, lastRow - 1, 2).getValues();
    return data.filter(row => row[0]).map(row => ({ 
      nombre: row[0].toString().trim(), 
      correo: row[1] ? row[1].toString().trim() : "" 
    }));
  } catch (e) { return []; }
}

function obtenerProcesos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Procesos");
    if (!hoja) return {};
    const lastCol = hoja.getLastColumn();
    const encabezados = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
    const procesos = {};
    encabezados.forEach((proceso, index) => {
      if (proceso) {
        const nombreProceso = proceso.toString().trim();
        const columna = index + 1;
        const lastRow = hoja.getLastRow();
        let subprocesos = [];
        if (lastRow > 1) {
          subprocesos = hoja.getRange(2, columna, lastRow - 1, 1).getValues().flat().filter(s => s).map(s => s.toString().trim());
        }
        procesos[nombreProceso] = subprocesos;
      }
    });
    return procesos;
  } catch (e) { return {}; }
}

function obtenerNormas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Normas");
    if (!hoja) return {};
    const lastCol = hoja.getLastColumn();
    const lastRow = hoja.getLastRow();
    const encabezados = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
    const normas = {};
    encabezados.forEach((norma, index) => {
      if (norma) {
        const nombreNorma = norma.toString().trim();
        const columna = index + 1;
        if (lastRow > 1) {
          const rango = hoja.getRange(2, columna, lastRow - 1, 1);
          const valores = rango.getValues();
          const colores = rango.getBackgrounds();
          const verdes = [];
          for (let i = 0; i < valores.length; i++) {
            if (valores[i][0] && colores[i][0].toLowerCase() === "#ccffcc") {
              verdes.push(valores[i][0].toString().trim());
            }
          }
          normas[nombreNorma] = verdes;
        } else { normas[nombreNorma] = []; }
      }
    });
    return normas;
  } catch (e) { return {}; }
}

function obtenerAuditoriasCalendario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro Auditoría");
  if (!hoja) return [];
  const data = hoja.getDataRange().getValues();
  const output = [];
  const tz = ss.getSpreadsheetTimeZone();
  
  // Columna L (11) es Fecha.
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (!row[0] || !row[11]) continue; 

    let fechaRaw = row[11];
    let fechaStr = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, tz, "yyyy-MM-dd") : String(fechaRaw).substring(0, 10);
    const fmtHora = (v) => (v instanceof Date) ? Utilities.formatDate(v, tz, "HH:mm") : v;

    output.push({
      "id": row[0],
      "nombre": row[1],
      "fecha": fechaStr, 
      "inicio": fmtHora(row[12]),
      "fin": fmtHora(row[13]),
      "auditor": row[5], 
      "auditado": row[7] 
    });
  }
  return output;
}

function obtenerIDsAuditorias() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Registro Auditoría");
    if (!hoja || hoja.getLastRow() < 2) return [];
    const idsRaw = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
    return [...new Set(idsRaw.flat().filter(String).map(id => String(id)))].sort();
  } catch (e) { return []; }
}

function obtenerAuditoriaPorID(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Registro Auditoría");
    if (!hoja) return null;

    const data = hoja.getDataRange().getValues();
    const idBuscado = String(id).trim().toUpperCase();
    const filas = data.filter(r => String(r[0]).trim().toUpperCase() === idBuscado);

    if (filas.length === 0) return null;

    const base = filas[0];
    const resultado = {
      id: base[0],
      nombre: base[1],
      objetivos: base[2] ? String(base[2]).split(", ") : [],
      alcance: base[3],
      modalidad: base[9],
      tipo: base[10],
      observaciones: base[17],
      horarios: []
    };

    const tz = ss.getSpreadsheetTimeZone();
    const fmtHora = (v) => (v instanceof Date) ? Utilities.formatDate(v, tz, "HH:mm") : v;
    const parsearBloque = (texto) => {
      if (!texto) return [];
      return String(texto).split(" | ").map(bloque => {
        const partes = bloque.split(": ");
        return { titulo: partes[0].trim(), items: (partes[1] || "").split(", ").filter(x => x.trim() !== "") };
      });
    };

    filas.forEach(f => {
      const audNoms = f[5] ? String(f[5]).split(", ") : [];
      const audMails = f[6] ? String(f[6]).split(", ") : [];
      const auditadosNoms = f[7] ? String(f[7]).split(", ") : [];
      const auditadosMails = f[8] ? String(f[8]).split(", ") : [];

      let fechaStr = f[11];
      if (fechaStr instanceof Date) fechaStr = Utilities.formatDate(fechaStr, tz, "yyyy-MM-dd");

      resultado.horarios.push({
        fecha: fechaStr,
        inicio: fmtHora(f[12]),
        fin: fmtHora(f[13]),
        descripcion: f[14],
        auditores: audNoms.map((n, i) => ({ nombre: n, correo: audMails[i] || "" })),
        auditados: auditadosNoms.map((n, i) => ({ nombre: n, correo: auditadosMails[i] || "" })),
        procesos: parsearBloque(f[15]).map(p => ({ proceso: p.titulo, subprocesos: p.items })),
        normas: parsearBloque(f[16]).map(n => ({ norma: n.titulo, requisitos: n.items }))
      });
    });
    return resultado;
  } catch (e) { return null; }
}

// 📝 ESCRITURA
function crearAuditoria(datos) { return guardarAuditoriaGeneral(datos, true); }
function editarAuditoria(datos) { return guardarAuditoriaGeneral(datos, false); }

function guardarAuditoriaGeneral(datos, esNuevo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Registro Auditoría");
    
    if (!hoja) {
      hoja = ss.insertSheet("Registro Auditoría");
      hoja.appendRow(["ID", "Nombre Auditoría", "Objetivos", "Alcance", "Riesgos", "Auditor(es)", "Correo Auditor(es)", "Auditado(s)", "Correo Auditado(s)", "Modalidad", "Tipo de auditoría", "Fecha", "Hora inicio", "Hora Fin", "Descripción", "Procesos", "Requisitos", "Observaciones", "Fecha de creación"]);
    }

    let idFinal = datos.id;
    let indiceInsercion = hoja.getLastRow() + 1;

    if (esNuevo) {
      idFinal = generarNuevoID(hoja);
    } else {
      const lastRow = hoja.getLastRow();
      let filasEliminadas = 0;
      let filaEncontrada = -1;
      
      const auditoriaAntigua = obtenerAuditoriaPorID(idFinal);
      if (auditoriaAntigua) desmarcarRequisitosAuditados(auditoriaAntigua);

      // Borramos filas antiguas y guardamos el índice para insertar en el mismo sitio
      for (let i = lastRow; i >= 2; i--) {
        const idEnFila = String(hoja.getRange(i, 1).getValue());
        if (idEnFila === String(idFinal)) {
          filaEncontrada = i; 
          hoja.deleteRow(i);
          filasEliminadas++;
        }
      }
      if (filasEliminadas > 0) indiceInsercion = filaEncontrada;
    }

    const fechaLog = obtenerFechaBogota();
    const nuevasFilas = [];

    datos.horarios.forEach(h => {
      const procesosTxt = (h.procesos || []).map(p => `${p.proceso}: ${p.subprocesos.join(', ')}`).join(' | ');
      const normasTxt = (h.normas || []).map(n => `${n.norma}: ${n.requisitos.join(', ')}`).join(' | ');
      
      const auditoresNombres = (h.auditores || []).map(a => a.nombre || a).join(", ");
      const auditoresCorreos = (h.correosAuditores || []).join(", ");
      const auditadosNombres = (h.auditados || []).map(a => a.nombre || a).join(", ");
      const auditadosCorreos = (h.correosAuditados || []).join(", ");

      nuevasFilas.push([
        idFinal, datos.nombre, (datos.objetivos || []).join(", "), datos.alcance, "Ver los riesgos en el documento 'Programa de auditorías'",
        auditoresNombres, auditoresCorreos, auditadosNombres, auditadosCorreos, 
        datos.modalidad, datos.tipo, h.fecha, h.inicio, h.fin, h.descripcion,    
        procesosTxt, normasTxt, datos.observaciones, fechaLog          
      ]);
    });

    if (nuevasFilas.length > 0) {
      if (indiceInsercion > hoja.getLastRow()) {
        nuevasFilas.forEach(fila => hoja.appendRow(fila));
      } else {
        hoja.insertRows(indiceInsercion, nuevasFilas.length);
        hoja.getRange(indiceInsercion, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
      }
    }

    marcarRequisitosAuditados(datos);       
    crearEventosCalendario(datos, idFinal); 
    enviarCorreosAuditoria(datos, idFinal); 

    return { success: true, id: idFinal };

  } catch (e) {
    Logger.log("❌ Error guardando auditoría: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function generarNuevoID(hoja) {
  try {
    if (hoja.getLastRow() < 2) return "PA001";
    const ids = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().flat();
    let max = 0;
    ids.forEach(id => {
      const num = parseInt(String(id).replace("PA", ""), 10);
      if (!isNaN(num) && num > max) max = num;
    });
    return "PA" + String(max + 1).padStart(3, "0");
  } catch (e) { return "PA001"; }
}

function marcarRequisitosAuditados(datos) { cambiarColorRequisitos(datos, "#ccffcc"); }
function desmarcarRequisitosAuditados(datos) { cambiarColorRequisitos(datos, null); }

function cambiarColorRequisitos(datos, color) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisitos faltantes por auditar");
    if (!hoja) return;
    const lastRow = hoja.getLastRow();
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    const reqSeleccionados = new Set();

    datos.horarios.forEach(h => {
      if (h.normas) h.normas.forEach(n => n.requisitos.forEach(r => reqSeleccionados.add(`${n.norma}::${r.trim()}`)));
    });

    encabezados.forEach((normaHead, colIndex) => {
      if (!normaHead) return;
      const nombreNorma = normaHead.toString().trim();
      if (lastRow > 1) {
        const rango = hoja.getRange(2, colIndex + 1, lastRow - 1, 1);
        const valores = rango.getValues();
        const fondos = rango.getBackgrounds();
        let cambios = false;
        for (let i = 0; i < valores.length; i++) {
          const req = valores[i][0] ? valores[i][0].toString().trim() : "";
          if (req && reqSeleccionados.has(`${nombreNorma}::${req}`)) {
            fondos[i][0] = color;
            cambios = true;
          }
        }
        if (cambios) rango.setBackgrounds(fondos);
      }
    });
  } catch (e) { Logger.log("⚠️ Error colores: " + e.toString()); }
}

function consolidarRequisitos() {
  // Misma lógica de consolidación
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaReq = ss.getSheetByName("Requ por procesos");
  const hojaNormas = ss.getSheetByName("Normas");
  if (!hojaReq || !hojaNormas) { SpreadsheetApp.getUi().alert("❌ Faltan hojas"); return; }
  const mapeo = [{ req: 5, normas: 1 }, { req: 6, normas: 2 }, { req: 7, normas: 3 }, { req: 8, normas: 4 }, { req: 9, normas: 5 }, { req: 10, normas: 6 }, { req: 11, normas: 7 }, { req: 12, normas: 8 }, { req: 13, normas: 9 }];
  const lastRowReq = hojaReq.getLastRow();
  const lastRowNormas = hojaNormas.getLastRow();
  if (lastRowReq < 2 || lastRowNormas < 2) return;
  hojaNormas.getRange(2, 1, lastRowNormas - 1, 9).setBackground(null);
  mapeo.forEach(par => {
    const valReq = hojaReq.getRange(2, par.req, lastRowReq - 1, 1).getValues().map(r => normalizar(r[0])).filter(v => v !== "");
    const setReq = new Set(valReq);
    const rangoDest = hojaNormas.getRange(2, par.normas, lastRowNormas - 1, 1);
    const valDest = rangoDest.getValues();
    const nuevosColores = valDest.map(fila => {
      const v = normalizar(fila[0]);
      return (v !== "" && setReq.has(v)) ? ["#ccffcc"] : (v !== "" ? ["#ffcccc"] : [null]);
    });
    rangoDest.setBackgrounds(nuevosColores);
  });
  SpreadsheetApp.getUi().alert("✅ Consolidación completada.");
}

function finDeCiclo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Requisitos faltantes por auditar");
  if (hoja && hoja.getLastRow() >= 2) {
    hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).setBackground(null);
    SpreadsheetApp.getUi().alert("✅ Ciclo reiniciado.");
  }
}
