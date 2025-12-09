/***************************************************
 * 📘 PLAN DE AUDITORÍAS INTEGRADAS - VERSIÓN MAESTRA FINAL
 * Archivo: plan.gs
 ***************************************************/

// ✅ MENÚ PRINCIPAL
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🔎 Auditorías');
  
  // Herramientas de visualización
  menu.addItem('📋 Abrir visor', 'mostrarVisorAuditorias');
  menu.addItem('🌐 Abrir web Plan de Auditoría', 'abrirWebAuditoria');
  
  menu.addSeparator();
  
  // Herramientas de Gestión
  menu.addItem('🔄 Consolidar requisitos', 'consolidarRequisitos');
  menu.addItem('🏁 Fin de ciclo (Limpiar)', 'finDeCiclo');
  
  menu.addToUi();
}

/***************************************************
 * 🌐 FUNCIÓN WEB: doGet()
 ***************************************************/
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('planaudi')
    .setTitle('🗓️ Planear Auditoría Integrada')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***************************************************
 * 🚀 FUNCIONES DE APERTURA DE MODALES
 ***************************************************/
function abrirWebAuditoria() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif; text-align:center; padding:30px">
      <h2>🌐 Acceso a la Web</h2>
      <a href="${url}" target="_blank"
         style="display:inline-block; background:#9e1a18; color:white; padding:12px 24px; border-radius:6px; text-decoration:none; font-weight:bold;">
        🔗 Abrir Planificador
      </a>
    </div>
  `).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '🌐 Web del Plan de Auditoría');
}

function mostrarVisorAuditorias() {
  const html = HtmlService.createHtmlOutputFromFile('Modal')
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Visor de Auditorías');
}

/***************************************************
 * 🔄 CONSOLIDAR REQUISITOS (Lógica Original)
 ***************************************************/
function consolidarRequisitos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaReq = ss.getSheetByName("Requ por procesos");
    const hojaNormas = ss.getSheetByName("Normas");

    if (!hojaReq || !hojaNormas) {
      SpreadsheetApp.getUi().alert("❌ No se encontró alguna de las hojas requeridas.");
      return;
    }

    const mapeo = [
      { req: 5, normas: 1 }, { req: 6, normas: 2 }, { req: 7, normas: 3 },
      { req: 8, normas: 4 }, { req: 9, normas: 5 }, { req: 10, normas: 6 },
      { req: 11, normas: 7 }, { req: 12, normas: 8 }, { req: 13, normas: 9 }
    ];

    const lastRowReq = hojaReq.getLastRow();
    const lastRowNormas = hojaNormas.getLastRow();

    if (lastRowReq < 2 || lastRowNormas < 2) {
      SpreadsheetApp.getUi().alert("⚠️ No hay datos suficientes.");
      return;
    }

    const COLOR_OK = "#ccffcc";
    const COLOR_NO = "#ffcccc";

    hojaNormas.getRange(2, 1, lastRowNormas - 1, 9).setBackground(null);

    mapeo.forEach(par => {
      const valoresReq = hojaReq.getRange(2, par.req, lastRowReq - 1, 1).getValues()
        .map(r => normalizar(r[0])).filter(v => v !== "");
      const setReq = new Set(valoresReq);
      const rangoNormas = hojaNormas.getRange(2, par.normas, lastRowNormas - 1, 1);
      const valoresNormas = rangoNormas.getValues();

      const colores = valoresNormas.map(fila => {
        const valor = normalizar(fila[0]);
        if (valor === "") return [null];
        return setReq.has(valor) ? [COLOR_OK] : [COLOR_NO];
      });
      rangoNormas.setBackgrounds(colores);
    });

    SpreadsheetApp.getUi().alert("✅ Consolidación completada correctamente.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.toString());
  }
}

function normalizar(v) {
  if (typeof v !== "string") v = String(v || "");
  return v.trim().toLowerCase().replace(/\s+/g, " ");
}

/***************************************************
 * 🏁 FIN DE CICLO (Limpia la hoja de control)
 ***************************************************/
function finDeCiclo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisitos faltantes por auditar");
    
    if (!hoja) {
      SpreadsheetApp.getUi().alert("❌ No se encontró la hoja 'Requisitos faltantes por auditar'");
      return;
    }

    const lastRow = hoja.getLastRow();
    const lastCol = hoja.getLastColumn();
    
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("⚠️ La hoja está vacía, no hay nada que limpiar.");
      return;
    }

    // Limpiar formatos (colores) desde fila 2
    hoja.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);
    
    SpreadsheetApp.getUi().alert("✅ ¡Ciclo finalizado! Se han limpiado los indicadores verdes.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.toString());
  }
}

/***************************************************
 * 💾 CRUD AUDITORÍAS (CREAR Y EDITAR) - SOPORTE MULTI-PERSONA
 ***************************************************/
function crearAuditoria(datos) {
  return procesarGuardado(datos, true);
}

function editarAuditoria(datos) {
  return procesarGuardado(datos, false);
}

function procesarGuardado(datos, esNuevo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Registro Auditoría");
    
    if (!hoja) {
      hoja = ss.insertSheet("Registro Auditoría");
      hoja.appendRow([
        "ID", "Nombre Auditoría", "Objetivos", "Alcance", "Riesgos", 
        "Modalidad", "Tipo de auditoría", "Fecha", "Hora Inicio", "Hora Fin", 
        "Descripción", "Auditores (Nombres)", "Auditores (Correos)", 
        "Auditados (Nombres)", "Auditados (Correos)", "Procesos", "Requisitos", 
        "Observaciones", "Fecha de creación"
      ]);
    }

    const id = esNuevo ? generarNuevoID() : datos.id;
    
    // Si editamos, borramos registros anteriores de este ID
    if (!esNuevo) {
      const ultimaFila = hoja.getLastRow();
      if (ultimaFila >= 2) {
        const rangoIds = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
        // Borramos de abajo hacia arriba para no alterar índices
        for (let i = rangoIds.length - 1; i >= 0; i--) {
          if (rangoIds[i][0] == id) {
            hoja.deleteRow(i + 2); // +2 porque el array es base 0 y hoja tiene header
          }
        }
      }
    }

    // Insertar nuevas filas (Una por cada Horario)
    datos.horarios.forEach(h => {
      // Aplanar Procesos y Normas a String
      const procesosTexto = h.procesos.map(p => `${p.proceso}: ${p.subprocesos.join(', ')}`).join(' | ');
      const requisitosTexto = h.normas.map(n => `${n.norma}: ${n.requisitos.join(', ')}`).join(' | ');
      
      // Aplanar Personas a String (separado por comas para guardar en celda)
      // Nota: h.auditores es un array de objetos {nombre, correo}
      const auditoresNom = h.auditores ? h.auditores.map(a => a.nombre).join(", ") : "";
      const auditoresMail = h.auditores ? h.auditores.map(a => a.correo).join(", ") : "";
      const auditadosNom = h.auditados ? h.auditados.map(a => a.nombre).join(", ") : "";
      const auditadosMail = h.auditados ? h.auditados.map(a => a.correo).join(", ") : "";

      hoja.appendRow([
        id,
        datos.nombre,
        datos.objetivos.join(", "),
        datos.alcance,
        "Ver los riesgos en el documento 'Programa de auditorías'",
        datos.modalidad,
        datos.tipo,
        h.fecha,
        h.inicio,
        h.fin,
        h.descripcion,
        auditoresNom,
        auditoresMail,
        auditadosNom,
        auditadosMail,
        procesosTexto,
        requisitosTexto,
        datos.observaciones,
        new Date()
      ]);
    });

    // 1. Marcar requisitos en hoja de control (Verde)
    marcarRequisitosEnHojaControl(datos);

    // 2. Crear eventos en calendario (Múltiples invitados)
    crearEventosCalendario(datos, id);

    // 3. Enviar correos (Formato limpio)
    enviarCorreosAuditoria(datos, id);

    return { success: true, id: id };
    
  } catch (error) {
    Logger.log("Error Procesando: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/***************************************************
 * 🟢 MARCAR REQUISITOS EN VERDE
 ***************************************************/
function marcarRequisitosEnHojaControl(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisitos faltantes por auditar");
    if (!hoja) return;

    const data = hoja.getDataRange().getValues(); 
    const encabezados = data[0]; // Fila 1 tiene las Normas (ISO 9001, etc.)

    datos.horarios.forEach(h => {
      if (h.normas) {
        h.normas.forEach(n => {
          const normaNombre = n.norma; 
          const requisitos = n.requisitos; 

          const colIndex = encabezados.indexOf(normaNombre);
          if (colIndex !== -1) {
            requisitos.forEach(req => {
              for (let fila = 1; fila < data.length; fila++) {
                if (String(data[fila][colIndex]).trim() === String(req).trim()) {
                  hoja.getRange(fila + 1, colIndex + 1).setBackground("#ccffcc");
                }
              }
            });
          }
        });
      }
    });
  } catch (e) {
    console.error("Error marcando verdes: " + e);
  }
}

/***************************************************
 * 📧 ENVÍO DE CORREOS (FORMATO LIMPIO)
 ***************************************************/
function enviarCorreosAuditoria(datos, idAuditoria) {
  const todosCorreos = new Set();
  
  // Recolectar todos los correos involucrados en la auditoría
  datos.horarios.forEach(h => {
    if (h.auditores) h.auditores.forEach(p => { if(p.correo) todosCorreos.add(p.correo.trim()); });
    if (h.auditados) h.auditados.forEach(p => { if(p.correo) todosCorreos.add(p.correo.trim()); });
  });

  if (todosCorreos.size === 0) return;

  const template = HtmlService.createTemplateFromFile('correo_auditoria');
  template.datos = datos;
  template.id = idAuditoria;
  const mensajeHtml = template.evaluate().getContent();

  todosCorreos.forEach(email => {
    try {
      MailApp.sendEmail({
        to: email,
        subject: `🗓️ Citación Auditoría: ${idAuditoria} - ${datos.nombre}`,
        htmlBody: mensajeHtml
      });
    } catch (e) {
      Logger.log("Error enviando a " + email + ": " + e);
    }
  });
}

/***************************************************
 * 📅 CALENDARIO (MÚLTIPLES INVITADOS)
 ***************************************************/
function crearEventosCalendario(datos, id) {
  datos.horarios.forEach(h => {
    try {
      const inicio = new Date(h.fecha + 'T' + h.inicio);
      const fin = new Date(h.fecha + 'T' + h.fin);
      const titulo = `🗓️ Auditoría ${id}: ${datos.nombre}`;
      
      let invitados = [];
      if (h.auditores) invitados = invitados.concat(h.auditores.map(a => a.correo));
      if (h.auditados) invitados = invitados.concat(h.auditados.map(a => a.correo));
      
      const descripcion = `
Auditoría: ${datos.nombre}
Modalidad: ${datos.modalidad}
------------------
${h.descripcion || "Sin descripción adicional"}
      `.trim();

      CalendarApp.getDefaultCalendar().createEvent(titulo, inicio, fin, {
        description: descripcion,
        guests: invitados.join(','),
        sendInvites: true
      });
    } catch (e) {
      Logger.log("Error calendario: " + e);
    }
  });
}

/***************************************************
 * 🔍 FUNCIONES DEL VISOR (ORIGINALES)
 ***************************************************/
function obtenerDatos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requ por procesos");
    if (!hoja) return {};
    const data = hoja.getDataRange().getValues();
    if (data.length < 2) return {};
    const encabezados = data[0];
    const procesos = {};
    let procesoActual = "", subprocesoActual = "", responsableActual = "";
    const columnasNormas = [4, 5, 6, 7, 8, 9, 11, 12, 13, 14];
    
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      if (fila[0]) procesoActual = fila[0].toString().trim();
      if (!procesos[procesoActual]) procesos[procesoActual] = {};
      const codSub = fila[1] ? fila[1].toString().trim() : "";
      const nomSub = fila[2] ? fila[2].toString().trim() : "";
      if (codSub || nomSub) subprocesoActual = (codSub + " " + nomSub).trim();
      if (fila[3]) responsableActual = fila[3].toString().trim();
      
      if (!procesoActual || !subprocesoActual || !responsableActual) continue;
      
      if (!procesos[procesoActual][subprocesoActual]) procesos[procesoActual][subprocesoActual] = {};
      if (!procesos[procesoActual][subprocesoActual][responsableActual]) procesos[procesoActual][subprocesoActual][responsableActual] = {};
      
      columnasNormas.forEach((idx) => {
        const norma = encabezados[idx] ? encabezados[idx].toString().trim() : "";
        const requisito = fila[idx] ? fila[idx].toString().trim() : "";
        if (norma && requisito) {
          if (!procesos[procesoActual][subprocesoActual][responsableActual][norma]) 
            procesos[procesoActual][subprocesoActual][responsableActual][norma] = [];
          const lista = procesos[procesoActual][subprocesoActual][responsableActual][norma];
          if (!lista.includes(requisito)) lista.push(requisito);
        }
      });
    }
    return procesos;
  } catch (e) { return {}; }
}

function obtenerDescripcion(norma, numero) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Requisitos");
    if (!hoja) return null;
    const data = hoja.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == norma && data[i][1] == numero) {
        return { titulo: data[i][2], subtitulo: data[i][3], descripcion: data[i][4] };
      }
    }
    return null;
  } catch (e) { return null; }
}

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
    let procesoActual = "", subprocesoActual = "", responsableActual = "";

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      if (fila[0]) procesoActual = fila[0].toString().trim();
      const codSub = fila[1] ? fila[1].toString().trim() : "";
      const nomSub = fila[2] ? fila[2].toString().trim() : "";
      if (codSub || nomSub) subprocesoActual = (codSub + " " + nomSub).trim();
      if (fila[3]) responsableActual = fila[3].toString().trim();
      if (!procesoActual || !subprocesoActual || !responsableActual) continue;

      columnasNormasIdx.forEach((idx) => {
        if (idx >= encabezados.length) return;
        const norma = encabezados[idx] ? encabezados[idx].toString().trim() : "";
        if (idx >= fila.length) return;
        const requisito = fila[idx] ? fila[idx].toString().trim() : "";
        if (norma && requisito) {
          if (!datosInvertidos[norma]) datosInvertidos[norma] = {};
          if (!datosInvertidos[norma][requisito]) datosInvertidos[norma][requisito] = {};
          if (!datosInvertidos[norma][requisito][procesoActual]) datosInvertidos[norma][requisito][procesoActual] = {};
          if (!datosInvertidos[norma][requisito][procesoActual][subprocesoActual]) datosInvertidos[norma][requisito][procesoActual][subprocesoActual] = [];
          const listaResp = datosInvertidos[norma][requisito][procesoActual][subprocesoActual];
          if (!listaResp.includes(responsableActual)) listaResp.push(responsableActual);
        }
      });
    }
    return datosInvertidos;
  } catch (e) { return {}; }
}

/***************************************************
 * 🛠️ UTILIDADES (GETTERS Y PARSERS)
 ***************************************************/
function generarNuevoID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro Auditoría");
  if (!hoja || hoja.getLastRow() < 2) return "PA001";
  
  const data = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().flat();
  const ids = data.filter(String);
  if (ids.length === 0) return "PA001";
  
  let maxNum = 0;
  ids.forEach(idStr => {
    const num = parseInt(idStr.replace("PA", ""), 10);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  });
  return "PA" + String(maxNum + 1).padStart(3, "0");
}

function obtenerIDsAuditorias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro Auditoría");
  if (!hoja) return [];
  const raw = hoja.getRange(2, 1, hoja.getLastRow(), 1).getValues().flat();
  return [...new Set(raw.filter(String))].sort();
}

// NUEVO: Obtener solo los verdes si se requiere, o todos. Asumo todos para el selector.
function obtenerAuditores() { return obtenerPersonas("Auditores"); }
function obtenerAuditados() { return obtenerPersonas("Auditados"); }

function obtenerPersonas(hojaNombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(hojaNombre);
  if (!hoja) return [];
  const data = hoja.getRange(2, 1, hoja.getLastRow(), 2).getValues();
  return data.filter(r => r[0]).map(r => ({ nombre: r[0], correo: r[1] || "" }));
}

function obtenerProcesos() { return obtenerEstructuraSimple("Procesos"); }
function obtenerNormas() { 
  // Usa tu lógica original de filtrar por VERDE (Background #ccffcc)
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
          const valores = hoja.getRange(2, columna, lastRow - 1, 1).getValues();
          const colores = hoja.getRange(2, columna, lastRow - 1, 1).getBackgrounds();
          const verdes = [];
          for (let i = 0; i < valores.length; i++) {
            if (valores[i][0] && colores[i][0].toLowerCase() === "#ccffcc") {
              verdes.push(valores[i][0].toString().trim());
            }
          }
          normas[nombreNorma] = verdes;
        } else {
          normas[nombreNorma] = [];
        }
      }
    });
    return normas;
  } catch(e) { return {}; }
}

function obtenerEstructuraSimple(hojaNombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(hojaNombre);
  if (!hoja) return {};
  const data = hoja.getDataRange().getValues();
  const headers = data[0];
  const resultado = {};
  
  headers.forEach((h, colIdx) => {
    if (h) {
      resultado[h] = [];
      for (let r = 1; r < data.length; r++) {
        if (data[r][colIdx]) resultado[h].push(data[r][colIdx]);
      }
    }
  });
  return resultado;
}

function obtenerAuditoriaPorID(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro Auditoría");
  if (!hoja) return null;
  const data = hoja.getDataRange().getValues();
  
  const filas = data.filter(r => r[0] == id);
  if (filas.length === 0) return null;
  
  const base = filas[0];
  const resultado = {
    id: base[0],
    nombre: base[1],
    objetivos: base[2] ? base[2].split(", ") : [],
    alcance: base[3],
    modalidad: base[5],
    tipo: base[6],
    observaciones: base[17],
    horarios: []
  };

  filas.forEach(f => {
    // Parseo de personas
    const audNoms = f[11] ? f[11].split(", ") : [];
    const audMails = f[12] ? f[12].split(", ") : [];
    const auditoresObj = audNoms.map((n, i) => ({ nombre: n, correo: audMails[i] || "" }));

    const auditadosNoms = f[13] ? f[13].split(", ") : [];
    const auditadosMails = f[14] ? f[14].split(", ") : [];
    const auditadosObj = auditadosNoms.map((n, i) => ({ nombre: n, correo: auditadosMails[i] || "" }));

    // Parseo de criterios
    const parsearBloque = (texto) => {
      if (!texto) return [];
      return texto.split(" | ").map(bloque => {
        const partes = bloque.split(": ");
        return { titulo: partes[0], items: partes[1] ? partes[1].split(", ") : [] };
      });
    };

    let fechaStr = f[7];
    if (fechaStr instanceof Date) fechaStr = fechaStr.toISOString().split("T")[0];
    
    const fmtHora = (v) => {
       if (v instanceof Date) return v.getHours().toString().padStart(2,'0') + ":" + v.getMinutes().toString().padStart(2,'0');
       return v;
    };

    resultado.horarios.push({
      fecha: fechaStr,
      inicio: fmtHora(f[8]),
      fin: fmtHora(f[9]),
      descripcion: f[10],
      auditores: auditoresObj,
      auditados: auditadosObj,
      procesos: parsearBloque(f[15]).map(p => ({ proceso: p.titulo, subprocesos: p.items })),
      normas: parsearBloque(f[16]).map(n => ({ norma: n.titulo, requisitos: n.items }))
    });
  });

  return resultado;
}

// Función para el calendario visual en la web (JSON simplificado)
function obtenerAuditoriasCalendario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro Auditoría");
  if (!hoja) return [];
  const data = hoja.getDataRange().getValues();
  const output = [];
  for(let i=1; i<data.length; i++) {
      let row = data[i];
      let obj = {};
      obj["ID"] = row[0];
      obj["Nombre Auditoría"] = row[1];
      obj["Auditado"] = row[13];
      
      let fecha = row[7];
      if(fecha instanceof Date) fecha = fecha.toISOString().split('T')[0];
      let inicio = row[8];
      if(inicio instanceof Date) inicio = inicio.getHours().toString().padStart(2,'0') + ":" + inicio.getMinutes().toString().padStart(2,'0');
      let fin = row[9];
      if(fin instanceof Date) fin = fin.getHours().toString().padStart(2,'0') + ":" + fin.getMinutes().toString().padStart(2,'0');

      obj["Horarios JSON"] = JSON.stringify([{fecha: fecha, inicio: inicio, fin: fin}]);
      output.push(obj);
  }
  return output;
}