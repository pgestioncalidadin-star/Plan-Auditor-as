/***************************************************
 * 📘 PLAN DE AUDITORÍAS INTEGRADAS
 * Archivo: plan.gs - VERSIÓN FINAL
 ***************************************************/

// ✅ Menú principal en la hoja de cálculo
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🔎 Auditorías');
  menu.addItem('📋 Abrir visor', 'mostrarVisorAuditorias');
  menu.addItem('🌐 Abrir web Plan de Auditoría', 'abrirWebAuditoria');
  menu.addItem('🔄 Consolidar requisitos', 'consolidarRequisitos');
  menu.addItem('🔚 Fin de ciclo', 'finDeCiclo');
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
 * 🚀 Abrir el link público de la aplicación web
 ***************************************************/
function abrirWebAuditoria() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif; text-align:center; padding:30px">
      <h2>🌐 Acceso a la Web del Plan de Auditorías</h2>
      <p>Haz clic en el siguiente botón para abrir la web:</p>
      <a href="${url}" target="_blank"
         style="display:inline-block; background:#9e1a18; color:white; padding:10px 20px; border-radius:6px; text-decoration:none;">
        🔗 Abrir aplicación web
      </a>
    </div>
  `)
  .setWidth(400)
  .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '🌐 Web del Plan de Auditoría');
}

/***************************************************
 * 🔍 Mantener visor existente
 ***************************************************/
function mostrarVisorAuditorias() {
  const html = HtmlService.createHtmlOutputFromFile('Modal')
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Visor de Auditorías');
}

/***************************************************
 * 🔄 CONSOLIDAR REQUISITOS
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
      { req: 5, normas: 1 },
      { req: 6, normas: 2 },
      { req: 7, normas: 3 },
      { req: 8, normas: 4 },
      { req: 9, normas: 5 },
      { req: 10, normas: 6 },
      { req: 11, normas: 7 },
      { req: 12, normas: 8 },
      { req: 13, normas: 9 }
    ];

    const lastRowReq = hojaReq.getLastRow();
    const lastRowNormas = hojaNormas.getLastRow();

    if (lastRowReq < 2 || lastRowNormas < 2) {
      SpreadsheetApp.getUi().alert("⚠️ No hay datos suficientes para comparar.");
      return;
    }

    const COLOR_OK = "#ccffcc";
    const COLOR_NO = "#ffcccc";

    hojaNormas.getRange(2, 1, lastRowNormas - 1, 9).setBackground(null);

    mapeo.forEach(par => {
      const valoresReq = hojaReq
        .getRange(2, par.req, lastRowReq - 1, 1)
        .getValues()
        .map(r => normalizar(r[0]))
        .filter(v => v !== "");

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
    SpreadsheetApp.getUi().alert("❌ Error en consolidación: " + e.toString());
  }
}

function normalizar(v) {
  if (typeof v !== "string") v = String(v || "");
  return v.trim().toLowerCase().replace(/\s+/g, " ");
}

/***************************************************
 * 🔚 FIN DE CICLO
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
    
    if (lastRow < 2 || lastCol < 1) {
      SpreadsheetApp.getUi().alert("⚠️ No hay datos en la hoja.");
      return;
    }

    // Limpiar colores desde fila 2 hacia abajo
    hoja.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);
    
    SpreadsheetApp.getUi().alert("✅ Colores limpiados. Ciclo reiniciado correctamente.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.toString());
  }
}

/***************************************************
 * 🕐 Obtener fecha/hora en formato Bogotá
 ***************************************************/
function obtenerFechaBogota() {
  const ahora = new Date();
  const opciones = {
    timeZone: 'America/Bogota',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
  };
  const formatter = new Intl.DateTimeFormat('es-CO', opciones);
  const partes = formatter.formatToParts(ahora);
  
  const dia = partes.find(p => p.type === 'day').value;
  const mes = partes.find(p => p.type === 'month').value;
  const año = partes.find(p => p.type === 'year').value;
  const hora = partes.find(p => p.type === 'hour').value;
  const minuto = partes.find(p => p.type === 'minute').value;
  
  return `${dia}/${mes}/${año} ${hora}:${minuto}`;
}

/***************************************************
 * 📋 Obtener auditores desde hoja "Auditores"
 ***************************************************/
function obtenerAuditores() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Auditores");
    
    if (!hoja) {
      Logger.log("⚠️ Hoja 'Auditores' no encontrada");
      return [];
    }

    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return [];

    const data = hoja.getRange(2, 1, lastRow - 1, 2).getValues();
    const auditores = data
      .filter(row => row[0] && row[1])
      .map(row => ({ 
        nombre: row[0].toString().trim(), 
        correo: row[1].toString().trim() 
      }));
    
    Logger.log(`✅ Auditores cargados: ${auditores.length}`);
    return auditores;
  } catch (e) {
    Logger.log("❌ Error en obtenerAuditores: " + e.toString());
    return [];
  }
}

/***************************************************
 * 📋 Obtener auditados desde hoja "Auditados"
 ***************************************************/
function obtenerAuditados() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Auditados");
    
    if (!hoja) {
      Logger.log("⚠️ Hoja 'Auditados' no encontrada");
      return [];
    }

    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return [];

    const data = hoja.getRange(2, 1, lastRow - 1, 2).getValues();
    const auditados = data
      .filter(row => row[0] && row[1])
      .map(row => ({ 
        nombre: row[0].toString().trim(), 
        correo: row[1].toString().trim() 
      }));
    
    Logger.log(`✅ Auditados cargados: ${auditados.length}`);
    return auditados;
  } catch (e) {
    Logger.log("❌ Error en obtenerAuditados: " + e.toString());
    return [];
  }
}

/***************************************************
 * 📋 Obtener procesos desde hoja "Procesos"
 ***************************************************/
function obtenerProcesos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Procesos");
    
    if (!hoja) {
      Logger.log("⚠️ Hoja 'Procesos' no encontrada");
      return {};
    }

    const lastCol = hoja.getLastColumn();
    const encabezados = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
    const procesos = {};

    encabezados.forEach((proceso, index) => {
      if (proceso) {
        const nombreProceso = proceso.toString().trim();
        const columna = index + 1;
        const lastRow = hoja.getLastRow();
        
        if (lastRow > 1) {
          const subprocesos = hoja.getRange(2, columna, lastRow - 1, 1)
            .getValues()
            .flat()
            .filter(sub => sub)
            .map(sub => sub.toString().trim());
          
          procesos[nombreProceso] = subprocesos;
        } else {
          procesos[nombreProceso] = [];
        }
      }
    });

    Logger.log(`✅ Procesos cargados: ${Object.keys(procesos).length}`);
    return procesos;
  } catch (e) {
    Logger.log("❌ Error en obtenerProcesos: " + e.toString());
    return {};
  }
}

/***************************************************
 * 📋 Obtener normas desde hoja "Normas" (solo verdes)
 ***************************************************/
function obtenerNormas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Normas");
    
    if (!hoja) {
      Logger.log("⚠️ Hoja 'Normas' no encontrada");
      return {};
    }

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
          
          const requisitosVerdes = [];
          for (let i = 0; i < valores.length; i++) {
            const valor = valores[i][0];
            const color = colores[i][0];
            
            if (valor && color.toLowerCase() === "#ccffcc") {
              requisitosVerdes.push(valor.toString().trim());
            }
          }
          
          normas[nombreNorma] = requisitosVerdes;
        } else {
          normas[nombreNorma] = [];
        }
      }
    });

    Logger.log(`✅ Normas cargadas: ${Object.keys(normas).length}`);
    return normas;
  } catch (e) {
    Logger.log("❌ Error en obtenerNormas: " + e.toString());
    return {};
  }
}

/***************************************************
 * 📅 Obtener auditorías del calendario
 ***************************************************/
/* ========================================
   CORRECCIÓN 1: CALENDARIO WEB
   Devuelve formato plano compatible con el Frontend
======================================== */
function obtenerAuditoriasCalendario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro Auditoría");
  if (!hoja) return [];
  
  const data = hoja.getDataRange().getValues();
  const output = [];
  
  // Empezamos en 1 para saltar encabezados
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    
    // Validar que la fila tenga ID y Fecha
    if (!row[0] || !row[7]) continue;

    // Formatear Fecha para que coincida con el Frontend (YYYY-MM-DD)
    let fecha = row[7];
    if (fecha instanceof Date) {
      // Ajuste de zona horaria simple para evitar problemas de día anterior/posterior
      fecha = Utilities.formatDate(fecha, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    }

    // Formatear Horas (Inicio y Fin)
    const fmtHora = (v) => {
       if (v instanceof Date) {
         return Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), "HH:mm");
       }
       return v;
    };

    output.push({
      "id": row[0],
      "nombre": row[1],
      "fecha": fecha,        // <--- ¡ESTO ES LO QUE FALTABA!
      "inicio": fmtHora(row[8]),
      "fin": fmtHora(row[9]),
      "auditor": row[11],    // Nombres de auditores
      "auditado": row[13]    // Nombres de auditados
    });
  }
  return output;
}
/***************************************************
 * 📋 Obtener IDs de auditorías para edición
 ***************************************************/
function obtenerIDsAuditorias() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Registro Auditoría");
    
    if (!hoja || hoja.getLastRow() < 2) return [];

    const lastRow = hoja.getLastRow();
    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    
    // Obtener IDs únicos y convertir a string
    const idsUnicos = [...new Set(ids.flat().map(id => String(id)))].sort();
    
    Logger.log(`✅ IDs cargados: ${idsUnicos.length}`);
    return idsUnicos;
  } catch (e) {
    Logger.log("❌ Error en obtenerIDsAuditorias: " + e.toString());
    return [];
  }
}

/***************************************************
 * 📋 Obtener auditoría por ID para edición (CORREGIDO)
 ***************************************************/
function obtenerAuditoriaPorID(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Registro Auditoría");
    if (!hoja) return null;

    const data = hoja.getDataRange().getValues();
    const idBuscado = String(id).trim().toUpperCase();

    // Buscar todas las filas donde la columna A coincida con el ID
    const filas = data.filter(r => String(r[0]).trim().toUpperCase() === idBuscado);

    if (filas.length === 0) return null;

    const base = filas[0];

    // Objeto principal de auditoría
    const resultado = {
      id: base[0],
      nombre: base[1],
      objetivos: base[2] ? String(base[2]).split(", ") : [],
      alcance: base[3],
      modalidad: base[5],
      tipo: base[6],
      observaciones: base[17],
      horarios: []
    };

    // Función para parsear bloques "Proceso: A, B | Otro: C"
    const parsearBloque = (texto) => {
      if (!texto) return [];
      return String(texto).split(" | ").map(bloque => {
        const partes = bloque.split(": ");
        return {
          titulo: partes[0],
          items: partes[1] ? partes[1].split(", ") : []
        };
      });
    };

    const tz = ss.getSpreadsheetTimeZone();

    // Función format hora
    const fmtHora = (v) => {
      if (v instanceof Date) {
        return Utilities.formatDate(v, tz, "HH:mm");
      }
      return v;
    };

    // Recorrer TODAS las filas de esa auditoría
    filas.forEach(f => {
      // ---- Auditores ----
      const audNoms = f[11] ? String(f[11]).split(", ") : [];
      const audMails = f[12] ? String(f[12]).split(", ") : [];
      const auditoresObj = audNoms.map((n, i) => ({ nombre: n, correo: audMails[i] || "" }));

      // ---- Auditados ----
      const auditadosNoms = f[13] ? String(f[13]).split(", ") : [];
      const auditadosMails = f[14] ? String(f[14]).split(", ") : [];
      const auditadosObj = auditadosNoms.map((n, i) => ({ nombre: n, correo: auditadosMails[i] || "" }));

      // ---- Fecha ----
      let fechaStr = f[7];
      if (fechaStr instanceof Date) {
        fechaStr = Utilities.formatDate(fechaStr, tz, "yyyy-MM-dd");
      }

      // Agregar horario
      resultado.horarios.push({
        fecha: fechaStr,
        inicio: fmtHora(f[8]),
        fin: fmtHora(f[9]),
        descripcion: f[10],
        auditores: auditoresObj,
        auditados: auditadosObj,
        procesos: parsearBloque(f[15]).map(p => ({
          proceso: p.titulo,
          subprocesos: p.items
        })),
        normas: parsearBloque(f[16]).map(n => ({
          norma: n.titulo,
          requisitos: n.items
        }))
      });
    });

    return resultado;

  } catch (e) {
    Logger.log("❌ Error en obtenerAuditoriaPorID: " + e.toString());
    return null;
  }
}


/***************************************************
 * 💾 Crear nueva auditoría
 ***************************************************/
function crearAuditoria(datos) {
  try {
    Logger.log("📥 Iniciando creación de auditoría...");
    Logger.log("Datos recibidos: " + JSON.stringify(datos));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("Registro Auditoría");
    
    // Crear hoja si no existe
    if (!hoja) {
      hoja = ss.insertSheet("Registro Auditoría");
      hoja.appendRow([
        "ID",
        "Nombre Auditoría",
        "Objetivos",
        "Alcance",
        "Riesgos",
        "Auditor",
        "Correo Auditor",
        "Auditado",
        "Correo Auditado",
        "Modalidad",
        "Tipo de auditoría",
        "Fecha",
        "Hora Inicio",
        "Hora Fin",
        "Descripción",
        "Procesos",
        "Requisitos",
        "Observaciones",
        "Fecha de creación"
      ]);
      Logger.log("✅ Hoja 'Registro Auditoría' creada");
    }

    // … (resto del código lo dejé EXACTAMENTE igual)


    // Generar ID
    const nuevoID = generarNuevoID();
    Logger.log(`✅ ID generado: ${nuevoID}`);
    
    // Validaciones
    if (!datos.nombre) throw new Error("Nombre de auditoría requerido");
    if (!datos.horarios || datos.horarios.length === 0) throw new Error("Debe agregar al menos un horario");

    // Obtener fecha/hora de Bogotá
    const fechaCreacion = obtenerFechaBogota();

    // Procesar cada horario como una fila independiente
    datos.horarios.forEach((horario, index) => {
      // Procesar procesos
      let procesosTexto = "";
      if (horario.procesos && horario.procesos.length > 0) {
        procesosTexto = horario.procesos.map(p => 
          `${p.proceso}: ${p.subprocesos.join(', ')}`
        ).join(' | ');
      }
      
      // Procesar requisitos
      let requisitosTexto = "";
      if (horario.normas && horario.normas.length > 0) {
        requisitosTexto = horario.normas.map(n => 
          `${n.norma}: ${n.requisitos.join(', ')}`
        ).join(' | ');
      }
      
      const fila = [
        nuevoID,
        datos.nombre || "",
        Array.isArray(datos.objetivos) ? datos.objetivos.join(", ") : "",
        datos.alcance || "",
        "Ver los riesgos en el documento 'Programa de auditorías'",
        Array.isArray(horario.auditores) ? horario.auditores.join(", ") : "",
        Array.isArray(horario.correosAuditores) ? horario.correosAuditores.join(", ") : "",
        Array.isArray(horario.auditados) ? horario.auditados.join(", ") : "",
        Array.isArray(horario.correosAuditados) ? horario.correosAuditados.join(", ") : "",
        datos.modalidad || "",
        datos.tipo || "",
        horario.fecha || "",
        horario.inicio || "",
        horario.fin || "",
        horario.descripcion || "",
        procesosTexto,
        requisitosTexto,
        datos.observaciones || "",
        fechaCreacion
      ];

      hoja.appendRow(fila);
      Logger.log(`✅ Fila ${index + 1} agregada para horario del día ${horario.fecha}`);
    });

    // Marcar requisitos en verde
    try {
      marcarRequisitosAuditados(datos);
      Logger.log("✅ Requisitos marcados en verde");
    } catch (e) {
      Logger.log("⚠️ Error marcando requisitos: " + e.toString());
    }

    // Crear eventos en calendario
    try {
      crearEventosCalendario(datos, nuevoID);
      Logger.log("✅ Eventos de calendario creados");
    } catch (e) {
      Logger.log("⚠️ Error en calendario: " + e.toString());
    }

    // Enviar correos a TODOS los auditores y auditados
    try {
      enviarCorreosAuditoria(datos, nuevoID);
      Logger.log("✅ Correos enviados");
    } catch (e) {
      Logger.log("⚠️ Error enviando correos: " + e.toString());
    }

    Logger.log("🎉 Auditoría creada exitosamente: " + nuevoID);
    return { success: true, id: nuevoID };
    
  } catch (error) {
    Logger.log("❌ Error en crearAuditoria: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/***************************************************
 * ✏️ EDITAR AUDITORÍA EXISTENTE
 ***************************************************/
function editarAuditoria(datos) {
  try {
    Logger.log("📝 Iniciando edición de auditoría: " + datos.id);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Registro Auditoría");
    
    if (!hoja) throw new Error("Hoja 'Registro Auditoría' no encontrada");

    // Convertir ID a string para comparación consistente
    const idBuscado = String(datos.id).trim();

    // Eliminar filas anteriores con este ID
    const lastRow = hoja.getLastRow();
    for (let i = lastRow; i >= 2; i--) {
      const idFila = String(hoja.getRange(i, 1).getValue()).trim();
      if (idFila === idBuscado) {
        hoja.deleteRow(i);
        Logger.log(`✅ Fila ${i} eliminada (ID: ${idFila})`);
      }
    }

    Logger.log("✅ Filas anteriores eliminadas");

    // Obtener fecha/hora de Bogotá
    const fechaActualizacion = obtenerFechaBogota();

    // Insertar nuevas filas (reutilizar lógica de crearAuditoria)
    datos.horarios.forEach((horario, index) => {
      let procesosTexto = "";
      if (horario.procesos && horario.procesos.length > 0) {
        procesosTexto = horario.procesos.map(p => 
          `${p.proceso}: ${p.subprocesos.join(', ')}`
        ).join(' | ');
      }
      
      let requisitosTexto = "";
      if (horario.normas && horario.normas.length > 0) {
        requisitosTexto = horario.normas.map(n => 
          `${n.norma}: ${n.requisitos.join(', ')}`
        ).join(' | ');
      }
      
      const fila = [
        datos.id,
        datos.nombre || "",
        Array.isArray(datos.objetivos) ? datos.objetivos.join(", ") : "",
        datos.alcance || "",
        "Ver los riesgos en el documento 'Programa de auditorías'",
        Array.isArray(horario.auditores) ? horario.auditores.join(", ") : "",
        Array.isArray(horario.correosAuditores) ? horario.correosAuditores.join(", ") : "",
        Array.isArray(horario.auditados) ? horario.auditados.join(", ") : "",
        Array.isArray(horario.correosAuditados) ? horario.correosAuditados.join(", ") : "",
        datos.modalidad || "",
        datos.tipo || "",
        horario.fecha || "",
        horario.inicio || "",
        horario.fin || "",
        horario.descripcion || "",
        procesosTexto,
        requisitosTexto,
        datos.observaciones || "",
        fechaActualizacion
      ];

      hoja.appendRow(fila);
    });

    Logger.log("✅ Nuevas filas insertadas");

    // Marcar requisitos
    try {
      marcarRequisitosAuditados(datos);
    } catch (e) {
      Logger.log("⚠️ Error marcando requisitos: " + e.toString());
    }

    // Crear eventos calendario
    try {
      crearEventosCalendario(datos, datos.id);
    } catch (e) {
      Logger.log("⚠️ Error en calendario: " + e.toString());
    }

    // Enviar correos
    try {
      enviarCorreosAuditoria(datos, datos.id);
} catch (e) {
Logger.log("⚠️ Error enviando correos: " + e.toString());
}
Logger.log("🎉 Auditoría editada exitosamente: " + datos.id);
return { success: true, id: datos.id };
} catch (error) {
Logger.log("❌ Error en editarAuditoria: " + error.toString());
return { success: false, error: error.toString() };
}
}
/***************************************************

🎨 MARCAR REQUISITOS AUDITADOS EN VERDE
***************************************************/
function marcarRequisitosAuditados(datos) {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const hoja = ss.getSheetByName("Requisitos faltantes por auditar");

if (!hoja) {
Logger.log("⚠️ Hoja 'Requisitos faltantes por auditar' no encontrada");
return;
}
const lastCol = hoja.getLastColumn();
const lastRow = hoja.getLastRow();
const encabezados = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
// Recopilar todos los requisitos seleccionados
const requisitosSeleccionados = new Set();
datos.horarios.forEach(h => {
if (h.normas && h.normas.length > 0) {
h.normas.forEach(n => {
n.requisitos.forEach(req => {
requisitosSeleccionados.add(`${n.norma}::${req.trim()}`);
});
});
}
});
// Marcar en verde
encabezados.forEach((norma, colIndex) => {
if (!norma) return;
const nombreNorma = norma.toString().trim();
const columna = colIndex + 1;
if (lastRow > 1) {
const valores = hoja.getRange(2, columna, lastRow - 1, 1).getValues();
  valores.forEach((fila, filaIndex) => {
    const requisito = fila[0] ? fila[0].toString().trim() : "";
    if (requisito) {
      const clave = `${nombreNorma}::${requisito}`;
      if (requisitosSeleccionados.has(clave)) {
        hoja.getRange(filaIndex + 2, columna).setBackground("#ccffcc");
      }
    }
  });
}
});
Logger.log("✅ Requisitos marcados en verde correctamente");
}
/***************************************************

🔢 Generar nuevo ID
***************************************************/
function generarNuevoID() {
try {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const hoja = ss.getSheetByName("Registro Auditoría");
if (!hoja || hoja.getLastRow() < 2) {
return "PA001";
}
const ultimaFila = hoja.getLastRow();
const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
const numero = parseInt(ultimoID.replace("PA", ""), 10);
const nuevoNumero = numero + 1;
return "PA" + String(nuevoNumero).padStart(3, "0");
} catch (e) {
Logger.log("⚠️ Error generando ID, usando PA001: " + e.toString());
return "PA001";
}
}

/***************************************************

📅 Crear eventos en Google Calendar (CORREGIDO)
***************************************************/
function crearEventosCalendario(datos, idAuditoria) {
if (!datos.horarios || datos.horarios.length === 0) {
Logger.log("⚠️ No se pueden crear eventos: faltan horarios");
return;
}

// Agrupar horarios por fecha
const horariosPorFecha = {};
datos.horarios.forEach(h => {
if (!horariosPorFecha[h.fecha]) {
horariosPorFecha[h.fecha] = [];
}
horariosPorFecha[h.fecha].push(h);
});
// Crear eventos por fecha única
const fechasOrdenadas = Object.keys(horariosPorFecha).sort();
fechasOrdenadas.forEach((fecha, diaIndex) => {
const horariosDelDia = horariosPorFecha[fecha];
const numeroDia = diaIndex + 1;
const titulo = `🗓️ ${datos.nombre || 'Auditoría'} ${idAuditoria} - Día ${numeroDia}`;
const descripcion = construirDescripcionEventoAgrupado(datos, fecha, horariosDelDia, numeroDia, idAuditoria);
try {
const primerHorario = horariosDelDia[0];
const ultimoHorario = horariosDelDia[horariosDelDia.length - 1];
  const fechaInicio = new Date(fecha + 'T' + primerHorario.inicio);
  const fechaFin = new Date(fecha + 'T' + ultimoHorario.fin);

  // Recopilar todos los correos únicos
  const todosLosCorreos = new Set();
  horariosDelDia.forEach(h => {
    if (h.correosAuditores) h.correosAuditores.forEach(c => todosLosCorreos.add(c));
    if (h.correosAuditados) h.correosAuditados.forEach(c => todosLosCorreos.add(c));
  });

  const invitados = Array.from(todosLosCorreos).join(",");

  CalendarApp.getDefaultCalendar().createEvent(titulo, fechaInicio, fechaFin, {
    description: descripcion,
    guests: invitados,
    sendInvites: true
  });

  Logger.log(`✅ Evento creado: ${titulo} - ${fecha}`);
} catch (e) {
  Logger.log("❌ Error creando evento: " + e.toString());
}
});
}
function construirDescripcionEventoAgrupado(datos, fecha, horariosDelDia, numeroDia, idAuditoria) {
  let desc = `📋 Auditoría: ${datos.nombre || ''}  
🆔 ID: ${idAuditoria}  
📅 Día ${numeroDia} - Fecha: ${fecha}  
🎯 Objetivos: ${Array.isArray(datos.objetivos) ? datos.objetivos.join(', ') : ''}  
🧭 Alcance: ${datos.alcance || ''}  
⏰ HORARIOS DEL DÍA:`.trim();

  horariosDelDia.forEach((h, idx) => {
    desc += `\n\n${idx + 1}. ${h.inicio} - ${h.fin}`;
    
    if (h.descripcion) {
      desc += ` | ${h.descripcion}`;
    }
  });

  return desc;
}

function construirDescripcionEventoAgrupado(datos, fecha, horariosDelDia, numeroDia, idAuditoria) {

  let desc = `📋 Auditoría: ${datos.nombre || ''}  
🆔 ID: ${idAuditoria}  
📅 Día ${numeroDia} - Fecha: ${fecha}  
🎯 Objetivos: ${Array.isArray(datos.objetivos) ? datos.objetivos.join(', ') : ''}  
🧭 Alcance: ${datos.alcance || ''}  
⏰ HORARIOS DEL DÍA:`.trim();

  horariosDelDia.forEach((h, idx) => {

    desc += `\n\n${idx + 1}. ${h.inicio} - ${h.fin}`;

    // Descripción
    if (h.descripcion) {
      desc += ` | ${h.descripcion}`;
    }

    // Auditores
    if (h.auditores && h.auditores.length > 0) {
      desc += `\n   👤 Auditores: ${h.auditores.join(', ')}`;
    }

    // Auditados
    if (h.auditados && h.auditados.length > 0) {
      desc += `\n   👥 Auditados: ${h.auditados.join(', ')}`;
    }

    // Procesos
    if (h.procesos && h.procesos.length > 0) {
      desc += `\n   🧩 Procesos:`;
      h.procesos.forEach(p => {
        desc += `\n      • ${p.proceso}: ${p.subprocesos.join(', ')}`;
      });
    }

    // Normas
    if (h.normas && h.normas.length > 0) {
      desc += `\n   📋 Normas:`;
      h.normas.forEach(n => {
        desc += `\n      • ${n.norma}: ${n.requisitos.join(', ')}`;
      });
    }

  });

  return desc;
}



/***************************************************

📧 Enviar correos a TODOS (auditores y auditados)
***************************************************/
function enviarCorreosAuditoria(datos, idAuditoria) {
try {
const datosAgrupados = agruparHorariosPorFecha(datos);
// Recopilar todos los correos únicos
const correosAuditores = new Set();
const correosAuditados = new Set();
datos.horarios.forEach(h => {
if (h.correosAuditores) h.correosAuditores.forEach(c => correosAuditores.add(c));
if (h.correosAuditados) h.correosAuditados.forEach(c => correosAuditados.add(c));
});
// Enviar a cada auditor
correosAuditores.forEach(correo => {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('correo_auditoria');
    htmlTemplate.datos = datosAgrupados;
    htmlTemplate.idAuditoria = idAuditoria;

    const mensaje = htmlTemplate.evaluate().getContent();

    MailApp.sendEmail({
      to: correo,
      subject: `🗓️ Nueva Auditoría Asignada - ${idAuditoria}`,
      htmlBody: mensaje
    });

    Logger.log(`✅ Correo enviado a auditor: ${correo}`);

  } catch (e) {
    Logger.log(`❌ Error enviando a auditor ${correo}: ${e.toString()}`);
  }
});

// Enviar a cada auditado
correosAuditados.forEach(correo => {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('auditados');
    htmlTemplate.datos = datosAgrupados;
    htmlTemplate.idAuditoria = idAuditoria;

    const mensaje = htmlTemplate.evaluate().getContent();

    MailApp.sendEmail({
      to: correo,
      subject: `📋 Notificación de Auditoría Programada - ${idAuditoria}`,
      htmlBody: mensaje
    });

    Logger.log(`✅ Correo enviado a auditado: ${correo}`);

  } catch (e) {
    Logger.log(`❌ Error enviando a auditado ${correo}: ${e.toString()}`);
  }
});

} catch (e) {
  Logger.log("❌ Error en enviarCorreosAuditoria: " + e.toString());
  throw e;
}
}

/***************************************************

🔧 FUNCIÓN AUXILIAR: Agrupar horarios por fecha
***************************************************/
function agruparHorariosPorFecha(datos) {
const horariosPorFecha = {};

datos.horarios.forEach(h => {
if (!horariosPorFecha[h.fecha]) {
horariosPorFecha[h.fecha] = [];
}
horariosPorFecha[h.fecha].push(h);
});
const horariosAgrupados = [];
const fechasOrdenadas = Object.keys(horariosPorFecha).sort();
fechasOrdenadas.forEach((fecha, diaIndex) => {
horariosAgrupados.push({
numeroDia: diaIndex + 1,
fecha: fecha,
horarios: horariosPorFecha[fecha]
});
});
return {
...datos,
horariosAgrupados: horariosAgrupados,
horarios: datos.horarios
};
}
/***************************************************

🔍 Funciones para el VISOR (Modal.html)
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
let procesoActual = "";
let subprocesoActual = "";
let responsableActual = "";
const columnasNormas = [4, 5, 6, 7, 8, 9, 11, 12, 13, 14];
for (let i = 1; i < data.length; i++) {
const fila = data[i];
if (fila[0]) {
procesoActual = fila[0].toString().trim();
if (!procesos[procesoActual]) procesos[procesoActual] = {};
}
const codSub = fila[1] ? fila[1].toString().trim() : "";
const nomSub = fila[2] ? fila[2].toString().trim() : "";
if (codSub || nomSub) {
subprocesoActual = (codSub + " " + nomSub).trim();
}
if (fila[3]) {
responsableActual = fila[3].toString().trim();
}
if (!procesoActual || !subprocesoActual || !responsableActual) continue;
if (!procesos[procesoActual][subprocesoActual])
procesos[procesoActual][subprocesoActual] = {};
if (!procesos[procesoActual][subprocesoActual][responsableActual])
procesos[procesoActual][subprocesoActual][responsableActual] = {};
columnasNormas.forEach((idx) => {
const norma = encabezados[idx] ? encabezados[idx].toString().trim() : "";
const requisito = fila[idx] ? fila[idx].toString().trim() : "";
if (norma && requisito) {
if (!procesos[procesoActual][subprocesoActual][responsableActual][norma]) {
procesos[procesoActual][subprocesoActual][responsableActual][norma] = [];
}
const lista = procesos[procesoActual][subprocesoActual][responsableActual][norma];
if (!lista.includes(requisito)) lista.push(requisito);
}
});
}
return procesos;
} catch (e) {
Logger.log("❌ Error en obtenerDatos: " + e.toString());
return {};
}
}

function obtenerDescripcion(norma, numero) {
try {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const hoja = ss.getSheetByName("Requisitos");
if (!hoja) return null;
const data = hoja.getDataRange().getValues();
for (let i = 1; i < data.length; i++) {
if (data[i][0] == norma && data[i][1] == numero) {
return {
titulo: data[i][2],
subtitulo: data[i][3],
descripcion: data[i][4],
};
}
}
return null;
} catch (e) {
Logger.log("❌ Error en obtenerDescripcion: " + e.toString());
return null;
}
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
let procesoActual = "";
let subprocesoActual = "";
let responsableActual = "";

for (let i = 1; i < data.length; i++) {
  const fila = data[i];

  if (fila[0]) procesoActual = fila[0].toString().trim();

  const codSub = fila[1] ? fila[1].toString().trim() : "";
  const nomSub = fila[2] ? fila[2].toString().trim() : "";
  if (codSub || nomSub) {
    subprocesoActual = (codSub + " " + nomSub).trim();
  }

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
      if (!datosInvertidos[norma][requisito][procesoActual]) {
        datosInvertidos[norma][requisito][procesoActual] = {};
      }
      if (!datosInvertidos[norma][requisito][procesoActual][subprocesoActual]) {
        datosInvertidos[norma][requisito][procesoActual][subprocesoActual] = [];
      }

      const listaResp = datosInvertidos[norma][requisito][procesoActual][subprocesoActual];
      if (!listaResp.includes(responsableActual)) {
        listaResp.push(responsableActual);
      }
    }
  });
}
return datosInvertidos;
} catch (e) {
Logger.log("❌ Error en obtenerDatosReversa: " + e.toString());
return {};
}
}