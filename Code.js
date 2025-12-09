/***************************************************
 * 📘 SISTEMA DE PLANEACIÓN DE AUDITORÍAS
 * Backend mejorado con manejo robusto de errores
 ***************************************************/

/* ========================================
   MENÚ Y CONFIGURACIÓN INICIAL
======================================== */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Auditoría")
      .addItem("Abrir visor", "mostrarModal")
      .addItem("🌐 Abrir web Plan de Auditoría", "abrirWebAuditoria")
      .addSeparator()
      .addItem("Consolida requisitos", "consolidarRequisitos")
      .addItem("Eliminar duplicados Auditados", "eliminarDuplicadosAuditados")
      .addToUi();
  } catch (e) {
    Logger.log("❌ Error en onOpen: " + e.toString());
  }
}

/* ========================================
   FUNCIÓN WEB - doGet()
======================================== */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('planaudi')
    .setTitle('🗓️ Planear Auditoría Integrada')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ========================================
   ABRIR WEB DE AUDITORÍA
======================================== */
function abrirWebAuditoria() {
  try {
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
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error al abrir web: " + e.toString());
  }
}

/* ========================================
   MOSTRAR VISOR (MODAL)
======================================== */
function mostrarModal() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Modal')
      .setWidth(1000)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, '🔍 Visor de Auditorías');
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error al abrir visor: " + e.toString());
  }
}

/* ========================================
   OBTENER AUDITORES
======================================== */
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

/* ========================================
   OBTENER AUDITADOS
======================================== */
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

/* ========================================
   OBTENER PROCESOS (BD Procesos)
======================================== */
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

/* ========================================
   OBTENER NORMAS Y REQUISITOS (BD Normas)
   Solo devuelve requisitos en verde (#ccffcc)
======================================== */
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
            
            // Solo incluir si tiene valor y es verde claro
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

/* ========================================
   OBTENER AUDITORÍAS DEL CALENDARIO
======================================== */
function obtenerAuditoriasCalendario() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Registro Auditoría");
    
    if (!hoja || hoja.getLastRow() < 2) return [];

    const lastRow = hoja.getLastRow();
    const lastCol = hoja.getLastColumn();
    const encabezados = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
    const datos = hoja.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const auditorias = datos.map(fila => {
      const obj = {};
      encabezados.forEach((header, idx) => {
        obj[header] = fila[idx];
      });
      return obj;
    });

    Logger.log(`✅ Auditorías cargadas: ${auditorias.length}`);
    return auditorias;
  } catch (e) {
    Logger.log("❌ Error en obtenerAuditoriasCalendario: " + e.toString());
    return [];
  }
}

/* ========================================
   CREAR NUEVA AUDITORÍA
======================================== */
function crearAuditoria(datos) {
  try {
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
        "Auditor",
        "Correo Auditor",
        "Auditado",
        "Correo Auditado",
        "Modalidad",
        "Tipo",
        "Horarios JSON",
        "Observaciones",
        "Fecha Creación"
      ]);
    }

    // Generar ID
    const nuevoID = generarNuevoID();
    
    // Validaciones
    if (!datos.nombre) throw new Error("Nombre de auditoría requerido");
    if (!datos.auditorNombre || !datos.auditorCorreo) throw new Error("Auditor requerido");
    if (!datos.auditadoNombre || !datos.auditadoCorreo) throw new Error("Auditado requerido");
    if (!datos.horarios || datos.horarios.length === 0) throw new Error("Debe agregar al menos un horario");

    // Preparar fila
    const fila = [
      nuevoID,
      datos.nombre || "",
      Array.isArray(datos.objetivos) ? datos.objetivos.join(", ") : "",
      datos.alcance || "",
      datos.auditorNombre || "",
      datos.auditorCorreo || "",
      datos.auditadoNombre || "",
      datos.auditadoCorreo || "",
      datos.modalidad || "",
      datos.tipo || "",
      JSON.stringify(datos.horarios || []),
      datos.observaciones || "",
      new Date().toISOString()
    ];

    // Guardar en hoja
    hoja.appendRow(fila);
    Logger.log(`✅ Auditoría creada: ${nuevoID}`);

    // Crear eventos en calendario
    try {
      crearEventosCalendario(datos, nuevoID);
    } catch (e) {
      Logger.log("⚠️ Error en calendario: " + e.toString());
    }

    // Enviar correos
    try {
      enviarCorreoAuditor(datos, nuevoID);
      enviarCorreoAuditado(datos, nuevoID);
    } catch (e) {
      Logger.log("⚠️ Error enviando correos: " + e.toString());
    }

    return { success: true, id: nuevoID };
    
  } catch (error) {
    Logger.log("❌ Error en crearAuditoria: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/* ========================================
   GENERAR NUEVO ID
======================================== */
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
    Logger.log("⚠️ Error generando ID, usando PA001");
    return "PA001";
  }
}

/* ========================================
   CREAR EVENTOS EN CALENDARIO
======================================== */
function crearEventosCalendario(datos, idAuditoria) {
  if (!datos.auditorCorreo || !datos.horarios || datos.horarios.length === 0) return;

  datos.horarios.forEach((horario, index) => {
    const titulo = `🗓️ ${datos.nombre || 'Auditoría'} ${idAuditoria} - Día ${index + 1}`;
    const descripcion = construirDescripcionEvento(datos, horario, idAuditoria);
    
    try {
      const fechaInicio = new Date(horario.fecha + 'T' + horario.inicio);
      const fechaFin = new Date(horario.fecha + 'T' + horario.fin);

      CalendarApp.getDefaultCalendar().createEvent(titulo, fechaInicio, fechaFin, {
        description: descripcion,
        guests: datos.auditorCorreo + "," + datos.auditadoCorreo,
        sendInvites: true
      });
      Logger.log(`✅ Evento creado: ${titulo}`);
    } catch (e) {
      Logger.log("❌ Error creando evento: " + e.toString());
    }
  });
}

function construirDescripcionEvento(datos, horario, idAuditoria) {
  let desc = `
📋 Auditoría: ${datos.nombre || ''}
🆔 ID: ${idAuditoria}
👤 Auditor: ${datos.auditorNombre || ''}
👥 Auditado: ${datos.auditadoNombre || ''}
📅 Fecha: ${horario.fecha}
⏰ Horario: ${horario.inicio} - ${horario.fin}
📝 Descripción: ${horario.descripcion || 'Sin descripción'}
🎯 Objetivos: ${Array.isArray(datos.objetivos) ? datos.objetivos.join(', ') : ''}
🧭 Alcance: ${datos.alcance || ''}
  `.trim();

  // Agregar procesos y subprocesos
  if (horario.procesos && horario.procesos.length > 0) {
    desc += "\n\n🧩 Procesos y Subprocesos:";
    horario.procesos.forEach(p => {
      desc += `\n• ${p.proceso}: ${p.subprocesos.join(', ')}`;
    });
  }

  // Agregar normas y requisitos
  if (horario.normas && horario.normas.length > 0) {
    desc += "\n\n📋 Normas y Requisitos:";
    horario.normas.forEach(n => {
      desc += `\n• ${n.norma}: ${n.requisitos.join(', ')}`;
    });
  }

  return desc;
}

/* ========================================
   ENVIAR CORREO AL AUDITOR
======================================== */
function enviarCorreoAuditor(datos, idAuditoria) {
  if (!datos.auditorCorreo) {
    Logger.log("⚠️ No hay correo de auditor");
    return;
  }

  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('correo_auditoria');
    htmlTemplate.datos = datos;
    htmlTemplate.idAuditoria = idAuditoria;
    
    const mensaje = htmlTemplate.evaluate().getContent();
    
    MailApp.sendEmail({
      to: datos.auditorCorreo,
      subject: `🗓️ Nueva Auditoría Asignada - ${idAuditoria}`,
      htmlBody: mensaje
    });
    
    Logger.log("✅ Correo enviado a auditor: " + datos.auditorCorreo);
  } catch (e) {
    Logger.log("❌ Error enviando correo auditor: " + e.toString());
    throw e;
  }
}

/* ========================================
   ENVIAR CORREO AL AUDITADO
======================================== */
function enviarCorreoAuditado(datos, idAuditoria) {
  if (!datos.auditadoCorreo) {
    Logger.log("⚠️ No hay correo de auditado");
    return;
  }

  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('auditados');
    htmlTemplate.datos = datos;
    htmlTemplate.idAuditoria = idAuditoria;
    
    const mensaje = htmlTemplate.evaluate().getContent();
    
    MailApp.sendEmail({
      to: datos.auditadoCorreo,
      subject: `📋 Notificación de Auditoría Programada - ${idAuditoria}`,
      htmlBody: mensaje
    });
    
    Logger.log("✅ Correo enviado a auditado: " + datos.auditadoCorreo);
  } catch (e) {
    Logger.log("❌ Error enviando correo auditado: " + e.toString());
    throw e;
  }
}

/* ========================================
   CONSOLIDAR REQUISITOS
======================================== */
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

/* ========================================
   ELIMINAR DUPLICADOS EN AUDITADOS
======================================== */
function eliminarDuplicadosAuditados() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Auditados");
    
    if (!hoja) {
      SpreadsheetApp.getUi().alert("No se encontró la hoja 'Auditados'.");
      return;
    }

    const lastRow = hoja.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("No hay datos en la hoja.");
      return;
    }

    const lastCol = hoja.getLastColumn();
    const datos = hoja.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const vistos = new Set();
    const datosFiltrados = [];

    datos.forEach(fila => {
      const nombre = (fila[0] || "").toString().trim().toUpperCase();
      if (!vistos.has(nombre)) {
        vistos.add(nombre);
        datosFiltrados.push(fila);
      }
    });

    hoja.getRange(2, 1, lastRow - 1, lastCol).clearContent();

    if (datosFiltrados.length > 0) {
      hoja.getRange(2, 1, datosFiltrados.length, lastCol).setValues(datosFiltrados);
    }

    SpreadsheetApp.getUi().alert(`✅ Duplicados eliminados: ${datosFiltrados.length} registros únicos.`);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.toString());
  }
}

/* ========================================
   FUNCIONES PARA EL VISOR (MODAL)
======================================== */
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