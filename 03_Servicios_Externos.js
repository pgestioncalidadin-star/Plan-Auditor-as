/**
 * ============================================================================
 * 📁 ARCHIVO: 03_Servicios_Externos.gs
 * 🎯 OBJETIVO: Integración con Google Calendar y Gmail.
 * ============================================================================
 */

function solicitarNuevaData() {
  const destinatario = "ymarin@noel.com.co";
  const asunto = "Solicitud actualización Data Maestra - SIG";
  const cuerpo = `Cordial saludo,\n\nPor medio del presente solicitamos amablemente la data maestra actualizada correspondiente al mes en curso.\n\nAtentamente,\nEquipo SIG`;
  try { MailApp.sendEmail({ to: destinatario, subject: asunto, body: cuerpo }); SpreadsheetApp.getUi().alert(`✅ Correo enviado a ${destinatario}`); } 
  catch (e) { SpreadsheetApp.getUi().alert(`❌ Error enviando correo: ${e.toString()}`); }
}

function crearEventosCalendario(datos, idAuditoria) {
  if (!datos.horarios || datos.horarios.length === 0) return;

  const calendar = CalendarApp.getDefaultCalendar();
  
  // 1. LIMPIEZA: Buscar y eliminar eventos previos de esta auditoría para evitar duplicados
  const now = new Date();
  // Buscamos en un rango amplio (6 meses atrás a 1 año adelante)
  const startSearch = new Date(now.getTime() - (180 * 24 * 60 * 60 * 1000));
  const endSearch = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000));
  
  try {
    // Busca eventos que contengan el ID (ej: PA001)
    const eventosAntiguos = calendar.getEvents(startSearch, endSearch, { search: idAuditoria });
    eventosAntiguos.forEach(evento => {
      // Verificación extra: que el título contenga el ID
      if (evento.getTitle().includes(idAuditoria)) {
        evento.deleteEvent(); // Esto borra el evento y notifica la cancelación si aplica
      }
    });
  } catch (e) {
    Logger.log("⚠️ No se pudieron limpiar eventos antiguos: " + e.toString());
  }

  // 2. CREACIÓN: Generar los nuevos eventos con las nuevas fechas
  const horariosPorFecha = {};
  datos.horarios.forEach(h => { if (!horariosPorFecha[h.fecha]) horariosPorFecha[h.fecha] = []; horariosPorFecha[h.fecha].push(h); });

  Object.keys(horariosPorFecha).sort().forEach((fecha, idx) => {
    const horariosDelDia = horariosPorFecha[fecha];
    const numeroDia = idx + 1;
    const titulo = `🗓️ ${datos.nombre || 'Auditoría'} ${idAuditoria} - Día ${numeroDia}`;
    const descripcion = construirDescripcionEvento(datos, fecha, horariosDelDia, numeroDia, idAuditoria);
    try {
      const primerH = horariosDelDia[0];
      const ultimoH = horariosDelDia[horariosDelDia.length - 1];
      const fechaInicio = new Date(fecha + 'T' + primerH.inicio);
      const fechaFin = new Date(fecha + 'T' + ultimoH.fin);
      const invitadosSet = new Set();
      horariosDelDia.forEach(h => {
        if (h.correosAuditores) h.correosAuditores.forEach(c => invitadosSet.add(c));
        if (h.correosAuditados) h.correosAuditados.forEach(c => invitadosSet.add(c));
      });
      
      // Crear evento con invitados (enviará invitación actualizada)
      CalendarApp.getDefaultCalendar().createEvent(titulo, fechaInicio, fechaFin, { 
        description: descripcion, 
        guests: Array.from(invitadosSet).join(","), 
        sendInvites: true 
      });
    } catch (e) { Logger.log("❌ Error Calendar: " + e.toString()); }
  });
}

function construirDescripcionEvento(datos, fecha, horariosDelDia, numeroDia, idAuditoria) {
  let desc = `📋 Auditoría: ${datos.nombre || ''}\n🆔 ID: ${idAuditoria}\n🎯 Objetivos: ${Array.isArray(datos.objetivos) ? datos.objetivos.join(', ') : ''}\n🧭 Alcance: ${datos.alcance || ''}\n`;
  horariosDelDia.forEach((h, idx) => {
    desc += `\n🔹 Bloque ${idx + 1}: ${h.inicio} - ${h.fin}`;
    if (h.descripcion) desc += ` | ${h.descripcion}`;
    
    const auditoresNoms = (h.auditores || []).map(a => a.nombre || a).join(', ');
    const auditadosNoms = (h.auditados || []).map(a => a.nombre || a).join(', ');
    if (auditoresNoms) desc += `\n   👤 Auditores: ${auditoresNoms}`;
    if (auditadosNoms) desc += `\n   👥 Auditados: ${auditadosNoms}`;

    if (h.procesos && h.procesos.length) {
      desc += `\n   🧩 Procesos:`;
      h.procesos.forEach(p => {
        desc += `\n   ${p.proceso}:`;
        p.subprocesos.forEach(sub => desc += `\n   • ${sub}`);
      });
    }
    if (h.normas && h.normas.length) {
      desc += `\n   📋 Normas:`;
      h.normas.forEach(n => {
        desc += `\n   ${n.norma}:`;
        n.requisitos.forEach(req => desc += `\n   • ${req}`);
      });
    }
  });
  return desc;
}

function enviarCorreosAuditoria(datos, idAuditoria) {
  try {
    const datosAgrupados = agruparHorariosPorFecha(datos);
    const correosAuditores = new Set();
    const correosAuditados = new Set();
    datos.horarios.forEach(h => {
      if (h.correosAuditores) h.correosAuditores.forEach(c => correosAuditores.add(c));
      if (h.correosAuditados) h.correosAuditados.forEach(c => correosAuditados.add(c));
    });
    correosAuditores.forEach(correo => { if(correo) enviarEmailConPlantilla(correo, 'correo_auditoria', `🗓️ Nueva Auditoría Asignada - ${idAuditoria}`, datosAgrupados, idAuditoria); });
    correosAuditados.forEach(correo => { if(correo) enviarEmailConPlantilla(correo, 'auditados', `📋 Notificación de Auditoría Programada - ${idAuditoria}`, datosAgrupados, idAuditoria); });
  } catch (e) { Logger.log("❌ Error enviando correos: " + e.toString()); }
}

function enviarEmailConPlantilla(destinatario, nombreArchivoHtml, asunto, datos, id) {
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile(nombreArchivoHtml);
    htmlTemplate.datos = datos; htmlTemplate.idAuditoria = id;
    MailApp.sendEmail({ to: destinatario, subject: asunto, htmlBody: htmlTemplate.evaluate().getContent() });
  } catch (e) { Logger.log(`⚠️ Fallo envío a ${destinatario}: ${e}`); }
}

function agruparHorariosPorFecha(datos) {
  const horariosPorFecha = {};
  datos.horarios.forEach(h => { if (!horariosPorFecha[h.fecha]) horariosPorFecha[h.fecha] = []; horariosPorFecha[h.fecha].push(h); });
  const horariosAgrupados = [];
  Object.keys(horariosPorFecha).sort().forEach((fecha, index) => {
    horariosAgrupados.push({ numeroDia: index + 1, fecha: fecha, horarios: horariosPorFecha[fecha] });
  });
  return { ...datos, horariosAgrupados: horariosAgrupados };
}
