/**
 * ============================================================================
 * 📁 ARCHIVO: 01_Configuracion.gs
 * 🎯 OBJETIVO: Configuración inicial, menús y lanzamiento de Web Apps.
 * ============================================================================
 */

/**
 * 🚀 Función especial que se ejecuta al abrir la Hoja de Cálculo.
 * Crea el menú personalizado en la barra superior.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🔎 Auditorías'); // Nombre del menú principal
  
  // --------------------------------------------------------
  // 👁️ SECCIÓN: VISUALIZACIÓN
  // --------------------------------------------------------
  menu.addItem('📋 Abrir visor de requisitos', 'mostrarVisorAuditorias');
  menu.addItem('🌐 Abrir Planificador Web', 'abrirWebAuditoria');
  
  menu.addSeparator(); // Separador visual
  
  // --------------------------------------------------------
  // ⚙️ SECCIÓN: GESTIÓN DE DATA MAESTRA
  // --------------------------------------------------------
  menu.addItem('🔄 Refrescar Data (Cargar Maestra)', 'mostrarModalImportarData');
  menu.addItem('📧 Solicitar Data (Correo)', 'solicitarNuevaData');
  
  menu.addSeparator(); // Separador visual

  // --------------------------------------------------------
  // 🛠️ MANTENIMIENTO
  // --------------------------------------------------------
  menu.addItem('🔄 Consolidar requisitos (Colorear)', 'consolidarRequisitos');
  menu.addItem('🔚 Fin de ciclo (Limpiar colores)', 'finDeCiclo');
  
  menu.addToUi(); // Renderiza el menú
}

/**
 * 🌐 Función especial para servir la Web App (HTML).
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('planaudi')
    .setTitle('🗓️ Planear Auditoría Integrada')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 🖥️ Abre la Web App desde un modal.
 */
function abrirWebAuditoria() {
  const url = "https://script.google.com/macros/s/AKfycbykb4n5qbL3lyi4QHus6pCrkKMInZoatJ6UteLI-jNRHBHPBAByd3JBHkMkmooVxUyA0g/exec";
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif; text-align:center; padding:30px">
      <h2 style="color:#9e1a18;">🌐 Acceso al Planificador</h2>
      <p>Haz clic en el botón para gestionar las auditorías en pantalla completa:</p>
      <a href="${url}" target="_blank"
         style="display:inline-block; background:#9e1a18; color:white; padding:12px 24px; border-radius:6px; text-decoration:none; font-weight:bold; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        🔗 Abrir Aplicación Web
      </a>
    </div>
  `)
  .setWidth(450)
  .setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(html, '🌐 Web del Plan de Auditoría');
}

/**
 * 🔍 Abre el visor "drill-down" (Modal.html).
 */
function mostrarVisorAuditorias() {
  const html = HtmlService.createHtmlOutputFromFile('Modal')
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Visor de Auditorías');
}

/**
 * 📥 Abre el modal para ingresar el link de la Data Maestra.
 */
function mostrarModalImportarData() {
  const html = HtmlService.createHtmlOutputFromFile('ModalImportar')
    .setWidth(500)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 Refrescar Data Maestra');
}

/**
 * 🛠️ UTILIDAD: Normaliza texto.
 */
function normalizar(v) {
  if (typeof v !== "string") v = String(v || "");
  return v.trim().toLowerCase().replace(/\s+/g, " ");
}

/**
 * 🛠️ UTILIDAD: Obtiene fecha Bogotá.
 */
function obtenerFechaBogota() {
  const ahora = new Date();
  const opciones = {
    timeZone: 'America/Bogota',
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit', hour12: false
  };
  return new Intl.DateTimeFormat('es-CO', opciones).format(ahora);
}
