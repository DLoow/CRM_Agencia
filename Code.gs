// ============================================================
// CRM AGENCIA DE VIAJES - Code.gs
// ============================================================

var CONFIG = {
  SHEETS: {
    CLIENTES:    'Clientes',
    DESTINOS:    'Destinos',
    COTIZACIONES:'Cotizaciones',
    RESERVAS:    'Reservas',
    PASAJEROS:   'Pasajeros',
    PAGOS:       'Historial_Pagos'
  },
  ETAPAS: [
    'Nuevos Interesados',
    'Contactados',
    'Cotización por Enviar',
    'Cotización Enviada',
    'Seguimiento',
    'Separó Cupo',
    'Pago Completo',
    'Viaje Realizado',
    'Posventa'
  ],
  DRIVE_FOLDER: 'CRM_Comprobantes_Pagos'
};

// ─── MENÚ (uso desde Sheets en navegador) ─────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✈️ CRM Viajes')
    .addItem('🏠 Panel CRM',           'abrirPanel')
    .addSeparator()
    .addItem('➕ Añadir Cliente',       'abrirFormularioCliente')
    .addItem('🔍 Buscar Cliente',       'abrirBuscarCliente')
    .addSeparator()
    .addItem('📊 Tablero Kanban',       'abrirKanban')
    .addItem('💳 Registrar Pago',       'abrirFormularioPago')
    .addItem('🗺️ Gestionar Destinos',  'abrirFormularioDestino')
    .addSeparator()
    .addItem('⚙️ Configurar Hojas',    'configurarHojas')
    .addItem('🔗 Ver URL WebApp Móvil','mostrarUrlWebApp')
    .addToUi();
}

// ─── WEBAPP MÓVIL ─────────────────────────────────────────
function doGet(e) {
  var tmpl = HtmlService.createTemplateFromFile('WebApp');
  return tmpl.evaluate()
    .setTitle('CRM Agencia de Viajes')
    .addMetaTag('viewport','width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// include() inserta el contenido de archivos HTML en WebApp
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── MOSTRAR URL ──────────────────────────────────────────
function mostrarUrlWebApp() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      '⚠️ La WebApp aún no está publicada.\n\n' +
      'Ve a: Implementar → Nueva implementación → WebApp\n' +
      'Ejecutar como: Yo\n' +
      'Quién accede: Cualquier usuario\n\n' +
      'Luego vuelve a este menú para ver la URL.'
    );
    return;
  }
  SpreadsheetApp.getUi().alert(
    '📱 URL WebApp Móvil:\n\n' + url +
    '\n\nAbre este enlace en el navegador del celular y\n' +
    'guárdalo en la pantalla de inicio como acceso directo.'
  );
}

// ─── CONFIGURAR HOJAS ─────────────────────────────────────
function configurarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = [
    { nombre: CONFIG.SHEETS.CLIENTES,
      cols: ['ID_Cliente','Nombre','Cedula','Celular','Email','Municipio','Origen','Fecha_Registro'] },
    { nombre: CONFIG.SHEETS.DESTINOS,
      cols: ['ID_Destino','Viaje_Ruta','Valor_Persona','Cupo','Cupos_Disponibles','Punto_Encuentro','Incluye','Fecha_Viaje','Estado'] },
    { nombre: CONFIG.SHEETS.COTIZACIONES,
      cols: ['ID_Cotizacion','ID_Cliente','Nombre_Cliente','ID_Destino','Viaje_Ruta','Cantidad_Cupos','Valor_Cotizado','Fecha_Cotizacion','Estado_Cotizacion','Notas'] },
    { nombre: CONFIG.SHEETS.RESERVAS,
      cols: ['ID_Reserva','ID_Cotizacion','ID_Cliente','Nombre_Cliente','ID_Destino','Viaje_Ruta','Fecha_Viaje','Etapa_Actual','Total_Pagar','Total_Abono','Saldo_Pendiente','Fecha_Registro','Notas'] },
    { nombre: CONFIG.SHEETS.PASAJEROS,
      cols: ['ID_Pasajero','ID_Reserva','ID_Cliente','Nombre_Pasajero','Cedula','Celular','Email','Tipo_Documento','Fecha_Nacimiento','Notas'] },
    { nombre: CONFIG.SHEETS.PAGOS,
      cols: ['ID_Pago','ID_Reserva','ID_Cliente','Nombre_Cliente','Fecha_Pago','Monto_Abonado','Metodo_Pago','Referencia_Comprobante','Enlace_Comprobante','Cupos_Reservados','Notas'] }
  ];
  hojas.forEach(function(h) {
    var hoja = ss.getSheetByName(h.nombre);
    if (!hoja) hoja = ss.insertSheet(h.nombre);
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(h.cols);
      hoja.getRange(1,1,1,h.cols.length)
          .setBackground('#1e3a5f').setFontColor('#ffffff')
          .setFontWeight('bold').setFontSize(11);
      hoja.setFrozenRows(1);
    }
  });
  obtenerOCrearCarpetaDrive();
  SpreadsheetApp.getUi().alert(
    '✅ Hojas configuradas correctamente.\n\n' +
    'Siguiente paso: publicar la WebApp desde\n' +
    'Implementar → Nueva implementación → WebApp'
  );
}

// ─── DRIVE ────────────────────────────────────────────────
function obtenerOCrearCarpetaDrive() {
  var it = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.DRIVE_FOLDER);
}

// ─── ABRIR VISTAS EN SHEETS (sidebar / modal) ─────────────
function abrirPanel() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Panel')
      .setTitle('CRM Agencia de Viajes').setWidth(420));
}
function abrirFormularioCliente() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('FormCliente')
      .setTitle('Añadir Cliente').setWidth(420));
}
function abrirBuscarCliente() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('BuscarCliente')
      .setTitle('Buscar Cliente').setWidth(420));
}
function abrirKanban() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('Kanban')
      .setTitle('Tablero Kanban'),
    '📊 Tablero de Reservas');
}
function abrirFormularioPago() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('FormPago')
      .setTitle('Registrar Pago').setWidth(420));
}
function abrirFormularioDestino() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('FormDestino')
      .setTitle('Gestionar Destinos').setWidth(420));
}

// ─── HELPER ID ────────────────────────────────────────────
function generarId(prefijo) {
  return prefijo + '-' + new Date().getTime().toString().slice(-6)
       + Math.random().toString(36).slice(-3).toUpperCase();
}