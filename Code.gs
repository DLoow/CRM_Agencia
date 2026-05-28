// ============================================================
// CRM AGENCIA DE VIAJES - Code.gs
// ============================================================

const CONFIG = {
  SHEETS: {
    CLIENTES:    'Clientes',
    DESTINOS:    'Destinos',
    COTIZACIONES:'Cotizaciones',
    RESERVAS:    'Reservas',
    PASAJEROS:   'Pasajeros',
    PAGOS:       'Historial_Pagos',
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

// ─── MENÚ PRINCIPAL (escritorio / Sheets web) ─────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✈️ CRM Viajes')
    .addItem('🏠 Panel CRM',              'abrirPanel')
    .addSeparator()
    .addItem('➕ Añadir Cliente',          'abrirFormularioCliente')
    .addItem('🔍 Buscar Cliente',          'abrirBuscarCliente')
    .addSeparator()
    .addItem('📊 Tablero Kanban',          'abrirKanban')
    .addItem('💳 Registrar Pago',          'abrirFormularioPago')
    .addItem('🗺️ Gestionar Destinos',      'abrirFormularioDestino')
    .addSeparator()
    .addItem('⚙️ Configurar Hojas',        'configurarHojas')
    .addItem('🔗 Ver URL WebApp Móvil',    'mostrarUrlWebApp')
    .addToUi();
}

// ─── WEBAPP MÓVIL: doGet ───────────────────────────────────
// Publicar como WebApp permite acceso desde el celular sin
// necesitar la app de Sheets ni el menú "Extensiones".
// Despliegue: Implementar > Nueva implementación > WebApp
//   Ejecutar como: Yo  |  Quién puede acceder: Cualquier usuario
function doGet(e) {
  const pagina = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';
  const archivos = {
    'home':     'WebApp',
    'kanban':   'Kanban',
    'cliente':  'FormCliente',
    'buscar':   'BuscarCliente',
    'pago':     'FormPago',
    'destino':  'FormDestino',
  };
  const archivo = archivos[pagina] || 'WebApp';
  return HtmlService.createTemplateFromFile(archivo)
    .evaluate()
    .setTitle('CRM Agencia de Viajes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── MOSTRAR URL DE LA WEBAPP ─────────────────────────────
function mostrarUrlWebApp() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      '⚠️ La WebApp aún no está publicada.\n\n' +
      'Ve a: Implementar → Nueva implementación → WebApp\n' +
      'Ejecutar como: Yo\n' +
      'Quién accede: Cualquier usuario (o Usuario de Google si prefieres login)\n\n' +
      'Luego vuelve aquí para ver la URL.'
    );
    return;
  }
  SpreadsheetApp.getUi().alert(
    '📱 URL de la WebApp Móvil:\n\n' + url +
    '\n\nGuarda este enlace en tu celular como acceso directo.\n' +
    'También puedes usar ?page=kanban, ?page=buscar, ?page=pago'
  );
}

// ─── CONFIGURAR HOJAS ─────────────────────────────────────
function configurarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojas = [
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
      cols: ['ID_Pago','ID_Reserva','ID_Cliente','Nombre_Cliente','Fecha_Pago','Monto_Abonado','Metodo_Pago','Referencia_Comprobante','Enlace_Comprobante','Notas'] },
  ];

  hojas.forEach(h => {
    let hoja = ss.getSheetByName(h.nombre);
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
    '✅ Hojas configuradas.\n\n' +
    'Próximo paso: publica la WebApp desde\n' +
    'Implementar → Nueva implementación → WebApp\n' +
    'para poder usar el CRM desde el celular.'
  );
}

// ─── DRIVE ────────────────────────────────────────────────
function obtenerOCrearCarpetaDrive() {
  const it = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.DRIVE_FOLDER);
}

// ─── ABRIR VISTAS (sidebar / modal) ──────────────────────
function abrirPanel() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('Panel').evaluate()
      .setTitle('CRM Agencia de Viajes').setWidth(420));
}
function abrirFormularioCliente() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('FormCliente').evaluate()
      .setTitle('Añadir Cliente').setWidth(420));
}
function abrirBuscarCliente() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('BuscarCliente').evaluate()
      .setTitle('Buscar Cliente').setWidth(420));
}
function abrirKanban() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('Kanban').evaluate()
      .setTitle('Tablero Kanban')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME),
    '📊 Tablero de Reservas');
}
function abrirFormularioPago() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('FormPago').evaluate()
      .setTitle('Registrar Pago').setWidth(420));
}
function abrirFormularioDestino() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('FormDestino').evaluate()
      .setTitle('Gestionar Destinos').setWidth(420));
}

// ─── HELPERS ─────────────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function generarId(prefijo) {
  return prefijo + '-' + new Date().getTime().toString().slice(-6)
       + Math.random().toString(36).slice(-3).toUpperCase();
}