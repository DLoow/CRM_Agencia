// ============================================================
// CRM AGENCIA DE VIAJES - API_Reservas_Pagos.gs
// ============================================================

// ════════════════════════════════════════════════════════════
//  RESERVAS
// ════════════════════════════════════════════════════════════

// Obtener todas las entradas del tablero (seguimientos + reservas reales)
function obtenerReservas() {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);
    if (!hoja) return [];
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return [];
    const filas = datos.slice(1);
    return filas
      .filter(f => f[0] && f[0].toString().trim() !== '')
      .map((f, idx) => ({
        fila: idx + 2,
        id: f[0].toString(),
        idCotizacion: f[1] ? f[1].toString() : '',
        idCliente: f[2] ? f[2].toString() : '',
        nombreCliente: f[3] ? f[3].toString() : '',
        idDestino: f[4] ? f[4].toString() : '',
        ruta: f[5] ? f[5].toString() : '',
        fechaViaje: f[6] ? (f[6] instanceof Date ? f[6].toLocaleDateString('es-CO') : f[6].toString()) : '',
        etapa: f[7] ? f[7].toString() : 'Nuevos Interesados',
        totalPagar: Number(f[8]) || 0,
        totalAbono: Number(f[9]) || 0,
        saldoPendiente: Number(f[10]) || 0,
        fechaReserva: f[11] ? (f[11] instanceof Date ? f[11].toLocaleDateString('es-CO') : f[11].toString()) : '',
        notas: f[12] ? f[12].toString() : '',
        esReservaReal: !String(f[0]).startsWith('SEG-')
      }));
  } catch (e) {
    console.error('Error en obtenerReservas:', e);
    return [];
  }
}


// Actualizar etapa en el tablero
// Cuando llega a "Separó Cupo" y es un SEG-, se convierte en reserva real
function actualizarEtapaReserva(idActual, nuevaEtapa, datosConfirmacion) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);
  const filas = hoja.getDataRange().getValues();

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === String(idActual).trim()) {

      // Si es SEG- y la nueva etapa es "Separó Cupo", confirmar reserva real
      if (String(filas[i][0]).startsWith('SEG-') && nuevaEtapa === 'Separó Cupo') {
        if (!datosConfirmacion) {
          // Señal al front para que pida los datos de confirmación
          return { ok: false, requiereConfirmacion: true, idSeguimiento: filas[i][0], idCliente: filas[i][2] };
        }
        // Confirmar con datos
        datosConfirmacion.idSeguimiento = filas[i][0];
        return confirmarReserva(datosConfirmacion);
      }

      // Cambio normal de etapa
      hoja.getRange(i + 1, 8).setValue(nuevaEtapa);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Entrada no encontrada' };
}

function obtenerReservasPorEtapa() {
  try {
    const reservas = obtenerReservas();
    const tablero = {};
    CONFIG.ETAPAS.forEach(e => tablero[e] = []);
    reservas.forEach(r => {
      const etapa = r.etapa || 'Nuevos Interesados';
      if (tablero[etapa]) tablero[etapa].push(r);
      else tablero['Nuevos Interesados'].push(r);
    });
    return tablero;
  } catch (e) {
    console.error('Error en obtenerReservasPorEtapa:', e);
    const vacio = {};
    CONFIG.ETAPAS.forEach(e => vacio[e] = []);
    return vacio;
  }
}

// Agregar nueva entrada de seguimiento para un cliente que vuelve a preguntar
function nuevaEntradaSeguimiento(idCliente) {
  const cliente = obtenerClientePorId(idCliente);
  if (!cliente) return { ok: false, error: 'Cliente no encontrado' };
  crearEntradaTablero(cliente);
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  PASAJEROS
// ════════════════════════════════════════════════════════════

function agregarPasajero(datos) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.PASAJEROS);
  const id   = generarId('PAS');
  hoja.appendRow([
    id,
    datos.idReserva,
    datos.idCliente || '',
    datos.nombre,
    datos.cedula,
    datos.celular   || '',
    datos.email     || '',
    datos.tipoDocumento  || 'CC',
    datos.fechaNacimiento|| '',
    datos.notas     || ''
  ]);
  return { ok: true, id };
}

function obtenerPasajerosPorReserva(idReserva) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.PASAJEROS);
  const filas = hoja.getDataRange().getValues().slice(1);
  return filas
    .filter(f => String(f[1]).trim() === String(idReserva).trim())
    .map(f => ({
      id: f[0], idReserva: f[1], nombre: f[3], cedula: f[4],
      celular: f[5], email: f[6], tipoDocumento: f[7], fechaNacimiento: f[8]
    }));
}

// ════════════════════════════════════════════════════════════
//  HISTORIAL DE PAGOS
// ════════════════════════════════════════════════════════════

function registrarPago(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hPago = ss.getSheetByName(CONFIG.SHEETS.PAGOS);
  const hRes  = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);

  // Solo se puede pagar sobre reservas reales (con ID_Reserva)
  if (String(datos.idReserva).startsWith('SEG-')) {
    return { ok: false, error: 'No se puede registrar pago en una entrada de seguimiento. El cliente debe separar cupo primero.' };
  }

  const cliente = obtenerClientePorId(datos.idCliente);
  const id      = generarId('PAG');

  // Subir comprobante a Drive si viene en base64
  let enlaceComprobante = datos.enlaceComprobante || '';
  if (datos.archivoBase64 && datos.nombreArchivo) {
    enlaceComprobante = subirComprobanteDrive(datos.archivoBase64, datos.nombreArchivo, datos.mimeType, id);
  }

  hPago.appendRow([
    id,
    datos.idReserva,
    datos.idCliente,
    cliente ? cliente.nombre : '',
    datos.fechaPago || new Date().toLocaleDateString('es-CO'),
    datos.montoAbonado,
    datos.metodoPago,
    datos.referencia || '',
    enlaceComprobante,
    datos.notas || ''
  ]);

  // Actualizar abonos y saldo en la fila de reserva
  const filas = hRes.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === String(datos.idReserva).trim()) {
      const nuevoAbono = Number(filas[i][9]) + Number(datos.montoAbonado);
      const nuevoSaldo = Math.max(0, Number(filas[i][8]) - nuevoAbono);
      hRes.getRange(i + 1, 10).setValue(nuevoAbono);
      hRes.getRange(i + 1, 11).setValue(nuevoSaldo);
      // Etapa automática al pagar completo
      if (nuevoSaldo === 0 && filas[i][7] !== 'Viaje Realizado') {
        hRes.getRange(i + 1, 8).setValue('Pago Completo');
      }
      break;
    }
  }

  return { ok: true, id, enlace: enlaceComprobante };
}

function obtenerPagosPorReserva(idReserva) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.PAGOS);
  const filas = hoja.getDataRange().getValues().slice(1);
  return filas
    .filter(f => String(f[1]).trim() === String(idReserva).trim())
    .map(f => ({
      id: f[0], fecha: f[4], monto: f[5],
      metodo: f[6], ref: f[7], enlace: f[8], notas: f[9]
    }));
}

// ─── SUBIR COMPROBANTE A DRIVE ────────────────────────────
function subirComprobanteDrive(base64Data, nombreArchivo, mimeType, idPago) {
  try {
    const carpeta = obtenerOCrearCarpetaDrive();
    const bytes   = Utilities.base64Decode(base64Data);
    const blob    = Utilities.newBlob(bytes, mimeType, nombreArchivo);
    const archivo = carpeta.createFile(blob);
    archivo.setName('COMP_' + idPago + '_' + nombreArchivo);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (e) {
    Logger.log('Error subiendo comprobante: ' + e.toString());
    return '';
  }
}

// ─── RESUMEN GENERAL ─────────────────────────────────────
function obtenerResumenGeneral() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hClientes = ss.getSheetByName(CONFIG.SHEETS.CLIENTES);
    const hCotiz = ss.getSheetByName(CONFIG.SHEETS.COTIZACIONES);
    const hReservas = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);
    const hPagos = ss.getSheetByName(CONFIG.SHEETS.PAGOS);

    const totalClientes = Math.max(0, hClientes.getLastRow() - 1);
    const totalCotiz = Math.max(0, hCotiz.getLastRow() - 1);

    const filasRes = hReservas.getDataRange().getValues().slice(1).filter(f => f[0]);
    const totalReservas = filasRes.filter(f => !String(f[0]).startsWith('SEG-')).length;

    let totalRecaudado = 0;
    hPagos.getDataRange().getValues().slice(1).forEach(f => { totalRecaudado += Number(f[5]) || 0; });

    return { totalClientes, totalCotiz, totalReservas, totalRecaudado };
  } catch (e) {
    console.error('Error en obtenerResumenGeneral:', e);
    return { totalClientes: 0, totalCotiz: 0, totalReservas: 0, totalRecaudado: 0 };
  }
}