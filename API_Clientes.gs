// ============================================================
// CRM AGENCIA DE VIAJES - API_Clientes.gs
// ============================================================

// ─── SOLO VERIFICAR (sin crear) ───────────────────────────
// Usado por el formulario para saber si la cédula ya existe
function verificarCedulaExiste(cedula) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.CLIENTES);
  const filas = hoja.getDataRange().getValues();

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][2]).trim() === String(cedula).trim()) {
      return {
        existe: true,
        cliente: {
          id:        filas[i][0],
          nombre:    filas[i][1],
          cedula:    filas[i][2],
          celular:   filas[i][3],
          email:     filas[i][4],
          municipio: filas[i][5],
          origen:    filas[i][6],
        }
      };
    }
  }
  return { existe: false };
}

// ─── CREAR CLIENTE NUEVO ──────────────────────────────────
// Solo se llama cuando ya se verificó que NO existe
function crearCliente(datos) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.CLIENTES);

  // Doble-check: no crear duplicados
  const check = verificarCedulaExiste(datos.cedula);
  if (check.existe) {
    return { ok: false, existe: true, cliente: check.cliente };
  }

  const id = generarId('CLI');
  hoja.appendRow([
    id,
    datos.nombre,
    datos.cedula,
    datos.celular,
    datos.email    || '',
    datos.municipio|| '',
    datos.origen   || '',
    new Date().toLocaleDateString('es-CO')
  ]);
  aplicarEstiloFila(hoja, hoja.getLastRow());

  const cliente = {
    id, nombre: datos.nombre, cedula: datos.cedula,
    celular: datos.celular, email: datos.email || '',
    municipio: datos.municipio || '', origen: datos.origen || ''
  };

  // ── Todo cliente nuevo entra al tablero en "Nuevos Interesados"
  // sin destino ni reserva formal (ID temporal de seguimiento)
  crearEntradaTablero(cliente);

  return { ok: true, existe: false, cliente };
}

// ─── CREAR ENTRADA EN TABLERO (sin ID_Reserva real) ───────
// Representa al cliente en el Kanban desde que llega
// El ID_Reserva real se genera solo al "Separar Cupo"
function crearEntradaTablero(cliente) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);

  // ID de seguimiento provisional (distinguible de un ID_Reserva real)
  const idSeguimiento = 'SEG-' + new Date().getTime().toString().slice(-6);

  hoja.appendRow([
    idSeguimiento,    // ID_Reserva (provisional hasta Separó Cupo)
    '',               // ID_Cotizacion
    cliente.id,       // ID_Cliente
    cliente.nombre,   // Nombre_Cliente
    '',               // ID_Destino
    '',               // Viaje_Ruta
    '',               // Fecha_Viaje
    'Nuevos Interesados', // Etapa_Actual
    0,                // Total_Pagar
    0,                // Total_Abono
    0,                // Saldo_Pendiente
    new Date().toLocaleDateString('es-CO'), // Fecha_Reserva
    'Lead inicial'    // Notas
  ]);
}

// ─── CONFIRMAR RESERVA (al llegar a "Separó Cupo") ────────
// Reemplaza el ID provisional por un ID_Reserva real
// y vincula destino, cotización y valor
function confirmarReserva(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hRes  = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);
  const hDest = ss.getSheetByName(CONFIG.SHEETS.DESTINOS);
  const filas = hRes.getDataRange().getValues();
  

  // Buscar la fila de seguimiento del cliente
  let filaIdx = -1;
  for (let i = 1; i < filas.length; i++) {
    // Buscar por idSeguimiento o por idCliente con etapa aún no reservada
    if (String(filas[i][0]).trim() === String(datos.idSeguimiento).trim()) {
      filaIdx = i + 1;
      break;
    }
  }
  if (filaIdx === -1) return { ok: false, error: 'Fila de seguimiento no encontrada' };

  // Obtener datos del destino
  let rutaNombre = '', fechaViaje = '';
if (datos.idDestino) {
  const destinos = hDest.getDataRange().getValues();
  let encontrado = false;
  for (let i = 1; i < destinos.length; i++) {
    if (String(destinos[i][0]).trim() === String(datos.idDestino).trim()) {
      encontrado = true;
      rutaNombre = destinos[i][1];
      // Manejo robusto de fecha
      let fechaRaw = destinos[i][7];
      if (fechaRaw) {
        if (fechaRaw instanceof Date) {
          fechaViaje = Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (typeof fechaRaw === 'string') {
          // Intentar parsear string como fecha
          let fechaParse = new Date(fechaRaw);
          if (!isNaN(fechaParse.getTime())) {
            fechaViaje = Utilities.formatDate(fechaParse, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else {
            fechaViaje = fechaRaw;
          }
        } else {
          fechaViaje = fechaRaw.toString();
        }
      }
      hRes.getRange(filaIdx, 14).setValue(datos.cantidadCupos || 1); // Cupos_Reservados
      // Descontar cupos
      const cuposActuales   = Number(destinos[i][4]);
      const cuposReservados = Number(datos.cantidadCupos || 1);
      if (cuposActuales >= cuposReservados) {
        hDest.getRange(i + 1, 5).setValue(cuposActuales - cuposReservados);
      }
      break;
    }
  }
  if (!encontrado) {
    console.error('Destino no encontrado: ' + datos.idDestino);
  }
}

  // Generar ID_Reserva real
  const idReserva = generarId('RES');

  // Actualizar la fila con todos los datos reales
  hRes.getRange(filaIdx, 1).setValue(idReserva);
  hRes.getRange(filaIdx, 2).setValue(datos.idCotizacion || '');
  hRes.getRange(filaIdx, 5).setValue(datos.idDestino    || '');
  hRes.getRange(filaIdx, 6).setValue(rutaNombre);
  hRes.getRange(filaIdx, 7).setValue(fechaViaje);
  hRes.getRange(filaIdx, 8).setValue('Separó Cupo');
  hRes.getRange(filaIdx, 9).setValue(datos.totalPagar   || 0);
  hRes.getRange(filaIdx, 11).setValue(datos.totalPagar  || 0); // saldo = total
  hRes.getRange(filaIdx, 13).setValue(datos.notas       || '');

  // Marcar cotización como Reservada
  if (datos.idCotizacion) {
    actualizarEstadoCotizacion(datos.idCotizacion, 'Reservada');
  }

  return { ok: true, idReserva, rutaNombre };
}

// ─── BUSCAR CLIENTES ──────────────────────────────────────
function buscarClientes(query) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.CLIENTES);
  const filas = hoja.getDataRange().getValues();
  const q     = String(query).toLowerCase().trim();

  const resultados = [];
  for (let i = 1; i < filas.length; i++) {
    if (!filas[i][0]) continue;
    const texto = [filas[i][0],filas[i][1],filas[i][2],filas[i][3],filas[i][4],filas[i][5],filas[i][6]]
      .join(' ').toLowerCase();
    if (texto.includes(q)) {
      resultados.push({
        id: filas[i][0], nombre: filas[i][1], cedula: filas[i][2],
        celular: filas[i][3], email: filas[i][4], municipio: filas[i][5], origen: filas[i][6]
      });
    }
  }
  return resultados;
}

// ─── OBTENER CLIENTE POR ID ───────────────────────────────
function obtenerClientePorId(id) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.CLIENTES);
  const filas = hoja.getDataRange().getValues();

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === String(id).trim()) {
      return {
        id: filas[i][0], nombre: filas[i][1], cedula: filas[i][2],
        celular: filas[i][3], email: filas[i][4], municipio: filas[i][5], origen: filas[i][6]
      };
    }
  }
  return null;
}

// ─── ACTUALIZAR CLIENTE ───────────────────────────────────
function actualizarCliente(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.CLIENTES);
  const hRes  = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);
  const filas = hoja.getDataRange().getValues();

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === String(datos.id).trim()) {
      const f = i + 1;
      hoja.getRange(f, 2).setValue(datos.nombre);
      hoja.getRange(f, 3).setValue(datos.cedula);
      hoja.getRange(f, 4).setValue(datos.celular);
      hoja.getRange(f, 5).setValue(datos.email);
      hoja.getRange(f, 6).setValue(datos.municipio);
      hoja.getRange(f, 7).setValue(datos.origen);

      // Sincronizar nombre en la hoja de Reservas
      const filasRes = hRes.getDataRange().getValues();
      for (let j = 1; j < filasRes.length; j++) {
        if (String(filasRes[j][2]).trim() === String(datos.id).trim()) {
          hRes.getRange(j + 1, 4).setValue(datos.nombre);
        }
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Cliente no encontrado' };
}

// ─── HISTORIAL DE CLIENTE ─────────────────────────────────
function obtenerHistorialCliente(idCliente) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const hCotiz    = ss.getSheetByName(CONFIG.SHEETS.COTIZACIONES);
  const hReservas = ss.getSheetByName(CONFIG.SHEETS.RESERVAS);
  const hPagos    = ss.getSheetByName(CONFIG.SHEETS.PAGOS);

  const cotizaciones = hCotiz.getDataRange().getValues().slice(1)
    .filter(f => String(f[1]).trim() === String(idCliente).trim())
    .map(f => ({
      id: f[0], idDestino: f[3], destino: f[4],
      cupos: f[5], valor: f[6], fecha: f[7], estado: f[8]
    }));

  const reservas = hReservas.getDataRange().getValues().slice(1)
    .filter(f => String(f[2]).trim() === String(idCliente).trim())
    .map(f => ({
      id: f[0], idCotizacion: f[1], idDestino: f[4],
      destino: f[5], fechaViaje: f[6], etapa: f[7],
      total: f[8], abono: f[9], saldo: f[10]
    }));

  const pagos = hPagos.getDataRange().getValues().slice(1)
    .filter(f => String(f[2]).trim() === String(idCliente).trim())
    .map(f => ({
      id: f[0], idReserva: f[1], fecha: f[4],
      monto: f[5], metodo: f[6], ref: f[7], enlace: f[8]
    }));

  return { cotizaciones, reservas, pagos };
}

// ─── HELPER ESTILO FILA ───────────────────────────────────
function aplicarEstiloFila(hoja, numFila) {
  hoja.getRange(numFila, 1, 1, hoja.getLastColumn()).setFontSize(10);
}