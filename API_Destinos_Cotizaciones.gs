// ============================================================
// CRM AGENCIA DE VIAJES - API_Destinos_Cotizaciones.gs
// ============================================================

// ════════════════════════════════════════════════════════════
//  DESTINOS
// ════════════════════════════════════════════════════════════

function obtenerDestinos() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.DESTINOS);
  if (!hoja) return [];
  const filas = hoja.getDataRange().getValues().slice(1);
  return filas.filter(f => f[0]).map(f => ({
    id:              f[0],
    ruta:            f[1],
    valorPersona:    f[2],
    cupo:            f[3],
    cuposDisponibles:f[4],
    puntoEncuentro:  f[5],
    incluye:         f[6],
    fechaViaje:      f[7] ? (f[7] instanceof Date ? f[7].toLocaleDateString('es-CO') : f[7]) : '',
    estado:          f[8]
  }));
}

function crearDestino(datos) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.SHEETS.DESTINOS);
  const id   = generarId('DST');
  hoja.appendRow([
    id,
    datos.ruta,
    Number(datos.valorPersona) || 0,
    Number(datos.cupo) || 0,
    Number(datos.cupo) || 0,   // cupos disponibles = cupo inicial
    datos.puntoEncuentro || '',
    datos.incluye        || '',
    datos.fechaViaje     || '',
    datos.estado         || 'Activo'
  ]);
  return { ok: true, id };
}

function actualizarDestino(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.DESTINOS);
  const filas = hoja.getDataRange().getValues();

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() !== String(datos.id).trim()) continue;
    const f = i + 1;

    // Modo solo-estado (toggle activo/inactivo desde la lista)
    if (datos.ruta === '_keep_') {
      hoja.getRange(f, 9).setValue(datos.estado);
      return { ok: true };
    }

    // Edición completa
    hoja.getRange(f, 2).setValue(datos.ruta);
    hoja.getRange(f, 3).setValue(Number(datos.valorPersona) || 0);
    hoja.getRange(f, 4).setValue(Number(datos.cupo) || 0);
    // Recalcular cupos disponibles si cambió el cupo total
    const cupoAnterior = Number(filas[i][3]) || 0;
    const cuposDispo   = Number(filas[i][4]) || 0;
    const diff         = (Number(datos.cupo) || 0) - cupoAnterior;
    hoja.getRange(f, 5).setValue(Math.max(0, cuposDispo + diff));
    hoja.getRange(f, 6).setValue(datos.puntoEncuentro || '');
    hoja.getRange(f, 7).setValue(datos.incluye        || '');
    hoja.getRange(f, 8).setValue(datos.fechaViaje     || '');
    hoja.getRange(f, 9).setValue(datos.estado         || 'Activo');
    return { ok: true };
  }
  return { ok: false, error: 'Destino no encontrado' };
}

// ════════════════════════════════════════════════════════════
//  COTIZACIONES
// ════════════════════════════════════════════════════════════

function crearCotizacion(datos) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hCot  = ss.getSheetByName(CONFIG.SHEETS.COTIZACIONES);
  const hDest = ss.getSheetByName(CONFIG.SHEETS.DESTINOS);

  const destinos = hDest.getDataRange().getValues();
  let valorPersona = 0, rutaNombre = '';
  for (let i = 1; i < destinos.length; i++) {
    if (String(destinos[i][0]).trim() === String(datos.idDestino).trim()) {
      valorPersona = destinos[i][2];
      rutaNombre   = destinos[i][1];
      break;
    }
  }

  const valorTotal = valorPersona * (datos.cantidadCupos || 1);
  const id         = generarId('COT');
  const cliente    = obtenerClientePorId(datos.idCliente);

  hCot.appendRow([
    id,
    datos.idCliente,
    cliente ? cliente.nombre : '',
    datos.idDestino,
    rutaNombre,
    datos.cantidadCupos || 1,
    valorTotal,
    new Date().toLocaleDateString('es-CO'),
    'Pendiente',
    datos.notas || ''
  ]);

  return { ok: true, id, valorTotal, rutaNombre };
}

function obtenerCotizaciones(idCliente) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.COTIZACIONES);
  const filas = hoja.getDataRange().getValues().slice(1);

  const todas = filas.filter(f => f[0]).map(f => ({
    id:            f[0],
    idCliente:     f[1],
    nombreCliente: f[2],
    idDestino:     f[3],
    ruta:          f[4],
    cupos:         f[5],
    valor:         f[6],
    fecha:         f[7],
    estado:        f[8],
    notas:         f[9]
  }));

  if (idCliente) return todas.filter(c => String(c.idCliente) === String(idCliente));
  return todas;
}

function actualizarEstadoCotizacion(idCotizacion, nuevoEstado) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(CONFIG.SHEETS.COTIZACIONES);
  const filas = hoja.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === String(idCotizacion).trim()) {
      hoja.getRange(i + 1, 9).setValue(nuevoEstado);
      return { ok: true };
    }
  }
  return { ok: false };
}