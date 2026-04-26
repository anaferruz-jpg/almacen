var SHEET_ID = '1YFu6po3AC_bg0k6lOMTy_IuBnVJX5Ng9_28C9TNWPYg';

function doGet(e) {
  var action = e.parameter.action;
  // Servir páginas HTML
  if (!action) {
    var page = e.parameter.page || 'index';
    if (page === 'facturacion') {
      return HtmlService.createHtmlOutputFromFile('Facturacion')
        .setTitle('SEGUFIJA — Facturación')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('SEGUFIJA — Almacén')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // API
  if (action === 'getInventario')    return json(getInventario());
  if (action === 'getHistorial')     return json(getHistorial());
  if (action === 'getObras')         return json(getObras());
  if (action === 'getFacturas')      return json(getFacturas());
  if (action === 'getGastos')        return json(getGastos());
  if (action === 'getClientes')      return json(getClientes());
  if (action === 'getProveedores')   return json(getProveedores());
  if (action === 'getCobros')        return json(getCobros());
  if (action === 'getGastosInt')     return json(getGastosInternos());
  if (action === 'getCRM')           return json(getCRM());
  if (action === 'getDocsPersonal')  return json(getDocsPersonal(e.parameter.trabajador));
  if (action === 'getDocsGestoria')  return json(getDocsGestoria());
  if (action === 'getDocsObra')      return json(getDocsObra(e.parameter.obra));
  if (action === 'getOperarios')     return json(getOperarios());
  if (action === 'getPresupuestos')  return json(getPresupuestos());
  if (action === 'getCostesMateriales') return json(getCostesMateriales());
  if (action === 'getProformas')        return json(getProformas());
  return json({error: 'accion no valida'});
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action === 'updateStock')            return json(updateStock(data));
  if (data.action === 'restoreStock')           return json(restoreStock(data));
  if (data.action === 'addHistorial')           return json(addHistorial(data));
  if (data.action === 'updatePrecios')          return json(updatePrecios(data));
  if (data.action === 'addArticulo')            return json(addArticulo(data));
  if (data.action === 'editArticulo')           return json(editArticulo(data));
  if (data.action === 'deleteArticulo')         return json(deleteArticulo(data));
  if (data.action === 'saveObra')               return json(saveObra(data));
  if (data.action === 'deleteObra')             return json(deleteObra(data));
  if (data.action === 'deleteHistorial')        return json(deleteHistorial(data));
  if (data.action === 'marcarPedidoEntregado')  return json(marcarPedidoEntregado(data));
  if (data.action === 'editarHistorial')        return json(editarHistorial(data));
  if (data.action === 'renameObraHistorial')    return json(renameObraHistorial(data));
  if (data.action === 'saveFactura')            return json(saveFactura(data));
  if (data.action === 'deleteFactura')          return json(deleteFactura(data));
  if (data.action === 'saveGasto')              return json(saveGasto(data));
  if (data.action === 'deleteGasto')            return json(deleteGasto(data));
  if (data.action === 'saveCliente')            return json(saveCliente(data));
  if (data.action === 'deleteCliente')          return json(deleteCliente(data));
  if (data.action === 'saveProveedor')          return json(saveProveedor(data));
  if (data.action === 'deleteProveedor')        return json(deleteProveedor(data));
  if (data.action === 'saveCobro')              return json(saveCobro(data));
  if (data.action === 'deleteCobro')            return json(deleteCobro(data));
  if (data.action === 'saveGastoInt')           return json(saveGastoInterno(data));
  if (data.action === 'deleteGastoInt')         return json(deleteGastoInterno(data));
  if (data.action === 'ocrFactura')             return json(ocrFactura(data));
  if (data.action === 'saveFileToDrive')        return json(saveFileToDrive(data));
  if (data.action === 'saveDocPersonal')        return json(saveDocPersonal(data));
  if (data.action === 'deleteDocPersonal')      return json(deleteDocPersonal(data));
  if (data.action === 'saveDocGestoria')        return json(saveDocGestoria(data));
  if (data.action === 'deleteDocGestoria')      return json(deleteDocGestoria(data));
  if (data.action === 'saveDocObra')            return json(saveDocObra(data));
  if (data.action === 'deleteDocObra')          return json(deleteDocObra(data));
  if (data.action === 'saveOperario')           return json(saveOperario(data));
  if (data.action === 'deleteOperario')         return json(deleteOperario(data));
  if (data.action === 'saveCRM')               return json(saveCRM(data));
  if (data.action === 'deleteCRM')             return json(deleteCRM(data));
  if (data.action === 'sendEmailPedido')        return json(sendEmailPedido(data));
  if (data.action === 'saveProforma')            return json(saveProforma(data));
  if (data.action === 'savePresupuesto')         return json(savePresupuesto(data));
  if (data.action === 'deletePresupuesto')       return json(deletePresupuesto(data));
  if (data.action === 'saveFileToFolder')        return json(saveFileToFolder(data));
  return json({error: 'accion no valida'});
}

// ══════════════════════════════════════════════════════
// INVENTARIO
// ══════════════════════════════════════════════════════
function getInventario() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  if (!sheet) return [];
  asegurarColumnas(sheet);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      nombre:   rows[i][0],
      stock:    rows[i][1] || 0,
      nec:      rows[i][2] || 0,
      pvp:      rows[i][3] || 0,
      pcompra:  rows[i][4] || 0,
      stockMin: rows[i][5] || 0,
      cod:      rows[i][6] || ''
    });
  }
  return result;
}

function asegurarColumnas(sheet) {
  var h = sheet.getRange(1,1,1,7).getValues()[0];
  if (!h[0]) sheet.getRange(1,1).setValue('Nombre');
  if (!h[1]) sheet.getRange(1,2).setValue('Stock');
  if (!h[2]) sheet.getRange(1,3).setValue('Necesario');
  if (!h[3]) sheet.getRange(1,4).setValue('PVP');
  if (!h[4]) sheet.getRange(1,5).setValue('PrecioCompra');
  if (!h[5]) sheet.getRange(1,6).setValue('StockMinimo');
  if (!h[6]) sheet.getRange(1,7).setValue('Codigo');
}

// ══════════════════════════════════════════════════════
// HISTORIAL
// ══════════════════════════════════════════════════════
function getHistorial() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Historial');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var col9 = String(rows[i][8] || '');
    var anulada = false;
    var entregado = false;
    var fechaEntrega = '';
    var estadoPedido = String(rows[i][9] || '');
    var fechaLimite = String(rows[i][10] || '');
    if (!estadoPedido || estadoPedido === 'false' || estadoPedido === '') {
      if (col9.indexOf('entregado:') >= 0) {
        entregado = true;
        estadoPedido = 'entregado';
        fechaEntrega = col9.replace('entregado:','').trim();
      } else if (col9.indexOf('estado:') >= 0) {
        var partes = col9.replace('estado:','').split('|fl:');
        estadoPedido = partes[0].trim();
        if (!fechaLimite && partes[1]) fechaLimite = partes[1].trim();
      } else if (col9 === 'true' || col9 === true) {
        anulada = true;
        estadoPedido = '';
      } else {
        estadoPedido = 'pendiente';
      }
    } else if (estadoPedido === 'entregado') {
      entregado = true;
      fechaEntrega = fechaLimite;
    } else if (col9 === 'true' || col9 === true) {
      anulada = true;
    }
    result.push({
      rowIndex:     i + 1,
      tipo:         rows[i][0],
      fecha:        rows[i][1],
      obra:         rows[i][2],
      resp:         rows[i][3] || '',
      prov:         rows[i][4] || '',
      ref:          rows[i][5] || '',
      obs:          rows[i][6] || '',
      items:        JSON.parse(rows[i][7] || '[]'),
      anulada:      anulada,
      entregado:    entregado,
      fechaEntrega: fechaEntrega,
      estadoPedido: estadoPedido || 'pendiente',
      fechaLimite:  fechaLimite
    });
  }
  return result;
}

function addHistorial(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Historial');
  if (!sheet) {
    sheet = ss.insertSheet('Historial');
    sheet.appendRow(['Tipo','Fecha','Obra','Responsable','Proveedor','Referencia','Observaciones','Items','Anulada','EstadoPedido','FechaLimite']);
  }
  sheet.appendRow([
    data.tipo, data.fecha, data.obra||'',
    data.resp||'', data.prov||'', data.ref||'', data.obs||'',
    JSON.stringify(data.items), data.anulada||false,
    'pendiente', ''
  ]);
  return {ok:true};
}

function marcarPedidoEntregado(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Historial');
  if (!sheet) return {ok:false, msg:'No hay hoja Historial'};
  var estadoPedido = data.estadoPedido || 'entregado';
  var fechaLimite = data.fechaLimite || '';
  if (data.rowIndex && data.rowIndex > 1) {
    sheet.getRange(data.rowIndex, 10).setValue(estadoPedido);
    if (fechaLimite) sheet.getRange(data.rowIndex, 11).setValue(fechaLimite);
    if (estadoPedido === 'entregado') {
      sheet.getRange(data.rowIndex, 9).setValue('entregado:' + (data.fechaEntrega||''));
    }
    return {ok:true};
  }
  var rows = sheet.getDataRange().getValues();
  for (var j = rows.length-1; j >= 1; j--) {
    if (rows[j][0] === 'pedido' && rows[j][2] === data.obra) {
      sheet.getRange(j+1, 10).setValue(estadoPedido);
      if (fechaLimite) sheet.getRange(j+1, 11).setValue(fechaLimite);
      if (estadoPedido === 'entregado') {
        sheet.getRange(j+1, 9).setValue('entregado:' + (data.fechaEntrega||''));
      }
      return {ok:true};
    }
  }
  return {ok:true};
}

function editarHistorial(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Historial');
  if (!sheet) return {ok:false};
  if (data.rowIndex && data.rowIndex > 1) {
    if (data.prov  !== undefined) sheet.getRange(data.rowIndex,5).setValue(data.prov);
    if (data.ref   !== undefined) sheet.getRange(data.rowIndex,6).setValue(data.ref);
    if (data.obs   !== undefined) sheet.getRange(data.rowIndex,7).setValue(data.obs);
    if (data.items !== undefined) sheet.getRange(data.rowIndex,8).setValue(JSON.stringify(data.items));
    return {ok:true};
  }
  return {ok:false};
}

function deleteHistorial(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Historial');
  if (!sheet) return {ok:false};
  if (data.rowIndex && data.rowIndex > 1) {
    sheet.deleteRow(data.rowIndex);
    return {ok:true};
  }
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.tipo && rows[i][2] === data.obra) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// OBRAS
// ══════════════════════════════════════════════════════
function getObras() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Obras');
  if (!sheet) return [];
  asegurarColumnasObras(sheet);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      nombre:            rows[i][0],
      lugar:             rows[i][1] || '—',
      presupuesto:       rows[i][2] || 0,
      inicio:            rows[i][3] || '',
      fin:               rows[i][4] || '',
      estado:            rows[i][5] || 'activa',
      obs:               rows[i][6] || '',
      partidas:          JSON.parse(rows[i][7]  || '[]'),
      personal:          JSON.parse(rows[i][8]  || '[]'),
      pctAlquiler:       rows[i][9]  || 0,
      diasAlquiler:      rows[i][10] || 0,
      previsionPersonal: JSON.parse(rows[i][11] || '[]'),
      tipo:              rows[i][12] || 'fija',
      ml:                rows[i][13] || 0,
      meses:             rows[i][14] || 0,
      formaPago:         rows[i][15] || '',
      plazo:             rows[i][16] || '',
      contacto:          JSON.parse(rows[i][17] || '{}'),
      codObra:           rows[i][18] || '',
      constructora:      (function(v){ return (v && v.length > 20 && !/\s/.test(v)) ? '' : v; })(String(rows[i][19]||'')),
      carpetaDriveId:    (function(v19,v20){ return (v20 && v20.length > 20 && !/\s/.test(v20)) ? v20 : (v19 && v19.length > 20 && !/\s/.test(v19)) ? v19 : v20; })(String(rows[i][19]||''),String(rows[i][20]||''))
    });
  }
  return result;
}

function asegurarColumnasObras(sheet) {
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!h[18]) sheet.getRange(1, 19).setValue('CodObra');
  if (!h[19]) sheet.getRange(1, 20).setValue('Constructora');
  if (!h[20]) sheet.getRange(1, 21).setValue('CarpetaDriveId');
}

function saveObra(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Obras');
  if (!sheet) {
    sheet = ss.insertSheet('Obras');
    sheet.appendRow(['Nombre','Lugar','Presupuesto','Inicio','Fin','Estado','Obs','Partidas','Personal','PctAlquiler','DiasAlquiler','PrevisionPersonal','Tipo','ML','Meses','FormaPago','Plazo','Contacto','CodObra','Constructora','CarpetaDriveId']);
  }
  asegurarColumnasObras(sheet);
  var rows = sheet.getDataRange().getValues();
  var rowData = [
    data.nombre, data.lugar||'—', data.presupuesto||0,
    data.inicio||'', data.fin||'', data.estado||'activa',
    data.obs||'', JSON.stringify(data.partidas||[]),
    JSON.stringify(data.personal||[]), data.pctAlquiler||0,
    data.diasAlquiler||0, JSON.stringify(data.previsionPersonal||[]),
    data.tipo||'fija', data.ml||0, data.meses||0,
    data.formaPago||'', data.plazo||'', JSON.stringify(data.contacto||{}),
    data.codObra||'', data.constructora||data.cliente||'', data.carpetaDriveId||''
  ];
  var codObra = (data.codObra||'').trim();
  for (var i = 1; i < rows.length; i++) {
    // 1. Buscar por codObra (más fiable, evita duplicados por nombre distinto)
    if (codObra && (rows[i][18]||'').trim() === codObra) {
      sheet.getRange(i+1,1,1,21).setValues([rowData]);
      return {ok:true};
    }
    // 2. Buscar por nombreOriginal o nombre
    if (rows[i][0] === data.nombreOriginal || rows[i][0] === data.nombre) {
      sheet.getRange(i+1,1,1,21).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteObra(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Obras');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombre) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// STOCK
// ══════════════════════════════════════════════════════
function updateStock(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  var rows = sheet.getDataRange().getValues();
  for (var j = 0; j < data.items.length; j++) {
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.items[j].nombre) {
        sheet.getRange(i+1,2).setValue(Math.max(0,(rows[i][1]||0)-data.items[j].qty));
        break;
      }
    }
  }
  return {ok:true};
}

function restoreStock(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  var rows = sheet.getDataRange().getValues();
  for (var j = 0; j < data.items.length; j++) {
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.items[j].nombre) {
        sheet.getRange(i+1,2).setValue((rows[i][1]||0)+(data.items[j].qty||0));
        break;
      }
    }
  }
  return {ok:true};
}

function updatePrecios(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombre) {
      sheet.getRange(i+1,4).setValue(data.pvp||0);
      sheet.getRange(i+1,5).setValue(data.pcompra||0);
      if (data.stockMin !== undefined) sheet.getRange(i+1,6).setValue(data.stockMin||0);
      break;
    }
  }
  return {ok:true};
}

function addArticulo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  if (!sheet) {
    sheet = ss.insertSheet('Inventario');
    sheet.appendRow(['Nombre','Stock','Necesario','PVP','PrecioCompra','StockMinimo']);
  }
  asegurarColumnas(sheet);
  sheet.appendRow([data.nombre,data.stock||0,data.nec||0,data.pvp||0,data.pcompra||0,data.stockMin||0,data.cod||'']);
  return {ok:true};
}

function editArticulo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombreOriginal) {
      // Nombre bloqueado — no se modifica. Solo datos y código.
      sheet.getRange(i+1,2).setValue(data.stock||0);
      sheet.getRange(i+1,3).setValue(data.nec||0);
      sheet.getRange(i+1,4).setValue(data.pvp||0);
      sheet.getRange(i+1,5).setValue(data.pcompra||0);
      sheet.getRange(i+1,6).setValue(data.stockMin||0);
      sheet.getRange(i+1,7).setValue(data.cod||'');
      return {ok:true};
    }
  }
  return {ok:false};
}

function deleteArticulo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombre) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// DEDUPLICAR OBRAS — ejecutar UNA SOLA VEZ desde el editor
// Elimina obras duplicadas: conserva la que tiene codObra, borra la sin código
// ══════════════════════════════════════════════════════
function deduplicarObras() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Obras');
  if (!sheet) { Logger.log('Hoja Obras no encontrada'); return; }

  function norm(s) {
    return (s||'').toLowerCase()
      .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i').replace(/ó/g,'o').replace(/ú/g,'u')
      .replace(/[^a-z0-9 ]/g,' ').replace(/\s+/g,' ').trim();
  }
  function similar(a, b) {
    if (a === b) return true;
    var shorter = a.length <= b.length ? a : b;
    var longer  = a.length <= b.length ? b : a;
    return shorter.length >= 6 && longer.indexOf(shorter) !== -1;
  }

  var rows = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  var procesados = []; // { normNom, idx, codObra, partidasLen }

  for (var i = 1; i < rows.length; i++) {
    var nombre = (rows[i][0]||'').toString().trim();
    if (!nombre) continue;
    var n = norm(nombre);
    var codObra = (rows[i][18]||'').toString().trim();
    var partidasLen = (rows[i][7]||'').toString().length;

    // Buscar si ya hay una entrada similar
    var matchIdx = -1;
    for (var j = 0; j < procesados.length; j++) {
      if (similar(procesados[j].normNom, n)) { matchIdx = j; break; }
    }

    if (matchIdx === -1) {
      procesados.push({ normNom: n, idx: i, codObra: codObra, partidasLen: partidasLen });
    } else {
      // Decidir cuál conservar: la que tiene codObra gana; si empate, la de más partidas
      var existente = procesados[matchIdx];
      var existTiene = existente.codObra ? 1 : 0;
      var nuevaTiene = codObra ? 1 : 0;
      if (nuevaTiene > existTiene || (nuevaTiene === existTiene && partidasLen > existente.partidasLen)) {
        // La nueva es mejor — borrar la existente
        rowsToDelete.push(existente.idx + 1);
        Logger.log('Borrar fila '+(existente.idx+1)+': '+rows[existente.idx][0]+' (reemplazada por '+nombre+')');
        procesados[matchIdx] = { normNom: n, idx: i, codObra: codObra, partidasLen: partidasLen };
      } else {
        // La existente es mejor — borrar la nueva
        rowsToDelete.push(i + 1);
        Logger.log('Borrar fila '+(i+1)+': '+nombre+' (duplicado de '+rows[existente.idx][0]+')');
      }
    }
  }

  rowsToDelete.sort(function(a,b){ return b-a; });
  rowsToDelete.forEach(function(r){ sheet.deleteRow(r); });
  Logger.log('✅ Listo. Filas eliminadas: ' + rowsToDelete.length);
}

// ══════════════════════════════════════════════════════
// MIGRACIÓN NOMBRES STOCK — ejecutar UNA SOLA VEZ desde el editor
// ══════════════════════════════════════════════════════
function migrateInventarioNames() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  if (!sheet) { Logger.log('Hoja Inventario no encontrada'); return; }

  // Nombres a RENOMBRAR: clave = nombre actual exacto en la hoja, valor = nombre limpio nuevo
  var RENOMBRAR = {
    'BARANDILLA D/40 2500mm UNE EN13374 (P 169 U.)':                          'Barandilla D/35 2500mm',
    'BARANDILLA D/40 100-150 EN 13374 UNETRA (palé 169 uds)':                 'Barandilla D/35 2500mm',
    'BASE RECTA para Guardacuerpo UNETRA':                                      'Pie recto',
    'BASE PLASTICO CVTOOLS  POSTE LDV 60 X 40 UNETRA  (B.114 U)':             'Base plástico poste LDV 60x40',
    'BASQUIT / CONIX para guardacuerpo (bolsa 200 uds)':                       'Basquit / Conix bolsa 200 uds',
    'BRIDA   4,8X200 P100':                                                     'Brida 4,8x200',
    'ESLINGA PLATAFORMA - ANCLAJE VERTICAL DE TECHO 1,5MTS':                  'Eslinga plataforma 1,5m',
    'GUARDACUERPO 1,2M EN13374  (pale 250 ud)':                                'Guardacuerpo 1,200',
    'GUARDACUERPO 1,5M EN13374  (pale 250 ud)':                                'Guardacuerpo 1,500',
    'GUARDACUERPO SARGENTO MORDAZA 1.3m  - UNETRA  (Palet 120 Ud.)':          'Guardacuerpo usillo 1,300',
    'GUARDACUERPO C/MORDAZA 1.3m  (inde-k PRECIO ALQUILER DIA/UNIDAD)  codigo 108541 (Falta barandilla)': 'Guardacuerpo usillo 1,300 (inde-k)',
    'Guardacuerpo C/Mor. 1.3M INDE-K ref. 108541':                            'Guardacuerpo usillo 1,300 (inde-k)',
    'LINEA DE VIDA 20 ML HORIZONTAL- LUISA 300  (2 personas) - CON MOCHILA -':'Línea vida horizontal 20m',
    'MOSQUETON -  19mm-25KN':                                                   'Mosquetón 19mm 25KN',
    'MOSQUETON SEG.C/ROSCA 18MM (más de 20 uds) - ':                          'Mosquetón 19mm 25KN',
    'MOSQUITERA BLANCA 3X20 ML (34,20€ entera)':                              'Mosquitera blanca 3x20m',
    'POSTE / MASTIL LINEA DE VIDA   BASE 60*40':                               'Poste mástil LDV 60x40',
    'PUNTO DE ANCLAJE - AC02 (2 US) - oreja':                                  'Punto anclaje AC02 estela',
    'RED 1X10 PP BAJO FORJADO) UNE 81652 Q100 (NARANJA)':                     'Red 1x10 PP bajo forjado',
    'RED 1.10 X 10 BAJO FORJADO':                                               'Red 1x10 PP bajo forjado',
    'RED 2,10 X 10 PP (BAJO FORJADO ) UNE 81652 Q100 (NARANJA)':              'Red 2x10 PP bajo forjado',
    'RED 2.1 x 10 BF naranja':                                                  'Red 2x10 PP bajo forjado',
    'RED 3,5 X 10 PP EN1263 (NARANJA)':                                        'Red 3,5x10 PP EN1263',
    'Red 3 x 10 tipo U naranja':                                                'Red 3x10 PP tipo U',
    'RED 5 X 10 PP EN1263,1 UA2 Q100 5 x 10':                                 'Red 5x10 PP EN1263',
    'RODAPIE 2,5M  EN 13374  (P 100)':                                         'Rodapié 2,5m',
    'Rodapie Acero INDE-K ; L=250 cm ref. 116211':                            'Rodapié 2,5m (inde-k)',
    'SETA PROTECTORA 12-22 MM P300':                                            'Seta protectora 12-22mm',
    'TACO D12  LARGO  M10X100':                                                 'Taco D12 largo M10x100',
    'TACO D12 CORTO  M10X70  - 50 UDS':                                        'Taco D12 corto M10x70',
    'CANCAMO ARGOLLA D12- M10X70 - ACHA12C':                                   'Cáncamo argolla D12 M10x70',
    'CANCAMO ARGOLLA  D10  M 8X60 -  ACHA10C':                                 'Cáncamo argolla D10 M8x60',
    'TELA ONIX SEÑALIZACIÓN. (malla naranja) - 1 M X 50 M':                   'Tela señalización naranja 1x50m',
    'ZOCALO RAFIA   0,50 X 100 ML BLANCO':                                     'Zócalo rafia 0,50x100m',
    'zocalo - MOSQUITERA MALLA 90GR/M2  1 X 100 ML blanco':                   'Zócalo rafia 0,50x100m',
    'BASE CODO  DE 160 EN L  guardacuderpo  UNETRA':                          'Soporte 160 en L',
    'CINTA SEÑALIZACION ROJA/BLANCA 200MTS X 7CM':                            'Cinta señalización roja/blanca 200m',
    'BAUL 120 X 60 BRICOTRAIL':                                                 'Baúl 120x60',
    'MALLA CORPORATIVA 60% OCULTACION BLANCO/LOGO 2mts x 50 ml (min 3 rollos) - PRECIO M2': 'Malla corporativa 60% 2x50m',
    'BARANDILLA D/40 120 EN 13374  (P169 U)':                                  'Barandilla D/35 2500mm',
    'GANCHO RED BAJO FORJADO 8MM':                                              'Gancho red bajo forjado 8mm',
    '118861 Polea M15 INDE-K RED ; Inoxidable':                               'Polea M15 inde-k',
    'Rollo Cuerda Atado INDE-K RED ; Ø10mm ; 100 ml ; UNE':                   'Cuerda atado inde-k 100m'
  };

  var rows = sheet.getDataRange().getValues();
  var seenNombres = {}; // para detectar y eliminar duplicados tras renombrar
  var rowsToDelete = [];

  // Primera pasada: renombrar
  for (var i = 1; i < rows.length; i++) {
    var nombreActual = (rows[i][0] || '').toString().trim();
    if (!nombreActual) continue;
    var nuevoNombre = RENOMBRAR[nombreActual];
    if (nuevoNombre) {
      sheet.getRange(i + 1, 1).setValue(nuevoNombre);
      rows[i][0] = nuevoNombre; // actualizar en memoria
      Logger.log('Renombrado: "' + nombreActual + '" → "' + nuevoNombre + '"');
    }
  }

  // Segunda pasada: eliminar duplicados (quedarse con el de mayor stock)
  rows = sheet.getDataRange().getValues();
  var bestRow = {}; // nombre limpio → { rowIdx, stock }
  for (var j = 1; j < rows.length; j++) {
    var nom = (rows[j][0] || '').toString().trim();
    if (!nom) { rowsToDelete.push(j + 1); continue; }
    var stk = parseFloat(rows[j][1]) || 0;
    if (bestRow[nom] === undefined) {
      bestRow[nom] = { rowIdx: j, stock: stk };
    } else {
      // Ya existe — marcar el de menor stock para eliminar, conservar el mayor
      if (stk >= bestRow[nom].stock) {
        rowsToDelete.push(bestRow[nom].rowIdx + 1);
        bestRow[nom] = { rowIdx: j, stock: stk };
      } else {
        rowsToDelete.push(j + 1);
      }
      Logger.log('Duplicado eliminado fila ' + (j + 1) + ': ' + nom);
    }
  }

  // Eliminar de abajo hacia arriba para no desplazar índices
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(r) { sheet.deleteRow(r); });

  Logger.log('✅ Migración Inventario completada. ' + rowsToDelete.length + ' filas eliminadas.');
}

// ── Paso 2: migrar partidas de obras ─────────────────────────────
// Ejecutar DESPUÉS de migrateInventarioNames(). Actualiza partidas.material
// en todas las obras usando el mismo mapa de renombrado.
function migrateObraPartidaNames() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Obras');
  if (!sheet) { Logger.log('Hoja Obras no encontrada'); return; }

  // Mismo mapa que migrateInventarioNames — nombre antiguo → nombre limpio
  var RENOMBRAR = {
    'BARANDILLA D/40 2500mm UNE EN13374 (P 169 U.)':        'Barandilla D/35 2500mm',
    'BARANDILLA D/40 100-150 EN 13374 UNETRA (palé 169 uds)':'Barandilla D/35 2500mm',
    'BARANDILLA D/40 120 EN 13374  (P169 U)':                'Barandilla D/35 2500mm',
    'BARANDILLA D/40 2500mm':                                'Barandilla D/35 2500mm',
    'BARANDILLA D/35 100-150 EN 13374':                      'Barandilla D/35 2500mm',
    'BASE RECTA para Guardacuerpo UNETRA':                   'Pie recto',
    'BASE PLASTICO CVTOOLS  POSTE LDV 60 X 40 UNETRA  (B.114 U)': 'Base plástico poste LDV 60x40',
    'BASQUIT / CONIX para guardacuerpo (bolsa 200 uds)':     'Basquit / Conix bolsa 200 uds',
    'BRIDA   4,8X200 P100':                                  'Brida 4,8x200',
    'BRIDA 4,8X200 P100':                                    'Brida 4,8x200',
    'ESLINGA PLATAFORMA - ANCLAJE VERTICAL DE TECHO 1,5MTS': 'Eslinga plataforma 1,5m',
    'GUARDACUERPO 1,2M EN13374  (pale 250 ud)':              'Guardacuerpo 1,200',
    'GUARDACUERPO 1,2M EN13374 (pale 250 ud)':               'Guardacuerpo 1,200',
    'GUARDACUERPO 1,5M EN13374  (pale 250 ud)':              'Guardacuerpo 1,500',
    'GUARDACUERPO 1,5M EN13374 (pale 250 ud)':               'Guardacuerpo 1,500',
    'GUARDACUERPO SARGENTO MORDAZA 1.3m  - UNETRA  (Palet 120 Ud.)': 'Guardacuerpo usillo 1,300',
    'GUARDACUERPO SARGENTO MORDAZA 1.3m - UNETRA':           'Guardacuerpo usillo 1,300',
    'GUARDACUERPO SARGENTO MORDAZA 1.3M':                    'Guardacuerpo usillo 1,300',
    'GUARDACUERPO C/MORDAZA 1.3m  (inde-k PRECIO ALQUILER DIA/UNIDAD)  codigo 108541 (Falta barandilla)': 'Guardacuerpo usillo 1,300',
    'guardacuerpo usillo 1,300 doble cuerpo':                'Guardacuerpo usillo 1,300',
    'LINEA DE VIDA 20 ML HORIZONTAL- LUISA 300  (2 personas) - CON MOCHILA -': 'Línea vida horizontal 20m',
    'MOSQUETON -  19mm-25KN':                                'Mosquetón 19mm 25KN',
    'MOSQUETON SEG.C/ROSCA 18MM (más de 20 uds) - ':         'Mosquetón 19mm 25KN',
    'mosqueton seg. c/rosca 18mm':                           'Mosquetón 19mm 25KN',
    'MOSQUITERA BLANCA 3X20 ML (34,20€ entera)':            'Mosquitera blanca 3x20m',
    'POSTE / MASTIL LINEA DE VIDA   BASE 60*40':             'Poste mástil LDV 60x40',
    'PUNTO DE ANCLAJE - AC02 (2 US) - oreja':               'Punto anclaje AC02 estela',
    'punto anclaje a1 cesar 1,5m eslinga':                   'Eslinga plataforma 1,5m',
    'RED 1X10 PP BAJO FORJADO) UNE 81652 Q100 (NARANJA)':   'Red 1x10 PP bajo forjado',
    'RED 1.10 X 10 BAJO FORJADO':                            'Red 1x10 PP bajo forjado',
    'red pp en1263-1 1,10x10 bajo forjado':                  'Red 1x10 PP bajo forjado',
    'RED 2,10 X 10 PP (BAJO FORJADO ) UNE 81652 Q100 (NARANJA)': 'Red 2x10 PP bajo forjado',
    'RED 2.1 x 10 BF naranja':                              'Red 2x10 PP bajo forjado',
    'red 2x10 pp en1263 ua2 q100 (verde)':                  'Red 2x10 PP bajo forjado',
    'RED 3,5 X 10 PP EN1263 (NARANJA)':                     'Red 3,5x10 PP EN1263',
    'Red 3 x 10 tipo U naranja':                            'Red 3x10 PP tipo U',
    'RED 5 X 10 PP EN1263,1 UA2 Q100 5 x 10':              'Red 5x10 PP EN1263',
    'red pp en1263-1 va2 q100 5x10m (horca)':              'Red 5x10 PP EN1263',
    'RODAPIE 2,5M  EN 13374  (P 100)':                     'Rodapié 2,5m',
    'rodapie 2,5m. en 13374':                              'Rodapié 2,5m',
    'SETA PROTECTORA 12-22 MM P300':                        'Seta protectora 12-22mm',
    'seta protectora 12-24':                               'Seta protectora 12-22mm',
    'TACO D12  LARGO  M10X100':                             'Taco D12 largo M10x100',
    'taco d12 m-10x70':                                    'Taco D12 corto M10x70',
    'TACO D12 CORTO  M10X70  - 50 UDS':                    'Taco D12 corto M10x70',
    'taco argolla d10 m-8x60':                             'Cáncamo argolla D10 M8x60',
    'CANCAMO ARGOLLA D12- M10X70 - ACHA12C':               'Cáncamo argolla D12 M10x70',
    'CANCAMO ARGOLLA  D10  M 8X60 -  ACHA10C':             'Cáncamo argolla D10 M8x60',
    'TELA ONIX SEÑALIZACIÓN. (malla naranja) - 1 M X 50 M':'Tela señalización naranja 1x50m',
    'ZOCALO RAFIA   0,50 X 100 ML BLANCO':                 'Zócalo rafia 0,50x100m',
    'zocalo - MOSQUITERA MALLA 90GR/M2  1 X 100 ML blanco':'Zócalo rafia 0,50x100m',
    'rafia zocalo h 50 cms 100m':                          'Zócalo rafia 0,50x100m',
    'BASE CODO  DE 160 EN L  guardacuderpo  UNETRA':       'Soporte 160 en L',
    'soporte de 160 en l para guardacuerpo':               'Soporte 160 en L',
    'CINTA SEÑALIZACION ROJA/BLANCA 200MTS X 7CM':         'Cinta señalización roja/blanca 200m',
    'BAUL 120 X 60 BRICOTRAIL':                            'Baúl 120x60',
    'anclaje m 10x100 d/12 tornillo (taco d12 m10x100) 50uds': 'Taco D12 largo M10x100',
    'PIE RECTO':                                           'Pie recto',
    'pie recto':                                           'Pie recto',
    'SOPORTE DE 160 EN L PARA GUARDACUERPO':               'Soporte 160 en L'
  };

  var rows = sheet.getDataRange().getValues();
  var obrasActualizadas = 0;
  var partidasActualizadas = 0;

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var partidas;
    try { partidas = JSON.parse(rows[i][7] || '[]'); } catch(e) { continue; }
    var modificado = false;
    partidas.forEach(function(p) {
      var mat = (p.material || '').trim();
      // Buscar coincidencia exacta primero
      var nuevo = RENOMBRAR[mat] || RENOMBRAR[mat.toUpperCase()];
      // Si no hay exacta, buscar insensible a mayúsculas
      if (!nuevo) {
        var matLow = mat.toLowerCase();
        for (var k in RENOMBRAR) {
          if (k.toLowerCase() === matLow) { nuevo = RENOMBRAR[k]; break; }
        }
      }
      if (nuevo && nuevo !== mat) {
        p.material = nuevo;
        modificado = true;
        partidasActualizadas++;
      }
    });
    if (modificado) {
      sheet.getRange(i + 1, 8).setValue(JSON.stringify(partidas));
      obrasActualizadas++;
      Logger.log('Obra actualizada: ' + rows[i][0] + ' — ' + partidas.filter(function(p){return RENOMBRAR[p.material];}).length + ' partidas');
    }
  }

  Logger.log('✅ Partidas migradas. Obras: ' + obrasActualizadas + ', Partidas: ' + partidasActualizadas);
}

// ══════════════════════════════════════════════════════
// FACTURAS EMITIDAS
// ══════════════════════════════════════════════════════
function getFacturas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Facturas');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id:          rows[i][0],
      numero:      rows[i][1] || '',
      tipo:        rows[i][2] || 'factura',
      fecha:       rows[i][3] || '',
      vencimiento: rows[i][4] || '',
      cliente:     rows[i][5] || '',
      cifCliente:  rows[i][6] || '',
      dirCliente:  rows[i][7] || '',
      obra:        rows[i][8] || '',
      lineas:      JSON.parse(rows[i][9]  || '[]'),
      base:        rows[i][10] || 0,
      pctIva:      rows[i][11] || 21,
      iva:         rows[i][12] || 0,
      total:       rows[i][13] || 0,
      estado:      rows[i][14] || 'pendiente',
      condiciones: rows[i][15] || '',
      obs:         rows[i][16] || '',
      driveUrl:    rows[i][17] || ''
    });
  }
  return result;
}

function saveFactura(data) {
  var f = data.factura || data;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Facturas');
  if (!sheet) {
    sheet = ss.insertSheet('Facturas');
    sheet.appendRow(['ID','Numero','Tipo','Fecha','Vencimiento','Cliente','CIF','Direccion','Obra','Lineas','Base','PctIVA','IVA','Total','Estado','Condiciones','Obs','DriveUrl']);
  }
  var rowData = [
    f.id, f.numero||'', f.tipo||'factura',
    f.fecha||'', f.vencimiento||'',
    f.cliente||'', f.cifCliente||'', f.dirCliente||'',
    f.obra||'', JSON.stringify(f.lineas||[]),
    f.base||0, f.pctIva||21, f.iva||0, f.total||0,
    f.estado||'pendiente', f.condiciones||'', f.obs||'', f.driveUrl||''
  ];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === f.id) {
      sheet.getRange(i+1,1,1,18).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteFactura(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Facturas');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// GASTOS (FACTURAS DE PROVEEDORES)
// ══════════════════════════════════════════════════════
function getGastos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Gastos');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id:          rows[i][0],
      proveedor:   rows[i][1] || '',
      numFac:      rows[i][2] || '',
      fecha:       rows[i][3] || '',
      vencimiento: rows[i][4] || '',
      obra:        rows[i][5] || '',
      concepto:    rows[i][6] || '',
      base:        rows[i][7] || 0,
      pctIva:      rows[i][8] || 21,
      iva:         rows[i][9] || 0,
      total:       rows[i][10] || 0,
      estado:      rows[i][11] || 'pendiente',
      formaPago:   rows[i][12] || '',
      driveUrl:    rows[i][13] || ''
    });
  }
  return result;
}

function saveGasto(data) {
  var g = data.gasto || data;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Gastos');
  if (!sheet) {
    sheet = ss.insertSheet('Gastos');
    sheet.appendRow(['ID','Proveedor','NumFac','Fecha','Vencimiento','Obra','Concepto','Base','PctIVA','IVA','Total','Estado','FormaPago','DriveUrl']);
  }
  var rowData = [
    g.id, g.proveedor||'', g.numFac||'',
    g.fecha||'', g.vencimiento||'', g.obra||'',
    g.concepto||'', g.base||0, g.pctIva||21,
    g.iva||0, g.total||0, g.estado||'pendiente', g.formaPago||'', g.driveUrl||''
  ];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === g.id) {
      sheet.getRange(i+1,1,1,14).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteGasto(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Gastos');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// COBROS
// ══════════════════════════════════════════════════════
function getCobros() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cobros');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id:        rows[i][0],
      facturaId: rows[i][1] || '',
      numero:    rows[i][2] || '',
      cliente:   rows[i][3] || '',
      importe:   rows[i][4] || 0,
      fecha:     rows[i][5] || '',
      formaPago: rows[i][6] || '',
      obs:       rows[i][7] || ''
    });
  }
  return result;
}

function saveCobro(data) {
  var c = data.cobro || data;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cobros');
  if (!sheet) {
    sheet = ss.insertSheet('Cobros');
    sheet.appendRow(['ID','FacturaID','Numero','Cliente','Importe','Fecha','FormaPago','Obs']);
  }
  var rowData = [
    c.id, c.facturaId||'', c.numero||'',
    c.cliente||'', c.importe||0, c.fecha||'',
    c.formaPago||'', c.obs||''
  ];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === c.id) {
      sheet.getRange(i+1,1,1,8).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteCobro(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cobros');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// CLIENTES
// ══════════════════════════════════════════════════════
function getClientes() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Clientes');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      nombre:    rows[i][0],
      cif:       rows[i][1] || '',
      dir:       rows[i][2] || '',
      email:     rows[i][3] || '',
      tel:       rows[i][4] || '',
      contacto:  rows[i][5] || '',
      formaPago: rows[i][6] || '',
      plazo:     rows[i][7] || '',
      iban:      rows[i][8] || '',
      obs:       rows[i][9] || ''
    });
  }
  return result;
}

function saveCliente(data) {
  var c = data.cliente || data;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Clientes');
  if (!sheet) {
    sheet = ss.insertSheet('Clientes');
    sheet.appendRow(['Nombre','CIF','Direccion','Email','Telefono','Contacto','FormaPago','Plazo','IBAN','Obs']);
  }
  var rowData = [c.nombre, c.cif||'', c.dir||'', c.email||'', c.tel||'', c.contacto||'', c.formaPago||'', c.plazo||'', c.iban||'', c.obs||''];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === c.nombreOriginal || rows[i][0] === c.nombre) {
      sheet.getRange(i+1,1,1,10).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteCliente(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Clientes');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombre) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// PROVEEDORES
// ══════════════════════════════════════════════════════
function getProveedores() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Proveedores');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      nombre:    rows[i][0],
      cif:       rows[i][1] || '',
      dir:       rows[i][2] || '',
      email:     rows[i][3] || '',
      tel:       rows[i][4] || '',
      contacto:  rows[i][5] || '',
      formaPago: rows[i][6] || '',
      plazo:     rows[i][7] || '',
      iban:      rows[i][8] || '',
      categoria: rows[i][9] || '',
      dto:       rows[i][10] || 0,
      obs:       rows[i][11] || ''
    });
  }
  return result;
}

function saveProveedor(data) {
  var p = data.proveedor || data;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Proveedores');
  if (!sheet) {
    sheet = ss.insertSheet('Proveedores');
    sheet.appendRow(['Nombre','CIF','Direccion','Email','Telefono','Contacto','FormaPago','Plazo','IBAN','Categoria','Dto','Obs']);
  }
  var rowData = [p.nombre, p.cif||'', p.dir||'', p.email||'', p.tel||'', p.contacto||'', p.formaPago||'', p.plazo||'', p.iban||'', p.categoria||'', p.dto||0, p.obs||''];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.nombreOriginal || rows[i][0] === p.nombre) {
      sheet.getRange(i+1,1,1,12).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteProveedor(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Proveedores');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombre) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// GASTOS INTERNOS
// ══════════════════════════════════════════════════════
function getGastosInternos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('GastosInternos');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id:         rows[i][0],
      categoria:  rows[i][1] || '',
      concepto:   rows[i][2] || '',
      proveedor:  rows[i][3] || '',
      fecha:      rows[i][4] || '',
      importe:    rows[i][5] || 0,
      base:       rows[i][5] || 0,
      recurrente: rows[i][6] || '',
      formaPago:  rows[i][7] || '',
      obs:        rows[i][8] || '',
      obra:       rows[i][9] || '',
      creado:     rows[i][10] || '',
      driveUrl:   rows[i][11] || ''
    });
  }
  return result;
}

function saveGastoInterno(data) {
  var g = data.gasto || data;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('GastosInternos');
  if (!sheet) {
    sheet = ss.insertSheet('GastosInternos');
    sheet.appendRow(['ID','Categoria','Concepto','Proveedor','Fecha','Importe','Recurrente','FormaPago','Obs','Obra','Creado','DriveUrl']);
  }
  var rowData = [
    g.id, g.categoria||'', g.concepto||'', g.proveedor||'',
    g.fecha||'', g.importe||g.base||0,
    g.recurrente||'', g.formaPago||'', g.obs||'', g.obra||'', g.creado||'', g.driveUrl||''
  ];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === g.id) {
      sheet.getRange(i+1,1,1,12).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteGastoInterno(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('GastosInternos');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// OCR — LECTURA DE FACTURAS CON GOOGLE DRIVE
// ══════════════════════════════════════════════════════
function ocrFactura(data) {
  try {
    var decoded = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(decoded, data.mimeType, 'factura_ocr');
    var resource = { title: 'factura_ocr_tmp', mimeType: 'application/vnd.google-apps.document' };
    var options = { ocr: true, ocrLanguage: 'es' };
    var file = Drive.Files.insert(resource, blob, options);
    var doc = DocumentApp.openById(file.id);
    var texto = doc.getBody().getText();
    Drive.Files.remove(file.id);
    return { ok: true, texto: texto };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════
// GOOGLE DRIVE — GUARDADO DE DOCUMENTOS
// ══════════════════════════════════════════════════════
var CARPETA_RAIZ_ID = '1VOzHctjI9aOeeQxxRXrpqB3caEIxu88l'; // Carpeta ALMACEN APP en Drive compartido SEGUFIJA

function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function saveFileToDrive(data) {
  try {
    var current = DriveApp.getFolderById(CARPETA_RAIZ_ID);
    for (var i = 0; i < (data.carpeta || []).length; i++) {
      current = getOrCreateFolder(current, data.carpeta[i]);
    }
    var decoded = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(decoded, data.mimeType, data.fileName);
    var file = current.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { ok: true, url: file.getUrl(), id: file.getId(), nombre: data.fileName };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ── DOCUMENTOS PERSONAL ──
function getDocsPersonal(trabajador) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DocsPersonal');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (trabajador && rows[i][1] !== trabajador) continue;
    result.push({
      id:         rows[i][0],
      trabajador: rows[i][1] || '',
      tipo:       rows[i][2] || '',
      nombre:     rows[i][3] || '',
      driveUrl:   rows[i][4] || '',
      fecha:      rows[i][5] || '',
      obs:        rows[i][6] || ''
    });
  }
  return result;
}

function saveDocPersonal(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DocsPersonal');
  if (!sheet) {
    sheet = ss.insertSheet('DocsPersonal');
    sheet.appendRow(['ID','Trabajador','Tipo','Nombre','DriveUrl','Fecha','Obs']);
  }
  var rowData = [data.id, data.trabajador||'', data.tipo||'', data.nombre||'', data.driveUrl||'', data.fecha||'', data.obs||''];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i+1,1,1,7).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteDocPersonal(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DocsPersonal');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ── DOCUMENTOS GESTORÍA ──
function getDocsGestoria() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Gestoria');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id:       rows[i][0],
      tipo:     rows[i][1] || '',
      nombre:   rows[i][2] || '',
      driveUrl: rows[i][3] || '',
      fecha:    rows[i][4] || '',
      obs:      rows[i][5] || ''
    });
  }
  return result;
}

function saveDocGestoria(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Gestoria');
  if (!sheet) {
    sheet = ss.insertSheet('Gestoria');
    sheet.appendRow(['ID','Tipo','Nombre','DriveUrl','Fecha','Obs']);
  }
  var rowData = [data.id, data.tipo||'', data.nombre||'', data.driveUrl||'', data.fecha||'', data.obs||''];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i+1,1,1,6).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteDocGestoria(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Gestoria');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ── OPERARIOS ──
function getOperarios() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Operarios');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      nombre:    rows[i][0] || '',
      dni:       rows[i][1] || '',
      categoria: rows[i][2] || '',
      coste:     rows[i][3] || 0,
      tel:       rows[i][4] || '',
      email:     rows[i][5] || '',
      alta:      rows[i][6] || '',
      baja:      rows[i][7] || null,
      obs:       rows[i][8] || ''
    });
  }
  return result;
}

function saveOperario(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Operarios');
  if (!sheet) {
    sheet = ss.insertSheet('Operarios');
    sheet.appendRow(['Nombre','DNI','Categoria','Coste','Tel','Email','Alta','Baja','Obs']);
  }
  var op = data.operario;
  var rowData = [op.nombre||'', op.dni||'', op.categoria||'', op.coste||0, op.tel||'', op.email||'', op.alta||'', op.baja||'', op.obs||''];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === op.nombre) {
      sheet.getRange(i+1,1,1,9).setValues([rowData]);
      return {ok:true};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true};
}

function deleteOperario(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Operarios');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombre) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ── DOCUMENTOS OBRA ──
function getDocsObra(obra) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DocsObras');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (obra && rows[i][1] !== obra) continue;
    result.push({
      id:       rows[i][0],
      obra:     rows[i][1] || '',
      tipo:     rows[i][2] || '',
      nombre:   rows[i][3] || '',
      driveUrl: rows[i][4] || '',
      fecha:    rows[i][5] || '',
      obs:      rows[i][6] || ''
    });
  }
  return result;
}

function saveDocObra(data) {
  // Guarda el archivo en Drive: SEGUFIJA/Obras/[Obra]/[Tipo]/
  var driveRes = saveFileToDrive({
    base64:   data.base64,
    mimeType: data.mimeType,
    fileName: data.fileName,
    carpeta:  ['Obras', data.obra, data.tipo]
  });
  if (!driveRes.ok) return driveRes;

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DocsObras');
  if (!sheet) {
    sheet = ss.insertSheet('DocsObras');
    sheet.appendRow(['ID','Obra','Tipo','Nombre','DriveUrl','Fecha','Obs']);
  }
  var rowData = [data.id, data.obra||'', data.tipo||'', data.nombre||'', driveRes.url, data.fecha||'', data.obs||''];
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i+1,1,1,7).setValues([rowData]);
      return {ok:true, url: driveRes.url};
    }
  }
  sheet.appendRow(rowData);
  return {ok:true, url: driveRes.url};
}

function deleteDocObra(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DocsObras');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// CRM OBRAS
// ══════════════════════════════════════════════════════
var CRM_COLS = ['oferta','constructora','obra','estado','prioridad','importe','ofertaEnviada','contacto','telefono','email','ultimaAccion','proximaAccion','fechaRec','carpetaDriveId'];

function getCRM() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('CRM_Obras');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] && !rows[i][1]) continue;
    var o = {};
    for (var c = 0; c < CRM_COLS.length; c++) o[CRM_COLS[c]] = rows[i][c] !== undefined ? String(rows[i][c]) : '';
    o.importe = parseFloat(rows[i][5]) || 0;
    o._row = i + 1;
    result.push(o);
  }
  return result;
}

function saveCRM(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('CRM_Obras');
  if (!sheet) {
    sheet = ss.insertSheet('CRM_Obras');
    sheet.appendRow(CRM_COLS);
    sheet.getRange(1,1,1,CRM_COLS.length).setFontWeight('bold').setBackground('#1a3353').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  var rowData = CRM_COLS.map(function(k){ return data[k] !== undefined ? data[k] : ''; });
  rowData[5] = parseFloat(data.importe) || 0;
  if (data._row && data._row > 1) {
    sheet.getRange(data._row, 1, 1, CRM_COLS.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return {ok:true};
}

function deleteCRM(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('CRM_Obras');
  if (!sheet) return {ok:false};
  if (data._row && data._row > 1) {
    sheet.deleteRow(data._row);
    return {ok:true};
  }
  return {ok:false};
}

// ══════════════════════════════════════════════════════
// UTILIDADES
// ══════════════════════════════════════════════════════
function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════
// ENVÍO EMAIL PEDIDO PROVEEDOR
// ══════════════════════════════════════════════════════
function renameObraHistorial(data) {
  var oldObra = (data.oldObra||'').trim();
  var newObra = (data.newObra||'').trim();
  if (!oldObra || !newObra) return {ok:false, msg:'Faltan parámetros'};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Historial');
  if (!sheet) return {ok:false, msg:'No hay hoja Historial'};
  var rows = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][2]||'').trim().toLowerCase() === oldObra.toLowerCase()) {
      sheet.getRange(i+1, 3).setValue(newObra);
      count++;
    }
  }
  return {ok:true, updated:count};
}

function sendEmailPedido(data) {
  try {
    var email = data.email || "";
    if (!email) return {ok: false, error: "Email no especificado"};
    var prov = data.prov || "Proveedor";
    var fecha = data.fecha || "";
    var obs = data.obs || "";
    var items = data.items || [];
    var NL = String.fromCharCode(10);
    var lineas = items.map(function(it) { return "  - " + it.nombre + ": " + it.qty + " ud"; }).join(NL);
    var subject = "Pedido SEGUFIJA - " + prov.split(" ")[0] + " - " + fecha;
    var body = "Estimados," + NL + NL;
    body += "Les enviamos el pedido de reposicion:" + NL + NL;
    body += "Proveedor: " + prov + NL;
    body += "Fecha: " + fecha + NL;
    if (obs) body += "Observaciones: " + obs + NL;
    body += NL + "Material:" + NL + lineas + NL + NL;
    body += "Muchas gracias," + NL + "Segufija SL" + NL + "Tel: 623782259 | info@segufija.com";
    var htmlBody = "<p>Estimados,</p><p>Pedido de reposicion:</p>";
    htmlBody += "<table style=\"border-collapse:collapse;width:100%;max-width:500px\">";
    htmlBody += "<tr><td><b>Proveedor</b></td><td>" + prov + "</td></tr>";
    htmlBody += "<tr><td><b>Fecha</b></td><td>" + fecha + "</td></tr>";
    if (obs) htmlBody += "<tr><td><b>Obs</b></td><td>" + obs + "</td></tr>";
    htmlBody += "</table><h3>Material solicitado:</h3><ul>";
    items.forEach(function(it) { htmlBody += "<li>" + it.nombre + ": <b>" + it.qty + " ud</b></li>"; });
    htmlBody += "</ul><p>Por favor confirmen recepcion.</p><p><b>Segufija SL</b> | Tel: 623782259</p>";
    GmailApp.sendEmail(email, subject, body, {htmlBody: htmlBody, name: "Segufija SL"});
    return {ok: true};
  } catch(e) {
    return {ok: false, error: e.message};
  }
}

// ══════════════════════════════════════════════════════
// GUARDAR ARCHIVO EN CARPETA DE DRIVE
// ══════════════════════════════════════════════════════
function saveFileToFolder(data) {
  try {
    var folderId = data.folderId || '';
    var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getFolderById(CARPETA_RAIZ_ID);
    var decoded = Utilities.base64Decode(data.content);
    var blob = Utilities.newBlob(decoded, data.mimeType || 'text/html', data.fileName);
    // Eliminar archivo anterior con mismo nombre si existe
    var existing = folder.getFilesByName(data.fileName);
    while (existing.hasNext()) { existing.next().setTrashed(true); }
    var file = folder.createFile(blob);
    return {ok: true, fileUrl: file.getUrl(), fileId: file.getId()};
  } catch(e) {
    return {ok: false, error: e.message};
  }
}

// ══════════════════════════════════════════════════════
// PRESUPUESTOS
// ══════════════════════════════════════════════════════
var PRES_COLS = ['numOferta','proyecto','cliente','estado','fechaOferta','importe','data'];

function getPresupuestos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Presupuestos');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] && !rows[i][1]) continue;
    var obj = {};
    for (var c = 0; c < PRES_COLS.length; c++) obj[PRES_COLS[c]] = rows[i][c] !== undefined ? String(rows[i][c]) : '';
    obj.importe = parseFloat(rows[i][5]) || 0;
    obj._row = i + 1;
    if (obj.data) {
      try { var parsed = JSON.parse(obj.data); Object.assign(obj, parsed); obj._row = i + 1; } catch(e2) {}
    }
    result.push(obj);
  }
  return result;
}

function savePresupuesto(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Presupuestos');
  if (!sheet) {
    sheet = ss.insertSheet('Presupuestos');
    sheet.appendRow(PRES_COLS);
    sheet.getRange(1,1,1,PRES_COLS.length).setFontWeight('bold').setBackground('#1a3353').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  var presData = data.presupuesto || data;
  var rowData = [
    presData.numOferta || data.numOferta || '',
    presData.proyecto  || data.proyecto  || '',
    presData.cliente   || data.cliente   || '',
    presData.estado    || data.estado    || '',
    presData.fechaOferta || data.fechaOferta || '',
    parseFloat(presData.importe || data.importe) || 0,
    data.data || JSON.stringify(presData)
  ];
  var _row = parseInt(data._row) || 0;
  if (_row > 1) {
    sheet.getRange(_row, 1, 1, PRES_COLS.length).setValues([rowData]);
    return {ok: true, _row: _row};
  }
  var numOferta = presData.numOferta || data.numOferta || '';
  if (numOferta) {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(numOferta).trim()) {
        sheet.getRange(i + 1, 1, 1, PRES_COLS.length).setValues([rowData]);
        return {ok: true, _row: i + 1};
      }
    }
  }
  sheet.appendRow(rowData);
  var newRow = sheet.getLastRow();
  return {ok: true, _row: newRow};
}

function deletePresupuesto(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Presupuestos');
  if (!sheet) return {ok: false};
  var _row = parseInt(data._row) || 0;
  if (_row > 1) {
    sheet.deleteRow(_row);
    return {ok: true};
  }
  return {ok: false};
}

// ══════════════════════════════════════════════════════
// PROFORMAS / CERTIFICACIONES
// ══════════════════════════════════════════════════════
function getProformas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Proformas');
  if (!sheet) return {};
  var rows = sheet.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < rows.length; i++) {
    var obraKey = String(rows[i][0] || '').trim();
    if (!obraKey) continue;
    try { result[obraKey] = JSON.parse(rows[i][1] || '[]'); } catch(e) { result[obraKey] = []; }
  }
  return result;
}

function saveProforma(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Proformas');
  if (!sheet) {
    sheet = ss.insertSheet('Proformas');
    sheet.appendRow(['obraKey','data','fecha']);
    sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a3353').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  var obraKey = data.obraKey || '';
  var proformasStr = JSON.stringify(data.proformas || []);
  var fecha = new Date().toISOString().slice(0,10);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(obraKey).trim()) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[obraKey, proformasStr, fecha]]);
      return {ok: true};
    }
  }
  sheet.appendRow([obraKey, proformasStr, fecha]);
  return {ok: true};
}

// ══════════════════════════════════════════════════════
// COSTES DE MATERIALES POR OBRA (historial × inventario)
// ══════════════════════════════════════════════════════
function getCostesMateriales() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Mapa nombre → pcompra desde Inventario
  var invSheet = ss.getSheetByName('Inventario');
  var precioMap = {};
  if (invSheet) {
    var invRows = invSheet.getDataRange().getValues();
    for (var j = 1; j < invRows.length; j++) {
      var nombre = String(invRows[j][0] || '').trim();
      var pcompra = parseFloat(invRows[j][4]) || 0;
      if (nombre) precioMap[nombre.toLowerCase()] = pcompra;
    }
  }

  // Leer historial y sumar coste por obra
  var hSheet = ss.getSheetByName('Historial');
  var result = {}; // { obraKey: { coste, unidades } }
  if (!hSheet) return result;
  var hRows = hSheet.getDataRange().getValues();
  var TIPOS_SALIDA = ['salida', 'entrega'];
  for (var i = 1; i < hRows.length; i++) {
    var tipo   = String(hRows[i][0] || '').toLowerCase();
    var obra   = String(hRows[i][2] || '').trim();
    var anulada = String(hRows[i][8] || '');
    if (!obra || obra === 'ALMACEN') continue;
    if (anulada === 'true' || anulada === 'anulada') continue;
    if (TIPOS_SALIDA.indexOf(tipo) === -1) continue;
    var items = [];
    try { items = JSON.parse(hRows[i][7] || '[]'); } catch(e2) {}
    if (!result[obra]) result[obra] = { coste: 0, unidades: 0 };
    items.forEach(function(it) {
      var qty  = parseFloat(it.qty || it.cantidad || 0);
      var p    = parseFloat(it.pcompra || it.precio || 0);
      if (!p) p = precioMap[(it.nombre || '').toLowerCase().trim()] || 0;
      result[obra].coste    += qty * p;
      result[obra].unidades += qty;
    });
  }
  return result;
}
