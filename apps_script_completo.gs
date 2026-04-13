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
  if (action === 'getDocsPersonal')  return json(getDocsPersonal(e.parameter.trabajador));
  if (action === 'getDocsGestoria')  return json(getDocsGestoria());
  if (action === 'getDocsObra')      return json(getDocsObra(e.parameter.obra));
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
      stockMin: rows[i][5] || 0
    });
  }
  return result;
}

function asegurarColumnas(sheet) {
  var h = sheet.getRange(1,1,1,6).getValues()[0];
  if (!h[0]) sheet.getRange(1,1).setValue('Nombre');
  if (!h[1]) sheet.getRange(1,2).setValue('Stock');
  if (!h[2]) sheet.getRange(1,3).setValue('Necesario');
  if (!h[3]) sheet.getRange(1,4).setValue('PVP');
  if (!h[4]) sheet.getRange(1,5).setValue('PrecioCompra');
  if (!h[5]) sheet.getRange(1,6).setValue('StockMinimo');
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
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      nombre:           rows[i][0],
      lugar:            rows[i][1] || '—',
      presupuesto:      rows[i][2] || 0,
      inicio:           rows[i][3] || '',
      fin:              rows[i][4] || '',
      estado:           rows[i][5] || 'activa',
      obs:              rows[i][6] || '',
      partidas:         JSON.parse(rows[i][7]  || '[]'),
      personal:         JSON.parse(rows[i][8]  || '[]'),
      pctAlquiler:      rows[i][9]  || 0,
      diasAlquiler:     rows[i][10] || 0,
      previsionPersonal: JSON.parse(rows[i][11] || '[]'),
      tipo:             rows[i][12] || 'fija',
      ml:               rows[i][13] || 0,
      meses:            rows[i][14] || 0,
      formaPago:        rows[i][15] || '',
      plazo:            rows[i][16] || '',
      contacto:         JSON.parse(rows[i][17] || '{}')
    });
  }
  return result;
}

function saveObra(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Obras');
  if (!sheet) {
    sheet = ss.insertSheet('Obras');
    sheet.appendRow(['Nombre','Lugar','Presupuesto','Inicio','Fin','Estado','Obs','Partidas','Personal','PctAlquiler','DiasAlquiler','PrevisionPersonal','Tipo','ML','Meses','FormaPago','Plazo','Contacto']);
  }
  var rows = sheet.getDataRange().getValues();
  var rowData = [
    data.nombre, data.lugar||'—', data.presupuesto||0,
    data.inicio||'', data.fin||'', data.estado||'activa',
    data.obs||'', JSON.stringify(data.partidas||[]),
    JSON.stringify(data.personal||[]), data.pctAlquiler||0,
    data.diasAlquiler||0, JSON.stringify(data.previsionPersonal||[]),
    data.tipo||'fija', data.ml||0, data.meses||0,
    data.formaPago||'', data.plazo||'', JSON.stringify(data.contacto||{})
  ];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombreOriginal || rows[i][0] === data.nombre) {
      sheet.getRange(i+1,1,1,18).setValues([rowData]);
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
  sheet.appendRow([data.nombre,data.stock||0,data.nec||0,data.pvp||0,data.pcompra||0,data.stockMin||0]);
  return {ok:true};
}

function editArticulo(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Inventario');
  if (!sheet) return {ok:false};
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.nombreOriginal) {
      sheet.getRange(i+1,1).setValue(data.nombre);
      sheet.getRange(i+1,2).setValue(data.stock||0);
      sheet.getRange(i+1,3).setValue(data.nec||0);
      sheet.getRange(i+1,4).setValue(data.pvp||0);
      sheet.getRange(i+1,5).setValue(data.pcompra||0);
      sheet.getRange(i+1,6).setValue(data.stockMin||0);
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
// UTILIDADES
// ══════════════════════════════════════════════════════
function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
