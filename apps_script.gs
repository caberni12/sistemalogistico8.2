/* ==================================================
   CONFIGURACIÓN GENERAL
================================================== */
const SPREADSHEET_ID = '1Pwlf0YYZrGYTU5rAZbGQc3PpS_v7Bus-2V9owIN78-o';
const SHEET_MAESTRA = 'MAESTRA';
const SHEET_MOV     = 'MOVIMIENTO';
// Nueva BD para pedidos de cliente
const SHEET_PEDIDOS = 'PEDIDOS';
// BD para ubicaciones de productos
const SHEET_UBICACIONES = 'UBICACIONES';
// BD para movimientos de productos en ubicaciones con referencia a pedido
const SHEET_MOV_UBI = 'MOVIMIENTO_UBICACION';
// Nueva BD para movimientos de pedidos (registro de cambios de estado)
const SHEET_MOV_PEDIDOS = 'MOVIMIENTO_PEDIDOS';
// Callback temporal para permitir respuestas JSONP desde doGet y evitar bloqueos CORS en navegador
var JSONP_CALLBACK = '';

/* ==================================================
   INIT LIGERO
================================================== */
function initSistema(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let shM = ss.getSheetByName(SHEET_MAESTRA);
  if(!shM){
    shM = ss.insertSheet(SHEET_MAESTRA);
    shM.appendRow(['codigo','descripcion']);
    shM.getRange('A:A').setNumberFormat('@');
  }else if(shM.getLastRow() === 0){
    shM.appendRow(['codigo','descripcion']);
    shM.getRange('A:A').setNumberFormat('@');
  }

  let shV = ss.getSheetByName(SHEET_MOV);
  if(!shV){
    shV = ss.insertSheet(SHEET_MOV);
    shV.appendRow([
      'id',
      'fecha_registro',
      'fecha_entrada',
      'fecha_salida',
      'ubicacion',
      'codigo',
      'descripcion',
      'cantidad',
      'responsable',
      'status',
      'origen'
    ]);
    shV.getRange('F:F').setNumberFormat('@');
  }else if(shV.getLastRow() === 0){
    shV.appendRow([
      'id',
      'fecha_registro',
      'fecha_entrada',
      'fecha_salida',
      'ubicacion',
      'codigo',
      'descripcion',
      'cantidad',
      'responsable',
      'status',
      'origen'
    ]);
    shV.getRange('F:F').setNumberFormat('@');
  }

    // Crear BD de pedidos si no existe
  let shP = ss.getSheetByName(SHEET_PEDIDOS);
  if(!shP){
    shP = ss.insertSheet(SHEET_PEDIDOS);
    // Encabezados: id, codigo, descripcion, ubicacion, pedido, cliente, vendedor, status, fecha
    // La fecha se mantiene como columna propia para que llegue al HTML y al PDF.
    shP.appendRow(['id','codigo','descripcion','ubicacion','pedido','cliente','vendedor','status','fecha']);
    // Establecer formato de texto para columnas críticas y preservar ceros a la izquierda.
    shP.getRange('B:B').setNumberFormat('@');
    shP.getRange('E:I').setNumberFormat('@');
  } else if(shP.getLastRow() === 0){
    shP.appendRow(['id','codigo','descripcion','ubicacion','pedido','cliente','vendedor','status','fecha']);
    shP.getRange('B:B').setNumberFormat('@');
    shP.getRange('E:I').setNumberFormat('@');
  }

  // Si la BD-PEDIDOS existe pero no tiene columna de estado, agregarla
  const header = shP.getRange(1,1,1,shP.getLastColumn()).getDisplayValues()[0];
  if(header.indexOf('status') === -1){
    const lastCol = shP.getLastColumn();
    shP.insertColumnAfter(lastCol);
    // Write header and default status
    shP.getRange(1, lastCol+1).setValue('status');
    if(shP.getLastRow() > 1){
      const numRows = shP.getLastRow() - 1;
      const valores = new Array(numRows).fill(['pendiente']);
      shP.getRange(2, lastCol+1, numRows, 1).setValues(valores);
    }
    // Format entire status column as text to preserve values
    shP.getRange(1, lastCol+1, shP.getMaxRows(), 1).setNumberFormat('@');
  }

  // Asegurar que PEDIDOS siempre tenga columna fecha y formatos de texto.
  asegurarColumnas_(shP, ['id','codigo','descripcion','ubicacion','pedido','cliente','vendedor','status','fecha']);
  shP.getRange('B:B').setNumberFormat('@');
  shP.getRange('E:I').setNumberFormat('@');

  // Crear BD de ubicaciones si no existe
  let shU = ss.getSheetByName(SHEET_UBICACIONES);
  if(!shU){
    shU = ss.insertSheet(SHEET_UBICACIONES);
    // Cabeceras: código, descripción, cantidad y ubicación
    shU.appendRow(['codigo','descripcion','cantidad','ubicacion']);
    // Aseguramos que la columna de código sea texto
    shU.getRange('A:A').setNumberFormat('@');
  } else if(shU.getLastRow() === 0){
    shU.appendRow(['codigo','descripcion','cantidad','ubicacion']);
    shU.getRange('A:A').setNumberFormat('@');
  }

  // Crear BD de movimientos de ubicaciones si no existe
  let shMU = ss.getSheetByName(SHEET_MOV_UBI);
  if(!shMU){
    shMU = ss.insertSheet(SHEET_MOV_UBI);
    shMU.appendRow(['id','fecha','codigo','descripcion','cantidad','ubicacion','pedido','cliente','vendedor','status','origen']);
    shMU.getRange('C:C').setNumberFormat('@');
  } else if(shMU.getLastRow() === 0){
    shMU.appendRow(['id','fecha','codigo','descripcion','cantidad','ubicacion','pedido','cliente','vendedor','status','origen']);
    shMU.getRange('C:C').setNumberFormat('@');
  } else {
    // Asegurar columnas nuevas sin borrar información existente
    const headersMU = shMU.getRange(1,1,1,shMU.getLastColumn()).getDisplayValues()[0].map(h => txt(h).toLowerCase());
    const requiredMU = ['id','fecha','codigo','descripcion','cantidad','ubicacion','pedido','cliente','vendedor','status','origen'];
    requiredMU.forEach(h => {
      if(headersMU.indexOf(h) === -1){
        shMU.getRange(1, shMU.getLastColumn()+1).setValue(h);
        headersMU.push(h);
      }
    });
  }
  shMU.getRange('C:C').setNumberFormat('@');

  // Crear BD de movimientos de pedidos si no existe
  let shMP = ss.getSheetByName(SHEET_MOV_PEDIDOS);
  if(!shMP){
    shMP = ss.insertSheet(SHEET_MOV_PEDIDOS);
    // Encabezados: id, fecha, pedido, status, cliente, vendedor, cantidad_items, origen
    shMP.appendRow(['id','fecha','pedido','status','cliente','vendedor','cantidad_items','origen']);
    // Formatear pedido como texto para preservar ceros
    shMP.getRange('C:C').setNumberFormat('@');
  } else if(shMP.getLastRow() === 0){
    shMP.appendRow(['id','fecha','pedido','status','cliente','vendedor','cantidad_items','origen']);
    shMP.getRange('C:C').setNumberFormat('@');
  } else {
    asegurarColumnas_(shMP, ['id','fecha','pedido','status','cliente','vendedor','cantidad_items','origen']);
  }

  // Reparación automática: cada fila debe tener un ID único.
  // Esto corrige registros antiguos que pudieron quedar con IDs iguales.
  asegurarIdsUnicos_(shV, 'MV');
  asegurarIdsUnicos_(shP, 'PD');
  asegurarIdsUnicos_(shMU, 'MU');
  asegurarIdsUnicos_(shMP, 'MVP');

}

/* ==================================================
   UTILIDADES
================================================== */
function generarIdUnico_(prefijo){
  /*
    ID único real por fila.
    Antes se usaba sólo Date.now(), pero cuando se importan o sincronizan
    muchas filas en el mismo segundo/milisegundo podían quedar IDs repetidos.
    Ahora combinamos timestamp + UUID + aleatorio para que cada fila tenga
    un identificador diferente aunque se creen muchas filas al mismo tiempo.
  */
  const ts = new Date().getTime().toString(36).toUpperCase();
  const uuid = Utilities.getUuid().replace(/-/g, '').toUpperCase().substring(0, 14);
  const rnd = Math.floor(Math.random() * 1000000000).toString(36).toUpperCase();
  return prefijo + '-' + ts + '-' + uuid + '-' + rnd;
}

function generarId(){
  return generarIdUnico_('MV');
}

/* ==================================================
   ID PARA PEDIDOS
================================================== */
function generarIdPedido(){
  return generarIdUnico_('PD');
}

/* ==================================================
   ID PARA MOVIMIENTO DE UBICACIÓN
================================================== */
function generarIdMovUbi(){
  return generarIdUnico_('MU');
}

/* ==================================================
   ID PARA MOVIMIENTO DE PEDIDOS
================================================== */
function generarIdMovPedido(){
  return generarIdUnico_('MVP');
}

function asegurarIdsUnicos_(sh, prefijo){
  /*
    Repara IDs antiguos vacíos o duplicados sin borrar información.
    Mantiene el primer ID encontrado y cambia sólo las filas repetidas/vacías.
  */
  if(!sh || sh.getLastRow() < 2) return 0;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(h => txt(h).toLowerCase());
  const idxId = headers.indexOf('id');
  if(idxId === -1) return 0;

  const numRows = sh.getLastRow() - 1;
  const range = sh.getRange(2, idxId + 1, numRows, 1);
  const values = range.getDisplayValues();
  const usados = {};
  let cambios = 0;

  const nuevos = values.map(row => {
    let id = txt(row[0]);
    if(!id || usados[id]){
      do {
        id = generarIdUnico_(prefijo);
      } while(usados[id]);
      cambios++;
    }
    usados[id] = true;
    return [id];
  });

  if(cambios){
    range.setNumberFormat('@');
    range.setValues(nuevos);
  }
  return cambios;
}

/* ==================================================
   BÚSQUEDA RÁPIDA EN UBICACIONES
================================================== */
function buscarUbicacionRapido(codigo){
  const codBuscado = normalizarCodigo(codigo);
  if(!codBuscado) return null;

  const cache = CacheService.getScriptCache();
  const cacheKey = 'ubi_' + codBuscado;
  const cacheData = cache.get(cacheKey);
  if(cacheData){
    return JSON.parse(cacheData);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_UBICACIONES);
  if(!sh) return null;

  const lastRow = sh.getLastRow();
  if(lastRow < 2) return null;

  const rangoCodigos = sh.getRange(2,1,lastRow-1,1);
  const celda = rangoCodigos
    .createTextFinder(codBuscado)
    .matchEntireCell(true)
    .findNext();
  if(!celda) return null;
  const row = celda.getRow();
  const codigoReal = normalizarCodigo(sh.getRange(row,1).getDisplayValue());
  const descripcion = txt(sh.getRange(row,2).getDisplayValue());
  const cantidad = Number(sh.getRange(row,3).getDisplayValue() || 0);
  const ubicacion = txt(sh.getRange(row,4).getDisplayValue());
  const resultado = {
    ok:true,
    codigo: codigoReal,
    descripcion: descripcion,
    cantidad: cantidad,
    ubicacion: ubicacion
  };
  cache.put(cacheKey, JSON.stringify(resultado), 21600);
  return resultado;
}

/* ==================================================
   BÚSQUEDA DE ÚLTIMA UBICACIÓN POR MOVIMIENTO
================================================== */
/**
 * Devuelve la ubicación más reciente de un producto a partir de la BD
 * MOVIMIENTO_UBICACION. Recorre la BD desde el final para encontrar la
 * última ocurrencia del código indicado y devuelve su descripción, cantidad
 * y ubicación. Si no hay registros para ese código, devuelve null.
 */
function buscarUbicacionDesdeMovimientos(codigo){
  const codBuscado = normalizarCodigo(codigo);
  if(!codBuscado) return null;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shMU = ss.getSheetByName(SHEET_MOV_UBI);
  if(!shMU) return null;
  const data = shMU.getDataRange().getDisplayValues();
  // Recorremos de abajo hacia arriba para encontrar la última ubicación
  for(let i = data.length - 1; i >= 1; i--){
    const rowCodigo = normalizarCodigo(data[i][2]);
    if(rowCodigo === codBuscado){
      const descripcion = txt(data[i][3]);
      const cantidad = Number(data[i][4] || 0);
      const ubicacion = txt(data[i][5]);
      return {
        ok:true,
        codigo: codBuscado,
        descripcion: descripcion,
        cantidad: cantidad,
        ubicacion: ubicacion
      };
    }
  }
  return null;
}

/* ==================================================
   BÚSQUEDA DE UBICACIONES EN BD-MOVIMIENTO
================================================== */
/**
 * Devuelve TODAS las ubicaciones donde aparece un producto según la BD-MOVIMIENTO.
 * Se recorre desde abajo hacia arriba para tomar la información más reciente por
 * combinación código + ubicación. La respuesta incluye un texto unido para mostrar
 * en tablas/PDF y también un arreglo de ubicaciones con cantidad.
 */
function buscarUbicacionesDesdeMovimientoSheet(codigo){
  const codBuscado = normalizarCodigo(codigo);
  if(!codBuscado) return null;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shMov = ss.getSheetByName(SHEET_MOV);
  if(!shMov) return null;
  const data = shMov.getDataRange().getDisplayValues();
  const porUbicacion = {};
  let descripcion = '';
  let fecha = '';

  for(let i = data.length - 1; i >= 1; i--){
    const rowCodigo = normalizarCodigo(data[i][5]);
    if(rowCodigo !== codBuscado) continue;

    const ubicacion = txt(data[i][4]);
    if(!ubicacion) continue;

    const key = ubicacion.toUpperCase();
    if(!porUbicacion[key]){
      const cantidad = Number(data[i][7] || 0);
      const status = txt(data[i][9]);
      porUbicacion[key] = {
        ubicacion: ubicacion,
        cantidad: cantidad,
        status: status,
        fecha: data[i][1] || ''
      };
    }

    if(!descripcion) descripcion = txt(data[i][6]);
    if(!fecha) fecha = data[i][1] || '';
  }

  const ubicaciones = Object.values(porUbicacion);
  if(!ubicaciones.length) return null;

  const ubicacionTexto = ubicaciones
    .map(u => u.cantidad ? `${u.ubicacion} (${u.cantidad})` : u.ubicacion)
    .join(' | ');

  return {
    ok:true,
    codigo: codBuscado,
    descripcion: descripcion,
    cantidad: ubicaciones.reduce((acc,u)=>acc + (Number(u.cantidad) || 0), 0),
    ubicacion: ubicacionTexto,
    ubicaciones: ubicaciones,
    fecha: fecha
  };
}

// Compatibilidad con llamadas antiguas: ahora retorna todas las ubicaciones, no sólo una.
function buscarUbicacionDesdeMovimientoSheet(codigo){
  return buscarUbicacionesDesdeMovimientoSheet(codigo);
}

/* ==================================================
   LISTADO OPTIMIZADO DE UBICACIONES ACTUALES
================================================== */
/**
 * Devuelve una lista compacta con TODAS las ubicaciones registradas por código
 * desde la BD-MOVIMIENTO. Esto evita enviar toda la BD al navegador y
 * acelera la sincronización de ubicaciones.
 */
function listarUbicacionesActualesDesdeMovimiento(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shMov = ss.getSheetByName(SHEET_MOV);
  if(!shMov) return [];

  const data = shMov.getDataRange().getDisplayValues();
  const mapa = {};

  for(let i = data.length - 1; i >= 1; i--){
    const codigo = normalizarCodigo(data[i][5]);
    const ubicacion = txt(data[i][4]);
    if(!codigo || !ubicacion) continue;

    if(!mapa[codigo]){
      mapa[codigo] = {
        codigo: codigo,
        descripcion: txt(data[i][6]),
        cantidad: 0,
        fecha: data[i][1] || '',
        ubicaciones: [],
        _keys: {}
      };
    }

    const key = ubicacion.toUpperCase();
    if(!mapa[codigo]._keys[key]){
      const cantidad = Number(data[i][7] || 0);
      mapa[codigo]._keys[key] = true;
      mapa[codigo].ubicaciones.push({
        ubicacion: ubicacion,
        cantidad: cantidad,
        status: txt(data[i][9]),
        fecha: data[i][1] || ''
      });
      mapa[codigo].cantidad += cantidad;
    }
  }

  return Object.values(mapa).map(item => {
    delete item._keys;
    item.ubicacion = item.ubicaciones
      .map(u => u.cantidad ? `${u.ubicacion} (${u.cantidad})` : u.ubicacion)
      .join(' | ');
    return item;
  });
}

function resolverUbicacionProducto(codigo, fallback){
  const desdeMovimiento = buscarUbicacionDesdeMovimientoSheet(codigo);
  if(desdeMovimiento && desdeMovimiento.ubicacion) return txt(desdeMovimiento.ubicacion);

  const desdeMovimientoUbicacion = buscarUbicacionDesdeMovimientos(codigo);
  if(desdeMovimientoUbicacion && desdeMovimientoUbicacion.ubicacion) return txt(desdeMovimientoUbicacion.ubicacion);

  const desdeUbicaciones = buscarUbicacionRapido(codigo);
  if(desdeUbicaciones && desdeUbicaciones.ubicacion) return txt(desdeUbicaciones.ubicacion);

  return txt(fallback);
}

function txt(v){
  return String(v ?? '').trim();
}

function normalizarCodigo(v){
  return String(v ?? '')
    .trim()
    .replace(/^'+/, '')
    .replace(/\s+/g, '')
    .toUpperCase();
}

function numeroPedidoOrden_(v){
  const nums = txt(v).match(/\d+/g);
  if(!nums) return -1;
  const n = Number(nums.join(''));
  return Number.isFinite(n) ? n : -1;
}

function fechaOrdenPedido_(v){
  const raw = txt(v);
  if(!raw) return 0;

  let m = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})(?:[\s,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(a\.?\s*m\.?|p\.?\s*m\.?|am|pm)?)?/i);
  if(m){
    let year = String(m[3]);
    if(year.length === 2) year = '20' + year;
    let hour = Number(m[4] || 0);
    const minute = Number(m[5] || 0);
    const second = Number(m[6] || 0);
    const ap = txt(m[7]).toLowerCase().replace(/\s|\./g,'');
    if(ap === 'pm' && hour < 12) hour += 12;
    if(ap === 'am' && hour === 12) hour = 0;
    const t = new Date(Number(year), Number(m[2]) - 1, Number(m[1]), hour, minute, second).getTime();
    return Number.isNaN(t) ? 0 : t;
  }

  m = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[\s,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if(m){
    const t = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)).getTime();
    return Number.isNaN(t) ? 0 : t;
  }

  const t = new Date(raw).getTime();
  return Number.isNaN(t) ? 0 : t;
}

function parseBody_(e){
  // Acepta JSON puro, formulario con campo data y parámetros directos.
  // Esto evita problemas CORS cuando el front envía application/x-www-form-urlencoded.
  try{
    if(e && e.parameter && e.parameter.data){
      return JSON.parse(e.parameter.data || '{}');
    }
  }catch(err){}

  try{
    const raw = e && e.postData ? (e.postData.contents || '') : '';
    if(raw){
      return JSON.parse(raw);
    }
  }catch(err){}

  try{
    if(e && e.parameter){
      return Object.assign({}, e.parameter);
    }
  }catch(err){}

  return {};
}

function normalizarEstadoPedido_(value){
  const estado = txt(value || 'pendiente')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if(estado === 'terminado') return 'terminado';
  if(estado === 'cancelado') return 'cancelado';
  if(estado === 'atrasado') return 'atrasado';
  return 'pendiente';
}

function estadoVisibleListado_(value){
  // En BD-PEDIDOS deben mantenerse visibles todos los estados distintos de TERMINADO.
  // Sólo TERMINADO se mueve a MOVIMIENTO_UBICACION y deja de aparecer en pedidos.
  const estado = normalizarEstadoPedido_(value || 'pendiente');
  return estado !== 'terminado';
}

function asegurarColumnas_(sh, headers){
  if(!sh || !headers || !headers.length) return;
  const lastCol = Math.max(1, sh.getLastColumn());
  const existentes = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h => txt(h));
  headers.forEach(h => {
    if(existentes.indexOf(h) === -1){
      sh.insertColumnAfter(sh.getLastColumn());
      sh.getRange(1, sh.getLastColumn()).setValue(h);
      existentes.push(h);
    }
  });
}

function appendMovimientoPedido_(ss, pedido, estado, cliente, vendedor, cantidadItems, origen){
  const shMP = ss.getSheetByName(SHEET_MOV_PEDIDOS);
  asegurarColumnas_(shMP, ['id','fecha','pedido','status','cliente','vendedor','cantidad_items','origen']);
  shMP.appendRow([
    generarIdMovPedido(),
    new Date(),
    txt(pedido),
    normalizarEstadoPedido_(estado),
    txt(cliente),
    txt(vendedor),
    Number(cantidadItems || 0),
    txt(origen || 'CAMBIO_ESTADO')
  ]);
}

function appendMovimientosUbicacionPedido_(ss, filas, estado, origen){
  const shMU = ss.getSheetByName(SHEET_MOV_UBI);
  asegurarColumnas_(shMU, ['id','fecha','codigo','descripcion','cantidad','ubicacion','pedido','cliente','vendedor','status','origen']);
  if(!filas || !filas.length) return 0;
  const values = filas.map(item => [
    generarIdMovUbi(),
    new Date(),
    normalizarCodigo(item.codigo),
    txt(item.descripcion),
    Number(item.cantidad || 1),
    txt(item.ubicacion),
    txt(item.pedido),
    txt(item.cliente),
    txt(item.vendedor),
    normalizarEstadoPedido_(estado),
    txt(origen || 'CAMBIO_ESTADO_PEDIDO')
  ]);
  shMU.getRange('C:C').setNumberFormat('@');
  shMU.getRange(shMU.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
  return values.length;
}



function normalizarPedido_(v){
  return String(v ?? '')
    .trim()
    .replace(/^'+/, '')
    .replace(/\s+/g, '')
    .toUpperCase();
}

function parseItemsPedido_(value){
  if(!value) return [];
  if(Array.isArray(value)) return value;
  try{
    const parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? parsed : [];
  }catch(err){
    return [];
  }
}

function obtenerValorItem_(item, names, fallback){
  item = item || {};
  for(const n of names){
    if(item[n] !== undefined && item[n] !== null && String(item[n]).trim() !== ''){
      return item[n];
    }
  }
  return fallback;
}

function filasDesdeItemsBody_(items, pedidoNum, estado){
  return (items || []).map(item => {
    const codigo = normalizarCodigo(obtenerValorItem_(item, ['codigo','Código','CODIGO','producto','PRODUCTO','sku','SKU'], ''));
    const descripcion = txt(obtenerValorItem_(item, ['descripcion','Descripción','DESCRIPCION','DESCRIPCIÓN','detalle','DETALLE'], ''));
    const ubicacion = resolverUbicacionProducto(codigo, obtenerValorItem_(item, ['ubicacion','Ubicación','UBICACION','UBICACIÓN'], ''));
    const pedido = txt(obtenerValorItem_(item, ['pedido','Pedido','PEDIDO'], pedidoNum));
    const cliente = txt(obtenerValorItem_(item, ['cliente','Cliente','CLIENTE'], ''));
    const vendedor = txt(obtenerValorItem_(item, ['vendedor','Vendedor','VENDEDOR'], ''));
    const fecha = txt(obtenerValorItem_(item, ['fecha','Fecha','FECHA'], ''));
    return {
      sheetRow: 0,
      id: txt(obtenerValorItem_(item, ['id','ID'], '')),
      codigo,
      descripcion,
      ubicacion,
      pedido: pedido || pedidoNum,
      cliente,
      vendedor,
      fecha,
      cantidad: Number(obtenerValorItem_(item, ['cantidad','Cantidad','CANTIDAD'], 1) || 1),
      status: normalizarEstadoPedido_(estado)
    };
  }).filter(item => item.codigo && item.pedido);
}

function appendPedidosSiNoExisten_(shP, filas, estado){
  if(!filas || !filas.length) return 0;
  asegurarColumnas_(shP, ['id','codigo','descripcion','ubicacion','pedido','cliente','vendedor','status','fecha']);
  const lastCol = Math.max(shP.getLastColumn(), 9);
  const header = shP.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h => txt(h).toLowerCase());
  const idx = (name, fallback) => {
    const i = header.indexOf(name);
    return i === -1 ? fallback : i;
  };
  const idxId = idx('id',0), idxCodigo=idx('codigo',1), idxDesc=idx('descripcion',2), idxUbi=idx('ubicacion',3), idxPed=idx('pedido',4), idxCli=idx('cliente',5), idxVen=idx('vendedor',6), idxStatus=idx('status',7), idxFecha=idx('fecha',8);
  const existentes = {};
  const data = shP.getLastRow() > 1 ? shP.getRange(2,1,shP.getLastRow()-1,lastCol).getDisplayValues() : [];
  data.forEach(r => {
    const key = normalizarCodigo(r[idxCodigo]) + '|' + normalizarPedido_(r[idxPed]);
    existentes[key] = true;
  });
  const rows = [];
  filas.forEach(item => {
    const key = normalizarCodigo(item.codigo) + '|' + normalizarPedido_(item.pedido);
    if(existentes[key]) return;
    const row = new Array(lastCol).fill('');
    row[idxId] = generarIdPedido();
    row[idxCodigo] = normalizarCodigo(item.codigo);
    row[idxDesc] = txt(item.descripcion);
    row[idxUbi] = txt(item.ubicacion);
    row[idxPed] = txt(item.pedido);
    row[idxCli] = txt(item.cliente);
    row[idxVen] = txt(item.vendedor);
    row[idxStatus] = normalizarEstadoPedido_(estado);
    row[idxFecha] = txt(item.fecha);
    rows.push(row);
    existentes[key] = true;
  });
  if(rows.length){
    shP.getRange(1, idxCodigo + 1, shP.getMaxRows(), 1).setNumberFormat('@');
    shP.getRange(shP.getLastRow()+1,1,rows.length,lastCol).setValues(rows);
  }
  return rows.length;
}

/**
 * Actualiza el estado de un pedido desde GET o POST.
 * Estados soportados: pendiente, terminado, atrasado, cancelado.
 * - pendiente/atrasado/cancelado: actualiza PEDIDOS y registra movimientos.
 * - terminado: copia los productos a MOVIMIENTO_UBICACION, registra MOVIMIENTO_PEDIDOS
 *   y elimina el pedido de PEDIDOS.
 */
function actualizarEstadoPedido_(body){
  initSistema();
  body = body || {};

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shP = ss.getSheetByName(SHEET_PEDIDOS);
  if(!shP){
    return { ok:false, msg:'No existe la BD-PEDIDOS' };
  }

  const pedidoNum = txt(body.pedido || body.numero || body.order || body.nro_pedido);
  const nuevoEstado = normalizarEstadoPedido_(body.nuevo_estado || body.status || body.estado);
  const itemsBody = parseItemsPedido_(body.items || body.productos || body.detalle || body.rows);

  if(!pedidoNum || !nuevoEstado){
    return { ok:false, msg:'Pedido o estado vacíos' };
  }

  // Aseguramos columnas base antes de leer, para que status y fecha siempre existan.
  asegurarColumnas_(shP, ['id','codigo','descripcion','ubicacion','pedido','cliente','vendedor','status','fecha']);

  const lastCol = Math.max(shP.getLastColumn(), 9);
  const dataP = shP.getRange(1,1,Math.max(shP.getLastRow(),1),lastCol).getDisplayValues();
  if(dataP.length < 1){
    return { ok:false, msg:'La BD-PEDIDOS no tiene encabezados' };
  }

  const headerP = dataP[0].map(h => txt(h).toLowerCase());
  const getIndex = function(names, fallback){
    for(let i=0; i<names.length; i++){
      const idx = headerP.indexOf(names[i]);
      if(idx !== -1) return idx;
    }
    return fallback;
  };

  const idxId = getIndex(['id'], 0);
  const idxCodigo = getIndex(['codigo','código','producto','sku'], 1);
  const idxDescripcion = getIndex(['descripcion','descripción','detalle'], 2);
  const idxUbicacion = getIndex(['ubicacion','ubicación'], 3);
  const idxPedido = getIndex(['pedido','numero','número','nro pedido','nro_pedido'], 4);
  const idxCliente = getIndex(['cliente'], 5);
  const idxVendedor = getIndex(['vendedor'], 6);
  const idxStatus = getIndex(['status','estado'], 7);
  const idxFecha = getIndex(['fecha','fecha pedido','fecha_pedido'], 8);

  const filasEncontradas = [];
  const pedidoBuscadoNorm = normalizarPedido_(pedidoNum);

  for(let i = 1; i < dataP.length; i++){
    const pedidoBD = txt(dataP[i][idxPedido]);
    if(normalizarPedido_(pedidoBD) === pedidoBuscadoNorm){
      const codigoPedido = normalizarCodigo(dataP[i][idxCodigo]);
      const ubicacionPedido = resolverUbicacionProducto(codigoPedido, dataP[i][idxUbicacion]);
      filasEncontradas.push({
        sheetRow: i + 1,
        id: txt(dataP[i][idxId]),
        codigo: codigoPedido,
        descripcion: txt(dataP[i][idxDescripcion]),
        ubicacion: ubicacionPedido,
        pedido: pedidoBD || pedidoNum,
        cliente: txt(dataP[i][idxCliente]),
        vendedor: txt(dataP[i][idxVendedor]),
        fecha: txt(dataP[i][idxFecha]),
        cantidad: 1
      });
    }
  }

  // Si el pedido no estaba todavía en PEDIDOS, usamos los productos enviados desde el HTML.
  let filasParaProcesar = filasEncontradas;
  let insertadosPedidos = 0;
  if(!filasParaProcesar.length && itemsBody.length){
    filasParaProcesar = filasDesdeItemsBody_(itemsBody, pedidoNum, nuevoEstado);
    if(nuevoEstado !== 'terminado'){
      insertadosPedidos = appendPedidosSiNoExisten_(shP, filasParaProcesar, nuevoEstado);
      // Volvemos a leer para poder actualizar filas reales y mantener consistencia.
      const dataReload = shP.getRange(1,1,Math.max(shP.getLastRow(),1),Math.max(shP.getLastColumn(),9)).getDisplayValues();
      for(let i=1; i<dataReload.length; i++){
        if(normalizarPedido_(dataReload[i][idxPedido]) === pedidoBuscadoNorm){
          filasEncontradas.push({
            sheetRow: i+1,
            id: txt(dataReload[i][idxId]),
            codigo: normalizarCodigo(dataReload[i][idxCodigo]),
            descripcion: txt(dataReload[i][idxDescripcion]),
            ubicacion: resolverUbicacionProducto(dataReload[i][idxCodigo], dataReload[i][idxUbicacion]),
            pedido: txt(dataReload[i][idxPedido]) || pedidoNum,
            cliente: txt(dataReload[i][idxCliente]),
            vendedor: txt(dataReload[i][idxVendedor]),
            fecha: txt(dataReload[i][idxFecha]),
            cantidad: 1
          });
        }
      }
      if(filasEncontradas.length) filasParaProcesar = filasEncontradas;
    }
  }

  if(!filasParaProcesar.length){
    return {
      ok:false,
      msg:'Pedido no encontrado en PEDIDOS y no llegaron productos desde el HTML: ' + pedidoNum,
      pedido: pedidoNum,
      estado: nuevoEstado
    };
  }

  const clientePedido = filasParaProcesar[0].cliente || txt(body.cliente || '');
  const vendedorPedido = filasParaProcesar[0].vendedor || txt(body.vendedor || '');

  // Registro de trazabilidad en MOVIMIENTO_PEDIDOS.
  appendMovimientoPedido_(
    ss,
    pedidoNum,
    nuevoEstado,
    clientePedido,
    vendedorPedido,
    filasParaProcesar.length,
    nuevoEstado === 'terminado' ? 'PEDIDO_TERMINADO' : 'CAMBIO_ESTADO'
  );

  // Registro de productos asociados en MOVIMIENTO_UBICACION.
  const origenUbi = nuevoEstado === 'terminado'
    ? 'PEDIDO_TERMINADO'
    : 'CAMBIO_ESTADO_' + nuevoEstado.toUpperCase();
  const movidos = appendMovimientosUbicacionPedido_(ss, filasParaProcesar, nuevoEstado, origenUbi);

  if(nuevoEstado === 'terminado'){
    // Eliminar de abajo hacia arriba sólo las filas que realmente existen en PEDIDOS.
    filasParaProcesar
      .map(item => Number(item.sheetRow || 0))
      .filter(rowNumber => rowNumber > 1)
      .sort((a,b) => b - a)
      .forEach(rowNumber => shP.deleteRow(rowNumber));

    return {
      ok:true,
      msg:'Pedido terminado: movido a MOVIMIENTO_UBICACION y eliminado de PEDIDOS',
      pedido: pedidoNum,
      estado: nuevoEstado,
      removido_de_pedidos: true,
      movidos: movidos,
      movimiento_pedido: true,
      insertados_pedidos: insertadosPedidos
    };
  }

  // Para pendiente, atrasado y cancelado se actualiza el status en PEDIDOS.
  filasParaProcesar.forEach(item => {
    if(Number(item.sheetRow || 0) > 1){
      shP.getRange(item.sheetRow, idxStatus + 1).setValue(nuevoEstado);
    }
  });

  return {
    ok:true,
    msg:'Estado de pedido actualizado correctamente',
    pedido: pedidoNum,
    estado: nuevoEstado,
    removido_de_pedidos: false,
    movidos: movidos,
    movimiento_pedido: true,
    insertados_pedidos: insertadosPedidos
  };
}

/* ==================================================
   BÚSQUEDA RÁPIDA EN MAESTRA
================================================== */
function normHeader_(v){
  return String(v || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,'')
    .trim();
}

function indexOfHeader_(headers, names, fallback){
  const normalized = headers.map(normHeader_);
  for(const name of names){
    const idx = normalized.indexOf(normHeader_(name));
    if(idx !== -1) return idx;
  }
  return fallback;
}

function getMaestraIndexes_(sh){
  const lastCol = Math.max(sh.getLastColumn(), 3);
  const headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  return {
    codigo: indexOfHeader_(headers, ['CODIGO','CÓDIGO','PRODUCTO','SKU','ITEM','COD'], 0),
    descripcion: indexOfHeader_(headers, ['DESCRIPCION','DESCRIPCIÓN','DETALLE','NOMBRE','PRODUCTO_DESCRIPCION'], 1),
    cantidad: indexOfHeader_(headers, ['CANTIDAD','UNIDADES','STOCK','EXISTENCIA'], 2)
  };
}

function productoDesdeFilaMaestra_(row, idx, codigoBuscado){
  const codigoReal = normalizarCodigo(row[idx.codigo]);
  if(!codigoReal) return null;
  if(codigoBuscado && codigoReal !== codigoBuscado) return null;
  return {
    ok: true,
    codigo: codigoReal,
    descripcion: txt(row[idx.descripcion]),
    cantidad: Number(String(row[idx.cantidad] || '0').replace(',','.')) || 0
  };
}

/**
 * Devuelve todos los productos de MAESTRA.
 * Usa getDisplayValues() para preservar ceros a la izquierda y códigos en texto.
 * El resultado queda liviano para el front: codigo, descripcion y cantidad.
 */
function listarProductosMaestra(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_MAESTRA);
  if(!sh || sh.getLastRow() < 2) return [];

  const idx = getMaestraIndexes_(sh);
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
  const out = [];
  const vistos = {};

  values.forEach(row => {
    const p = productoDesdeFilaMaestra_(row, idx, '');
    if(p && !vistos[p.codigo]){
      out.push({
        codigo: p.codigo,
        descripcion: p.descripcion,
        cantidad: p.cantidad
      });
      vistos[p.codigo] = true;
    }
  });
  return out;
}

/**
 * Búsqueda rápida por código contra la BD-MAESTRA.
 * 1) Revisa caché individual por código.
 * 2) Busca con TextFinder.
 * 3) Si no encuentra, recorre displayValues normalizados para soportar apóstrofos y ceros.
 */
function buscarProductoRapido(codigo){
  const codigoBuscado = normalizarCodigo(codigo);
  if(!codigoBuscado) return null;

  const cache = CacheService.getScriptCache();
  const cacheKey = 'prod_' + codigoBuscado;
  const cached = cache.get(cacheKey);
  if(cached){
    try{
      return JSON.parse(cached);
    }catch(err){}
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_MAESTRA);
  if(!sh || sh.getLastRow() < 2) return null;

  const idx = getMaestraIndexes_(sh);
  const codeCol = idx.codigo + 1;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  // TextFinder suele ser rápido cuando el código está exactamente como texto.
  try{
    const cell = sh.getRange(2, codeCol, lastRow - 1, 1)
      .createTextFinder(codigoBuscado)
      .matchEntireCell(true)
      .findNext();
    if(cell){
      const row = sh.getRange(cell.getRow(), 1, 1, lastCol).getDisplayValues()[0];
      const prod = productoDesdeFilaMaestra_(row, idx, codigoBuscado);
      if(prod){
        cache.put(cacheKey, JSON.stringify(prod), 21600);
        return prod;
      }
    }
  }catch(err){}

  // Fallback robusto: compara todos los códigos normalizados.
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  for(let i=0; i<values.length; i++){
    const prod = productoDesdeFilaMaestra_(values[i], idx, codigoBuscado);
    if(prod){
      cache.put(cacheKey, JSON.stringify(prod), 21600);
      return prod;
    }
  }
  return null;
}

/* ==================================================
   OBTENER STOCK ACTUAL
================================================== */
function obtenerStockActual(codigo){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_MOV);
  const data = sh.getDataRange().getDisplayValues();

  const codigoBuscado = normalizarCodigo(codigo);
  let stock = 0;

  for(let i = 1; i < data.length; i++){
    const codigoFila = normalizarCodigo(data[i][5]);
    if(codigoFila === codigoBuscado){
      stock = Number(data[i][7] || 0);
    }
  }

  return stock;
}

/* ==================================================
   GET
================================================== */
function doGet(e){
  initSistema();

  e = e || { parameter:{} };
  JSONP_CALLBACK = txt(e.parameter.callback || '');
  const accion = txt(e.parameter.accion || '');

  // Permite actualizar estado por GET para evitar problemas de CORS con POST.
  // Ejemplo: ?accion=actualizar_estado_pedido&pedido=P001&nuevo_estado=terminado
  if(accion === 'actualizar_estado_pedido' || accion === 'cambiar_estado_pedido'){
    return salida(actualizarEstadoPedido_(e.parameter));
  }


  // Listar cualquier BD permitida del sistema para los módulos HTML de consulta.
  // Devuelve encabezados y datos con getDisplayValues() para conservar ceros a la izquierda.
  if(accion === 'listar_bd'){
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requested = txt(e.parameter.sheet || e.parameter.bd || e.parameter.BD || '').toUpperCase();
    const allowed = {
      'PEDIDOS': SHEET_PEDIDOS,
      'MAESTRA': SHEET_MAESTRA,
      'MOVIMIENTO': SHEET_MOV,
      'UBICACIONES': SHEET_UBICACIONES,
      'MOVIMIENTO_UBICACION': SHEET_MOV_UBI,
      'MOVIMIENTO_PEDIDOS': SHEET_MOV_PEDIDOS
    };
    const sheetName = allowed[requested];
    if(!sheetName){
      return salida({ ok:false, msg:'BD no permitida o no encontrada: ' + requested });
    }
    const sh = ss.getSheetByName(sheetName);
    if(!sh){
      return salida({ ok:false, msg:'No existe la BD: ' + sheetName });
    }
    const values = sh.getDataRange().getDisplayValues();
    const headers = values.length ? values[0].map(h => txt(h)) : [];
    const data = values.length > 1 ? values.slice(1) : [];
    return salida({ ok:true, sheet:sheetName, headers:headers, data:data });
  }

  if(accion === 'buscar' || accion === 'buscar_maestra' || accion === 'buscar_producto'){
    const codigo = normalizarCodigo(e.parameter.codigo);

    if(!codigo){
      return salida({ ok:false, msg:'Código vacío' });
    }

    const producto = buscarProductoRapido(codigo);
    if(producto) return salida(producto);

    return salida({ ok:false, msg:'Código no encontrado' });
  }

  // Catálogo liviano para que el módulo de ubicaciones pueda buscar descripción rápido desde memoria.
  if(accion === 'listar_maestra'){
    return salida({ ok:true, data:listarProductosMaestra() });
  }

  if(accion === 'listar_ubicaciones_actuales'){
    return salida({
      ok:true,
      data:listarUbicacionesActualesDesdeMovimiento()
    });
  }

  if(accion === 'listar'){
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_MOV);
    // getDisplayValues preserva códigos con ceros a la izquierda.
    const data = sh.getDataRange().getDisplayValues();

    if(data.length > 0) data.shift();

    return salida({
      ok:true,
      data:data
    });
  }

  // Listar pedidos de cliente
  if(accion === 'listar_pedidos'){
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_PEDIDOS);
    asegurarColumnas_(sh, ['id','codigo','descripcion','ubicacion','pedido','cliente','vendedor','status','fecha']);

    // Usar getDisplayValues para preservar ceros a la izquierda y fecha visible.
    const data = sh.getDataRange().getDisplayValues();
    if(data.length > 0) data.shift();

    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(h => txt(h).toLowerCase());
    const idxOf = (names, fallback) => {
      for(const name of names){
        const idx = headers.indexOf(name);
        if(idx !== -1) return idx;
      }
      return fallback;
    };

    const idxId = idxOf(['id'],0);
    const idxCodigo = idxOf(['codigo','código','producto','sku'],1);
    const idxDesc = idxOf(['descripcion','descripción','detalle'],2);
    const idxUbi = idxOf(['ubicacion','ubicación'],3);
    const idxPedido = idxOf(['pedido','numero','número','nro pedido','nro_pedido'],4);
    const idxCliente = idxOf(['cliente'],5);
    const idxVendedor = idxOf(['vendedor'],6);
    const idxStatus = idxOf(['status','estado'],7);
    const idxFecha = idxOf(['fecha','fecha pedido','fecha_pedido'],8);

    // Devolver siempre en el orden esperado por el HTML:
    // [id,codigo,descripcion,ubicacion,pedido,cliente,vendedor,status,fecha]
    const visibles = data
      .filter(r => estadoVisibleListado_(r[idxStatus] || 'pendiente'))
      .map(r => [
        txt(r[idxId]),
        normalizarCodigo(r[idxCodigo]),
        txt(r[idxDesc]),
        txt(r[idxUbi]),
        txt(r[idxPedido]),
        txt(r[idxCliente]),
        txt(r[idxVendedor]),
        normalizarEstadoPedido_(r[idxStatus] || 'pendiente'),
        txt(r[idxFecha])
      ])
      // Siempre enviar primero el pedido más nuevo.
      // Prioridad: número de pedido mayor; si se repite, fecha más reciente; si se repite, ID más reciente.
      .sort((a,b) => {
        const np = numeroPedidoOrden_(b[4]) - numeroPedidoOrden_(a[4]);
        if(np !== 0) return np;
        const nf = fechaOrdenPedido_(b[8]) - fechaOrdenPedido_(a[8]);
        if(nf !== 0) return nf;
        return txt(b[0]).localeCompare(txt(a[0]));
      });

    return salida({
      ok:true,
      data:visibles
    });
  }

  // Listar ubicaciones de productos
  if(accion === 'listar_ubicaciones'){
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_UBICACIONES);
    const data = sh.getDataRange().getDisplayValues();
    if(data.length > 0) data.shift();
    return salida({ ok:true, data:data });
  }

  // Listar movimientos de ubicaciones
  if(accion === 'listar_movimientos_ubi'){
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_MOV_UBI);
    const data = sh.getDataRange().getDisplayValues();
    if(data.length > 0) data.shift();
    return salida({ ok:true, data:data });
  }

  // Listar movimientos de pedidos
  if(accion === 'listar_movimientos_pedidos'){
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_MOV_PEDIDOS);
    const data = sh.getDataRange().getDisplayValues();
    if(data.length > 0) data.shift();
    return salida({ ok:true, data:data });
  }

  // Buscar ubicación por código
  if(accion === 'buscar_ubicacion'){
    const codigo = normalizarCodigo(e.parameter.codigo);
    if(!codigo){
      return salida({ ok:false, msg:'Código vacío' });
    }

    // 1. La fuente principal es MOVIMIENTO, porque ahí queda la última ubicación real.
    const movSheetUbi = buscarUbicacionDesdeMovimientoSheet(codigo);
    if(movSheetUbi) return salida(movSheetUbi);

    // 2. Luego se consulta MOVIMIENTO_UBICACION.
    const movUbi = buscarUbicacionDesdeMovimientos(codigo);
    if(movUbi) return salida(movUbi);

    // 3. Como respaldo se consulta UBICACIONES.
    const ubi = buscarUbicacionRapido(codigo);
    if(ubi) return salida(ubi);

    return salida({ ok:false, msg:'Ubicación no encontrada' });
  }

  return salida({
    ok:false,
    msg:'Acción no válida'
  });
}

/* ==================================================
   POST
================================================== */
function doPost(e){
  initSistema();

  const body = parseBody_(e);
  const accion = txt(body.accion);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_MOV);
  const data = sh.getDataRange().getDisplayValues();

  // BD de pedidos
  const shP = ss.getSheetByName(SHEET_PEDIDOS);
  const dataP = shP.getDataRange().getDisplayValues();

  // BD de ubicaciones
  const shU = ss.getSheetByName(SHEET_UBICACIONES);
  const dataU = shU.getDataRange().getDisplayValues();

  // BD de movimientos de ubicaciones
  const shMU = ss.getSheetByName(SHEET_MOV_UBI);
  const dataMU = shMU.getDataRange().getDisplayValues();

  const codigo = normalizarCodigo(body.codigo);
  const cantidadNueva = Number(body.cantidad || 0);
  const status = body.status === 'RETIRADO' ? 'RETIRADO' : 'VIGENTE';

  if(!accion){
    return salida({ ok:false, msg:'Acción vacía' });
  }

  // Para acciones distintas a eliminar, importar_pedidos y actualizar/cambiar estado no requerimos código
  if(!codigo && accion !== 'eliminar' && accion !== 'importar_pedidos' && accion !== 'actualizar_estado_pedido' && accion !== 'cambiar_estado_pedido'){
    return salida({ ok:false, msg:'Código vacío' });
  }

  // Solo aplicamos validación de stock negativo en acción 'agregar' de movimientos
  if(cantidadNueva < 0 && accion === 'agregar'){
    return salida({
      ok:false,
      msg:'Stock insuficiente (validación servidor)'
    });
  }

  if(accion === 'agregar'){
    sh.appendRow([
      generarId(),
      new Date(),
      body.fecha_entrada || '',
      body.fecha_salida || '',
      txt(body.ubicacion),
      "'" + codigo,
      txt(body.descripcion),
      cantidadNueva,
      txt(body.responsable),
      status,
      txt(body.origen || 'WEB')
    ]);

    return salida({ ok:true, msg:'Registro agregado correctamente' });
  }

  if(accion === 'editar'){
    for(let i = 1; i < data.length; i++){
      if(txt(data[i][0]) === txt(body.id)){
        sh.getRange(i + 1, 3, 1, 9).setValues([[
          body.fecha_entrada || '',
          body.fecha_salida || '',
          txt(body.ubicacion),
          "'" + codigo,
          txt(body.descripcion),
          cantidadNueva,
          txt(body.responsable),
          status,
          txt(body.origen || 'EDIT')
        ]]);

        return salida({ ok:true, msg:'Registro editado correctamente' });
      }
    }

    return salida({ ok:false, msg:'ID no encontrado' });
  }

  if(accion === 'eliminar'){
    for(let i = 1; i < data.length; i++){
      if(txt(data[i][0]) === txt(body.id)){
        sh.deleteRow(i + 1);
        return salida({ ok:true, msg:'Registro eliminado correctamente' });
      }
    }

    return salida({ ok:false, msg:'ID no encontrado' });
  }

  /* =====================================================
     ACCIONES PARA PEDIDOS
  ==================================================== */
  if(accion === 'agregar_pedido'){
    // Agregar un nuevo pedido preservando ceros a la izquierda y fecha.
    const estado = normalizarEstadoPedido_(body.status || 'pendiente');
    const fila = filasDesdeItemsBody_([{
      id: body.id,
      codigo: body.codigo,
      descripcion: body.descripcion,
      ubicacion: body.ubicacion,
      pedido: body.pedido,
      cliente: body.cliente,
      vendedor: body.vendedor,
      fecha: body.fecha,
      cantidad: body.cantidad || 1
    }], txt(body.pedido), estado);

    const insertados = appendPedidosSiNoExisten_(shP, fila, estado);
    return salida({
      ok:true,
      msg: insertados ? 'Pedido agregado correctamente' : 'Pedido ya existía, no se duplicó',
      insertados: insertados
    });
  }

  if(accion === 'editar_pedido'){
    // Editar pedido existente, preservando código como texto y fecha.
    asegurarColumnas_(shP, ['id','codigo','descripcion','ubicacion','pedido','cliente','vendedor','status','fecha']);
    const headers = shP.getRange(1,1,1,shP.getLastColumn()).getDisplayValues()[0].map(h => txt(h).toLowerCase());
    const idxOf = (name, fallback) => {
      const idx = headers.indexOf(name);
      return idx === -1 ? fallback : idx;
    };
    const idxId = idxOf('id',0), idxCodigo=idxOf('codigo',1), idxDesc=idxOf('descripcion',2), idxUbi=idxOf('ubicacion',3), idxPedido=idxOf('pedido',4), idxCliente=idxOf('cliente',5), idxVendedor=idxOf('vendedor',6), idxStatus=idxOf('status',7), idxFecha=idxOf('fecha',8);
    const dataEdit = shP.getDataRange().getDisplayValues();

    for(let i = 1; i < dataEdit.length; i++){
      if(txt(dataEdit[i][idxId]) === txt(body.id)){
        const row = i + 1;
        shP.getRange(row, idxCodigo+1).setValue("'" + normalizarCodigo(body.codigo));
        shP.getRange(row, idxDesc+1).setValue(txt(body.descripcion));
        shP.getRange(row, idxUbi+1).setValue(resolverUbicacionProducto(body.codigo, body.ubicacion));
        shP.getRange(row, idxPedido+1).setValue(txt(body.pedido));
        shP.getRange(row, idxCliente+1).setValue(txt(body.cliente));
        shP.getRange(row, idxVendedor+1).setValue(txt(body.vendedor));
        shP.getRange(row, idxStatus+1).setValue(normalizarEstadoPedido_(body.status || dataEdit[i][idxStatus] || 'pendiente'));
        shP.getRange(row, idxFecha+1).setValue(txt(body.fecha));
        return salida({ ok:true, msg:'Pedido editado correctamente' });
      }
    }
    return salida({ ok:false, msg:'ID de pedido no encontrado' });
  }

  if(accion === 'eliminar_pedido'){
    // Eliminar pedido
    for(let i = 1; i < dataP.length; i++){
      if(txt(dataP[i][0]) === txt(body.id)){
        shP.deleteRow(i + 1);
        return salida({ ok:true, msg:'Pedido eliminado correctamente' });
      }
    }
    return salida({ ok:false, msg:'ID de pedido no encontrado' });
  }

  if(accion === 'importar_pedidos'){
    // Importar múltiples pedidos sin duplicar: sólo se agregan los que no existen y están pendientes.
    const items = Array.isArray(body.items) ? body.items : [];
    const pendientes = items.filter(item => normalizarEstadoPedido_(item.status || 'pendiente') === 'pendiente');
    const filas = filasDesdeItemsBody_(pendientes, '', 'pendiente');
    const insertados = appendPedidosSiNoExisten_(shP, filas, 'pendiente');
    return salida({
      ok:true,
      msg:'Pedidos importados correctamente',
      recibidos: items.length,
      pendientes: pendientes.length,
      insertados: insertados
    });
  }

  // Cambiar o actualizar el estado de un pedido
  if(accion === 'actualizar_estado_pedido' || accion === 'cambiar_estado_pedido'){
    return salida(actualizarEstadoPedido_(body));
  }


  if(accion === 'editar_ubicacion'){
    // Editar registro existente de UBICACIONES
    const cod = normalizarCodigo(body.codigo);
    for(let i=1; i<dataU.length; i++){
      if(txt(dataU[i][0]) === cod){
        shU.getRange(i+1, 2, 1, 3).setValues([[
          txt(body.descripcion),
          Number(body.cantidad || 0),
          txt(body.ubicacion)
        ]]);
        return salida({ ok:true, msg:'Ubicación editada correctamente' });
      }
    }
    return salida({ ok:false, msg:'Código no encontrado en ubicaciones' });
  }

  if(accion === 'eliminar_ubicacion'){
    // Eliminar registro de UBICACIONES
    const cod = normalizarCodigo(body.codigo);
    for(let i=1; i<dataU.length; i++){
      if(txt(dataU[i][0]) === cod){
        shU.deleteRow(i+1);
        return salida({ ok:true, msg:'Ubicación eliminada correctamente' });
      }
    }
    return salida({ ok:false, msg:'Código no encontrado en ubicaciones' });
  }

  if(accion === 'registrar_movimiento_ubicacion'){
    /*
      Registra un movimiento de producto en una ubicación y ajusta el stock.
      Parámetros esperados en body:
        codigo, descripcion (opcional), cantidad (positivo suma, negativo resta), ubicacion, pedido
    */
    const cod = normalizarCodigo(body.codigo);
    const desc = txt(body.descripcion);
    const mov = Number(body.cantidad || 0);
    const ubi = txt(body.ubicacion);
    const ped = txt(body.pedido);
    if(!cod || !ubi || !ped){
      return salida({ ok:false, msg:'Datos incompletos' });
    }
    // Encontrar el índice en UBICACIONES
    let idx = -1;
    let stock = 0;
    for(let i=1; i<dataU.length; i++){
      if(txt(dataU[i][0]) === cod){
        idx = i;
        stock = Number(dataU[i][2] || 0);
        break;
      }
    }
    if(idx === -1){
      return salida({ ok:false, msg:'Código no registrado en ubicaciones' });
    }
    const nuevoStock = stock + mov;
    if(nuevoStock < 0){
      return salida({ ok:false, msg:'Stock insuficiente' });
    }
    // Actualizar stock
    shU.getRange(idx+1, 3).setValue(nuevoStock);
    // Registrar movimiento en la BD de movimientos
    const descripcionMov = desc || txt(dataU[idx][1]);
    shMU.appendRow([
      generarIdMovUbi(),
      new Date(),
      cod,
      descripcionMov,
      mov,
      ubi,
      ped
    ]);
    return salida({ ok:true, msg:'Movimiento registrado correctamente' });
  }

  return salida({ ok:false, msg:'Acción no válida' });
}

/* ==================================================
   SALIDA JSON
================================================== */
function salida(obj){
  const json = JSON.stringify(obj);
  const cb = txt(typeof JSONP_CALLBACK !== 'undefined' ? JSONP_CALLBACK : '');
  if(cb && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(cb)){
    return ContentService
      .createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}