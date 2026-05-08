/* =========================================================
   VERSIÓN REFORZADA 2026-05-06
   - MOVIMIENTO/UBICACIONES soporta agregar, editar y eliminar por POST.
   - Compatible con acciones usadas por web y Android: agregar, editar, eliminar,
     guardar_movimiento, actualizar_movimiento, guardar_ubicacion, actualizar_ubicacion, etc.
   - Alertas inmediatas por token único en PEDIDOS + ALERTAS_PEDIDOS.
   - Cualquier modificación de STATUS/ESTADO genera alerta + mensaje de voz.
========================================================= */

/* =========================================================
   APPS SCRIPT FULL FUNCIONAL - ORDEN DE PEDIDOS
   Lee hoja PEDIDOS, carga PIKEADORES, MAESTRA y envía alertas.
   Compatible con JSONP (?callback=) para evitar CORS en HTML.
========================================================= */
const SPREADSHEET_ID = '1Pwlf0YYZrGYTU5rAZbGQc3PpS_v7Bus-2V9owIN78-o';
const SHEET_PEDIDOS = 'PEDIDOS';
const SHEET_PEDIDO_ALT = 'PEDIDO';
const SHEET_PIKEADORES = 'PIKEADORES';
const SHEET_MAESTRA = 'MAESTRA';
const SHEET_MOV = 'MOVIMIENTO';
const SHEET_ALERTAS = 'ALERTAS_PEDIDOS';
const TZ = 'America/Santiago';

function doGet(e){
  e = e || {parameter:{}};
  const p = e.parameter || {};
  const accion = t_(p.accion || p.action || 'diagnostico').toLowerCase();
  const cb = t_(p.callback || '');
  try{
    let r;
    if(['diagnostico','test'].includes(accion)) r = diagnostico_();
    else if(['reparar_estados','reparar_estados_pedidos','limpiar_estados'].includes(accion)) r = repararEstadosPedidos_();
    else if(['agregar','agregar_movimiento','guardar_movimiento','guardar_captura','agregar_ubicacion','guardar_ubicacion','guardar_ubicaciones','crear_ubicacion','crear_movimiento'].includes(accion)) r = agregarMovimiento_(p);
    else if(['editar','editar_movimiento','actualizar_movimiento','editar_ubicacion','actualizar_ubicacion','modificar_ubicacion','modificar_movimiento'].includes(accion)) r = editarMovimiento_(p);
    else if(['eliminar','eliminar_movimiento','borrar_movimiento','eliminar_ubicacion','borrar_ubicacion'].includes(accion)) r = eliminarMovimiento_(p);
    else if(['reparar_pikeadores','reparar_pikeadores_pedidos','limpiar_pikeadores'].includes(accion)) r = repararPikeadoresPedidos_();
    else if(['listar','listar_movimiento','listar_movimientos','cargar_movimiento','cargar_movimientos','listar_ubicaciones','cargar_ubicaciones','listar_movimiento_ubicaciones'].includes(accion)) r = listarMovimiento_(p);
    else if(['listar_pedidos','listar_pedidos_rapido','cargar_pedidos'].includes(accion)) r = listarPedidos_(p);
    else if(['seguimiento_pedido','consulta_cliente_pedido','consultar_pedido_cliente','estado_pedido_cliente'].includes(accion)) r = seguimientoPedido_(p);
    else if(['listar_alertas','listar_alertas_pedidos','cargar_alertas','alertas_pedidos','ultimas_alertas'].includes(accion)) r = listarAlertas_(p);
    else if(accion === 'listar_bd') r = listarBD_(p.sheet || p.bd || p.hoja || '');
    else if(['listar_pikeadores','obtener_pikeadores','cargar_pikeadores','select_pikeadores'].includes(accion)) r = listarPikeadores_();
    else if(['agregar_pikeador','crear_pikeador'].includes(accion)) r = agregarPikeador_(p.nombre || p.pikeador || '');
    else if(['listar_maestra'].includes(accion)) r = listarMaestra_();
    else if(['buscar','buscar_producto','buscar_maestra'].includes(accion)) r = buscarProducto_(p.codigo || p.cod || '');
    else if(['actualizar_estado_pedido','cambiar_estado_pedido','actualizar_estado','cambiar_estado'].includes(accion)) r = actualizarEstadoPedido_(p);
    else if(['preparar_pedido','enviar_preparacion','enviar_a_preparacion'].includes(accion)) r = alertaPedido_(p, 'PREPARACION', 'Pedido enviado a preparación', true);
    else if(['reenviar_alerta','reenviar_alerta_pedido','enviar_alerta','disparar_alerta_pedido'].includes(accion)) r = alertaPedido_(p, 'ALERTA', 'Pedido enviado a preparación', false);
    else if(['asignar_pikeador','asignar_pikeador_pedido'].includes(accion)) r = asignarPikeador_(p);
    else if(['guardar_pedido','guardar_pedido_rapido','importar_pedidos','importar_pedidos_rapido','importar_listado','enviar_base_datos','enviar_bd'].includes(accion)) r = guardarPedido_(p);
    else r = {ok:false,msg:'Acción no válida: '+accion};
    return out_(r, cb);
  }catch(err){ return out_({ok:false,msg:String(err.message||err),stack:String(err.stack||'')}, cb); }
}

function doPost(e){
  const b = body_(e);
  const accion = t_(b.accion || b.action || b.operacion || '').toLowerCase();
  try{
    let r;
    if(['reparar_estados','reparar_estados_pedidos','limpiar_estados'].includes(accion)) r = repararEstadosPedidos_();
    else if(['agregar','agregar_movimiento','guardar_movimiento','guardar_captura','agregar_ubicacion','guardar_ubicacion','guardar_ubicaciones','crear_ubicacion','crear_movimiento'].includes(accion)) r = agregarMovimiento_(b);
    else if(['editar','editar_movimiento','actualizar_movimiento','editar_ubicacion','actualizar_ubicacion','modificar_ubicacion','modificar_movimiento'].includes(accion)) r = editarMovimiento_(b);
    else if(['eliminar','eliminar_movimiento','borrar_movimiento','eliminar_ubicacion','borrar_ubicacion'].includes(accion)) r = eliminarMovimiento_(b);
    else if(['reparar_pikeadores','reparar_pikeadores_pedidos','limpiar_pikeadores'].includes(accion)) r = repararPikeadoresPedidos_();
    else if(['listar_alertas','listar_alertas_pedidos','cargar_alertas','alertas_pedidos','ultimas_alertas'].includes(accion)) r = listarAlertas_(b);
    else if(['seguimiento_pedido','consulta_cliente_pedido','consultar_pedido_cliente','estado_pedido_cliente'].includes(accion)) r = seguimientoPedido_(b);
    else if(['actualizar_estado_pedido','cambiar_estado_pedido','actualizar_estado','cambiar_estado'].includes(accion)) r = actualizarEstadoPedido_(b);
    else if(['preparar_pedido','enviar_preparacion','enviar_a_preparacion'].includes(accion)) r = alertaPedido_(b, 'PREPARACION', 'Pedido enviado a preparación', true);
    else if(['reenviar_alerta','reenviar_alerta_pedido','enviar_alerta','disparar_alerta_pedido'].includes(accion)) r = alertaPedido_(b, 'ALERTA', 'Pedido enviado a preparación', false);
    else if(['asignar_pikeador','asignar_pikeador_pedido'].includes(accion)) r = asignarPikeador_(b);
    else if(['agregar_pikeador','crear_pikeador'].includes(accion)) r = agregarPikeador_(b.nombre || b.pikeador || '');
    else if(['guardar_pedido','guardar_pedido_rapido','importar_pedidos','importar_pedidos_rapido','importar_listado','enviar_base_datos','enviar_bd'].includes(accion)) r = guardarPedido_(b);
    else r = {ok:false,msg:'Acción POST no válida: '+accion};
    return out_(r, '');
  }catch(err){ return out_({ok:false,msg:String(err.message||err),stack:String(err.stack||'')}, ''); }
}


/* =========================================================
   MOVIMIENTO / UBICACIONES - COMPATIBILIDAD WEB Y ANDROID
   Acciones POST soportadas: agregar, editar, eliminar.
   Acciones GET soportadas: listar, listar_movimiento, listar_bd&sheet=MOVIMIENTO.
========================================================= */
function movimientoHeaders_(){
  return ['id','fecha_registro','fecha_entrada','fecha_salida','ubicacion','codigo','descripcion','cantidad','responsable','status','origen'];
}

function movimientoSheet_(crear){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_MOV);
  if(!sh && crear) sh = getOrCreate_(SHEET_MOV, movimientoHeaders_());
  if(sh && crear){
    ensureCols_(sh, movimientoHeaders_());
    try{
      const h = values_(sh).headers;
      const ixCodigo = find_(h, ['codigo','código','cod','sku'], -1);
      if(ixCodigo >= 0) sh.getRange(1, ixCodigo + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    }catch(e){}
  }
  return sh;
}

function movimientoIdx_(headers){
  return {
    id:find_(headers,['id'],0),
    fecha_registro:find_(headers,['fecha_registro','fecha registro','fecha'],1),
    fecha_entrada:find_(headers,['fecha_entrada','fecha entrada','entrada'],2),
    fecha_salida:find_(headers,['fecha_salida','fecha salida','salida'],3),
    ubicacion:find_(headers,['ubicacion','ubicación','bodega'],4),
    codigo:find_(headers,['codigo','código','cod','sku'],5),
    descripcion:find_(headers,['descripcion','descripción','detalle','nombre'],6),
    cantidad:find_(headers,['cantidad','stock','unidades'],7),
    responsable:find_(headers,['responsable','operador','usuario'],8),
    status:find_(headers,['status','estado'],9),
    origen:find_(headers,['origen','fuente'],10)
  };
}

function listarMovimiento_(p){
  const sh = movimientoSheet_(true);
  const v = values_(sh);
  return {
    ok:true,
    sheet:sh.getName(),
    headers:v.headers,
    data:v.data,
    rows:v.data,
    values:v.data,
    total:v.data.length,
    serverTime:new Date().toISOString()
  };
}

function agregarMovimiento_(b){
  const sh = movimientoSheet_(true);
  const v = values_(sh);
  const headers = v.headers.length ? v.headers : movimientoHeaders_();
  const ix = movimientoIdx_(headers);
  const codigo = code_(b.codigo || b.cod || b.sku || '');
  const cantidad = n_(b.cantidad || b.stock || b.unidades || 0);
  if(!codigo) return {ok:false,msg:'Código vacío'};
  if(cantidad < 0) return {ok:false,msg:'Stock insuficiente o cantidad inválida'};

  const obj = {};
  obj[headers[ix.id] || 'id'] = id_('MV');
  obj[headers[ix.fecha_registro] || 'fecha_registro'] = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  obj[headers[ix.fecha_entrada] || 'fecha_entrada'] = t_(b.fecha_entrada || b.fechaEntrada || '');
  obj[headers[ix.fecha_salida] || 'fecha_salida'] = t_(b.fecha_salida || b.fechaSalida || '');
  obj[headers[ix.ubicacion] || 'ubicacion'] = t_(b.ubicacion || b.ubicación || '');
  obj[headers[ix.codigo] || 'codigo'] = codigo;
  obj[headers[ix.descripcion] || 'descripcion'] = t_(b.descripcion || b.descripción || b.detalle || '');
  obj[headers[ix.cantidad] || 'cantidad'] = cantidad;
  obj[headers[ix.responsable] || 'responsable'] = t_(b.responsable || b.operador || b.usuario || '');
  obj[headers[ix.status] || 'status'] = t_(b.status || b.estado || 'VIGENTE') || 'VIGENTE';
  obj[headers[ix.origen] || 'origen'] = t_(b.origen || 'WEB');

  appendObjects_(sh, headers, [obj]);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Movimiento agregado correctamente',accion:'agregar',id:obj[headers[ix.id] || 'id'],codigo:codigo,cantidad:cantidad,serverTime:new Date().toISOString()};
}

function editarMovimiento_(b){
  const sh = movimientoSheet_(true);
  const v = values_(sh);
  const headers = v.headers;
  const ix = movimientoIdx_(headers);
  const idBuscado = t_(b.id || b.ID || '');
  if(!idBuscado) return {ok:false,msg:'ID vacío'};

  let rowNumber = 0;
  for(let i=0; i<v.data.length; i++){
    if(t_(getCell_(v.data[i], ix.id)) === idBuscado){ rowNumber = i + 2; break; }
  }
  if(!rowNumber) return {ok:false,msg:'ID no encontrado'};

  const updates = {
    fecha_entrada:t_(b.fecha_entrada || b.fechaEntrada || ''),
    fecha_salida:t_(b.fecha_salida || b.fechaSalida || ''),
    ubicacion:t_(b.ubicacion || b.ubicación || ''),
    codigo:code_(b.codigo || b.cod || b.sku || ''),
    descripcion:t_(b.descripcion || b.descripción || b.detalle || ''),
    cantidad:n_(b.cantidad || b.stock || b.unidades || 0),
    responsable:t_(b.responsable || b.operador || b.usuario || ''),
    status:t_(b.status || b.estado || 'VIGENTE') || 'VIGENTE',
    origen:t_(b.origen || 'EDIT')
  };
  if(!updates.codigo) return {ok:false,msg:'Código vacío'};
  if(updates.cantidad < 0) return {ok:false,msg:'Cantidad inválida'};

  const map = [
    ['fecha_entrada', ix.fecha_entrada], ['fecha_salida', ix.fecha_salida], ['ubicacion', ix.ubicacion],
    ['codigo', ix.codigo], ['descripcion', ix.descripcion], ['cantidad', ix.cantidad],
    ['responsable', ix.responsable], ['status', ix.status], ['origen', ix.origen]
  ];
  map.forEach(([field, col])=>{ if(col >= 0) sh.getRange(rowNumber, col + 1).setValue(updates[field]); });
  SpreadsheetApp.flush();
  return {ok:true,msg:'Movimiento editado correctamente',accion:'editar',id:idBuscado,serverTime:new Date().toISOString()};
}

function eliminarMovimiento_(b){
  const sh = movimientoSheet_(true);
  const v = values_(sh);
  const ix = movimientoIdx_(v.headers);
  const idBuscado = t_(b.id || b.ID || '');
  if(!idBuscado) return {ok:false,msg:'ID vacío'};
  for(let i=0; i<v.data.length; i++){
    if(t_(getCell_(v.data[i], ix.id)) === idBuscado){
      sh.deleteRow(i + 2);
      SpreadsheetApp.flush();
      return {ok:true,msg:'Movimiento eliminado correctamente',accion:'eliminar',id:idBuscado,serverTime:new Date().toISOString()};
    }
  }
  return {ok:false,msg:'ID no encontrado'};
}

function diagnostico_(){
  const ss = ss_();
  const ped = pedidosSheet_(false);
  const pik = getOrCreate_(SHEET_PIKEADORES, ['nombre','fecha_creacion']);
  return {ok:true,msg:'doGet funcionando correctamente',spreadsheet:ss.getName(),hoja_pedidos_usada:ped?ped.getName():'NO EXISTE PEDIDOS/PEDIDO',filas_pedidos:ped?Math.max(0,ped.getLastRow()-1):0,filas_pikeadores:pik?Math.max(0,pik.getLastRow()-1):0,sheets:ss.getSheets().map(s=>({nombre:s.getName(),filas:s.getLastRow(),columnas:s.getLastColumn(),maxFilas:s.getMaxRows(),maxColumnas:s.getMaxColumns()})),serverTime:new Date().toISOString()};
}

function listarPedidos_(p){
  const sh = pedidosSheet_(false);
  if(!sh) return {ok:true,sheet:SHEET_PEDIDOS,headers:[],data:[],pedidos:[],msg:'No existe hoja PEDIDOS'};
  // Repara estados inválidos antes de devolver la tabla.
  // Esto corrige datos antiguos donde DESCRIPCIÓN cayó por error en STATUS/ESTADO.
  repararEstadosPedidos_(false);
  repararPikeadoresPedidos_(false);
  const v = values_(sh);
  let rows = v.data;
  const idx = idx_(v.headers);
  const mostrar = truth_(p.mostrarTodo || p.todo || '1');
  if(!mostrar && idx.status >= 0) rows = rows.filter(r => estadoSeguro_(r[idx.status], 'PENDIENTE') !== 'TERMINADO');
  return {ok:true,sheet:sh.getName(),headers:v.headers,data:rows,rawHeaders:v.headers,rawData:v.data,pedidos:agrupar_(v.headers,rows),totalFilasHoja:v.data.length,serverTime:new Date().toISOString()};
}


function seguimientoPedido_(p){
  const pedido = valorPedido_(p, p.order || p.nro_pedido || p.numero || p['número']);
  const clienteFiltro = nh_(p.cliente || p.rut || p.codigo_cliente || '');
  if(!pedido) return {ok:false,msg:'Ingresa el número de pedido'};
  const sh = pedidosSheet_(false);
  if(!sh) return {ok:false,msg:'No existe hoja PEDIDOS'};
  const v = values_(sh), ix = idx_(v.headers), key = pedido_(pedido);
  const rows = v.data.filter(r => pedido_(getCell_(r, ix.pedido)) === key);
  if(!rows.length) return {ok:false,msg:'Pedido no encontrado',pedido:pedido};
  const ped = agrupar_(v.headers, rows)[0];
  if(clienteFiltro && nh_(ped.cliente).indexOf(clienteFiltro) === -1 && clienteFiltro.indexOf(nh_(ped.cliente)) === -1){
    return {ok:false,msg:'El pedido no coincide con el cliente informado',pedido:pedido};
  }
  const min = minutosPreparacion_(ped.hora_inicio, ped.hora_termino, ped.tiempo_preparacion_min);
  return {
    ok:true,
    pedido:ped.pedido,
    cliente:ped.cliente,
    vendedor:ped.vendedor,
    pikeador:ped.pikeador,
    status:ped.status,
    estado:ped.status,
    fecha:ped.fecha,
    total_productos:ped.total_productos,
    total_unidades:ped.total_unidades,
    hora_inicio:ped.hora_inicio,
    hora_termino:ped.hora_termino,
    tiempo_preparacion_min:min,
    tiempo_preparacion_texto:formatearMinutos_(min),
    avance:avancePedido_(ped.status),
    pasos:pasosPedido_(ped.status),
    serverTime:new Date().toISOString()
  };
}

function listarBD_(name){
  name = normName_(name);
  let sh = null;
  if(name === 'PEDIDOS' || name === 'PEDIDO') sh = pedidosSheet_(false);
  else if(name === 'PIKEADORES') sh = ss_().getSheetByName(SHEET_PIKEADORES);
  else if(name === 'MAESTRA') sh = ss_().getSheetByName(SHEET_MAESTRA);
  else if(name === 'ALERTAS_PEDIDOS') sh = ss_().getSheetByName(SHEET_ALERTAS);
  else if(name === 'MOVIMIENTO' || name === 'BD_MOVIMIENTO' || name === 'BD-MOVIMIENTO') sh = movimientoSheet_(true);
  else sh = ss_().getSheetByName(name);
  if(!sh) return {ok:true,sheet:name,headers:[],data:[],rows:[],values:[]};
  const v = values_(sh);
  return {ok:true,sheet:sh.getName(),headers:v.headers,data:v.data,rows:v.data,values:v.data,total:v.data.length};
}

function listarPikeadores_(){
  const sh = getOrCreate_(SHEET_PIKEADORES, ['nombre','fecha_creacion']);
  ensureCols_(sh, ['nombre','fecha_creacion']);
  const v = values_(sh);
  const iN = find_(v.headers, ['nombre','pikeador','picker','preparador','responsable','usuario'], 0);
  const iF = find_(v.headers, ['fecha_creacion','fecha'], 1);
  const seen = {}, data = [];
  v.data.forEach(r=>{ const n=t_(r[iN]); if(!n || seen[n.toUpperCase()]) return; seen[n.toUpperCase()]=1; data.push({nombre:n,fecha_creacion:t_(r[iF])}); });
  data.sort((a,b)=>a.nombre.localeCompare(b.nombre));
  return {ok:true,sheet:SHEET_PIKEADORES,headers:v.headers,data:data,pikeadores:data,total:data.length};
}

function agregarPikeador_(nombre){
  nombre = t_(nombre);
  if(!nombre) return {ok:false,msg:'Nombre de pikeador vacío'};
  const sh = getOrCreate_(SHEET_PIKEADORES, ['nombre','fecha_creacion']);
  ensureCols_(sh, ['nombre','fecha_creacion']);
  const v = values_(sh);
  const iN = find_(v.headers, ['nombre','pikeador','picker','preparador','responsable','usuario'],0);
  const existe = v.data.some(r=>t_(r[iN]).toUpperCase() === nombre.toUpperCase());
  if(!existe) sh.appendRow([nombre, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
  return {ok:true,msg:existe?'Pikeador ya existía':'Pikeador creado',nombre:nombre,data:listarPikeadores_().data};
}

function listarMaestra_(){
  const sh = ss_().getSheetByName(SHEET_MAESTRA);
  if(!sh) return {ok:true,sheet:SHEET_MAESTRA,headers:[],data:[],productos:[]};
  const v = values_(sh);
  const iC = find_(v.headers, ['codigo','código','cod','sku','producto','item'],0);
  const iD = find_(v.headers, ['descripcion','descripción','detalle','nombre'],1);
  const iQ = find_(v.headers, ['cantidad','unidades','stock'],2);
  const data = v.data.map(r=>({codigo:code_(r[iC]),descripcion:t_(r[iD]),cantidad:n_(r[iQ])})).filter(x=>x.codigo||x.descripcion);
  return {ok:true,sheet:SHEET_MAESTRA,headers:v.headers,data:data,productos:data,total:data.length};
}
function buscarProducto_(codigo){
  codigo = code_(codigo);
  if(!codigo) return {ok:false,msg:'Código vacío'};
  const p = (listarMaestra_().productos || []).find(x=>code_(x.codigo)===codigo);
  return p ? Object.assign({ok:true},p) : {ok:false,msg:'Código no encontrado',codigo:codigo};
}


function collectRawHeaders_(items){
  const out = [];
  (items || []).forEach(it => {
    let headers = [];
    if(it && Array.isArray(it.__rawHeaders)) headers = it.__rawHeaders;
    else if(it && it.__raw && typeof it.__raw === 'object') headers = Object.keys(it.__raw);
    headers.forEach(h => {
      h = t_(h);
      if(h && !out.some(x => nh_(x) === nh_(h))) out.push(h);
    });
  });
  return out;
}
function isCanonicalPedidoHeader_(h){
  const n = nh_(h);
  const canonical = ['id','fecha','pedido','cliente','vendedor','pikeador','codigo','código','cod','sku','descripcion','descripción','detalle','ubicacion','ubicación','cantidad','unidades','status','estado','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora inicio','inicio_preparacion','fecha_asignacion','hora_termino','hora termino','hora término','termino_preparacion','fecha_termino','tiempo_preparacion_min','minutos_preparacion','import_raw_json','import_headers_json'];
  return canonical.some(x => nh_(x) === n);
}
function safeJson_(v){
  try{return JSON.stringify(v || {});}catch(e){return '';}
}

function valorPedido_(obj, fallback){
  // Regla de importación: NÚMERO / NUMERO / Nº / N° también representa el campo PEDIDO.
  // Se revisan variantes frecuentes para que el servidor no pierda el dato aunque el front cambie.
  obj = obj || {};
  const keys = Object.keys(obj);
  const aliases = ['pedido','nro pedido','numero pedido','número pedido','numero','número','n°','nº','no','nro','orden','orden pedido','orden de pedido'];
  for(const a of aliases){
    const k = keys.find(x => nh_(x) === nh_(a));
    if(k && t_(obj[k])) return t_(obj[k]);
  }
  return t_(fallback || '');
}
function addRawExtrasToRow_(rowObj, item){
  if(!item || !item.__raw || typeof item.__raw !== 'object') return rowObj;
  Object.keys(item.__raw).forEach(h => {
    if(!h || isCanonicalPedidoHeader_(h)) return;
    rowObj[h] = item.__raw[h];
  });
  rowObj.import_raw_json = safeJson_(item.__raw || {});
  rowObj.import_headers_json = safeJson_(Array.isArray(item.__rawHeaders) ? item.__rawHeaders : Object.keys(item.__raw || {}));
  return rowObj;
}

function guardarPedido_(b){
  const sh = pedidosSheet_(true);
  const baseCols = ['id','fecha','pedido','cliente','vendedor','pikeador','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min','import_raw_json','import_headers_json'];
  const items = parseItems_(b.items || b.productos || b.rows || b.detalle);
  // Preserva información importada tal como viene desde el XLSX/CSV: guarda JSON completo
  // y agrega columnas extra si el archivo trae campos que no existen en la BD.
  const rawCols = collectRawHeaders_(items.concat([b])).filter(h => !isCanonicalPedidoHeader_(h));
  ensureCols_(sh, baseCols.concat(rawCols));
  const before = values_(sh);
  const headers = before.headers;
  const ix = idx_(headers);
  const pedido = valorPedido_(b, b.order || b.nro_pedido);
  if(!pedido && !items.length) return {ok:false,msg:'Pedido vacío'};

  const base = {
    fecha:t_(b.fecha)||Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HH:mm:ss'),
    pedido:pedido,
    cliente:t_(b.cliente),
    vendedor:t_(b.vendedor),
    pikeador:limpiarPikeador_(b.pikeador, b.cliente),
    status:estadoSeguro_(b.status||b.estado,'PENDIENTE')
  };

  const porClave = {};
  const porPedidoCodigo = {};
  before.data.forEach((r,i)=>{
    const obj = rowToObj_(r, ix);
    const clave = uniquePedidoProductoKey_(obj.pedido, obj.codigo, obj.descripcion, obj.ubicacion);
    if(clave) porClave[clave] = {row:i+2,obj:obj};
    const pc = pedidoCodigoKey_(obj.pedido, obj.codigo);
    if(pc && !porPedidoCodigo[pc]) porPedidoCodigo[pc] = {row:i+2,obj:obj};
  });

  const rows = (items.length?items:[b]).map(it=>{
    const row = {
      id:id_('PD'),
      fecha:t_(it.fecha || base.fecha) || base.fecha,
      pedido:valorPedido_(Object.assign({}, it.__raw || {}, it), pedido),
      cliente:t_(it.cliente||base.cliente),
      vendedor:t_(it.vendedor||base.vendedor),
      pikeador:limpiarPikeador_(it.pikeador || base.pikeador, it.cliente || base.cliente),
      codigo:code_(it.codigo||it.CODIGO||it.producto||it.sku),
      descripcion:t_(it.descripcion||it.DESCRIPCION||it.detalle||it.nombre),
      ubicacion:t_(it.ubicacion||it.UBICACION),
      cantidad:n_(it.cantidad||it.CANTIDAD||1)||1,
      status:estadoSeguro_(it.status||it.estado,base.status||'PENDIENTE'),
      hora_inicio:t_(it.hora_inicio || it.horaInicio || it.inicio_preparacion || it.fecha_inicio || ''),
      hora_termino:t_(it.hora_termino || it.horaTermino || it.termino_preparacion || it.fecha_termino || ''),
      tiempo_preparacion_min:n_(it.tiempo_preparacion_min || it.tiempoPreparacionMin || it.minutos_preparacion || 0),
      import_raw_json:safeJson_(it.__raw || {}),
      import_headers_json:safeJson_(it.__rawHeaders || [])
    };
    return addRawExtrasToRow_(row, it);
  }).filter(x=>x.pedido && (x.codigo || x.descripcion));

  const nuevos = [];
  let actualizados = 0, sinCambios = 0, duplicadosArchivo = 0;
  const vistos = {};
  const statusAlertas = {};

  rows.forEach(row=>{
    const key = uniquePedidoProductoKey_(row.pedido, row.codigo, row.descripcion, row.ubicacion);
    const pc = pedidoCodigoKey_(row.pedido, row.codigo);
    const marker = key || pc || (pedido_(row.pedido)+'|'+nh_(row.descripcion));
    if(vistos[marker]){ duplicadosArchivo++; return; }
    vistos[marker] = true;

    const found = porClave[key] || porPedidoCodigo[pc];
    if(!found){
      nuevos.push(row);
      if(key) porClave[key] = {row:0,obj:row};
      if(pc) porPedidoCodigo[pc] = {row:0,obj:row};
      registrarStatusPendiente_(statusAlertas, row.pedido, row.status, row.cliente, row.vendedor, row.pikeador, 'NUEVO_PEDIDO');
      return;
    }

    const changes = diffImportRow_(found.obj, row);
    if(!changes.length){ sinCambios++; return; }

    changes.forEach(ch=>{
      const colIndex = ix[ch.field];
      if(colIndex >= 0){ sh.getRange(found.row, colIndex + 1).setValue(ch.value); }
    });
    if(changes.some(ch => ch.field === 'status') && estadoSeguro_(found.obj.status, '') !== estadoSeguro_(row.status, '')){
      registrarStatusPendiente_(statusAlertas, row.pedido, row.status, row.cliente || found.obj.cliente, row.vendedor || found.obj.vendedor, row.pikeador || found.obj.pikeador, 'IMPORTACION_STATUS');
    }
    actualizados++;
  });

  appendObjects_(sh, headers, nuevos);
  if(base.pikeador) agregarPikeador_(base.pikeador);
  nuevos.forEach(x=>{ if(x.pikeador) agregarPikeador_(x.pikeador); });
  rows.forEach(x=>{ if(x.pikeador) agregarPikeador_(x.pikeador); });

  const alertas_emitidas = Object.keys(statusAlertas).map(k => {
    const a = statusAlertas[k];
    return emitirAlertaStatus_(a.pedido, a.estado, a.cliente, a.vendedor, a.pikeador, a.origen, true);
  }).filter(x => x && x.ok);

  return {
    ok:true,
    msg:'Importación validada contra PEDIDOS: solo nuevos y cambios fueron aplicados',
    recibidos:rows.length,
    insertados:nuevos.length,
    actualizados:actualizados,
    sin_cambios:sinCambios,
    duplicados_archivo:duplicadosArchivo,
    alertas_emitidas:alertas_emitidas.length,
    voces:alertas_emitidas.map(a => a.voice),
    pedido:pedido
  };
}

function rowToObj_(r, ix){
  return {
    fecha:getCell_(r, ix.fecha),
    pedido:getCell_(r, ix.pedido),
    cliente:getCell_(r, ix.cliente),
    vendedor:getCell_(r, ix.vendedor),
    pikeador:limpiarPikeador_(getCell_(r, ix.pikeador), getCell_(r, ix.cliente)),
    codigo:code_(getCell_(r, ix.codigo)),
    descripcion:getCell_(r, ix.descripcion),
    ubicacion:getCell_(r, ix.ubicacion),
    cantidad:getNumCell_(r, ix.cantidad) || 1,
    status:estadoSeguro_(getCell_(r, ix.status), 'PENDIENTE'),
    hora_inicio:getCell_(r, ix.hora_inicio),
    hora_termino:getCell_(r, ix.hora_termino),
    tiempo_preparacion_min:getNumCell_(r, ix.tiempo_preparacion_min)
  };
}
function pedidoCodigoKey_(pedido, codigo){
  const p = pedido_(pedido), c = code_(codigo);
  return p && c ? p + '|' + c : '';
}
function diffImportRow_(oldObj, newObj){
  const fields = ['fecha','cliente','vendedor','pikeador','descripcion','ubicacion','cantidad','status'];
  const changes = [];
  fields.forEach(f=>{
    const oldVal = f === 'cantidad' ? n_(oldObj[f]) : nh_(oldObj[f]);
    const newVal = f === 'cantidad' ? n_(newObj[f]) : nh_(newObj[f]);
    if(String(oldVal) !== String(newVal)) changes.push({field:f,value:newObj[f]});
  });
  return changes;
}


/**
 * Limpia la columna STATUS/ESTADO en PEDIDOS.
 * Regla estricta: el estado solo puede ser uno de los estados permitidos.
 * Si encuentra una descripción de producto u otro texto en ESTADO, lo cambia a PENDIENTE.
 * Esto evita que la TD muestre descripciones en el campo Estado.
 */
function repararEstadosPedidos_(includeData){
  const sh = pedidosSheet_(false);
  if(!sh) return {ok:true,msg:'No existe hoja PEDIDOS',corregidos:0};
  ensureCols_(sh, ['status']);
  const v = values_(sh);
  const ix = idx_(v.headers);
  if(ix.status < 0) return {ok:false,msg:'No se encontró columna status/estado',corregidos:0};
  if(!v.data.length) return {ok:true,msg:'Sin datos para reparar',corregidos:0};

  const range = sh.getRange(2, ix.status + 1, v.data.length, 1);
  const vals = range.getDisplayValues();
  let corregidos = 0;
  const salida = vals.map(row => {
    const raw = t_(row[0]);
    const seguro = estadoSeguro_(raw, '');
    if(!seguro){
      corregidos++;
      return ['PENDIENTE'];
    }
    if(raw !== seguro){
      corregidos++;
      return [seguro];
    }
    return [raw || 'PENDIENTE'];
  });
  if(corregidos){
    range.setValues(salida);
    SpreadsheetApp.flush();
  }
  return {
    ok:true,
    msg:corregidos ? 'Estados reparados correctamente' : 'Estados correctos',
    corregidos:corregidos,
    hoja:sh.getName(),
    columna:v.headers[ix.status],
    serverTime:new Date().toISOString()
  };
}


/**
 * Limpia la columna PIKEADOR en PEDIDOS.
 * Regla: PIKEADOR no puede ser igual al CLIENTE, ni ser un estado.
 * Si detecta ese cruce, deja PIKEADOR vacío para que se asigne desde la maestra PIKEADORES.
 */
function repararPikeadoresPedidos_(includeData){
  const sh = pedidosSheet_(false);
  if(!sh) return {ok:true,msg:'No existe hoja PEDIDOS',corregidos:0};
  ensureCols_(sh, ['pikeador']);
  const v = values_(sh);
  const ix = idx_(v.headers);
  if(ix.pikeador < 0) return {ok:false,msg:'No se encontró columna pikeador',corregidos:0};
  if(!v.data.length) return {ok:true,msg:'Sin datos para reparar',corregidos:0};

  const range = sh.getRange(2, ix.pikeador + 1, v.data.length, 1);
  const vals = range.getDisplayValues();
  let corregidos = 0;
  const salida = vals.map((row, i) => {
    const actual = t_(row[0]);
    const cliente = getCell_(v.data[i], ix.cliente);
    const limpio = limpiarPikeador_(actual, cliente);
    if(actual !== limpio){ corregidos++; return [limpio]; }
    return [actual];
  });
  if(corregidos){
    range.setValues(salida);
    SpreadsheetApp.flush();
  }
  return {
    ok:true,
    msg:corregidos ? 'Pikeadores cruzados corregidos' : 'Pikeadores correctos',
    corregidos:corregidos,
    hoja:sh.getName(),
    columna:v.headers[ix.pikeador],
    serverTime:new Date().toISOString()
  };
}

function uniquePedidoProductoKey_(pedido, codigo, descripcion, ubicacion){
  const p = pedido_(pedido);
  if(!p) return '';
  const c = code_(codigo);
  const d = nh_(descripcion);
  const u = nh_(ubicacion);
  return [p,c,d,u].join('|');
}

function registrarStatusPendiente_(bucket, pedido, estado, cliente, vendedor, pikeador, origen){
  const p = t_(pedido);
  const e = estadoSeguro_(estado, '');
  if(!p || !e) return;
  const k = pedido_(p) + '|' + e;
  if(!bucket[k]) bucket[k] = {pedido:p, estado:e, cliente:t_(cliente), vendedor:t_(vendedor), pikeador:t_(pikeador), origen:t_(origen || 'STATUS')};
}

function alertaTipoEstado_(estado){
  estado = estadoSeguro_(estado, '');
  return estado ? 'STATUS_' + nh_(estado) : 'STATUS';
}

function clientePedido_(pedido){
  pedido = t_(pedido);
  if(!pedido) return '';
  try{
    const sh = pedidosSheet_(false);
    if(!sh) return '';
    const v = values_(sh);
    const ix = idx_(v.headers);
    const key = pedido_(pedido);
    for(let i=0; i<v.data.length; i++){
      if(pedido_(getCell_(v.data[i], ix.pedido)) === key){
        const cliente = getCell_(v.data[i], ix.cliente);
        if(cliente) return cliente;
      }
    }
  }catch(e){}
  return '';
}

function mensajeVozEstado_(pedido, estado, pikeador, cliente){
  pedido = t_(pedido);
  estado = estadoSeguro_(estado, 'PENDIENTE');
  cliente = t_(cliente) || clientePedido_(pedido);
  const quien = t_(pikeador) ? ' Pikeador asignado: ' + t_(pikeador) + '.' : '';
  const pedidoTxt = pedido ? 'pedido ' + pedido : 'pedido sin número';
  const clienteTxt = cliente ? ', cliente ' + cliente : '';
  const mensajes = {
    'PENDIENTE':'Pedido ' + pedido + ' quedó pendiente.' + quien,
    'PREPARACION':'Pedido ' + pedido + ' enviado a preparación.' + quien,
    'RECIBIDO':'Pedido ' + pedido + ' fue recibido.' + quien,
    'RECEPCIONADO':'Pedido ' + pedido + ' fue recepcionado.' + quien,
    'DESPACHADO':'Pedido ' + pedido + ' fue despachado.' + quien,
    'EN RUTA':'Pedido ' + pedido + ' está en ruta.' + quien,
    'ENTREGADO':'Pedido ' + pedido + ' fue entregado.' + quien,
    'TERMINADO':'Pedido terminado, ' + pedidoTxt + clienteTxt,
    'CANCELADO':'Pedido ' + pedido + ' fue cancelado.' + quien
  };
  return mensajes[estado] || ('Pedido ' + pedido + ' cambió a estado ' + estado + '.' + quien);
}

function emitirAlertaStatus_(pedido, estado, cliente, vendedor, pikeador, origen, actualizarPedidos){
  pedido = t_(pedido);
  estado = estadoSeguro_(estado, '');
  if(!pedido || !estado) return {ok:false,msg:'Pedido o estado vacío'};
  const tipo = alertaTipoEstado_(estado);
  const mensaje = mensajeVozEstado_(pedido, estado, pikeador, cliente);
  const token = 'ALERTA-' + nh_(estado) + '-' + Date.now() + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  let count = 0;

  if(actualizarPedidos){
    const sh = pedidosSheet_(true);
    const cols = ['id','fecha','pedido','cliente','vendedor','pikeador','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min'];
    ensureCols_(sh, cols);
    const v = values_(sh), ix = idx_(v.headers);
    const key = pedido_(pedido);
    v.data.forEach((r,i)=>{
      if(pedido_(getCell_(r, ix.pedido)) !== key) return;
      const row = i + 2;
      if(ix.alerta_token >= 0) sh.getRange(row, ix.alerta_token + 1).setValue(token);
      if(ix.alerta_ts >= 0) sh.getRange(row, ix.alerta_ts + 1).setValue(ts);
      if(ix.alerta_tipo >= 0) sh.getRange(row, ix.alerta_tipo + 1).setValue(tipo);
      if(ix.alerta_mensaje >= 0) sh.getRange(row, ix.alerta_mensaje + 1).setValue(mensaje);
      let inicioFinal = getCell_(r, ix.hora_inicio);
      if(estado === 'PREPARACION' && ix.hora_inicio >= 0 && !inicioFinal){
        sh.getRange(row, ix.hora_inicio + 1).setValue(ts);
        inicioFinal = ts;
      }
      if(estado === 'TERMINADO'){
        if(ix.hora_inicio >= 0 && !inicioFinal){
          sh.getRange(row, ix.hora_inicio + 1).setValue(ts);
          inicioFinal = ts;
        }
        if(ix.hora_termino >= 0) sh.getRange(row, ix.hora_termino + 1).setValue(ts);
        if(ix.tiempo_preparacion_min >= 0) sh.getRange(row, ix.tiempo_preparacion_min + 1).setValue(minutosPreparacion_(inicioFinal, ts, 0));
      }
      if(!cliente) cliente = getCell_(r, ix.cliente);
      if(!vendedor) vendedor = getCell_(r, ix.vendedor);
      if(!pikeador) pikeador = limpiarPikeador_(getCell_(r, ix.pikeador), getCell_(r, ix.cliente));
      count++;
    });
  }

  registrarAlerta_(pedido, cliente, pikeador, tipo, mensaje, token, ts, estado, vendedor, origen || 'STATUS');
  SpreadsheetApp.flush();
  return {ok:true,msg:mensaje,pedido:pedido,status:estado,actualizados:count,alerta_token:token,alerta_ts:ts,alerta_tipo:tipo,alerta_mensaje:mensaje,voice:mensaje};
}

function listarAlertas_(p){
  const sh = getOrCreate_(SHEET_ALERTAS, ['id','fecha','pedido','cliente','pikeador','tipo','mensaje','token','status','vendedor','origen']);
  ensureCols_(sh, ['id','fecha','pedido','cliente','pikeador','tipo','mensaje','token','status','vendedor','origen']);
  const v = values_(sh);
  const limit = Math.max(1, Math.min(500, n_(p.limit || p.limite || 100) || 100));
  const data = v.data.slice(-limit).reverse();
  return {ok:true,sheet:sh.getName(),headers:v.headers,data:data,rows:data,values:data,total:v.data.length,limit:limit,serverTime:new Date().toISOString()};
}

function actualizarEstadoPedido_(b){
  const pedido = valorPedido_(b, b.order || b.nro_pedido);
  const estado = estadoSeguro_(b.status || b.estado || b.nuevo_estado || b.state, '');
  const pikeador = t_(b.pikeador || b.picker || b.preparador || '');
  if(!pedido) return {ok:false,msg:'Pedido vacío'};
  if(!estado) return {ok:false,msg:'Estado vacío'};

  const sh = pedidosSheet_(true);
  const cols = ['id','fecha','pedido','cliente','vendedor','pikeador','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min'];
  ensureCols_(sh, cols);
  const v = values_(sh), ix = idx_(v.headers);
  const key = pedido_(pedido);
  let count = 0, cliente = '', vendedor = '', pikeadorFinal = pikeador;
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  v.data.forEach((r,i)=>{
    if(pedido_(getCell_(r, ix.pedido)) !== key) return;
    const row = i + 2;
    if(ix.status >= 0) sh.getRange(row, ix.status + 1).setValue(estado);
    if(pikeador && ix.pikeador >= 0) sh.getRange(row, ix.pikeador + 1).setValue(pikeador);
    let inicioFinal = getCell_(r, ix.hora_inicio);
    if((estado === 'PREPARACION' || pikeador) && ix.hora_inicio >= 0 && !inicioFinal){
      sh.getRange(row, ix.hora_inicio + 1).setValue(ts);
      inicioFinal = ts;
    }
    if(estado === 'TERMINADO'){
      if(ix.hora_inicio >= 0 && !inicioFinal){
        sh.getRange(row, ix.hora_inicio + 1).setValue(ts);
        inicioFinal = ts;
      }
      if(ix.hora_termino >= 0) sh.getRange(row, ix.hora_termino + 1).setValue(ts);
      if(ix.tiempo_preparacion_min >= 0) sh.getRange(row, ix.tiempo_preparacion_min + 1).setValue(minutosPreparacion_(inicioFinal, ts, 0));
    }
    if(!cliente) cliente = getCell_(r, ix.cliente);
    if(!vendedor) vendedor = getCell_(r, ix.vendedor);
    if(!pikeadorFinal) pikeadorFinal = limpiarPikeador_(getCell_(r, ix.pikeador), getCell_(r, ix.cliente));
    count++;
  });

  if(!count){
    guardarPedido_(Object.assign({}, b, {pedido:pedido, status:estado, pikeador:pikeador}));
    return actualizarEstadoPedido_(b);
  }

  if(pikeadorFinal) agregarPikeador_(pikeadorFinal);
  registrarMovimientoEstado_(pedido, estado, cliente, vendedor, pikeadorFinal, count, ts);

  // Cualquier modificación de STATUS/ESTADO genera alerta inmediata y texto de voz.
  const alerta = emitirAlertaStatus_(pedido, estado, cliente, vendedor, pikeadorFinal, 'ACTUALIZAR_ESTADO', true);

  return {
    ok:true,
    msg:'Estado actualizado correctamente',
    pedido:pedido,
    status:estado,
    pikeador:pikeadorFinal,
    actualizados:count,
    alerta_token:alerta.alerta_token,
    alerta_ts:alerta.alerta_ts,
    alerta_tipo:alerta.alerta_tipo,
    alerta_mensaje:alerta.alerta_mensaje,
    voice:alerta.voice,
    serverTime:new Date().toISOString()
  };
}

function registrarMovimientoEstado_(pedido, estado, cliente, vendedor, pikeador, cantidadFilas, ts){
  const sh = getOrCreate_('MOVIMIENTO_PEDIDOS', ['id','fecha','pedido','status','cliente','vendedor','pikeador','cantidad_items','origen']);
  ensureCols_(sh, ['id','fecha','pedido','status','cliente','vendedor','pikeador','cantidad_items','origen']);
  sh.appendRow([id_('MVP'), ts, pedido, estado, cliente, vendedor, pikeador, Number(cantidadFilas||0), 'ACTUALIZAR_ESTADO']);
}

function alertaPedido_(b, tipo, mensaje, cambiarEstado){
  const pedido = valorPedido_(b, b.order || b.nro_pedido);
  if(!pedido) return {ok:false,msg:'Pedido vacío'};

  // Protección: si esta función se llama desde una acción genérica de alerta
  // pero viene un status/estado distinto de PREPARACION, la voz y el mensaje
  // deben corresponder al estado real. Ejemplo: TERMINADO => "Pedido terminado".
  const estadoSolicitado = estadoSeguro_(b.status || b.estado || b.nuevo_estado || b.state || (cambiarEstado ? 'PREPARACION' : ''), '');
  if(estadoSolicitado && estadoSolicitado !== 'PREPARACION'){
    tipo = alertaTipoEstado_(estadoSolicitado);
    mensaje = mensajeVozEstado_(pedido, estadoSolicitado, b.pikeador || b.picker || b.preparador || '', b.cliente || '');
    cambiarEstado = false;
  }

  const sh = pedidosSheet_(true);
  const cols = ['id','fecha','pedido','cliente','vendedor','pikeador','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min'];
  ensureCols_(sh, cols);
  const v = values_(sh), ix = idx_(v.headers);
  const token = 'ALERTA-' + Date.now() + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const key = pedido_(pedido);
  let count = 0, cliente = t_(b.cliente), pikeador = t_(b.pikeador);
  v.data.forEach((r,i)=>{
    if(pedido_(getCell_(r, ix.pedido)) !== key) return;
    const row = i+2;
    if(ix.alerta_token>=0) sh.getRange(row,ix.alerta_token+1).setValue(token);
    if(ix.alerta_ts>=0) sh.getRange(row,ix.alerta_ts+1).setValue(ts);
    if(ix.alerta_tipo>=0) sh.getRange(row,ix.alerta_tipo+1).setValue(tipo);
    if(ix.alerta_mensaje>=0) sh.getRange(row,ix.alerta_mensaje+1).setValue(mensaje);
    if(cambiarEstado && ix.status>=0) sh.getRange(row,ix.status+1).setValue('PREPARACION');
    if(pikeador && ix.pikeador>=0) sh.getRange(row,ix.pikeador+1).setValue(pikeador);
    if((cambiarEstado || pikeador) && ix.hora_inicio>=0 && !getCell_(r, ix.hora_inicio)) sh.getRange(row,ix.hora_inicio+1).setValue(ts);
    if(!cliente) cliente = getCell_(r, ix.cliente);
    if(!pikeador) pikeador = getCell_(r, ix.pikeador);
    count++;
  });
  if(!count){ guardarPedido_(Object.assign({},b,{pedido:pedido,status:cambiarEstado?'PREPARACION':'PENDIENTE'})); return alertaPedido_(b,tipo,mensaje,cambiarEstado); }
  if(pikeador) agregarPikeador_(pikeador);
  registrarAlerta_(pedido,cliente,pikeador,tipo,mensaje,token,ts);
  return {ok:true,msg:mensaje,pedido:pedido,actualizados:count,alerta_token:token,alerta_ts:ts,alerta_tipo:tipo,alerta_mensaje:mensaje,voice:mensaje};
}

function asignarPikeador_(b){
  const pedido = t_(b.pedido || b.numero || b.nro_pedido || b.order), pikeador = t_(b.pikeador || b.nombre);
  if(!pedido || !pikeador) return {ok:false,msg:'Pedido o pikeador vacío'};
  const sh = pedidosSheet_(true);
  ensureCols_(sh,['pedido','pikeador','hora_inicio','hora_termino','tiempo_preparacion_min']);
  const v = values_(sh), ix = idx_(v.headers);
  let count=0, inicioRegistrado=0;
  const key=pedido_(pedido), ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  v.data.forEach((r,i)=>{
    if(pedido_(getCell_(r, ix.pedido))!==key) return;
    const row=i+2;
    if(ix.pikeador>=0) sh.getRange(row,ix.pikeador+1).setValue(pikeador);
    if(ix.hora_inicio>=0 && !getCell_(r, ix.hora_inicio)){
      sh.getRange(row,ix.hora_inicio+1).setValue(ts);
      inicioRegistrado++;
    }
    count++;
  });
  agregarPikeador_(pikeador);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Pikeador asignado. Hora de inicio registrada',pedido:pedido,pikeador:pikeador,hora_inicio:ts,inicio_registrado:inicioRegistrado,actualizados:count};
}

function registrarAlerta_(pedido,cliente,pikeador,tipo,mensaje,token,ts,status,vendedor,origen){
  const sh = getOrCreate_(SHEET_ALERTAS, ['id','fecha','pedido','cliente','pikeador','tipo','mensaje','token','status','vendedor','origen']);
  ensureCols_(sh, ['id','fecha','pedido','cliente','pikeador','tipo','mensaje','token','status','vendedor','origen']);
  const headers = values_(sh).headers;
  const obj = {
    id:id_('AL'),
    fecha:ts || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    pedido:pedido,
    cliente:cliente,
    pikeador:pikeador,
    tipo:tipo,
    mensaje:mensaje,
    token:token,
    status:estadoSeguro_(status || tipo, '') || '',
    vendedor:t_(vendedor),
    origen:t_(origen || 'ALERTA')
  };
  appendObjects_(sh, headers, [obj]);
}

function agrupar_(headers, rows){
  const ix = idx_(headers), map={};
  rows.forEach((r,i)=>{
    const ped = getCell_(r, ix.pedido) || ('SIN_PEDIDO_'+(i+1));
    const key = pedido_(ped);
    if(!map[key]){
      map[key]={
        key:key,
        pedido:ped,
        fecha:getCell_(r, ix.fecha),
        cliente:getCell_(r, ix.cliente),
        vendedor:getCell_(r, ix.vendedor),
        pikeador:limpiarPikeador_(getCell_(r, ix.pikeador), getCell_(r, ix.cliente)),
        status:estadoSeguro_(getCell_(r, ix.status), 'PENDIENTE'),
        alerta_token:getCell_(r, ix.alerta_token),
        alerta_ts:getCell_(r, ix.alerta_ts),
        alerta_tipo:getCell_(r, ix.alerta_tipo),
        alerta_mensaje:getCell_(r, ix.alerta_mensaje),
        hora_inicio:getCell_(r, ix.hora_inicio),
        hora_termino:getCell_(r, ix.hora_termino),
        tiempo_preparacion_min:getNumCell_(r, ix.tiempo_preparacion_min),
        total_unidades:0,total_productos:0,productos:[],rows:[]
      };
    }
    const p = map[key], cant = getNumCell_(r, ix.cantidad) || 1;
    ['fecha','cliente','vendedor','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino'].forEach(k=>{
      const val = getCell_(r, ix[k]);
      if(val) p[k]=val;
    });
    const minutosPrep = getNumCell_(r, ix.tiempo_preparacion_min);
    if(minutosPrep > 0) p.tiempo_preparacion_min = minutosPrep;
    const pkSeguro = limpiarPikeador_(getCell_(r, ix.pikeador), p.cliente || getCell_(r, ix.cliente));
    if(pkSeguro) p.pikeador = pkSeguro;
    const estadoLeido = estadoSeguro_(getCell_(r, ix.status), '');
    if(estadoLeido) p.status = estadoLeido;
    const prod = {codigo:getCell_(r, ix.codigo),descripcion:getCell_(r, ix.descripcion),ubicacion:getCell_(r, ix.ubicacion),cantidad:cant};
    if(prod.codigo || prod.descripcion){ p.productos.push(prod); p.total_productos++; p.total_unidades += cant; }
    p.rows.push(r);
  });
  return Object.values(map).sort((a,b)=>numPedido_(b.pedido)-numPedido_(a.pedido));
}


function fechaHoraPedido_(v){
  v = t_(v);
  if(!v) return null;
  let m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if(m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0));
  m = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if(m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0));
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function minutosPreparacion_(inicio, termino, directo){
  const d = n_(directo);
  if(d > 0) return d;
  const a = fechaHoraPedido_(inicio), b = fechaHoraPedido_(termino);
  if(!a || !b) return 0;
  return Math.max(0, Math.round((b.getTime() - a.getTime()) / 60000));
}
function formatearMinutos_(min){
  min = Math.round(n_(min));
  if(min <= 0) return '0 min';
  const h = Math.floor(min/60), m = min % 60;
  return h ? (h + 'h ' + String(m).padStart(2,'0') + 'm') : (m + ' min');
}
function avancePedido_(estado){
  estado = estadoSeguro_(estado, 'PENDIENTE');
  if(estado === 'CANCELADO') return 0;
  const orden = ['PENDIENTE','PREPARACION','RECIBIDO','DESPACHADO','TERMINADO'];
  const ix = Math.max(0, orden.indexOf(estado));
  return Math.round((ix / (orden.length - 1)) * 100);
}
function pasosPedido_(estado){
  const orden = ['PENDIENTE','PREPARACION','RECIBIDO','DESPACHADO','TERMINADO'];
  const actual = estadoSeguro_(estado, 'PENDIENTE');
  const ixActual = orden.indexOf(actual);
  return orden.map((x,i)=>({estado:x,activo:x===actual,completado:ixActual>=0 && i<=ixActual}));
}

function ss_(){ try{ if(SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID); }catch(e){} const ss=SpreadsheetApp.getActiveSpreadsheet(); if(!ss) throw new Error('No se pudo abrir la planilla. Revisa SPREADSHEET_ID.'); return ss; }
function pedidosSheet_(crear){ const ss=ss_(); let sh=ss.getSheetByName(SHEET_PEDIDOS)||ss.getSheetByName(SHEET_PEDIDO_ALT); if(!sh && crear) sh=getOrCreate_(SHEET_PEDIDOS,['id','fecha','pedido','cliente','vendedor','pikeador','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min']); return sh; }
function getOrCreate_(name, headers){ const ss=ss_(); let sh=ss.getSheetByName(name); if(!sh){ sh=ss.insertSheet(name); try{ if(sh.getMaxRows()>300) sh.deleteRows(301,sh.getMaxRows()-300); }catch(e){} try{ if(sh.getMaxColumns()>Math.max(headers.length,10)) sh.deleteColumns(Math.max(headers.length,10)+1,sh.getMaxColumns()-Math.max(headers.length,10)); }catch(e){} sh.getRange(1,1,1,headers.length).setValues([headers]); } else if(sh.getLastRow()===0){ sh.getRange(1,1,1,headers.length).setValues([headers]); } return sh; }
function values_(sh){ const lr=sh.getLastRow(), lc=sh.getLastColumn(); if(lr<1||lc<1) return {headers:[],data:[]}; const v=sh.getRange(1,1,lr,lc).getDisplayValues(); return {headers:v[0].map(t_),data:v.slice(1).filter(r=>r.some(c=>t_(c)))}; }
function ensureCols_(sh, req){ let lc=Math.max(1,sh.getLastColumn()); let h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(t_); if(h.every(x=>!x)){ sh.getRange(1,1,1,req.length).setValues([req]); return; } req.forEach(name=>{ if(find_(h,[name],-1)!==-1) return; let blank=h.findIndex(x=>!x); if(blank!==-1){sh.getRange(1,blank+1).setValue(name); h[blank]=name; return;} const max=sh.getMaxColumns(); if(max>lc){lc++; sh.getRange(1,lc).setValue(name); h.push(name); return;} const total=ss_().getSheets().reduce((s,x)=>s+x.getMaxRows()*x.getMaxColumns(),0); if(total+sh.getMaxRows()>9900000) throw new Error('No se puede agregar la columna '+name+' por límite de 10.000.000 de celdas. Elimina filas/columnas vacías.'); sh.insertColumnAfter(lc); lc++; sh.getRange(1,lc).setValue(name); h.push(name); }); }
function appendObjects_(sh, headers, objs){ if(!objs.length) return 0; const rows=objs.map(o=>headers.map(h=>{ const k=Object.keys(o).find(x=>nh_(x)===nh_(h)); return k?o[k]:''; })); sh.getRange(sh.getLastRow()+1,1,rows.length,headers.length).setValues(rows); return rows.length; }
function idx_(h){
  // Mapeo estricto por encabezado. Esto evita que una columna por posición cruce datos.
  // Si un encabezado falta, devuelve -1 y el código no lee/escribe en una columna equivocada.
  return {
    id:find_(h,['id'],-1),
    fecha:find_(h,['fecha','fecha ingreso','fecha_pedido','fecha pedido','fecha creacion','fecha creación'],-1),
    pedido:find_(h,['pedido','nro pedido','numero pedido','número pedido','numero','número','n°','nº','nro','no','orden','orden pedido','orden de pedido','n° pedido','nº pedido'],-1),
    cliente:find_(h,['cliente','nombre cliente','razon social','razón social'],-1),
    vendedor:find_(h,['vendedor','responsable','ejecutivo'],-1),
    pikeador:find_(h,['pikeador','picker','preparador','asignado','asignado a'],-1),
    codigo:find_(h,['codigo','código','cod','sku','codigo producto','código producto'],-1),
    descripcion:find_(h,['descripcion','descripción','detalle','nombre','producto','nombre producto','descripcion producto','descripción producto'],-1),
    ubicacion:find_(h,['ubicacion','ubicación','bodega','ubicacion producto','ubicación producto'],-1),
    cantidad:find_(h,['cantidad','unidades','cajas','qty','cant'],-1),
    status:find_(h,['status','estado','estado pedido','estado del pedido'],-1),
    alerta_token:find_(h,['alerta_token','token_alerta','alert token','alerta token'],-1),
    alerta_ts:find_(h,['alerta_ts','fecha_alerta','alert_ts','alerta fecha'],-1),
    alerta_tipo:find_(h,['alerta_tipo','tipo_alerta','tipo alerta'],-1),
    alerta_mensaje:find_(h,['alerta_mensaje','mensaje_alerta','mensaje alerta'],-1),
    hora_inicio:find_(h,['hora_inicio','hora inicio','inicio_preparacion','inicio preparación','fecha_inicio','fecha inicio','fecha_asignacion','fecha asignacion','asignado_en','asignado en'],-1),
    hora_termino:find_(h,['hora_termino','hora termino','hora término','termino_preparacion','término preparación','termino preparación','fecha_termino','fecha término','finalizado_en','finalizado en'],-1),
    tiempo_preparacion_min:find_(h,['tiempo_preparacion_min','tiempo preparacion min','tiempo preparación min','minutos_preparacion','minutos preparación','duracion_min','duración min'],-1)
  };
}
function getCell_(r, idx){ return (idx == null || idx < 0 || !Array.isArray(r)) ? '' : t_(r[idx]); }
function getNumCell_(r, idx){ return (idx == null || idx < 0 || !Array.isArray(r)) ? 0 : n_(r[idx]); }
function find_(h,names,fb){ const a=h.map(nh_); for(const n of names){ const i=a.indexOf(nh_(n)); if(i!==-1) return i; } return fb; }
function out_(o,cb){ const j=JSON.stringify(o); return ContentService.createTextOutput(cb?cb+'('+j+');':j).setMimeType(cb?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON); }
function body_(e){ try{ if(e&&e.parameter&&e.parameter.data) return JSON.parse(e.parameter.data); }catch(x){} try{ if(e&&e.postData&&e.postData.contents) return JSON.parse(e.postData.contents); }catch(x){} return e&&e.parameter?Object.assign({},e.parameter):{}; }
function parseItems_(v){ if(!v) return []; if(Array.isArray(v)) return v; try{ const x=JSON.parse(String(v)); return Array.isArray(x)?x:[]; }catch(e){ return []; } }

function limpiarPikeador_(pk, cliente){
  pk = t_(pk);
  cliente = t_(cliente);
  if(!pk) return '';
  if(cliente && nh_(pk) === nh_(cliente)) return '';
  if(estadoSeguro_(pk, '')) return '';
  return pk;
}
function t_(v){ return String(v==null?'':v).trim(); }
function n_(v){ const x=Number(String(v==null?'':v).replace(',','.')); return Number.isFinite(x)?x:0; }
function nh_(v){ return t_(v).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[\s_\-.]+/g,''); }
function normName_(v){ return t_(v).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[\s\-]+/g,'_'); }
function code_(v){ return t_(v).replace(/^'+/,'').replace(/\s+/g,'').toUpperCase(); }
function pedido_(v){ return t_(v).replace(/^'+/,'').replace(/\s+/g,'').toUpperCase(); }
function estado_(v){ return t_(v||'PENDIENTE').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function estadoSeguro_(v, defecto){
  const e = nh_(v);
  const map = {
    'PENDIENTE':'PENDIENTE',
    'PREPARACION':'PREPARACION',
    'ENPREPARACION':'PREPARACION',
    'RECIBIDO':'RECIBIDO',
    'RECEPCIONADO':'RECEPCIONADO',
    'DESPACHADO':'DESPACHADO',
    'ENRUTA':'EN RUTA',
    'ENTREGADO':'ENTREGADO',
    'TERMINADO':'TERMINADO',
    'CANCELADO':'CANCELADO',
    'ANULADO':'CANCELADO'
  };
  if(map[e]) return map[e];
  if(defecto === '') return '';
  const d = nh_(defecto || 'PENDIENTE');
  return map[d] || 'PENDIENTE';
}
function truth_(v){ v=t_(v).toLowerCase(); return ['1','true','si','sí','yes','todo','all'].includes(v); }
function id_(p){ return p+'-'+Date.now().toString(36).toUpperCase()+'-'+Utilities.getUuid().slice(0,8).toUpperCase(); }
function numPedido_(v){ const m=t_(v).match(/\d+/g); return m?Number(m.join('')):0; }
