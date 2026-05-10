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
const SHEET_VENDEDORES = 'VENDEDORES';
const SHEET_BODEGAS = 'BODEGAS';
const SHEET_CLIENTES = 'CLIENTES';
const SHEET_MAESTRA = 'MAESTRA';
const SHEET_MOV = 'MOVIMIENTO';
const SHEET_MOV_UBICACION = 'MOVIMIENTO_UBICACION';
const SHEET_MOV_PEDIDOS = 'MOVIMIENTO_PEDIDOS';
const SHEET_ALERTAS = 'ALERTAS_PEDIDOS';
const TZ = 'America/Santiago';

function doGet(e){
  e = e || {parameter:{}};
  const p = e.parameter || {};
  const accion = t_(p.accion || p.action || 'diagnostico').toLowerCase();
  const cb = t_(p.callback || p.jsonp || p.cb || p.callback_jsonp || '');
  try{
    let r;
    if((['diagnostico','test',''].includes(accion)) && (p.sheet || p.bd || p.hoja)) r = listarBD_(p.sheet || p.bd || p.hoja || '');
    else if(['diagnostico','test',''].includes(accion)) r = diagnostico_();
    else if(['reparar_estados','reparar_estados_pedidos','limpiar_estados'].includes(accion)) r = repararEstadosPedidos_();
    else if(['agregar','agregar_movimiento','guardar_movimiento','guardar_captura','agregar_ubicacion','guardar_ubicacion','guardar_ubicaciones','crear_ubicacion','crear_movimiento'].includes(accion)) r = agregarMovimiento_(p);
    else if(['editar','editar_movimiento','actualizar_movimiento','editar_ubicacion','actualizar_ubicacion','modificar_ubicacion','modificar_movimiento'].includes(accion)) r = editarMovimiento_(p);
    else if(['eliminar','eliminar_movimiento','borrar_movimiento','eliminar_ubicacion','borrar_ubicacion'].includes(accion)) r = eliminarMovimiento_(p);
    else if(['reparar_pikeadores','reparar_pikeadores_pedidos','limpiar_pikeadores'].includes(accion)) r = repararPikeadoresPedidos_();
    else if(['listar','listar_movimiento','listar_movimientos','cargar_movimiento','cargar_movimientos','listar_ubicaciones','cargar_ubicaciones','listar_movimiento_ubicaciones'].includes(accion)) r = listarMovimiento_(p);
    else if(['importar_movimiento','importar_movimientos','cargar_archivo_movimiento'].includes(accion)) r = importarMovimiento_(p);
    else if(['listar_pedidos','listar_pedidos_rapido','cargar_pedidos'].includes(accion)) r = listarPedidos_(p);
    else if(['siguiente_numero_pedido','generar_numero_pedido','proximo_pedido','next_pedido','next_order'].includes(accion)) r = siguienteNumeroPedido_(p);
    else if(['seguimiento_pedido','consulta_cliente_pedido','consultar_pedido_cliente','estado_pedido_cliente'].includes(accion)) r = seguimientoPedido_(p);
    else if(['listar_alertas','listar_alertas_pedidos','cargar_alertas','alertas_pedidos','ultimas_alertas'].includes(accion)) r = listarAlertas_(p);

    else if(['listar_movimiento_ubicacion','listar_movimientos_ubicacion','listar_bd_movimiento_ubicacion'].includes(accion)) r = listarMovimientoUbicacionBD_();
    else if(['agregar_movimiento_ubicacion','crear_movimiento_ubicacion','guardar_movimiento_ubicacion'].includes(accion)) r = agregarMovimientoUbicacionBD_(p);
    else if(['editar_movimiento_ubicacion','actualizar_movimiento_ubicacion','modificar_movimiento_ubicacion'].includes(accion)) r = editarMovimientoUbicacionBD_(p);
    else if(['eliminar_movimiento_ubicacion','borrar_movimiento_ubicacion'].includes(accion)) r = eliminarMovimientoUbicacionBD_(p);
    else if(['importar_movimiento_ubicacion','importar_movimientos_ubicacion'].includes(accion)) r = importarMovimientoUbicacionBD_(p);
    else if(['listar_movimiento_pedidos','listar_movimientos_pedidos','listar_bd_movimiento_pedidos'].includes(accion)) r = listarMovimientoPedidosBD_();
    else if(['agregar_movimiento_pedidos','crear_movimiento_pedidos','guardar_movimiento_pedidos'].includes(accion)) r = agregarMovimientoPedidosBD_(p);
    else if(['editar_movimiento_pedidos','actualizar_movimiento_pedidos','modificar_movimiento_pedidos'].includes(accion)) r = editarMovimientoPedidosBD_(p);
    else if(['eliminar_movimiento_pedidos','borrar_movimiento_pedidos'].includes(accion)) r = eliminarMovimientoPedidosBD_(p);
    else if(['importar_movimiento_pedidos','importar_movimientos_pedidos'].includes(accion)) r = importarMovimientoPedidosBD_(p);
    else if(accion === 'listar_bd') r = listarBD_(p.sheet || p.bd || p.hoja || '');
    else if(['listar_pikeadores','obtener_pikeadores','cargar_pikeadores','select_pikeadores'].includes(accion)) r = listarPikeadores_();
    else if(['agregar_pikeador','crear_pikeador'].includes(accion)) r = agregarPikeador_(p.nombre || p.pikeador || '');
    else if(['listar_vendedores','obtener_vendedores','cargar_vendedores','select_vendedores'].includes(accion)) r = listarVendedores_();
    else if(['agregar_vendedor','crear_vendedor','guardar_vendedor'].includes(accion)) r = agregarVendedor_(p.nombre || p.vendedor || '');
    else if(['listar_clientes','obtener_clientes','cargar_clientes','select_clientes'].includes(accion)) r = listarClientes_();
    else if(['agregar_cliente','crear_cliente','guardar_cliente'].includes(accion)) r = agregarCliente_(p);
    else if(['editar_cliente','actualizar_cliente','modificar_cliente'].includes(accion)) r = editarCliente_(p);
    else if(['eliminar_cliente','borrar_cliente'].includes(accion)) r = eliminarCliente_(p);
    else if(['listar_bodegas','obtener_bodegas','cargar_bodegas','select_bodegas'].includes(accion)) r = listarBodegas_();
    else if(['agregar_bodega','crear_bodega','guardar_bodega'].includes(accion)) r = agregarBodega_(p);
    else if(['editar_bodega','actualizar_bodega','modificar_bodega'].includes(accion)) r = editarBodega_(p);
    else if(['eliminar_bodega','borrar_bodega'].includes(accion)) r = eliminarBodega_(p);
    else if(['listar_maestra_bd','maestra_listar','cargar_maestra','bd_maestra','maestra','obtener_maestra','ver_maestra'].includes(accion)) r = listarMaestraBD_(p);
    else if(['agregar_maestra','crear_maestra','guardar_maestra'].includes(accion)) r = agregarMaestra_(p);
    else if(['editar_maestra','actualizar_maestra','modificar_maestra'].includes(accion)) r = editarMaestra_(p);
    else if(['eliminar_maestra','borrar_maestra'].includes(accion)) r = eliminarMaestra_(p);
    else if(['importar_maestra','importar_maestra_xlsx','cargar_archivo_maestra'].includes(accion)) r = importarMaestra_(p);
    else if(['listar_maestra'].includes(accion)) r = listarMaestra_();
    else if(['buscar','buscar_producto','buscar_maestra'].includes(accion)) r = buscarProducto_(p.codigo || p.cod || '');
    else if(['actualizar_estado_pedido','cambiar_estado_pedido','actualizar_estado','cambiar_estado'].includes(accion)) r = actualizarEstadoPedido_(p);
    else if(['editar_pedido_completo','editar_pedido','actualizar_pedido','modificar_pedido'].includes(accion)) r = editarPedidoCompleto_(p);
    else if(['preparar_pedido','enviar_preparacion','enviar_a_preparacion'].includes(accion)) r = alertaPedido_(p, 'PREPARACION', 'Pedido enviado a preparación', true);
    else if(['reenviar_alerta','reenviar_alerta_pedido','enviar_alerta','disparar_alerta_pedido'].includes(accion)) r = alertaPedido_(p, 'ALERTA', 'Pedido enviado a preparación', false);
    else if(['asignar_pikeador','asignar_pikeador_pedido'].includes(accion)) r = asignarPikeador_(p);
    else if(['verificar_importacion_pedidos','verificar_pedidos_importados','verificar_importacion'].includes(accion)) r = verificarImportacionPedidos_(p);
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
    if((['diagnostico','test',''].includes(accion)) && (b.sheet || b.bd || b.hoja)) r = listarBD_(b.sheet || b.bd || b.hoja || '');
    else if(['reparar_estados','reparar_estados_pedidos','limpiar_estados'].includes(accion)) r = repararEstadosPedidos_();
    else if(['agregar','agregar_movimiento','guardar_movimiento','guardar_captura','agregar_ubicacion','guardar_ubicacion','guardar_ubicaciones','crear_ubicacion','crear_movimiento'].includes(accion)) r = agregarMovimiento_(b);
    else if(['editar','editar_movimiento','actualizar_movimiento','editar_ubicacion','actualizar_ubicacion','modificar_ubicacion','modificar_movimiento'].includes(accion)) r = editarMovimiento_(b);
    else if(['eliminar','eliminar_movimiento','borrar_movimiento','eliminar_ubicacion','borrar_ubicacion'].includes(accion)) r = eliminarMovimiento_(b);
    else if(['reparar_pikeadores','reparar_pikeadores_pedidos','limpiar_pikeadores'].includes(accion)) r = repararPikeadoresPedidos_();
    else if(['listar_alertas','listar_alertas_pedidos','cargar_alertas','alertas_pedidos','ultimas_alertas'].includes(accion)) r = listarAlertas_(b);
    else if(['listar','listar_movimiento','listar_movimientos','cargar_movimiento','cargar_movimientos','listar_ubicaciones','cargar_ubicaciones','listar_movimiento_ubicaciones'].includes(accion)) r = listarMovimiento_(b);

    else if(['listar_movimiento_ubicacion','listar_movimientos_ubicacion','listar_bd_movimiento_ubicacion'].includes(accion)) r = listarMovimientoUbicacionBD_();
    else if(['agregar_movimiento_ubicacion','crear_movimiento_ubicacion','guardar_movimiento_ubicacion'].includes(accion)) r = agregarMovimientoUbicacionBD_(b);
    else if(['editar_movimiento_ubicacion','actualizar_movimiento_ubicacion','modificar_movimiento_ubicacion'].includes(accion)) r = editarMovimientoUbicacionBD_(b);
    else if(['eliminar_movimiento_ubicacion','borrar_movimiento_ubicacion'].includes(accion)) r = eliminarMovimientoUbicacionBD_(b);
    else if(['importar_movimiento_ubicacion','importar_movimientos_ubicacion'].includes(accion)) r = importarMovimientoUbicacionBD_(b);
    else if(['listar_movimiento_pedidos','listar_movimientos_pedidos','listar_bd_movimiento_pedidos'].includes(accion)) r = listarMovimientoPedidosBD_();
    else if(['agregar_movimiento_pedidos','crear_movimiento_pedidos','guardar_movimiento_pedidos'].includes(accion)) r = agregarMovimientoPedidosBD_(b);
    else if(['editar_movimiento_pedidos','actualizar_movimiento_pedidos','modificar_movimiento_pedidos'].includes(accion)) r = editarMovimientoPedidosBD_(b);
    else if(['eliminar_movimiento_pedidos','borrar_movimiento_pedidos'].includes(accion)) r = eliminarMovimientoPedidosBD_(b);
    else if(['importar_movimiento_pedidos','importar_movimientos_pedidos'].includes(accion)) r = importarMovimientoPedidosBD_(b);
    else if(accion === 'listar_bd') r = listarBD_(b.sheet || b.bd || b.hoja || '');
    else if(['importar_movimiento','importar_movimientos','cargar_archivo_movimiento'].includes(accion)) r = importarMovimiento_(b);
    else if(['seguimiento_pedido','consulta_cliente_pedido','consultar_pedido_cliente','estado_pedido_cliente'].includes(accion)) r = seguimientoPedido_(b);
    else if(['siguiente_numero_pedido','generar_numero_pedido','proximo_pedido','next_pedido','next_order'].includes(accion)) r = siguienteNumeroPedido_(b);
    else if(['actualizar_estado_pedido','cambiar_estado_pedido','actualizar_estado','cambiar_estado'].includes(accion)) r = actualizarEstadoPedido_(b);
    else if(['editar_pedido_completo','editar_pedido','actualizar_pedido','modificar_pedido'].includes(accion)) r = editarPedidoCompleto_(b);
    else if(['preparar_pedido','enviar_preparacion','enviar_a_preparacion'].includes(accion)) r = alertaPedido_(b, 'PREPARACION', 'Pedido enviado a preparación', true);
    else if(['reenviar_alerta','reenviar_alerta_pedido','enviar_alerta','disparar_alerta_pedido'].includes(accion)) r = alertaPedido_(b, 'ALERTA', 'Pedido enviado a preparación', false);
    else if(['asignar_pikeador','asignar_pikeador_pedido'].includes(accion)) r = asignarPikeador_(b);
    else if(['agregar_pikeador','crear_pikeador'].includes(accion)) r = agregarPikeador_(b.nombre || b.pikeador || '');
    else if(['listar_vendedores','obtener_vendedores','cargar_vendedores','select_vendedores'].includes(accion)) r = listarVendedores_();
    else if(['agregar_vendedor','crear_vendedor','guardar_vendedor'].includes(accion)) r = agregarVendedor_(b.nombre || b.vendedor || '');
    else if(['listar_clientes','obtener_clientes','cargar_clientes','select_clientes'].includes(accion)) r = listarClientes_();
    else if(['agregar_cliente','crear_cliente','guardar_cliente'].includes(accion)) r = agregarCliente_(b);
    else if(['editar_cliente','actualizar_cliente','modificar_cliente'].includes(accion)) r = editarCliente_(b);
    else if(['eliminar_cliente','borrar_cliente'].includes(accion)) r = eliminarCliente_(b);
    else if(['listar_bodegas','obtener_bodegas','cargar_bodegas','select_bodegas'].includes(accion)) r = listarBodegas_();
    else if(['agregar_bodega','crear_bodega','guardar_bodega'].includes(accion)) r = agregarBodega_(b);
    else if(['editar_bodega','actualizar_bodega','modificar_bodega'].includes(accion)) r = editarBodega_(b);
    else if(['eliminar_bodega','borrar_bodega'].includes(accion)) r = eliminarBodega_(b);
    else if(['listar_maestra_bd','maestra_listar','cargar_maestra','bd_maestra','maestra','obtener_maestra','ver_maestra'].includes(accion)) r = listarMaestraBD_(b);
    else if(['agregar_maestra','crear_maestra','guardar_maestra'].includes(accion)) r = agregarMaestra_(b);
    else if(['editar_maestra','actualizar_maestra','modificar_maestra'].includes(accion)) r = editarMaestra_(b);
    else if(['eliminar_maestra','borrar_maestra'].includes(accion)) r = eliminarMaestra_(b);
    else if(['importar_maestra','importar_maestra_xlsx','cargar_archivo_maestra'].includes(accion)) r = importarMaestra_(b);
    else if(['listar_maestra'].includes(accion)) r = listarMaestra_();
    else if(['verificar_importacion_pedidos','verificar_pedidos_importados','verificar_importacion'].includes(accion)) r = verificarImportacionPedidos_(b);
    else if(['guardar_pedido','guardar_pedido_rapido','importar_pedidos','importar_pedidos_rapido','importar_listado','enviar_base_datos','enviar_bd'].includes(accion)) r = guardarPedido_(b);
    else r = {ok:false,msg:'Acción POST no válida: '+accion};
    return outPost_(r, b);
  }catch(err){ return outPost_({ok:false,msg:String(err.message||err),stack:String(err.stack||'')}, (typeof b!=='undefined'?b:{})); }
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
  const headers = v.headers.length ? v.headers : movimientoHeaders_();
  const ix = movimientoIdx_(headers);
  const idBuscado = t_(b.id || b.ID || '');
  const rowParam = Number(b.rowNumber || b.fila || b.row || 0);
  let rowNumber = 0;
  if(rowParam >= 2 && rowParam <= sh.getLastRow()) rowNumber = rowParam;
  if(!rowNumber && idBuscado && ix.id >= 0){
    for(let i=0; i<v.data.length; i++){
      if(t_(getCell_(v.data[i], ix.id)) === idBuscado){ rowNumber = i + 2; break; }
    }
  }
  if(!rowNumber) return {ok:false,msg:'Movimiento no encontrado para editar'};

  const updates = {
    fecha_entrada:t_(b.fecha_entrada || b.fechaEntrada || ''),
    fecha_salida:t_(b.fecha_salida || b.fechaSalida || ''),
    ubicacion:t_(b.ubicacion || b['ubicación'] || ''),
    codigo:code_(b.codigo || b.cod || b.sku || ''),
    descripcion:t_(b.descripcion || b['descripción'] || b.detalle || ''),
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
  return {ok:true,msg:'Movimiento editado correctamente',accion:'editar',id:idBuscado,rowNumber:rowNumber,serverTime:new Date().toISOString()};
}

function eliminarMovimiento_(b){
  const sh = movimientoSheet_(true);
  const v = values_(sh);
  const ix = movimientoIdx_(v.headers);
  const idBuscado = t_(b.id || b.ID || '');
  const rowParam = Number(b.rowNumber || b.fila || b.row || 0);
  if(rowParam >= 2 && rowParam <= sh.getLastRow()){
    sh.deleteRow(rowParam);
    SpreadsheetApp.flush();
    return {ok:true,msg:'Movimiento eliminado correctamente',accion:'eliminar',id:idBuscado,rowNumber:rowParam,serverTime:new Date().toISOString()};
  }
  if(!idBuscado) return {ok:false,msg:'ID o fila vacía'};
  for(let i=0; i<v.data.length; i++){
    if(t_(getCell_(v.data[i], ix.id)) === idBuscado){
      sh.deleteRow(i + 2);
      SpreadsheetApp.flush();
      return {ok:true,msg:'Movimiento eliminado correctamente',accion:'eliminar',id:idBuscado,serverTime:new Date().toISOString()};
    }
  }
  return {ok:false,msg:'Movimiento no encontrado'};
}

function importarMovimiento_(b){
  let items = [];
  if(Array.isArray(b.rows)) items = b.rows;
  else if(Array.isArray(b.data)) items = b.data;
  else items = parseItems_(b.rows || b.data || b.items || b.movimientos || b.detalle);
  const modo = t_(b.modo || 'append').toLowerCase();
  let insertados = 0, actualizados = 0, omitidos = 0;
  items.forEach(raw => {
    raw = raw || {};
    const id = t_(raw.id || raw.ID || '');
    const rowNumber = Number(raw.rowNumber || raw.fila || raw.row || 0);
    const codigo = code_(raw.codigo || raw['código'] || raw.cod || raw.sku || '');
    if(!codigo){ omitidos++; return; }
    if((id || rowNumber) && modo !== 'append'){
      const r = editarMovimiento_(Object.assign({}, raw, {id:id,rowNumber:rowNumber,codigo:codigo}));
      if(r && r.ok) actualizados++; else { agregarMovimiento_(raw); insertados++; }
    }else{
      agregarMovimiento_(raw);
      insertados++;
    }
  });
  SpreadsheetApp.flush();
  return {ok:true,msg:'Importación de BD-MOVIMIENTO finalizada',accion:'importar_movimiento',insertados:insertados,actualizados:actualizados,omitidos:omitidos,total:listarMovimiento_({}).total};
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
    bodega_preparacion:ped.bodega_preparacion,
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
  else if(name === 'VENDEDORES' || name === 'BD-VENDEDORES' || name === 'BD_VENDEDORES') sh = ss_().getSheetByName(SHEET_VENDEDORES);
  else if(name === 'BODEGAS' || name === 'BD-BODEGAS' || name === 'BD_BODEGAS') sh = bodegaSheet_(true);
  else if(name === 'CLIENTES' || name === 'BD-CLIENTES' || name === 'BD_CLIENTES') sh = clienteSheet_(true);
  else if(name === 'MAESTRA' || name === 'BD_MAESTRA' || name === 'BD-MAESTRA' || name === 'BDMAESTRA') sh = maestraSheet_(true);
  else if(name === 'ALERTAS_PEDIDOS') sh = ss_().getSheetByName(SHEET_ALERTAS);
  else if(name === 'MOVIMIENTO' || name === 'BD_MOVIMIENTO' || name === 'BD-MOVIMIENTO') sh = movimientoSheet_(true);
  else if(name === 'MOVIMIENTO_UBICACION' || name === 'BD_MOVIMIENTO_UBICACION' || name === 'BD-MOVIMIENTO-UBICACION' || name === 'MOVIMIENTOUBICACION') sh = gestionSheet_(SHEET_MOV_UBICACION, movimientoUbicacionHeaders_(), true);
  else if(name === 'MOVIMIENTO_PEDIDOS' || name === 'BD_MOVIMIENTO_PEDIDOS' || name === 'BD-MOVIMIENTO-PEDIDOS' || name === 'MOVIMIENTOPEDIDOS') sh = gestionSheet_(SHEET_MOV_PEDIDOS, movimientoPedidosHeaders_(), true);
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

function listarVendedores_(){
  const sh = getOrCreate_(SHEET_VENDEDORES, ['nombre','fecha_creacion']);
  ensureCols_(sh, ['nombre','fecha_creacion']);
  const v = values_(sh);
  const iN = find_(v.headers, ['nombre','vendedor','ejecutivo','responsable','usuario'], 0);
  const iF = find_(v.headers, ['fecha_creacion','fecha'], 1);
  const seen = {}, data = [];
  v.data.forEach(r=>{ const n=t_(r[iN]); if(!n || seen[n.toUpperCase()]) return; seen[n.toUpperCase()]=1; data.push({nombre:n,fecha_creacion:t_(r[iF])}); });
  data.sort((a,b)=>a.nombre.localeCompare(b.nombre));
  return {ok:true,sheet:SHEET_VENDEDORES,headers:v.headers,data:data,vendedores:data,total:data.length};
}

function agregarVendedor_(nombre){
  nombre = t_(nombre);
  if(!nombre) return {ok:false,msg:'Nombre de vendedor vacío'};
  const sh = getOrCreate_(SHEET_VENDEDORES, ['nombre','fecha_creacion']);
  ensureCols_(sh, ['nombre','fecha_creacion']);
  const v = values_(sh);
  const iN = find_(v.headers, ['nombre','vendedor','ejecutivo','responsable','usuario'],0);
  const existe = v.data.some(r=>t_(r[iN]).toUpperCase() === nombre.toUpperCase());
  if(!existe) sh.appendRow([nombre, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
  return {ok:true,msg:existe?'Vendedor ya existía':'Vendedor creado',nombre:nombre,data:listarVendedores_().data};
}

function listarMaestra_(){
  return listarMaestraBD_();
}
function buscarProducto_(codigo){
  codigo = code_(codigo);
  if(!codigo) return {ok:false,msg:'Código vacío'};

  // Búsqueda exacta directa en toda la hoja MAESTRA.
  // No usa listarMaestra_(), porque esa acción puede venir paginada y solo devolver
  // la primera página; eso hacía que algunos productos no aparecieran en Orden de Pedido.
  const sh = maestraSheet_(true);
  const v = values_(sh);
  const headers = v.headers && v.headers.length ? v.headers : maestraHeaders_();
  const ix = maestraIdx_(headers);
  let p = null;
  if(ix.codigo >= 0){
    for(let i=0;i<v.data.length;i++){
      const row = v.data[i];
      if(code_(getCell_(row, ix.codigo)) === codigo){
        p = {
          ok:true,
          rowNumber:i+2,
          id: ix.id >= 0 ? t_(getCell_(row, ix.id)) : '',
          codigo: code_(getCell_(row, ix.codigo)),
          descripcion: ix.descripcion >= 0 ? t_(getCell_(row, ix.descripcion)) : '',
          cantidad: ix.cantidad >= 0 ? t_(getCell_(row, ix.cantidad)) : '',
          categoria: ix.categoria >= 0 ? t_(getCell_(row, ix.categoria)) : '',
          unidad: ix.unidad >= 0 ? t_(getCell_(row, ix.unidad)) : '',
          status: ix.status >= 0 ? t_(getCell_(row, ix.status)) : '',
          ubicacion: ix.ubicacion >= 0 ? t_(getCell_(row, ix.ubicacion)) : '',
          fuente:'MAESTRA'
        };
        break;
      }
    }
  }

  const ubicacion = ubicacionPorCodigo_(codigo);
  if(p) return Object.assign(p,{ubicacion:ubicacion || p.ubicacion || ''});
  return ubicacion ? {ok:true,codigo:codigo,descripcion:'',cantidad:0,ubicacion:ubicacion,fuente:'MOVIMIENTO'} : {ok:false,msg:'Código no encontrado en MAESTRA/MOVIMIENTO',codigo:codigo};
}


function textoUbicacionCantidad_(ubicacion,cantidad){
  const u=t_(ubicacion);
  const c=t_(cantidad);
  if(!u) return '';
  return c && c !== '0' ? u + ' (' + c + ')' : u;
}
function unirUbicacionTexto_(actual,nuevo){
  nuevo=t_(nuevo);
  if(!nuevo) return t_(actual);
  const partes=t_(actual).split('|').map(function(x){return t_(x);}).filter(Boolean);
  const baseNuevo=nh_(nuevo.replace(/\s*\([^)]*\)\s*$/,''));
  const existe=partes.some(function(p){ return nh_(p.replace(/\s*\([^)]*\)\s*$/,'')) === baseNuevo; });
  if(!existe) partes.push(nuevo);
  return partes.join(' | ');
}
function ubicacionesMovimientoMap_(){
  const out={};
  [SHEET_MOV, SHEET_MOV_UBICACION].forEach(function(name){
    const sh=ss_().getSheetByName(name);
    if(!sh) return;
    const v=values_(sh);
    const h=v.headers||[];
    const iC=find_(h,['codigo','código','cod','sku','producto','codigo producto','código producto'], -1);
    const iU=find_(h,['ubicacion','ubicación','ubicaciones','bodega','ubic'], -1);
    const iQ=find_(h,['cantidad','stock','existencia','unidades','qty','cant'], -1);
    const iS=find_(h,['status','estado','vigencia'], -1);
    if(iC<0 || iU<0) return;
    v.data.forEach(function(r){
      const estado=iS>=0 ? nh_(getCell_(r,iS)) : '';
      if(['RETIRADO','ELIMINADO','ANULADO','INACTIVO'].indexOf(estado)>=0) return;
      const codigo=code_(getCell_(r,iC));
      const ubic=textoUbicacionCantidad_(getCell_(r,iU), iQ>=0?getCell_(r,iQ):'');
      if(codigo && ubic) out[codigo]=unirUbicacionTexto_(out[codigo]||'', ubic);
    });
  });
  return out;
}
function ubicacionPorCodigo_(codigo){
  codigo=code_(codigo);
  if(!codigo) return '';
  const map=ubicacionesMovimientoMap_();
  return map[codigo] || '';
}

/* =========================================================
   BD-MAESTRA - CARGA, CRUD E IMPORTACIÓN XLSX/CSV
   Acciones:
   - listar_maestra_bd / cargar_maestra
   - agregar_maestra / editar_maestra / eliminar_maestra
   - importar_maestra
========================================================= */
function maestraHeaders_(){
  return ['id','codigo','descripcion','cantidad','categoria','unidad','status','ubicacion','fecha_actualizacion'];
}
function maestraSheet_(crear){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_MAESTRA);
  if(!sh && crear) sh = getOrCreate_(SHEET_MAESTRA, maestraHeaders_());
  if(sh && crear){
    ensureCols_(sh, maestraHeaders_());
    try{
      const h = values_(sh).headers;
      const ix = maestraIdx_(h);
      if(ix.codigo >= 0) sh.getRange(1, ix.codigo + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    }catch(e){}
  }
  return sh;
}
function maestraIdx_(headers){
  return {
    id:find_(headers,['id'],0),
    codigo:find_(headers,['codigo','código','cod','sku','producto','item','codigo producto','código producto'],1),
    descripcion:find_(headers,['descripcion','descripción','detalle','nombre','nombre producto','descripcion producto','descripción producto'],2),
    cantidad:find_(headers,['cantidad','stock','unidades','existencia'],3),
    categoria:find_(headers,['categoria','categoría','familia','grupo','linea','línea'],4),
    unidad:find_(headers,['unidad','um','u/m','medida'],5),
    status:find_(headers,['status','estado','vigencia'],6),
    ubicacion:find_(headers,['ubicacion','ubicación','bodega','ubicacion sugerida','ubicación sugerida'],7),
    fecha_actualizacion:find_(headers,['fecha_actualizacion','fecha actualización','fecha','actualizado'],8)
  };
}
function listarMaestraBD_(p){
  p = p || {};
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_MAESTRA);
  if(!sh) sh = getOrCreate_(SHEET_MAESTRA, maestraHeaders_());

  const minHeaders = maestraHeaders_();
  let lastRow = Math.max(1, sh.getLastRow());
  let lastCol = Math.max(sh.getLastColumn(), minHeaders.length);
  let headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(t_);
  if(headers.every(h => !h)){
    sh.getRange(1, 1, 1, minHeaders.length).setValues([minHeaders]);
    headers = minHeaders.slice();
    lastCol = minHeaders.length;
    lastRow = Math.max(1, sh.getLastRow());
  }
  try{ ensureCols_(sh, minHeaders); }catch(e){}

  lastCol = Math.max(sh.getLastColumn(), minHeaders.length);
  headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(t_);
  lastRow = Math.max(1, sh.getLastRow());

  const totalDataRows = Math.max(0, lastRow - 1);
  const limit = Math.max(1, Math.min(Number(p.limit || p.pageSize || p.tamano || 300) || 300, 500));
  const offset = Math.max(0, Math.min(Number(p.offset || p.desde || p.start || 0) || 0, totalDataRows));
  const rowsToRead = Math.max(0, Math.min(limit, totalDataRows - offset));
  const full = rowsToRead > 0 ? sh.getRange(2 + offset, 1, rowsToRead, lastCol).getDisplayValues() : [];

  const data = [];
  const rowNumbers = [];
  for(let i = 0; i < full.length; i++){
    const row = full[i];
    if(row.some(c => t_(c))){
      data.push(row);
      rowNumbers.push(2 + offset + i);
    }
  }

  const ix = maestraIdx_(headers);
  const items = data.map((r, i) => ({
    rowNumber: rowNumbers[i],
    id: getCell_(r, ix.id),
    codigo: code_(getCell_(r, ix.codigo)),
    descripcion: t_(getCell_(r, ix.descripcion)),
    cantidad: t_(getCell_(r, ix.cantidad)),
    categoria: t_(getCell_(r, ix.categoria)),
    unidad: t_(getCell_(r, ix.unidad)),
    status: t_(getCell_(r, ix.status) || 'ACTIVO').toUpperCase(),
    ubicacion: t_(getCell_(r, ix.ubicacion)),
    fecha_actualizacion: t_(getCell_(r, ix.fecha_actualizacion))
  })).filter(x => x.codigo || x.descripcion);

  return {
    ok:true,
    sheet:sh.getName(),
    headers:headers,
    data:data,
    rows:data,
    values:data,
    items:items,
    maestra:items,
    productos:items,
    rowNumbers:rowNumbers,
    total:items.length,
    totalFilasMaestra:totalDataRows,
    offset:offset,
    limit:limit,
    paged:true,
    hasMore:(offset + rowsToRead) < totalDataRows,
    msg:'MAESTRA cargada por página para evitar bloqueo por exceso de registros',
    serverTime:new Date().toISOString()
  };
}
function valorMaestra_(b, names, fallback){
  b = b || {};
  const keys = Object.keys(b);
  for(const n of names){
    const k = keys.find(x => nh_(x) === nh_(n));
    if(k && t_(b[k])) return t_(b[k]);
  }
  return t_(fallback || '');
}
function maestraPayload_(b){
  return {
    id:t_(b.id || b.ID || ''),
    rowNumber:Number(b.rowNumber || b.fila || b.row || 0),
    originalCode:code_(b.originalCode || b.codigo_original || b.codigoAnterior || ''),
    codigo:code_(valorMaestra_(b,['codigo','código','cod','sku','producto','item','codigo producto','código producto'], b.codigo || b.cod || b.sku || '')),
    descripcion:t_(valorMaestra_(b,['descripcion','descripción','detalle','nombre','nombre producto','descripcion producto','descripción producto'], b.descripcion || b.detalle || b.nombre || '')),
    cantidad:n_(valorMaestra_(b,['cantidad','stock','unidades','existencia'], b.cantidad || b.stock || 0)),
    categoria:t_(valorMaestra_(b,['categoria','categoría','familia','grupo','linea','línea'], b.categoria || b.familia || '')),
    unidad:t_(valorMaestra_(b,['unidad','um','u/m','medida'], b.unidad || b.um || '')),
    status:t_(valorMaestra_(b,['status','estado','vigencia'], b.status || b.estado || 'ACTIVO') || 'ACTIVO').toUpperCase(),
    ubicacion:t_(valorMaestra_(b,['ubicacion','ubicación','bodega','ubicacion sugerida','ubicación sugerida'], b.ubicacion || b.bodega || ''))
  };
}
function encontrarFilaMaestra_(sh, headers, item){
  const v = values_(sh);
  const ix = maestraIdx_(headers || v.headers);
  const id = t_(item.id || '');
  const code = code_(item.originalCode || item.codigo || '');
  const rowNumber = Number(item.rowNumber || 0);
  if(rowNumber >= 2 && rowNumber <= sh.getLastRow()) return rowNumber;
  if(id && ix.id >= 0){
    for(let i=0;i<v.data.length;i++) if(t_(getCell_(v.data[i], ix.id)) === id) return i+2;
  }
  if(code && ix.codigo >= 0){
    for(let i=0;i<v.data.length;i++) if(code_(getCell_(v.data[i], ix.codigo)) === code) return i+2;
  }
  return 0;
}
function escribirMaestraEnFila_(sh, headers, rowNumber, item, mantenerId){
  const ix = maestraIdx_(headers);
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const valuesMap = {
    id: mantenerId ? '' : (item.id || id_('MA')),
    codigo:item.codigo,
    descripcion:item.descripcion,
    cantidad:item.cantidad,
    categoria:item.categoria,
    unidad:item.unidad,
    status:item.status || 'ACTIVO',
    ubicacion:item.ubicacion,
    fecha_actualizacion:ts
  };
  Object.keys(valuesMap).forEach(field => {
    const col = ix[field];
    if(col < 0) return;
    if(field === 'id' && mantenerId) return;
    sh.getRange(rowNumber, col + 1).setValue(valuesMap[field]);
  });
  return ts;
}
function agregarMaestra_(b){
  const sh = maestraSheet_(true);
  const v = values_(sh), headers = v.headers.length ? v.headers : maestraHeaders_();
  const item = maestraPayload_(b);
  if(!item.codigo) return {ok:false,msg:'Código vacío'};
  if(!item.descripcion) return {ok:false,msg:'Descripción vacía'};
  const existe = encontrarFilaMaestra_(sh, headers, {codigo:item.codigo});
  if(existe) return editarMaestra_(Object.assign({}, b, {rowNumber:existe, originalCode:item.codigo}));
  const obj = {id:id_('MA'), codigo:item.codigo, descripcion:item.descripcion, cantidad:item.cantidad, categoria:item.categoria, unidad:item.unidad, status:item.status || 'ACTIVO', ubicacion:item.ubicacion, fecha_actualizacion:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')};
  appendObjects_(sh, headers, [obj]);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Producto agregado a BD-MAESTRA',accion:'agregar_maestra',codigo:item.codigo,total:listarMaestraBD_().total};
}
function editarMaestra_(b){
  const sh = maestraSheet_(true);
  const v = values_(sh), headers = v.headers.length ? v.headers : maestraHeaders_();
  const item = maestraPayload_(b);
  if(!item.codigo) return {ok:false,msg:'Código vacío'};
  if(!item.descripcion) return {ok:false,msg:'Descripción vacía'};
  const rowNumber = encontrarFilaMaestra_(sh, headers, item);
  if(!rowNumber) return {ok:false,msg:'Producto no encontrado en BD-MAESTRA'};
  escribirMaestraEnFila_(sh, headers, rowNumber, item, true);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Producto editado correctamente',accion:'editar_maestra',codigo:item.codigo,rowNumber:rowNumber};
}
function eliminarMaestra_(b){
  const sh = maestraSheet_(true);
  const v = values_(sh), headers = v.headers.length ? v.headers : maestraHeaders_();
  const item = maestraPayload_(b);
  const rowNumber = encontrarFilaMaestra_(sh, headers, item);
  if(!rowNumber) return {ok:false,msg:'Producto no encontrado para eliminar'};
  sh.deleteRow(rowNumber);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Producto eliminado de BD-MAESTRA',accion:'eliminar_maestra',codigo:item.codigo,rowNumber:rowNumber};
}
function importarMaestra_(b){
  const sh = maestraSheet_(true);
  const v = values_(sh), headers = v.headers.length ? v.headers : maestraHeaders_();
  let items = [];
  if(Array.isArray(b.rows)) items = b.rows;
  else if(Array.isArray(b.data)) items = b.data;
  else items = parseItems_(b.rows || b.data || b.items || b.productos || b.detalle);
  const modo = t_(b.modo || 'upsert').toLowerCase();
  let insertados = 0, actualizados = 0, omitidos = 0;
  items.forEach(raw => {
    const item = maestraPayload_(raw || {});
    if(!item.codigo || !item.descripcion){ omitidos++; return; }
    const existe = encontrarFilaMaestra_(sh, headers, {codigo:item.codigo});
    if(existe && modo !== 'append'){
      escribirMaestraEnFila_(sh, headers, existe, item, true);
      actualizados++;
    }else{
      const obj = {id:id_('MA'), codigo:item.codigo, descripcion:item.descripcion, cantidad:item.cantidad, categoria:item.categoria, unidad:item.unidad, status:item.status || 'ACTIVO', ubicacion:item.ubicacion, fecha_actualizacion:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')};
      appendObjects_(sh, headers, [obj]);
      insertados++;
    }
  });
  SpreadsheetApp.flush();
  return {ok:true,msg:'Importación de BD-MAESTRA finalizada',accion:'importar_maestra',insertados:insertados,actualizados:actualizados,omitidos:omitidos,total:listarMaestraBD_().total};
}


/* =========================================================
   MOVIMIENTO_UBICACION / MOVIMIENTO_PEDIDOS - CRUD COMPLETO
   Acciones soportadas:
   - listar_movimiento_ubicacion / agregar / editar / eliminar / importar
   - listar_movimiento_pedidos / agregar / editar / eliminar / importar
========================================================= */
function movimientoUbicacionHeaders_(){
  return ['id','fecha_registro','codigo','descripcion','ubicacion','tipo_movimiento','cantidad','pedido','responsable','status','observaciones','origen'];
}
function movimientoPedidosHeaders_(){
  return ['id','fecha','pedido','status','cliente','vendedor','pikeador','cantidad_items','origen','observaciones'];
}
function gestionSheet_(name, headers, crear){
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if(!sh && crear) sh = getOrCreate_(name, headers);
  if(sh && crear){
    ensureCols_(sh, headers);
    try{
      const h = values_(sh).headers;
      const ixCodigo = find_(h, ['codigo','código','cod','sku'], -1);
      const ixPedido = find_(h, ['pedido','nro pedido','numero pedido','número pedido'], -1);
      if(ixCodigo >= 0) sh.getRange(1, ixCodigo + 1, sh.getMaxRows(), 1).setNumberFormat('@');
      if(ixPedido >= 0) sh.getRange(1, ixPedido + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    }catch(e){}
  }
  return sh;
}
function listarGestionBD_(sheetName, headers){
  const sh = gestionSheet_(sheetName, headers, true);
  const lastRow = Math.max(1, sh.getLastRow());
  const lastCol = Math.max(sh.getLastColumn(), headers.length);
  const raw = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  let h = (raw[0] || []).map(t_);
  if(h.every(x => !x)){ h = headers; sh.getRange(1, 1, 1, h.length).setValues([h]); }
  ensureCols_(sh, headers);
  h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getDisplayValues()[0].map(t_);
  const full = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), Math.max(sh.getLastColumn(), headers.length)).getDisplayValues();
  const data = [];
  const items = [];
  const rowNumbers = [];
  for(let i=1; i<full.length; i++){
    const row = full[i];
    if(!row.some(c => t_(c))) continue;
    data.push(row);
    rowNumbers.push(i + 1);
    const obj = {rowNumber:i + 1};
    h.forEach((col, idx) => { if(col) obj[col] = t_(row[idx]); });
    items.push(obj);
  }
  return {ok:true,sheet:sh.getName(),headers:h,data:data,rows:data,values:data,items:items,rowNumbers:rowNumbers,total:data.length,serverTime:new Date().toISOString()};
}
function valorGestion_(b, header){
  b = b || {};
  const keys = Object.keys(b);
  const aliases = [header];
  const n = nh_(header);
  if(n === 'CODIGO') aliases.push('código','cod','sku');
  if(n === 'DESCRIPCION') aliases.push('descripción','detalle','nombre');
  if(n === 'UBICACION') aliases.push('ubicación','bodega');
  if(n === 'FECHAREGISTRO') aliases.push('fecha registro','fecha','registro');
  if(n === 'TIPOMOVIMIENTO') aliases.push('tipo movimiento','tipo','movimiento');
  if(n === 'CANTIDADITEMS') aliases.push('cantidad items','items','total_items');
  if(n === 'STATUS') aliases.push('estado','status');
  const k = keys.find(key => aliases.some(a => nh_(a) === nh_(key)));
  return k ? t_(b[k]) : '';
}
function filaGestion_(sh, headers, b){
  const rowNumber = Number(b.rowNumber || b.fila || b.row || 0);
  if(rowNumber >= 2 && rowNumber <= sh.getLastRow()) return rowNumber;
  const id = t_(b.id || b.ID || '');
  if(!id) return 0;
  const v = values_(sh);
  const ix = find_(headers, ['id'], 0);
  if(ix < 0) return 0;
  for(let i=0; i<v.data.length; i++) if(t_(getCell_(v.data[i], ix)) === id) return i + 2;
  return 0;
}
function objetoGestion_(b, headers, prefix){
  const obj = {};
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  headers.forEach(h => {
    const n = nh_(h);
    let value = valorGestion_(b, h);
    if(n === 'ID') value = t_(b.id || b.ID || '') || id_(prefix);
    if((n === 'FECHAREGISTRO' || n === 'FECHA') && !value) value = now;
    if(n === 'CODIGO') value = code_(value);
    if(n === 'PEDIDO') value = pedido_(value);
    if(n === 'CANTIDAD' || n === 'CANTIDADITEMS') value = value === '' ? '' : n_(value);
    if(n === 'STATUS' && !value) value = 'VIGENTE';
    if(n === 'ORIGEN' && !value) value = 'WEB';
    obj[h] = value;
  });
  return obj;
}
function agregarGestionBD_(sheetName, headers, b, prefix){
  const sh = gestionSheet_(sheetName, headers, true);
  const h = values_(sh).headers.length ? values_(sh).headers : headers;
  const obj = objetoGestion_(b, h, prefix);
  appendObjects_(sh, h, [obj]);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Registro agregado en '+sheetName,accion:'agregar_'+sheetName.toLowerCase(),id:obj.id || obj.ID || '',total:listarGestionBD_(sheetName, headers).total};
}
function editarGestionBD_(sheetName, headers, b, prefix){
  const sh = gestionSheet_(sheetName, headers, true);
  const h = values_(sh).headers.length ? values_(sh).headers : headers;
  const row = filaGestion_(sh, h, b);
  if(!row) return {ok:false,msg:'Registro no encontrado en '+sheetName};
  const obj = objetoGestion_(b, h, prefix);
  h.forEach((col, idx) => {
    if(nh_(col) === 'ID' && !t_(valorGestion_(b, col))) return;
    sh.getRange(row, idx + 1).setValue(obj[col]);
  });
  SpreadsheetApp.flush();
  return {ok:true,msg:'Registro editado en '+sheetName,rowNumber:row};
}
function eliminarGestionBD_(sheetName, headers, b){
  const sh = gestionSheet_(sheetName, headers, true);
  const h = values_(sh).headers.length ? values_(sh).headers : headers;
  const row = filaGestion_(sh, h, b);
  if(!row) return {ok:false,msg:'Registro no encontrado para eliminar en '+sheetName};
  sh.deleteRow(row);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Registro eliminado en '+sheetName,rowNumber:row};
}
function importarGestionBD_(sheetName, headers, b, prefix){
  let items = [];
  if(Array.isArray(b.rows)) items = b.rows;
  else if(Array.isArray(b.data)) items = b.data;
  else items = parseItems_(b.rows || b.data || b.items || b.detalle);
  const modo = t_(b.modo || 'append').toLowerCase();
  let insertados = 0, actualizados = 0, omitidos = 0;
  items.forEach(raw => {
    raw = raw || {};
    if(!Object.keys(raw).length){ omitidos++; return; }
    if(modo !== 'append' && (raw.id || raw.ID || raw.rowNumber || raw.fila || raw.row)){
      const r = editarGestionBD_(sheetName, headers, raw, prefix);
      if(r && r.ok) actualizados++; else { agregarGestionBD_(sheetName, headers, raw, prefix); insertados++; }
    }else{
      agregarGestionBD_(sheetName, headers, raw, prefix);
      insertados++;
    }
  });
  SpreadsheetApp.flush();
  return {ok:true,msg:'Importación finalizada en '+sheetName,insertados:insertados,actualizados:actualizados,omitidos:omitidos,total:listarGestionBD_(sheetName, headers).total};
}
function listarMovimientoUbicacionBD_(){ return listarGestionBD_(SHEET_MOV_UBICACION, movimientoUbicacionHeaders_()); }
function agregarMovimientoUbicacionBD_(b){ return agregarGestionBD_(SHEET_MOV_UBICACION, movimientoUbicacionHeaders_(), b, 'MVU'); }
function editarMovimientoUbicacionBD_(b){ return editarGestionBD_(SHEET_MOV_UBICACION, movimientoUbicacionHeaders_(), b, 'MVU'); }
function eliminarMovimientoUbicacionBD_(b){ return eliminarGestionBD_(SHEET_MOV_UBICACION, movimientoUbicacionHeaders_(), b); }
function importarMovimientoUbicacionBD_(b){ return importarGestionBD_(SHEET_MOV_UBICACION, movimientoUbicacionHeaders_(), b, 'MVU'); }
function listarMovimientoPedidosBD_(){ return listarGestionBD_(SHEET_MOV_PEDIDOS, movimientoPedidosHeaders_()); }
function agregarMovimientoPedidosBD_(b){ return agregarGestionBD_(SHEET_MOV_PEDIDOS, movimientoPedidosHeaders_(), b, 'MVP'); }
function editarMovimientoPedidosBD_(b){ return editarGestionBD_(SHEET_MOV_PEDIDOS, movimientoPedidosHeaders_(), b, 'MVP'); }
function eliminarMovimientoPedidosBD_(b){ return eliminarGestionBD_(SHEET_MOV_PEDIDOS, movimientoPedidosHeaders_(), b); }
function importarMovimientoPedidosBD_(b){ return importarGestionBD_(SHEET_MOV_PEDIDOS, movimientoPedidosHeaders_(), b, 'MVP'); }




/* =========================================================
   CLIENTES - REGISTRO MAESTRO PARA PEDIDOS
   Hoja: CLIENTES
========================================================= */
function clienteHeaders_(){ return ['id','cliente','rut','vendedor','clase_cliente','direccion','giro','medio_pago','telefono','correo_electronico','responsable_empresa','fecha_creacion','status']; }
function clienteSheet_(crear){
  const sh = getOrCreate_(SHEET_CLIENTES, clienteHeaders_());
  if(crear) ensureCols_(sh, clienteHeaders_());
  return sh;
}
function clienteIdx_(h){
  return {
    id:find_(h,['id'],0),
    cliente:find_(h,['cliente','nombre','razon_social','razón social'],1),
    rut:find_(h,['rut','r.u.t'],2),
    vendedor:find_(h,['vendedor','ejecutivo','asesor','responsable_venta','responsable venta'],3),
    clase_cliente:find_(h,['clase_cliente','clase cliente','tipo_cliente','tipo cliente','categoria_cliente','categoría cliente','categoria'],4),
    direccion:find_(h,['direccion','dirección'],5),
    giro:find_(h,['giro'],6),
    medio_pago:find_(h,['medio_pago','medio pago','forma_pago','forma pago'],7),
    telefono:find_(h,['telefono','teléfono','fono'],8),
    correo_electronico:find_(h,['correo_electronico','correo electrónico','correo','email','mail'],9),
    responsable_empresa:find_(h,['responsable_empresa','responsable empresa','responsable'],10),
    fecha_creacion:find_(h,['fecha_creacion','fecha creación','fecha'],11),
    status:find_(h,['status','estado'],12)
  };
}
function listarClientes_(){
  const sh = clienteSheet_(true);
  const v = values_(sh), h = v.headers.length ? v.headers : clienteHeaders_(), ix = clienteIdx_(h);
  const items = v.data.map((r,i)=>({
    rowNumber:i+2,
    id:getCell_(r,ix.id),
    cliente:getCell_(r,ix.cliente),
    rut:getCell_(r,ix.rut),
    vendedor:getCell_(r,ix.vendedor),
    clase_cliente:(getCell_(r,ix.clase_cliente)||'CLIENTE NORMAL').toUpperCase(),
    direccion:getCell_(r,ix.direccion),
    giro:getCell_(r,ix.giro),
    medio_pago:getCell_(r,ix.medio_pago),
    telefono:getCell_(r,ix.telefono),
    correo_electronico:getCell_(r,ix.correo_electronico),
    responsable_empresa:getCell_(r,ix.responsable_empresa),
    fecha_creacion:getCell_(r,ix.fecha_creacion),
    status:(getCell_(r,ix.status)||'ACTIVO').toUpperCase()
  })).filter(x=>x.cliente);
  return {ok:true,sheet:sh.getName(),headers:h,data:v.data,rows:v.data,values:v.data,items:items,clientes:items,total:items.length,serverTime:new Date().toISOString()};
}
function clientePayload_(b){
  const st = nh_(b.status || b.estado || 'ACTIVO');
  const status = st === 'INACTIVO' ? 'INACTIVO' : st === 'BLOQUEADO' ? 'BLOQUEADO' : 'ACTIVO';
  return {
    id:t_(b.id || b.ID || ''),
    rowNumber:Number(b.rowNumber || b.fila || b.row || 0),
    cliente:t_(b.cliente || b.nombre || b.razon_social || b['razón social'] || ''),
    rut:t_(b.rut || b.RUT || ''),
    vendedor:t_(b.vendedor || b.ejecutivo || b.asesor || ''),
    clase_cliente:(function(){ var c=nh_(b.clase_cliente || b.claseCliente || b.tipo_cliente || b.tipoCliente || b.categoria_cliente || b.categoria || 'CLIENTE NORMAL'); return c === 'FRECUENTE' ? 'FRECUENTE' : c === 'PREMIUM' ? 'PREMIUM' : c === 'SUPER PREMIUM' || c === 'SUPERPREMIUM' ? 'SUPER PREMIUM' : 'CLIENTE NORMAL'; })(),
    direccion:t_(b.direccion || b['dirección'] || ''),
    giro:t_(b.giro || ''),
    medio_pago:t_(b.medio_pago || b.medioPago || b.forma_pago || ''),
    telefono:t_(b.telefono || b.fono || ''),
    correo_electronico:t_(b.correo_electronico || b.correo || b.email || b.mail || ''),
    responsable_empresa:t_(b.responsable_empresa || b.responsable || ''),
    status:status
  };
}
function filaCliente_(sh,h,item){
  const row = Number(item.rowNumber || 0);
  if(row >= 2 && row <= sh.getLastRow()) return row;
  const v = values_(sh), ix = clienteIdx_(h);
  if(item.id && ix.id >= 0){ for(let i=0;i<v.data.length;i++) if(t_(getCell_(v.data[i],ix.id)) === item.id) return i+2; }
  const rut = nh_(item.rut);
  if(rut && ix.rut >= 0){ for(let i=0;i<v.data.length;i++) if(nh_(getCell_(v.data[i],ix.rut)) === rut) return i+2; }
  const cliente = nh_(item.cliente);
  if(cliente && ix.cliente >= 0){ for(let i=0;i<v.data.length;i++) if(nh_(getCell_(v.data[i],ix.cliente)) === cliente) return i+2; }
  return 0;
}
function agregarCliente_(b){
  const sh = clienteSheet_(true);
  const h = values_(sh).headers.length ? values_(sh).headers : clienteHeaders_();
  const item = clientePayload_(b);
  if(!item.cliente) return {ok:false,msg:'Nombre de cliente vacío'};
  const existe = filaCliente_(sh,h,item);
  if(existe) return editarCliente_(Object.assign({}, b, {rowNumber:existe}));
  const obj = Object.assign({}, item, {id:id_('CLI'),fecha_creacion:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')});
  appendObjects_(sh,h,[obj]);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Cliente creado correctamente',cliente:item.cliente,total:listarClientes_().total};
}
function editarCliente_(b){
  const sh = clienteSheet_(true);
  const h = values_(sh).headers.length ? values_(sh).headers : clienteHeaders_();
  const ix = clienteIdx_(h), item = clientePayload_(b);
  const row = filaCliente_(sh,h,item);
  if(!row) return {ok:false,msg:'Cliente no encontrado'};
  const set = (field,val)=>{ const col=ix[field]; if(col>=0) sh.getRange(row,col+1).setValue(val); };
  set('cliente',item.cliente); set('rut',item.rut); set('vendedor',item.vendedor); set('clase_cliente',item.clase_cliente); set('direccion',item.direccion); set('giro',item.giro); set('medio_pago',item.medio_pago); set('telefono',item.telefono); set('correo_electronico',item.correo_electronico); set('responsable_empresa',item.responsable_empresa); set('status',item.status);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Cliente actualizado',cliente:item.cliente,rowNumber:row};
}
function eliminarCliente_(b){
  const sh = clienteSheet_(true);
  const h = values_(sh).headers.length ? values_(sh).headers : clienteHeaders_();
  const row = filaCliente_(sh,h,clientePayload_(b));
  if(!row) return {ok:false,msg:'Cliente no encontrado para eliminar'};
  sh.deleteRow(row);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Cliente eliminado',rowNumber:row,total:listarClientes_().total};
}

/* =========================================================
   BODEGAS - CONFIGURACIÓN DE BODEGAS DE PREPARACIÓN
   Hoja: BODEGAS
   Campos: nombre_bodega, direccion, fecha_creacion, status, radio_metros, latitud, longitud
========================================================= */
function bodegaHeaders_(){ return ['id','nombre_bodega','direccion','fecha_creacion','status','radio_metros','latitud','longitud']; }
function bodegaSheet_(crear){
  const sh = getOrCreate_(SHEET_BODEGAS, bodegaHeaders_());
  if(crear) ensureCols_(sh, bodegaHeaders_());
  return sh;
}
function bodegaIdx_(h){
  return {
    id:find_(h,['id'],0),
    nombre_bodega:find_(h,['nombre_bodega','nombre bodega','bodega','nombre'],1),
    direccion:find_(h,['direccion','dirección'],2),
    fecha_creacion:find_(h,['fecha_creacion','fecha creación','fecha'],3),
    status:find_(h,['status','estado'],4),
    radio_metros:find_(h,['radio_metros','radio metros','radio'],5),
    latitud:find_(h,['latitud','lat'],6),
    longitud:find_(h,['longitud','lng','lon'],7)
  };
}
function listarBodegas_(){
  const sh = bodegaSheet_(true);
  const v = values_(sh), h = v.headers.length ? v.headers : bodegaHeaders_(), ix = bodegaIdx_(h);
  const items = v.data.map((r,i)=>({
    rowNumber:i+2,
    id:getCell_(r,ix.id),
    nombre:getCell_(r,ix.nombre_bodega),
    nombre_bodega:getCell_(r,ix.nombre_bodega),
    direccion:getCell_(r,ix.direccion),
    fecha_creacion:getCell_(r,ix.fecha_creacion),
    status:(getCell_(r,ix.status)||'ACTIVA').toUpperCase(),
    radio_metros:getCell_(r,ix.radio_metros)||'200',
    latitud:getCell_(r,ix.latitud),
    longitud:getCell_(r,ix.longitud)
  })).filter(x=>x.nombre_bodega);
  return {ok:true,sheet:sh.getName(),headers:h,data:v.data,rows:v.data,values:v.data,items:items,bodegas:items,total:items.length,serverTime:new Date().toISOString()};
}
function bodegaPayload_(b){
  const status = nh_(b.status || b.estado || 'ACTIVA');
  const st = status === 'INACTIVA' ? 'INACTIVA' : status === 'BLOQUEADA' ? 'BLOQUEADA' : 'ACTIVA';
  return {
    id:t_(b.id || b.ID || ''),
    rowNumber:Number(b.rowNumber || b.fila || b.row || 0),
    nombre_bodega:t_(b.nombre_bodega || b.nombreBodega || b.bodega || b.nombre || ''),
    direccion:t_(b.direccion || b['dirección'] || ''),
    status:st,
    radio_metros:n_(b.radio_metros || b.radio || 200) || 200,
    latitud:t_(b.latitud || b.lat || ''),
    longitud:t_(b.longitud || b.lng || b.lon || '')
  };
}
function filaBodega_(sh, h, item){
  const row = Number(item.rowNumber || 0);
  if(row >= 2 && row <= sh.getLastRow()) return row;
  const v = values_(sh), ix = bodegaIdx_(h);
  if(item.id && ix.id >= 0){ for(let i=0;i<v.data.length;i++) if(t_(getCell_(v.data[i],ix.id)) === item.id) return i+2; }
  const nombre = nh_(item.nombre_bodega);
  if(nombre && ix.nombre_bodega >= 0){ for(let i=0;i<v.data.length;i++) if(nh_(getCell_(v.data[i],ix.nombre_bodega)) === nombre) return i+2; }
  return 0;
}
function agregarBodega_(b){
  const sh = bodegaSheet_(true);
  const h = values_(sh).headers.length ? values_(sh).headers : bodegaHeaders_();
  const item = bodegaPayload_(b);
  if(!item.nombre_bodega) return {ok:false,msg:'Nombre de bodega vacío'};
  if(!item.direccion) return {ok:false,msg:'Dirección de bodega vacía'};
  const existe = filaBodega_(sh,h,item);
  if(existe) return editarBodega_(Object.assign({}, b, {rowNumber:existe}));
  const obj = {id:id_('BDG'),nombre_bodega:item.nombre_bodega,direccion:item.direccion,fecha_creacion:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),status:item.status,radio_metros:item.radio_metros,latitud:item.latitud,longitud:item.longitud};
  appendObjects_(sh,h,[obj]);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Bodega creada correctamente',bodega:item.nombre_bodega,total:listarBodegas_().total};
}
function editarBodega_(b){
  const sh = bodegaSheet_(true);
  const h = values_(sh).headers.length ? values_(sh).headers : bodegaHeaders_();
  const ix = bodegaIdx_(h), item = bodegaPayload_(b);
  const row = filaBodega_(sh,h,item);
  if(!row) return {ok:false,msg:'Bodega no encontrada'};
  const set = (field,val)=>{ const col=ix[field]; if(col>=0) sh.getRange(row,col+1).setValue(val); };
  set('nombre_bodega',item.nombre_bodega); set('direccion',item.direccion); set('status',item.status); set('radio_metros',item.radio_metros); set('latitud',item.latitud); set('longitud',item.longitud);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Bodega actualizada',bodega:item.nombre_bodega,rowNumber:row};
}
function eliminarBodega_(b){
  const sh = bodegaSheet_(true);
  const h = values_(sh).headers.length ? values_(sh).headers : bodegaHeaders_();
  const row = filaBodega_(sh,h,bodegaPayload_(b));
  if(!row) return {ok:false,msg:'Bodega no encontrada para eliminar'};
  sh.deleteRow(row);
  SpreadsheetApp.flush();
  return {ok:true,msg:'Bodega eliminada',rowNumber:row};
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
  const canonical = ['id','fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','bodega preparación','bodega preparacion','bodega pedido','bodega de preparacion','bodega de preparación','nombre_bodega','nombre bodega','bodega','codigo','código','cod','sku','descripcion','descripción','detalle','ubicacion','ubicación','cantidad','unidades','status','estado','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora inicio','inicio_preparacion','fecha_asignacion','hora_termino','hora termino','hora término','termino_preparacion','fecha_termino','tiempo_preparacion_min','minutos_preparacion','import_raw_json','import_headers_json'];
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
  const usarAuto = bool_(b && (b.autogenerar_pedido || b.auto_pedido || b.autoPedido || b.generar_pedido || b.generarPedido));
  const lock = usarAuto ? LockService.getScriptLock() : null;
  if(lock) lock.waitLock(25000);
  try{
    return guardarPedidoImpl_(b || {}, usarAuto);
  }finally{
    if(lock){ try{ lock.releaseLock(); }catch(e){} }
  }
}

function guardarPedidoImpl_(b, usarAuto){
  const sh = pedidosSheet_(true);
  const baseCols = ['id','fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min','import_raw_json','import_headers_json'];
  const items = parseItems_(b.items || b.productos || b.rows || b.detalle);
  // Preserva información importada tal como viene desde el XLSX/CSV: guarda JSON completo
  // y agrega columnas extra si el archivo trae campos que no existen en la BD.
  const rawCols = collectRawHeaders_(items.concat([b])).filter(h => !isCanonicalPedidoHeader_(h));
  ensureCols_(sh, baseCols.concat(rawCols));
  const before = values_(sh);
  const headers = before.headers;
  const ix = idx_(headers);
  const ubicacionesPorCodigo = ubicacionesMovimientoMap_();
  if(ix.pedido >= 0){
    try{ sh.getRange(1, ix.pedido + 1, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat('@'); }catch(e){}
  }
  let pedido = valorPedido_(b, b.order || b.nro_pedido);
  if(usarAuto){
    const next = siguienteNumeroPedidoDesdeValues_(before, ix);
    pedido = next.numero;
  }
  if(!pedido && !items.length) return {ok:false,msg:'Pedido vacío'};

  const base = {
    fecha:fechaVisible_(b.fecha)||Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HH:mm:ss'),
    pedido:pedido,
    cliente:t_(b.cliente),
    vendedor:t_(b.vendedor),
    pikeador:limpiarPikeador_(b.pikeador, b.cliente),
    bodega_preparacion:t_(b.bodega_preparacion || b.bodegaPreparacion || b.bodega || ''),
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
      fecha:fechaVisible_(it.fecha || base.fecha) || base.fecha,
      pedido:usarAuto ? pedido : valorPedido_(Object.assign({}, it.__raw || {}, it), pedido),
      cliente:t_(it.cliente||base.cliente),
      vendedor:t_(it.vendedor||base.vendedor),
      pikeador:limpiarPikeador_(it.pikeador || base.pikeador, it.cliente || base.cliente),
      bodega_preparacion:t_(it.bodega_preparacion || it.bodegaPreparacion || it.bodega || base.bodega_preparacion),
      codigo:code_(it.codigo||it.CODIGO||it.producto||it.sku),
      descripcion:t_(it.descripcion||it.DESCRIPCION||it.detalle||it.nombre),
      ubicacion:t_(it.ubicacion||it.UBICACION) || ubicacionesPorCodigo[code_(it.codigo||it.CODIGO||it.producto||it.sku)] || '',
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

    // Respeta varias líneas del mismo pedido/código cuando tienen descripción o ubicación distinta.
    // El fallback por pedido+código solo se usa para archivos antiguos sin ubicación ni descripción.
    const usarFallbackPedidoCodigo = pc && !t_(row.ubicacion) && !t_(row.descripcion);
    const found = porClave[key] || (usarFallbackPedidoCodigo ? porPedidoCodigo[pc] : null);
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
  if(base.vendedor) agregarVendedor_(base.vendedor);
  nuevos.forEach(x=>{ if(x.pikeador) agregarPikeador_(x.pikeador); if(x.vendedor) agregarVendedor_(x.vendedor); });
  rows.forEach(x=>{ if(x.pikeador) agregarPikeador_(x.pikeador); if(x.vendedor) agregarVendedor_(x.vendedor); });

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


function verificarImportacionPedidos_(p){
  const sh = pedidosSheet_(false);
  if(!sh) return {ok:false,msg:'No existe hoja PEDIDOS',recibidos:0,encontrados:0,faltantes:0};
  const items = parseItems_(p.items || p.keys || p.rows || p.detalle);
  if(!items.length) return {ok:true,msg:'Sin filas para verificar',recibidos:0,encontrados:0,faltantes:0};
  const v = values_(sh), ix = idx_(v.headers);
  const exactos = {};
  const pedidoCodigo = {};
  v.data.forEach((r,i)=>{
    const obj = rowToObj_(r, ix);
    const key = uniquePedidoProductoKey_(obj.pedido, obj.codigo, obj.descripcion, obj.ubicacion);
    if(key) exactos[key] = true;
    const pc = pedidoCodigoKey_(obj.pedido, obj.codigo);
    if(pc) pedidoCodigo[pc] = true;
  });
  let encontrados = 0;
  const faltantes = [];
  items.forEach(it=>{
    const codigo = code_(it.codigo || it.CODIGO || it.cod || it.sku || it.producto);
    const obj = {
      pedido:valorPedido_(it, it.pedido || it.numero || it['número'] || it.nro || ''),
      codigo:codigo,
      descripcion:t_(it.descripcion || it.DESCRIPCION || it.detalle || it.nombre || ''),
      ubicacion:t_(it.ubicacion || it.UBICACION || '')
    };
    const key = uniquePedidoProductoKey_(obj.pedido, obj.codigo, obj.descripcion, obj.ubicacion);
    const pc = pedidoCodigoKey_(obj.pedido, obj.codigo);
    if(exactos[key] || (!obj.ubicacion && !obj.descripcion && pedidoCodigo[pc])) encontrados++;
    else faltantes.push(obj);
  });
  return {
    ok:true,
    recibidos:items.length,
    encontrados:encontrados,
    faltantes:faltantes.length,
    detalle_faltantes:faltantes.slice(0,100),
    serverTime:new Date().toISOString()
  };
}


function editarPedidoCompleto_(b){
  const sh = pedidosSheet_(true);
  const cols = ['id','fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min','import_raw_json','import_headers_json'];
  ensureCols_(sh, cols);
  const v = values_(sh), headers = v.headers, ix = idx_(headers);
  const pedidoOriginal = t_(b.pedido_original || b.pedidoOriginal || b.original || b.pedido_anterior || b.pedidoAnterior) || valorPedido_(b, b.pedido);
  const pedidoNuevo = valorPedido_(b, b.pedido || pedidoOriginal);
  if(!pedidoOriginal) return {ok:false,msg:'Pedido original vacío'};
  if(!pedidoNuevo) return {ok:false,msg:'Pedido nuevo vacío'};

  const ubicacionesPorCodigo = ubicacionesMovimientoMap_();
  const items = parseItems_(b.items || b.productos || b.detalle || b.rows || '[]').map(it => {
    const codigo = code_(it.codigo || it.CODIGO || it.cod || it.sku || it.producto);
    return {
      codigo:codigo,
      descripcion:t_(it.descripcion || it.DESCRIPCION || it.detalle || it.nombre || it.producto),
      ubicacion:t_(it.ubicacion || it.UBICACION || it.bodega) || ubicacionesPorCodigo[codigo] || '',
      cantidad:n_(it.cantidad || it.CANTIDAD || it.unidades || it.qty || 1) || 1
    };
  }).filter(x => x.codigo || x.descripcion || x.ubicacion || x.cantidad > 0);
  if(!items.length) return {ok:false,msg:'El pedido debe tener al menos un producto'};

  const keyOriginal = pedido_(pedidoOriginal);
  const filas = [];
  v.data.forEach((r,i)=>{ if(pedido_(getCell_(r, ix.pedido)) === keyOriginal) filas.push({row:i+2, data:r}); });

  if(!filas.length){
    return guardarPedido_(Object.assign({}, b, {pedido:pedidoNuevo, items:JSON.stringify(items)}));
  }

  const primera = filas[0].data;
  const estadoAnterior = estadoSeguro_(getCell_(primera, ix.status), 'PENDIENTE');
  const fecha = fechaVisible_(b.fecha) || fechaVisible_(getCell_(primera, ix.fecha)) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const cliente = t_(b.cliente) || getCell_(primera, ix.cliente);
  const vendedor = t_(b.vendedor) || getCell_(primera, ix.vendedor);
  const pikeador = limpiarPikeador_(b.pikeador || getCell_(primera, ix.pikeador), cliente);
  const status = estadoSeguro_(b.status || b.estado || getCell_(primera, ix.status), 'PENDIENTE');

  function setField(row, field, value){
    const col = ix[field];
    if(col >= 0) sh.getRange(row, col + 1).setValue(value);
  }

  let actualizados = 0, insertados = 0, eliminados = 0;
  items.forEach((it, i)=>{
    if(i < filas.length){
      const row = filas[i].row;
      setField(row,'fecha',fecha);
      setField(row,'pedido',pedidoNuevo);
      setField(row,'cliente',cliente);
      setField(row,'vendedor',vendedor);
      setField(row,'pikeador',pikeador);
      setField(row,'bodega_preparacion',t_(b.bodega_preparacion || b.bodegaPreparacion || b.bodega || getCell_(primera, ix.bodega_preparacion)));
      setField(row,'codigo',it.codigo);
      setField(row,'descripcion',it.descripcion);
      setField(row,'ubicacion',it.ubicacion);
      setField(row,'cantidad',it.cantidad);
      setField(row,'status',status);
      actualizados++;
    }else{
      const obj = {
        id:id_('PD'),
        fecha:fecha,
        pedido:pedidoNuevo,
        cliente:cliente,
        vendedor:vendedor,
        pikeador:pikeador,
        bodega_preparacion:t_(b.bodega_preparacion || b.bodegaPreparacion || b.bodega || getCell_(primera, ix.bodega_preparacion)),
        codigo:it.codigo,
        descripcion:it.descripcion,
        ubicacion:it.ubicacion,
        cantidad:it.cantidad,
        status:status,
        alerta_token:'',
        alerta_ts:'',
        alerta_tipo:'',
        alerta_mensaje:'',
        hora_inicio:'',
        hora_termino:'',
        tiempo_preparacion_min:0,
        import_raw_json:'',
        import_headers_json:''
      };
      appendObjects_(sh, headers, [obj]);
      insertados++;
    }
  });

  if(filas.length > items.length){
    filas.slice(items.length).sort((a,b)=>b.row-a.row).forEach(x=>{ sh.deleteRow(x.row); eliminados++; });
  }

  if(pikeador) agregarPikeador_(pikeador);
  if(vendedor) agregarVendedor_(vendedor);

  let alerta = null;
  if(estadoAnterior !== status){
    alerta = emitirAlertaStatus_(pedidoNuevo, status, cliente, vendedor, pikeador, 'EDITAR_PEDIDO', true);
  }else{
    SpreadsheetApp.flush();
  }

  return {
    ok:true,
    msg:'Pedido editado correctamente',
    pedido:pedidoNuevo,
    pedido_original:pedidoOriginal,
    actualizados:actualizados,
    insertados:insertados,
    eliminados:eliminados,
    status:status,
    pikeador:pikeador,
    alerta_token:alerta ? alerta.alerta_token : '',
    alerta_mensaje:alerta ? alerta.alerta_mensaje : '',
    voice:alerta ? alerta.voice : '',
    serverTime:new Date().toISOString()
  };
}

function rowToObj_(r, ix){
  return {
    fecha:fechaVisible_(getCell_(r, ix.fecha)),
    pedido:getCell_(r, ix.pedido),
    cliente:getCell_(r, ix.cliente),
    vendedor:getCell_(r, ix.vendedor),
    pikeador:limpiarPikeador_(getCell_(r, ix.pikeador), getCell_(r, ix.cliente)),
    bodega_preparacion:getCell_(r, ix.bodega_preparacion),
    codigo:code_(getCell_(r, ix.codigo)),
    descripcion:getCell_(r, ix.descripcion),
    ubicacion:getCell_(r, ix.ubicacion),
    cantidad:getNumCell_(r, ix.cantidad) || 1,
    status:estadoSeguro_(getCell_(r, ix.status), 'PENDIENTE'),
    hora_inicio:fechaHoraVisible_(getCell_(r, ix.hora_inicio)),
    hora_termino:fechaHoraVisible_(getCell_(r, ix.hora_termino)),
    tiempo_preparacion_min:getNumCell_(r, ix.tiempo_preparacion_min)
  };
}
function pedidoCodigoKey_(pedido, codigo){
  const p = pedido_(pedido), c = code_(codigo);
  return p && c ? p + '|' + c : '';
}
function diffImportRow_(oldObj, newObj){
  const fields = ['fecha','cliente','vendedor','pikeador','bodega_preparacion','descripcion','ubicacion','cantidad','status'];
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

function mensajeVozEstado_(pedido, estado, pikeador, cliente, vendedor){
  pedido = t_(pedido);
  estado = estadoSeguro_(estado, 'PENDIENTE');
  cliente = t_(cliente) || clientePedido_(pedido);
  vendedor = t_(vendedor);
  pikeador = limpiarPikeador_(t_(pikeador), cliente);
  const terminado = estado === 'TERMINADO';
  const partes = [];
  if(terminado) partes.push('Pedido terminado');
  else partes.push('Pedido cambió a estado ' + estado);
  partes.push(pedido ? 'número de pedido ' + pedido : 'número de pedido sin registrar');
  partes.push(cliente ? 'cliente ' + cliente : 'cliente sin registrar');
  if(terminado){
    partes.push(vendedor ? 'vendedor asociado ' + vendedor : 'vendedor sin registrar');
  }else{
    partes.push(pikeador ? 'pikeador asignado ' + pikeador : 'pikeador sin asignar');
  }
  return partes.join(', ');
}

function emitirAlertaStatus_(pedido, estado, cliente, vendedor, pikeador, origen, actualizarPedidos){
  pedido = t_(pedido);
  estado = estadoSeguro_(estado, '');
  if(!pedido || !estado) return {ok:false,msg:'Pedido o estado vacío'};
  const tipo = alertaTipoEstado_(estado);
  const token = 'ALERTA-' + nh_(estado) + '-' + Date.now() + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  let count = 0;

  if(actualizarPedidos){
    const sh = pedidosSheet_(true);
    const cols = ['id','fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min'];
    ensureCols_(sh, cols);
    const v = values_(sh), ix = idx_(v.headers);
    const key = pedido_(pedido);
    v.data.forEach((r,i)=>{
      if(pedido_(getCell_(r, ix.pedido)) !== key) return;
      const row = i + 2;
      if(ix.alerta_token >= 0) sh.getRange(row, ix.alerta_token + 1).setValue(token);
      if(ix.alerta_ts >= 0) sh.getRange(row, ix.alerta_ts + 1).setValue(ts);
      if(ix.alerta_tipo >= 0) sh.getRange(row, ix.alerta_tipo + 1).setValue(tipo);
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

  const mensaje = mensajeVozEstado_(pedido, estado, pikeador, cliente, vendedor);
  if(actualizarPedidos){
    const shMsg = pedidosSheet_(false);
    if(shMsg){
      const vMsg = values_(shMsg), ixMsg = idx_(vMsg.headers), keyMsg = pedido_(pedido);
      if(ixMsg.alerta_mensaje >= 0){
        vMsg.data.forEach((r,i)=>{
          if(pedido_(getCell_(r, ixMsg.pedido)) === keyMsg) shMsg.getRange(i + 2, ixMsg.alerta_mensaje + 1).setValue(mensaje);
        });
      }
    }
  }

  registrarAlerta_(pedido, cliente, pikeador, tipo, mensaje, token, ts, estado, vendedor, origen || 'STATUS');
  SpreadsheetApp.flush();
  return {ok:true,msg:mensaje,pedido:pedido,status:estado,cliente:cliente,vendedor:vendedor,pikeador:pikeador,actualizados:count,alerta_token:token,alerta_ts:ts,alerta_tipo:tipo,alerta_mensaje:mensaje,voice:mensaje};
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
  const cols = ['id','fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min'];
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
    cliente:cliente,
    vendedor:vendedor,
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

  const estadoFinal = estadoSeguro_(b.status || b.estado || b.nuevo_estado || b.state || (cambiarEstado ? 'PREPARACION' : 'PREPARACION'), 'PREPARACION');
  tipo = alertaTipoEstado_(estadoFinal) || tipo || 'ALERTA';
  cambiarEstado = cambiarEstado || estadoFinal === 'PREPARACION';

  const sh = pedidosSheet_(true);
  const cols = ['id','fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min'];
  ensureCols_(sh, cols);
  const v = values_(sh), ix = idx_(v.headers);
  const token = 'ALERTA-' + nh_(estadoFinal) + '-' + Date.now() + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const key = pedido_(pedido);
  let count = 0;
  let cliente = t_(b.cliente);
  let vendedor = t_(b.vendedor || b.ejecutivo || b.responsable_venta);
  let pikeador = limpiarPikeador_(t_(b.pikeador || b.picker || b.preparador || ''), cliente);

  v.data.forEach((r,i)=>{
    if(pedido_(getCell_(r, ix.pedido)) !== key) return;
    const row = i + 2;
    if(!cliente) cliente = getCell_(r, ix.cliente);
    if(!vendedor) vendedor = getCell_(r, ix.vendedor);
    if(!pikeador) pikeador = limpiarPikeador_(getCell_(r, ix.pikeador), getCell_(r, ix.cliente));
    if(ix.alerta_token >= 0) sh.getRange(row, ix.alerta_token + 1).setValue(token);
    if(ix.alerta_ts >= 0) sh.getRange(row, ix.alerta_ts + 1).setValue(ts);
    if(ix.alerta_tipo >= 0) sh.getRange(row, ix.alerta_tipo + 1).setValue(tipo);
    if((cambiarEstado || estadoFinal) && ix.status >= 0) sh.getRange(row, ix.status + 1).setValue(estadoFinal);
    if(pikeador && ix.pikeador >= 0) sh.getRange(row, ix.pikeador + 1).setValue(pikeador);
    if((cambiarEstado || pikeador || estadoFinal === 'PREPARACION') && ix.hora_inicio >= 0 && !getCell_(r, ix.hora_inicio)) sh.getRange(row, ix.hora_inicio + 1).setValue(ts);
    if(estadoFinal === 'TERMINADO'){
      const inicio = getCell_(r, ix.hora_inicio) || ts;
      if(ix.hora_inicio >= 0 && !getCell_(r, ix.hora_inicio)) sh.getRange(row, ix.hora_inicio + 1).setValue(inicio);
      if(ix.hora_termino >= 0) sh.getRange(row, ix.hora_termino + 1).setValue(ts);
      if(ix.tiempo_preparacion_min >= 0) sh.getRange(row, ix.tiempo_preparacion_min + 1).setValue(minutosPreparacion_(inicio, ts, 0));
    }
    count++;
  });

  if(!count){
    guardarPedido_(Object.assign({}, b, {pedido:pedido, status:estadoFinal, pikeador:pikeador}));
    return alertaPedido_(b, tipo, mensaje, cambiarEstado);
  }

  const mensajeFinal = mensajeVozEstado_(pedido, estadoFinal, pikeador, cliente, vendedor);
  if(ix.alerta_mensaje >= 0){
    v.data.forEach((r,i)=>{
      if(pedido_(getCell_(r, ix.pedido)) === key) sh.getRange(i + 2, ix.alerta_mensaje + 1).setValue(mensajeFinal);
    });
  }

  if(pikeador) agregarPikeador_(pikeador);
  if(vendedor) agregarVendedor_(vendedor);
  registrarAlerta_(pedido, cliente, pikeador, tipo, mensajeFinal, token, ts, estadoFinal, vendedor, 'ALERTA_PEDIDO');
  SpreadsheetApp.flush();
  return {ok:true,msg:mensajeFinal,pedido:pedido,cliente:cliente,vendedor:vendedor,pikeador:pikeador,status:estadoFinal,actualizados:count,alerta_token:token,alerta_ts:ts,alerta_tipo:tipo,alerta_mensaje:mensajeFinal,voice:mensajeFinal};
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
        fecha:fechaVisible_(getCell_(r, ix.fecha)),
        cliente:getCell_(r, ix.cliente),
        vendedor:getCell_(r, ix.vendedor),
        pikeador:limpiarPikeador_(getCell_(r, ix.pikeador), getCell_(r, ix.cliente)),
        bodega_preparacion:getCell_(r, ix.bodega_preparacion),
        status:estadoSeguro_(getCell_(r, ix.status), 'PENDIENTE'),
        alerta_token:getCell_(r, ix.alerta_token),
        alerta_ts:getCell_(r, ix.alerta_ts),
        alerta_tipo:getCell_(r, ix.alerta_tipo),
        alerta_mensaje:getCell_(r, ix.alerta_mensaje),
        hora_inicio:fechaHoraVisible_(getCell_(r, ix.hora_inicio)),
        hora_termino:fechaHoraVisible_(getCell_(r, ix.hora_termino)),
        tiempo_preparacion_min:getNumCell_(r, ix.tiempo_preparacion_min),
        total_unidades:0,total_productos:0,productos:[],rows:[]
      };
    }
    const p = map[key], cant = getNumCell_(r, ix.cantidad) || 1;
    ['fecha','cliente','vendedor','bodega_preparacion','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino'].forEach(k=>{
      const val = getCell_(r, ix[k]);
      if(val) p[k] = k === 'fecha' ? fechaVisible_(val) : (k === 'hora_inicio' || k === 'hora_termino' ? fechaHoraVisible_(val) : val);
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
  const serial = Number(String(v).replace(',','.'));
  if(isFinite(serial) && serial > 20000 && serial < 90000){
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000));
    return isNaN(d.getTime()) ? null : d;
  }
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

function siguienteNumeroPedido_(b){
  const sh = pedidosSheet_(true);
  ensureCols_(sh, ['pedido']);
  const v = values_(sh);
  const ix = idx_(v.headers);
  if(ix.pedido >= 0){
    try{ sh.getRange(1, ix.pedido + 1, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat('@'); }catch(e){}
  }
  const next = siguienteNumeroPedidoDesdeValues_(v, ix);
  return {ok:true, pedido:next.numero, numero:next.numero, usados:next.usados, msg:'Número de pedido generado automáticamente'};
}

function siguienteNumeroPedidoDesdeValues_(v, ix){
  const usados = {};
  (v.data || []).forEach(r => {
    const raw = getCell_(r, ix.pedido);
    const n = numeroPedidoConsecutivo_(raw);
    if(n > 0) usados[n] = true;
  });
  let n = 1;
  while(usados[n]) n++;
  return {numero:String(n), usados:Object.keys(usados).length};
}

function numeroPedidoConsecutivo_(v){
  const raw = t_(v);
  if(!raw || /^SIN_PEDIDO/i.test(raw)) return 0;
  const m = raw.match(/\d+/g);
  if(!m) return 0;
  const n = Number(m.join(''));
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

function pasosPedido_(estado){
  const orden = ['PENDIENTE','PREPARACION','RECIBIDO','DESPACHADO','TERMINADO'];
  const actual = estadoSeguro_(estado, 'PENDIENTE');
  const ixActual = orden.indexOf(actual);
  return orden.map((x,i)=>({estado:x,activo:x===actual,completado:ixActual>=0 && i<=ixActual}));
}

function ss_(){ try{ if(SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID); }catch(e){} const ss=SpreadsheetApp.getActiveSpreadsheet(); if(!ss) throw new Error('No se pudo abrir la planilla. Revisa SPREADSHEET_ID.'); return ss; }
function pedidosSheet_(crear){ const ss=ss_(); let sh=ss.getSheetByName(SHEET_PEDIDOS)||ss.getSheetByName(SHEET_PEDIDO_ALT); if(!sh && crear) sh=getOrCreate_(SHEET_PEDIDOS,['id','fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','codigo','descripcion','ubicacion','cantidad','status','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino','tiempo_preparacion_min']); return sh; }
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
    bodega_preparacion:find_(h,['bodega_preparacion','bodega preparación','bodega preparacion','bodega pedido','bodega_origen','bodega de preparacion','bodega de preparación','nombre_bodega','nombre bodega','bodega'],-1),
    codigo:find_(h,['codigo','código','cod','sku','codigo producto','código producto'],-1),
    descripcion:find_(h,['descripcion','descripción','detalle','nombre','producto','nombre producto','descripcion producto','descripción producto'],-1),
    ubicacion:find_(h,['ubicacion','ubicación','ubicacion producto','ubicación producto','posicion','posición','rack','pasillo'],-1),
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

function fechaVisible_(valor){
  let v = t_(valor);
  if(!v) return '';
  if(v.charAt(0)==="'") v = v.slice(1).trim();
  const n = Number(String(v).replace(',', '.'));
  if(isFinite(n) && n > 20000 && n < 90000){
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n * 86400000));
    const dd = String(d.getUTCDate()).padStart(2,'0');
    const mm = String(d.getUTCMonth()+1).padStart(2,'0');
    const yy = d.getUTCFullYear();
    const frac = n - Math.floor(n);
    if(frac > 0.0007){
      const mins = Math.round(frac * 1440);
      return dd + '-' + mm + '-' + yy + ' ' + String(Math.floor(mins/60)%24).padStart(2,'0') + ':' + String(mins%60).padStart(2,'0');
    }
    return dd + '-' + mm + '-' + yy;
  }
  let m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::\d{2})?)?/);
  if(m) return String(Number(m[3])).padStart(2,'0') + '-' + String(Number(m[2])).padStart(2,'0') + '-' + m[1] + (m[4] ? ' ' + String(Number(m[4])).padStart(2,'0') + ':' + m[5] : '');
  m = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[T\s](\d{1,2}):(\d{2})(?::\d{2})?)?/);
  if(m) return String(Number(m[1])).padStart(2,'0') + '-' + String(Number(m[2])).padStart(2,'0') + '-' + m[3] + (m[4] ? ' ' + String(Number(m[4])).padStart(2,'0') + ':' + m[5] : '');
  return v;
}
function fechaHoraVisible_(valor){ return fechaVisible_(valor); }

function getNumCell_(r, idx){ return (idx == null || idx < 0 || !Array.isArray(r)) ? 0 : n_(r[idx]); }
function find_(h,names,fb){ const a=h.map(nh_); for(const n of names){ const i=a.indexOf(nh_(n)); if(i!==-1) return i; } return fb; }
function out_(o,cb){ const j=JSON.stringify(o); return ContentService.createTextOutput(cb?cb+'('+j+');':j).setMimeType(cb?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON); }
function outPost_(o,b){
  b = b || {};
  if(b.__iframe_post || b.iframe_post || b.iframe || b.__post_token){
    const token = t_(b.__post_token || b.post_token || '');
    const payload = JSON.stringify({source:'orden_pedidos_post',token:token,data:o}).replace(/</g,'\\u003c');
    return HtmlService.createHtmlOutput('<!doctype html><html><body><script>try{parent.postMessage('+payload+',"*");}catch(e){}<\/script></body></html>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return out_(o,'');
}
function body_(e){
  try{ if(e&&e.parameter&&e.parameter.data) return JSON.parse(e.parameter.data); }catch(x){}
  try{ if(e&&e.parameter&&e.parameter.payload) return JSON.parse(e.parameter.payload); }catch(x){}
  try{ if(e&&e.postData&&e.postData.contents){
    const raw = String(e.postData.contents || '');
    if(raw && raw.trim().charAt(0)==='{') return JSON.parse(raw);
  }}catch(x){}
  return e&&e.parameter?Object.assign({},e.parameter):{};
}
function parseItems_(v){ if(!v) return []; if(Array.isArray(v)) return v; try{ const x=JSON.parse(String(v)); return Array.isArray(x)?x:[]; }catch(e){ return []; } }

function limpiarPikeador_(pk, cliente){
  pk = t_(pk);
  cliente = t_(cliente);
  if(!pk) return '';
  if(cliente && nh_(pk) === nh_(cliente)) return '';
  if(estadoSeguro_(pk, '')) return '';
  return pk;
}
function bool_(v){ if(v === true) return true; const x = nh_(v); return ['1','TRUE','SI','SÍ','YES','ON','AUTO','AUTOGENERAR'].includes(x); }
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
