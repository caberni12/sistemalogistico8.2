const API_URL='https://script.google.com/macros/s/AKfycbw47SxU42yTG1yb8Tlc7H3fnhH77SZVEYSlqyTMI8xzsXEHDK0i7PDDFaj3-Y29vQo7Ng/exec';

const state={
  pedidos:[],
  filtrados:[],
  pikeadores:[],
  maestra:[],
  maestraMap:{},
  ubicacionesMap:{},
  maestraListaCargada:false,
  sel:null,
  importados:[],
  importStats:{nuevos:0,cambios:0,iguales:0,errores:0},
  importHeaders:[],
  manualProductos:[],
  tokens:JSON.parse(localStorage.getItem('orden_tokens')||'{}')
};

const $=id=>document.getElementById(id);

/* ================= LOADER BOTONES ================= */
function loaderStart(btn){
  if(!btn || btn.dataset.noLoader==='true') return;
  btn.dataset.loaderStartedAt=String(Date.now());
  btn.classList.add('btn-loading');
  btn.setAttribute('aria-busy','true');
  btn.disabled=true;
}
function loaderStop(btn){
  if(!btn) return;
  const started=Number(btn.dataset.loaderStartedAt||Date.now());
  const wait=Math.max(0,450-(Date.now()-started));
  setTimeout(()=>{
    btn.classList.remove('btn-loading');
    btn.removeAttribute('aria-busy');
    btn.disabled=false;
    delete btn.dataset.loaderStartedAt;
  },wait);
}
async function withLoader(btn,fn){
  loaderStart(btn);
  try{return await fn();}
  finally{loaderStop(btn);}
}
function withLoaderEvent(fn){
  return function(e){
    const btn=e&&e.currentTarget?e.currentTarget:null;
    return withLoader(btn,()=>fn(e));
  };
}
function stopClickedLoader(e){
  const btn=e&&e.currentTarget?e.currentTarget:null;
  setTimeout(()=>loaderStop(btn),650);
}

window.addEventListener('DOMContentLoaded',()=>{
  bind();
  cargarTodo();
  setInterval(verificarAlertas,1000);
});

function bind(){
  $('btnCargar')?.addEventListener('click',withLoaderEvent(()=>cargarTodo()));
  $('btnSync')?.addEventListener('click',withLoaderEvent(()=>cargarTodo()));
  $('btnSyncMaestra')?.addEventListener('click',withLoaderEvent(()=>cargarMaestra(false)));
  $('btnSyncUbicaciones')?.addEventListener('click',withLoaderEvent(()=>sincronizarUbicacionesManual(false)));
  $('btnPDF')?.addEventListener('click',withLoaderEvent(()=>pdfGeneral()));
  $('btnImportar')?.addEventListener('click',e=>{abrirImportar();stopClickedLoader(e);});
  $('btnManual')?.addEventListener('click',e=>{abrirManual();stopClickedLoader(e);});
  $('txtBuscar')?.addEventListener('input',filtrar);
  $('selEstado')?.addEventListener('change',filtrar);
  $('selFiltroPikeador')?.addEventListener('change',filtrar);
  $('btnAgregarPikeador')?.addEventListener('click',withLoaderEvent(()=>agregarPikeador()));
  $('btnTogglePanel')?.addEventListener('click',e=>{togglePanel();stopClickedLoader(e);});
  $('btnCerrarModal')?.addEventListener('click',e=>{cerrarModal();stopClickedLoader(e);});
  $('modalPedido')?.addEventListener('click',e=>{if(e.target.id==='modalPedido')cerrarModal();});
  $('btnAsignarPikeador')?.addEventListener('click',withLoaderEvent(()=>asignarPikeadorActual()));
  $('btnActualizarEstado')?.addEventListener('click',withLoaderEvent(()=>actualizarEstadoActual()));
  $('btnActualizarUbicacionesModal')?.addEventListener('click',withLoaderEvent(()=>sincronizarUbicacionesManual(true)));
  $('btnPreparar')?.addEventListener('click',withLoaderEvent(()=>prepararActual()));
  $('btnReenviar')?.addEventListener('click',withLoaderEvent(()=>reenviarActual()));
  $('btnPdfPedido')?.addEventListener('click',withLoaderEvent(()=>pdfPedidoA4()));
  $('btnPdfPedido80')?.addEventListener('click',withLoaderEvent(()=>pdfPedido80()));

  $('btnCerrarManual')?.addEventListener('click',e=>{cerrarManual();stopClickedLoader(e);});
  $('modalManual')?.addEventListener('click',e=>{if(e.target.id==='modalManual')cerrarManual();});
  $('btnManualAddProducto')?.addEventListener('click',withLoaderEvent(()=>agregarProductoManual()));
  $('btnBuscarCodigoManual')?.addEventListener('click',withLoaderEvent(()=>buscarProductoManualPorCodigo(true)));
  $('btnGuardarManual')?.addEventListener('click',withLoaderEvent(()=>guardarPedidoManual()));
  $('btnLimpiarManual')?.addEventListener('click',e=>{limpiarManual();stopClickedLoader(e);});
  ['manFecha','manPedido','manCliente','manVendedor','manPikeador','manStatus','manDescripcion','manUbicacion','manCantidad'].forEach(id=>$(id)?.addEventListener('input',renderManualPreview));
  $('manCodigo')?.addEventListener('input',()=>{ renderManualPreview(); debounceBuscarCodigoManual(); });
  $('manCodigo')?.addEventListener('change',()=>buscarProductoManualPorCodigo(false));
  $('manCodigo')?.addEventListener('blur',()=>buscarProductoManualPorCodigo(false));
  $('manPikeador')?.addEventListener('change',renderManualPreview);
  $('manStatus')?.addEventListener('change',renderManualPreview);

  $('btnCerrarImportar')?.addEventListener('click',e=>{cerrarImportar();stopClickedLoader(e);});
  $('modalImportar')?.addEventListener('click',e=>{if(e.target.id==='modalImportar')cerrarImportar();});
  $('btnProcesarImport')?.addEventListener('click',withLoaderEvent(()=>procesarImportacion()));
  $('btnEnviarImport')?.addEventListener('click',withLoaderEvent(()=>enviarImportacionBD()));
  $('btnLimpiarImport')?.addEventListener('click',e=>{limpiarImportacion();stopClickedLoader(e);});
  $('fileImportar')?.addEventListener('change',()=>limpiarImportacion(false));

  document.addEventListener('click',e=>{
    const b=e.target.closest('[data-ver-pedido]');
    if(b){ loaderStart(b); try{abrirModal(b.dataset.verPedido);} finally{setTimeout(()=>loaderStop(b),650);} }
  });
}

/* ================= CARGA PRINCIPAL ================= */
async function cargarTodo(){
  setStatus('Cargando hoja PEDIDOS y ubicaciones...');
  try{
    const ubicacionesPromise=cargarUbicacionesPedido().catch(err=>{
      console.warn('No se pudieron cargar ubicaciones múltiples', err);
      return state.ubicacionesMap || {};
    });
    let r=await api('listar_pedidos',{mostrarTodo:1});
    if(!r?.ok) throw new Error(r?.msg||'Respuesta inválida');
    state.pedidos=normPedidos(r);
    if(!state.pedidos.length){
      const b=await api('listar_bd',{sheet:'PEDIDOS'}).catch(()=>null);
      if(b?.ok) state.pedidos=normPedidos(b);
    }
    state.ubicacionesMap=await ubicacionesPromise;
    aplicarUbicacionesAPedidos();
    filtrar();
    setStatus('Conectado. Pedidos cargados: '+state.pedidos.length+' | Productos con ubicaciones: '+Object.keys(state.ubicacionesMap||{}).length);
    marcarTokens();
  }catch(e){
    console.error(e);
    setStatus('Error cargando PEDIDOS: '+e.message);
    errorTabla('No se pudo cargar PEDIDOS. Verifica doGet/listar_pedidos en Apps Script.');
  }
  cargarPikeadores().catch(console.warn);
  cargarMaestra(true).catch(console.warn);
}

async function cargarPikeadores(){
  const r=await api('listar_pikeadores',{});
  let l=[];
  if(r?.ok) l=normPikeadores(r);
  // Fuente única del select: hoja PIKEADORES.
  // No se rellenan pikeadores desde PEDIDOS para evitar que CLIENTE caiga en el select.
  state.pikeadores=[...new Set(l.map(x=>String(x||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  llenarSelects();
}
function normPikeadores(r){
  const a=r.pikeadores||r.data||r.rows||r.values||[];
  return Array.isArray(a)?a.map(x=>typeof x==='string'?x:Array.isArray(x)?x[0]:(x.nombre||x.pikeador||x.picker||x.preparador||x.responsable||'')).filter(Boolean):[];
}
function llenarSelects(){
  [$('selFiltroPikeador'),$('selModalPikeador'),$('impPikeador'),$('manPikeador')].filter(Boolean).forEach(s=>{
    const v=s.value;
    const first=s.id==='selFiltroPikeador'?'<option value="">Todos</option>':((s.id==='impPikeador'||s.id==='manPikeador')?'<option value="">Sin asignar</option>':'<option value="">-- Seleccionar Pikeador --</option>');
    s.innerHTML=first+state.pikeadores.map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join('');
    if(v)s.value=v;
  });
}


/* ================= MAESTRA DE PRODUCTOS =================
   Se sincroniza desde la hoja MAESTRA y se usa para completar
   la descripción del producto al digitar el código en Pedido Manual.
================================================== */
function normProductosMaestra(r){
  const arr = r?.productos || r?.data || r?.rows || r?.values || r?.items || [];
  if(!Array.isArray(arr)) return [];
  return arr.map(x=>{
    if(Array.isArray(x)) return {codigo:String(x[0]||'').trim(), descripcion:String(x[1]||'').trim(), cantidad:Number(String(x[2]||0).replace(',','.'))||0, ubicacion:String(x[3]||'').trim()};
    return {
      codigo:String(x.codigo||x.CODIGO||x['Código']||x['CÓDIGO']||x.cod||x.sku||x.producto||'').trim(),
      descripcion:String(x.descripcion||x.DESCRIPCION||x['Descripción']||x['DESCRIPCIÓN']||x.detalle||x.nombre||x.producto_descripcion||'').trim(),
      cantidad:Number(String(x.cantidad||x.CANTIDAD||x.unidades||x.stock||0).replace(',','.'))||0,
      ubicacion:String(x.ubicacion||x.UBICACION||x['Ubicación']||x['UBICACIÓN']||x.bodega||'').trim()
    };
  }).filter(x=>x.codigo || x.descripcion);
}
function guardarCacheMaestra(){
  try{ localStorage.setItem('orden_maestra_productos', JSON.stringify({ts:Date.now(), data:state.maestra.slice(0,8000)})); }catch(e){}
}
function cargarCacheMaestra(){
  try{
    const obj=JSON.parse(localStorage.getItem('orden_maestra_productos')||'{}');
    if(Array.isArray(obj.data) && obj.data.length){ setMaestra(obj.data); return true; }
  }catch(e){}
  return false;
}
function setMaestra(lista){
  state.maestra=(lista||[]).map(p=>({
    codigo:String(p.codigo||'').trim(),
    descripcion:String(p.descripcion||'').trim(),
    cantidad:Number(p.cantidad||0)||0,
    ubicacion:String(p.ubicacion||'').trim()
  })).filter(p=>p.codigo||p.descripcion);
  state.maestraMap={};
  state.maestra.forEach(p=>{
    const k=normCodigo(p.codigo);
    if(!k) return;
    if(!state.maestraMap[k]) state.maestraMap[k]={codigo:p.codigo,descripcion:p.descripcion,cantidad:0,ubicacion:''};
    const acc=state.maestraMap[k];
    if(!acc.descripcion && p.descripcion) acc.descripcion=p.descripcion;
    acc.cantidad=(Number(acc.cantidad)||0)+(Number(p.cantidad)||0);
    if(p.ubicacion){
      const txt=textoUbicacionCantidad(p.ubicacion,p.cantidad);
      acc.ubicacion=unirUbicacionTexto(acc.ubicacion,txt);
    }
  });
  state.maestraListaCargada=true;
  llenarDatalistMaestra();
}
function llenarDatalistMaestra(){
  const dl=$('dlProductosMaestra');
  if(!dl) return;
  const top=state.maestra.slice(0,2000);
  dl.innerHTML=top.map(p=>`<option value="${esc(p.codigo)}">${esc(p.descripcion)}</option>`).join('');
}
async function cargarMaestra(silencioso=true){
  if(silencioso && !state.maestra.length) cargarCacheMaestra();
  const msg=$('manualMaestraMsg');
  if(msg && !silencioso) msg.textContent='Sincronizando MAESTRA de productos...';
  const r=await api('listar_maestra',{});
  if(!r?.ok) throw new Error(r?.msg||'No se pudo leer MAESTRA');
  const lista=normProductosMaestra(r);
  setMaestra(lista);
  guardarCacheMaestra();
  if(msg) msg.textContent='MAESTRA sincronizada: '+state.maestra.length+' productos disponibles.';
  if(!silencioso) toast('MAESTRA sincronizada: '+state.maestra.length+' productos.');
  return state.maestra;
}

/* ================= UBICACIONES MÚLTIPLES POR PRODUCTO =================
   No modifica Apps Script. Toma la respuesta actual de MOVIMIENTO/listar_ubicaciones_actuales
   y acumula todas las ubicaciones del mismo código con su cantidad.
================================================== */
function textoUbicacionCantidad(ubicacion,cantidad){
  const u=String(ubicacion||'').trim();
  const c=String(cantidad??'').trim();
  if(!u) return '';
  return c && c!=='0' ? `${u} (${c})` : u;
}
function unirUbicacionTexto(actual,nuevo){
  const n=String(nuevo||'').trim();
  if(!n) return String(actual||'').trim();
  const partes=String(actual||'').split('|').map(x=>x.trim()).filter(Boolean);
  const baseNuevo=n.replace(/\s*\([^)]*\)\s*$/,'').trim().toUpperCase();
  const yaExiste=partes.some(p=>p.replace(/\s*\([^)]*\)\s*$/,'').trim().toUpperCase()===baseNuevo);
  if(yaExiste) return partes.join(' | ');
  partes.push(n);
  return partes.join(' | ');
}
function agregarUbicacionMultiple(map,codigo,ubicacion,cantidad){
  const k=normCodigo(codigo);
  const txt=textoUbicacionCantidad(ubicacion,cantidad);
  if(!k || !txt) return;
  map[k]=unirUbicacionTexto(map[k]||'',txt);
}
function parseUbicacionesPedido(r){
  const out={};
  const arr=r?.data||r?.rows||r?.values||r?.ubicaciones||r?.items||[];
  const headers=Array.isArray(r?.headers)?r.headers:[];
  const h=headers.map(norm);
  const idx=(names)=>{ for(const name of names){ const i=h.indexOf(norm(name)); if(i>=0) return i; } return -1; };
  const ix={
    codigo:idx(['codigo','código','cod','sku','producto','codigo producto','código producto']),
    ubicacion:idx(['ubicacion','ubicación','ubicaciones','bodega','ubic']),
    cantidad:idx(['cantidad','stock','existencia','unidades','qty','cant']),
    status:idx(['status','estado','vigencia'])
  };
  (Array.isArray(arr)?arr:[]).forEach(row=>{
    if(Array.isArray(row)){
      const estado=ix.status>=0 ? row[ix.status] : row[9];
      const estadoNorm=norm(estado);
      if(['RETIRADO','ELIMINADO','ANULADO','INACTIVO'].includes(estadoNorm)) return;
      const codigo=ix.codigo>=0 ? row[ix.codigo] : row[5];
      const ubicacion=ix.ubicacion>=0 ? row[ix.ubicacion] : row[4];
      const cantidad=ix.cantidad>=0 ? row[ix.cantidad] : row[7];
      agregarUbicacionMultiple(out,codigo,ubicacion,cantidad);
      return;
    }
    if(row && typeof row==='object'){
      const codigo=row.codigo||row.CODIGO||row['Código']||row['CÓDIGO']||row.cod||row.sku||row.producto||'';
      if(Array.isArray(row.ubicaciones)){
        row.ubicaciones.forEach(u=>{
          if(u && typeof u==='object') agregarUbicacionMultiple(out,codigo,u.ubicacion||u.UBICACION||u['Ubicación']||u.ubic||u.bodega,u.cantidad||u.CANTIDAD||u.stock||u.unidades);
          else agregarUbicacionMultiple(out,codigo,u,'');
        });
      }else{
        agregarUbicacionMultiple(out,codigo,row.ubicacion||row.UBICACION||row['Ubicación']||row['UBICACIÓN']||row.ubic||row.bodega,row.cantidad||row.CANTIDAD||row.stock||row.unidades);
      }
    }
  });
  return out;
}
async function cargarUbicacionesPedido(){
  const fuentes=[
    ['listar_bd',{sheet:'MOVIMIENTO'}],
    ['listar_bd',{sheet:'BD-MOVIMIENTO'}],
    ['listar_movimiento',{}],
    ['listar_ubicaciones_actuales',{}]
  ];
  let ultimoError=null;
  for(const [accion,params] of fuentes){
    try{
      const r=await api(accion,params);
      if(!r?.ok) throw new Error(r?.msg||('No se pudo leer '+accion));
      const mapa=parseUbicacionesPedido(r);
      if(Object.keys(mapa).length) return mapa;
      ultimoError=new Error('Sin ubicaciones en '+accion);
    }catch(err){ ultimoError=err; }
  }
  throw ultimoError||new Error('No se pudieron leer ubicaciones');
}
function aplicarUbicacionesAPedidos(){
  const mapa=state.ubicacionesMap||{};
  if(!Object.keys(mapa).length) return;
  state.pedidos=state.pedidos.map(p=>{
    const productos=(p.productos||[]).map(item=>{
      const loc=mapa[normCodigo(item.codigo)]||'';
      return loc ? {...item,ubicacion:loc} : item;
    });
    return {...p,productos};
  });
}

async function sincronizarUbicacionesManual(desdeModal=false){
  const pedidoAbierto=state.sel ? (state.sel.key || state.sel.pedido || '') : '';
  setStatus('Sincronizando ubicaciones desde MOVIMIENTO...');
  const mapa=await cargarUbicacionesPedido();
  state.ubicacionesMap=mapa||{};
  aplicarUbicacionesAPedidos();
  filtrar();
  const total=Object.keys(state.ubicacionesMap).length;
  setStatus('Ubicaciones sincronizadas desde MOVIMIENTO: '+total+' código(s).');
  toast('Ubicaciones actualizadas: '+total+' código(s).');
  if(desdeModal && pedidoAbierto){
    const actualizado=state.pedidos.find(x=>x.key===pedidoAbierto||x.pedido===pedidoAbierto);
    if(actualizado) abrirModal(actualizado.key||actualizado.pedido);
  }
  return state.ubicacionesMap;
}
function productoMaestraPorCodigo(codigo){
  const k=normCodigo(codigo);
  if(!k) return null;
  return state.maestraMap[k] || null;
}
let timerBuscarManual=null;
function debounceBuscarCodigoManual(){
  clearTimeout(timerBuscarManual);
  timerBuscarManual=setTimeout(()=>buscarProductoManualPorCodigo(false),450);
}
async function buscarProductoManualPorCodigo(forzar=false){
  const input=$('manCodigo');
  const desc=$('manDescripcion');
  const ubi=$('manUbicacion');
  const msg=$('manualMaestraMsg');
  const codigo=(input?.value||'').trim();
  if(!codigo) return;

  if(!state.maestraListaCargada){
    cargarCacheMaestra();
    if(!state.maestra.length){
      try{ await cargarMaestra(true); }catch(e){ console.warn(e); }
    }
  }

  let prod=productoMaestraPorCodigo(codigo);
  if(!prod){
    try{
      const r=await api('buscar_producto',{codigo});
      if(r?.ok) prod={codigo:r.codigo||codigo, descripcion:r.descripcion||'', cantidad:r.cantidad||0, ubicacion:r.ubicacion||''};
    }catch(e){ console.warn(e); }
  }

  if(prod){
    if(desc && (forzar || !desc.value.trim())) desc.value=prod.descripcion||desc.value;
    if(ubi && (forzar || !ubi.value.trim()) && prod.ubicacion) ubi.value=prod.ubicacion;
    if(msg) msg.textContent='Producto encontrado en MAESTRA: '+(prod.descripcion||prod.codigo);
    renderManualPreview();
    return prod;
  }
  if(msg) msg.textContent='Código no encontrado en MAESTRA. Puedes ingresar la descripción manualmente.';
  return null;
}



/* ================= PROTECCIÓN PIKEADOR / CLIENTE =================
   Pikeador nunca debe copiarse desde Cliente ni desde columnas de la TD.
   Solo se acepta si viene desde columna PIKEADOR/PICKER/PREPARADOR o desde la maestra PIKEADORES.
================================================== */
function mismoTexto(a,b){ return norm(a) && norm(a) === norm(b); }
function limpiarPikeador(pk, cliente){
  pk = String(pk || '').trim();
  cliente = String(cliente || '').trim();
  if(!pk) return '';
  // Si por error vino igual al cliente, lo limpiamos para no mostrar cliente como pikeador.
  if(cliente && mismoTexto(pk, cliente)) return '';
  // Estados o textos genéricos tampoco son pikeadores.
  if(normalizarEstadoPermitido(pk)) return '';
  return pk;
}
function pikeadorVisible(pk){ return String(pk || '').trim() || 'Sin asignar'; }

/* ================= NORMALIZACIÓN PEDIDOS ================= */
function normPedidos(r){
  if(Array.isArray(r.pedidos)&&r.pedidos.length){
    return r.pedidos.map(p=>({
      key:String(p.key||p.pedido||'').trim().toUpperCase(),
      pedido:String(p.pedido||'').trim(),
      fecha:String(p.fecha||'').trim(),
      cliente:String(p.cliente||'').trim(),
      vendedor:String(p.vendedor||'').trim(),
      pikeador:limpiarPikeador(p.pikeador, p.cliente),
      status:estadoSeguroImport(p.status||p.estado,'PENDIENTE'),
      alerta_token:String(p.alerta_token||'').trim(),
      alerta_ts:String(p.alerta_ts||'').trim(),
      alerta_tipo:String(p.alerta_tipo||'').trim(),
      alerta_mensaje:String(p.alerta_mensaje||'').trim(),
      hora_inicio:String(p.hora_inicio||p.horaInicio||p.inicio_preparacion||p.fecha_inicio||p.fecha_asignacion||'').trim(),
      hora_termino:String(p.hora_termino||p.horaTermino||p.termino_preparacion||p.fecha_termino||'').trim(),
      tiempo_preparacion_min:Number(p.tiempo_preparacion_min||p.tiempoPreparacionMin||p.tiempo_preparacion||0)||0,
      total_productos:Number(p.total_productos||(p.productos?p.productos.length:0)||0),
      total_unidades:Number(p.total_unidades||0),
      productos:Array.isArray(p.productos)?p.productos:[],
      rows:Array.isArray(p.rows)?p.rows:[]
    }));
  }
  return agrupar(r.headers||r.rawHeaders||[],r.data||r.rawData||r.rows||r.values||[]);
}
function agrupar(h,rows){
  const ix=mapH(h); const m=new Map(); if(!Array.isArray(rows))return[];
  rows.forEach((r,i)=>{
    if(!Array.isArray(r))return;
    const ped=pick(r,ix.pedido,'SIN_PEDIDO_'+(i+1));
    const key=String(ped).replace(/\s+/g,'').toUpperCase();
    if(!m.has(key))m.set(key,{key,pedido:ped,fecha:pick(r,ix.fecha,''),cliente:pick(r,ix.cliente,''),vendedor:pick(r,ix.vendedor,''),pikeador:limpiarPikeador(pick(r,ix.pikeador,''), pick(r,ix.cliente,'')),status:estadoSeguroImport(pick(r,ix.status,''),'PENDIENTE'),alerta_token:pick(r,ix.alerta_token,''),alerta_ts:pick(r,ix.alerta_ts,''),alerta_tipo:pick(r,ix.alerta_tipo,''),alerta_mensaje:pick(r,ix.alerta_mensaje,''),hora_inicio:pick(r,ix.hora_inicio,''),hora_termino:pick(r,ix.hora_termino,''),tiempo_preparacion_min:Number(String(pick(r,ix.tiempo_preparacion_min,0)).replace(',','.'))||0,total_productos:0,total_unidades:0,productos:[],rows:[]});
    const p=m.get(key);
    ['fecha','cliente','vendedor','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino'].forEach(k=>{if(ix[k]>=0&&pick(r,ix[k],''))p[k]=pick(r,ix[k],'')});
    if(ix.tiempo_preparacion_min>=0){ const tm=Number(String(pick(r,ix.tiempo_preparacion_min,0)).replace(',','.'))||0; if(tm>0)p.tiempo_preparacion_min=tm; }
    if(ix.pikeador>=0){ const pk=limpiarPikeador(pick(r,ix.pikeador,''), p.cliente||pick(r,ix.cliente,'')); if(pk) p.pikeador=pk; }
    if(ix.status>=0 && normalizarEstadoPermitido(pick(r,ix.status,''))) p.status=normalizarEstadoPermitido(pick(r,ix.status,''));
    const cant=Number(String(pick(r,ix.cantidad,1)).replace(',','.'))||1;
    const prod={codigo:pick(r,ix.codigo,''),descripcion:pick(r,ix.descripcion,''),ubicacion:pick(r,ix.ubicacion,''),cantidad:cant};
    if(prod.codigo||prod.descripcion){p.productos.push(prod);p.total_productos++;p.total_unidades+=cant;}
    p.rows.push(r);
  });
  return Array.from(m.values()).sort((a,b)=>numPed(b.pedido)-numPed(a.pedido));
}
function headerIndexStrict(headers, names){
  const normalized = (headers || []).map(norm);
  for(const name of names){
    const idx = normalized.indexOf(norm(name));
    if(idx !== -1) return idx;
  }
  return -1;
}
function hasUsableHeaders(headers){
  const h = (headers || []).map(norm).filter(Boolean);
  const core = ['pedido','codigo','descripcion','cliente','estado','status'];
  return core.some(c => h.includes(norm(c)));
}
function mapH(h){
  // Mapeo seguro por encabezado. No usamos posiciones por defecto porque eso cruzaba datos.
  // Si el Apps Script entrega headers, cada columna se toma sólo por su nombre real.
  const headers = Array.isArray(h) ? h : [];
  const i = (names) => headerIndexStrict(headers, names);
  return {
    id:i(['id','ID']),
    fecha:i(['fecha','fecha ingreso','fecha_pedido','fecha pedido','fecha creacion','fecha creación']),
    pedido:i(['pedido','nro pedido','numero pedido','número pedido','n° pedido','nº pedido','orden','orden pedido']),
    cliente:i(['cliente','nombre cliente','razon social','razón social']),
    vendedor:i(['vendedor','responsable','ejecutivo']),
    pikeador:i(['pikeador','picker','preparador','asignado','asignado a']),
    codigo:i(['codigo','código','cod','sku','codigo producto','código producto']),
    descripcion:i(['descripcion','descripción','detalle','nombre','producto','nombre producto','descripcion producto','descripción producto']),
    ubicacion:i(['ubicacion','ubicación','bodega','ubicacion producto','ubicación producto']),
    cantidad:i(['cantidad','unidades','cajas','qty','cant']),
    status:i(['status','estado','estado pedido','estado del pedido']),
    alerta_token:i(['alerta_token','token_alerta','alert token','alerta token']),
    alerta_ts:i(['alerta_ts','fecha_alerta','alert_ts','alerta fecha']),
    alerta_tipo:i(['alerta_tipo','tipo_alerta','tipo alerta']),
    alerta_mensaje:i(['alerta_mensaje','mensaje_alerta','mensaje alerta']),
    hora_inicio:i(['hora_inicio','hora inicio','inicio_preparacion','inicio preparación','fecha_inicio','fecha inicio','fecha_asignacion','fecha asignacion','asignado_en','asignado en']),
    hora_termino:i(['hora_termino','hora termino','hora término','termino_preparacion','término preparación','termino preparación','fecha_termino','fecha término','finalizado_en','finalizado en']),
    tiempo_preparacion_min:i(['tiempo_preparacion_min','tiempo preparacion min','tiempo preparación min','minutos_preparacion','minutos preparación','duracion_min','duración min'])
  };
}

/* ================= FILTRO / RENDER ================= */
function filtrar(){
  const q=($('txtBuscar')?.value||'').toLowerCase().trim(), e=($('selEstado')?.value||'').toUpperCase().trim(), pk=($('selFiltroPikeador')?.value||'').toLowerCase().trim();
  state.filtrados=state.pedidos.filter(p=>{
    const txt=[p.pedido,p.cliente,p.vendedor,p.pikeador,p.status,...p.productos.map(x=>`${x.codigo} ${x.descripcion} ${x.ubicacion}`)].join(' ').toLowerCase();
    if(q&&!txt.includes(q))return false;
    if(e&&String(p.status||'').toUpperCase()!==e)return false;
    if(pk&&String(p.pikeador||'').toLowerCase()!==pk)return false;
    return true;
  });
  render(); kpis();
}
function render(){
  const tb=$('tbodyPedidos'), mob=$('mobileList');
  if(!state.filtrados.length){
    tb.innerHTML='<tr><td colspan="13" style="text-align:center;padding:26px;color:#64748b">Sin registros para mostrar. Presiona Cargar lista o revisa la hoja PEDIDOS.</td></tr>';
    mob.innerHTML='<div class="pedido-card">Sin registros para mostrar.</div>';
    return;
  }
  tb.innerHTML=state.filtrados.map(p=>`<tr><td>${esc(p.fecha)}</td><td><b>${esc(p.pedido)}</b></td><td>${esc(p.cliente)}</td><td>${esc(p.vendedor)}</td><td>${esc(pikeadorVisible(p.pikeador))}</td><td>${p.total_productos}</td><td><b>${p.total_unidades}</b></td><td>${badgeEstado(p.status)}</td><td>${esc(horaCorta(p.hora_inicio)||'-')}</td><td>${esc(horaCorta(p.hora_termino)||'-')}</td><td><b>${esc(tiempoPreparacionTexto(p))}</b></td><td>${badgeAlerta(p)}</td><td class="actions-cell"><button data-ver-pedido="${esc(p.key)}">Ver</button></td></tr>`).join('');
  mob.innerHTML=state.filtrados.map(p=>`<div class="pedido-card"><div class="top"><span>${esc(p.pedido)}</span>${badgeEstado(p.status)}</div><p><b>Cliente:</b> ${esc(p.cliente||'-')}</p><p><b>Pikeador:</b> ${esc(pikeadorVisible(p.pikeador))}</p><p><b>Productos:</b> ${p.total_productos} | <b>Unidades:</b> ${p.total_unidades}</p><p><b>Inicio:</b> ${esc(horaCorta(p.hora_inicio)||'-')} | <b>Término:</b> ${esc(horaCorta(p.hora_termino)||'-')}</p><p><b>Tiempo preparación:</b> ${esc(tiempoPreparacionTexto(p))}</p><p>${badgeAlerta(p)}</p><button data-ver-pedido="${esc(p.key)}">Ver opciones</button></div>`).join('');
}
function kpis(){
  const a=state.filtrados;
  const procesados=a.filter(esPedidoProcesado);
  const tiempos=procesados.map(minutosPreparacionPedido).filter(x=>x>0);
  const setTxt=(id,val)=>{ const el=$(id); if(el) el.textContent=val; };
  setTxt('kPedidos',a.length);
  setTxt('kProductos',a.reduce((s,p)=>s+p.total_productos,0));
  setTxt('kUnidades',a.reduce((s,p)=>s+p.total_unidades,0));
  setTxt('kPendientes',a.filter(p=>String(p.status||'').toUpperCase()!=='TERMINADO').length);
  setTxt('kProcesados',procesados.length);
  setTxt('kPromedioProductos',procesados.length?(procesados.reduce((s,p)=>s+p.total_productos,0)/procesados.length).toFixed(1):'0');
  setTxt('kPromedioUnidades',procesados.length?(procesados.reduce((s,p)=>s+p.total_unidades,0)/procesados.length).toFixed(1):'0');
  setTxt('kTiempoPromedio',tiempos.length?formatearMinutos(tiempos.reduce((s,x)=>s+x,0)/tiempos.length):'0 min');
  renderRotacionYRendimiento(a);
}

/* ================= MODAL PEDIDO ================= */
function abrirModal(key){
  const p=state.pedidos.find(x=>x.key===key||x.pedido===key);
  if(!p)return toast('No se encontró el pedido');
  state.sel=p;
  $('modalTitle').textContent='Pedido '+p.pedido;
  $('modalSub').textContent=(p.cliente||'')+' | Total unidades: '+(p.total_unidades||0);
  $('detallePedido').innerHTML=[['Fecha',p.fecha],['Pedido',p.pedido],['Cliente',p.cliente],['Vendedor',p.vendedor],['Pikeador',pikeadorVisible(p.pikeador)],['Estado',p.status],['Hora inicio',horaCorta(p.hora_inicio)||'-'],['Hora término',horaCorta(p.hora_termino)||'-'],['Tiempo preparación',tiempoPreparacionTexto(p)],['Productos',p.total_productos],['Total unidades',p.total_unidades],['Avance cliente',porcentajeAvancePedido(p)+'%']].map(([l,v])=>`<div class="detail-item"><div class="l">${l}</div><div class="v">${esc(v)}</div></div>`).join('');
  $('tbodyProductos').innerHTML=p.productos.length?p.productos.map(x=>`<tr><td>${esc(x.codigo)}</td><td>${esc(x.descripcion)}</td><td>${esc(x.ubicacion)}</td><td><b>${x.cantidad}</b></td></tr>`).join(''):'<tr><td colspan="4" style="text-align:center;color:#64748b">Este pedido no tiene productos detallados.</td></tr>';
  llenarSelects();
  $('selModalPikeador').value=p.pikeador||'';
  if($('selModalEstado')) $('selModalEstado').value=(p.status||'PENDIENTE').toUpperCase();
  $('modalPedido').classList.add('show');
}
function cerrarModal(){ $('modalPedido').classList.remove('show'); }
async function agregarPikeador(){
  const nombre=($('txtNuevoPikeador')?.value||'').trim();
  if(!nombre)return toast('Ingresa el nombre del pikeador');
  const r=await api('agregar_pikeador',{nombre});
  if(!r.ok)return toast(r.msg||'No se pudo agregar');
  $('txtNuevoPikeador').value='';
  await cargarPikeadores();
  toast('Pikeador guardado');
}
async function asignarPikeadorActual(){
  const p=state.sel, pk=$('selModalPikeador').value;
  if(!p||!pk)return toast('Selecciona un pikeador');
  const r=await api('asignar_pikeador_pedido',{pedido:p.pedido,pikeador:pk});
  if(!r.ok)return toast(r.msg||'No se pudo asignar');
  toast('Pikeador asignado. Hora de inicio registrada.');
  await cargarTodo(); abrirModal(p.key);
}
async function actualizarEstadoActual(){
  const p=state.sel;
  const estado=($('selModalEstado')?.value||'').trim().toUpperCase();
  const pk=$('selModalPikeador')?.value||p?.pikeador||'';
  if(!p)return toast('No hay pedido seleccionado');
  if(!estado)return toast('Selecciona un estado');
  if(!confirm(`Actualizar pedido ${p.pedido} al estado ${estado}?`))return;
  const r=await api('actualizar_estado_pedido',{pedido:p.pedido,status:estado,pikeador:pk});
  if(!r.ok)return toast(r.msg||'No se pudo actualizar estado');
  toast('Estado actualizado: '+estado);
  const vozMsg = estado==='TERMINADO' ? mensajePedidoTerminado(p) : (r.voice || (estado==='PREPARACION' ? 'Pedido enviado a preparación' : ''));
  if(vozMsg) voz(vozMsg);
  await cargarTodo();
  abrirModal(p.key);
}
async function prepararActual(){
  const p=state.sel; if(!p)return;
  const pk=$('selModalPikeador').value||p.pikeador||'';
  if(!confirm(`Enviar pedido ${p.pedido} a preparación?\nProductos: ${p.total_productos}\nUnidades: ${p.total_unidades}`))return;
  const r=await api('preparar_pedido',{pedido:p.pedido,pikeador:pk});
  if(!r.ok)return toast(r.msg||'No se pudo enviar');
  voz('Pedido enviado a preparación'); toast('Pedido enviado a preparación. Hora de inicio registrada.');
  await cargarTodo(); await verificarAlertas(); abrirModal(p.key);
}
async function reenviarActual(){
  const p=state.sel; if(!p)return;
  const r=await api('reenviar_alerta_pedido',{pedido:p.pedido,pikeador:$('selModalPikeador').value||p.pikeador||''});
  if(!r.ok)return toast(r.msg||'No se pudo reenviar');
  voz('Pedido enviado a preparación'); toast('Alerta reenviada');
  await cargarTodo(); await verificarAlertas(); abrirModal(p.key);
}
async function verificarAlertas(){
  try{
    const r=await api('listar_pedidos',{mostrarTodo:1});
    if(!r?.ok)return;
    const pedidos=normPedidos(r);
    let hubo=false;
    pedidos.forEach(p=>{
      if(!p.alerta_token)return;
      const old=state.tokens[p.key];
      if(old&&old!==p.alerta_token){
        hubo=true;
        const estadoAlerta = String(p.status || p.estado || '').trim().toUpperCase();
        const msgAlerta = p.alerta_mensaje || (estadoAlerta === 'TERMINADO' ? mensajePedidoTerminado(p) : (estadoAlerta === 'PREPARACION' ? 'Pedido enviado a preparación' : ('Pedido ' + p.pedido + ' cambió de estado')));
        toast('🔔 ' + msgAlerta + ': ' + p.pedido);
        voz(msgAlerta);
      }
      state.tokens[p.key]=p.alerta_token;
    });
    localStorage.setItem('orden_tokens',JSON.stringify(state.tokens));
    if(hubo){state.pedidos=pedidos;filtrar();}
  }catch(e){}
}
function marcarTokens(){ state.pedidos.forEach(p=>{if(p.alerta_token)state.tokens[p.key]=p.alerta_token}); localStorage.setItem('orden_tokens',JSON.stringify(state.tokens)); }



/* ================= PEDIDO MANUAL ================= */
function abrirManual(){
  llenarSelects();
  if(!state.maestraListaCargada) cargarMaestra(true).catch(console.warn);
  limpiarManual(false);
  const hoy=new Date();
  const yyyy=hoy.getFullYear(), mm=String(hoy.getMonth()+1).padStart(2,'0'), dd=String(hoy.getDate()).padStart(2,'0');
  if($('manFecha')&&!$('manFecha').value) $('manFecha').value=`${yyyy}-${mm}-${dd}`;
  $('modalManual')?.classList.add('show');
  setTimeout(()=>$('manPedido')?.focus(),80);
  renderManualPreview();
}
function cerrarManual(){ $('modalManual')?.classList.remove('show'); }
function limpiarManual(clearHeader=true){
  state.manualProductos=[];
  if(clearHeader){
    ['manPedido','manCliente','manVendedor','manCodigo','manDescripcion','manUbicacion'].forEach(id=>{ if($(id)) $(id).value=''; });
    if($('manCantidad')) $('manCantidad').value='1';
    if($('manStatus')) $('manStatus').value='PENDIENTE';
    if($('manPikeador')) $('manPikeador').value='';
  }else{
    ['manCodigo','manDescripcion','manUbicacion'].forEach(id=>{ if($(id)) $(id).value=''; });
    if($('manCantidad')) $('manCantidad').value='1';
  }
  renderManualProductos();
  renderManualPreview();
}
function datosManualBase(){
  return {
    fecha:$('manFecha')?.value||'',
    pedido:($('manPedido')?.value||'').trim(),
    cliente:($('manCliente')?.value||'').trim(),
    vendedor:($('manVendedor')?.value||'').trim(),
    pikeador:($('manPikeador')?.value||'').trim(),
    status:estadoSeguroImport($('manStatus')?.value||'PENDIENTE','PENDIENTE')
  };
}
function agregarProductoManual(){
  const base=datosManualBase();
  const codigo=($('manCodigo')?.value||'').trim();
  const descripcion=($('manDescripcion')?.value||'').trim();
  const ubicacion=($('manUbicacion')?.value||'').trim();
  const cantidad=Number(String($('manCantidad')?.value||'1').replace(',','.'))||1;
  if(!base.pedido) return toast('Ingresa el número de pedido antes de agregar productos.');
  if(!codigo && !descripcion) return toast('Ingresa código o descripción del producto.');
  state.manualProductos.push({codigo,descripcion,ubicacion,cantidad});
  ['manCodigo','manDescripcion','manUbicacion'].forEach(id=>{ if($(id)) $(id).value=''; });
  if($('manCantidad')) $('manCantidad').value='1';
  renderManualProductos();
  renderManualPreview();
}
function quitarProductoManual(i){
  state.manualProductos.splice(i,1);
  renderManualProductos();
  renderManualPreview();
}
window.quitarProductoManual=quitarProductoManual;
function renderManualProductos(){
  const tbody=$('tbodyManualProductos');
  if(!tbody) return;
  if(!state.manualProductos.length){
    tbody.innerHTML='<tr><td colspan="6" style="text-align:center;color:#64748b;padding:18px">Aún no hay productos agregados.</td></tr>';
    return;
  }
  tbody.innerHTML=state.manualProductos.map((p,i)=>`<tr><td>${i+1}</td><td>${esc(p.codigo)}</td><td>${esc(p.descripcion)}</td><td>${esc(p.ubicacion)}</td><td>${p.cantidad}</td><td><button class="secondary" onclick="quitarProductoManual(${i})">Quitar</button></td></tr>`).join('');
}
function renderManualPreview(){
  const box=$('manualPreview');
  if(!box) return;
  const base=datosManualBase();
  const total=state.manualProductos.reduce((s,x)=>s+(Number(x.cantidad)||0),0);
  const parcial={codigo:$('manCodigo')?.value||'',descripcion:$('manDescripcion')?.value||'',ubicacion:$('manUbicacion')?.value||'',cantidad:$('manCantidad')?.value||''};
  const parcialTxt=(parcial.codigo||parcial.descripcion)?`<br><span style="color:#0f766e"><b>Producto en edición:</b> ${esc(parcial.codigo)} ${esc(parcial.descripcion)} | Ubicación: ${esc(parcial.ubicacion)} | Cantidad: ${esc(parcial.cantidad||1)}</span>`:'';
  box.innerHTML=`<b>Pedido:</b> ${esc(base.pedido||'Sin número')} | <b>Cliente:</b> ${esc(base.cliente||'Sin cliente')} | <b>Pikeador:</b> ${esc(base.pikeador||'Sin asignar')} | <b>Estado:</b> ${esc(base.status)} | <b>Productos agregados:</b> ${state.manualProductos.length} | <b>Total unidades:</b> ${total}${parcialTxt}`;
}
async function guardarPedidoManual(){
  const base=datosManualBase();
  if(!base.pedido) return toast('Ingresa el número de pedido.');
  if(!base.cliente) return toast('Ingresa el cliente.');
  if(!state.manualProductos.length) return toast('Agrega al menos un producto al pedido.');
  const items=state.manualProductos.map(p=>({
    fecha:base.fecha,
    pedido:base.pedido,
    cliente:base.cliente,
    vendedor:base.vendedor,
    pikeador:base.pikeador,
    status:base.status,
    codigo:p.codigo,
    descripcion:p.descripcion,
    ubicacion:p.ubicacion,
    cantidad:p.cantidad,
    __raw:{fecha:base.fecha,pedido:base.pedido,cliente:base.cliente,vendedor:base.vendedor,pikeador:base.pikeador,status:base.status,codigo:p.codigo,descripcion:p.descripcion,ubicacion:p.ubicacion,cantidad:p.cantidad,origen:'MANUAL'},
    __rawHeaders:['fecha','pedido','cliente','vendedor','pikeador','status','codigo','descripcion','ubicacion','cantidad','origen']
  }));
  const r=await api('guardar_pedido',{pedido:base.pedido,cliente:base.cliente,vendedor:base.vendedor,pikeador:base.pikeador,status:base.status,fecha:base.fecha,items:JSON.stringify(items)});
  if(!r?.ok) return toast('No se pudo guardar: '+(r?.msg||'error'));
  toast(`Pedido guardado. Insertados: ${r.insertados||0}. Actualizados: ${r.actualizados||0}. Sin cambio: ${r.sin_cambios||0}.`);
  cerrarManual();
  await cargarTodo();
}

/* ================= IMPORTAR LISTADO ================= */
function abrirImportar(){
  llenarSelects();
  limpiarImportacion(false);
  const hoy=new Date();
  const yyyy=hoy.getFullYear(), mm=String(hoy.getMonth()+1).padStart(2,'0'), dd=String(hoy.getDate()).padStart(2,'0');
  if($('impFecha')&&!$('impFecha').value) $('impFecha').value=`${yyyy}-${mm}-${dd}`;
  $('modalImportar')?.classList.add('show');
}
function cerrarImportar(){ $('modalImportar')?.classList.remove('show'); }
function setImportProgress(percent,msg='',detail='',mode=''){
  const p=Math.max(0,Math.min(100,Math.round(Number(percent)||0)));
  const wrap=$('importProgress');
  const fill=$('importProgressFill');
  const pct=$('importProgressPercent');
  const text=$('importProgressMsg');
  const det=$('importProgressDetail');
  if(wrap){
    wrap.classList.remove('ok','error');
    if(mode) wrap.classList.add(mode);
  }
  if(fill) fill.style.width=p+'%';
  if(pct) pct.textContent=p+'%';
  if(text && msg) text.textContent=msg;
  if(det && detail) det.textContent=detail;
}
function resetImportProgress(){
  setImportProgress(0,'Esperando archivo para importar','Selecciona un XLSX/CSV y presiona Validar y previsualizar.');
}
function limpiarImportacion(clearFile=true){
  state.importados=[];
  state.importHeaders=[];
  if(clearFile && $('fileImportar')) $('fileImportar').value='';
  setImportStats(0,0,0,0,0,0,0);
  resetImportProgress();
  if($('tbodyImportar')) $('tbodyImportar').innerHTML='<tr><td colspan="12" style="text-align:center;color:#64748b;padding:20px">Selecciona un archivo y presiona Previsualizar.</td></tr>';
}
async function procesarImportacion(){
  const file=$('fileImportar')?.files?.[0];
  if(!file) return toast('Selecciona un archivo Excel o CSV');
  try{
    setImportProgress(3,'Preparando lectura','Archivo: '+file.name+' | Tamaño: '+Math.round(file.size/1024)+' KB');
    setStatus('Leyendo archivo y validando contra PEDIDOS...');
    const rows=await leerArchivoListado(file, pct=>{
      setImportProgress(5+Math.round(pct*0.35),'Leyendo archivo XLSX/CSV','Procesando archivo: '+pct+'% leído');
    });
    setImportProgress(42,'Archivo leído','Filas detectadas en el archivo: '+Math.max(0,rows.length-1));
    if(!rows.length){
      setImportProgress(0,'Archivo vacío','El archivo no contiene filas para importar.','error');
      return toast('El archivo no contiene filas');
    }
    const headers=normalizarHeadersImport(rows[0]);
    state.importHeaders=headers;
    const data=rows.slice(1).filter(r=>r.some(c=>String(c??'').trim()));
    setImportProgress(55,'Normalizando columnas','Columnas detectadas: '+headers.filter(Boolean).length+' | Filas útiles: '+data.length);
    const ix=mapImportHeaders(headers);
    const fechaDefault=$('impFecha')?.value||'';
    const statusDefault=$('impStatus')?.value||'PENDIENTE';
    const pikeadorDefault=$('impPikeador')?.value||'';
    const items=data.map((r,i)=>{
      if(i%50===0 || i===data.length-1){
        const pct=55+Math.round(((i+1)/Math.max(1,data.length))*20);
        setImportProgress(pct,'Validando filas del archivo',`Fila ${i+1} de ${data.length}`);
      }
      return normalizarFilaImport(r,headers,ix,i,fechaDefault,statusDefault,pikeadorDefault);
    }).filter(x=>x._hasData);
    setImportProgress(78,'Comparando con PEDIDOS','Verificando nuevos, cambios y duplicados contra la tabla actual.');
    state.importados=validarImportadosContraPedidos(items);
    renderImportPreview();
    const enviables=(state.importados||[]).filter(x=>!x._error && x._hasData && x._estadoImport!=='SIN CAMBIO').length;
    setImportProgress(100,'Validación terminada',`Nuevos: ${state.importStats.nuevos} | Cambios: ${state.importStats.cambios} | Sin cambio: ${state.importStats.iguales} | Errores: ${state.importStats.errores} | Listos para enviar: ${enviables}`, state.importStats.errores?'':'ok');
    toast('Listado validado: '+items.length+' filas');
    setStatus('Importación validada. Nuevos: '+state.importStats.nuevos+' | Cambios: '+state.importStats.cambios+' | Sin cambio: '+state.importStats.iguales+' | Errores: '+state.importStats.errores);
  }catch(err){
    console.error(err);
    setImportProgress(0,'Error al validar archivo',err.message||String(err),'error');
    toast('Error al validar archivo: '+(err.message||err));
  }
}
function leerArchivoListado(file,onProgress){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onerror=()=>reject(new Error('No se pudo leer el archivo'));
    reader.onprogress=e=>{
      if(e.lengthComputable && typeof onProgress==='function'){
        onProgress(Math.round((e.loaded/e.total)*100));
      }
    };
    reader.onload=e=>{
      try{
        if(typeof onProgress==='function') onProgress(100);
        const name=file.name.toLowerCase();
        if(name.endsWith('.csv')){
          const text=String(e.target.result||'');
          const wb=XLSX.read(text,{type:'string'});
          const sheet=wb.Sheets[wb.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json(sheet,{header:1,defval:''}));
        }else{
          const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
          const sheet=wb.Sheets[wb.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json(sheet,{header:1,defval:''}));
        }
      }catch(err){reject(err)}
    };
    if(file.name.toLowerCase().endsWith('.csv')) reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
  });
}

function normalizarHeadersImport(row){
  const usados={};
  return (row||[]).map((h,i)=>{
    let base=String(h??'').trim();
    if(!base) base='COLUMNA_'+(i+1);
    let name=base, n=2;
    while(usados[norm(name)]){ name=base+'_'+n; n++; }
    usados[norm(name)]=true;
    return name;
  });
}
function rawObjectFromRow(headers,row){
  const obj={};
  (headers||[]).forEach((h,i)=>{
    if(!h) return;
    const val=row && row[i]!==undefined && row[i]!==null ? String(row[i]).trim() : '';
    obj[h]=val;
  });
  return obj;
}
function pickRaw(raw,names){
  if(!raw) return '';
  const keys=Object.keys(raw);
  for(const name of names){
    const nk=norm(name);
    const k=keys.find(x=>norm(x)===nk);
    if(k && String(raw[k]??'').trim()!=='') return String(raw[k]).trim();
  }
  return '';
}

function idxHeader(headers, names){
  const normalized = headers.map(norm);
  for(const name of names){
    const i = normalized.indexOf(norm(name));
    if(i !== -1) return i;
  }
  return -1;
}
function valorHeader(headers, idx){
  return idx >= 0 && idx < headers.length ? norm(headers[idx]) : '';
}
function estadoSeguroImport(v, defecto){
  const def = normalizarEstadoPermitido(defecto || 'PENDIENTE') || 'PENDIENTE';
  return normalizarEstadoPermitido(v) || def;
}
function normalizarEstadoPermitido(v){
  const e = normTxt(v).replace(/\s+/g,'');
  const map = {
    'PENDIENTE':'PENDIENTE',
    'PREPARACION':'PREPARACION',
    'PREPARACIÓN':'PREPARACION',
    'ENPREPARACION':'PREPARACION',
    'ENPREPARACIÓN':'PREPARACION',
    'RECIBIDO':'RECIBIDO',
    'RECEPCIONADO':'RECEPCIONADO',
    'DESPACHADO':'DESPACHADO',
    'ENRUTA':'EN RUTA',
    'ENTREGADO':'ENTREGADO',
    'TERMINADO':'TERMINADO',
    'CANCELADO':'CANCELADO',
    'ANULADO':'CANCELADO'
  };
  return map[e] || '';
}
function mapImportHeaders(headers){
  // Mapeo estricto por encabezado para evitar que DESCRIPCIÓN caiga en ESTADO.
  // Sólo se usa posición fija para campos base si el archivo no trae encabezados reconocibles.
  const ix = {
    fecha:idxHeader(headers,['fecha','fecha ingreso','fecha_pedido','fecha pedido','fecha creacion','fecha creación']),
    pedido:idxHeader(headers,['pedido','nro pedido','numero pedido','número pedido','numero','número','n°','nº','nro','no','orden','orden pedido','orden de pedido']),
    cliente:idxHeader(headers,['cliente','nombre cliente','razon social','razón social']),
    vendedor:idxHeader(headers,['vendedor','responsable','ejecutivo']),
    pikeador:idxHeader(headers,['pikeador','picker','preparador','asignado','asignado a']),
    codigo:idxHeader(headers,['codigo','código','cod','sku','codigo producto','código producto']),
    descripcion:idxHeader(headers,['descripcion','descripción','detalle','nombre','producto','nombre producto','descripcion producto','descripción producto']),
    ubicacion:idxHeader(headers,['ubicacion','ubicación','bodega','ubicacion producto','ubicación producto']),
    cantidad:idxHeader(headers,['cantidad','unidades','cajas','qty','cant']),
    status:idxHeader(headers,['status','estado','estado pedido','estado del pedido']),
    observacion:idxHeader(headers,['observacion','observación','obs','comentario','comentarios'])
  };

  const coreHits = ['pedido','codigo','descripcion','cliente'].filter(k => ix[k] >= 0).length;
  if(coreHits < 2){
    // Compatibilidad con plantillas sin encabezado: se aplican posiciones base,
    // pero ESTADO sigue sin posición por defecto para no copiar DESCRIPCIÓN.
    return {
      fecha:0,pedido:1,cliente:2,vendedor:3,pikeador:-1,codigo:5,descripcion:6,ubicacion:7,cantidad:8,status:-1,observacion:10
    };
  }
  return ix;
}

function obtenerPedidoDesdeImport(r,headers,ix,raw){
  // Regla solicitada: si el archivo trae columna NÚMERO / NUMERO / Nº / N°, ese valor ES el PEDIDO.
  // Se respeta PEDIDO si existe, pero NÚMERO funciona como respaldo directo para que no se pierda el pedido.
  const desdePedido = pick(r,ix.pedido,'');
  const desdeNumero = pickRaw(raw,['numero','número','NUMERO','NÚMERO','n°','N°','nº','Nº','no','nro','nro pedido','numero pedido','número pedido']);
  return desdePedido || desdeNumero;
}

function normalizarFilaImport(r,headers,ix,i,fechaDefault,statusDefault,pikeadorDefault){
  const raw = rawObjectFromRow(headers,r);
  const rawStatus = pick(r,ix.status,'');
  const statusFinal = estadoSeguroImport(rawStatus, statusDefault || 'PENDIENTE');
  const clienteImport = pick(r,ix.cliente,'');
  const pikeadorDesdeArchivo = ix.pikeador >= 0 ? pick(r,ix.pikeador,'') : '';
  const pikeadorFinal = limpiarPikeador(pikeadorDesdeArchivo || pikeadorDefault, clienteImport);
  const item={
    fecha:pick(r,ix.fecha,fechaDefault)||pickRaw(raw,['fecha','fecha ingreso','fecha pedido'])||fechaDefault,
    pedido:obtenerPedidoDesdeImport(r,headers,ix,raw)||pickRaw(raw,['pedido','orden','orden pedido','orden de pedido']),
    cliente:clienteImport,
    vendedor:pick(r,ix.vendedor,''),
    pikeador:pikeadorFinal,
    codigo:pick(r,ix.codigo,''),
    descripcion:pick(r,ix.descripcion,''),
    ubicacion:pick(r,ix.ubicacion,''),
    cantidad:Number(String(pick(r,ix.cantidad,1)).replace(',','.'))||1,
    status:statusFinal,
    observacion:pick(r,ix.observacion,''),
    __raw:raw,
    __rawHeaders:headers,
    _row:i+2
  };
  item._hasData=Boolean(item.pedido||item.codigo||item.descripcion||item.cliente);
  item._error=!item.pedido?'Falta pedido':'';
  if(rawStatus && !normalizarEstadoPermitido(rawStatus)){
    item._statusAdvertencia = 'Estado inválido ignorado: '+rawStatus;
  }
  return item;
}
function validarImportadosContraPedidos(items){
  const porKey=new Map();
  const porPedidoCodigo=new Map();
  (state.pedidos||[]).forEach(p=>{
    (p.productos||[]).forEach(prod=>{
      const base={fecha:p.fecha||'',pedido:p.pedido||'',cliente:p.cliente||'',vendedor:p.vendedor||'',pikeador:p.pikeador||'',codigo:prod.codigo||'',descripcion:prod.descripcion||'',ubicacion:prod.ubicacion||'',cantidad:Number(prod.cantidad||0),status:p.status||'PENDIENTE'};
      porKey.set(importKey(base),base);
      const pc=pedidoCodigoKey(base);
      if(pc&&!porPedidoCodigo.has(pc)) porPedidoCodigo.set(pc,base);
    });
  });
  const vistosArchivo=new Set();
  const stats={nuevos:0,cambios:0,iguales:0,errores:0};
  const out=items.map(item=>{
    item._changes=[];
    item._estadoImport='NUEVO';
    item._exists=false;
    if(!item.pedido){item._error='Falta pedido'; item._estadoImport='ERROR'; stats.errores++; return item;}
    if(!item.codigo&&!item.descripcion){item._error='Falta código o descripción'; item._estadoImport='ERROR'; stats.errores++; return item;}
    const k=importKey(item);
    if(vistosArchivo.has(k)){ item._error='Duplicado dentro del archivo'; item._estadoImport='ERROR'; stats.errores++; return item; }
    vistosArchivo.add(k);
    const actual=porKey.get(k)||porPedidoCodigo.get(pedidoCodigoKey(item));
    if(!actual){ stats.nuevos++; item._estadoImport='NUEVO'; return item; }
    item._exists=true;
    const campos=['fecha','cliente','vendedor','pikeador','descripcion','ubicacion','cantidad','status'];
    campos.forEach(c=>{
      const a=c==='cantidad'?Number(actual[c]||0):String(actual[c]||'').trim().toUpperCase();
      const b=c==='cantidad'?Number(item[c]||0):String(item[c]||'').trim().toUpperCase();
      if(String(a)!==String(b)) item._changes.push(c+': '+(actual[c]||'-')+' → '+(item[c]||'-'));
    });
    if(item._changes.length){ stats.cambios++; item._estadoImport='CAMBIO'; }
    else { stats.iguales++; item._estadoImport='SIN CAMBIO'; }
    return item;
  });
  state.importStats=stats;
  return out;
}
function importKey(x){ return [normPedido(x.pedido),normCodigo(x.codigo),normTxt(x.descripcion),normTxt(x.ubicacion)].join('|'); }
function pedidoCodigoKey(x){ const p=normPedido(x.pedido), c=normCodigo(x.codigo); return p&&c?p+'|'+c:''; }
function normPedido(v){return String(v||'').replace(/^'+/,'').replace(/\s+/g,'').toUpperCase();}
function normCodigo(v){return String(v||'').replace(/^'+/,'').replace(/\s+/g,'').toUpperCase();}
function normTxt(v){return String(v||'').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');}
function estadoImportBadge(x){
  const e=String(x._estadoImport||'NUEVO').toUpperCase();
  if(e==='ERROR') return '<span class="estado-import estado-error">Error</span>';
  if(e==='CAMBIO') return '<span class="estado-import estado-cambio">Cambio</span>';
  if(e==='SIN CAMBIO') return '<span class="estado-import estado-igual">Sin cambio</span>';
  return '<span class="estado-import estado-nuevo">Nuevo</span>';
}
function renderImportPreview(){
  const items=state.importados||[];
  const pedidos=new Set(items.map(x=>x.pedido).filter(Boolean));
  const unidades=items.reduce((s,x)=>s+(Number(x.cantidad)||0),0);
  const errores=items.filter(x=>x._error).length;
  setImportStats(pedidos.size,items.length,unidades,errores,state.importStats.nuevos||0,state.importStats.cambios||0,state.importStats.iguales||0);

  const tbody=$('tbodyImportar');
  if(!tbody) return;
  const table=tbody.closest('table');
  const thead=table?.querySelector('thead');
  const rawHeaders=(state.importHeaders||[]).filter(Boolean);
  const view=items.slice(0,200);
  const baseHeads=['#','Validación','Cambios detectados'];
  const heads=[...baseHeads,...rawHeaders];
  if(thead){
    thead.innerHTML='<tr>'+heads.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr>';
  }
  if(!view.length){
    tbody.innerHTML=`<tr><td colspan="${Math.max(3,heads.length)}" style="text-align:center;color:#64748b;padding:20px">No hay datos válidos para importar.</td></tr>`;
    return;
  }
  tbody.innerHTML=view.map((x,i)=>{
    const raw=x.__raw||{};
    const cambios=esc(x._error||(([...(x._statusAdvertencia?[x._statusAdvertencia]:[]), ...(x._changes||[])]).join(' | ')||'—'));
    const rawCells=rawHeaders.map(h=>`<td>${esc(raw[h]??'')}</td>`).join('');
    return `<tr style="${x._error?'background:#fee2e2':''}"><td>${i+1}</td><td>${estadoImportBadge(x)}</td><td class="changes-list">${cambios}</td>${rawCells}</tr>`;
  }).join('');
}
function setImportStats(pedidos,filas,unidades,errores,nuevos=0,cambios=0,iguales=0){
  if($('impPedidos')) $('impPedidos').textContent=pedidos;
  if($('impFilas')) $('impFilas').textContent=filas;
  if($('impUnidades')) $('impUnidades').textContent=unidades;
  if($('impErrores')) $('impErrores').textContent=errores;
  if($('impNuevos')) $('impNuevos').textContent=nuevos;
  if($('impCambios')) $('impCambios').textContent=cambios;
  if($('impIguales')) $('impIguales').textContent=iguales;
}
async function enviarImportacionBD(){
  const items=(state.importados||[]).filter(x=>!x._error && x._hasData && x._estadoImport!=='SIN CAMBIO');
  if(!items.length){
    setImportProgress(100,'Sin cambios para enviar','Todo está igual o existen errores pendientes de corregir.','ok');
    return toast('No hay cambios para enviar. Todo está igual o tiene errores.');
  }
  if(!confirm(`Enviar ${items.length} filas con cambios/nuevos a la Base de Datos PEDIDOS?`)) return;
  let enviados=0, fallidos=0, insertados=0, actualizados=0, sinCambios=0;
  const chunkSize=10;
  const totalChunks=Math.ceil(items.length/chunkSize);
  setImportProgress(0,'Iniciando envío a Base de Datos',`Se enviarán ${items.length} filas en ${totalChunks} bloque(s).`);
  for(let i=0,chunkIndex=0;i<items.length;i+=chunkSize,chunkIndex++){
    const batch=items.slice(i,i+chunkSize).map(({_row,_hasData,_error,_estadoImport,_exists,_changes,_statusAdvertencia,...x})=>x);
    const desde=i+1;
    const hasta=Math.min(i+chunkSize,items.length);
    const pctInicio=Math.round((i/items.length)*100);
    setImportProgress(pctInicio,'Enviando a Base de Datos',`Bloque ${chunkIndex+1} de ${totalChunks} | Filas ${desde}-${hasta} de ${items.length}`);
    setStatus(`Importando filas ${desde}-${hasta} de ${items.length}...`);
    try{
      const r=await api('importar_pedidos',{items:JSON.stringify(batch)});
      if(r?.ok){
        enviados+=batch.length;
        insertados+=Number(r.insertados||0);
        actualizados+=Number(r.actualizados||0);
        sinCambios+=Number(r.sin_cambios||0);
      }else{
        fallidos+=batch.length;
      }
    }catch(err){
      console.error(err);
      fallidos+=batch.length;
    }
    const pctFin=Math.round((Math.min(i+chunkSize,items.length)/items.length)*100);
    setImportProgress(pctFin,'Enviando a Base de Datos',`Avance ${pctFin}% | Enviadas: ${enviados} | Insertados: ${insertados} | Actualizados: ${actualizados} | Fallidas: ${fallidos}`);
  }
  if(fallidos){
    setImportProgress(100,'Importación terminada con observaciones',`Enviadas: ${enviados}. Insertados: ${insertados}. Actualizados: ${actualizados}. Sin cambio: ${sinCambios}. Fallidas: ${fallidos}.`,'error');
  }else{
    setImportProgress(100,'Importación completada',`Enviadas: ${enviados}. Insertados: ${insertados}. Actualizados: ${actualizados}. Sin cambio: ${sinCambios}.`,'ok');
  }
  toast(`Importación terminada. Enviadas: ${enviados}. Insertados: ${insertados}. Actualizados: ${actualizados}. Sin cambio: ${sinCambios}. Fallidas: ${fallidos}.`);
  await cargarTodo();
  if(!fallidos) setTimeout(()=>cerrarImportar(),900);
}

/* ================= PDFS MEJORADOS ================= */

function abrirDialogoImpresionPdf(doc,nombreArchivo){
  try{
    if(doc && typeof doc.setProperties==='function'){
      doc.setProperties({title:String(nombreArchivo||'documento').replace(/\.pdf$/i,'')});
    }
    if(doc && typeof doc.autoPrint==='function') doc.autoPrint();
    const blob=doc.output('blob');
    const url=URL.createObjectURL(blob);
    const iframe=document.createElement('iframe');
    iframe.style.position='fixed';
    iframe.style.right='0';
    iframe.style.bottom='0';
    iframe.style.width='0';
    iframe.style.height='0';
    iframe.style.border='0';
    iframe.onload=()=>{
      setTimeout(()=>{
        try{
          iframe.contentWindow.focus();
          iframe.contentWindow.print();
        }catch(e){
          window.open(url,'_blank','noopener');
        }
      },350);
    };
    iframe.src=url;
    document.body.appendChild(iframe);
    setTimeout(()=>{
      try{document.body.removeChild(iframe);}catch(e){}
      try{URL.revokeObjectURL(url);}catch(e){}
    },120000);
    toast('Abriendo cuadro de impresión');
  }catch(err){
    console.error(err);
    try{ doc.save(nombreArchivo || 'documento.pdf'); }
    catch(e){ toast('No se pudo abrir impresión'); }
  }
}

function pdfGeneral(){
  const {jsPDF}=window.jspdf||{}; if(!jsPDF)return toast('No cargó jsPDF');
  const d=new jsPDF('l','mm','a4');
  const totalPedidos=state.filtrados.length;
  const totalProductos=state.filtrados.reduce((s,p)=>s+p.total_productos,0);
  const totalUnidades=state.filtrados.reduce((s,p)=>s+p.total_unidades,0);
  d.setFont('helvetica','bold'); d.setFontSize(15); d.text('Listado de Orden de Pedidos',14,13);
  d.setFont('helvetica','normal'); d.setFontSize(8); d.text('Generado: '+new Date().toLocaleString('es-CL'),14,18);
  d.setDrawColor(220); d.roundedRect(14,22,268,14,2,2);
  d.setFont('helvetica','bold'); d.text('Pedidos: '+totalPedidos,18,31); d.text('Productos: '+totalProductos,70,31); d.text('TOTAL UNIDADES: '+totalUnidades,125,31); d.text('Filtrados por pantalla',205,31);
  d.autoTable({
    startY:42,
    head:[['Fecha','Pedido','Cliente','Vendedor','Pikeador','Productos','Unidades','Estado']],
    body:state.filtrados.map(p=>[p.fecha,p.pedido,p.cliente,p.vendedor,pikeadorVisible(p.pikeador),p.total_productos,p.total_unidades,p.status]),
    theme:'grid',
    styles:{fontSize:7.5,cellPadding:1.8,overflow:'linebreak',valign:'middle'},
    headStyles:{fillColor:[15,118,110],textColor:255,fontStyle:'bold'},
    columnStyles:{0:{cellWidth:28},1:{cellWidth:24},2:{cellWidth:55},3:{cellWidth:38},4:{cellWidth:38},5:{cellWidth:22,halign:'center'},6:{cellWidth:24,halign:'center',fontStyle:'bold'},7:{cellWidth:28}},
    margin:{left:14,right:14},
    didDrawPage:data=>{d.setFontSize(7);d.text('Página '+d.internal.getNumberOfPages(),270,202);}
  });
  const y=(d.lastAutoTable?.finalY||42)+8;
  d.setFont('helvetica','bold'); d.setFontSize(10); d.text('TOTAL DE UNIDADES: '+totalUnidades,14,Math.min(y,198));
  abrirDialogoImpresionPdf(d,'orden_pedidos_A4.pdf');
}
function pdfPedidoA4(){
  const p=state.sel; if(!p)return;
  const {jsPDF}=window.jspdf||{}; if(!jsPDF)return toast('No cargó jsPDF');
  const d=new jsPDF('p','mm','a4');
  d.setFont('helvetica','bold'); d.setFontSize(16); d.text('Orden de Pedido',14,14);
  d.setFont('helvetica','normal'); d.setFontSize(9); d.text('Generado: '+new Date().toLocaleString('es-CL'),14,20);
  d.setDrawColor(210); d.roundedRect(14,25,182,34,2,2);
  d.setFontSize(9);
  infoLine(d,18,33,'Pedido',p.pedido); infoLine(d,75,33,'Fecha',p.fecha||'-'); infoLine(d,132,33,'Estado',p.status||'-');
  infoLine(d,18,42,'Cliente',p.cliente||'-'); infoLine(d,105,42,'Vendedor',p.vendedor||'-');
  infoLine(d,18,51,'Pikeador',pikeadorVisible(p.pikeador)); infoLine(d,105,51,'Total unidades',String(p.total_unidades||0));
  d.autoTable({
    startY:66,
    head:[['Código','Descripción','Ubicación','Cantidad']],
    body:(p.productos.length?p.productos:[{codigo:'-',descripcion:'Sin productos detallados',ubicacion:'-',cantidad:0}]).map(x=>[x.codigo,x.descripcion,x.ubicacion,x.cantidad]),
    theme:'grid',
    styles:{fontSize:8,cellPadding:2,overflow:'linebreak',valign:'middle'},
    headStyles:{fillColor:[15,118,110],textColor:255,fontStyle:'bold'},
    columnStyles:{0:{cellWidth:32},1:{cellWidth:82},2:{cellWidth:42},3:{cellWidth:24,halign:'center',fontStyle:'bold'}},
    margin:{left:14,right:14}
  });
  const y=(d.lastAutoTable?.finalY||70)+8;
  d.setFont('helvetica','bold'); d.setFontSize(11); d.text('TOTAL DE UNIDADES: '+(p.total_unidades||0),14,Math.min(y,284));
  abrirDialogoImpresionPdf(d,'pedido_'+safe(p.pedido)+'_A4.pdf');
}
function infoLine(d,x,y,label,value){
  d.setFont('helvetica','bold'); d.text(label+':',x,y);
  d.setFont('helvetica','normal'); d.text(short(String(value??'-'),32),x+22,y);
}
function pdfPedido80(){
  const p=state.sel; if(!p)return;
  const {jsPDF}=window.jspdf||{}; if(!jsPDF)return toast('No cargó jsPDF');
  const filas=Math.max(1,p.productos.length);
  const alto=Math.max(140,70+(filas*12));
  const d=new jsPDF({orientation:'p',unit:'mm',format:[80,alto]});
  let y=6;
  d.setFont('helvetica','bold'); d.setFontSize(11); d.text('ORDEN DE PEDIDO',40,y,{align:'center'}); y+=6;
  d.setFont('helvetica','normal'); d.setFontSize(7); d.text(new Date().toLocaleString('es-CL'),40,y,{align:'center'}); y+=5;
  line80(d,y); y+=4;
  d.setFontSize(8);
  y=ticketText(d,'Pedido',p.pedido,y); y=ticketText(d,'Cliente',p.cliente||'-',y); y=ticketText(d,'Pikeador',pikeadorVisible(p.pikeador),y); y=ticketText(d,'Estado',p.status||'-',y); y=ticketText(d,'Total unidades',String(p.total_unidades||0),y); y+=2;
  line80(d,y); y+=4;
  d.setFont('helvetica','bold'); d.text('DETALLE PRODUCTOS',4,y); y+=5;
  d.setFont('helvetica','normal');
  const productos=p.productos.length?p.productos:[{codigo:'-',descripcion:'Sin productos detallados',ubicacion:'-',cantidad:0}];
  productos.forEach((x,i)=>{
    d.setFont('helvetica','bold'); d.text(`${i+1}. ${short(x.codigo||'-',18)}  Cant: ${x.cantidad}`,4,y); y+=4;
    d.setFont('helvetica','normal');
    const desc=d.splitTextToSize(String(x.descripcion||'-'),70);
    d.text(desc,4,y); y+=desc.length*3.5;
    if(x.ubicacion){d.text('Ubic: '+short(x.ubicacion,32),4,y); y+=4;}
    y+=2;
  });
  line80(d,y); y+=5;
  d.setFont('helvetica','bold'); d.setFontSize(10); d.text('TOTAL UNIDADES: '+(p.total_unidades||0),40,y,{align:'center'});
  abrirDialogoImpresionPdf(d,'pedido_'+safe(p.pedido)+'_80mm.pdf');
}
function ticketText(d,label,value,y){
  d.setFont('helvetica','bold'); d.text(label+':',4,y);
  d.setFont('helvetica','normal');
  const txt=d.splitTextToSize(String(value||'-'),50);
  d.text(txt,28,y);
  return y+(txt.length*3.8);
}
function line80(d,y){ d.setDrawColor(120); d.line(4,y,76,y); }


/* ================= ANALÍTICA / TIEMPOS ================= */
function estadoPedido(p){ return String(p?.status||'').toUpperCase().trim(); }
function esPedidoProcesado(p){ return ['TERMINADO','DESPACHADO','RECIBIDO'].includes(estadoPedido(p)); }
function fechaHoraValida(v){
  v=String(v||'').trim();
  if(!v) return null;
  let d=null;
  const m=v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if(m) d=new Date(Number(m[3]),Number(m[2])-1,Number(m[1]),Number(m[4]||0),Number(m[5]||0),Number(m[6]||0));
  else d=new Date(v.replace(' ','T'));
  return isNaN(d.getTime())?null:d;
}
function minutosPreparacionPedido(p){
  const directo=Number(String(p?.tiempo_preparacion_min||0).replace(',','.'))||0;
  if(directo>0) return directo;
  const ini=fechaHoraValida(p?.hora_inicio), fin=fechaHoraValida(p?.hora_termino);
  if(!ini||!fin) return 0;
  return Math.max(0,Math.round((fin.getTime()-ini.getTime())/60000));
}
function formatearMinutos(min){
  min=Math.round(Number(min)||0);
  if(min<=0) return '0 min';
  const h=Math.floor(min/60), m=min%60;
  return h?`${h}h ${String(m).padStart(2,'0')}m`:`${m} min`;
}
function tiempoPreparacionTexto(p){
  const min=minutosPreparacionPedido(p);
  if(min>0) return formatearMinutos(min);
  if(p?.hora_inicio && !p?.hora_termino) return 'En preparación';
  return '-';
}
function horaCorta(v){
  v=String(v||'').trim();
  if(!v) return '';
  return v.replace('T',' ').replace(/\.\d+Z?$/,'').slice(0,19);
}
function porcentajeAvancePedido(p){
  const orden=['PENDIENTE','PREPARACION','RECIBIDO','DESPACHADO','TERMINADO'];
  const st=estadoPedido(p);
  if(st==='CANCELADO') return 0;
  const ix=Math.max(0,orden.indexOf(st));
  return Math.round((ix/(orden.length-1))*100);
}
function renderRotacionYRendimiento(lista){
  const tbodyRot=$('tbodyRotacionProductos');
  const tbodyPik=$('tbodyRendimientoPikeador');
  if(tbodyRot){
    const map={};
    (lista||[]).forEach(p=>{
      (p.productos||[]).forEach(prod=>{
        const k=(normCodigo(prod.codigo)||norm(prod.descripcion)||'SIN_CODIGO');
        if(!map[k]) map[k]={codigo:prod.codigo||'-',descripcion:prod.descripcion||'-',unidades:0,pedidos:new Set()};
        map[k].unidades += Number(prod.cantidad||0)||0;
        map[k].pedidos.add(p.pedido);
      });
    });
    const top=Object.values(map).sort((a,b)=>b.unidades-a.unidades).slice(0,8);
    tbodyRot.innerHTML=top.length?top.map((x,i)=>`<tr><td>${i+1}</td><td><b>${esc(x.codigo)}</b></td><td>${esc(short(x.descripcion,42))}</td><td><b>${x.unidades}</b></td><td>${x.pedidos.size}</td></tr>`).join(''):'<tr><td colspan="5" style="text-align:center;color:#64748b">Sin productos para calcular rotación.</td></tr>';
  }
  if(tbodyPik){
    const map={};
    (lista||[]).forEach(p=>{
      const pk=pikeadorVisible(p.pikeador);
      if(pk==='Sin asignar') return;
      if(!map[pk]) map[pk]={pikeador:pk,pedidos:0,procesados:0,unidades:0,tiempos:[]};
      map[pk].pedidos++;
      map[pk].unidades+=Number(p.total_unidades||0)||0;
      if(esPedidoProcesado(p)) map[pk].procesados++;
      const min=minutosPreparacionPedido(p); if(min>0) map[pk].tiempos.push(min);
    });
    const rows=Object.values(map).sort((a,b)=>b.procesados-a.procesados || b.unidades-a.unidades).slice(0,8);
    tbodyPik.innerHTML=rows.length?rows.map(x=>`<tr><td><b>${esc(x.pikeador)}</b></td><td>${x.pedidos}</td><td>${x.procesados}</td><td>${x.unidades}</td><td><b>${x.tiempos.length?formatearMinutos(x.tiempos.reduce((s,t)=>s+t,0)/x.tiempos.length):'-'}</b></td></tr>`).join(''):'<tr><td colspan="5" style="text-align:center;color:#64748b">Sin pikeadores asignados.</td></tr>';
  }
}

/* ================= UTILIDADES ================= */
function togglePanel(){
  const b=[$('panelControl'),$('panelKpis'),$('panelAnalitica'),$('panelTabla')].filter(Boolean);
  const h=b[0].classList.toggle('hidden');
  b.slice(1).forEach(x=>x.classList.toggle('hidden',h));
  $('btnTogglePanel').textContent=h?'Mostrar panel y tabla':'Ocultar panel';
}
function api(accion,params={}){
  return new Promise((resolve,reject)=>{
    const cb='cbOP_'+Date.now()+'_'+Math.floor(Math.random()*99999);
    const url=new URL(API_URL);
    url.searchParams.set('accion',accion);
    url.searchParams.set('callback',cb);
    url.searchParams.set('_',Date.now());
    Object.entries(params).forEach(([k,v])=>{if(v!=null)url.searchParams.set(k,String(v));});
    const s=document.createElement('script');
    const to=setTimeout(()=>done(new Error('Tiempo agotado consultando Web App')),30000);
    window[cb]=data=>done(null,data);
    function done(err,data){clearTimeout(to); try{delete window[cb]}catch(e){} s.remove(); err?reject(err):resolve(data);}
    s.onerror=()=>done(new Error('No se pudo conectar con Web App'));
    s.src=url.toString();
    document.body.appendChild(s);
  });
}
function errorTabla(m){ $('tbodyPedidos').innerHTML=`<tr><td colspan="13" style="padding:28px;text-align:center;color:#991b1b">${esc(m)}</td></tr>`; }
function badgeEstado(s){
  s=String(s||'PENDIENTE').toUpperCase(); let c='pendiente';
  if(s.includes('PREPAR'))c='preparacion'; else if(s.includes('TERMIN'))c='terminado'; else if(s.includes('CANCEL'))c='cancelado';
  return `<span class="badge ${c}">${esc(s)}</span>`;
}
function badgeAlerta(p){ return p.alerta_token?`<span class="badge alerta">${esc(p.alerta_tipo||'ALERTA')}</span>`:'<span style="color:#94a3b8">Sin alerta</span>'; }
function setStatus(t){ if($('statusLine')) $('statusLine').textContent=t; }
function toast(m){ const e=$('toast'); if(!e)return alert(m); e.textContent=m; e.classList.add('show'); setTimeout(()=>e.classList.remove('show'),4200); }
function mensajePedidoTerminado(p){
  p = p || {};
  const pedido = String(p.pedido || '').trim();
  const cliente = String(p.cliente || '').trim();
  return 'Pedido terminado' + (pedido ? ', pedido ' + pedido : '') + (cliente ? ', cliente ' + cliente : '');
}
function voz(t){ try{const u=new SpeechSynthesisUtterance(t);u.lang='es-CL';speechSynthesis.cancel();speechSynthesis.speak(u);}catch(e){} }
function pick(r,i,f){return (i == null || i < 0 || !Array.isArray(r) || i >= r.length || r[i] == null || String(r[i]).trim()==='') ? f : String(r[i]).trim();}
function norm(v){return String(v||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[\s_\-.]+/g,'')}
function numPed(v){const m=String(v||'').match(/\d+/g);return m?Number(m.join('')):0}
function esc(v){return String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function safe(v){return String(v||'pedido').replace(/[^a-z0-9_-]+/gi,'_')}
function short(v,n){v=String(v||''); return v.length>n?v.slice(0,n-1)+'…':v;}
