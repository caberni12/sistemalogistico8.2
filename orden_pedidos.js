const API_URL='https://script.google.com/macros/s/AKfycbxDPLaKDy5LqC9US-DQcCicPDIPb0XlxnPPA-y6N1AdDvbHZPxLzM0awD-NoFTcVk8Fkw/exec';

const state={
  pedidos:[],
  filtrados:[],
  pikeadores:[],
  maestra:[],
  maestraMap:{},
  ubicacionesMap:{},
  vendedores:[],
  bodegas:[],
  clientes:[],
  bodegaMapa:null,
  bodegaMarker:null,
  bodegaCircle:null,
  maestraListaCargada:false,
  sel:null,
  selectedKey:null,
  editandoPedido:false,
  importados:[],
  importFallidos:[],
  importStats:{nuevos:0,cambios:0,iguales:0,errores:0},
  importHeaders:[],
  manualProductos:[],
  tokens:JSON.parse(localStorage.getItem('orden_tokens')||'{}')
};

const $=id=>document.getElementById(id);

/* ================= FORMATO FECHAS =================
   Evita que fechas importadas desde Excel/Sheets aparezcan como serial numérico.
   Formato visible: dd-mm-yyyy o dd-mm-yyyy HH:MM.
================================================== */
function fechaVisiblePedido(valor){
  const v0 = valor;
  let v = String(valor == null ? '' : valor).trim();
  if(!v) return '';
  if(v.startsWith("'")) v = v.slice(1).trim();
  const n = Number(v.replace(',', '.'));
  if(Number.isFinite(n) && n > 20000 && n < 90000){
    const ms = Math.round(n * 86400000);
    const d = new Date(Date.UTC(1899, 11, 30) + ms);
    const dd = String(d.getUTCDate()).padStart(2,'0');
    const mm = String(d.getUTCMonth()+1).padStart(2,'0');
    const yy = d.getUTCFullYear();
    const frac = n - Math.floor(n);
    if(frac > 0.0007){
      const mins = Math.round(frac * 1440);
      const hh = String(Math.floor(mins/60)%24).padStart(2,'0');
      const mi = String(mins%60).padStart(2,'0');
      return `${dd}-${mm}-${yy} ${hh}:${mi}`;
    }
    return `${dd}-${mm}-${yy}`;
  }
  let m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::\d{2})?)?/);
  if(m){
    const dd = String(Number(m[3])).padStart(2,'0');
    const mm = String(Number(m[2])).padStart(2,'0');
    const yy = m[1];
    return m[4] ? `${dd}-${mm}-${yy} ${String(Number(m[4])).padStart(2,'0')}:${m[5]}` : `${dd}-${mm}-${yy}`;
  }
  m = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[T\s](\d{1,2}):(\d{2})(?::\d{2})?)?/);
  if(m){
    const dd = String(Number(m[1])).padStart(2,'0');
    const mm = String(Number(m[2])).padStart(2,'0');
    const yy = m[3];
    return m[4] ? `${dd}-${mm}-${yy} ${String(Number(m[4])).padStart(2,'0')}:${m[5]}` : `${dd}-${mm}-${yy}`;
  }
  return String(v0 == null ? '' : v0).trim();
}
function fechaHoraVisiblePedido(valor){ return fechaVisiblePedido(valor); }
function esHeaderFechaPedido(h){ return /fecha/i.test(String(h||'')); }


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
  inicializarPanelesOrden();
  bind();
  prepararCargaManualOrden();
});


function inicializarPanelesOrden(){
  document.body.classList.add('metricas-ocultas');
  [$('panelKpis'),$('panelAnalitica')].filter(Boolean).forEach(x=>x.classList.add('hidden'));
  const btn=$('btnToggleMetricas');
  if(btn) btn.textContent='Mostrar KPI / Rendimiento';
  const btnPanel=$('btnTogglePanel');
  if(btnPanel) btnPanel.textContent='Ocultar filtros';
}

function prepararCargaManualOrden(){
  setStatus('Carga automática desactivada. Presiona Sincronizar para consultar la base de datos.');
  const tb=$('tbodyPedidos');
  if(tb){
    tb.innerHTML='<tr><td colspan="13" style="text-align:center;padding:30px;color:#64748b">Carga automática desactivada. Presiona <b>Sincronizar</b> para mostrar pedidos.</td></tr>';
  }
  const mob=$('mobileList');
  if(mob){
    mob.innerHTML='<div class="pedido-card">Carga automática desactivada. Presiona Sincronizar.</div>';
  }
  limpiarMetricasOrdenSinCarga();
}

function limpiarMetricasOrdenSinCarga(){
  ['kPedidos','kProductos','kUnidades','kPendientes','kProcesados','kPromedioProductos','kPromedioUnidades','kUnidadesDia','kUnidadesMes','kUnidadesAnio'].forEach(id=>{
    const el=$(id);
    if(el) el.textContent='0';
  });
  const t=$('kTiempoPromedio');
  if(t) t.textContent='0 min';
  const rot=$('tbodyRotacionProductos');
  if(rot) rot.innerHTML='<tr><td colspan="5" style="text-align:center;color:#64748b">Presiona Sincronizar para calcular.</td></tr>';
  const pik=$('tbodyRendimientoPikeador');
  if(pik) pik.innerHTML='<tr><td colspan="5" style="text-align:center;color:#64748b">Presiona Sincronizar para calcular.</td></tr>';
  const ven=$('tbodyRendimientoVendedor');
  if(ven) ven.innerHTML='<tr><td colspan="5" style="text-align:center;color:#64748b">Presiona Sincronizar para calcular.</td></tr>';
  const cli=$('tbodyTopClientes');
  if(cli) cli.innerHTML='<tr><td colspan="5" style="text-align:center;color:#64748b">Presiona Sincronizar para calcular.</td></tr>';
}

function bind(){
  instalarEdicionPedidoUI();
  $('btnCargar')?.addEventListener('click',withLoaderEvent(()=>cargarTodo()));
  $('btnSync')?.addEventListener('click',withLoaderEvent(()=>cargarTodo()));
  $('btnSyncMaestra')?.addEventListener('click',withLoaderEvent(()=>cargarMaestra(false)));
  $('btnSyncUbicaciones')?.addEventListener('click',withLoaderEvent(()=>sincronizarUbicacionesManual(false)));
  $('btnPDF')?.addEventListener('click',withLoaderEvent(()=>pdfGeneral()));
  $('btnImportar')?.addEventListener('click',e=>{abrirImportar();stopClickedLoader(e);});
  $('btnAbrirClientes')?.addEventListener('click',e=>{abrirClientesModal();stopClickedLoader(e);});
  $('btnCerrarClientes')?.addEventListener('click',e=>{cerrarClientesModal();stopClickedLoader(e);});
  $('modalClientes')?.addEventListener('click',e=>{if(e.target.id==='modalClientes')cerrarClientesModal();});
  $('btnGuardarCliente')?.addEventListener('click',withLoaderEvent(()=>guardarCliente()));
  $('btnNuevoCliente')?.addEventListener('click',e=>{limpiarClienteForm();stopClickedLoader(e);});
  $('btnRecargarClientes')?.addEventListener('click',withLoaderEvent(()=>cargarClientes()));
  $('tbodyClientes')?.addEventListener('click',manejarClickClientes);
  $('btnManual')?.addEventListener('click',e=>{abrirManual();stopClickedLoader(e);});
  $('btnConfigOrden')?.addEventListener('click',e=>{abrirConfigOrden();stopClickedLoader(e);});
  $('btnCerrarConfigOrden')?.addEventListener('click',e=>{cerrarConfigOrden();stopClickedLoader(e);});
  $('modalConfigOrden')?.addEventListener('click',e=>{if(e.target.id==='modalConfigOrden')cerrarConfigOrden();});
  $('btnGenerarPedidoManual')?.addEventListener('click',withLoaderEvent(()=>generarNumeroPedidoManual(true)));
  $('txtBuscar')?.addEventListener('input',filtrar);
  $('selEstado')?.addEventListener('change',filtrar);
  $('selFiltroPikeador')?.addEventListener('change',filtrar);
  $('selFiltroVendedor')?.addEventListener('change',filtrar);
  ['selPeriodoKpi','fechaDesdeKpi','fechaHastaKpi'].forEach(id=>$(id)?.addEventListener('change',filtrar));
  $('tbodyPedidos')?.addEventListener('click',abrirModalDesdeTabla);
  $('tbodyPedidos')?.addEventListener('keydown',manejarTeclasFilaPedido);
  $('mobileList')?.addEventListener('click',abrirModalDesdeTarjetaMovil);
  document.addEventListener('keydown',manejarFlechasPedidos);
  $('btnAgregarPikeador')?.addEventListener('click',withLoaderEvent(()=>agregarPikeador()));
  $('btnAgregarVendedor')?.addEventListener('click',withLoaderEvent(()=>agregarVendedor()));
  $('btnBuscarBodegaMapa')?.addEventListener('click',withLoaderEvent(()=>buscarBodegaEnMapa()));
  $('btnGuardarBodega')?.addEventListener('click',withLoaderEvent(()=>agregarBodega()));
  $('txtBodegaDireccion')?.addEventListener('change',()=>buscarBodegaEnMapa(false));
  $('txtBodegaRadio')?.addEventListener('input',()=>actualizarRadioBodegaMapa());
  $('btnTogglePanel')?.addEventListener('click',e=>{togglePanel();stopClickedLoader(e);});
  $('btnToggleMetricas')?.addEventListener('click',e=>{toggleMetricas();stopClickedLoader(e);});
  $('btnCerrarModal')?.addEventListener('click',e=>{cerrarModal();stopClickedLoader(e);});
  $('modalPedido')?.addEventListener('click',e=>{if(e.target.id==='modalPedido')cerrarModal();});
  $('btnEditarPedido')?.addEventListener('click',e=>{activarEdicionPedido(true);stopClickedLoader(e);});
  $('btnCancelarEdicionPedido')?.addEventListener('click',e=>{activarEdicionPedido(false);stopClickedLoader(e);});
  $('btnGuardarEdicionPedido')?.addEventListener('click',withLoaderEvent(()=>guardarEdicionPedidoActual()));
  $('editPedidoBox')?.addEventListener('click',manejarClickEdicionPedido);
  $('editPedidoBox')?.addEventListener('change',manejarCambioEdicionPedido);
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
  ['manFecha','manPedido','manCliente','manDescripcion','manUbicacion','manCantidad'].forEach(id=>$(id)?.addEventListener('input',renderManualPreview));
  $('manCliente')?.addEventListener('change',aplicarClienteManual);
  ['manVendedor','manPikeador','manBodega','manStatus'].forEach(id=>$(id)?.addEventListener('change',renderManualPreview));
  $('manCodigo')?.addEventListener('input',()=>{ renderManualPreview(); debounceBuscarCodigoManual(); });
  $('manCodigo')?.addEventListener('change',()=>buscarProductoManualPorCodigo(false));
  $('manCodigo')?.addEventListener('blur',()=>buscarProductoManualPorCodigo(false));


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
  setStatus('Sincronizando con base de datos...');
  try{
    const ubicacionesPromise=cargarUbicacionesPedido().catch(err=>{
      console.warn('No se pudieron cargar ubicaciones múltiples', err);
      return state.ubicacionesMap || {};
    });
    let r=await api('listar_pedidos',{mostrarTodo:1});
    if(!r?.ok) throw new Error(r?.msg||'Respuesta inválida');
    state.pedidos=normPedidos(r);
    llenarSelectEstadosOrden();
    if(!state.pedidos.length){
      const b=await api('listar_bd',{sheet:'PEDIDOS'}).catch(()=>null);
      if(b?.ok) state.pedidos=normPedidos(b);
      llenarSelectEstadosOrden();
    }
    state.ubicacionesMap=await ubicacionesPromise;
    aplicarUbicacionesAPedidos();
    llenarSelectEstadosOrden();
    filtrar();
    setStatus('Sincronización con base de datos completada. Pedidos: '+state.pedidos.length+' | Productos con ubicaciones: '+Object.keys(state.ubicacionesMap||{}).length);
    marcarTokens();
  }catch(e){
    console.error(e);
    setStatus('Error sincronizando con base de datos: '+e.message);
    errorTabla('No se pudo sincronizar con base de datos. Verifica doGet/listar_pedidos en Apps Script.');
  }
  cargarPikeadores().catch(console.warn);
  cargarVendedores().catch(console.warn);
  cargarBodegas().catch(console.warn);
  cargarClientes().catch(console.warn);
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
  llenarSelectVendedores();
  llenarSelectBodegas();
}
async function cargarVendedores(){
  const r=await api('listar_vendedores',{});
  let l=[];
  if(r?.ok) l=normVendedores(r);
  const desdePedidos=[...new Set((state.pedidos||[]).map(p=>String(p.vendedor||'').trim()).filter(Boolean))];
  state.vendedores=[...new Set([...l,...desdePedidos].map(x=>String(x||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  llenarSelectVendedores();
}
function normVendedores(r){
  const a=r.vendedores||r.data||r.rows||r.values||[];
  return Array.isArray(a)?a.map(x=>typeof x==='string'?x:Array.isArray(x)?x[0]:(x.nombre||x.vendedor||x.ejecutivo||x.responsable||'')).filter(Boolean):[];
}
function llenarSelectVendedores(){
  [$('selFiltroVendedor'),$('manVendedor'),$('editVendedorPedido'),$('cliVendedor')].filter(Boolean).forEach(s=>{
    const v=s.value;
    const first=s.id==='selFiltroVendedor'?'<option value="">Todos</option>':'<option value="">Sin vendedor</option>';
    s.innerHTML=first+state.vendedores.map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join('');
    if(v){
      if(!state.vendedores.some(x=>String(x)===String(v))) s.insertAdjacentHTML('beforeend',`<option value="${esc(v)}">${esc(v)}</option>`);
      s.value=v;
    }
  });
}

function llenarSelectEstadosOrden(){
  const s=$('selEstado');
  if(!s) return;
  const actual=s.value;
  const base=['PENDIENTE','PREPARACION','EN RUTA','RECIBIDO','DESPACHADO','ENTREGADO','TERMINADO','CANCELADO'];
  const desdePedidos=[...new Set((state.pedidos||[]).map(p=>String(p.status||p.estado||'').trim().toUpperCase()).filter(Boolean))];
  const estados=[...new Set([...base,...desdePedidos])].sort((a,b)=>{
    const ia=base.indexOf(a), ib=base.indexOf(b);
    if(ia>=0 && ib>=0) return ia-ib;
    if(ia>=0) return -1;
    if(ib>=0) return 1;
    return a.localeCompare(b);
  });
  s.innerHTML='<option value="">Todos</option>'+estados.map(e=>`<option value="${esc(e)}">${esc(e)}</option>`).join('');
  if(actual){
    if(!estados.includes(actual)) s.insertAdjacentHTML('beforeend',`<option value="${esc(actual)}">${esc(actual)}</option>`);
    s.value=actual;
  }
}

function opcionesVendedores(actual=''){
  const lista=[...new Set([...(state.vendedores||[]), actual].map(x=>String(x||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  return '<option value="">Sin vendedor</option>'+lista.map(n=>`<option value="${esc(n)}" ${String(n)===String(actual||'')?'selected':''}>${esc(n)}</option>`).join('');
}
async function agregarVendedor(){
  const nombre=($('txtNuevoVendedor')?.value||'').trim();
  if(!nombre)return toast('Ingresa el nombre del vendedor');
  const r=await api('agregar_vendedor',{nombre});
  if(!r?.ok)return toast(r?.msg||'No se pudo agregar vendedor');
  $('txtNuevoVendedor').value='';
  await cargarVendedores();
  toast('Vendedor guardado');
}


/* ================= CLIENTES ================= */
async function cargarClientes(){
  const r=await api('listar_clientes',{});
  if(r?.ok){
    state.clientes=normClientes(r).sort((a,b)=>String(a.cliente||'').localeCompare(String(b.cliente||'')));
    llenarDatalistClientes();
    renderClientesTabla();
  }
  return state.clientes;
}
function normClientes(r){
  const arr=r.clientes||r.items||r.data||r.rows||r.values||[];
  if(!Array.isArray(arr)) return [];
  if(arr.length && Array.isArray(arr[0])){
    const headers=(r.headers||[]).map(norm);
    const idx=(names)=>{for(const n of names){const i=headers.indexOf(norm(n)); if(i>=0)return i;} return -1;};
    const ix={id:idx(['id']),cliente:idx(['cliente','nombre','razon_social','razón social']),rut:idx(['rut','r.u.t']),vendedor:idx(['vendedor','ejecutivo','asesor','responsable_venta','responsable venta']),clase_cliente:idx(['clase_cliente','clase cliente','tipo_cliente','tipo cliente','categoria_cliente','categoría cliente','categoria']),direccion:idx(['direccion','dirección']),giro:idx(['giro']),medio_pago:idx(['medio_pago','medio pago','forma_pago','forma pago']),telefono:idx(['telefono','teléfono','fono']),correo:idx(['correo_electronico','correo electrónico','correo','email','mail']),responsable:idx(['responsable_empresa','responsable empresa','responsable']),fecha:idx(['fecha_creacion','fecha creación','fecha']),status:idx(['status','estado'])};
    return arr.map((x,i)=>({rowNumber:(r.rowNumbers&&r.rowNumbers[i])||i+2,id:String(x[ix.id]||'').trim(),cliente:String(x[ix.cliente]||'').trim(),rut:String(x[ix.rut]||'').trim(),vendedor:String(x[ix.vendedor]||'').trim(),clase_cliente:String(x[ix.clase_cliente]||'CLIENTE NORMAL').trim().toUpperCase(),direccion:String(x[ix.direccion]||'').trim(),giro:String(x[ix.giro]||'').trim(),medio_pago:String(x[ix.medio_pago]||'').trim(),telefono:String(x[ix.telefono]||'').trim(),correo_electronico:String(x[ix.correo]||'').trim(),responsable_empresa:String(x[ix.responsable]||'').trim(),fecha_creacion:fechaHoraVisiblePedido(x[ix.fecha]),status:String(x[ix.status]||'ACTIVO').trim().toUpperCase()})).filter(x=>x.cliente);
  }
  return arr.map(x=>({rowNumber:x.rowNumber||x.fila||x.row||'',id:String(x.id||x.ID||'').trim(),cliente:String(x.cliente||x.nombre||x.razon_social||x['razón social']||'').trim(),rut:String(x.rut||x.RUT||'').trim(),vendedor:String(x.vendedor||x.ejecutivo||x.asesor||'').trim(),clase_cliente:String(x.clase_cliente||x.claseCliente||x.tipo_cliente||x.tipoCliente||x.categoria_cliente||x.categoria||'CLIENTE NORMAL').trim().toUpperCase(),direccion:String(x.direccion||x['dirección']||'').trim(),giro:String(x.giro||'').trim(),medio_pago:String(x.medio_pago||x.medioPago||x.forma_pago||'').trim(),telefono:String(x.telefono||x.fono||'').trim(),correo_electronico:String(x.correo_electronico||x.correo||x.email||x.mail||'').trim(),responsable_empresa:String(x.responsable_empresa||x.responsable||'').trim(),fecha_creacion:fechaHoraVisiblePedido(x.fecha_creacion||x.fecha),status:String(x.status||x.estado||'ACTIVO').trim().toUpperCase()})).filter(x=>x.cliente);
}
function llenarDatalistClientes(){
  const dl=$('dlClientes'); if(!dl) return;
  dl.innerHTML=(state.clientes||[]).filter(c=>String(c.status||'ACTIVO').toUpperCase()!=='BLOQUEADO').map(c=>`<option value="${esc(c.cliente)}">${esc([c.rut,c.vendedor,c.clase_cliente].filter(Boolean).join(' | ') || c.direccion || '')}</option>`).join('');
}
function abrirClientesModal(){
  $('modalClientes')?.classList.add('show');
  cargarClientes().catch(err=>{console.warn(err);toast('No se pudo cargar CLIENTES');});
}
function cerrarClientesModal(){ $('modalClientes')?.classList.remove('show'); }
function limpiarClienteForm(){
  ['cliRowNumber','cliId','cliCliente','cliRut','cliDireccion','cliGiro','cliTelefono','cliCorreo','cliResponsable'].forEach(id=>{ if($(id)) $(id).value=''; });
  if($('cliMedioPago')) $('cliMedioPago').value='';
  if($('cliVendedor')) $('cliVendedor').value='';
  if($('cliClaseCliente')) $('cliClaseCliente').value='CLIENTE NORMAL';
  if($('cliStatus')) $('cliStatus').value='ACTIVO';
}
function clienteFormPayload(){
  return {
    rowNumber:($('cliRowNumber')?.value||'').trim(),
    id:($('cliId')?.value||'').trim(),
    cliente:($('cliCliente')?.value||'').trim(),
    rut:($('cliRut')?.value||'').trim(),
    vendedor:($('cliVendedor')?.value||'').trim(),
    clase_cliente:($('cliClaseCliente')?.value||'CLIENTE NORMAL').trim().toUpperCase(),
    direccion:($('cliDireccion')?.value||'').trim(),
    giro:($('cliGiro')?.value||'').trim(),
    medio_pago:($('cliMedioPago')?.value||'').trim(),
    telefono:($('cliTelefono')?.value||'').trim(),
    correo_electronico:($('cliCorreo')?.value||'').trim(),
    responsable_empresa:($('cliResponsable')?.value||'').trim(),
    status:($('cliStatus')?.value||'ACTIVO').trim().toUpperCase()
  };
}
async function guardarCliente(){
  const data=clienteFormPayload();
  if(!data.cliente) return toast('Ingresa el nombre del cliente');
  const accion=data.rowNumber||data.id?'editar_cliente':'agregar_cliente';
  const r=await api(accion,data);
  if(!r?.ok) return toast(r?.msg||'No se pudo guardar cliente');
  limpiarClienteForm();
  await cargarClientes();
  toast('Cliente guardado en base de datos CLIENTES');
}
function renderClientesTabla(){
  const tb=$('tbodyClientes'); if(!tb) return;
  const lista=state.clientes||[];
  if(!lista.length){ tb.innerHTML='<tr><td colspan="13" style="text-align:center;color:#64748b;padding:18px">Sin clientes registrados.</td></tr>'; return; }
  tb.innerHTML=lista.map((c,i)=>`<tr><td>${esc(c.cliente)}</td><td>${esc(c.rut||'-')}</td><td>${esc(c.vendedor||'-')}</td><td>${esc(c.clase_cliente||'CLIENTE NORMAL')}</td><td>${esc(c.direccion||'-')}</td><td>${esc(c.giro||'-')}</td><td>${esc(c.medio_pago||'-')}</td><td>${esc(c.telefono||'-')}</td><td>${esc(c.correo_electronico||'-')}</td><td>${esc(c.responsable_empresa||'-')}</td><td>${esc(c.fecha_creacion||'-')}</td><td><span class="badge ${String(c.status).includes('BLOQUE')?'cancelado':String(c.status).includes('INACT')?'pendiente':'terminado'}">${esc(c.status||'ACTIVO')}</span></td><td class="actions-cell"><button class="secondary" data-edit-cliente="${i}">Editar</button><button class="warn" data-del-cliente="${i}">Eliminar</button></td></tr>`).join('');
}
function manejarClickClientes(e){
  const edit=e.target.closest('[data-edit-cliente]');
  const del=e.target.closest('[data-del-cliente]');
  if(edit){ cargarClienteEnForm(Number(edit.dataset.editCliente)); return; }
  if(del){ eliminarCliente(Number(del.dataset.delCliente)); return; }
}
function cargarClienteEnForm(i){
  const c=(state.clientes||[])[i]; if(!c) return;
  if($('cliRowNumber')) $('cliRowNumber').value=c.rowNumber||'';
  if($('cliId')) $('cliId').value=c.id||'';
  if($('cliCliente')) $('cliCliente').value=c.cliente||'';
  if($('cliRut')) $('cliRut').value=c.rut||'';
  if($('cliVendedor')) { if(c.vendedor && ![...$('cliVendedor').options].some(o=>o.value===c.vendedor)) $('cliVendedor').insertAdjacentHTML('beforeend',`<option value="${esc(c.vendedor)}">${esc(c.vendedor)}</option>`); $('cliVendedor').value=c.vendedor||''; }
  if($('cliClaseCliente')) $('cliClaseCliente').value=c.clase_cliente||'CLIENTE NORMAL';
  if($('cliDireccion')) $('cliDireccion').value=c.direccion||'';
  if($('cliGiro')) $('cliGiro').value=c.giro||'';
  if($('cliMedioPago')) $('cliMedioPago').value=c.medio_pago||'';
  if($('cliTelefono')) $('cliTelefono').value=c.telefono||'';
  if($('cliCorreo')) $('cliCorreo').value=c.correo_electronico||'';
  if($('cliResponsable')) $('cliResponsable').value=c.responsable_empresa||'';
  if($('cliStatus')) $('cliStatus').value=c.status||'ACTIVO';
  toast('Cliente cargado para edición');
}
async function eliminarCliente(i){
  const c=(state.clientes||[])[i]; if(!c) return;
  if(!confirm('¿Eliminar cliente '+c.cliente+'?')) return;
  const r=await api('eliminar_cliente',{rowNumber:c.rowNumber,id:c.id,cliente:c.cliente});
  if(!r?.ok) return toast(r?.msg||'No se pudo eliminar cliente');
  await cargarClientes();
  toast('Cliente eliminado');
}
function aplicarClienteManual(){
  const nombre=($('manCliente')?.value||'').trim();
  if(!nombre) return;
  const c=(state.clientes||[]).find(x=>String(x.cliente||'').trim().toLowerCase()===nombre.toLowerCase());
  if(c){
    if(String(c.status||'ACTIVO').toUpperCase()==='BLOQUEADO') toast('Cliente bloqueado. Revisa configuración antes de usarlo.');
    if(c.vendedor && $('manVendedor')){
      if(![...$('manVendedor').options].some(o=>o.value===c.vendedor)) $('manVendedor').insertAdjacentHTML('beforeend',`<option value="${esc(c.vendedor)}">${esc(c.vendedor)}</option>`);
      $('manVendedor').value=c.vendedor;
    }
    if(c.clase_cliente) toast('Cliente: '+c.clase_cliente+(c.vendedor?' | Vendedor: '+c.vendedor:''));
  }
  renderManualPreview();
}

async function cargarBodegas(){
  const r=await api('listar_bodegas',{});
  let lista=[];
  if(r?.ok) lista=normBodegas(r);
  const desdePedidos=[...new Set((state.pedidos||[]).map(p=>String(p.bodega_preparacion||'').trim()).filter(Boolean))].map(nombre=>({nombre,status:'ACTIVA'}));
  const map={};
  [...lista,...desdePedidos].forEach(b=>{ const n=String(b.nombre||'').trim(); if(n && !map[n]) map[n]=b; });
  state.bodegas=Object.values(map).sort((a,b)=>String(a.nombre).localeCompare(String(b.nombre)));
  llenarSelectBodegas();
}
function normBodegas(r){
  const arr=r.bodegas||r.items||r.data||r.rows||r.values||[];
  if(!Array.isArray(arr)) return [];
  if(arr.length && Array.isArray(arr[0])){
    const headers=(r.headers||[]).map(norm);
    const idx=(names)=>{for(const n of names){const i=headers.indexOf(norm(n)); if(i>=0)return i;} return -1;};
    const ix={nombre:idx(['nombre_bodega','nombre bodega','bodega','nombre']),direccion:idx(['direccion','dirección']),fecha:idx(['fecha_creacion','fecha creación','fecha']),status:idx(['status','estado']),radio:idx(['radio_metros','radio','radio metros']),lat:idx(['latitud','lat']),lng:idx(['longitud','lng','lon'])};
    return arr.map(x=>({nombre:String(x[ix.nombre]||'').trim(),direccion:String(x[ix.direccion]||'').trim(),fecha_creacion:fechaHoraVisiblePedido(x[ix.fecha]),status:String(x[ix.status]||'ACTIVA').trim().toUpperCase(),radio_metros:Number(x[ix.radio]||200)||200,latitud:String(x[ix.lat]||'').trim(),longitud:String(x[ix.lng]||'').trim()})).filter(x=>x.nombre);
  }
  return arr.map(x=>({nombre:String(x.nombre||x.nombre_bodega||x.bodega||'').trim(),direccion:String(x.direccion||x.dirección||'').trim(),fecha_creacion:fechaHoraVisiblePedido(x.fecha_creacion||x.fecha),status:String(x.status||x.estado||'ACTIVA').trim().toUpperCase(),radio_metros:Number(x.radio_metros||x.radio||200)||200,latitud:String(x.latitud||x.lat||'').trim(),longitud:String(x.longitud||x.lng||x.lon||'').trim()})).filter(x=>x.nombre);
}
function llenarSelectBodegas(){
  [$('manBodega'),$('editBodegaPedido'),$('impBodega')].filter(Boolean).forEach(s=>{
    const v=s.value;
    const activas=(state.bodegas||[]).filter(b=>String(b.status||'ACTIVA').toUpperCase()==='ACTIVA');
    const first=s.id==='impBodega'?'<option value="">Tomar desde archivo / Sin bodega</option>':'<option value="">Sin bodega</option>';
    s.innerHTML=first+activas.map(b=>`<option value="${esc(b.nombre)}">${esc(b.nombre)}${b.direccion?' — '+esc(b.direccion):''}</option>`).join('');
    if(v){
      if(!activas.some(b=>String(b.nombre)===String(v))) s.insertAdjacentHTML('beforeend',`<option value="${esc(v)}">${esc(v)}</option>`);
      s.value=v;
    }
  });
}
function opcionesBodegas(actual=''){
  const activas=(state.bodegas||[]).filter(b=>String(b.status||'ACTIVA').toUpperCase()==='ACTIVA').map(b=>b.nombre);
  const lista=[...new Set([...activas, String(actual||'').trim()].filter(Boolean))];
  return '<option value="">Sin bodega</option>'+lista.map(n=>`<option value="${esc(n)}" ${String(n)===String(actual||'')?'selected':''}>${esc(n)}</option>`).join('');
}
function iniciarMapaBodega(){
  const el=$('bodegaMap');
  if(!el || !window.L) return null;
  if(state.bodegaMapa) return state.bodegaMapa;
  el.innerHTML='';
  state.bodegaMapa=L.map(el).setView([-33.45,-70.66],11);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap'}).addTo(state.bodegaMapa);
  setTimeout(()=>state.bodegaMapa.invalidateSize(),250);
  return state.bodegaMapa;
}
function pintarBodegaMapa(lat,lng,radio){
  lat=Number(lat); lng=Number(lng); radio=Number(radio)||200;
  if(!Number.isFinite(lat)||!Number.isFinite(lng)) return;
  const map=iniciarMapaBodega(); if(!map) return;
  if(state.bodegaMarker) state.bodegaMarker.remove();
  if(state.bodegaCircle) state.bodegaCircle.remove();
  state.bodegaMarker=L.marker([lat,lng]).addTo(map);
  state.bodegaCircle=L.circle([lat,lng],{radius:radio}).addTo(map);
  map.setView([lat,lng],16);
  if($('txtBodegaLat')) $('txtBodegaLat').value=String(lat.toFixed(7));
  if($('txtBodegaLng')) $('txtBodegaLng').value=String(lng.toFixed(7));
  setTimeout(()=>map.invalidateSize(),250);
}
function actualizarRadioBodegaMapa(){
  const lat=Number($('txtBodegaLat')?.value||'');
  const lng=Number($('txtBodegaLng')?.value||'');
  if(Number.isFinite(lat)&&Number.isFinite(lng)) pintarBodegaMapa(lat,lng,Number($('txtBodegaRadio')?.value||200)||200);
}
async function buscarBodegaEnMapa(mostrarToast=true){
  const dir=($('txtBodegaDireccion')?.value||'').trim();
  const status=$('bodegaMapStatus');
  if(!dir){ if(mostrarToast) toast('Ingresa la dirección de la bodega'); return null; }
  if(status) status.textContent='Buscando dirección en mapa...';
  try{
    const q=encodeURIComponent(dir+', Chile');
    const res=await fetch('https://nominatim.openstreetmap.org/search?format=json&limit=1&q='+q,{headers:{'Accept':'application/json'}});
    const data=await res.json();
    if(!Array.isArray(data)||!data.length) throw new Error('Dirección no encontrada');
    const lat=Number(data[0].lat), lng=Number(data[0].lon);
    pintarBodegaMapa(lat,lng,Number($('txtBodegaRadio')?.value||200)||200);
    if(status) status.textContent='Ubicación encontrada. Revisa el marcador y el radio antes de guardar.';
    return {lat,lng};
  }catch(e){
    if(status) status.textContent='No se pudo geolocalizar automáticamente. Puedes revisar la dirección o guardar sin coordenadas.';
    if(mostrarToast) toast('No se pudo ubicar la dirección en el mapa.');
    return null;
  }
}
async function agregarBodega(){
  const nombre=($('txtBodegaNombre')?.value||'').trim();
  const direccion=($('txtBodegaDireccion')?.value||'').trim();
  const status=($('selBodegaStatus')?.value||'ACTIVA').trim().toUpperCase();
  const radio=Number($('txtBodegaRadio')?.value||200)||200;
  let lat=($('txtBodegaLat')?.value||'').trim();
  let lng=($('txtBodegaLng')?.value||'').trim();
  if(!nombre) return toast('Ingresa el nombre de la bodega');
  if(!direccion) return toast('Ingresa la dirección de la bodega');
  if(!lat || !lng){ const geo=await buscarBodegaEnMapa(false); if(geo){ lat=String(geo.lat); lng=String(geo.lng); } }
  const r=await api('agregar_bodega',{nombre_bodega:nombre,direccion,status,radio_metros:radio,latitud:lat,longitud:lng});
  if(!r?.ok) return toast(r?.msg||'No se pudo guardar la bodega');
  ['txtBodegaNombre','txtBodegaDireccion','txtBodegaLat','txtBodegaLng'].forEach(id=>{ if($(id)) $(id).value=''; });
  if($('txtBodegaRadio')) $('txtBodegaRadio').value='200';
  if($('selBodegaStatus')) $('selBodegaStatus').value='ACTIVA';
  await cargarBodegas();
  toast('Bodega guardada en base de datos BODEGAS');
}



/* ================= MAESTRA DE PRODUCTOS =================
   Se sincroniza desde la base de datos MAESTRA y se usa para completar
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
  // Se deja vacío intencionalmente: cargar miles de opciones en un datalist
  // relentece la digitación del código. Ahora se usa búsqueda exacta/debounced.
  const dl=$('dlProductosMaestra');
  if(dl) dl.innerHTML='';
}
async function cargarMaestra(silencioso=true){
  if(silencioso && !state.maestra.length) cargarCacheMaestra();
  const msg=$('manualMaestraMsg');
  if(msg && !silencioso) msg.textContent='Sincronizando MAESTRA de productos desde base de datos...';

  const pageSize=500;
  let offset=0;
  let todos=[];
  let vueltas=0;
  let ultimo=null;

  // La MAESTRA quedó paginada para evitar timeout; Orden de Pedido debe recorrer
  // todas las páginas para que aparezcan todos los códigos, no solo los primeros 300/500.
  while(true){
    vueltas++;
    const r=await api('listar_maestra_bd',{limit:pageSize,offset});
    if(!r?.ok) throw new Error(r?.msg||'No se pudo leer MAESTRA');
    ultimo=r;
    const lista=normProductosMaestra(r);
    todos=todos.concat(lista);
    if(msg && !silencioso){
      const totalRef=Number(r.totalFilasMaestra||r.total||0)||'';
      msg.textContent='Sincronizando MAESTRA: '+todos.length+(totalRef?' de '+totalRef:'')+' productos leídos...';
    }
    const hasMore=Boolean(r.hasMore);
    const nextOffset=Number(r.offset||offset)+Number(r.limit||pageSize);
    if(!hasMore || !lista.length || nextOffset<=offset || vueltas>80) break;
    offset=nextOffset;
    await new Promise(res=>setTimeout(res,25));
  }

  // Respaldo para Web Apps antiguos que todavía no tengan listar_maestra_bd paginado.
  if(!todos.length){
    const r=await api('listar_maestra',{});
    if(!r?.ok) throw new Error(r?.msg||'No se pudo leer MAESTRA');
    todos=normProductosMaestra(r);
    ultimo=r;
  }

  setMaestra(todos);
  guardarCacheMaestra();
  state.maestraTotal=Number(ultimo?.totalFilasMaestra||todos.length)||todos.length;
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
  const base=state.maestraMap[k] || null;
  return enriquecerProductoConUbicacionMovimiento(base,codigo);
}
function ubicacionMovimientoPorCodigo(codigo){
  const k=normCodigo(codigo);
  if(!k) return '';
  return (state.ubicacionesMap && state.ubicacionesMap[k]) ? String(state.ubicacionesMap[k]).trim() : '';
}
function enriquecerProductoConUbicacionMovimiento(prod,codigo){
  const ubiMov=ubicacionMovimientoPorCodigo(codigo || prod?.codigo || '');
  if(prod){
    return {...prod, ubicacion: ubiMov || prod.ubicacion || ''};
  }
  return ubiMov ? {codigo:String(codigo||'').trim(), descripcion:'', cantidad:0, ubicacion:ubiMov} : null;
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
  // Flujo optimizado: no cargar listas dinámicas completas al digitar.
  // Solo busca exacto localmente y consulta al servidor cuando el código tiene largo suficiente
  // o cuando el usuario presiona explícitamente Buscar en Maestra.
  if(!forzar && codigo.length < 3){
    const loc=ubicacionMovimientoPorCodigo(codigo);
    if(loc && ubi && !ubi.value.trim()) ubi.value=loc;
    if(msg) msg.textContent='Escribe al menos 3 caracteres o presiona Buscar para consultar MAESTRA.';
    return null;
  }

  if(!state.maestraListaCargada){
    cargarCacheMaestra();
    if(!state.maestra.length && forzar){
      try{ await cargarMaestra(true); }catch(e){ console.warn(e); }
    }
  }

  let prod=productoMaestraPorCodigo(codigo);
  if(!prod && forzar){
    try{
      const r=await api('buscar_producto',{codigo});
      if(r?.ok) prod=enriquecerProductoConUbicacionMovimiento({codigo:r.codigo||codigo, descripcion:r.descripcion||'', cantidad:r.cantidad||0, ubicacion:r.ubicacion||''},codigo);
    }catch(e){ console.warn(e); }
  }
  if(!prod && forzar){
    try{
      const r=await api('buscar_maestra',{codigo});
      if(r?.ok) prod=enriquecerProductoConUbicacionMovimiento({codigo:r.codigo||codigo, descripcion:r.descripcion||'', cantidad:r.cantidad||0, ubicacion:r.ubicacion||''},codigo);
    }catch(e){ console.warn(e); }
  }

  const loc=ubicacionMovimientoPorCodigo(codigo);
  if(prod || loc){
    if(desc && prod && (forzar || !desc.value.trim())) desc.value=prod.descripcion||desc.value;
    if(ubi && (forzar || !ubi.value.trim())) ubi.value=(prod?.ubicacion||loc||ubi.value);
    if(msg) msg.textContent=(prod?.descripcion?'Producto encontrado: '+prod.descripcion:'Ubicación encontrada en MOVIMIENTO')+(ubi?.value?' | Ubicación: '+ubi.value:'');
    renderManualPreview();
    return prod || {codigo,descripcion:'',ubicacion:loc,cantidad:0};
  }
  if(msg) msg.textContent='Código no encontrado en MAESTRA/MOVIMIENTO. Puedes ingresar la descripción y ubicación manualmente.';
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
      fecha:fechaVisiblePedido(p.fecha),
      cliente:String(p.cliente||'').trim(),
      vendedor:String(p.vendedor||'').trim(),
      pikeador:limpiarPikeador(p.pikeador, p.cliente),
      bodega_preparacion:String(p.bodega_preparacion||p.bodegaPreparacion||p.bodega_pedido||p.bodega||'').trim(),
      status:estadoSeguroImport(p.status||p.estado,'PENDIENTE'),
      alerta_token:String(p.alerta_token||'').trim(),
      alerta_ts:String(p.alerta_ts||'').trim(),
      alerta_tipo:String(p.alerta_tipo||'').trim(),
      alerta_mensaje:String(p.alerta_mensaje||'').trim(),
      hora_inicio:fechaHoraVisiblePedido(p.hora_inicio||p.horaInicio||p.inicio_preparacion||p.fecha_inicio||p.fecha_asignacion),
      hora_termino:fechaHoraVisiblePedido(p.hora_termino||p.horaTermino||p.termino_preparacion||p.fecha_termino),
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
    if(!m.has(key))m.set(key,{key,pedido:ped,fecha:fechaVisiblePedido(pick(r,ix.fecha,'')),cliente:pick(r,ix.cliente,''),vendedor:pick(r,ix.vendedor,''),pikeador:limpiarPikeador(pick(r,ix.pikeador,''), pick(r,ix.cliente,'')),bodega_preparacion:pick(r,ix.bodega_preparacion,''),status:estadoSeguroImport(pick(r,ix.status,''),'PENDIENTE'),alerta_token:pick(r,ix.alerta_token,''),alerta_ts:pick(r,ix.alerta_ts,''),alerta_tipo:pick(r,ix.alerta_tipo,''),alerta_mensaje:pick(r,ix.alerta_mensaje,''),hora_inicio:fechaHoraVisiblePedido(pick(r,ix.hora_inicio,'')),hora_termino:fechaHoraVisiblePedido(pick(r,ix.hora_termino,'')),tiempo_preparacion_min:Number(String(pick(r,ix.tiempo_preparacion_min,0)).replace(',','.'))||0,total_productos:0,total_unidades:0,productos:[],rows:[]});
    const p=m.get(key);
    ['fecha','cliente','vendedor','bodega_preparacion','alerta_token','alerta_ts','alerta_tipo','alerta_mensaje','hora_inicio','hora_termino'].forEach(k=>{if(ix[k]>=0&&pick(r,ix[k],'')){const val=pick(r,ix[k],'');p[k]=k==='fecha'?fechaVisiblePedido(val):(k==='hora_inicio'||k==='hora_termino'?fechaHoraVisiblePedido(val):val);}});
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
    bodega_preparacion:i(['bodega_preparacion','bodega preparación','bodega preparacion','bodega pedido','bodega_origen','bodega preparación pedido','bodega de preparacion','bodega de preparación','nombre_bodega','nombre bodega']),
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
  const q=($('txtBuscar')?.value||'').toLowerCase().trim();
  const e=($('selEstado')?.value||'').toUpperCase().trim();
  const pk=($('selFiltroPikeador')?.value||'').toLowerCase().trim();
  const vend=($('selFiltroVendedor')?.value||'').toLowerCase().trim();
  const rango=obtenerRangoFechaOrden();
  state.filtrados=state.pedidos.filter(p=>{
    const txt=[p.pedido,p.cliente,p.vendedor,p.pikeador,p.bodega_preparacion,p.status,...p.productos.map(x=>`${x.codigo} ${x.descripcion} ${x.ubicacion}`)].join(' ').toLowerCase();
    if(q&&!txt.includes(q))return false;
    if(e&&String(p.status||'').toUpperCase()!==e)return false;
    if(pk&&String(p.pikeador||'').toLowerCase()!==pk)return false;
    if(vend&&String(p.vendedor||'').toLowerCase()!==vend)return false;
    const f=fechaPedidoOrden(p);
    if(rango.desde && (!f || f<rango.desde)) return false;
    if(rango.hasta && (!f || f>rango.hasta)) return false;
    return true;
  });
  render(); kpis();
}
function render(){
  const tb=$('tbodyPedidos'), mob=$('mobileList');
  if(!state.filtrados.length){
    tb.innerHTML='<tr><td colspan="13" style="text-align:center;padding:26px;color:#64748b">Sin registros para mostrar. Presiona Sincronizar o revisa la base de datos de pedidos.</td></tr>';
    mob.innerHTML='<div class="pedido-card">Sin registros para mostrar.</div>';
    return;
  }
  tb.innerHTML=state.filtrados.map(p=>`<tr class="pedido-row ${state.selectedKey===p.key?'row-selected':''}" data-pedido-key="${esc(p.key)}" tabindex="0" title="Seleccionar para abrir pedido"><td>${esc(fechaVisiblePedido(p.fecha))}</td><td><b>${esc(p.pedido)}</b></td><td>${esc(p.cliente)}</td><td>${esc(p.vendedor)}</td><td>${esc(pikeadorVisible(p.pikeador))}</td><td>${p.total_productos}</td><td><b>${p.total_unidades}</b></td><td>${badgeEstado(p.status)}</td><td>${esc(horaCorta(p.hora_inicio)||'-')}</td><td>${esc(horaCorta(p.hora_termino)||'-')}</td><td><b>${esc(tiempoPreparacionTexto(p))}</b></td><td>${badgeAlerta(p)}</td><td class="actions-cell"><button data-ver-pedido="${esc(p.key)}">Ver</button></td></tr>`).join('');
  mob.innerHTML=state.filtrados.map(p=>`<div class="pedido-card ${state.selectedKey===p.key?'row-selected':''}" data-pedido-key="${esc(p.key)}" tabindex="0"><div class="top"><span>${esc(p.pedido)}</span>${badgeEstado(p.status)}</div><p><b>Cliente:</b> ${esc(p.cliente||'-')}</p><p><b>Pikeador:</b> ${esc(pikeadorVisible(p.pikeador))}</p><p><b>Productos:</b> ${p.total_productos} | <b>Unidades:</b> ${p.total_unidades}</p><p><b>Inicio:</b> ${esc(horaCorta(p.hora_inicio)||'-')} | <b>Término:</b> ${esc(horaCorta(p.hora_termino)||'-')}</p><p><b>Tiempo preparación:</b> ${esc(tiempoPreparacionTexto(p))}</p><p>${badgeAlerta(p)}</p><button data-ver-pedido="${esc(p.key)}">Ver opciones</button></div>`).join('');
}
function kpis(){
  const a=state.filtrados||[];
  const procesados=a.filter(esPedidoProcesado);
  const productos=a.reduce((s,p)=>s+Number(p.total_productos||0),0);
  const unidades=a.reduce((s,p)=>s+Number(p.total_unidades||0),0);
  const pendientes=a.filter(p=>estadoPedido(p)==='PENDIENTE').length;
  const tiempos=procesados.map(minutosPreparacionPedido).filter(x=>x>0);
  const set=(id,v)=>{const el=$(id); if(el) el.textContent=v;};
  set('kPedidos',a.length);
  set('kProductos',productos);
  set('kUnidades',unidades);
  set('kPendientes',pendientes);
  set('kProcesados',procesados.length);
  set('kPromedioProductos',procesados.length?dec(productos/procesados.length,1):'0');
  set('kPromedioUnidades',procesados.length?dec(unidades/procesados.length,1):'0');
  set('kTiempoPromedio',tiempos.length?formatearMinutos(tiempos.reduce((s,x)=>s+x,0)/tiempos.length):'0 min');
  const hoy=new Date();
  set('kUnidadesDia', unidadesPorPeriodoOrden(a,'DIA',hoy));
  set('kUnidadesMes', unidadesPorPeriodoOrden(a,'MES',hoy));
  set('kUnidadesAnio', unidadesPorPeriodoOrden(a,'ANIO',hoy));
  renderRotacionYRendimiento(a);
}
function dec(n,d=1){n=Number(n)||0;return n.toFixed(d).replace(/\.0$/,'');}
function abrirModalDesdeTabla(e){
  if(e.target.closest('button,a,input,select,textarea,label')) return;
  const tr=e.target.closest('tr[data-pedido-key]');
  if(!tr) return;
  seleccionarFilaPedido(tr.dataset.pedidoKey,true);
}
function abrirModalDesdeTarjetaMovil(e){
  if(e.target.closest('button,a,input,select,textarea,label')) return;
  const card=e.target.closest('[data-pedido-key]');
  if(!card) return;
  seleccionarFilaPedido(card.dataset.pedidoKey,true);
}
function manejarTeclasFilaPedido(e){
  const tr=e.target.closest('tr[data-pedido-key]');
  if(!tr) return;
  if(e.key==='Enter' || e.key===' '){
    e.preventDefault();
    seleccionarFilaPedido(tr.dataset.pedidoKey,true);
  }
}
function manejarFlechasPedidos(e){
  if(e.key!=='ArrowDown' && e.key!=='ArrowUp') return;
  const tag=(document.activeElement?.tagName||'').toLowerCase();
  if(['input','select','textarea'].includes(tag)) return;
  const modalAbierto=$('modalPedido')?.classList.contains('show');
  const dentroOtroModal=document.activeElement?.closest?.('#modalImportar,#modalManual');
  if(dentroOtroModal) return;
  e.preventDefault();
  moverSeleccionPedido(e.key==='ArrowDown'?1:-1,modalAbierto || !!state.selectedKey);
}
function seleccionarFilaPedido(key,abrir){
  const p=state.filtrados.find(x=>x.key===key||x.pedido===key) || state.pedidos.find(x=>x.key===key||x.pedido===key);
  if(!p) return;
  state.selectedKey=p.key;
  marcarFilaSeleccionada(p.key);
  if(abrir) abrirModal(p.key);
}
function marcarFilaSeleccionada(key){
  document.querySelectorAll('[data-pedido-key].row-selected').forEach(x=>x.classList.remove('row-selected'));
  document.querySelectorAll(`[data-pedido-key="${cssEscapeSafe(key)}"]`).forEach(x=>x.classList.add('row-selected'));
  const row=document.querySelector(`#tbodyPedidos [data-pedido-key="${cssEscapeSafe(key)}"]`);
  if(row){
    row.focus({preventScroll:true});
    row.scrollIntoView({block:'nearest',inline:'nearest'});
  }
}
function moverSeleccionPedido(delta,abrir){
  const lista=state.filtrados.length?state.filtrados:state.pedidos;
  if(!lista.length) return;
  const actual=state.selectedKey || state.sel?.key || lista[0].key;
  let idx=lista.findIndex(x=>x.key===actual || x.pedido===actual);
  if(idx<0) idx=delta>0?-1:0;
  const next=(idx+delta+lista.length)%lista.length;
  seleccionarFilaPedido(lista[next].key,abrir);
}
function cssEscapeSafe(v){
  v=String(v??'');
  if(window.CSS && typeof CSS.escape==='function') return CSS.escape(v);
  return v.replace(/\\/g,'\\\\').replace(/"/g,'\\"');
}

/* ================= MODAL PEDIDO ================= */
function abrirModal(key){
  const p=state.pedidos.find(x=>x.key===key||x.pedido===key);
  if(!p)return toast('No se encontró el pedido');
  state.sel=p;
  state.selectedKey=p.key;
  marcarFilaSeleccionada(p.key);
  $('modalTitle').textContent='Pedido '+p.pedido;
  $('modalSub').textContent=(p.cliente||'')+' | Total unidades: '+(p.total_unidades||0);
  $('detallePedido').innerHTML=[['Fecha',fechaVisiblePedido(p.fecha)],['Pedido',p.pedido],['Cliente',p.cliente],['Vendedor',p.vendedor],['Pikeador',pikeadorVisible(p.pikeador)],['Bodega preparación',p.bodega_preparacion||'-'],['Estado',p.status],['Hora inicio',horaCorta(p.hora_inicio)||'-'],['Hora término',horaCorta(p.hora_termino)||'-'],['Tiempo preparación',tiempoPreparacionTexto(p)],['Productos',p.total_productos],['Total unidades',p.total_unidades],['Avance cliente',porcentajeAvancePedido(p)+'%']].map(([l,v])=>`<div class="detail-item"><div class="l">${l}</div><div class="v">${esc(v)}</div></div>`).join('');
  $('tbodyProductos').innerHTML=p.productos.length?p.productos.map(x=>`<tr><td>${esc(x.codigo)}</td><td>${esc(x.descripcion)}</td><td>${esc(x.ubicacion)}</td><td><b>${x.cantidad}</b></td></tr>`).join(''):'<tr><td colspan="4" style="text-align:center;color:#64748b">Este pedido no tiene productos detallados.</td></tr>';
  llenarSelects();
  renderEdicionPedido(p);
  activarEdicionPedido(false);
  $('selModalPikeador').value=p.pikeador||'';
  if($('selModalEstado')) $('selModalEstado').value=(p.status||'PENDIENTE').toUpperCase();
  $('modalPedido').classList.add('show');
}
function cerrarModal(){ activarEdicionPedido(false); $('modalPedido').classList.remove('show'); }


/* ================= EDICIÓN COMPLETA PEDIDO ================= */
function instalarEdicionPedidoUI(){
  if($('btnEditarPedido')) return;
  const actions=document.querySelector('#modalPedido .modal-actions-top');
  const body=document.querySelector('#modalPedido .modal-body');
  if(!actions||!body) return;
  const editBtns=document.createElement('span');
  editBtns.className='edit-actions-inline';
  editBtns.innerHTML=`
    <button id="btnEditarPedido" class="ok" type="button">✏️ Editar Pedido</button>
    <button id="btnGuardarEdicionPedido" class="ok hidden" type="button">💾 Guardar Edición</button>
    <button id="btnCancelarEdicionPedido" class="secondary hidden" type="button">Cancelar Edición</button>
  `;
  actions.insertBefore(editBtns, $('btnPdfPedido') || null);
  const box=document.createElement('div');
  box.id='editPedidoBox';
  box.className='edit-pedido-box hidden';
  const detalle=$('detallePedido');
  if(detalle && detalle.parentNode) detalle.parentNode.insertBefore(box, detalle.nextSibling);
  else body.insertBefore(box, body.firstChild);
  if(!document.getElementById('cssEdicionPedido')){
    const st=document.createElement('style');
    st.id='cssEdicionPedido';
    st.textContent=`
      .hidden{display:none!important}
      .edit-actions-inline{display:contents}
      .edit-pedido-box{border:1px solid #bfdbfe;background:#eff6ff;border-radius:18px;padding:14px;margin:12px 0;box-shadow:0 8px 24px rgba(15,23,42,.08)}
      .edit-grid{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:10px;margin-bottom:12px}
      .edit-field{display:flex;flex-direction:column;gap:5px;grid-column:span 3;min-width:0}
      .edit-field.span-2{grid-column:span 2}.edit-field.span-4{grid-column:span 4}.edit-field.span-6{grid-column:span 6}
      .edit-field label{font-size:12px;font-weight:800;color:#1e3a8a;text-transform:uppercase;letter-spacing:.03em}
      .edit-field input,.edit-field select{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:12px;padding:10px 11px;background:#fff;color:#0f172a;font-size:14px;outline:none}
      .edit-field input:focus,.edit-field select:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.14)}
      .edit-products-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:10px 0;flex-wrap:wrap}
      .edit-products-wrap{overflow:auto;border:1px solid #dbeafe;border-radius:14px;background:#fff;max-height:42vh}
      .edit-products-wrap table{min-width:820px;width:100%;border-collapse:collapse}
      .edit-products-wrap th,.edit-products-wrap td{padding:8px;border-bottom:1px solid #e5e7eb;vertical-align:middle}
      .edit-products-wrap input{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:10px;padding:8px;background:#fff;color:#0f172a}
      .edit-products-wrap .cantidad{max-width:96px}
      .btn-mini{padding:8px 10px;border-radius:10px;border:0;font-weight:800;cursor:pointer}
      .btn-mini.danger{background:#fee2e2;color:#991b1b}.btn-mini.info{background:#dbeafe;color:#1d4ed8}
      @media(max-width:900px){.edit-grid{grid-template-columns:1fr}.edit-field,.edit-field.span-2,.edit-field.span-4,.edit-field.span-6{grid-column:span 1}.edit-pedido-box{padding:10px}.edit-products-wrap{max-height:50vh}}
    `;
    document.head.appendChild(st);
  }
}
function renderEdicionPedido(p){
  const box=$('editPedidoBox');
  if(!box||!p) return;
  const productos=(Array.isArray(p.productos)&&p.productos.length?p.productos:[{codigo:'',descripcion:'',ubicacion:'',cantidad:1}]);
  const pikeadorOptions=opcionesPikeadorEdicion(p.pikeador);
  box.innerHTML=`
    <div class="edit-grid">
      <div class="edit-field span-2"><label>Fecha</label><input id="editFechaPedido" value="${esc(fechaVisiblePedido(p.fecha)||'')}"></div>
      <div class="edit-field span-2"><label>Pedido</label><input id="editNumeroPedido" value="${esc(p.pedido||'')}"></div>
      <div class="edit-field span-4"><label>Cliente</label><input id="editClientePedido" value="${esc(p.cliente||'')}"></div>
      <div class="edit-field span-2"><label>Vendedor</label><select id="editVendedorPedido">${opcionesVendedores(p.vendedor||'')}</select></div>
      <div class="edit-field span-2"><label>Estado</label><select id="editEstadoPedido">${opcionesEstadoEdicion(p.status)}</select></div>
      <div class="edit-field span-4"><label>Pikeador</label><select id="editPikeadorPedido">${pikeadorOptions}</select></div>
      <div class="edit-field span-4"><label>Bodega preparación</label><select id="editBodegaPedido">${opcionesBodegas(p.bodega_preparacion||'')}</select></div>
      <div class="edit-field span-4"><label>Hora inicio</label><input id="editHoraInicioPedido" value="${esc(p.hora_inicio||'')}" readonly></div>
      <div class="edit-field span-4"><label>Hora término</label><input id="editHoraTerminoPedido" value="${esc(p.hora_termino||'')}" readonly></div>
    </div>
    <div class="edit-products-head">
      <h3 style="margin:0">Editar productos del pedido</h3>
      <button id="btnAddProductoEdit" class="info" type="button">➕ Agregar producto</button>
    </div>
    <div class="edit-products-wrap">
      <table>
        <thead><tr><th style="width:52px">#</th><th>Código</th><th>Descripción</th><th>Ubicación</th><th style="width:110px">Cantidad</th><th style="width:160px">Acción</th></tr></thead>
        <tbody id="tbodyEditProductos">${productos.map((x,i)=>filaProductoEdicion(x,i)).join('')}</tbody>
      </table>
    </div>
    <div class="hint" style="margin-top:10px">Al guardar, se actualizan los datos del pedido y sus productos en la base de datos PEDIDOS. Los ceros iniciales del código se conservan como texto.</div>
  `;
}
function opcionesEstadoEdicion(actual){
  const estados=['PENDIENTE','PREPARACION','RECIBIDO','DESPACHADO','TERMINADO','CANCELADO'];
  actual=estadoSeguroImport(actual,'PENDIENTE');
  return estados.map(e=>`<option value="${e}" ${e===actual?'selected':''}>${e}</option>`).join('');
}
function opcionesPikeadorEdicion(actual){
  const lista=[...new Set([String(actual||'').trim(),...state.pikeadores].filter(Boolean))];
  return '<option value="">Sin asignar</option>'+lista.map(n=>`<option value="${esc(n)}" ${String(n)===String(actual||'')?'selected':''}>${esc(n)}</option>`).join('');
}
function filaProductoEdicion(x,i){
  return `<tr data-edit-row="${i}">
    <td><b>${i+1}</b></td>
    <td><input class="editCodigo" value="${esc(x.codigo||'')}" placeholder="Código/SKU"></td>
    <td><input class="editDescripcion" value="${esc(x.descripcion||'')}" placeholder="Descripción"></td>
    <td><input class="editUbicacion" value="${esc(x.ubicacion||'')}" placeholder="Ubicación"></td>
    <td><input class="editCantidad cantidad" type="number" min="0" step="1" value="${esc(x.cantidad||1)}"></td>
    <td><button class="btn-mini info" type="button" data-edit-buscar="1">Buscar</button> <button class="btn-mini danger" type="button" data-edit-del="1">Eliminar</button></td>
  </tr>`;
}
function activarEdicionPedido(on){
  state.editandoPedido=!!on;
  const box=$('editPedidoBox');
  if(box) box.classList.toggle('hidden',!on);
  const det=$('detallePedido');
  if(det) det.style.display=on?'none':'';
  const wrap=document.querySelector('#modalPedido .products-wrap');
  if(wrap) wrap.style.display=on?'none':'';
  const h3=[...document.querySelectorAll('#modalPedido .modal-body h3')].find(x=>(x.textContent||'').includes('Productos preagregados'));
  if(h3) h3.style.display=on?'none':'';
  $('btnEditarPedido')?.classList.toggle('hidden',on);
  $('btnGuardarEdicionPedido')?.classList.toggle('hidden',!on);
  $('btnCancelarEdicionPedido')?.classList.toggle('hidden',!on);
}
function manejarClickEdicionPedido(e){
  const add=e.target.closest('#btnAddProductoEdit');
  if(add){ agregarFilaProductoEdicion(); return; }
  const del=e.target.closest('[data-edit-del]');
  if(del){
    const tr=del.closest('tr');
    const tb=$('tbodyEditProductos');
    if(tb && tb.querySelectorAll('tr').length<=1) return toast('El pedido debe mantener al menos un producto');
    tr?.remove(); renumerarFilasEdicion(); return;
  }
  const buscar=e.target.closest('[data-edit-buscar]');
  if(buscar){ buscarProductoFilaEdicion(buscar.closest('tr')); }
}
function manejarCambioEdicionPedido(e){
  if(e.target?.classList?.contains('editCodigo')) buscarProductoFilaEdicion(e.target.closest('tr'),false);
}
function agregarFilaProductoEdicion(){
  const tb=$('tbodyEditProductos'); if(!tb) return;
  tb.insertAdjacentHTML('beforeend',filaProductoEdicion({codigo:'',descripcion:'',ubicacion:'',cantidad:1},tb.querySelectorAll('tr').length));
}
function renumerarFilasEdicion(){
  document.querySelectorAll('#tbodyEditProductos tr').forEach((tr,i)=>{tr.dataset.editRow=i; const c=tr.querySelector('td b'); if(c)c.textContent=i+1;});
}
async function buscarProductoFilaEdicion(tr,forzar=true){
  if(!tr) return;
  const codigo=(tr.querySelector('.editCodigo')?.value||'').trim();
  if(!codigo) return;
  if(!forzar && codigo.length < 3) return;
  let prod=productoMaestraPorCodigo(codigo);
  if(!prod && forzar){ try{ const r=await api('buscar_producto',{codigo}); if(r?.ok) prod=enriquecerProductoConUbicacionMovimiento(r,codigo); }catch(e){} }
  const loc=ubicacionMovimientoPorCodigo(codigo);
  if(!prod && !loc){ if(forzar) toast('Código no encontrado en MAESTRA/MOVIMIENTO'); return; }
  const desc=tr.querySelector('.editDescripcion'), ubi=tr.querySelector('.editUbicacion');
  if(desc && prod && (forzar || !desc.value.trim())) desc.value=prod.descripcion||desc.value;
  if(ubi && (forzar || !ubi.value.trim())) ubi.value=(prod?.ubicacion||loc||ubi.value);
}
function productosDesdeEdicion(){
  return [...document.querySelectorAll('#tbodyEditProductos tr')].map(tr=>({
    codigo:String(tr.querySelector('.editCodigo')?.value||'').replace(/^'+/,''),
    descripcion:(tr.querySelector('.editDescripcion')?.value||'').trim(),
    ubicacion:(tr.querySelector('.editUbicacion')?.value||'').trim(),
    cantidad:Number(String(tr.querySelector('.editCantidad')?.value||'1').replace(',','.'))||0
  })).filter(x=>x.codigo||x.descripcion||x.ubicacion||x.cantidad>0);
}
async function guardarEdicionPedidoActual(){
  const p=state.sel;
  if(!p) return toast('No hay pedido seleccionado');
  const items=productosDesdeEdicion();
  if(!items.length) return toast('Agrega al menos un producto al pedido');
  const pedidoNuevo=($('editNumeroPedido')?.value||'').trim();
  if(!pedidoNuevo) return toast('El número de pedido no puede quedar vacío');
  const payload={
    pedido_original:p.pedido,
    pedido:pedidoNuevo,
    fecha:fechaVisiblePedido(($('editFechaPedido')?.value||'').trim()),
    cliente:($('editClientePedido')?.value||'').trim(),
    vendedor:($('editVendedorPedido')?.value||'').trim(),
    pikeador:($('editPikeadorPedido')?.value||'').trim(),
    bodega_preparacion:($('editBodegaPedido')?.value||'').trim(),
    status:($('editEstadoPedido')?.value||'PENDIENTE').trim().toUpperCase(),
    items:JSON.stringify(items)
  };
  const r=await api('editar_pedido_completo',payload);
  if(!r?.ok) return toast('No se pudo editar el pedido: '+(r?.msg||'error'));
  toast('Pedido editado correctamente. Filas actualizadas: '+(r.actualizados||0)+'. Nuevas: '+(r.insertados||0)+'. Eliminadas: '+(r.eliminados||0));
  if(r.voice) voz(r.voice);
  await cargarTodo();
  const key=String(r.pedido||pedidoNuevo).replace(/\s+/g,'').toUpperCase();
  abrirModal(key);
}
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
  const datosVoz = {pedido:r.pedido||p.pedido, cliente:r.cliente||p.cliente, vendedor:r.vendedor||p.vendedor, pikeador:r.pikeador||pk||p.pikeador, status:r.status||estado};
  const vozMsg = mensajeEstadoPedido(p, estado, datosVoz);
  toast(vozMsg);
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
  const vozPrep = mensajeEstadoPedido(p, 'PREPARACION', {pedido:r.pedido||p.pedido, cliente:r.cliente||p.cliente, vendedor:r.vendedor||p.vendedor, pikeador:r.pikeador||pk||p.pikeador});
  voz(vozPrep); toast(vozPrep + '. Hora de inicio registrada.');
  await cargarTodo(); await verificarAlertas(); abrirModal(p.key);
}
async function reenviarActual(){
  const p=state.sel; if(!p)return;
  const r=await api('reenviar_alerta_pedido',{pedido:p.pedido,pikeador:$('selModalPikeador').value||p.pikeador||''});
  if(!r.ok)return toast(r.msg||'No se pudo reenviar');
  const vozReenvio = mensajeEstadoPedido(p, 'PREPARACION', {pedido:r.pedido||p.pedido, cliente:r.cliente||p.cliente, vendedor:r.vendedor||p.vendedor, pikeador:r.pikeador||$('selModalPikeador').value||p.pikeador});
  voz(vozReenvio); toast('Alerta reenviada: ' + vozReenvio);
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
        const msgAlerta = p.alerta_mensaje || mensajeEstadoPedido(p, estadoAlerta || p.status || 'ACTUALIZADO');
        toast('🔔 ' + msgAlerta);
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
function cargarDatosConfiguracionPedidoManual(){
  if(!state.pikeadores.length) cargarPikeadores().catch(console.warn);
  if(!state.vendedores.length) cargarVendedores().catch(console.warn);
  if(!state.bodegas.length) cargarBodegas().catch(console.warn);
  if(!state.clientes.length) cargarClientes().catch(console.warn);
  if(!state.maestraListaCargada) cargarMaestra(true).catch(console.warn);
}

function abrirManual(){
  llenarSelects();
  cargarDatosConfiguracionPedidoManual();
  limpiarManual(false);
  const hoy=new Date();
  const yyyy=hoy.getFullYear(), mm=String(hoy.getMonth()+1).padStart(2,'0'), dd=String(hoy.getDate()).padStart(2,'0');
  if($('manFecha')&&!$('manFecha').value) $('manFecha').value=`${yyyy}-${mm}-${dd}`;
  prepararPedidoAutogeneradoLocal();
  generarNumeroPedidoManual(false).catch(()=>{});
  $('modalManual')?.classList.add('show');
  setTimeout(()=>$('manCliente')?.focus(),80);
  renderManualPreview();
}
function cerrarManual(){ $('modalManual')?.classList.remove('show'); }
function limpiarManual(clearHeader=true){
  state.manualProductos=[];
  if(clearHeader){
    ['manCliente','manVendedor','manBodega','manCodigo','manDescripcion','manUbicacion'].forEach(id=>{ if($(id)) $(id).value=''; });
    prepararPedidoAutogeneradoLocal();
    generarNumeroPedidoManual(false).catch(()=>{});
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

function numeroPedidoConsecutivoJS(v){
  const raw=String(v||'').trim();
  if(!raw || /^SIN_PEDIDO/i.test(raw)) return 0;
  const m=raw.match(/\d+/g);
  if(!m) return 0;
  const n=Number(m.join(''));
  return Number.isFinite(n)&&n>0?Math.floor(n):0;
}
function calcularSiguientePedidoLocal(){
  const usados=new Set();
  (state.pedidos||[]).forEach(p=>{
    const n=numeroPedidoConsecutivoJS(p.pedido||p.key||'');
    if(n>0) usados.add(n);
  });
  let n=1;
  while(usados.has(n)) n++;
  return String(n);
}
function prepararPedidoAutogeneradoLocal(){
  const input=$('manPedido');
  if(!input) return;
  input.value=calcularSiguientePedidoLocal();
  input.readOnly=true;
  input.classList.add('input-readonly');
  input.title='Número generado automáticamente desde la base de datos PEDIDOS';
  const msg=$('manualNumeroMsg');
  if(msg) msg.textContent='Número sugerido automáticamente. Se validará nuevamente al guardar para evitar duplicados.';
}
async function generarNumeroPedidoManual(mostrarToast=false){
  const input=$('manPedido');
  if(!input) return;
  const fallback=calcularSiguientePedidoLocal();
  if(!input.value) input.value=fallback;
  const msg=$('manualNumeroMsg');
  if(msg) msg.textContent='Generando número consecutivo desde la BD...';
  try{
    const r=await api('siguiente_numero_pedido',{});
    const numero=String(r?.numero||r?.pedido||'').trim();
    if(r?.ok && numero){
      input.value=numero;
      if(msg) msg.textContent='Pedido N° '+numero+' reservado visualmente. Al guardar, Apps Script confirmará el siguiente disponible.';
      if(mostrarToast) toast('Número de pedido generado: '+numero);
    }else{
      input.value=fallback;
      if(msg) msg.textContent='Usando número local sugerido. Al guardar, Apps Script validará que no exista.';
    }
  }catch(e){
    input.value=fallback;
    if(msg) msg.textContent='Sin respuesta del generador remoto. Se usará '+fallback+' y se validará al guardar.';
  }
  renderManualPreview();
}

function datosManualBase(){
  return {
    fecha:fechaVisiblePedido($('manFecha')?.value||''),
    pedido:($('manPedido')?.value||'').trim(),
    cliente:($('manCliente')?.value||'').trim(),
    clase_cliente:(()=>{ const n=($('manCliente')?.value||'').trim().toLowerCase(); const c=(state.clientes||[]).find(x=>String(x.cliente||'').trim().toLowerCase()===n); return c?.clase_cliente||''; })(),
    vendedor:($('manVendedor')?.value||'').trim(),
    pikeador:($('manPikeador')?.value||'').trim(),
    bodega_preparacion:($('manBodega')?.value||'').trim(),
    status:estadoSeguroImport($('manStatus')?.value||'PENDIENTE','PENDIENTE')
  };
}
function agregarProductoManual(){
  const base=datosManualBase();
  const codigo=($('manCodigo')?.value||'').trim();
  const descripcion=($('manDescripcion')?.value||'').trim();
  let ubicacion=($('manUbicacion')?.value||'').trim();
  if(!ubicacion && codigo) ubicacion=ubicacionMovimientoPorCodigo(codigo);
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
  box.innerHTML=`<b>Pedido:</b> ${esc(base.pedido||'Sin número')} | <b>Cliente:</b> ${esc(base.cliente||'Sin cliente')}${base.clase_cliente?' ('+esc(base.clase_cliente)+')':''} | <b>Pikeador:</b> ${esc(base.pikeador||'Sin asignar')} | <b>Bodega:</b> ${esc(base.bodega_preparacion||'Sin bodega')} | <b>Estado:</b> ${esc(base.status)} | <b>Productos agregados:</b> ${state.manualProductos.length} | <b>Total unidades:</b> ${total}${parcialTxt}`;
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
    bodega_preparacion:base.bodega_preparacion,
    status:base.status,
    codigo:p.codigo,
    descripcion:p.descripcion,
    ubicacion:p.ubicacion,
    cantidad:p.cantidad,
    __raw:{fecha:base.fecha,pedido:base.pedido,cliente:base.cliente,vendedor:base.vendedor,pikeador:base.pikeador,bodega_preparacion:base.bodega_preparacion,status:base.status,codigo:p.codigo,descripcion:p.descripcion,ubicacion:p.ubicacion,cantidad:p.cantidad,origen:'MANUAL'},
    __rawHeaders:['fecha','pedido','cliente','vendedor','pikeador','bodega_preparacion','status','codigo','descripcion','ubicacion','cantidad','origen']
  }));
  const r=await api('guardar_pedido',{autogenerar_pedido:1,pedido:base.pedido,cliente:base.cliente,vendedor:base.vendedor,pikeador:base.pikeador,bodega_preparacion:base.bodega_preparacion,status:base.status,fecha:base.fecha,items:JSON.stringify(items)});
  if(!r?.ok) return toast('No se pudo guardar: '+(r?.msg||'error'));
  const pedidoFinal=String(r.pedido||base.pedido||'').trim();
  toast(`Pedido ${pedidoFinal} guardado. Insertados: ${r.insertados||0}. Actualizados: ${r.actualizados||0}. Sin cambio: ${r.sin_cambios||0}.`);
  cerrarManual();
  await cargarTodo();
}

/* ================= IMPORTAR LISTADO ================= */
function abrirImportar(){
  llenarSelects();
  if(!state.bodegas.length) cargarBodegas().catch(console.warn);
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

function resetImportFinal(){
  const box=$('importFinalBox');
  if(box) box.classList.add('hidden');
  ['chkImportConexion','chkImportEnvio','chkImportVerificacion','chkImportFinal'].forEach(id=>{
    const el=$(id);
    if(el) el.textContent='⏳';
  });
  const title=$('importFinalTitle');
  if(title) title.textContent='Resultado de importación';
  const msg=$('importFinalMsg');
  if(msg) msg.textContent='El resultado aparecerá al finalizar el envío.';
  state.importFallidos=[];
  renderImportFailedList();
}
function setImportCheck(id,estado){
  const el=$(id);
  if(!el) return;
  const e=String(estado||'pendiente').toLowerCase();
  el.textContent=e==='ok'?'✅':(e==='error'?'❌':(e==='warn'?'⚠️':'⏳'));
}
function mostrarResultadoImportacion(res){
  const box=$('importFinalBox');
  if(box) box.classList.remove('hidden');
  const ok=!Number(res?.fallidos||0);
  setImportCheck('chkImportConexion','ok');
  setImportCheck('chkImportEnvio', Number(res?.enviados||0)>0 ? 'ok' : 'error');
  setImportCheck('chkImportVerificacion', ok ? 'ok' : 'warn');
  setImportCheck('chkImportFinal', ok ? 'ok' : 'error');
  const title=$('importFinalTitle');
  if(title) title.textContent=ok?'✅ Importación finalizada correctamente':'❌ Importación finalizada con errores';
  const msg=$('importFinalMsg');
  if(msg){
    const detalle=Number(res?.fallidos||0)>0?' Revisa abajo los registros marcados en rojo para corregirlos y volver a importar.':'';
    msg.textContent=`Enviadas: ${res?.enviados||0}. Insertados: ${res?.insertados||0}. Actualizados: ${res?.actualizados||0}. Sin cambio: ${res?.sinCambios||0}. Fallidas: ${res?.fallidos||0}.${detalle}`;
  }
  renderImportFailedList();
}
function limpiarImportacion(clearFile=true){
  state.importados=[];
  state.importFallidos=[];
  state.importHeaders=[];
  if(clearFile && $('fileImportar')) $('fileImportar').value='';
  setImportStats(0,0,0,0,0,0,0);
  resetImportProgress();
  resetImportFinal();
  if($('tbodyImportar')) $('tbodyImportar').innerHTML='<tr><td colspan="14" style="text-align:center;color:#64748b;padding:20px">Selecciona un archivo y presiona Previsualizar.</td></tr>';
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
    const bodegaDefault=$('impBodega')?.value||'';
    const items=data.map((r,i)=>{
      if(i%50===0 || i===data.length-1){
        const pct=55+Math.round(((i+1)/Math.max(1,data.length))*20);
        setImportProgress(pct,'Validando filas del archivo',`Fila ${i+1} de ${data.length}`);
      }
      return normalizarFilaImport(r,headers,ix,i,fechaDefault,statusDefault,pikeadorDefault,bodegaDefault);
    }).filter(x=>x._hasData);
    setImportProgress(78,'Comparando con PEDIDOS','Verificando nuevos, cambios y duplicados contra la tabla actual.');
    state.importados=validarImportadosContraPedidos(items);
    renderImportPreview();
    const enviables=(state.importados||[]).filter(x=>!x._error && x._hasData).length;
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
    bodega_preparacion:idxHeader(headers,['bodega_preparacion','bodega preparación','bodega preparacion','bodega pedido','bodega de preparacion','bodega de preparación','nombre_bodega','nombre bodega','bodega']),
    codigo:idxHeader(headers,['codigo','código','cod','sku','codigo producto','código producto']),
    descripcion:idxHeader(headers,['descripcion','descripción','detalle','nombre','producto','nombre producto','descripcion producto','descripción producto']),
    ubicacion:idxHeader(headers,['ubicacion','ubicación','ubicacion producto','ubicación producto','ubicacion producto pedido','ubicación producto pedido','posicion','posición','rack','pasillo']),
    cantidad:idxHeader(headers,['cantidad','unidades','cajas','qty','cant']),
    status:idxHeader(headers,['status','estado','estado pedido','estado del pedido']),
    observacion:idxHeader(headers,['observacion','observación','obs','comentario','comentarios'])
  };

  const coreHits = ['pedido','codigo','descripcion','cliente'].filter(k => ix[k] >= 0).length;
  if(coreHits < 2){
    // Compatibilidad con plantillas sin encabezado: se aplican posiciones base,
    // pero ESTADO sigue sin posición por defecto para no copiar DESCRIPCIÓN.
    return {
      fecha:1,pedido:0,cliente:2,vendedor:3,pikeador:4,bodega_preparacion:5,codigo:6,descripcion:7,ubicacion:8,cantidad:9,status:10,observacion:11
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

function normalizarFilaImport(r,headers,ix,i,fechaDefault,statusDefault,pikeadorDefault,bodegaDefault){
  const raw = rawObjectFromRow(headers,r);
  const rawStatus = pick(r,ix.status,'');
  const statusFinal = estadoSeguroImport(rawStatus, statusDefault || 'PENDIENTE');
  const clienteImport = pick(r,ix.cliente,'');
  const pikeadorDesdeArchivo = ix.pikeador >= 0 ? pick(r,ix.pikeador,'') : '';
  const pikeadorFinal = limpiarPikeador(pikeadorDesdeArchivo || pikeadorDefault, clienteImport);
  const bodegaDesdeArchivo = ix.bodega_preparacion >= 0 ? pick(r,ix.bodega_preparacion,'') : pickRaw(raw,['bodega_preparacion','bodega preparación','bodega preparacion','bodega pedido','bodega de preparacion','bodega de preparación','nombre_bodega','nombre bodega','bodega']);
  const bodegaFinal = bodegaDesdeArchivo || bodegaDefault || '';
  const item={
    fecha:fechaVisiblePedido(pick(r,ix.fecha,fechaDefault)||pickRaw(raw,['fecha','fecha ingreso','fecha pedido'])||fechaDefault),
    pedido:obtenerPedidoDesdeImport(r,headers,ix,raw)||pickRaw(raw,['pedido','orden','orden pedido','orden de pedido']),
    cliente:clienteImport,
    vendedor:pick(r,ix.vendedor,''),
    pikeador:pikeadorFinal,
    bodega_preparacion:bodegaFinal,
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
      const base={fecha:fechaVisiblePedido(p.fecha)||'',pedido:p.pedido||'',cliente:p.cliente||'',vendedor:p.vendedor||'',pikeador:p.pikeador||'',bodega_preparacion:p.bodega_preparacion||'',codigo:prod.codigo||'',descripcion:prod.descripcion||'',ubicacion:prod.ubicacion||'',cantidad:Number(prod.cantidad||0),status:p.status||'PENDIENTE'};
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
    // Si el mismo pedido/código tiene varias ubicaciones o descripciones, se debe respetar como línea distinta.
    // El fallback por pedido+código solo se usa cuando la fila importada no trae ubicación ni descripción.
    const pcKey = pedidoCodigoKey(item);
    const usarFallbackPedidoCodigo = pcKey && !String(item.ubicacion||'').trim() && !String(item.descripcion||'').trim();
    const actual=porKey.get(k)||(usarFallbackPedidoCodigo?porPedidoCodigo.get(pcKey):null);
    if(!actual){ stats.nuevos++; item._estadoImport='NUEVO'; return item; }
    item._exists=true;
    const campos=['fecha','cliente','vendedor','pikeador','bodega_preparacion','descripcion','ubicacion','cantidad','status'];
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
  const e=String((x._postError||x._error)?'ERROR':(x._estadoImport||'NUEVO')).toUpperCase();
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
  const fallidosVista=items.filter(x=>x._postError||x._error);
  const normalesVista=items.filter(x=>!(x._postError||x._error));
  const view=(fallidosVista.length?[...fallidosVista,...normalesVista]:items).slice(0,200);
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
    const errorTxt=x._postError||x._error||'';
    const cambios=esc(errorTxt||(([...(x._statusAdvertencia?[x._statusAdvertencia]:[]), ...(x._changes||[])]).join(' | ')||'—'));
    const rawCells=rawHeaders.map(h=>`<td>${esc(esHeaderFechaPedido(h)?fechaVisiblePedido(raw[h]):(raw[h]??''))}</td>`).join('');
    const cls=(x._postError||x._error)?'import-row-error':'';
    return `<tr class="${cls}"><td>${i+1}</td><td>${estadoImportBadge(x)}</td><td class="changes-list">${cambios}</td>${rawCells}</tr>`;
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
function importFailureKey(x){ return importKey(x||{}); }
function matchFaltantesImport(batchOriginal, detalleFaltantes){
  const detalles=Array.isArray(detalleFaltantes)?detalleFaltantes:[];
  if(!detalles.length) return batchOriginal.slice();
  const faltantesKeys=new Set(detalles.map(importFailureKey).filter(Boolean));
  return batchOriginal.filter(x=>faltantesKeys.has(importFailureKey(x)));
}
function registrarFallosImportacion(batchOriginal,motivo,detalleFaltantes=[]){
  const afectados=matchFaltantesImport(batchOriginal||[],detalleFaltantes);
  afectados.forEach(x=>{
    x._postError=motivo||'No se pudo guardar o verificar esta fila en la base de datos';
    x._estadoImport='ERROR';
  });
  const existentes=new Set((state.importFallidos||[]).map(x=>String(x._row||'')+'|'+importFailureKey(x)));
  const nuevos=afectados.filter(x=>{
    const k=String(x._row||'')+'|'+importFailureKey(x);
    if(existentes.has(k)) return false;
    existentes.add(k);
    return true;
  });
  state.importFallidos=[...(state.importFallidos||[]),...nuevos];
  return afectados.length;
}
function renderImportFailedList(){
  const box=$('importFailedBox');
  const tbody=$('tbodyImportFallidos');
  if(!box || !tbody) return;
  const fallidos=state.importFallidos||[];
  if(!fallidos.length){
    box.classList.add('hidden');
    tbody.innerHTML='';
    return;
  }
  box.classList.remove('hidden');
  tbody.innerHTML=fallidos.map((x,i)=>`<tr class="import-row-error"><td>${i+1}</td><td>${esc(x._row||'')}</td><td>${esc(x.pedido||'')}</td><td>${esc(x.cliente||'')}</td><td>${esc(x.bodega_preparacion||'')}</td><td>${esc(x.codigo||'')}</td><td>${esc(x.descripcion||'')}</td><td>${esc(x.ubicacion||'')}</td><td>${esc(x.cantidad||'')}</td><td class="changes-list">${esc(x._postError||x._error||'No se pudo guardar')}</td></tr>`).join('');
}

function sleepImport(ms){ return new Promise(resolve=>setTimeout(resolve,ms)); }
function clavesVerificacionImport(items){
  return (items||[]).map(x=>({pedido:x.pedido||'',codigo:x.codigo||'',descripcion:x.descripcion||'',ubicacion:x.ubicacion||''}));
}
async function verificarImportacionBatch(batch,intentos=4,esperaMs=1200){
  const keys=clavesVerificacionImport(batch);
  if(!keys.length) return {ok:true,encontrados:0,faltantes:0};
  let last=null;
  for(let i=0;i<intentos;i++){
    if(i>0) await sleepImport(esperaMs);
    try{
      const r=await api('verificar_importacion_pedidos',{items:JSON.stringify(keys),_ver:i});
      last=r;
      if(r?.ok && Number(r.encontrados||0)>=keys.length) return r;
    }catch(err){ last={ok:false,msg:err.message||String(err)}; }
  }
  return last || {ok:false,msg:'No se pudo verificar importación'};
}

async function enviarImportacionBD(){
  const items=(state.importados||[]).filter(x=>!x._error && x._hasData);
  if(!items.length){
    setImportProgress(100,'Sin filas válidas para enviar','Existen errores pendientes de corregir o el archivo está vacío.','error');
    mostrarResultadoImportacion({enviados:0,insertados:0,actualizados:0,sinCambios:0,fallidos:0});
    setImportCheck('chkImportEnvio','error');
    setImportCheck('chkImportFinal','error');
    return toast('No hay filas válidas para enviar. Corrige los errores o revisa el archivo.');
  }
  if(!confirm(`Enviar ${items.length} fila(s) válida(s) a la base de datos PEDIDOS?`)) return;
  resetImportFinal();
  state.importFallidos=[];
  (state.importados||[]).forEach(x=>{ delete x._postError; });
  renderImportPreview();
  const finalBox=$('importFinalBox');
  if(finalBox) finalBox.classList.remove('hidden');
  setImportCheck('chkImportConexion','pendiente');
  setImportCheck('chkImportEnvio','pendiente');
  setImportCheck('chkImportVerificacion','pendiente');
  setImportCheck('chkImportFinal','pendiente');
  let enviados=0, fallidos=0, insertados=0, actualizados=0, sinCambios=0;
  const chunkSize=20;
  const totalChunks=Math.ceil(items.length/chunkSize);
  setImportProgress(0,'Iniciando sincronización con base de datos',`Se enviarán ${items.length} filas en ${totalChunks} bloque(s).`);
  setImportCheck('chkImportConexion','ok');
  for(let i=0,chunkIndex=0;i<items.length;i+=chunkSize,chunkIndex++){
    const batchOriginal=items.slice(i,i+chunkSize);
    const batch=batchOriginal.map(({_row,_hasData,_error,_postError,_estadoImport,_exists,_changes,_statusAdvertencia,...x})=>x);
    const desde=i+1;
    const hasta=Math.min(i+chunkSize,items.length);
    const pctInicio=Math.round((i/items.length)*100);
    setImportProgress(pctInicio,'Sincronizando con base de datos',`Bloque ${chunkIndex+1} de ${totalChunks} | Filas ${desde}-${hasta} de ${items.length}`);
    setStatus(`Sincronizando con base de datos. Filas ${desde}-${hasta} de ${items.length}...`);
    try{
      const r=await apiPostIframe('importar_pedidos',{items:JSON.stringify(batch)},180000).catch(err=>({ok:false,msg:err.message||String(err),pendiente_verificacion:true}));
      const ver=await verificarImportacionBatch(batch,5,1600);
      if(ver.ok && Number(ver.encontrados||0)>=batch.length){
        enviados+=batch.length;
        insertados+=Number(r.insertados||0);
        actualizados+=Number(r.actualizados||0);
        sinCambios+=Number(r.sin_cambios||0);
        setImportCheck('chkImportEnvio','ok');
        setImportCheck('chkImportVerificacion','ok');
        if(r.pendiente_verificacion) console.warn('Bloque guardado y verificado sin respuesta directa del iframe', r);
      }else if(ver.ok && Number(ver.encontrados||0)>0){
        const faltan=batch.length-Number(ver.encontrados||0);
        enviados+=Number(ver.encontrados||0);
        fallidos+=faltan;
        const motivo=`No se verificó en base de datos después del envío. Bloque ${chunkIndex+1}.`;
        const marcados=registrarFallosImportacion(batchOriginal,motivo,ver.detalle_faltantes||[]);
        if(marcados!==faltan) console.warn('Diferencia marcando fallidos', {faltan,marcados,ver});
        setImportCheck('chkImportEnvio','warn');
        setImportCheck('chkImportVerificacion','warn');
        console.warn('Bloque importado parcialmente', {respuesta:r,verificacion:ver});
      }else{
        fallidos+=batch.length;
        registrarFallosImportacion(batchOriginal,`No se guardó o no se pudo verificar el bloque ${chunkIndex+1}: ${(ver&&ver.msg)|| (r&&r.msg) || 'sin respuesta válida de Apps Script'}`);
        setImportCheck('chkImportEnvio','error');
        setImportCheck('chkImportVerificacion','error');
        console.warn('Bloque no verificado en BD', {respuesta:r,verificacion:ver});
      }
    }catch(err){
      console.error(err);
      fallidos+=batch.length;
      registrarFallosImportacion(batchOriginal,`Error enviando bloque ${chunkIndex+1}: ${err.message||err}`);
      setImportCheck('chkImportEnvio','error');
      setImportCheck('chkImportVerificacion','error');
      setImportProgress(pctInicio,'Error enviando bloque',`Bloque ${chunkIndex+1}: ${err.message||err}`,'error');
    }
    renderImportPreview();
    renderImportFailedList();
    const pctFin=Math.round((Math.min(i+chunkSize,items.length)/items.length)*100);
    setImportProgress(pctFin,'Sincronizando con base de datos',`Avance ${pctFin}% | Enviadas: ${enviados} | Insertados: ${insertados} | Actualizados: ${actualizados} | Fallidas: ${fallidos}`);
  }
  if(fallidos){
    setImportProgress(100,'❌ Importación finalizada con errores',`Enviadas: ${enviados}. Insertados: ${insertados}. Actualizados: ${actualizados}. Sin cambio: ${sinCambios}. Fallidas: ${fallidos}. Revisa las filas en rojo.`, 'error');
  }else{
    setImportProgress(100,'✅ Importación finalizada correctamente',`Enviadas: ${enviados}. Insertados: ${insertados}. Actualizados: ${actualizados}. Sin cambio: ${sinCambios}.`,'ok');
  }
  mostrarResultadoImportacion({enviados,insertados,actualizados,sinCambios,fallidos});
  toast(`${fallidos?'Importación finalizada con errores. Revisa las filas en rojo.':'Importación finalizada correctamente'}. Enviadas: ${enviados}. Insertados: ${insertados}. Actualizados: ${actualizados}. Sin cambio: ${sinCambios}. Fallidas: ${fallidos}.`);
  if(!fallidos){
    await cargarTodo();
    setTimeout(()=>cerrarImportar(),900);
  }else{
    setStatus(`Importación finalizada con errores. ${fallidos} fila(s) quedaron destacadas en rojo para corregir.`);
  }
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
    body:state.filtrados.map(p=>[fechaVisiblePedido(p.fecha),p.pedido,p.cliente,p.vendedor,pikeadorVisible(p.pikeador),p.total_productos,p.total_unidades,p.status]),
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
  d.setFont('helvetica','bold'); d.setFontSize(16); d.text('Orden de Pedido',105,14,{align:'center'});
  d.setFont('helvetica','bold'); d.setFontSize(24); d.text('N° PEDIDO '+String(p.pedido||'-'),105,26,{align:'center'});
  d.setFont('helvetica','normal'); d.setFontSize(9); d.text('Generado: '+new Date().toLocaleString('es-CL'),105,33,{align:'center'});
  d.setDrawColor(210); d.roundedRect(14,38,182,42,2,2);
  d.setFontSize(9);
  infoLine(d,18,46,'Pedido',p.pedido); infoLine(d,75,46,'Fecha',fechaVisiblePedido(p.fecha)||'-'); infoLine(d,132,46,'Estado',p.status||'-');
  infoLine(d,18,55,'Cliente',p.cliente||'-'); infoLine(d,105,55,'Vendedor',p.vendedor||'-');
  infoLine(d,18,64,'Pikeador',pikeadorVisible(p.pikeador)); infoLine(d,105,64,'Bodega',p.bodega_preparacion||'-');
  infoLine(d,18,73,'Total unidades',String(p.total_unidades||0));
  d.autoTable({
    startY:88,
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
  const alto=Math.max(160,86+(filas*12));
  const d=new jsPDF({orientation:'p',unit:'mm',format:[80,alto]});
  let y=6;
  d.setFont('helvetica','bold'); d.setFontSize(11); d.text('ORDEN DE PEDIDO',40,y,{align:'center'}); y+=6;
  d.setFont('helvetica','bold'); d.setFontSize(15); d.text('N° '+String(p.pedido||'-'),40,y,{align:'center'}); y+=7;
  d.setFont('helvetica','normal'); d.setFontSize(7); d.text(new Date().toLocaleString('es-CL'),40,y,{align:'center'}); y+=5;
  line80(d,y); y+=4;
  d.setFontSize(8);
  y=ticketText(d,'Pedido',p.pedido,y); y=ticketText(d,'Cliente',p.cliente||'-',y); y=ticketText(d,'Vendedor',p.vendedor||'-',y); y=ticketText(d,'Pikeador',pikeadorVisible(p.pikeador),y); y=ticketText(d,'Bodega',p.bodega_preparacion||'-',y); y=ticketText(d,'Estado',p.status||'-',y); y=ticketText(d,'Total unidades',String(p.total_unidades||0),y); y+=2;
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
  const serial=Number(v.replace(',','.'));
  if(Number.isFinite(serial) && serial>20000 && serial<90000){
    const ms=Math.round(serial*86400000);
    const d=new Date(Date.UTC(1899,11,30)+ms);
    return isNaN(d.getTime())?null:d;
  }
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
  const tbodyVen=$('tbodyRendimientoVendedor');
  const tbodyCli=$('tbodyTopClientes');
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
    const top=Object.values(map).sort((a,b)=>b.unidades-a.unidades || b.pedidos.size-a.pedidos.size).slice(0,10);
    tbodyRot.innerHTML=top.length?top.map((x,i)=>`<tr><td>${i+1}</td><td><b>${esc(x.codigo)}</b></td><td>${esc(short(x.descripcion,42))}</td><td><b>${x.unidades}</b></td><td>${x.pedidos.size}</td></tr>`).join(''):'<tr><td colspan="5" style="text-align:center;color:#64748b">Sin productos para calcular rotación.</td></tr>';
  }
  if(tbodyPik){
    const rows=resumenRendimientoPorCampoOrden(lista,'pikeador',p=>pikeadorVisible(p.pikeador), 'Sin asignar').slice(0,10);
    tbodyPik.innerHTML=rows.length?rows.map(x=>`<tr><td><b>${esc(x.nombre)}</b></td><td>${x.pedidos}</td><td>${x.procesados}</td><td>${x.unidades}</td><td><b>${x.tiempos.length?formatearMinutos(x.tiempos.reduce((s,t)=>s+t,0)/x.tiempos.length):'-'}</b></td></tr>`).join(''):'<tr><td colspan="5" style="text-align:center;color:#64748b">Sin pikeadores asignados.</td></tr>';
  }
  if(tbodyVen){
    const rows=resumenRendimientoPorCampoOrden(lista,'vendedor',p=>String(p.vendedor||'').trim()||'Sin vendedor', 'Sin vendedor').slice(0,10);
    tbodyVen.innerHTML=rows.length?rows.map(x=>`<tr><td><b>${esc(x.nombre)}</b></td><td>${x.pedidos}</td><td>${x.procesados}</td><td>${x.unidades}</td><td><b>${x.tiempos.length?formatearMinutos(x.tiempos.reduce((s,t)=>s+t,0)/x.tiempos.length):'-'}</b></td></tr>`).join(''):'<tr><td colspan="5" style="text-align:center;color:#64748b">Sin vendedores registrados.</td></tr>';
  }
  if(tbodyCli){
    const map={};
    (lista||[]).forEach(p=>{
      const c=String(p.cliente||'').trim()||'Sin cliente';
      if(!map[c]) map[c]={cliente:c,unidades:0,pedidos:0,procesados:0};
      map[c].pedidos++;
      if(esPedidoProcesado(p)){
        map[c].procesados++;
        map[c].unidades += Number(p.total_unidades||0)||0;
      }
    });
    const rows=Object.values(map).sort((a,b)=>b.unidades-a.unidades || b.procesados-a.procesados || b.pedidos-a.pedidos).slice(0,10);
    tbodyCli.innerHTML=rows.length?rows.map((x,i)=>`<tr><td>${i+1}</td><td><b>${esc(short(x.cliente,42))}</b></td><td>${x.unidades}</td><td>${x.pedidos}</td><td>${x.procesados}</td></tr>`).join(''):'<tr><td colspan="5" style="text-align:center;color:#64748b">Sin clientes para calcular ranking.</td></tr>';
  }
}
function resumenRendimientoPorCampoOrden(lista,campo,getNombre,omitir){
  const map={};
  (lista||[]).forEach(p=>{
    const nombre=getNombre(p);
    if(!nombre || nombre===omitir) return;
    if(!map[nombre]) map[nombre]={nombre,pedidos:0,procesados:0,unidades:0,tiempos:[]};
    map[nombre].pedidos++;
    map[nombre].unidades+=Number(p.total_unidades||0)||0;
    if(esPedidoProcesado(p)) map[nombre].procesados++;
    const min=minutosPreparacionPedido(p); if(min>0) map[nombre].tiempos.push(min);
  });
  return Object.values(map).sort((a,b)=>b.procesados-a.procesados || b.unidades-a.unidades || b.pedidos-a.pedidos);
}
function fechaPedidoOrden(p){
  return fechaHoraValida(p?.fecha) || fechaHoraValida(p?.hora_inicio) || fechaHoraValida(p?.hora_termino);
}
function inicioDia(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate(),0,0,0,0);}
function finDia(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate(),23,59,59,999);}
function obtenerRangoFechaOrden(){
  const periodo=String($('selPeriodoKpi')?.value||'').toUpperCase();
  const hoy=new Date();
  let desde=null,hasta=null;
  if(periodo==='HOY'){desde=inicioDia(hoy); hasta=finDia(hoy);}
  else if(periodo==='MES'){desde=new Date(hoy.getFullYear(),hoy.getMonth(),1,0,0,0,0); hasta=new Date(hoy.getFullYear(),hoy.getMonth()+1,0,23,59,59,999);}
  else if(periodo==='ANIO'){desde=new Date(hoy.getFullYear(),0,1,0,0,0,0); hasta=new Date(hoy.getFullYear(),11,31,23,59,59,999);}
  const d1=$('fechaDesdeKpi')?.value, d2=$('fechaHastaKpi')?.value;
  if(periodo==='RANGO' || d1 || d2){
    desde=d1?new Date(d1+'T00:00:00'):null;
    hasta=d2?new Date(d2+'T23:59:59'):null;
  }
  return {desde,hasta};
}
function unidadesPorPeriodoOrden(lista,periodo,base=new Date()){
  const rango = periodo==='DIA' ? {desde:inicioDia(base), hasta:finDia(base)} : periodo==='MES' ? {desde:new Date(base.getFullYear(),base.getMonth(),1,0,0,0,0), hasta:new Date(base.getFullYear(),base.getMonth()+1,0,23,59,59,999)} : {desde:new Date(base.getFullYear(),0,1,0,0,0,0), hasta:new Date(base.getFullYear(),11,31,23,59,59,999)};
  return (lista||[]).filter(p=>{const f=fechaPedidoOrden(p); return f && f>=rango.desde && f<=rango.hasta;}).reduce((s,p)=>s+(Number(p.total_unidades||0)||0),0);
}


/* ================= UTILIDADES ================= */

function abrirConfigOrden(){
  const m=$('modalConfigOrden');
  if(m) m.classList.add('show');
  setTimeout(()=>{ iniciarMapaBodega(); if(state.bodegaMapa) state.bodegaMapa.invalidateSize(); },250);
}
function cerrarConfigOrden(){
  const m=$('modalConfigOrden');
  if(m) m.classList.remove('show');
}
function togglePanel(){
  const panel=$('panelControl');
  if(!panel) return;
  const h=panel.classList.toggle('hidden');
  document.body.classList.toggle('control-oculto',h);
  const btn=$('btnTogglePanel');
  if(btn) btn.textContent=h?'Mostrar filtros':'Ocultar filtros';
}
function toggleMetricas(){
  const ocultar=!document.body.classList.contains('metricas-ocultas');
  document.body.classList.toggle('metricas-ocultas',ocultar);
  [$('panelKpis'),$('panelAnalitica')].filter(Boolean).forEach(x=>x.classList.toggle('hidden',ocultar));
  const btn=$('btnToggleMetricas');
  if(btn) btn.textContent=ocultar?'Mostrar KPI / Rendimiento':'Ocultar KPI / Rendimiento';
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

function apiPostIframe(accion,params={},timeoutMs=120000){
  return new Promise((resolve,reject)=>{
    const token='op_post_'+Date.now()+'_'+Math.floor(Math.random()*999999);
    const iframe=document.createElement('iframe');
    const form=document.createElement('form');
    const input=document.createElement('textarea');
    const frameName='frame_'+token;
    let doneCalled=false;
    let submitted=false;
    let submittedAt=0;
    iframe.name=frameName;
    iframe.style.display='none';
    iframe.onload=()=>{
      if(!submitted || doneCalled) return;
      if(Date.now()-submittedAt<500) return;
      // Apps Script a veces no permite leer la respuesta del iframe/postMessage,
      // pero el POST ya fue recibido. Se resuelve como pendiente y luego se verifica contra PEDIDOS.
      setTimeout(()=>done(null,{ok:true,pendiente_verificacion:true,msg:'POST enviado; verificación contra PEDIDOS pendiente'}),900);
    };
    form.method='POST';
    form.action=API_URL;
    form.target=frameName;
    form.style.display='none';
    input.name='data';
    input.value=JSON.stringify(Object.assign({},params,{accion:accion,__iframe_post:1,__post_token:token}));
    form.appendChild(input);
    document.body.appendChild(iframe);
    document.body.appendChild(form);
    const timer=setTimeout(()=>done(new Error('Tiempo agotado enviando a Apps Script. Revisa publicación del Web App.')),timeoutMs);
    function cleanup(){
      clearTimeout(timer);
      window.removeEventListener('message',onMsg);
      setTimeout(()=>{try{form.remove();iframe.remove();}catch(e){}},200);
    }
    function done(err,data){
      if(doneCalled) return;
      doneCalled=true;
      cleanup();
      err?reject(err):resolve(data);
    }
    function onMsg(ev){
      const msg=ev && ev.data;
      if(!msg || typeof msg!=='object') return;
      if(msg.source!=='orden_pedidos_post' || msg.token!==token) return;
      done(null,msg.data||msg.result||msg);
    }
    window.addEventListener('message',onMsg);
    try{submitted=true; submittedAt=Date.now(); form.submit();}catch(err){done(err);}
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
function datosPedidoParaVoz(p, extra){
  p = p || {};
  extra = extra || {};
  return {
    pedido: String(extra.pedido || p.pedido || '').trim(),
    cliente: String(extra.cliente || p.cliente || '').trim(),
    vendedor: String(extra.vendedor || p.vendedor || p.vendedor_asociado || '').trim(),
    pikeador: String(extra.pikeador || p.pikeador || '').trim(),
    status: String(extra.status || extra.estado || p.status || p.estado || '').trim().toUpperCase()
  };
}
function mensajeEstadoPedido(p, estado, extra){
  const d = datosPedidoParaVoz(p, Object.assign({}, extra || {}, {status: estado || (extra && (extra.status || extra.estado))}));
  const estadoTxt = d.status || 'ACTUALIZADO';
  const terminado = estadoTxt === 'TERMINADO';
  const partes = [];
  if(terminado) partes.push('Pedido terminado');
  else partes.push('Pedido cambió a estado ' + estadoTxt);
  partes.push(d.pedido ? 'número de pedido ' + d.pedido : 'número de pedido sin registrar');
  partes.push(d.cliente ? 'cliente ' + d.cliente : 'cliente sin registrar');
  if(terminado){
    partes.push(d.vendedor ? 'vendedor asociado ' + d.vendedor : 'vendedor sin registrar');
  }else{
    partes.push(d.pikeador ? 'pikeador asignado ' + d.pikeador : 'pikeador sin asignar');
  }
  return partes.join(', ');
}
function mensajePedidoTerminado(p, extra){
  return mensajeEstadoPedido(p, 'TERMINADO', extra);
}
function voz(t){ try{const u=new SpeechSynthesisUtterance(t);u.lang='es-CL';speechSynthesis.cancel();speechSynthesis.speak(u);}catch(e){} }
function pick(r,i,f){return (i == null || i < 0 || !Array.isArray(r) || i >= r.length || r[i] == null || String(r[i]).trim()==='') ? f : String(r[i]).trim();}
function norm(v){return String(v||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[\s_\-.]+/g,'')}
function numPed(v){const m=String(v||'').match(/\d+/g);return m?Number(m.join('')):0}
function esc(v){return String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function safe(v){return String(v||'pedido').replace(/[^a-z0-9_-]+/gi,'_')}
function short(v,n){v=String(v||''); return v.length>n?v.slice(0,n-1)+'…':v;}
