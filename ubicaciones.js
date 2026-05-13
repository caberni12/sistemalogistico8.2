let CARD_STATE = {}; // idFila => true (expandida) | false (colapsada)
/* =====================================================
   CONFIGURACIÓN
===================================================== */
const DEFAULT_URL_GS =
  'https://script.google.com/macros/s/AKfycbxDPLaKDy5LqC9US-DQcCicPDIPb0XlxnPPA-y6N1AdDvbHZPxLzM0awD-NoFTcVk8Fkw/exec';

function abrirDialogoImpresionPdf(doc,nombreArchivo){
  try{
    if(doc && typeof doc.setProperties==='function') doc.setProperties({title:String(nombreArchivo||'documento').replace(/\.pdf$/i,'')});
    if(doc && typeof doc.autoPrint==='function') doc.autoPrint();
    const blob=doc.output('blob');
    const url=URL.createObjectURL(blob);
    const iframe=document.createElement('iframe');
    iframe.style.position='fixed'; iframe.style.right='0'; iframe.style.bottom='0'; iframe.style.width='0'; iframe.style.height='0'; iframe.style.border='0';
    iframe.onload=()=>setTimeout(()=>{try{iframe.contentWindow.focus(); iframe.contentWindow.print();}catch(e){window.open(url,'_blank','noopener');}},350);
    iframe.src=url; document.body.appendChild(iframe);
    setTimeout(()=>{try{document.body.removeChild(iframe);}catch(e){} try{URL.revokeObjectURL(url);}catch(e){}},120000);
  }catch(err){ console.error(err); try{doc.save(nombreArchivo||'documento.pdf');}catch(e){} }
}

const API_URL_STORAGE_KEY = 'sistema_pedidos_api_url';
let URL_GS = '';

let DATA = [];
let DATA_FILTRADA = [];
let DATA_SEARCH = [];
let timerFiltroUbicaciones = null;
let ULTIMA_RECARGA_MANUAL = '';
const SYNC_MANUAL_ONLY = true; // Ubicaciones sólo consulta la BD al presionar Recargar
const RENDER_LIMIT_UBICACIONES = 700; // evita congelar el navegador con miles de nodos visibles
let ORIGEN = /android|iphone|ipad|mobile/i.test(navigator.userAgent)
  ? 'MOBILE'
  : 'WEB';

let timerBuscar = null;
let TIPO_MOV = null;
let productoFetchController = null;
let ultimoCodigoBuscado = '';

const PRODUCT_CACHE_KEY = 'ubicaciones_productos_cache_v1';
const PRODUCT_CACHE_TS_KEY = 'ubicaciones_productos_cache_v1_ts';
const PRODUCT_CACHE_TTL = 6 * 60 * 60 * 1000; // 6 horas
let PRODUCT_MAP = {}; // codigo normalizado => {codigo, descripcion, cantidad}

/* =====================================================
   UTILIDADES
===================================================== */
function $(id){ return document.getElementById(id); }

function normalizarCodigo(v){
  return String(v ?? '')
    .replace(/^['’`´]+/, '')
    .trim()
    .replace(/\s+/g, '')
    .toUpperCase();
}


function parseConfigText(text){
  const raw = String(text || '').trim();
  if(!raw) return '';
  try{
    const obj = JSON.parse(raw);
    return String(obj.API_URL || obj.apiUrl || obj.url || obj.URL_GS || '').trim();
  }catch(err){}
  const line = raw.split(/\r?\n/).map(x => x.trim()).find(x => x && !x.startsWith('#')) || '';
  const eq = line.match(/^(API_URL|URL_GS|URL)\s*=\s*(.+)$/i);
  if(eq) return eq[2].trim().replace(/^['"]|['"]$/g,'');
  if(/^https?:\/\//i.test(line)) return line.trim();
  return '';
}

async function cargarConfiguracionUrl(){
  URL_GS = DEFAULT_URL_GS;
  try{ localStorage.removeItem(API_URL_STORAGE_KEY); }catch(err){}
  const input = $('configUrlGs');
  if(input) input.value = URL_GS;
  const status = $('configStatus');
  if(status) status.textContent = 'URL activa interna: ' + URL_GS;
  return URL_GS;
}


function abrirConfiguracion(){
  const input = $('configUrlGs');
  if(input) input.value = URL_GS || DEFAULT_URL_GS;
  const status = $('configStatus');
  if(status) status.textContent = 'URL activa: ' + (URL_GS || DEFAULT_URL_GS);
  $('configModal').classList.add('active');
}

function cerrarConfiguracion(){
  $('configModal').classList.remove('active');
}

function guardarConfiguracionUrl(){
  const input = $('configUrlGs');
  const status = $('configStatus');
  const url = String(input?.value || '').trim();
  if(!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/i.test(url)){
    if(status) status.textContent = 'URL inválida. Debe ser una URL de Apps Script y terminar en /exec.';
    return false;
  }
  const anterior = URL_GS;
  URL_GS = url;
  localStorage.setItem(API_URL_STORAGE_KEY, URL_GS);
  if(anterior && anterior !== URL_GS){
    localStorage.removeItem(PRODUCT_CACHE_KEY);
    localStorage.removeItem(PRODUCT_CACHE_TS_KEY);
    PRODUCT_MAP = {};
    estadoSincronizacionManual('URL actualizada. Presiona Recargar para consultar la BD; la MAESTRA sólo se sincroniza con su botón.');
  }
  if(status) status.textContent = 'Configuración guardada: ' + URL_GS;
  return true;
}

async function probarConfiguracionUrl(){
  if(!guardarConfiguracionUrl()) return;
  const status = $('configStatus');
  if(status) status.textContent = 'Probando conexión...';
  try{
    const res = await fetch(URL_GS + '?accion=listar_maestra&test=' + Date.now(), { cache:'no-store' });
    const data = await res.json();
    if(data && data.ok){
      if(status) status.textContent = 'Conexión correcta. Apps Script respondió OK.';
    }else{
      if(status) status.textContent = 'La URL respondió, pero no entregó OK. Revisa Apps Script.';
    }
  }catch(err){
    if(status) status.textContent = 'No se pudo conectar. Revisa URL y permisos del despliegue.';
  }
}

function descargarConfigTxt(){
  const input = $('configUrlGs');
  const url = String(input?.value || URL_GS || DEFAULT_URL_GS).trim();
  const blob = new Blob([`API_URL=${url}\n`], {type:'text/plain;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'config.txt';
  a.click();
  URL.revokeObjectURL(a.href);
}

function cargarConfigTxtArchivo(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const url = parseConfigText(reader.result);
    const status = $('configStatus');
    if(url){
      $('configUrlGs').value = url;
      guardarConfiguracionUrl();
    }else if(status){
      status.textContent = 'No se encontró una URL válida dentro del TXT.';
    }
  };
  reader.readAsText(file);
}

function escapeHtml(value){
  return String(value ?? '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}

function normalizarBusqueda(value){
  return String(value ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .trim();
}

function actualizarIndiceBusquedaUbicaciones(){
  DATA_SEARCH = DATA.map(r => normalizarBusqueda((r || []).join(' ')));
}

function jsonAttr(value){
  return escapeHtml(JSON.stringify(value));
}

function vistaMovilUbicaciones(){
  return !!(window.matchMedia && window.matchMedia('(max-width: 820px)').matches);
}

function estadoSincronizacionManual(msg){
  const texto = msg || 'Sin sincronización automática. Presiona Recargar para consultar BD-MOVIMIENTO.';
  const targets = ['ubicacionesSyncStatus','statusUbicaciones','configStatus'];
  for(const id of targets){
    const el = $(id);
    if(el){ el.textContent = texto; return; }
  }
  const cont = document.querySelector('.toolbar,.topbar,.actions,.header-actions,.panel-actions') || document.body;
  if(!cont || document.getElementById('ubicacionesSyncStatus')) return;
  const div = document.createElement('div');
  div.id = 'ubicacionesSyncStatus';
  div.className = 'sync-manual-status';
  div.textContent = texto;
  cont.appendChild(div);
}

/* ===== FECHA PARA TABLA / TARJETA ===== */
function formatFechaTabla(f){
  if(!f) return '';
  const d = new Date(f);
  if(isNaN(d)) return '';
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2,'0');
  const mi = String(d.getMinutes()).padStart(2,'0');
  return `${dd}-${mm}-${yy} ${hh}:${mi}`;
}

/* ===== FECHA PARA INPUT DATE ===== */
function formatFechaInput(f){
  if(!f) return '';
  const d = new Date(f);
  if(isNaN(d)) return '';
  return d.toISOString().slice(0,10);
}

/* =====================================================
   PROGRESS BAR
===================================================== */
function startProgress(){
  const bar = $('progress-bar');
  if(!bar) return;
  bar.classList.add('active');
  bar.style.width = '30%';
}

function endProgress(){
  const bar = $('progress-bar');
  if(!bar) return;
  bar.style.width = '100%';
  setTimeout(()=>{
    bar.classList.remove('active');
    bar.style.width = '0%';
  },300);
}

/* =====================================================
   LOADER BOTONES
===================================================== */
function startBtnLoader(btn){
  if(!btn) return;
  btn.disabled = true;
  btn.dataset.txt = btn.innerHTML;
  btn.classList.add('loading');
  btn.innerHTML = '<span class="btn-loader"></span>';
}

function endBtnLoader(btn){
  if(!btn) return;
  btn.disabled = false;
  btn.classList.remove('loading');
  btn.innerHTML = btn.dataset.txt || 'Guardar';
}

/* =====================================================
   MODAL
===================================================== */
function abrirModal(){
  limpiarFormulario();
  $('origen').value = ORIGEN;
  $('modal').classList.add('active');
}

function cerrarModal(){
  cerrarScanner();
  $('modal').classList.remove('active');
}

/* =====================================================
   TIPO MOVIMIENTO
===================================================== */
function setMovimiento(tipo){
  TIPO_MOV = tipo;
  $('btnEntrada').classList.remove('active');
  $('btnSalida').classList.remove('active');
  $(tipo === 'ENTRADA' ? 'btnEntrada' : 'btnSalida')
    .classList.add('active');
}

/* =====================================================
   AUTOCOMPLETE RÁPIDO CON MEMORIA LOCAL
===================================================== */
function cargarCatalogoDesdeMemoria(){
  try{
    const raw = localStorage.getItem(PRODUCT_CACHE_KEY);
    if(!raw) return false;
    const parsed = JSON.parse(raw);
    if(!parsed || typeof parsed !== 'object') return false;
    PRODUCT_MAP = parsed;
    return Object.keys(PRODUCT_MAP).length > 0;
  }catch(err){
    console.warn('No se pudo leer catálogo local', err);
    return false;
  }
}

function guardarCatalogoEnMemoria(){
  try{
    localStorage.setItem(PRODUCT_CACHE_KEY, JSON.stringify(PRODUCT_MAP || {}));
    localStorage.setItem(PRODUCT_CACHE_TS_KEY, String(Date.now()));
  }catch(err){
    console.warn('No se pudo guardar catálogo local', err);
  }
}

function setProductoEnCache(codigo, descripcion, cantidad){
  const key = normalizarCodigo(codigo);
  if(!key) return;
  PRODUCT_MAP[key] = {
    codigo: normalizarCodigo(codigo),
    descripcion: String(descripcion || '').trim(),
    cantidad: Number(cantidad || 0)
  };
  guardarCatalogoEnMemoria();
}

function productoLocal(cod){
  const key = normalizarCodigo(cod);
  return PRODUCT_MAP[key] || null;
}

function pintarProducto(prod, mostrarSugerencia=true){
  if(!prod) return;
  $('descripcion').value = prod.descripcion || '';
  $('cantidad').value = Number(prod.cantidad || 0);
  if(mostrarSugerencia){
    const sug = $('suggest');
    sug.innerHTML = `
      <div onclick="selectProducto('${escapeHtml(prod.codigo)}','${escapeHtml(prod.descripcion)}',${Number(prod.cantidad || 0)})">
        ${escapeHtml(prod.codigo)} – ${escapeHtml(prod.descripcion)}
      </div>`;
    sug.style.display = 'block';
  }
}

function limpiarProductoUI(){
  $('descripcion').value = '';
  $('cantidad').value = '';
  $('suggest').style.display = 'none';
}

async function precargarCatalogo(force=false){
  const ts = Number(localStorage.getItem(PRODUCT_CACHE_TS_KEY) || 0);
  const vigente = ts && (Date.now() - ts < PRODUCT_CACHE_TTL);

  if(!force && cargarCatalogoDesdeMemoria() && vigente){
    return true;
  }

  try{
    const controller = new AbortController();
    const timer = setTimeout(()=>controller.abort(), 25000);
    const res = await fetch(`${URL_GS}?accion=listar_maestra`, { signal: controller.signal });
    clearTimeout(timer);
    const data = await res.json();
    if(data && data.ok && Array.isArray(data.data)){
      const nuevo = {};
      data.data.forEach(item => {
        const codigo = normalizarCodigo(item.codigo || item[0]);
        const descripcion = String(item.descripcion || item[1] || '').trim();
        if(codigo){
          nuevo[codigo] = { codigo, descripcion, cantidad: Number(item.cantidad || item[2] || 0) };
        }
      });
      if(Object.keys(nuevo).length){
        PRODUCT_MAP = nuevo;
        guardarCatalogoEnMemoria();
        return true;
      }
    }
  }catch(err){
    // Mantener catálogo local si falla la red
    cargarCatalogoDesdeMemoria();
  }
  return false;
}

function buscarCodigo(){
  clearTimeout(timerBuscar);

  const cod = normalizarCodigo($('codigo').value);

  if(!cod){
    limpiarProductoUI();
    return;
  }

  // Respuesta inmediata desde memoria local. Esto evita que el sistema piense lento.
  const local = productoLocal(cod);
  if(local){
    pintarProducto(local, true);
    return;
  }

  // Si no está en memoria, consultar la BD-MAESTRA por URL casi de inmediato.
  // Usamos buscar_maestra para leer directo desde la BD-MAESTRA del Apps Script.
  timerBuscar = setTimeout(()=>{
    if(ultimoCodigoBuscado === cod) return;
    ultimoCodigoBuscado = cod;

    if(productoFetchController){
      productoFetchController.abort();
    }
    productoFetchController = new AbortController();

    const url = `${URL_GS}?accion=buscar_maestra&codigo=${encodeURIComponent(cod)}&t=${Date.now()}`;
    fetch(url, { signal: productoFetchController.signal, cache:'no-store' })
      .then(r=>r.json())
      .then(d=>{
        if(normalizarCodigo($('codigo').value) !== cod) return;
        if(d && d.ok){
          const prod = {
            codigo: normalizarCodigo(d.codigo || cod),
            descripcion: d.descripcion || '',
            cantidad: Number(d.cantidad || 0)
          };
          setProductoEnCache(prod.codigo, prod.descripcion, prod.cantidad);
          pintarProducto(prod, true);
        }else{
          $('descripcion').value = '';
          $('cantidad').value = '';
          $('suggest').style.display = 'none';
        }
      })
      .catch(err=>{
        if(err && err.name === 'AbortError') return;
        $('descripcion').value = '';
        $('cantidad').value = '';
        $('suggest').style.display = 'none';
      });
  },30);
}

async function sincronizarMaestra(){
  const btn = event?.target || null;
  startBtnLoader(btn);
  try{
    const ok = await precargarCatalogo(true);
    alert(ok ? 'Maestra sincronizada correctamente.' : 'No se pudo sincronizar la maestra.');
  }finally{
    endBtnLoader(btn);
  }
}

function selectProducto(c,d,stock){
  $('codigo').value = normalizarCodigo(c);
  $('descripcion').value = d;
  $('cantidad').value = Number(stock || 0);
  setProductoEnCache(c, d, stock);
  $('suggest').style.display = 'none';
}

/* =====================================================
   LISTAR / RENDER TABLA + TARJETAS
===================================================== */
function cargar(){
  startProgress();
  estadoSincronizacionManual('Recargando BD-MOVIMIENTO manualmente...');
  return fetch(`${URL_GS}?accion=listar&t=${Date.now()}`, { cache:'no-store' })
    .then(r=>r.json())
    .then(d=>{
      DATA = d.data || [];
      DATA_FILTRADA = DATA;
      actualizarIndiceBusquedaUbicaciones();
      renderTabla(DATA);
      endProgress();
      ULTIMA_RECARGA_MANUAL = new Date().toLocaleString();
      estadoSincronizacionManual(`BD-MOVIMIENTO cargada manualmente. Registros: ${DATA.length}. Última recarga: ${ULTIMA_RECARGA_MANUAL}`);
      return DATA;
    })
    .catch(()=>{
      DATA = [];
      DATA_FILTRADA = [];
      renderTabla([]);
      endProgress();
      estadoSincronizacionManual('No se pudo recargar BD-MOVIMIENTO. Revisa conexión o Apps Script.');
      return DATA;
    });
}

function renderTabla(arr){
  arr = Array.isArray(arr) ? arr : [];
  const tbody = $('tabla');
  const cards = $('cards');
  if(!tbody || !cards) return;

  const esMovil = vistaMovilUbicaciones();
  const visibles = arr.slice(0, RENDER_LIMIT_UBICACIONES);
  const quedanOcultas = Math.max(0, arr.length - visibles.length);

  tbody.innerHTML = '';
  cards.innerHTML = '';

  if(!arr.length){
    tbody.innerHTML = `
      <tr>
        <td colspan="11" style="text-align:center;padding:20px">
          Sin registros
        </td>
      </tr>`;
    return;
  }

  if(!esMovil){
    const html = visibles.map(r=>{
      const id = String(r[0] ?? '');
      return `
      <tr class="fila-clickable" title="Clic para abrir detalle/edición" onclick='abrirFilaUbicacion(event, ${jsonAttr(r)})'>
        <td>${escapeHtml(r[5])}</td>
        <td>${escapeHtml(r[6])}</td>
        <td>${escapeHtml(r[4])}</td>
        <td>${escapeHtml(r[7])}</td>
        <td>${escapeHtml(formatFechaTabla(r[1]))}</td>
        <td>${escapeHtml(formatFechaTabla(r[2]))}</td>
        <td>${escapeHtml(formatFechaTabla(r[3]))}</td>
        <td>${escapeHtml(r[8])}</td>
        <td>${escapeHtml(r[9])}</td>
        <td>${escapeHtml(r[10])}</td>
        <td class="actions-td">
          <button class="edit" onclick='event.stopPropagation(); editar(${jsonAttr(r)})' title="Editar">✏️</button>
          <button class="del" onclick='event.stopPropagation(); eliminar(${jsonAttr(id)},this)' title="Eliminar">🗑️</button>
        </td>
      </tr>`;
    }).join('');

    tbody.innerHTML = html + (quedanOcultas ? `
      <tr>
        <td colspan="11" style="text-align:center;padding:14px;color:#64748b;font-weight:800">
          Vista optimizada: mostrando ${visibles.length} de ${arr.length} registros. Usa el buscador para reducir el resultado.
        </td>
      </tr>` : '');
    return;
  }

  cards.innerHTML = visibles.map(r=>{
    const id = String(r[0] ?? '');
    const open = CARD_STATE[id] === true;
    return `
  <div class="card-item fila-clickable" data-id="${escapeHtml(id)}" title="Clic para abrir detalle/edición" onclick='abrirFilaUbicacion(event, ${jsonAttr(r)})'>

    <div class="card-head">
      <div class="desc">${escapeHtml(r[6])}</div>
      <button onclick='event.stopPropagation(); toggleCard(${jsonAttr(id)}, this)'>
        ${open ? '−' : '+'}
      </button>
    </div>

    <div class="card-row"><b>Código</b><span>${escapeHtml(r[5])}</span></div>
    <div class="card-row"><b>Ubicación</b><span>${escapeHtml(r[4])}</span></div>
    <div class="card-row"><b>Stock</b><span>${escapeHtml(r[7])}</span></div>

    <div class="card-body" style="display:${open ? 'block' : 'none'}">

      <div class="card-row"><b>Responsable</b><span>${escapeHtml(r[8])}</span></div>
      <div class="card-row"><b>Origen</b><span>${escapeHtml(r[10])}</span></div>

      <div class="card-fechas">
        <div>📥 ${escapeHtml(formatFechaTabla(r[2]))}</div>
        <div>📤 ${escapeHtml(formatFechaTabla(r[3]))}</div>
      </div>

      <span class="badge ${r[9] === 'VIGENTE' ? 'vigente' : 'retirado'}">
        ${escapeHtml(r[9])}
      </span>

      <div class="card-actions">
        <button class="edit" onclick='event.stopPropagation(); editar(${jsonAttr(r)})' title="Editar">✏️</button>
        <button class="del" onclick='event.stopPropagation(); eliminar(${jsonAttr(id)},this)' title="Eliminar">🗑️</button>
      </div>

    </div>

  </div>`;
  }).join('') + (quedanOcultas ? `
    <div class="card-item" style="text-align:center;color:#64748b;font-weight:800">
      Vista optimizada: mostrando ${visibles.length} de ${arr.length} registros. Usa el buscador para reducir el resultado.
    </div>` : '');
}

/* =====================================================
   EDITAR / ELIMINAR
===================================================== */
function abrirFilaUbicacion(ev, r){
  if(ev){
    const target = ev.target;
    if(target && target.closest && target.closest('button,a,input,select,textarea,.actions-td,.card-actions,.card-head button')){
      return;
    }
  }
  editar(r);
}

function editar(r){
  abrirModal();
  $('id').value = r[0];
  $('fecha_entrada').value = formatFechaInput(r[2]);
  $('fecha_salida').value  = formatFechaInput(r[3]);
  $('ubicacion').value = r[4];
  $('codigo').value = r[5];
  $('descripcion').value = r[6];
  $('cantidad').value = Number(r[7] || 0);
  $('cantidad_mov').value = '';
  $('responsable').value = r[8];
  $('status').value = r[9];
  $('origen').value = r[10];
}

function eliminar(idFila, btn){
  if(!confirm('¿Eliminar este movimiento?')) return;
  startBtnLoader(btn);

  fetch(URL_GS,{
    method:'POST',
    body:JSON.stringify({accion:'eliminar_movimiento',id:idFila})
  })
  .then(()=>{
    endBtnLoader(btn);
    DATA = DATA.filter(r => String(r[0] ?? '') !== String(idFila ?? ''));
    DATA_FILTRADA = DATA_FILTRADA.filter(r => String(r[0] ?? '') !== String(idFila ?? ''));
    actualizarIndiceBusquedaUbicaciones();
    renderTabla(DATA_FILTRADA.length ? DATA_FILTRADA : DATA);
    estadoSincronizacionManual('Registro eliminado. La BD no se vuelve a consultar hasta presionar Recargar.');
  })
  .catch(()=>{ endBtnLoader(btn); alert('Error al eliminar'); });
}

/* =====================================================
   GUARDAR
===================================================== */
function guardar(){
  const btn = $('btnGuardar');

  const stockActual = Number($('cantidad').value || 0);
  const mov = Number($('cantidad_mov').value || 0);

  if(!TIPO_MOV){
    alert('Seleccione ENTRADA o SALIDA');
    return;
  }

  if(mov <= 0){
    alert('Ingrese una cantidad válida');
    return;
  }

  let nuevoStock = stockActual;

  if(TIPO_MOV === 'SALIDA'){
    if(mov > stockActual){
      alert(`Stock insuficiente\nStock: ${stockActual}\nSalida: ${mov}`);
      return;
    }
    nuevoStock -= mov;
  }else{
    nuevoStock += mov;
  }

  startBtnLoader(btn);

  fetch(URL_GS,{
    method:'POST',
    body:JSON.stringify({
      accion: $('id').value ? 'editar_movimiento' : 'guardar_movimiento',
      id: $('id').value,
      fecha_entrada: $('fecha_entrada').value,
      fecha_salida: $('fecha_salida').value,
      ubicacion: $('ubicacion').value,
      codigo: $('codigo').value,
      descripcion: $('descripcion').value,
      cantidad: nuevoStock,
      responsable: $('responsable').value,
      status: $('status').value,
      origen: ORIGEN
    })
  })
  .then(r=>r.json())
  .then(res=>{
    endBtnLoader(btn);
    if(res.ok === false){
      alert(res.msg || 'Error');
      return;
    }
    cerrarModal();
    estadoSincronizacionManual('Registro guardado. La vista no consulta nuevamente la BD hasta presionar Recargar.');
  })
  .catch(()=>{
    endBtnLoader(btn);
    alert('Error al guardar');
  });
}

/* =====================================================
   FILTRO
===================================================== */
function filtrar(txt){
  const q = normalizarBusqueda(txt);
  clearTimeout(timerFiltroUbicaciones);
  timerFiltroUbicaciones = setTimeout(()=>{
    if(!q){
      DATA_FILTRADA = DATA;
    }else{
      DATA_FILTRADA = DATA.filter((r, i) => (DATA_SEARCH[i] || '').includes(q));
    }
    renderTabla(DATA_FILTRADA);
  }, 160);
}

/* =====================================================
   LIMPIAR
===================================================== */
function limpiarFormulario(){
  document.querySelectorAll('#modal input, #modal select')
    .forEach(i=>i.value='');
  TIPO_MOV = null;
  $('suggest').style.display='none';
}

/* =====================================================
   SCANNER
===================================================== */
let scanner = null;
let torchOn = false;

function abrirScanner(){
  if(!/android|iphone|ipad|mobile/i.test(navigator.userAgent)){
    alert('Scanner solo disponible en móvil');
    return;
  }

  $('scannerBox').style.display='block';
  $('torchBtn').style.display='block';

  scanner = new Html5Qrcode('scannerBox');
  scanner.start(
    {facingMode:{exact:'environment'}},
    {fps:10,qrbox:220},
    txt=>{
      $('codigo').value = txt.trim();
      cerrarScanner();
      buscarCodigo();
    }
  );
}

function toggleTorch(){
  if(!scanner) return;
  torchOn = !torchOn;
  scanner.applyVideoConstraints({advanced:[{torch:torchOn}]});
  $('torchBtn').classList.toggle('active',torchOn);
}

function cerrarScanner(){
  if(scanner){
    scanner.stop().then(()=>scanner.clear()).catch(()=>{});
    scanner = null;
  }
  $('scannerBox').style.display='none';
  $('torchBtn').style.display='none';
  torchOn = false;
}

   
function recargar(){
  return cargar();
}


/* =====================================================
   IMPORTAR ARCHIVOS XLSX / CSV A BD-UBICACIONES
   Flujo: descargar preformato -> leer archivo -> previsualizar -> confirmar.
   Usa importación masiva por Apps Script mediante importar_movimiento.
===================================================== */
let IMPORT_PREVIEW_ROWS = [];
let IMPORT_PREVIEW_VALIDAS = [];
let IMPORT_PREVIEW_FILE_NAME = '';
let IMPORT_PREVIEW_SHEET_NAME = '';
let IMPORT_PREVIEW_TOTAL_ARCHIVO = 0;
const IMPORT_PREVIEW_MAX_RENDER = 350;

function abrirImportadorUbicaciones(){
  const input = $('inputImportUbicaciones');
  if(!input){ alert('No se encontró el selector de archivo.'); return; }
  input.value = '';
  input.click();
}

function descargarFormatoUbicaciones(){
  if(!window.XLSX){
    const csv = generarCsvPreformatoUbicaciones();
    descargarBlob('Preformato_Importar_Ubicaciones.csv', csv, 'text/csv;charset=utf-8');
    return;
  }

  const headers = [
    'CODIGO',
    'DESCRIPCION',
    'UBICACION',
    'CANTIDAD',
    'RESPONSABLE',
    'STATUS',
    'FECHA_ENTRADA',
    'FECHA_SALIDA',
    'ORIGEN'
  ];

  const ejemplo = [
    {
      CODIGO:'000123456789',
      DESCRIPCION:'PRODUCTO EJEMPLO CON CÓDIGO COMO TEXTO',
      UBICACION:'A-01-01',
      CANTIDAD:10,
      RESPONSABLE:'Alejandro Silva',
      STATUS:'VIGENTE',
      FECHA_ENTRADA:todayIso(),
      FECHA_SALIDA:'',
      ORIGEN:'WEB'
    },
    {
      CODIGO:'000987654321',
      DESCRIPCION:'OTRO PRODUCTO EJEMPLO',
      UBICACION:'B-02-03',
      CANTIDAD:5,
      RESPONSABLE:'',
      STATUS:'VIGENTE',
      FECHA_ENTRADA:todayIso(),
      FECHA_SALIDA:'',
      ORIGEN:'WEB'
    }
  ];

  const instrucciones = [
    ['INSTRUCCIONES'],
    ['1. No cambies los nombres de las columnas.'],
    ['2. El campo CODIGO debe escribirse como texto para conservar ceros iniciales.'],
    ['3. Las columnas obligatorias son CODIGO y UBICACION.'],
    ['4. CANTIDAD acepta números enteros o decimales. No acepta valores negativos.'],
    ['5. STATUS recomendado: VIGENTE o RETIRADO.'],
    ['6. Antes de importar, el sistema mostrará una previsualización con válidas, errores y duplicadas.'],
    ['7. Al confirmar, el sistema enviará todas las líneas válidas en una sola carga masiva.']
  ];

  const ws = XLSX.utils.json_to_sheet(ejemplo, {header: headers});
  XLSX.utils.sheet_add_aoa(ws, [headers], {origin:'A1'});
  ws['!cols'] = [
    {wch:18},{wch:38},{wch:16},{wch:12},{wch:20},{wch:12},{wch:16},{wch:16},{wch:12}
  ];
  const wsInfo = XLSX.utils.aoa_to_sheet(instrucciones);
  wsInfo['!cols'] = [{wch:110}];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'BD-UBICACIONES');
  XLSX.utils.book_append_sheet(wb, wsInfo, 'INSTRUCCIONES');
  XLSX.writeFile(wb, 'Preformato_Importar_Ubicaciones.xlsx');
}

function generarCsvPreformatoUbicaciones(){
  const rows = [
    ['CODIGO','DESCRIPCION','UBICACION','CANTIDAD','RESPONSABLE','STATUS','FECHA_ENTRADA','FECHA_SALIDA','ORIGEN'],
    ['000123456789','PRODUCTO EJEMPLO CON CODIGO COMO TEXTO','A-01-01','10','Alejandro Silva','VIGENTE',todayIso(),'','WEB'],
    ['000987654321','OTRO PRODUCTO EJEMPLO','B-02-03','5','','VIGENTE',todayIso(),'','WEB']
  ];
  return rows.map(row => row.map(csvEscape).join(';')).join('\n');
}

function csvEscape(value){
  const raw = String(value ?? '');
  return /[";\n\r]/.test(raw) ? '"' + raw.replace(/"/g,'""') + '"' : raw;
}

function descargarBlob(filename, content, type){
  const blob = new Blob([content], {type});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{
    try{ document.body.removeChild(a); }catch(e){}
    try{ URL.revokeObjectURL(a.href); }catch(e){}
  }, 500);
}

function todayIso(){
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function normalizarHeaderImport(h){
  return String(h ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .trim().toLowerCase()
    .replace(/[\._-]+/g,' ')
    .replace(/\s+/g,' ');
}

function valorImport(row, aliases){
  const keys = Object.keys(row || {});
  for(const alias of aliases){
    const objetivo = normalizarHeaderImport(alias);
    const key = keys.find(k => normalizarHeaderImport(k) === objetivo);
    if(key && String(row[key] ?? '').trim() !== '') return String(row[key] ?? '').trim();
  }
  return '';
}

function numeroImport(v, def=0){
  const rawOriginal = String(v ?? '').trim();
  if(!rawOriginal) return def;

  // Si viene como 1.234,56 lo convierte a 1234.56.
  // Si viene como 1234.56 lo mantiene.
  let raw = rawOriginal;
  if(/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(raw)) raw = raw.replace(/\./g,'').replace(',', '.');
  else raw = raw.replace(',', '.');

  const n = Number(raw);
  return Number.isFinite(n) ? n : def;
}

function estadoVisibleImportacion(item){
  if(!item.valida) return '⚠️ Error';
  if(item.duplicada) return '🟡 Válida duplicada';
  return '✅ Válida';
}

function observacionImportacion(item){
  const obs = [];
  if(item.errores && item.errores.length) obs.push(...item.errores);
  if(item.duplicada) obs.push('Código + ubicación repetidos en el archivo');
  if(item.payload && Number(item.payload.cantidad) === 0) obs.push('Cantidad en 0');
  return obs.join(' · ') || 'Lista para importar';
}

function mapearFilaUbicacion(row, index=0){
  const codigo = normalizarCodigo(valorImport(row, ['codigo','código','cod','sku','producto','item','codigo producto','código producto']));
  const descripcion = valorImport(row, ['descripcion','descripción','detalle','nombre','producto descripcion','producto descripción','nombre producto']);
  const ubicacion = valorImport(row, ['ubicacion','ubicación','bodega','posicion','posición','location','ubicacion producto','ubicación producto']);
  const cantidadRaw = valorImport(row, ['cantidad','cant','stock','unidades','saldo','cantidad actual','stock actual']);
  const cantidad = numeroImport(cantidadRaw, 0);
  const status = (valorImport(row, ['status','estado']) || 'VIGENTE').toUpperCase();

  const payload = {
    accion: 'guardar_movimiento',
    fecha_entrada: valorImport(row, ['fecha entrada','fecha_entrada','entrada','fecha ingreso','fecha']),
    fecha_salida: valorImport(row, ['fecha salida','fecha_salida','salida']),
    ubicacion,
    codigo,
    descripcion,
    cantidad,
    responsable: valorImport(row, ['responsable','operador','usuario','encargado']),
    status,
    origen: valorImport(row, ['origen','fuente']) || ORIGEN
  };

  const errores = [];
  if(!codigo) errores.push('Falta código');
  if(!ubicacion) errores.push('Falta ubicación');
  if(cantidad < 0) errores.push('Cantidad negativa');
  if(status && !['VIGENTE','RETIRADO','ACTIVO','INACTIVO'].includes(status)) errores.push('Status no reconocido');

  return {
    fila: index + 2,
    valida: errores.length === 0,
    errores,
    duplicada:false,
    clave: codigo && ubicacion ? `${codigo}||${normalizarHeaderImport(ubicacion)}` : '',
    payload
  };
}

function marcarDuplicadasImportacion(rows){
  const conteo = {};
  rows.forEach(item => {
    if(!item.clave) return;
    conteo[item.clave] = (conteo[item.clave] || 0) + 1;
  });
  rows.forEach(item => { item.duplicada = !!item.clave && conteo[item.clave] > 1; });
  return rows;
}

async function leerArchivoImportacion(file){
  const ext = String(file?.name || '').split('.').pop().toLowerCase();
  if(['csv','txt'].includes(ext)){
    const text = await file.text();
    return XLSX.read(text, {type:'string', raw:false});
  }
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, {type:'array', cellDates:false, raw:false});
}

async function importarArchivoUbicaciones(file){
  if(!file) return;
  if(!window.XLSX){ alert('No está disponible el lector Excel/CSV.'); return; }

  startProgress();
  try{
    const wb = await leerArchivoImportacion(file);
    const sheetName = wb.SheetNames[0];
    if(!sheetName) throw new Error('El archivo no contiene hojas válidas.');

    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, {defval:'', raw:false, blankrows:false});
    const previewRows = marcarDuplicadasImportacion(rows.map((row, index) => mapearFilaUbicacion(row, index)));
    const validas = previewRows.filter(x => x.valida).map(x => x.payload);

    if(!previewRows.length){
      alert('El archivo está vacío o no contiene encabezados válidos.');
      return;
    }

    IMPORT_PREVIEW_ROWS = previewRows;
    IMPORT_PREVIEW_VALIDAS = validas;
    IMPORT_PREVIEW_FILE_NAME = file.name || 'archivo seleccionado';
    IMPORT_PREVIEW_SHEET_NAME = sheetName;
    IMPORT_PREVIEW_TOTAL_ARCHIVO = previewRows.length;

    const filtro = $('importPreviewFiltro');
    const buscar = $('importPreviewBuscar');
    if(filtro) filtro.value = 'todas';
    if(buscar) buscar.value = '';

    renderPreviewImportacionUbicaciones();
  }catch(err){
    alert('Error al leer archivo: ' + (err.message || err));
  }finally{
    endProgress();
    const input = $('inputImportUbicaciones');
    if(input) input.value = '';
  }
}

function filtrarPreviewImportacionUbicaciones(){
  const filtro = String($('importPreviewFiltro')?.value || 'todas');
  const q = normalizarHeaderImport($('importPreviewBuscar')?.value || '');
  return IMPORT_PREVIEW_ROWS.filter(item => {
    if(filtro === 'validas' && !item.valida) return false;
    if(filtro === 'errores' && item.valida) return false;
    if(filtro === 'duplicadas' && !item.duplicada) return false;
    if(q){
      const p = item.payload || {};
      const texto = normalizarHeaderImport([
        item.fila,
        estadoVisibleImportacion(item),
        observacionImportacion(item),
        p.codigo,
        p.descripcion,
        p.ubicacion,
        p.cantidad,
        p.responsable,
        p.status,
        p.fecha_entrada,
        p.fecha_salida
      ].join(' '));
      if(!texto.includes(q)) return false;
    }
    return true;
  });
}

function statPreviewHtml(titulo, valor, detalle){
  return `<div style="border:1px solid #e2e8f0;border-radius:16px;padding:12px;background:#f8fafc">
    <div style="font-size:11px;color:#64748b;font-weight:900;text-transform:uppercase;letter-spacing:.05em">${escapeHtml(titulo)}</div>
    <div style="font-size:24px;font-weight:1000;margin-top:4px">${escapeHtml(valor)}</div>
    <div style="font-size:12px;color:#64748b;margin-top:2px">${escapeHtml(detalle || '')}</div>
  </div>`;
}

function renderPreviewImportacionUbicaciones(){
  const modal = $('importPreviewModal');
  const resumen = $('importPreviewResumen');
  const stats = $('importPreviewStats');
  const tabla = $('importPreviewTabla');
  const nota = $('importPreviewNota');
  const btn = $('btnConfirmarImportUbicaciones');
  if(!modal || !resumen || !tabla) return;

  const total = IMPORT_PREVIEW_ROWS.length;
  const validas = IMPORT_PREVIEW_ROWS.filter(x => x.valida).length;
  const errores = IMPORT_PREVIEW_ROWS.filter(x => !x.valida).length;
  const duplicadas = IMPORT_PREVIEW_ROWS.filter(x => x.duplicada).length;
  const cantidadTotal = IMPORT_PREVIEW_ROWS
    .filter(x => x.valida)
    .reduce((acc, x) => acc + (Number(x.payload?.cantidad) || 0), 0);
  const filtradas = filtrarPreviewImportacionUbicaciones();
  const visibles = filtradas.slice(0, IMPORT_PREVIEW_MAX_RENDER);

  resumen.innerHTML = `Archivo: <b>${escapeHtml(IMPORT_PREVIEW_FILE_NAME)}</b>` +
    (IMPORT_PREVIEW_SHEET_NAME ? ` · Hoja leída: <b>${escapeHtml(IMPORT_PREVIEW_SHEET_NAME)}</b>` : '') + `<br>` +
    `<b>Revisión previa:</b> aquí se muestra exactamente lo que se enviará a BD-MOVIMIENTO. ` +
    `Al confirmar, el sistema hará una sola carga masiva, no una solicitud por fila.`;

  if(stats){
    stats.innerHTML = [
      statPreviewHtml('Detectadas', total, 'filas leídas'),
      statPreviewHtml('Válidas', validas, 'se importarán'),
      statPreviewHtml('Con errores', errores, 'no se enviarán'),
      statPreviewHtml('Duplicadas', duplicadas, 'repetidas en archivo'),
      statPreviewHtml('Cantidad total', cantidadTotal, 'solo filas válidas')
    ].join('');
  }

  tabla.innerHTML = '';
  visibles.forEach(item => {
    const p = item.payload || {};
    const tr = document.createElement('tr');
    tr.style.background = item.valida ? '#ffffff' : '#fff7ed';
    tr.innerHTML = `
      <td>${escapeHtml(item.fila)}</td>
      <td>${escapeHtml(estadoVisibleImportacion(item))}</td>
      <td>${escapeHtml(observacionImportacion(item))}</td>
      <td>${escapeHtml(p.codigo)}</td>
      <td>${escapeHtml(p.descripcion)}</td>
      <td>${escapeHtml(p.ubicacion)}</td>
      <td>${escapeHtml(p.cantidad)}</td>
      <td>${escapeHtml(p.responsable)}</td>
      <td>${escapeHtml(p.status)}</td>
      <td>${escapeHtml(p.fecha_entrada)}</td>
      <td>${escapeHtml(p.fecha_salida)}</td>`;
    tabla.appendChild(tr);
  });

  if(!filtradas.length){
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="11" style="text-align:center;color:#64748b">No hay filas para mostrar con el filtro actual.</td>`;
    tabla.appendChild(tr);
  }

  if(filtradas.length > IMPORT_PREVIEW_MAX_RENDER){
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="11" style="text-align:center;color:#64748b;font-weight:800">Vista rápida limitada a ${IMPORT_PREVIEW_MAX_RENDER} filas de ${filtradas.length} filtradas. Al confirmar se importarán todas las filas válidas: ${validas}.</td>`;
    tabla.appendChild(tr);
  }

  if(nota){
    nota.innerHTML = `Mostrando <b>${Math.min(filtradas.length, IMPORT_PREVIEW_MAX_RENDER)}</b> de <b>${filtradas.length}</b> filas filtradas. ` +
      `Se importarán <b>${validas}</b> filas válidas. Las filas con errores quedan fuera automáticamente.`;
  }

  if(btn){
    btn.disabled = validas === 0;
    btn.textContent = validas ? `Confirmar importación masiva (${validas} líneas)` : 'No hay líneas válidas para importar';
  }
  modal.classList.add('active');
}

function cerrarPreviewImportacionUbicaciones(){
  const modal = $('importPreviewModal');
  if(modal) modal.classList.remove('active');
}

function csvImportPreviewLine(values){
  return values.map(v => {
    const raw = String(v ?? '');
    return /[";\n\r]/.test(raw) ? '"' + raw.replace(/"/g,'""') + '"' : raw;
  }).join(';');
}

function descargarObservacionesImportacionUbicaciones(){
  if(!IMPORT_PREVIEW_ROWS.length){
    alert('No hay previsualización cargada.');
    return;
  }
  const rows = [
    ['FILA','ESTADO','OBSERVACION','CODIGO','DESCRIPCION','UBICACION','CANTIDAD','RESPONSABLE','STATUS','FECHA_ENTRADA','FECHA_SALIDA']
  ];
  IMPORT_PREVIEW_ROWS.forEach(item => {
    const p = item.payload || {};
    rows.push([
      item.fila,
      estadoVisibleImportacion(item),
      observacionImportacion(item),
      p.codigo,
      p.descripcion,
      p.ubicacion,
      p.cantidad,
      p.responsable,
      p.status,
      p.fecha_entrada,
      p.fecha_salida
    ]);
  });
  descargarBlob('Observaciones_Importacion_Ubicaciones.csv', rows.map(csvImportPreviewLine).join('\n'), 'text/csv;charset=utf-8');
}

async function confirmarImportacionUbicaciones(){
  const mapeadas = IMPORT_PREVIEW_VALIDAS.slice();
  if(!mapeadas.length){
    alert('No hay filas válidas para importar.');
    return;
  }

  const btn = $('btnConfirmarImportUbicaciones');
  const bar = $('progress-bar');
  startBtnLoader(btn);
  startProgress();
  if(bar) bar.style.width = '18%';

  try{
    // Limpia la acción individual para que el servidor reciba solo filas de datos.
    const rows = mapeadas.map(item => {
      const row = Object.assign({}, item);
      delete row.accion;
      row.origen = row.origen || ORIGEN;
      return row;
    });

    if(bar) bar.style.width = '42%';

    const res = await fetch(URL_GS, {
      method:'POST',
      body: JSON.stringify({
        accion:'importar_movimiento',
        modo:'append',
        origen:ORIGEN,
        rows
      })
    });

    if(bar) bar.style.width = '72%';

    const json = await res.json().catch(()=>null);
    if(!res.ok || !json){
      throw new Error('Apps Script no respondió con JSON válido.');
    }
    if(json.ok === false){
      throw new Error(json.msg || 'Apps Script rechazó la importación masiva.');
    }

    rows.forEach(payload => {
      if(payload.codigo && payload.descripcion){
        setProductoEnCache(payload.codigo, payload.descripcion, payload.cantidad);
      }
    });

    if(bar) bar.style.width = '88%';
    cerrarPreviewImportacionUbicaciones();
    estadoSincronizacionManual('Importación enviada correctamente. Presiona Recargar para consultar nuevamente BD-MOVIMIENTO.');

    const recibidos = Number(json.recibidos ?? rows.length) || rows.length;
    const insertados = Number(json.insertados || 0);
    const actualizados = Number(json.actualizados || 0);
    const omitidos = Number(json.omitidos || 0);
    const procesados = Number(json.procesados ?? (insertados + actualizados)) || (insertados + actualizados);
    const errores = Array.isArray(json.errores) ? json.errores : [];
    const completas = procesados === rows.length && omitidos === 0 && errores.length === 0;

    const resumen = [
      completas ? 'Importación confirmada: todas las líneas válidas fueron importadas correctamente.' : 'Importación finalizada con observaciones.',
      `Filas detectadas en previsualización: ${IMPORT_PREVIEW_TOTAL_ARCHIVO || IMPORT_PREVIEW_ROWS.length}`,
      `Filas válidas enviadas: ${rows.length}`,
      `Filas recibidas por Apps Script: ${recibidos}`,
      `Insertadas: ${insertados}`,
      `Actualizadas: ${actualizados}`,
      `Omitidas por Apps Script: ${omitidos}`,
      `Procesadas correctamente: ${procesados}`
    ];

    if(json.total !== undefined) resumen.push(`Total actual en BD-MOVIMIENTO: ${json.total}`);
    if(errores.length){
      resumen.push('');
      resumen.push('Primeras observaciones:');
      resumen.push(errores.slice(0,10).join('\n'));
    }

    alert(resumen.join('\n'));
  }catch(err){
    alert('Error al importar archivo: ' + (err.message || err));
  }finally{
    endBtnLoader(btn);
    endProgress();
  }
}


function exportarPDF(){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('l','pt','a4');

  const origen = DATA_FILTRADA.length ? DATA_FILTRADA : DATA;

  doc.text('Reporte de Ubicaciones',40,40);

  doc.autoTable({
    startY:60,
    head:[['Código','Descripción','Ubicación','Stock','Registro','Entrada','Salida','Responsable','Status','Origen']],
    body:origen.map(r=>[
      r[5], r[6], r[4], r[7],
      formatFechaTabla(r[1]),
      formatFechaTabla(r[2]),
      formatFechaTabla(r[3]),
      r[8], r[9], r[10]
    ]),
    styles:{ fontSize:9 },
    headStyles:{ fillColor:[20,184,166], textColor:255 }
  });

  abrirDialogoImpresionPdf(doc,'ubicaciones.pdf');
}

/* =====================================================
   EXPORTAR XLSX
===================================================== */
function exportarXLSX(){
  const origen = DATA_FILTRADA.length ? DATA_FILTRADA : DATA;

  const ws = XLSX.utils.json_to_sheet(
    origen.map(r=>({
      Codigo:r[5],
      Descripcion:r[6],
      Ubicacion:r[4],
      Stock:r[7],
      Registro:formatFechaTabla(r[1]),
      Entrada:formatFechaTabla(r[2]),
      Salida:formatFechaTabla(r[3]),
      Responsable:r[8],
      Status:r[9],
      Origen:r[10]
    }))
  );

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Ubicaciones');
  XLSX.writeFile(wb,'ubicaciones.xlsx');
}

function toggleCard(id, btn){
  CARD_STATE[id] = !CARD_STATE[id];

  const card = document.querySelector(`.card-item[data-id="${id}"]`);
  if(!card) return;

  const body = card.querySelector('.card-body');
  if(!body) return;

  body.style.display = CARD_STATE[id] ? 'block' : 'none';
  btn.textContent = CARD_STATE[id] ? '−' : '+';
}

/* =====================================================
   INICIO
===================================================== */
document.addEventListener('DOMContentLoaded', async ()=>{
  await cargarConfiguracionUrl();
  cargarCatalogoDesdeMemoria();
  DATA = [];
  DATA_FILTRADA = [];
  actualizarIndiceBusquedaUbicaciones();
  renderTabla([]);
  estadoSincronizacionManual('Sin sincronización automática. Presiona Recargar para consultar BD-MOVIMIENTO.');
});

let timerResizeUbicaciones = null;
window.addEventListener('resize', ()=>{
  clearTimeout(timerResizeUbicaciones);
  timerResizeUbicaciones = setTimeout(()=>renderTabla(DATA_FILTRADA.length ? DATA_FILTRADA : DATA), 180);
});
