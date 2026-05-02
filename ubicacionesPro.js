let CARD_STATE = {}; // idFila => true (expandida) | false (colapsada)
/* =====================================================
   CONFIGURACIÓN
===================================================== */
const DEFAULT_URL_GS =
  'https://script.google.com/macros/s/AKfycbwOG9W7fdVqDrhacHp3Ry1A-pM5eKVFXwFXl4Q2V2SYS1uclU9Ko7XcqS5iOcP2BKEb7g/exec';
const API_URL_STORAGE_KEY = 'sistema_pedidos_api_url';
let URL_GS = '';

let DATA = [];
let DATA_FILTRADA = [];
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
  const saved = localStorage.getItem(API_URL_STORAGE_KEY);
  if(saved && saved.trim()){
    URL_GS = saved.trim();
  }else{
    try{
      const res = await fetch('config.txt?ts=' + Date.now(), { cache:'no-store' });
      if(res.ok){
        const urlTxt = parseConfigText(await res.text());
        if(urlTxt){
          URL_GS = urlTxt;
          localStorage.setItem(API_URL_STORAGE_KEY, URL_GS);
        }
      }
    }catch(err){
      // En file:// algunos navegadores bloquean config.txt.
    }
  }
  if(!URL_GS) URL_GS = DEFAULT_URL_GS;
  const input = $('configUrlGs');
  if(input) input.value = URL_GS;
  const status = $('configStatus');
  if(status) status.textContent = 'URL activa: ' + URL_GS;
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
    precargarCatalogo(true);
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
  fetch(`${URL_GS}?accion=listar`)
    .then(r=>r.json())
    .then(d=>{
      DATA = d.data || [];
      DATA_FILTRADA = DATA;
      renderTabla(DATA);
      endProgress();
    })
    .catch(()=>{
      DATA = [];
      DATA_FILTRADA = [];
      renderTabla([]);
      endProgress();
    });
}

function renderTabla(arr){
  const tbody = $('tabla');
  const cards = $('cards');

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

  arr.forEach(r=>{

    /* ===== TABLA ESCRITORIO ===== */
    tbody.innerHTML += `
      <tr>
        <td>${r[5]}</td>
        <td>${r[6]}</td>
        <td>${r[4]}</td>
        <td>${r[7]}</td>
        <td>${formatFechaTabla(r[1])}</td>
        <td>${formatFechaTabla(r[2])}</td>
        <td>${formatFechaTabla(r[3])}</td>
        <td>${r[8]}</td>
        <td>${r[9]}</td>
        <td>${r[10]}</td>
        <td class="actions-td">
          <button class="edit" onclick='editar(${JSON.stringify(r)})'>✏️</button>
          <button class="del" onclick='eliminar("${r[0]}",this)'>🗑️</button>
        </td>
      </tr>`;
/* ===== TARJETA MÓVIL ===== */
const id = r[0];
const open = CARD_STATE[id] === true;

cards.innerHTML += `
  <div class="card-item" data-id="${id}">

    <div class="card-head">
      <div class="desc">${r[6]}</div>
      <button onclick="toggleCard('${id}', this)">
        ${open ? '−' : '+'}
      </button>
    </div>

    <!-- VISIBLE POR DEFECTO -->
    <div class="card-row"><b>Código</b><span>${r[5]}</span></div>
    <div class="card-row"><b>Ubicación</b><span>${r[4]}</span></div>
    <div class="card-row"><b>Stock</b><span>${r[7]}</span></div>

    <!-- OCULTO / EXPANDIBLE -->
    <div class="card-body" style="display:${open ? 'block' : 'none'}">

      <div class="card-row"><b>Responsable</b><span>${r[8]}</span></div>
      <div class="card-row"><b>Origen</b><span>${r[10]}</span></div>

      <div class="card-fechas">
        <div>📥 ${formatFechaTabla(r[2])}</div>
        <div>📤 ${formatFechaTabla(r[3])}</div>
      </div>

      <span class="badge ${r[9] === 'VIGENTE' ? 'vigente' : 'retirado'}">
        ${r[9]}
      </span>

      <div class="card-actions">
        <button class="edit" onclick='editar(${JSON.stringify(r)})'>✏️</button>
        <button class="del" onclick='eliminar("${r[0]}",this)'>🗑️</button>
      </div>

    </div>

  </div>`;
  });
}

/* =====================================================
   EDITAR / ELIMINAR
===================================================== */
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
    body:JSON.stringify({accion:'eliminar',id:idFila})
  })
  .then(()=>{ endBtnLoader(btn); cargar(); })
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
      accion: $('id').value ? 'editar' : 'agregar',
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
    cargar();
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
  txt = txt.toLowerCase();
  DATA_FILTRADA = DATA.filter(r =>
    r.join(' ').toLowerCase().includes(txt)
  );
  renderTabla(DATA_FILTRADA);
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
  cargar();
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

  doc.save('ubicaciones.pdf');
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
  precargarCatalogo(false); // sincroniza catálogo en segundo plano sólo si falta o está vencido
  cargar();
});
