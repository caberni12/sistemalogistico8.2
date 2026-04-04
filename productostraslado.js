/***************************************************
 CONSULTA RÁPIDA DE PRODUCTO POR CÓDIGO / SELECT
 - Mantiene API propia de productos
 - Precarga catálogo para respuesta inmediata
 - Completa producto + descripción
 - Si no encuentra en caché, consulta backend
***************************************************/
/* API de productos separada de la API de pedidos. Si desea cambiarla, defina window.API_PRODUCTOS antes de cargar este archivo. */
const API_PRODUCTOS = (window.API_PRODUCTOS || "https://script.google.com/macros/s/AKfycbxzSkxz-rVSMLBEGy7k0FPd1EJpfeufZXEzxzf3JXOAQ7ONJ8O3tpxkTXYzdwDbjb7s/exec");
const PRODUCT_CACHE_KEY = "catalogo_producto_rapido_v2";
const PRODUCT_CACHE_TTL_MS = 6 * 60 * 60 * 1000;

let CONSULTANDO_PRODUCTO = false;
let ULTIMO_CODIGO = "";
let PRODUCTOS_CACHE = [];
let PRODUCTOS_CACHE_CARGADO = false;
let PRODUCTOS_CACHE_PROMISE = null;

function el(id){ return document.getElementById(id); }
function txt(v){ return String(v ?? "").trim(); }
function limpiarCodigo(v){ return String(v ?? "").trim().replace(/\s+/g, ""); }
function normalizarProducto(v){
  return String(v ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function setMensajeCodigo(msg, color = "#64748b"){
  const box = el("msgCodigoProducto");
  if(box){
    box.textContent = msg;
    box.style.color = color;
  }
}

function setLoadingBoton(btn, estado){
  if(!btn) return;
  if(estado){
    btn.disabled = true;
    btn.dataset.txtOriginal = btn.textContent;
    btn.textContent = "Buscando...";
    btn.classList.add("btn-loading");
  }else{
    btn.disabled = false;
    btn.textContent = btn.dataset.txtOriginal || "🔎 Consultar";
    btn.classList.remove("btn-loading");
  }
}

function limpiarCamposProductoCodigo(){
  const pProducto = el("pProducto");
  const pDetalle = el("pDetalle");
  if(pProducto) pProducto.value = "";
  if(pDetalle) pDetalle.value = "";
}

function guardarCatalogoLocal(items){
  try{
    localStorage.setItem(PRODUCT_CACHE_KEY, JSON.stringify({
      at: Date.now(),
      items: items || []
    }));
  }catch(err){}
}

function leerCatalogoLocal(){
  try{
    const raw = localStorage.getItem(PRODUCT_CACHE_KEY);
    if(!raw) return [];
    const parsed = JSON.parse(raw);
    if(!parsed || !Array.isArray(parsed.items)) return [];
    if(Date.now() - Number(parsed.at || 0) > PRODUCT_CACHE_TTL_MS) return parsed.items;
    return parsed.items;
  }catch(err){
    return [];
  }
}

function dedupeCatalogo(items){
  const out = [];
  const seen = new Set();
  (items || []).forEach(item => {
    const codigo = limpiarCodigo(item.codigo || item.producto || "");
    const descripcion = txt(item.descripcion || item.detalle || "");
    if(!codigo && !descripcion) return;
    const key = `${normalizarProducto(codigo)}|${normalizarProducto(descripcion)}`;
    if(seen.has(key)) return;
    seen.add(key);
    out.push({ codigo, producto: codigo, descripcion });
  });
  return out;
}

function setCatalogo(items){
  PRODUCTOS_CACHE = dedupeCatalogo(items).sort((a,b) => a.codigo.localeCompare(b.codigo, 'es'));
  PRODUCTOS_CACHE_CARGADO = true;
  renderCatalogoSelect(PRODUCTOS_CACHE);
  renderCatalogoDatalist(PRODUCTOS_CACHE);
  guardarCatalogoLocal(PRODUCTOS_CACHE);
}

function renderCatalogoSelect(items){
  let select = el("selectCodigoProducto");
  if(!select) return;
  const top = (items || []).slice(0, 500);
  select.innerHTML = '<option value="">Seleccionar producto rápido...</option>' + top.map(item => {
    const codigo = item.codigo || item.producto || "";
    const desc = item.descripcion || "";
    return `<option value="${codigo.replace(/"/g,'&quot;')}" data-desc="${desc.replace(/"/g,'&quot;')}">${codigo} - ${desc}</option>`;
  }).join("");
}

function renderCatalogoDatalist(items){
  let list = el("listaCodigosProducto");
  if(!list) return;
  list.innerHTML = (items || []).slice(0, 1000).map(item => {
    const codigo = item.codigo || item.producto || "";
    const desc = item.descripcion || "";
    return `<option value="${codigo.replace(/"/g,'&quot;')}">${desc}</option>`;
  }).join("");
}

function buscarEnCatalogoLocal(termino){
  const q = normalizarProducto(termino);
  if(!q) return null;
  for(const item of PRODUCTOS_CACHE){
    if(normalizarProducto(item.codigo) === q || normalizarProducto(item.producto) === q){
      return item;
    }
  }
  return null;
}

function filtrarCatalogoLocal(termino, limit = 80){
  const q = normalizarProducto(termino);
  if(!q) return PRODUCTOS_CACHE.slice(0, limit);
  const exactos = [];
  const prefijos = [];
  const contiene = [];
  for(const item of PRODUCTOS_CACHE){
    const codigo = normalizarProducto(item.codigo);
    const descripcion = normalizarProducto(item.descripcion);
    if(codigo === q || descripcion === q) exactos.push(item);
    else if(codigo.startsWith(q) || descripcion.startsWith(q)) prefijos.push(item);
    else if(codigo.includes(q) || descripcion.includes(q)) contiene.push(item);
    if(exactos.length + prefijos.length + contiene.length >= limit) break;
  }
  return exactos.concat(prefijos, contiene).slice(0, limit);
}

async function cargarCatalogoProducto(force = false){
  if(PRODUCTOS_CACHE_CARGADO && !force) return PRODUCTOS_CACHE;
  if(PRODUCTOS_CACHE_PROMISE && !force) return PRODUCTOS_CACHE_PROMISE;

  const local = leerCatalogoLocal();
  if(local.length && !force){
    setCatalogo(local);
  }

  const mapearLista = (data) => {
    let lista = [];
    if(data && Array.isArray(data.data)) lista = data.data;
    else if(Array.isArray(data)) lista = data;
    return lista.map(row => ({
      codigo: limpiarCodigo(Array.isArray(row) ? row[0] : (row.codigo || row.CODIGO || row.producto || row.PRODUCTO || "")),
      producto: limpiarCodigo(Array.isArray(row) ? row[0] : (row.codigo || row.CODIGO || row.producto || row.PRODUCTO || "")),
      descripcion: txt(Array.isArray(row) ? row[1] : (row.descripcion || row.DESCRIPCION || row.detalle || row.DETALLE || ""))
    })).filter(item => item.codigo || item.descripcion);
  };

  PRODUCTOS_CACHE_PROMISE = (async () => {
    try{
      let lista = [];

      try{
        const res = await fetch(`${API_PRODUCTOS}?action=catalogoProductos&_=${Date.now()}`, { method: 'GET', cache: 'no-store' });
        const data = await res.json();
        lista = mapearLista(data);
      }catch(err){}

      if(!lista.length){
        try{
          const resPost = await fetch(API_PRODUCTOS, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ action: "catalogoProductos" }),
            cache: "no-store"
          });
          const dataPost = await resPost.json();
          lista = mapearLista(dataPost);
        }catch(err){}
      }

      if(!lista.length){
        try{
          const resCompat = await fetch(`${API_PRODUCTOS}?accion=listar&_=${Date.now()}`, { method: 'GET', cache: 'no-store' });
          const dataCompat = await resCompat.json();
          lista = mapearLista(dataCompat);
        }catch(err){}
      }

      if(lista.length){
        setCatalogo(lista);
      } else if(local.length) {
        setCatalogo(local);
      }
      return PRODUCTOS_CACHE;
    }catch(err){
      if(local.length){
        setCatalogo(local);
        return PRODUCTOS_CACHE;
      }
      console.error('No se pudo cargar catálogo rápido de productos:', err);
      return PRODUCTOS_CACHE;
    }finally{
      PRODUCTOS_CACHE_PROMISE = null;
    }
  })();

  return PRODUCTOS_CACHE_PROMISE;
}

function aplicarProductoEncontrado(item){
  const pProducto = el("pProducto");
  const pDetalle = el("pDetalle");
  const pCantidad = el("pCantidad");
  const pCodigo = el("pCodigo");
  if(!item) return false;
  const codigo = limpiarCodigo(item.codigo || item.producto || "");
  const descripcion = txt(item.descripcion || item.detalle || "");
  if(pCodigo) pCodigo.value = codigo;
  if(pProducto) pProducto.value = codigo;
  if(pDetalle) pDetalle.value = descripcion;
  if(pCantidad && !txt(pCantidad.value)) pCantidad.value = "1";
  ULTIMO_CODIGO = codigo;
  setMensajeCodigo("Producto cargado correctamente.", "#16a34a");
  return true;
}

function crearCampoCodigoSiNoExiste(){
  const modal = el("trasladoModal");
  const productoAdd = modal ? modal.querySelector(".producto-add") : null;
  if(!modal || !productoAdd) return;
  if(el("grupoCodigoProducto")) return;

  const bloque = document.createElement("div");
  bloque.className = "codigo-producto-wrap";
  bloque.id = "grupoCodigoProducto";
  bloque.style.margin = "10px 0 14px 0";
  bloque.innerHTML = `
    <label>Código rápido</label>
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <input id="pCodigo" type="text" inputmode="text" placeholder="Ingrese código del producto" autocomplete="off" list="listaCodigosProducto" style="flex:1 1 220px;">
      <button type="button" id="btnBuscarCodigoProducto" style="white-space:nowrap;">🔎 Consultar</button>
    </div>
    <datalist id="listaCodigosProducto"></datalist>
    <select id="selectCodigoProducto" style="margin-top:8px;width:100%;"></select>
    <div id="msgCodigoProducto" style="font-size:12px;color:#64748b;margin-top:6px;"></div>
    <div id="pDetalleEstado" style="font-size:12px;color:#64748b;margin-top:4px;"></div>
  `;

  productoAdd.parentNode.insertBefore(bloque, productoAdd);

  const inputCodigo = el("pCodigo");
  const btnBuscar = el("btnBuscarCodigoProducto");
  const select = el("selectCodigoProducto");

  if(inputCodigo){
    inputCodigo.addEventListener("keydown", function(e){
      if(e.key === "Enter"){
        e.preventDefault();
        consultarCodigoProducto();
      }
    });

    inputCodigo.addEventListener("input", function(){
      ULTIMO_CODIGO = "";
      const code = limpiarCodigo(this.value);
      if(!code){
        setMensajeCodigo("");
        const estado = el("pDetalleEstado");
        if(estado) estado.textContent = "";
        renderCatalogoSelect(filtrarCatalogoLocal(""));
        return;
      }
      const local = buscarEnCatalogoLocal(code);
      if(local){
        aplicarProductoEncontrado(local);
      }else{
        setMensajeCodigo("");
      }
      renderCatalogoSelect(filtrarCatalogoLocal(code));
    });

    inputCodigo.addEventListener("focus", function(){
      renderCatalogoSelect(filtrarCatalogoLocal(this.value || ""));
    });
  }

  if(btnBuscar){
    btnBuscar.addEventListener("click", consultarCodigoProducto);
  }

  if(select){
    select.addEventListener("change", function(){
      const codigo = limpiarCodigo(this.value);
      if(!codigo) return;
      const item = buscarEnCatalogoLocal(codigo);
      if(item){
        aplicarProductoEncontrado(item);
        const pCantidad = el("pCantidad");
        if(pCantidad){
          pCantidad.focus();
          if(typeof pCantidad.select === "function") pCantidad.select();
        }
      }
    });
  }
}

async function consultarCodigoProducto(){
  if(CONSULTANDO_PRODUCTO) return;

  const inputCodigo = el("pCodigo");
  const btnBuscar = el("btnBuscarCodigoProducto");
  const pCantidad = el("pCantidad");
  if(!inputCodigo){
    console.error("No existe #pCodigo");
    return;
  }

  const codigo = limpiarCodigo(inputCodigo.value);
  if(!codigo){
    limpiarCamposProductoCodigo();
    setMensajeCodigo("Ingrese un código.", "#dc2626");
    inputCodigo.focus();
    return;
  }

  const local = buscarEnCatalogoLocal(codigo);
  if(local){
    aplicarProductoEncontrado(local);
    if(pCantidad){ pCantidad.focus(); pCantidad.select?.(); }
    return;
  }

  if(codigo === ULTIMO_CODIGO && txt(el("pDetalle")?.value)){ return; }

  CONSULTANDO_PRODUCTO = true;
  setLoadingBoton(btnBuscar, true);
  setMensajeCodigo("Consultando código...", "#2563eb");

  const mapearItem = (payload) => {
    const raw = payload?.data || payload;
    if(!raw) return null;
    const codigoMap = limpiarCodigo(raw.codigo || raw.CODIGO || raw.producto || raw.PRODUCTO || codigo);
    const descripcionMap = txt(raw.descripcion || raw.DESCRIPCION || raw.detalle || raw.DETALLE || "");
    if(!codigoMap && !descripcionMap) return null;
    return { codigo: codigoMap, producto: codigoMap, descripcion: descripcionMap };
  };

  try{
    let item = null;

    try{
      const url = `${API_PRODUCTOS}?action=buscarProductoExacto&codigo=${encodeURIComponent(codigo)}&_=${Date.now()}`;
      const res = await fetch(url, { method: "GET", cache: "no-store" });
      const data = await res.json();
      if(data?.ok) item = mapearItem(data);
    }catch(err){}

    if(!item){
      try{
        const resPost = await fetch(API_PRODUCTOS, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ action: "buscarProductoExacto", codigo, q: codigo, producto: codigo }),
          cache: "no-store"
        });
        const dataPost = await resPost.json();
        if(dataPost?.ok) item = mapearItem(dataPost);
      }catch(err){}
    }

    if(!item){
      try{
        const urlCompat = `${API_PRODUCTOS}?accion=buscar&codigo=${encodeURIComponent(codigo)}&_=${Date.now()}`;
        const resCompat = await fetch(urlCompat, { method: "GET", cache: "no-store" });
        const dataCompat = await resCompat.json();
        if(dataCompat?.ok) item = mapearItem(dataCompat);
      }catch(err){}
    }

    if(!item || !item.descripcion){
      limpiarCamposProductoCodigo();
      setMensajeCodigo("Código no encontrado.", "#dc2626");
      return;
    }

    PRODUCTOS_CACHE.unshift(item);
    setCatalogo(PRODUCTOS_CACHE);
    aplicarProductoEncontrado(item);
    if(pCantidad){ pCantidad.focus(); pCantidad.select?.(); }
  }catch(error){
    console.error("Error consulta producto:", error);
    limpiarCamposProductoCodigo();
    setMensajeCodigo(error.message || "No se pudo consultar el código.", "#dc2626");
  }finally{
    CONSULTANDO_PRODUCTO = false;
    setLoadingBoton(btnBuscar, false);
  }
}

function initConsultaProductoCodigo(){
  crearCampoCodigoSiNoExiste();
  cargarCatalogoProducto(false);
}

function refrescarConsultaProductoRapida(){
  crearCampoCodigoSiNoExiste();
  cargarCatalogoProducto(false);
}

if(document.readyState === "loading"){
  document.addEventListener("DOMContentLoaded", initConsultaProductoCodigo);
}else{
  initConsultaProductoCodigo();
}
window.consultarCodigoProducto = consultarCodigoProducto;
window.initConsultaProductoCodigo = initConsultaProductoCodigo;
window.refrescarConsultaProductoRapida = refrescarConsultaProductoRapida;
window.cargarCatalogoProducto = cargarCatalogoProducto;
