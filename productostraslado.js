/***************************************************
 CONSULTA DE PRODUCTO POR CÓDIGO
***************************************************/
const API_PRODUCTOS = "https://script.google.com/macros/s/AKfycbz8IUHeDvRiDUvlq_JIrTY1Rb0ZeMVhRt7al_V3NacKfWAMmK-J7Vngjr-hZ19o0woOaQ/exec";

let CONSULTANDO_PRODUCTO = false;
let ULTIMO_CODIGO = "";
let TIMER_CODIGO = null;

function el(id){ return document.getElementById(id); }
function txt(v){ return String(v ?? "").trim(); }
function limpiarCodigo(v){ return String(v ?? "").trim().replace(/\s+/g, ""); }
function setMensajeCodigo(msg, color = "#64748b"){
  const box = el("msgCodigoProducto");
  if(box){ box.textContent = msg || ""; box.style.color = color; }
}
function setLoadingBoton(btn, estado){
  if(!btn) return;
  if(estado){
    btn.disabled = true;
    btn.classList.add('loading');
  }else{
    btn.disabled = false;
    btn.classList.remove('loading');
  }
}
function limpiarCamposProductoSoloConsulta(){
  const pProducto = el("pProducto");
  const pDescripcion = el("pDescripcion");
  if(pProducto) pProducto.value = "";
  if(pDescripcion) pDescripcion.value = "";
}

function crearCampoCodigoSiNoExiste(){
  const modalProducto = el("modalProducto");
  const pProducto = el("pProducto");
  if(!modalProducto || !pProducto || el("pCodigo")) return;

  const grupoProducto = pProducto.closest(".form-group") || pProducto.parentNode;
  if(!grupoProducto || !grupoProducto.parentNode) return;

  const bloque = document.createElement("div");
  bloque.className = "form-group";
  bloque.id = "grupoCodigoProducto";
  bloque.innerHTML = `
    <label>Código</label>
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <input id="pCodigo" type="text" inputmode="numeric" placeholder="Ingrese código del producto" autocomplete="off" style="flex:1;min-width:220px;">
      <button type="button" id="btnBuscarCodigoProducto" style="white-space:nowrap;">🔎 Consultar</button>
    </div>
    <div id="msgCodigoProducto" style="font-size:12px;color:#64748b;margin-top:6px;"></div>
  `;

  grupoProducto.parentNode.insertBefore(bloque, grupoProducto);

  const inputCodigo = el("pCodigo");
  const btnBuscar = el("btnBuscarCodigoProducto");

  if(inputCodigo){
    inputCodigo.addEventListener("keydown", function(e){
      if(e.key === "Enter"){
        e.preventDefault();
        consultarCodigoProducto();
      }
    });
    inputCodigo.addEventListener("input", function(){
      ULTIMO_CODIGO = "";
      setMensajeCodigo("");
      clearTimeout(TIMER_CODIGO);
      const codigo = limpiarCodigo(this.value);
      if(codigo.length >= 3){
        TIMER_CODIGO = setTimeout(()=>consultarCodigoProducto(false), 250);
      }
    });
    inputCodigo.addEventListener("blur", function(){
      const codigo = limpiarCodigo(this.value);
      if(codigo) consultarCodigoProducto(false);
    });
  }

  if(btnBuscar){
    btnBuscar.addEventListener("click", function(){ consultarCodigoProducto(true); });
  }
}

function normalizarRespuestaProducto(data, codigoFallback){
  if(!data) return null;
  if(Array.isArray(data)){
    if(!data.length) return null;
    return normalizarRespuestaProducto(data[0], codigoFallback);
  }
  if(data.data) return normalizarRespuestaProducto(data.data, codigoFallback);
  if(data.item) return normalizarRespuestaProducto(data.item, codigoFallback);
  if(data.producto && typeof data.producto === 'object') return normalizarRespuestaProducto(data.producto, codigoFallback);
  if(data.ok === false) return null;

  const codigo = limpiarCodigo(data.codigo || data.CODIGO || data.producto || data.PRODUCTO || codigoFallback || "");
  const descripcion = txt(data.descripcion || data.DESCRIPCION || data.detalle || data.DETALLE || data.nombre || data.NOMBRE || "");
  if(!codigo && !descripcion) return null;
  return { codigo, descripcion };
}

async function consultarCodigoProducto(mostrarMensajes = true){
  if(CONSULTANDO_PRODUCTO) return;
  const inputCodigo = el("pCodigo");
  const btnBuscar = el("btnBuscarCodigoProducto");
  const pProducto = el("pProducto");
  const pDescripcion = el("pDescripcion");
  const pCantidad = el("pCantidad");
  if(!inputCodigo || !pProducto || !pDescripcion) return;

  const codigo = limpiarCodigo(inputCodigo.value);
  if(!codigo){
    limpiarCamposProductoSoloConsulta();
    if(mostrarMensajes) setMensajeCodigo("Ingrese un código.", "#dc2626");
    inputCodigo.focus();
    return;
  }
  if(codigo === ULTIMO_CODIGO && txt(pDescripcion.value)) return;

  CONSULTANDO_PRODUCTO = true;
  setLoadingBoton(btnBuscar, true);
  if(mostrarMensajes) setMensajeCodigo("Consultando código...", "#2563eb");

  try{
    let normalizado = null;
    const urls = [
      `${API_PRODUCTOS}?accion=buscar&codigo=${encodeURIComponent(codigo)}&_=${Date.now()}`,
      `${API_PRODUCTOS}?action=buscar&codigo=${encodeURIComponent(codigo)}&_=${Date.now()}`,
      `${API_PRODUCTOS}?q=${encodeURIComponent(codigo)}&_=${Date.now()}`,
      `${API_PRODUCTOS}?codigo=${encodeURIComponent(codigo)}&_=${Date.now()}`
    ];

    for(const url of urls){
      try{
        const res = await fetch(url, { method:'GET', cache:'no-store' });
        if(!res.ok) continue;
        const text = await res.text();
        let data;
        try{ data = JSON.parse(text); }catch(e){ continue; }
        normalizado = normalizarRespuestaProducto(data, codigo);
        if(normalizado) break;
      }catch(e){}
    }

    if(!normalizado){
      const params = new URLSearchParams();
      params.append('accion','buscar');
      params.append('codigo', codigo);
      try{
        const res = await fetch(API_PRODUCTOS, { method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'}, body: params.toString() });
        if(res.ok){
          const text = await res.text();
          try{ normalizado = normalizarRespuestaProducto(JSON.parse(text), codigo); }catch(e){}
        }
      }catch(e){}
    }

    if(!normalizado || !txt(normalizado.descripcion)){
      limpiarCamposProductoSoloConsulta();
      if(mostrarMensajes) setMensajeCodigo("Código no encontrado.", "#dc2626");
      return;
    }

    inputCodigo.value = normalizado.codigo || codigo;
    pProducto.value = normalizado.codigo || codigo;
    pDescripcion.value = normalizado.descripcion || "";
    if(pCantidad && !txt(pCantidad.value)) pCantidad.value = '1';
    ULTIMO_CODIGO = normalizado.codigo || codigo;
    if(mostrarMensajes) setMensajeCodigo("Producto cargado correctamente.", "#16a34a");
    if(pCantidad){ pCantidad.focus(); pCantidad.select && pCantidad.select(); }
  }catch(error){
    console.error('Error consulta producto:', error);
    limpiarCamposProductoSoloConsulta();
    if(mostrarMensajes) setMensajeCodigo(error.message || 'No se pudo consultar el código.', '#dc2626');
  }finally{
    CONSULTANDO_PRODUCTO = false;
    setLoadingBoton(btnBuscar, false);
  }
}

function initConsultaProductoCodigo(){ crearCampoCodigoSiNoExiste(); }
document.addEventListener('DOMContentLoaded', initConsultaProductoCodigo);
window.consultarCodigoProducto = consultarCodigoProducto;
window.initConsultaProductoCodigo = initConsultaProductoCodigo;
