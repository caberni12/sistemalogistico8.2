/***************************************************
 CONSULTA DE PRODUCTO POR CÓDIGO
 - Agrega campo Código al modal de productos
 - Consulta al presionar ENTER o al hacer click
 - Soporta códigos con ceros a la izquierda
 - Producto = código
 - Descripción = descripción
***************************************************/
const API_PRODUCTOS = "https://script.google.com/macros/s/AKfycbz8IUHeDvRiDUvlq_JIrTY1Rb0ZeMVhRt7al_V3NacKfWAMmK-J7Vngjr-hZ19o0woOaQ/exec";

let CONSULTANDO_PRODUCTO = false;
let ULTIMO_CODIGO = "";

/***************************************************
 HELPERS
***************************************************/
function el(id){
  return document.getElementById(id);
}

function txt(v){
  return String(v ?? "").trim();
}

function limpiarCodigo(v){
  return String(v ?? "")
    .trim()
    .replace(/\s+/g, "");
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

function limpiarCamposProducto(){
  const pProducto = el("pProducto");
  const pDescripcion = el("pDescripcion");

  if(pProducto) pProducto.value = "";
  if(pDescripcion) pDescripcion.value = "";
}

/***************************************************
 CREA CAMPO CÓDIGO SI NO EXISTE
***************************************************/
function crearCampoCodigoSiNoExiste(){
  const modalProducto = el("modalProducto");
  const pProducto = el("pProducto");

  if(!modalProducto || !pProducto) return;
  if(el("pCodigo")) return;

  const grupoProducto = pProducto.closest(".form-group");
  if(!grupoProducto) return;

  const bloque = document.createElement("div");
  bloque.className = "form-group";
  bloque.id = "grupoCodigoProducto";
  bloque.innerHTML = `
    <label>Código</label>
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <input id="pCodigo" type="text" inputmode="numeric" placeholder="Ingrese código del producto" autocomplete="off">
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
    });
  }

  if(btnBuscar){
    btnBuscar.addEventListener("click", function(){
      consultarCodigoProducto();
    });
  }
}

/***************************************************
 CONSULTAR PRODUCTO
***************************************************/
async function consultarCodigoProducto(){
  if(CONSULTANDO_PRODUCTO) return;

  const inputCodigo = el("pCodigo");
  const btnBuscar = el("btnBuscarCodigoProducto");
  const pProducto = el("pProducto");
  const pDescripcion = el("pDescripcion");
  const pCantidad = el("pCantidad");

  if(!inputCodigo){
    console.error("No existe #pCodigo");
    return;
  }

  if(!pProducto){
    console.error("No existe #pProducto");
    return;
  }

  if(!pDescripcion){
    console.error("No existe #pDescripcion");
    return;
  }

  const codigo = limpiarCodigo(inputCodigo.value);

  if(!codigo){
    limpiarCamposProducto();
    setMensajeCodigo("Ingrese un código.", "#dc2626");
    inputCodigo.focus();
    return;
  }

  if(codigo === ULTIMO_CODIGO && txt(pDescripcion.value)){
    return;
  }

  CONSULTANDO_PRODUCTO = true;
  setLoadingBoton(btnBuscar, true);
  setMensajeCodigo("Consultando código...", "#2563eb");

  try{
    const url = `${API_PRODUCTOS}?accion=buscar&codigo=${encodeURIComponent(codigo)}&_=${Date.now()}`;

    const res = await fetch(url, {
      method: "GET",
      cache: "no-store"
    });

    const rawText = await res.text();

    if(!res.ok){
      throw new Error("Error de conexión con el servidor.");
    }

    let data;
    try{
      data = JSON.parse(rawText);
    }catch(err){
      throw new Error("La respuesta no vino en JSON. Revisa el Web App.");
    }

    if(!data.ok){
      limpiarCamposProducto();
      setMensajeCodigo(data.msg || "Código no encontrado.", "#dc2626");
      return;
    }

    const codigoResp = limpiarCodigo(data.codigo || codigo);
    const descripcion = txt(data.descripcion);

    if(!descripcion){
      limpiarCamposProducto();
      setMensajeCodigo("El código existe pero no tiene descripción.", "#dc2626");
      return;
    }

    if(el("pCodigo")) el("pCodigo").value = codigoResp;

    // CORREGIDO:
    // Producto = código
    // Descripción = descripción
    pProducto.value = codigoResp;
    pDescripcion.value = descripcion;

    if(pCantidad && !txt(pCantidad.value)){
      pCantidad.value = "1";
    }

    ULTIMO_CODIGO = codigoResp;
    setMensajeCodigo("Producto cargado correctamente.", "#16a34a");

    if(pCantidad){
      pCantidad.focus();
      pCantidad.select?.();
    }

  }catch(error){
    console.error("Error consulta producto:", error);
    limpiarCamposProducto();
    setMensajeCodigo(error.message || "No se pudo consultar el código.", "#dc2626");
  }finally{
    CONSULTANDO_PRODUCTO = false;
    setLoadingBoton(btnBuscar, false);
  }
}

/***************************************************
 INICIALIZAR
***************************************************/
function initConsultaProductoCodigo(){
  crearCampoCodigoSiNoExiste();
}

document.addEventListener("DOMContentLoaded", initConsultaProductoCodigo);

window.consultarCodigoProducto = consultarCodigoProducto;
window.initConsultaProductoCodigo = initConsultaProductoCodigo;
