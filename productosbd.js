/***************************************************
 CONSULTA DE PRODUCTO POR CÓDIGO
 - Busca solo con ENTER o botón CONSULTAR
 - Incluye botón LIMPIAR
 - Usa select de resultados
 - Precarga catálogo para búsqueda más rápida
***************************************************/
const API_PRODUCTOS = "https://script.google.com/macros/s/AKfycbz8IUHeDvRiDUvlq_JIrTY1Rb0ZeMVhRt7al_V3NacKfWAMmK-J7Vngjr-hZ19o0woOaQ/exec";

const PRODUCT_CACHE_KEY = "catalogo_productos_modal_v1";
const PRODUCT_CACHE_TTL = 1000 * 60 * 30; // 30 min

let CONSULTANDO_PRODUCTO = false;
let ULTIMO_CODIGO = "";
let PRODUCTOS_CACHE = [];
let PRODUCTOS_CACHE_CARGADO = false;
let PRODUCTOS_CACHE_PROMISE = null;

function el(id){ return document.getElementById(id); }
function txt(v){ return String(v ?? "").trim(); }
function limpiarCodigo(v){ return String(v ?? "").trim().replace(/\s+/g, ""); }

function setMensajeCodigo(msg, color = "#64748b"){
  const box = el("msgCodigoProducto");
  if(box){
    box.textContent = msg || "";
    box.style.color = color;
  }
}

function setLoadingBoton(btn, estado){
  if(!btn) return;
  if(estado){
    btn.disabled = true;
    btn.classList.add("loading");
  }else{
    btn.disabled = false;
    btn.classList.remove("loading");
  }
}

function limpiarCamposProductoSoloConsulta(){
  const pProducto = el("pProducto");
  const pDescripcion = el("pDescripcion");
  if(pProducto) pProducto.value = "";
  if(pDescripcion) pDescripcion.value = "";
}

function limpiarBusquedaProducto(){
  const pCodigo = el("pCodigo");
  const pProducto = el("pProducto");
  const pDescripcion = el("pDescripcion");
  const pCantidad = el("pCantidad");

  if(pCodigo) pCodigo.value = "";
  if(pProducto) pProducto.value = "";
  if(pDescripcion) pDescripcion.value = "";
  if(pCantidad) pCantidad.value = "";

  ULTIMO_CODIGO = "";
  ocultarSelectProductos();
  setMensajeCodigo("");

  if(pCodigo) pCodigo.focus();
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
      <button type="button" id="btnLimpiarCodigoProducto" style="white-space:nowrap;">🧹 Limpiar</button>
    </div>

    <select id="selectCodigoProducto" size="6" style="display:none;width:100%;margin-top:8px;padding:8px;border:1px solid #d1d5db;border-radius:10px;background:#fff;"></select>

    <div id="msgCodigoProducto" style="font-size:12px;color:#64748b;margin-top:6px;"></div>
  `;

  grupoProducto.parentNode.insertBefore(bloque, grupoProducto);

  const inputCodigo = el("pCodigo");
  const btnBuscar = el("btnBuscarCodigoProducto");
  const btnLimpiar = el("btnLimpiarCodigoProducto");
  const selectCodigo = el("selectCodigoProducto");

  if(inputCodigo){
    inputCodigo.addEventListener("keydown", function(e){
      if(e.key === "Enter"){
        e.preventDefault();
        consultarCodigoProducto(true);
      }

      if(e.key === "ArrowDown"){
        if(selectCodigo && selectCodigo.style.display !== "none" && selectCodigo.options.length > 0){
          e.preventDefault();
          selectCodigo.focus();
          selectCodigo.selectedIndex = 0;
        }
      }

      if(e.key === "Escape"){
        ocultarSelectProductos();
      }
    });

    inputCodigo.addEventListener("input", function(){
      ULTIMO_CODIGO = "";
      setMensajeCodigo("");
      ocultarSelectProductos();
      limpiarCamposProductoSoloConsulta();
    });
  }

  if(btnBuscar){
    btnBuscar.addEventListener("click", function(){
      consultarCodigoProducto(true);
    });
  }

  if(btnLimpiar){
    btnLimpiar.addEventListener("click", function(){
      limpiarBusquedaProducto();
    });
  }

  if(selectCodigo){
    selectCodigo.addEventListener("dblclick", aplicarProductoDesdeSelect);
    selectCodigo.addEventListener("keydown", function(e){
      if(e.key === "Enter"){
        e.preventDefault();
        aplicarProductoDesdeSelect();
      }
      if(e.key === "Escape"){
        ocultarSelectProductos();
        if(inputCodigo) inputCodigo.focus();
      }
    });
    selectCodigo.addEventListener("change", aplicarProductoDesdeSelect);
  }
}

/***************************************************
 NORMALIZACIÓN
***************************************************/
function normalizarProductoItem(item){
  if(!item || typeof item !== "object") return null;

  const codigo = limpiarCodigo(
    item.codigo ||
    item.CODIGO ||
    item.cod ||
    item.Cod ||
    item.producto ||
    item.PRODUCTO ||
    item.id ||
    item.ID ||
    ""
  );

  const descripcion = txt(
    item.descripcion ||
    item.DESCRIPCION ||
    item.detalle ||
    item.DETALLE ||
    item.nombre ||
    item.NOMBRE ||
    item.productoDescripcion ||
    item.PRODUCTODESCRIPCION ||
    ""
  );

  if(!codigo && !descripcion) return null;

  return { codigo, descripcion };
}

function normalizarListaProductos(data){
  if(!data) return [];

  let lista = [];

  if(Array.isArray(data)) lista = data;
  else if(Array.isArray(data.data)) lista = data.data;
  else if(Array.isArray(data.items)) lista = data.items;
  else if(Array.isArray(data.productos)) lista = data.productos;
  else if(Array.isArray(data.resultado)) lista = data.resultado;
  else if(Array.isArray(data.results)) lista = data.results;
  else if(Array.isArray(data.rows)) lista = data.rows;
  else return [];

  return lista.map(normalizarProductoItem).filter(x => x && x.codigo);
}

function normalizarRespuestaProducto(data, codigoFallback){
  if(!data) return null;

  if(Array.isArray(data)){
    if(!data.length) return null;
    return normalizarRespuestaProducto(data[0], codigoFallback);
  }

  if(data.data && !Array.isArray(data.data)) return normalizarRespuestaProducto(data.data, codigoFallback);
  if(data.item) return normalizarRespuestaProducto(data.item, codigoFallback);
  if(data.producto && typeof data.producto === "object" && !Array.isArray(data.producto)){
    return normalizarRespuestaProducto(data.producto, codigoFallback);
  }
  if(data.ok === false) return null;

  const codigo = limpiarCodigo(
    data.codigo ||
    data.CODIGO ||
    data.cod ||
    data.Cod ||
    data.producto ||
    data.PRODUCTO ||
    codigoFallback ||
    ""
  );

  const descripcion = txt(
    data.descripcion ||
    data.DESCRIPCION ||
    data.detalle ||
    data.DETALLE ||
    data.nombre ||
    data.NOMBRE ||
    ""
  );

  if(!codigo && !descripcion) return null;
  return { codigo, descripcion };
}

/***************************************************
 CACHE LOCAL
***************************************************/
function guardarCacheProductos(lista){
  try{
    localStorage.setItem(PRODUCT_CACHE_KEY, JSON.stringify({
      time: Date.now(),
      data: lista || []
    }));
  }catch(e){}
}

function leerCacheProductos(){
  try{
    const raw = localStorage.getItem(PRODUCT_CACHE_KEY);
    if(!raw) return [];
    const json = JSON.parse(raw);
    if(!json || !Array.isArray(json.data)) return [];
    if((Date.now() - Number(json.time || 0)) > PRODUCT_CACHE_TTL) return [];
    return json.data.map(normalizarProductoItem).filter(Boolean);
  }catch(e){
    return [];
  }
}

/***************************************************
 PRECARGA DE CATÁLOGO
***************************************************/
async function precargarCatalogoProductos(force = false){
  if(PRODUCTOS_CACHE_CARGADO && !force) return PRODUCTOS_CACHE;
  if(PRODUCTOS_CACHE_PROMISE && !force) return PRODUCTOS_CACHE_PROMISE;

  const cacheLocal = leerCacheProductos();
  if(cacheLocal.length && !force){
    PRODUCTOS_CACHE = cacheLocal;
    PRODUCTOS_CACHE_CARGADO = true;
  }

  PRODUCTOS_CACHE_PROMISE = (async () => {
    const urls = [
      `${API_PRODUCTOS}?accion=listar&_=${Date.now()}`,
      `${API_PRODUCTOS}?action=listar&_=${Date.now()}`,
      `${API_PRODUCTOS}?listar=1&_=${Date.now()}`,
      `${API_PRODUCTOS}?all=1&_=${Date.now()}`,
      `${API_PRODUCTOS}?_=${Date.now()}`
    ];

    for(const url of urls){
      try{
        const res = await fetch(url, { method: "GET", cache: "no-store" });
        if(!res.ok) continue;

        const text = await res.text();
        let data = null;

        try{
          data = JSON.parse(text);
        }catch(e){
          continue;
        }

        const lista = normalizarListaProductos(data);
        if(lista.length){
          PRODUCTOS_CACHE = lista;
          PRODUCTOS_CACHE_CARGADO = true;
          guardarCacheProductos(lista);
          return lista;
        }
      }catch(e){}
    }

    if(cacheLocal.length){
      PRODUCTOS_CACHE = cacheLocal;
      PRODUCTOS_CACHE_CARGADO = true;
      return cacheLocal;
    }

    PRODUCTOS_CACHE = [];
    PRODUCTOS_CACHE_CARGADO = true;
    return [];
  })();

  try{
    return await PRODUCTOS_CACHE_PROMISE;
  }finally{
    PRODUCTOS_CACHE_PROMISE = null;
  }
}

/***************************************************
 SELECT
***************************************************/
function ocultarSelectProductos(){
  const select = el("selectCodigoProducto");
  if(!select) return;
  select.innerHTML = "";
  select.style.display = "none";
}

function mostrarSelectProductos(lista){
  const select = el("selectCodigoProducto");
  if(!select) return;

  select.innerHTML = "";

  if(!lista || !lista.length){
    select.style.display = "none";
    return;
  }

  lista.forEach((item, i) => {
    const op = document.createElement("option");
    op.value = item.codigo;
    op.dataset.codigo = item.codigo;
    op.dataset.descripcion = item.descripcion || "";
    op.textContent = `${item.codigo} - ${item.descripcion || "Sin descripción"}`;
    if(i === 0) op.selected = true;
    select.appendChild(op);
  });

  select.style.display = "block";
}

function aplicarProductoSeleccionado(item){
  const inputCodigo = el("pCodigo");
  const pProducto = el("pProducto");
  const pDescripcion = el("pDescripcion");
  const pCantidad = el("pCantidad");

  if(!item) return;

  if(inputCodigo) inputCodigo.value = item.codigo || "";
  if(pProducto) pProducto.value = item.codigo || "";
  if(pDescripcion) pDescripcion.value = item.descripcion || "";
  if(pCantidad && !txt(pCantidad.value)) pCantidad.value = "1";

  ULTIMO_CODIGO = item.codigo || "";
  ocultarSelectProductos();
  setMensajeCodigo("Producto cargado correctamente.", "#16a34a");

  if(pCantidad){
    pCantidad.focus();
    if(pCantidad.select) pCantidad.select();
  }
}

function aplicarProductoDesdeSelect(){
  const select = el("selectCodigoProducto");
  if(!select || !select.value) return;

  const op = select.options[select.selectedIndex];
  if(!op) return;

  aplicarProductoSeleccionado({
    codigo: txt(op.dataset.codigo || op.value),
    descripcion: txt(op.dataset.descripcion || "")
  });
}

/***************************************************
 BÚSQUEDA LOCAL
***************************************************/
function buscarProductosLocales(texto){
  const q = limpiarCodigo(texto).toLowerCase();
  if(!q) return [];

  const exactos = [];
  const empiezan = [];
  const contiene = [];

  for(const item of PRODUCTOS_CACHE){
    const codigo = limpiarCodigo(item.codigo).toLowerCase();
    const descripcion = txt(item.descripcion).toLowerCase();

    if(!codigo && !descripcion) continue;

    if(codigo === q){
      exactos.push(item);
      continue;
    }

    if(codigo.startsWith(q)){
      empiezan.push(item);
      continue;
    }

    if(codigo.includes(q) || descripcion.includes(q)){
      contiene.push(item);
    }
  }

  return [...exactos, ...empiezan, ...contiene].slice(0, 30);
}

/***************************************************
 CONSULTA SOLO CON BOTÓN O ENTER
***************************************************/
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
    ocultarSelectProductos();
    if(mostrarMensajes) setMensajeCodigo("Ingrese un código.", "#dc2626");
    inputCodigo.focus();
    return;
  }

  if(codigo === ULTIMO_CODIGO && txt(pDescripcion.value)) return;

  await precargarCatalogoProductos(false);

  const encontrados = buscarProductosLocales(codigo);

  if(encontrados.length){
    mostrarSelectProductos(encontrados);

    const exacto = encontrados.find(x => limpiarCodigo(x.codigo) === codigo);
    if(exacto){
      aplicarProductoSeleccionado(exacto);
      return;
    }

    if(mostrarMensajes){
      setMensajeCodigo(`Seleccione un producto de la lista (${encontrados.length} resultado(s)).`, "#2563eb");
    }
    return;
  }

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
        const res = await fetch(url, { method: "GET", cache: "no-store" });
        if(!res.ok) continue;

        const text = await res.text();
        let data = null;

        try{
          data = JSON.parse(text);
        }catch(e){
          continue;
        }

        normalizado = normalizarRespuestaProducto(data, codigo);
        if(normalizado) break;
      }catch(e){}
    }

    if(!normalizado){
      const params = new URLSearchParams();
      params.append("accion", "buscar");
      params.append("codigo", codigo);

      try{
        const res = await fetch(API_PRODUCTOS, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
          body: params.toString()
        });

        if(res.ok){
          const text = await res.text();
          try{
            normalizado = normalizarRespuestaProducto(JSON.parse(text), codigo);
          }catch(e){}
        }
      }catch(e){}
    }

    if(!normalizado || !txt(normalizado.descripcion)){
      limpiarCamposProductoSoloConsulta();
      ocultarSelectProductos();
      if(mostrarMensajes) setMensajeCodigo("Código no encontrado.", "#dc2626");
      return;
    }

    inputCodigo.value = normalizado.codigo || codigo;
    pProducto.value = normalizado.codigo || codigo;
    pDescripcion.value = normalizado.descripcion || "";
    if(pCantidad && !txt(pCantidad.value)) pCantidad.value = "1";
    ULTIMO_CODIGO = normalizado.codigo || codigo;

    const yaExiste = PRODUCTOS_CACHE.some(x => limpiarCodigo(x.codigo) === limpiarCodigo(normalizado.codigo));
    if(!yaExiste){
      PRODUCTOS_CACHE.unshift(normalizado);
      guardarCacheProductos(PRODUCTOS_CACHE);
    }

    ocultarSelectProductos();
    if(mostrarMensajes) setMensajeCodigo("Producto cargado correctamente.", "#16a34a");

    if(pCantidad){
      pCantidad.focus();
      if(pCantidad.select) pCantidad.select();
    }

  }catch(error){
    console.error("Error consulta producto:", error);
    limpiarCamposProductoSoloConsulta();
    ocultarSelectProductos();
    if(mostrarMensajes) setMensajeCodigo(error.message || "No se pudo consultar el código.", "#dc2626");
  }finally{
    CONSULTANDO_PRODUCTO = false;
    setLoadingBoton(btnBuscar, false);
  }
}

/***************************************************
 LIMPIAR AL CERRAR MODAL
***************************************************/
function vincularLimpiezaAlCerrarModal(){
  const modal = el("modalProducto");
  if(!modal || modal.dataset.cleanupBound === "1") return;

  modal.dataset.cleanupBound = "1";

  modal.addEventListener("click", function(e){
    if(e.target === modal){
      limpiarBusquedaProducto();
    }
  });

  document.addEventListener("keydown", function(e){
    const modalVisible = modal.style.display === "flex" ||
                         modal.style.display === "block" ||
                         modal.classList.contains("show") ||
                         !modal.hasAttribute("hidden");

    if(e.key === "Escape" && modalVisible){
      limpiarBusquedaProducto();
    }
  });
}

/***************************************************
 INIT
***************************************************/
async function initConsultaProductoCodigo(){
  crearCampoCodigoSiNoExiste();
  vincularLimpiezaAlCerrarModal();
  precargarCatalogoProductos(false).catch(() => {});
}

document.addEventListener("DOMContentLoaded", initConsultaProductoCodigo);

window.consultarCodigoProducto = consultarCodigoProducto;
window.initConsultaProductoCodigo = initConsultaProductoCodigo;
window.precargarCatalogoProductos = precargarCatalogoProductos;
window.limpiarBusquedaProducto = limpiarBusquedaProducto;
