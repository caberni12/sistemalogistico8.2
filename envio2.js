
//const API="https://script.google.com/macros/s/AKfycbzj3sRVqYDgGVak1PNHrycYQ6FI5Mk5UyADOL0uI4CDAprlT7LDv3ZVWrfMCkwPMCgW/exec";


const API="https://script.google.com/macros/s/AKfycbyMhSW9JBm6zb90K1V_qHTuSZ9GqR7XNPAgV3j9upGq66OMQNK9RtEii2gT5QXlTpFD/exec";
//https://script.google.com/macros/s/AKfycbzj3sRVqYDgGVak1PNHrycYQ6FI5Mk5UyADOL0uI4CDAprlT7LDv3ZVWrfMCkwPMCgW/exec

let RAW=[];
let FILT=[];
let visibleCount=0;
let EDIT_ROW=null;
let TRASLADO_ROW=null;

const CHUNK=20;

/* ELEMENTOS */

const cardsGrid=document.getElementById("cardsGrid");

const fBuscar=document.getElementById("fBuscar");
const fStatus=document.getElementById("fStatus");
const fDesde=document.getElementById("fDesde");
const fHasta=document.getElementById("fHasta");

const totalPedidos=document.getElementById("totalPedidos");
const totalCajas=document.getElementById("totalCajas");

const btnReload=document.getElementById("btnReload");
const btnGuardar=document.getElementById("btnGuardar");

const editModal=document.getElementById("editModal");

/* CAMPOS MODAL */

const mFechaIngreso=document.getElementById("mFechaIngreso");
const mPedido=document.getElementById("mPedido");
const mTipoDocumento=document.getElementById("mTipoDocumento");
const mNumeroDocumento=document.getElementById("mNumeroDocumento");
const mCliente=document.getElementById("mCliente");
const mDireccion=document.getElementById("mDireccion");
const mComuna=document.getElementById("mComuna");
const mTransporte=document.getElementById("mTransporte");
const mCajas=document.getElementById("mCajas");
const mResponsable=document.getElementById("mResponsable");
const mFechaEntrega=document.getElementById("mFechaEntrega");
const mStatus=document.getElementById("mStatus");
const mStatusEntrega=document.getElementById("mStatusEntrega");
const mSemaforo=document.getElementById("mSemaforo");
const mDiasAtraso=document.getElementById("mDiasAtraso");
const mObservaciones=document.getElementById("mObservaciones");

const mFoto=document.getElementById("mFoto");
const mPDF=document.getElementById("mPDF");
const mPDFTraslado=document.getElementById("mPDFTraslado");

const boxFoto=document.getElementById("boxFoto");
const boxPDF=document.getElementById("boxPDF");
const boxTraslado=document.getElementById("boxTraslado");

const previewFoto=document.getElementById("previewFoto");
const previewPDF=document.getElementById("previewPDF");
const previewTraslado=document.getElementById("previewTraslado");

/* LOADER */

function setLoading(btn,state){
btn.disabled=state;
btn.classList.toggle("loading",state);
}

/* FECHA */


function parseFechaSoloDia(str){
if(!str) return null;
const parsed = new Date(str);
if(isNaN(parsed.getTime())) return null;
return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function normalizeSearchText(value){
return String(value || "")
.toLowerCase()
.normalize("NFD")
.replace(/[\u0300-\u036f]/g,"")
.trim();
}

function buildPedidoSearchBlob(item){
return normalizeSearchText([
item?.fechaIngreso,
item?.pedido,
item?.tipoDocumento,
item?.numeroDocumento,
item?.cliente,
item?.direccion,
item?.comuna,
item?.transporte,
item?.etiquetas,
item?.responsable,
item?.status,
item?.fechaEntrega,
item?.alerta,
item?.statusEntrega,
item?.diasAtraso,
item?.semaforo,
item?.observaciones,
item?.TR,
item?.solicitudTraslado
].join(" | "));
}

function debounce(fn, wait = 220){
let timer = null;
return (...args)=>{
clearTimeout(timer);
timer = setTimeout(()=>fn(...args), wait);
};
}

function mapPedidoRow(row){
return {
_row: row?._row || "",
fechaIngreso: row?.fechaIngreso || row?.Fecha || row?.["FECHA INGRESO"] || row?.["Fecha Ingreso"] || "",
pedido: row?.pedido || row?.["PEDIDO"] || "",
tipoDocumento: row?.tipoDocumento || row?.["TIPO DOCUMENTO"] || row?.["Tipo Documento"] || "",
numeroDocumento: row?.numeroDocumento || row?.["NUMERO DOCUMENTO"] || row?.["Nº Documento"] || row?.["N° Documento"] || "",
cliente: row?.cliente || row?.["CLIENTE"] || row?.["Cliente"] || "",
direccion: row?.direccion || row?.["DIRECCION"] || row?.["Dirección"] || "",
comuna: row?.comuna || row?.["COMUNA"] || row?.["Comuna"] || "",
transporte: row?.transporte || row?.["TRANSPORTE"] || row?.["Transporte"] || "",
etiquetas: row?.etiquetas || row?.["ETIQUETAS"] || row?.["Etiquetas"] || 0,
observaciones: row?.observaciones || row?.["OBSERVACIONES"] || row?.["Observaciones"] || "",
status: String(row?.status || row?.["STATUS"] || row?.["Status"] || "PENDIENTE").trim().toUpperCase(),
fechaEntrega: row?.fechaEntrega || row?.["FECHA ENTREGA"] || row?.["FECHA ESTIMADA ENTREGA"] || row?.["Fecha Entrega"] || "",
alerta: row?.alerta || row?.["ALERTA"] || row?.["Alerta"] || "",
statusEntrega: row?.statusEntrega || row?.["STATUS ENTREGA"] || "",
diasAtraso: row?.diasAtraso || row?.["DIAS ATRASO"] || "",
semaforo: row?.semaforo || row?.["SEMAFORO"] || "",
responsable: row?.responsable || row?.["RESPONSABLE"] || row?.["Responsable"] || "",
foto: row?.foto || row?.["FOTO"] || "",
pdf: row?.pdf || row?.["PDF"] || "",
pdfTraslado: row?.pdfTraslado || row?.["PDF TRASLADO"] || "",
TR: row?.TR || row?.solicitudTraslado || row?.["TR"] || row?.["SOLICITUD TRASLADO"] || "",
solicitudTraslado: row?.solicitudTraslado || row?.TR || row?.["TR"] || row?.["SOLICITUD TRASLADO"] || "",
_fechaObj: parseFechaSoloDia(row?.fechaIngreso || row?.Fecha || row?.["FECHA INGRESO"] || row?.["Fecha Ingreso"] || "")
};
}


/* ALERTAS */




function calcularAlertas(r){
  // Sin fecha de entrega no hay cálculos
  if (!r.fechaEntrega) return r;
  const ahora = new Date();
  const entrega = new Date(r.fechaEntrega);
  const diffHoras = (entrega - ahora) / (1000 * 60 * 60);
  const diffDias = Math.floor((ahora - entrega) / (1000 * 60 * 60 * 24));
  const status = (r.status || '').toString().trim().toUpperCase();
  // Casos de pedidos finalizados: ENTREGADO o TERMINADO
  if (status === 'ENTREGADO' || status === 'TERMINADO') {
    r.alerta = '';
    r.diasAtraso = '';
    r.statusEntrega = status === 'ENTREGADO' ? 'ENTREGADO A TIEMPO' : 'FINALIZADO';
    r.semaforo = status === 'ENTREGADO' ? 'VERDE' : 'AZUL';
    return r;
  }
  // Pedido atrasado
  if (ahora > entrega) {
    r.alerta = 'PEDIDO ATRASADO';
    r.semaforo = 'ROJO';
    r.diasAtraso = Math.max(diffDias, 1);
    r.statusEntrega = 'ATRASADO';
  } else if (diffHoras <= 48) {
    // Próximo a vencer
    r.alerta = 'ENTREGA EN MENOS DE 48H';
    r.semaforo = 'AMARILLO';
    r.diasAtraso = 0;
    r.statusEntrega = 'POR VENCER';
  } else {
    // Sin alerta
    r.alerta = '';
    r.semaforo = 'VERDE';
    r.diasAtraso = 0;
    r.statusEntrega = 'EN TIEMPO';
  }
  return r;
}

// ------------------------------------------------------------------
//  Función auxiliar para obtener productos del backend por pedido.
//  Se declara fuera de calcularAlertas para poder utilizarla en
//  openTraslado sin interferir con la lógica de alertas.
async function obtenerProductosPedidoBD(pedido) {
  if (!pedido) return [];
  try {
    const payload = { action: "obtenerProductos", pedido: pedido };
    const resp = await fetch(API, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    const text = await resp.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      console.error("Respuesta inválida productos:", text);
      return [];
    }
    if (data && data.ok && Array.isArray(data.productos)) {
      return data.productos;
    }
    return [];
  } catch (err) {
    console.error("Error al obtener productos desde backend:", err);
    return [];
  }
}
/* LOAD */


async 
async function load(){

  try{

    setLoading(btnReload,true);

    const r = await fetch(API);
    const data = await r.json();
    const rows = Array.isArray(data) ? data : [];

    RAW = rows.map(row => calcularAlertas(mapPedidoRow(row)));

    RAW.sort((a,b)=> Number(b._row) - Number(a._row));

    applyFilter();

  }catch(e){

    console.error("ERROR API",e);

  }

  setLoading(btnReload,false);

}

/* FILTROS */


function applyFilter(){

    const texto = normalizeSearchText(fBuscar.value || "");
    const status = fStatus.value;
    const d1 = fDesde.value ? parseFechaSoloDia(fDesde.value) : null;
    const d2 = fHasta.value ? parseFechaSoloDia(fHasta.value) : null;

    FILT = RAW.filter(r=>{

    if(texto && !buildPedidoSearchBlob(r).includes(texto)) return false;

    if(status==="ATRASO" && r.semaforo!=="ROJO") return false;

    if(status && status!=="ATRASO" && r.status!==status) return false;

    if(d1 || d2){

    const fr=r._fechaObj;

    if(!fr) return false;

    if(d1 && fr < d1) return false;

    if(d2 && fr > d2) return false;

    }

    return true;

    });

    FILT.sort((a,b)=> Number(b._row) - Number(a._row));

    totalPedidos.textContent = FILT.length;

    totalCajas.textContent = FILT.reduce((s,r)=>s+Number(r.etiquetas||0),0);

    visibleCount = 0;
    cardsGrid.innerHTML = "";

    renderMore();

}


/* STATUS COLOR */

function statusClass(s){

if(s==="PENDIENTE") return "estado estado-pendiente";
if(s==="EN RUTA") return "estado estado-ruta";
if(s==="RECIBIDO") return "estado estado-recibido";
if(s==="ENTREGADO") return "estado estado-entregado";
if(s==="CANCELADO") return "estado estado-cancelado";
 if(s==="TERMINADO") return "estado estado-terminado";

return "estado";

}

/* TARJETAS */

// ============================================
// Renderizado de tarjetas y scroll infinito
// ============================================
function renderMore() {

  const fragment = document.createDocumentFragment();
  const slice = FILT.slice(visibleCount, visibleCount + CHUNK);

  slice.forEach(r => {

    /* 🔒 VALIDAR TR */
    const tieneTR = r.TR || r.solicitudTraslado;

    let clase = "card";
    if (r.semaforo === "ROJO") clase += " card-atraso";
    if (r.semaforo === "AMARILLO") clase += " card-alerta";

    const mapId = "map_" + r._row;

    const card = document.createElement("div");
    card.className = clase;

    card.innerHTML = `
      
      <div class="card-header">
        <div class="pedido-numero">#${r.pedido}</div>
        <div class="${statusClass(r.status)}">${r.status}</div>
      </div>

      <div class="section">
        <div class="section-title">Cliente</div>
        <div class="section-value">${r.cliente}</div>
      </div>

      <div class="section">
        <div class="section-title">Dirección</div>
        <div class="section-value">${r.direccion} (${r.comuna})</div>
      </div>

      <div class="section">
        <div class="section-title">Responsable</div>
        <div class="section-value">${r.responsable || ""}</div>
      </div>

      <div class="cajas-box">
        UNIDADES <span>${r.etiquetas || 0}</span>
      </div>

      ${tieneTR ? `
        <div class="traslado-ref">
          📦 Traslado: <b>${r.TR || r.solicitudTraslado}</b>
        </div>
      ` : ""}

      ${r.alerta ? `
        <div style="color:#dc2626;font-weight:800">
          ${r.alerta}
          ${r.diasAtraso > 0 ? `<br>Días atraso: ${r.diasAtraso}` : ""}
        </div>
      ` : ""}

      <div class="card-actions">

        ${r.pdfTraslado 
          ? `<a href="${r.pdfTraslado}" target="_blank" class="btn-pdf">📄 Documento Generado</a>` 
          : ""}

        <button onclick="toggleMap('${mapId}',this)">🗺 Mapa</button>

        ${
          tieneTR
          ? ""  // 🔒 BLOQUEADO SI EXISTE TR
          : `
            ${(r.status === "ENTREGADO" || r.status === "TERMINADO")
              ? `<button onclick="openTraslado(${r._row})">📦 Traslado</button>`
              : ""}
            <button onclick="openEdit(${r._row})">✏️ Editar</button>
          `
        }

      </div>

      <div class="map-container" id="${mapId}">
        <iframe src="https://maps.google.com/maps?q=${encodeURIComponent(r.direccion + " " + r.comuna)}&z=15&output=embed"></iframe>
      </div>
    `;

    fragment.appendChild(card);

  });

  cardsGrid.appendChild(fragment);
  visibleCount += CHUNK;
}
  
  /* SCROLL INFINITO */
  window.addEventListener("scroll", () => {
    if (window.innerHeight + window.scrollY >= document.body.offsetHeight - 400) {
      renderMore();
    }
  });
  
  /* ===== FUNCION PARA ACTUALIZAR EL BOTÓN "Documento Generado" DESPUÉS DE CREAR EL TRASLADO ===== */
  function actualizarBotonPdfTraslado(row, pdfUrl, trasladoNum) {
    const cards = document.querySelectorAll(".card");
    cards.forEach(card => {
      const numeroPedido = card.querySelector(".pedido-numero").textContent.replace("#", "");
      if (numeroPedido == row) {
        const acciones = card.querySelector(".card-actions");
        const existe = acciones.querySelector(".btn-pdf");
        if (!existe) {
          const a = document.createElement("a");
          a.href = pdfUrl;
          a.target = "_blank";
          a.className = "btn-pdf";
          a.textContent = "📄 Documento Generado";
          acciones.insertBefore(a, acciones.firstChild);
        }
        // Actualizar número de traslado visible
        const trasladoRef = card.querySelector(".traslado-ref");
        if (!trasladoRef) {
          const div = document.createElement("div");
          div.className = "traslado-ref";
          div.innerHTML = `📦 Traslado: <b>${trasladoNum}</b>`;
          acciones.insertBefore(div, acciones.firstChild.nextSibling);
        }
      }
    });
  }
  

  
  // ============================================
  // Función que se llama al generar traslado
  // ============================================
  async function generarTraslado(row) {
    try {
      // Llamada a tu API global
      const response = await fetch(API, {
        method: "POST",
        body: JSON.stringify({ row }),
        headers: { "Content-Type": "application/json" }
      });
  
      const data = await response.json();
  
      if (data.ok) {
        // Actualizar tarjeta inmediatamente
        actualizarBotonPdfTraslado(row, data.pdf, data.traslado);
        alert("Traslado generado correctamente: " + data.traslado);
      } else {
        alert("Error al generar traslado: " + (data.error || "Desconocido"));
      }
  
    } catch (err) {
      console.error(err);
      alert("Error al generar traslado: Failed to fetch");
    }
  }
/* EDITAR */



function openEdit(row){

    const r=RAW.find(x=>x._row==row);
    if(!r) return;
    
    if(r.status === "ENTREGADO" || r.status === "CANCELADO" || r.status === "TERMINADO"){
      alert("Este pedido está cerrado y no puede ser editado.");
      return;
    }
    
    EDIT_ROW=row;
    
    /* CARGAR DATOS */
    
    // Mostrar la fecha de ingreso incluso si el backend utiliza nombres
    // diferentes como "Fecha" o "FECHA INGRESO".  Tomamos el
    // primero que exista.
    mFechaIngreso.value = r.fechaIngreso || r.Fecha || r["Fecha Ingreso"] || r["FECHA INGRESO"] || r["FECHA DE INGRESO"] || "";
    mPedido.value=r.pedido||"";
    mTipoDocumento.value=r.tipoDocumento||"";
    mNumeroDocumento.value=r.numeroDocumento||"";
    mCliente.value=r.cliente||"";
    mDireccion.value=r.direccion||"";
    mComuna.value=r.comuna||"";
    mTransporte.value=r.transporte||"";
    mCajas.value=r.etiquetas||1;
    mResponsable.value=r.responsable||"";
    mFechaEntrega.value=r.fechaEntrega||"";
    mStatus.value=r.status||"";
    
    mStatusEntrega.value=r.statusEntrega||"";
    mSemaforo.value=r.semaforo||"";
    mDiasAtraso.value=r.diasAtraso||0;
    
    mObservaciones.value=r.observaciones||"";
    
    mFoto.value=r.foto||"";
    mPDF.value=r.pdf||"";
    mPDFTraslado.value=r.pdfTraslado||"";
    
    /* MOSTRAR DOCUMENTOS */
    
    boxFoto.classList.toggle("hidden",!r.foto);
    boxPDF.classList.toggle("hidden",!r.pdf);
    boxTraslado.classList.toggle("hidden",!r.pdfTraslado);
    
    /* BLOQUEAR CAMPOS */
    
    mFechaIngreso.disabled=true;
    mPedido.disabled=true;
    mTipoDocumento.disabled=true;
    mNumeroDocumento.disabled=true;
    mCliente.disabled=true;
    mDireccion.disabled=true;
    mComuna.disabled=true;
    mTransporte.disabled=true;
    mCajas.disabled=true;
    mResponsable.disabled=true;
    
    mStatusEntrega.disabled=true;
    mSemaforo.disabled=true;
    mDiasAtraso.disabled=true;
    
    /* CAMPOS EDITABLES */
    
    mObservaciones.disabled=false;
    mStatus.disabled=false;
    mFechaEntrega.disabled=false;
    
    /* ABRIR MODAL */
    
    editModal.style.display="flex";
    
    }

/* PREVIEW */

previewFoto.onclick=()=>{ if(mFoto.value) window.open(mFoto.value); };
previewPDF.onclick=()=>{ if(mPDF.value) window.open(mPDF.value); };
previewTraslado.onclick=()=>{ if(mPDFTraslado.value) window.open(mPDFTraslado.value); };

/* GUARDAR */

async function guardar(){

    setLoading(btnGuardar,true);
    
    /* OBTENER REGISTRO ORIGINAL */
    
    const r = RAW.find(x=>x._row==EDIT_ROW);
    
    await fetch(API,{
    method:"POST",
    body:JSON.stringify({
    
    action:"update",
    row:EDIT_ROW,
    
    FECHAINGRESO:r.fechaIngreso,
    PEDIDO:r.pedido,
    TIPODOCUMENTO:r.tipoDocumento,
    NUMERODOCUMENTO:r.numeroDocumento,
    CLIENTE:r.cliente,
    DIRECCION:r.direccion,
    COMUNA:r.comuna,
    TRANSPORTE:r.transporte,
    ETIQUETAS:r.etiquetas,
    RESPONSABLE:r.responsable,
    
    STATUS:mStatus.value,
    OBSERVACIONES:mObservaciones.value,
    "FECHA ENTREGA":mFechaEntrega.value
    
    })
    });
    
    closeEdit();
    await load();
    
    setLoading(btnGuardar,false);
    
    }
/* MAPA */

function toggleMap(id,btn){

const el=document.getElementById(id);

if(!el.style.display||el.style.display==="none"){
el.style.display="block";
btn.textContent="➖ Ocultar mapa";
}else{
el.style.display="none";
btn.textContent="🗺 Mostrar mapa";
}

}

/* TRASLADO */

/* ================= TRASLADO ================= */

async function openTraslado(row){

    /* BUSCAR REGISTRO */
    
    const r = RAW.find(x => Number(x._row) === Number(row));
    if(!r){
    console.warn("Pedido no encontrado",row);
    return;
    }
    
    TRASLADO_ROW = row;
    
    /* OBTENER MODAL */
    
    const modal = document.getElementById("trasladoModal");
    if(!modal){
    console.error("No existe el modal trasladoModal");
    return;
    }
    
    /* ABRIR MODAL */
    
    modal.style.display = "flex";
    
    /* CARGAR DATOS */
    
    const tPedido = document.getElementById("tPedido");
    const tCliente = document.getElementById("tCliente");
    const tDireccion = document.getElementById("tDireccion");
    const tComuna = document.getElementById("tComuna");
    const tTransporte = document.getElementById("tTransporte");
    
    if(tPedido) tPedido.value = r.pedido || "";
    if(tCliente) tCliente.value = r.cliente || "";
    if(tDireccion) tDireccion.value = r.direccion || "";
    if(tComuna) tComuna.value = r.comuna || "";
    if(tTransporte) tTransporte.value = r.transporte || "";
    
    /* LIMPIAR TABLA PRODUCTOS */
    const tabla = document.getElementById("detalleTable");
    if(tabla) tabla.innerHTML="";
    // Cargar productos desde backend.  Si no hay productos, agregar una línea vacía.
    try{
      const productosGuardados = await obtenerProductosPedidoBD(r.pedido);
      if(Array.isArray(productosGuardados) && productosGuardados.length){
        productosGuardados.forEach(p => {
          const prodVal = p.producto || "";
          const detVal  = p.detalle || p.descripcion || "";
          const cantVal = Number(p.cantidad || 1);
          addProducto(prodVal, detVal, cantVal);
        });
      } else {
        // Si no hay productos guardados, agregar una fila vacía para ingresar datos
        addProducto();
      }
    }catch(err){
      console.error("Error cargando productos para traslado:", err);
      // En caso de error, dejar una fila vacía
      addProducto();
    }
    }

    /**
     * Agrega una fila de producto al detalle del traslado.  Si se
     * entregan argumentos, éstos se utilizan directamente para
     * poblar los campos (producto, detalle, cantidad).  Si no se
     * entregan argumentos, se utilizan los valores de los inputs
     * superiores (pProducto, pDetalle, pCantidad) y se valida que
     * el producto no esté vacío.
     * @param {string=} prodArg Nombre del producto
     * @param {string=} detArg  Detalle o descripción del producto
     * @param {number=} cantArg Cantidad del producto
     */
    function addProducto(prodArg, detArg, cantArg){
        const tabla = document.getElementById("detalleTable");
        if(!tabla) return;
        // Si hay argumentos se usa un flujo corto
        if (arguments.length > 0) {
          const prodVal = String(prodArg || "");
          const detVal  = String(detArg || "");
          const cantVal = Number(cantArg || 1);
          const fila = document.createElement("tr");
          fila.innerHTML = `
            <td><input class="prod" value="${prodVal}" placeholder="Producto"></td>
            <td><input class="det" value="${detVal}" placeholder="Detalle"></td>
            <td><input type="number" class="cant" value="${cantVal}" min="1" oninput="calcTotal()"></td>
            <td><button type="button" onclick="this.closest('tr').remove();calcTotal()">✖</button></td>
          `;
          tabla.appendChild(fila);
          calcTotal();
          return;
        }
        // Leer valores desde los inputs y validar
        const prod = document.getElementById("pProducto").value.trim();
        const det  = document.getElementById("pDetalle").value.trim();
        const cant = document.getElementById("pCantidad").value || 1;
        if(prod === ""){ alert("Favor Informe los Productos a Trasladar y Genere Numero de Traslado"); return; }
        const fila = document.createElement("tr");
        fila.innerHTML = `
          <td><input class="prod" value="${prod}" placeholder="Producto"></td>
          <td><input class="det" value="${det}" placeholder="Detalle"></td>
          <td><input type="number" class="cant" value="${cant}" min="1" oninput="calcTotal()"></td>
          <td><button type="button" onclick="this.closest('tr').remove();calcTotal()">✖</button></td>
        `;
        tabla.appendChild(fila);
        // Limpiar inputs superiores
        document.getElementById("pProducto").value = "";
        document.getElementById("pDetalle").value = "";
        document.getElementById("pCantidad").value = "1";
        document.getElementById("pProducto").focus();
        calcTotal();
    }
    /* ================= CALCULAR TOTAL ================= */
    
    function calcTotal(){
    
    let total = 0;
    
    document.querySelectorAll(".cant").forEach(input=>{
    total += Number(input.value || 0);
    });
    
    const campoTotal = document.getElementById("tTotal");
    
    if(campoTotal){
    campoTotal.value = total;
    }
    
    }
    
 
    

/* ================= GUARDAR TRASLADO ================= */

async function guardarTraslado(){

    const btn = document.getElementById("btnGuardarTraslado");
    setLoading(btn,true);

    try{

        const r = RAW.find(x => Number(x._row) === Number(TRASLADO_ROW));

        if(!r){
            alert("Pedido no encontrado");
            return;
        }

        /* ================= PRODUCTOS ================= */

        const productos = [];

        document.querySelectorAll("#detalleTable tr").forEach(tr=>{

            const prod = tr.querySelector(".prod")?.value?.trim() || "";
            const det  = tr.querySelector(".det")?.value?.trim() || "";
            const cant = Number(tr.querySelector(".cant")?.value || 0);

            if(prod && cant>0){
                productos.push({
                    producto: prod,
                    detalle: det,
                    descripcion: det,
                    cantidad: cant
                });
            }

        });

        if(productos.length===0){
            alert("Debe agregar al menos un producto");
            return;
        }

        /* ================= LOGO ================= */

        const logoEmpresa="https://lh3.googleusercontent.com/d/11T8x616pxgYq0QYV51JpTulsF2s4szkk";

        /* ================= PAYLOAD ================= */

        const payload={

            action:"crearTraslado",

            row:Number(TRASLADO_ROW),

            pedido:r.pedido||"",
            cliente:r.cliente||"",
            direccion:r.direccion||"",
            comuna:r.comuna||"",
            transporte:r.transporte||"",

            observaciones:document.getElementById("tObs")?.value||"",
            total:Number(document.getElementById("tTotal")?.value||0),

            productos:productos,
            logo:logoEmpresa
        };



        /* ================= FETCH ================= */

        const resp = await fetch(API,{
            method:"POST",
            headers:{
                "Content-Type":"text/plain;charset=utf-8"
            },
            body:JSON.stringify(payload)
        });

        const text = await resp.text();

        let data;
        try{
            data = JSON.parse(text);
        }catch(e){
            console.error("Respuesta servidor:",text);
            throw new Error("Respuesta inválida del servidor");
        }

        if(!data.ok){
            throw new Error(data.error || "Error generando traslado");
        }

        /* ================= ACTUALIZAR DATOS ================= */

        if(data.traslado){
            r.solicitudTraslado=data.traslado;
        }

        if(data.pdf){
            r.pdfTraslado=data.pdf;
        }

        await load();

        if(data.pdf){
            window.open(data.pdf,"_blank");
        }

        alert("Traslado generado: "+data.traslado);

        closeTraslado();

    }
    catch(e){

        console.error(e);
        alert("Error generando traslado: "+e.message);

    }
    finally{

        setLoading(btn,false);

    }

}





/************************************************
 * CERRAR MODAL
 ************************************************/
function closeTraslado() {
    const modal = document.getElementById("trasladoModal");
    if (modal) modal.style.display = "none";
}
  
/* MODAL */

function closeEdit(){
editModal.style.display="none";
}

/* EVENTOS */

btnReload.onclick=load;
btnGuardar.onclick=guardar;

const debouncedApplyFilter = debounce(applyFilter, 220);
fBuscar.oninput=debouncedApplyFilter;
fStatus.onchange=applyFilter;
fDesde.onchange=applyFilter;
fHasta.onchange=applyFilter;


/* ================= AUTO-RELOAD ================= */
let autoLoad = setInterval(() => {
    if ((editModal.style.display === "none" || !editModal.style.display) &&
        (document.getElementById("trasladoModal").style.display === "none" || !document.getElementById("trasladoModal").style.display)) {
        load();
    }
}, 15000); // cada 15 segundos

let autoAlertas = setInterval(() => {
    if ((editModal.style.display === "none" || !editModal.style.display) &&
        (document.getElementById("trasladoModal").style.display === "none" || !document.getElementById("trasladoModal").style.display)) {
        RAW = RAW.map(r => calcularAlertas(r));
        applyFilter();
    }
}, 60000); // cada 1 minuto


/* LIMPIAR localStorage AL CERRAR MODAL */
function closeTraslado() {
    const modal = document.getElementById("trasladoModal");
    if (modal) modal.style.display = "none";
    localStorage.removeItem("trasladoModalRow");
}



load();
