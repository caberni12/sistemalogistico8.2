/***************************************************
API
***************************************************/
//const API="https://script.google.com/macros/s/AKfycbxtLfg0gSUBPCBgDZZeVC-yO7KElDU5RLbTmvj68K9UPOthpdtgLfrk_MRTGTpRaa1M/exec";

const API="https://script.google.com/macros/s/AKfycbyMhSW9JBm6zb90K1V_qHTuSZ9GqR7XNPAgV3j9upGq66OMQNK9RtEii2gT5QXlTpFD/exec";

/***************************************************
DOM
***************************************************/
const tbody=document.getElementById("tbody");
const mobileList=document.getElementById("mobileList");

const search=document.getElementById("search");
const fStatus=document.getElementById("fStatus");
const fDesde=document.getElementById("fDesde");
const fHasta=document.getElementById("fHasta");

const btnReload=document.getElementById("btnReload");
const btnNuevo=document.getElementById("btnNuevo");
const btnGuardar=document.getElementById("btnGuardar");
const btnCancelar=document.getElementById("btnCancelar");

const btnPDF=document.getElementById("btnPDF");
const btnExcel=document.getElementById("btnExcel");

const modalForm=document.getElementById("modalForm");

const mPedido=document.getElementById("mPedido");
const mTipoDoc=document.getElementById("mTipoDoc");
const mNumeroDoc=document.getElementById("mNumeroDoc");
const mCliente=document.getElementById("mCliente");
const mDireccion=document.getElementById("mDireccion");
const mComuna=document.getElementById("mComuna");
const mTransporte=document.getElementById("mTransporte");
const mCajas=document.getElementById("mCajas");
const mStatus=document.getElementById("mStatus");
const mHoraEntrega=document.getElementById("mHoraEntrega");
const mResponsable=document.getElementById("mResponsable");
const mObs=document.getElementById("mObs");

const mFotos=document.getElementById("mFotos");
const mPdf=document.getElementById("mPdf");

const kpis=document.getElementById("kpis");

const fotoModal=document.getElementById("fotoModal");
const fotoGrande=document.getElementById("fotoGrande");
const btnCerrarFoto=document.getElementById("btnCerrarFoto");
const btnDescargarFoto=document.getElementById("btnDescargarFoto");

const mapModal=document.getElementById("mapModal");
const mapFrame=document.getElementById("mapFrame");
const btnCerrarMapa=document.getElementById("btnCerrarMapa");

/* 🔒 BLOQUEAR PEDIDO */
if(mPedido){
  mPedido.readOnly = true;
}

/***************************************************
VARIABLES
***************************************************/
let RAW=[];
let FILT=[];
let EDIT=null;
let KPI_CHARTS={};
let isEditing = false;
let selectedIndex = -1;
let SELECTED_ROW_KEY = null;
function getRowDataByRow(row){
  return RAW.find(r => Number(r._row) === Number(row)) || null;
}

function getSelectedRowData(){
  if(SELECTED_ROW_KEY === null || SELECTED_ROW_KEY === undefined) return null;
  return getRowDataByRow(SELECTED_ROW_KEY);
}

window.getPedidosData = () => Array.isArray(RAW) ? RAW.map(r => ({...r})) : [];
window.getPedidoByRow = (row) => {
  const item = getRowDataByRow(row);
  return item ? {...item} : null;
};
window.getSelectedPedidoData = () => {
  const item = getSelectedRowData();
  return item ? {...item} : null;
};

/***************************************************
PAGINACION
***************************************************/
let PAGE = 1;
let PAGE_SIZE = 20;
let TOTAL_PAGES = 1;

/***************************************************
UTIL
***************************************************/
function setLoading(btn,state){
 if(!btn) return;
 btn.disabled=state;
 btn.classList.toggle("loading",state);
}

function pad2(n){
 return String(n).padStart(2,"0");
}

function parseLocalDate(value,endOfDay=false){
 if(value===null || value===undefined || value==="") return null;

 if(value instanceof Date && !isNaN(value.getTime())){
  return new Date(value.getTime());
 }

 const s=String(value).trim();
 if(!s) return null;

 // dd/MM/yyyy HH:mm o dd/MM/yyyyTHH:mm
 let m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})[T ](\d{2}):(\d{2})(?::(\d{2}))?$/);
 if(m){
  return new Date(
   Number(m[3]),
   Number(m[2])-1,
   Number(m[1]),
   Number(m[4]),
   Number(m[5]),
   Number(m[6]||0),
   0
  );
 }

 // dd/MM/yyyy
 m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
 if(m){
  return new Date(
   Number(m[3]),
   Number(m[2])-1,
   Number(m[1]),
   endOfDay?23:0,
   endOfDay?59:0,
   endOfDay?59:0,
   0
  );
 }

 // yyyy-MM-dd HH:mm o yyyy-MM-ddTHH:mm
 m=s.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?$/);
 if(m){
  return new Date(
   Number(m[1]),
   Number(m[2])-1,
   Number(m[3]),
   Number(m[4]),
   Number(m[5]),
   Number(m[6]||0),
   0
  );
 }

 // yyyy-MM-dd
 m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
 if(m){
  return new Date(
   Number(m[1]),
   Number(m[2])-1,
   Number(m[3]),
   endOfDay?23:0,
   endOfDay?59:0,
   endOfDay?59:0,
   0
  );
 }

 const d=new Date(s);
 return isNaN(d.getTime()) ? null : d;
}

function formatDate(d){
 const date=parseLocalDate(d);
 if(!date) return "";
 return `${pad2(date.getDate())}/${pad2(date.getMonth()+1)}/${date.getFullYear()}`;
}

function toDateFilterValue(value,endOfDay=false){
 return parseLocalDate(value,endOfDay);
}

function formatDateTimeLocalValue(value){
 const date=parseLocalDate(value);
 if(!date) return "";
 return `${date.getFullYear()}-${pad2(date.getMonth()+1)}-${pad2(date.getDate())}T${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

/***************************************************
BASE64
***************************************************/
function fileToBase64(file){
 return new Promise((resolve,reject)=>{
  const reader=new FileReader();
  reader.onload=()=>resolve(reader.result);
  reader.onerror=e=>reject(e);
  reader.readAsDataURL(file);
 });
}

function blobToBase64(blob){
 return new Promise((resolve,reject)=>{
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result);
  reader.onerror = (e) => reject(e);
  reader.readAsDataURL(blob);
 });
}

/**
 * Obtiene la lista de productos almacenados para un pedido desde el
 * backend de Apps Script.  Se envía un payload con la acción
 * "obtenerProductos" y el número de pedido.  Devuelve un arreglo de
 * productos o un arreglo vacío si no hay datos o ocurre un error.
 * @param {string} pedido Número de pedido
 * @return {Promise<Array<Object>>}
 */
async function obtenerProductosPedidoBD(pedido){
  if(!pedido) return [];
  try{
    const payload = {
      action: "obtenerProductos",
      pedido: pedido
    };
    const res = await fetch(API, {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if(data && data.ok && Array.isArray(data.productos)){
      return data.productos;
    }
    return [];
  }catch(err){
    console.error("Error obteniendo productos desde backend:", err);
    return [];
  }
}

function construirPayloadPedido(base = {}, overrides = {}){
 return {
  action: base && base._row ? "update" : "add",
  row: base?._row || "",
  "TIPO DOCUMENTO": base?.tipoDocumento || "",
  "NUMERO DOCUMENTO": base?.numeroDocumento || "",
  "CLIENTE": base?.cliente || "",
  "DIRECCION": base?.direccion || "",
  "COMUNA": base?.comuna || "",
  "TRANSPORTE": base?.transporte || "",
  "ETIQUETAS": base?.etiquetas || "",
  "STATUS": base?.status || "PENDIENTE",
  "FECHA ENTREGA": base?.fechaEntrega || "",
  "RESPONSABLE": base?.responsable || "",
  "OBSERVACIONES": base?.observaciones || "",
  FOTO: base?.foto || "",
  PDF: base?.pdf || "",
  ...overrides
 };
}

async function postPedidoData(data){
 const params = new URLSearchParams();
 params.append("data", JSON.stringify(data));

 const res = await fetch(API,{
  method:"POST",
  headers:{ "Content-Type":"application/x-www-form-urlencoded" },
  body: params
 });

 if(!res.ok){
  throw new Error("Error de conexión");
 }

 let response = {};
 try{
  response = await res.json();
 }catch(e){
  response = {};
 }

 if(response.ok === false){
  throw new Error(response.error || "Error backend");
 }

 return response;
}

async function actualizarPdfPedidoDesdeDataUrl(row, pdfDataUrl){
 const registro = getRowDataByRow(row);
 if(!registro) throw new Error("No se encontró el pedido para actualizar el PDF");

 const nombreArchivo = `Productos_Pedido_${registro.pedido || row}_${Date.now()}.pdf`;

 const payload = construirPayloadPedido(registro,{
  action: "update",
  row: registro._row || row,
  PDF_BASE64: pdfDataUrl,
  PDF_NOMBRE: nombreArchivo
 });

 const response = await postPedidoData(payload);
 const pdfUrl = response?.pdfUrl || response?.url || response?.pdf || response?.PDF || "";

 if(!pdfUrl){
  throw new Error("El backend no devolvió la URL del PDF. Debes guardar el archivo en Drive y retornar pdfUrl.");
 }

 const itemRaw = RAW.find(r => Number(r._row) === Number(row));
 if(itemRaw) itemRaw.pdf = pdfUrl;
 const itemFilt = FILT.find(r => Number(r._row) === Number(row));
 if(itemFilt) itemFilt.pdf = pdfUrl;

 const pdfPreview = document.getElementById("pdfPreview");
 if(pdfPreview && EDIT && Number(EDIT) === Number(row)) {
  pdfPreview.innerHTML = `
    <div class="preview-box">
      <a href="${pdfUrl}" target="_blank">📄 Ver documento actual</a>
    </div>
  `;
 }

 return pdfUrl;
}

/***************************************************
SEMAFORO
***************************************************/
/**
 * Calcula un indicador visual de semáforo basado en el estado del pedido
 * o la fecha de entrega.  Devuelve un círculo de color acorde al
 * estado calculado.  Si el registro ya contiene una propiedad
 * "semaforo", se utiliza ésta para determinar el color; de lo
 * contrario se calcula a partir de la fecha de entrega.
 * @param {Object} registro Registro del pedido
 * @return {string} HTML con un punto de color
 */
function renderSemaforoValue(registro){
 if(!registro) return "";
 let color;
 let sem = (registro.semaforo || "").toString().trim().toUpperCase();
 if(sem){
  if(sem.includes("ROJO")) color = "#dc2626";
  else if(sem.includes("AMARILLO")) color = "#eab308";
  else if(sem.includes("VERDE")) color = "#16a34a";
  else if(sem.includes("AZUL")) color = "#2563eb";
 }
 // Si no hay propiedad semáforo, calcular por fecha de entrega
 if(!color){
   const fechaEntrega = registro.fechaEntrega;
   if(!fechaEntrega) return "";
   const hoy = new Date();
   const entrega = parseLocalDate(fechaEntrega,true);
   if(!entrega) return "";
   const diff = Math.floor((entrega - hoy)/(1000*60*60*24));
   if(diff>1) color = "#16a34a"; // verde
   else if(diff===1) color = "#eab308"; // amarillo
   else if(diff<0) color = "#dc2626"; // rojo
   else color = "#2563eb"; // azul
 }
 return `<span class="sem-dot" style="background:${color}"></span>`;
}

/**
 * Conservamos esta función para uso en las tarjetas móviles y en
 * otros componentes que muestran un texto en lugar de un punto.  Se
 * calcula basándose en la fecha de entrega.
 * @param {string|Date} fechaEntrega
 * @return {string} HTML con una etiqueta estilizada
 */
function calcularSemaforo(fechaEntrega){
 if(!fechaEntrega) return "";
 const hoy = new Date();
 const entrega = parseLocalDate(fechaEntrega,true);
 if(!entrega) return "";
 const diff = Math.floor((entrega - hoy)/(1000*60*60*24));
 if(diff>1) return `<span class="sem-verde">OK</span>`;
 if(diff===1) return `<span class="sem-amarillo">HOY</span>`;
 if(diff<0) return `<span class="sem-rojo">ATRASO</span>`;
 return `<span class="sem-azul">PROX</span>`;
}

/***************************************************
ALERTA
***************************************************/
function renderAlerta(alerta){

 if(!alerta) return `<span class="alerta alerta-verde">OK</span>`;

 alerta=alerta.toUpperCase();

 if(alerta.includes("ATRASO")) return `<span class="alerta alerta-rojo">ATRASADO</span>`;
 if(alerta.includes("48")) return `<span class="alerta alerta-amarillo">POR VENCER</span>`;

 return `<span class="alerta alerta-azul">${alerta}</span>`;
}

/***************************************************
PDF ICON
***************************************************/
function renderPDF(url){
 if(!url) return "";
 return `<a href="${url}" target="_blank" class="icon-pdf">📄</a>`;
}

/***************************************************
ESTADO
***************************************************/
function renderEstado(status){

 let color="#fff";

  if(status==="PENDIENTE") color="#facc15";
  if(status==="EN RUTA") color="#ef4444";
  if(status==="ENTREGADO") color="#22c55e";
  if(status==="RECIBIDO") color="#fb923c";
  if(status==="CANCELADO") color="#3b82f6";
  if(status==="TERMINADO") color="#a855f7";

 return `<span style="background:#000;color:${color};padding:3px 8px;border-radius:6px">${status||""}</span>`;
}

/***************************************************
KPI CHART
***************************************************/
function crearKPI(id,valor,total,color){

 const canvas=document.getElementById(id);
 if(!canvas) return;

 const ctx=canvas.getContext("2d");

 if(KPI_CHARTS[id]) KPI_CHARTS[id].destroy();

 KPI_CHARTS[id]=new Chart(ctx,{
  type:"doughnut",
  data:{
   datasets:[{
    data:[valor,total-valor],
    backgroundColor:[color,"#e5e7eb"],
    borderWidth:0
   }]
  },
  options:{
   responsive:true,
   maintainAspectRatio:false,
   cutout:"70%",
   plugins:{
    legend:{display:false}
   }
  }
 });

}

/***************************************************
LOAD
***************************************************/

async function load(){

  try{

    setLoading(btnReload,true);

    const r = await fetch(API);
    let rawData = await r.json();

    if(!Array.isArray(rawData)) rawData=[];

    /* 🔹 MAPEAMOS TODOS LOS CAMPOS */
    RAW = rawData.map(r => ({

      _row: r._row || "",

      fechaIngreso:
        r.fechaIngreso ||
        r['Fecha'] ||
        r['Fecha Ingreso'] ||
        r['FECHA INGRESO'] ||
        r['FECHA DE INGRESO'] ||
        "",

      pedido: r['Nº Pedido/OC'] || r.pedido || "",
      tipoDocumento: r['Tipo Documento'] || r.tipoDocumento || "",
      numeroDocumento: r['Nº Documento'] || r.numeroDocumento || "",

      cliente: r['Cliente'] || r.cliente || "",
      direccion: r['Dirección'] || r.direccion || "",
      comuna: r['Comuna'] || r.comuna || "",
      transporte: r['Transporte'] || r.transporte || "",

      etiquetas: r['Etiquetas'] || r.etiquetas || "",

      TR:
        r['TR'] ||
        r['Tr'] ||
        r['tr'] ||
        r['N° TR'] ||
        r['Nº TR'] ||
        r['Traslado'] ||
        r['TRASLADO'] ||
        r['SOLICITUD TRASLADO'] ||
        r['Solicitud Traslado'] ||
        r.TR ||
        "",

      status: r['Status'] || r.status || "PENDIENTE",

      fechaEntrega:
        r.fechaEntrega ||
        r['FECHA ENTREGA'] ||
        r['Fecha Entrega'] ||
        r['FECHA ESTIMADA ENTREGA'] ||
        r['Fecha Estimada Entrega'] ||
        r['FECHA ESTIMADA DE ENTREGA'] ||
        r['Fecha Estimada de Entrega'] ||
        "",

      responsable: r['Responsable'] || r.responsable || "",
      observaciones: r['Observaciones'] || r.observaciones || "",

      foto: r['FOTO'] || r.foto || "",
      pdf: r['PDF'] || r.pdf || "",
      pdfTraslado: r['PDF TRASLADO'] || r['PDF Traslado'] || r.pdfTraslado || "",

      alerta: r['Alerta'] || r.alerta || "",
      diasAtraso: r['Días Atraso'] || r.diasAtraso || "",

      horaEntrega:
        r.horaEntrega ||
        r.fechaEntrega ||
        r['FECHA ENTREGA'] ||
        r['FECHA ESTIMADA ENTREGA'] ||
        ""
    }));

    /* 🔹 ORDENAR DEL MAS NUEVO AL MAS ANTIGUO */
    RAW.sort((a,b)=> Number(b._row) - Number(a._row));

    applyFilters();

  }catch(e){

    console.error("ERROR LOAD:", e);
    alert("Error cargando datos: " + e.message);

  }

  setLoading(btnReload,false);

}

/***************************************************
FILTROS
***************************************************/
function applyFilters(){

  const q=(search.value||"").toLowerCase();

  FILT=RAW.filter(r=>{

   let ok=true;

   if(q){

    const txt = (
      (r.fechaIngreso || "") + " " +
      (r.pedido || "") + " " +
      (r.tipoDocumento || "") + " " +
      (r.numeroDocumento || "") + " " +
      (r.cliente || "") + " " +
      (r.direccion || "") + " " +
      (r.comuna || "") + " " +
      (r.transporte || "") + " " +
      (r.TR || "") + " " +
      (r.etiquetas || "") + " " +
      (r.status || "") + " " +
      (r.fechaEntrega || "") + " " +
      (r.alerta || "") + " " +
      (r.diasAtraso || "") + " " +
      (r.responsable || "") + " " +
      (r.observaciones || "")
    ).toLowerCase();

    ok = txt.includes(q);

   }

   if(ok && fStatus.value) ok = r.status === fStatus.value;

   if(ok && fDesde.value){
    const fechaRegistro = toDateFilterValue(r.fechaIngreso);
    const fechaDesde = toDateFilterValue(fDesde.value);
    ok = !!fechaRegistro && !!fechaDesde && fechaRegistro >= fechaDesde;
   }

   if(ok && fHasta.value){
    const fechaRegistro = toDateFilterValue(r.fechaIngreso,true);
    const fechaHasta = toDateFilterValue(fHasta.value,true);
    ok = !!fechaRegistro && !!fechaHasta && fechaRegistro <= fechaHasta;
   }

   return ok;

  });

  render();

 }

 search.oninput=applyFilters;
 fStatus.onchange=applyFilters;
 fDesde.onchange=applyFilters;
 fHasta.onchange=applyFilters;

/***************************************************
RENDER
***************************************************/
function render(){

  TOTAL_PAGES = Math.ceil(FILT.length / PAGE_SIZE) || 1;
  if(PAGE > TOTAL_PAGES) PAGE = TOTAL_PAGES;
  if(PAGE < 1) PAGE = 1;

  const start = (PAGE - 1) * PAGE_SIZE;
  const end = start + PAGE_SIZE;

  const data = FILT.slice(start,end);

  renderTable(data);
  renderCards(data);
  renderKPIs();

  renderPagination();

  if(typeof syncProductoPedidosOnRender === "function") syncProductoPedidosOnRender();

 }

/***************************************************
TABLA
***************************************************/

function renderTable(data){

  tbody.innerHTML = "";

  if(!data.length){
    tbody.innerHTML = "<tr><td colspan='20'>Sin datos</td></tr>";
    return;
  }

  data.forEach((r,i)=>{

    const fi =
      r.fechaIngreso ||
      r.Fecha ||
      r["Fecha Ingreso"] ||
      r["FECHA INGRESO"] ||
      r["FECHA DE INGRESO"] ||
      "";

    const fe =
      r.fechaEntrega ||
      r["FECHA ENTREGA"] ||
      r["Fecha Entrega"] ||
      r["FECHA ESTIMADA ENTREGA"] ||
      r["Fecha Estimada Entrega"] ||
      r["FECHA ESTIMADA DE ENTREGA"] ||
      "";

    const semaforoHtml = renderSemaforoValue(r);

    const tr = `
<tr 
  data-index="${i}" 
  onclick="selectRow(${i}); openModal(${r._row})" 
  style="cursor:pointer"
>
<td>${formatDate(fi)}</td>
<td>${r.pedido||""}</td>
<td>${r.tipoDocumento||""}</td>
<td>${r.numeroDocumento||""}</td>
<td>${r.cliente||""}</td>
<td>
  <a href="#" onclick="event.stopPropagation(); verMapa(\`${r.direccion||""}\`)">
    ${r.direccion||""}
  </a>
</td>
<td>${r.comuna||""}</td>
<td>${r.transporte||""}</td>
<td>${r.TR||""}</td>
<td>${r.etiquetas||""}</td>
<td>${renderEstado(r.status)}</td>
<td>${formatDate(fe)}</td>
<td>${renderAlerta(r.alerta)}</td>
<td>${r.diasAtraso||""}</td>
<td>${semaforoHtml}</td>
<td>${r.responsable||""}</td>
<td>
  ${
    r.foto
    ? `<img 
        src="${r.foto}" 
        class="foto-thumb" 
        onclick="event.stopPropagation(); verFoto('${r.foto}')"
      >`
    : ""
  }
</td>
<td>${renderPDF(r.pdf)}</td>
<td>${renderPDF(r.pdfTraslado)}</td>
<td class="actions">
  <button title="Editar" onclick="event.stopPropagation(); openModal(${r._row})">✏️</button>
  <button title="Productos" onclick="event.stopPropagation(); abrirModalProducto(${r._row})">📦</button>
  <button title="Eliminar" onclick="event.stopPropagation(); deleteRow(${r._row})">🗑️</button>
</td>

</tr>
`;

    tbody.insertAdjacentHTML("beforeend", tr);

  });

}

function selectRow(index){

  const rows = tbody.querySelectorAll("tr");

  rows.forEach(r => r.classList.remove("selected"));

  const row = rows[index];

  if(row){
    row.classList.add("selected");
    selectedIndex = index;

    const globalIndex = ((PAGE - 1) * PAGE_SIZE) + index;
    const rowData = FILT[globalIndex];
    SELECTED_ROW_KEY = rowData ? Number(rowData._row) : null;
  }

}

document.addEventListener("keydown", (e)=>{

  const rows = tbody.querySelectorAll("tr");
  if(!rows.length) return;

  if(e.key === "ArrowDown"){

    e.preventDefault();

    if(selectedIndex < rows.length - 1){
      selectedIndex++;
    }

    selectRow(selectedIndex);
    rows[selectedIndex].scrollIntoView({block:"nearest"});

  }

  if(e.key === "ArrowUp"){

    e.preventDefault();

    if(selectedIndex > 0){
      selectedIndex--;
    }

    selectRow(selectedIndex);
    rows[selectedIndex].scrollIntoView({block:"nearest"});

  }

  if(e.key === "Enter"){

    if(selectedIndex >= 0){
      const rowData = FILT[(PAGE-1)*PAGE_SIZE + selectedIndex];
      if(rowData){
        openModal(rowData._row);
      }
    }

  }

});

function renderCards(data){

  mobileList.innerHTML="";

  data.forEach(r=>{

   const semaforo = calcularSemaforo(r.fechaEntrega);

   let entregaHTML = "";

   if(r.status === "ENTREGADO" && r.fechaEntrega){
     const fecha = formatDate(r.fechaEntrega);
     entregaHTML = `
     <div style="margin-top:6px">
       <span style="color:#22c55e">✅</span>
       <b> Entregado:</b> ${fecha}
     </div>`;
   }

   const card = `
   <div class="card">

    <div class="card-title">
      📦 Pedido #${r.pedido||""}
      ${renderEstado(r.status)}
    </div>

    <div style="margin-top:6px">
      <span style="color:#3b82f6">👤</span>
      <b> Cliente:</b> ${r.cliente||""}
    </div>

    <div style="margin-top:6px;padding:6px 0"
         onclick="verMapa(\`${r.direccion||""}\`)">
      <span style="color:#ef4444">📍</span>
      <b> Dirección:</b> ${r.direccion||""}
    </div>

    <div style="margin-top:6px">
      <span style="color:#14b8a6">🏙️</span>
      <b> Comuna:</b> ${r.comuna||""}
    </div>

    <div style="margin-top:6px">
      <span style="color:#f97316">🚚</span>
      <b> Transporte:</b> ${r.transporte||""}
    </div>

    <div style="margin-top:6px">
      <span style="color:#10b981">🔢</span>
      <b> TR:</b> ${r.TR||""}
    </div>

    <div style="margin-top:6px">
      <span style="color:#8b5cf6">📦</span>
      <b> Unidades:</b> ${r.etiquetas||""}
    </div>

    <div style="margin-top:6px">
      <span style="color:#eab308">⏱️</span>
      <b> Semáforo:</b> ${semaforo}
    </div>

    ${entregaHTML}

    <div style="
        margin-top:12px;
        display:flex;
        gap:14px;
        align-items:center;
        flex-wrap:wrap;
    ">

      ${r.foto
        ? `<img src="${r.foto}" 
             class="foto-thumb" 
             style="width:60px;height:60px;border-radius:8px"
             onclick="verFoto('${r.foto}')">`
        : ""}

      ${r.pdf
        ? `<a href="${r.pdf}" target="_blank"
            title="PDF Documento"
            style="
            font-size:28px;
            text-decoration:none;
            background:#fee2e2;
            padding:8px 10px;
            border-radius:8px;
            ">
            📄
           </a>`
        : ""}

      ${r.pdfTraslado
        ? `<a href="${r.pdfTraslado}" target="_blank"
            title="PDF Traslado"
            style="
            font-size:28px;
            text-decoration:none;
            background:#fee2e2;
            padding:8px 10px;
            border-radius:8px;
            ">
            📄
           </a>`
        : ""}

    </div>

    <div style="
        margin-top:14px;
        display:flex;
        gap:8px;
        flex-wrap:wrap;
    ">

      <button onclick="openModal(${r._row})"
        style="padding:6px 10px">✏️ Editar</button>

      <button onclick="abrirModalProducto(${r._row})"
        style="padding:6px 10px">📦 Productos</button>

      <button onclick="deleteRow(${r._row})"
        style="padding:6px 10px">🗑️ Eliminar</button>

    </div>

   </div>
   `;

   mobileList.insertAdjacentHTML("beforeend",card);

  });

}

function renderKPIs(){

 const total=RAW.length;
 const pendientes=RAW.filter(x=>x.status==="PENDIENTE").length;
 const ruta=RAW.filter(x=>x.status==="EN RUTA").length;
 const entregado=RAW.filter(x=>x.status==="ENTREGADO").length;

 kpis.innerHTML=`

<div class="kpi"><div style="height:80px"><canvas id="k1"></canvas></div><b>${total}</b><div>Total</div></div>
<div class="kpi"><div style="height:80px"><canvas id="k2"></canvas></div><b>${pendientes}</b><div>Pendiente</div></div>
<div class="kpi"><div style="height:80px"><canvas id="k3"></canvas></div><b>${ruta}</b><div>En Ruta</div></div>
<div class="kpi"><div style="height:80px"><canvas id="k4"></canvas></div><b>${entregado}</b><div>Entregado</div></div>

`;

 crearKPI("k1",total,total,"#14b8a6");
 crearKPI("k2",pendientes,total,"#facc15");
 crearKPI("k3",ruta,total,"#ef4444");
 crearKPI("k4",entregado,total,"#22c55e");

}

/***************************************************
FOTO
***************************************************/
function verFoto(src){
 fotoGrande.src=src;
 btnDescargarFoto.onclick=()=>window.open(src);
 fotoModal.style.display="flex";
}

btnCerrarFoto.onclick=()=>fotoModal.style.display="none";

/***************************************************
MAPA
***************************************************/
function verMapa(dir){
 mapFrame.src="https://maps.google.com/maps?q="+encodeURIComponent(dir)+"&output=embed";
 mapModal.style.display="flex";
}

btnCerrarMapa.onclick=()=>mapModal.style.display="none";

/***************************************************
MODAL
***************************************************/
function openModal(row){

  EDIT=row;
  isEditing = true;
  SELECTED_ROW_KEY = Number(row);

  const data=RAW.find(r=>Number(r._row)===Number(row));
  if(!data) return;

  const mtitle=document.getElementById("mtitle");
  if(mtitle) mtitle.textContent = "Editar Pedido";

  mPedido.value=data.pedido||"";
  mTipoDoc.value=data.tipoDocumento||"";
  mNumeroDoc.value=data.numeroDocumento||"";
  mCliente.value=data.cliente||"";
  mDireccion.value=data.direccion||"";
  mComuna.value=data.comuna||"";
  mTransporte.value=data.transporte||"";
  mCajas.value=data.etiquetas||"";
  mStatus.value=data.status||"PENDIENTE";
  mResponsable.value=data.responsable||"";
  mObs.value=data.observaciones||"";

  if(data.fechaEntrega){
   mHoraEntrega.value=formatDateTimeLocalValue(data.fechaEntrega);
  }else{
   mHoraEntrega.value="";
  }

  const fotoPreview=document.getElementById("fotoPreview");

  if(fotoPreview){

   if(data.foto){
    fotoPreview.innerHTML=`
    <div class="preview-box">
    <a href="${data.foto}" target="_blank">
    <img src="${data.foto}">
    </a>
    <div class="preview-label">Imagen actual</div>
    </div>
    `;
   }else{
    fotoPreview.innerHTML="";
   }

  }

  const pdfPreview=document.getElementById("pdfPreview");

  if(pdfPreview){

   if(data.pdf){
    pdfPreview.innerHTML=`
    <div class="preview-box">
    <a href="${data.pdf}" target="_blank">
    📄 Ver documento actual
    </a>
    </div>
    `;
   }else{
    pdfPreview.innerHTML="";
   }

  }

  modalForm.style.display="flex";

}

/***************************************************
NUEVO
***************************************************/
btnNuevo.onclick = () => {

  EDIT = null;
  isEditing = false;

  modalForm.querySelectorAll("input,select,textarea").forEach(el => {
    el.value = "";
  });

  if (mFotos) mFotos.value = "";
  if (mPdf) mPdf.value = "";

  const fotoPreview = document.getElementById("fotoPreview");
  const pdfPreview = document.getElementById("pdfPreview");

  if (fotoPreview) fotoPreview.innerHTML = "";
  if (pdfPreview) pdfPreview.innerHTML = "";

  if (mStatus) mStatus.value = "PENDIENTE";

  const mtitle = document.getElementById("mtitle");
  if(mtitle) mtitle.textContent = "Nuevo Pedido";

  selectedIndex = -1;
  SELECTED_ROW_KEY = null;
  isEditing = true;

  modalForm.style.display = "flex";

  setTimeout(()=>{
    if(mPedido) mPedido.focus();
  },100);

};

/***************************************************
CANCELAR
***************************************************/
btnCancelar.onclick=()=>{ modalForm.style.display="none"; isEditing = false; };

/***************************************************
GUARDAR
***************************************************/
let guardando = false;

btnGuardar.onclick = async () => {

  if (guardando) {
    alerta("⏳ Procesando...", "warn");
    return;
  }

  guardando = true;
  setLoading(btnGuardar, true);

  try {

    let foto = "";
    let pdf = "";

    if (mFotos && mFotos.files.length) {
      const file = mFotos.files[0];

      if (file.size > 2 * 1024 * 1024) {
        alerta("⚠️ Imagen muy pesada (máx 2MB)", "warn");
        return;
      }

      foto = await fileToBase64(file);
    }

    if (mPdf && mPdf.files.length) {
      const file = mPdf.files[0];

      if (file.size > 3 * 1024 * 1024) {
        alerta("⚠️ PDF muy pesado (máx 3MB)", "warn");
        return;
      }

      pdf = await fileToBase64(file);
    }

    const data = {
      action: EDIT ? "update" : "add",
      row: EDIT || "",
      "TIPO DOCUMENTO": mTipoDoc.value,
      "NUMERO DOCUMENTO": mNumeroDoc.value,
      "CLIENTE": mCliente.value,
      "DIRECCION": mDireccion.value,
      "COMUNA": mComuna.value,
      "TRANSPORTE": mTransporte.value,
      "ETIQUETAS": mCajas.value,
      "STATUS": mStatus.value,
      "FECHA ENTREGA": mHoraEntrega.value,
      "RESPONSABLE": mResponsable.value,
      "OBSERVACIONES": mObs.value,
      FOTO: foto,
      PDF: pdf
    };

    await postPedidoData(data);

    alerta("✅ Guardado correctamente", "ok");

    setTimeout(async () => {

      modalForm.style.display = "none";
      isEditing = false;

      PAGE = 1;
      if(Number(SELECTED_ROW_KEY) === Number(EDIT)) SELECTED_ROW_KEY = null;
      await load();

    }, 500);

  } catch (err) {

    console.error("ERROR GUARDAR:", err);
    alerta("❌ Error al guardar", "error");

  } finally {

    guardando = false;
    setLoading(btnGuardar, false);

  }
};

function alerta(msg, tipo = "ok") {

  if (!msg) return;

  let container = document.getElementById("toastContainer");

  if (!container) {
    container = document.createElement("div");
    container.id = "toastContainer";

    Object.assign(container.style, {
      position: "fixed",
      top: "20px",
      right: "20px",
      zIndex: "999999",
      display: "flex",
      flexDirection: "column",
      gap: "10px"
    });

    document.body.appendChild(container);
  }

  const div = document.createElement("div");
  div.textContent = msg;

  let bg = "#16a34a";
  if (tipo === "error") bg = "#dc2626";
  if (tipo === "warn") bg = "#eab308";
  if (tipo === "info") bg = "#2563eb";

  Object.assign(div.style, {
    minWidth: "220px",
    maxWidth: "320px",
    padding: "12px 16px",
    borderRadius: "12px",
    color: tipo === "warn" ? "#000" : "#fff",
    fontSize: "13px",
    fontWeight: "600",
    boxShadow: "0 10px 25px rgba(0,0,0,.2)",
    opacity: "0",
    transform: "translateY(-10px)",
    transition: "all .3s ease",
    background: bg
  });

  container.appendChild(div);

  requestAnimationFrame(() => {
    div.style.opacity = "1";
    div.style.transform = "translateY(0)";
  });

  setTimeout(() => {
    div.style.opacity = "0";
    div.style.transform = "translateY(-10px)";

    setTimeout(() => div.remove(), 300);
  }, 2500);
}
/***************************************************
DELETE
***************************************************/
async function deleteRow(row){

  console.log("ROW A ELIMINAR:", row);

  if(!row){
    alert("Error: fila inválida");
    return;
  }

  if(!confirm("¿Eliminar registro?")) return;

  try{

    const params = new URLSearchParams();
    params.append("data", JSON.stringify({
      action: "delete",
      row: Number(row)
    }));

    const res = await fetch(API,{
      method: "POST",
      headers:{
        "Content-Type":"application/x-www-form-urlencoded"
      },
      body: params
    });

    const response = await res.json();

    console.log("RESPUESTA DELETE:", response);

    if(!response.ok){
      throw new Error(response.error || "Error al eliminar");
    }

    await load();

  }catch(err){

    console.error("ERROR DELETE:", err);
    alert("No se pudo eliminar: " + err.message);

  }

}

/* ======================================================
   EXPORTES
====================================================== */

btnPDF.onclick = ()=> exportPDF(btnPDF);
btnExcel.onclick = ()=> exportExcel(btnExcel);

/* ======================================================
   EXPORT PDF CON ENCABEZADO
====================================================== */
function exportPDF(btn){ 

  setLoading(btn,true);

  setTimeout(()=>{

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation:'landscape' });

    const totalPedidos = FILT.length;
    const totalBultos = FILT.reduce(
      (sum,r)=> sum + Number(r.etiquetas || 0), 0
    );

    doc.setFontSize(18);
    doc.text("REPORTE LOGÍSTICO DE PEDIDOS",14,15);

    doc.setFontSize(10);
    doc.text("Sistema Logístico",14,22);
    doc.text("Fecha de generación: " + new Date().toLocaleString(),14,28);

    doc.setFontSize(11);
    doc.text(`Total Pedidos: ${totalPedidos}`,250,20);
    doc.text(`Total Unidades: ${totalBultos}`,250,26);

    doc.autoTable({
      startY:35,

      head:[[
        "Fecha",
        "Pedido",
        "Tipo Documento",
        "Número Documento",
        "Cliente",
        "Dirección",
        "Comuna",
        "Transporte",
        "Unidades",
        "Responsable",
        "Fecha de Entrega",
        "Estado",
        "Nro.Tr",
      ]],

      body: FILT.map(r=>[
        formatDate(r.fechaIngreso) || '',
        r.pedido || '',
        r.tipoDocumento || '',
        r.numeroDocumento || '',
        r.cliente || '',
        r.direccion || '',
        r.comuna || '',
        r.transporte || '',
        r.etiquetas || 0,
        r.responsable || '',
        formatDate(r.fechaEntrega) || '',
        r.status || '',
        r.TR || '',
      ]),

      foot:[[
        '',
        '',
        '',
        '',
        '',
        'TOTALES →',
        `UNIDADES: ${totalBultos}`,
        '',
        '',
        `PEDIDOS: ${totalPedidos}`,
        '',
        '',
        '',
        ''
      ]],

      styles:{
        fontSize:9,
        textColor: [0,0,0]
      },

      headStyles:{
        fillColor: [220,220,220],
        textColor: [0,0,0]
      },

      footStyles:{
        fillColor: [220,220,220],
        textColor: [0,0,0]
      }

    });

    doc.save("Reporte_Pedidos_Logisticos.pdf");

    setLoading(btn,false);

  },300);
}

/* ======================================================
   EXPORT EXCEL CON ENCABEZADO
====================================================== */
function exportExcel(btn){
  setLoading(btn,true);

  setTimeout(()=>{

    const totalPedidos = FILT.length;
    const totalBultos = FILT.reduce((sum,r)=> sum + Number(r.etiquetas || 0),0);

    const encabezados = [
      "Fecha",
      "Pedido",
      "Tipo Documento",
      "Número Documento",
      "Cliente",
      "Dirección",
      "Comuna",
      "Transporte",
      "Unidades",
      "Responsable",
      "Fecha de Entrega",
      "Estado",
      "Observaciones",
      "Nro de Traslado",
    ];

    const data = FILT.map(r => [
      formatDate(r.fechaIngreso) || '',
      r.pedido || '',
      r.tipoDocumento || '',
      r.numeroDocumento || '',
      r.cliente || '',
      r.direccion || '',
      r.comuna || '',
      r.transporte || '',
      r.etiquetas || 0,
      r.responsable || '',
      formatDate(r.fechaEntrega) || '',
      r.status || '',
      r.observaciones || '',
      r.TR || '',
    ]);

    const ws = XLSX.utils.aoa_to_sheet([encabezados, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pedidos");

    const resumen = [
      ["REPORTE LOGÍSTICO"],
      [""],
      ["Fecha generación", new Date().toLocaleString()],
      ["Total Pedidos", totalPedidos],
      ["Total Unidades", totalBultos]
    ];
    const wsResumen = XLSX.utils.aoa_to_sheet(resumen);
    XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen");

    XLSX.writeFile(wb,"Reporte_Pedidos_Logisticos.xlsx");

    setLoading(btn,false);

  },300);
}

function renderPagination(){

  const pag = document.getElementById("pagination");

  if(!pag) return;

  pag.innerHTML=`

 <button onclick="prevPage()">◀</button>

 <span style="padding:0 10px">
 Página ${PAGE} de ${TOTAL_PAGES}
 </span>

 <button onclick="nextPage()">▶</button>

 `;

}

function nextPage(){

  if(PAGE < TOTAL_PAGES){
   PAGE++;
   render();
  }

}

function prevPage(){

  if(PAGE > 1){
   PAGE--;
   render();
  }

}

/* ===============================
   OCULTAR / MOSTRAR DASHBOARD
================================ */

const btnTogglePanel = document.getElementById("btnTogglePanel");
const panelDashboard = document.getElementById("panelDashboard");

if(btnTogglePanel && panelDashboard){

btnTogglePanel.onclick = () => {

const hidden = panelDashboard.classList.toggle("panel-hidden");

btnTogglePanel.textContent = hidden
? "📊 Mostrar Panel"
: "📊 Ocultar Panel";

};

}

function autoRefresh(){

  setInterval(async ()=>{

    try{

      if(isEditing) return;
      if(typeof modalProducto !== "undefined" && modalProducto && modalProducto.style.display === "flex") return;

      await load();

    }catch(err){

      console.error("AutoRefresh error:", err);

    }

  },60000);

}

/***************************************************
PRODUCTOS POR PEDIDO
***************************************************/
const PRODUCTOS_STORAGE_KEY = "productos_pedidos_v1";

const modalProducto = document.getElementById("modalProducto");
const pSelectPedido = document.getElementById("pSelectPedido");
const pPedido = document.getElementById("pPedido");
const pCliente = document.getElementById("pCliente");
const pDireccion = document.getElementById("pDireccion");
const pComuna = document.getElementById("pComuna");
const pProducto = document.getElementById("pProducto");
const pDescripcion = document.getElementById("pDescripcion");
const pCantidad = document.getElementById("pCantidad");
const pTablaProductos = document.getElementById("pTablaProductos");
const pResumenPedido = document.getElementById("pResumenPedido");
const pResumenTotales = document.getElementById("pResumenTotales");
const btnAddProducto = document.getElementById("btnAddProducto");
const btnGuardarProducto = document.getElementById("btnGuardarProducto");
const btnCerrarProducto = document.getElementById("btnCerrarProducto");
const btnGenerarProductoPDF = document.getElementById("btnGenerarProductoPDF");

let PRODUCTOS_DB = cargarProductosDB();
let PRODUCTOS_ACTUALES = [];
let PRODUCTO_ROW_ACTUAL = "";
let PEDIDO_META_ACTUAL = null;

function cargarProductosDB(){
  try{
    const raw = localStorage.getItem(PRODUCTOS_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" ? parsed : {};
  }catch(e){
    console.error("Error leyendo productos guardados:", e);
    return {};
  }
}

function persistirProductosDB(){
  try{
    localStorage.setItem(PRODUCTOS_STORAGE_KEY, JSON.stringify(PRODUCTOS_DB));
  }catch(e){
    console.error("Error guardando productos:", e);
    alerta("❌ No se pudieron guardar los productos en este navegador", "error");
  }
}

function obtenerPedidosDisponibles(){
  return Array.isArray(RAW)
    ? RAW.filter(item => item && item.pedido)
    : [];
}

function buscarPedido(ref){
  if(ref === null || ref === undefined || ref === "") return null;

  const valor = String(ref);

  return RAW.find(item =>
    String(item._row) === valor ||
    String(item.pedido) === valor
  ) || null;
}

function poblarSelectPedidos(preferido = ""){
  if(!pSelectPedido) return;

  const pedidos = obtenerPedidosDisponibles();
  const actual = preferido ? String(preferido) : String(PRODUCTO_ROW_ACTUAL || "");

  pSelectPedido.innerHTML = '<option value="">-- Seleccionar Pedido --</option>' + pedidos.map(item => {
    const etiqueta = `#${item.pedido || ""} · ${item.cliente || "Sin cliente"}`;
    return `<option value="${item._row}">${etiqueta}</option>`;
  }).join("");

  if(actual && pedidos.some(item => String(item._row) === actual)){
    pSelectPedido.value = actual;
  }
}

function limpiarInputsProducto(){
  if(pProducto) pProducto.value = "";
  if(pDescripcion) pDescripcion.value = "";
  if(pCantidad) pCantidad.value = "";
}

function limpiarCabeceraProducto(){
  if(pPedido) pPedido.value = "";
  if(pCliente) pCliente.value = "";
  if(pDireccion) pDireccion.value = "";
  if(pComuna) pComuna.value = "";
}

function totalProductosActuales(){
  return PRODUCTOS_ACTUALES.reduce((sum, item) => sum + Number(item.cantidad || 0), 0);
}

function actualizarResumenPedido(){
  if(!pResumenPedido) return;

  if(!PEDIDO_META_ACTUAL){
    pResumenPedido.textContent = "Selecciona un pedido para cargar sus datos.";
    return;
  }

  const totalLineas = PRODUCTOS_ACTUALES.length;
  const totalUnidades = totalProductosActuales();

  pResumenPedido.innerHTML = `Pedido <b>#${PEDIDO_META_ACTUAL.pedido || ""}</b> · Cliente <b>${PEDIDO_META_ACTUAL.cliente || "Sin cliente"}</b> · ${totalLineas} producto(s) · ${totalUnidades} unidad(es)`;
}

function renderTablaProductos(){
  if(!pTablaProductos) return;

  if(!PRODUCTOS_ACTUALES.length){
    pTablaProductos.innerHTML = '<tr><td colspan="4" style="padding:12px;text-align:center;color:#64748b;">No hay productos agregados.</td></tr>';
  }else{
    pTablaProductos.innerHTML = PRODUCTOS_ACTUALES.map((item, index) => `
      <tr>
        <td style="padding:8px;border-bottom:1px solid #e2e8f0;">${item.producto || ""}</td>
        <td style="padding:8px;border-bottom:1px solid #e2e8f0;">${item.detalle || item.descripcion || ""}</td>
        <td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:center;">${item.cantidad || 0}</td>
        <td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:center;">
          <button type="button" onclick="eliminarProductoActual(${index})">🗑️</button>
        </td>
      </tr>
    `).join("");
  }

  if(pResumenTotales){
    pResumenTotales.innerHTML = PRODUCTOS_ACTUALES.length
      ? `Total de líneas: <b>${PRODUCTOS_ACTUALES.length}</b> · Total de unidades: <b>${totalProductosActuales()}</b>`
      : "Aún no hay productos agregados.";
  }

  actualizarResumenPedido();
}

async function cargarProductosDePedido(meta){
  PEDIDO_META_ACTUAL = meta || null;
  if(!meta){
    PRODUCTO_ROW_ACTUAL = "";
    PRODUCTOS_ACTUALES = [];
    limpiarCabeceraProducto();
    renderTablaProductos();
    return;
  }
  PRODUCTO_ROW_ACTUAL = String(meta._row || "");
  if(pPedido) pPedido.value = meta.pedido || "";
  if(pCliente) pCliente.value = meta.cliente || "";
  if(pDireccion) pDireccion.value = meta.direccion || "";
  if(pComuna) pComuna.value = meta.comuna || "";
  PRODUCTOS_ACTUALES = [];
  renderTablaProductos();
  const guardadoLocal = PRODUCTOS_DB[String(meta.pedido)] || {};
  if(guardadoLocal && Array.isArray(guardadoLocal.items) && guardadoLocal.items.length){
    PRODUCTOS_ACTUALES = guardadoLocal.items.map(item => ({...item}));
    renderTablaProductos();
  }
  try{
    const productosRemotos = await obtenerProductosPedidoBD(meta.pedido);
    if(Array.isArray(productosRemotos) && productosRemotos.length){
      PRODUCTOS_ACTUALES = productosRemotos.map(p => ({
        producto: p.producto || "",
        detalle: p.detalle || p.descripcion || "",
        descripcion: p.descripcion || p.detalle || "",
        cantidad: Number(p.cantidad || 0)
      }));
      PRODUCTOS_DB[String(meta.pedido)] = {
        row: meta._row || "",
        pedido: meta.pedido || "",
        cliente: meta.cliente || "",
        direccion: meta.direccion || "",
        comuna: meta.comuna || "",
        updatedAt: new Date().toISOString(),
        items: PRODUCTOS_ACTUALES.map(item => ({...item})),
        totalItems: PRODUCTOS_ACTUALES.length,
        totalCantidad: totalProductosActuales()
      };
      persistirProductosDB();
      renderTablaProductos();
    }
  }catch(err){
    console.error("Error cargando productos desde backend:", err);
  }
}

function agregarProductoActual(){
  if(!PEDIDO_META_ACTUAL){
    alerta("⚠️ Primero selecciona un pedido", "warn");
    return;
  }

  const producto = (pProducto?.value || "").trim();
  const descripcion = (pDescripcion?.value || "").trim();
  const cantidad = Number(pCantidad?.value || 0);

  if(!producto){
    alerta("⚠️ Debes ingresar el nombre del producto", "warn");
    pProducto?.focus();
    return;
  }

  if(!cantidad || cantidad < 1){
    alerta("⚠️ Debes ingresar una cantidad válida", "warn");
    pCantidad?.focus();
    return;
  }

  PRODUCTOS_ACTUALES.push({
    producto,
    descripcion,
    detalle: descripcion,
    cantidad
  });

  renderTablaProductos();
  limpiarInputsProducto();
  if(pProducto) pProducto.focus();
}

function eliminarProductoActual(index){
  PRODUCTOS_ACTUALES.splice(index, 1);
  renderTablaProductos();
}
window.eliminarProductoActual = eliminarProductoActual;

async function guardarProductosPedido(){
  if(!PEDIDO_META_ACTUAL){
    alerta("⚠️ No hay pedido seleccionado", "warn");
    return;
  }
  const pedidoKey = String(PEDIDO_META_ACTUAL.pedido || "");
  PRODUCTOS_DB[pedidoKey] = {
    row: PEDIDO_META_ACTUAL._row || "",
    pedido: PEDIDO_META_ACTUAL.pedido || "",
    cliente: PEDIDO_META_ACTUAL.cliente || "",
    direccion: PEDIDO_META_ACTUAL.direccion || "",
    comuna: PEDIDO_META_ACTUAL.comuna || "",
    updatedAt: new Date().toISOString(),
    items: PRODUCTOS_ACTUALES.map(item => ({...item})),
    totalItems: PRODUCTOS_ACTUALES.length,
    totalCantidad: totalProductosActuales()
  };
  persistirProductosDB();
  try{
    const payload = {
      action: "guardarProductos",
      pedido: pedidoKey,
      productos: PRODUCTOS_ACTUALES.map(item => ({
        producto: item.producto || item.descripcion || item.detalle || "",
        detalle: item.descripcion || item.detalle || "",
        cantidad: Number(item.cantidad || 0)
      }))
    };
    const res = await fetch(API, {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if(!data.ok) throw new Error(data.error || "Error al guardar productos");
    alerta("✅ Productos guardados para el pedido #" + pedidoKey, "ok");
  }catch(e){
    console.error("Error guardando productos en backend:", e);
    alerta("⚠️ Los productos se guardaron localmente pero no en el backend", "warn");
  }
}

async function generarPDFProductos(){
  if(!PEDIDO_META_ACTUAL){
    alerta("⚠️ Selecciona un pedido antes de generar el PDF", "warn");
    return;
  }

  if(!PRODUCTOS_ACTUALES.length){
    alerta("⚠️ Debes agregar al menos un producto", "warn");
    return;
  }

  try{
    setLoading(btnGenerarProductoPDF, true);

    await guardarProductosPedido();

    const payload = construirPayloadPedido(PEDIDO_META_ACTUAL, {
      action: "generarPdfPedido",
      row: PEDIDO_META_ACTUAL._row || "",
      pedido: PEDIDO_META_ACTUAL.pedido || "",
      productos: PRODUCTOS_ACTUALES.map(item => ({
        producto: item.producto || "",
        detalle: item.detalle || item.descripcion || "",
        descripcion: item.descripcion || item.detalle || "",
        cantidad: Number(item.cantidad || 0)
      })),
      total: totalProductosActuales(),
      totalItems: PRODUCTOS_ACTUALES.length,
      PDF_NOMBRE: `Pedido_${PEDIDO_META_ACTUAL.pedido || "sin_numero"}_${Date.now()}.pdf`
    });

    const response = await postPedidoData(payload);
    const pdfUrl = response?.pdfUrl || response?.pdf || "";

    if(!pdfUrl){
      throw new Error("El backend no devolvió la URL del PDF");
    }

    const itemRaw = RAW.find(r => Number(r._row) === Number(PEDIDO_META_ACTUAL._row));
    if(itemRaw) itemRaw.pdf = pdfUrl;

    const itemFilt = FILT.find(r => Number(r._row) === Number(PEDIDO_META_ACTUAL._row));
    if(itemFilt) itemFilt.pdf = pdfUrl;

    if(EDIT && Number(EDIT) === Number(PEDIDO_META_ACTUAL._row)){
      const pdfPreview = document.getElementById("pdfPreview");
      if(pdfPreview){
        pdfPreview.innerHTML = `
          <div class="preview-box">
            <a href="${pdfUrl}" target="_blank">📄 Ver documento actual</a>
          </div>
        `;
      }
    }

    render();
    window.open(pdfUrl, "_blank");
    alerta(`✅ PDF generado con logo + QR y cargado en la columna PDF del pedido #${PEDIDO_META_ACTUAL.pedido || ""}`, "ok");
  }catch(err){
    console.error("ERROR PDF PRODUCTOS:", err);
    alerta("❌ No se pudo generar el PDF con logo y QR", "error");
  }finally{
    setLoading(btnGenerarProductoPDF, false);
  }
}

function cerrarModalProducto(){
  if(modalProducto) modalProducto.style.display = "none";
}

function abrirModalProducto(ref = ""){
  if(!modalProducto){
    alerta("❌ No se encontró el modal de productos", "error");
    return;
  }

  const pedidos = obtenerPedidosDisponibles();

  if(!pedidos.length){
    alerta("⚠️ Aún no hay pedidos cargados", "warn");
    return;
  }

  const seleccionado = buscarPedido(ref) || getSelectedRowData() || pedidos[0];

  poblarSelectPedidos(seleccionado ? seleccionado._row : "");

  if(seleccionado && pSelectPedido){
    pSelectPedido.value = String(seleccionado._row);
  }

  cargarProductosDePedido(seleccionado);
  limpiarInputsProducto();
  modalProducto.style.display = "flex";
}
window.abrirModalProducto = abrirModalProducto;

function syncProductoPedidosOnRender(){
  if(!modalProducto || modalProducto.style.display !== "flex") return;

  const referencia = PRODUCTO_ROW_ACTUAL || (pSelectPedido ? pSelectPedido.value : "");
  poblarSelectPedidos(referencia);

  const meta = buscarPedido(referencia) || getSelectedRowData();
  if(meta){
    if(pSelectPedido) pSelectPedido.value = String(meta._row);
    cargarProductosDePedido(meta);
  }
}

if(pSelectPedido){
  pSelectPedido.onchange = () => {
    const meta = buscarPedido(pSelectPedido.value);
    cargarProductosDePedido(meta);
  };
}

if(btnAddProducto) btnAddProducto.onclick = agregarProductoActual;
if(btnGuardarProducto) btnGuardarProducto.onclick = guardarProductosPedido;
if(btnCerrarProducto) btnCerrarProducto.onclick = cerrarModalProducto;
if(btnGenerarProductoPDF) btnGenerarProductoPDF.onclick = generarPDFProductos;

document.addEventListener("keydown", (e) => {
  if(e.key === "Escape" && modalProducto && modalProducto.style.display === "flex"){
    cerrarModalProducto();
  }
});

/***************************************************
INIT
***************************************************/
btnReload.onclick = load;

window.onload = () => {

  load();
  autoRefresh();

};
