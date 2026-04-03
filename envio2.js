//const API="https://script.google.com/macros/s/AKfycbzj3sRVqYDgGVak1PNHrycYQ6FI5Mk5UyADOL0uI4CDAprlT7LDv3ZVWrfMCkwPMCgW/exec";
//const API_OLD="https://script.google.com/macros/s/AKfycbyMhSW9JBm6zb90K1V_qHTuSZ9GqR7XNPAgV3j9upGq66OMQNK9RtEii2gT5QXlTpFD/exec";

// URL base del endpoint de Apps Script. Se actualiza cuando se publica una nueva versión.
const API="https://script.google.com/macros/s/AKfycbxzSkxz-rVSMLBEGy7k0FPd1EJpfeufZXEzxzf3JXOAQ7ONJ8O3tpxkTXYzdwDbjb7s/exec";
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
  if(!btn) return;
  btn.disabled=state;
  btn.classList.toggle("loading",state);
}

/* FECHA */

function pad2(n){
  return String(n).padStart(2,"0");
}

function parseFechaSoloDia(str,endOfDay=false){
  if(str===null || str===undefined || str==="") return null;
  if(str instanceof Date && !isNaN(str.getTime())) return new Date(str.getTime());

  const s=String(str).trim();
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

function formatFechaVista(value){
  const d=parseFechaSoloDia(value);
  if(!d) return "";
  return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`;
}

function formatFechaInput(value){
  const d=parseFechaSoloDia(value);
  if(!d) return "";
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
}

function setFechaField(el, value){
  if(!el) return;
  if((el.type || "").toLowerCase() === "date"){
    el.value = formatFechaInput(value);
  }else{
    el.value = formatFechaVista(value) || String(value || "");
  }
}

/* ALERTAS */

function calcularAlertas(r){
  if (!r.fechaEntrega) return r;

  const ahora = new Date();
  const entrega = parseFechaSoloDia(r.fechaEntrega,true);
  if(!entrega) return r;

  const diffHoras = (entrega - ahora) / (1000 * 60 * 60);
  const diffDias = Math.floor((ahora - entrega) / (1000 * 60 * 60 * 24));
  const status = (r.status || '').toString().trim().toUpperCase();

  if (status === 'ENTREGADO' || status === 'TERMINADO') {
    r.alerta = '';
    r.diasAtraso = '';
    r.statusEntrega = status === 'ENTREGADO' ? 'ENTREGADO A TIEMPO' : 'FINALIZADO';
    r.semaforo = status === 'ENTREGADO' ? 'VERDE' : 'AZUL';
    return r;
  }

  if (ahora > entrega) {
    r.alerta = 'PEDIDO ATRASADO';
    r.semaforo = 'ROJO';
    r.diasAtraso = Math.max(diffDias, 1);
    r.statusEntrega = 'ATRASADO';
  } else if (diffHoras <= 48) {
    r.alerta = 'ENTREGA EN MENOS DE 48H';
    r.semaforo = 'AMARILLO';
    r.diasAtraso = 0;
    r.statusEntrega = 'POR VENCER';
  } else {
    r.alerta = '';
    r.semaforo = 'VERDE';
    r.diasAtraso = 0;
    r.statusEntrega = 'EN TIEMPO';
  }

  return r;
}

// ------------------------------------------------------------------
//  Función auxiliar para obtener productos del backend por pedido.
// ------------------------------------------------------------------
async function obtenerProductosPedidoBD(pedido) {
  if (!pedido) return [];
  try {
    const payload = { action: "obtenerProductos", pedido: pedido };

    const params = new URLSearchParams();
    params.append("data", JSON.stringify(payload));

    const resp = await fetch(API, {
      method: "POST",
      headers: { "Content-Type":"application/x-www-form-urlencoded;charset=UTF-8" },
      body: params.toString()
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

async function load(){

  try{

    setLoading(btnReload,true);

    const r = await fetch(API);
    let data = await r.json();
    if(!Array.isArray(data)) data = [];

    RAW = data.map(row=>{

      const fechaIngreso =
        row.fechaIngreso ||
        row.Fecha ||
        row["Fecha"] ||
        row["Fecha Ingreso"] ||
        row["FECHA INGRESO"] ||
        row["FECHA DE INGRESO"] ||
        "";

      const fechaEntrega =
        row.fechaEntrega ||
        row["FECHA ENTREGA"] ||
        row["Fecha Entrega"] ||
        row["FECHA ESTIMADA ENTREGA"] ||
        row["Fecha Estimada Entrega"] ||
        row["FECHA ESTIMADA DE ENTREGA"] ||
        row["Fecha Estimada de Entrega"] ||
        "";

      let obj={
        _row:row._row || "",
        fechaIngreso:fechaIngreso,
        pedido:row.pedido || row["PEDIDO"] || row["Nº Pedido/OC"] || "",
        tipoDocumento:row.tipoDocumento || row["Tipo Documento"] || row["TIPO DOCUMENTO"] || "",
        numeroDocumento:row.numeroDocumento || row["Nº Documento"] || row["NUMERO DOCUMENTO"] || "",
        cliente:row.cliente || row["Cliente"] || row["CLIENTE"] || "",
        direccion:row.direccion || row["Dirección"] || row["DIRECCION"] || "",
        comuna:row.comuna || row["Comuna"] || row["COMUNA"] || "",
        transporte:row.transporte || row["Transporte"] || row["TRANSPORTE"] || "",
        etiquetas:row.etiquetas || row["Etiquetas"] || row["ETIQUETAS"] || "",
        observaciones:row.observaciones || row["Observaciones"] || row["OBSERVACIONES"] || "",
        status:String(row.status || row["Status"] || row["STATUS"] || "").trim().toUpperCase() || "PENDIENTE",
        fechaEntrega:fechaEntrega,
        alerta:row.alerta || row["Alerta"] || row["ALERTA"] || "",
        statusEntrega:row.statusEntrega || row["Status Entrega"] || row["STATUS ENTREGA"] || "",
        diasAtraso:row.diasAtraso || row["Días Atraso"] || row["DIAS ATRASO"] || "",
        semaforo:row.semaforo || row["Semaforo"] || row["SEMAFORO"] || "",
        responsable:row.responsable || row["Responsable"] || row["RESPONSABLE"] || "",
        foto:row.foto || row["FOTO"] || "",
        pdf:row.pdf || row["PDF"] || "",
        pdfTraslado:row.pdfTraslado || row["PDF TRASLADO"] || row["PDF Traslado"] || "",
        TR: row.TR || row["TR"] || row["N° TR"] || row["Nº TR"] || "",
        solicitudTraslado: row.solicitudTraslado || row["SOLICITUD TRASLADO"] || "",
        _fechaObj:parseFechaSoloDia(fechaIngreso)
      };

      return calcularAlertas(obj);
    });

    RAW.sort((a,b)=> Number(b._row) - Number(a._row));

    applyFilter();

  }catch(e){
    console.error("ERROR API",e);
  }

  setLoading(btnReload,false);
}

/* FILTROS */

function applyFilter(){

  const texto = (fBuscar?.value || "").toLowerCase().trim();
  const status = fStatus?.value || "";

  const d1 = fDesde?.value ? parseFechaSoloDia(fDesde.value) : null;
  const d2 = fHasta?.value ? parseFechaSoloDia(fHasta.value,true) : null;

  FILT = RAW.filter(r=>{

    const combo=(
      (r.pedido || "") + " " +
      (r.cliente || "") + " " +
      (r.comuna || "") + " " +
      (r.responsable || "") + " " +
      (r.fechaIngreso || "") + " " +
      (r.fechaEntrega || "")
    ).toLowerCase();

    if(texto && !combo.includes(texto)) return false;

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

  if(totalPedidos) totalPedidos.textContent = FILT.length;
  if(totalCajas) totalCajas.textContent = FILT.reduce((s,r)=>s+Number(r.etiquetas||0),0);

  visibleCount = 0;
  if(cardsGrid) cardsGrid.innerHTML = "";

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

function renderMore() {

  if(!cardsGrid) return;
  if(visibleCount >= FILT.length) return;

  const fragment = document.createDocumentFragment();
  const slice = FILT.slice(visibleCount, visibleCount + CHUNK);

  slice.forEach(r => {

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
        <div class="section-title">Fecha ingreso</div>
        <div class="section-value">${formatFechaVista(r.fechaIngreso)}</div>
      </div>

      <div class="section">
        <div class="section-title">Fecha entrega</div>
        <div class="section-value">${formatFechaVista(r.fechaEntrega)}</div>
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
          ${Number(r.diasAtraso) > 0 ? `<br>Días atraso: ${r.diasAtraso}` : ""}
        </div>
      ` : ""}

      <div class="card-actions">

        ${r.pdfTraslado 
          ? `<a href="${r.pdfTraslado}" target="_blank" class="btn-pdf">📄 Documento Generado</a>` 
          : ""}

        <button onclick="toggleMap('${mapId}',this)">🗺 Mapa</button>

        ${
          tieneTR
          ? ""
          : `
            ${(r.status === "ENTREGADO" || r.status === "TERMINADO")
              ? `<button onclick="openTraslado(${r._row})">📦 Traslado</button>`
              : ""}
            <button onclick="openEdit(${r._row})">✏️ Editar</button>
          `
        }

      </div>

      <div class="map-container" id="${mapId}">
        <iframe src="https://maps.google.com/maps?q=${encodeURIComponent((r.direccion || "") + " " + (r.comuna || ""))}&z=15&output=embed"></iframe>
      </div>
    `;

    fragment.appendChild(card);
  });

  cardsGrid.appendChild(fragment);
  visibleCount += CHUNK;
}

window.addEventListener("scroll", () => {
  if (window.innerHeight + window.scrollY >= document.body.offsetHeight - 400) {
    renderMore();
  }
});

function actualizarBotonPdfTraslado(row, pdfUrl, trasladoNum) {
  const cards = document.querySelectorAll(".card");
  cards.forEach(card => {
    const pedidoNode = card.querySelector(".pedido-numero");
    if(!pedidoNode) return;
    const numeroPedido = pedidoNode.textContent.replace("#", "");
    if (String(numeroPedido) === String(row)) {
      const acciones = card.querySelector(".card-actions");
      if(!acciones) return;
      const existe = acciones.querySelector(".btn-pdf");
      if (!existe) {
        const a = document.createElement("a");
        a.href = pdfUrl;
        a.target = "_blank";
        a.className = "btn-pdf";
        a.textContent = "📄 Documento Generado";
        acciones.insertBefore(a, acciones.firstChild);
      }
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

async function generarTraslado(row) {
  try {
    const response = await fetch(API, {
      method: "POST",
      body: JSON.stringify({ row }),
      headers: { "Content-Type": "application/json" }
    });

    const data = await response.json();

    if (data.ok) {
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

  const r=RAW.find(x=>Number(x._row)===Number(row));
  if(!r) return;

  if(r.status === "ENTREGADO" || r.status === "CANCELADO" || r.status === "TERMINADO"){
    alert("Este pedido está cerrado y no puede ser editado.");
    return;
  }

  EDIT_ROW=row;

  setFechaField(mFechaIngreso, r.fechaIngreso);
  if(mPedido) mPedido.value=r.pedido||"";
  if(mTipoDocumento) mTipoDocumento.value=r.tipoDocumento||"";
  if(mNumeroDocumento) mNumeroDocumento.value=r.numeroDocumento||"";
  if(mCliente) mCliente.value=r.cliente||"";
  if(mDireccion) mDireccion.value=r.direccion||"";
  if(mComuna) mComuna.value=r.comuna||"";
  if(mTransporte) mTransporte.value=r.transporte||"";
  if(mCajas) mCajas.value=r.etiquetas||1;
  if(mResponsable) mResponsable.value=r.responsable||"";
  if(mFechaEntrega) mFechaEntrega.value=formatFechaInput(r.fechaEntrega);
  if(mStatus) mStatus.value=r.status||"";

  if(mStatusEntrega) mStatusEntrega.value=r.statusEntrega||"";
  if(mSemaforo) mSemaforo.value=r.semaforo||"";
  if(mDiasAtraso) mDiasAtraso.value=r.diasAtraso||0;

  if(mObservaciones) mObservaciones.value=r.observaciones||"";

  if(mFoto) mFoto.value=r.foto||"";
  if(mPDF) mPDF.value=r.pdf||"";
  if(mPDFTraslado) mPDFTraslado.value=r.pdfTraslado||"";

  if(boxFoto) boxFoto.classList.toggle("hidden",!r.foto);
  if(boxPDF) boxPDF.classList.toggle("hidden",!r.pdf);
  if(boxTraslado) boxTraslado.classList.toggle("hidden",!r.pdfTraslado);

  if(mFechaIngreso) mFechaIngreso.disabled=true;
  if(mPedido) mPedido.disabled=true;
  if(mTipoDocumento) mTipoDocumento.disabled=true;
  if(mNumeroDocumento) mNumeroDocumento.disabled=true;
  if(mCliente) mCliente.disabled=true;
  if(mDireccion) mDireccion.disabled=true;
  if(mComuna) mComuna.disabled=true;
  if(mTransporte) mTransporte.disabled=true;
  if(mCajas) mCajas.disabled=true;
  if(mResponsable) mResponsable.disabled=true;

  if(mStatusEntrega) mStatusEntrega.disabled=true;
  if(mSemaforo) mSemaforo.disabled=true;
  if(mDiasAtraso) mDiasAtraso.disabled=true;

  if(mObservaciones) mObservaciones.disabled=false;
  if(mStatus) mStatus.disabled=false;
  if(mFechaEntrega) mFechaEntrega.disabled=false;

  if(editModal) editModal.style.display="flex";
}

/* PREVIEW */

if(previewFoto) previewFoto.onclick=()=>{ if(mFoto && mFoto.value) window.open(mFoto.value); };
if(previewPDF) previewPDF.onclick=()=>{ if(mPDF && mPDF.value) window.open(mPDF.value); };
if(previewTraslado) previewTraslado.onclick=()=>{ if(mPDFTraslado && mPDFTraslado.value) window.open(mPDFTraslado.value); };

/* GUARDAR */

async function guardar(){

  setLoading(btnGuardar,true);

  try{
    const r = RAW.find(x=>Number(x._row)===Number(EDIT_ROW));
    if(!r) throw new Error("No se encontró el registro a editar");

    const payload = {
      action: "update",
      row: Number(EDIT_ROW),

      "TIPO DOCUMENTO": r.tipoDocumento || "",
      "NUMERO DOCUMENTO": r.numeroDocumento || "",
      "CLIENTE": r.cliente || "",
      "DIRECCION": r.direccion || "",
      "COMUNA": r.comuna || "",
      "TRANSPORTE": r.transporte || "",
      "ETIQUETAS": r.etiquetas || "",
      "RESPONSABLE": r.responsable || "",

      "STATUS": mStatus ? mStatus.value : "",
      "OBSERVACIONES": mObservaciones ? mObservaciones.value : "",
      "FECHA ENTREGA": mFechaEntrega ? mFechaEntrega.value : ""
    };

    const params = new URLSearchParams();
    params.append("data", JSON.stringify(payload));

    const resp = await fetch(API,{
      method:"POST",
      headers:{ "Content-Type":"application/x-www-form-urlencoded;charset=UTF-8" },
      body: params.toString()
    });

    if(!resp.ok){
      throw new Error("Error de conexión con el servidor");
    }

    const text = await resp.text();

    let data = {};
    try{
      data = JSON.parse(text);
    }catch(e){
      console.error("Respuesta inválida del backend:", text);
      throw new Error("El servidor devolvió una respuesta inválida");
    }

    if(data.ok === false){
      throw new Error(data.error || "No se pudo guardar");
    }

    closeEdit();
    await load();
    alert("Registro actualizado correctamente");

  }catch(err){
    console.error("Error al guardar:", err);
    alert("No se pudo guardar: " + err.message);
  }finally{
    setLoading(btnGuardar,false);
  }
}

/* MAPA */

function toggleMap(id,btn){

  const el=document.getElementById(id);
  if(!el || !btn) return;

  if(!el.style.display||el.style.display==="none"){
    el.style.display="block";
    btn.textContent="➖ Ocultar mapa";
  }else{
    el.style.display="none";
    btn.textContent="🗺 Mapa";
  }
}

/* TRASLADO */

async function openTraslado(row){

  const r = RAW.find(x => Number(x._row) === Number(row));
  if(!r){
    console.warn("Pedido no encontrado",row);
    return;
  }

  TRASLADO_ROW = row;

  const modal = document.getElementById("trasladoModal");
  if(!modal){
    console.error("No existe el modal trasladoModal");
    return;
  }

  modal.style.display = "flex";

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

  const tabla = document.getElementById("detalleTable");
  if(tabla) tabla.innerHTML="";

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
      addProducto();
    }
  }catch(err){
    console.error("Error cargando productos para traslado:", err);
    addProducto();
  }
}

function addProducto(prodArg, detArg, cantArg){
  const tabla = document.getElementById("detalleTable");
  if(!tabla) return;

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

  const prodEl = document.getElementById("pProducto");
  const detEl  = document.getElementById("pDetalle");
  const cantEl = document.getElementById("pCantidad");

  const prod = prodEl ? prodEl.value.trim() : "";
  const det  = detEl ? detEl.value.trim() : "";
  const cant = cantEl ? (cantEl.value || 1) : 1;

  if(prod === ""){
    alert("Favor Informe los Productos a Trasladar y Genere Numero de Traslado");
    return;
  }

  const fila = document.createElement("tr");
  fila.innerHTML = `
    <td><input class="prod" value="${prod}" placeholder="Producto"></td>
    <td><input class="det" value="${det}" placeholder="Detalle"></td>
    <td><input type="number" class="cant" value="${cant}" min="1" oninput="calcTotal()"></td>
    <td><button type="button" onclick="this.closest('tr').remove();calcTotal()">✖</button></td>
  `;
  tabla.appendChild(fila);

  if(prodEl) prodEl.value = "";
  if(detEl) detEl.value = "";
  if(cantEl) cantEl.value = "1";
  if(prodEl) prodEl.focus();

  calcTotal();
}

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

    const productos = [];

    document.querySelectorAll("#detalleTable tr").forEach(tr=>{
      const prod = tr.querySelector(".prod")?.value?.trim() || "";
      const det  = tr.querySelector(".det")?.value?.trim() || "";
      const cant = Number(tr.querySelector(".cant")?.value || 0);

      if(prod && cant>0){
        productos.push({
          producto: prod,
          detalle: det,
          cantidad: cant
        });
      }
    });

    if(productos.length===0){
      alert("Debe agregar al menos un producto");
      return;
    }

    const logoEmpresa="https://lh3.googleusercontent.com/d/11T8x616pxgYq0QYV51JpTulsF2s4szkk";

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

    if(data.traslado){
      r.solicitudTraslado=data.traslado;
      r.TR=data.traslado;
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

/* MODAL */

function closeEdit(){
  if(editModal) editModal.style.display="none";
}

function closeTraslado() {
  const modal = document.getElementById("trasladoModal");
  if (modal) modal.style.display = "none";
  localStorage.removeItem("trasladoModalRow");
}

/* EVENTOS */

if(btnReload) btnReload.onclick=load;
if(btnGuardar) btnGuardar.onclick=guardar;

if(fBuscar) fBuscar.oninput=applyFilter;
if(fStatus) fStatus.onchange=applyFilter;
if(fDesde) fDesde.onchange=applyFilter;
if(fHasta) fHasta.onchange=applyFilter;

function modalAbierto(id){
  const el = document.getElementById(id);
  return !!(el && el.style && el.style.display === "flex");
}

/* ================= AUTO-RELOAD ================= */
let autoLoad = setInterval(() => {
  if (!modalAbierto("editModal") && !modalAbierto("trasladoModal")) {
    load();
  }
}, 15000);

let autoAlertas = setInterval(() => {
  if (!modalAbierto("editModal") && !modalAbierto("trasladoModal")) {
    RAW = RAW.map(r => calcularAlertas(r));
    applyFilter();
  }
}, 60000);

load();