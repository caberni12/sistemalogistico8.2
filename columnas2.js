/***************************************************
API CONFIG
***************************************************/
const API_CONFIG = "https://script.google.com/macros/s/AKfycbxbT3crefQXWJplHog8ogP-XzSo3q22ouaDg_uNlXsGjsR9Mj7ByNLNPB59siHVZcAj/exec";

/***************************************************
COLUMNAS
***************************************************/
const COLUMNAS = [
  {key:"fechaIngreso", label:"Fecha", index:0},
  {key:"pedido", label:"Pedido", index:1},
  {key:"tipoDocumento", label:"Tipo", index:2},
  {key:"numeroDocumento", label:"N°", index:3},
  {key:"cliente", label:"Cliente", index:4},
  {key:"direccion", label:"Dirección", index:5},
  {key:"comuna", label:"Comuna", index:6},
  {key:"transporte", label:"Transporte", index:7},
  {key:"TR", label:"Nro.TR", index:8},
  {key:"etiquetas", label:"Unidades", index:9},
  {key:"status", label:"Estado", index:10},
  {key:"fechaEntrega", label:"Entrega", index:11},
  {key:"alerta", label:"Alerta", index:12},
  {key:"diasAtraso", label:"Atraso", index:13},
  {key:"semaforo", label:"Semáforo", index:14},
  {key:"responsable", label:"Responsable", index:15},
  {key:"foto", label:"Foto", index:16},
  {key:"pdf", label:"PDF", index:17},
  {key:"pdfTraslado", label:"Traslado", index:18},
  {key:"acciones", label:"Acciones", index:19}
];

let columnasVisibles = {};
let columnasBloqueadas = {};
let accionesHabilitadas = true;
let nuevoHabilitado = true;
let guardarHabilitado = true;
let cancelarHabilitado = true;

/***************************************************
LOADER
***************************************************/
function activarLoader(btn){ btn.classList.add("loading"); }
function quitarLoader(btn){ btn.classList.remove("loading"); }

/***************************************************
CARGAR CONFIG
***************************************************/
async function cargarColumnas(){
  try{
    const res = await fetch(API_CONFIG + "?action=getColumnas");
    const data = await res.json();

    columnasVisibles = data.visibles || {};
    columnasBloqueadas = data.bloqueadas || {};
    accionesHabilitadas = data.acciones ?? true;
    nuevoHabilitado = data.nuevo ?? true;
    guardarHabilitado = data.guardar ?? true;
    cancelarHabilitado = data.cancelar ?? true;

  }catch(e){
    COLUMNAS.forEach(c=>{
      columnasVisibles[c.key] = true;
      columnasBloqueadas[c.key] = false;
    });
  }
}

/***************************************************
GUARDAR CONFIG
***************************************************/
async function guardarColumnasServer(){
  await fetch(API_CONFIG,{
    method:"POST",
    body:JSON.stringify({
      action:"guardarColumnas",
      data:{
        visibles:columnasVisibles,
        bloqueadas:columnasBloqueadas,
        acciones:accionesHabilitadas,
        nuevo:nuevoHabilitado,
        guardar:guardarHabilitado,
        cancelar:cancelarHabilitado
      }
    })
  });
}

/***************************************************
APLICAR COLUMNAS + BLOQUEOS
***************************************************/
function aplicarColumnas(){

  const table = document.querySelector("table");
  if(table){
    table.querySelectorAll("tr").forEach(row=>{
      const cells = row.querySelectorAll("th,td");
      COLUMNAS.forEach(col=>{
        const cell = cells[col.index];
        if(!cell) return;

        cell.style.display = columnasVisibles[col.key] ? "" : "none";

        if(columnasBloqueadas[col.key]){
          cell.style.opacity = "0.5";
          cell.style.cursor = "not-allowed";
          cell.onclick = (e)=>{ e.stopPropagation(); e.preventDefault(); return false; };
        }else{
          cell.style.opacity = "";
          cell.style.cursor = "";
          cell.onclick = null;
        }

        if(col.key === "acciones"){
          const botones = cell.querySelectorAll("button");
          botones.forEach(b=>{
            b.disabled = !accionesHabilitadas;
            b.style.opacity = accionesHabilitadas ? "1" : "0.4";
            b.style.pointerEvents = accionesHabilitadas ? "" : "none";
          });
        }
      });
    });
  }

  // Botones principales
  const botonesMap = {
    btnNuevo: nuevoHabilitado,
    btnGuardar: guardarHabilitado,
    btnCancelar: cancelarHabilitado
  };

  Object.keys(botonesMap).forEach(id=>{
    const btn = document.getElementById(id);
    if(!btn) return;
    const activo = botonesMap[id];
    btn.disabled = !activo;
    btn.style.opacity = activo?"1":"0.5";
    btn.style.cursor = activo?"":"not-allowed";
    btn.onclick = (e)=>{
      if(!activo){ e.preventDefault(); e.stopPropagation(); return false; }
    };
  });
}

/***************************************************
MODAL
***************************************************/
function abrirModalColumnas(){

  const lista = document.getElementById("listaColumnas");
  const modal = document.getElementById("modalColumnas");
  lista.innerHTML = "";

  lista.insertAdjacentHTML("beforeend",`
    <label><input type="checkbox" id="selectAll"> Seleccionar todo</label><hr>
  `);

  // Botones principales
  lista.insertAdjacentHTML("beforeend",`
    <label><input type="checkbox" id="chkAcciones" ${accionesHabilitadas?"checked":""}> Activar acciones</label>
  `);
  lista.insertAdjacentHTML("beforeend",`
    <label><input type="checkbox" id="chkNuevo" ${nuevoHabilitado?"checked":""}> Activar NUEVO</label>
  `);
  lista.insertAdjacentHTML("beforeend",`
    <label><input type="checkbox" id="chkGuardar" ${guardarHabilitado?"checked":""}> Activar GUARDAR</label>
  `);
  lista.insertAdjacentHTML("beforeend",`
    <label><input type="checkbox" id="chkCancelar" ${cancelarHabilitado?"checked":""}> Activar CANCELAR</label><hr>
  `);

  // Columnas
  COLUMNAS.forEach(col=>{
    const checked = columnasVisibles[col.key] ? "checked":"";
    const locked  = columnasBloqueadas[col.key] ? "🔒":"🔓";
    lista.insertAdjacentHTML("beforeend",`
      <label style="display:flex;justify-content:space-between">
        <div>
          <input type="checkbox" data-key="${col.key}" ${checked}>
          ${col.label}
        </div>
        <button class="btnLock" data-key="${col.key}">${locked}</button>
      </label>
    `);
  });

  modal.style.display = "flex";

  document.getElementById("selectAll").onchange = (e)=>{
    document.querySelectorAll("#listaColumnas input[data-key]").forEach(chk=>{ chk.checked = e.target.checked; });
  };
  document.getElementById("chkAcciones").onchange = (e)=> accionesHabilitadas = e.target.checked;
  document.getElementById("chkNuevo").onchange = (e)=> nuevoHabilitado = e.target.checked;
  document.getElementById("chkGuardar").onchange = (e)=> guardarHabilitado = e.target.checked;
  document.getElementById("chkCancelar").onchange = (e)=> cancelarHabilitado = e.target.checked;

  document.querySelectorAll(".btnLock").forEach(btn=>{
    btn.onclick = ()=>{
      const key = btn.dataset.key;
      columnasBloqueadas[key] = !columnasBloqueadas[key];
      btn.textContent = columnasBloqueadas[key] ? "🔒":"🔓";
    };
  });

}

/***************************************************
INIT
***************************************************/
window.addEventListener("DOMContentLoaded", async ()=>{

  await cargarColumnas();

  const btn = document.getElementById("btnColumnas");
  const modal = document.getElementById("modalColumnas");
  const btnCerrar = document.getElementById("btnCerrarColumnas");
  const btnGuardar = document.getElementById("btnGuardarColumnas");

  btn.onclick = abrirModalColumnas;
  btnCerrar.onclick = ()=> modal.style.display = "none";

  btnGuardar.onclick = async ()=>{
    activarLoader(btnGuardar);
    document.querySelectorAll("#listaColumnas input[data-key]").forEach(chk=>{
      columnasVisibles[chk.dataset.key] = chk.checked;
    });
    await guardarColumnasServer();
    modal.style.display = "none";
    aplicarColumnas();
    quitarLoader(btnGuardar);
  };

  modal.addEventListener("click",(e)=>{
    if(e.target === modal) modal.style.display = "none";
  });

  const tbody = document.getElementById("tbody");
  if(tbody){
    const observer = new MutationObserver(()=> aplicarColumnas());
    observer.observe(tbody,{childList:true,subtree:true});
  }

  setTimeout(aplicarColumnas, 800);

});