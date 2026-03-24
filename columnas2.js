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

/***************************************************
ESTADOS
***************************************************/
let columnasVisibles = {};
let columnasBloqueadas = {};
let accionesHabilitadas = true;
let nuevoHabilitado = true;

let botonesVisibles = {
  btnNuevo:true,
  btnPDF:true,
  btnExcel:true,
  btnReload:true,
  btnColumnas:true
};

/***************************************************
LOADER
***************************************************/
function activarLoader(btn){ if(btn) btn.classList.add("loading"); }
function quitarLoader(btn){ if(btn) btn.classList.remove("loading"); }

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
    botonesVisibles = data.botones || botonesVisibles;

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
        botones:botonesVisibles
      }
    })
  });
}

/***************************************************
VISIBILIDAD BOTONES
***************************************************/
function aplicarVisibilidadBotones(){
  Object.keys(botonesVisibles).forEach(id=>{
    const btn = document.getElementById(id);
    if(btn){
      btn.style.display = botonesVisibles[id] ? "" : "none";
    }
  });
}

/***************************************************
BTN NUEVO
***************************************************/
function actualizarVisibilidadBtnNuevo(){
  const btn = document.getElementById("btnNuevo");
  if(btn) btn.style.display = nuevoHabilitado ? "" : "none";
}

/***************************************************
APLICAR COLUMNAS
***************************************************/
function aplicarColumnas(){

  const table = document.querySelector("table");
  if(!table) return;

  table.querySelectorAll("tr").forEach(row=>{
    const cells = row.querySelectorAll("th,td");

    COLUMNAS.forEach(col=>{
      const cell = cells[col.index];
      if(!cell) return;

      const visible = columnasVisibles[col.key];
      const bloqueado = columnasBloqueadas[col.key];

      cell.style.display = visible ? "" : "none";

      cell.style.opacity = bloqueado ? "0.5" : "";
      cell.style.pointerEvents = bloqueado ? "none" : "";
      cell.style.userSelect = bloqueado ? "none" : "";

      const elementos = cell.querySelectorAll("input,select,textarea,button,a");
      elementos.forEach(el=>{
        el.disabled = bloqueado || (col.key==="acciones" && !accionesHabilitadas);
      });

      if(col.key==="acciones"){
        const botones = cell.querySelectorAll("button");
        botones.forEach(b=>{
          b.disabled = !accionesHabilitadas || bloqueado;
          b.style.opacity = (!accionesHabilitadas || bloqueado) ? "0.4":"1";
        });
      }

    });
  });

  actualizarVisibilidadBtnNuevo();
  aplicarVisibilidadBotones();
}

/***************************************************
MODAL COLUMNAS
***************************************************/
function abrirModalColumnas(){

  const lista = document.getElementById("listaColumnas");
  const modal = document.getElementById("modalColumnas");
  if(!lista || !modal) return;

  lista.innerHTML = "";

  // BOTONES
  lista.insertAdjacentHTML("beforeend",`
    <h4>Botones</h4>
    <label><input type="checkbox" data-btn="btnNuevo" ${botonesVisibles.btnNuevo?"checked":""}> Nuevo</label>
    <label><input type="checkbox" data-btn="btnPDF" ${botonesVisibles.btnPDF?"checked":""}> PDF</label>
    <label><input type="checkbox" data-btn="btnExcel" ${botonesVisibles.btnExcel?"checked":""}> Excel</label>
    <label><input type="checkbox" data-btn="btnReload" ${botonesVisibles.btnReload?"checked":""}> Recargar</label>
    <label><input type="checkbox" data-btn="btnColumnas" ${botonesVisibles.btnColumnas?"checked":""}> Config</label>
    <hr>
  `);

  // CONTROLES
  lista.insertAdjacentHTML("beforeend",`
    <button id="btnSelectAll">Seleccionar todo</button>
    <button id="btnUnselectAll">Deseleccionar</button>
    <button id="btnBloquearTodo">Bloquear todo</button>
    <button id="btnDesbloquearTodo">Desbloquear todo</button>
    <hr>
  `);

  COLUMNAS.forEach(col=>{
    lista.insertAdjacentHTML("beforeend",`
      <label style="display:flex;justify-content:space-between">
        <div>
          <input type="checkbox" data-key="${col.key}" ${columnasVisibles[col.key]?"checked":""}>
          ${col.label}
        </div>
        <button class="btnLock" data-key="${col.key}">
          ${columnasBloqueadas[col.key] ? "🔒":"🔓"}
        </button>
      </label>
    `);
  });

  modal.style.display = "flex";

  // CHECKBOX COLUMNAS EN VIVO
  document.querySelectorAll("#listaColumnas input[data-key]").forEach(chk=>{
    chk.onchange = ()=>{
      columnasVisibles[chk.dataset.key] = chk.checked;
      aplicarColumnas();
    };
  });

  // CHECKBOX BOTONES EN VIVO
  document.querySelectorAll("#listaColumnas input[data-btn]").forEach(chk=>{
    chk.onchange = ()=>{
      botonesVisibles[chk.dataset.btn] = chk.checked;
      aplicarVisibilidadBotones();
    };
  });

  // LOCK
  document.querySelectorAll(".btnLock").forEach(btn=>{
    btn.onclick = e=>{
      e.stopPropagation();
      const key = btn.dataset.key;
      columnasBloqueadas[key] = !columnasBloqueadas[key];
      btn.textContent = columnasBloqueadas[key] ? "🔒":"🔓";
      aplicarColumnas();
    };
  });

  // SELECT ALL
  document.getElementById("btnSelectAll").onclick = ()=>{
    document.querySelectorAll("#listaColumnas input[data-key]").forEach(c=>{
      c.checked = true;
      columnasVisibles[c.dataset.key] = true;
    });
    aplicarColumnas();
  };

  // UNSELECT ALL
  document.getElementById("btnUnselectAll").onclick = ()=>{
    document.querySelectorAll("#listaColumnas input[data-key]").forEach(c=>{
      c.checked = false;
      columnasVisibles[c.dataset.key] = false;
    });
    aplicarColumnas();
  };

  // BLOQUEAR TODO
  document.getElementById("btnBloquearTodo").onclick = ()=>{
    COLUMNAS.forEach(c=>columnasBloqueadas[c.key]=true);
    abrirModalColumnas();
    aplicarColumnas();
  };

  // DESBLOQUEAR TODO
  document.getElementById("btnDesbloquearTodo").onclick = ()=>{
    COLUMNAS.forEach(c=>columnasBloqueadas[c.key]=false);
    abrirModalColumnas();
    aplicarColumnas();
  };
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

  if(btn) btn.onclick = abrirModalColumnas;
  if(btnCerrar) btnCerrar.onclick = ()=> modal.style.display="none";

  if(btnGuardar){
    btnGuardar.onclick = async ()=>{
      activarLoader(btnGuardar);

      await guardarColumnasServer();

      modal.style.display="none";
      aplicarColumnas();

      quitarLoader(btnGuardar);
    };
  }

  const tbody = document.getElementById("tbody");
  if(tbody){
    let t;
    const observer = new MutationObserver(()=>{
      clearTimeout(t);
      t = setTimeout(aplicarColumnas,100);
    });
    observer.observe(tbody,{childList:true,subtree:true});
  }

  aplicarColumnas();
});