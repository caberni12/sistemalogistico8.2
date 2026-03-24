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
        nuevo:nuevoHabilitado
      }
    })
  });
}

/***************************************************
BTN NUEVO
***************************************************/
function actualizarVisibilidadBtnNuevo() {
  const btnNuevo = document.getElementById("btnNuevo");
  if(!btnNuevo) return;
  btnNuevo.style.display = nuevoHabilitado ? "" : "none";
}

/***************************************************
APLICAR COLUMNAS (🔥 BLOQUEO REAL)
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

      // visibilidad
      cell.style.display = visible ? "" : "none";

      // reset
      cell.style.opacity = "";
      cell.style.pointerEvents = "";
      cell.style.userSelect = "";

      // bloqueo real
      if(bloqueado){
        cell.style.opacity = "0.5";
        cell.style.pointerEvents = "none";
        cell.style.userSelect = "none";
      }

      // bloquear inputs internos
      const inputs = cell.querySelectorAll("input,select,textarea,button");
      inputs.forEach(el=>{
        el.disabled = bloqueado || (col.key === "acciones" && !accionesHabilitadas);
      });

      // acciones
      if(col.key === "acciones"){
        const botones = cell.querySelectorAll("button");
        botones.forEach(b=>{
          b.disabled = !accionesHabilitadas || bloqueado;
          b.style.opacity = (!accionesHabilitadas || bloqueado) ? "0.4" : "1";
        });
      }

    });
  });

  actualizarVisibilidadBtnNuevo();
}

/***************************************************
MODAL COLUMNAS
***************************************************/
function abrirModalColumnas(){

  const lista = document.getElementById("listaColumnas");
  const modal = document.getElementById("modalColumnas");

  if(!lista || !modal) return;

  lista.innerHTML = "";

  COLUMNAS.forEach(col=>{
    const checked = columnasVisibles[col.key] ? "checked": "";
    const locked  = columnasBloqueadas[col.key] ? "🔒":"🔓";

    lista.insertAdjacentHTML("beforeend",`
      <label style="display:flex;justify-content:space-between">
        <div>
          <input type="checkbox" data-key="${col.key}" ${checked}>
          ${col.label}
        </div>
        <button class="btnLock" data-key="${col.key}" 
        style="border:0;background:none;font-size:16px;cursor:pointer">
          ${locked}
        </button>
      </label>
    `);
  });

  modal.style.display = "flex";

  // sync checks
  document.getElementById("chkAcciones").checked = accionesHabilitadas;
  document.getElementById("chkNuevo").checked = nuevoHabilitado;

  // eventos básicos
  document.getElementById("chkAcciones").onchange = e=>{
    accionesHabilitadas = e.target.checked;
  };

  document.getElementById("chkNuevo").onchange = e=>{
    nuevoHabilitado = e.target.checked;
    actualizarVisibilidadBtnNuevo();
  };

  // 🔒 lock individual
  document.querySelectorAll(".btnLock").forEach(btn=>{
    btn.onclick = (e)=>{
      e.stopPropagation();
      const key = btn.dataset.key;
      columnasBloqueadas[key] = !columnasBloqueadas[key];
      btn.textContent = columnasBloqueadas[key] ? "🔒" : "🔓";
      aplicarColumnas();
    };
  });

  // ✅ seleccionar todo
  document.getElementById("btnSelectAll").onclick = ()=>{
    document.querySelectorAll("#listaColumnas input[data-key]").forEach(chk=>{
      chk.checked = true;
    });
  };

  // ❌ deseleccionar todo
  document.getElementById("btnUnselectAll").onclick = ()=>{
    document.querySelectorAll("#listaColumnas input[data-key]").forEach(chk=>{
      chk.checked = false;
    });
  };

  // 🔒 bloquear todo
  document.getElementById("btnBloquearTodo").onclick = ()=>{
    COLUMNAS.forEach(c=> columnasBloqueadas[c.key] = true);
    abrirModalColumnas();
  };

  // 🔓 desbloquear todo
  document.getElementById("btnDesbloquearTodo").onclick = ()=>{
    COLUMNAS.forEach(c=> columnasBloqueadas[c.key] = false);
    abrirModalColumnas();
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

  if(btnCerrar){
    btnCerrar.onclick = ()=> modal.style.display = "none";
  }

  if(btnGuardar){
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
  }

  // observer tabla
  const tbody = document.getElementById("tbody");
  if(tbody){
    let timeoutTabla = null;

    const observer = new MutationObserver(()=>{
      clearTimeout(timeoutTabla);
      timeoutTabla = setTimeout(()=>{
        aplicarColumnas();
      },100);
    });

    observer.observe(tbody,{childList:true,subtree:true});
  }

  aplicarColumnas();
});