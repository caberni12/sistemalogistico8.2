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
ESTADO GLOBAL
***************************************************/
let columnasVisibles = {};
let columnasBloqueadas = {};

let botonesConfig = {
  nuevo:true,
  guardar:true,
  cancelar:true,
  pdf:true,
  excel:true,
  reload:true,
  acciones:true
};

/***************************************************
LOADER
***************************************************/
function activarLoader(btn){ btn.classList.add("loading"); }
function quitarLoader(btn){ btn.classList.remove("loading"); }

/***************************************************
CARGAR CONFIG
***************************************************/
async function cargarConfig(){
  try{
    const res = await fetch(API_CONFIG + "?action=getColumnas");
    const data = await res.json();

    columnasVisibles = data.visibles || {};
    columnasBloqueadas = data.bloqueadas || {};
    botonesConfig = data.botones || botonesConfig;

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
async function guardarConfig(){
  await fetch(API_CONFIG,{
    method:"POST",
    body:JSON.stringify({
      action:"guardarColumnas",
      data:{
        visibles:columnasVisibles,
        bloqueadas:columnasBloqueadas,
        botones:botonesConfig
      }
    })
  });
}

/***************************************************
APLICAR TODO
***************************************************/
function aplicarTodo(){

  /******** TABLA ********/
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

          cell.onclick = (e)=>{
            e.stopPropagation();
            e.preventDefault();
          };

        }else{
          cell.style.opacity = "";
          cell.style.cursor = "";
          cell.onclick = null;
        }

        if(col.key === "acciones"){
          const botones = cell.querySelectorAll("button");

          botones.forEach(b=>{
            b.disabled = !botonesConfig.acciones;
            b.style.opacity = botonesConfig.acciones ? "1":"0.4";
            b.style.pointerEvents = botonesConfig.acciones ? "" : "none";
          });
        }

      });

    });
  }

  /******** BOTONES SUPERIORES ********/
  aplicarBoton("btnNuevo", botonesConfig.nuevo);
  aplicarBoton("btnGuardar", botonesConfig.guardar);
  aplicarBoton("btnCancelar", botonesConfig.cancelar);
  aplicarBoton("btnPDF", botonesConfig.pdf);
  aplicarBoton("btnExcel", botonesConfig.excel);
  aplicarBoton("btnReload", botonesConfig.reload);
}

/***************************************************
CONTROL BOTÓN
***************************************************/
function aplicarBoton(id, activo){

  const btn = document.getElementById(id);
  if(!btn) return;

  btn.disabled = !activo;
  btn.style.opacity = activo ? "1":"0.5";
  btn.style.cursor = activo ? "" : "not-allowed";

  btn.onclick = (e)=>{
    if(!activo){
      e.preventDefault();
      e.stopPropagation();
      return false;
    }
  };

}

/***************************************************
MODAL
***************************************************/
function abrirModalColumnas(){

  const lista = document.getElementById("listaColumnas");
  const modal = document.getElementById("modalColumnas");

  lista.innerHTML = "";

  lista.insertAdjacentHTML("beforeend",`
    <h4>Botones</h4>
    ${crearCheck("nuevo","Nuevo")}
    ${crearCheck("guardar","Guardar")}
    ${crearCheck("cancelar","Cancelar")}
    ${crearCheck("pdf","Exportar PDF")}
    ${crearCheck("excel","Exportar Excel")}
    ${crearCheck("reload","Recargar")}
    ${crearCheck("acciones","Acciones Tabla")}
    <hr>
  `);

  lista.insertAdjacentHTML("beforeend",`
    <label><input type="checkbox" id="selectAll"> Seleccionar todo</label>
    <hr>
  `);

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

  /******** EVENTOS ********/
  document.getElementById("selectAll").onchange = (e)=>{
    document.querySelectorAll("#listaColumnas input[data-key]").forEach(chk=>{
      chk.checked = e.target.checked;
    });
  };

  document.querySelectorAll(".chkBtn").forEach(chk=>{
    chk.onchange = ()=>{
      botonesConfig[chk.dataset.key] = chk.checked;
    };
  });

  document.querySelectorAll(".btnLock").forEach(btn=>{
    btn.onclick = ()=>{
      const key = btn.dataset.key;
      columnasBloqueadas[key] = !columnasBloqueadas[key];
      btn.textContent = columnasBloqueadas[key] ? "🔒":"🔓";
    };
  });

}

/***************************************************
CHECK GENERADOR
***************************************************/
function crearCheck(key,label){
  return `
    <label>
      <input type="checkbox" class="chkBtn" data-key="${key}" ${botonesConfig[key]?"checked":""}>
      ${label}
    </label>
  `;
}

/***************************************************
INIT
***************************************************/
window.addEventListener("DOMContentLoaded", async ()=>{

  await cargarConfig();

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

    await guardarConfig();

    modal.style.display = "none";
    aplicarTodo();

    quitarLoader(btnGuardar);
  };

  modal.addEventListener("click",(e)=>{
    if(e.target === modal){
      modal.style.display = "none";
    }
  });

  const tbody = document.getElementById("tbody");

  if(tbody){
    new MutationObserver(aplicarTodo)
    .observe(tbody,{childList:true,subtree:true});
  }

  setTimeout(aplicarTodo, 800);

});