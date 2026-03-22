/***************************************************
JS COMPLETO DE CONFIGURACIÓN DE COLUMNAS Y BOTONES
***************************************************/
const API_CONFIG = "https://script.google.com/macros/s/AKfycbxbT3crefQXWJplHog8ogP-XzSo3q22ouaDg_uNlXsGjsR9Mj7ByNLNPB59siHVZcAj/exec";

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

// ESTADO GLOBAL
let columnasVisibles = {};
let columnasBloqueadas = {};
let botonesConfig = {nuevo:true, guardar:true, cancelar:true, pdf:true, excel:true, reload:true, acciones:true};

// LOADER
function activarLoader(btn){ btn.classList.add("btn-loading"); }
function quitarLoader(btn){ btn.classList.remove("btn-loading"); }

// ------------------- CARGAR CONFIG -------------------
async function cargarConfig(){
  try{
    const res = await fetch(`${API_CONFIG}?action=getColumnas`);
    const data = await res.json();
    columnasVisibles = data.visibles || {};
    columnasBloqueadas = data.bloqueadas || {};
    botonesConfig = data.botones || botonesConfig;
  }catch(e){
    console.warn("No se pudo cargar la configuración, usando default");
    COLUMNAS.forEach(c=>{
      columnasVisibles[c.key] = true;
      columnasBloqueadas[c.key] = false;
    });
  }
}

// ------------------- GUARDAR CONFIG -------------------
async function guardarConfig(){
  try{
    await fetch(API_CONFIG,{
      method:"POST",
      body: JSON.stringify({
        action:"guardarColumnas",
        data:{visibles:columnasVisibles, bloqueadas:columnasBloqueadas, botones:botonesConfig}
      })
    });
  }catch(e){ console.error(e); }
}

// ------------------- APLICAR CONFIGURACIÓN -------------------
function aplicarTodo(){
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
          cell.onclick = e=>{ e.stopPropagation(); e.preventDefault(); };
        } else {
          cell.style.opacity = "";
          cell.style.cursor = "";
          cell.onclick = null;
        }

        if(col.key==="acciones"){
          const botones = cell.querySelectorAll("button");
          botones.forEach(b=>{
            b.disabled = !botonesConfig.acciones;
            b.style.opacity = botonesConfig.acciones?"1":"0.4";
            b.style.pointerEvents = botonesConfig.acciones?"":"none";
          });
        }
      });
    });
  }

  // Botones superiores
  const btnMap = {nuevo:"btnNuevo", guardar:"btnGuardar", cancelar:"btnCancelar", pdf:"btnPDF", excel:"btnExcel", reload:"btnReload"};
  Object.keys(btnMap).forEach(key=>{
    const btn = document.getElementById(btnMap[key]);
    if(!btn) return;
    const activo = botonesConfig[key];
    btn.disabled = !activo;
    btn.style.opacity = activo?"1":"0.5";
    btn.style.cursor = activo?"":"not-allowed";
    btn.onclick = e=>{
      if(!activo){ e.preventDefault(); e.stopPropagation(); return false; }
    };
  });
}

// ------------------- ABRIR MODAL COLUMNAS -------------------
function abrirModalColumnas(){
  const modal = document.getElementById("modalColumnas");
  const lista = document.getElementById("listaColumnas");
  lista.innerHTML = "";

  // Botones superiores
  ["nuevo","guardar","cancelar","pdf","excel","reload","acciones"].forEach(key=>{
    lista.insertAdjacentHTML("beforeend",`
      <label>
        <input type="checkbox" class="chkBtn" data-key="${key}" ${botonesConfig[key]?"checked":""}>
        ${key.charAt(0).toUpperCase() + key.slice(1)}
      </label>
    `);
  });
  lista.insertAdjacentHTML("beforeend","<hr>");

  // Columnas
  COLUMNAS.forEach(col=>{
    const checked = columnasVisibles[col.key]?"checked":"";
    const locked  = columnasBloqueadas[col.key]?"🔒":"🔓";
    lista.insertAdjacentHTML("beforeend",`
      <label style="display:flex;justify-content:space-between">
        <div>
          <input type="checkbox" data-key="${col.key}" ${checked}>
          ${col.label}
        </div>
        <button type="button" class="btnLock" data-key="${col.key}">${locked}</button>
      </label>
    `);
  });

  // Eventos Lock
  document.querySelectorAll(".btnLock").forEach(btn=>{
    btn.onclick = ()=>{
      const key = btn.dataset.key;
      columnasBloqueadas[key] = !columnasBloqueadas[key];
      btn.textContent = columnasBloqueadas[key]?"🔒":"🔓";
    };
  });

  // Eventos botones superiores
  document.querySelectorAll(".chkBtn").forEach(chk=>{
    chk.onchange = ()=>{ botonesConfig[chk.dataset.key] = chk.checked; };
  });

  modal.style.display = "flex";
}

// ------------------- INIT -------------------
window.addEventListener("DOMContentLoaded", async ()=>{
  await cargarConfig();
  aplicarTodo();

  const btnColumnas = document.getElementById("btnColumnas");
  const modal = document.getElementById("modalColumnas");
  const btnGuardar = document.getElementById("btnGuardarColumnas");
  const btnCerrar = document.getElementById("btnCerrarColumnas");

  // Guardar snapshot temporal para cancelar
  let columnasTemp = {...columnasVisibles};
  let bloqueadasTemp = {...columnasBloqueadas};
  let botonesTemp = {...botonesConfig};

  btnColumnas.onclick = ()=>{
    columnasTemp = {...columnasVisibles};
    bloqueadasTemp = {...columnasBloqueadas};
    botonesTemp = {...botonesConfig};
    abrirModalColumnas();
  };

  btnCerrar.onclick = ()=>{
    columnasVisibles = {...columnasTemp};
    columnasBloqueadas = {...bloqueadasTemp};
    botonesConfig = {...botonesTemp};
    modal.style.display = "none";
    aplicarTodo();
  };

  btnGuardar.onclick = async ()=>{
    document.querySelectorAll("#listaColumnas input[type='checkbox']").forEach(chk=>{
      columnasVisibles[chk.dataset.key] = chk.checked;
    });
    await guardarConfig();
    modal.style.display = "none";
    aplicarTodo();
  };

  modal.addEventListener("click", e=>{
    if(e.target===modal){
      columnasVisibles = {...columnasTemp};
      columnasBloqueadas = {...bloqueadasTemp};
      botonesConfig = {...botonesTemp};
      modal.style.display = "none";
      aplicarTodo();
    }
  });

  // Observador tabla (aplica cambios automáticamente)
  const tbody = document.getElementById("tbody");
  if(tbody){
    new MutationObserver(aplicarTodo).observe(tbody,{childList:true,subtree:true});
  }
});