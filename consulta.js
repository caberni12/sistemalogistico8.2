/************************************************
API
************************************************/
const API_CLIENTES = "https://script.google.com/macros/s/AKfycbwKFSsdlYwfzCcNHpTnbMrEHkDe5BLsb5jOA0qzr08xOU6Nu7PQ57ltUpFB3fk1DnUckQ/exec";

/************************************************
DOM
************************************************/
const btnConsultar = document.getElementById("btnConsultar");
const inputConsulta = document.getElementById("mConsulta");
const msgConsulta = document.getElementById("consultaMsg");

/************************************************
CREAR SELECT DINÁMICO (SI NO EXISTE)
************************************************/
let selectResultados = document.getElementById("selectResultados");

if(!selectResultados){
  selectResultados = document.createElement("select");
  selectResultados.id = "selectResultados";
  selectResultados.style.marginTop = "10px";
  selectResultados.style.width = "100%";
  selectResultados.style.padding = "8px";
  selectResultados.style.display = "none";

  inputConsulta.parentNode.appendChild(selectResultados);
}

/************************************************
LIMPIAR CAMPOS
************************************************/
function limpiarCampos(){
  const ids = [
    "mNumeroDoc","mCliente","mDireccion","mComuna",
    "mTransporte","mResponsable","mFecha","mFecReg"
  ];

  ids.forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.value = "";
  });
}

/************************************************
SET VALUE SEGURO
************************************************/
function setValue(id, value){
  const el = document.getElementById(id);
  if(el){
    el.value = value ? String(value) : "";
  }
}

/************************************************
RENDER RESULTADOS EN SELECT
************************************************/
function renderResultados(lista){

  selectResultados.innerHTML = "";

  if(!lista || lista.length === 0){
    selectResultados.style.display = "none";
    return;
  }

  const optDefault = document.createElement("option");
  optDefault.value = "";
  optDefault.textContent = "Seleccione un resultado";
  selectResultados.appendChild(optDefault);

  lista.forEach((item, index)=>{
    const opt = document.createElement("option");
    opt.value = index;
    opt.textContent = `${item.numero} - ${item.cliente} - ${item.comuna}`;
    selectResultados.appendChild(opt);
  });

  selectResultados.style.display = "block";

  selectResultados.onchange = function(){
    const i = this.value;
    if(i === "") return;

    const fila = lista[i];

    setValue("mNumeroDoc", fila.numero);
    setValue("mCliente", fila.cliente);
    setValue("mDireccion", fila.direccion);
    setValue("mComuna", fila.comuna);
    setValue("mTransporte", fila.transporte);
    setValue("mResponsable", fila.responsable);
    setValue("mFecha", fila.fecha);
    setValue("mFecReg", fila.fecreg);
  };
}

/************************************************
CONSULTAR PEDIDO GLOBAL
************************************************/
async function consultarPedidoGlobal(){

  const busqueda = inputConsulta.value.trim();

  if(!busqueda){
    msgConsulta.textContent = "Ingrese término de búsqueda";
    msgConsulta.style.color = "#dc2626";
    inputConsulta.focus();
    return;
  }

  btnConsultar.classList.add("loading");
  msgConsulta.textContent = "Consultando...";
  msgConsulta.style.color = "#64748b";

  try{
    const url = API_CLIENTES + "?q=" + encodeURIComponent(busqueda);
    const res = await fetch(url);

    if(!res.ok){
      throw new Error("Error conexión API");
    }

    const data = await res.json();

    console.log("DATA:", data);

    if(data.error || !data.data || data.data.length === 0){
      limpiarCampos();
      renderResultados([]);
      msgConsulta.textContent = "No encontrado";
      msgConsulta.style.color = "#dc2626";
      return;
    }

    renderResultados(data.data);

    msgConsulta.textContent = `${data.total} resultados encontrados`;
    msgConsulta.style.color = "#16a34a";

  }catch(err){
    console.error("ERROR:", err);
    limpiarCampos();
    renderResultados([]);
    msgConsulta.textContent = "Error al consultar";
    msgConsulta.style.color = "#dc2626";
  }

  btnConsultar.classList.remove("loading");
}

/************************************************
EVENTOS
************************************************/
btnConsultar.addEventListener("click", consultarPedidoGlobal);

inputConsulta.addEventListener("keypress", function(e){
  if(e.key === "Enter"){
    e.preventDefault();
    consultarPedidoGlobal();
  }
});