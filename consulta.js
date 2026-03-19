/************************************************
API
************************************************/
const API_CLIENTES = "https://script.google.com/macros/s/AKfycbzY_pqPBF8PtLhp1cnhVQVR0iw9tAGiGdQpz56Av1rZQQ8IFmyNSiPDzRPUxA25lmOgaQ/exec";

/************************************************
DOM
************************************************/
const btnConsultar = document.getElementById("btnConsultar");
const inputConsulta = document.getElementById("mConsulta");
const msgConsulta = document.getElementById("consultaMsg");

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
CONSULTAR PEDIDO
************************************************/
btnConsultar.addEventListener("click", consultarPedido);

async function consultarPedido(){

  const numero = inputConsulta.value.trim();

  if(!numero){
    msgConsulta.textContent = "Ingrese Nº";
    msgConsulta.style.color = "#dc2626";
    inputConsulta.focus();
    return;
  }

  btnConsultar.classList.add("loading");
  msgConsulta.textContent = "Consultando...";
  msgConsulta.style.color = "#64748b";

  try{

    const url = API_CLIENTES + "?numero=" + encodeURIComponent(numero);
    const res = await fetch(url);

    if(!res.ok){
      throw new Error("Error conexión API");
    }

    const data = await res.json();

    console.log("DATA:", data);

    if(data.error){
      limpiarCampos();
      msgConsulta.textContent = "No encontrado";
      msgConsulta.style.color = "#dc2626";
      btnConsultar.classList.remove("loading");
      return;
    }

    /************************************************
    AUTOCOMPLETAR CAMPOS
    ************************************************/

    setValue("mNumeroDoc", data.numero);
    setValue("mCliente", data.cliente);
    setValue("mDireccion", data.direccion);
    setValue("mComuna", data.comuna);
    setValue("mTransporte", data.transporte);
    setValue("mResponsable", data.responsable);

    // Fechas
    setValue("mFecha", data.fecha);
    setValue("mFecReg", data.fecreg);

    msgConsulta.textContent = "Datos cargados correctamente";
    msgConsulta.style.color = "#16a34a";

  }catch(err){

    console.error("ERROR:", err);
    limpiarCampos();
    msgConsulta.textContent = "Error al consultar";
    msgConsulta.style.color = "#dc2626";

  }

  btnConsultar.classList.remove("loading");
}

/************************************************
ENTER PARA CONSULTAR
************************************************/
inputConsulta.addEventListener("keypress", function(e){
  if(e.key === "Enter"){
    e.preventDefault();
    consultarPedido();
  }
});