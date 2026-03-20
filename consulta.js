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
CONSULTAR PEDIDO GLOBAL
************************************************/
btnConsultar.addEventListener("click", consultarPedidoGlobal);

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
    const url = API_CLIENTES + "?q=" + encodeURIComponent(busqueda); // <-- PARAMETRO GLOBAL 'q'
    const res = await fetch(url);

    if(!res.ok){
      throw new Error("Error conexión API");
    }

    const data = await res.json();

    console.log("DATA:", data);

    if(data.error || !data.data || data.data.length === 0){
      limpiarCampos();
      msgConsulta.textContent = "No encontrado";
      msgConsulta.style.color = "#dc2626";
      btnConsultar.classList.remove("loading");
      return;
    }

    // Tomamos solo el primer resultado para autocompletar (puedes adaptarlo si quieres mostrar varios)
    const fila = data.data[0];

    /************************************************
    AUTOCOMPLETAR CAMPOS
    ************************************************/
    setValue("mNumeroDoc", fila.numero);
    setValue("mCliente", fila.cliente);
    setValue("mDireccion", fila.direccion);
    setValue("mComuna", fila.comuna);
    setValue("mTransporte", fila.transporte);
    setValue("mResponsable", fila.responsable);

    // Fechas
    setValue("mFecha", fila.fecha);
    setValue("mFecReg", fila.fecreg);

    msgConsulta.textContent = `Datos cargados (${data.total} encontrados)`;
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
    consultarPedidoGlobal();
  }
});
