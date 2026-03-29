/**
 * Módulo de consulta de pedidos a nivel global.
 *
 * Proporciona una interfaz para buscar pedidos en un API externo (API_CLIENTES) y
 * completar automáticamente los campos del formulario de edición de pedidos en
 * dashboard.html.  Incluye mejoras de rendimiento como caché en memoria,
 * temporizador de cancelación (timeout) y respaldo de búsqueda local en caso
 * de que la API demore o falle.
 */

/*
  CONSTANTES
  ==========

  API_CLIENTES: URL del endpoint de Apps Script que permite consultar
  información de pedidos del sistema control.  Este valor puede ser
  reemplazado por el usuario si cambia la URL de su API.
*/
const API_CLIENTES = "https://script.google.com/macros/s/AKfycbyf979_TUvV-I01GtfU837GWO5OIPJTKRo0t77QcMmZ95m9b4dR9znQIRjw16GW61gMew/exec";

/*
  SELECTORES Y ELEMENTOS DOM
  =========================

  Referencias a los elementos de la UI utilizados en el módulo.  Se
  asume que estos IDs están definidos en dashboard.html.
*/
const btnConsultar  = document.getElementById("btnConsultar");
const inputConsulta = document.getElementById("mConsulta");
const msgConsulta   = document.getElementById("consultaMsg");

// Select dinámico para mostrar resultados.  Si no existe en el DOM,
// lo creamos y lo anexamos al contenedor del input de consulta.
let selectResultados = document.getElementById("selectResultados");
if(!selectResultados){
  selectResultados = document.createElement("select");
  selectResultados.id = "selectResultados";
  selectResultados.style.marginTop = "10px";
  selectResultados.style.width = "100%";
  selectResultados.style.padding = "8px";
  selectResultados.style.display = "none";
  // Insertar a continuación del input de búsqueda
  inputConsulta && inputConsulta.parentNode && inputConsulta.parentNode.appendChild(selectResultados);
}

/*
  CAMPOS QUE SE RELLENAN
  =====================

  IDs de los campos de formulario que se actualizan cuando el usuario
  selecciona un pedido de la lista de resultados.  Se utilizan en
  limpiarCampos() y setValue().
*/
const CAMPOS_FORM = [
  "mNumeroDoc",
  "mCliente",
  "mDireccion",
  "mComuna",
  "mTransporte",
  "mResponsable",
  "mFecha",
  "mFecReg"
];

/**
 * Limpia los valores de los campos del formulario de consulta.
 */
function limpiarCampos(){
  CAMPOS_FORM.forEach(id => {
    const el = document.getElementById(id);
    if(el) el.value = "";
  });
}

/**
 * Asigna un valor a un campo del formulario si existe.
 * @param {string} id ID del elemento
 * @param {string} value Valor a asignar
 */
function setValue(id, value){
  const el = document.getElementById(id);
  if(el){
    el.value = value ? String(value) : "";
  }
}

/**
 * Renderiza la lista de resultados en el select.  Oculta el select si
 * no hay resultados.  Cada opción muestra número, cliente y comuna.
 * @param {Array<Object>} lista Lista de resultados a mostrar
 */
function renderResultados(lista){
  // Vaciar select
  selectResultados.innerHTML = "";
  if(!Array.isArray(lista) || lista.length === 0){
    selectResultados.style.display = "none";
    return;
  }
  // Opción por defecto
  const optDefault = document.createElement("option");
  optDefault.value = "";
  optDefault.textContent = "Seleccione un resultado";
  selectResultados.appendChild(optDefault);
  // Recorrer resultados
  lista.forEach((item,index) => {
    const opt = document.createElement("option");
    opt.value = index;
    opt.textContent = `${item.numero} - ${item.cliente} - ${item.comuna}`;
    selectResultados.appendChild(opt);
  });
  // Mostrar select
  selectResultados.style.display = "block";
  // Evento de selección
  selectResultados.onchange = function(){
    const i = this.value;
    if(i === "") return;
    const fila = lista[i];
    setValue("mNumeroDoc", fila.numero);
    setValue("mCliente",    fila.cliente);
    setValue("mDireccion",  fila.direccion);
    setValue("mComuna",     fila.comuna);
    setValue("mTransporte", fila.transporte);
    setValue("mResponsable",fila.responsable);
    setValue("mFecha",      fila.fecha);
    setValue("mFecReg",     fila.fecreg);
  };
}

/*
  CACHÉ DE CONSULTAS
  ===================

  Guardamos resultados de consultas previas para no volver a hacer
  peticiones idénticas al API.  La clave es la cadena de búsqueda en
  minúsculas.
*/
const cacheClientes = {};

/**
 * Convierte una respuesta del backend en un arreglo de objetos con
 * la forma esperada por renderResultados().  Se filtran claves
 * redundantes y se renombran propiedades.
 * @param {Array<Object>} data Datos crudos del API
 * @return {Array<Object>} Datos procesados
 */
function procesarData(data){
  if(!Array.isArray(data)) return [];
  return data.map(item => ({
    numero:      item.numero || item.numeroDocumento || item.numDocumento || "",
    cliente:     item.cliente || item.Cliente || "",
    direccion:   item.direccion || item.Direccion || item.Dirección || "",
    comuna:      item.comuna || item.Comuna || "",
    transporte:  item.transporte || item.Transporte || "",
    responsable: item.responsable || item.Responsable || "",
    // La API podría enviar fechas con nombre distinto o formato
    // diferente.  Las dejamos tal cual y permitimos que el
    // desarrollador las formatee en el front si lo necesita.
    fecha:       item.fecha || item.fechaEntrega || item.fechaIngreso || item.fechaPedido || "",
    fecreg:      item.fecreg || item.fechaRegistro || item.FechaRegistro || item.FechaIngreso || ""
  }));
}

/**
 * Busca pedidos en la API externa.  Aplica un timeout para abortar
 * solicitudes lentas y utiliza una caché para evitar repetir
 * solicitudes previas.  Si la API falla o no devuelve resultados,
 * intenta buscar coincidencias en la lista local de pedidos
 * cargados en la página (window.getPedidosData()) para ofrecer
 * una respuesta rápida.
 */
async function consultarPedidoGlobal(){
  const busquedaRaw = inputConsulta.value || "";
  const busqueda = busquedaRaw.trim();
  // Validar input
  if(!busqueda){
    msgConsulta.textContent = "Ingrese término de búsqueda";
    msgConsulta.style.color = "#dc2626";
    inputConsulta.focus();
    return;
  }
  // Indicar carga
  btnConsultar && (btnConsultar.disabled = true);
  btnConsultar && btnConsultar.classList.add("loading");
  msgConsulta.textContent = "Consultando...";
  msgConsulta.style.color = "#64748b";
  // Convertir clave a minúsculas para caché
  const clave = busqueda.toLowerCase();
  try{
    // 1) Intentar usar la caché si existe
    if(cacheClientes[clave]){
      renderResultados(cacheClientes[clave]);
      const totalCached = cacheClientes[clave].length;
      msgConsulta.textContent = `${totalCached} resultado${totalCached!==1?"s":""} encontrado${totalCached!==1?"s":""} (caché)`;
      msgConsulta.style.color = totalCached ? "#16a34a" : "#dc2626";
      return;
    }
    // 2) Ejecutar búsqueda en el backend con timeout
    const url = API_CLIENTES + "?q=" + encodeURIComponent(busqueda);
    const controller = new AbortController();
    const idTimeout = setTimeout(() => controller.abort(), 10000); // 10 segundos
    let res;
    try{
      res = await fetch(url, { signal: controller.signal });
      clearTimeout(idTimeout);
    }catch(fetchErr){
      clearTimeout(idTimeout);
      throw fetchErr;
    }
    // Validar respuesta
    if(!res.ok){
      throw new Error("Error conexión API");
    }
    let data;
    try{
      data = await res.json();
    }catch(parseErr){
      throw new Error("Respuesta inválida del servidor");
    }
    // Validar estructura
    const total = data && data.total != null ? Number(data.total) : (data.data ? data.data.length : 0);
    const items = data && data.data ? procesarData(data.data) : [];
    // Almacenar en caché
    cacheClientes[clave] = items;
    // Mostrar resultados
    if(!items.length){
      // Si no hay resultados en el API, intentamos buscar localmente
      const local = buscarLocalmente(busqueda);
      if(local.length){
        renderResultados(local);
        msgConsulta.textContent = `${local.length} resultado${local.length!==1?"s":""} encontrado${local.length!==1?"s":""} (local)`;
        msgConsulta.style.color = "#16a34a";
      }else{
        limpiarCampos();
        renderResultados([]);
        msgConsulta.textContent = "No encontrado";
        msgConsulta.style.color = "#dc2626";
      }
    }else{
      renderResultados(items);
      msgConsulta.textContent = `${total} resultado${total!==1?"s":""} encontrado${total!==1?"s":""}`;
      msgConsulta.style.color = "#16a34a";
    }
  }catch(err){
    console.error("ERROR en consulta global:", err);
    // Búsqueda local de emergencia
    const local = buscarLocalmente(busqueda);
    if(local.length){
      renderResultados(local);
      msgConsulta.textContent = `${local.length} resultado${local.length!==1?"s":""} encontrado${local.length!==1?"s":""} (local)`;
      msgConsulta.style.color = "#16a34a";
    }else{
      limpiarCampos();
      renderResultados([]);
      msgConsulta.textContent = err.name === 'AbortError' ? "Tiempo de espera agotado" : "Error al consultar";
      msgConsulta.style.color = "#dc2626";
    }
  }finally{
    // Restablecer estado del botón
    if(btnConsultar){
      btnConsultar.disabled = false;
      btnConsultar.classList.remove("loading");
    }
  }
}

/**
 * Busca coincidencias en la lista local de pedidos cargados en la página.
 * Se utilizan las funciones window.getPedidosData() y window.getPedidoByRow()
 * expuestas por envio.js para obtener los datos cargados.  La búsqueda
 * se realiza sobre los campos número de pedido y nombre de cliente.
 * @param {string} termino Término de búsqueda
 * @return {Array<Object>} Lista de coincidencias locales
 */
function buscarLocalmente(termino){
  const q = termino.toString().toLowerCase().trim();
  const getPedidosData = window.getPedidosData;
  if(typeof getPedidosData !== 'function'){ return []; }
  const pedidos = getPedidosData();
  if(!Array.isArray(pedidos) || !pedidos.length) return [];
  // Filtro
  const results = [];
  pedidos.forEach(item => {
    const numero  = (item.numeroDocumento || item.numero || item.numDocumento || item.pedido || "").toString();
    const cliente = (item.cliente || "").toString().toLowerCase();
    if(numero.includes(q) || cliente.includes(q)){
      results.push({
        numero:      numero,
        cliente:     item.cliente || "",
        direccion:   item.direccion || "",
        comuna:      item.comuna || "",
        transporte:  item.transporte || "",
        responsable: item.responsable || "",
        fecha:       item.fechaEntrega || item.fecha || "",
        fecreg:      item.fechaIngreso || item.fechaRegistro || ""
      });
    }
  });
  return results;
}

// Asociar eventos
if(btnConsultar){
  btnConsultar.addEventListener("click", consultarPedidoGlobal);
}
if(inputConsulta){
  inputConsulta.addEventListener("keypress", function(e){
    if(e.key === "Enter"){
      e.preventDefault();
      consultarPedidoGlobal();
    }
  });
}