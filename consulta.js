/**
 * Consulta rápida de pedidos/documentos.
 *
 * Esta consulta usa SOLO la API externa del sistema de control.
 * No debe tomar datos locales de la tabla cargada en pantalla para evitar
 * mezclar procesos o traer información desde una base incorrecta.
 */
const API_CLIENTES = "https://script.google.com/macros/s/AKfycbz2bLTTGdOOZ59w21r2zVcobqYnlFF-sVgkCRT_9CHuXu0cdARIUfAve3M7WEN-J72kfA/exec";

const btnConsultar  = document.getElementById("btnConsultar");
const inputConsulta = document.getElementById("mConsulta");
const msgConsulta   = document.getElementById("consultaMsg");

let selectResultados = document.getElementById("selectResultados");
if(!selectResultados && inputConsulta && inputConsulta.parentNode){
  selectResultados = document.createElement("select");
  selectResultados.id = "selectResultados";
  selectResultados.style.display = "none";
  selectResultados.style.minWidth = "180px";
  selectResultados.style.flex = "1 1 220px";
  inputConsulta.parentNode.insertBefore(selectResultados, inputConsulta.nextSibling);
}

const CACHE_CONSULTA = Object.create(null);
let ULTIMA_CONSULTA = "";

function normalizarConsulta(v){
  return String(v || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function setMsgConsulta(msg, color){
  if(!msgConsulta) return;
  msgConsulta.textContent = msg || "";
  if(color) msgConsulta.style.color = color;
}

function setLoadingConsulta(state){
  if(!btnConsultar) return;
  btnConsultar.disabled = state;
  btnConsultar.classList.toggle("loading", state);
}

function setValue(id, value){
  const el = document.getElementById(id);
  if(el) el.value = value == null ? "" : String(value);
}

function limpiarCamposConsulta(){
  [
    "mNumeroDoc",
    "mCliente",
    "mDireccion",
    "mComuna",
    "mTransporte",
    "mResponsable",
    "mTipoDoc"
  ].forEach(id => setValue(id, ""));
}

function adaptarItemConsulta(item){
  if(!item || typeof item !== "object") return null;
  return {
    pedido: item.pedido || item.PEDIDO || item.numeroPedido || item["Nº Pedido/OC"] || "",
    numero: item.numero || item.numeroDocumento || item.numDocumento || item["NUMERO DOCUMENTO"] || item["Nº Documento"] || "",
    tipoDocumento: item.tipoDocumento || item["TIPO DOCUMENTO"] || item["Tipo Documento"] || "",
    cliente: item.cliente || item.Cliente || item.CLIENTE || "",
    direccion: item.direccion || item.Direccion || item["Dirección"] || item.DIRECCION || "",
    comuna: item.comuna || item.Comuna || item.COMUNA || "",
    transporte: item.transporte || item.Transporte || item.TRANSPORTE || "",
    responsable: item.responsable || item.Responsable || item.RESPONSABLE || "",
    fecha: item.fecha || item.fechaEntrega || item["FECHA ENTREGA"] || item.fechaIngreso || "",
    fecreg: item.fecreg || item.fechaRegistro || item.FechaRegistro || item.fechaIngreso || item["FECHA INGRESO"] || ""
  };
}

function aplicarResultadoConsulta(item){
  if(!item) return;
  setValue("mNumeroDoc", item.numero || "");
  setValue("mCliente", item.cliente || "");
  setValue("mDireccion", item.direccion || "");
  setValue("mComuna", item.comuna || "");
  setValue("mTransporte", item.transporte || "");
  setValue("mResponsable", item.responsable || "");
  if(item.tipoDocumento) setValue("mTipoDoc", item.tipoDocumento);
}

function renderResultadosConsulta(lista){
  if(!selectResultados) return;
  selectResultados.innerHTML = "";
  if(!Array.isArray(lista) || !lista.length){
    selectResultados.style.display = "none";
    return;
  }

  const def = document.createElement("option");
  def.value = "";
  def.textContent = "Seleccione un resultado";
  selectResultados.appendChild(def);

  lista.forEach((item, index) => {
    const opt = document.createElement("option");
    opt.value = String(index);
    const etiquetaPedido = item.pedido ? `Pedido ${item.pedido} · ` : "";
    const etiquetaNumero = item.numero ? `${item.numero} · ` : "";
    opt.textContent = `${etiquetaPedido}${etiquetaNumero}${item.cliente || item.comuna || "Resultado"}`;
    selectResultados.appendChild(opt);
  });

  selectResultados.style.display = "block";
  selectResultados.onchange = function(){
    const idx = Number(this.value);
    if(Number.isNaN(idx) || !lista[idx]) return;
    aplicarResultadoConsulta(lista[idx]);
    setMsgConsulta("Resultado aplicado.", "#16a34a");
  };

  if(lista.length === 1){
    selectResultados.value = "0";
    aplicarResultadoConsulta(lista[0]);
  }
}

function buscarLocalmenteConsulta(termino){
  // Se desactiva la búsqueda local para que la consulta use únicamente
  // la base externa correspondiente a este proceso.
  return [];
}

async function fetchConsultaRemota(termino){
  const variantes = [
    `${API_CLIENTES}?q=${encodeURIComponent(termino)}`,
    `${API_CLIENTES}?pedido=${encodeURIComponent(termino)}`,
    `${API_CLIENTES}?numero=${encodeURIComponent(termino)}`,
    `${API_CLIENTES}?documento=${encodeURIComponent(termino)}`
  ];

  for(const url of variantes){
    try{
      const controller = new AbortController();
      const to = setTimeout(() => controller.abort(), 5000);
      const res = await fetch(url, { signal: controller.signal, cache: "no-store" });
      clearTimeout(to);
      if(!res.ok) continue;
      const data = await res.json();
      let lista = [];
      if(Array.isArray(data)) lista = data;
      else if(Array.isArray(data.data)) lista = data.data;
      else if(Array.isArray(data.resultados)) lista = data.resultados;
      else if(data && typeof data === 'object' && (data.numero || data.numeroDocumento || data.cliente)) lista = [data];
      lista = lista.map(adaptarItemConsulta).filter(Boolean);
      if(lista.length) return lista;
    }catch(err){}
  }
  return [];
}

async function consultarPedidoGlobal(){
  const busqueda = String(inputConsulta?.value || "").trim();
  if(!busqueda){
    setMsgConsulta("Ingrese término de búsqueda", "#dc2626");
    inputConsulta && inputConsulta.focus();
    return;
  }

  const clave = normalizarConsulta(busqueda);
  ULTIMA_CONSULTA = clave;
  setLoadingConsulta(true);
  setMsgConsulta("Consultando...", "#64748b");

  try{
    if(CACHE_CONSULTA[clave]){
      const listaCache = CACHE_CONSULTA[clave];
      renderResultadosConsulta(listaCache);
      setMsgConsulta(`${listaCache.length} resultado(s) encontrado(s)`, listaCache.length ? "#16a34a" : "#dc2626");
      return;
    }

    const remotos = await fetchConsultaRemota(busqueda);
    if(ULTIMA_CONSULTA !== clave) return;

    if(remotos.length){
      CACHE_CONSULTA[clave] = remotos;
      renderResultadosConsulta(remotos);
      setMsgConsulta(`${remotos.length} resultado(s) encontrado(s)`, "#16a34a");
    }else{
      limpiarCamposConsulta();
      renderResultadosConsulta([]);
      setMsgConsulta("No encontrado", "#dc2626");
    }
  }catch(err){
    console.error("Error consulta pedido:", err);
    setMsgConsulta("Error al consultar", "#dc2626");
  }finally{
    setLoadingConsulta(false);
  }
}

if(btnConsultar){
  btnConsultar.addEventListener("click", consultarPedidoGlobal);
}
if(inputConsulta){
  inputConsulta.addEventListener("keydown", function(e){
    if(e.key === "Enter"){
      e.preventDefault();
      consultarPedidoGlobal();
    }
  });
  inputConsulta.addEventListener("input", function(){
    if(!this.value.trim()){
      renderResultadosConsulta([]);
      setMsgConsulta("");
    }
  });
}

window.consultarPedidoGlobal = consultarPedidoGlobal;
