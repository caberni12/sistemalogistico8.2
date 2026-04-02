/* =====================================================
   APP PRO UNIFICADO – CORREGIDO Y FUNCIONAL
   ===================================================== */

/* ===================== APIs ===================== */
// SOLO CONSULTA (maestra / búsqueda / importar)
const API = "https://script.google.com/macros/s/AKfycbzC_qrSyXeTw9NcO40ap4x2cfs3FZIBKqMZLV9kKhYYh7n2XTPAuj1Vb2ckpFBWi8Ys/exec";

// SOLO GUARDAR CAPTURAS
const API_GUARDAR = "https://script.google.com/macros/s/AKfycbz-_cZbe36eaQyopjw1HURuE4Zwbvuo4Lewsn0S393ocCLiQRbdouSUwpiAFOSwVzXwyA/exec";

/* ===================== ESTADO ===================== */
let productos = [];
let capturas = JSON.parse(localStorage.getItem("capturas") || "[]");

let scannerActivo = null;
let scannerTarget = null;
let torch = false;
let editIndex = -1;

let timerConsulta = null;
let filasConsulta = [];
let indexConsulta = -1;

let bufferImportacion = JSON.parse(localStorage.getItem("bufferImportacion") || "null");
let estadoImportacion = JSON.parse(localStorage.getItem("estadoImportacion") || "null");

let ajustePendiente = null;
let ultimoTapPorInput = {};
let ultimoInputScanner = null;

/* ===================== HELPERS DOM ===================== */
const $ = (id) => document.getElementById(id);

function el(id) {
  return document.getElementById(id);
}

function esMobile() {
  return /android|iphone|ipad|ipod/i.test(navigator.userAgent);
}

function obtenerJSONSeguro(texto) {
  try {
    return JSON.parse(texto);
  } catch {
    return null;
  }
}

function formatearFechaLocal() {
  return new Date().toLocaleString("es-CL");
}

function hoyISO() {
  return new Date().toISOString().slice(0, 10);
}

function safeTrim(valor) {
  return String(valor || "").trim();
}

function safeLower(valor) {
  return String(valor || "").trim().toLowerCase();
}

/* ===================== CARGA INICIAL ===================== */
document.addEventListener("DOMContentLoaded", async () => {
  const operador = el("operador");
  const ubicacion = el("ubicacion");
  const inputCodigo = el("codigo");
  const inputUbicacion = el("ubicacion");
  const barra = el("barra");
  const mensaje = el("mensaje");

  if (operador) operador.value = localStorage.getItem("operador") || "";
  if (ubicacion) ubicacion.value = localStorage.getItem("ubicacion") || "";

  await cargarProductosInicial();

  render();
  configurarInputsScanner();
  aplicarModoDesktop();

  if (estadoImportacion && estadoImportacion.enProceso) {
    openTab("importar");
    if (barra) barra.style.width = (estadoImportacion.progreso || 0) + "%";
    if (mensaje) mensaje.innerText = estadoImportacion.mensaje || "⏳ Importando...";
  }

  if (inputCodigo) {
    inputCodigo.addEventListener("input", () => {
      buscarDescripcion();
      previewIngreso();
    });
  }

  if (inputUbicacion) {
    inputUbicacion.addEventListener("input", previewIngreso);
  }

  ["descripcion", "cantidad", "operador"].forEach((id) => {
    const nodo = el(id);
    if (nodo) nodo.addEventListener("input", previewIngreso);
  });
});

/* ===================== PRODUCTOS ===================== */
async function cargarProductosInicial() {
  try {
    const r = await fetch(API);
    if (!r.ok) throw new Error("No se pudo consultar API");
    const d = await r.json();
    productos = Array.isArray(d) ? d : [];
    localStorage.setItem("productos", JSON.stringify(productos));
  } catch {
    const cache = localStorage.getItem("productos");
    productos = cache ? JSON.parse(cache) : [];
  }
}

/* ===================== TABS ===================== */
function openTab(id) {
  document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
  const tab = el(id);
  if (tab) tab.classList.add("active");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ===================== CAPTURADOR ===================== */
function limpiarUbicacion() {
  const ubicacion = el("ubicacion");
  if (!ubicacion) return;
  ubicacion.value = "";
  localStorage.removeItem("ubicacion");
  previewIngreso();
}

function buscarDescripcion() {
  const codigo = el("codigo");
  const descripcion = el("descripcion");
  const cantidad = el("cantidad");

  if (!codigo || !descripcion || !cantidad) return;

  const c = safeLower(codigo.value);
  if (!c) return;

  const p = productos.find((x) => safeLower(x.CODIGO) === c);
  if (p) {
    descripcion.value = p.DESCRIPCION || "";
    if (!cantidad.value || Number(cantidad.value) <= 0) cantidad.value = 1;
  }
}

function previewIngreso() {
  const codigo = el("codigo");
  const descripcion = el("descripcion");
  const ubicacion = el("ubicacion");
  const operador = el("operador");
  const cantidad = el("cantidad");
  const preview = el("preview");

  if (!codigo || !descripcion || !ubicacion || !operador || !cantidad || !preview) return;

  if (!codigo.value && !descripcion.value) {
    preview.innerHTML = "";
    return;
  }

  preview.innerHTML = `
    <div class="row preview">
      <b>🕒 PREVISUALIZANDO</b><br><br>
      <b>${codigo.value || "-"}</b> – ${descripcion.value || "-"}<br>
      <span class="small">
        ${ubicacion.value || "SIN UBICACIÓN"} |
        ${operador.value || "-"} |
        Cant: ${cantidad.value || 1}
      </span>
    </div>
  `;
}

/* ===================== SCANNER ===================== */
function activarScan(tipo) {
  if (typeof Html5Qrcode === "undefined") {
    alert("❌ Librería Html5Qrcode no encontrada");
    return;
  }

  cerrarScanner();
  scannerTarget = tipo;
  ultimoInputScanner = tipo;

  const cont = el("scanner-" + tipo);
  if (!cont) return;

  cont.innerHTML = `<div id="scannerBox"></div>`;

  scannerActivo = new Html5Qrcode("scannerBox");

  scannerActivo.start(
    { facingMode: "environment" },
    { fps: 10, qrbox: 260 },
    (txt) => {
      if (!txt) return;

      el("beep")?.play?.();
      navigator.vibrate?.(150);

      if (scannerTarget === "codigo") {
        const codigo = el("codigo");
        if (codigo) codigo.value = txt;
        buscarDescripcion();
      }

      if (scannerTarget === "ubicacion") {
        const ubicacion = el("ubicacion");
        if (ubicacion) {
          ubicacion.value = txt;
          localStorage.setItem("ubicacion", txt);
        }
      }

      previewIngreso();
      cerrarScanner();
    },
    () => {}
  ).then(() => {
    setTimeout(agregarOverlayScanner, 250);
    programarAjuste();
  }).catch(() => {
    cerrarScanner();
    alert("❌ No fue posible iniciar la cámara");
  });
}

function agregarOverlayScanner() {
  const box = el("scannerBox");
  if (!box) return;

  if (box.querySelector(".scanner-overlay")) return;

  box.insertAdjacentHTML("afterbegin", `
    <div class="scanner-overlay">
      <button class="scanner-btn close" type="button" onclick="cerrarScanner()">✖</button>
      <button class="scanner-btn torch" type="button" onclick="toggleTorch()">🔦</button>
      <button class="scanner-btn keyboard" type="button" onclick="abrirTecladoDesdeScanner()">⌨️</button>
    </div>
    <div class="scan-frame"></div>
    <div class="scan-line"></div>
  `);
}

function cerrarScanner() {
  if (!scannerActivo) {
    document.querySelectorAll(".scanner-slot").forEach((d) => d.innerHTML = "");
    scannerTarget = null;
    torch = false;
    return;
  }

  const actual = scannerActivo;
  scannerActivo = null;

  actual.stop()
    .then(() => actual.clear())
    .catch(() => {})
    .finally(() => {
      scannerTarget = null;
      torch = false;
      document.querySelectorAll(".scanner-slot").forEach((d) => d.innerHTML = "");
    });
}

function toggleTorch() {
  if (!scannerActivo) return;

  torch = !torch;

  scannerActivo.applyVideoConstraints({
    advanced: [{ torch }]
  }).then(() => {
    document.querySelector(".scanner-btn.torch")?.classList.toggle("on", torch);
  }).catch(() => {
    torch = false;
    document.querySelector(".scanner-btn.torch")?.classList.remove("on");
  });
}

function abrirTecladoDesdeScanner() {
  if (!ultimoInputScanner) return;

  const input = el(ultimoInputScanner);
  if (!input) return;

  cerrarScanner();
  setTimeout(() => {
    input.removeAttribute("readonly");
    input.focus();
  }, 120);
}

/* ===================== AJUSTE SCANNER / TECLADO ===================== */
if (window.visualViewport) {
  window.visualViewport.addEventListener("resize", () => {
    const hWin = window.innerHeight;
    const hVV = window.visualViewport.height;
    const tecladoVisible = hVV < hWin - 120;
    if (tecladoVisible) programarAjuste();
  });
}

function programarAjuste() {
  clearTimeout(ajustePendiente);
  ajustePendiente = setTimeout(ajustarScanner, 350);
}

function ajustarScanner() {
  const box = el("scannerBox");
  if (!box) return;

  const rect = box.getBoundingClientRect();
  const offset = window.visualViewport?.offsetTop || 0;
  const y = rect.top + window.scrollY - offset - 16;

  window.scrollTo({
    top: Math.max(0, y),
    behavior: "smooth"
  });
}

document.addEventListener("focusin", (e) => {
  if (e.target instanceof HTMLInputElement) {
    programarAjuste();
  }
});

/* ===================== INPUTS SCANNER ===================== */
function configurarInputsScanner() {
  ["codigo", "ubicacion"].forEach((id) => {
    const input = el(id);
    if (!input) return;

    if (!esMobile()) {
      input.removeAttribute("readonly");
      return;
    }

    input.setAttribute("readonly", "true");

    const handler = (e) => {
      const now = Date.now();
      const lastTap = ultimoTapPorInput[id] || 0;
      const diff = now - lastTap;
      ultimoTapPorInput[id] = now;

      if (diff < 350) {
        e.preventDefault();
        input.blur();
        activarScan(id);
        return;
      }

      cerrarScanner();
      input.removeAttribute("readonly");
      input.focus();

      setTimeout(() => {
        if (document.activeElement !== input) {
          input.setAttribute("readonly", "true");
        }
      }, 800);
    };

    input.addEventListener("touchend", handler, { passive: false });
    input.addEventListener("click", handler);
    input.addEventListener("blur", () => {
      if (esMobile()) input.setAttribute("readonly", "true");
    });
  });
}

function aplicarModoDesktop() {
  if (esMobile()) return;

  const inputCodigo = el("codigo");
  const inputUbicacion = el("ubicacion");

  if (inputCodigo) {
    inputCodigo.removeAttribute("readonly");
    inputCodigo.style.pointerEvents = "auto";
  }

  if (inputUbicacion) {
    inputUbicacion.removeAttribute("readonly");
    inputUbicacion.style.pointerEvents = "auto";
  }
}

/* ===================== GUARDAR / EDITAR ===================== */
function ingresar() {
  const codigo = el("codigo");
  const descripcion = el("descripcion");
  const cantidad = el("cantidad");
  const operador = el("operador");
  const ubicacion = el("ubicacion");

  if (!codigo || !descripcion || !cantidad || !operador || !ubicacion) return;

  if (!safeTrim(codigo.value)) {
    alert("❌ Digite un código válido");
    return;
  }

  localStorage.setItem("operador", operador.value || "");
  if (ubicacion.value) localStorage.setItem("ubicacion", ubicacion.value);
  else localStorage.removeItem("ubicacion");

  const d = {
    Fecha: formatearFechaLocal(),
    Operador: operador.value || "",
    Ubicación: ubicacion.value || "SIN UBICACIÓN",
    Código: codigo.value || "",
    Descripción: descripcion.value || "",
    Cantidad: Number(cantidad.value || 1)
  };

  if (editIndex >= 0) capturas[editIndex] = d;
  else capturas.push(d);

  editIndex = -1;
  localStorage.setItem("capturas", JSON.stringify(capturas));

  limpiar();
  render();
}

function cancelarEdicion() {
  editIndex = -1;
  limpiar();
  render();
}

function limpiar() {
  const codigo = el("codigo");
  const descripcion = el("descripcion");
  const cantidad = el("cantidad");
  const preview = el("preview");

  if (codigo) codigo.value = "";
  if (descripcion) descripcion.value = "";
  if (cantidad) cantidad.value = 1;
  if (preview) preview.innerHTML = "";
}

function cargarParaEditar(i) {
  const operador = el("operador");
  const ubicacion = el("ubicacion");
  const codigo = el("codigo");
  const descripcion = el("descripcion");
  const cantidad = el("cantidad");

  const c = capturas[i];
  if (!c) return;

  if (operador) operador.value = c.Operador || "";
  if (ubicacion) ubicacion.value = c.Ubicación === "SIN UBICACIÓN" ? "" : (c.Ubicación || "");
  if (codigo) codigo.value = c.Código || "";
  if (descripcion) descripcion.value = c.Descripción || "";
  if (cantidad) cantidad.value = c.Cantidad || 1;

  editIndex = i;
  previewIngreso();
  render();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function eliminarItem(i) {
  if (!confirm("¿Eliminar este registro?")) return;
  capturas.splice(i, 1);
  localStorage.setItem("capturas", JSON.stringify(capturas));
  if (editIndex === i) editIndex = -1;
  render();
}

/* ===================== RENDER ===================== */
function render() {
  const tabla = el("tabla");
  const totalizador = el("totalizador");
  if (!tabla || !totalizador) return;

  tabla.innerHTML = "";
  let total = 0;

  capturas.forEach((c, i) => {
    total += Number(c.Cantidad) || 0;

    tabla.innerHTML += `
      <div class="row ${editIndex === i ? "editing" : ""}">
        <button class="delbtn" type="button" onclick="event.stopPropagation();eliminarItem(${i})">×</button>
        <div onclick="cargarParaEditar(${i})">
          <b>${c.Código || ""}</b> – ${c.Descripción || ""}<br>
          <span class="small">
            ${c.Ubicación || "SIN UBICACIÓN"} | ${c.Operador || ""} | ${c.Fecha || ""} | Cant: ${c.Cantidad || 0}
          </span>
        </div>
      </div>
    `;
  });

  totalizador.innerText = "Total unidades: " + total;
}

/* ===================== EXPORTAR PDF ===================== */
function exportarPDF() {
  if (!capturas.length) {
    alert("❌ Sin datos");
    return;
  }

  const w = window.open("", "_blank");
  if (!w) {
    alert("❌ El navegador bloqueó la ventana");
    return;
  }

  let h = `
    <html>
    <head>
      <title>Reporte de Captura</title>
      <style>
        body{font-family:Arial;padding:20px}
        table{border-collapse:collapse;width:100%}
        th,td{border:1px solid #000;padding:8px;font-size:12px}
        th{background:#f0f0f0}
      </style>
    </head>
    <body>
      <h3>Reporte de Captura</h3>
      <table>
        <tr>
  `;

  Object.keys(capturas[0]).forEach((k) => h += `<th>${k}</th>`);
  h += `</tr>`;

  capturas.forEach((r) => {
    h += `<tr>`;
    Object.values(r).forEach((v) => h += `<td>${v ?? ""}</td>`);
    h += `</tr>`;
  });

  h += `
      </table>
    </body>
    </html>
  `;

  w.document.open();
  w.document.write(h);
  w.document.close();
  w.focus();
  w.print();
}

/* ===================== FINALIZAR ===================== */
async function finalizar() {
  if (!capturas.length) {
    alert("❌ No hay capturas para enviar");
    return;
  }

  if (!confirm(`📤 Enviar ${capturas.length} registros a la hoja?`)) return;

  try {
    for (const r of capturas) {
      const payload = {
        accion: "agregar",
        fecha_entrada: hoyISO(),
        fecha_salida: "",
        ubicacion: r.Ubicación || "SIN UBICACIÓN",
        codigo: String(r.Código || ""),
        descripcion: r.Descripción || "",
        cantidad: Number(r.Cantidad || 0),
        responsable: r.Operador || "",
        status: "VIGENTE",
        origen: esMobile() ? "CAPTURADOR_MOBILE" : "CAPTURADOR_WEB"
      };

      const resp = await fetch(API_GUARDAR, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      const texto = await resp.text();
      const json = obtenerJSONSeguro(texto);

      if (!resp.ok) {
        throw new Error("Error HTTP al guardar");
      }

      if (json && json.ok === false) {
        throw new Error(json.error || "El servidor rechazó el registro");
      }
    }

    alert("✅ Datos enviados correctamente a Base de Datos");
    localStorage.removeItem("capturas");
    capturas = [];
    limpiar();
    render();

  } catch (e) {
    alert("❌ Error al enviar los datos: " + (e.message || e));
  }
}

/* ===================== IMPORTADOR ===================== */
function importarMaestra() {
  const fileExcel = el("fileExcel");
  if (!fileExcel) return;

  const file = fileExcel.files?.[0];

  if (!file && bufferImportacion) {
    enviarMaestra(bufferImportacion);
    return;
  }

  if (!file) {
    alert("Selecciona Excel");
    return;
  }

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const hoja = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(hoja);

      bufferImportacion = data;
      localStorage.setItem("bufferImportacion", JSON.stringify(data));
      enviarMaestra(data);
    } catch {
      alert("❌ Error leyendo el archivo Excel");
    }
  };

  reader.readAsBinaryString(file);
}

async function enviarMaestra(data) {
  const barra = el("barra");
  const mensaje = el("mensaje");

  estadoImportacion = {
    enProceso: true,
    progreso: 0,
    mensaje: "⏳ Importando..."
  };

  localStorage.setItem("estadoImportacion", JSON.stringify(estadoImportacion));

  if (barra) barra.style.width = "0%";
  if (mensaje) mensaje.innerText = estadoImportacion.mensaje;

  let p = 0;
  const t = setInterval(() => {
    p += 10;
    if (p > 90) p = 90;

    if (barra) barra.style.width = p + "%";
    estadoImportacion.progreso = p;
    localStorage.setItem("estadoImportacion", JSON.stringify(estadoImportacion));

    if (p >= 90) clearInterval(t);
  }, 200);

  try {
    const resp = await fetch(API, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ accion: "importar", data })
    });

    const texto = await resp.text();
    const json = obtenerJSONSeguro(texto);

    clearInterval(t);

    if (!resp.ok) throw new Error("Error HTTP en importación");
    if (json && json.ok === false) throw new Error(json.error || "La importación fue rechazada");

    if (barra) barra.style.width = "100%";
    if (mensaje) mensaje.innerText = "✅ Importación exitosa";

    localStorage.removeItem("estadoImportacion");
    localStorage.removeItem("bufferImportacion");
    estadoImportacion = null;
    bufferImportacion = null;

    await cargarProductosInicial();

  } catch (e) {
    clearInterval(t);
    if (mensaje) mensaje.innerText = "❌ Error al importar";
    console.error(e);
  }
}

/* ===================== CONSULTA ===================== */
function abrirModalConsulta() {
  const modalConsulta = el("modalConsulta");
  const buscarConsulta = el("buscarConsulta");
  const resultadoConsulta = el("resultadoConsulta");
  const scrollConsulta = el("scrollConsulta");
  const msgConsulta = el("msgConsulta");

  if (!modalConsulta || !buscarConsulta || !resultadoConsulta || !scrollConsulta || !msgConsulta) return;

  modalConsulta.classList.add("show");
  buscarConsulta.value = "";
  resultadoConsulta.innerHTML = "";
  scrollConsulta.style.display = "none";
  msgConsulta.innerText = "Escriba para consultar";

  filasConsulta = [];
  indexConsulta = -1;

  buscarConsulta.focus();
}

function cerrarModalConsulta() {
  const modalConsulta = el("modalConsulta");
  if (modalConsulta) modalConsulta.classList.remove("show");
}

function filtrarConsulta() {
  clearTimeout(timerConsulta);
  timerConsulta = setTimeout(filtrarConsultaReal, 250);
}

function filtrarConsultaReal() {
  const buscarConsulta = el("buscarConsulta");
  const resultadoConsulta = el("resultadoConsulta");
  const scrollConsulta = el("scrollConsulta");
  const msgConsulta = el("msgConsulta");

  if (!buscarConsulta || !resultadoConsulta || !scrollConsulta || !msgConsulta) return;

  const q = safeLower(buscarConsulta.value);

  resultadoConsulta.innerHTML = "";
  filasConsulta = [];
  indexConsulta = -1;

  if (q.length < 2) {
    scrollConsulta.style.display = "none";
    msgConsulta.innerText = "Escriba al menos 2 caracteres";
    return;
  }

  for (const p of productos) {
    if (
      safeLower(p.CODIGO).includes(q) ||
      safeLower(p.DESCRIPCION).includes(q)
    ) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${p.CODIGO || ""}</td><td>${p.DESCRIPCION || ""}</td>`;
      tr.onclick = () => seleccionarProductoConsulta(p, filasConsulta.indexOf(tr));
      resultadoConsulta.appendChild(tr);
      filasConsulta.push(tr);

      if (filasConsulta.length >= 50) break;
    }
  }

  if (!filasConsulta.length) {
    scrollConsulta.style.display = "none";
    msgConsulta.innerText = "❌ Sin coincidencias";
    return;
  }

  scrollConsulta.style.display = "block";
  msgConsulta.innerText = `✅ ${filasConsulta.length} resultado(s)`;
  activarFilaConsulta(0);
}

function activarFilaConsulta(i) {
  const scrollConsulta = el("scrollConsulta");
  if (!scrollConsulta) return;
  if (i < 0 || i >= filasConsulta.length) return;

  filasConsulta.forEach((r) => r.classList.remove("selected"));

  const fila = filasConsulta[i];
  fila.classList.add("selected");
  indexConsulta = i;

  const filaTop = fila.offsetTop;
  const filaBottom = filaTop + fila.offsetHeight;

  const contTop = scrollConsulta.scrollTop;
  const contBottom = contTop + scrollConsulta.clientHeight;

  if (filaBottom > contBottom) {
    scrollConsulta.scrollTop = filaBottom - scrollConsulta.clientHeight + 8;
  }

  if (filaTop < contTop) {
    scrollConsulta.scrollTop = filaTop - 8;
  }
}

function seleccionarProductoConsulta(p, i) {
  const codigo = el("codigo");
  const descripcion = el("descripcion");
  const cantidad = el("cantidad");

  if (codigo) codigo.value = p.CODIGO || "";
  if (descripcion) descripcion.value = p.DESCRIPCION || "";
  if (cantidad && (!cantidad.value || Number(cantidad.value) <= 0)) cantidad.value = 1;

  activarFilaConsulta(i);
  previewIngreso();
  cerrarModalConsulta();
}

document.addEventListener("keydown", (e) => {
  const modalConsulta = el("modalConsulta");
  if (!modalConsulta || !modalConsulta.classList.contains("show")) return;
  if (!filasConsulta.length) return;

  if (["ArrowDown", "ArrowUp"].includes(e.key)) e.preventDefault();

  if (e.key === "ArrowDown") {
    activarFilaConsulta(Math.min(indexConsulta + 1, filasConsulta.length - 1));
  }

  if (e.key === "ArrowUp") {
    activarFilaConsulta(Math.max(indexConsulta - 1, 0));
  }

  if (e.key === "Enter") {
    filasConsulta[indexConsulta]?.click();
  }

  if (e.key === "Escape") {
    cerrarModalConsulta();
  }
});

/* ===================== EXPORTAR MAESTRA ===================== */
function exportarMaestraProductos() {
  if (!productos || !productos.length) {
    alert("❌ No hay productos para exportar");
    return;
  }

  const data = productos.map((p) => ({
    CODIGO: String(p.CODIGO || ""),
    DESCRIPCION: p.DESCRIPCION || ""
  }));

  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Maestra_Productos");

  const excel = XLSX.write(wb, {
    bookType: "xlsx",
    type: "array"
  });

  const blob = new Blob([excel], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "maestra_productos.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

/* ===================== PROTEGER RECARGA ===================== */
window.addEventListener("beforeunload", (e) => {
  const est = JSON.parse(localStorage.getItem("estadoImportacion") || "null");
  if (est && est.enProceso) {
    e.preventDefault();
    e.returnValue = "";
  }
});