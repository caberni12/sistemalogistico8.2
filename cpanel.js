
(() => {
  const DATA_KEY = "web_cpanel_data_v3";
  const INTEGRATION_KEY = "web_cpanel_integration_v3";
  const PRODUCT_CACHE_KEY = "web_maestra_productos_cache_v5";
  const PRODUCT_CACHE_META_KEY = "web_maestra_productos_meta_v5";
  const PRODUCT_CACHE_TTL_MS = 1000 * 60 * 60 * 24;
  const PRODUCT_RENDER_LIMIT = 250;
  const $ = (s) => document.querySelector(s);
  const $$ = (s) => [...document.querySelectorAll(s)];
  const clone = (obj) => JSON.parse(JSON.stringify(obj || {}));
  const money = (n) => Number(n || 0).toLocaleString("es-CL", { style: "currency", currency: "CLP", maximumFractionDigits: 0 });
  const safe = (v) => String(v ?? "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[m]));
  const id = (prefix) => `${prefix}-${Date.now()}-${Math.floor(Math.random() * 999)}`;

  let data = loadData();
  let productImportBuffer = [];
  let productFilters = { search: "", category: "" };
  let productRenderLimit = PRODUCT_RENDER_LIMIT;
  let productFilterTimer = null;
  let currentDocument = null;
  const VOICE_KEY = "web_cpanel_voice_orders_v1";
  const KNOWN_ORDERS_KEY = "web_cpanel_known_orders_v1";
  let voiceAlertsEnabled = localStorage.getItem(VOICE_KEY) !== "false";
  let orderWatcherTimer = null;
  let watcherFirstRemoteLoad = true;
  let knownOrderIds = new Set(loadKnownOrderIds());

  function loadData() {
    const fallback = clone(window.DEFAULT_WEB_DATA || {});
    try {
      const local = JSON.parse(localStorage.getItem(DATA_KEY) || "null");
      if (local) return normalize(local, fallback);
    } catch (err) {
      console.warn("Sin configuración local válida", err);
    }
    return fallback;
  }

  function normalize(incoming, fallback = window.DEFAULT_WEB_DATA || {}) {
    const base = clone(fallback);
    return {
      ...base,
      ...incoming,
      site: { ...(base.site || {}), ...(incoming.site || {}) },
      integracion: { ...(base.integracion || {}), ...(incoming.integracion || {}) },
      banners: Array.isArray(incoming.banners) ? incoming.banners : (base.banners || []),
      modulos: Array.isArray(incoming.modulos) ? incoming.modulos : (base.modulos || []),
      servicios: Array.isArray(incoming.servicios) ? incoming.servicios : (base.servicios || []),
      productos: Array.isArray(incoming.productos) ? incoming.productos : (base.productos || []),
      cotizaciones: Array.isArray(incoming.cotizaciones) ? incoming.cotizaciones : (base.cotizaciones || []),
      pedidos: Array.isArray(incoming.pedidos) ? incoming.pedidos : (base.pedidos || [])
    };
  }

  function saveData({ render = true, remote = true, productSource = "CPANEL_LOCAL" } = {}) {
    data.productos = normalizeProductList(data.productos || []);
    localStorage.setItem(DATA_KEY, JSON.stringify(data));
    if ((data.productos || []).length) saveProductsCache(data.productos, productSource);
    updateProductLocalStatus();
    if (render) {
      renderAll();
      reloadClientPreview();
    }
    const auto = String(data.integracion?.sincronizacionAutomatica) === "true" || data.integracion?.sincronizacionAutomatica === true;
    if (remote && auto && getRemoteUrl()) pushRemote(false);
  }

  function activeSorted(list) {
    return (list || []).slice().sort((a, b) => Number(a.orden || 0) - Number(b.orden || 0));
  }


  function normalizeProductList(list) {
    const unique = new Map();
    (Array.isArray(list) ? list : []).forEach((p, index) => {
      const codigo = cleanText(p.codigo || p.CODIGO || p.id || p.ID || "");
      const nombre = cleanText(p.nombre || p.descripcion || p.DESCRIPCION || p.detalle || codigo || "Producto");
      const key = (codigo || cleanText(p.id || p.ID || `prod-${index}`)).toUpperCase();
      if (!key) return;
      unique.set(key, {
        id: cleanText(p.id || p.ID || codigo || `prod-${index}`),
        codigo,
        nombre,
        categoria: cleanText(p.categoria || p.CATEGORIA || "General"),
        descripcion: cleanText(p.descripcion || p.DESCRIPCION || p.detalle || ""),
        precio: Number(p.precio || p.PRECIO || p.valor || 0),
        stock: Number(p.stock ?? p.cantidad ?? p.CANTIDAD ?? 0),
        imagen: cleanText(p.imagen || p.IMAGEN || p.foto || "assets/img/placeholder.svg"),
        orden: Number(p.orden || p.ORDEN || index + 1),
        activo: String(p.activo ?? p.status ?? "true").toUpperCase() !== "FALSE" && String(p.status || "ACTIVO").toUpperCase() !== "INACTIVO"
      });
    });
    return [...unique.values()].sort((a, b) => Number(a.orden || 0) - Number(b.orden || 0));
  }

  function saveProductsCache(products, source = "CPANEL") {
    const clean = normalizeProductList(products);
    if (!clean.length) return;
    try {
      localStorage.setItem(PRODUCT_CACHE_KEY, JSON.stringify(clean));
      localStorage.setItem(PRODUCT_CACHE_META_KEY, JSON.stringify({
        source,
        total: clean.length,
        syncedAt: new Date().toISOString(),
        url: getRemoteUrl()
      }));
    } catch (err) {
      console.warn("No se pudo guardar MAESTRA local", err);
    }
  }

  function reloadClientPreview() {
    const frame = document.getElementById("previewFrame");
    if (!frame) return;
    const src = (frame.getAttribute("src") || "index.html").split("?")[0] || "index.html";
    frame.setAttribute("src", src + "?preview=" + Date.now());
  }

  function getProductsCacheMeta() {
    try { return JSON.parse(localStorage.getItem(PRODUCT_CACHE_META_KEY) || "{}"); }
    catch { return {}; }
  }

  function loadProductsCache() {
    try {
      const raw = localStorage.getItem(PRODUCT_CACHE_KEY);
      if (!raw) return { ok:false, total:0 };
      const cached = normalizeProductList(JSON.parse(raw));
      if (!cached.length) return { ok:false, total:0 };
      data.productos = cached;
      localStorage.setItem(DATA_KEY, JSON.stringify(data));
      return { ok:true, total:cached.length, meta:getProductsCacheMeta() };
    } catch (err) {
      console.warn("No se pudo leer MAESTRA local", err);
      return { ok:false, total:0 };
    }
  }

  function updateProductLocalStatus(message, tone = "ok") {
    const el = $("#productLocalSyncStatus");
    if (!el) return;
    if (message) {
      el.textContent = message;
      el.dataset.tone = tone;
      return;
    }
    const meta = getProductsCacheMeta();
    const total = (data.productos || []).length;
    if (!total) {
      el.textContent = "Sin productos en memoria local. Usa Sincronizar productos.";
      el.dataset.tone = "warn";
      return;
    }
    const fecha = meta.syncedAt ? new Date(meta.syncedAt).toLocaleString("es-CL") : "sin fecha";
    el.textContent = `Productos en memoria local: ${total.toLocaleString("es-CL")} · Última sincronización: ${fecha}. Los filtros trabajan localmente.`;
    el.dataset.tone = "ok";
  }

  function isProductCacheFresh() {
    const meta = getProductsCacheMeta();
    if (!meta.syncedAt) return false;
    const d = new Date(meta.syncedAt);
    if (Number.isNaN(d.getTime())) return false;
    return Date.now() - d.getTime() < PRODUCT_CACHE_TTL_MS;
  }

  function getRemoteUrl() {
    // La URL oficial del archivo config.default.js tiene prioridad para evitar
    // que una URL antigua guardada en localStorage envíe pedidos a otro Web App.
    const direct = (window.APP_REMOTE_URL || "").trim();
    if (direct) return direct;
    try {
      const integration = JSON.parse(localStorage.getItem(INTEGRATION_KEY) || "{}");
      return (integration.appsScriptUrl || data.integracion?.appsScriptUrl || "").trim();
    } catch {
      return (data.integracion?.appsScriptUrl || "").trim();
    }
  }

  function persistOfficialRemoteUrl() {
    const url = (window.APP_REMOTE_URL || "").trim();
    if (!url) return;
    data.integracion = { ...(data.integracion || {}), appsScriptUrl: url };
    try {
      const integration = JSON.parse(localStorage.getItem(INTEGRATION_KEY) || "{}");
      integration.appsScriptUrl = url;
      localStorage.setItem(INTEGRATION_KEY, JSON.stringify(integration));
      localStorage.setItem(DATA_KEY, JSON.stringify(data));
    } catch (err) {
      console.warn("No se pudo fijar URL oficial", err);
    }
  }

  function buildPayload(action, payload = {}) {
    return { action, accion: action, ...payload, origenWeb: location.href, enviadoDesde: "cpanel" };
  }

  function createRemoteWebPayload() {
    const webData = {
      version: data.version || "3.0.0",
      site: data.site || {},
      integracion: data.integracion || {},
      banners: Array.isArray(data.banners) ? data.banners : [],
      modulos: Array.isArray(data.modulos) ? data.modulos : [],
      servicios: Array.isArray(data.servicios) ? data.servicios : []
    };
    return {
      webData,
      site: webData.site,
      banners: webData.banners,
      modulos: webData.modulos,
      servicios: webData.servicios,
      webDataJson: JSON.stringify(webData)
    };
  }

  function successMessageForAction(action) {
    const a = String(action || "").toLowerCase();
    if (a.includes("cotizacion")) return "Cotización creada correctamente";
    if (a.includes("pedido")) return "Pedido creado exitosamente";
    return "Información enviada correctamente";
  }

  function toFormBody(action, payload = {}) {
    const body = new URLSearchParams();
    body.set("data", JSON.stringify(buildPayload(action, payload)));
    body.set("action", action);
    return body;
  }

  async function api(action, payload = {}) {
    const url = getRemoteUrl();
    if (!url) throw new Error("Falta URL Apps Script.");
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
      body: toFormBody(action, payload)
    });
    const txt = await res.text();
    try {
      return JSON.parse(txt);
    } catch {
      throw new Error("Respuesta no válida: " + txt.slice(0, 240));
    }
  }

  function apiNoCors(action, payload = {}) {
    const url = getRemoteUrl();
    if (!url) return Promise.reject(new Error("Falta URL Apps Script."));
    return new Promise((resolve) => {
      const iframeName = "apps_script_sink_" + Date.now();
      const iframe = document.createElement("iframe");
      iframe.name = iframeName;
      iframe.style.display = "none";

      const form = document.createElement("form");
      form.method = "POST";
      form.action = url;
      form.target = iframeName;
      form.enctype = "application/x-www-form-urlencoded";
      form.style.display = "none";

      const input = document.createElement("input");
      input.type = "hidden";
      input.name = "data";
      input.value = JSON.stringify(buildPayload(action, payload));
      form.appendChild(input);

      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        setTimeout(() => { iframe.remove(); form.remove(); }, 500);
        resolve({ ok: true, fallback: true, msg: successMessageForAction(action), message: successMessageForAction(action) });
      };
      iframe.onload = finish;
      document.body.appendChild(iframe);
      document.body.appendChild(form);
      form.submit();
      setTimeout(finish, 4500);
    });
  }

  function apiBeacon(action, payload = {}) {
    const url = getRemoteUrl();
    if (!url || !navigator.sendBeacon) return false;
    try {
      return navigator.sendBeacon(url, toFormBody(action, payload));
    } catch {
      return false;
    }
  }

  async function apiWrite(action, payload = {}) {
    try {
      const res = await api(action, payload);
      if (res && res.ok === false) throw new Error(res.error || res.msg || "Apps Script rechazó la operación.");
      return res;
    } catch (err) {
      console.warn("Fetch directo falló; se usará respaldo de envío.", err);
      const beaconOk = apiBeacon(action, payload);
      if (beaconOk) return { ok: true, fallback: true, msg: successMessageForAction(action), message: successMessageForAction(action) };
      return apiNoCors(action, payload);
    }
  }

  function apiJsonp(action, payload = {}) {
    const url = getRemoteUrl();
    if (!url) return Promise.reject(new Error("Falta URL Apps Script."));
    return new Promise((resolve, reject) => {
      const cb = "jsonp_cb_" + Date.now() + "_" + Math.floor(Math.random() * 100000);
      const script = document.createElement("script");
      const params = new URLSearchParams({ action, accion: action, callback: cb });
      Object.entries(payload || {}).forEach(([key, value]) => {
        if (value === undefined || value === null) return;
        params.set(key, typeof value === "object" ? JSON.stringify(value) : String(value));
      });
      const cleanup = () => { delete window[cb]; script.remove(); };
      window[cb] = (res) => { cleanup(); resolve(res || {}); };
      script.onerror = () => { cleanup(); reject(new Error("No se pudo leer Apps Script por JSONP.")); };
      script.src = `${url}?${params.toString()}`;
      document.body.appendChild(script);
      setTimeout(() => {
        if (window[cb]) { cleanup(); reject(new Error("Tiempo agotado leyendo Apps Script.")); }
      }, 60000);
    });
  }

  function rowsToObjects(headers, rows) {
    headers = Array.isArray(headers) ? headers : [];
    rows = Array.isArray(rows) ? rows : [];
    return rows.map((row) => {
      const obj = {};
      headers.forEach((h, idx) => obj[String(h || "").trim()] = row[idx]);
      return obj;
    });
  }

  function normalizeRemoteProduct(p, index) {
    return {
      id: String(p.codigo || p.id || `maestra-${index}`),
      codigo: String(p.codigo || ""),
      nombre: String(p.descripcion || p.nombre || p.codigo || "Producto"),
      descripcion: String(p.descripcion || p.detalle || ""),
      categoria: String(p.categoria || "Sin categoría"),
      precio: Number(p.precio || 0),
      stock: Number(p.stock ?? p.cantidad ?? 0),
      imagen: String(p.imagen || "assets/img/placeholder.svg"),
      activo: true,
      orden: index + 1,
      unidad: String(p.unidad || "UN")
    };
  }

  function normalizePortalPedido(p, type = "pedido") {
    const numero = String(p.pedido || p.numero || p.id || "").trim();
    const items = Array.isArray(p.productos) ? p.productos.map((it, idx) => ({
      id: `${numero}-${idx}`,
      codigo: String(it.codigo || ""),
      nombre: String(it.descripcion || it.nombre || ""),
      descripcion: String(it.descripcion || it.nombre || ""),
      cantidad: Number(it.cantidad || 1),
      precio: Number(it.precio || 0),
      subtotal: Number(it.subtotal || 0)
    })) : [];
    return {
      id: numero,
      numero,
      fecha: p.fecha || "",
      cliente: p.cliente || p.nombre || "",
      nombre: p.cliente || p.nombre || "",
      vendedor: p.vendedor || "",
      pikeador: p.pikeador || "",
      telefono: p.telefono || "",
      direccion: p.direccion || "",
      estado: p.status || p.estado || "PENDIENTE",
      items,
      productos: items,
      total: Number(p.total || 0),
      total_productos: Number(p.total_productos || items.length),
      total_unidades: Number(p.total_unidades || 0),
      tipo_solicitud: type === "cotizacion" ? "COTIZACION" : "PEDIDO",
      pdfUrl: p.pdfUrl || p.pdf || ""
    };
  }

  function remoteActive(value) {
    const v = String(value ?? "").trim().toUpperCase();
    return !(v === "FALSE" || v === "NO" || v === "0" || v === "INACTIVO");
  }

  function normalizeRemoteBanner(b) {
    return {
      id: String(b.id || b.ID || "ban-" + Date.now()),
      titulo: String(b.titulo || b.TITULO || ""),
      subtitulo: String(b.subtitulo || b.SUBTITULO || ""),
      imagen: String(b.imagen || b.IMAGEN || "assets/img/banner-1.svg"),
      boton: String(b.boton || b.boton_texto || b.botonTexto || b.BOTON_TEXTO || "Ver más"),
      link: String(b.link || b.boton_link || b.botonLink || b.BOTON_LINK || "#catalogo"),
      orden: Number(b.orden || b.ORDEN || 1),
      activo: remoteActive(b.activo ?? b.ACTIVO ?? true)
    };
  }

  function normalizeRemoteModule(m) {
    return {
      id: String(m.id || m.ID || "mod-" + Date.now()),
      titulo: String(m.titulo || m.TITULO || ""),
      tipo: String(m.tipo || m.TIPO || "contenido"),
      icono: String(m.icono || m.ICONO || "🧩"),
      descripcion: String(m.descripcion || m.DESCRIPCION || ""),
      contenido: String(m.contenido || m.CONTENIDO || ""),
      boton: String(m.boton || m.boton_texto || m.botonTexto || "Ver"),
      link: String(m.link || m.boton_link || m.botonLink || "#"),
      orden: Number(m.orden || m.ORDEN || 1),
      activo: remoteActive(m.activo ?? m.ACTIVO ?? true)
    };
  }

  function normalizeRemoteService(srv) {
    return {
      id: String(srv.id || srv.ID || "srv-" + Date.now()),
      titulo: String(srv.titulo || srv.TITULO || ""),
      descripcion: String(srv.descripcion || srv.DESCRIPCION || ""),
      icono: String(srv.icono || srv.ICONO || "✨"),
      precio: String(srv.precio || srv.PRECIO || ""),
      imagen: String(srv.imagen || srv.IMAGEN || ""),
      orden: Number(srv.orden || srv.ORDEN || 1),
      activo: remoteActive(srv.activo ?? srv.ACTIVO ?? true)
    };
  }

  function mapConfigToSite(config) {
    config = config || {};
    const webJson = config.webDataJson || config.web_data_json || config.webData || config.web_data || "";
    if (webJson) {
      try { return normalize(JSON.parse(webJson), data); } catch (err) { console.warn("No se pudo leer webDataJson", err); }
    }
    return {
      ...data,
      site: {
        ...(data.site || {}),
        nombre: config.nombre_sitio || config.siteNombre || data.site?.nombre || "",
        logoTexto: config.logo_texto || config.logoTexto || data.site?.logoTexto || "",
        slogan: config.slogan || data.site?.slogan || "",
        descripcion: config.descripcion || data.site?.descripcion || "",
        telefono: config.telefono || data.site?.telefono || "",
        whatsapp: config.whatsapp || data.site?.whatsapp || "",
        correo: config.correo || data.site?.correo || "",
        direccion: config.direccion || data.site?.direccion || "",
        instagram: config.instagram || data.site?.instagram || "",
        facebook: config.facebook || data.site?.facebook || "",
        tiktok: config.tiktok || data.site?.tiktok || "",
        linkedin: config.linkedin || data.site?.linkedin || "",
        colorPrimario: config.color_primario || config.colorPrimario || data.site?.colorPrimario || "#2563eb",
        colorSecundario: config.color_secundario || config.colorSecundario || data.site?.colorSecundario || "#0f172a",
        colorAcento: config.color_acento || config.colorAcento || data.site?.colorAcento || "#f97316",
        fondo: config.fondo || data.site?.fondo || "#f8fafc"
      }
    };
  }

  async function syncWebStructureFromRemote() {
    try {
      const [cfg, banners, modulos, servicios] = await Promise.all([
        apiJsonp("get_config").catch(() => ({ ok:false })),
        apiJsonp("listar_banners").catch(() => ({ ok:false, banners: [] })),
        apiJsonp("listar_modulos").catch(() => ({ ok:false, modulos: [] })),
        apiJsonp("listar_servicios").catch(() => ({ ok:false, servicios: [] }))
      ]);
      if (cfg?.ok && cfg.config) data = mapConfigToSite(cfg.config);
      if (banners?.ok && Array.isArray(banners.banners) && banners.banners.length) data.banners = banners.banners.map(normalizeRemoteBanner);
      if (modulos?.ok && Array.isArray(modulos.modulos) && modulos.modulos.length) data.modulos = modulos.modulos.map(normalizeRemoteModule);
      if (servicios?.ok && Array.isArray(servicios.servicios) && servicios.servicios.length) data.servicios = servicios.servicios.map(normalizeRemoteService);
      localStorage.setItem(DATA_KEY, JSON.stringify(data));
      return { ok:true };
    } catch (err) {
      console.warn("No se pudo cargar estructura web desde Apps Script", err);
      return { ok:false, msg:String(err.message || err) };
    }
  }

  async function pullPortalScriptData() {
    await syncWebStructureFromRemote();
    const catalogo = await apiJsonp("portal_catalogo_cliente").catch(() => ({ ok:false }));

    if (catalogo?.ok) {
      const products = (catalogo.productos || catalogo.items || []).map(normalizeRemoteProduct).filter((x) => x.codigo || x.nombre);
      if (products.length) {
        data.productos = normalizeProductList(products);
        saveProductsCache(data.productos, "APPS_SCRIPT_MAESTRA");
      }
    }

    localStorage.setItem(DATA_KEY, JSON.stringify(data));
    return { productos: data.productos || [] };
  }

  async function syncProductsFromRemote(show = true, { force = true } = {}) {
    try {
      if (show) updateProductLocalStatus("Sincronizando productos desde Apps Script...", "warn");
      if (!force && isProductCacheFresh() && (data.productos || []).length) {
        updateProductLocalStatus();
        if (show) toast("Productos cargados desde memoria local.");
        return { ok:true, local:true, total:(data.productos || []).length };
      }
      const catalogo = await apiJsonp("portal_catalogo_cliente");
      if (!catalogo || catalogo.ok === false) throw new Error(catalogo?.msg || catalogo?.error || "Apps Script no entregó productos.");
      const products = (catalogo.productos || catalogo.items || []).map(normalizeRemoteProduct).filter((x) => x.codigo || x.nombre);
      if (!products.length) throw new Error("Apps Script respondió sin productos en MAESTRA.");
      data.productos = normalizeProductList(products);
      saveProductsCache(data.productos, "CPANEL_APPS_SCRIPT_MAESTRA");
      localStorage.setItem(DATA_KEY, JSON.stringify(data));
      productRenderLimit = PRODUCT_RENDER_LIMIT;
      renderProducts();
      enhanceResponsiveTables();
      updateProductLocalStatus(`Productos sincronizados localmente: ${data.productos.length.toLocaleString("es-CL")} registros.`, "ok");
      if (show) toast("Productos sincronizados localmente.");
      return { ok:true, total:data.productos.length };
    } catch (err) {
      console.warn("No se pudo sincronizar productos", err);
      updateProductLocalStatus(err.message || "No se pudo sincronizar productos. Se mantiene memoria local.", "err");
      if (show) toast(err.message || "No se pudo sincronizar productos.");
      return { ok:false, msg:err.message || String(err) };
    }
  }

  function toast(msg) {
    const el = $("#toast");
    if (!el) return;
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(window.__toastTimer);
    window.__toastTimer = setTimeout(() => el.classList.remove("show"), 2600);
  }
  function loadKnownOrderIds() {
    try {
      const raw = JSON.parse(localStorage.getItem(KNOWN_ORDERS_KEY) || "[]");
      return Array.isArray(raw) ? raw.map(String) : [];
    } catch {
      return [];
    }
  }

  function persistKnownOrderIds() {
    localStorage.setItem(KNOWN_ORDERS_KEY, JSON.stringify([...knownOrderIds].slice(-500)));
  }

  function getPedidoId(pedido) {
    return String(pedido?.id || pedido?.numero || "").trim();
  }

  function markKnownOrders(pedidos) {
    (pedidos || []).forEach((pedido) => {
      const pedidoId = getPedidoId(pedido);
      if (pedidoId) knownOrderIds.add(pedidoId);
    });
    persistKnownOrderIds();
  }

  function setOrderWatcherStatus(message, type = "") {
    const el = $("#orderWatcherStatus");
    if (!el) return;
    el.textContent = message;
    el.className = "status-line " + type;
  }

  function speakNewOrder(pedido) {
    if (!voiceAlertsEnabled || !("speechSynthesis" in window)) return;
    const pedidoId = getPedidoId(pedido);
    const cliente = pedido?.cliente || pedido?.nombre || "cliente sin nombre";
    const text = `Tiene un pedido nuevo. Pedido ${pedidoId}. Cliente ${cliente}.`;
    try {
      window.speechSynthesis.cancel();
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.lang = "es-CL";
      utterance.rate = 0.95;
      utterance.pitch = 1;
      window.speechSynthesis.speak(utterance);
    } catch (err) {
      console.warn("No se pudo emitir voz", err);
    }
  }

  function enableVoiceAlerts() {
    voiceAlertsEnabled = true;
    localStorage.setItem(VOICE_KEY, "true");
    setOrderWatcherStatus("Voz activada. Monitoreando pedidos nuevos cada 8 segundos.", "ok");
    toast("Voz de pedidos activada.");
    if ("speechSynthesis" in window) {
      try {
        const utterance = new SpeechSynthesisUtterance("Alertas de voz activadas.");
        utterance.lang = "es-CL";
        window.speechSynthesis.speak(utterance);
      } catch (err) {
        console.warn(err);
      }
    }
  }

  function disableVoiceAlerts() {
    voiceAlertsEnabled = false;
    localStorage.setItem(VOICE_KEY, "false");
    setOrderWatcherStatus("Voz desactivada. El monitoreo sigue cargando pedidos, pero no hablará.", "");
    toast("Voz de pedidos desactivada.");
  }

  async function checkNewOrdersFromRemote({ manual = false } = {}) {
    if (!getRemoteUrl()) {
      setOrderWatcherStatus("Falta configurar la URL de Apps Script.", "err");
      return;
    }
    try {
      const remote = await pullPortalScriptData();
      renderAll();

      const pedidos = Array.isArray(remote.pedidos) ? remote.pedidos : [];
      const nuevos = pedidos.filter((pedido) => {
        const pedidoId = getPedidoId(pedido);
        return pedidoId && !knownOrderIds.has(pedidoId);
      });

      if (watcherFirstRemoteLoad && knownOrderIds.size === 0) {
        markKnownOrders(pedidos);
        watcherFirstRemoteLoad = false;
        setOrderWatcherStatus(`Monitoreo activo. Pedidos actuales en Apps Script: ${pedidos.length}.`, "ok");
        return;
      }

      watcherFirstRemoteLoad = false;
      markKnownOrders(pedidos);

      if (nuevos.length) {
        const primero = nuevos[0];
        setOrderWatcherStatus(`Pedido nuevo detectado desde Apps Script: ${getPedidoId(primero)}.`, "ok");
        toast("Pedido nuevo detectado desde Apps Script.");
        speakNewOrder(primero);
      } else if (manual) {
        setOrderWatcherStatus("Sin pedidos nuevos. Pedidos, cotizaciones y catálogo actualizados desde Apps Script.", "ok");
        toast("Datos cargados desde Apps Script.");
      } else {
        setOrderWatcherStatus(`Monitoreo activo. Última revisión: ${new Date().toLocaleTimeString("es-CL")}.`, "ok");
      }
    } catch (err) {
      setOrderWatcherStatus(err.message || "No se pudo revisar pedidos nuevos.", "err");
    }
  }

  function startOrderWatcher() {
    if (orderWatcherTimer) clearInterval(orderWatcherTimer);
    orderWatcherTimer = null;
    setOrderWatcherStatus("Módulo de pedidos oculto: el CPanel queda dedicado a configuración web.", "");
  }


  function status(msg, type = "") {
    const el = $("#integrationStatus");
    if (!el) return;
    el.textContent = msg;
    el.className = "status-line " + type;
  }

  function bool(v) {
    return String(v) === "true" || v === true;
  }

  function applyTheme() {
    const s = data.site || {};
    document.documentElement.style.setProperty("--primary", s.colorPrimario || "#2563eb");
    document.documentElement.style.setProperty("--secondary", s.colorSecundario || "#0f172a");
    document.documentElement.style.setProperty("--accent", s.colorAcento || "#f97316");
    document.documentElement.style.setProperty("--bg", s.fondo || "#f8fafc");
  }

  function switchTab(tab) {
    $$(".admin-nav button").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
    $$(".admin-section").forEach((s) => s.classList.toggle("active", s.id === `tab-${tab}`));
    const titles = {
      dashboard: ["Dashboard", "Resumen general del sistema web."],
      general: ["Datos generales", "Administra marca, colores y contacto."],
      banners: ["Banners", "Carrusel principal de la página cliente."],
      modulos: ["Módulos dinámicos", "Crea más módulos para ampliar el sistema web."],
      servicios: ["Servicios", "Tarjetas de servicios visibles para los clientes."],
      productos: ["Productos", "Catálogo conectado a la página cliente."],
      integracion: ["Apps Script", "Conexión con Google Sheets como base de datos."],
      preview: ["Vista cliente", "Previsualización directa de index.html."]
    };
    $("#adminTitle").textContent = titles[tab]?.[0] || "CPanel";
    $("#adminSubtitle").textContent = titles[tab]?.[1] || "";
  }

  function fillSiteForm() {
    const s = data.site || {};
    $("#siteNombre").value = s.nombre || "";
    $("#siteLogoTexto").value = s.logoTexto || "";
    $("#siteSloganInput").value = s.slogan || "";
    $("#siteDescripcion").value = s.descripcion || "";
    $("#siteTelefono").value = s.telefono || "";
    $("#siteWhatsapp").value = s.whatsapp || "";
    $("#siteCorreo").value = s.correo || "";
    $("#siteDireccion").value = s.direccion || "";
    if ($("#siteInstagram")) $("#siteInstagram").value = s.instagram || s.instagramUrl || "";
    if ($("#siteFacebook")) $("#siteFacebook").value = s.facebook || s.facebookUrl || "";
    if ($("#siteTiktok")) $("#siteTiktok").value = s.tiktok || s.tiktokUrl || "";
    if ($("#siteLinkedin")) $("#siteLinkedin").value = s.linkedin || s.linkedinUrl || "";
    $("#siteBotonPrincipal").value = s.botonPrincipal || "";
    $("#siteBotonSecundario").value = s.botonSecundario || "";
    $("#siteColorPrimario").value = s.colorPrimario || "#2563eb";
    $("#siteColorSecundario").value = s.colorSecundario || "#0f172a";
    $("#siteColorAcento").value = s.colorAcento || "#f97316";
    $("#siteFondo").value = s.fondo || "#f8fafc";
    $("#generalPreviewTitle").textContent = s.nombre || "Tu Empresa";
    $("#generalPreviewDesc").textContent = s.descripcion || "";
    $("#generalPreviewColor").style.background = s.colorPrimario || "#2563eb";
    $("#generalPreviewColor").style.color = "#fff";
  }

  function readSiteForm() {
    data.site = {
      ...(data.site || {}),
      nombre: $("#siteNombre").value.trim(),
      logoTexto: $("#siteLogoTexto").value.trim(),
      slogan: $("#siteSloganInput").value.trim(),
      descripcion: $("#siteDescripcion").value.trim(),
      telefono: $("#siteTelefono").value.trim(),
      whatsapp: $("#siteWhatsapp").value.trim(),
      correo: $("#siteCorreo").value.trim(),
      direccion: $("#siteDireccion").value.trim(),
      instagram: $("#siteInstagram") ? $("#siteInstagram").value.trim() : (data.site?.instagram || ""),
      facebook: $("#siteFacebook") ? $("#siteFacebook").value.trim() : (data.site?.facebook || ""),
      tiktok: $("#siteTiktok") ? $("#siteTiktok").value.trim() : (data.site?.tiktok || ""),
      linkedin: $("#siteLinkedin") ? $("#siteLinkedin").value.trim() : (data.site?.linkedin || ""),
      botonPrincipal: $("#siteBotonPrincipal").value.trim(),
      botonSecundario: $("#siteBotonSecundario").value.trim(),
      colorPrimario: $("#siteColorPrimario").value,
      colorSecundario: $("#siteColorSecundario").value,
      colorAcento: $("#siteColorAcento").value,
      fondo: $("#siteFondo").value,
      modo: "claro"
    };
  }

  function fillIntegrationForm() {
    $("#appsScriptUrl").value = getRemoteUrl();
    $("#syncAuto").value = String(data.integracion?.sincronizacionAutomatica === true || data.integracion?.sincronizacionAutomatica === "true");
  }

  function renderStats() {
    const stats = [
      ["Banners", data.banners?.length || 0],
      ["Módulos", data.modulos?.length || 0],
      ["Servicios", data.servicios?.length || 0],
      ["Productos", data.productos?.length || 0]
    ];
    $("#statsGrid").innerHTML = stats.map(([label, value]) => `<div class="stat-card"><span>${safe(label)}</span><strong>${value}</strong></div>`).join("");
  }

  function badge(active) {
    return `<span class="badge ${bool(active) ? "on" : "off"}">${bool(active) ? "Activo" : "Inactivo"}</span>`;
  }

  function renderBanners() {
    $("#bannersTable").innerHTML = activeSorted(data.banners).map((b) => `
      <tr>
        <td>${safe(b.orden || "")}</td>
        <td><strong>${safe(b.titulo)}</strong><br><small>${safe((b.subtitulo || "").slice(0, 80))}</small></td>
        <td>${badge(b.activo)}</td>
        <td><small>${safe((b.imagen || "").slice(0, 45))}</small></td>
        <td><div class="row-actions">
          <button class="icon-btn" data-edit-banner="${safe(b.id)}">Editar</button>
          <button class="icon-btn" data-toggle-banner="${safe(b.id)}">${bool(b.activo) ? "Desactivar" : "Activar"}</button>
          <button class="icon-btn" data-delete-banner="${safe(b.id)}">Eliminar</button>
        </div></td>
      </tr>
    `).join("") || `<tr><td colspan="5">No hay banners creados.</td></tr>`;
  }

  function clearBannerForm() {
    $("#bannerForm").reset();
    $("#bannerId").value = "";
    $("#bannerOrden").value = (data.banners?.length || 0) + 1;
    $("#bannerActivo").value = "true";
    $("#bannerFormTitle").textContent = "Crear banner";
  }

  function fillBannerForm(b) {
    $("#bannerId").value = b.id || "";
    $("#bannerTitulo").value = b.titulo || "";
    $("#bannerSubtitulo").value = b.subtitulo || "";
    $("#bannerBoton").value = b.boton || "";
    $("#bannerLink").value = b.link || "";
    $("#bannerImagen").value = b.imagen || "";
    $("#bannerOrden").value = b.orden || 1;
    $("#bannerActivo").value = String(bool(b.activo));
    $("#bannerFormTitle").textContent = "Editar banner";
    switchTab("banners");
  }

  function saveBanner(e) {
    e.preventDefault();
    const currentId = $("#bannerId").value || id("ban");
    const item = {
      id: currentId,
      titulo: $("#bannerTitulo").value.trim(),
      subtitulo: $("#bannerSubtitulo").value.trim(),
      boton: $("#bannerBoton").value.trim(),
      link: $("#bannerLink").value.trim() || "#catalogo",
      imagen: $("#bannerImagen").value.trim() || "assets/img/banner-1.svg",
      orden: Number($("#bannerOrden").value || 1),
      activo: bool($("#bannerActivo").value)
    };
    const idx = data.banners.findIndex((x) => String(x.id) === String(currentId));
    if (idx >= 0) data.banners[idx] = item;
    else data.banners.push(item);
    clearBannerForm();
    saveData();
    toast("Banner guardado.");
  }

  function renderModules() {
    $("#modulesTable").innerHTML = activeSorted(data.modulos).map((m) => `
      <tr>
        <td>${safe(m.orden || "")}</td>
        <td><strong>${safe(m.icono || "🧩")} ${safe(m.titulo)}</strong><br><small>${safe((m.descripcion || "").slice(0, 90))}</small></td>
        <td><span class="badge">${safe(m.tipo || "contenido")}</span></td>
        <td>${badge(m.activo)}</td>
        <td><div class="row-actions">
          <button class="icon-btn" data-edit-module="${safe(m.id)}">Editar</button>
          <button class="icon-btn" data-toggle-module="${safe(m.id)}">${bool(m.activo) ? "Desactivar" : "Activar"}</button>
          <button class="icon-btn" data-delete-module="${safe(m.id)}">Eliminar</button>
        </div></td>
      </tr>
    `).join("") || `<tr><td colspan="5">No hay módulos creados.</td></tr>`;
  }

  function clearModuleForm() {
    $("#moduleForm").reset();
    $("#moduleId").value = "";
    $("#moduleOrden").value = (data.modulos?.length || 0) + 1;
    $("#moduleActivo").value = "true";
    $("#moduleTipo").value = "contenido";
    $("#moduleFormTitle").textContent = "Crear módulo";
  }

  function fillModuleForm(m) {
    $("#moduleId").value = m.id || "";
    $("#moduleTitulo").value = m.titulo || "";
    $("#moduleTipo").value = m.tipo || "contenido";
    $("#moduleIcono").value = m.icono || "";
    $("#moduleDescripcion").value = m.descripcion || "";
    $("#moduleContenido").value = m.contenido || "";
    $("#moduleBoton").value = m.boton || "";
    $("#moduleLink").value = m.link || "";
    $("#moduleOrden").value = m.orden || 1;
    $("#moduleActivo").value = String(bool(m.activo));
    $("#moduleFormTitle").textContent = "Editar módulo";
    switchTab("modulos");
  }

  function saveModule(e) {
    e.preventDefault();
    const currentId = $("#moduleId").value || id("mod");
    const slug = ($("#moduleTitulo").value || "modulo").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
    const item = {
      id: currentId,
      titulo: $("#moduleTitulo").value.trim(),
      tipo: $("#moduleTipo").value,
      icono: $("#moduleIcono").value.trim() || "🧩",
      descripcion: $("#moduleDescripcion").value.trim(),
      contenido: $("#moduleContenido").value.trim(),
      boton: $("#moduleBoton").value.trim() || "Ver más",
      link: $("#moduleLink").value.trim() || `#${slug}`,
      orden: Number($("#moduleOrden").value || 1),
      activo: bool($("#moduleActivo").value)
    };
    const idx = data.modulos.findIndex((x) => String(x.id) === String(currentId));
    if (idx >= 0) data.modulos[idx] = item;
    else data.modulos.push(item);
    clearModuleForm();
    saveData();
    toast("Módulo guardado.");
  }

  function renderServices() {
    $("#servicesTable").innerHTML = activeSorted(data.servicios).map((s) => `
      <tr>
        <td>${safe(s.orden || "")}</td>
        <td><strong>${safe(s.icono || "✨")} ${safe(s.titulo)}</strong><br><small>${safe((s.descripcion || "").slice(0, 90))}</small></td>
        <td>${safe(s.precio || "")}</td>
        <td>${badge(s.activo)}</td>
        <td><div class="row-actions">
          <button class="icon-btn" data-edit-service="${safe(s.id)}">Editar</button>
          <button class="icon-btn" data-toggle-service="${safe(s.id)}">${bool(s.activo) ? "Desactivar" : "Activar"}</button>
          <button class="icon-btn" data-delete-service="${safe(s.id)}">Eliminar</button>
        </div></td>
      </tr>
    `).join("") || `<tr><td colspan="5">No hay servicios creados.</td></tr>`;
  }

  function clearServiceForm() {
    $("#serviceForm").reset();
    $("#serviceId").value = "";
    $("#serviceOrden").value = (data.servicios?.length || 0) + 1;
    $("#serviceActivo").value = "true";
    $("#serviceFormTitle").textContent = "Crear servicio";
  }

  function fillServiceForm(s) {
    $("#serviceId").value = s.id || "";
    $("#serviceTitulo").value = s.titulo || "";
    $("#serviceDescripcion").value = s.descripcion || "";
    $("#serviceIcono").value = s.icono || "";
    $("#servicePrecio").value = s.precio || "";
    $("#serviceOrden").value = s.orden || 1;
    $("#serviceActivo").value = String(bool(s.activo));
    $("#serviceFormTitle").textContent = "Editar servicio";
    switchTab("servicios");
  }

  function saveService(e) {
    e.preventDefault();
    const currentId = $("#serviceId").value || id("srv");
    const item = {
      id: currentId,
      titulo: $("#serviceTitulo").value.trim(),
      descripcion: $("#serviceDescripcion").value.trim(),
      icono: $("#serviceIcono").value.trim() || "✨",
      precio: $("#servicePrecio").value.trim(),
      orden: Number($("#serviceOrden").value || 1),
      activo: bool($("#serviceActivo").value)
    };
    const idx = data.servicios.findIndex((x) => String(x.id) === String(currentId));
    if (idx >= 0) data.servicios[idx] = item;
    else data.servicios.push(item);
    clearServiceForm();
    saveData();
    toast("Servicio guardado.");
  }

  function productText(value) {
    return cleanText(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  function getProductCategories() {
    return [...new Set((data.productos || []).map((p) => cleanText(p.categoria || "General")).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, "es"));
  }

  function getVisibleProducts() {
    const search = productText(productFilters.search);
    const category = productText(productFilters.category);
    return activeSorted(data.productos).filter((p) => {
      const haystack = productText([p.codigo, p.nombre, p.categoria, p.descripcion, p.precio, p.stock].join(" "));
      const sameCategory = !category || productText(p.categoria || "General") === category;
      const matchSearch = !search || haystack.includes(search);
      return sameCategory && matchSearch;
    });
  }

  function renderProductTools(visible) {
    const all = data.productos || [];
    const categories = getProductCategories();
    const categorySelect = $("#productCategoryFilter");
    if (categorySelect) {
      const selected = productFilters.category || categorySelect.value || "";
      categorySelect.innerHTML = `<option value="">Todas las categorías</option>` + categories.map((cat) => `<option value="${safe(cat)}">${safe(cat)}</option>`).join("");
      categorySelect.value = categories.includes(selected) ? selected : "";
      productFilters.category = categorySelect.value;
    }
    const totalStock = all.reduce((sum, p) => sum + Number(p.stock || 0), 0);
    const setText = (selector, value) => { const el = $(selector); if (el) el.textContent = value; };
    setText("#productStatTotal", all.length);
    setText("#productStatActive", all.filter((p) => bool(p.activo)).length);
    setText("#productStatCategories", categories.length);
    setText("#productStatStock", totalStock.toLocaleString("es-CL"));
    setText("#productResultCount", visible.length.toLocaleString("es-CL"));
    updateProductLocalStatus();
  }

  function renderProducts() {
    const categories = getProductCategories();
    if (productFilters.category && !categories.includes(productFilters.category)) productFilters.category = "";
    const visibleAll = getVisibleProducts();
    const visible = visibleAll.slice(0, productRenderLimit);
    renderProductTools(visibleAll);
    const html = visible.map((p) => {
      const img = p.imagen || "assets/img/placeholder.svg";
      const desc = cleanText(p.descripcion || "").slice(0, 95);
      return `
        <tr>
          <td data-label="Producto">
            <div class="product-cell">
              <div class="product-thumb"><img src="${safe(img)}" alt="" loading="lazy" onerror="this.src='assets/img/placeholder.svg'"></div>
              <div class="product-cell-info">
                <strong>${safe(p.nombre || "Sin nombre")}</strong>
                <small>${safe(desc || "Sin descripción")}</small>
              </div>
            </div>
          </td>
          <td data-label="Código"><span class="code-chip">${safe(p.codigo || "")}</span></td>
          <td data-label="Categoría">${safe(p.categoria || "General")}</td>
          <td data-label="Precio"><strong>${money(p.precio)}</strong></td>
          <td data-label="Stock">${safe(p.stock ?? 0)}</td>
          <td data-label="Orden">${safe(p.orden || "")}</td>
          <td data-label="Estado">${badge(p.activo)}</td>
          <td data-label="Acciones"><div class="row-actions product-row-actions">
            <button class="icon-btn" data-edit-product="${safe(p.id)}">Editar</button>
            <button class="icon-btn" data-toggle-product="${safe(p.id)}">${bool(p.activo) ? "Desactivar" : "Activar"}</button>
            <button class="icon-btn danger-soft" data-delete-product="${safe(p.id)}">Eliminar</button>
          </div></td>
        </tr>
      `;
    }).join("");

    const moreRow = visibleAll.length > visible.length ? `
      <tr>
        <td colspan="8" data-label="Vista optimizada">
          <div class="empty-products">
            Mostrando ${visible.length.toLocaleString("es-CL")} de ${visibleAll.length.toLocaleString("es-CL")} productos para evitar lentitud.
            <button class="icon-btn" type="button" id="btnMoreProductRows">Ver más</button>
          </div>
        </td>
      </tr>` : "";

    $("#productsTable").innerHTML = html || `<tr><td colspan="8" data-label="Productos"><div class="empty-products">No hay productos para mostrar con el filtro actual.</div></td></tr>`;
    if (moreRow) $("#productsTable").insertAdjacentHTML("beforeend", moreRow);
    $("#btnMoreProductRows")?.addEventListener("click", () => {
      productRenderLimit += PRODUCT_RENDER_LIMIT;
      renderProducts();
      enhanceResponsiveTables();
    });
  }

  function clearProductForm() {
    $("#productForm").reset();
    $("#productId").value = "";
    $("#productOrden").value = (data.productos?.length || 0) + 1;
    $("#productActivo").value = "true";
    $("#productFormTitle").textContent = "Crear producto";
  }

  function openProductModal() {
    const modal = $("#productModal");
    if (!modal) return;
    modal.classList.add("show");
    modal.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
    setTimeout(() => $("#productCodigo")?.focus(), 60);
  }

  function closeProductModal() {
    const modal = $("#productModal");
    if (!modal) return;
    modal.classList.remove("show");
    modal.setAttribute("aria-hidden", "true");
    document.body.style.overflow = "";
  }

  function fillProductForm(p) {
    $("#productId").value = p.id || "";
    $("#productCodigo").value = p.codigo || "";
    $("#productNombre").value = p.nombre || "";
    $("#productCategoria").value = p.categoria || "";
    $("#productDescripcion").value = p.descripcion || "";
    $("#productPrecio").value = p.precio || 0;
    $("#productStock").value = p.stock || 0;
    $("#productImagen").value = p.imagen || "";
    $("#productOrden").value = p.orden || 1;
    $("#productActivo").value = String(bool(p.activo));
    $("#productFormTitle").textContent = "Editar producto";
    switchTab("productos");
    openProductModal();
  }

  function saveProduct(e) {
    e.preventDefault();
    const currentId = $("#productId").value || id("prod");
    const item = {
      id: currentId,
      codigo: $("#productCodigo").value.trim(),
      nombre: $("#productNombre").value.trim(),
      categoria: $("#productCategoria").value.trim() || "General",
      descripcion: $("#productDescripcion").value.trim(),
      precio: Number($("#productPrecio").value || 0),
      stock: Number($("#productStock").value || 0),
      imagen: $("#productImagen").value.trim() || "assets/img/placeholder.svg",
      orden: Number($("#productOrden").value || 1),
      activo: bool($("#productActivo").value)
    };
    const idx = data.productos.findIndex((x) => String(x.id) === String(currentId));
    if (idx >= 0) data.productos[idx] = item;
    else data.productos.push(item);
    data.productos = normalizeProductList(data.productos);
    saveProductsCache(data.productos, "CPANEL_PRODUCTO_GUARDADO");
    updateProductLocalStatus();
    clearProductForm();
    closeProductModal();
    saveData({ productSource: "CPANEL_PRODUCTO_GUARDADO" });
    toast("Producto guardado y actualizado en memoria local.");
  }


  function normalizeHeader(value) {
    return String(value || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }

  function getRowValue(row, names) {
    const wanted = names.map(normalizeHeader);
    for (const key of Object.keys(row || {})) {
      if (wanted.includes(normalizeHeader(key))) return row[key];
    }
    return "";
  }

  function cleanText(value) {
    return String(value ?? "").trim();
  }

  function cleanNumber(value) {
    const txt = cleanText(value).replace(/\./g, "").replace(/,/g, ".").replace(/[^0-9.-]/g, "");
    const n = Number(txt || 0);
    return Number.isFinite(n) ? n : 0;
  }

  function parseActive(value) {
    if (value === "" || value === null || value === undefined) return true;
    const txt = cleanText(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    return ["si", "s", "true", "1", "activo", "activa", "yes", "y"].includes(txt);
  }

  function normalizeProductImportRow(row, index) {
    const codigo = cleanText(getRowValue(row, ["CODIGO", "CÓDIGO", "SKU", "COD", "CODE"]));
    const nombre = cleanText(getRowValue(row, ["NOMBRE", "PRODUCTO", "TITULO", "TÍTULO", "DESCRIPCION CORTA"]));
    const categoria = cleanText(getRowValue(row, ["CATEGORIA", "CATEGORÍA", "FAMILIA", "LINEA", "LÍNEA"]));
    const descripcion = cleanText(getRowValue(row, ["DESCRIPCION", "DESCRIPCIÓN", "DETALLE", "CONTENIDO"]));
    const precio = cleanNumber(getRowValue(row, ["PRECIO", "VALOR", "MONTO", "PRECIO VENTA"]));
    const stock = cleanNumber(getRowValue(row, ["STOCK", "CANTIDAD", "EXISTENCIA", "UNIDADES"]));
    const imagen = cleanText(getRowValue(row, ["IMAGEN", "URL IMAGEN", "FOTO", "URL", "IMAGE"]));
    const activoRaw = getRowValue(row, ["ACTIVO", "ESTADO", "VISIBLE", "PUBLICADO"]);
    const ordenRaw = getRowValue(row, ["ORDEN", "POSICION", "POSICIÓN", "N"]);
    const idRaw = cleanText(getRowValue(row, ["ID", "IDENTIFICADOR"]));
    const errors = [];
    if (!codigo) errors.push("Falta CODIGO");
    if (!nombre) errors.push("Falta NOMBRE");
    const existing = data.productos.find((p) => cleanText(p.codigo).toLowerCase() === codigo.toLowerCase());
    return {
      id: idRaw || existing?.id || id("prod"),
      codigo,
      nombre,
      categoria: categoria || "General",
      descripcion,
      precio,
      stock,
      imagen: imagen || "assets/img/placeholder.svg",
      activo: parseActive(activoRaw),
      orden: ordenRaw === "" || ordenRaw === null || ordenRaw === undefined ? (existing?.orden || ((data.productos?.length || 0) + index + 1)) : cleanNumber(ordenRaw),
      __action: existing ? "update" : "new",
      __error: errors.join(", ")
    };
  }

  function openProductImportPanel() {
    const panel = $("#productImportPanel");
    if (!panel) return;
    panel.hidden = false;
    panel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function closeProductImportPanel() {
    const panel = $("#productImportPanel");
    if (!panel) return;
    panel.hidden = true;
    clearProductImportPreview();
  }

  function clearProductImportPreview() {
    productImportBuffer = [];
    const wrap = $("#productImportPreviewWrap");
    if (wrap) wrap.hidden = true;
    const tbody = $("#productImportPreview");
    if (tbody) tbody.innerHTML = "";
    const summary = $("#productImportSummary");
    if (summary) summary.textContent = "";
  }

  function renderProductImportPreview(items) {
    const valid = items.filter((p) => !p.__error);
    const nuevos = valid.filter((p) => p.__action === "new").length;
    const actualizados = valid.filter((p) => p.__action === "update").length;
    const errores = items.length - valid.length;
    $("#productImportSummary").textContent = `${items.length} filas leídas · ${nuevos} nuevos · ${actualizados} actualizaciones · ${errores} con error.`;
    $("#productImportPreview").innerHTML = items.slice(0, 200).map((p) => `
      <tr>
        <td><span class="import-chip ${p.__error ? "error" : p.__action}">${p.__error ? "Error" : (p.__action === "update" ? "Actualizar" : "Nuevo")}</span></td>
        <td>${safe(p.codigo)}</td>
        <td><strong>${safe(p.nombre)}</strong><br><small>${safe((p.descripcion || "").slice(0, 70))}</small></td>
        <td>${safe(p.categoria)}</td>
        <td>${money(p.precio)}</td>
        <td>${safe(p.stock)}</td>
        <td>${badge(p.activo)}</td>
        <td>${safe(p.__error || "Listo para importar")}</td>
      </tr>
    `).join("") || `<tr><td colspan="8">No hay filas para importar.</td></tr>`;
    if (items.length > 200) {
      $("#productImportPreview").insertAdjacentHTML("beforeend", `<tr><td colspan="8">Vista previa limitada a 200 filas. Al aplicar se importan todas las filas válidas.</td></tr>`);
    }
    $("#productImportPreviewWrap").hidden = false;
    enhanceResponsiveTables();
  }

  function parseCsv(text) {
    const rows = [];
    let row = [], current = "", inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const char = text[i], next = text[i + 1];
      if (char === '"' && inQuotes && next === '"') { current += '"'; i++; continue; }
      if (char === '"') { inQuotes = !inQuotes; continue; }
      if (char === ',' && !inQuotes) { row.push(current); current = ""; continue; }
      if ((char === '\n' || char === '\r') && !inQuotes) {
        if (char === '\r' && next === '\n') i++;
        row.push(current); rows.push(row); row = []; current = ""; continue;
      }
      current += char;
    }
    if (current || row.length) { row.push(current); rows.push(row); }
    const headers = (rows.shift() || []).map((h) => cleanText(h));
    return rows.filter((r) => r.some((c) => cleanText(c))).map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ""])));
  }

  async function readProductImportFile(file) {
    const name = file.name.toLowerCase();
    let rows = [];
    if (name.endsWith(".csv")) {
      rows = parseCsv(await file.text());
    } else {
      if (!window.XLSX) throw new Error("No se pudo cargar la librería Excel. Usa CSV o revisa conexión a internet para XLSX.");
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellText: true, cellDates: false, raw: false });
      const sheetName = workbook.SheetNames[0];
      rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", raw: false });
    }
    const mapped = rows.map((row, index) => normalizeProductImportRow(row, index + 1));
    productImportBuffer = mapped;
    renderProductImportPreview(mapped);
  }

  function applyProductImport() {
    const valid = productImportBuffer.filter((p) => !p.__error).map((p) => {
      const copy = { ...p };
      delete copy.__action;
      delete copy.__error;
      return copy;
    });
    if (!valid.length) {
      toast("No hay productos válidos para importar.");
      return;
    }
    valid.forEach((item) => {
      const idx = data.productos.findIndex((p) => cleanText(p.codigo).toLowerCase() === item.codigo.toLowerCase());
      if (idx >= 0) data.productos[idx] = { ...data.productos[idx], ...item, id: data.productos[idx].id || item.id };
      else data.productos.push(item);
    });
    clearProductImportPreview();
    saveData();
    data.productos = normalizeProductList(data.productos);
    saveProductsCache(data.productos, "CPANEL_IMPORTACION_PRODUCTOS");
    updateProductLocalStatus();
    toast(`${valid.length} productos importados/actualizados en memoria local.`);
  }

  function downloadProductTemplate() {
    const rows = [
      ["CODIGO", "NOMBRE", "CATEGORIA", "DESCRIPCION", "PRECIO", "STOCK", "IMAGEN", "ACTIVO", "ORDEN"],
      ["000001", "", "", "", "", "", "", "SI", "1"]
    ];
    if (window.XLSX) {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = [{ wch: 18 }, { wch: 28 }, { wch: 20 }, { wch: 38 }, { wch: 14 }, { wch: 12 }, { wch: 36 }, { wch: 12 }, { wch: 10 }];
      XLSX.utils.book_append_sheet(wb, ws, "PRODUCTOS_IMPORTAR");
      XLSX.writeFile(wb, "plantilla_importacion_productos.xlsx");
      return;
    }
    const csv = rows.map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "plantilla_importacion_productos.csv";
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function itemCount(doc) {
    return Array.isArray(doc.items) ? doc.items.length : Number(doc.items || 0);
  }

  function renderQuotes() {
    if (!$("#quotesTable")) return;
    $("#quotesTable").innerHTML = (data.cotizaciones || []).map((q) => {
      const docId = safe(q.id || q.numero || "");
      return `
      <tr>
        <td>${safe(formatDate(q.fecha))}</td>
        <td><strong>${docId}</strong></td>
        <td>${safe(q.nombre || q.cliente || "")}</td>
        <td>${safe(q.telefono || "")}</td>
        <td>${itemCount(q)}</td>
        <td>${money(q.total || 0)}</td>
        <td><span class="badge on">${safe(q.estado || "Nueva")}</span></td>
        <td>
          <div class="doc-row-actions">
            <button class="icon-btn" data-view-doc="cotizacion" data-doc-id="${docId}">Ver</button>
            <button class="icon-btn" data-pdf-doc="cotizacion" data-doc-id="${docId}" data-pdf-format="a4">PDF A4</button>
            <button class="icon-btn" data-pdf-doc="cotizacion" data-doc-id="${docId}" data-pdf-format="ticket">Ticket 80</button>
          </div>
        </td>
      </tr>`;
    }).join("") || `<tr><td colspan="8">No hay cotizaciones recibidas.</td></tr>`;
  }


  function renderOrders() {
    const tbody = $("#ordersTable");
    if (!tbody) return;
    if (!tbody) return;
    tbody.innerHTML = (data.pedidos || []).map((p) => {
      const docId = safe(p.id || p.numero || "");
      return `
      <tr>
        <td>${safe(formatDate(p.fecha))}</td>
        <td><strong>${docId}</strong></td>
        <td>${safe(p.cliente || p.nombre || "")}</td>
        <td>${safe(p.telefono || "")}</td>
        <td>${safe(p.direccion || "")}</td>
        <td>${itemCount(p)}</td>
        <td>${money(p.total || 0)}</td>
        <td><span class="badge on">${safe(p.estado || "Nuevo")}</span></td>
        <td>
          <div class="doc-row-actions">
            <button class="icon-btn" data-view-doc="pedido" data-doc-id="${docId}">Ver</button>
            <button class="icon-btn" data-pdf-doc="pedido" data-doc-id="${docId}" data-pdf-format="a4">PDF A4</button>
            <button class="icon-btn" data-pdf-doc="pedido" data-doc-id="${docId}" data-pdf-format="ticket">Ticket 80</button>
          </div>
        </td>
      </tr>`;
    }).join("") || `<tr><td colspan="9">No hay pedidos recibidos.</td></tr>`;
  }

  function findDocument(kind, docId) {
    const list = kind === "pedido" ? data.pedidos : data.cotizaciones;
    return (list || []).find((doc) => String(doc.id || doc.numero || "") === String(docId));
  }

  function docClientName(kind, doc) {
    return kind === "pedido" ? (doc.cliente || doc.nombre || "") : (doc.nombre || doc.cliente || "");
  }

  function normalizeDocItems(doc) {
    let items = doc?.items || [];
    if (typeof items === "string") {
      try { items = JSON.parse(items); } catch { items = []; }
    }
    return Array.isArray(items) ? items : [];
  }

  function openDocumentModal(kind, docId) {
    const doc = findDocument(kind, docId);
    if (!doc) {
      toast("No se encontró el documento solicitado.");
      return;
    }
    currentDocument = { kind, doc };
    const title = kind === "pedido" ? "Pedido solicitado" : "Cotización solicitada";
    const items = normalizeDocItems(doc);
    $("#documentModalTitle").textContent = `${title} ${doc.id || doc.numero || ""}`;
    $("#documentModalSubtitle").textContent = "Detalle completo recibido desde la página cliente.";
    $("#documentModalBody").innerHTML = `
      <div class="document-detail-grid">
        <div class="document-detail-box"><span>Número</span><strong>${safe(doc.id || doc.numero || "")}</strong></div>
        <div class="document-detail-box"><span>Fecha</span><strong>${safe(formatDate(doc.fecha))}</strong></div>
        <div class="document-detail-box"><span>Cliente</span><strong>${safe(docClientName(kind, doc))}</strong></div>
        <div class="document-detail-box"><span>Estado</span><strong>${safe(doc.estado || (kind === "pedido" ? "Nuevo" : "Nueva"))}</strong></div>
        <div class="document-detail-box"><span>RUT</span><strong>${safe(doc.rut || "")}</strong></div>
        <div class="document-detail-box"><span>Teléfono</span><strong>${safe(doc.telefono || "")}</strong></div>
        <div class="document-detail-box"><span>Correo</span><strong>${safe(doc.correo || "")}</strong></div>
        <div class="document-detail-box"><span>Dirección</span><strong>${safe(doc.direccion || "")}</strong></div>
      </div>
      <div class="document-detail-box" style="margin-bottom:16px"><span>Mensaje / observación</span><strong>${safe(doc.mensaje || "")}</strong></div>
      <div class="document-items-card">
        <h3>Productos solicitados</h3>
        <div class="table-wrap">
          <table class="document-items-table">
            <thead><tr><th>Código</th><th>Producto</th><th>Precio</th><th>Cantidad</th><th>Subtotal</th></tr></thead>
            <tbody>
              ${items.map((item) => `
                <tr>
                  <td>${safe(item.codigo || "")}</td>
                  <td>${safe(item.nombre || item.producto || "")}</td>
                  <td>${money(item.precio || 0)}</td>
                  <td>${safe(item.cantidad || 0)}</td>
                  <td>${money(item.subtotal || Number(item.precio || 0) * Number(item.cantidad || 0))}</td>
                </tr>`).join("") || `<tr><td colspan="5">No hay productos en el detalle.</td></tr>`}
            </tbody>
          </table>
        </div>
        <div class="document-total-row"><strong>Total:</strong><strong>${money(doc.total || 0)}</strong></div>
      </div>
    `;
    const modal = $("#documentModal");
    modal?.classList.add("show");
    modal?.setAttribute("aria-hidden", "false");
  }

  function closeDocumentModal() {
    const modal = $("#documentModal");
    if (!modal) return;
    modal.classList.remove("show");
    modal.setAttribute("aria-hidden", "true");
  }

  async function generateDocumentPdf(format = "a4", source) {
    const payload = source || currentDocument;
    if (!payload?.doc) {
      toast("Primero abre un pedido o cotización.");
      return;
    }

    const { kind, doc } = payload;
    const docId = doc.id || doc.numero || "";
    const remoteFormat = format === "ticket" ? "ticket" : "a4";

    if (getRemoteUrl() && docId) {
      try {
        toast("Generando PDF en Apps Script...");
        const action = kind === "pedido" ? "generarPdfPedido" : "generarPdfCotizacion";
        const res = await api(action, { id: docId, formato: remoteFormat });
        if (res && res.ok && res.pdf && res.pdf.url) {
          const urlKey = remoteFormat === "ticket" ? "pdf80Url" : "pdfA4Url";
          doc[urlKey] = res.pdf.url;
          saveLocalData();
          renderQuotes();
          renderOrders();
          window.open(res.pdf.url, "_blank");
          toast(`PDF ${remoteFormat === "ticket" ? "ticket 80 mm" : "A4"} generado y guardado en Apps Script.`);
          return;
        }
        if (res && res.ok === false) throw new Error(res.error || "Apps Script rechazó la generación del PDF.");
      } catch (err) {
        console.warn("No se pudo generar PDF en Apps Script. Se usará PDF local.", err);
        toast("No se pudo generar en Apps Script; se generará PDF local.");
      }
    }

    if (!window.jspdf?.jsPDF) {
      toast("No se pudo cargar jsPDF. Revisa conexión a internet.");
      return;
    }

    const { jsPDF } = window.jspdf;
    const items = normalizeDocItems(doc);
    const isTicket = format === "ticket";
    const pageHeight = isTicket ? Math.max(160, 92 + (items.length * 13)) : 297;
    const pdf = new jsPDF({ unit: "mm", format: isTicket ? [80, pageHeight] : "a4", orientation: "portrait" });
    const nombreDoc = kind === "pedido" ? "Pedido" : "Cotización";
    const pageWidth = pdf.internal.pageSize.getWidth();
    let y = isTicket ? 8 : 16;

    const line = () => { pdf.line(isTicket ? 4 : 12, y, pageWidth - (isTicket ? 4 : 12), y); y += isTicket ? 5 : 7; };
    const text = (txt, x, yy, opts = {}) => pdf.text(String(txt ?? ""), x, yy, opts);
    const addWrapped = (txt, x, maxWidth, size = 9) => {
      pdf.setFontSize(size);
      const lines = pdf.splitTextToSize(String(txt ?? ""), maxWidth);
      lines.forEach((ln) => { text(ln, x, y); y += isTicket ? 4 : 5; });
    };

    if (isTicket) {
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(12);
      text(data.site?.nombre || "Mi Empresa", pageWidth / 2, y, { align: "center" }); y += 5;
      pdf.setFontSize(9); pdf.setFont("helvetica", "normal");
      text(data.site?.telefono || "", pageWidth / 2, y, { align: "center" }); y += 4;
      text(data.site?.correo || "", pageWidth / 2, y, { align: "center" }); y += 5;
      line();
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(10);
      text(`${nombreDoc}: ${doc.id || doc.numero || ""}`, 4, y); y += 5;
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(8);
      addWrapped(`Fecha: ${formatDate(doc.fecha)}`, 4, 72, 8);
      addWrapped(`Cliente: ${docClientName(kind, doc)}`, 4, 72, 8);
      addWrapped(`Teléfono: ${doc.telefono || ""}`, 4, 72, 8);
      addWrapped(`Dirección: ${doc.direccion || ""}`, 4, 72, 8);
      line();
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(8);
      text("Producto", 4, y); text("Cant", 52, y); text("Total", 66, y, { align: "right" }); y += 4;
      pdf.setFont("helvetica", "normal");
      items.forEach((item) => {
        const subtotal = Number(item.subtotal || Number(item.precio || 0) * Number(item.cantidad || 0));
        addWrapped(`${item.codigo ? item.codigo + " - " : ""}${item.nombre || item.producto || ""}`, 4, 46, 7.5);
        text(String(item.cantidad || 0), 54, y - 4);
        text(money(subtotal).replace("CLP", "").trim(), 76, y - 4, { align: "right" });
      });
      line();
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(11);
      text("TOTAL", 4, y); text(money(doc.total || 0).replace("CLP", "").trim(), 76, y, { align: "right" }); y += 7;
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(8);
      text("Documento generado desde CPanel", pageWidth / 2, y, { align: "center" });
    } else {
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(18);
      text(data.site?.nombre || "Mi Empresa", 12, y); y += 8;
      pdf.setFontSize(12); text(`${nombreDoc} ${doc.id || doc.numero || ""}`, 12, y); y += 8;
      line();
      pdf.setFont("helvetica", "normal"); pdf.setFontSize(10);
      const fields = [
        ["Fecha", formatDate(doc.fecha)], ["Cliente", docClientName(kind, doc)], ["RUT", doc.rut || ""], ["Teléfono", doc.telefono || ""],
        ["Correo", doc.correo || ""], ["Dirección", doc.direccion || ""], ["Estado", doc.estado || ""], ["Observación", doc.mensaje || ""]
      ];
      fields.forEach(([label, value], idx) => {
        const x = idx % 2 === 0 ? 12 : 108;
        if (idx > 0 && idx % 2 === 0) y += 8;
        pdf.setFont("helvetica", "bold"); text(label + ":", x, y);
        pdf.setFont("helvetica", "normal");
        const lines = pdf.splitTextToSize(String(value || ""), 70);
        text(lines[0] || "", x + 24, y);
      });
      y += 14;
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(9);
      pdf.rect(12, y - 5, 186, 8);
      text("Código", 14, y); text("Producto", 42, y); text("Precio", 124, y); text("Cant.", 152, y); text("Subtotal", 173, y, { align: "right" }); y += 6;
      pdf.setFont("helvetica", "normal");
      items.forEach((item) => {
        if (y > 274) { pdf.addPage(); y = 16; }
        const subtotal = Number(item.subtotal || Number(item.precio || 0) * Number(item.cantidad || 0));
        const productLines = pdf.splitTextToSize(String(item.nombre || item.producto || ""), 74);
        const rowHeight = Math.max(8, productLines.length * 5 + 3);
        pdf.rect(12, y - 4, 186, rowHeight);
        text(String(item.codigo || ""), 14, y);
        text(productLines, 42, y);
        text(money(item.precio || 0).replace("CLP", "").trim(), 144, y, { align: "right" });
        text(String(item.cantidad || 0), 158, y);
        text(money(subtotal).replace("CLP", "").trim(), 194, y, { align: "right" });
        y += rowHeight;
      });
      y += 8;
      if (y > 270) { pdf.addPage(); y = 16; }
      pdf.setFont("helvetica", "bold"); pdf.setFontSize(13);
      text("TOTAL", 152, y); text(money(doc.total || 0), 198, y, { align: "right" });
    }

    const fileSafe = String(doc.id || doc.numero || Date.now()).replace(/[^a-z0-9_-]+/gi, "_");
    pdf.save(`${nombreDoc.toLowerCase()}_${fileSafe}_${isTicket ? "ticket_80mm" : "a4"}.pdf`);
    toast(`PDF ${isTicket ? "ticket 80 mm" : "A4"} generado.`);
  }

  function formatDate(value) {
    if (value === null || value === undefined || value === "") return "";

    const pad = (n) => String(n).padStart(2, "0");
    const format = (d) => {
      if (!(d instanceof Date) || Number.isNaN(d.getTime())) return "";
      return `${pad(d.getDate())}-${pad(d.getMonth() + 1)}-${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    };

    if (value instanceof Date) return format(value) || "";

    let raw = String(value).trim();
    if (!raw || raw.toLowerCase() === "invalid date") return "";
    if (raw.startsWith("'")) raw = raw.slice(1).trim();

    // Google Sheets puede entregar fechas como número serial de Excel/Sheets.
    const numeric = Number(raw.replace(",", "."));
    if (Number.isFinite(numeric)) {
      if (numeric > 20000 && numeric < 90000) {
        const ms = Date.UTC(1899, 11, 30) + Math.round(numeric * 86400000);
        return format(new Date(ms));
      }
      if (numeric > 100000000000) {
        const byTimestamp = format(new Date(numeric));
        if (byTimestamp) return byTimestamp;
      }
    }

    // Formato ISO o similar: 2026-05-24 / 2026-05-24 13:45 / 2026-05-24T13:45:00Z
    let m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) {
      const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
      const out = format(d);
      if (out) return out;
    }

    // Formato chileno/frecuente: 24-05-2026 13:45 o 24/05/2026 13:45
    m = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:[T\s,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) {
      const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
      const out = format(d);
      if (out) return out;
    }

    const parsed = new Date(raw);
    const out = format(parsed);
    return out || raw;
  }

  function bindNav() {
    $$("#adminNav button").forEach((btn) => btn.addEventListener("click", () => switchTab(btn.dataset.tab)));
  }

  function bindForms() {
    $("#siteForm").addEventListener("submit", (e) => {
      e.preventDefault();
      readSiteForm();
      saveData();
      toast("Datos generales guardados.");
    });

    $("#btnResetLocal").addEventListener("click", () => {
      if (!confirm("¿Restablecer los datos demo locales?")) return;
      localStorage.removeItem(DATA_KEY);
      data = loadData();
      renderAll();
      toast("Datos demo restablecidos.");
    });

    $("#bannerForm").addEventListener("submit", saveBanner);
    $("#btnClearBanner").addEventListener("click", clearBannerForm);

    $("#moduleForm").addEventListener("submit", saveModule);
    $("#btnClearModule").addEventListener("click", clearModuleForm);

    $("#serviceForm").addEventListener("submit", saveService);
    $("#btnClearService").addEventListener("click", clearServiceForm);

    $("#productForm").addEventListener("submit", saveProduct);
    $("#btnClearProduct").addEventListener("click", clearProductForm);
    $("#btnOpenProductModal")?.addEventListener("click", () => {
      clearProductForm();
      openProductModal();
    });
    $("#btnCloseProductModal")?.addEventListener("click", closeProductModal);
    $("#btnCancelProduct")?.addEventListener("click", closeProductModal);
    $("#btnClearProductListView")?.addEventListener("click", () => {
      productFilters = { search: "", category: "" };
      const search = $("#productSearch");
      const category = $("#productCategoryFilter");
      if (search) search.value = "";
      if (category) category.value = "";
      productRenderLimit = PRODUCT_RENDER_LIMIT;
      renderProducts();
      enhanceResponsiveTables();
      updateProductLocalStatus();
      toast("Vista de productos actualizada desde memoria local.");
    });
    $("#btnSyncProductsRemote")?.addEventListener("click", () => syncProductsFromRemote(true, { force:true }));
    $("#productSearch")?.addEventListener("input", (e) => {
      productFilters.search = e.target.value || "";
      productRenderLimit = PRODUCT_RENDER_LIMIT;
      clearTimeout(productFilterTimer);
      productFilterTimer = setTimeout(() => {
        renderProducts();
        enhanceResponsiveTables();
      }, 120);
    });
    $("#productCategoryFilter")?.addEventListener("change", (e) => {
      productFilters.category = e.target.value || "";
      productRenderLimit = PRODUCT_RENDER_LIMIT;
      renderProducts();
      enhanceResponsiveTables();
    });
    $("#btnOpenImportProducts")?.addEventListener("click", openProductImportPanel);
    $("#btnCloseImportProducts")?.addEventListener("click", closeProductImportPanel);
    $("#btnCancelProductImport")?.addEventListener("click", clearProductImportPreview);
    $("#btnApplyProductImport")?.addEventListener("click", applyProductImport);
    $("#btnDownloadProductTemplate")?.addEventListener("click", downloadProductTemplate);
    $("#productImportFile")?.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await readProductImportFile(file);
        toast("Archivo leído. Revisa la vista previa y aplica la importación.");
      } catch (err) {
        toast(err.message || "No se pudo leer el archivo.");
      }
      e.target.value = "";
    });
    document.querySelectorAll("[data-close-product-modal]").forEach((el) => el.addEventListener("click", closeProductModal));
    document.querySelectorAll("[data-close-document-modal]").forEach((el) => el.addEventListener("click", closeDocumentModal));
    $("#btnCloseDocumentModal")?.addEventListener("click", closeDocumentModal);
    $("#btnCloseDocumentModalFooter")?.addEventListener("click", closeDocumentModal);
    $("#btnDocPdfA4")?.addEventListener("click", () => generateDocumentPdf("a4"));
    $("#btnDocPdfTicket")?.addEventListener("click", () => generateDocumentPdf("ticket"));
    document.addEventListener("keydown", (ev) => {
      if (ev.key === "Escape") {
        closeProductModal();
        closeDocumentModal();
      }
    });

    document.getElementById("btnPublishWeb")?.addEventListener("click", async () => {
      readSiteForm();
      saveData({ remote: false });
      await pushRemote(true);
      toast("Web guardada y publicada.");
    });

    $("#integrationForm").addEventListener("submit", (e) => {
      e.preventDefault();
      data.integracion = {
        ...(data.integracion || {}),
        appsScriptUrl: $("#appsScriptUrl").value.trim(),
        sincronizacionAutomatica: bool($("#syncAuto").value)
      };
      localStorage.setItem(INTEGRATION_KEY, JSON.stringify(data.integracion));
      saveData({ remote: false });
      status("URL guardada correctamente.", "ok");
      toast("Integración guardada.");
    });

    $("#btnPing").addEventListener("click", async () => {
      try {
        status("Probando conexión...");
        const res = await api("ping");
        if (!res.ok) throw new Error(res.error || "Sin respuesta OK");
        status("Conexión correcta con Apps Script.", "ok");
      } catch (err) {
        status(err.message, "err");
      }
    });

    $("#btnSetup").addEventListener("click", async () => {
      try {
        status("Creando hojas de base de datos...");
        const res = await api("setup");
        if (!res.ok) throw new Error(res.error || "No se pudo crear BD");
        status("Hojas creadas/verificadas correctamente.", "ok");
      } catch (err) {
        status(err.message, "err");
      }
    });

    $("#btnPushRemote").addEventListener("click", () => pushRemote(true));
    $("#btnPullRemote").addEventListener("click", () => pullRemote(true));
    $("#btnLoadQuotes")?.addEventListener("click", () => pullRemote(true));

    $("#btnClearQuotes").addEventListener("click", () => {
      if (!confirm("¿Limpiar cotizaciones locales?")) return;
      data.cotizaciones = [];
      saveData({ remote: false });
      toast("Cotizaciones locales eliminadas.");
    });

    $("#btnLoadOrders")?.addEventListener("click", () => checkNewOrdersFromRemote({ manual: true }));
    $("#btnEnableOrderVoice")?.addEventListener("click", enableVoiceAlerts);
    $("#btnDisableOrderVoice")?.addEventListener("click", disableVoiceAlerts);

    $("#btnClearOrders")?.addEventListener("click", () => {
      if (!confirm("¿Limpiar pedidos locales?")) return;
      data.pedidos = [];
      saveData({ remote: false });
      toast("Pedidos locales eliminados.");
    });

    $("#btnReloadPreview").addEventListener("click", () => {
      $("#previewFrame").src = "index.html?t=" + Date.now();
    });

    $("#btnExport").addEventListener("click", () => {
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "configuracion_web_cpanel.json";
      a.click();
      URL.revokeObjectURL(a.href);
    });

    $("#importFile").addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        data = normalize(parsed, window.DEFAULT_WEB_DATA || {});
        saveData({ remote: false });
        toast("Configuración importada.");
      } catch (err) {
        toast("Archivo JSON no válido.");
      }
      e.target.value = "";
    });
  }

  async function pushRemote(show = true) {
    try {
      if (show) status("Enviando información a Apps Script...");
      const res = await apiWrite("saveAll", createRemoteWebPayload());
      if (res && res.ok === false) throw new Error(res.error || "Error al enviar.");
      if (show) status(res?.msg || res?.message || "Información enviada correctamente.", "ok");
      if (show) toast(res?.msg || res?.message || "Sincronizado hacia Apps Script.");
    } catch (err) {
      if (show) status(err.message, "err");
      else console.warn(err);
    }
  }

  async function pullRemote(show = true) {
    try {
      if (show) status("Cargando configuración web y catálogo desde Apps Script...");
      await pullPortalScriptData();
      saveData({ remote: false });
      renderAll();
      if (show) status("Configuración web cargada desde Apps Script y productos guardados en memoria local.", "ok");
      if (show) toast("Configuración web y MAESTRA local sincronizadas.");
    } catch (err) {
      if (show) status(err.message, "err");
    }
  }

  function bindTableActions() {
    document.addEventListener("click", (e) => {
      const docBtn = e.target.closest("[data-view-doc],[data-pdf-doc]");
      if (docBtn) {
        const kind = docBtn.dataset.viewDoc || docBtn.dataset.pdfDoc;
        const docId = docBtn.dataset.docId;
        const doc = findDocument(kind, docId);
        if (!doc) { toast("No se encontró el documento."); return; }
        if (docBtn.dataset.viewDoc) openDocumentModal(kind, docId);
        if (docBtn.dataset.pdfDoc) generateDocumentPdf(docBtn.dataset.pdfFormat || "a4", { kind, doc });
        return;
      }

      const btn = e.target.closest("[data-edit-banner],[data-toggle-banner],[data-delete-banner],[data-edit-module],[data-toggle-module],[data-delete-module],[data-edit-service],[data-toggle-service],[data-delete-service],[data-edit-product],[data-toggle-product],[data-delete-product]");
      if (!btn) return;

      const action = Object.keys(btn.dataset)[0];
      const value = btn.dataset[action];

      if (action === "editBanner") {
        const item = data.banners.find((x) => String(x.id) === String(value));
        if (item) fillBannerForm(item);
      }
      if (action === "toggleBanner") toggleItem("banners", value);
      if (action === "deleteBanner") deleteItem("banners", value, "banner");

      if (action === "editModule") {
        const item = data.modulos.find((x) => String(x.id) === String(value));
        if (item) fillModuleForm(item);
      }
      if (action === "toggleModule") toggleItem("modulos", value);
      if (action === "deleteModule") deleteItem("modulos", value, "módulo");

      if (action === "editService") {
        const item = data.servicios.find((x) => String(x.id) === String(value));
        if (item) fillServiceForm(item);
      }
      if (action === "toggleService") toggleItem("servicios", value);
      if (action === "deleteService") deleteItem("servicios", value, "servicio");

      if (action === "editProduct") {
        const item = data.productos.find((x) => String(x.id) === String(value));
        if (item) fillProductForm(item);
      }
      if (action === "toggleProduct") toggleItem("productos", value);
      if (action === "deleteProduct") deleteItem("productos", value, "producto");
    });
  }

  function toggleItem(collection, itemId) {
    const item = data[collection].find((x) => String(x.id) === String(itemId));
    if (!item) return;
    item.activo = !bool(item.activo);
    saveData();
    toast("Estado actualizado.");
  }

  function deleteItem(collection, itemId, label) {
    if (!confirm(`¿Eliminar ${label}?`)) return;
    data[collection] = data[collection].filter((x) => String(x.id) !== String(itemId));
    saveData();
    toast(`${label} eliminado.`);
  }

  function enhanceResponsiveTables() {
    $$(".admin-table").forEach((table) => {
      const headers = [...table.querySelectorAll("thead th")].map((th) => th.textContent.trim());
      table.querySelectorAll("tbody tr").forEach((row) => {
        [...row.children].forEach((cell, index) => {
          if (!cell.hasAttribute("data-label") && headers[index]) cell.setAttribute("data-label", headers[index]);
        });
      });
    });
  }

  function renderAll() {
    applyTheme();
    fillSiteForm();
    fillIntegrationForm();
    renderStats();
    renderBanners();
    renderModules();
    renderServices();
    renderProducts();
    renderQuotes();
    renderOrders();
    enhanceResponsiveTables();
  }

  document.addEventListener("DOMContentLoaded", () => {
    persistOfficialRemoteUrl();
    const cached = loadProductsCache();
    if (cached.ok) updateProductLocalStatus();
    bindNav();
    bindForms();
    bindTableActions();
    renderAll();
    startOrderWatcher();
  });
})();
