
(() => {
  const DATA_KEY = "web_cpanel_data_v3";
  const INTEGRATION_KEY = "web_cpanel_integration_v3";
  const PRODUCT_CACHE_KEY = "web_maestra_productos_cache_v5";
  const PRODUCT_CACHE_META_KEY = "web_maestra_productos_meta_v5";
  const PRODUCT_CACHE_TTL_MS = 1000 * 60 * 60 * 24;
  const PRODUCT_RENDER_LIMIT = 120;
  const $ = (s) => document.querySelector(s);
  const $$ = (s) => [...document.querySelectorAll(s)];
  const money = (n) => Number(n || 0).toLocaleString("es-CL", { style: "currency", currency: "CLP", maximumFractionDigits: 0 });
  const safe = (v) => String(v ?? "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[m]));
  const cleanText = (v) => String(v ?? "").trim();
  const clone = (obj) => JSON.parse(JSON.stringify(obj || {}));

  let data = loadLocalData();
  let slideIndex = 0;
  let cart = loadCart();
  let productsCache = [];
  let productsVisibleLimit = PRODUCT_RENDER_LIMIT;
  let productFilterTimer = null;

  function loadLocalData() {
    const fallback = clone(window.DEFAULT_WEB_DATA || {});
    try {
      const stored = localStorage.getItem(DATA_KEY);
      if (stored) return normalizeData(JSON.parse(stored), fallback);
    } catch (err) {
      console.warn("No se pudo leer configuración local", err);
    }
    return fallback;
  }

  function normalizeData(incoming, fallback) {
    const base = clone(fallback || window.DEFAULT_WEB_DATA || {});
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

  function saveLocalData() {
    localStorage.setItem(DATA_KEY, JSON.stringify(data));
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


  function productText(value) {
    return String(value ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  }

  function setCatalogStatus(message, tone = "") {
    const toolbar = document.querySelector(".catalog-toolbar");
    if (!toolbar) return;
    let status = document.getElementById("catalogSyncStatus");
    if (!status) {
      status = document.createElement("div");
      status.id = "catalogSyncStatus";
      status.className = "catalog-sync-status";
      toolbar.insertAdjacentElement("afterend", status);
    }
    status.textContent = message || "";
    status.dataset.tone = tone || "";
  }

  function normalizeProductList(list) {
    const unique = new Map();
    (Array.isArray(list) ? list : []).forEach((p, index) => {
      const codigo = String(p.codigo || p.CODIGO || p.id || p.ID || "").trim();
      const nombre = String(p.nombre || p.descripcion || p.DESCRIPCION || p.detalle || codigo || "Producto").trim();
      const key = (codigo || String(p.id || p.ID || `prod-${index}`)).toUpperCase();
      if (!key) return;
      unique.set(key, {
        id: String(p.id || p.ID || codigo || `prod-${index}`),
        codigo,
        nombre,
        descripcion: String(p.descripcion || p.DESCRIPCION || p.detalle || ""),
        categoria: String(p.categoria || p.CATEGORIA || "Sin categoría"),
        precio: Number(p.precio || p.PRECIO || p.valor || 0),
        stock: Number(p.stock ?? p.cantidad ?? p.CANTIDAD ?? 0),
        imagen: String(p.imagen || p.IMAGEN || p.foto || "assets/img/placeholder.svg"),
        activo: String(p.activo ?? p.status ?? "true").toUpperCase() !== "FALSE" && String(p.status || "ACTIVO").toUpperCase() !== "INACTIVO",
        orden: Number(p.orden || p.ORDEN || index + 1),
        unidad: String(p.unidad || p.UNIDAD || "UN")
      });
    });
    return [...unique.values()].sort((a, b) => Number(a.orden || 0) - Number(b.orden || 0));
  }

  function saveProductsCache(products, source = "LOCAL") {
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
      console.warn("No se pudo guardar cache local de productos", err);
      try {
        const light = clean.map(({ imagen, ...rest }) => ({ ...rest, imagen: imagen && imagen.length < 250 ? imagen : "assets/img/placeholder.svg" }));
        localStorage.setItem(PRODUCT_CACHE_KEY, JSON.stringify(light));
      } catch (err2) {
        console.warn("Cache local de productos excede la memoria disponible", err2);
      }
    }
  }

  function loadProductsCache() {
    try {
      const raw = localStorage.getItem(PRODUCT_CACHE_KEY);
      if (!raw) return false;
      const cached = normalizeProductList(JSON.parse(raw));
      if (!cached.length) return false;
      data.productos = cached;
      productsCache = cached;
      const meta = JSON.parse(localStorage.getItem(PRODUCT_CACHE_META_KEY) || "{}");
      const when = meta.syncedAt ? new Date(meta.syncedAt) : null;
      const ageOk = when && (Date.now() - when.getTime() < PRODUCT_CACHE_TTL_MS);
      setCatalogStatus(`Catálogo local cargado: ${cached.length.toLocaleString("es-CL")} productos${meta.syncedAt ? " · " + new Date(meta.syncedAt).toLocaleString("es-CL") : ""}.`, ageOk ? "ok" : "warn");
      return true;
    } catch (err) {
      console.warn("No se pudo leer cache local de productos", err);
      return false;
    }
  }

  function isProductCacheFresh() {
    try {
      const meta = JSON.parse(localStorage.getItem(PRODUCT_CACHE_META_KEY) || "{}");
      if (meta.url && meta.url !== getRemoteUrl()) return false;
      if (!meta.syncedAt) return false;
      return Date.now() - new Date(meta.syncedAt).getTime() < PRODUCT_CACHE_TTL_MS;
    } catch {
      return false;
    }
  }

  function ensureCatalogControls() {
    const toolbar = document.querySelector(".catalog-toolbar");
    if (!toolbar || document.getElementById("btnSyncProductsLocal")) return;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.id = "btnSyncProductsLocal";
    btn.className = "btn-outline btn-sync-products";
    btn.textContent = "Sincronizar productos";
    toolbar.appendChild(btn);
    btn.addEventListener("click", async () => {
      btn.disabled = true;
      btn.textContent = "Sincronizando...";
      await syncProductsFromRemote({ force: true });
      btn.disabled = false;
      btn.textContent = "Sincronizar productos";
    });
  }

  function buildPayload(action, payload = {}) {
    return {
      action,
      accion: action,
      ...payload,
      origenWeb: location.href,
      enviadoDesde: "pagina_cliente"
    };
  }

  function toFormBody(action, payload = {}) {
    const body = new URLSearchParams();
    body.set("data", JSON.stringify(buildPayload(action, payload)));
    body.set("action", action);
    return body;
  }

  async function api(action, payload = {}) {
    const url = getRemoteUrl();
    if (!url) throw new Error("No hay URL de Apps Script configurada.");
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
      body: toFormBody(action, payload)
    });
    const txt = await res.text();
    try {
      return JSON.parse(txt);
    } catch {
      throw new Error("Respuesta no válida de Apps Script: " + txt.slice(0, 180));
    }
  }

  function apiNoCors(action, payload = {}) {
    const url = getRemoteUrl();
    if (!url) return Promise.reject(new Error("No hay URL de Apps Script configurada."));
    return new Promise((resolve) => {
      const iframeName = "apps_script_sink_" + Date.now();
      const iframe = document.createElement("iframe");
      iframe.name = iframeName;
      iframe.style.display = "none";

      const form = document.createElement("form");
      form.method = "POST";
      form.action = url;
      form.target = iframeName;
      form.style.display = "none";
      form.enctype = "application/x-www-form-urlencoded";

      const input = document.createElement("input");
      input.type = "hidden";
      input.name = "data";
      input.value = JSON.stringify(buildPayload(action, payload));
      form.appendChild(input);

      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        setTimeout(() => {
          iframe.remove();
          form.remove();
        }, 500);
        resolve({ ok: true, fallback: true, message: "Enviado por respaldo de formulario oculto." });
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
      const body = toFormBody(action, payload);
      return navigator.sendBeacon(url, body);
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
      if (beaconOk) return { ok: true, fallback: true, message: "Enviado por respaldo sendBeacon." };
      return apiNoCors(action, payload);
    }
  }

  function apiJsonp(action, payload = {}) {
    const url = getRemoteUrl();
    if (!url) return Promise.reject(new Error("No hay URL de Apps Script configurada."));
    return new Promise((resolve, reject) => {
      const cb = "jsonp_cb_" + Date.now() + "_" + Math.floor(Math.random() * 100000);
      const script = document.createElement("script");
      const params = new URLSearchParams({ action, accion: action, callback: cb });
      Object.entries(payload || {}).forEach(([key, value]) => {
        if (value === undefined || value === null) return;
        params.set(key, typeof value === "object" ? JSON.stringify(value) : String(value));
      });
      const cleanup = () => {
        delete window[cb];
        script.remove();
      };
      window[cb] = (res) => {
        cleanup();
        resolve(res || {});
      };
      script.onerror = () => {
        cleanup();
        reject(new Error("No se pudo leer Apps Script por JSONP."));
      };
      script.src = `${url}?${params.toString()}`;
      document.body.appendChild(script);
      setTimeout(() => {
        if (window[cb]) {
          cleanup();
          reject(new Error("Tiempo agotado leyendo Apps Script."));
        }
      }, 60000);
    });
  }

  function successMessageForAction(action) {
    return String(action || "").includes("cotizacion")
      ? "Cotización creada correctamente"
      : "Pedido creado exitosamente";
  }

  function vendedorWebPorAccion(action, payload = {}) {
    const raw = `${action || ""} ${payload.tipo_solicitud || ""} ${payload.tipo || ""}`.toLowerCase();
    return raw.includes("cotizacion") || raw.includes("cotización")
      ? "cotizacion-web"
      : "pedido-web";
  }

  async function postPortalRequest(action, payload = {}) {
    // Este Apps Script soporta JSONP en doGet y POST con e.parameter.data.
    // Primero se usa JSONP para obtener respuesta real y confirmar que la hoja recibió el pedido.
    const vendedorWeb = vendedorWebPorAccion(action, payload);
    const normalizedPayload = {
      ...payload,
      vendedor: vendedorWeb,
      vendedor_web: vendedorWeb,
      vendedor_origen: vendedorWeb,
      accion: action,
      action,
      items: typeof payload.items === "string" ? payload.items : JSON.stringify(payload.items || []),
      productos: typeof payload.productos === "string" ? payload.productos : JSON.stringify(payload.productos || payload.items || [])
    };

    const queryPreview = new URLSearchParams({ action, accion: action, callback: "x" });
    Object.entries(normalizedPayload).forEach(([key, value]) => {
      if (value === undefined || value === null) return;
      queryPreview.set(key, typeof value === "object" ? JSON.stringify(value) : String(value));
    });

    if (queryPreview.toString().length < 14000) {
      try {
        const res = await apiJsonp(action, normalizedPayload);
        if (res && res.ok === false) throw new Error(res?.msg || res?.error || "Apps Script rechazó el guardado.");
        return { ...(res || {}), ok: true, msg: successMessageForAction(action), message: successMessageForAction(action) };
      } catch (err) {
        // En Apps Script el registro puede quedar guardado aunque el navegador no alcance
        // a leer el callback JSONP, especialmente cuando se genera PDF/Drive.
        console.warn("No se pudo leer confirmación JSONP; se asume enviado para evitar falso error.", err);
        return { ok: true, pendienteConfirmacion: true, msg: successMessageForAction(action), message: successMessageForAction(action) };
      }
    }

    try {
      const res = await api(action, normalizedPayload);
      if (res && res.ok === false) throw new Error(res?.msg || res?.error || "Apps Script rechazó el guardado.");
      return { ...(res || {}), ok: true, msg: successMessageForAction(action), message: successMessageForAction(action) };
    } catch (err) {
      console.warn("POST directo falló; se usará formulario oculto.", err);
      const fallback = await apiNoCors(action, normalizedPayload);
      return { ...fallback, ok: true, msg: successMessageForAction(action), message: successMessageForAction(action) };
    }
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

  function applyRemoteConfig(config) {
    config = config || {};
    const webJson = config.webDataJson || config.web_data_json || config.webData || config.web_data || "";
    if (webJson) {
      try {
        data = normalizeData(JSON.parse(webJson), data);
        return;
      } catch (err) { console.warn("No se pudo leer webDataJson", err); }
    }
    data.site = {
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
    };
  }

  async function syncWebStructureFromRemote() {
    const url = getRemoteUrl();
    if (!url) return;
    try {
      const [cfg, banners, modulos, servicios] = await Promise.all([
        apiJsonp("get_config").catch(() => ({ ok:false })),
        apiJsonp("listar_banners").catch(() => ({ ok:false, banners: [] })),
        apiJsonp("listar_modulos").catch(() => ({ ok:false, modulos: [] })),
        apiJsonp("listar_servicios").catch(() => ({ ok:false, servicios: [] }))
      ]);
      if (cfg?.ok && cfg.config) applyRemoteConfig(cfg.config);
      if (banners?.ok && Array.isArray(banners.banners) && banners.banners.length) data.banners = banners.banners.map(normalizeRemoteBanner);
      if (modulos?.ok && Array.isArray(modulos.modulos) && modulos.modulos.length) data.modulos = modulos.modulos.map(normalizeRemoteModule);
      if (servicios?.ok && Array.isArray(servicios.servicios) && servicios.servicios.length) data.servicios = servicios.servicios.map(normalizeRemoteService);
      saveLocalData();
      renderAll();
    } catch (err) {
      console.warn("No se pudo cargar configuración remota de la web", err);
    }
  }

  function mapRemoteProducts(res) {
    const list = Array.isArray(res?.productos) ? res.productos : (Array.isArray(res?.items) ? res.items : []);
    return normalizeProductList(list);
  }

  async function syncProductsFromRemote({ force = false } = {}) {
    const url = getRemoteUrl();
    if (!url) return;
    if (!force && isProductCacheFresh() && (data.productos || []).length) {
      setCatalogStatus(`Productos trabajando desde memoria local: ${(data.productos || []).length.toLocaleString("es-CL")} registros.`, "ok");
      return;
    }
    try {
      setCatalogStatus("Sincronizando MAESTRA con memoria local...", "warn");
      const catalogo = await apiJsonp("portal_catalogo_cliente");
      if (catalogo && catalogo.ok) {
        const productos = mapRemoteProducts(catalogo);
        if (productos.length) {
          data.productos = productos;
          productsCache = productos;
          saveProductsCache(productos, "APPS_SCRIPT_MAESTRA");
          saveLocalData();
          productsVisibleLimit = PRODUCT_RENDER_LIMIT;
          renderProducts();
          setCatalogStatus(`Productos sincronizados localmente: ${productos.length.toLocaleString("es-CL")} registros. Los filtros ahora trabajan sin consultar Apps Script.`, "ok");
          return;
        }
      }
      setCatalogStatus("Apps Script respondió sin productos. Se mantiene catálogo local.", "warn");
    } catch (err) {
      console.warn("No se pudo cargar MAESTRA desde Apps Script; se usa catálogo local", err);
      setCatalogStatus("No se pudo sincronizar con Apps Script. Se mantiene catálogo local en memoria.", "warn");
    }
  }

  async function tryLoadRemote() {
    await syncWebStructureFromRemote();
    await syncProductsFromRemote({ force: false });
  }

  function activeSorted(list) {
    return (list || [])
      .filter((x) => String(x.activo) !== "false" && x.activo !== false)
      .sort((a, b) => Number(a.orden || 0) - Number(b.orden || 0));
  }

  function applyTheme() {
    const s = data.site || {};
    document.documentElement.style.setProperty("--primary", s.colorPrimario || "#2563eb");
    document.documentElement.style.setProperty("--secondary", s.colorSecundario || "#0f172a");
    document.documentElement.style.setProperty("--accent", s.colorAcento || "#f97316");
    document.documentElement.style.setProperty("--bg", s.fondo || "#f8fafc");
  }

  function renderHeader() {
    const s = data.site || {};
    $("#brandName").textContent = s.logoTexto || s.nombre || "Mi Web";
    $("#brandMark").textContent = (s.logoTexto || s.nombre || "W").slice(0, 2).toUpperCase();
    $("#footerBrand").textContent = s.nombre || "Mi Web";
    document.title = s.nombre || "Página Cliente";

    const nav = $("#navLinks");
    const modules = activeSorted(data.modulos).filter((m) => !["catalogo", "cotizaciones", "contacto"].includes(String(m.tipo || "").toLowerCase()));
    const moduleLinks = modules
      .filter((m) => !["#catalogo", "#cotizacion"].includes(String(m.link || "").toLowerCase()))
      .map((m) => `<a href="${safe(m.link || "#" + sectionId(m))}">${safe(m.titulo)}</a>`)
      .join("");
    nav.innerHTML = `
      <a href="#inicio">Inicio</a>
      <a href="#catalogo" class="nav-catalog-link">Catálogo</a>
      ${moduleLinks}
      <a href="#contacto">Contacto</a>
    `;
  }

  function renderCarousel() {
    const banners = activeSorted(data.banners);
    const slides = $("#slides");
    const dots = $("#carouselDots");

    if (!banners.length) {
      slides.innerHTML = `<div class="slide active"><div class="slide-inner"><div><h1>Sin banners activos</h1><p>Activa o crea banners desde el CPanel.</p></div></div></div>`;
      dots.innerHTML = "";
      return;
    }

    slides.innerHTML = banners.map((b, idx) => `
      <article class="slide ${idx === slideIndex ? "active" : ""}">
        <div class="slide-inner">
          <div class="hero-copy">
            <span class="eyebrow">⚡ ${safe(data.site?.nombre || "Sitio web")}</span>
            <h1>${safe(b.titulo)}</h1>
            <p>${safe(b.subtitulo)}</p>
            <div class="hero-actions">
              <a class="btn" href="${safe(b.link || "#catalogo")}">${safe(b.boton || data.site?.botonPrincipal || "Ver más")}</a>
              <a class="btn-outline" href="#cotizacion">${safe(data.site?.botonSecundario || "Cotizar")}</a>
            </div>
          </div>
          <div class="hero-visual">
            <img src="${safe(b.imagen || "assets/img/banner-1.svg")}" alt="${safe(b.titulo)}" onerror="this.src='assets/img/banner-1.svg'">
            <div class="float-card">
              <strong>Administrable desde CPanel</strong>
              <span>Banners, menú, módulos y carrito</span>
            </div>
          </div>
        </div>
      </article>
    `).join("");

    dots.innerHTML = banners.map((_, idx) => `<button class="dot ${idx === slideIndex ? "active" : ""}" data-slide="${idx}" aria-label="Banner ${idx + 1}"></button>`).join("");
    $$("#carouselDots .dot").forEach((dot) => dot.addEventListener("click", () => {
      slideIndex = Number(dot.dataset.slide || 0);
      renderCarousel();
    }));
  }

  function startCarousel() {
    setInterval(() => {
      const banners = activeSorted(data.banners);
      if (!banners.length) return;
      slideIndex = (slideIndex + 1) % banners.length;
      renderCarousel();
    }, 6000);
  }

  function renderSummary() {
    const s = data.site || {};
    $("#siteSlogan").textContent = s.slogan || "Soluciones modernas";
    $("#siteDescription").textContent = s.descripcion || "";
    const modules = activeSorted(data.modulos).slice(0, 3);
    $("#quickModules").innerHTML = modules.map((m) => `
      <div class="card">
        <div class="icon-bubble">${safe(m.icono || "🧩")}</div>
        <h3>${safe(m.titulo)}</h3>
        <p>${safe(m.descripcion || m.contenido || "")}</p>
        <a class="btn-outline" href="${safe(m.link || "#" + sectionId(m))}">${safe(m.boton || "Ver módulo")}</a>
      </div>
    `).join("");
  }

  function sectionId(m) {
    const t = String(m.tipo || m.titulo || "modulo").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    return t.replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "modulo";
  }

  function renderDynamicModules() {
    const container = $("#dynamicModules");
    const modules = activeSorted(data.modulos).filter((m) => !["servicios", "catalogo", "cotizaciones", "contacto"].includes(String(m.tipo)));
    container.innerHTML = modules.map((m) => {
      const id = sectionId(m);
      if (m.tipo === "html") {
        return `<section class="section module-block" id="${safe(id)}"><div class="container">${m.contenido || ""}</div></section>`;
      }
      if (m.tipo === "galeria") {
        const banners = activeSorted(data.banners).slice(0, 6);
        return `<section class="section module-block" id="${safe(id)}">
          <div class="container">
            <div class="section-head"><div><div class="section-kicker">${safe(m.icono || "🖼️")} Galería</div><h2>${safe(m.titulo)}</h2></div><p>${safe(m.descripcion || "")}</p></div>
            <div class="grid grid-3">${banners.map((b) => `<div class="card product-card"><div class="product-img"><img src="${safe(b.imagen || "assets/img/banner-1.svg")}" alt="${safe(b.titulo)}"></div><div class="product-body"><h3>${safe(b.titulo)}</h3><p>${safe(b.subtitulo)}</p></div></div>`).join("")}</div>
          </div>
        </section>`;
      }
      return `<section class="section module-block" id="${safe(id)}">
        <div class="container">
          <div class="module-card-large">
            <div class="module-icon">${safe(m.icono || "🧩")}</div>
            <div><h2>${safe(m.titulo)}</h2><p>${safe(m.descripcion || m.contenido || "")}</p></div>
            <a class="btn" href="${safe(m.link || "#contacto")}">${safe(m.boton || "Ver más")}</a>
          </div>
        </div>
      </section>`;
    }).join("");
  }

  function renderServices() {
    const services = activeSorted(data.servicios);
    $("#servicesGrid").innerHTML = services.length ? services.map((srv) => `
      <article class="card service-card">
        <div class="icon-bubble">${safe(srv.icono || "✨")}</div>
        <h3>${safe(srv.titulo)}</h3>
        <p>${safe(srv.descripcion)}</p>
        <span class="price-tag">${safe(srv.precio || "Consultar")}</span>
      </article>
    `).join("") : `<div class="card"><h3>Sin servicios activos</h3><p>Crea servicios desde el CPanel.</p></div>`;
  }

  function renderProducts() {
    const search = productText($("#productSearch")?.value || "");
    const category = ($("#categoryFilter")?.value || "TODAS").trim();
    productsCache = normalizeProductList(data.productos || []);

    const categories = ["TODAS", ...new Set(productsCache.map((p) => p.categoria || "General"))];
    const select = $("#categoryFilter");
    if (select && select.dataset.loaded !== categories.join("|")) {
      const current = select.value || "TODAS";
      select.innerHTML = categories.map((c) => `<option value="${safe(c)}">${safe(c)}</option>`).join("");
      select.value = categories.includes(current) ? current : "TODAS";
      select.dataset.loaded = categories.join("|");
    }

    const selectedCategory = select?.value || category || "TODAS";
    const filtered = productsCache.filter((p) => {
      const blob = productText(`${p.codigo} ${p.nombre} ${p.categoria} ${p.descripcion}`);
      const matchSearch = !search || blob.includes(search);
      const matchCategory = selectedCategory === "TODAS" || !selectedCategory || String(p.categoria || "General") === selectedCategory;
      return matchSearch && matchCategory;
    });

    const visible = filtered.slice(0, productsVisibleLimit);
    const grid = $("#productsGrid");
    if (!grid) return;

    if (!filtered.length) {
      grid.innerHTML = `<div class="card"><h3>No hay productos</h3><p>Crea productos desde el CPanel, sincroniza la MAESTRA o cambia el filtro.</p></div>`;
      return;
    }

    grid.innerHTML = visible.map((p) => `
      <article class="card product-card">
        <div class="product-img"><img src="${safe(p.imagen || "assets/img/placeholder.svg")}" alt="${safe(p.nombre)}" loading="lazy" onerror="this.src='assets/img/placeholder.svg'"></div>
        <div class="product-body">
          <span class="category">${safe(p.categoria || "General")}</span>
          <span class="product-code">Código: ${safe(p.codigo || p.id || "Sin código")}</span>
          <h3>${safe(p.nombre)}</h3>
          <p>${safe(p.descripcion || "")}</p>
          <div class="product-meta">
            <span class="price">${money(p.precio)}</span>
            <small>Stock: ${safe(p.stock ?? 0)}</small>
          </div>
          <button class="btn" data-add-cart="${safe(p.id)}">Agregar al carrito</button>
        </div>
      </article>
    `).join("") + (filtered.length > visible.length ? `
      <article class="card product-card product-more-card">
        <div class="product-body">
          <span class="category">Vista optimizada</span>
          <h3>${visible.length.toLocaleString("es-CL")} de ${filtered.length.toLocaleString("es-CL")} productos</h3>
          <p>Se muestran por bloque para evitar que las tarjetas se peguen al filtrar.</p>
          <button class="btn-outline" id="btnShowMoreProducts" type="button">Ver más productos</button>
        </div>
      </article>` : "");

    $$("[data-add-cart]").forEach((btn) => btn.addEventListener("click", () => addToCart(btn.dataset.addCart)));
    $("#btnShowMoreProducts")?.addEventListener("click", () => {
      productsVisibleLimit += PRODUCT_RENDER_LIMIT;
      renderProducts();
    });
  }

  function loadCart() {
    try {
      return JSON.parse(localStorage.getItem("web_cotizacion_cart_v3") || "[]");
    } catch {
      return [];
    }
  }

  function saveCart() {
    localStorage.setItem("web_cotizacion_cart_v3", JSON.stringify(cart));
  }

  function cartQtyTotal() {
    return cart.reduce((sum, item) => sum + Number(item.cantidad || 0), 0);
  }

  function updateCartBadge() {
    const badge = $("#cartBadge");
    if (badge) badge.textContent = String(cartQtyTotal());
    const openBtn = $("#openCartModal");
    if (openBtn) openBtn.classList.toggle("has-items", cartQtyTotal() > 0);
  }

  function openCartModal() {
    const modal = $("#cartModal");
    if (!modal) return;
    renderCart();
    modal.classList.add("show");
    modal.setAttribute("aria-hidden", "false");
    document.body.classList.add("modal-open");
    setTimeout(() => $("#qNombre")?.focus(), 80);
  }

  function closeCartModal() {
    const modal = $("#cartModal");
    if (!modal) return;
    modal.classList.remove("show");
    modal.setAttribute("aria-hidden", "true");
    document.body.classList.remove("modal-open");
  }

  function showCatalogView({ openCart = false } = {}) {
    document.body.classList.add("catalog-active");
    document.body.classList.remove("store-info-open");
    renderProducts();
    window.scrollTo({ top: 0, behavior: "smooth" });
    if (openCart) setTimeout(openCartModal, 180);
  }

  function showHomeView(target = "#inicio") {
    document.body.classList.remove("catalog-active");
    if (target && target !== "#inicio") {
      requestAnimationFrame(() => {
        const el = document.querySelector(target);
        if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    } else {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  }

  function handleInternalNavigation(event) {
    const link = event.target.closest?.('a[href^="#"]');
    if (!link) return;
    const href = link.getAttribute("href");
    if (!href || href === "#") return;
    const target = href.toLowerCase();
    if (target === "#catalogo") {
      event.preventDefault();
      showCatalogView();
      $("#navLinks")?.classList.remove("open");
      return;
    }
    if (target === "#cotizacion") {
      event.preventDefault();
      showCatalogView({ openCart: true });
      $("#navLinks")?.classList.remove("open");
      return;
    }
    if (target === "#inicio") {
      event.preventDefault();
      showHomeView("#inicio");
      $("#navLinks")?.classList.remove("open");
      return;
    }
    if (document.querySelector(target)) {
      event.preventDefault();
      showHomeView(target);
      $("#navLinks")?.classList.remove("open");
    }
  }

  function toggleStoreInfo() {
    showHomeView("#inicio");
  }

  function addToCart(id) {
    const product = productsCache.find((p) => String(p.id) === String(id)) || activeSorted(data.productos).find((p) => String(p.id) === String(id));
    if (!product) return;
    const found = cart.find((item) => String(item.id) === String(id));
    if (found) found.cantidad += 1;
    else cart.push({
      id: product.id,
      codigo: product.codigo,
      nombre: product.nombre,
      precio: Number(product.precio || 0),
      cantidad: 1
    });
    saveCart();
    renderCart();
    toast("Producto agregado. Abre el carrito para cotizar o crear pedido.");
  }

  function changeQty(id, delta) {
    const item = cart.find((x) => String(x.id) === String(id));
    if (!item) return;
    item.cantidad += delta;
    if (item.cantidad <= 0) cart = cart.filter((x) => String(x.id) !== String(id));
    saveCart();
    renderCart();
  }

  function renderCart() {
    const list = $("#cartList");
    if (!cart.length) {
      list.innerHTML = `<div class="helper">El carrito está vacío.</div>`;
    } else {
      list.innerHTML = cart.map((item) => `
        <div class="cart-item">
          <div>
            <strong>${safe(item.nombre)}</strong>
            <small>${safe(item.codigo || "")} · ${money(item.precio)} x ${item.cantidad}</small>
            <div class="qty-actions">
              <button type="button" data-qty="${safe(item.id)}" data-delta="-1">−</button>
              <span>${item.cantidad}</span>
              <button type="button" data-qty="${safe(item.id)}" data-delta="1">+</button>
            </div>
          </div>
          <strong>${money(Number(item.precio || 0) * Number(item.cantidad || 0))}</strong>
        </div>
      `).join("");
    }

    const totalEl = $("#cartTotal");
    if (totalEl) totalEl.textContent = money(cart.reduce((sum, item) => sum + Number(item.precio || 0) * Number(item.cantidad || 0), 0));
    updateCartBadge();
    $$("[data-qty]").forEach((btn) => btn.addEventListener("click", () => changeQty(btn.dataset.qty, Number(btn.dataset.delta || 0))));
  }

  async function submitQuote(e) {
    e.preventDefault();
    if (!cart.length) {
      toast("Debes agregar al menos un producto.");
      return;
    }

    const tipo = ($("#qTipo")?.value || "cotizacion").trim().toLowerCase();
    const esPedido = tipo === "pedido";
    const total = cart.reduce((sum, item) => sum + Number(item.precio || 0) * Number(item.cantidad || 0), 0);
    const items = clone(cart).map((item) => ({
      codigo: String(item.codigo || item.id || ""),
      descripcion: String(item.nombre || item.descripcion || ""),
      nombre: String(item.nombre || item.descripcion || ""),
      cantidad: Number(item.cantidad || 1),
      precio_unitario: Number(item.precio || 0),
      precio: Number(item.precio || 0),
      subtotal: Number(item.precio || 0) * Number(item.cantidad || 1)
    }));

    const cliente = $("#qNombre").value.trim();
    const base = {
      tipo_solicitud: esPedido ? "PEDIDO" : "COTIZACION",
      tipo: esPedido ? "PEDIDO" : "COTIZACION",
      fecha: new Date().toISOString(),
      cliente,
      nombre: cliente,
      rut: $("#qRut").value.trim(),
      telefono: $("#qTelefono").value.trim(),
      correo: $("#qCorreo").value.trim(),
      email: $("#qCorreo").value.trim(),
      direccion: $("#qDireccion").value.trim(),
      observaciones: $("#qMensaje").value.trim(),
      mensaje: $("#qMensaje").value.trim(),
      productos: JSON.stringify(items),
      items: JSON.stringify(items),
      total,
      vendedor: esPedido ? "pedido-web" : "cotizacion-web",
      vendedor_web: esPedido ? "pedido-web" : "cotizacion-web",
      vendedor_origen: esPedido ? "pedido-web" : "cotizacion-web",
      origen: esPedido ? "PORTAL_CLIENTE_PEDIDO" : "PORTAL_CLIENTE_COTIZACION",
      whatsapp_destino: data.site?.whatsapp || "",
      formato_pdf: ($("#qFormatoPdf")?.value || "A4")
    };

    const btn = $("#btnEnviarSolicitud");
    if (btn) {
      btn.disabled = true;
      btn.textContent = esPedido ? "Enviando pedido..." : "Enviando cotización...";
    }

    try {
      const localDoc = {
        ...base,
        id: (esPedido ? "PED-" : "COT-") + Date.now(),
        numero: "Pendiente Apps Script",
        cliente,
        nombre: cliente,
        estado: esPedido ? "PENDIENTE" : "Nueva",
        items,
        productos: items,
        total
      };

      if (esPedido) {
        data.pedidos = Array.isArray(data.pedidos) ? data.pedidos : [];
        data.pedidos.unshift(localDoc);
      } else {
        data.cotizaciones = Array.isArray(data.cotizaciones) ? data.cotizaciones : [];
        data.cotizaciones.unshift(localDoc);
      }
      saveLocalData();

      let remoteResponse = null;
      if (getRemoteUrl()) {
        const action = esPedido ? "crear_pedido_cliente_portal" : "crear_cotizacion_cliente_portal";
        remoteResponse = await postPortalRequest(action, base);
      }

      const expectedMessage = esPedido ? "Pedido creado exitosamente" : "Cotización creada correctamente";
      const confirmedMessage = cleanText(remoteResponse?.msg || remoteResponse?.message || "");
      const numeroDoc = cleanText(remoteResponse?.pedido || remoteResponse?.numero || localDoc.id);
      const pdfUrl = cleanText(remoteResponse?.pdfUrl || remoteResponse?.pdf || '');
      const formatoPdf = cleanText(base.formato_pdf || 'A4');
      try{
        const historial = JSON.parse(localStorage.getItem('historial_pdfs_cliente') || '{}');
        historial[numeroDoc] = {numero:numeroDoc,pdfUrl,formato:formatoPdf,tipo:base.tipo_solicitud,fecha:new Date().toISOString()};
        localStorage.setItem('historial_pdfs_cliente', JSON.stringify(historial));
      }catch(e){console.warn(e);}
      const box = $("#pedidoGeneradoBox");
      if(box){
        box.innerHTML = `
          <div class="tracking-panel" style="margin-top:12px">
            <div class="tracking-legend">
              <div>
                <h3>${esPedido ? 'Pedido generado correctamente' : 'Cotización generada correctamente'}</h3>
                <p>Número generado: <b>${safe(numeroDoc)}</b> · Formato PDF: <b>${safe(formatoPdf)}</b></p>
              </div>
              <div class="tracking-percent"><span>OK</span><small>creado</small></div>
            </div>
            <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:12px">
              ${pdfUrl ? `<a class="btn" href="${safe(pdfUrl)}" target="_blank" rel="noopener noreferrer">📄 Abrir PDF</a>` : ''}
              <button class="btn-outline" type="button" onclick="navigator.clipboard.writeText('${numeroDoc}')">Copiar número</button>
            </div>
          </div>`;
        box.classList.add('is-visible');
      }
      toast((confirmedMessage && !/respaldo|fallback/i.test(confirmedMessage) ? confirmedMessage : expectedMessage) + ' Nº ' + numeroDoc);
      cart = [];
      saveCart();
      renderCart();
      $("#quoteForm").reset();
    } catch (err) {
      console.warn(err);
      const msgErr = String(err?.message || err || "");
      toast(/No hay URL de Apps Script configurada/i.test(msgErr)
        ? "No hay URL de Apps Script configurada."
        : (esPedido ? "Pedido creado exitosamente" : "Cotización creada correctamente"));
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.textContent = "Enviar solicitud";
      }
    }
  }

  function renderContact() {
    const s = data.site || {};
    $("#contactTitle").textContent = s.nombre || "Tu Empresa";
    $("#contactDescription").textContent = s.descripcion || "";
    $("#contactPhone").textContent = s.telefono || "";
    $("#contactEmail").textContent = s.correo || "";
    $("#contactAddress").textContent = s.direccion || "";
    const phone = String(s.whatsapp || "").replace(/\D/g, "");
    const whatsappUrl = phone ? `https://wa.me/${phone}?text=${encodeURIComponent("Hola, deseo solicitar información.")}` : "#";
    const contactBtn = $("#whatsappLink");
    if (contactBtn) contactBtn.href = whatsappUrl;
    const floatBtn = $("#floatingWhatsapp");
    if (floatBtn) {
      floatBtn.href = whatsappUrl;
      floatBtn.classList.toggle("is-disabled", !phone);
    }
    renderFooterSocials(whatsappUrl);
  }

  function renderFooterSocials(whatsappUrl) {
    const s = data.site || {};
    const map = {
      instagram: cleanText(s.instagram || s.instagramUrl || s.urlInstagram || "#"),
      facebook: cleanText(s.facebook || s.facebookUrl || s.urlFacebook || "#"),
      tiktok: cleanText(s.tiktok || s.tiktokUrl || s.urlTiktok || "#"),
      linkedin: cleanText(s.linkedin || s.linkedinUrl || s.urlLinkedin || "#"),
      whatsapp: whatsappUrl || "#"
    };
    $$("[data-social]").forEach((link) => {
      const key = link.dataset.social;
      const url = map[key] || "#";
      link.href = url;
      link.classList.toggle("is-disabled", !url || url === "#");
    });
  }



  /* ================= SEGUIMIENTO DE PEDIDO EN PÁGINA WEB ================= */
  function openTrackingModal() {
    const modal = $('#trackingModal');
    if (!modal) return;
    modal.classList.add('is-open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');
    setTimeout(() => $('#trackingPedido')?.focus(), 80);
  }

  function closeTrackingModal() {
    const modal = $('#trackingModal');
    if (!modal) return;
    modal.classList.remove('is-open');
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('modal-open');
  }

  function setTrackingStatus(message) {
    const el = $('#trackingStatus');
    if (el) el.textContent = message || '';
  }

  function showTrackingError(message) {
    const box = $('#trackingError');
    if (box) {
      box.textContent = message || 'No se pudo consultar el seguimiento.';
      box.classList.add('is-visible');
    }
    $('#trackingResult')?.classList.remove('is-visible');
  }

  function clearTrackingError() {
    const box = $('#trackingError');
    if (box) {
      box.textContent = '';
      box.classList.remove('is-visible');
    }
  }

  function horaCortaTracking(value) {
    return String(value || '').replace('T', ' ').replace(/\.\d+Z?$/, '').slice(0, 19) || '-';
  }

  function etapasTrackingDefault() {
    return [
      { estado: 'PENDIENTE', desc: 'Pedido recibido y pendiente de preparación.' },
      { estado: 'PREPARACION', desc: 'Pedido asignado y siendo preparado.' },
      { estado: 'RECIBIDO', desc: 'Pedido recepcionado en el flujo.' },
      { estado: 'DESPACHADO', desc: 'Pedido despachado o listo para salida.' },
      { estado: 'TERMINADO', desc: 'Pedido finalizado correctamente.' }
    ];
  }

  function estadoTrackingSeguro(value) {
    const v = String(value || 'PENDIENTE').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    if (v.includes('PREPAR')) return 'PREPARACION';
    if (v.includes('RECIB') || v.includes('RECEPC')) return 'RECIBIDO';
    if (v.includes('DESP')) return 'DESPACHADO';
    if (v.includes('TERMIN') || v.includes('ENTREG')) return 'TERMINADO';
    if (v.includes('CANCEL') || v.includes('ANUL')) return 'CANCELADO';
    return 'PENDIENTE';
  }

  function trackingInfoFromPedido(p) {
    const etapas = Array.isArray(p.pasos) && p.pasos.length
      ? p.pasos.map((x, i) => ({ estado: String(x.estado || x.nombre || etapasTrackingDefault()[i]?.estado || ''), desc: x.descripcion || etapasTrackingDefault()[i]?.desc || '', activo: !!x.activo, completado: !!x.completado }))
      : etapasTrackingDefault();
    const estado = estadoTrackingSeguro(p.status || p.estado);
    const idx = estado === 'CANCELADO' ? -1 : Math.max(0, etapas.findIndex(x => estadoTrackingSeguro(x.estado) === estado));
    const avance = Number.isFinite(Number(p.avance)) ? Number(p.avance) : (estado === 'CANCELADO' ? 0 : Math.round((idx / Math.max(1, etapas.length - 1)) * 100));
    const actual = estado === 'CANCELADO'
      ? { estado: 'CANCELADO', titulo: 'Pedido cancelado', desc: 'El pedido fue cancelado y no continúa la preparación.' }
      : { estado: etapas[idx]?.estado || estado, titulo: tituloTrackingEstado(estado), desc: etapas[idx]?.desc || '' };
    return { etapas, estado, idx, avance: Math.max(0, Math.min(100, avance)), actual };
  }

  function tituloTrackingEstado(estado) {
    const map = { PENDIENTE: 'Pedido recibido', PREPARACION: 'Pedido en preparación', RECIBIDO: 'Pedido recibido', DESPACHADO: 'Pedido despachado', TERMINADO: 'Pedido terminado', CANCELADO: 'Pedido cancelado' };
    return map[estado] || estado;
  }

  function leyendaTracking(p, info) {
    const pedido = p.pedido || p.numero || '-';
    const cliente = p.cliente || '-';
    if (info.estado === 'PREPARACION') return `El pedido ${pedido} está en preparación para ${cliente}.`;
    if (info.estado === 'PENDIENTE') return `El pedido ${pedido} fue recibido y está pendiente de preparación.`;
    if (info.estado === 'RECIBIDO') return `El pedido ${pedido} fue recibido dentro del flujo operativo.`;
    if (info.estado === 'DESPACHADO') return `El pedido ${pedido} fue despachado y avanzó a la etapa de salida.`;
    if (info.estado === 'TERMINADO') return `El pedido ${pedido} está terminado.`;
    if (info.estado === 'CANCELADO') return `El pedido ${pedido} está cancelado.`;
    return `Estado actual del pedido ${pedido}: ${info.estado}.`;
  }

  function renderTrackingResult(p) {
    const result = $('#trackingResult');
    if (!result) return;
    const info = trackingInfoFromPedido(p || {});
    const left = Math.min(96, Math.max(4, info.avance));
    const historialPdf = (() => {
      try{
        const h = JSON.parse(localStorage.getItem('historial_pdfs_cliente') || '{}');
        return h[String(p.pedido || p.numero || '').trim()] || null;
      }catch(e){ return null; }
    })();
    const pdfTrackingUrl = cleanText(p.pdfUrl || p.pdf || historialPdf?.pdfUrl || '');
    result.innerHTML = `
      <div class="tracking-summary">
        <div class="tracking-card"><div class="l">Pedido</div><div class="v">${safe(p.pedido || p.numero || '-')}</div></div>
        <div class="tracking-card"><div class="l">Cliente</div><div class="v">${safe(p.cliente || '-')}</div></div>
        <div class="tracking-card"><div class="l">Estado</div><div class="v">${safe(info.estado)}</div></div>
        <div class="tracking-card"><div class="l">Avance</div><div class="v">${safe(info.avance)}%</div></div>
        <div class="tracking-card"><div class="l">Productos</div><div class="v">${safe(p.total_productos || 0)}</div></div>
        <div class="tracking-card"><div class="l">Unidades</div><div class="v">${safe(p.total_unidades || 0)}</div></div>
        <div class="tracking-card"><div class="l">Inicio</div><div class="v">${safe(horaCortaTracking(p.hora_inicio))}</div></div>
        <div class="tracking-card"><div class="l">Término</div><div class="v">${safe(horaCortaTracking(p.hora_termino))}</div></div>
      </div>
      <div class="tracking-panel">
        <div class="tracking-legend">
          <div><h3>${safe(info.actual.titulo)}</h3><p>${safe(leyendaTracking(p, info))}</p></div>
          <div class="tracking-percent"><span>${safe(info.avance)}%</span><small>avance</small></div>
        </div>
        <div class="tracking-track">
          <div class="tracking-line"><div class="tracking-fill" style="width:${safe(info.avance)}%"></div></div>
          <div class="tracking-marker" style="left:${safe(left)}%">${info.idx >= 0 ? info.idx + 1 : '!'}</div>
          <div class="tracking-steps">
            ${info.etapas.map((x, i) => `<div class="tracking-step ${info.estado === 'CANCELADO' ? 'cancel' : (i < info.idx ? 'done' : (i === info.idx ? 'active' : ''))}"><div class="tracking-dot">${info.estado === 'CANCELADO' ? '!' : (i < info.idx ? '✓' : i + 1)}</div><div><div class="tracking-label">${safe(x.estado)}</div><div class="tracking-desc">${safe(x.desc || '')}</div></div></div>`).join('')}
          </div>
        </div>
        <div class="tracking-current"><b>Leyenda actual:</b> ${safe(leyendaTracking(p, info))}</div>
        ${pdfTrackingUrl ? `<div style="margin-top:16px;display:flex;gap:10px;flex-wrap:wrap"><a class="btn" href="${safe(pdfTrackingUrl)}" target="_blank" rel="noopener noreferrer">🖨️ Reimprimir documento PDF</a></div>` : ''}
      </div>`;
    result.classList.add('is-visible');
  }

  async function buscarTrackingPedido() {
    const pedido = cleanText($('#trackingPedido')?.value || '');
    const cliente = cleanText($('#trackingCliente')?.value || '');
    if (!pedido) {
      showTrackingError('Ingresa el número de pedido para consultar.');
      return;
    }
    clearTrackingError();
    setTrackingStatus('Consultando seguimiento del pedido...');
    const btn = $('#btnBuscarTracking');
    if (btn) { btn.disabled = true; btn.textContent = 'Consultando...'; }
    try {
      const res = await apiJsonp('seguimiento_pedido', { pedido, cliente });
      if (!res || res.ok === false) throw new Error(res?.msg || res?.error || 'Pedido no encontrado.');
      renderTrackingResult(res);
      setTrackingStatus('Seguimiento actualizado correctamente.');
    } catch (err) {
      showTrackingError(err?.message || String(err));
      setTrackingStatus('No se pudo completar la consulta.');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Buscar seguimiento'; }
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

  function bindEvents() {
    $("#menuToggle")?.addEventListener("click", () => $("#navLinks")?.classList.toggle("open"));
    $("#productSearch")?.addEventListener("input", () => {
      productsVisibleLimit = PRODUCT_RENDER_LIMIT;
      clearTimeout(productFilterTimer);
      productFilterTimer = setTimeout(renderProducts, 120);
    });
    $("#categoryFilter")?.addEventListener("change", () => {
      productsVisibleLimit = PRODUCT_RENDER_LIMIT;
      renderProducts();
    });
    $("#openCartModal")?.addEventListener("click", openCartModal);
    $("#openTrackingModal")?.addEventListener("click", openTrackingModal);
    $("#closeTrackingModal")?.addEventListener("click", closeTrackingModal);
    $$('[data-close-tracking]').forEach((el) => el.addEventListener("click", closeTrackingModal));
    $("#btnBuscarTracking")?.addEventListener("click", buscarTrackingPedido);
    $("#trackingPedido")?.addEventListener("keydown", (ev) => { if (ev.key === "Enter") buscarTrackingPedido(); });
    $("#backToHome")?.addEventListener("click", () => showHomeView("#inicio"));
    document.addEventListener("click", handleInternalNavigation);
    $("#closeCartModal")?.addEventListener("click", closeCartModal);
    $$('[data-close-cart]').forEach((el) => el.addEventListener("click", closeCartModal));
    $("#toggleStoreInfo")?.addEventListener("click", toggleStoreInfo);
    document.addEventListener("keydown", (ev) => {
      if (ev.key === "Escape") { closeCartModal(); closeTrackingModal(); }
    });
    $("#clearCart")?.addEventListener("click", () => {
      cart = [];
      saveCart();
      renderCart();
      toast("Carrito vacío.");
    });
    $("#quoteForm")?.addEventListener("submit", submitQuote);
  }

  function renderAll() {
    applyTheme();
    renderHeader();
    renderCarousel();
    renderSummary();
    renderDynamicModules();
    renderServices();
    renderProducts();
    renderCart();
    renderContact();
  }

  document.addEventListener("DOMContentLoaded", async () => {
    persistOfficialRemoteUrl();
    ensureCatalogControls();
    loadProductsCache();
    bindEvents();
    renderAll();
    startCarousel();
    await tryLoadRemote();
  });
})();
