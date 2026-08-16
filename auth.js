/* =====================================================
   AUTH.JS — SESIÓN + PROTECCIÓN DE MÓDULOS
   - Valida token contra Apps Script
   - Renueva token automáticamente
   - Comprueba permiso por ARCHIVO en la hoja MODULOS
   - Si no tiene permiso: vuelve a main.html
   - Si no tiene sesión: vuelve a index.html
===================================================== */

/* ========= CONFIG ========= */
const API_AUTH =
  "https://script.google.com/macros/s/AKfycbxDUcEzMGw9LWnzn-YUV89So3AFCEnplOHQuGmzq-EV_hnIMhQvgQIPY1AxvlKJPZNx/exec";

const LOGIN_PAGE = "index.html";
const HOME_PAGE  = "main.html";

// Páginas que deliberadamente pueden abrirse sin sesión.
const AUTH_PUBLIC_PAGES = new Set([
  "index.html",
  "index - copia.html",
  "webindex.html",
  "seguimiento_cliente.html"
]);

/* ========= ESTADO ========= */
let AUTH_USER = null;
let AUTH_MODULES = null;
let AUTH_ACCESS_PROMISE = null;

/* =====================================================
   HELPERS DE RUTA
===================================================== */
function authFileName(value = location.pathname){
  try{
    const clean = String(value || "")
      .split("?")[0]
      .split("#")[0]
      .replace(/\\/g, "/");
    return decodeURIComponent(clean.substring(clean.lastIndexOf("/") + 1)).toLowerCase();
  }catch{
    return "";
  }
}

function authNormalizePermissions(value){
  if(Array.isArray(value)){
    return value.map(v => String(v || "").trim()).filter(Boolean);
  }

  if(typeof value === "string"){
    const raw = value.trim();
    if(!raw) return [];
    try{
      const parsed = JSON.parse(raw);
      if(Array.isArray(parsed)){
        return parsed.map(v => String(v || "").trim()).filter(Boolean);
      }
    }catch{}
    return raw.split(",").map(v => v.trim()).filter(Boolean);
  }

  return [];
}

function authIsAdmin(user){
  return String(user?.rol || "").trim().toUpperCase() === "ADMIN";
}

function authRedirect(page){
  try{
    localStorage.removeItem("moduloActivo");
  }catch{}

  // Si el módulo está dentro de un iframe, sacar también al contenedor.
  try{
    if(window.top && window.top !== window){
      window.top.location.replace(page);
      return;
    }
  }catch{}

  location.replace(page);
}

function redirigirLogin(){
  if(authFileName() !== LOGIN_PAGE.toLowerCase()){
    authRedirect(LOGIN_PAGE);
  }
}

function redirigirInicio(){
  if(authFileName() !== HOME_PAGE.toLowerCase()){
    authRedirect(HOME_PAGE);
  }
}

/* =====================================================
   VALIDAR SESIÓN GLOBAL
===================================================== */
async function validarSesionGlobal(){
  const current = authFileName();

  // Nunca validar páginas públicas desde esta función automática.
  if(AUTH_PUBLIC_PAGES.has(current)) return null;

  const token = localStorage.getItem("token");
  if(!token){
    limpiarSesion();
    redirigirLogin();
    return null;
  }

  try{
    const r = await fetch(
      `${API_AUTH}?action=verify&token=${encodeURIComponent(token)}&_=${Date.now()}`,
      { cache:"no-store" }
    );
    const res = await r.json();

    if(!res.valid){
      limpiarSesion();
      redirigirLogin();
      return null;
    }

    if(res.token){
      localStorage.setItem("token", res.token);
    }

    AUTH_USER = res.data || {};
    AUTH_USER.permisos = authNormalizePermissions(AUTH_USER.permisos);

    // Mantener cache local coherente con lo que realmente validó el servidor.
    localStorage.setItem("usuario", AUTH_USER.user || "");
    localStorage.setItem("nombre", AUTH_USER.nombre || "");
    localStorage.setItem("rol", AUTH_USER.rol || "");
    localStorage.setItem("permisos", JSON.stringify(AUTH_USER.permisos));
    sessionStorage.setItem("user", JSON.stringify(AUTH_USER));

    return AUTH_USER;

  }catch(err){
    console.error("AUTH ERROR:", err);
    limpiarSesion();
    redirigirLogin();
    return null;
  }
}

/* =====================================================
   CARGAR REGISTRO DE MÓDULOS
===================================================== */
async function cargarModulosAuth(force = false){
  if(AUTH_MODULES && !force) return AUTH_MODULES;

  const token = localStorage.getItem("token") || "";
  const r = await fetch(
    `${API_AUTH}?action=listarModulos&token=${encodeURIComponent(token)}&_=${Date.now()}`,
    { cache:"no-store" }
  );
  const res = await r.json();

  if(res && res.auth === false){
    limpiarSesion();
    redirigirLogin();
    return [];
  }

  if(!res || res.ok === false || !Array.isArray(res.data)){
    throw new Error(res?.msg || "No fue posible cargar los módulos");
  }

  AUTH_MODULES = res.data;
  return AUTH_MODULES;
}

function buscarModuloPorArchivo(modulos, archivo){
  const target = authFileName(archivo);
  return (modulos || []).find(m => authFileName(m?.[2]) === target) || null;
}

function usuarioPuedeModulo(user, modulo){
  if(!user || !modulo) return false;
  if(authIsAdmin(user)) return true;

  const activo = String(modulo[5] || "").trim().toUpperCase();
  if(activo !== "SI") return false;

  const permiso = String(modulo[4] || "").trim();
  if(!permiso) return false; // fail-closed: módulo sin permiso configurado

  return authNormalizePermissions(user.permisos).includes(permiso);
}

function authParentFile(){
  try{
    if(!document.referrer) return "";
    const u = new URL(document.referrer, location.href);
    if(u.origin !== location.origin) return "";
    return authFileName(u.pathname);
  }catch{
    return "";
  }
}

/* =====================================================
   PROTECCIÓN DE LA PÁGINA/MÓDULO ACTUAL
===================================================== */
async function validarAccesoModuloGlobal(){
  const current = authFileName();

  if(!current || AUTH_PUBLIC_PAGES.has(current)){
    return { ok:true, public:true };
  }

  // main.html valida su sesión desde su propio arranque para evitar doble petición.
  if(current === HOME_PAGE.toLowerCase()){
    return { ok:true, home:true };
  }

  const user = await validarSesionGlobal();
  if(!user) return { ok:false, reason:"session" };

  // ADMIN tiene acceso a cualquier página interna autenticada.
  if(authIsAdmin(user)){
    return { ok:true, admin:true, user };
  }

  try{
    const modulos = await cargarModulosAuth();
    const modulo = buscarModuloPorArchivo(modulos, current);

    // Si la página está registrada como módulo, exige exactamente su permiso.
    if(modulo){
      if(usuarioPuedeModulo(user, modulo)){
        return { ok:true, user, modulo };
      }
      redirigirInicio();
      return { ok:false, reason:"permission" };
    }

    // Compatibilidad para subpáginas internas abiertas dentro de un módulo autorizado.
    // Un acceso directo a una página NO registrada se bloquea.
    const parentFile = authParentFile();
    if(parentFile){
      const parentModule = buscarModuloPorArchivo(modulos, parentFile);
      if(parentModule && usuarioPuedeModulo(user, parentModule)){
        return { ok:true, user, inheritedFrom:parentModule };
      }
    }

    console.warn(`Acceso bloqueado: ${current} no está registrado como módulo autorizado.`);
    redirigirInicio();
    return { ok:false, reason:"unregistered" };

  }catch(err){
    console.error("MODULE AUTH ERROR:", err);
    // Fail-closed: si no se puede comprobar el permiso, no mostrar el módulo.
    redirigirInicio();
    return { ok:false, reason:"verification-error" };
  }
}

/* =====================================================
   UTILIDADES
===================================================== */
function limpiarSesion(){
  localStorage.removeItem("token");
  localStorage.removeItem("usuario");
  localStorage.removeItem("nombre");
  localStorage.removeItem("rol");
  localStorage.removeItem("permisos");
  localStorage.removeItem("login_ok");
  localStorage.removeItem("moduloActivo");
  sessionStorage.removeItem("user");
  sessionStorage.removeItem("modulosPermitidos");
  AUTH_USER = null;
  AUTH_MODULES = null;
}

function cerrarSesionGlobal(){
  limpiarSesion();
  redirigirLogin();
}

function getUsuarioActual(){
  return AUTH_USER;
}

/* =====================================================
   AUTO-GUARD
   Se ejecuta al incluir auth.js en la cabecera de módulos.
===================================================== */
(function iniciarProteccionAutomatica(){
  const current = authFileName();
  if(!current || AUTH_PUBLIC_PAGES.has(current) || current === HOME_PAGE.toLowerCase()) return;

  // Evita que el contenido se vea mientras se decide el acceso.
  const style = document.createElement("style");
  style.id = "auth-guard-style";
  style.textContent = "html{visibility:hidden!important}";
  (document.head || document.documentElement).appendChild(style);

  AUTH_ACCESS_PROMISE = validarAccesoModuloGlobal()
    .then(result => {
      if(result?.ok){
        style.remove();
      }
      return result;
    })
    .catch(err => {
      console.error(err);
      redirigirInicio();
      return { ok:false, reason:"unexpected" };
    });
})();
