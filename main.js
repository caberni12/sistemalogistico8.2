/***************************************************
CONFIGURACIÓN GENERAL
***************************************************/
const API =
"https://script.google.com/macros/s/AKfycbxDUcEzMGw9LWnzn-YUV89So3AFCEnplOHQuGmzq-EV_hnIMhQvgQIPY1AxvlKJPZNx/exec";

let USER_IP = "—";
let WATCH_ID = null;
let LAST_GEOCODE = 0;
let MAPA = null;
let MARCADOR = null;
let CIRCULO = null;
let mapaVisible = false; // Estado del mapa

/***************************************************
🔥 RESTAURACIÓN INMEDIATA
***************************************************/
(function(){
  const data = localStorage.getItem("moduloActivo");
  if(data){
    try{
      const {url} = JSON.parse(data);
      if(url){
        window.addEventListener("DOMContentLoaded", ()=>{
          const viewer = document.getElementById("viewer");
          const frame = document.getElementById("frame");
          if(viewer && frame){
            viewer.style.display = "flex";
            frame.src = url;
          }
        });
      }
    }catch(e){}
  }
})();

/***************************************************
ICONOS
***************************************************/
function crearIconoAuto(rot = 0){
  return L.divIcon({
    className: "auto-icon",
    iconSize: [48,48],
    iconAnchor: [24,24],
    html: `<svg viewBox="0 0 64 64" style="transform:rotate(${rot}deg)">
      <rect x="10" y="24" width="44" height="16" rx="6" fill="#111"/>
      <rect x="18" y="20" width="28" height="10" rx="4" fill="#222"/>
      <circle cx="20" cy="42" r="4" fill="#000"/>
      <circle cx="44" cy="42" r="4" fill="#000"/>
    </svg>`
  });
}

const ICONO_ESTATICO = L.divIcon({
  className:"auto-icon",
  iconSize:[32,32],
  iconAnchor:[16,16],
  html:`<svg viewBox="0 0 24 24" fill="#dc2626">
    <path d="M12 2C8 2 4 6 4 10c0 6 8 14 8 14s8-8 8-14c0-4-4-8-8-8z"/>
  </svg>`
});

/***************************************************
MENÚ
***************************************************/
function toggleMenu(){
  const menu = document.getElementById("menuLateral");
  if(menu) menu.classList.toggle("open");
}

/***************************************************
LOADER
***************************************************/
function iniciarProgreso(){
  const overlay = document.getElementById("loadingOverlay");
  const bar = document.getElementById("progressBar");
  if(overlay) overlay.style.display = "flex";
  if(bar) bar.style.display = "block";
}

function finalizarProgreso(){
  const overlay = document.getElementById("loadingOverlay");
  const bar = document.getElementById("progressBar");
  if(overlay) overlay.style.display = "none";
  if(bar) bar.style.display = "none";
}

/***************************************************
INIT PRINCIPAL
***************************************************/
window.addEventListener("DOMContentLoaded", ()=>{

  iniciarProgreso();

  // Botón menú
  const btnMenu = document.getElementById("btnMenu");
  if(btnMenu) btnMenu.addEventListener("click", toggleMenu);

  // Botón mostrar mapa
  const btnMapa = document.getElementById("btnMostrarMapa");
  if(btnMapa) btnMapa.addEventListener("click", toggleMapa);

  // Inicializar app
  iniciarApp();
});

/***************************************************
INICIAR APP
***************************************************/
async function iniciarApp(){
  try{
    const user = validarSesionGlobal ? await validarSesionGlobal() : null;
    if(!user) return cerrarSesion();

    document.getElementById("usuario").textContent =
      `👤 ${user.nombre} · ${user.rol}`;

    if(typeof cargarEmpresaHeader === "function") await cargarEmpresaHeader();
    await cargarMenu(user);

    setTimeout(obtenerIP,0);
    setTimeout(iniciarReloj,0);

    // Restaurar módulo activo
    const moduloGuardado = localStorage.getItem("moduloActivo");
    if(moduloGuardado){
      try{
        const {titulo, url} = JSON.parse(moduloGuardado);
        if(url){
          const viewer = document.getElementById("viewer");
          const frame = document.getElementById("frame");
          if(viewer && frame){
            viewer.style.display = "flex";
            frame.src = url;
          }
        }
        if(titulo) document.getElementById("tituloSistema").textContent = titulo;
      }catch(e){}
    }

  }catch(e){ console.error(e); }
  finally{ finalizarProgreso(); }
}

/***************************************************
CARGAR MENÚ
***************************************************/
async function cargarMenu(user){
  try{
    const r = await fetch(`${API}?action=listarModulos`);
    const res = await r.json();
    const cont = document.getElementById("menuModulos");
    if(!cont) return;
    cont.innerHTML = "";

    if(!Array.isArray(res.data)) return;

    res.data.forEach(m=>{
      const [id,nombre,archivo,icono,permiso,activo] = m;
      if(activo !== "SI") return;
      if(user.rol !== "ADMIN" && !user.permisos.includes(permiso)) return;

      const item = document.createElement("div");
      item.className = "menu-item";
      item.innerHTML = `${icono||"📦"} ${nombre}`;
      item.addEventListener("click", ()=>{
        abrirModulo(archivo,nombre);
        toggleMenu();
      });
      cont.appendChild(item);
    });
  }catch(e){ console.error(e); }
}

/***************************************************
VISOR
***************************************************/
function abrirModulo(url,titulo){
  localStorage.setItem("moduloActivo", JSON.stringify({url,titulo}));
  const viewer = document.getElementById("viewer");
  const frame = document.getElementById("frame");
  if(viewer && frame){
    viewer.style.display = "flex";
    frame.src = url;
    document.getElementById("tituloSistema").textContent = titulo;
  }
}

function volver(){
  const viewer = document.getElementById("viewer");
  const frame = document.getElementById("frame");
  if(viewer && frame){
    viewer.style.display = "none";
    frame.src = "";
    document.getElementById("tituloSistema").textContent = "Panel Logístico";
  }
  localStorage.removeItem("moduloActivo");
}

/***************************************************
MAPA SHOW/HIDE
***************************************************/
async function toggleMapa(){
  const mapaBox = document.getElementById("mapaBox");
  const btn = document.getElementById("btnMostrarMapa");

  if(!mapaBox || !btn) return;

  if(!mapaVisible){
    mapaBox.style.display = "block";
    btn.textContent = "🗺️ Ocultar Mapa";
    if(!MAPA) await iniciarMapa();
  }else{
    mapaBox.style.display = "none";
    btn.textContent = "🗺️ Mostrar Mapa";
  }
  mapaVisible = !mapaVisible;
}

/***************************************************
MAPA + GPS (CARGA SOLO CUANDO SE NECESITE)
***************************************************/
async function iniciarMapa(){
  if(!navigator.geolocation || !window.L) return;

  WATCH_ID = navigator.geolocation.watchPosition(pos=>{
    const {latitude:lat, longitude:lng, speed=0, heading=0} = pos.coords;

    if(!MAPA){
      MAPA = L.map("mapa").setView([lat,lng],16);
      L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png").addTo(MAPA);
      MARCADOR = L.marker([lat,lng],{icon:ICONO_ESTATICO}).addTo(MAPA);
      CIRCULO = L.circle([lat,lng],{radius:80,color:"#2563eb",fillOpacity:.15}).addTo(MAPA);
    }

    MARCADOR.setLatLng([lat,lng]);
    MARCADOR.setIcon(speed>2?crearIconoAuto(heading):ICONO_ESTATICO);

    const conn = navigator.connection || {};
    let r=80;
    if(conn.effectiveType==="4g") r=150;
    if(conn.effectiveType==="3g") r=100;
    if(conn.effectiveType==="2g") r=60;
    CIRCULO.setLatLng([lat,lng]);
    CIRCULO.setRadius(r);

    actualizarRedVelocidad(speed);

    if(Date.now()-LAST_GEOCODE>15000){
      LAST_GEOCODE=Date.now();
      actualizarDireccion(lat,lng);
    }

  },()=>{}, {enableHighAccuracy:true,maximumAge:2000,timeout:10000});
}

/***************************************************
DIRECCIÓN, RED Y VELOCIDAD
***************************************************/
async function actualizarDireccion(lat,lng){
  try{
    const r = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}`);
    const d = await r.json();
    document.getElementById("dirTexto").textContent = d.display_name || "—";
  }catch{ document.getElementById("dirTexto").textContent = "—"; }
}

function actualizarRedVelocidad(speed){
  const kmh=(speed*3.6).toFixed(1);
  const conn=navigator.connection||{};
  document.getElementById("netTexto").textContent=
    `${navigator.onLine?"Online":"Offline"} · ${conn.effectiveType||"—"}`;
  document.getElementById("speedTexto").textContent=`🚗 ${kmh} km/h`;
}

/***************************************************
IP + RELOJ
***************************************************/
async function obtenerIP(){
  try{ USER_IP=(await(await fetch("https://api.ipify.org?format=json")).json()).ip; }catch{}
}

function iniciarReloj(){
  setInterval(()=>{
    const n=new Date();
    document.getElementById("conexionInfo").innerHTML=
      `📅 ${n.toLocaleDateString("es-CL")}<br>⏰ ${n.toLocaleTimeString("es-CL")}<br>🌐 IP: ${USER_IP}`;
  },1000);
}

/***************************************************
BOTONES
***************************************************/
function recargarPanel(){ location.reload(); }
function cerrarSesion(){
  if(WATCH_ID) navigator.geolocation.clearWatch(WATCH_ID);
  sessionStorage.clear();
  localStorage.clear();
  location.href="index.html";
}