window.APP_REMOTE_URL = "https://script.google.com/macros/s/AKfycbxZ2GtzJPlMX3TO0Nj7kfNZV0Roqn_CSHMiNCWydDYl940zuR3HymkZVv2tOgzXM_71mQ/exec";
window.DEFAULT_WEB_DATA = {
  "version": "3.0.0",
  "site": {
    "nombre": "Tu Empresa",
    "slogan": "Soluciones modernas para clientes, cotizaciones y ventas online",
    "descripcion": "Administra tu sitio web, productos, módulos, banners, servicios, cotizaciones y pedidos desde un CPanel claro conectado a Google Sheets mediante Apps Script.",
    "telefono": "+56 9 0000 0000",
    "whatsapp": "56900000000",
    "correo": "contacto@tuempresa.cl",
    "direccion": "Santiago, Chile",
    "instagram": "#",
    "facebook": "#",
    "tiktok": "#",
    "linkedin": "#",
    "logoTexto": "Mi Web",
    "botonPrincipal": "Ver productos",
    "botonSecundario": "Solicitar cotización",
    "colorPrimario": "#2563eb",
    "colorSecundario": "#0f172a",
    "colorAcento": "#f97316",
    "fondo": "#f8fafc",
    "modo": "claro"
  },
  "integracion": {
    "appsScriptUrl": "https://script.google.com/macros/s/AKfycbxZ2GtzJPlMX3TO0Nj7kfNZV0Roqn_CSHMiNCWydDYl940zuR3HymkZVv2tOgzXM_71mQ/exec",
    "sincronizacionAutomatica": false
  },
  "banners": [
    {
      "id": "ban-001",
      "titulo": "Página web moderna y administrable",
      "subtitulo": "Banner principal tipo carrusel, menú moderno, módulos dinámicos, catálogo, carrito y cotizaciones.",
      "boton": "Comprar / Cotizar",
      "link": "#catalogo",
      "imagen": "assets/img/banner-1.svg",
      "activo": true,
      "orden": 1
    },
    {
      "id": "ban-002",
      "titulo": "CPanel completo en tema claro",
      "subtitulo": "Crea banners, servicios, productos y nuevos módulos sin modificar código.",
      "boton": "Ver servicios",
      "link": "#servicios",
      "imagen": "assets/img/banner-2.svg",
      "activo": true,
      "orden": 2
    },
    {
      "id": "ban-003",
      "titulo": "Conectado a Google Sheets",
      "subtitulo": "El backend Apps Script incluido permite guardar configuración, productos, cotizaciones y pedidos.",
      "boton": "Contacto",
      "link": "#contacto",
      "imagen": "assets/img/banner-3.svg",
      "activo": true,
      "orden": 3
    }
  ],
  "modulos": [
    {
      "id": "mod-servicios",
      "titulo": "Servicios",
      "tipo": "servicios",
      "icono": "✨",
      "descripcion": "Muestra tarjetas de servicios administradas desde el CPanel.",
      "contenido": "Nuestros servicios principales.",
      "boton": "Ver servicios",
      "link": "#servicios",
      "activo": true,
      "orden": 1
    },
    {
      "id": "mod-catalogo",
      "titulo": "Catálogo",
      "tipo": "catalogo",
      "icono": "🛒",
      "descripcion": "Catálogo de productos con carrito y cotización.",
      "contenido": "Productos disponibles para cotizar y comprar.",
      "boton": "Ver catálogo",
      "link": "#catalogo",
      "activo": true,
      "orden": 2
    },
    {
      "id": "mod-cotizaciones",
      "titulo": "Cotizaciones",
      "tipo": "cotizaciones",
      "icono": "📄",
      "descripcion": "Formulario para que el cliente envíe su cotización.",
      "contenido": "Agrega productos al carrito y solicita tu cotización.",
      "boton": "Solicitar",
      "link": "#cotizacion",
      "activo": true,
      "orden": 3
    },
    {
      "id": "mod-contacto",
      "titulo": "Contacto",
      "tipo": "contacto",
      "icono": "📞",
      "descripcion": "Datos de contacto y acceso rápido a WhatsApp.",
      "contenido": "Comunícate con nosotros para recibir soporte personalizado.",
      "boton": "Escribir",
      "link": "#contacto",
      "activo": true,
      "orden": 4
    }
  ],
  "servicios": [
    {
      "id": "srv-001",
      "titulo": "Diseño Web",
      "descripcion": "Sitio responsive, moderno y rápido, administrable desde CPanel.",
      "icono": "💻",
      "precio": "Desde $150.000",
      "activo": true,
      "orden": 1
    },
    {
      "id": "srv-002",
      "titulo": "Catálogo Online",
      "descripcion": "Productos con imágenes, filtros, carrito y cotización.",
      "icono": "🛍️",
      "precio": "Incluido",
      "activo": true,
      "orden": 2
    },
    {
      "id": "srv-003",
      "titulo": "Automatización Apps Script",
      "descripcion": "Base de datos Google Sheets conectada al sitio web y al CPanel.",
      "icono": "⚙️",
      "precio": "Configurable",
      "activo": true,
      "orden": 3
    }
  ],
  "productos": [
    {
      "id": "prod-001",
      "codigo": "SKU-001",
      "nombre": "Producto destacado",
      "categoria": "General",
      "descripcion": "Producto de ejemplo administrable desde CPanel.",
      "precio": 12990,
      "stock": 25,
      "imagen": "assets/img/producto-1.svg",
      "activo": true,
      "orden": 1
    },
    {
      "id": "prod-002",
      "codigo": "SKU-002",
      "nombre": "Servicio premium",
      "categoria": "Servicios",
      "descripcion": "Servicio editable para cotización directa.",
      "precio": 24990,
      "stock": 10,
      "imagen": "assets/img/producto-2.svg",
      "activo": true,
      "orden": 2
    },
    {
      "id": "prod-003",
      "codigo": "SKU-003",
      "nombre": "Pack empresa",
      "categoria": "Empresa",
      "descripcion": "Pack completo para empresas y emprendimientos.",
      "precio": 39990,
      "stock": 7,
      "imagen": "assets/img/producto-3.svg",
      "activo": true,
      "orden": 3
    }
  ],
  "cotizaciones": [],
  "pedidos": []
};
