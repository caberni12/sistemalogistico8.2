(function(){
  // Data store
  let orders = [];
  let filteredOrders = [];
  
  // DOM elements
  // Body for the summary table (one row per pedido)
  const summaryTableBody = document.getElementById('summaryTable');
  // Detail modal elements
  const detailsModal = document.getElementById('detailsModal');
  const detailsTitle = document.getElementById('detailsTitle');
  const detailsPedidoMeta = document.getElementById('detailsPedidoMeta');
  const detailsBody = document.getElementById('detailsBody');
  const detailsPrintFormat = document.getElementById('detailsPrintFormat');
  const detailsPrintBtn = document.getElementById('detailsPrintBtn');
  const detailsEstadoSelect = document.getElementById('detailsEstadoSelect');
  const detailsEstadoBtn = document.getElementById('detailsEstadoBtn');
  const btnCloseDetails = document.getElementById('btnCloseDetails');

  const modal = document.getElementById('modal');
  const btnAdd = document.getElementById('btnAdd');
  const btnImport = document.getElementById('btnImport');
  const btnSyncPedidos = document.getElementById('btnSyncPedidos');
  const btnSyncUbicaciones = document.getElementById('btnSyncUbicaciones');
  const btnConfigSistema = document.getElementById('btnConfigSistema');
  const configModal = document.getElementById('configModal');
  const configApiUrl = document.getElementById('configApiUrl');
  const configStatus = document.getElementById('configStatus');
  const configTxtInput = document.getElementById('configTxtInput');
  const btnLoadConfigTxt = document.getElementById('btnLoadConfigTxt');
  const btnDownloadConfigTxt = document.getElementById('btnDownloadConfigTxt');
  const btnTestConfig = document.getElementById('btnTestConfig');
  const btnSaveConfig = document.getElementById('btnSaveConfig');
  const btnCloseConfig = document.getElementById('btnCloseConfig');
  const syncInfo = document.getElementById('syncInfo');
  const fileInput = document.getElementById('fileInput');
  const searchInput = document.getElementById('searchInput');
  const statusFilter = document.getElementById('statusFilter');
  const fechaDesde = document.getElementById('fechaDesde');
  const fechaHasta = document.getElementById('fechaHasta');
  const btnClearFilters = document.getElementById('btnClearFilters');
  const btnCancel = document.getElementById('btnCancel');
  const btnSave = document.getElementById('btnSave');
  const editIndexInput = document.getElementById('editIndex');
  const inputCodigo = document.getElementById('inputCodigo');
  const inputDescripcion = document.getElementById('inputDescripcion');
  const inputUbicacion = document.getElementById('inputUbicacion');
  const inputCliente = document.getElementById('inputCliente');
  const inputVendedor = document.getElementById('inputVendedor');
  const inputPedido = document.getElementById('inputPedido');
  const inputFecha = document.getElementById('inputFecha');

  // Configuración de API dinámica
  // No vuelvas a editar el código para cambiar la URL: usa el módulo Configuración o config.txt.
  const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbwOG9W7fdVqDrhacHp3Ry1A-pM5eKVFXwFXl4Q2V2SYS1uclU9Ko7XcqS5iOcP2BKEb7g/exec';
  const API_URL_STORAGE_KEY = 'sistema_pedidos_api_url';
  let API_URL = '';
  let currentDetailsOrder = '';
  let currentDetailsCliente = '';
  let currentDetailsVendedor = '';

  // Actualización automática de pedidos cada 20 segundos.
  // Consulta pedidos pendientes, atrasados y cancelados, no duplica productos
  // y mantiene los pedidos importados localmente que aún no tienen ID de Google Sheets.
  const AUTO_REFRESH_MS = 20000;
  let autoRefreshTimer = null;
  let isAutoRefreshing = false;

  // Mapa de ubicaciones por código.
  // Se guarda en memoria local para que el sistema no se quede pegado consultando Google Sheets a cada rato.
  let ubicacionesMap = {};
  const UBICACIONES_CACHE_KEY = 'pedidos_ubicaciones_cache_v3_all_locations';
  const UBICACIONES_CACHE_TS_KEY = 'pedidos_ubicaciones_cache_v3_all_locations_ts';

  /**
   * Normaliza un código de producto quitando comillas iniciales, apóstrofes,
   * espacios y convirtiéndolo a minúsculas. Esto permite cruzar códigos con
   * ceros a la izquierda aunque vengan desde Excel, Google Sheets o scanner.
   */
  function normalizeCode(c){
    return String(c || '')
      .replace(/^['’`´]+/, '')
      .trim()
      .toLowerCase();
  }

  // Limpia el código sólo para mostrar/enviar, sin convertirlo a número ni quitar ceros a la izquierda.
  function displayCode(c){
    return String(c ?? '')
      .replace(/^['’`´]+/, '')
      .trim();
  }

  async function postToAppsScript(payload){
    if(!API_URL) throw new Error('API_URL no configurada');
    const body = new URLSearchParams();
    body.append('data', JSON.stringify(payload || {}));
    const res = await fetch(API_URL, {
      method: 'POST',
      body
    });
    const text = await res.text();
    try{
      return JSON.parse(text);
    }catch(err){
      throw new Error('Respuesta inválida del Apps Script: ' + text.slice(0, 160));
    }
  }


  function parseConfigText(text){
    const raw = String(text || '').trim();
    if(!raw) return '';
    try{
      const obj = JSON.parse(raw);
      return String(obj.API_URL || obj.apiUrl || obj.url || obj.URL_GS || '').trim();
    }catch(err){
      // No es JSON, seguimos con formato texto simple
    }
    const line = raw.split(/\r?\n/).map(x => x.trim()).find(x => x && !x.startsWith('#')) || '';
    const eq = line.match(/^(API_URL|URL_GS|URL)\s*=\s*(.+)$/i);
    if(eq) return eq[2].trim().replace(/^['"]|['"]$/g,'');
    if(/^https?:\/\//i.test(line)) return line.trim();
    return '';
  }

  async function loadApiConfig(){
    const saved = localStorage.getItem(API_URL_STORAGE_KEY);
    if(saved && saved.trim()){
      API_URL = saved.trim();
    }else{
      try{
        const res = await fetch('config.txt?ts=' + Date.now(), { cache:'no-store' });
        if(res.ok){
          const urlFromTxt = parseConfigText(await res.text());
          if(urlFromTxt){
            API_URL = urlFromTxt;
            localStorage.setItem(API_URL_STORAGE_KEY, API_URL);
          }
        }
      }catch(err){
        // En file:// algunos navegadores bloquean fetch de config.txt; se mantiene fallback.
      }
    }
    if(!API_URL) API_URL = DEFAULT_API_URL;
    if(configApiUrl) configApiUrl.value = API_URL;
    updateConfigStatus('URL activa: ' + API_URL);
    return API_URL;
  }

  function updateConfigStatus(text){
    if(configStatus) configStatus.textContent = text || '';
  }

  function abrirConfiguracion(){
    if(configApiUrl) configApiUrl.value = API_URL || DEFAULT_API_URL;
    updateConfigStatus('URL activa: ' + (API_URL || DEFAULT_API_URL));
    configModal.classList.add('active');
  }

  function cerrarConfiguracion(){
    configModal.classList.remove('active');
  }

  function guardarConfiguracion(){
    const url = (configApiUrl.value || '').trim();
    if(!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/i.test(url)){
      updateConfigStatus('URL inválida. Debe terminar en /exec y ser de script.google.com/macros/s/.');
      return;
    }
    API_URL = url;
    localStorage.setItem(API_URL_STORAGE_KEY, API_URL);
    updateConfigStatus('Configuración guardada. URL activa: ' + API_URL);
  }

  async function probarConfiguracion(){
    guardarConfiguracion();
    if(!API_URL) return;
    updateConfigStatus('Probando conexión...');
    try{
      const res = await fetch(API_URL + '?accion=listar_ubicaciones_actuales&test=' + Date.now(), { cache:'no-store' });
      const data = await res.json();
      if(data && data.ok){
        updateConfigStatus('Conexión correcta. Apps Script respondió OK.');
      }else{
        updateConfigStatus('La URL respondió, pero no entregó OK. Revisa que pegaste el Apps Script actualizado.');
      }
    }catch(err){
      updateConfigStatus('No se pudo conectar. Revisa URL, permisos del despliegue y conexión.');
    }
  }

  function descargarConfigTxt(){
    const url = (configApiUrl.value || API_URL || DEFAULT_API_URL).trim();
    const contenido = `API_URL=${url}\n`;
    const blob = new Blob([contenido], {type:'text/plain;charset=utf-8'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'config.txt';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function cargarConfigTxtArchivo(file){
    if(!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const url = parseConfigText(reader.result);
      if(url){
        configApiUrl.value = url;
        guardarConfiguracion();
      }else{
        updateConfigStatus('No se encontró una URL válida dentro del archivo.');
      }
    };
    reader.readAsText(file);
  }

  function updateSyncInfo(texto){
    if(syncInfo){
      syncInfo.textContent = texto || '';
    }
  }

  function cargarUbicacionesDesdeMemoria(){
    try{
      const raw = localStorage.getItem(UBICACIONES_CACHE_KEY);
      if(!raw) return false;
      const parsed = JSON.parse(raw);
      if(!parsed || typeof parsed !== 'object') return false;
      ubicacionesMap = parsed;
      const ts = localStorage.getItem(UBICACIONES_CACHE_TS_KEY);
      const total = Object.keys(ubicacionesMap).length;
      const fecha = ts ? new Date(Number(ts)).toLocaleString() : 'sin fecha';
      updateSyncInfo(`Ubicaciones en memoria: ${total} | Última sincronización: ${fecha}`);
      return total > 0;
    }catch(err){
      console.error('No se pudo leer memoria de ubicaciones', err);
      return false;
    }
  }

  function guardarUbicacionesEnMemoria(){
    try{
      localStorage.setItem(UBICACIONES_CACHE_KEY, JSON.stringify(ubicacionesMap || {}));
      localStorage.setItem(UBICACIONES_CACHE_TS_KEY, String(Date.now()));
      const total = Object.keys(ubicacionesMap || {}).length;
      updateSyncInfo(`Ubicaciones sincronizadas: ${total} | ${new Date().toLocaleString()}`);
    }catch(err){
      console.error('No se pudo guardar memoria de ubicaciones', err);
    }
  }

  function aplicarUbicacionesEnPedidos(){
    let actualizados = 0;
    orders.forEach(item => {
      const loc = getUbicacion(item.codigo);
      if(loc && item.ubicacion !== loc){
        item.ubicacion = loc;
        actualizados++;
      }
    });
    return actualizados;
  }

  // Carga ubicaciones.
  // Por defecto usa la memoria local; sólo consulta Google Sheets cuando force=true
  // o cuando aún no existe caché.
  async function loadUbicaciones(options = {}){
    const force = options.force === true;

    if(!force && cargarUbicacionesDesdeMemoria()){
      return true;
    }

    if(!API_URL){
      updateSyncInfo('API_URL no configurada. No se pueden sincronizar ubicaciones.');
      return false;
    }

    try{
      updateSyncInfo('Sincronizando ubicaciones desde Google Sheets...');
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 25000);

      // Endpoint optimizado: devuelve TODAS las ubicaciones por código desde la BD-MOVIMIENTO.
      const res = await fetch(`${API_URL}?accion=listar_ubicaciones_actuales`, { signal: controller.signal });
      clearTimeout(timer);
      const data = await res.json();

      if(data && data.ok && Array.isArray(data.data)){
        const nuevoMapa = {};
        const movimientos = data.data;

        // Si el endpoint devuelve objetos, cada objeto ya viene con todas las ubicaciones unidas.
        for(let i = movimientos.length - 1; i >= 0; i--){
          const row = movimientos[i];

          let codigoRow = '';
          let ubic = '';

          if(Array.isArray(row)){
            // Compatibilidad con acción listar: [id, fecha_registro, fecha_entrada, fecha_salida, ubicacion, codigo, ...]
            codigoRow = normalizeCode(row[5]);
            ubic = String(row[4] || '').trim();
          }else if(row && typeof row === 'object'){
            // Acción optimizada listar_ubicaciones_actuales: {codigo, ubicacion, ubicaciones:[...]}
            codigoRow = normalizeCode(row.codigo);
            if(Array.isArray(row.ubicaciones)){
              ubic = row.ubicaciones
                .map(u => (u && typeof u === 'object') ? (u.cantidad ? `${u.ubicacion} (${u.cantidad})` : u.ubicacion) : String(u || ''))
                .filter(Boolean)
                .join(' | ');
            }else{
              ubic = String(row.ubicacion || '').trim();
            }
          }

          if(!codigoRow || !ubic) continue;
          if(!nuevoMapa[codigoRow]){
            nuevoMapa[codigoRow] = ubic;
          }
        }

        ubicacionesMap = nuevoMapa;
        guardarUbicacionesEnMemoria();
        aplicarUbicacionesEnPedidos();
        updateFilter();
        return true;
      }

      updateSyncInfo('No se recibieron ubicaciones válidas desde la BD-MOVIMIENTO.');
      return false;
    }catch(err){
      console.error('Error al sincronizar ubicaciones', err);
      updateSyncInfo('No se pudo sincronizar. Se mantiene la memoria local.');
      return false;
    }
  }

  // Obtener ubicación por código desde la memoria local
  function getUbicacion(cod){
    const key = normalizeCode(cod);
    return ubicacionesMap[key] || '';
  }

  function escapeHtml(value){
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function ubicacionHtml(value){
    return escapeHtml(value).replace(/\s*\|\s*/g, '<br>');
  }

  function ubicacionPdf(value){
    return String(value || '').replace(/\s*\|\s*/g, '\n');
  }

  function normalizarEstado(value){
    const estado = String(value || 'pendiente')
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
    if(estado === 'terminado') return 'terminado';
    if(estado === 'cancelado') return 'cancelado';
    if(estado === 'atrasado') return 'atrasado';
    return 'pendiente';
  }

  function labelEstado(value){
    const estado = normalizarEstado(value);
    return estado.charAt(0).toUpperCase() + estado.slice(1);
  }

  function estadoVisibleEnListado(value){
    // Sólo se oculta/elimina visualmente el estado TERMINADO.
    // Todo estado distinto de terminado debe permanecer visible en la TD/listado de pedidos.
    const estado = normalizarEstado(value || 'pendiente');
    return estado !== 'terminado';
  }

  function appsScriptJsonp(params, timeoutMs=25000){
    return new Promise((resolve, reject) => {
      if(!API_URL){
        reject(new Error('API_URL no configurada'));
        return;
      }
      const cbName = '__sgl_cb_' + Date.now() + '_' + Math.floor(Math.random() * 1000000);
      const script = document.createElement('script');
      const timer = setTimeout(() => {
        cleanup();
        reject(new Error('Tiempo de espera agotado al actualizar Apps Script'));
      }, timeoutMs);
      function cleanup(){
        clearTimeout(timer);
        try{ delete window[cbName]; }catch(e){ window[cbName] = undefined; }
        if(script && script.parentNode) script.parentNode.removeChild(script);
      }
      window[cbName] = function(data){
        cleanup();
        resolve(data);
      };
      const qs = new URLSearchParams(Object.assign({}, params, {
        callback: cbName,
        ts: Date.now()
      }));
      script.onerror = function(){
        cleanup();
        reject(new Error('No se pudo conectar con Apps Script'));
      };
      script.src = API_URL + '?' + qs.toString();
      document.head.appendChild(script);
    });
  }

  async function appsScriptPostForm(params, timeoutMs=25000){
    if(!API_URL) throw new Error('API_URL no configurada');
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    try{
      const res = await fetch(API_URL, {
        method: 'POST',
        body: new URLSearchParams(params),
        cache: 'no-store',
        signal: controller.signal
      });
      const text = await res.text();
      try{
        return JSON.parse(text);
      }catch(err){
        throw new Error('Respuesta no JSON de Apps Script: ' + text.slice(0,180));
      }
    }finally{
      clearTimeout(timeout);
    }
  }

  function itemsPedidoParaServidor(orderNum){
    return orders
      .filter(item => String(item.pedido || '') === String(orderNum || ''))
      .map(item => ({
        id: item.id || '',
        codigo: displayCode(item.codigo || ''),
        descripcion: item.descripcion || '',
        ubicacion: item.ubicacion || getUbicacion(item.codigo) || '',
        pedido: item.pedido || orderNum,
        cliente: item.cliente || '',
        vendedor: item.vendedor || '',
        fecha: item.fecha || '',
        status: item.status || 'pendiente',
        cantidad: item.cantidad || 1
      }));
  }

  function pedidoProductoKey(item){
    return String(item?.pedido || '').trim() + '|' + normalizeCode(item?.codigo || '');
  }

  function existePedidoProducto(item){
    const key = pedidoProductoKey(item);
    if(!key || key === '|') return false;
    return orders.some(actual => pedidoProductoKey(actual) === key);
  }

  function agregarPedidoPendienteSinDuplicar(item){
    if(!item || !item.codigo || !item.pedido) return false;
    if(!estadoVisibleEnListado(item.status || 'pendiente')) return false;
    if(existePedidoProducto(item)) return false;
    orders.push(item);
    return true;
  }


  function esItemServidor(item){
    return String(item?.id || '').trim().startsWith('PD-');
  }

  function reconciliarPedidosServidor(pedidosServidor, options = {}){
    const silent = !!options.silent;
    const serverByKey = new Map();
    pedidosServidor.forEach(item => {
      const key = pedidoProductoKey(item);
      if(key && key !== '|') serverByKey.set(key, item);
    });

    let agregados = 0;
    let actualizados = 0;
    let removidos = 0;

    const nuevos = [];
    orders.forEach(actual => {
      const key = pedidoProductoKey(actual);
      if(serverByKey.has(key)){
        const servidor = serverByKey.get(key);
        const combinado = Object.assign({}, actual, servidor);
        const antes = JSON.stringify({
          id: actual.id || '', codigo: actual.codigo || '', descripcion: actual.descripcion || '',
          ubicacion: actual.ubicacion || '', pedido: actual.pedido || '', cliente: actual.cliente || '',
          vendedor: actual.vendedor || '', status: actual.status || '', fecha: actual.fecha || ''
        });
        const despues = JSON.stringify({
          id: combinado.id || '', codigo: combinado.codigo || '', descripcion: combinado.descripcion || '',
          ubicacion: combinado.ubicacion || '', pedido: combinado.pedido || '', cliente: combinado.cliente || '',
          vendedor: combinado.vendedor || '', status: combinado.status || '', fecha: combinado.fecha || ''
        });
        nuevos.push(combinado);
        serverByKey.delete(key);
        if(antes !== despues) actualizados++;
      }else if(esItemServidor(actual)){
        // No ocultar ni eliminar de la vista pedidos con estados distintos de TERMINADO.
        // Si Apps Script todavía no devuelve un atrasado/cancelado por caché o despliegue,
        // se conserva en la TD/listado local para evitar que desaparezca.
        const estadoActual = normalizarEstado(actual.status || 'pendiente');
        if(estadoActual === 'terminado'){
          removidos++;
        }else{
          nuevos.push(actual);
        }
      }else{
        // Pedido importado manualmente y aún no sincronizado: se conserva.
        nuevos.push(actual);
      }
    });

    serverByKey.forEach(item => {
      nuevos.push(item);
      agregados++;
    });

    orders = nuevos;
    orders.sort(compararPedidoMasNuevo);

    if(!silent && (agregados || actualizados || removidos)){
      updateSyncInfo(`Pedidos actualizados automáticamente. Nuevos: ${agregados}. Actualizados: ${actualizados}. Retirados terminados: ${removidos}.`);
    }

    return { agregados, actualizados, removidos };
  }

  async function loadPedidosFromServer(options = {}){
    if(!API_URL) return { agregados:0, actualizados:0, removidos:0 };
    const silent = !!options.silent;
    try{
      const res = await fetch(`${API_URL}?accion=listar_pedidos&ts=${Date.now()}`, { cache:'no-store' });
      const data = await res.json();
      if(!data || !data.ok || !Array.isArray(data.data)) return { agregados:0, actualizados:0, removidos:0 };

      const pedidos = data.data
        .filter(r => Array.isArray(r) && String(r[0] || '').startsWith('PD-'))
        .map(r => ({
          id: String(r[0] || ''),
          codigo: displayCode(r[1] || ''),
          descripcion: String(r[2] || '').trim(),
          ubicacion: String(r[3] || '').trim(),
          pedido: String(r[4] || '').trim(),
          cliente: String(r[5] || '').trim(),
          vendedor: String(r[6] || '').trim(),
          status: String(r[7] || 'pendiente').trim() || 'pendiente',
          fecha: String(r[8] || '').trim()
        }))
        .filter(x => x.codigo && x.pedido && estadoVisibleEnListado(x.status || 'pendiente'));

      const resultado = reconciliarPedidosServidor(pedidos, { silent });
      await ensureLocationsForItems(orders);
      updateFilter();
      return resultado;
    }catch(err){
      console.error('No se pudieron cargar pedidos desde Apps Script', err);
      if(!silent){
        updateSyncInfo('No se pudieron actualizar pedidos desde Apps Script. Revisa la conexión o la URL.');
      }
      return { agregados:0, actualizados:0, removidos:0, error:true };
    }
  }

  async function autoActualizarPedidos(){
    if(isAutoRefreshing || !API_URL) return;
    if(document.hidden) return;
    isAutoRefreshing = true;
    try{
      const resultado = await loadPedidosFromServer({ silent:true });
      aplicarUbicacionesEnPedidos();
      updateFilter();
      const totalCambios = (resultado.agregados || 0) + (resultado.actualizados || 0) + (resultado.removidos || 0);
      if(totalCambios){
        updateSyncInfo(`Actualización automática cada 20 segundos. Nuevos: ${resultado.agregados || 0}. Actualizados: ${resultado.actualizados || 0}. Retirados terminados: ${resultado.removidos || 0}. ${new Date().toLocaleTimeString()}`);
      }
    }finally{
      isAutoRefreshing = false;
    }
  }

  function iniciarActualizacionAutomatica(){
    if(autoRefreshTimer) clearInterval(autoRefreshTimer);
    autoRefreshTimer = setInterval(autoActualizarPedidos, AUTO_REFRESH_MS);
    updateSyncInfo('Actualización automática activa cada 20 segundos.');
  }

  function numeroPedidoOrden(value){
    const raw = String(value || '').trim();
    const nums = raw.match(/\d+/g);
    if(!nums) return -1;
    const joined = nums.join('');
    const n = Number(joined);
    return Number.isFinite(n) ? n : -1;
  }

  function fechaHoraOrden(value){
    const raw = String(value || '').trim();
    if(!raw) return 0;

    // Formatos comunes: dd-mm-yyyy, dd/mm/yyyy, yyyy-mm-dd, con o sin hora.
    let m = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})(?:[\s,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(a\.?\s*m\.?|p\.?\s*m\.?|am|pm)?)?/i);
    if(m){
      let year = String(m[3]);
      if(year.length === 2) year = '20' + year;
      let hour = Number(m[4] || 0);
      const min = Number(m[5] || 0);
      const sec = Number(m[6] || 0);
      const ap = String(m[7] || '').toLowerCase().replace(/\s|\./g,'');
      if(ap === 'pm' && hour < 12) hour += 12;
      if(ap === 'am' && hour === 12) hour = 0;
      return new Date(Number(year), Number(m[2])-1, Number(m[1]), hour, min, sec).getTime() || 0;
    }

    m = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[\s,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if(m){
      return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)).getTime() || 0;
    }

    const d = new Date(raw);
    return Number.isNaN(d.getTime()) ? 0 : d.getTime();
  }

  function compararPedidoMasNuevo(a, b){
    const pedidoB = numeroPedidoOrden(b && b.pedido);
    const pedidoA = numeroPedidoOrden(a && a.pedido);
    if(pedidoB !== pedidoA) return pedidoB - pedidoA;

    const fechaB = fechaHoraOrden(b && b.fecha);
    const fechaA = fechaHoraOrden(a && a.fecha);
    if(fechaB !== fechaA) return fechaB - fechaA;

    return String(b && b.id || '').localeCompare(String(a && a.id || ''));
  }

  // Group orders by pedido to generate summary rows
  function groupOrders(arr){
    const map = {};
    arr.forEach(item => {
      const key = item.pedido;
      if(!map[key]){
        map[key] = {
          pedido: key,
          cliente: item.cliente || '',
          vendedor: item.vendedor || '',
          fecha: item.fecha || '',
          status: normalizarEstado(item.status || 'pendiente')
        };
      }
    });
    return Object.values(map).sort(compararPedidoMasNuevo);
  }

  // Render summary table (one row per pedido)
  function renderSummary(data){
    summaryTableBody.innerHTML = '';
    if(!data.length){
      summaryTableBody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px">Sin registros</td></tr>';
      return;
    }
    data.forEach(order => {
      const estado = normalizarEstado(order.status);
      const pedidoArg = JSON.stringify(order.pedido || '');
      const clienteArg = JSON.stringify(order.cliente || '');
      const vendedorArg = JSON.stringify(order.vendedor || '');
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(order.pedido)}</td>
        <td>${escapeHtml(order.cliente)}</td>
        <td>${escapeHtml(order.vendedor)}</td>
        <td>${escapeHtml(order.fecha || '')}</td>
        <td><span class="status-badge status-${estado}">${labelEstado(estado)}</span></td>
        <td class="actions-cell">
          <button class="actions-button" onclick='showDetails(${pedidoArg})'>👁️ Ver / Acciones</button>
        </td>
      `;
      summaryTableBody.appendChild(tr);
    });
  }
  
  function fechaToComparable(value){
    const raw = String(value || '').trim();
    if(!raw) return '';
    const firstPart = raw.split(/[,\s]+/)[0].trim();

    let m = firstPart.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if(m){
      return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
    }

    m = firstPart.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
    if(m){
      let year = String(m[3]);
      if(year.length === 2) year = '20' + year;
      return `${year}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
    }

    const d = new Date(raw);
    if(!Number.isNaN(d.getTime())){
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    }
    return '';
  }

  // Update filtered orders and render summary
  function updateFilter(){
    const txt = (searchInput.value || '').toLowerCase().trim();
    const estadoFiltro = normalizarEstado(statusFilter ? statusFilter.value : '');
    const desde = fechaDesde ? fechaDesde.value : '';
    const hasta = fechaHasta ? fechaHasta.value : '';

    const filtered = orders.filter(item => {
      const itemStatus = normalizarEstado(item.status || 'pendiente');
      const itemFecha = fechaToComparable(item.fecha || '');

      const matchTexto = !txt || (
        displayCode(item.codigo || '').toLowerCase().includes(txt) ||
        String(item.descripcion || '').toLowerCase().includes(txt) ||
        String(item.ubicacion || '').toLowerCase().includes(txt) ||
        String(item.pedido || '').toLowerCase().includes(txt) ||
        String(item.cliente || '').toLowerCase().includes(txt) ||
        String(item.vendedor || '').toLowerCase().includes(txt) ||
        String(item.fecha || '').toLowerCase().includes(txt) ||
        String(item.status || '').toLowerCase().includes(txt)
      );

      const matchEstado = !estadoFiltro || itemStatus === estadoFiltro;
      const matchDesde = !desde || (itemFecha && itemFecha >= desde);
      const matchHasta = !hasta || (itemFecha && itemFecha <= hasta);

      return matchTexto && matchEstado && matchDesde && matchHasta;
    });

    const summary = groupOrders(filtered);
    renderSummary(summary);
  }

  // Show modal for adding/editing
  function openModal(editIndex){
    if(editIndex != null){
      editIndexInput.value = editIndex;
      const item = orders[editIndex];
      inputCodigo.value = displayCode(item.codigo);
      inputDescripcion.value = item.descripcion;
      inputUbicacion.value = item.ubicacion;
      inputCliente.value = item.cliente || '';
      inputVendedor.value = item.vendedor || '';
      inputFecha.value = item.fecha || '';
      inputPedido.value = item.pedido;
    }else{
      editIndexInput.value = '';
      inputCodigo.value = '';
      inputDescripcion.value = '';
      inputUbicacion.value = '';
      inputCliente.value = '';
      inputVendedor.value = '';
      inputFecha.value = '';
      inputPedido.value = '';
    }
    modal.classList.add('active');
  }
  function closeModal(){
    modal.classList.remove('active');
  }
  
  // Save entry from modal
  function saveEntry(){
    const codigo = displayCode(inputCodigo.value);
    const descripcion = inputDescripcion.value.trim();
    // Obtener ubicación desde el input o, si está vacío, desde el mapa
    let ubicacion = inputUbicacion.value.trim();
    if(!ubicacion){
      const locMap = getUbicacion(codigo);
      if(locMap){
        ubicacion = locMap;
        inputUbicacion.value = locMap;
      }
    }
    const cliente = inputCliente.value.trim();
    const vendedor = inputVendedor.value.trim();
    const pedido = inputPedido.value.trim();
    const fecha = inputFecha.value.trim();
    if(!codigo || !descripcion || !pedido){
      alert('Ingrese todos los campos obligatorios: código, descripción y pedido.');
      return;
    }
    const idx = editIndexInput.value;
    if(idx){
      const existing = orders[idx] || {};
      orders[idx] = {
        codigo,
        descripcion,
        ubicacion,
        cliente,
        vendedor,
        pedido,
        fecha,
        status: existing.status || 'pendiente'
      };
    }else{
      orders.push({
        codigo,
        descripcion,
        ubicacion,
        cliente,
        vendedor,
        pedido,
        fecha,
        status: 'pendiente'
      });
    }
    closeModal();
    updateFilter();
  }
  
  // Edit entry
  window.editEntry = function(index){
    openModal(index);
  };
  
  // Delete entry
  window.deleteEntry = function(index){
    if(confirm('¿Eliminar este registro?')){
      orders.splice(index,1);
      updateFilter();
    }
  };

  window.cambiarEstadoPedido = async function(orderNum, btn){
    const select = (String(currentDetailsOrder || '') === String(orderNum || '') && detailsEstadoSelect)
      ? detailsEstadoSelect
      : document.getElementById('estado_' + encodeURIComponent(orderNum || ''));
    const nuevoEstado = normalizarEstado(select ? select.value : 'pendiente');
    const estadoAnterior = orders.find(item => String(item.pedido) === String(orderNum))?.status || 'pendiente';
    const itemsPedido = itemsPedidoParaServidor(orderNum);

    if(!itemsPedido.length){
      alert('No se encontraron productos asociados al pedido ' + orderNum + '.');
      return;
    }

    await ensureLocationsForItems(itemsPedido);

    const originalText = btn ? btn.textContent : '';
    if(btn){
      btn.disabled = true;
      btn.textContent = 'Guardando...';
      btn.classList.add('is-loading');
    }

    try{
      if(!API_URL){
        throw new Error('API_URL no configurada. Revise el módulo Configuración.');
      }

      const basePayload = {
        accion: 'actualizar_estado_pedido',
        pedido: orderNum,
        nuevo_estado: nuevoEstado,
        cliente: itemsPedido[0]?.cliente || '',
        vendedor: itemsPedido[0]?.vendedor || '',
        items: JSON.stringify(itemsPedido)
      };

      // Primero usamos JSONP por GET para evitar bloqueos CORS en Apps Script.
      // Si el URL queda demasiado largo, usamos POST form como respaldo.
      let data = null;
      const testQuery = new URLSearchParams(Object.assign({}, basePayload, { callback:'x', ts:Date.now() })).toString();
      if((API_URL + '?' + testQuery).length < 7500){
        data = await appsScriptJsonp(basePayload);
      }else{
        data = await appsScriptPostForm(basePayload);
      }

      if(!data || !data.ok){
        throw new Error((data && data.msg) ? data.msg : 'Apps Script no confirmó la actualización');
      }

      // Confirmado por Apps Script: recién aquí actualizamos la vista.
      if(nuevoEstado === 'terminado' && data.removido_de_pedidos){
        orders = orders.filter(item => String(item.pedido) !== String(orderNum));
        if(String(currentDetailsOrder || '') === String(orderNum || '')){
          detailsModal.classList.remove('active');
        }
      }else{
        orders.forEach(item => {
          if(String(item.pedido) === String(orderNum)){
            item.status = nuevoEstado;
          }
        });
        if(String(currentDetailsOrder || '') === String(orderNum || '') && detailsEstadoSelect){
          detailsEstadoSelect.value = nuevoEstado;
        }
      }
      updateFilter();
      alert(data.msg || 'Estado actualizado correctamente.');
    }catch(err){
      console.error('Error al actualizar estado del pedido', err);
      orders.forEach(item => {
        if(String(item.pedido) === String(orderNum)){
          item.status = normalizarEstado(estadoAnterior);
        }
      });
      updateFilter();
      if(select) select.value = normalizarEstado(estadoAnterior);
      alert('No se pudo guardar el estado en Apps Script. Se dejó el estado anterior.\nDetalle: ' + (err && err.message ? err.message : err));
    }finally{
      if(btn){
        btn.disabled = false;
        btn.textContent = originalText || 'Cambiar';
        btn.classList.remove('is-loading');
      }
    }
  };

  // Show order details in modal
  // Completa ubicaciones sólo desde la memoria local. No consulta Google Sheets aquí
  // para evitar que el sistema se pegue al abrir pedidos o generar PDF.
  async function ensureLocationsForItems(items){
    if(!items || !items.length) return;
    if(!ubicacionesMap || !Object.keys(ubicacionesMap).length){
      cargarUbicacionesDesdeMemoria();
    }

    for(const item of items){
      const loc = getUbicacion(item.codigo);
      if(loc){
        item.ubicacion = loc;
      }
    }
  }

  // Show order details in modal
  window.showDetails = async function(orderNum){
    const items = orders.filter(item => item.pedido === orderNum);
    // Intentar completar las ubicaciones vacías
    await ensureLocationsForItems(items);
    const firstItem = items[0] || {};
    currentDetailsOrder = String(orderNum || '');
    currentDetailsCliente = firstItem.cliente || '';
    currentDetailsVendedor = firstItem.vendedor || '';
    detailsTitle.textContent = 'Pedido: ' + orderNum;
    if(detailsPedidoMeta){
      detailsPedidoMeta.textContent = `Cliente: ${currentDetailsCliente || '—'} | Vendedor: ${currentDetailsVendedor || '—'} | Fecha: ${firstItem.fecha || '—'}`;
    }
    if(detailsEstadoSelect){
      detailsEstadoSelect.value = normalizarEstado(firstItem.status || 'pendiente');
    }
    if(detailsPrintFormat){
      detailsPrintFormat.value = 'a4';
    }
    detailsBody.innerHTML = '';
    if(!items.length){
      detailsBody.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:20px">Sin productos</td></tr>';
    } else {
      items.forEach(item => {
        const idx = orders.indexOf(item);
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(displayCode(item.codigo))}</td>
          <td>${escapeHtml(item.descripcion)}</td>
          <td class="ubicacion-cell">${ubicacionHtml(item.ubicacion)}</td>
          <td>
            <button class="actions-button" onclick="editEntry(${idx})">✏️</button>
            <button class="actions-button" onclick="deleteEntry(${idx})">🗑️</button>
          </td>
        `;
        detailsBody.appendChild(tr);
      });
    }
    detailsModal.classList.add('active');
  };

  // Close details modal
  btnCloseDetails.addEventListener('click', () => {
    detailsModal.classList.remove('active');
  });

  if(detailsPrintBtn){
    detailsPrintBtn.addEventListener('click', () => {
      if(!currentDetailsOrder){
        alert('No hay pedido seleccionado.');
        return;
      }
      generatePDF(
        currentDetailsOrder,
        currentDetailsCliente,
        currentDetailsVendedor,
        detailsPrintFormat && detailsPrintFormat.value ? detailsPrintFormat.value : 'a4'
      );
    });
  }

  if(detailsEstadoBtn){
    detailsEstadoBtn.addEventListener('click', () => {
      if(!currentDetailsOrder){
        alert('No hay pedido seleccionado.');
        return;
      }
      cambiarEstadoPedido(currentDetailsOrder, detailsEstadoBtn);
    });
  }
  


  window.getPrintFormat = function(orderNum){
    const sel = document.getElementById('formato_' + encodeURIComponent(orderNum || ''));
    return sel && sel.value ? sel.value : 'a4';
  };


  // Generate PDF for a specific order
  window.getPrintFormat = window.getPrintFormat || function(orderNum){
    const sel = document.getElementById('formato_' + encodeURIComponent(orderNum || ''));
    return sel && sel.value ? sel.value : 'a4';
  };

  window.generatePDF = async function(orderNum, clienteName='', vendedorName='', formato='a4'){
    // Collect all entries with this order number
    const items = orders.filter(item => item.pedido === orderNum);
    if(!items.length){
      alert('No se encontraron productos para este pedido.');
      return;
    }

    // Asegurar que las ubicaciones estén completas antes de generar el PDF
    await ensureLocationsForItems(items);

    const firstItem = items[0] || {};
    const cliente = clienteName || firstItem.cliente || '';
    const vendedor = vendedorName || firstItem.vendedor || '';
    const fechaPedido = firstItem.fecha || new Date().toLocaleDateString('es-CL');
    const { jsPDF } = window.jspdf;

    if(String(formato || '').toLowerCase() === 'ticket80'){
      generarTicket80PDF({ jsPDF, items, orderNum, cliente, vendedor, fechaPedido });
      return;
    }

    generarA4PDF({ jsPDF, items, orderNum, cliente, vendedor, fechaPedido });
  };

  function generarA4PDF({ jsPDF, items, orderNum, cliente, vendedor, fechaPedido }){
    const doc = new jsPDF('p', 'pt', 'a4');
    const totalPagesExp = '{total_pages_count_string}';
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    function drawPdfHeader(data){
      const pageNumber = data && data.pageNumber ? data.pageNumber : doc.internal.getNumberOfPages();
      const marginLeft = 36;
      const marginTop = 26;
      const boxWidth = pageWidth - 72;
      const titleH = 30;
      const metaH = 34;
      const infoH = 64;
      const boxHeight = titleH + metaH + infoH;

      function fitOneLine(text, maxWidth){
        let value = String(text || '');
        if(doc.getTextWidth(value) <= maxWidth) return value;
        while(value.length > 0 && doc.getTextWidth(value + '...') > maxWidth){
          value = value.slice(0, -1);
        }
        return value + '...';
      }

      function splitLimited(text, maxWidth, maxLines){
        const raw = String(text || '');
        let lines = doc.splitTextToSize(raw, maxWidth);
        if(lines.length > maxLines){
          lines = lines.slice(0, maxLines);
          lines[maxLines - 1] = fitOneLine(lines[maxLines - 1], maxWidth);
        }
        return lines;
      }

      function drawInfoCell(x, y, w, h, label, value, maxLines){
        doc.setDrawColor(15, 23, 42);
        doc.setLineWidth(0.45);
        doc.rect(x, y, w, h);
        doc.setTextColor(15, 23, 42);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8);
        doc.text(String(label || ''), x + 6, y + 13, { maxWidth: w - 12 });
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8);
        const lines = splitLimited(value, w - 12, maxLines || 3);
        doc.text(lines, x + 6, y + 28, { maxWidth: w - 12, lineHeightFactor: 1.12 });
      }

      // Marco general de página
      doc.setDrawColor(30, 41, 59);
      doc.setLineWidth(0.7);
      doc.rect(24, 20, pageWidth - 48, pageHeight - 44);

      // Encabezado principal en cuadro
      doc.setDrawColor(15, 23, 42);
      doc.setLineWidth(0.8);
      doc.rect(marginLeft, marginTop, boxWidth, boxHeight);

      // Franja de título
      doc.setFillColor(15, 118, 110);
      doc.rect(marginLeft, marginTop, boxWidth, titleH, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(13);
      doc.text('SISTEMA DE GESTIÓN LOGÍSTICA, SERVICIOS AS', marginLeft + 12, marginTop + 20, { maxWidth: boxWidth - 24 });

      // Fila de detalle y fecha del PDF
      const metaY = marginTop + titleH;
      const halfW = boxWidth / 2;
      doc.setDrawColor(15, 23, 42);
      doc.setLineWidth(0.45);
      doc.rect(marginLeft, metaY, halfW, metaH);
      doc.rect(marginLeft + halfW, metaY, halfW, metaH);

      doc.setTextColor(15, 23, 42);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      doc.text('DETALLE DE PEDIDO', marginLeft + 12, metaY + 21, { maxWidth: halfW - 24 });
      doc.text('FECHA PDF', marginLeft + halfW + 12, metaY + 21, { maxWidth: 70 });
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8.5);
      doc.text(new Date().toLocaleString('es-CL'), marginLeft + halfW + 82, metaY + 21, { maxWidth: halfW - 94 });

      // Fila de información: cada dato queda dentro de su celda, con texto ajustado.
      const infoY = metaY + metaH;
      const pedidoW = 82;
      const fechaW = 88;
      const clienteW = 178;
      const vendedorW = boxWidth - pedidoW - clienteW - fechaW;
      let x = marginLeft;
      drawInfoCell(x, infoY, pedidoW, infoH, 'PEDIDO:', orderNum, 2);
      x += pedidoW;
      drawInfoCell(x, infoY, clienteW, infoH, 'CLIENTE:', cliente, 3);
      x += clienteW;
      drawInfoCell(x, infoY, vendedorW, infoH, 'VENDEDOR:', vendedor, 3);
      x += vendedorW;
      drawInfoCell(x, infoY, fechaW, infoH, 'FECHA PEDIDO:', fechaPedido, 3);

      // Pie de página con numeración
      doc.setDrawColor(148, 163, 184);
      doc.setLineWidth(0.4);
      doc.line(36, pageHeight - 38, pageWidth - 36, pageHeight - 38);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(71, 85, 105);
      doc.text('Documento generado automáticamente por Sistema de Gestión Logística, Servicios AS', 36, pageHeight - 22, { maxWidth: pageWidth - 190 });
      doc.text('Página ' + pageNumber + ' de ' + totalPagesExp, pageWidth - 112, pageHeight - 22);
    }

    // Prepare table data
    const body = items.map((it, idx) => [
      String(idx + 1),
      displayCode(it.codigo || ''),
      String(it.descripcion || ''),
      ubicacionPdf(it.ubicacion || '')
    ]);

    doc.autoTable({
      startY: 166,
      margin: { top: 166, left: 36, right: 36, bottom: 54 },
      head: [['#', 'Código', 'Descripción', 'Ubicación']],
      body: body,
      theme: 'grid',
      showHead: 'everyPage',
      pageBreak: 'auto',
      rowPageBreak: 'avoid',
      tableLineColor: [15, 23, 42],
      tableLineWidth: 0.6,
      styles: {
        font: 'helvetica',
        fontSize: 8.5,
        cellPadding: 5,
        valign: 'top',
        overflow: 'linebreak',
        lineColor: [51, 65, 85],
        lineWidth: 0.35,
        textColor: [15, 23, 42]
      },
      headStyles: {
        fillColor: [20, 184, 166],
        textColor: 255,
        fontStyle: 'bold',
        halign: 'center',
        lineColor: [15, 23, 42],
        lineWidth: 0.55
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      columnStyles: {
        0: { cellWidth: 28, halign: 'center' },
        1: { cellWidth: 82 },
        2: { cellWidth: 230 },
        3: { cellWidth: 'auto' }
      },
      didDrawPage: drawPdfHeader
    });

    if(typeof doc.putTotalPages === 'function'){
      doc.putTotalPages(totalPagesExp);
    }

    // Show PDF in new tab
    doc.output('dataurlnewwindow');
  }

  function estimarAltoTicket(items){
    const base = 58;
    const pie = 22;
    const filas = items.reduce((acc, it) => {
      const desc = String(it.descripcion || '');
      const ubic = ubicacionPdf(it.ubicacion || '');
      const lineasDesc = Math.max(1, Math.ceil(desc.length / 30));
      const lineasUbic = Math.max(1, Math.ceil(String(ubic).length / 17));
      return acc + Math.max(9, 4 + (lineasDesc * 3.2) + (lineasUbic * 3.2));
    }, 0);
    return Math.max(160, Math.min(2500, base + filas + pie));
  }

  function generarTicket80PDF({ jsPDF, items, orderNum, cliente, vendedor, fechaPedido }){
    const ticketHeight = estimarAltoTicket(items);
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [80, ticketHeight] });
    const pageWidth = 80;
    const margin = 4;

    doc.setTextColor(15, 23, 42);
    doc.setDrawColor(15, 23, 42);
    doc.setLineWidth(0.25);
    doc.rect(3, 3, pageWidth - 6, ticketHeight - 6);

    doc.setFillColor(15, 118, 110);
    doc.rect(3, 3, pageWidth - 6, 10, 'F');
    doc.setTextColor(255,255,255);
    doc.setFont('helvetica','bold');
    doc.setFontSize(8);
    doc.text('SISTEMA DE GESTIÓN LOGÍSTICA', pageWidth / 2, 9.5, { align:'center', maxWidth: pageWidth - 8 });

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(7);
    doc.text('SERVICIOS AS', pageWidth / 2, 17, { align:'center' });
    doc.setFont('helvetica','bold');
    doc.setFontSize(8);
    doc.text('PEDIDO: ' + String(orderNum || ''), margin, 24, { maxWidth: pageWidth - (margin * 2) });

    doc.setFont('helvetica','normal');
    doc.setFontSize(6.5);
    doc.text('Cliente: ' + String(cliente || ''), margin, 30, { maxWidth: pageWidth - (margin * 2) });
    doc.text('Vendedor: ' + String(vendedor || ''), margin, 36, { maxWidth: pageWidth - (margin * 2) });
    doc.text('Fecha pedido: ' + String(fechaPedido || ''), margin, 42, { maxWidth: pageWidth - (margin * 2) });
    doc.text('Fecha PDF: ' + new Date().toLocaleString('es-CL'), margin, 48, { maxWidth: pageWidth - (margin * 2) });

    const body = items.map((it, idx) => [
      String(idx + 1),
      displayCode(it.codigo || ''),
      String(it.descripcion || ''),
      ubicacionPdf(it.ubicacion || '')
    ]);

    doc.autoTable({
      startY: 53,
      margin: { left: margin, right: margin, bottom: 8 },
      head: [['#', 'Código', 'Descripción', 'Ubicación']],
      body,
      theme: 'grid',
      tableWidth: pageWidth - (margin * 2),
      showHead: 'everyPage',
      pageBreak: 'auto',
      rowPageBreak: 'avoid',
      styles: {
        font: 'helvetica',
        fontSize: 5.4,
        cellPadding: 1.1,
        valign: 'top',
        overflow: 'linebreak',
        lineColor: [51, 65, 85],
        lineWidth: 0.15,
        textColor: [15, 23, 42]
      },
      headStyles: {
        fillColor: [20, 184, 166],
        textColor: 255,
        fontStyle: 'bold',
        halign: 'center',
        fontSize: 5.5,
        lineWidth: 0.15
      },
      columnStyles: {
        0: { cellWidth: 6.5, halign: 'center' },
        1: { cellWidth: 15.5 },
        2: { cellWidth: 32.5 },
        3: { cellWidth: 17.5 }
      }
    });

    const finalY = Math.min(ticketHeight - 11, (doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY + 6 : ticketHeight - 18));
    doc.setFont('helvetica','normal');
    doc.setFontSize(6);
    doc.setTextColor(71, 85, 105);
    doc.text('Formato ticket 80mm - Documento generado automáticamente', pageWidth / 2, finalY, { align:'center', maxWidth: pageWidth - 8 });
    doc.output('dataurlnewwindow');
  }

  // Import file and optionally send new orders to the Apps Script
  async function handleFile(evt){
    const file = evt.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = async function(e){
      // Antes de procesar el archivo, asegúrate de que el mapa de ubicaciones esté cargado.
      // Esto garantiza que getUbicacion devuelva resultados correctos al importar pedidos.
      await loadUbicaciones();

      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      // Parsear como matriz usando texto formateado (raw:false). Esto es clave para
      // conservar códigos como 00123 cuando Excel los trae como texto/formato visible.
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false, blankrows: false });
      const itemsForServer = [];
      if(!rows.length){
        updateSyncInfo('Archivo vacío o sin encabezados.');
        fileInput.value = '';
        return;
      }

      const normalizeHeader = value => String(value || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-zA-Z0-9]/g, '')
        .toLowerCase();

      const headers = rows[0].map(normalizeHeader);
      const colIndex = aliases => {
        const normalizedAliases = aliases.map(normalizeHeader);
        for(const alias of normalizedAliases){
          const idx = headers.indexOf(alias);
          if(idx !== -1) return idx;
        }
        return -1;
      };
      const getValue = (row, aliases) => {
        const idx = colIndex(aliases);
        if(idx === -1) return '';
        return row[idx] ?? '';
      };

      for(let r = 1; r < rows.length; r++){
        const row = rows[r];
        const codigo = displayCode(getValue(row, ['codigo','código','producto','sku']));
        const descripcion = String(getValue(row, ['descripcion','descripción','detalle','producto descripcion','descripcion producto']) || '').trim();
        const ubicacion = String(getValue(row, ['ubicacion','ubicación','posicion','posición']) || '').trim();
        const pedido = String(getValue(row, ['pedido','numero','número','nro pedido','nro_pedido','orden','orden compra']) || '').trim();
        const cliente = String(getValue(row, ['cliente','nombre cliente','razon social','razón social']) || '').trim();
        const vendedor = String(getValue(row, ['vendedor','asesor','ejecutivo']) || '').trim();
        const fecha = String(getValue(row, ['fecha','fecha pedido','fecha_pedido','fecha ingreso']) || '').trim();
        const estadoRaw = String(getValue(row, ['status','estado']) || '').trim() || 'pendiente';

        if(codigo && descripcion && pedido){
          const locFromMap = getUbicacion(codigo);
          const finalUbicacion = (locFromMap || ubicacion || '').trim();
          const newItem = {
            codigo: displayCode(codigo),
            descripcion,
            ubicacion: finalUbicacion,
            cliente,
            vendedor,
            pedido,
            fecha,
            status: normalizarEstado(estadoRaw || 'pendiente')
          };
          if(agregarPedidoPendienteSinDuplicar(newItem)){
            itemsForServer.push(newItem);
          }
        }
      }
      // El archivo queda cargado en pantalla. La BD-PEDIDOS se actualiza sólo al presionar Sincronizar pedidos.
      updateFilter();
      if(itemsForServer.length){
        updateSyncInfo(`Archivo cargado: ${itemsForServer.length} producto(s) nuevo(s). Presiona Sincronizar pedidos para enviarlos a Google Sheets.`);
      }else{
        updateSyncInfo('Archivo leído: no se agregaron productos nuevos visibles.');
      }
      fileInput.value = ''; // reset
    };
    reader.readAsArrayBuffer(file);
  }
  

  async function syncPedidosToServer(){
    if(!API_URL){
      alert('Primero configura la URL del Apps Script.');
      return;
    }
    const pendientes = orders
      .filter(item => normalizarEstado(item.status || 'pendiente') === 'pendiente')
      .map(item => ({
        codigo: displayCode(item.codigo || ''),
        descripcion: String(item.descripcion || '').trim(),
        ubicacion: String(item.ubicacion || '').trim(),
        pedido: String(item.pedido || '').trim(),
        cliente: String(item.cliente || '').trim(),
        vendedor: String(item.vendedor || '').trim(),
        fecha: String(item.fecha || '').trim(),
        status: 'pendiente'
      }))
      .filter(item => item.codigo && item.descripcion && item.pedido);

    if(!pendientes.length){
      updateSyncInfo('No hay productos locales pendientes. Actualizando listado desde Google Sheets en tiempo real...');
      await loadPedidosFromServer({ silent:false, force:true });
      aplicarUbicacionesEnPedidos();
      updateFilter();
      updateSyncInfo(`Listado actualizado en tiempo real desde Google Sheets. ${new Date().toLocaleTimeString()}`);
      return;
    }

    const originalText = btnSyncPedidos.textContent;
    btnSyncPedidos.disabled = true;
    btnSyncPedidos.textContent = 'Sincronizando...';

    try{
      const result = await postToAppsScript({
        accion: 'importar_pedidos',
        items: pendientes
      });

      if(result && result.ok){
        const insertados = Number(result.insertados || 0);
        const recibidos = Number(result.recibidos || pendientes.length);
        updateSyncInfo(`Sincronización en tiempo real ejecutada. Recibidos: ${recibidos}. Nuevos guardados en PEDIDOS: ${insertados}. Actualizando listado...`);
        // Sincronización inmediata: después de guardar en Apps Script, se consulta PEDIDOS al instante.
        // Esto deja la pantalla exactamente igual que la BD, sin esperar los 20 segundos del automático.
        await loadPedidosFromServer({ silent:false, force:true });
        aplicarUbicacionesEnPedidos();
        updateFilter();
        updateSyncInfo(`Sincronización lista en tiempo real. Recibidos: ${recibidos}. Nuevos guardados: ${insertados}. Ya existentes no se duplicaron. ${new Date().toLocaleTimeString()}`);
        alert(`Sincronización lista.\nRecibidos: ${recibidos}\nNuevos guardados: ${insertados}\nLos existentes no se duplicaron.`);
      }else{
        const msg = (result && result.msg) ? result.msg : 'No se pudo sincronizar con Apps Script.';
        throw new Error(msg);
      }
    }catch(err){
      console.error('Error al sincronizar pedidos', err);
      alert('No se pudo sincronizar pedidos: ' + (err.message || err));
    }finally{
      btnSyncPedidos.textContent = originalText;
      btnSyncPedidos.disabled = false;
    }
  }

  // Event listeners
  btnAdd.addEventListener('click', () => openModal());
  btnImport.addEventListener('click', () => fileInput.click());
  btnSyncPedidos.addEventListener('click', syncPedidosToServer);
  fileInput.addEventListener('change', handleFile);
  searchInput.addEventListener('input', updateFilter);
  statusFilter.addEventListener('change', updateFilter);
  fechaDesde.addEventListener('change', updateFilter);
  fechaHasta.addEventListener('change', updateFilter);
  btnClearFilters.addEventListener('click', () => {
    searchInput.value = '';
    statusFilter.value = '';
    fechaDesde.value = '';
    fechaHasta.value = '';
    updateFilter();
  });
  btnCancel.addEventListener('click', closeModal);
  btnSave.addEventListener('click', saveEntry);

  btnConfigSistema.addEventListener('click', abrirConfiguracion);
  btnCloseConfig.addEventListener('click', cerrarConfiguracion);
  btnSaveConfig.addEventListener('click', guardarConfiguracion);
  btnTestConfig.addEventListener('click', probarConfiguracion);
  btnDownloadConfigTxt.addEventListener('click', descargarConfigTxt);
  btnLoadConfigTxt.addEventListener('click', () => configTxtInput.click());
  configTxtInput.addEventListener('change', e => cargarConfigTxtArchivo(e.target.files[0]));

  btnSyncUbicaciones.addEventListener('click', async () => {
    const originalText = btnSyncUbicaciones.textContent;
    btnSyncUbicaciones.disabled = true;
    btnSyncUbicaciones.textContent = 'Sincronizando...';
    await loadUbicaciones({ force:true });
    const actualizados = aplicarUbicacionesEnPedidos();
    updateFilter();
    btnSyncUbicaciones.textContent = originalText;
    btnSyncUbicaciones.disabled = false;
    if(actualizados){
      updateSyncInfo(`${syncInfo.textContent} | Pedidos actualizados: ${actualizados}`);
    }
  });

    // Cuando se ingresa código manualmente, rellenar ubicación si existe
    inputCodigo.addEventListener('blur', () => {
      const cod = inputCodigo.value.trim();
      if(!cod) return;
      const loc = getUbicacion(cod);
      if(loc){
        inputUbicacion.value = loc;
      }
    });
  
    // Initialize: usar ubicaciones desde memoria y cargar pedidos.
    // Sólo sincroniza ubicaciones desde Google Sheets si no existe memoria local.
    (async function init(){
      await loadApiConfig();
      await loadUbicaciones();
      await loadPedidosFromServer();
      aplicarUbicacionesEnPedidos();
      updateFilter();
      iniciarActualizacionAutomatica();
    })();
})();
