(function(){
  const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbyOR_9D7lshIw2sLwWSKvwiHwfhHMXaR9qV2he9gh5cxc3nKdIOjuq3FZgbtJWwrRR48g/exec';
  const AUTO_REFRESH_MS = 2000;
  const DEVICE_ID = (() => {
    const key = 'sgsa_device_id';
    let id = localStorage.getItem(key);
    if(!id){
      id = 'DEV-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
      localStorage.setItem(key, id);
    }
    return id;
  })();
  const SHEET_NAME = window.SHEET_NAME || 'PEDIDOS';
  const SHEET_TITLE = window.SHEET_TITLE || SHEET_NAME;
  const STATUS_VALUES = ['pendiente','atrasado','cancelado','terminado','recibido','despachado','recepcionado'];

  let API_URL = DEFAULT_API_URL;
  let autoTimer = null;
  let isLoading = false;
  let rows = [];
  let headers = [];

  const $ = (id) => document.getElementById(id);
  const tableHead = $('tableHead');
  const tableBody = $('tableBody');
  const searchInput = $('searchInput');
  const statusFilter = $('statusFilter');
  const fechaDesde = $('fechaDesde');
  const fechaHasta = $('fechaHasta');
  const syncInfo = $('syncInfo');
  const btnSync = $('btnSync');
  const btnClear = $('btnClear');
  const btnExport = $('btnExport');
  const titleEl = $('sheetTitle');
  const subtitleEl = $('sheetSubtitle');
  const counterEl = $('counter');

  function txt(v){ return String(v ?? '').trim(); }

  function isCodeHeader(header){
    const h = normalizeHeader(header);
    return ['codigo','codigo producto','codigoproducto','sku','producto'].includes(h);
  }

  function cellText(header, value){
    let v = txt(value);
    if(isCodeHeader(header)){
      v = v.replace(/^['’`´]+/, '').replace(/\.0$/, '');
    }
    return v;
  }

  function normalizeHeader(value){
    return txt(value).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  }

  function jsonpListData(url){
    return new Promise((resolve,reject)=>{
      const cb='cb_sheet_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const script=document.createElement('script');
      const clean=()=>{try{delete window[cb];}catch(e){window[cb]=undefined;} if(script.parentNode)script.parentNode.removeChild(script); clearTimeout(timer);};
      const timer=setTimeout(()=>{clean(); reject(new Error('Sin respuesta del Apps Script'));},30000);
      window[cb]=(data)=>{clean(); resolve(data);};
      script.onerror=()=>{clean(); reject(new Error('No se pudo cargar Apps Script'));};
      script.src=url+'&callback='+encodeURIComponent(cb);
      document.head.appendChild(script);
    });
  }

  function updateInfo(message){
    if(syncInfo) syncInfo.textContent = message || '';
  }

  function setButtonLoading(btn, loading, text){
    if(!btn) return;
    if(loading){
      btn.dataset.originalText = btn.textContent;
      btn.disabled = true;
      btn.innerHTML = '<span class="mini-spinner"></span>' + (text || 'Sincronizando...');
    }else{
      btn.disabled = false;
      btn.textContent = btn.dataset.originalText || btn.textContent;
    }
  }

  async function loadApiConfig(){
    API_URL = DEFAULT_API_URL;
    try{
      localStorage.removeItem('API_URL_SGL');
      localStorage.removeItem('API_URL_PEDIDOS');
      localStorage.removeItem('API_URL_UBICACIONES');
      localStorage.removeItem('sistema_pedidos_api_url');
    }catch(err){}
  }

  function findIndex(possibleNames){
    const names = possibleNames.map(normalizeHeader);
    return headers.findIndex(h => names.includes(normalizeHeader(h)));
  }

  function parseDateValue(value){
    const raw = txt(value);
    if(!raw) return 0;
    let m = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})(?:[\s,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if(m){
      let year = String(m[3]);
      if(year.length === 2) year = '20' + year;
      const t = new Date(Number(year), Number(m[2])-1, Number(m[1]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0)).getTime();
      return Number.isNaN(t) ? 0 : t;
    }
    m = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[\s,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if(m){
      const t = new Date(Number(m[1]), Number(m[2])-1, Number(m[3]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0)).getTime();
      return Number.isNaN(t) ? 0 : t;
    }
    const t = new Date(raw).getTime();
    return Number.isNaN(t) ? 0 : t;
  }

  function getDateIndex(){
    return findIndex(['fecha','fecha_registro','fecha ingreso','fecha_pedido','fecha pedido','fecha_entrada','fecha salida','fecha_salida']);
  }

  function escapeHtml(value){
    return txt(value).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
  }

  function visibleRows(){
    const q = txt(searchInput && searchInput.value).toLowerCase();
    const status = txt(statusFilter && statusFilter.value).toLowerCase();
    const d1 = fechaDesde && fechaDesde.value ? new Date(fechaDesde.value + 'T00:00:00').getTime() : 0;
    const d2 = fechaHasta && fechaHasta.value ? new Date(fechaHasta.value + 'T23:59:59').getTime() : 0;
    const statusIdx = findIndex(['status','estado']);
    const dateIdx = getDateIndex();

    return rows.filter(row => {
      if(q){
        const joined = row.map(v => txt(v).toLowerCase()).join(' | ');
        if(!joined.includes(q)) return false;
      }
      if(status && statusIdx !== -1){
        const rowStatus = txt(row[statusIdx]).toLowerCase();
        if(rowStatus !== status) return false;
      }
      if((d1 || d2) && dateIdx !== -1){
        const t = parseDateValue(row[dateIdx]);
        if(d1 && (!t || t < d1)) return false;
        if(d2 && (!t || t > d2)) return false;
      }
      return true;
    });
  }

  function render(){
    const data = visibleRows();
    if(counterEl) counterEl.textContent = `${data.length} registro(s)`;
    if(!headers.length){
      tableHead.innerHTML = '<tr><th>Sin datos</th></tr>';
      tableBody.innerHTML = '<tr><td>No hay información para mostrar.</td></tr>';
      return;
    }
    tableHead.innerHTML = '<tr>' + headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr>';
    tableBody.innerHTML = data.map(row => {
      return '<tr>' + headers.map((_,idx) => `<td>${escapeHtml(cellText(headers[idx], row[idx] || ''))}</td>`).join('') + '</tr>';
    }).join('') || `<tr><td colspan="${headers.length}">No hay registros con los filtros seleccionados.</td></tr>`;
  }

  function fillStatusOptions(){
    if(!statusFilter) return;
    const idx = findIndex(['status','estado']);
    const current = statusFilter.value;
    const found = new Set();
    if(idx !== -1){
      rows.forEach(row => {
        const st = txt(row[idx]).toLowerCase();
        if(st) found.add(st);
      });
    }
    STATUS_VALUES.forEach(st => { if(found.has(st)) return; });
    const values = Array.from(new Set([...STATUS_VALUES.filter(st => found.has(st)), ...Array.from(found)])).sort();
    statusFilter.innerHTML = '<option value="">Todos los status</option>' + values.map(st => `<option value="${escapeHtml(st)}">${escapeHtml(st.toUpperCase())}</option>`).join('');
    statusFilter.value = current;
  }

  async function loadData(manual=false){
    if(isLoading) return;
    if(!API_URL){
      updateInfo('API_URL no configurada.');
      return;
    }
    isLoading = true;
    if(manual) setButtonLoading(btnSync, true, 'Sincronizando...');
    try{
      const url = `${API_URL}?accion=listar_bd&sheet=${encodeURIComponent(SHEET_NAME)}&device=${encodeURIComponent(DEVICE_ID)}&rt=${Date.now()}&nocache=${Math.random().toString(36).slice(2)}`;
      const data = await jsonpListData(url);
      if(!data || !data.ok){
        throw new Error((data && data.msg) || 'No se pudo leer la BD');
      }
      headers = Array.isArray(data.headers) ? data.headers.map(txt) : [];
      rows = Array.isArray(data.data) ? data.data : [];
      fillStatusOptions();
      render();
      updateInfo(`${manual ? 'Sincronización en tiempo real' : 'Actualización automática cada 2 segundos'} | ${SHEET_NAME}: ${rows.length} registro(s) | ${new Date().toLocaleTimeString()}`);
    }catch(err){
      console.error(err);
      updateInfo('No se pudo sincronizar la BD: ' + (err.message || err));
    }finally{
      if(manual) setButtonLoading(btnSync, false);
      isLoading = false;
    }
  }

  function exportXlsx(){
    const data = visibleRows();
    const aoa = [headers, ...data.map(row => headers.map((h, idx) => cellText(h, row[idx] || '')))];
    if(window.XLSX){
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      headers.forEach((h, c) => {
        if(isCodeHeader(h)){
          for(let r=1; r<=data.length; r++){
            const addr = XLSX.utils.encode_cell({r, c});
            if(ws[addr]){ ws[addr].t = 's'; ws[addr].z = '@'; }
          }
        }
      });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME.substring(0, 31));
      XLSX.writeFile(wb, `${SHEET_NAME}_${new Date().toISOString().slice(0,10)}.xlsx`);
      return;
    }
    const csvRows = aoa.map(row => row.map(v => '"' + txt(v).replace(/"/g,'""') + '"').join(';'));
    const blob = new Blob([csvRows.join('\n')], { type:'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${SHEET_NAME}_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a);
    a.click();
    URL.revokeObjectURL(a.href);
    a.remove();
  }

  function startAuto(){
    if(autoTimer) clearInterval(autoTimer);
    autoTimer = setInterval(() => loadData(false), AUTO_REFRESH_MS);
  }

  if(titleEl) titleEl.textContent = SHEET_TITLE;
  if(subtitleEl) subtitleEl.textContent = `${SHEET_NAME} | Actualización automática cada 2 segundos y sincronización manual en tiempo real`;
  btnSync && btnSync.addEventListener('click', () => loadData(true));
  btnClear && btnClear.addEventListener('click', () => {
    searchInput.value = '';
    statusFilter.value = '';
    fechaDesde.value = '';
    fechaHasta.value = '';
    render();
  });
  btnExport && btnExport.addEventListener('click', exportXlsx);
  searchInput && searchInput.addEventListener('input', render);
  statusFilter && statusFilter.addEventListener('change', render);
  fechaDesde && fechaDesde.addEventListener('change', render);
  fechaHasta && fechaHasta.addEventListener('change', render);
  document.addEventListener('visibilitychange', () => { if(!document.hidden) loadData(false); });

  (async function init(){
    await loadApiConfig();
    await loadData(true);
    startAuto();
  })();
})();
