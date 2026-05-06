/* =====================================================
   LOADER GLOBAL DE BOTONES
   Compatible con fetch, JSONP y acciones locales.
   También se usa manualmente desde orden_pedidos.js.
===================================================== */
(function(){
  if(window.ButtonLoader) return;
  const pendingByButton = new WeakMap();
  let lastClickedButton = null;
  let lastClickTime = 0;
  const MIN_VISIBLE_MS = 450;
  const FALLBACK_MS = 1200;
  const MAX_VISIBLE_MS = 20000;

  function canUseLoader(btn){
    return !!btn && !btn.classList.contains('no-loader') && btn.dataset.noLoader !== 'true' && !btn.disabled;
  }

  function start(btn){
    if(!canUseLoader(btn) && !(btn && btn.classList.contains('btn-loading'))) return;
    btn.dataset.loaderStartedAt = String(Date.now());
    btn.classList.add('btn-loading');
    btn.setAttribute('aria-busy','true');
    btn.disabled = true;
    clearTimeout(btn._loaderFallbackTimer);
    btn._loaderFallbackTimer = setTimeout(() => stop(btn, true), MAX_VISIBLE_MS);
  }

  function hasPending(btn){ return (pendingByButton.get(btn) || 0) > 0; }
  function addPending(btn){ if(btn) pendingByButton.set(btn, (pendingByButton.get(btn) || 0) + 1); }
  function removePending(btn){
    if(!btn) return;
    const next = Math.max(0, (pendingByButton.get(btn) || 0) - 1);
    pendingByButton.set(btn, next);
    if(next === 0) stop(btn);
  }

  function stop(btn, force){
    if(!btn) return;
    const startedAt = Number(btn.dataset.loaderStartedAt || Date.now());
    const elapsed = Date.now() - startedAt;
    const wait = force ? 0 : Math.max(0, MIN_VISIBLE_MS - elapsed);
    setTimeout(() => {
      if(!force && hasPending(btn)) return;
      btn.classList.remove('btn-loading');
      btn.removeAttribute('aria-busy');
      btn.disabled = false;
      delete btn.dataset.loaderStartedAt;
      clearTimeout(btn._loaderFallbackTimer);
    }, wait);
  }

  async function run(btn, fn){
    start(btn);
    try { return await fn(); }
    finally { stop(btn); }
  }

  document.addEventListener('click', function(e){
    const btn = e.target.closest('button');
    if(!btn || !canUseLoader(btn)) return;
    start(btn);
    lastClickedButton = btn;
    lastClickTime = Date.now();
    setTimeout(() => { if(!hasPending(btn)) stop(btn); }, FALLBACK_MS);
  }, true);

  if(window.fetch && !window.fetch.__loaderPatched){
    const originalFetch = window.fetch.bind(window);
    const patchedFetch = function(){
      const btn = (lastClickedButton && (Date.now() - lastClickTime) < 1500) ? lastClickedButton : null;
      if(btn){ start(btn); addPending(btn); }
      return originalFetch.apply(window, arguments).finally(() => { if(btn) removePending(btn); });
    };
    patchedFetch.__loaderPatched = true;
    window.fetch = patchedFetch;
  }

  window.ButtonLoader = { start, stop, run };
})();
