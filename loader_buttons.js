/* =====================================================
   LOADER GLOBAL PARA BOTONES
   - Aplica loader a todos los botones del sistema.
   - Mantiene el loader mientras haya llamadas fetch asociadas al botón.
   - Se detiene automáticamente para acciones locales rápidas.
===================================================== */
(function(){
  if(window.ButtonLoader) return;

  const pendingByButton = new WeakMap();
  let lastClickedButton = null;
  let lastClickTime = 0;
  let clearCandidateTimer = null;
  const MIN_VISIBLE_MS = 650;
  const FALLBACK_MS = 1200;
  const MAX_VISIBLE_MS = 12000;

  function canUseLoader(btn){
    if(!btn) return false;
    if(btn.classList.contains('no-loader')) return false;
    if(btn.dataset.noLoader === 'true') return false;
    if(btn.disabled) return false;
    return true;
  }

  function start(btn){
    if(!canUseLoader(btn)) return;
    if(!btn.dataset.loaderOriginalHtml){
      btn.dataset.loaderOriginalHtml = btn.innerHTML;
    }
    btn.dataset.loaderStartedAt = String(Date.now());
    btn.classList.add('btn-loading');
    btn.setAttribute('aria-busy','true');
    btn.disabled = true;

    clearTimeout(btn._loaderFallbackTimer);
    btn._loaderFallbackTimer = setTimeout(() => stop(btn, true), MAX_VISIBLE_MS);
  }

  function hasPending(btn){
    return (pendingByButton.get(btn) || 0) > 0;
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

  function addPending(btn){
    if(!btn) return;
    pendingByButton.set(btn, (pendingByButton.get(btn) || 0) + 1);
  }

  function removePending(btn){
    if(!btn) return;
    const next = Math.max(0, (pendingByButton.get(btn) || 0) - 1);
    pendingByButton.set(btn, next);
    if(next === 0) stop(btn);
  }

  document.addEventListener('click', function(e){
    const btn = e.target.closest('button');
    if(!btn) return;
    if(!canUseLoader(btn)) return;

    start(btn);
    lastClickedButton = btn;
    lastClickTime = Date.now();

    clearTimeout(clearCandidateTimer);
    clearCandidateTimer = setTimeout(() => {
      lastClickedButton = null;
    }, 450);

    // Para acciones que no usan fetch: abrir/cerrar modal, generar archivo local, etc.
    setTimeout(() => {
      if(!hasPending(btn)) stop(btn);
    }, FALLBACK_MS);
  }, true);

  // Asociar las llamadas fetch realizadas inmediatamente después de un click al botón activo.
  if(window.fetch && !window.fetch.__loaderPatched){
    const originalFetch = window.fetch.bind(window);
    const patchedFetch = function(){
      const btn = (lastClickedButton && (Date.now() - lastClickTime) < 1200) ? lastClickedButton : null;
      if(btn){
        start(btn);
        addPending(btn);
      }
      return originalFetch.apply(window, arguments)
        .then(resp => resp)
        .catch(err => { throw err; })
        .finally(() => {
          if(btn) removePending(btn);
        });
    };
    patchedFetch.__loaderPatched = true;
    window.fetch = patchedFetch;
  }

  window.ButtonLoader = { start, stop };
})();
