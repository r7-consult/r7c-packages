(function(window){
 'use strict';
 if(window.LegendModule) return;
 function _safeKey(s){ return String(s==null? '' : s).trim().replace(/[^A-Za-z0-9_]/g,'_'); }
 function allBoxes(){ return [...document.querySelectorAll('.cat-filter')]; }
 function computeState(){ const boxes=allBoxes(); const active=boxes.filter(b=>b.checked).map(b=>b.getAttribute('data-cat')); return (active.length===0||active.length===boxes.length)? null: active; }

 function applyStateFromUI(){ const normalized=computeState(); if(window.State) State.set({ enabledCategories: normalized }); if(window.PluginActions && window.PluginActions.refreshLegend) window.PluginActions.refreshLegend(); }
 function selectAll(){ allBoxes().forEach(b=>b.checked=true); if(window.State) State.set({ enabledCategories:null }); if(window.PluginActions && window.PluginActions.refreshLegend) window.PluginActions.refreshLegend(); }
 function clearAll(){ allBoxes().forEach(b=>b.checked=false); if(window.State) State.set({ enabledCategories:null }); if(window.PluginActions && window.PluginActions.refreshLegend) window.PluginActions.refreshLegend(); }
 function reset(){ 
   allBoxes().forEach(b=>b.checked=true); 
   if(window.State) State.set({ enabledCategories:null }); 
   if(window.PluginActions && window.PluginActions.resetLegend) window.PluginActions.resetLegend(); 
   if(window.PluginActions && window.PluginActions.refreshLegend) window.PluginActions.refreshLegend(); 
 }
 function updateCounts(counts){ try { document.querySelectorAll('#legendSection .cat-count').forEach(sp=>{ const cat=sp.getAttribute('data-count'); const v=counts && counts[cat]; if(v>0){ sp.textContent=v; sp.style.display='inline'; } else { sp.textContent=''; sp.style.display='none'; } }); } catch(_e){} }
 function resetCounts(){ document.querySelectorAll('#legendSection .cat-count').forEach(sp=>{ sp.textContent=''; sp.style.display='none'; }); dimUnavailable(null); }
 function dimUnavailable(activeSet){ document.querySelectorAll('#legendSection li').forEach(li=>{ if(!activeSet||!activeSet.size){ li.classList.remove('legend-dim'); return; } const t=li.getAttribute('data-difftype'); if(activeSet.has(t)) li.classList.remove('legend-dim'); else li.classList.add('legend-dim'); }); }
 function handleChange(e){ if(e.target && e.target.classList && e.target.classList.contains('cat-filter')) applyStateFromUI(); }
 function handleClick(e){ const btn=e.target.closest && e.target.closest('button'); if(!btn) return; switch(btn.id){ case 'btnCatSelectAll': e.preventDefault(); selectAll(); break; case 'btnCatClearAll': e.preventDefault(); clearAll(); break; case 'btnCatReset': e.preventDefault(); reset(); break; default: break; } }
 function init(){ 
  try {
    const actions=document.getElementById('legendActions');
    if(actions && !actions.getAttribute('data-testid')) actions.setAttribute('data-testid','xlsx-legend-actions');
    const btnSel=document.getElementById('btnCatSelectAll');
    const btnClr=document.getElementById('btnCatClearAll');
    const btnReset=document.getElementById('btnCatReset');
    if(btnSel && !btnSel.getAttribute('data-testid')) btnSel.setAttribute('data-testid','xlsx-legend-select-all');
    if(btnClr && !btnClr.getAttribute('data-testid')) btnClr.setAttribute('data-testid','xlsx-legend-clear-all');
    if(btnReset && !btnReset.getAttribute('data-testid')) btnReset.setAttribute('data-testid','xlsx-legend-reset');
    allBoxes().forEach((cb, idx)=>{
      if(!cb || cb.getAttribute('data-testid')) return;
      const cat=_safeKey(cb.getAttribute('data-cat') || cb.value || idx);
      cb.setAttribute('data-testid','xlsx-legend-filter-'+cat+'-'+idx);
    });
    document.querySelectorAll('#legendSection .cat-count').forEach((sp, idx)=>{
      if(!sp || sp.getAttribute('data-testid')) return;
      const cat=_safeKey(sp.getAttribute('data-count') || idx);
      sp.setAttribute('data-testid','xlsx-legend-count-'+cat+'-'+idx);
    });
  } catch(_e){}
  document.addEventListener('change', handleChange); 
  document.getElementById('legendActions')?.addEventListener('click', handleClick);
 }
 window.LegendModule={ init, updateCounts, resetCounts, dimUnavailable, selectAll, clearAll, reset };
 if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', init); else init();
})(window);