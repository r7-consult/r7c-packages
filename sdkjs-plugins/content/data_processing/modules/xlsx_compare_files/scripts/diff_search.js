/**
 * @fileoverview Diff search & navigation extracted from plugin.js
 */
(function(window){
 'use strict';
 // Exact match filtering: show only rows whose type/from/to EXACTLY equals query (case-insensitive)
 let _indices=[]; let _cursor=-1; let _open=false; let _lastQuery='';
 function els(){ return {
   toggle: document.getElementById('btnDiffSearchToggle'),
   fields: document.getElementById('diffSearchFields'),
   input: document.getElementById('diffSearchInput'),
   count: document.getElementById('diffSearchCount'),
   prev: document.getElementById('btnDiffSearchPrev'),
   next: document.getElementById('btnDiffSearchNext'),
   clear: document.getElementById('btnDiffSearchClear')
 }; }
 function open(){ const {fields,input,toggle}=els(); if(!fields) return; fields.classList.remove('hidden'); _open=true; if(toggle) toggle.setAttribute('aria-expanded','true'); if(input){ setTimeout(()=>input.focus(),15); input.select(); } }
 function close(){ const {fields,toggle}=els(); if(!fields) return; fields.classList.add('hidden'); _open=false; if(toggle) toggle.setAttribute('aria-expanded','false'); }
 function toggle(){ _open? close(): open(); }
function reset(){ _indices=[]; _cursor=-1; _lastQuery=''; const rows=_rows(); rows.forEach(r=>{ r.style.display=''; r.classList.remove('search-current'); if(r.classList.contains('active')){ r.classList.remove('active'); r.removeAttribute('aria-selected'); } }); updateCounter(); document.dispatchEvent(new CustomEvent('diffSearchReset')); }
// Clear preview focus when search reset triggered externally
document.addEventListener('diffSearchReset', ()=>{
  document.querySelectorAll('#basePreview td.preview-focus, #targetPreview td.preview-focus').forEach(td=> td.classList.remove('preview-focus'));
});
 function _rows(){ return [...document.querySelectorAll('#diffContainer .diff-row')].filter(r=>!r.classList.contains('diff-header') && !r.classList.contains('more')); }
 function updateCounter(){ const {count}=els(); if(count){ count.textContent= _indices.length? (_cursor+1)+'/'+_indices.length : '0/0'; } }
 function searchExact(q){ const query=(q||'').trim(); const rows=_rows(); rows.forEach(r=>{ r.style.display=''; r.classList.remove('search-current'); });
   if(!query){ reset(); return; }
   const lower=query.toLowerCase(); const st=State.get(); if(!st.diff || !st.diff.raw){ reset(); return; }
   function colName(c){ let s=''; c++; while(c>0){ let m=(c-1)%26; s=String.fromCharCode(65+m)+s; c=Math.floor((c-1)/26); } return s; }
   _indices=[]; (st.diff.raw||[]).forEach((d,i)=>{ if(!d) return; const t=(d.type||'').toString().toLowerCase(); const from=(d.from==null?'':d.from).toString().toLowerCase(); const to=(d.to==null?'':d.to).toString().toLowerCase(); let addr=''; if(d.r!=null && d.c!=null){ addr=(colName(d.c)+(d.r+1)).toLowerCase(); }
     if(t===lower || from===lower || to===lower || (addr && addr===lower)){ _indices.push(i); } });
  // Do NOT hide or bulk-highlight matches; only the currently focused one will get 'search-current'.
  rows.forEach(r=>{ r.classList.remove('search-match'); });
   _cursor=_indices.length?0:-1; _lastQuery=query; updateCounter(); focusCurrent(true); }
 function focusCurrent(scroll){ if(_cursor<0){ updateCounter(); return; } const rows=_rows(); rows.forEach(r=>{ r.classList.remove('search-current'); r.classList.remove('active'); }); const targetIndex=_indices[_cursor]; const row=rows.find(r=> parseInt(r.getAttribute('data-diff-index'),10)===targetIndex); if(!row) return; row.classList.add('search-current');
  // Ensure parent diff group (if any) is expanded so the row is visible
  try {
    const section = row.closest('.diff-group');
    if(section && section.classList.contains('collapsed')){
      section.classList.remove('collapsed');
      const toggleEl = section.querySelector('.diff-group-header .dg-toggle');
      if(toggleEl){ toggleEl.textContent='▾'; toggleEl.setAttribute('data-open','1'); }
    }
  } catch(_e){}
   // also mark as active diff row (reuse plugin logic for preview focus) – ensure exclusivity
   if(typeof window.activateDiffRow === 'function'){ try { window.activateDiffRow(row); } catch(e){} }
   else if(row && row.click){ row.classList.add('active'); }
   if(scroll){ const sc=document.querySelector('#diffContainer .diff-scroll'); if(sc){ const top=row.offsetTop; sc.scrollTop=Math.max(0, top-60); } } }
 function next(){ if(!_indices.length) return; _cursor=( _cursor+1 ) % _indices.length; focusCurrent(true); updateCounter(); }
 function prev(){ if(!_indices.length) return; _cursor=( _cursor-1+_indices.length ) % _indices.length; focusCurrent(true); updateCounter(); }
function clearSearch(hide=true){ const {input}=els(); if(input) input.value=''; reset(); if(hide) close(); }
 function wire(){ const {toggle:tg,input,prev:pb,next:nb,clear:cb}=els(); if(tg && !tg._wired){ tg._wired=true; tg.addEventListener('click', toggle); tg.setAttribute('aria-expanded','false'); }
  if(input && !input._wired){ input._wired=true; input.addEventListener('keydown', e=>{
      if(e.key==='Enter'){
        const val=input.value.trim();
        if(val && val.toLowerCase()!==_lastQuery.toLowerCase()){
          searchExact(val);
        } else {
          if(e.shiftKey) prev(); else next();
        }
        e.preventDefault();
      } else if(e.key==='ArrowDown') { next(); e.preventDefault(); }
        else if(e.key==='ArrowUp'){ prev(); e.preventDefault(); }
        else if(e.key==='Escape'){ if(input.value){ clearSearch(false); } else { close(); } }
    }); }
   if(pb && !pb._wired){ pb._wired=true; pb.addEventListener('click', prev); }
   if(nb && !nb._wired){ nb._wired=true; nb.addEventListener('click', next); }
   if(cb && !cb._wired){ cb._wired=true; cb.addEventListener('click', ()=> clearSearch(true)); }
   document.addEventListener('keydown', e=>{ if(!_open) return; if(e.key==='F3'){ e.preventDefault(); e.shiftKey? prev(): next(); } });
 }
 State.subscribe(st=>{ if(!st.diff){ reset(); return; } if(_lastQuery){ // reapply current filter after diff change
     const {input}=els(); if(input && input.value.trim()) searchExact(input.value); }
 });
 window.DiffSearch={ searchExact, next, prev, reset, open, close, toggle, clearSearch };
 document.addEventListener('DOMContentLoaded', wire);
})(window);
