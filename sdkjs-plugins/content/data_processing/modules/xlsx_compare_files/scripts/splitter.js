(function(window){
 'use strict';
 function initHorizontal(){
   const split=document.getElementById('horizontalSplitter');
   const previews=document.getElementById('regionPreviews');
   const diff=document.getElementById('regionDiff');
   if(!split||!previews||!diff) return;
   try {
     if(!split.getAttribute('data-testid')) split.setAttribute('data-testid','xlsx-splitter-horizontal');
     if(!previews.getAttribute('data-testid')) previews.setAttribute('data-testid','xlsx-region-previews');
     if(!diff.getAttribute('data-testid')) diff.setAttribute('data-testid','xlsx-region-diff');
     const status=document.getElementById('regionStatus');
     if(status && !status.getAttribute('data-testid')) status.setAttribute('data-testid','xlsx-region-status');
   } catch(_e){}
   let startY=0; let startH=0;
  const STORE_KEY='compareSheets_splitterH_v1';
   function bounds(){
     const total = window.innerHeight;
     const statusH = document.getElementById('regionStatus')?.getBoundingClientRect().height || 60;
     const splitterH = split.getBoundingClientRect().height || 6;
     const minDiff = 160; // нижний предел для diff
     // динамический нижний предел для превью (10 строк или реальное число если <10)
     const dyn = window.__previewMinRowsHeight || 120;
     const minPrev = dyn;
     const available = total - statusH - splitterH - 20; // padding safety
     const maxPrev = available - minDiff; // максимум превью чтобы diff сохранил минимум
     const maxClamped = Math.max(minPrev, maxPrev);
     return { minPrev, maxPrev:maxClamped };
   }
   function setH(px){
     const {minPrev,maxPrev}=bounds();
     // clamp respecting dynamic baseline (>=10 rows -> baseline10, <10 rows -> exact small)
     let h=Math.min(maxPrev, Math.max(minPrev, px));
     // Snap very close drags to the baseline to avoid 1-2px off-by jitter hiding scrollbar
     if(Math.abs(h-minPrev)<4) h=minPrev;
     previews.style.flex='0 0 '+h+'px';
     previews.style.height=h+'px';
     // Let diff flex-fill; keep a stable functional minimum only
     diff.style.minHeight='160px';
     adaptPreviewRowCount(h);
    try {
      // Recompute scroll viewport inside each grid to keep vertical scrollbar consistent
      const regionPreviews = document.getElementById('regionPreviews');
      if(regionPreviews){
        const cols = regionPreviews.querySelectorAll('.sheet-preview-col');
        cols.forEach(col=>{
          const title = col.querySelector('.preview-title');
          const grid = col.querySelector('.sheet-grid');
          if(!grid) return;
          const tbl = grid.querySelector('table');
          const bodyRows = tbl && tbl.tBodies && tbl.tBodies[0] ? tbl.tBodies[0].rows.length : 0;
          if(bodyRows <= 10){
            grid.style.height='auto';
            grid.style.maxHeight='100%';
            grid.style.overflowY='auto';
            return;
          }
          const titleHNow = title? (title.getBoundingClientRect().height||18):18;
          const scrollbarGuess = (navigator.platform && /win/i.test(navigator.platform)) ? 18 : 16;
          const inner = Math.max(60, h - titleHNow - 6 - scrollbarGuess);
          grid.style.height=inner+'px';
          grid.style.maxHeight=inner+'px';
          grid.style.overflowY='auto';
          grid.style.overflowX='auto';
        });
      }
    } catch(_e){}
  // Splitter now lives as its own flex item between previews and diff; no re-parenting needed
  split.classList.add('attached');
     try { localStorage.setItem(STORE_KEY, String(h)); } catch(_e){}
   }
   function onDown(e){ startY=e.clientY; startH=previews.getBoundingClientRect().height; document.addEventListener('mousemove',onMove); document.addEventListener('mouseup',onUp); split.classList.add('dragging'); previews.classList.add('resizing'); diff.classList.add('resizing'); e.preventDefault(); }
   function onMove(e){ const dy=e.clientY-startY; setH(startH+dy); }
  function onUp(){ document.removeEventListener('mousemove',onMove); document.removeEventListener('mouseup',onUp); split.classList.remove('dragging'); previews.classList.remove('resizing'); diff.classList.remove('resizing'); try { window.__splitterUserResized = true; } catch(_e){} }
   split.addEventListener('mousedown', onDown);
  split.addEventListener('keydown', e=>{ if(e.key==='ArrowUp'){ setH(previews.getBoundingClientRect().height-24); } else if(e.key==='ArrowDown'){ setH(previews.getBoundingClientRect().height+24); } });
  // Recompute bounds on window resize
  window.addEventListener('resize', ()=>{ setH(previews.getBoundingClientRect().height); });

  // expose applyStored so preview renderer can re-evaluate after dynamic min changes
  function applyStored(){
     try {
       if(!window.__splitterUserResized) return; // only honor after explicit user adjustment in session
       const raw=localStorage.getItem(STORE_KEY);
       if(!raw) return; const val=parseInt(raw,10); if(isNaN(val)) return; setH(val);
     } catch(_e){}
  }
  window.__splitterApplyStoredHeight = applyStored;

  // initial apply (after small timeout to let preview min propagate if already set)
  setTimeout(()=>{ applyStored(); }, 30);
 }
 // Increase/decrease visible rows based on height (approx row height 18px after removing header + controls row)
 function adaptPreviewRowCount(height){ try { const tables=document.querySelectorAll('#basePreview table, #targetPreview table'); if(!tables.length) return; const headerOffset=38; const avail=height-headerOffset; const approxRow=18; const rowsVisible=Math.max(5, Math.floor(avail/approxRow)); tables.forEach(tbl=>{ tbl.dataset.visibleRows=rowsVisible; });
 } catch(_e){}
 }
 document.addEventListener('DOMContentLoaded', ()=>{ initHorizontal(); });
 window.__splitterReady=true;
})(window);
