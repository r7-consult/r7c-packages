/** Plugin entry */
(function(window){
 'use strict';
let diffEngine=null; let cancelFlag=false; let legendDynamic=true; // legacy flag (now selection-driven)
const JSON_SCHEMA_VERSION = '1.0.0';
function _attemptRestoreSelections(){
  const st=State.get();
  if(st.base && st.target) return; // nothing to do
  // Heuristic: pick first virtual sheet whose name hints side, else any
  const allVS = window.VirtualSheets? Object.keys(window.VirtualSheets): [];
  if(!st.base){
    const cand = allVS.find(k=>/Imported_Base_/i.test(k)) || allVS[0];
    if(cand) State.set({ base:cand });
  }
  if(!st.target){
    const cand = allVS.find(k=>/Imported_Target_/i.test(k) && k!==State.get().base) || allVS.find(k=>k!==State.get().base);
    if(cand) State.set({ target:cand });
  }
}
// Build quick row-level classification from diffs
function _classify(diff){ const rows={}, cells={}; if(!diff||!diff.raw) return {rows,cells}; diff.raw.forEach(d=>{ if(d.r==null || d.c==null) return; const t=d.type; if(!rows[d.r]) rows[d.r]=new Set(); rows[d.r].add(t); const key=d.r+':'+d.c; if(!cells[key]) cells[key]=new Set(); cells[key].add(t); }); return {rows,cells}; }
function _rowCssAll(set){ if(!set||!set.size) return []; const order=['inserted-row','deleted-row','inserted-col','deleted-col','formula','formulaToValue','valueToFormula','value','type','format']; const out=[]; order.forEach(k=>{ if(set.has(k)) out.push('diff-row-'+k.replace(/[^A-Za-z0-9]/g,'')); }); return out; }
function _cellCssAll(set){ if(!set||!set.size) return []; // Exclude structural categories from direct cell coloring
  const order=['formula','formulaToValue','valueToFormula','value','type','format']; const out=[]; order.forEach(k=>{ if(set.has(k)) out.push('diff-cell-'+k.replace(/[^A-Za-z0-9]/g,'')); }); return out; }
async function _renderPreviews(baseName,targetName,diff){
  // Reset splitter user-resize flag for a fresh preview build (new run shows only baseline 10 rows)
  try { window.__splitterUserResized = false; } catch(_e){}
  if(window.PreviewRenderer && PreviewRenderer.buildPreviews){ try { await PreviewRenderer.buildPreviews(baseName,targetName,diff,diffEngine); } catch(e){ Logger.error('Preview build failed', e); } }
  // Dynamic preview titles (workbook | sheet) to match diff header naming
  try {
    const st=(State && State.get)? State.get(): {};
    const ph1=(window.__i18n && window.__i18n['placeholder.source1'])||'Src1';
    const ph2=(window.__i18n && window.__i18n['placeholder.source2'])||'Src2';
    function stripExt(n){ return (n||'').replace(/\.[^.]+$/,''); }
    const meta = st.diff && st.diff.meta || {};
    const baseWB = stripExt(st.baseOriginalFile || meta.baseWorkbook || meta.base || baseName || ph1);
    const targetWB = stripExt(st.targetOriginalFile || meta.targetWorkbook || meta.target || targetName || ph2);
    const baseSheet = document.getElementById('baseSheetSelect')?.value || meta.baseSheet || meta.baseSheetName || '';
    const targetSheet = document.getElementById('targetSheetSelect')?.value || meta.targetSheet || meta.targetSheetName || '';
    const baseLabel = baseWB + (baseSheet? ' | '+baseSheet : '');
    const targetLabel = targetWB + (targetSheet? ' | '+targetSheet : '');
    const baseTitleEl=document.querySelector('#sheetPreviewWrapper .sheet-preview-col:first-child .preview-title');
    const targetTitleEl=document.querySelector('#sheetPreviewWrapper .sheet-preview-col:last-child .preview-title');
    if(baseTitleEl) baseTitleEl.textContent = baseLabel;
    if(targetTitleEl) targetTitleEl.textContent = targetLabel;
  } catch(_e){}
  // Advanced sync scroll with diff-based row & column segment mapping (anchor follow removed)
  try { const baseDiv=document.getElementById('basePreview'); const targetDiv=document.getElementById('targetPreview'); const syncCheckbox=document.getElementById('chkSyncScroll'); if(!baseDiv||!targetDiv||!syncCheckbox) return; const isSyncEnabled=()=> syncCheckbox.checked; let rafPending=false; let syncing=false; let activeSource=null;
    const diffObj = State.get().diff; const diffRaw = diffObj && diffObj.raw || [];
    // Collect structural changes
    const insertedRowsTarget=new Set(); const deletedRowsBase=new Set();
    const insertedColsTarget=new Set(); const deletedColsBase=new Set();
    diffRaw.forEach(d=>{ if(!d) return; switch(d.type){ case 'inserted-row': insertedRowsTarget.add(d.r); break; case 'deleted-row': deletedRowsBase.add(d.r); break; case 'inserted-col': insertedColsTarget.add(d.c); break; case 'deleted-col': deletedColsBase.add(d.c); break; default: break; }});
    // Row segments (aligned contiguous blocks)
    function buildRowSegments(){ const baseRows=baseDiv.querySelectorAll('tbody tr').length; const targetRows=targetDiv.querySelectorAll('tbody tr').length; const segments=[]; let t=0; let segStartB=null, segStartT=null; function flush(b){ if(segStartB!=null){ segments.push({ baseStart:segStartB, baseEnd:b, targetStart:segStartT, targetEnd:t }); segStartB=null; segStartT=null; } }
      for(let b=0;b<baseRows;b++){ if(deletedRowsBase.has(b)){ flush(b); continue; } while(t<targetRows && insertedRowsTarget.has(t)){ flush(b); t++; } if(t>=targetRows){ flush(b); break; } if(segStartB==null){ segStartB=b; segStartT=t; } t++; } flush(baseRows); return { segments, baseRows, targetRows }; }
    const rowMap=buildRowSegments();
    function locateSegByBaseRow(r){ return rowMap.segments.find(s=> r>=s.baseStart && r < s.baseEnd); }
    function locateSegByTargetRow(r){ return rowMap.segments.find(s=> r>=s.targetStart && r < s.targetEnd); }
    // Column offset helpers (approx, fixed width assumption)
    const DEFAULT_W = 120; // keep in sync with preview_renderer
    function buildColShiftArrays(){ const baseCols=(baseDiv.querySelector('thead tr')?.cells.length||1)-1; const targetCols=(targetDiv.querySelector('thead tr')?.cells.length||1)-1; const baseDelShift=[]; let del=0; for(let c=0;c<baseCols;c++){ if(deletedColsBase.has(c)) del++; baseDelShift[c]=del; } const targetInsShift=[]; let ins=0; for(let c=0;c<targetCols;c++){ if(insertedColsTarget.has(c)) ins++; targetInsShift[c]=ins; } return { baseCols,targetCols, baseDelShift,targetInsShift }; }
    const colMap=buildColShiftArrays();
    function mapBaseLeft(x){ const approx=Math.floor(x/DEFAULT_W); const delShift=colMap.baseDelShift[Math.min(approx,colMap.baseDelShift.length-1)]||0; const logical=approx - delShift; const insShift=colMap.targetInsShift[Math.min(logical,colMap.targetInsShift.length-1)]||0; const targetVisual=logical + insShift; return targetVisual*DEFAULT_W + (x%DEFAULT_W); }
    function mapTargetLeft(x){ const approx=Math.floor(x/DEFAULT_W); const insShift=colMap.targetInsShift[Math.min(approx,colMap.targetInsShift.length-1)]||0; const logical=approx - insShift; const delShift=colMap.baseDelShift[Math.min(logical,colMap.baseDelShift.length-1)]||0; const baseVisual=logical + delShift; return baseVisual*DEFAULT_W + (x%DEFAULT_W); }
    // Vertical mapping via segments
    function mapBaseTop(y){
      const rowH=(baseDiv.querySelector('tbody tr')?.offsetHeight)||18; if(rowH<=0) return y;
      const r=Math.floor(y/rowH); const intra=y - r*rowH; const seg=locateSegByBaseRow(r);
      if(!seg){ // proportional fallback with precise fraction
        const maxB=Math.max(1, baseDiv.scrollHeight-baseDiv.clientHeight);
        const maxT=Math.max(1, targetDiv.scrollHeight-targetDiv.clientHeight);
        const frac = y / maxB; return frac * maxT;
      }
      const offset=r - seg.baseStart; const tRow = seg.targetStart + offset;
      const tRowH=(targetDiv.querySelector('tbody tr')?.offsetHeight)||18;
      const mapped = tRow * tRowH + Math.min(intra, tRowH-1);
      return mapped;
    }
    function mapTargetTop(y){
      const rowH=(targetDiv.querySelector('tbody tr')?.offsetHeight)||18; if(rowH<=0) return y;
      const r=Math.floor(y/rowH); const intra=y - r*rowH; const seg=locateSegByTargetRow(r);
      if(!seg){ const maxT=Math.max(1, targetDiv.scrollHeight-targetDiv.clientHeight); const maxB=Math.max(1, baseDiv.scrollHeight-baseDiv.clientHeight); const frac=y/maxT; return frac*maxB; }
      const offset=r - seg.targetStart; const bRow=seg.baseStart + offset; const bRowH=(baseDiv.querySelector('tbody tr')?.offsetHeight)||18; const mapped=bRow * bRowH + Math.min(intra, bRowH-1); return mapped; }
  function schedule(from){ if(!isSyncEnabled()||syncing||!from||rafPending) return; rafPending=true; requestAnimationFrame(()=>{ rafPending=false; if(!isSyncEnabled()) return; const other=from===baseDiv?targetDiv:baseDiv; syncing=true; other._ignoreNextScroll=true; if(from===baseDiv){ other.scrollTop=mapBaseTop(from.scrollTop); other.scrollLeft=mapBaseLeft(from.scrollLeft); } else { other.scrollTop=mapTargetTop(from.scrollTop); other.scrollLeft=mapTargetLeft(from.scrollLeft); } syncing=false; }); }
  function markActive(div){ if(!isSyncEnabled()) return; activeSource=div; }
  [baseDiv,targetDiv].forEach(div=>{ ['pointerdown','mousedown','wheel','touchstart'].forEach(ev=> div.addEventListener(ev, ()=>markActive(div), { passive:true })); });
  function onScroll(div){ if(!isSyncEnabled()) return; if(div._ignoreNextScroll){ div._ignoreNextScroll=false; return; } if(activeSource && activeSource!==div) return; if(!syncing) schedule(div); }
  if(!syncCheckbox._bound){ syncCheckbox._bound=true; syncCheckbox.addEventListener('change',()=>{ if(syncCheckbox.checked){ // Choose last active pane (user's recent scroll) as source for initial alignment
    const src = activeSource && (activeSource===baseDiv||activeSource===targetDiv)? activeSource : baseDiv;
    const dst = (src===baseDiv)? targetDiv : baseDiv;
    if(src===baseDiv){ dst.scrollLeft=mapBaseLeft(src.scrollLeft); dst.scrollTop=mapBaseTop(src.scrollTop); }
    else { dst.scrollLeft=mapTargetLeft(src.scrollLeft); dst.scrollTop=mapTargetTop(src.scrollTop); }
    schedule(src); }
    else { clearSharedHover(); }}); }
  if(!baseDiv._syncScrollBound){ baseDiv.addEventListener('scroll',()=>onScroll(baseDiv), { passive:true }); baseDiv._syncScrollBound=true; }
  if(!targetDiv._syncScrollBound){ targetDiv.addEventListener('scroll',()=>onScroll(targetDiv), { passive:true }); targetDiv._syncScrollBound=true; }
  // === Shared row hover highlight (mirrors row across previews) ===
  const HOVER_CLASS='preview-hover-shared';
  function clearSharedHover(){ [baseDiv,targetDiv].forEach(div=>{ div.querySelectorAll('tbody tr.'+HOVER_CLASS).forEach(r=>r.classList.remove(HOVER_CLASS)); }); }
  function applySharedHover(fromDiv, rowIdx){
    if(!isSyncEnabled()) return;
    if(rowIdx==null||rowIdx<0){ clearSharedHover(); return; }
    const baseRowsTotal = baseDiv.querySelectorAll('tbody tr').length;
    const targetRowsTotal = targetDiv.querySelectorAll('tbody tr').length;
    let baseRowIdx=null, targetRowIdx=null;
    if(fromDiv===baseDiv){
      baseRowIdx=rowIdx;
      const seg=locateSegByBaseRow(rowIdx);
      if(seg){
        targetRowIdx = seg.targetStart + (rowIdx - seg.baseStart);
      } else {
        // Fallback: proportional mapping (handles trailing inserted rows in target)
        if(baseRowsTotal>1 && targetRowsTotal>0){
          const frac = rowIdx / (baseRowsTotal-1);
          targetRowIdx = Math.min(targetRowsTotal-1, Math.max(0, Math.round(frac * (targetRowsTotal-1))));
        } else if(targetRowsTotal){
          targetRowIdx = 0;
        }
      }
    } else {
      targetRowIdx=rowIdx;
      const seg=locateSegByTargetRow(rowIdx);
      if(seg){
        baseRowIdx = seg.baseStart + (rowIdx - seg.targetStart);
      } else {
        // Fallback: proportional mapping for rows existing only on target side
        if(targetRowsTotal>1 && baseRowsTotal>0){
          const frac = rowIdx / (targetRowsTotal-1);
          baseRowIdx = Math.min(baseRowsTotal-1, Math.max(0, Math.round(frac * (baseRowsTotal-1))));
        } else if(baseRowsTotal){
          baseRowIdx = 0;
        }
      }
    }
    clearSharedHover();
    if(baseRowIdx!=null){ const rB = baseDiv.querySelectorAll('tbody tr')[baseRowIdx]; if(rB) rB.classList.add(HOVER_CLASS); }
    if(targetRowIdx!=null){ const rT = targetDiv.querySelectorAll('tbody tr')[targetRowIdx]; if(rT) rT.classList.add(HOVER_CLASS); }
  }
  function bindHover(div){ if(div._sharedHoverBound) return; div._sharedHoverBound=true; let last=-1; div.addEventListener('mousemove', e=>{ if(!isSyncEnabled()) return; const tr=e.target.closest('tbody tr'); if(!tr){ if(last!==-1){ last=-1; clearSharedHover(); } return; } const rows=[...div.querySelectorAll('tbody tr')]; const idx=rows.indexOf(tr); if(idx===-1 || idx===last) return; last=idx; applySharedHover(div, idx); }, { passive:true }); div.addEventListener('mouseleave', ()=>{ if(last!==-1){ last=-1; clearSharedHover(); } }, { passive:true }); }
  bindHover(baseDiv); bindHover(targetDiv);
  } catch(_e){}
}
 function attachAscHandlers(){
   if(!window.Asc || !window.Asc.plugin){ setTimeout(attachAscHandlers,50); return; }
   const asc = window.Asc.plugin;
   if(asc.__compareTablesHandlersAttached) return; // idempotent
   asc.__compareTablesHandlersAttached = true;
   asc.init = function(){ Logger.info('Plugin init'); UI.init(); };
   asc.button = function(id){ if(id===-1){ Logger.info('Close button'); } };
   Logger.info('Asc handlers attached');
 }
 attachAscHandlers();
 async function startCompare(){ let st=State.get(); if(st.running) return; if(!st.base||!st.target){ _attemptRestoreSelections(); st=State.get(); }
  if(st.importDirtyForRun){
  const msg = (document.querySelector('[data-i18n="msg.importBlocked"]')?.textContent) || (window.__i18n && window.__i18n['msg.importBlocked']) || 'Перед запуском очистите импорт';
    State.set({ statusMessage: msg, statusKind:'warn' });
    if(window.PluginActions && window.PluginActions._pushUserMessage) window.PluginActions._pushUserMessage('warn', msg);
    return;
  }
  if(!st.base||!st.target){
    const msgEl=document.querySelector('[data-i18n="status.missingSelection"]');
    State.set({ statusMessage: msgEl? msgEl.textContent: 'Выберите базовый и целевой листы', statusKind:'warn' });
    return;
  }
  cancelFlag=false; legendDynamic=true; const opts={...st.options, enabledCategories: st.enabledCategories};
  State.set({ running:true, progress:0, diff:null, statusMessage:null, statusKind:null });
    // Lock further uploads (until user clears imports) once a comparison run is initiated with both sides present
    try { if(st.base && st.target){ State.set({ lockUploads:true }); } } catch(_e){}
  const btnCancel=document.getElementById('btnCancel'); if(btnCancel){ btnCancel.style.display='inline-block'; btnCancel.disabled=false; }
  clearMessages(); renderDiff(null); filterLegend(null); Logger.info('Compare start', st.base, st.target);
  // show previews before diff processing (previous diff for highlighting if exists)
  await _renderPreviews(st.base, st.target, st.diff);
   try {
  // Select provider (JS or Python) based on size / config
  let provider; try { provider = await (window.DiffProviders? DiffProviders.select(st.base, st.target, opts): Promise.resolve(new DiffEngine(opts))); } catch(e){ provider = new DiffEngine(opts); }
  diffEngine = (provider instanceof DiffEngine)? provider : null;
  const diff = await provider.compare(st.base, st.target, (p)=>{ if(cancelFlag && diffEngine && diffEngine.cancel){ diffEngine.cancel(); } else State.set({ progress:p }); });
  if(diff.canceled){ const msgCanceled=document.querySelector('[data-i18n="msg.canceled"]')?.textContent || 'Сравнение отменено'; Logger.warn('Comparison canceled'); pushMessage('warn',msgCanceled); State.set({ running:false, progress:0 }); return; }
  State.set({ diff, running:false, progress:1 });
    if(btnCancel){ btnCancel.style.display='none'; }
    // Ensure counts visible again (may have been hidden by reset)
    document.querySelectorAll('.cat-count').forEach(sp=>{ sp.style.display='inline-block'; });
  // Re-render previews with latest diff to update highlighting
  await _renderPreviews(st.base, st.target, diff);
  // Capture actual rendered (trimmed) preview sizes for delta metrics (exclude trailing empty rows/cols)
  try {
    const baseTable = document.querySelector('#basePreview table');
    const targetTable = document.querySelector('#targetPreview table');
    if(baseTable && targetTable){
      const rowsBase = (baseTable.tBodies[0] && baseTable.tBodies[0].rows.length) || 0;
      const rowsTarget = (targetTable.tBodies[0] && targetTable.tBodies[0].rows.length) || 0;
      const colsBase = ((baseTable.tHead && baseTable.tHead.rows[0] && baseTable.tHead.rows[0].cells.length)||1) - 1; // minus row number column
      const colsTarget = ((targetTable.tHead && targetTable.tHead.rows[0] && targetTable.tHead.rows[0].cells.length)||1) - 1;
      const m = diff.meta.metrics || (diff.meta.metrics = {});
      m.previewBaseRows = rowsBase; m.previewTargetRows = rowsTarget;
      m.previewBaseCols = colsBase; m.previewTargetCols = colsTarget;
    }
  } catch(_e){}
  renderDiff(diff); filterLegend(diff);
     Logger.info('Compare complete', diff.stats);
     // Diagnostics message (localized delta summary)
     try {
       if(window.PluginActions && window.PluginActions._pushUserMessage){
         const m=diff.meta.metrics||{}; const t=diff.meta.timings||{};
         const trD=(k)=> (window.__i18n && window.__i18n[k]) || k;
         const density = m.diffDensity!=null? (m.diffDensity*100).toFixed(1)+'%':'?';
         const timeMs = t.totalMs || diff.meta.durationMs || 0;
         const unitTime = (document.documentElement && /ru/i.test(document.documentElement.lang))? 'мс':'ms';
         const rBase = m.previewBaseRows!=null? m.previewBaseRows : m.baseRows;
         const rTarget = m.previewTargetRows!=null? m.previewTargetRows : m.targetRows;
         const cBase = m.previewBaseCols!=null? m.previewBaseCols : m.baseCols;
         const cTarget = m.previewTargetCols!=null? m.previewTargetCols : m.targetCols;
         const parts=[
           `${trD('delta.changedCells')} ${m.changedCells||0}`,
           `${trD('delta.structuralDiffs')} ${m.structuralDiffs||0}`,
           `${trD('delta.diffPercent')} ${density}`,
           `${trD('delta.rows')} ${rBase!=null?rBase:'?'}→${rTarget!=null?rTarget:'?'}`,
           `${trD('delta.cols')} ${cBase!=null?cBase:'?'}→${cTarget!=null?cTarget:'?'}`,
           `${trD('delta.time')} ${timeMs} ${unitTime}`
         ];
         window.PluginActions._pushUserMessage('info', parts.join('  |  '));
       }
     } catch(_e){}
  } catch(err){ Logger.error('Compare failed', err); State.set({ running:false }); const btnCancel=document.getElementById('btnCancel'); if(btnCancel){ btnCancel.style.display='none'; } }
 }
function renderDiff(diff){ if(window.DiffListRenderer){ DiffListRenderer.renderDiffList(diff); try { applyCategoryFilterToDiffList(); } catch(_e){} } }
function applyCategoryFilterToDiffList(){
  const st=State.get(); const diff=st.diff; if(!diff) return;
  const boxes=[...document.querySelectorAll('.cat-filter')];
  const active=boxes.filter(cb=>cb.checked).map(cb=>cb.getAttribute('data-cat'));
  const allSelected = active.length===boxes.length && boxes.length>0;
  const noneSelected = active.length===0;
  const groups=[...document.querySelectorAll('#diffContainer .diff-group')];
  const foundLabel = (document.querySelector('[data-i18n="label.found"]')?.textContent) || (window.__i18n && window.__i18n['label.found']) || 'Found:';
  groups.forEach(group=>{
    const gType=group.getAttribute('data-diff-type')||'';
    // Decide group visibility
    if(noneSelected){ group.style.display='none'; return; }
    if(!allSelected && !active.includes(gType)){ group.style.display='none'; return; }
    group.style.display='';
    // Now filter rows inside this group
    const rows=[...group.querySelectorAll('.diff-row')];
    let visibleCount=0;
    rows.forEach(row=>{
      const t=row.getAttribute('data-diff-type') || row.dataset.diffType || '';
      if(noneSelected){ row.style.display='none'; }
      else if(allSelected || active.includes(t)){ row.style.display=''; visibleCount++; }
      else { row.style.display='none'; }
    });
    const cntEl=group.querySelector('.diff-group-header .dg-count'); if(cntEl){ cntEl.textContent=foundLabel+" "+visibleCount; }
    if(visibleCount===0){ group.classList.add('no-visible'); } else { group.classList.remove('no-visible'); }
  });
}
function filterLegend(diff){
  const items=document.querySelectorAll('#legendSection li[data-difftype]'); if(!items) return;
  
  // Check if "Compare formatting" option is enabled
  const formatCheckbox = document.getElementById('optFormat');
  const isFormatEnabled = formatCheckbox && formatCheckbox.checked;
  
  if(!diff){ // до запуска: показываем все (как раньше)
    items.forEach(li=>{ 
      const isFormatItem = li.getAttribute('data-difftype') === 'format';
      // Don't show format item if formatting comparison is disabled
      if(isFormatItem && !isFormatEnabled) {
        li.style.display='none';
      } else {
        li.style.display=''; 
      }
      li.classList.remove('legend-dim'); 
    });
    updateCounts({}, false, false, false);
    return;
  }
  // Подсчет наличия типов в текущем diff
  const counts={}; (diff.raw||[]).forEach(d=>{ counts[d.type]=(counts[d.type]||0)+1; });
  items.forEach(li=>{
    const t=li.getAttribute('data-difftype');
    const isFormatItem = t === 'format';
    
    if(counts[t]>0){
      // Don't show format item if formatting comparison is disabled
      if(isFormatItem && !isFormatEnabled) {
        li.style.display='none';
      } else {
        li.style.display=''; // всегда видимы, даже если чекбокс снят
      }
    } else {
      li.style.display='none'; // тип не встречается в текущем сравнении
    }
    li.classList.remove('legend-dim');
  });
  updateCounts(counts, true, true, true);
}
function updateCounts(presentCounts, _sel, _treatAll, hasDiff){
  const spans=document.querySelectorAll('.cat-count[data-count]');
  spans.forEach(sp=>{ const cat=sp.getAttribute('data-count'); let val=''; if(hasDiff && presentCounts[cat]){ val=String(presentCounts[cat]); sp.style.display='inline-block'; } else if(hasDiff){ sp.style.display='none'; } else { sp.style.display='none'; } sp.textContent=val; });
}
 function toA1(r,c){ return colName(c)+ (r+1); }
 function colName(c){ let s=''; c++; while(c>0){ let m=(c-1)%26; s=String.fromCharCode(65+m)+s; c=Math.floor((c-1)/26); } return s; }
 function fmt(v){ if(v==null) return ''; if(typeof v==='object') return JSON.stringify(v); return String(v); }
 function escapeHtml(s){ return s.replace(/[&<>'"]/g, ch=>({ '&':'&amp;','<':'&lt;','>':'&gt;','\'':'&#39;','"':'&quot;' }[ch])); }
function insertReport(){ const st=State.get(); if(!st.diff) return; Logger.info('Insert report start');
  window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand(function(){ try {
     var reportName='CompareReport';
     var existing=Api.GetSheet(reportName);
     if(!existing){ Api.AddSheet(reportName); existing=Api.GetSheet(reportName); }
     if(existing){ existing.SetVisible(true); existing.SetActive(); var r=0; function w(row,col,text){ var cell=existing.GetCells(row,col); cell.SetValue(String(text)); }
       var metrics = Asc.scope.diffMetrics || {}; var stats=Asc.scope.diffStats || {}; var dl = Asc.scope.deltaLabels || { changed:'Changed cells:', structural:'Structural changes:', percent:'Change %:', rows:'Rows:', cols:'Cols:', time:'Time:' };
       w(r++,0,'Compare Report');
       w(r++,0,'Base'); w(r-1,1,Asc.scope.baseName||'');
       w(r++,0,'Target'); w(r-1,1,Asc.scope.targetName||'');
       // Legacy stats section
       w(r++,0,'Stats');
       for(var k in stats){ w(r,0,k); w(r,1,stats[k]); r++; }
       // Delta metrics section (localized)
       w(r++,0,'Delta');
       var densityPct = (metrics.diffDensity!=null)? (metrics.diffDensity*100).toFixed(2)+'%' : '';
       var changed = metrics.changedCells!=null? metrics.changedCells: '';
       var structural = metrics.structuralDiffs!=null? metrics.structuralDiffs: '';
  var rowsLine = ((metrics.previewBaseRows!=null?metrics.previewBaseRows:metrics.baseRows)!=null? (metrics.previewBaseRows!=null?metrics.previewBaseRows:metrics.baseRows):'')+'→'+((metrics.previewTargetRows!=null?metrics.previewTargetRows:metrics.targetRows)!=null? (metrics.previewTargetRows!=null?metrics.previewTargetRows:metrics.targetRows):'');
  var colsLine = ((metrics.previewBaseCols!=null?metrics.previewBaseCols:metrics.baseCols)!=null? (metrics.previewBaseCols!=null?metrics.previewBaseCols:metrics.baseCols):'')+'→'+((metrics.previewTargetCols!=null?metrics.previewTargetCols:metrics.targetCols)!=null? (metrics.previewTargetCols!=null?metrics.previewTargetCols:metrics.targetCols):'');
       w(r,0,dl.changed); w(r++,1,changed);
       w(r,0,dl.structural); w(r++,1,structural);
       w(r,0,dl.percent); w(r++,1,densityPct);
       w(r,0,dl.rows); w(r++,1,rowsLine);
       w(r,0,dl.cols); w(r++,1,colsLine);
       w(r,0,'CellsCompared'); w(r++,1, metrics.cellsCompared!=null? metrics.cellsCompared: '');
     }
   } catch(e){ }
   }, function(){});
 }
 function tr(k, fallback){ const el=document.querySelector(`[data-i18n="${k}"]`); return el? el.textContent : (fallback||k); }
 function pushMessage(kind,text){ const area=document.getElementById('messageArea'); if(!area) return; const span=document.createElement('span'); span.className='msg-'+kind; span.textContent=text; area.innerHTML=''; area.appendChild(span); }
 // internal helper for other modules (ui.js) to surface user messages
 window.PluginActions = window.PluginActions || {}; window.PluginActions._pushUserMessage = pushMessage;
 function clearMessages(){ const area=document.getElementById('messageArea'); if(area) area.innerHTML=''; }
 function resetLegend(){ legendDynamic=false; filterLegend(State.get().diff||null); }
 function enableLegendDynamic(){ legendDynamic=true; filterLegend(State.get().diff||null); }
 function refreshLegend(){ filterLegend(State.get().diff||null); }
// Also refilter diff rows when legend checkboxes change
document.addEventListener('change', e=>{ if(e.target && e.target.classList && e.target.classList.contains('cat-filter')){ applyCategoryFilterToDiffList(); }});
State.subscribe(st=>{ if(st.diff){ // when diff changes reapply filter
  setTimeout(applyCategoryFilterToDiffList,0);
} });
// Merge into existing PluginActions to avoid wiping methods injected earlier (exportReportXlsx, _pushUserMessage)
window.PluginActions = Object.assign(window.PluginActions || {}, { startCompare, insertReport, /* exportReportXlsx preserved if already defined */ resetLegend, enableLegendDynamic, refreshLegend });
 // Cancel button behavior
 document.addEventListener('click', e=>{ if(e.target && e.target.id==='btnCancel'){ if(!cancelFlag){ cancelFlag=true; const btn=e.target; btn.disabled=true; pushMessage('warn', (document.querySelector('[data-i18n="msg.cancelRequested"]')?.textContent)||'Запрошена отмена'); } }});
function prepareScopeForReport(){ const st=State.get(); if(!st.diff) return; if(!window.Asc){ window.Asc={ scope:{} }; }
  if(!window.Asc.scope) window.Asc.scope={};
  try {
    window.Asc.scope.baseName=st.diff.meta.base;
    window.Asc.scope.targetName=st.diff.meta.target;
    window.Asc.scope.diffStats=st.diff.stats;
    window.Asc.scope.diffMetrics=st.diff.meta.metrics || {};
    const i18n=window.__i18n||{};
    window.Asc.scope.deltaLabels={
      changed: i18n['delta.changedCells'] || 'Изменение ячеек:',
      structural: i18n['delta.structuralDiffs'] || 'Изменение структуры:',
      percent: i18n['delta.diffPercent'] || 'Процент изменений:',
      rows: i18n['delta.rows'] || 'Количество строк:',
      cols: i18n['delta.cols'] || 'Количество столбцов:',
      time: i18n['delta.time'] || 'Время сравнения:'
    };
  } catch(_e){}
}
State.subscribe(st=>{ if(st.diff) prepareScopeForReport(); });
// === Diff row navigation / preview focus (dual row indices support) ===
function focusPreviewCellDual(rBase, rTarget, c, diffType){
  const baseDiv=document.getElementById('basePreview');
  const targetDiv=document.getElementById('targetPreview');
  const structuralRow = diffType==='inserted-row' || diffType==='deleted-row';
  const structuralCol = diffType==='inserted-col' || diffType==='deleted-col';
  // Clear previous highlights
  [baseDiv,targetDiv].forEach(div=>{ if(!div) return; div.querySelectorAll('tr.preview-focus-row').forEach(tr=>tr.classList.remove('preview-focus-row')); div.querySelectorAll('.preview-focus, .preview-focus-col, .preview-focus-col-top, .preview-focus-col-bottom').forEach(td=>td.classList.remove('preview-focus','preview-focus-col','preview-focus-col-top','preview-focus-col-bottom')); });
  
  // Helper: map original column index to visual column index (accounting for hidden phantom columns)
  function getVisualColIndex(div, originalCol) {
    if(originalCol == null) return null;
    const tbl = div?.querySelector('table');
    if(!tbl) return originalCol;
    const thead = tbl.querySelector('thead');
    if(!thead) return originalCol;
    const headers = thead.querySelectorAll('th[data-col-index]');
    for(let i=0; i<headers.length; i++) {
      const dataCol = parseInt(headers[i].dataset.colIndex, 10);
      if(dataCol === originalCol) return i;
    }
    return originalCol; // fallback
  }
  
  function apply(div, rowIndex){ if(!div || rowIndex==null || rowIndex<0) return; const tbl=div.querySelector('table'); if(!tbl) return; const body=tbl.querySelector('tbody'); if(!body) return; const tr=body.querySelectorAll('tr')[rowIndex]; if(!tr) return; const cells=tr.querySelectorAll('td'); 
    // Map original column index to visual index
    const visualCol = getVisualColIndex(div, c);
    const targetCell=visualCol!=null? cells[visualCol]: null; 
    if(!targetCell && !structuralCol && !structuralRow) return; const refCell = targetCell || cells[0] || null; if(refCell){ const top=refCell.offsetTop; const left=refCell.offsetLeft; div.scrollTop=Math.max(0, top-40); div.scrollLeft=Math.max(0,left-40); }
    if(structuralRow){ tr.classList.add('preview-focus-row'); }
    if(structuralCol && visualCol != null){ let firstCell=null,lastCell=null; body.querySelectorAll('tr').forEach(rtr=>{ const tds=rtr.querySelectorAll('td'); const colCell=tds[visualCol]; if(colCell){ colCell.classList.add('preview-focus-col'); if(!firstCell) firstCell=colCell; lastCell=colCell; } }); if(firstCell) firstCell.classList.add('preview-focus-col-top'); if(lastCell) lastCell.classList.add('preview-focus-col-bottom'); }
    if(!structuralRow && !structuralCol && targetCell){ targetCell.classList.add('preview-focus'); }
  }
  // For inserted-row: only target side has the row; for deleted-row: only base side.
  if(diffType==='inserted-row'){ apply(targetDiv, rTarget); }
  else if(diffType==='deleted-row'){ apply(baseDiv, rBase); }
  else { apply(baseDiv, rBase!=null? rBase : rTarget); apply(targetDiv, rTarget!=null? rTarget : rBase); }
}
function activateDiffRow(row){ const diffType = row.dataset.diffType || ''; const cStr=row.dataset.c; const c = cStr!=null? parseInt(cStr,10): null; const rTargetStr = row.dataset.rTarget || row.dataset.r; const rBaseStr = row.dataset.rBase || row.dataset.r; const rTarget = rTargetStr!=null? parseInt(rTargetStr,10): null; const rBase = rBaseStr!=null? parseInt(rBaseStr,10): null; document.querySelectorAll('#diffContainer .diff-row.active').forEach(el=>{ el.classList.remove('active'); el.setAttribute('aria-selected','false'); }); row.classList.add('active'); row.setAttribute('aria-selected','true'); focusPreviewCellDual(rBase, rTarget, c, diffType); }
// expose for search navigation
window.activateDiffRow = activateDiffRow;
document.addEventListener('click', e=>{ let tgt=e.target; if(tgt && !tgt.closest && tgt.parentElement) tgt=tgt.parentElement; if(!(tgt instanceof Element)) return; const row=tgt.closest('#diffContainer .diff-row'); if(!row) return; activateDiffRow(row); });
document.addEventListener('keydown', e=>{ if(e.key==='Enter' || e.key===' '){ const ae=document.activeElement; if(ae && ae.classList && ae.classList.contains('diff-row')){ e.preventDefault(); activateDiffRow(ae); } } });
// Arrow key navigation between diff rows (preserves structural highlighting)
document.addEventListener('keydown', e=>{
  if(e.key!=='ArrowDown' && e.key!=='ArrowUp') return; const active=document.querySelector('#diffContainer .diff-row.active')||document.activeElement; if(!active || !active.classList || !active.classList.contains('diff-row')) return;
  e.preventDefault(); const dir = e.key==='ArrowDown'? 1 : -1; let cur=active; while(true){ cur = (dir>0)? cur.nextElementSibling : cur.previousElementSibling; if(!cur) return; if(cur.classList && cur.classList.contains('diff-row')) break; }
  cur.focus(); activateDiffRow(cur);
});
 // Highlight & Diff search now handled by separate modules (highlight.js, diff_search.js)
 // Progress bar UI update
// Progress bar with debounce & minimum visible duration to prevent flicker on very short tasks
(function(){
  const SHOW_DELAY_MS = 180;        // wait before showing bar
  const MIN_VISIBLE_MS = 400;       // keep visible at least this long once shown
  let showTimer=null; let hideTimer=null; let shownAt=0;
  function upsertStatusMessage(kind, text){
    const area=document.getElementById('messageArea');
    if(!area) return;
    let el=area.querySelector('span[data-source="status"]');
    if(!text){
      if(el) el.remove();
      return;
    }
    if(!el){
      el=document.createElement('span');
      el.dataset.source='status';
      el.setAttribute('data-testid','xlsx-status-message');
      area.appendChild(el);
    }
    el.className='msg-'+(kind||'info');
    el.textContent=text;
  }
  function ensureShow(bar){ if(bar.style.visibility==='visible') return; bar.style.visibility='visible'; shownAt=performance.now(); }
  function scheduleShow(bar){ if(showTimer) return; showTimer=setTimeout(()=>{ showTimer=null; ensureShow(bar); }, SHOW_DELAY_MS); }
  function scheduleHide(bar, span){ if(hideTimer){ clearTimeout(hideTimer); hideTimer=null; } const elapsed=performance.now()-shownAt; const wait = Math.max(0, MIN_VISIBLE_MS - elapsed); hideTimer=setTimeout(()=>{ bar.style.visibility='hidden'; span.style.width='0%'; hideTimer=null; }, wait); }
  State.subscribe(st=>{ const bar=document.getElementById('progressBar'); const pText=document.getElementById('progressText'); const btnCancel=document.getElementById('btnCancel'); if(!bar) return; const span=bar.querySelector('span'); if(!span) return; const pct=Math.round((st.progress||0)*100);
    const active = st.importing || st.running; const completed = !!st.diff && !active;
    // Active long task states
    if(active){
      scheduleShow(bar);
      span.style.width=pct+'%';
      if(active){
        if(st.importing){ const lblImp=document.querySelector('[data-i18n="status.importing"]'); pText.textContent=(lblImp?lblImp.textContent:'Импорт...')+' '+pct+'%'; }
        else { const lblRun=document.querySelector('[data-i18n="status.running"]'); pText.textContent=(lblRun?lblRun.textContent:'Выполнение...')+' '+pct+'%'; }
        pText.className='';
      }
      if(btnCancel) btnCancel.style.display = st.running ? 'inline-block':'none';
      return;
    }
    // Completed (diff available) -> show 100% then hide after min duration
  if(completed){
      clearTimeout(showTimer); showTimer=null; ensureShow(bar);
      span.style.width='100%';
  if(st.statusMessage){ upsertStatusMessage(st.statusKind, st.statusMessage); }
  else { upsertStatusMessage(null, null); }
  if(pText){ const lblDone=document.querySelector('[data-i18n="status.complete"]'); pText.textContent= (lblDone? lblDone.textContent : 'Готово'); pText.className=''; }
      if(btnCancel) btnCancel.style.display='none';
      scheduleHide(bar, span);
      return;
    }
    // Idle
    clearTimeout(showTimer); showTimer=null;
    if(bar.style.visibility==='visible'){ scheduleHide(bar, span); }
  if(st.statusMessage){ upsertStatusMessage(st.statusKind, st.statusMessage); }
  else { upsertStatusMessage(null, null); }
  if(pText){ const lblIdle=document.querySelector('[data-i18n="status.idle"]'); pText.textContent = (lblIdle? lblIdle.textContent : 'Ожидание'); pText.className=''; }
    if(btnCancel) btnCancel.style.display='none';
    // Not completed yet: clear delta if no diff
  });
})();

// Sidebar collapse feature
(function(){
  const body=document.body;
  try {
    const pb=document.getElementById('progressBar');
    const pt=document.getElementById('progressText');
    if(pb && !pb.getAttribute('data-testid')) pb.setAttribute('data-testid','xlsx-progress-bar');
    if(pt && !pt.getAttribute('data-testid')) pt.setAttribute('data-testid','xlsx-progress-text');
    const btnCancel=document.getElementById('btnCancel');
    if(btnCancel && !btnCancel.getAttribute('data-testid')) btnCancel.setAttribute('data-testid','xlsx-btn-cancel');
    const toggleBtn=document.getElementById('btnToggleSidebar');
    if(toggleBtn && !toggleBtn.getAttribute('data-testid')) toggleBtn.setAttribute('data-testid','xlsx-btn-toggle-sidebar');
    const btnAnalyze=document.getElementById('btnAnalyze');
    if(btnAnalyze && !btnAnalyze.getAttribute('data-testid')) btnAnalyze.setAttribute('data-testid','xlsx-btn-analyze');
  } catch(_e) {}
  function setCollapsed(on){
    body.classList.toggle('sidebar-collapsed', !!on);
    const tBtn=document.getElementById('btnToggleSidebar');
  // Support possible duplicate markup: pick icon inside button
  const btn=document.getElementById('btnToggleSidebar');
    if(tBtn){
  const hideTitle=document.querySelector('[data-i18n-title="btn.hidePanel"]')?.getAttribute('title') || document.querySelector('[data-i18n="btn.hidePanel"]')?.textContent || 'Скрыть панель';
  const showTitle=document.querySelector('[data-i18n-title="btn.showPanel"]')?.getAttribute('title') || document.querySelector('[data-i18n="btn.showPanel"]')?.textContent || 'Показать панель';
  if(on){ tBtn.title=showTitle; tBtn.setAttribute('aria-label', showTitle); }
  else { tBtn.title=hideTitle; tBtn.setAttribute('aria-label', hideTitle); }
    }
  }
  function toggle(){ setCollapsed(!document.body.classList.contains('sidebar-collapsed')); }
  document.addEventListener('click', e=>{
    if(e.target && (e.target.id==='btnToggleSidebar' || (e.target.closest && e.target.closest('#btnToggleSidebar')))){ toggle(); }
  });
  document.addEventListener('keydown', e=>{ if((e.ctrlKey||e.metaKey) && e.key==='\"'){ toggle(); }});
  // Auto show button & collapse after run starts first time; hide/reset on clear imports
  let autoCollapsed=false;
  State.subscribe(st=>{
    const tBtn=document.getElementById('btnToggleSidebar');
    if(!tBtn) return;
    if(st.running && !autoCollapsed){ setTimeout(()=>{ setCollapsed(true); }, 80); autoCollapsed=true; }
  });
  // Expose reset for Clear Imports (ui.js can call window.PluginActions.resetSidebarState())
  function resetSidebarState(){
    const tBtn=document.getElementById('btnToggleSidebar');
  setCollapsed(false);
  autoCollapsed=false;
  }
  window.PluginActions = window.PluginActions || {}; window.PluginActions.resetSidebarState = resetSidebarState;
  // Initial state: ensure not collapsed AND sync icon explicitly
  setCollapsed(false);
  // Force icon reset if mismatch
  // Icons handled purely by CSS classes now
})();

// (Theme handling removed)

// ===================================================================
// Memory Cleanup on Window Close
// ===================================================================
(function() {
  function cleanupXLSXResources() {
    try {
      // 1. Очистка workbook cache
      if (window.ImportModule && window.ImportModule.workbookCache) {
        window.ImportModule.workbookCache.clear();
      }
      
      // 2. Очистка format cache
      if (window.ImportModule && window.ImportModule.formatCache) {
        window.ImportModule.formatCache.clear();
      }
      
      // 3. Очистка virtual sheets
      if (window.VirtualSheets) {
        Object.keys(window.VirtualSheets).forEach(key => {
          delete window.VirtualSheets[key];
        });
      }
      
      // 4. Очистка diff engine
      if (diffEngine) {
        diffEngine = null;
      }
      
      // 5. Очистка State если есть большие diff данные
      if (window.State && window.State.get) {
        const st = window.State.get();
        if (st.diff && st.diff.raw) {
          // Не чистим полностью State, только тяжелые данные
          st.diff.raw = null;
        }
      }
      
      // 6. Сброс cancel flag
      cancelFlag = false;
      
      if (window.Logger) {
        Logger.info('XLSX plugin resources cleaned up');
      }
    } catch (error) {
      if (window.Logger) {
        Logger.error('Failed to cleanup XLSX resources', error);
      }
    }
  }
  
  // Регистрируем cleanup на beforeunload
  window.addEventListener('beforeunload', cleanupXLSXResources);
  
  // Дополнительная страховка на unload
  window.addEventListener('unload', cleanupXLSXResources);
  
  // Expose для ручного вызова если нужно
  window.PluginActions = window.PluginActions || {};
  window.PluginActions.cleanupXLSXResources = cleanupXLSXResources;
})();

})(window);
