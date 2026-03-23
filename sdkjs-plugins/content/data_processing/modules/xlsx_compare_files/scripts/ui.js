/** UI bindings */
(function(window){
 'use strict';
 // === Helper DOM ===
 function qs(id){ return document.getElementById(id); }
 const baseSel=()=>qs('baseSheetSelect');
 const targetSel=()=>qs('targetSheetSelect');
 const baseFileLabel=()=>qs('baseFileNameLabel');
 const targetFileLabel=()=>qs('targetFileNameLabel');
 // sideMeta moved to ImportModule; fallback local structure if module not yet loaded
 const sideMeta = (window.ImportModule && window.ImportModule.sideMeta)? window.ImportModule.sideMeta : { base:{ fileName:null, sheetOrder:[], sheetMap:{} }, target:{ fileName:null, sheetOrder:[], sheetMap:{} } };

 // Disable upload buttons while imports present
 // Upload button state delegated to ImportModule
 function updateUploadButtons(){ if(window.ImportModule && ImportModule.updateUploadButtons){ try { ImportModule.updateUploadButtons(); } catch(_e){} } }

 // === Virtual sheet store & persistence keys ===
 const virtualSheets = (window.ImportModule && window.ImportModule.virtualSheets) ? window.ImportModule.virtualSheets : (window.VirtualSheets||{}); // delegate
 const VIRTUAL_STORE_KEY = 'compareTables_virtualSheets_v1';
 const CATS_STORE_KEY = 'compareTables_categories_v1'; // deprecated (no persistence now)
 const SELECTION_STORE_KEY = 'compareTables_selection_v1'; // store {base,target}
 const _MIGRATION_FLAG = 'compareTables_storage_migrated_v1';

 function storage(){ return (function(){ try { if(window.localStorage){ return window.localStorage; } } catch(e){} return window.sessionStorage; })(); }
 function migrateIfNeeded(){
      try {
           const s=storage();
           if(!s) return;
           if(s.getItem(_MIGRATION_FLAG)) return; // already migrated
           // Migrate known keys (selection persistence reintroduced, options removed)
           const pairs=[VIRTUAL_STORE_KEY, CATS_STORE_KEY, SELECTION_STORE_KEY];
           pairs.forEach(k=>{ try { if(typeof localStorage!=='undefined' && localStorage && !localStorage.getItem(k) && sessionStorage.getItem(k)){ localStorage.setItem(k, sessionStorage.getItem(k)); } } catch(_e){} });
           s.setItem(_MIGRATION_FLAG,'1');
      } catch(e){ /* ignore */ }
 }

// Persistence disabled – always start with empty virtual sheet set
function persistVirtualSheets(){ /* disabled */ }
function restoreVirtualSheets(){ return []; }
function persistCategories(_cats){ /* disabled persistence (legend selection not persisted) */ }
function restoreCategories(){ return null; }
function persistSelection(_sel){ /* disabled */ }
function restoreSelection(){ return null; }

 // === Render & State Binding (custom limited to imported file sheets) ===
 function updateSideSelect(kind){
      // Gate: do not allow enabling selects before any virtual sheets imported
      const importedCount = Object.keys(window.VirtualSheets||{}).length;
     const sel = kind==='base'? baseSel(): targetSel();
     const meta = sideMeta[kind];
     const fileLabel = kind==='base'? baseFileLabel(): targetFileLabel();
     if(!sel) return;
     sel.innerHTML=''; sel.disabled=true; // keep disabled until populated
     if(meta.fileName){
          meta.sheetOrder.forEach((sh,idx)=>{ const o=document.createElement('option'); o.value=sh; o.textContent=sh; const safe=String(sh).replace(/[^A-Za-z0-9_]/g,'_'); o.setAttribute('data-testid', (kind==='base'? 'xlsx-base-sheet-option-' : 'xlsx-target-sheet-option-')+safe+'-'+idx); sel.appendChild(o); });
          // enable will happen after value selection below
          if(fileLabel){
               fileLabel.textContent=''; // will set after selection below via updateDynamicLabel
               fileLabel.classList.add('visible');
          }
          const st=State.get();
          const virtWanted = (kind==='base'? st.base: st.target);
          if(virtWanted){
               const found = Object.entries(meta.sheetMap).find(([disp,v])=>v===virtWanted);
               if(found) sel.value=found[0];
          }
          if(!sel.value && sel.options.length){
               const st2=State.get();
               const hasSel = kind==='base'? !!st2.base: !!st2.target;
               if(!hasSel){
                    sel.selectedIndex=0;
                    const virt=meta.sheetMap[sel.options[0].value];
                    if(kind==='base') State.set({ base:virt }); else State.set({ target:virt });
               }
          }
     // Enable only if options exist
     if(sel.options.length>0 && importedCount>0) sel.disabled=false; // extra guard against premature enable
          // After selection decisions, update label text
          updateDynamicBookSheetLabel(kind);
     } else {
          sel.disabled=true;
          if(fileLabel){ fileLabel.textContent=''; fileLabel.classList.remove('visible'); }
          // Ensure placeholder label (Источник 1/2) shown even with no file imported yet
          updateDynamicBookSheetLabel(kind);
     }
 }

// Build truncated book name + sheet name label: show start & end of long names
function truncateMiddle(name, maxLen){
     if(!name) return '';
     if(name.length<=maxLen) return name;
     const keep = maxLen-3; if(keep<4) return name.slice(0,maxLen); // fallback
     const first = Math.ceil(keep/2);
     const last = keep-first;
     return name.slice(0,first)+'…'+name.slice(name.length-last);
}
function stripExtension(fileName){ return fileName? fileName.replace(/\.[^.]+$/,''): fileName; }
function updateDynamicBookSheetLabel(kind){
     const meta=sideMeta[kind];
     const sel = kind==='base'? baseSel(): targetSel();
     const labelEl = kind==='base'? document.querySelector('label[for="baseSheetSelect"]'): document.querySelector('label[for="targetSheetSelect"]');
     if(!labelEl) return;
     const fileOriginal = stripExtension(meta && meta.fileName? meta.fileName: '');
     if(!fileOriginal){
          // i18n placeholders
          const ph1 = (window.__i18n && window.__i18n['placeholder.source1']) || 'Источник 1';
          const ph2 = (window.__i18n && window.__i18n['placeholder.source2']) || 'Источник 2';
          labelEl.textContent = kind==='base'? ph1: ph2; return;
     }
     labelEl.classList.add('visible');
     // Set full name first
     labelEl.textContent = fileOriginal;
     labelEl.classList.add('visible');
     // Match label width to select (if available) for fitting
     try {
          if(sel && sel.offsetWidth){
               labelEl.style.display='inline-block';
               labelEl.style.maxWidth = sel.offsetWidth + 'px';
               // If overflows, iteratively middle-truncate to fit
               if(labelEl.scrollWidth > labelEl.clientWidth){
                    // Heuristic initial target length proportional to available space
                    let full = fileOriginal;
                    let approxLen = Math.max(6, Math.floor(full.length * (labelEl.clientWidth / labelEl.scrollWidth)) - 1);
                    // Clamp
                    if(approxLen > full.length) approxLen = full.length;
                    // Binary style narrowing
                    let low=6, high=approxLen, best=approxLen;
                    function apply(len){ return truncateMiddle(full, len); }
                    // Ensure high not zero
                    if(high < 6) high = Math.min(12, full.length);
                    for(let i=0;i<20;i++){
                         const mid = Math.max(low, Math.min(high, Math.floor((low+high)/2)));
                         labelEl.textContent = apply(mid);
                         if(labelEl.scrollWidth <= labelEl.clientWidth){ best = mid; low = mid+1; }
                         else { high = mid-1; }
                         if(low>high) break;
                    }
                    labelEl.textContent = apply(best);
               }
          }
     } catch(_e){}
}
 function renderSheets(){ updateSideSelect('base'); updateSideSelect('target'); }
 // expose for ImportModule to call after registrations
 window.UI = window.UI || {}; window.UI.renderSheets = renderSheets; window.UI.updateDynamicBookSheetLabel = updateDynamicBookSheetLabel;
// MutationObserver removed – selects are managed only by internal code now
function startSelectObservers(){}
function bindState(){ State.subscribe(st=>{ const btnRun=qs('btnRun'), btnReport=qs('btnReport'); const hasDiff=!!st.diff; if(btnRun) btnRun.disabled=st.running; if(btnReport) btnReport.disabled=!hasDiff; /* progress text managed in plugin.js to avoid conflicts */ }); }
// Manage clear imports button enablement
State.subscribe(st=>{ const btnClr=qs('btnClearImports'); if(btnClr){ const virtNames = Object.keys(window.VirtualSheets||{}); btnClr.disabled = virtNames.length===0 || st.importing || st.running; } });
// Enforce disabled state after any potential external re-enable glitches
function enforceClearImportsDisabled(){ const btnClr=qs('btnClearImports'); if(!btnClr) return; const virtNames=Object.keys(window.VirtualSheets||{}); if(!virtNames.length && !btnClr.disabled) btnClr.disabled=true; }
State.subscribe(()=> enforceClearImportsDisabled());
State.subscribe(st=>{ const blocker=document.getElementById('uiBlocker'); if(blocker){ if(st.importing){ blocker.classList.add('active'); } else { blocker.classList.remove('active'); } }
     const modal=document.getElementById('sheetSelectModal'); const modalOpen = modal && !modal.classList.contains('hidden');
     if(st.importing){
          document.querySelectorAll('button,select,input').forEach(el=>{
               const isModalChild = modalOpen && modal.contains(el);
               if(isModalChild){
                    // Force enable modal interactive elements
                    el.disabled=false; el.removeAttribute('data-prev-disabled');
               } else {
                    if(!el.hasAttribute('data-prev-disabled')) el.setAttribute('data-prev-disabled', el.disabled? '1':'0');
                    el.disabled=true;
               }
          });
     } else {
            document.querySelectorAll('button,select,input').forEach(el=>{
               const prev=el.getAttribute('data-prev-disabled');
               if(prev!==null){ el.disabled = prev==='1'; el.removeAttribute('data-prev-disabled'); }
               else if(el.id!=='btnReport' && el.id!=='btnClearImports'){
                         // Generic enable for ordinary controls; skip special buttons with their own logic.
                         // Additional guard: keep sheet selects disabled until at least one sheet imported.
                         if((el.id==='baseSheetSelect' || el.id==='targetSheetSelect')){
                              const importedCount = Object.keys(window.VirtualSheets||{}).length;
                              if(importedCount===0){ el.disabled=true; return; }
                         }
                         el.disabled=false;
               }
          });
     }
});
function collectOptions(){ return { ignoreCase:qs('optIgnoreCase')?.checked, trim:qs('optTrim')?.checked, compareFormatting:qs('optFormat')?.checked }; }
function attachHandlers(){ 
     try {
          const map = {
               btnRun: 'xlsx-btn-run',
               btnCancel: 'xlsx-btn-cancel',
               btnReport: 'xlsx-btn-report',
               btnClearImports: 'xlsx-btn-clear-imports',
               baseSheetSelect: 'xlsx-base-sheet-select',
               targetSheetSelect: 'xlsx-target-sheet-select',
               chkSyncScroll: 'xlsx-sync-scroll',
               optIgnoreCase: 'xlsx-opt-ignore-case',
               optTrim: 'xlsx-opt-trim',
               optFormat: 'xlsx-opt-format',
               btnDiffSearchToggle: 'xlsx-diff-search-toggle',
               diffSearchFields: 'xlsx-diff-search-fields',
               diffSearchInput: 'xlsx-diff-search-input',
               diffSearchCount: 'xlsx-diff-search-count',
               btnDiffSearchPrev: 'xlsx-diff-search-prev',
               btnDiffSearchNext: 'xlsx-diff-search-next',
               btnDiffSearchClear: 'xlsx-diff-search-clear'
          };
          Object.entries(map).forEach(([id, tid]) => {
               const el = qs(id);
               if (el && !el.getAttribute('data-testid')) el.setAttribute('data-testid', tid);
          });
          const diffContainer = document.getElementById('diffContainer');
          if (diffContainer && !diffContainer.getAttribute('data-testid')) diffContainer.setAttribute('data-testid', 'xlsx-diff-container');
          const previewWrap = document.getElementById('sheetPreviewWrapper');
          if (previewWrap && !previewWrap.getAttribute('data-testid')) previewWrap.setAttribute('data-testid', 'xlsx-preview-wrapper');
          const basePreview = document.getElementById('basePreview');
          if (basePreview && !basePreview.getAttribute('data-testid')) basePreview.setAttribute('data-testid', 'xlsx-preview-base');
          const targetPreview = document.getElementById('targetPreview');
          if (targetPreview && !targetPreview.getAttribute('data-testid')) targetPreview.setAttribute('data-testid', 'xlsx-preview-target');
     } catch (_e) {}

     const opts=['optIgnoreCase','optTrim','optFormat']; 
     const run=qs('btnRun'), report=qs('btnReport'); 
     run&&run.addEventListener('click', ()=>window.PluginActions.startCompare()); 
     report&&report.addEventListener('click', ()=>{ if(window.PluginActions.exportReportXlsx) window.PluginActions.exportReportXlsx(); else if(window.PluginActions.insertReport) window.PluginActions.insertReport(); }); 
     
     opts.forEach(id=>{ 
          const el=qs(id); 
          el&&el.addEventListener('change', ()=>{ 
               const o=collectOptions(); 
               State.set({ options:o }); 
               
               // If "Compare formatting" checkbox changed
               if(id === 'optFormat') {
                    // Show/hide "format" legend item
                    const formatLegendItem = document.querySelector('#legendList li[data-difftype="format"]');
                    if(formatLegendItem) {
                         if(el.checked) {
                              formatLegendItem.style.display = '';
                         } else {
                              formatLegendItem.style.display = 'none';
                         }
                    }
                    
                    // Refresh preview if visible
                    const st = State.get();
                    const previewWrap = document.getElementById('sheetPreviewWrapper');
                    // Only refresh if we have diff data and preview is visible
                    if(st.diff && previewWrap && !previewWrap.classList.contains('hidden')) {
                         if(window.PreviewRenderer && PreviewRenderer.buildPreviews) {
                              try {
                                   PreviewRenderer.buildPreviews(st.base, st.target, st.diff, null);
                              } catch(e) {
                                   if(window.Logger) Logger.error('Preview refresh failed', e);
                              }
                         }
                    }
               }
          }); 
     });
     const clr=qs('btnClearImports'); if(clr){ clr.addEventListener('click', ()=>{ if(clr.disabled) return; clearImportedSheets(); }); }
      // Map display sheet name -> virtual sheet id on change
      baseSel()?.addEventListener('change', ()=>{ const disp=baseSel().value||null; const virt=disp? (sideMeta.base.sheetMap[disp]||null): null; State.set({ base:virt }); persistSelection({ base:State.get().base, target:State.get().target }); });
     targetSel()?.addEventListener('change', ()=>{ const disp=targetSel().value||null; const virt=disp? (sideMeta.target.sheetMap[disp]||null): null; State.set({ target:virt }); persistSelection({ base:State.get().base, target:State.get().target }); });
     // Update dynamic labels on selection changes
     baseSel()?.addEventListener('change', ()=>updateDynamicBookSheetLabel('base'));
     targetSel()?.addEventListener('change', ()=>updateDynamicBookSheetLabel('target'));
     // Legend logic removed (handled by LegendModule)
     }

 // === Type helpers & parsers moved to ImportModule ===
 // inferType, parseCSV, parseBinaryWorkbook, decodeBestCsv, scoreDecoded
 // All these functions are now in ImportModule
 
 function loadScriptOnce(primary){
  const sources = Array.isArray(primary)? primary: [primary];
     return new Promise((resolve,reject)=>{
          function tryNext(i){
               if(i>=sources.length){ return reject(new Error('All script sources failed: '+sources.join(', '))); }
               const src=sources[i];
               if(document.querySelector('script[data-src="'+src+'"]')){ // already loading/loaded
                    // Wait a tick to ensure onload maybe already fired
                    setTimeout(()=>resolve(),50);
                    return;
               }
               const s=document.createElement('script'); s.src=src; s.dataset.src=src; s.onload=()=>resolve(); s.onerror=()=>{ Logger.warn('Script load failed', src); tryNext(i+1); }; document.head.appendChild(s);
          }
          tryNext(0);
     });
 }

// Import logic removed (migrated to ImportModule)

 // (stale registerVirtual snippet removed; logic lives in ImportModule)

 // === Upload buttons ===
// Upload hooks moved to ImportModule

// clearImportedSheets delegated
window.clearImportedSheets = function(){ if(window.ImportModule && ImportModule.clearImportedSheets){ ImportModule.clearImportedSheets(); } };

// Soft reset for new import: clear diff, legend counts, search, keep imports
window.__softResetForNewImport = function(){
     try {
          const diffEl=document.getElementById('diffContainer'); if(diffEl) diffEl.innerHTML='';
          const previewWrap=document.getElementById('sheetPreviewWrapper'); if(previewWrap) previewWrap.classList.add('hidden');
          const syncCb=document.getElementById('chkSyncScroll'); if(syncCb){ syncCb.checked=false; }
          const sb=document.getElementById('diffSearchBar'); if(sb) sb.classList.add('hidden');
          const sf=document.getElementById('diffSearchFields'); if(sf) sf.classList.add('hidden');
          const si=document.getElementById('diffSearchInput'); if(si) si.value='';
          const scnt=document.getElementById('diffSearchCount'); if(scnt) scnt.textContent='0/0';
          if(window.__diffSearchState){ window.__diffSearchState.query=''; window.__diffSearchState.matches=[]; window.__diffSearchState.index=-1; }
          document.querySelectorAll('.cat-filter').forEach(cb=>cb.checked=true);
          document.querySelectorAll('.cat-count').forEach(sp=>{ sp.textContent=''; sp.style.display='none'; });
          const st=State.get(); State.set({ diff:null, statusMessage:null, statusKind:null, enabledCategories:null, highlightApplied:false, importDirtyForRun:false });
          if(window.PluginActions && window.PluginActions.resetLegend) window.PluginActions.resetLegend();
     } catch(_e){}
};

 // === Init ===
 async function waitForAPIReady(){ let guard=0; while(!API.isReady() && guard<200){ await new Promise(r=>setTimeout(r,50)); guard++; } }
 // Re-init helpers removed (repeated init disabled by request)
 async function init(){
      // Early guard to prevent race double-init (Asc.plugin.init + DOMContentLoaded)
      if(window.__UIInitStarted){ Logger.info('UI init skipped (already initialized)'); return; }
      window.__UIInitStarted=true;
      Logger.info('UI init start (first run)');
     try { window.addEventListener('r7api:real-ready', ()=>{ Logger.info('Editor API ready'); API._ready=true; setTimeout(async ()=>{ try { const sheets=await API.listSheets(); if(sheets && sheets.length){ State.set({ sheets }); renderSheets(); } } catch(e){ Logger.error('Refresh after API ready failed', e); } }, 200); }); } catch(_e){}
     // (flag already set at start)
     migrateIfNeeded();
     // Clear any previously persisted data to enforce clean start
     try { const s=storage(); if(s&&s.removeItem){ s.removeItem(VIRTUAL_STORE_KEY); s.removeItem(SELECTION_STORE_KEY); } } catch(e){}
      bindState();
     
     // Initialize format legend item visibility EARLY - before async operations
     // By default optFormat is unchecked in HTML, so format legend item should be hidden
     try {
          const formatCheckbox = qs('optFormat');
          const formatLegendItem = document.querySelector('#legendList li[data-difftype="format"]');
          if(formatLegendItem && formatCheckbox) {
               formatLegendItem.style.display = formatCheckbox.checked ? '' : 'none';
          }
     } catch(e) {
          Logger.warn('Failed to initialize format legend visibility', e);
     }
     
     startSelectObservers();
      // Subscribe guard: restore virtual sheets if some external action wiped them
      State.subscribe(st=>{
           if(window.__sheetGuarding) return;
           const virtNames=Object.keys(virtualSheets);
           if(!virtNames.length) return;
           const missing=virtNames.filter(v=>!st.sheets.includes(v));
           if(missing.length){
                try {
                     window.__sheetGuarding=true;
                     const union=[...st.sheets];
                     missing.forEach(v=>union.push(v));
                     Logger.warn('Sheet guard restored missing virtual sheets', {missing, before:st.sheets});
                     State.set({ sheets:union });
                     renderSheets();
                } finally { window.__sheetGuarding=false; }
           }
      });
       attachHandlers();
       // Initial hook attempts
       if(window.ImportModule){ ImportModule.hookUploads && ImportModule.hookUploads(); ImportModule.hookUploadMenus && ImportModule.hookUploadMenus(); }
       else {
            // Retry a few times if ImportModule not yet defined (script race safety)
            let retries=0; (function retryHooks(){
                 if(window.ImportModule){ try { ImportModule.hookUploads && ImportModule.hookUploads(); ImportModule.hookUploadMenus && ImportModule.hookUploadMenus(); updateUploadButtons(); } catch(_e){} return; }
                 if(retries++ < 10){ setTimeout(retryHooks,100); }
            })();
       }
     // Current-sheet buttons removed: no API readiness toggle needed here.
      Logger.info('UI init start');
      await waitForAPIReady();
      try {
           let sheets=await API.listSheets();
           let active=await API.getActiveSheet();
           if(!active || !sheets.includes(active)) active=null;
           const prev=State.get(); // should be empty first run
           const initState={ sheets };
           initState.base=null;
           initState.target=null;
           State.set(initState);
           // initial empty (wait for import)
           renderSheets();
      } catch(e){ Logger.error('Sheet list failed', e); }
     // Remove mock timeout handling (mock mode removed)
 }
 // Watchdog removed per user request
 if(document.readyState==='complete' || document.readyState==='interactive'){ setTimeout(()=>{ if(!window.__UIInitStarted) init(); },200); } else { document.addEventListener('DOMContentLoaded', ()=>{ if(!window.__UIInitStarted) init(); }); }
 window.UI=window.UI||{}; window.UI.init=init;
 // Provide a manual rebind helper (for debugging / dynamic reloads)
 window.UI.rebindUploads = function(){ try { if(window.ImportModule){ ImportModule.hookUploads && ImportModule.hookUploads(); ImportModule.hookUploadMenus && ImportModule.hookUploadMenus(); updateUploadButtons(); } } catch(_e){} };
 try { updateUploadButtons(); } catch(_e){}
 if(window.State && State.subscribe){ State.subscribe(()=> updateUploadButtons()); }
 // Initial state for upload buttons
 try { updateUploadButtons(); } catch(_e){}
 // Subscribe to state changes to refresh upload button lock
 State.subscribe(()=> { try { updateUploadButtons(); } catch(_e){} });
})(window);
