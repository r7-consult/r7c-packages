(function(window){
 'use strict';
 if(window.ImportModule) return; // guard
 function qs(id){ return document.getElementById(id); }
 const baseSel=()=>qs('baseSheetSelect');
 const targetSel=()=>qs('targetSheetSelect');
 const baseFileLabel=()=>qs('baseFileNameLabel');
 const targetFileLabel=()=>qs('targetFileNameLabel');

 const virtualSheets = {}; // exported
 const sideMeta={
   base:{ fileName:null, sheetOrder:[], sheetMap:{} },
   target:{ fileName:null, sheetOrder:[], sheetMap:{} }
 };
 window.VirtualSheets = virtualSheets;

 function updateUploadButtons(){
   try {
     const st = (window.State && State.get && State.get())? State.get(): {};
     const disabled = !!st.lockUploads;
     ['btnUploadBase','btnUploadTarget','btnUploadFileBase','btnUploadFileTarget'].forEach(id=>{
       const el=document.getElementById(id); if(!el) return;
       try {
         const map = {
           btnUploadBase: 'xlsx-upload-base-menu',
           btnUploadTarget: 'xlsx-upload-target-menu',
           btnUploadFileBase: 'xlsx-upload-base-file',
           btnUploadFileTarget: 'xlsx-upload-target-file'
         };
         const tid = map[id];
         if(tid && !el.getAttribute('data-testid')) el.setAttribute('data-testid', tid);
       } catch(_e) {}
       el.disabled=disabled;
       const cont=el.closest('.upload-menu-container');
       if(cont){ cont.classList.toggle('disabled', disabled); if(disabled) cont.classList.remove('open'); }
       if(disabled){
         const msg = (window.__i18n && (window.__i18n['hint.clearBeforeImport'] || window.__i18n['hint.clearImportsFirst'] || window.__i18n['msg.importBlocked'])) || 'Сначала очистите импорт, чтобы загрузить новые листы';
         if(id==='btnUploadBase' || id==='btnUploadTarget'){
           if(cont){ cont.setAttribute('data-disabled-hint', msg); }
           el.classList.add('upload-disabled-hint');
           el.removeAttribute('title');
           el.removeAttribute('aria-label');
         }
       } else {
         if(cont){ cont.removeAttribute('data-disabled-hint'); }
         el.classList.remove('upload-disabled-hint');
         if(id==='btnUploadBase' || id==='btnUploadTarget' || id==='btnUploadFileBase' || id==='btnUploadFileTarget'){
           el.removeAttribute('title');
           el.removeAttribute('aria-label');
         } else {
           const origKey = el.getAttribute('data-i18n-title');
           if(origKey && window.__i18n && window.__i18n[origKey]){
             el.setAttribute('title', window.__i18n[origKey]);
             el.setAttribute('aria-label', window.__i18n[origKey]);
           } else { el.removeAttribute('title'); el.removeAttribute('aria-label'); }
         }
       }
     });
   } catch(_e){}
 }

 function importSheet(kind,file){ if(State.get().importing) return; const importingLabel=(document.querySelector('[data-i18n="status.importing"]')?.textContent)||'Импорт...'; State.set({ importing:true, statusMessage:importingLabel, statusKind:null, progress:0 });
   try { const st=State.get(); if(st.diff){ State.set({ importDirtyForRun:true }); } } catch(_e){}
   document.querySelectorAll('button,select,input').forEach(el=>{ if(!el.hasAttribute('data-prev-disabled')) el.setAttribute('data-prev-disabled', el.disabled? '1':'0'); el.disabled=true; });
   function finishImport(success){
     const completeLabel=(document.querySelector('[data-i18n="msg.importComplete"]')?.textContent)||'Импорт завершен';
     const st=State.get();
     State.set({ importing:false, statusMessage: success? completeLabel: null, progress: success? 1:0 });
     setTimeout(()=>{ const cur=State.get(); if(!cur.running && !cur.importing && !cur.statusMessage){ const idleLbl=document.querySelector('[data-i18n="status.idle"]')?.textContent||'Ожидание'; State.set({ statusMessage:idleLbl }); } },1600);
     // Restore previous enabled/disabled states for controls
     document.querySelectorAll('button,select,input').forEach(el=>{ const prev=el.getAttribute('data-prev-disabled'); if(prev!==null){ el.disabled = prev==='1'; el.removeAttribute('data-prev-disabled'); } });
     // Only enable sheet selects after a successful full import and if they have options
     if(success){
       ['baseSheetSelect','targetSheetSelect'].forEach(id=>{ const sel=document.getElementById(id); if(sel && sel.options && sel.options.length>0){ sel.disabled=false; } });
     } else {
       ['baseSheetSelect','targetSheetSelect'].forEach(id=>{ const sel=document.getElementById(id); if(sel){ sel.disabled=true; } });
     }
   }
   const ext=(file.name.split('.').pop()||'').toLowerCase(); const baseName=file.name.replace(/\.[^.]+$/,''); const prefix=kind==='base'? 'Imported_Base_':'Imported_Target_'; let label=prefix+baseName; const stNow=State.get(); let counter=2; while(stNow.sheets.includes(label)){ label=prefix+baseName+'_'+counter++; }
   function resetFileInput(k){ try { (k==='base'? qs('fileBase'):qs('fileTarget')).value=''; }catch(e){} }
   function pushImportMessage(kindMsg,key,fallback){ if(window.PluginActions && window.PluginActions._pushUserMessage){ const sel=document.querySelector('[data-i18n="'+key+'"]'); const msg= sel? sel.textContent : fallback; window.PluginActions._pushUserMessage(kindMsg,msg||fallback); } }
   if(['csv','txt'].includes(ext)){
     const reader=new FileReader(); reader.onload=e=>{ try { const text=decodeBestCsv(e.target.result); const lines=text.split(/\r?\n/); const total=lines.length||1; const chunk=20; for(let i=0;i<lines.length;i+=chunk){ State.set({ progress: Math.min(0.99, i/total) }); } const snap=parseCSV(text); registerVirtual(label,snap,kind,{ fileName:file.name, sheetName: baseName }); pushImportMessage('info','msg.importComplete','Импорт завершен'); finishImport(true); } catch(err){ if(err && err.message==='ROW_LIMIT_EXCEEDED'){ pushImportMessage('error','msg.rowLimit','Превышен лимит строк (5000)'); } else { Logger.error('Import failed', err); pushImportMessage('error','msg.importFailed','Ошибка импорта'); } finishImport(false); } finally { resetFileInput(kind); } }; reader.onerror=()=>{ Logger.error('File read error'); pushImportMessage('error','msg.fileReadError','Ошибка чтения файла'); resetFileInput(kind); finishImport(false); }; reader.readAsArrayBuffer(file);
   } else if(['xls','xlsx','xlsm','xlsb','ods','fods'].includes(ext)){
     const reader=new FileReader(); reader.onload=async e=>{ try { if(!window.XLSX){ try { await loadScriptOnce([
       'vendor/xlsx.full.min.js','vendor/xlsx.mini.min.js','https://unpkg.com/xlsx@0.19.3/dist/xlsx.full.min.js','https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.19.3/xlsx.full.min.js','https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js'
    ]); } catch(eLoad){ pushImportMessage('error','msg.xlsxLibLoadFailed','Не удалось загрузить библиотеку XLSX'); if(window.State){ const el=document.querySelector('[data-i18n="msg.xlsxLibLoadFailed"]'); State.set({ statusMessage: el?el.textContent:'Не удалось загрузить парсер XLSX', statusKind:'error' }); } throw eLoad; } }
      if(!window.XLSX){ pushImportMessage('error','msg.xlsxLibMissing','Библиотека XLSX недоступна'); if(window.State){ const el2=document.querySelector('[data-i18n="msg.xlsxLibMissing"]'); State.set({ statusMessage: el2?el2.textContent:'Парсер XLSX недоступен', statusKind:'error' }); } return; }
      if(!window.XLSX.utils || !window.XLSX.read){ const diag='Библиотека XLSX загружена, но работает некорректно'; Logger.error(diag); pushImportMessage('error','msg.xlsxLibMissing','Библиотека XLSX недоступна'); if(window.State){ State.set({ statusMessage: diag, statusKind:'error' }); } return; }
       
       // ✅ УЛУЧШЕНО: используем новый API кеша с учетом типа импорта (base/target)
       const sig=file.name+':'+(file.size||0); 
       let wbBundle=_WorkbookCache.get(sig, kind); 
       if(!wbBundle){ 
         wbBundle=await parseBinaryWorkbook(e.target.result,file.name); 
         _WorkbookCache.set(sig, wbBundle, kind);
         // ✅ ОПТИМИЗАЦИЯ: Очищаем кеш форматов при загрузке нового workbook
         // так как форматы зависят от workbook.Styles
         clearFormatCache();
       }
       
       const importableNames=wbBundle.sheetNames.filter(n=> wbBundle.sheets[n] && wbBundle.sheets[n].maxR>0);
       const skipped=wbBundle.sheetNames.filter(n=> wbBundle.sheets[n] && wbBundle.sheets[n].maxR===0);
      showSheetSelectionModal(importableNames,(chosen)=>{ if(!chosen){ finishImport(false); return; } const selected=chosen && chosen.length? chosen: importableNames; const batch=[]; selected.forEach(sn=>{ const snap=wbBundle.sheets[sn]; const uniqueLabel=ensureUniqueLabel(label+'_'+sn); batch.push({uniqueLabel,snap,sn}); }); let done=0; const totalSel=batch.length||1; batch.forEach(item=>{ 
        registerVirtual(item.uniqueLabel,item.snap,kind,{ fileName:file.name, sheetName:item.sn }, true); done++; if(done%1===0){ State.set({ progress: Math.min(0.99, done/totalSel) }); } }); if(window.UI && window.UI.renderSheets) window.UI.renderSheets(); let msg = 'Импортировано листов: '+selected.length; if(skipped.length) msg += ', пропущено: '+skipped.length+' (>5000 строк)'; pushImportMessage('info','msg.workbookImported', msg); finishImport(true); });
    } catch(err){ Logger.error('Workbook import failed', err); pushImportMessage('error','msg.importFailed','Ошибка импорта'); finishImport(false); } finally { resetFileInput(kind); } }; reader.onerror=()=>{ Logger.error('File read error'); pushImportMessage('error','msg.fileReadError','Ошибка чтения файла'); resetFileInput(kind); finishImport(false); }; reader.readAsArrayBuffer(file);
   } else { Logger.warn('Unsupported extension', ext); }
 }
 function ensureUniqueLabel(base){ let candidate=base; let c=2; const stNow=State.get(); while(stNow.sheets.includes(candidate)){ candidate=base+'_'+c++; } return candidate; }
 function showSheetSelectionModal(sheetNames,onDone){ const modal=qs('sheetSelectModal'); if(!modal){ onDone(sheetNames); return; } const listDiv=qs('sheetSelectList'); listDiv.innerHTML=''; sheetNames.forEach((n,idx)=>{ const safe=n.replace(/[^A-Za-z0-9_]/g,'_'); const id='shsel_'+safe; const lbl=document.createElement('label'); lbl.setAttribute('data-testid', 'xlsx-sheet-select-item-'+safe+'-'+idx); const cb=document.createElement('input'); cb.type='checkbox'; cb.id=id; cb.value=n; cb.checked=true; cb.setAttribute('data-testid', 'xlsx-sheet-select-checkbox-'+safe+'-'+idx); lbl.appendChild(cb); const span=document.createElement('span'); span.textContent=n; span.setAttribute('data-testid', 'xlsx-sheet-select-name-'+safe+'-'+idx); lbl.appendChild(span); listDiv.appendChild(lbl); }); 
  try {
    if(modal && !modal.getAttribute('data-testid')) modal.setAttribute('data-testid', 'xlsx-sheet-select-modal');
    if(listDiv && !listDiv.getAttribute('data-testid')) listDiv.setAttribute('data-testid', 'xlsx-sheet-select-list');
    const okBtn = qs('btnSheetSelectOk');
    const cancelBtn = qs('btnSheetSelectCancel');
    const allBtn = qs('btnSheetSelectAll');
    const noneBtn = qs('btnSheetSelectNone');
    if(okBtn && !okBtn.getAttribute('data-testid')) okBtn.setAttribute('data-testid', 'xlsx-sheet-select-ok');
    if(cancelBtn && !cancelBtn.getAttribute('data-testid')) cancelBtn.setAttribute('data-testid', 'xlsx-sheet-select-cancel');
    if(allBtn && !allBtn.getAttribute('data-testid')) allBtn.setAttribute('data-testid', 'xlsx-sheet-select-all');
    if(noneBtn && !noneBtn.getAttribute('data-testid')) noneBtn.setAttribute('data-testid', 'xlsx-sheet-select-none');
  } catch(_e) {}
  function cleanup(){ ['btnSheetSelectOk','btnSheetSelectCancel','btnSheetSelectAll','btnSheetSelectNone'].forEach(id=>{ const el=qs(id); if(el){ const clone=el.cloneNode(true); clone.disabled=false; clone.removeAttribute('data-prev-disabled'); el.parentNode.replaceChild(clone, el); } }); }
  function closeOnly(){ modal.classList.add('hidden'); cleanup(); }
  function okH(){ const chosen=[...listDiv.querySelectorAll('input[type=checkbox]:checked')].map(i=>i.value); closeOnly(); onDone(chosen); }
  function cancelH(){ closeOnly(); onDone(null); }
   function allH(){ listDiv.querySelectorAll('input[type=checkbox]').forEach(i=>i.checked=true); }
   function noneH(){ listDiv.querySelectorAll('input[type=checkbox]').forEach(i=>i.checked=false); }
   ['btnSheetSelectOk','btnSheetSelectCancel','btnSheetSelectAll','btnSheetSelectNone'].forEach(id=>{ const b=qs(id); if(b){ b.disabled=false; b.removeAttribute('data-prev-disabled'); }});

   qs('btnSheetSelectOk').addEventListener('click',okH); qs('btnSheetSelectCancel').addEventListener('click',cancelH); qs('btnSheetSelectAll').addEventListener('click',allH); qs('btnSheetSelectNone').addEventListener('click',noneH); modal.classList.remove('hidden'); }
 function registerVirtual(name,snap,kind, meta, skipRender){ 
   const stg=window.State; 
   const virtual = virtualSheets; 
   virtual[name]=snap; 
   
   const st=State.get(); const sheets = st.sheets.includes(name)? st.sheets.slice(): st.sheets.concat([name]); const sm=sideMeta[kind]; if(!sm.fileName || (meta && meta.fileName && meta.fileName!==sm.fileName)){ sm.fileName = meta? meta.fileName: 'File'; sm.sheetOrder=[]; sm.sheetMap={}; }
  const displaySheet = meta && meta.sheetName? meta.sheetName: name; if(!sm.sheetOrder.includes(displaySheet)) sm.sheetOrder.push(displaySheet); sm.sheetMap[displaySheet]=name; const current=State.get(); const patch={ sheets }; if(kind==='base'){ if(!current.base){ patch.base=name; } if(meta && meta.fileName){ patch.baseOriginalFile = meta.fileName; } } else { if(!current.target){ patch.target=name; } if(meta && meta.fileName){ patch.targetOriginalFile = meta.fileName; } } State.set(patch); if(!skipRender){ if(window.UI && window.UI.renderSheets) window.UI.renderSheets(); } else { if(window.UI && window.UI.updateDynamicBookSheetLabel) window.UI.updateDynamicBookSheetLabel(kind); }
  // Do not enable selects here; wait until finishImport() to avoid premature interaction during multi-sheet import.
   if(window.persistVirtualSheets) window.persistVirtualSheets();
   if(window.Logger && Logger.info) Logger.info('Imported sheet', name, {kind, rows:snap.maxR, cols:snap.maxC, meta}); }
 function clearImportedSheets(){ 
   const imported = Object.keys(virtualSheets); 
   if(!imported.length){ 
     Logger && Logger.info && Logger.info('Clear imports: nothing to clear'); 
     return; 
   } 
   
   Logger && Logger.info && Logger.info('Clearing imported virtual sheets', {count:imported.length, imported}); 
   
   // ✅ УЛУЧШЕНО: Очищаем workbook cache для освобождения памяти
   _WorkbookCache.clear();
   
   // ✅ ОПТИМИЗАЦИЯ: Очищаем кеш форматов
   clearFormatCache();
   
   imported.forEach(n=>{ delete virtualSheets[n]; }); 
   if(window.persistVirtualSheets) window.persistVirtualSheets(); 
   
   // ✅ УЛУЧШЕНО: Явно обнуляем старые массивы перед присвоением новых
   if(sideMeta && sideMeta.base){ 
     sideMeta.base.sheetOrder = null;
     sideMeta.base.sheetMap = null;
     sideMeta.base.fileName = null; 
     sideMeta.base.sheetOrder = []; 
     sideMeta.base.sheetMap = {}; 
   } 
   if(sideMeta && sideMeta.target){ 
     sideMeta.target.sheetOrder = null;
     sideMeta.target.sheetMap = null;
     sideMeta.target.fileName = null; 
     sideMeta.target.sheetOrder = []; 
     sideMeta.target.sheetMap = {}; 
   } const st=State.get(); const remainingSheets = st.sheets.filter(n=>!imported.includes(n)); const patch={ sheets: remainingSheets, base:null, target:null }; patch.diff=null; patch.statusMessage=null; patch.statusKind=null; patch.importDirtyForRun=false; patch.lockUploads=false; State.set(patch); const diffEl=document.getElementById('diffContainer'); if(diffEl) diffEl.innerHTML=''; const previewWrap=document.getElementById('sheetPreviewWrapper'); if(previewWrap) previewWrap.classList.add('hidden'); const syncCb=document.getElementById('chkSyncScroll'); if(syncCb){ syncCb.checked=false; } try { if(window.DiffSearch && typeof window.DiffSearch.clearSearch==='function'){ window.DiffSearch.clearSearch(true); } const sBar=document.getElementById('diffSearchBar'); if(sBar) sBar.classList.add('hidden'); const sFields=document.getElementById('diffSearchFields'); if(sFields) sFields.classList.add('hidden'); const sCount=document.getElementById('diffSearchCount'); if(sCount) sCount.textContent='0/0'; } catch(_e){} const bSel=baseSel&&baseSel(); if(bSel){ bSel.innerHTML=''; bSel.disabled=true; } const tSel=targetSel&&targetSel(); if(tSel){ tSel.innerHTML=''; tSel.disabled=true; } const bLbl=baseFileLabel&&baseFileLabel(); if(bLbl){ bLbl.textContent=''; bLbl.classList.remove('visible'); } const tLbl=targetFileLabel&&targetFileLabel(); if(tLbl){ tLbl.textContent=''; tLbl.classList.remove('visible'); } if(window.UI && window.UI.renderSheets) window.UI.renderSheets(); const btnClr=document.getElementById('btnClearImports'); if(btnClr){ btnClr.disabled=true; } try { document.querySelectorAll('#legendSection .cat-count').forEach(sp=>{ sp.textContent=''; sp.style.display='none'; }); document.querySelectorAll('#legendSection li.legend-dim').forEach(li=>li.classList.remove('legend-dim')); } catch(_e){} try { if(window.PluginActions && window.PluginActions.resetLegend) window.PluginActions.resetLegend(); } catch(_e){} if(window.PluginActions && window.PluginActions._pushUserMessage){ const el=document.querySelector('[data-i18n="msg.importCleared"]'); const msg= el? el.textContent: 'Импорт очищен'; window.PluginActions._pushUserMessage('info', msg); } try { if(window.PluginActions && window.PluginActions.resetSidebarState){ window.PluginActions.resetSidebarState(); } } catch(_e){} updateUploadButtons(); }
 function hookUploads(){ const fBase=qs('fileBase'), fTarget=qs('fileTarget'); if(fBase){ fBase.addEventListener('change', ()=>{ if(fBase.files&&fBase.files[0]) importSheet('base', fBase.files[0]); }); } if(fTarget){ fTarget.addEventListener('change', ()=>{ if(fTarget.files&&fTarget.files[0]) importSheet('target', fTarget.files[0]); }); } }
 try {
   const fBase = qs('fileBase');
   const fTarget = qs('fileTarget');
   if(fBase && !fBase.getAttribute('data-testid')) fBase.setAttribute('data-testid', 'xlsx-file-base');
   if(fTarget && !fTarget.getAttribute('data-testid')) fTarget.setAttribute('data-testid', 'xlsx-file-target');
 } catch(_e) {}
function hookUploadMenus(){
  function closeAll(){ document.querySelectorAll('.upload-menu-container.open').forEach(c=>c.classList.remove('open')); }
  function attemptOpen(cont){ if(!cont || cont.classList.contains('disabled')) return false; cont.classList.add('open'); return true; }
  ['btnUploadBase','btnUploadTarget'].forEach(id=>{
    const btn=document.getElementById(id); if(!btn) return; const cont=btn.closest('.upload-menu-container');
    function openMenu(){ if(btn.disabled) return; document.querySelectorAll('.upload-menu-container.open').forEach(c=>{ if(c!==cont) c.classList.remove('open'); }); attemptOpen(cont); }
    ['mouseenter','pointerenter','focus'].forEach(ev=> btn.addEventListener(ev, openMenu));
  btn.addEventListener('click', e=>{ e.preventDefault(); if(btn.disabled) return; if(!cont.classList.contains('open')){ openMenu(); setTimeout(()=>{ const dd=cont.querySelector('.upload-dropdown'); if(!dd || getComputedStyle(dd).display==='none'){ const fileInput = id==='btnUploadBase'? document.getElementById('fileBase'): document.getElementById('fileTarget'); if(fileInput) fileInput.click(); } },60); } /* если открыто — оставляем открытым */ });
  });
  document.addEventListener('click', e=>{ if(!e.target.closest('.upload-menu-container')) closeAll(); });
  document.addEventListener('keydown', e=>{ if(e.key==='Escape') closeAll(); });
  // Авто-закрытие при уходе курсора из контейнера (без авто-открытия)
  document.querySelectorAll('.upload-menu-container').forEach(cont=>{
    cont.addEventListener('mouseleave', ()=>{
      if(cont.classList.contains('open')){
        // небольшая задержка чтобы позволить кликнуть по пункту
        setTimeout(()=>{ cont.classList.remove('open'); }, 180);
      }
    });
  });
  const uploadFileBase=document.getElementById('btnUploadFileBase');
  const uploadFileTarget=document.getElementById('btnUploadFileTarget');
  const fBase=document.getElementById('fileBase');
  const fTarget=document.getElementById('fileTarget');
  if(uploadFileBase&&fBase) uploadFileBase.addEventListener('click', ()=>{ if(uploadFileBase.classList.contains('disabled-api')) return; fBase.click(); closeAll(); });
  if(uploadFileTarget&&fTarget) uploadFileTarget.addEventListener('click', ()=>{ if(uploadFileTarget.classList.contains('disabled-api')) return; fTarget.click(); closeAll(); });
}
 function loadScriptOnce(primary){ const sources = Array.isArray(primary)? primary: [primary]; return new Promise((resolve,reject)=>{ function tryNext(i){ if(i>=sources.length){ return reject(new Error('All script sources failed: '+sources.join(', '))); } const src=sources[i]; if(document.querySelector('script[data-src="'+src+'"]')){ setTimeout(()=>resolve(),50); return; } const s=document.createElement('script'); s.src=src; s.dataset.src=src; s.onload=()=>resolve(); s.onerror=()=>{ Logger.warn('Script load failed', src); tryNext(i+1); }; document.head.appendChild(s); } tryNext(0); }); }
 
 // ===================================================================
 // Workbook Cache Management (max 2 files: base + target)
 // ===================================================================
 const _WorkbookCache = {
   base: null,    // { sig: string, bundle: object }
   target: null,  // { sig: string, bundle: object }
   
   /**
    * Получить кешированный workbook
    * @param {string} sig - Сигнатура файла (filename:size)
    * @param {string} kind - Тип импорта ('base' или 'target')
    * @returns {object|null} - Кешированный bundle или null
    */
   get(sig, kind) {
     const entry = this[kind];
     if (entry && entry.sig === sig) {
       return entry.bundle;
     }
     return null;
   },
   
   /**
    * Сохранить workbook в кеш
    * @param {string} sig - Сигнатура файла
    * @param {object} bundle - Распарсенный workbook
    * @param {string} kind - Тип импорта ('base' или 'target')
    */
   set(sig, bundle, kind) {
     // Очищаем старый entry если был
     if (this[kind] && this[kind].bundle) {
       try {
         // Пытаемся освободить память от старого workbook
         const old = this[kind];
         if (old.bundle && old.bundle.sheets) {
           Object.keys(old.bundle.sheets).forEach(sheetName => {
             delete old.bundle.sheets[sheetName];
           });
         }
         delete old.bundle;
       } catch (e) {
         // Игнорируем ошибки при очистке
       }
     }
     
     this[kind] = { sig, bundle };
   },
   
   /**
    * Очистить весь кеш
    */
   clear() {
     ['base', 'target'].forEach(kind => {
       if (this[kind] && this[kind].bundle) {
         try {
           const entry = this[kind];
           if (entry.bundle && entry.bundle.sheets) {
             Object.keys(entry.bundle.sheets).forEach(sheetName => {
               delete entry.bundle.sheets[sheetName];
             });
           }
           delete entry.bundle;
         } catch (e) {
           // Игнорируем ошибки при очистке
         }
       }
       this[kind] = null;
     });
   },
   
   /**
    * Очистить кеш конкретного типа
    * @param {string} kind - Тип ('base' или 'target')
    */
   clearKind(kind) {
     if (this[kind] && this[kind].bundle) {
       try {
         const entry = this[kind];
         if (entry.bundle && entry.bundle.sheets) {
           Object.keys(entry.bundle.sheets).forEach(sheetName => {
             delete entry.bundle.sheets[sheetName];
           });
         }
         delete entry.bundle;
       } catch (e) {
         // Игнорируем ошибки при очистке
       }
     }
     this[kind] = null;
   }
 };
 
 // ===================================================================
 // Format Signature Cache (reduces redundant format extraction)
 // ===================================================================
 const _FormatCache = new Map();
 
 /**
  * Очистить кеш форматов
  */
 function clearFormatCache() {
   _FormatCache.clear();
 }
 
 /**
  * Получить ключ кеша для cell.s объекта
  * Использует JSON.stringify для создания уникального ключа
  */
 function getFormatCacheKey(style) {
   if (!style) return null;
   try {
     // Создаем стабильный ключ из критичных свойств стиля
     const key = JSON.stringify({
       numFmt: style.numFmt,
       fontId: style.fontId ?? style.fontid,
       fillId: style.fillId ?? style.fillid,
       patternType: style.patternType,
       fgColor: style.fgColor,
       alignment: style.alignment
     });
     return key;
   } catch (e) {
     return null;
   }
 }
 
 // Extract formatting signature from SheetJS cell object
 function extractFormatSignature(cell, workbook){
   if(!cell || !cell.s) {
     return null;
   }
   
   const style = cell.s;
   
   // ✅ ОПТИМИЗАЦИЯ: Проверяем кеш перед вычислением
   const cacheKey = getFormatCacheKey(style);
   if (cacheKey && _FormatCache.has(cacheKey)) {
     return _FormatCache.get(cacheKey);
   }
   
   const parts = [];
   
   try {
     // ========================================
     // НОВОЕ: Извлекаем font и fill из workbook.Styles по индексам
     // ========================================
     let font = null;
     let fill = null;
     
     // Попытка 1: Если есть прямые индексы fontId/fillId в cell.s
     if (workbook && workbook.Styles) {
       const styles = workbook.Styles;
       
       // Получаем fontId из cell.s
       const fontId = style.fontId !== undefined ? parseInt(style.fontId) : 
                      style.fontid !== undefined ? parseInt(style.fontid) : null;
       if (fontId !== null && styles.Fonts && styles.Fonts[fontId]) {
         font = styles.Fonts[fontId];
       }
       
       // Получаем fillId из cell.s
       const fillId = style.fillId !== undefined ? parseInt(style.fillId) :
                      style.fillid !== undefined ? parseInt(style.fillid) : null;
       if (fillId !== null && styles.Fills && styles.Fills[fillId]) {
         fill = styles.Fills[fillId];
       }
       
       // Попытка 2: Если индексы отсутствуют, но есть развернутые свойства (SheetJS 0.20+)
       // Ищем соответствующий CellXf по совпадению fill
       if (!font && !fill && styles.CellXf && (style.fgColor || style.patternType)) {
         // Ищем CellXf, у которого fill совпадает с текущим
         for (let i = 0; i < styles.CellXf.length; i++) {
           const xf = styles.CellXf[i];
           const xfFillId = xf.fillId !== undefined ? parseInt(xf.fillId) : 
                           xf.fillid !== undefined ? parseInt(xf.fillid) : null;
           
           if (xfFillId !== null && styles.Fills && styles.Fills[xfFillId]) {
             const xfFill = styles.Fills[xfFillId];
             
             // Проверяем совпадение fill
             const patternMatches = xfFill.patternType === style.patternType;
             const fgColorMatches = (!xfFill.fgColor && !style.fgColor) || 
                                   (xfFill.fgColor && style.fgColor && xfFill.fgColor.rgb === style.fgColor.rgb);
             
             if (patternMatches && fgColorMatches) {
               // Получаем font из этого CellXf
               const xfFontId = xf.fontId !== undefined ? parseInt(xf.fontId) :
                               xf.fontid !== undefined ? parseInt(xf.fontid) : null;
               if (xfFontId !== null && styles.Fonts && styles.Fonts[xfFontId]) {
                 font = styles.Fonts[xfFontId];
               }
               
               // Используем fill из CellXf
               fill = xfFill;
               break;
             }
           }
         }
       }
     }
     
     // Fallback: если font/fill уже в cell.s (старые версии SheetJS)
     if (!font && style.font) font = style.font;
     if (!fill) {
       // fill может быть прямо в style (fgColor, patternType)
       if (style.fgColor || style.bgColor || style.patternType) {
         fill = {
           patternType: style.patternType,
           fgColor: style.fgColor,
           bgColor: style.bgColor
         };
       } else if (style.fill) {
         fill = style.fill;
       }
     }
     
     // ========================================
     // Number format
     if(style.numFmt != null) parts.push('nf:' + style.numFmt);
     
     // Font properties (теперь используем извлеченный font)
     if(font){
       // Bold
       if(font.bold !== undefined || font.b !== undefined) {
         const isBold = font.bold || font.b;
         parts.push('b:' + (isBold ? '1' : '0'));
       }
       // Italic
       if(font.italic !== undefined || font.i !== undefined) {
         const isItalic = font.italic || font.i;
         parts.push('i:' + (isItalic ? '1' : '0'));
       }
       // Underline
       if(font.underline !== undefined || font.u !== undefined) {
         const isUnderline = font.underline || font.u;
         parts.push('u:' + (isUnderline ? '1' : '0'));
       }
       // Font name
       if(font.name) parts.push('fn:' + font.name);
       // Font size
       if(font.sz != null) parts.push('fs:' + font.sz);
       // Font color
       if(font.color){
         let color = '';
         if(font.color.rgb) {
           // Убираем альфа-канал если есть (AARRGGBB -> RRGGBB)
           const rgb = font.color.rgb.length === 8 ? font.color.rgb.slice(2) : font.color.rgb;
           color = '#' + rgb;
         } else if(font.color.theme != null && workbook && workbook.Themes){
           color = 'theme' + font.color.theme;
         }
         if(color) parts.push('fc:' + color);
       }
     }
     
     // Fill (background color) - теперь используем извлеченный fill
     if(fill && fill.fgColor){
       let bgColor = '';
       if(fill.fgColor.rgb) {
         // Убираем альфа-канал если есть
         const rgb = fill.fgColor.rgb.length === 8 ? fill.fgColor.rgb.slice(2) : fill.fgColor.rgb;
         bgColor = '#' + rgb;
       } else if(fill.fgColor.theme != null && workbook && workbook.Themes){
         bgColor = 'theme' + fill.fgColor.theme;
       }
       if(bgColor) parts.push('bg:' + bgColor);
     }
     
     // Alignment
     if(style.alignment){
       const align = style.alignment;
       if(align.horizontal) parts.push('ah:' + align.horizontal);
       if(align.vertical) parts.push('av:' + align.vertical);
     }
   } catch(e) {
     if(window.Logger) Logger.warn('Format extraction failed', e);
   }
   
   const signature = parts.length > 0 ? parts.join('|') : null;
   
   // ✅ ОПТИМИЗАЦИЯ: Сохраняем в кеш
   if (cacheKey) {
     _FormatCache.set(cacheKey, signature);
   }
   
   return signature;
 }
 
 async function parseBinaryWorkbook(arrayBuffer,fileName){ 
   // Debug logging
   if(window.Logger && Logger.info) {
     Logger.info('parseBinaryWorkbook called', {
       fileName,
       arrayBufferSize: arrayBuffer?.byteLength
     });
   }
   
   if(!window.XLSX){ try { Logger.info('Loading XLSX parser'); await loadScriptOnce('vendor/xlsx.full.min.js'); } catch(e){ Logger.error('Parser load failed', e); throw e; } } 
   if(!window.XLSX) throw new Error('Parser library not available'); 
   const data=new Uint8Array(arrayBuffer); let wb; 
   try { wb=XLSX.read(data,{type:'array', cellDates:true, cellStyles:true}); } catch(e){ Logger.error('XLSX.read failed', e); throw e; } 
   
   if(window.Logger && Logger.info) {
     Logger.info('XLSX.read completed', {
       sheetNames: wb.SheetNames,
       sheetCount: wb.SheetNames.length,
       cellStylesEnabled: true
     });
   }
   
   const sheets={}; 
   
   function mapType(cell){ if(!cell) return 'empty'; if(cell.f) return (cell.t==='n'||cell.t==='d')? 'number':'formula'; switch(cell.t){ case 'n': return 'number'; case 'd': return 'date'; case 'b': return 'boolean'; case 'e': return 'error'; case 's': default: return inferType(cell.v); } } 
   
   let sheetIndex=0; 
   for(const n of wb.SheetNames){ try { 
     const ws=wb.Sheets[n]; 
     if(!ws||!ws['!ref']){ sheets[n]={ maxR:0,maxC:0, rows:[] }; continue; } 
     const range=XLSX.utils.decode_range(ws['!ref']); 
     const maxR=range.e.r - range.s.r + 1; 
     const maxC=range.e.c - range.s.c + 1; 
     if(maxR>5000){ sheets[n]={ maxR:0,maxC:0, rows:[] }; Logger.warn('Skipped sheet due to row limit', {sheet:n, rows:maxR}); continue; } 
     const rows=[]; 
     for(let r=0;r<maxR;r++){ 
       const rowArr=[]; 
       for(let c=0;c<maxC;c++){ 
         const addr=XLSX.utils.encode_cell({r:range.s.r + r, c:range.s.c + c}); 
         const cell=ws[addr]; 
         if(cell){ 
           let v=cell.v; 
           const f=cell.f||null; 
           const t=mapType(cell); 
           if(cell.t==='d' && v instanceof Date){ v=v.toISOString().split('T')[0]; }
           
           // Log first cell to verify we're parsing
           if(r === 0 && c === 0 && window.Logger && Logger.info) {
             Logger.info('Parsing first cell', {
               sheet: n,
               addr,
               cellKeys: Object.keys(cell),
               cellV: cell.v,
               cellT: cell.t,
               hasS: !!cell.s
             });
           }
           
           // Extract formatting signature - ALWAYS extract, show/hide is controlled by checkbox
           const fm = extractFormatSignature(cell, wb);
           
           const cellData = {v,f,t};
           if(fm) {
             cellData.fm = fm;
           }
           
           rowArr.push(cellData); 
         } else { 
           rowArr.push({v:null,f:null,t:'empty'}); 
         } 
       } 
       rows.push(rowArr); 
       if(r % 200 === 0){ const totalEst = wb.SheetNames.length; const baseProgress = sheetIndex/totalEst; const local = (r+1)/maxR; const overall = Math.min(0.99, baseProgress + local*(1/totalEst)); State.set({ progress: overall }); await new Promise(rf=>setTimeout(rf)); } 
     } 
     sheets[n]={ maxR, maxC, rows }; 
     sheetIndex++; 
     State.set({ progress: Math.min(0.99, sheetIndex / wb.SheetNames.length) }); 
   } catch(e){ Logger.error('Sheet parse failed', n, e); } } 
   return { sheetNames: wb.SheetNames, sheets }; 
 }
 function inferType(v){ if(v===''||v==null) return 'empty'; if(!isNaN(Number(v))) return 'number'; if(/^(\d{4}-\d{2}-\d{2})$/.test(v)) return 'date'; return 'string'; }
 function parseCSV(text){ const rows=text.split(/\r?\n/); const matrix=rows.filter(r=>r.trim().length>0).map(r=>r.split(/,|;|\t/)); if(matrix.length>5000){ throw new Error('ROW_LIMIT_EXCEEDED'); } return { maxR:matrix.length, maxC:matrix[0]?matrix[0].length:0, rows:matrix.map(r=> r.map(v=>({v:v,f:null,t:inferType(v)}))) }; }
 function scoreDecoded(text,enc){ const total=(text.match(/[A-Za-z\u0400-\u04FF]/g)||[]).length; const cyr=(text.match(/[\u0400-\u04FF]/g)||[]).length; const repl=(text.match(/[\uFFFD\?]/g)||[]).length; const mojibake=(text.match(/[ÃÂÐÑ][\x80-\xBF]?/g)||[]).length; const cyrRatio=total? cyr/total:0; return { text: text, enc, cyrRatio, repl, mojibake }; }
 function decodeBestCsv(arrayBuffer){ const encodings=['utf-8','windows-1251','koi8-r','ibm866','iso-8859-5']; const bytes=new Uint8Array(arrayBuffer); if(bytes.length>=3 && bytes[0]===0xEF && bytes[1]===0xBB && bytes[2]===0xBF){ try { return new TextDecoder('utf-8').decode(bytes);}catch(_){} } const results=[]; encodings.forEach(enc=>{ try{ const dec=new TextDecoder(enc); const text=dec.decode(bytes); results.push(scoreDecoded(text,enc)); }catch(e){} }); if(!results.length){ try { return new TextDecoder().decode(bytes);}catch(e){ return ''; } } results.sort((a,b)=>{ if(b.cyrRatio!==a.cyrRatio) return b.cyrRatio-a.cyrRatio; if(a.repl!==b.repl) return a.repl-b.repl; return a.mojibake-b.mojibake; }); return results[0].text; }

 function expose(){ window.ImportModule={
   updateUploadButtons, importSheet, hookUploads, hookUploadMenus, clearImportedSheets,
   registerVirtual, showSheetSelectionModal, ensureUniqueLabel, extractFormatSignature,
   sideMeta, virtualSheets, parseBinaryWorkbook, inferType, parseCSV, decodeBestCsv,
   renderSheets: function(){ if(window.UI && window.UI.renderSheets) window.UI.renderSheets(); },
   // ✅ НОВОЕ: API для управления workbook cache
   workbookCache: {
     clear: () => _WorkbookCache.clear(),
     clearBase: () => _WorkbookCache.clearKind('base'),
     clearTarget: () => _WorkbookCache.clearKind('target'),
     getStats: () => ({
       base: _WorkbookCache.base ? { sig: _WorkbookCache.base.sig, sheets: _WorkbookCache.base.bundle?.sheetNames?.length || 0 } : null,
       target: _WorkbookCache.target ? { sig: _WorkbookCache.target.sig, sheets: _WorkbookCache.target.bundle?.sheetNames?.length || 0 } : null
     })
   },
   // ✅ ОПТИМИЗАЦИЯ: API для управления format cache
   formatCache: {
     clear: clearFormatCache,
     getSize: () => _FormatCache.size,
     getStats: () => ({ size: _FormatCache.size })
   }
 };
 }
 expose();
 if(window.State && State.subscribe){ State.subscribe(()=> updateUploadButtons()); }
})(window);
