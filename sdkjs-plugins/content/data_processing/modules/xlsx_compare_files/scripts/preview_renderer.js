(function(window){
 'use strict';
 // Human-readable / localized label for diff category (used in tooltips)
 function catLabel(cat){
   const i18n = window.__i18n || {};
   switch(cat){
     case 'value': return i18n['legend.valueChange'] || 'Value changed';
     case 'formula': return i18n['legend.formulaChange'] || 'Formula changed';
     case 'formulaToValue': return i18n['legend.formulaToValue'] || 'Formula → Value';
     case 'valueToFormula': return i18n['legend.valueToFormula'] || 'Value → Formula';
     case 'type': return i18n['legend.typeChange'] || 'Type changed';
     case 'format': return i18n['legend.formatChange'] || 'Format changed';
     case 'inserted-row': return i18n['legend.insertedRow'] || 'Inserted row';
     case 'deleted-row': return i18n['legend.deletedRow'] || 'Deleted row';
     case 'inserted-col': return i18n['legend.insertedCol'] || 'Inserted column';
     case 'deleted-col': return i18n['legend.deletedCol'] || 'Deleted column';
     default: return cat;
   }
 }
 // Side-aware classification (base uses rBase, target uses r/rTarget)
 function classifySide(diff, isBase){
   const rows={}, cells={};
   if(!diff || !diff.raw) return {rows,cells};
   for(const d of diff.raw){
     if(d.type==='inserted-row'){
       if(isBase) continue; // inserted only highlighted on target side
       const rIns = d.r; if(rIns==null) continue; if(!rows[rIns]) rows[rIns]=new Set(); rows[rIns].add('inserted-row'); continue;
     }
     if(d.type==='deleted-row'){
       if(!isBase) continue; // deleted only highlighted on base side
       const rDel = d.r; if(rDel==null) continue; if(!rows[rDel]) rows[rDel]=new Set(); rows[rDel].add('deleted-row'); continue;
     }
     if(d.c==null) continue;
     const rIdx = isBase ? (d.rBase!=null? d.rBase : d.r) : (d.rTarget!=null? d.rTarget : d.r);
     if(rIdx==null) continue;
     if(!rows[rIdx]) rows[rIdx]=new Set(); rows[rIdx].add(d.type);
     const key=rIdx+':'+d.c; if(!cells[key]) cells[key]=new Set(); cells[key].add(d.type);
   }
   return {rows,cells};
 }
 function rowCssAll(set){ if(!set||!set.size) return []; const order=['inserted-row','deleted-row','inserted-col','deleted-col','formula','formulaToValue','valueToFormula','value','type','format']; const out=[]; order.forEach(k=>{ if(set.has(k)) out.push('diff-row-'+k.replace(/[^A-Za-z0-9]/g,'')); }); return out; }
 
 // Parse and apply cell formatting from signature string
 function applyFormatting(td, cell){
   // Increment call counter for debugging
   window._applyFormattingCalls = (window._applyFormattingCalls || 0) + 1;
   const callNum = window._applyFormattingCalls;
   
   // Check if "Compare formatting" checkbox is enabled
   const formatCheckbox = document.getElementById('optFormat');
   const isFormattingEnabled = formatCheckbox && formatCheckbox.checked;
   
   // Only apply formatting if checkbox is checked
   if(!isFormattingEnabled) {
     return;
   }
   
   if(!cell || !cell.fm) {
     return; // no formatting signature
   }
   
   const sig = cell.fm;
   if(!sig || typeof sig !== 'string') {
     return;
   }
   
   // Parse signature: "nf:0.00|b:1|i:0|fn:Arial|fs:12|fc:#000000|bg:#FFFFFF|ah:left|av:top"
   const parts = sig.split('|');
   const fmt = {};
   parts.forEach(p => {
     if(!p) return;
     const idx = p.indexOf(':');
     if(idx > 0) fmt[p.slice(0, idx)] = p.slice(idx + 1);
   });
   
   // Apply formatting to td element
   try {
     // Bold
     if(fmt.b === '1') td.style.fontWeight = 'bold';
     
     // Italic
     if(fmt.i === '1') td.style.fontStyle = 'italic';
     
     // Underline
     if(fmt.u === '1') td.style.textDecoration = 'underline';
     
     // Font name
     if(fmt.fn && fmt.fn !== '') td.style.fontFamily = fmt.fn;
     
     // Font size (convert to px if needed)
     if(fmt.fs && fmt.fs !== '') {
       const size = parseFloat(fmt.fs);
       if(!isNaN(size)) td.style.fontSize = size + 'pt';
     }
     
     // Font color
     if(fmt.fc && fmt.fc !== '' && fmt.fc !== 'undefined') {
       td.style.color = fmt.fc;
     }
     
     // Background color - only apply if cell doesn't have diff highlighting
     // Check if td has any diff-cell-* classes (they use !important backgrounds)
     const hasDiffClass = td.className && /diff-cell-/.test(td.className);
     if(!hasDiffClass && fmt.bg && fmt.bg !== '' && fmt.bg !== 'undefined') {
       td.style.backgroundColor = fmt.bg;
     }
     
     // Horizontal alignment
     if(fmt.ah && fmt.ah !== '') {
       const align = fmt.ah.toLowerCase();
       if(['left', 'center', 'right', 'justify'].includes(align)) {
         td.style.textAlign = align;
       }
     }
     
     // Vertical alignment
     if(fmt.av && fmt.av !== '') {
       const valign = fmt.av.toLowerCase();
       if(['top', 'middle', 'bottom'].includes(valign)) {
         td.style.verticalAlign = valign;
       }
     }
     
     // Mark that this cell has custom formatting
     if(Object.keys(fmt).length > 0) {
       td.classList.add('cell-formatted');
     }
     
     // Number format is handled by display value, not CSS
     // (the value should already be formatted by the engine)
   } catch(e) {
     const callNum = window._applyFormattingCalls || 0;
     if(window.Logger) Logger.error(`applyFormatting #${callNum}: ERROR`, e);
   }
 }
 
 let __colWidths = window.__previewColWidths || {}; // { index: px }
 const DEFAULT_COL_WIDTH = 120; // px fixed width for all data columns unless resized/persisted
 window.__previewColWidths = __colWidths;
 function applyWidthsToTables(baseTable,targetTable){
   Object.entries(__colWidths).forEach(([idx,w])=>{
     const i=parseInt(idx,10); if(isNaN(i)||!w) return;
     [baseTable,targetTable].forEach(tbl=>{
       if(!tbl) return;
       const cg = tbl.querySelector('colgroup');
       if(cg && cg.children[i+1]){ cg.children[i+1].style.width=w+'px'; }
     });
   });
 }
 function recalcTablesWidth(baseTable,targetTable){
   try {
     [baseTable,targetTable].forEach(tbl=>{
       if(!tbl) return;
       const cols = [...tbl.querySelectorAll('colgroup col[data-col]')].filter(c=>c.getAttribute('data-col')!=='_rownums');
       let total = 46; // row numbers col
       cols.forEach(colEl=>{
         const idx = parseInt(colEl.getAttribute('data-col'),10);
         const w = __colWidths[idx] || DEFAULT_COL_WIDTH;
         total += w;
       });
       tbl.style.width = total + 'px';
       // Ensure col elements have inline width (if missing) to prevent content-based expansion
       cols.forEach(colEl=>{
         const idx = parseInt(colEl.getAttribute('data-col'),10);
         if(!colEl.style.width){ colEl.style.width = (__colWidths[idx] || DEFAULT_COL_WIDTH)+'px'; }
       });
     });
   } catch(_e){}
 }
 function attachColumnResizers(baseTable,targetTable){
   const headerBase = baseTable && baseTable.tHead && baseTable.tHead.rows[0];
   const headerTarget = targetTable && targetTable.tHead && targetTable.tHead.rows[0];
   if(!headerBase && !headerTarget) return;
  function setup(th, idx, isBase){
     if(!th || th.dataset.resizable==='0') return;
     if(th.querySelector('.col-resizer')) return; // already
     const handle=document.createElement('span'); handle.className='col-resizer'; handle.title='Resize'; handle.setAttribute('data-testid', (isBase? 'xlsx-preview-base-col-resizer-' : 'xlsx-preview-target-col-resizer-')+String(idx)); th.appendChild(handle);
     let startX=0; let startW=0; let active=false;
     function onDown(e){ if(e.button!==0) return; active=true; startX=e.clientX; // prefer explicit width snapshot from colgroup
       const cg = (isBase? baseTable: targetTable).querySelector('colgroup');
       const colEl = cg && cg.children[idx+1];
       startW = colEl && colEl.offsetWidth? colEl.offsetWidth : th.offsetWidth;
       th.classList.add('resizing');
       document.addEventListener('mousemove', onMove);
       document.addEventListener('mouseup', onUp);
       e.preventDefault();
     }
     function onMove(e){ if(!active) return; const dx=e.clientX-startX; let newW=Math.max(40, startW+dx); if(newW>800) newW=800; __colWidths[idx]=newW; // apply to both tables if present
       [baseTable,targetTable].forEach(tbl=>{ if(!tbl) return; const cg=tbl.querySelector('colgroup'); if(cg && cg.children[idx+1]) cg.children[idx+1].style.width=newW+'px'; }); recalcTablesWidth(baseTable,targetTable); }
     function onUp(){ if(!active) return; active=false; th.classList.remove('resizing'); document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); recalcTablesWidth(baseTable,targetTable); try { localStorage.setItem('compareSheets_colWidths', JSON.stringify(__colWidths)); } catch(_e){} }
     handle.addEventListener('mousedown', onDown);
   }
   // Determine max columns across both
   const maxCols = Math.max(
     headerBase? headerBase.cells.length-1:0,
     headerTarget? headerTarget.cells.length-1:0
   );
   for(let i=0;i<maxCols;i++){
     const thB = headerBase && headerBase.cells[i+1];
     const thT = headerTarget && headerTarget.cells[i+1];
     // Only sync resize if column exists in both; if only one side, still allow local resize
     if(thB) setup(thB, i, true);
     if(thT) setup(thT, i, false);
   }
 }
 // Restore stored widths from localStorage (persisted across sessions)
 try { const stored=localStorage.getItem('compareSheets_colWidths'); if(stored){ const parsed=JSON.parse(stored); if(parsed && typeof parsed==='object'){ __colWidths=Object.assign(__colWidths, parsed); window.__previewColWidths=__colWidths; } } } catch(_e){}

 function buildPreviews(baseName,targetName,diff, diffEngineInstance){ return Promise.resolve().then(async ()=>{
   try {
     const wrap=document.getElementById('sheetPreviewWrapper'); if(!wrap) return;
     const baseDiv=document.getElementById('basePreview'); const targetDiv=document.getElementById('targetPreview'); if(!baseDiv||!targetDiv) return;
     
     // Check if we have valid sheet names
     if(!baseName || !targetName) {
       if(window.Logger) Logger.warn('buildPreviews: missing sheet names', {baseName, targetName});
       return;
     }
     
     const eng = diffEngineInstance || (window.DiffEngine? new window.DiffEngine(window.State && State.get()? State.get().options: {}): null);
     if(!eng) return;
  const [baseSnap,targetSnap] = await Promise.all([eng.snapshotSheet(baseName), eng.snapshotSheet(targetName)]);
  
  // Validate snapshots
  if(!baseSnap || !targetSnap) {
    if(window.Logger) Logger.warn('buildPreviews: failed to get snapshots', {baseSnap: !!baseSnap, targetSnap: !!targetSnap});
    return;
  }
  
  // Insert phantom columns/rows for visual alignment in preview
  // This ensures deleted columns in base and inserted columns in target appear at correct visual positions
  if(diff && diff.raw) {
    const deletedCols = new Set();
    const insertedCols = new Set();
    const deletedRows = new Set();
    const insertedRows = new Set();
    
    for(const d of diff.raw) {
      if(d.type === 'deleted-col') deletedCols.add(d.c);
      if(d.type === 'inserted-col') insertedCols.add(d.c);
      if(d.type === 'deleted-row') deletedRows.add(d.r);
      if(d.type === 'inserted-row') insertedRows.add(d.r);
    }
    
    // Insert phantom columns in target for deleted columns (reverse order to maintain indices)
    if(deletedCols.size > 0) {
      const sortedDeleted = Array.from(deletedCols).sort((a,b) => b-a);
      for(const colIdx of sortedDeleted) {
        for(let r=0; r<targetSnap.maxR; r++) {
          if(!targetSnap.rows[r]) targetSnap.rows[r] = [];
          targetSnap.rows[r].splice(colIdx, 0, {v:null, f:null, t:null, _phantom:true});
        }
        targetSnap.maxC++;
      }
    }
    
    // Insert phantom columns in base for inserted columns (reverse order)
    if(insertedCols.size > 0) {
      const sortedInserted = Array.from(insertedCols).sort((a,b) => b-a);
      for(const colIdx of sortedInserted) {
        for(let r=0; r<baseSnap.maxR; r++) {
          if(!baseSnap.rows[r]) baseSnap.rows[r] = [];
          baseSnap.rows[r].splice(colIdx, 0, {v:null, f:null, t:null, _phantom:true});
        }
        baseSnap.maxC++;
      }
    }
    
    // Insert phantom rows in target for deleted rows (reverse order)
    if(deletedRows.size > 0) {
      const sortedDeleted = Array.from(deletedRows).sort((a,b) => b-a);
      for(const rowIdx of sortedDeleted) {
        const phantomRow = new Array(targetSnap.maxC).fill(null).map(() => ({v:null, f:null, t:null, _phantom:true}));
        targetSnap.rows.splice(rowIdx, 0, phantomRow);
        targetSnap.maxR++;
      }
    }
    
    // Insert phantom rows in base for inserted rows (reverse order)
    if(insertedRows.size > 0) {
      const sortedInserted = Array.from(insertedRows).sort((a,b) => b-a);
      for(const rowIdx of sortedInserted) {
        const phantomRow = new Array(baseSnap.maxC).fill(null).map(() => ({v:null, f:null, t:null, _phantom:true}));
        baseSnap.rows.splice(rowIdx, 0, phantomRow);
        baseSnap.maxR++;
      }
    }
  }
  
  // --- Normalize snapshot shape (some APIs may return rowCount/colCount or omit maxR/maxC) ---
  function normalizeSnap(snap,label){ 
    if(!snap) {
      if(window.Logger) Logger.warn('normalizeSnap: null snapshot', {label});
      return;
    }
    
    let changed=false; 
    
    if(snap.maxR==null || snap.maxR===0){ 
      if(typeof snap.rowCount==='number' && snap.rowCount>0){ 
        snap.maxR=snap.rowCount; 
        changed=true; 
      } else if(Array.isArray(snap.rows)){ 
        snap.maxR=Math.max(snap.maxR||0, snap.rows.length); 
        changed=true; 
      } else { 
        snap.maxR=0; 
        changed=true; 
      } 
    }
    
    if(snap.maxC==null || snap.maxC===0){ 
      if(typeof snap.colCount==='number' && snap.colCount>0){ 
        snap.maxC=snap.colCount; 
        changed=true; 
      } else if(Array.isArray(snap.rows)){ 
        let m=0; 
        for(const r of snap.rows){ 
          if(r && r.length>m) m=r.length; 
        } 
        snap.maxC=Math.max(snap.maxC||0,m); 
        changed=true; 
      } else { 
        snap.maxC=0; 
        changed=true; 
      } 
    }
    
    if(!Array.isArray(snap.rows)){ 
      snap.rows=[]; 
      changed=true; 
    }
  }
  
  normalizeSnap(baseSnap,'base'); 
  normalizeSnap(targetSnap,'target');
  // Show FULL actual sheet dimensions; then trim trailing empty rows/cols (based on data or diff) before rendering.
  const baseMaps = classifySide(diff,true);
  const targetMaps = classifySide(diff,false);
     let colStructInserted=new Set(), colStructDeleted=new Set();
     if(diff && diff.raw){ diff.raw.forEach(d=>{ if(d.type==='inserted-col') colStructInserted.add(d.c); else if(d.type==='deleted-col') colStructDeleted.add(d.c); }); }
  // (Removed obsolete first computeEffectiveSize definition)
     function computeEffectiveSize(snap,maps,isBase){
       // Safety check
       if(!snap || snap.maxR == null || snap.maxC == null) {
         if(window.Logger) Logger.warn('computeEffectiveSize: invalid snapshot', {
           hasSnap: !!snap,
           maxR: snap?.maxR,
           maxC: snap?.maxC,
           isBase
         });
         return { rows: 0, cols: 0 };
       }
       
       if(!maps || !maps.rows || !maps.cells) {
         if(window.Logger) Logger.warn('computeEffectiveSize: invalid maps', {
           hasMaps: !!maps,
           hasRows: maps?.rows,
           hasCells: maps?.cells,
           isBase
         });
         return { rows: snap.maxR || 0, cols: snap.maxC || 0 };
       }
       
       const rowsTotal = snap.maxR, colsTotal = snap.maxC;
       const rowTypesLocal = maps.rows, cellTypesLocal = maps.cells;
       const hasStructCol = (c)=> (colStructInserted.has(c) && !isBase) || (colStructDeleted.has(c) && isBase);
       function cellMeaningful(cell){ if(!cell) return false; return !(cell.v==='' || cell.v==null) || !!cell.f; }
       function rowHasData(r){ const row = snap.rows[r]; if(row){ for(let c=0;c<row.length;c++){ if(cellMeaningful(row[c])) return true; } }
         if(rowTypesLocal[r] && rowTypesLocal[r].size) return true; return false; }
       function colHasData(c){ for(let r=0;r<rowsTotal;r++){ const row=snap.rows[r]; if(row && cellMeaningful(row[c])) return true; const key=r+':'+c; if(cellTypesLocal[key] && cellTypesLocal[key].size) return true; }
         if(hasStructCol(c)) return true; return false; }
       let lastRow=-1; for(let r=rowsTotal-1;r>=0;r--){ if(rowHasData(r)){ lastRow=r; break; } }
       let lastCol=-1; for(let c=colsTotal-1;c>=0;c--){ if(colHasData(c)){ lastCol=c; break; } }
       return { rows: lastRow+1, cols: lastCol+1 };
     }
  const effBase = computeEffectiveSize(baseSnap, baseMaps, true);
  const effTarget = computeEffectiveSize(targetSnap, targetMaps, false);
  // Логирование эффективных размеров после обрезки пустых строк/столбцов
  if(window.Logger) {
    Logger.info('Effective sizes computed', {
      base: {eff: effBase, maxR: baseSnap.maxR, maxC: baseSnap.maxC},
      target: {eff: effTarget, maxR: targetSnap.maxR, maxC: targetSnap.maxC}
    });
  }
  // If snapshots appear empty but diff has coordinates, derive fallback counts from diff
  function deriveCountsFromDiff(which){ if(!diff || !diff.raw) return {rows:0, cols:0}; let maxR=-1,maxC=-1; for(const d of diff.raw){ const rCandidate = which==='base'? (d.rBase!=null? d.rBase : (d.type==='deleted-row'? d.r : null)) : (d.rTarget!=null? d.rTarget : (d.type==='inserted-row'? d.r : null)); if(rCandidate!=null && rCandidate>maxR) maxR=rCandidate; if(d.c!=null && d.c>maxC) maxC=d.c; }
    return { rows:maxR+1, cols:maxC+1 };
  }
  // Fallbacks: if heuristic trimming produced 0 but snapshot has size, revert to full snapshot
  let rCountBase = effBase.rows>0 ? effBase.rows : (baseSnap.maxR||0);
  let rCountTarget = effTarget.rows>0 ? effTarget.rows : (targetSnap.maxR||0);
  let cCountBase = effBase.cols>0 ? effBase.cols : (baseSnap.maxC||0);
  let cCountTarget = effTarget.cols>0 ? effTarget.cols : (targetSnap.maxC||0);
  if(rCountBase===0){ const d=deriveCountsFromDiff('base'); if(d.rows>0) rCountBase=d.rows; if(d.cols>0 && cCountBase===0) cCountBase=d.cols; }
  if(rCountTarget===0){ const d=deriveCountsFromDiff('target'); if(d.rows>0) rCountTarget=d.rows; if(d.cols>0 && cCountTarget===0) cCountTarget=d.cols; }
  // Логирование финальных счетчиков строк/столбцов перед рендерингом
  if(window.Logger) {
    Logger.info('Final row/col counts before rendering', {
      base: {rows: rCountBase, cols: cCountBase},
      target: {rows: rCountTarget, cols: cCountTarget}
    });
  }
     function colName(c){ let s=''; c++; while(c>0){ let m=(c-1)%26; s=String.fromCharCode(65+m)+s; c=Math.floor((c-1)/26); } return s; }
  function buildTable(snap,isBase,rLimit,cLimit){
    // Filter out phantom columns and rows for display
    const visibleCols = [];
    for(let c=0; c<cLimit; c++) {
      const isPhantom = snap.rows[0] && snap.rows[0][c] && snap.rows[0][c]._phantom;
      if(!isPhantom) visibleCols.push(c);
    }
    const visibleRows = [];
    for(let r=0; r<rLimit; r++) {
      const isPhantom = snap.rows[r] && snap.rows[r][0] && snap.rows[r][0]._phantom;
      if(!isPhantom) visibleRows.push(r);
    }
    
    const tbl=document.createElement('table');
    tbl.setAttribute('data-testid', isBase? 'xlsx-preview-base-table' : 'xlsx-preview-target-table');
    const colgroup=document.createElement('colgroup');
    const cornerCol=document.createElement('col'); cornerCol.setAttribute('data-col','_rownums'); colgroup.appendChild(cornerCol);
    for(let i=0; i<visibleCols.length; i++){ 
      const c = visibleCols[i];
      const col=document.createElement('col'); 
      col.setAttribute('data-col',String(c)); 
      const w = __colWidths[c] || DEFAULT_COL_WIDTH; 
      col.style.width = w+'px'; 
      colgroup.appendChild(col); 
    }
    tbl.appendChild(colgroup);
    const thead=document.createElement('thead'); const hr=document.createElement('tr');
    const corner=document.createElement('th'); corner.textContent=''; corner.className='corner-cell'; corner.dataset.resizable='0'; corner.setAttribute('data-testid', isBase? 'xlsx-preview-base-corner' : 'xlsx-preview-target-corner'); hr.appendChild(corner);
    for(let i=0; i<visibleCols.length; i++){
      const c = visibleCols[i];
      const th=document.createElement('th'); 
      th.textContent=colName(c); 
      th.dataset.colIndex=String(c); 
      th.setAttribute('data-testid', (isBase? 'xlsx-preview-base-col-' : 'xlsx-preview-target-col-')+String(c));
      if(colStructInserted.has(c) && !isBase) th.classList.add('col-struct-inserted'); 
      if(colStructDeleted.has(c) && isBase) th.classList.add('col-struct-deleted'); 
      hr.appendChild(th);
    } 
    thead.appendChild(hr); tbl.appendChild(thead);
    const tb=document.createElement('tbody'); tbl.appendChild(tb);
  // --- Row rendering helpers (batched for large sheets) ---
  function renderRow(r){ const maps = isBase? baseMaps : targetMaps; const rowTypes = maps.rows; const cellTypes = maps.cells; const rawSet=rowTypes[r]; let rowClasses=rowCssAll(rawSet);
      const tr=document.createElement('tr');
      tr.setAttribute('data-testid', (isBase? 'xlsx-preview-base-row-' : 'xlsx-preview-target-row-')+String(r));
      // Filter structural classes based on side: inserted-row only target, deleted-row only base
      if(isBase){ rowClasses=rowClasses.filter(c=> c!=='diff-row-inserted-row'); } else { rowClasses=rowClasses.filter(c=> c!=='diff-row-deleted-row'); }
      rowClasses.forEach(cls=>tr.classList.add(cls));
      const th=document.createElement('th'); th.textContent=String(r+1);
      th.setAttribute('data-testid', (isBase? 'xlsx-preview-base-row-header-' : 'xlsx-preview-target-row-header-')+String(r));
      if(rawSet){ if(rawSet.has('inserted-row') && !isBase) th.classList.add('row-struct-inserted'); if(rawSet.has('deleted-row') && isBase) th.classList.add('row-struct-deleted'); }
      tr.appendChild(th);
      const row=snap.rows[r];
      for(let i=0; i<visibleCols.length; i++){
        const c = visibleCols[i];
        const td=document.createElement('td'); const cell=row?row[c]:null;
        td.setAttribute('data-testid', (isBase? 'xlsx-preview-base-cell-' : 'xlsx-preview-target-cell-')+String(r)+'-'+String(c));
        if(cell && !cell._phantom){ let v=cell.v; if(cell.f){ td.classList.add('cell-formula'); v='='+cell.f; } if(v===''||v==null){ td.classList.add('cell-empty'); v=''; } td.textContent=v; }
        else { td.classList.add('cell-empty'); td.textContent=''; }
    const set=cellTypes[r+':'+c];
        if(set && set.size){
          const ordered=(window.CellVisuals && CellVisuals.orderCategories(set))||Array.from(set);
          const baseCat=window.CellVisuals? CellVisuals.computeBaseCategory(ordered):null;
          if(baseCat){ td.classList.add('diff-cell-'+baseCat.replace(/[^A-Za-z0-9]/g,'')); }
          const stripe=window.CellVisuals && CellVisuals.computeStripeStyle(ordered); if(stripe){ td.classList.add('multi-cat'); td.style.backgroundImage=stripe; }
          // Tooltip: categories listed vertically (previous legacy behavior)
          try { td.title = ordered.map(catLabel).join('\n'); } catch(_e){}
        }
        
        // Apply formatting AFTER diff classes are applied
        if(cell) applyFormatting(td, cell);
        
        tr.appendChild(td);
      }
      return tr;
    }
    const LARGE_THRESHOLD = 1500; // rows
    const BATCH_SIZE = 500; // rows per async chunk (увеличено с 300 до 500 для быстрого рендеринга)
    if(visibleRows.length <= LARGE_THRESHOLD){ 
      for(let i=0; i<visibleRows.length; i++){ 
        const r = visibleRows[i];
        tb.appendChild(renderRow(r)); 
      } 
    }
    else {
      // initial slice quick render to display something immediately
      const initial = Math.min(1000, visibleRows.length); // Увеличено с 600 до 1000
      for(let i=0; i<initial; i++){ 
        const r = visibleRows[i];
        tb.appendChild(renderRow(r)); 
      }
      let next = initial;
      function pump(){ 
        const frag=document.createDocumentFragment(); 
        let count=0; 
        while(next<visibleRows.length && count<BATCH_SIZE){ 
          const r = visibleRows[next];
          frag.appendChild(renderRow(r)); 
          next++; 
          count++; 
        }
        tb.appendChild(frag);
        if(next<visibleRows.length){ 
          if(window.requestIdleCallback){ requestIdleCallback(pump,{timeout:120}); } 
          else { setTimeout(pump,0); } 
        } else {
          // Логирование завершения рендеринга всех строк
          if(window.Logger) Logger.info(`Preview rendering complete: ${next} rows rendered`);
        }
      }
      pump();
    }
    return tbl;
  }
  baseDiv.innerHTML=''; targetDiv.innerHTML=''; const baseTable = buildTable(baseSnap,true,rCountBase,cCountBase); const targetTable = buildTable(targetSnap,false,rCountTarget,cCountTarget);
  if(!baseTable.tBodies[0].rows.length && window.Logger) Logger.warn('Preview base empty', {rCountBase,cCountBase,maxR:baseSnap.maxR});
  if(!targetTable.tBodies[0].rows.length && window.Logger) Logger.warn('Preview target empty', {rCountTarget,cCountTarget,maxR:targetSnap.maxR});
  baseDiv.appendChild(baseTable); targetDiv.appendChild(targetTable);
  wrap.classList.remove('hidden');
  // (previewControlsRow removed; sync checkbox lives in diffSearchBar now)

  // Apply stored column widths & attach resizers
  try { applyWidthsToTables(baseTable,targetTable); recalcTablesWidth(baseTable,targetTable); attachColumnResizers(baseTable,targetTable); recalcTablesWidth(baseTable,targetTable); } catch(_e){}

  // === Dynamic height logic (auto-fit up to 10 rows; <10 rows exact; splitter expands beyond) ===
  try {
    const regionPreviews = document.getElementById('regionPreviews');
    if(regionPreviews){
      const sampleTable = baseDiv.querySelector('table') || targetDiv.querySelector('table');
      let headerH = 36; let rowH = 18; // conservative defaults
      if(sampleTable){
        const thead = sampleTable.tHead?.rows?.[0];
        if(thead){ const r=thead.getBoundingClientRect(); if(r.height) headerH=Math.round(r.height); }
        const firstBodyRow = sampleTable.tBodies?.[0]?.rows?.[0];
        if(firstBodyRow){ const rb=firstBodyRow.getBoundingClientRect(); if(rb.height) rowH=Math.round(rb.height); }
      }
  const maxRenderedRows = Math.max(rCountBase, rCountTarget);
  // Title (preview caption) height added so inner grid has full 10-row space
  const titleEl = document.querySelector('#sheetPreviewWrapper .preview-title');
  const titleH = titleEl? Math.round(titleEl.getBoundingClientRect().height||20) : 20;
      // Detect horizontal scrollbar need (after first layout flush)
      let needsHScroll = false;
      try { needsHScroll = (baseDiv.scrollWidth>baseDiv.clientWidth) || (targetDiv.scrollWidth>targetDiv.clientWidth); } catch(_e){}
  // Allowance reserves vertical space for potential horizontal scrollbar + breathing space
  const scrollbarGuess =  (navigator.platform && /win/i.test(navigator.platform)) ? 18 : 16; // crude heuristic
  const buffer = 6; // small visual breathing gap
  const baseAllowance = needsHScroll ? (scrollbarGuess + buffer) : buffer;
  const exactHeightCompact = titleH + headerH + (maxRenderedRows * rowH) + baseAllowance; // precise for <10 rows
  // Baseline for 10 rows: title + header + 10*rowH + scrollbar allowance ALWAYS
  const baseline10 = titleH + headerH + (10 * rowH) + (scrollbarGuess + buffer);
      const minHeight = (maxRenderedRows < 10) ? exactHeightCompact : baseline10;
      window.__previewMinRowsHeight = minHeight;
      let initialHeight = minHeight;
      if(maxRenderedRows >= 10){
        // Apply stored height only if user previously resized in this session
        if(window.__splitterUserResized){
          try {
            const raw = localStorage.getItem('compareSheets_splitterH_v1');
            if(raw){ const v=parseInt(raw,10); if(!isNaN(v) && v>baseline10){ initialHeight = v; } }
          } catch(_e){}
        }
      } else {
        // New behavior: no compact auto-fit; previews flex-fill region so splitter always governs visible area.
        // We still set a minimal initial height equal to exactHeightCompact to avoid collapsing.
      }
      regionPreviews.style.flex = '0 0 '+initialHeight+'px';
      regionPreviews.style.height = initialHeight+'px';
      // Restore scroll-based row viewing: ensure grids get proper viewport height and vertical scroll enabled when >10 rows
      setTimeout(()=>{
        const cols = regionPreviews.querySelectorAll('.sheet-preview-col');
        cols.forEach(col=>{
          const title = col.querySelector('.preview-title');
          const grid = col.querySelector('.sheet-grid');
          if(!grid) return;
          const tbl = grid.querySelector('table');
          const bodyRows = tbl && tbl.tBodies && tbl.tBodies[0] ? tbl.tBodies[0].rows.length : 0;
          if(bodyRows <= 10){
            // Fit height to content (no extra whitespace) but not exceeding region
            grid.style.overflowY = 'auto';
            grid.style.maxHeight = '100%';
            return;
          }
          const regionH = regionPreviews.getBoundingClientRect().height;
          const titleHNow = title? (title.getBoundingClientRect().height||18):18;
          const scrollbarGuess = (navigator.platform && /win/i.test(navigator.platform)) ? 18 : 16; // reserve for horizontal bar
          const inner = Math.max(60, regionH - titleHNow - 6 - scrollbarGuess);
          grid.style.height = inner+'px';
          grid.style.maxHeight = inner+'px';
          grid.style.overflowY='auto';
          grid.style.overflowX='auto';
        });
      },0);
      // If stored height was below min (rare), ignore; if rows <10 we intentionally ignore stored expansion
    }
  } catch(_e){}
     return true;
   } catch(e){ if(window.Logger) Logger.error('buildPreviews failed', e); }
 }); }
 window.PreviewRenderer = { buildPreviews };
})(window);
