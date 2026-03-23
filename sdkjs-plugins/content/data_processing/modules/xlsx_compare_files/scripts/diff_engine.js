/** Diff engine implementation */
(function(window){
 'use strict';
 class DiffEngine {
   constructor(opts){ this.opts=opts||{}; this._cancelRef={canceled:false}; }
   _isFormatEnabled(){
     if(!this.opts) return false;
     if(this.opts.compareFormatting) return true;
     if(Array.isArray(this.opts.enabledCategories)) return this.opts.enabledCategories.includes('format');
     return false;
   }
   cancel(){ this._cancelRef.canceled=true; }
  // Tokenizing normalization (functions & identifiers uppercased, whitespace stripped)
  _normFormula(f){
    if(f==null) return f;
    const raw=String(f);
    // Simple tokenizer: split by non-word boundaries while preserving cell refs & numbers
    // Uppercase words (function names / refs) but keep quoted strings
    let out='';
    let i=0; while(i<raw.length){
      const ch=raw[i];
      if(ch==='"' || ch==="'"){
        // string literal – copy verbatim
        const quote=ch; let j=i+1; while(j<raw.length && raw[j]!==quote){ if(raw[j]==='\\') j++; j++; }
        out+=raw.slice(i,j<raw.length? j+1: raw.length); i=j<raw.length? j+1: raw.length; continue;
      }
      if(/[A-Za-z_]/.test(ch)){
        let j=i+1; while(j<raw.length && /[A-Za-z0-9_.]/.test(raw[j])) j++;
        out+= raw.slice(i,j).toUpperCase(); i=j; continue;
      }
      if(/\s/.test(ch)){ i++; continue; }
      out+=ch; i++;
    }
    return out;
  }
  _rawFormula(f){ return f==null? null: String(f); }
  async snapshotSheet(sheetName){
    if(window.VirtualSheets && window.VirtualSheets[sheetName]){ 
      const vs=window.VirtualSheets[sheetName];
      const fmtOn=this._isFormatEnabled();
      
      const rows = vs.rows.map(row=> row.map(cell=>{
        // Always process cell to ensure it has proper type
        const v = cell.v;
        let t = cell.t;
        
        // Infer type if missing
        if(!t){
          if(v===''||v==null){ 
            t='empty'; 
          } else if(!isNaN(Number(v))){ 
            t='number'; 
          } else if(typeof v==='string' && /^\d{4}-\d{2}-\d{2}$/.test(v)){ 
            t='date'; 
          } else { 
            t=typeof v; 
          }
        }
        
        // Build cell object with type
        let processedCell = {...cell, t};
        
        // Convert string numbers to actual numbers if needed
        if(t === 'number' && typeof v === 'string'){
          processedCell.v = Number(v);
        }
        
        // ИЗМЕНЕНО: Всегда включаем форматирование в snapshot (если есть)
        // Решение о показе будет приниматься в preview_renderer.js
        if(cell.fm){
          processedCell.fm = cell.fm;
        }
        
        return processedCell;
      }));
      
      // Ensure maxR and maxC are valid numbers
      const maxR = (vs.maxR != null && vs.maxR > 0) ? vs.maxR : rows.length;
      const maxC = (vs.maxC != null && vs.maxC > 0) ? vs.maxC : (rows[0] ? rows[0].length : 0);
      
      return { name:sheetName, rows, maxR, maxC, virtual:true }; 
    }
     // Preferred modern API snapshot (may already include formatting signature if provided by host)
     if(window.R7API && window.R7API.getSheetSnapshot){ try { return await window.R7API.getSheetSnapshot(sheetName, { includeFormatting:this._isFormatEnabled() }); } catch(_e){} }
     // fallback (legacy)
     return new Promise((resolve,reject)=>{ window.Asc.plugin.callCommand(function(){ try { var sheet = Api.GetSheet(sheetName); if(!sheet){ Asc.scope._snapErr='SHEET_NOT_FOUND'; return; } var usedRange=sheet.GetUsedRange(); var maxR=usedRange?usedRange.GetRowCount():0; var maxC=usedRange?usedRange.GetColCount():0; var rows=[]; for(var r=0;r<maxR;r++){ var rowArr=[]; for(var c=0;c<maxC;c++){ var cell=sheet.GetCells(r,c); var val=null,f=null,t=null; try{f=cell.GetFormula&&cell.GetFormula();}catch(e){} try{val=cell.GetValue&&cell.GetValue();}catch(e){} try{t=cell.GetValueType&&cell.GetValueType();}catch(e){} rowArr.push({v:val,f:f||null,t:t||null}); } rows.push(rowArr);} Asc.scope._sheetSnapshot={ name:sheetName, rows:rows, maxR:maxR, maxC:maxC }; } catch(e){ Asc.scope._snapErr=e.message; } },()=>{ if(Asc.scope._snapErr) reject(new Error(Asc.scope._snapErr)); else resolve(Asc.scope._sheetSnapshot); }); });
   }
   _normString(str){ if(str==null) return str; if(this.opts.trim && typeof str==='string') str=str.trim(); if(this.opts.ignoreCase && typeof str==='string') str=str.toLowerCase(); return str; }
  _compareValues(a,b){
    if(a===b) return true;
    if(typeof a==='number' && typeof b==='number'){
      const eps = (this.opts && typeof this.opts.numericEpsilon==='number')? this.opts.numericEpsilon: 0;
      if(eps>0 && Math.abs(a-b) <= eps) return true;
      return a===b;
    }
    return false;
  }

  /**
   * Convert format signature to human-readable text
   * @param {string} sig - Format signature like "b:1|i:1|fn:Arial|fs:12|fc:#FF0000|bg:#FFFF00"
   * @param {string} compareSig - Optional signature to compare against (to highlight differences)
   * @returns {string} Human-readable description with <b> tags around differences
   */
  _formatSignatureToText(sig, compareSig) {
    if (!sig || sig === '') return '';
    
    const parts = sig.split('|');
    const props = {};
    
    // Parse signature into properties
    for (const part of parts) {
      const [key, val] = part.split(':');
      if (key && val !== undefined) {
        props[key] = val;
      }
    }
    
    // Parse compare signature if provided
    let compareProps = null;
    if (compareSig) {
      compareProps = {};
      const compareParts = compareSig.split('|');
      for (const part of compareParts) {
        const [key, val] = part.split(':');
        if (key && val !== undefined) {
          compareProps[key] = val;
        }
      }
    }
    
    // Helper to check if property differs
    const differs = (key) => {
      if (!compareProps) return false;
      return props[key] !== compareProps[key];
    };
    
    // Helper to wrap in <b> if differs
    const maybeWrap = (text, isDifferent) => {
      return isDifferent ? `<b>${text}</b>` : text;
    };
    
    // Helper to get translation
    const _t = (key, fallback) => {
      if (window.I18n && typeof window.I18n.t === 'function') {
        return window.I18n.t(key, null, fallback);
      }
      return (window.__i18n && window.__i18n[key]) || fallback;
    };
    
    const result = [];
    
    // Bold
    if (props.b === '1') {
      result.push(maybeWrap(_t('format.bold', 'Bold'), differs('b')));
    }
    
    // Italic
    if (props.i === '1') {
      result.push(maybeWrap(_t('format.italic', 'Italic'), differs('i')));
    }
    
    // Underline
    if (props.u === '1') {
      result.push(maybeWrap(_t('format.underline', 'Underline'), differs('u')));
    }
    
    // Font name and size
    if (props.fn || props.fs) {
      const fontDiffers = differs('fn') || differs('fs');
      let fontDesc = _t('format.font', 'Font:');
      if (props.fn) fontDesc += ' ' + props.fn;
      if (props.fs) fontDesc += ' ' + props.fs + 'pt';
      result.push(maybeWrap(fontDesc.trim(), fontDiffers));
    }
    
    // Font color
    if (props.fc && props.fc !== '#000000') {
      result.push(maybeWrap(_t('format.textColor', 'Text color:') + ' ' + props.fc, differs('fc')));
    }
    
    // Background color
    if (props.bg && props.bg !== '#FFFFFF' && props.bg !== '') {
      result.push(maybeWrap(_t('format.background', 'Background:') + ' ' + props.bg, differs('bg')));
    }
    
    // Horizontal alignment
    if (props.ah) {
      const alignMap = {
        'left': _t('format.align.left', 'left'),
        'center': _t('format.align.center', 'center'),
        'right': _t('format.align.right', 'right'),
        'justify': _t('format.align.justify', 'justify')
      };
      result.push(maybeWrap(_t('format.alignment', 'Alignment:') + ' ' + (alignMap[props.ah] || props.ah), differs('ah')));
    }
    
    // Vertical alignment
    if (props.av) {
      const valignMap = {
        'top': _t('format.valign.top', 'top'),
        'center': _t('format.valign.center', 'center'),
        'bottom': _t('format.valign.bottom', 'bottom')
      };
      result.push(maybeWrap(_t('format.verticalAlignment', 'Vert. align:') + ' ' + (valignMap[props.av] || props.av), differs('av')));
    }
    
    return result.length > 0 ? result.join(', ') : _t('format.none', 'No formatting');
  }
  _cellFormatSignature(cell){
    if(!cell) return '';
    if(cell.fm) return cell.fm; // precomputed signature from virtual snapshot
    if(!this._isFormatEnabled()) return '';
    // cache
    this._fmtCache = this._fmtCache || new WeakMap();
    if(this._fmtCache.has(cell)) return this._fmtCache.get(cell);
    const sel = (Array.isArray(this.opts.formatFields) && this.opts.formatFields.length)? new Set(this.opts.formatFields.map(f=>String(f).trim())): null;
    function allowed(tag){ return !sel || sel.has(tag); }
    try {
      const parts=[];
      let v;
      if(allowed('nf')){
        if(cell.GetNumberFormat){ try { v=cell.GetNumberFormat()||''; } catch(_e){ v=''; } parts.push('nf:'+v); }
        else if(cell.GetNumberFormatString){ try { v=cell.GetNumberFormatString()||''; } catch(_e){ v=''; } parts.push('nf:'+v); }
      }
      if(allowed('b') && (cell.IsBold || cell.GetBold)){
        try { const b=(cell.IsBold?cell.IsBold(): (cell.GetBold?cell.GetBold():false)); parts.push('b:'+(b?1:0)); } catch(_e){ parts.push('b:'); }
      }
      if(allowed('i') && (cell.IsItalic || cell.GetItalic)){
        try { const it=(cell.IsItalic?cell.IsItalic(): (cell.GetItalic?cell.GetItalic():false)); parts.push('i:'+(it?1:0)); } catch(_e){ parts.push('i:'); }
      }
      if(allowed('u') && (cell.IsUnderline || cell.GetUnderline)){
        try { const u=(cell.IsUnderline?cell.IsUnderline(): (cell.GetUnderline?cell.GetUnderline():false)); parts.push('u:'+(u?1:0)); } catch(_e){ parts.push('u:'); }
      }
      if(allowed('fn') && cell.GetFontName){ try { parts.push('fn:'+(cell.GetFontName()||'')); } catch(_e){ parts.push('fn:'); } }
      if(allowed('fs') && cell.GetFontSize){ try { parts.push('fs:'+(cell.GetFontSize()||'')); } catch(_e){ parts.push('fs:'); } }
      if(allowed('fc') && cell.GetFontColor){ try { parts.push('fc:'+(cell.GetFontColor()||'')); } catch(_e){ parts.push('fc:'); } }
      if(allowed('bg') && cell.GetFillColor){ try { parts.push('bg:'+(cell.GetFillColor()||'')); } catch(_e){ parts.push('bg:'); } }
      if(allowed('ah') && cell.GetAlignHorizontal){ try { parts.push('ah:'+(cell.GetAlignHorizontal()||'')); } catch(_e){ parts.push('ah:'); } }
      if(allowed('av') && cell.GetAlignVertical){ try { parts.push('av:'+(cell.GetAlignVertical()||'')); } catch(_e){ parts.push('av:'); } }
      const sig = parts.join('|');
      this._fmtCache.set(cell, sig);
      return sig;
    } catch(_e){ return ''; }
  }
  async compareSheets(baseName,targetName,progressCb){
     const startTs=Date.now();
     const baseSnap = await this.snapshotSheet(baseName);
     const targetSnap = await this.snapshotSheet(targetName);
  const afterSnapshotTs = Date.now();
  const maxR = Math.max(baseSnap.maxR, targetSnap.maxR);
     let maxC = Math.max(baseSnap.maxC, targetSnap.maxC); // let instead of const for recalculation after phantom insertions
  const enabled = Array.isArray(this.opts.enabledCategories)? this.opts.enabledCategories.filter(Boolean): null; // null => all
  function isCat(cat){ if(!enabled || !enabled.length) return true; return enabled.includes(cat); }
  const diffs=[]; const stats={ value:0, formula:0, formulaToValue:0, valueToFormula:0, type:0, format:0,
    'inserted-row':0,'deleted-row':0,'inserted-col':0,'deleted-col':0,
    // legacy keys kept for backward compatibility with report sheet
    rowInserted:0,rowDeleted:0,colInserted:0,colDeleted:0 };
  // Track structural row/col indices (for skipping cell-level comparisons)
  const structRowsInserted=new Set();
  const structRowsDeleted=new Set();
  const structColsInserted=new Set();
  const structColsDeleted=new Set();
     // === Structural diff (window alignment) ===
     // Alignment functions (extracted; fallback if external module not loaded)
     const AH = (typeof window!=='undefined' && window.AlignmentHeuristic) ? window.AlignmentHeuristic : null;
     const buildRowHashes = AH? AH.buildRowHashes.bind(AH): (snap=>{ const out=[]; for(let r=0;r<snap.maxR;r++){ const row=snap.rows[r]||[]; let nonEmpty=0; const parts=[]; for(let c=0;c<row.length;c++){ const cell=row[c]; if(!cell){ parts.push(''); continue; } let val=''; if(cell.f){ val='='+this._normFormula(cell.f); } else if(cell.v!=null && cell.v!==''){ val=String(cell.v).trim(); } parts.push(val); if(val!=='') nonEmpty++; if(nonEmpty>12) break; } out[r]=parts.join('\u0001'); } return out; });
     const buildColHashes = AH? AH.buildColHashes.bind(AH): (snap=>{ const out=[]; for(let c=0;c<snap.maxC;c++){ const parts=[]; for(let r=0;r<snap.maxR;r++){ const cell=snap.rows[r]?snap.rows[r][c]:null; parts.push(cell? (cell.v==null?'':String(cell.v)):''); } out[c]=parts.join('\u0001'); } return out; });
     const alignGeneric = AH? AH.align : function(baseHashes, targetHashes, opts){
       const {insertType, deleteType, markInserted, markDeleted, windowSize, isCat, collectPair} = opts;
       let i=0,j=0; const nA=baseHashes.length, nB=targetHashes.length;
       while(i<nA && j<nB){
         if(baseHashes[i]===targetHashes[j]){ collectPair && collectPair(i,j); i++; j++; continue; }
         let foundT=-1; for(let k=1;k<=windowSize && j+k<nB;k++){ if(baseHashes[i]===targetHashes[j+k]){ foundT=j+k; break; } }
         let foundA=-1; for(let k=1;k<=windowSize && i+k<nA;k++){ if(targetHashes[j]===baseHashes[i+k]){ foundA=i+k; break; } }
         if(foundT!==-1 && isCat(insertType)){
           for(let y=j;y<foundT;y++){ markInserted(y); }
           j=foundT; if(baseHashes[i]===targetHashes[j]){ collectPair && collectPair(i,j); i++; j++; }
         } else if(foundA!==-1 && isCat(deleteType)){
           for(let x=i;x<foundA;x++){ markDeleted(x); }
           i=foundA; if(baseHashes[i]===targetHashes[j]){ collectPair && collectPair(i,j); i++; j++; }
         } else { collectPair && collectPair(i,j); i++; j++; }
       }
       if(isCat(deleteType)) while(i<nA && j>=nB){ markDeleted(i); i++; }
       if(isCat(insertType)) while(j<nB && i>=nA){ markInserted(j); j++; }
     };
     // Rows alignment
    const rowPairs=[]; // {base, target}
    if(baseSnap.maxR || targetSnap.maxR){ 
      const rhA=buildRowHashes(baseSnap); 
      const rhB=buildRowHashes(targetSnap);
      alignGeneric(rhA,rhB,{ 
        insertType:'inserted-row', 
        deleteType:'deleted-row', 
        windowSize:30, 
        isCat, 
        markInserted:r=>{ 
          if(!structRowsInserted.has(r)){ 
            structRowsInserted.add(r); 
            diffs.push({ type:'inserted-row', r, c:0, from:null, to:'ROW_INSERTED' }); 
            stats['inserted-row']++; 
            stats.rowInserted++; 
          } 
        }, 
        markDeleted:r=>{ 
          if(!structRowsDeleted.has(r)){ 
            structRowsDeleted.add(r); 
            diffs.push({ type:'deleted-row', r, c:0, from:'ROW_DELETED', to:null }); 
            stats['deleted-row']++; 
            stats.rowDeleted++; 
          } 
        }, 
        collectPair:(baseRow, targetRow)=>{ 
          rowPairs.push({base: baseRow, target: targetRow}); 
        }
      });
    }
     // Columns alignment
     const colPairs = []; // {base, target} mapping for column alignment
     if(baseSnap.maxC || targetSnap.maxC){ 
       const chA=buildColHashes(baseSnap); 
       const chB=buildColHashes(targetSnap); 
       alignGeneric(chA,chB,{ 
         insertType:'inserted-col', 
         deleteType:'deleted-col', 
         windowSize:15, 
         isCat, 
         markInserted:c=>{ 
           if(!structColsInserted.has(c)){ 
             structColsInserted.add(c); 
             diffs.push({ type:'inserted-col', r:0, c, from:null, to:'COL_INSERTED' }); 
             stats['inserted-col']++; 
             stats.colInserted++; 
           } 
         }, 
         markDeleted:c=>{ 
           if(!structColsDeleted.has(c)){ 
             structColsDeleted.add(c); 
             diffs.push({ type:'deleted-col', r:0, c, from:'COL_DELETED', to:null }); 
             stats['deleted-col']++; 
             stats.colDeleted++; 
           } 
         }, 
         collectPair:(baseCol, targetCol)=>{ 
           colPairs.push({base: baseCol, target: targetCol}); 
         }
       }); 
     }
  
  const batchSize=500; let processed=0;
     // === Performance helpers (Priority 3) ===
     function isCellEmpty(cell){ return !cell || ((cell.v==null || cell.v==='') && !cell.f); }
     function rowAllEmpty(row){ if(!row||!row.length) return true; for(let k=0;k<row.length;k++){ if(!isCellEmpty(row[k])) return false; } return true; }
     const emptyAlignedRow = new Array(rowPairs.length);
     for(let idx=0; idx<rowPairs.length; idx++){
       const {base:rb, target:rt} = rowPairs[idx];
       const rowA = baseSnap.rows[rb];
       const rowB = targetSnap.rows[rt];
       emptyAlignedRow[idx] = rowAllEmpty(rowA) && rowAllEmpty(rowB);
     }
  let skippedEmptyRows=0, skippedStructuralRows= (structRowsInserted.size + structRowsDeleted.size) * colPairs.length;
  const formatCandidates = []; // lazy second pass
  const totalCells = rowPairs.length * colPairs.length;
  for(let pairIndex=0; pairIndex<rowPairs.length; pairIndex++){
       if(this._cancelRef.canceled) return { canceled:true };
       if(emptyAlignedRow[pairIndex]){ skippedEmptyRows += colPairs.length; processed += colPairs.length; continue; }
       const {base:rb, target:rt} = rowPairs[pairIndex];
       const rowA = baseSnap.rows[rb];
       const rowB = targetSnap.rows[rt];
       // Iterate over aligned column pairs instead of all columns
        for(let colPairIdx=0; colPairIdx<colPairs.length; colPairIdx++){
        const {base: cBase, target: cTarget} = colPairs[colPairIdx];
        const cellA = rowA? rowA[cBase]: undefined;
        const cellB = rowB? rowB[cTarget]: undefined;
        // Use cBase as the canonical column index for diffs (original base column position)
        const c = cBase;
        // Skip cell-level if row/col is structurally inserted/deleted
        if(!structColsInserted.has(cTarget) && !structColsDeleted.has(cBase)){
          if(cellA && cellB){
            const fA = cellA.f; const fB = cellB.f; const hasFA= !!fA; const hasFB= !!fB;
            if(hasFA!==hasFB){
              if(hasFA && !hasFB && isCat('formulaToValue')){ diffs.push({ type:'formulaToValue', r:rt, rTarget:rt, rBase:rb, c, from:this._rawFormula(fA), to:cellB.v, baseFormula:this._rawFormula(fA), targetFormula:null }); stats.formulaToValue++; }
              else if(!hasFA && hasFB && isCat('valueToFormula')){ diffs.push({ type:'valueToFormula', r:rt, rTarget:rt, rBase:rb, c, from:cellA.v, to:this._rawFormula(fB), baseFormula:null, targetFormula:this._rawFormula(fB) }); stats.valueToFormula++; }
            } else if(hasFA && hasFB){
              if(this._normFormula(fA) !== this._normFormula(fB) && isCat('formula')){ diffs.push({ type:'formula', r:rt, rTarget:rt, rBase:rb, c, from:this._rawFormula(fA), to:this._rawFormula(fB), baseFormula:this._rawFormula(fA), targetFormula:this._rawFormula(fB) }); stats.formula++; }
              else if(isCat('value')){
                const vA=this._normString(cellA.v); const vB=this._normString(cellB.v);
                if(!this._compareValues(vA,vB)){ diffs.push({ type:'value', r:rt, rTarget:rt, rBase:rb, c, from:cellA.v, to:cellB.v, note:'formulaResult', baseFormula:this._rawFormula(fA), targetFormula:this._rawFormula(fB) }); stats.value++; }
              }
            } else { // neither formula
              if(isCat('value')){ const vA=this._normString(cellA.v); const vB=this._normString(cellB.v); if(!(vA==null && vB==null) && !this._compareValues(vA,vB)){ diffs.push({type:'value', r:rt, rTarget:rt, rBase:rb, c, from:cellA.v, to:cellB.v}); stats.value++; } }
            }
            if(cellA.t!==cellB.t && isCat('type')){ diffs.push({ type:'type', r:rt, rTarget:rt, rBase:rb, c, from:cellA.t, to:cellB.t }); stats.type++; }
            if(this._isFormatEnabled() && isCat('format')){ formatCandidates.push({rb, rt, c, cellA, cellB}); }
          } else if(isCat('value') && (cellA || cellB)){
            // One-sided cell difference (present only in base or target)
            const fromV = cellA? cellA.v : null; const toV = cellB? cellB.v : null;
            if(!(fromV==null && toV==null)){ diffs.push({ type:'value', r:rt, rTarget:rt, rBase:rb, c, from:fromV, to:toV, note:'oneSide' }); stats.value++; }
          }
        }
        // formatting temporarily disabled
         processed++;
         if(processed % batchSize ===0){ progressCb && progressCb(processed/totalCells); await new Promise(rz=>setTimeout(rz)); }
       }
     }
    // Second pass: compute format diffs lazily
    if(this._isFormatEnabled() && isCat('format')){
      for(let i=0;i<formatCandidates.length;i++){
        const {rb, rt, c, cellA, cellB} = formatCandidates[i];
        
        // Skip format comparison if one of the cells is empty
        const isAEmpty = !cellA || cellA.v == null || cellA.v === '';
        const isBEmpty = !cellB || cellB.v == null || cellB.v === '';
        if(isAEmpty || isBEmpty) {
          continue; // Don't report format change if one cell is empty
        }
        
        const sigA=this._cellFormatSignature(cellA); const sigB=this._cellFormatSignature(cellB);
        if(sigA!==sigB){
          let fieldsDetail=null;
          if(this.opts && this.opts.formatExplain){
            const parse=(sig)=>{ const map={}; sig.split('|').forEach(p=>{ if(!p) return; const idx=p.indexOf(':'); if(idx>0) map[p.slice(0,idx)]=p.slice(idx+1); }); return map; };
            const ma=parse(sigA), mb=parse(sigB); const all=new Set([...Object.keys(ma), ...Object.keys(mb)]); fieldsDetail=[]; all.forEach(k=>{ if(ma[k]!==mb[k]) fieldsDetail.push({ field:k, from:ma[k]||'', to:mb[k]||'' }); });
          }
          
          // Format signatures as human-readable text with highlighting differences
          const fromText = this._formatSignatureToText(sigA, sigB);
          const toText = this._formatSignatureToText(sigB, sigA);
          
          diffs.push({ type:'format', r:rt, rTarget:rt, rBase:rb, c, from:fromText, to:toText, fromRaw:sigA, toRaw:sigB, ...(fieldsDetail?{fields:fieldsDetail}:{}) }); stats.format++;
        }
        if(i % 1000 ===0){ progressCb && progressCb(Math.min(0.99, processed/totalCells)); }
        if(this._cancelRef.canceled) return { canceled:true };
      }
    }
    const afterLoopTs = Date.now();
  // (Tail structural diffs already included in fallback branches above; advanced path adds explicit per-index diffs.)
  progressCb && progressCb(1);
  const endTs = Date.now();
  // Diagnostics & metrics
  const cellsCompared = maxR * maxC;
  const changedCells = stats.value + stats.formula + stats.formulaToValue + stats.valueToFormula + stats.type + stats.format;
  const structuralDiffs = stats['inserted-row'] + stats['deleted-row'] + stats['inserted-col'] + stats['deleted-col'];
  const metrics = {
    cellsCompared,
    changedCells,
    changedCellsPct: cellsCompared ? changedCells / cellsCompared : 0,
    structuralDiffs,
    diffDensity: cellsCompared ? changedCells / cellsCompared : 0,
    baseRows: baseSnap.maxR, baseCols: baseSnap.maxC,
  targetRows: targetSnap.maxR, targetCols: targetSnap.maxC,
  skippedEmptyRowCells: skippedEmptyRows,
  skippedStructuralRowCells: skippedStructuralRows,
  alignedRowPairs: rowPairs.length
  };
  const timings = {
    snapshotMs: afterSnapshotTs - startTs,
    diffLoopMs: afterLoopTs - afterSnapshotTs,
    totalMs: endTs - startTs
  };
  const result = { meta:{ base:baseName, target:targetName, generated:new Date().toISOString(), durationMs: endTs-startTs, metrics, timings, rowPairs }, stats, categories: this._groupCategories(diffs), raw: diffs };
  if(this.opts && this.opts.includeDiagnostics){
    result.meta.diagnostics = {
      alignment:{ rowPairs:rowPairs.length, structRowsInserted:structRowsInserted.size, structRowsDeleted:structRowsDeleted.size, structColsInserted:structColsInserted.size, structColsDeleted:structColsDeleted.size, rowWindow:30, colWindow:15 },
      options:{ numericEpsilon: this.opts.numericEpsilon||0, compareFormatting:this._isFormatEnabled() }
    };
  }
  return result;
   }
   _groupCategories(diffs){
     const catMap={ value:[], formula:[], formulaToValue:[], valueToFormula:[], type:[], format:[], 'inserted-row':[], 'deleted-row':[], 'inserted-col':[], 'deleted-col':[] };
     diffs.forEach(d=>{ if(!catMap[d.type]) catMap[d.type]=[]; catMap[d.type].push(d); });
     return Object.keys(catMap).filter(k=>catMap[k].length).map(k=>({ name:k, items:catMap[k] }));
   }
 }
 window.DiffEngine = DiffEngine;
})(window);
