(function(window){
  'use strict';
  if(window.AlignmentHeuristic) return;
  /**
   * Heuristic row/column alignment utilities (extracted from DiffEngine Sprint 1).
   * Pure functions – no DOM or plugin state dependencies.
   */
  function buildRowHashes(snap){
    const out=[]; if(!snap) return out;
    for(let r=0;r<(snap.maxR||0);r++){
      const row=snap.rows[r]||[];
      let nonEmpty=0; const parts=[];
      for(let c=0;c<row.length;c++){
        const cell=row[c]; if(!cell){ parts.push(''); continue; }
        let val='';
        if(cell.f){ val='='+String(cell.f).replace(/\s+/g,'').toUpperCase(); }
        else if(cell.v!=null && cell.v!==''){ val=String(cell.v).trim(); }
        parts.push(val);
        if(val!=='') nonEmpty++;
        if(nonEmpty>12) break; // cap width
      }
      out[r]=parts.join('\u0001');
    }
    return out;
  }
  function buildColHashes(snap){
    const out=[]; if(!snap) return out;
    for(let c=0;c<(snap.maxC||0);c++){
      const parts=[];
      for(let r=0;r<(snap.maxR||0);r++){
        const cell = snap.rows[r]? snap.rows[r][c]: null;
        // Use only VALUES for column hashes, not formulas (formulas contain cell references that change)
        parts.push(cell? (cell.v==null?'':String(cell.v)):'');
      }
      out[c]=parts.join('\u0001');
    }
    return out;
  }
  /** Generic alignment with bounded look-ahead window. */
  function align(baseHashes, targetHashes, opts){
    const {insertType, deleteType, markInserted, markDeleted, windowSize, isCat, collectPair} = opts;
    let i=0,j=0; const nA=baseHashes.length, nB=targetHashes.length;
    
    // Helper: calculate similarity between two column hashes (% of matching rows)
    function similarity(hashA, hashB) {
      if (hashA === hashB) return 1.0;
      const partsA = hashA.split('\u0001');
      const partsB = hashB.split('\u0001');
      const len = Math.min(partsA.length, partsB.length);
      if (len === 0) return 0;
      let matches = 0;
      for (let i = 0; i < len; i++) {
        if (partsA[i] === partsB[i]) matches++;
      }
      return matches / len;
    }
    
    // Debug logging disabled
    const isDebug = false;
    
    const SIMILARITY_THRESHOLD = 0.7; // 70% similarity required
    
    // Determine expected operation based on array lengths
    const expectDeletion = nA > nB;
    const expectInsertion = nB > nA;
    
    // For columns, lower threshold since formulas may change but values might be similar
    // Also, when rows are deleted/inserted, column hashes change (fewer/more values)
    const isColAlignment = (insertType === 'inserted-col' || deleteType === 'deleted-col');
    const effectiveThreshold = isColAlignment ? 0.4 : SIMILARITY_THRESHOLD;
    
    while(i<nA && j<nB){
      const sim = similarity(baseHashes[i], targetHashes[j]);
      if(sim >= effectiveThreshold){ 
        collectPair && collectPair(i,j); i++; j++; continue; 
      }
      let foundT=-1, bestSimT = 0;
      for(let k=1;k<=windowSize && j+k<nB;k++){ 
        const s = similarity(baseHashes[i], targetHashes[j+k]);
        if(s >= effectiveThreshold && s > bestSimT){ foundT=j+k; bestSimT=s; }
      }
      let foundA=-1, bestSimA = 0;
      for(let k=1;k<=windowSize && i+k<nA;k++){ 
        const s = similarity(targetHashes[j], baseHashes[i+k]);
        if(s >= effectiveThreshold && s > bestSimA){ foundA=i+k; bestSimA=s; }
      }
      
      if(foundT!==-1 && isCat(insertType)){
        for(let y=j;y<foundT;y++){ 
          markInserted(y); 
        }
        j=foundT; 
        const simAfter = similarity(baseHashes[i], targetHashes[j]);
        if(simAfter >= effectiveThreshold){ collectPair && collectPair(i,j); i++; j++; }
      } else if(foundA!==-1 && isCat(deleteType)){
        for(let x=i;x<foundA;x++){ 
          markDeleted(x); 
        }
        i=foundA; 
        const simAfter = similarity(baseHashes[i], targetHashes[j]);
        if(simAfter >= effectiveThreshold){ collectPair && collectPair(i,j); i++; j++; }
      } else {
        // No match found in window for either direction
        // Check if we should mark as deleted/inserted based on:
        // 1. Expected operation (deletion/insertion based on length difference)
        // 2. Very low similarity (elements are truly different, not just modified)
        // 3. For small differences (1-2 elements), be more conservative
        const lengthDiff = Math.abs(nA - nB);
        const isSmallDiff = lengthDiff <= 2;
        const simThreshold = isSmallDiff ? 0.2 : 0.3; // stricter threshold for small diffs
        
        if(expectDeletion && isCat(deleteType) && sim < simThreshold) {
          // Very low similarity AND base is longer - likely a true deletion
          markDeleted(i);
          i++;
        } else if(expectInsertion && isCat(insertType) && sim < simThreshold) {
          // Very low similarity AND target is longer - likely a true insertion
          markInserted(j);
          j++;
        } else {
          // Otherwise treat as a modified pair or skip both
          collectPair && collectPair(i,j); 
          i++; j++; 
        }
      }
    }
    if(isCat(deleteType)) {
      while(i<nA && j>=nB){ 
        markDeleted(i); i++; 
      }
    }
    if(isCat(insertType)) {
      while(j<nB && i>=nA){ 
        markInserted(j); j++; 
      }
    }
  }
  window.AlignmentHeuristic = { buildRowHashes, buildColHashes, align };
})(window);
