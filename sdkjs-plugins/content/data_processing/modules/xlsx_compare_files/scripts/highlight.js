/**
 * @fileoverview Sheet highlight application & clearing
 * @description Extracted from plugin.js to satisfy size/function limits.
 */
(function(window){
 'use strict';
 const HIGHLIGHT_LIMIT = 2000;
 const PRIORITY = Object.freeze(['formula','formulaToValue','valueToFormula','value','type','format','inserted-row','deleted-row','inserted-col','deleted-col']);
 let _highlightCache=null;
 function _safeRangeSetFill(cell, color){ try { if(cell && typeof cell.SetFillColor === 'function') cell.SetFillColor(color); } catch(e){} }
 function applyHighlights(){ const st=State.get(); if(!st.diff || !st.diff.raw.length){ _push('warn', tr('msg.noDiff','No diff')); return; } if(_highlightCache){ _push('info', tr('msg.highlightsAlready','Already highlighted')); return; }
   const targetSheet = st.target || st.base; if(!targetSheet){ _push('error', tr('msg.noSheet','No sheet')); return; }
   const slice=st.diff.raw.slice(0,HIGHLIGHT_LIMIT);
   const cellMap=new Map(); slice.forEach(d=>{ if(d.r==null||d.c==null) return; const key=d.r+':'+d.c; let set=cellMap.get(key); if(!set){ set=new Set(); cellMap.set(key,set);} set.add(d.type); });
   Logger.info('Apply highlights grouped cells', cellMap.size);
   if(!(window.Asc&&window.Asc.plugin&&typeof window.Asc.plugin.callCommand==='function')){ _push('error','API not ready'); return; }
   window.Asc.plugin.callCommand(function(){ try {
     var sh=Api.GetSheet(targetSheet); if(!sh){ return; }
     var cache=[]; var colors={ value:Api.CreateColorFromRGB(255,255,153), formula:Api.CreateColorFromRGB(187,222,251), formulaToValue:Api.CreateColorFromRGB(243,229,245), valueToFormula:Api.CreateColorFromRGB(255,243,224), type:Api.CreateColorFromRGB(236,239,241), format:Api.CreateColorFromRGB(207,216,220), 'inserted-row':Api.CreateColorFromRGB(200,230,201), 'deleted-row':Api.CreateColorFromRGB(255,205,210), 'inserted-col':Api.CreateColorFromRGB(197,225,165), 'deleted-col':Api.CreateColorFromRGB(239,154,154) };
     cellMap.forEach(function(set,key){ try { var p=key.split(':'); var r=parseInt(p[0],10), c=parseInt(p[1],10); var cell=sh.GetCells(r,c); var prevFill=null; try { prevFill=cell.GetFill && cell.GetFill(); }catch(e){}
       var chosen=PRIORITY.find(pv=>set.has(pv)) || 'value'; if(cell && typeof cell.SetFillColor==='function'){ cell.SetFillColor(colors[chosen]||colors.value); }
       cache.push({r:r,c:c,orig:{fill:prevFill}});
     } catch(e){} });
     Asc.scope._hlCache={ sheet:targetSheet, cells:cache };
    } catch(e){ }
   }, function(){ if(Asc.scope._hlCache){ _highlightCache=Asc.scope._hlCache; State.set({ highlightApplied:true }); _push('info', tr('msg.highlightApplied','Highlights applied')); } else { _push('error', tr('msg.highlightFailed','Highlight failed')); } });
 }
 function clearHighlights(){ if(!_highlightCache){ _push('info', tr('msg.noHighlights','No highlights to clear')); return; }
   if(!(window.Asc&&window.Asc.plugin&&typeof window.Asc.plugin.callCommand==='function')){ _push('error','API not ready'); return; }
   Logger.info('Clear highlights', _highlightCache.cells.length);
   window.Asc.plugin.callCommand(function(){ try { var sh=Api.GetSheet(_highlightCache.sheet); if(!sh){ return; } _highlightCache.cells.forEach(function(ent){ try { var cell=sh.GetCells(ent.r,ent.c); if(ent.orig.fill){ if(typeof cell.SetFill === 'function') cell.SetFill(ent.orig.fill); else if(typeof cell.SetFillColor==='function') cell.SetFillColor(Api.CreateColorFromRGB(255,255,255)); } else if(typeof cell.SetFillColor==='function'){ cell.SetFillColor(Api.CreateColorFromRGB(255,255,255)); } }catch(e){} }); } catch(e){} }, function(){ _highlightCache=null; State.set({ highlightApplied:false }); _push('info', tr('msg.highlightsCleared','Highlights cleared')); }); }
 function tr(k,f){ const el=document.querySelector(`[data-i18n="${k}"]`); return el? el.textContent: (f||k); }
 function _push(kind,text){ if(window.PluginActions && window.PluginActions._pushUserMessage){ window.PluginActions._pushUserMessage(kind,text); } }
 window.HighlightActions = { applyHighlights, clearHighlights };
})(window);
