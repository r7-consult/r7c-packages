(function(window){
 'use strict';
 const U = window.Utils || {};
 function renderDiffList(diff){ const cont=document.getElementById('diffContainer'); if(!cont) return; cont.innerHTML=''; cont.classList.add('fixed-header'); const sb=document.getElementById('diffSearchBar'); if(!diff){ if(sb) sb.classList.add('hidden'); return; } else { if(sb) sb.classList.remove('hidden'); }
 function trLocal(k,fb){ if(window.__i18n && window.__i18n[k]) return window.__i18n[k]; const el=document.querySelector(`[data-i18n="${k}"]`); return el? el.textContent.trim(): (fb||k); }
 function stripExt(n){ return n? n.replace(/\.[^.]+$/,''): ''; }
 let fromLabel=trLocal('report.header.from','From'); let toLabel=trLocal('report.header.to','To');
 try { const st=(window.State && State.get && State.get())? State.get(): {}; const ph1=(window.__i18n && window.__i18n['placeholder.source1'])||'Src1'; const ph2=(window.__i18n && window.__i18n['placeholder.source2'])||'Src2'; const baseWB=stripExt(st.baseOriginalFile || (diff.meta && (diff.meta.baseWorkbook||diff.meta.base)) || ph1); const targetWB=stripExt(st.targetOriginalFile || (diff.meta && (diff.meta.targetWorkbook||diff.meta.target)) || ph2); const baseSheet=document.getElementById('baseSheetSelect')?.value || (diff.meta.baseSheet||diff.meta.baseSheetName)||''; const targetSheet=document.getElementById('targetSheetSelect')?.value || (diff.meta.targetSheet||diff.meta.targetSheetName)||''; fromLabel=baseWB+(baseSheet?' | '+baseSheet:''); toLabel=targetWB+(targetSheet?' | '+targetSheet:''); } catch(_e){}
 const header=document.createElement('div'); header.className='diff-header'; header.innerHTML=`<div class="col type" data-i18n="report.header.type">${trLocal('report.header.type','Type')}</div>`+
   `<div class="col addr" data-i18n="report.header.cell">${trLocal('report.header.cell','Cell')}</div>`+
   `<div class="col from" data-i18n="report.header.from" title="${trLocal('report.header.from','From')}">${fromLabel}</div>`+
   `<div class="col to" data-i18n="report.header.to" title="${trLocal('report.header.to','To')}">${toLabel}</div>`;
 cont.appendChild(header); const scroll=document.createElement('div'); scroll.className='diff-scroll'; cont.appendChild(scroll); const bodyWrap=document.createElement('div'); bodyWrap.className='diff-rows'; scroll.appendChild(bodyWrap);
 // Grouping logic
 const items = (diff.raw||[]).slice(0,5000);
 const groupByType = window.__diffGroupByType !== false; // default on
 if(!groupByType){
   items.forEach((d,i)=> addRow(d,i,bodyWrap));
 } else {
   // Map diff type -> legend label (with i18n key + fallback)
   function getTypeLabel(t){
     switch(t){
       case 'value': return trLocal('legend.valueChange','value');
       case 'formula': return trLocal('legend.formulaChange','formula');
       case 'inserted-row': return trLocal('legend.insertedRow','inserted row');
       case 'deleted-row': return trLocal('legend.deletedRow','deleted row');
       case 'formulaToValue': return trLocal('legend.formulaToValue','formula→value');
       case 'valueToFormula': return trLocal('legend.valueToFormula','value→formula');
       case 'type': return trLocal('legend.typeChange','type changed');
       case 'format': return trLocal('legend.formatChange','format changed');
       case 'inserted-col': return trLocal('legend.insertedCol','inserted col');
       case 'deleted-col': return trLocal('legend.deletedCol','deleted col');
       default: return t;
     }
   }
  const foundLabel = trLocal('label.found','Found:');
   const groups=new Map();
   items.forEach((d,i)=>{ const t=d.type; if(!groups.has(t)) groups.set(t,[]); groups.get(t).push({d,i}); });
   const order=['inserted-row','deleted-row','inserted-col','deleted-col','formula','formulaToValue','valueToFormula','value','type','format'];
   order.filter(t=>groups.has(t)).forEach(t=>{
  const section=document.createElement('div'); const tClass=t.replace(/[^A-Za-z0-9]/g,''); section.className='diff-group diff-group-'+tClass; section.setAttribute('data-diff-type', t); section.setAttribute('data-testid', 'xlsx-diff-group-'+tClass);
  const headerRow=document.createElement('div'); headerRow.className='diff-group-header diff-group-header-'+tClass; section.classList.add('collapsed');
     const arr=groups.get(t);
     headerRow.innerHTML=`<span class="dg-toggle" data-open="0">▸</span><span class="dg-title">${getTypeLabel(t)}</span><span class="dg-count">${foundLabel} ${arr.length}</span>`;
     headerRow.setAttribute('data-testid', 'xlsx-diff-group-header-'+tClass);
     headerRow.addEventListener('click',()=>{ const open = headerRow.querySelector('.dg-toggle').getAttribute('data-open')==='1'; headerRow.querySelector('.dg-toggle').textContent=open?'▸':'▾'; headerRow.querySelector('.dg-toggle').setAttribute('data-open', open?'0':'1'); section.classList.toggle('collapsed', open); });
     section.appendChild(headerRow);
     const inner=document.createElement('div'); inner.className='diff-group-body'; section.appendChild(inner);
  // Flat list (no sub-groups for value diffs)
  arr.forEach(({d,i})=> addRow(d,i,inner));
     bodyWrap.appendChild(section);
   });
 }
 function addRow(d,i,container){ const row=document.createElement('div'); const cssType=d.type==='value'? 'value-change': (d.type==='formula'? 'formula-change': d.type); row.className='diff-row '+cssType; row.setAttribute('data-diff-index', i); row.dataset.diffType=d.type; row.setAttribute('data-diff-type', d.type); row.setAttribute('data-testid','xlsx-diff-row-'+i);
  // Derive address display using target coordinate (d.r) to preserve existing behavior
  let addr= U.toA1? U.toA1(d.r,d.c): (d.r!=null&&d.c!=null? ('C'+d.c+'R'+d.r):'');
  const rowPrefix = (window.__i18n && window.__i18n['diff.addr.rowPrefix']) || 'Row ';
  const colPrefix = (window.__i18n && window.__i18n['diff.addr.colPrefix']) || 'Col ';
  if(d.type==='inserted-row'||d.type==='deleted-row'){ addr=rowPrefix+(d.r+1);} else if(d.type==='inserted-col'||d.type==='deleted-col'){ addr=colPrefix+(U.colName? U.colName(d.c): d.c); }
  
  // For 'format' type, allow HTML (for bold highlighting), otherwise escape
  let fromHtml, toHtml;
  if(d.type === 'format') {
    fromHtml = d.from || '';
    toHtml = d.to || '';
  } else {
    fromHtml = U.escapeHtml ? U.escapeHtml(U.fmt ? U.fmt(d.from) : d.from) : d.from;
    toHtml = U.escapeHtml ? U.escapeHtml(U.fmt ? U.fmt(d.to) : d.to) : d.to;
  }
  
  row.innerHTML=`<div>${d.type}</div><div>${addr}</div><div>${fromHtml}</div><div>${toHtml}</div>`;
  // Store both base and target row indices when available for precise preview scrolling
  if(d.rBase!=null) row.dataset.rBase = d.rBase;
  if(d.rTarget!=null) row.dataset.rTarget = d.rTarget;
  if(d.r!=null && d.c!=null){ row.dataset.r=d.r; row.dataset.c=d.c; const goToTpl=(window.__i18n && window.__i18n['diff.addr.goTo'])||'Go to {ADDR}'; row.title=goToTpl.replace('{ADDR}', addr); row.tabIndex=0; row.setAttribute('role','button'); }
  container.appendChild(row); }
 if((diff.raw||[]).length>5000){ const more=document.createElement('div'); more.className='diff-row more'; more.setAttribute('data-testid','xlsx-diff-truncated-notice'); const lbl=document.querySelector('[data-i18n="msg.truncated"]')?.textContent || 'Only first 5000'; more.textContent=lbl.replace('{extra}', diff.raw.length-5000); bodyWrap.appendChild(more);} if(window.__diffSearchState && window.__diffSearchState.query){ if(window.refreshDiffSearchMatches) window.refreshDiffSearchMatches(); }
 }
 // Utility helpers now centralized in Utils (colName, toA1, fmt, escapeHtml)
 window.DiffListRenderer={ renderDiffList };
})(window);
