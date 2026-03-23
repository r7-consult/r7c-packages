(function(window){
  'use strict';
  if(window.PluginAnalysis && window.location.pathname.indexOf('analysis.html')===-1) return; // keep original in main page

  const VOLATILE_FUNCS = ['NOW','TODAY','RAND','RANDBETWEEN','OFFSET','INDIRECT','INFO','CELL','СЕГОДНЯ','ТДАТА','СЛЧИС','СЛУЧМЕЖДУ','СМЕЩ','ДВССЫЛ','ЯЧЕЙКА'];
  const ERROR_MARKERS = ['#DIV/0!','#N/A','#NAME?','#NULL!','#NUM!','#REF!','#VALUE!','#ДЕЛ/0!','#Н/Д','#ИМЯ?','#ПУСТО!','#ЧИСЛО!','#ССЫЛКА!','#ЗНАЧ!'];
  // SheetJS error code map (t==='e' and v numeric) based on Excel error codes
  const ERROR_CODE_MAP = {
    0x00:'#NULL!',
    0x07:'#DIV/0!',
    0x0F:'#VALUE!',
    0x17:'#REF!',
    0x1D:'#NAME?',
    0x24:'#NUM!',
    0x2A:'#N/A',
    0x2B:'#GETTING_DATA'
  };
  // Russian error names mapping (for display)
  const ERROR_RU_MAP = {
    '#DIV/0!':'#ДЕЛ/0!',
    '#N/A':'#Н/Д',
    '#NAME?':'#ИМЯ?',
    '#NULL!':'#ПУСТО!',
    '#NUM!':'#ЧИСЛО!',
    '#REF!':'#ССЫЛКА!',
    '#VALUE!':'#ЗНАЧ!'
  };
  
  function readFileAsArrayBuffer(file){
    return new Promise((resolve,reject)=>{ const fr=new FileReader(); fr.onerror=e=>reject(e); fr.onload=()=>resolve(fr.result); fr.readAsArrayBuffer(file); });
  }

  function parseWorkbook(arrayBuffer){ if(!window.XLSX) throw new Error('XLSX library not loaded'); const data=new Uint8Array(arrayBuffer); return XLSX.read(data,{ type:'array', cellFormula:true }); }

  function eachCell(sheet, cb){ const range=sheet['!ref']; if(!range) return; const [s,e]=range.split(':').map(XLSX.utils.decode_cell); for(let r=s.r;r<=e.r;r++){ for(let c=s.c;c<=e.c;c++){ const addr=XLSX.utils.encode_cell({r,c}); const cell=sheet[addr]; if(cell) cb(r,c,addr,cell); } } }

  function estimateIfDepth(f){ let depth=0,maxDepth=0; for(let i=0;i<f.length;i++){ if((f[i]==='I' && f.slice(i,i+3)==='IF(') || (f[i]==='Е' && f.slice(i,i+5)==='ЕСЛИ(')){ depth++; if(depth>maxDepth) maxDepth=depth; } if(f[i]===')' && depth>0) depth--; } return maxDepth; }
  function normalizeFormula(f){ return f.replace(/\$?[A-Z]{1,3}\$?\d{1,7}(:\$?[A-Z]{1,3}\$?\d{1,7})?/g,'REF').replace(/\d+/g,'N'); }
  function collectSelectedMetrics(){ const map={}; document.querySelectorAll('#metricsList input[type=checkbox]').forEach(cb=>map[cb.value]=cb.checked); return map; }
  function escapeHtml(s){ if(s==null) return ''; return String(s).replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[ch])); }

  function ensureTestIds(){
    try {
      const byId = {
        analysisFileInput: 'xlsx-analysis-file',
        btnStartAnalysis: 'xlsx-analysis-start',
        btnReportAnalysis: 'xlsx-analysis-report',
        metricsList: 'xlsx-analysis-metrics',
        analysisKpiGrid: 'xlsx-analysis-kpis',
        analysisDetails: 'xlsx-analysis-details',
        analysisMetricDetails: 'xlsx-analysis-metric-details',
        amdClose: 'xlsx-analysis-metric-details-close',
        amdTitle: 'xlsx-analysis-metric-details-title',
        amdBody: 'xlsx-analysis-metric-details-body',
        analysisStatusTop: 'xlsx-analysis-status',
        analysisProgressTop: 'xlsx-analysis-progress',
        analysisSummary: 'xlsx-analysis-summary',
        btnAnalyze: 'xlsx-btn-analyze'
      };
      Object.entries(byId).forEach(([id, tid]) => {
        const el = document.getElementById(id);
        if (el && !el.getAttribute('data-testid')) el.setAttribute('data-testid', tid);
      });

      const bar = document.querySelector('#analysisProgressTop span');
      if (bar && !bar.getAttribute('data-testid')) bar.setAttribute('data-testid', 'xlsx-analysis-progress-bar');

      const metrics = document.querySelectorAll('#metricsList input[type=checkbox]');
      metrics.forEach((cb, idx) => {
        if (!cb || cb.getAttribute('data-testid')) return;
        const safe = String(cb.value || idx).replace(/[^A-Za-z0-9_]/g, '_');
        cb.setAttribute('data-testid', 'xlsx-analysis-metric-' + safe);
      });
    } catch(_e) {}
  }

  function analyzeWorkbook(wb, metrics){
  const summary={ sheets:wb.SheetNames.length, totalCells:0, formulaCells:0, formulaCellsFallback:0, errorCells:0, volatileCells:0, longFormulas:0, numericConstants:0, errorTypes:{} };
    const sheets=[];
    // Detail collections for drill-down
    const details={
      formulas:[], // native formulas {sheet,addr,formula}
      formulasText:[], // fallback text formulas
      errors:[], // {sheet,addr,value,type}
      volatile:[], // {sheet,addr,formula}
      long:[], // {sheet,addr,len,formula}
      nestedIf:[], // {sheet,addr,depth,formula}
      numericConstants:[] // {sheet,addr,value,context:'in-formula'|'cell'}
    };

    wb.SheetNames.forEach(name=>{
      const sheet=wb.Sheets[name];
  const sheetInfo={ name,cells:0,formulas:0,formulasFallback:0,errors:0,volatile:0,longFormulas:[],nestedIf:0,numericConstants:0, errorTypes:{} };
      eachCell(sheet,(r,c,addr,cell)=>{
        sheetInfo.cells++; summary.totalCells++;
        const f=cell.f; const v=cell.v;
        // Check for formulas: native (cell.f) or fallback (text starting with '=')
        let isFormula = false;
        let formulaText = '';
        let countedNative = false;
        
        // Priority 1: Native formula (cell.f exists)
        if(f && metrics.formulasCount){
          isFormula = true;
          formulaText = f;
          countedNative = true;
          sheetInfo.formulas++; 
          summary.formulaCells++;
        }
        // Priority 2: Fallback text formula (v starts with '=' when cell.f is absent)
        else if(!f && typeof v==='string' && v.startsWith('=') && metrics.formulasCount){
          isFormula = true;
          formulaText = v.slice(1);
          sheetInfo.formulasFallback++; 
          summary.formulaCellsFallback++;
        }
        if(metrics.errorCells){
          let detected=null;
            // Direct string match (#REF! etc.)
          if(typeof v==='string' && ERROR_MARKERS.some(m=>v.indexOf(m)===0)) detected=ERROR_MARKERS.find(m=>v.indexOf(m)===0);
            // SheetJS error type (t==='e') numeric code
          else if(cell.t==='e'){
            if(typeof v==='number' && ERROR_CODE_MAP[v]) detected=ERROR_CODE_MAP[v];
            else if(typeof v==='string' && ERROR_MARKERS.includes(v)) detected=v; // fallback
            else if(cell.w && ERROR_MARKERS.some(m=>cell.w.indexOf(m)===0)) detected=ERROR_MARKERS.find(m=>cell.w.indexOf(m)===0);
          }
          if(detected){
            sheetInfo.errors++; summary.errorCells++;
            sheetInfo.errorTypes[detected]=(sheetInfo.errorTypes[detected]||0)+1;
            summary.errorTypes[detected]=(summary.errorTypes[detected]||0)+1;
            if(details.errors.length < 10000){ details.errors.push({sheet:name,addr,type:detected,value:(cell.w||cell.v||detected)}); }
          }
        }
        // Process formula metrics if we have a formula
        if(isFormula){ 
          const baseF = formulaText;
          const upperF = baseF.toUpperCase();
          
          // Collect formula details
          if(metrics.formulasCount){
            if(countedNative) details.formulas.push({sheet:name,addr,formula:baseF}); 
            else details.formulasText.push({sheet:name,addr,formula:baseF});
          }
          
          // Check for volatile functions
          if(metrics.volatileFunctions && VOLATILE_FUNCS.some(fn=>upperF.includes(fn+'('))){ 
            sheetInfo.volatile++; 
            summary.volatileCells++; 
            details.volatile.push({sheet:name,addr,formula:baseF}); 
          }

          // Check for long formulas (>100 chars)
          if(metrics.longFormulas && baseF.length>100){ 
            const lf={addr,len:baseF.length,formula:baseF}; 
            sheetInfo.longFormulas.push(lf); 
            summary.longFormulas++; 
            details.long.push({sheet:name,addr,len:lf.len,formula:baseF}); 
          }
          
          // Check for nested IF (depth > 3)
          if(metrics.nestedIf){ 
            const d=estimateIfDepth(upperF); 
            if(d>3){ 
              sheetInfo.nestedIf++; 
              details.nestedIf.push({sheet:name,addr,depth:d,formula:baseF}); 
            } 
          }
          
          // Check for numeric constants in formulas
          if(metrics.numericConstants && /(^|[+\-*/,( ])\d+(?:\.\d+)?/.test(baseF)){ 
            sheetInfo.numericConstants++; 
            details.numericConstants.push({sheet:name,addr,value:null,context:'in-formula',formula:baseF}); 
          }
        }
        if(metrics.numericConstants && !isFormula && typeof v==='number'){ sheetInfo.numericConstants++; summary.numericConstants++; details.numericConstants.push({sheet:name,addr,value:v,context:'cell'}); }
        if(metrics.errorCells && sheetInfo.errors && details.errors.length < 5000){ /* already counted above if detected */ }
      });
      sheets.push(sheetInfo);
    });
    return { summary, sheets, details };
  }

  function formatSummary(res,metrics){ const parts=[]; if(metrics.formulasCount) parts.push('Формулы: '+res.summary.formulaCells); if(metrics.volatileFunctions) parts.push('Волатильные: '+res.summary.volatileCells); if(metrics.errorCells) parts.push('Ошибки: '+res.summary.errorCells); if(metrics.longFormulas) parts.push('Длинные формулы: '+res.summary.longFormulas); if(metrics.numericConstants) parts.push('Числовые константы: '+res.summary.numericConstants); return parts.join(' | '); }

  function buildKpis(res,m){ const el=document.getElementById('analysisKpiGrid'); if(!el) return; const s=res.summary; const items=[ {k:'sheets',label:'Листов',val:s.sheets}, {k:'cells',label:'Ячеек',val:s.totalCells}, m.formulasCount?{k:'formulas',label:'Формулы (встроенные)',val:s.formulaCells, dkey:'formulas'}:null, (m.formulasCount && s.formulaCellsFallback)?{k:'formulas',label:'Формулы (текст)',val:s.formulaCellsFallback, dkey:'formulasText'}:null, m.volatileFunctions?{k:'volatile',label:'Волатильные',val:s.volatileCells,dkey:'volatile'}:null, m.errorCells?{k:'errors',label:'Ошибки',val:s.errorCells,dkey:'errors'}:null, m.longFormulas?{k:'long',label:'Длинные формулы',val:s.longFormulas,dkey:'long'}:null, m.nestedIf?{k:'nestedIf',label:'Глубокий ЕСЛИ',val:res.details.nestedIf.length,dkey:'nestedIf'}:null, m.numericConstants?{k:'nums',label:'Константы',val:s.numericConstants,dkey:'numericConstants'}:null ].filter(Boolean); const html=items.map(it=>{ const isDetail=!!it.dkey; const common='<div class="kpi" data-testid="xlsx-analysis-kpi-'+escapeHtml(it.k)+'" data-k="'+it.k+'"'+(isDetail?' data-detail-key="'+it.dkey+'" tabindex="0" role="button" aria-pressed="false" title="Раскрыть детали: '+escapeHtml(it.label)+'" aria-label="Раскрыть детали: '+escapeHtml(it.label)+'"':'')+'><h3 data-testid="xlsx-analysis-kpi-'+escapeHtml(it.k)+'-title">'+escapeHtml(it.label)+'</h3><div class="val" data-testid="xlsx-analysis-kpi-'+escapeHtml(it.k)+'-value">'+escapeHtml(it.val)+'</div></div>'; return common; }).join(''); el.innerHTML=html; }

  function showMetricDetails(res,key){ const wrap=document.getElementById('analysisMetricDetails'); const body=document.getElementById('amdBody'); const title=document.getElementById('amdTitle'); if(!wrap||!body||!title) return; const map={ formulas:'Формулы (встроенные)', formulasText:'Формулы (текст)', errors:'Ошибки', volatile:'Волатильные', long:'Длинные формулы', nestedIf:'Глубокий ЕСЛИ', numericConstants:'Числовые константы' }; const arr=res.details[key]; if(!arr||!arr.length){ body.innerHTML='<em>Нет данных</em>'; } else { let table=['<table class="analysis-table"><thead><tr><th>Лист</th><th>Адрес</th>'];     if(key==='errors') table.push('<th>Тип</th><th>Значение</th>'); else if(key==='long') table.push('<th>Длина</th><th>Формула</th>'); else if(key==='nestedIf') table.push('<th>Глубина</th><th>Формула</th>'); else if(key==='numericConstants') table.push('<th>Контекст</th><th>Значение/Формула</th>'); else table.push('<th>Формула</th>'); table.push('</tr></thead><tbody>'); const max= Math.min(arr.length, 2000); for(let i=0;i<max;i++){ const it=arr[i]; if(key==='errors') table.push('<tr><td>'+escapeHtml(it.sheet)+'</td><td>'+it.addr+'</td><td>'+escapeHtml(it.type||'')+'</td><td>'+escapeHtml(it.value||'')+'</td></tr>'); else if(key==='long') table.push('<tr><td>'+escapeHtml(it.sheet)+'</td><td>'+it.addr+'</td><td>'+it.len+'</td><td>'+escapeHtml(it.formula)+'</td></tr>');     else if(key==='numericConstants'){ const val = it.context==='cell'? String(it.value): it.formula; const ctxRu = it.context==='cell'? 'ячейка' : 'формула'; table.push('<tr><td>'+escapeHtml(it.sheet)+'</td><td>'+it.addr+'</td><td>'+escapeHtml(ctxRu)+'</td><td>'+escapeHtml(val)+'</td></tr>'); } else if(key==='nestedIf') table.push('<tr><td>'+escapeHtml(it.sheet)+'</td><td>'+it.addr+'</td><td>'+it.depth+'</td><td>'+escapeHtml(it.formula)+'</td></tr>'); else table.push('<tr><td>'+escapeHtml(it.sheet)+'</td><td>'+it.addr+'</td><td>'+escapeHtml(it.formula||'')+'</td></tr>'); }
      if(arr.length>max) table.push('<tr><td colspan="4">Показано '+max+' из '+arr.length+'</td></tr>');
      table.push('</tbody></table>'); body.innerHTML=table.join(''); }
    title.textContent = map[key] || key;
    wrap.style.display='';
  }

  function renderResults(res, metrics){
    ensureTestIds();
    const sumEl=document.getElementById('analysisSummary');
    const detEl=document.getElementById('analysisDetails');

    // Removed textual summary to avoid duplication with KPI grid per user request.
    // If the element exists, hide it (non-destructive change) instead of populating text.
    if(sumEl){
      sumEl.textContent='';
      sumEl.style.display='none';
    }
    buildKpis(res,metrics);
    // Wire KPI click handlers for drill-down
    const kpiGrid=document.getElementById('analysisKpiGrid');
    if(kpiGrid){ kpiGrid.querySelectorAll('.kpi[data-detail-key]').forEach(el=>{
      const key=el.getAttribute('data-detail-key');
      const handler=()=>showMetricDetails(res,key);
      el.addEventListener('click',handler);
      el.addEventListener('keydown',e=>{ if(e.key==='Enter' || e.key===' '){ e.preventDefault(); handler(); }});
    }); }
    const closeBtn=document.getElementById('amdClose'); if(closeBtn){ closeBtn.onclick=()=>{ const wrap=document.getElementById('analysisMetricDetails'); if(wrap) wrap.style.display='none'; }; }
    if(detEl){ const rows=[]; rows.push('<table class="analysis-table"><thead><tr><th>Лист</th><th>Ячеек</th>'+(metrics.formulasCount?'<th>Формулы (встроенные)</th>':'')+(metrics.formulasCount?'<th>Формулы (текст)</th>':'')+(metrics.errorCells?'<th>Ошибки</th>':'')+(metrics.volatileFunctions?'<th>Волатильные</th>':'')+(metrics.longFormulas?'<th>Длинные</th>':'')+(metrics.nestedIf?'<th>ЕСЛИ>3</th>':'')+(metrics.numericConstants?'<th>Константы</th>':'')+'</tr></thead><tbody>'); res.sheets.forEach(s=>{ const errCell = metrics.errorCells ? '<td'+(s.errors&&Object.keys(s.errorTypes).length? ' title="'+escapeHtml(Object.entries(s.errorTypes).map(([k,v])=>k+': '+v).join('\n'))+'"':'')+'>'+s.errors+'</td>' : ''; rows.push('<tr><td>'+escapeHtml(s.name)+'</td><td>'+s.cells+'</td>'+(metrics.formulasCount?'<td>'+s.formulas+'</td>':'')+(metrics.formulasCount?'<td>'+s.formulasFallback+'</td>':'')+ errCell +(metrics.volatileFunctions?'<td>'+s.volatile+'</td>':'')+(metrics.longFormulas?'<td>'+s.longFormulas.length+'</td>':'')+(metrics.nestedIf?'<td>'+s.nestedIf+'</td>':'')+(metrics.numericConstants?'<td>'+s.numericConstants+'</td>':'')+'</tr>'); }); rows.push('</tbody></table>'); detEl.innerHTML=rows.join(''); }

  }

  function updateStatus(t){ const st=document.getElementById('analysisStatusTop'); if(st) st.textContent=t; }
  function setProgress(p){ const bar=document.querySelector('#analysisProgressTop span'); if(bar) bar.style.width=Math.max(0,Math.min(100,p))+'%'; }

  // Export analysis report to XLSX
  function exportReport(res, metrics){
    if(!window.XLSX){
      alert('Библиотека XLSX не загружена');
      return;
    }
    
    const wb = XLSX.utils.book_new();
    
    // Sheet 1: Summary (Сводка)
    const summaryData = [
      ['Параметр', 'Значение'],
      ['Количество листов', res.summary.sheets],
      ['Всего ячеек', res.summary.totalCells],
      ['Формулы (встроенные)', res.summary.formulaCells],
      ['Формулы (текст)', res.summary.formulaCellsFallback],
      ['Волатильные функции', res.summary.volatileCells],
      ['Ошибки', res.summary.errorCells],
      ['Длинные формулы (>100)', res.summary.longFormulas],
      ['Глубокий ЕСЛИ (>3)', res.details.nestedIf ? res.details.nestedIf.length : 0],
      ['Числовые константы', res.summary.numericConstants]
    ];
    
    // Add error types breakdown
    if(Object.keys(res.summary.errorTypes).length > 0){
      summaryData.push(['', '']);
      summaryData.push(['Типы ошибок:', '']);
      Object.entries(res.summary.errorTypes).forEach(([type, count]) => {
        summaryData.push([type, count]);
      });
    }
    
    const ws_summary = XLSX.utils.aoa_to_sheet(summaryData);
    ws_summary['!cols'] = [{wch: 25}, {wch: 15}];
    XLSX.utils.book_append_sheet(wb, ws_summary, 'Сводка');
    
    // Sheet 2: Details by sheet (Детали по листам)
    const detailsHeaders = ['Лист', 'Всего ячеек'];
    if(metrics.formulasCount) detailsHeaders.push('Формулы (встроенные)', 'Формулы (текст)');
    if(metrics.errorCells) detailsHeaders.push('Ошибки');
    if(metrics.volatileFunctions) detailsHeaders.push('Волатильные');
    if(metrics.longFormulas) detailsHeaders.push('Длинные');
    if(metrics.nestedIf) detailsHeaders.push('ЕСЛИ>3');
    if(metrics.numericConstants) detailsHeaders.push('Константы');
    
    const detailsData = [detailsHeaders];
    res.sheets.forEach(s => {
      const row = [s.name, s.cells];
      if(metrics.formulasCount) row.push(s.formulas, s.formulasFallback);
      if(metrics.errorCells) row.push(s.errors);
      if(metrics.volatileFunctions) row.push(s.volatile);
      if(metrics.longFormulas) row.push(s.longFormulas.length);
      if(metrics.nestedIf) row.push(s.nestedIf);
      if(metrics.numericConstants) row.push(s.numericConstants);
      detailsData.push(row);
    });
    
    const ws_details = XLSX.utils.aoa_to_sheet(detailsData);
    XLSX.utils.book_append_sheet(wb, ws_details, 'Детали по листам');
    
    // Sheet 3: Formulas (if any)
    if(metrics.formulasCount && (res.details.formulas.length > 0 || res.details.formulasText.length > 0)){
      const formulasData = [['Лист', 'Адрес', 'Тип', 'Формула']];
      res.details.formulas.forEach(it => {
        formulasData.push([it.sheet, it.addr, 'встроенная', it.formula]);
      });
      res.details.formulasText.forEach(it => {
        formulasData.push([it.sheet, it.addr, 'текст', it.formula]);
      });
      const ws_formulas = XLSX.utils.aoa_to_sheet(formulasData);
      ws_formulas['!cols'] = [{wch: 20}, {wch: 10}, {wch: 10}, {wch: 60}];
      XLSX.utils.book_append_sheet(wb, ws_formulas, 'Формулы');
    }
    
    // Sheet 4: Errors (if any)
    if(metrics.errorCells && res.details.errors.length > 0){
      const errorsData = [['Лист', 'Адрес', 'Тип ошибки', 'Значение']];
      res.details.errors.forEach(it => {
        errorsData.push([it.sheet, it.addr, it.type, it.value]);
      });
      const ws_errors = XLSX.utils.aoa_to_sheet(errorsData);
      ws_errors['!cols'] = [{wch: 20}, {wch: 10}, {wch: 15}, {wch: 40}];
      XLSX.utils.book_append_sheet(wb, ws_errors, 'Ошибки');
    }
    
    // Sheet 5: Volatile Functions (if any)
    if(metrics.volatileFunctions && res.details.volatile.length > 0){
      const volatileData = [['Лист', 'Адрес', 'Формула']];
      res.details.volatile.forEach(it => {
        volatileData.push([it.sheet, it.addr, it.formula]);
      });
      const ws_volatile = XLSX.utils.aoa_to_sheet(volatileData);
      ws_volatile['!cols'] = [{wch: 20}, {wch: 10}, {wch: 60}];
      XLSX.utils.book_append_sheet(wb, ws_volatile, 'Волатильные');
    }
    
    // Sheet 6: Long Formulas (if any)
    if(metrics.longFormulas && res.details.long.length > 0){
      const longData = [['Лист', 'Адрес', 'Длина', 'Формула']];
      res.details.long.forEach(it => {
        longData.push([it.sheet, it.addr, it.len, it.formula]);
      });
      const ws_long = XLSX.utils.aoa_to_sheet(longData);
      ws_long['!cols'] = [{wch: 20}, {wch: 10}, {wch: 10}, {wch: 80}];
      XLSX.utils.book_append_sheet(wb, ws_long, 'Длинные формулы');
    }
    
    // Sheet 7: Nested IF (if any)
    if(metrics.nestedIf && res.details.nestedIf.length > 0){
      const nestedIfData = [['Лист', 'Адрес', 'Глубина', 'Формула']];
      res.details.nestedIf.forEach(it => {
        nestedIfData.push([it.sheet, it.addr, it.depth, it.formula]);
      });
      const ws_nestedIf = XLSX.utils.aoa_to_sheet(nestedIfData);
      ws_nestedIf['!cols'] = [{wch: 20}, {wch: 10}, {wch: 10}, {wch: 60}];
      XLSX.utils.book_append_sheet(wb, ws_nestedIf, 'Глубокий ЕСЛИ');
    }
    
    // Sheet 8: Numeric Constants (if any)
    if(metrics.numericConstants && res.details.numericConstants.length > 0){
      const constantsData = [['Лист', 'Адрес', 'Контекст', 'Значение/Формула']];
      res.details.numericConstants.forEach(it => {
        const ctxRu = it.context === 'cell' ? 'ячейка' : 'формула';
        const val = it.context === 'cell' ? String(it.value) : it.formula;
        constantsData.push([it.sheet, it.addr, ctxRu, val]);
      });
      const ws_constants = XLSX.utils.aoa_to_sheet(constantsData);
      ws_constants['!cols'] = [{wch: 20}, {wch: 10}, {wch: 12}, {wch: 60}];
      XLSX.utils.book_append_sheet(wb, ws_constants, 'Константы');
    }
    
    // Generate filename with timestamp
    const now = new Date();
    const timestamp = now.toISOString().slice(0, 19).replace(/:/g, '-').replace('T', '_');
    const filename = `отчёт_анализа_${timestamp}.xlsx`;
    
    // Write file
    XLSX.writeFile(wb, filename);
    updateStatus('Отчет экспортирован: ' + filename);
  }

  // Store last analysis results for export
  let lastAnalysisResults = null;
  let lastAnalysisMetrics = null;

  async function runAnalysis(){ const fileInput=document.getElementById('analysisFileInput'); if(!fileInput||!fileInput.files||!fileInput.files[0]) return; const file=fileInput.files[0]; const metrics=collectSelectedMetrics(); updateStatus('Чтение файла...'); setProgress(5); let buffer; try{ buffer=await readFileAsArrayBuffer(file);}catch(e){ updateStatus('Ошибка чтения'); console.error(e); return;} let wb; try{ updateStatus('Парсинг...'); setProgress(20); wb=parseWorkbook(buffer);}catch(e){ updateStatus('Ошибка парсинга'); console.error(e); return;} try{ updateStatus('Анализ...'); setProgress(50); const res=analyzeWorkbook(wb,metrics); updateStatus('Рендер...'); setProgress(80); renderResults(res,metrics); updateStatus('Готово'); setProgress(100); lastAnalysisResults=res; lastAnalysisMetrics=metrics; const reportBtn=document.getElementById('btnReportAnalysis'); if(reportBtn) reportBtn.disabled=false;}catch(e){ updateStatus('Ошибка анализа'); console.error(e);} }

  function attachStart(){ ensureTestIds(); const btn=document.getElementById('btnStartAnalysis'); if(!btn||btn._wiredAnalysis) return; btn._wiredAnalysis=true; btn.addEventListener('click',()=>{ setTimeout(runAnalysis,20); }); }
  
  function attachReport(){ ensureTestIds(); const btn=document.getElementById('btnReportAnalysis'); if(!btn||btn._wiredReport) return; btn._wiredReport=true; btn.addEventListener('click',()=>{ if(lastAnalysisResults && lastAnalysisMetrics){ exportReport(lastAnalysisResults, lastAnalysisMetrics); } else { updateStatus('Сначала запустите анализ'); } }); }

  // ===================================================================
  // Memory Cleanup for Analysis Module
  // ===================================================================
  function cleanupAnalysisData() {
    try {
      // Clear analysis results cache
      if (lastAnalysisResults) {
        if (lastAnalysisResults.details) {
          // Clear detail arrays
          Object.keys(lastAnalysisResults.details).forEach(key => {
            if (Array.isArray(lastAnalysisResults.details[key])) {
              lastAnalysisResults.details[key].length = 0;
              lastAnalysisResults.details[key] = null;
            }
          });
          lastAnalysisResults.details = null;
        }
        if (lastAnalysisResults.sheets) {
          lastAnalysisResults.sheets.length = 0;
          lastAnalysisResults.sheets = null;
        }
        lastAnalysisResults = null;
      }
      
      lastAnalysisMetrics = null;
      
      // Clear file input
      const fileInput = document.getElementById('analysisFileInput');
      if (fileInput) {
        fileInput.value = '';
      }
      
      // Clear UI
      const kpiGrid = document.getElementById('analysisKpiGrid');
      if (kpiGrid) kpiGrid.innerHTML = '';
      
      const detEl = document.getElementById('analysisDetails');
      if (detEl) detEl.innerHTML = '';
      
      updateStatus('Данные анализа очищены');
      
    } catch (error) {
      console.error('Failed to cleanup analysis data', error);
    }
  }

  if(window.location.pathname.indexOf('analysis.html')!==-1){ 
    // Setup cleanup on window close
    window.addEventListener('beforeunload', cleanupAnalysisData);
    window.addEventListener('unload', cleanupAnalysisData);
    
    ensureTestIds();
    if(!window.XLSX){ const script=document.createElement('script'); script.src='vendor/xlsx.full.min.js'; script.onload=()=>{attachStart(); attachReport();}; script.onerror=()=>updateStatus('Не удалось загрузить XLSX библиотеку'); document.head.appendChild(script); } else {attachStart(); attachReport();} 
  }

  window.PluginAnalysis={ run:runAnalysis, cleanup: cleanupAnalysisData };

  // Original popup support retained
  function enableButtonIfReady(){ try{ const st=(window.State&&State.get&&State.get())?State.get():{}; const btn=document.getElementById('btnAnalyze'); if(!btn) return; const hasAny=Array.isArray(st.sheets)&&st.sheets.length>0; btn.disabled=!hasAny; if(!hasAny) btn.setAttribute('aria-disabled','true'); else btn.removeAttribute('aria-disabled'); }catch(e){ if(window.Logger) Logger.warn('Analysis button enable check failed',e); } }
  function openAnalysisWindow(){ try{ const w=window.open('analysis.html','analysisWindow','width=1024,height=768,resizable=yes,scrollbars=yes'); if(!w && window.PluginActions && PluginActions._pushUserMessage) PluginActions._pushUserMessage('error','Не удалось открыть окно анализа'); }catch(e){ if(window.Logger) Logger.error('Open analysis window failed',e); } }
  function wire(){ const btn=document.getElementById('btnAnalyze'); if(btn && !btn._wired){ btn._wired=true; btn.addEventListener('click',()=>{ if(btn.disabled) return; openAnalysisWindow(); }); } enableButtonIfReady(); }
  if(window.State && State.subscribe) State.subscribe(enableButtonIfReady); document.addEventListener('DOMContentLoaded',wire);
})(window);