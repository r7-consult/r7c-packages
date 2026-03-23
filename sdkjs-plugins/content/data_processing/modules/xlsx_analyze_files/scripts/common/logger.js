/** Safe logger (no console override) */
(function(window){
 'use strict';
 const LOGS=[];
 function push(level,msg,extra){ const entry={ts:new Date().toISOString(),level,msg,extra}; LOGS.push(entry); if(LOGS.length>5000) LOGS.shift(); console[level==='error'?'error':'log'](`[${level}]`,msg,extra||''); }
 const Logger={ info:(m,...a)=>push('info',m,a), warn:(m,...a)=>push('warn',m,a), error:(m,...a)=>push('error',m,a), debug:(m,...a)=>push('debug',m,a), getAll:()=>LOGS.slice(), bridgeAllToLogWindow(){ if(!window.MessageOutput) return; LOGS.forEach(l=>{ if(l.level==='error'&&window.MessageOutput.ErrorLog) window.MessageOutput.ErrorLog(l.msg); else if(l.level==='warn'&&window.MessageOutput.WarningLog) window.MessageOutput.WarningLog(l.msg); else if(window.MessageOutput.InfoLog) window.MessageOutput.InfoLog(l.msg); }); } };
 window.Logger=Logger; window.debug=Logger; // convenience per standards
})(window);
