(function(window){
 'use strict';
 // Basic adaptive resize handler aligning with flexible window standard (guide 14)
 function onResize(){ try { const root=document.getElementById('appRoot'); if(!root) return; const h=window.innerHeight; const w=window.innerWidth; root.style.setProperty('--app-height', h+'px'); root.style.setProperty('--app-width', w+'px'); } catch(e){ if(window.Logger) Logger.warn('resize handler error', e); } }
 window.addEventListener('resize', ()=>{ requestAnimationFrame(onResize); });
 document.addEventListener('DOMContentLoaded', onResize);
 window.__resizeHandlerInstalled = true;
})(window);
