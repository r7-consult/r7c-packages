(function(window){
  'use strict';
  function colName(c){ let s=''; c=Number(c); if(isNaN(c)||c<0) return ''; c++; while(c>0){ let m=(c-1)%26; s=String.fromCharCode(65+m)+s; c=Math.floor((c-1)/26); } return s; }
  function toA1(r,c){ return colName(c)+(r+1); }
  function escapeHtml(s){ if(s==null) return ''; return String(s).replace(/[&<>'"]/g, ch=>({ '&':'&amp;','<':'&lt;','>':'&gt;','\'':'&#39;','"':'&quot;' }[ch])); }
  function fmt(v){ if(v==null) return ''; if(typeof v==='object') return JSON.stringify(v); return String(v); }
  window.Utils = window.Utils || Object.freeze({ colName, toA1, escapeHtml, fmt });
})(window);
