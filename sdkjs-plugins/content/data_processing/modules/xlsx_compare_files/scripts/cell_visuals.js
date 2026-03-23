/**
 * @fileoverview Cell visualization helper (extracted from plugin.js render logic)
 * @description Computes ordered categories and style decisions for diff preview cells.
 * Follows coding standards: small pure functions, no side-effects, <=30 lines per function.
 */
(function(window){
  'use strict';
  const PRIORITY_ORDER = Object.freeze(['formula','value','formulaToValue','valueToFormula','type','format','inserted-row','deleted-row','inserted-col','deleted-col']);
  const COLOR_MAP = Object.freeze({
    value:'#fff59d', formula:'#bbdefb', formulaToValue:'#d1c4e9', valueToFormula:'#ffe0b2', type:'#f8bbd0', format:'#cfd8dc',
    'inserted-row':'#c8e6c9','deleted-row':'#ffcdd2','inserted-col':'#c5e1a5','deleted-col':'#ef9a9a'
  });
  function orderCategories(set){ if(!set||!set.size) return null; return [...set].sort((a,b)=> PRIORITY_ORDER.indexOf(a)-PRIORITY_ORDER.indexOf(b)); }
  function computeBaseCategory(ordered){ if(!ordered||!ordered.length) return null; const base=ordered[0]; if(/^(inserted|deleted)-(row|col)$/.test(base)){ return null; } return base; }
  function computeStripeStyle(ordered){
    if(!ordered || ordered.length<2) return null;
    const nonStruct = ordered.filter(k=>!['inserted-row','deleted-row','inserted-col','deleted-col'].includes(k));
    if(nonStruct.length<2) return null;
    const baseCat = nonStruct[0]; const lastCat = nonStruct[nonStruct.length-1];
    const baseColor = COLOR_MAP[baseCat] || '#ffffff'; const stripeColor = COLOR_MAP[lastCat] || baseColor;
    return buildStripeStyle(baseColor, stripeColor);
  }
  function hexToRgb(hex){ const h=hex.replace('#',''); return { r:parseInt(h.substr(0,2),16), g:parseInt(h.substr(2,2),16), b:parseInt(h.substr(4,2),16) }; }
  function adjustForContrast(rgb, base){ const diff=Math.abs(rgb.r-base.r)+Math.abs(rgb.g-base.g)+Math.abs(rgb.b-base.b); if(diff<30){ return { r:Math.max(0,rgb.r-60), g:Math.max(0,rgb.g-60), b:Math.max(0,rgb.b-60) }; } return rgb; }
  function buildStripeStyle(baseColor, stripeColor){
    try {
      const baseRGB=hexToRgb(baseColor); let stripeRGB=hexToRgb(stripeColor); stripeRGB=adjustForContrast(stripeRGB, baseRGB); const rgba=`rgba(${stripeRGB.r},${stripeRGB.g},${stripeRGB.b},0.45)`;
      return `repeating-linear-gradient(45deg, ${rgba} 0 6px, rgba(0,0,0,0) 6px 12px)`;
    } catch(e){ return 'repeating-linear-gradient(45deg, rgba(0,0,0,0.25) 0 6px, rgba(0,0,0,0) 6px 12px)'; }
  }
  function buildTooltip(set){ if(!set||!set.size) return ''; const ordered=orderCategories(set); return ordered.join(','); }
  window.CellVisuals = { orderCategories, computeBaseCategory, computeStripeStyle, buildTooltip, PRIORITY_ORDER };
})(window);
