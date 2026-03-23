/** Immutable state manager */
(function(window){
 'use strict';
 class StateManager {
  constructor(){ this.#state=Object.freeze({ sheets:[], base:null, target:null, options:{ ignoreCase:true, trim:true, compareFormatting:false }, enabledCategories:null, diff:null, running:false, importing:false, progress:0, highlightApplied:false, statusMessage:null, statusKind:null }); this.#listeners=new Set(); }
   #state; #listeners;
   get(){ return this.#state; }
   subscribe(fn){ this.#listeners.add(fn); return ()=>this.#listeners.delete(fn); }
   set(patch){ const next=Object.freeze({...this.#state,...patch}); this.#state=next; this.#listeners.forEach(l=>{ try{ l(next);}catch(e){ console.error(e);} }); }
 }
 window.State = new StateManager();
})(window);
