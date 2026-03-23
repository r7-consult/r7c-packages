/** Simplified diff provider: JS only (Python/Hybrid removed) */
(function(window){
  'use strict';
  function log(){ if(window.Logger && Logger.info) Logger.info.apply(Logger, ['[DiffProvider]'].concat([].slice.call(arguments))); }

  class JsProvider {
    constructor(opts){ this.opts = opts; }
    async compare(base, target, progressCb){
      const engine = new window.DiffEngine(this.opts);
      return engine.compareSheets(base, target, progressCb);
    }
  }

  // Always return JS provider
  window.DiffProviders = {
    async select(_base,_target,opts){ log('JS provider selected (only option)'); return new JsProvider(opts); }
  };
})(window);
