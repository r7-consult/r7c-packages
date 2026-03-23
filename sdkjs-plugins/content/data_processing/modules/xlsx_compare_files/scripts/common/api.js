/** Unified API facade */
(function(window){
    'use strict';
    class APIFacade {
        constructor(){ this._ready=false; this._wait(); }
    _wait(){ if(window.R7API && window.R7API.isAvailable && window.R7API.isAvailable()){ this._ready=true; } else { setTimeout(()=>this._wait(),200); } }
        // Dynamic readiness check (do not rely only on cached _ready)
    isReady(){ if(this._ready) return true; if(window.R7API && window.R7API.isAvailable && window.R7API.isAvailable()){ this._ready=true; return true; } return false; }
        async listSheets(){ if(!this.isReady()) return []; try { return await window.R7API.getSheetsList(); } catch(_e){ return []; } }
        async getActiveSheet(){ if(!this.isReady()) return null; try { return await window.R7API.getActiveSheetName(); } catch(_e){ return null; } }
    }
    window.API = new APIFacade();
})(window);
