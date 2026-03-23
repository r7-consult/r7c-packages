/** Lightweight OnlyOffice API wrapper (no mock mode) */
(function(window){
    'use strict';
    class R7APIImpl {
        constructor(){
            this._ready = !!(window.Asc && window.Asc.plugin);
            if(!this._ready){
                // Poll until editor API injected
                this._timer = setInterval(()=>{
                    if(window.Asc && window.Asc.plugin){
                        this._ready=true;
                        clearInterval(this._timer); this._timer=null;
                        try { window.dispatchEvent(new CustomEvent('r7api:real-ready')); } catch(_e){}
                    }
                }, 500);
            }
        }
        isAvailable(){ return !!(window.Asc && window.Asc.plugin); }
        isReady(){ return this.isAvailable(); }
        get readinessDetails(){
            return {
                ascPresent: !!window.Asc,
                pluginPresent: !!(window.Asc && window.Asc.plugin),
                hasCallCommand: !!(window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand),
                callCommandArity: (window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand)? window.Asc.plugin.callCommand.length : -1,
                readyFlag: this._ready
            };
        }
        callCommand(fn, cb){
            if(!this.isReady()) return;
            try {
                const impl = window.Asc.plugin.callCommand;
                if(typeof impl === 'function'){
                    if(impl.length <= 2){
                        return impl.call(window.Asc.plugin, fn, cb);
                    } else {
                        return impl.call(window.Asc.plugin, fn, false, false, cb);
                    }
                }
            } catch(e){ console.error('[R7API] callCommand error', e); }
        }
        getSheetsList(){
            if(!this.isReady()) return Promise.resolve([]);
            return new Promise((resolve)=>{
                this.callCommand(function(){
                    var names=[]; try { var sheets=Api.GetSheets(); if(sheets && sheets.forEach){ sheets.forEach(function(sh){ try{ names.push(sh.GetName()); }catch(e){} }); } } catch(e){}
                    Asc.scope._sheetsList = names;
                }, ()=> resolve(Asc.scope._sheetsList||[]));
            });
        }
        getDocumentName(){
            if(!this.isReady()) return Promise.resolve(null);
            return new Promise((resolve)=>{
                this.callCommand(function(){ try { Asc.scope._docName = (typeof Api.asc_getDocumentName==='function')? Api.asc_getDocumentName(): null; } catch(e){ Asc.scope._docName=null; } }, ()=> resolve(Asc.scope._docName||null));
            });
        }
        getActiveSheetName(){
            if(!this.isReady()) return Promise.resolve(null);
            return new Promise((resolve)=>{
                this.callCommand(function(){ try { var sh=Api.GetActiveSheet(); Asc.scope._activeSheetName= sh? sh.GetName(): null; } catch(e){ Asc.scope._activeSheetName=null; } }, ()=> resolve(Asc.scope._activeSheetName||null));
            });
        }
        getSheetSnapshot(sheetName){
            if(!this.isReady()) return Promise.resolve(null);
            return new Promise((resolve,reject)=>{
                this.callCommand(function(){
                    try {
                        var sheet=Api.GetSheet(sheetName);
                        if(!sheet){ Asc.scope._snapErr='SHEET_NOT_FOUND'; return; }
                        var used=sheet.GetUsedRange(); var maxR=used?used.GetRowCount():0; var maxC=used?used.GetColCount():0; var rows=[];
                        for(var r=0;r<maxR;r++){ var rowArr=[]; for(var c=0;c<maxC;c++){ var cell=sheet.GetCells(r,c); var val=null,f=null,t=null; try{ f=cell.GetFormula&&cell.GetFormula(); }catch(e){} try{ val=cell.GetValue&&cell.GetValue(); }catch(e){} try{ t=cell.GetValueType&&cell.GetValueType(); }catch(e){} rowArr.push({v:val,f:f||null,t:t||null}); } rows.push(rowArr); }
                        Asc.scope._sheetSnapshot={ name:sheetName, rows:rows, maxR:maxR, maxC:maxC };
                    } catch(e){ Asc.scope._snapErr=e.message; }
                }, ()=>{ if(Asc.scope._snapErr) reject(new Error(Asc.scope._snapErr)); else resolve(Asc.scope._sheetSnapshot); });
            });
        }
    }
    window.R7API = new R7APIImpl();
})(window);
