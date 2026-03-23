(function(window){
  'use strict';
  if(window.I18n) return;
  const I18n = {
    lang: 'ru',
    dict: {},
    fallbackLang: 'en',
    loading: null,
    listeners: [],
    setLanguage(lang){ if(lang && lang!==this.lang){ this.lang=lang; this.loadLanguage(lang,true); } },
    onChange(cb){ if(typeof cb==='function') this.listeners.push(cb); },
    t(key, params, fallback){
      const raw = (this.dict && this.dict[key]) || fallback || key;
      if(params){ return raw.replace(/\{(\w+)\}/g,(m,k)=> (k in params? params[k]: m)); }
      return raw;
    },
    apply(root){
      const scope = root || document;
      // Elements with data-i18n = key (textContent)
      scope.querySelectorAll('[data-i18n]').forEach(el=>{
        const k = el.getAttribute('data-i18n'); if(!k) return; el.textContent = this.t(k, null, el.textContent.trim());
      });
      // Attributes data-i18n-attrName
      const attrMap = ['title','placeholder','aria-label','aria-description','value'];
      attrMap.forEach(attr=>{
        scope.querySelectorAll('[data-i18n-'+attr+']').forEach(el=>{
          const k = el.getAttribute('data-i18n-'+attr); if(!k) return; const translated = this.t(k, null, el.getAttribute(attr)||''); if(translated) el.setAttribute(attr, translated);
        });
      });
    },
    loadJson(path){ return fetch(path).then(r=> r.ok? r.json(): {}).catch(()=> ({})); },
    async loadLanguage(lang, fire){
      if(this.loading) try { await this.loading; } catch(_e){}
      const mainPath = 'translations/'+lang+'.json';
      this.loading = this.loadJson(mainPath);
      let dict = await this.loading; this.loading=null;
      if(!dict || Object.keys(dict).length===0 && lang!==this.fallbackLang){
        dict = await this.loadJson('translations/'+this.fallbackLang+'.json');
      }
      this.dict = dict || {}; window.__i18n = this.dict; // compatibility
      try { this.apply(); } catch(_e){}
      if(fire){ this.listeners.forEach(cb=>{ try { cb(lang); } catch(_e){} }); }
      return this.dict;
    },
    async init(options){
      if(options && options.lang) this.lang = options.lang;
      const htmlLang = document.documentElement.getAttribute('lang');
      if(htmlLang) this.lang = htmlLang;
      await this.loadLanguage(this.lang, true);
    }
  };
  window.I18n = I18n;
})(window);
