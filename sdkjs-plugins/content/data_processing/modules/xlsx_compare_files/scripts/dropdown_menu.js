// Simple dropdown menu for Start / About navigation
(function(){
  function init(){
    var icon=document.getElementById('pluginIcon');
    var dd=document.getElementById('pluginDropdown');
    if(!icon||!dd) return;
    try {
      function safeKey(s){ return String(s==null? '' : s).trim().replace(/[^A-Za-z0-9_]/g,'_'); }
      if(!icon.getAttribute('data-testid')) icon.setAttribute('data-testid','xlsx-plugin-menu-toggle');
      if(!dd.getAttribute('data-testid')) dd.setAttribute('data-testid','xlsx-plugin-menu-dropdown');
      var items = dd.querySelectorAll('.menu-item');
      Array.prototype.forEach.call(items, function(el, idx){
        if(!el || el.getAttribute('data-testid')) return;
        var key = el.id || el.getAttribute('data-action') || el.getAttribute('href') || el.textContent || idx;
        el.setAttribute('data-testid','xlsx-plugin-menu-item-'+safeKey(key)+'-'+idx);
      });
    } catch(_e){}
    function toggle(force){
      var open = force!=null? force : !dd.classList.contains('open');
      if(open){ dd.classList.add('open'); icon.setAttribute('aria-expanded','true'); }
      else { dd.classList.remove('open'); icon.setAttribute('aria-expanded','false'); }
    }
    icon.addEventListener('click', function(e){ e.stopPropagation(); toggle(); });
    icon.addEventListener('mouseenter', function(){ toggle(true); });
    dd.addEventListener('mouseleave', function(){ toggle(false); });
    document.addEventListener('click', function(e){ if(!e.target.closest('.plugin-menu')) toggle(false); });
    // Keyboard navigation
    icon.addEventListener('keydown', function(e){ if(e.key==='Enter'||e.key===' '||e.key==='ArrowDown'){ e.preventDefault(); toggle(true); var first=dd.querySelector('.menu-item'); first && first.focus(); } });
    dd.addEventListener('keydown', function(e){ var items=[...dd.querySelectorAll('.menu-item')]; var idx=items.indexOf(document.activeElement); if(e.key==='Escape'){ toggle(false); icon.focus(); } else if(e.key==='ArrowDown'){ e.preventDefault(); items[(idx+1)%items.length].focus(); } else if(e.key==='ArrowUp'){ e.preventDefault(); items[(idx-1+items.length)%items.length].focus(); } });
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
