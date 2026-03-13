(function(window) {
  'use strict';

  function resolveModuleUrl(modulePath) {
    try {
      return new URL(modulePath, window.location.href).toString();
    } catch (_) {
      const location = window.location;
      const idx = location.pathname.lastIndexOf('/') + 1;
      const fileName = location.pathname.slice(idx);
      return location.href.replace(fileName, modulePath);
    }
  }

  function normalizeModulePath(modulePath) {
    return String(modulePath || '').replace(/\\/g, '/');
  }



  function normalizeSize(viewport, requested, fallback) {
    const output = [fallback[0], fallback[1]];
    if (Array.isArray(requested) && requested.length === 2) {
      output[0] = requested[0];
      output[1] = requested[1];
    }

    if (viewport[0] > 0) {
      output[0] = Math.min(output[0], Math.floor(viewport[0] * 0.95));
    }
    if (viewport[1] > 0) {
      output[1] = Math.min(output[1], Math.floor(viewport[1] * 0.95));
    }

    output[0] = Math.max(320, output[0]);
    output[1] = Math.max(320, output[1]);

    return output;
  }

  function getViewport() {
    let w = 0;
    let h = 0;
    try {
      if (window.parent) {
        w = window.parent.innerWidth || 0;
        h = window.parent.innerHeight || 0;
      }
    } catch (_) {}
    return [w, h];
  }

  function createVariation(game, url, size, minSize, maxSize) {
    return {
      url: url,
      description: game.title,
      isVisual: true,
      isModal: true,
      isResizable: true,
      isInsideMode: false,
      isUpdateOleOnResize: false,
      EditorsSupport: ['cell'],
      size: size,
      minSize: minSize,
      maxSize: maxSize,
      buttons: []
    };
  }

  function openPopup(url, size) {
    const popup = window.open(
      url,
      '_blank',
      'noopener,noreferrer,resizable=yes,scrollbars=yes,width=' + size[0] + ',height=' + size[1]
    );

    if (popup) {
      try {
        popup.opener = null;
      } catch (_) {}
    }

    return popup;
  }

  class Launcher {
    open(game) {
      if (!game || !game.modulePath) {
        throw new Error('Game metadata is invalid');
      }

      const viewport = getViewport();
      const fallbackSize = [960, 720];
      const size = normalizeSize(viewport, game.size, fallbackSize);
      const minSize = normalizeSize(viewport, game.minSize || [640, 480], [640, 480]);
      const maxSize = normalizeSize(viewport, game.maxSize || [1600, 1000], [1600, 1000]);

      // Ensure `maxSize` is never below `minSize`.
      maxSize[0] = Math.max(maxSize[0], minSize[0]);
      maxSize[1] = Math.max(maxSize[1], minSize[1]);

      const modulePath = normalizeModulePath(game.modulePath);
      const moduleUrl = resolveModuleUrl(modulePath);
      const variation = createVariation(game, moduleUrl, size, minSize, maxSize);

      if (window.Asc && window.Asc.PluginWindow) {
        const modal = new window.Asc.PluginWindow();
        modal.show(variation);
        return modal;
      }

      const popup = openPopup(moduleUrl, size);
      if (popup) {
        return popup;
      }

      // Last fallback for strict popup policies.
      window.location.href = moduleUrl;
      return { type: 'redirect' };
    }
  }

  window.GameLauncher = new Launcher();
})(window);
