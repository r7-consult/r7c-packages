(function(window) {
  'use strict';

  var HOSTED_LAUNCH_CONTEXT_KEY = 'data_processing_hosted_launch_context';

  function normalizeSize(viewport, requested, fallback) {
    var output = [fallback[0], fallback[1]];

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
    var width = 0;
    var height = 0;

    try {
      if (window.parent) {
        width = window.parent.innerWidth || 0;
        height = window.parent.innerHeight || 0;
      }
    } catch (_) {}

    return [width, height];
  }

  function buildFullUrl(fileName) {
    var currentUrl = String(window.location.href || '');
    var baseUrl = currentUrl.substring(0, currentUrl.lastIndexOf('/') + 1);
    return baseUrl + fileName;
  }

  function persistHostedLaunchContext(context) {
    try {
      window.localStorage.setItem(HOSTED_LAUNCH_CONTEXT_KEY, JSON.stringify(context));
      return true;
    } catch (_) {
      return false;
    }
  }

  function createVariation(record, size, minSize, maxSize) {
    var base = record && record.variation && typeof record.variation === 'object' ? record.variation : {};
    var variation = {};
    var key;

    for (key in base) {
      if (Object.prototype.hasOwnProperty.call(base, key)) {
        variation[key] = base[key];
      }
    }

    variation.url = record.launchUrl;
    variation.description = base.description || record.title;
    variation.isViewer = typeof base.isViewer === 'boolean' ? base.isViewer : true;
    variation.isVisual = true;
    variation.isModal = true;
    variation.isResizable = typeof base.isResizable === 'boolean' ? base.isResizable : true;
    variation.isInsideMode = false;
    variation.isUpdateOleOnResize = !!base.isUpdateOleOnResize;
    variation.EditorsSupport = Array.isArray(base.EditorsSupport) && base.EditorsSupport.length ? base.EditorsSupport.slice() : ['cell'];
    variation.initDataType = base.initDataType || 'none';
    variation.initData = typeof base.initData === 'string' ? base.initData : '';
    variation.buttons = Array.isArray(base.buttons) ? base.buttons.slice() : [];
    variation.size = size;
    variation.minSize = minSize;
    variation.maxSize = maxSize;

    return variation;
  }

  function createHostedVariation(context) {
    var size = Array.isArray(context.windowSize) && context.windowSize.length === 2 ? context.windowSize.slice() : [980, 760];
    var minSize = Array.isArray(context.windowMinSize) && context.windowMinSize.length === 2 ? context.windowMinSize.slice() : [720, 560];
    var maxSize = Array.isArray(context.windowMaxSize) && context.windowMaxSize.length === 2 ? context.windowMaxSize.slice() : [1600, 1000];
    var buttons = Array.isArray(context.buttons)
      ? context.buttons.slice()
      : [];

    return {
      url: buildFullUrl('module-window-host.html'),
      description: context.title || 'Hosted Module Window',
      isViewer: true,
      isVisual: true,
      isModal: true,
      isResizable: true,
      isInsideMode: false,
      initDataType: 'none',
      initData: '',
      isUpdateOleOnResize: true,
      EditorsSupport: Array.isArray(context.editorsSupport) && context.editorsSupport.length ? context.editorsSupport.slice() : ['cell'],
      buttons: buttons,
      size: size,
      minSize: minSize,
      maxSize: maxSize
    };
  }

  function openPopup(url, size) {
    var popup = window.open(
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

  function openExternalLink(url) {
    var popup = null;
    var anchor = null;

    if (!url) {
      throw new Error('External link is not configured');
    }

    try {
      if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.executeMethod === 'function') {
        window.Asc.plugin.executeMethod('OpenLink', [url]);
        return { type: 'open-link-method', url: url };
      }
    } catch (_) {}

    try {
      popup = window.open(url, '_blank', 'noopener');
    } catch (_) {
      popup = null;
    }

    if (popup) {
      try {
        popup.opener = null;
      } catch (_) {}
      return popup;
    }

    try {
      if (window.top && window.top !== window && typeof window.top.open === 'function') {
        popup = window.top.open(url, '_blank', 'noopener');
      }
    } catch (_) {
      popup = null;
    }

    if (popup) {
      try {
        popup.opener = null;
      } catch (_) {}
      return popup;
    }

    try {
      anchor = document.createElement('a');
      anchor.href = url;
      anchor.target = '_blank';
      anchor.rel = 'noopener noreferrer';
      anchor.style.position = 'absolute';
      anchor.style.left = '-9999px';
      anchor.style.width = '1px';
      anchor.style.height = '1px';
      document.body.appendChild(anchor);
      anchor.click();
      return { type: 'anchor-click', url: url };
    } catch (_) {
      // fall through to the explicit error below
    } finally {
      if (anchor && anchor.parentNode) {
        anchor.parentNode.removeChild(anchor);
      }
    }

    throw new Error('External link could not be opened');
  }

  function getPluginInfo() {
    return window.Asc && window.Asc.plugin ? window.Asc.plugin.info || {} : {};
  }

  function createHostedContext(record, overrides) {
    var pluginInfo = getPluginInfo();
    var options = overrides || {};

    return {
      title: options.title || record.title || 'Hosted Module',
      moduleId: options.moduleId || record.module || '',
      moduleEntry: options.moduleEntry || record.moduleEntry || ('modules/' + record.module + '/' + (record.entry || 'index.html')),
      moduleGuid: options.moduleGuid || record.guid || '',
      hostGuid: options.hostGuid || pluginInfo.guid || '',
      initData: typeof options.initData === 'string' ? options.initData : '',
      objectId: typeof options.objectId === 'undefined' ? pluginInfo.objectId : options.objectId,
      resize: typeof options.resize === 'boolean' ? options.resize : !!pluginInfo.resize,
      width: options.width || pluginInfo.width || 70,
      height: options.height || pluginInfo.height || 70,
      mmToPx: options.mmToPx || pluginInfo.mmToPx || 3.78,
      lang: options.lang || pluginInfo.lang || 'ru-RU',
      theme: options.theme || pluginInfo.theme || 'theme-light',
      windowSize: options.windowSize || record.size || [980, 760],
      windowMinSize: options.windowMinSize || record.minSize || [720, 560],
      windowMaxSize: options.windowMaxSize || record.maxSize || [1600, 1000],
      buttons: Array.isArray(options.buttons) ? options.buttons.slice() : [],
      editorsSupport: options.editorsSupport || (record.variation && Array.isArray(record.variation.EditorsSupport) ? record.variation.EditorsSupport.slice() : ['cell'])
    };
  }

  function Launcher() {}

  Launcher.prototype.openHostedContext = function(context) {
    var variation = createHostedVariation(context);
    var size = variation.size || [980, 760];

    if (!persistHostedLaunchContext(context)) {
      throw new Error('Не удалось сохранить контекст запуска hosted window');
    }

    if (window.Asc && window.Asc.PluginWindow) {
      var modal = new window.Asc.PluginWindow();
      modal.show(variation);
      return modal;
    }

    var popup = openPopup(variation.url, size);
    if (popup) {
      return popup;
    }

    window.location.href = variation.url;
    return { type: 'redirect' };
  };

  Launcher.prototype.open = function(record) {
    if (!record) {
      throw new Error('Module launch metadata is invalid');
    }

    if (record.launchMode === 'external-link') {
      if (!record.launchUrl) {
        throw new Error('External link is not configured');
      }
      return openExternalLink(record.launchUrl);
    }

    if (record.launchMode === 'hosted-window') {
      return this.openHostedContext(createHostedContext(record));
    }

    if (!record.launchUrl) {
      throw new Error('Module launch metadata is invalid');
    }

    var viewport = getViewport();
    var size = normalizeSize(viewport, record.size, [960, 720]);
    var minSize = normalizeSize(viewport, record.minSize || [640, 480], [640, 480]);
    var maxSize = normalizeSize(viewport, record.maxSize || [1600, 1000], [1600, 1000]);

    maxSize[0] = Math.max(maxSize[0], minSize[0]);
    maxSize[1] = Math.max(maxSize[1], minSize[1]);

    var variation = createVariation(record, size, minSize, maxSize);

    if (window.Asc && window.Asc.PluginWindow) {
      var modal = new window.Asc.PluginWindow();
      modal.show(variation);
      return modal;
    }

    var popup = openPopup(record.launchUrl, size);
    if (popup) {
      return popup;
    }

    window.location.href = record.launchUrl;
    return { type: 'redirect' };
  };

  window.GameLauncher = new Launcher();
})(window);
