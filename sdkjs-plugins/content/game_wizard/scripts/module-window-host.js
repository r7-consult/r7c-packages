(function(window) {
  'use strict';

  var HOSTED_OLE_MARKER = '__game_wizard_hosted_module__';
  var HOSTED_OLE_VERSION = 1;
  var HOSTED_LAUNCH_CONTEXT_KEY = 'game_wizard_hosted_launch_context';
  var ACTIVE_CHILD_WINDOW_KEY = 'game_wizard_active_child_window_id';
  var CHILD_WINDOW_COMMAND_KEY = 'game_wizard_child_window_command';

  var runtimeState = {
    launchContext: null,
    isSaving: false
  };

  function byId(id) {
    return document.getElementById(id);
  }

  function setStatus(text) {
    var node = byId('module-window-status');
    if (node) {
      node.textContent = text;
    }
  }

  function log(message) {
    if (window.console && typeof window.console.log === 'function') {
      window.console.log('[module-window-host] ' + message);
    }
  }

  function getHostPlugin() {
    return window.Asc && window.Asc.plugin ? window.Asc.plugin : null;
  }

  function getRootBridgePlugin() {
    var parentWindow;
    var index;
    var candidate;
    var href;

    try {
      parentWindow = window.parent;
      if (!parentWindow || parentWindow === window || !parentWindow.frames) {
        return null;
      }

      for (index = 0; index < parentWindow.frames.length; index += 1) {
        candidate = parentWindow.frames[index];
        if (!candidate || candidate === window) {
          continue;
        }

        if (!candidate.Asc || !candidate.Asc.plugin) {
          continue;
        }

        href = '';
        try {
          href = String(candidate.location && candidate.location.href || '');
        } catch (_) {
          href = '';
        }

        if (href.indexOf('/module-window-host.html') !== -1) {
          continue;
        }

        if (href.indexOf('/index.html') !== -1 || !href) {
          return candidate.Asc.plugin;
        }
      }
    } catch (_) {
      return null;
    }

    return null;
  }

  function getBridgePlugin() {
    return getRootBridgePlugin() || getHostPlugin();
  }

  function getHostPluginInfo() {
    var hostPlugin = getBridgePlugin() || getHostPlugin();
    return hostPlugin ? hostPlugin.info || {} : {};
  }

  function buildFullUrl(fileName) {
    var currentUrl = String(window.location.href || '');
    var baseUrl = currentUrl.substring(0, currentUrl.lastIndexOf('/') + 1);
    return baseUrl + fileName;
  }

  function readStorageJson(key) {
    try {
      var raw = window.localStorage.getItem(key);
      return raw ? JSON.parse(raw) : null;
    } catch (_) {
      return null;
    }
  }

  function removeStorageKey(key) {
    try {
      window.localStorage.removeItem(key);
    } catch (_) {}
  }

  function readLaunchContext() {
    var stored = readStorageJson(HOSTED_LAUNCH_CONTEXT_KEY);
    removeStorageKey(HOSTED_LAUNCH_CONTEXT_KEY);
    return stored;
  }

  function normalizeLaunchContext(rawContext) {
    var hostInfo = getHostPluginInfo();
    var context = rawContext || {};

    return {
      title: context.title || 'Hosted Module',
      moduleId: context.moduleId || 'module',
      moduleEntry: context.moduleEntry || 'index.html',
      moduleGuid: context.moduleGuid || '',
      hostGuid: context.hostGuid || hostInfo.guid || '',
      initData: typeof context.initData === 'string' ? context.initData : '',
      objectId: context.objectId,
      resize: !!context.resize,
      width: context.width || hostInfo.width || 70,
      height: context.height || hostInfo.height || 70,
      mmToPx: context.mmToPx || hostInfo.mmToPx || 3.78,
      lang: context.lang || hostInfo.lang || 'ru-RU',
      theme: context.theme || hostInfo.theme || 'theme-light',
      windowSize: Array.isArray(context.windowSize) ? context.windowSize.slice(0, 2) : [980, 760],
      windowMinSize: Array.isArray(context.windowMinSize) ? context.windowMinSize.slice(0, 2) : [720, 560],
      windowMaxSize: Array.isArray(context.windowMaxSize) ? context.windowMaxSize.slice(0, 2) : [1600, 1000],
      buttons: Array.isArray(context.buttons) ? context.buttons.slice() : [],
      editorsSupport: Array.isArray(context.editorsSupport) ? context.editorsSupport.slice() : ['cell']
    };
  }

  function setActiveChildWindowId(windowId) {
    if (!windowId) {
      return;
    }

    try {
      window.localStorage.setItem(ACTIVE_CHILD_WINDOW_KEY, windowId);
    } catch (_) {}
  }

  function clearActiveChildWindowId(windowId) {
    try {
      var current = window.localStorage.getItem(ACTIVE_CHILD_WINDOW_KEY);
      if (!windowId || current === windowId) {
        window.localStorage.removeItem(ACTIVE_CHILD_WINDOW_KEY);
      }
    } catch (_) {}
  }

  function clearChildWindowCommand() {
    try {
      window.localStorage.removeItem(CHILD_WINDOW_COMMAND_KEY);
    } catch (_) {}
  }

  function resolveWindowId(explicitWindowId) {
    if (explicitWindowId) {
      return explicitWindowId;
    }

    try {
      var searchId = new URLSearchParams(window.location.search).get('windowID');
      if (searchId) {
        return searchId;
      }
    } catch (_) {}

    var hostPlugin = getHostPlugin();
    if (hostPlugin) {
      if (hostPlugin.windowID) {
        return hostPlugin.windowID;
      }
      if (hostPlugin.info && hostPlugin.info.windowID) {
        return hostPlugin.info.windowID;
      }
    }

    return '';
  }

  function closeWindow(explicitWindowId) {
    var bridgePlugin = getBridgePlugin();
    var hostPlugin = getHostPlugin();
    var windowId = resolveWindowId(explicitWindowId);
    clearActiveChildWindowId(windowId);

    if (bridgePlugin && typeof bridgePlugin.executeMethod === 'function') {
      try {
        if (windowId) {
          bridgePlugin.executeMethod('CloseWindow', [windowId], function() {});
          return;
        }
      } catch (_) {}

      try {
        bridgePlugin.executeMethod('CloseWindow', [], function() {});
        return;
      } catch (_) {}
    }

    if (hostPlugin && typeof hostPlugin.executeCommand === 'function') {
      try {
        hostPlugin.executeCommand('close', '');
        return;
      } catch (_) {}
    }

    try {
      window.close();
    } catch (_) {}
  }

  function buildHostedOleData(context, oleParams) {
    return JSON.stringify({
      version: HOSTED_OLE_VERSION,
      title: context.title,
      module: {
        id: context.moduleId,
        entry: context.moduleEntry,
        guid: context.moduleGuid
      },
      data: oleParams && oleParams.data != null ? String(oleParams.data) : '',
      width: oleParams && oleParams.width ? oleParams.width : context.width,
      height: oleParams && oleParams.height ? oleParams.height : context.height,
      resize: oleParams && typeof oleParams.resize !== 'undefined' ? !!oleParams.resize : !!context.resize,
      objectId: oleParams ? oleParams.objectId : context.objectId,
      windowSize: context.windowSize,
      windowMinSize: context.windowMinSize,
      windowMaxSize: context.windowMaxSize,
      buttons: context.buttons,
      editorsSupport: context.editorsSupport,
      __game_wizard_hosted_module__: true
    });
  }

  function buildHostedOleParams(originalParams, context) {
    var hostInfo = getHostPluginInfo();
    var params = {};
    var key;

    for (key in originalParams) {
      if (Object.prototype.hasOwnProperty.call(originalParams, key)) {
        params[key] = originalParams[key];
      }
    }

    params.guid = context.hostGuid || hostInfo.guid || params.guid || context.moduleGuid;
    params.width = params.width || context.width;
    params.height = params.height || context.height;
    params.mmToPx = params.mmToPx || context.mmToPx || hostInfo.mmToPx || 3.78;
    params.widthPix = params.widthPix || ((params.mmToPx * params.width) >> 0);
    params.heightPix = params.heightPix || ((params.mmToPx * params.height) >> 0);
    params.objectId = typeof params.objectId === 'undefined' ? context.objectId : params.objectId;
    params.resize = typeof params.resize === 'undefined' ? !!context.resize : !!params.resize;
    params.data = buildHostedOleData(context, params);

    return params;
  }

  function proxyExecuteMethod(method, args, callback) {
    var bridgePlugin = getBridgePlugin();
    var hostArgs = Array.isArray(args) ? args.slice() : [];

    if ((method === 'AddOleObject' || method === 'EditOleObject') && hostArgs[0] && typeof hostArgs[0] === 'object') {
      hostArgs = [buildHostedOleParams(hostArgs[0], runtimeState.launchContext)];
    }

    if (bridgePlugin && typeof bridgePlugin.executeMethod === 'function') {
      return bridgePlugin.executeMethod(method, hostArgs, typeof callback === 'function' ? callback : function() {});
    }

    if (typeof callback === 'function') {
      callback({
        shim: true,
        error: 'executeMethod unavailable',
        method: method,
        args: hostArgs
      });
    }
  }

  function proxyExecuteCommand(command, value) {
    var bridgePlugin = getBridgePlugin();

    if (command === 'close') {
      closeWindow();
      return;
    }

    if (bridgePlugin && typeof bridgePlugin.executeCommand === 'function') {
      return bridgePlugin.executeCommand(command, value);
    }
  }

  function applyShim(frameWindow) {
    var hostInfo = getHostPluginInfo();
    var context = runtimeState.launchContext;
    var scopeKey;

    frameWindow.Asc = frameWindow.Asc || {};
    frameWindow.Asc.plugin = frameWindow.Asc.plugin || {};
    frameWindow.Asc.plugin.info = frameWindow.Asc.plugin.info || {};

    frameWindow.Asc.plugin.info.guid = context.hostGuid || hostInfo.guid || frameWindow.Asc.plugin.info.guid || context.moduleGuid;
    frameWindow.Asc.plugin.info.originalGuid = context.moduleGuid || '';
    frameWindow.Asc.plugin.info.moduleId = context.moduleId;
    frameWindow.Asc.plugin.info.moduleEntry = context.moduleEntry;
    frameWindow.Asc.plugin.info.data = context.initData;
    frameWindow.Asc.plugin.info.objectId = context.objectId;
    frameWindow.Asc.plugin.info.resize = !!context.resize;
    frameWindow.Asc.plugin.info.lang = context.lang;
    frameWindow.Asc.plugin.info.theme = context.theme;
    frameWindow.Asc.plugin.info.windowID = resolveWindowId();
    frameWindow.Asc.plugin.info.width = context.width;
    frameWindow.Asc.plugin.info.height = context.height;
    frameWindow.Asc.plugin.info.mmToPx = context.mmToPx;
    frameWindow.Asc.plugin.info.widthPix = ((context.width * context.mmToPx) >> 0);
    frameWindow.Asc.plugin.info.heightPix = ((context.height * context.mmToPx) >> 0);

    frameWindow.Asc.plugin.executeMethod = function(method, args, callback) {
      return proxyExecuteMethod(method, args, callback);
    };

    frameWindow.Asc.scope = frameWindow.Asc.scope || {};

    frameWindow.Asc.plugin.executeCommand = function(command, value) {
      return proxyExecuteCommand(command, value);
    };

    frameWindow.Asc.plugin.callCommand = function(func, isClose, isCalc, callback) {
      var bridgePlugin = getBridgePlugin();

      if (window.Asc && frameWindow.Asc && window.Asc.scope && frameWindow.Asc.scope) {
        for (scopeKey in window.Asc.scope) {
          if (Object.prototype.hasOwnProperty.call(window.Asc.scope, scopeKey)) {
            delete window.Asc.scope[scopeKey];
          }
        }
        for (scopeKey in frameWindow.Asc.scope) {
          if (Object.prototype.hasOwnProperty.call(frameWindow.Asc.scope, scopeKey)) {
            window.Asc.scope[scopeKey] = frameWindow.Asc.scope[scopeKey];
          }
        }
      }

      if (bridgePlugin && typeof bridgePlugin.callCommand === 'function') {
        return bridgePlugin.callCommand(func, !!isClose, !!isCalc, typeof callback === 'function' ? callback : function() {});
      }

      if (typeof callback === 'function') {
        callback({ shim: true, error: 'callCommand unavailable' });
      }
    };

    frameWindow.Asc.plugin.resizeWindow = function(w, h, minW, minH, maxW, maxH) {
      var bridgePlugin = getBridgePlugin();
      if (bridgePlugin && typeof bridgePlugin.resizeWindow === 'function') {
        return bridgePlugin.resizeWindow(w, h, minW, minH, maxW, maxH);
      }
    };

    if (typeof frameWindow.Asc.plugin.button !== 'function') {
      frameWindow.Asc.plugin.button = function() {};
    }

    if (typeof frameWindow.Asc.plugin.onExternalMouseUp !== 'function') {
      frameWindow.Asc.plugin.onExternalMouseUp = function() {};
    }
  }

  function getHostedFramePlugin() {
    var frame = byId('module-window-frame');
    if (!frame || !frame.contentWindow || !frame.contentWindow.Asc || !frame.contentWindow.Asc.plugin) {
      return null;
    }
    return frame.contentWindow.Asc.plugin;
  }

  function forwardModuleButton(id) {
    var framePlugin = getHostedFramePlugin();

    if (framePlugin && typeof framePlugin.button === 'function') {
      if (String(id) === '0') {
        runtimeState.isSaving = true;
        setStatus('Сохранение состояния модуля...');
      }
      framePlugin.button(id);
      return;
    }

    if (id !== 0) {
      closeWindow();
    }
  }

  function handleChildWindowCommandValue(rawValue) {
    var command;
    var windowId;

    if (!rawValue) {
      return;
    }

    try {
      command = JSON.parse(rawValue);
    } catch (_) {
      return;
    }

    windowId = resolveWindowId();
    if (!command || command.windowId !== windowId) {
      return;
    }

    clearChildWindowCommand();

    if (command.action === 'save') {
      forwardModuleButton(0);
    } else if (command.action === 'close') {
      closeWindow(windowId);
    }
  }

  function bindCommandBridge() {
    if (window.__gameWizardChildCommandBound) {
      return;
    }

    window.__gameWizardChildCommandBound = true;

    window.addEventListener('storage', function(event) {
      if (event.key === CHILD_WINDOW_COMMAND_KEY) {
        handleChildWindowCommandValue(event.newValue);
      }
    });

    window.setInterval(function() {
      try {
        handleChildWindowCommandValue(window.localStorage.getItem(CHILD_WINDOW_COMMAND_KEY));
      } catch (_) {}
    }, 350);
  }

  function bootHostedFrame() {
    var frame = byId('module-window-frame');
    var context = runtimeState.launchContext;

    if (!frame || frame.__booted) {
      return;
    }

    frame.__booted = true;
    frame.src = buildFullUrl(context.moduleEntry);
    setStatus('Загрузка модуля...');

    frame.onload = function() {
      try {
        var frameWindow = frame.contentWindow;
        applyShim(frameWindow);

        if (frameWindow.Asc && frameWindow.Asc.plugin && typeof frameWindow.Asc.plugin.init === 'function') {
          frameWindow.Asc.plugin.init(context.initData || '');
        }

        setStatus('Модуль запущен');
      } catch (error) {
        setStatus('Ошибка запуска модуля');
        log('frame error -> ' + error.message);
      }
    };
  }

  function start() {
    runtimeState.launchContext = normalizeLaunchContext(readLaunchContext());
    setActiveChildWindowId(resolveWindowId());
    bindCommandBridge();
    bootHostedFrame();
  }

  document.addEventListener('DOMContentLoaded', start);

  window.addEventListener('beforeunload', function() {
    clearActiveChildWindowId(resolveWindowId());
  });

  if (window.Asc && window.Asc.plugin) {
    var previousInit = window.Asc.plugin.init;
    window.Asc.plugin.init = function() {
      if (typeof previousInit === 'function') {
        previousInit.apply(this, arguments);
      }
      start();
    };

    window.Asc.plugin.button = function(id, windowId) {
      if (String(id) === '0') {
        forwardModuleButton(0);
        return;
      }

      if (String(id) === '1') {
        closeWindow(windowId);
        return;
      }

      closeWindow(windowId);
    };
  }
})(window);
