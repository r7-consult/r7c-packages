(function(window) {
  'use strict';

  var HOSTED_OLE_MARKER = '__game_wizard_hosted_module__';
  var CHILD_WINDOW_COMMAND_KEY = 'game_wizard_child_window_command';
  var ACTIVE_CHILD_WINDOW_KEY = 'game_wizard_active_child_window_id';

  function getPluginInfo() {
    return window.Asc && window.Asc.plugin ? window.Asc.plugin.info || {} : {};
  }

  function parseHostedOlePayload(text) {
    var parsed;

    if (typeof text !== 'string' || !text.trim()) {
      return null;
    }

    try {
      parsed = JSON.parse(text);
    } catch (_) {
      return null;
    }

    if (!parsed || parsed[HOSTED_OLE_MARKER] !== true) {
      return null;
    }

    return parsed;
  }

  function GameWizardPlugin() {
    this.initialized = false;
    this.busy = false;
    this.service = null;
    this.initText = '';
    this.pendingHostedOleContext = null;
    this.hostedOleLaunchRequested = false;
    this.hostedOleLaunchAttempts = 0;
  }

  GameWizardPlugin.prototype.init = function() {
    if (this.initialized) {
      return;
    }

    if (!window.GameWizardConfig || !window.GameWizardModuleService || !window.GameWizardUI || !window.GameLauncher) {
      throw new Error('Game Wizard modules are not loaded');
    }

    this.resizePanel();
    this.applyTheme(this.readTheme());
    window.GameWizardUI.setBranding(window.GameWizardConfig.branding || {});
    this.service = new window.GameWizardModuleService(window.GameWizardConfig);
    this.initialized = true;
    this.refreshModules('Обновление списка игр...');
  };

  GameWizardPlugin.prototype.resizePanel = function() {
    if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.resizeWindow === 'function') {
      window.Asc.plugin.resizeWindow(460, 820, 360, 520, 900, 1280);
    }
  };

  GameWizardPlugin.prototype.readTheme = function() {
    if (window.Asc && window.Asc.plugin && window.Asc.plugin.theme) {
      return window.Asc.plugin.theme;
    }
    if (window.Asc && window.Asc.plugin && window.Asc.plugin.info && window.Asc.plugin.info.theme) {
      return window.Asc.plugin.info.theme;
    }
    return null;
  };

  GameWizardPlugin.prototype.applyTheme = function(theme) {
    var rawType = '';
    if (theme && typeof theme === 'object' && theme.type) {
      rawType = String(theme.type);
    } else if (typeof theme === 'string') {
      rawType = theme;
    }

    var isDark = rawType.toLowerCase().indexOf('dark') !== -1;
    document.body.classList.toggle('theme-dark', isDark);
  };

  GameWizardPlugin.prototype.setBusy = function(isBusy, message) {
    this.busy = !!isBusy;
    document.body.classList.toggle('gw-busy', this.busy);
    if (message) {
      window.GameWizardUI.setStatus(message);
    }
  };

  GameWizardPlugin.prototype.buildMeta = function(remote) {
    if (remote && remote.configured && remote.available) {
      return 'Локальные и удаленные модули';
    }

    if (remote && remote.configured && !remote.available) {
      return 'Локальный режим: GitHub недоступен';
    }

    return 'Локальный режим: GitHub не настроен';
  };

  GameWizardPlugin.prototype.buildIdleStatus = function(snapshot) {
    if (snapshot.remote && snapshot.remote.configured && snapshot.remote.available) {
      return 'Модулей в списке: ' + snapshot.records.length;
    }

    if (snapshot.remote && snapshot.remote.configured && !snapshot.remote.available) {
      return 'GitHub недоступен, показаны локальные игры';
    }

    return 'Показаны локальные игры';
  };

  GameWizardPlugin.prototype.applySnapshot = function(snapshot, statusMessage) {
    var self = this;

    window.GameWizardUI.renderModules(snapshot.records, {
      onOpen: function(recordId) {
        self.openModule(recordId);
      },
      onInstall: function(recordId) {
        self.installModule(recordId);
      },
      onRemove: function(recordId) {
        self.removeModule(recordId);
      },
      onUpdate: function(recordId) {
        self.updateModule(recordId);
      }
    });

    window.GameWizardUI.setMeta(this.buildMeta(snapshot.remote));
    window.GameWizardUI.setStatus(statusMessage || this.buildIdleStatus(snapshot));
  };

  GameWizardPlugin.prototype.refreshModules = function(statusMessage) {
    var self = this;
    this.setBusy(true, statusMessage || 'Обновление списка игр...');

    this.service.refresh().then(function(snapshot) {
      self.applySnapshot(snapshot);
    }).catch(function(error) {
      var details = error && error.message ? error.message : 'неизвестная ошибка';
      window.GameWizardUI.setMeta('Ошибка обновления каталога');
      window.GameWizardUI.setStatus('Не удалось обновить список игр: ' + details);
      console.error('[GameWizard] Refresh failed', error);
    }).then(function() {
      self.setBusy(false);
    });
  };

  GameWizardPlugin.prototype.withMutation = function(label, operation, successMessage) {
    var self = this;

    if (this.busy) {
      window.GameWizardUI.setStatus('Дождитесь окончания текущей операции');
      return;
    }

    this.setBusy(true, label);

    operation().then(function(snapshot) {
      self.applySnapshot(snapshot, successMessage);
    }).catch(function(error) {
      var details = error && error.message ? error.message : 'неизвестная ошибка';
      window.GameWizardUI.setStatus(label + ' Ошибка: ' + details);
      console.error('[GameWizard] Mutation failed', error);
    }).then(function() {
      self.setBusy(false);
    });
  };

  GameWizardPlugin.prototype.findRecord = function(recordId) {
    var record = this.service.findRecord(recordId);
    if (!record) {
      window.GameWizardUI.setStatus('Модуль не найден');
      return null;
    }
    return record;
  };

  GameWizardPlugin.prototype.openModule = function(recordId) {
    var record = this.findRecord(recordId);
    if (!record || !record.canLaunch) {
      return;
    }

    try {
      window.GameLauncher.open(record);
      window.GameWizardUI.setStatus('Открыто: ' + record.title);
    } catch (error) {
      var details = error && error.message ? error.message : 'неизвестная ошибка';
      window.GameWizardUI.setStatus('Ошибка запуска: ' + record.title + ' (' + details + ')');
      console.error('[GameWizard] Failed to open module', error);
    }
  };

  GameWizardPlugin.prototype.installModule = function(recordId) {
    var record = this.findRecord(recordId);
    if (!record || !record.canInstall) {
      return;
    }

    this.withMutation('Установка ' + record.title + '...', this.service.install.bind(this.service, recordId), 'Установлено: ' + record.title);
  };

  GameWizardPlugin.prototype.removeModule = function(recordId) {
    var record = this.findRecord(recordId);
    if (!record || !record.canRemove) {
      return;
    }

    this.withMutation('Удаление ' + record.title + '...', this.service.remove.bind(this.service, recordId), 'Удалено: ' + record.title);
  };

  GameWizardPlugin.prototype.updateModule = function(recordId) {
    var record = this.findRecord(recordId);
    if (!record || !record.canUpdate) {
      return;
    }

    this.withMutation('Обновление ' + record.title + '...', this.service.update.bind(this.service, recordId), 'Обновлено: ' + record.title);
  };

  GameWizardPlugin.prototype.buildHostedOleLaunchContext = function(text) {
    var payload = parseHostedOlePayload(text);
    var pluginInfo;

    if (!payload || !payload.module) {
      return null;
    }

    pluginInfo = getPluginInfo();

    return {
      title: payload.title || payload.module.id || 'Hosted Module',
      moduleId: payload.module.id || 'module',
      moduleEntry: payload.module.entry || 'index.html',
      moduleGuid: payload.module.guid || '',
      hostGuid: pluginInfo.guid || '',
      initData: typeof payload.data === 'string' ? payload.data : '',
      objectId: pluginInfo.objectId,
      resize: !!pluginInfo.resize,
      width: pluginInfo.width || payload.width || 70,
      height: pluginInfo.height || payload.height || 70,
      mmToPx: pluginInfo.mmToPx || 3.78,
      lang: pluginInfo.lang || 'ru-RU',
      theme: pluginInfo.theme || 'theme-light',
      windowSize: Array.isArray(payload.windowSize) ? payload.windowSize.slice(0, 2) : [980, 760],
      windowMinSize: Array.isArray(payload.windowMinSize) ? payload.windowMinSize.slice(0, 2) : [720, 560],
      windowMaxSize: Array.isArray(payload.windowMaxSize) ? payload.windowMaxSize.slice(0, 2) : [1600, 1000],
      buttons: Array.isArray(payload.buttons) ? payload.buttons.slice() : [
        { text: 'Ok', primary: true, isViewer: false },
        { text: 'Cancel', primary: false }
      ],
      editorsSupport: Array.isArray(payload.editorsSupport) ? payload.editorsSupport.slice() : ['cell']
    };
  };

  GameWizardPlugin.prototype.resolveHostedRecord = function(context) {
    var records = this.service && this.service.lastSnapshot && Array.isArray(this.service.lastSnapshot.records)
      ? this.service.lastSnapshot.records
      : [];
    var index;
    var record;

    for (index = 0; index < records.length; index += 1) {
      record = records[index];
      if (context.moduleGuid && record.guid === context.moduleGuid) {
        return record;
      }
    }

    for (index = 0; index < records.length; index += 1) {
      record = records[index];
      if (!context.moduleId) {
        continue;
      }
      if (record.module === context.moduleId || record.localModule === context.moduleId || record.remoteModule === context.moduleId) {
        return record;
      }
    }

    return null;
  };

  GameWizardPlugin.prototype.hydrateHostedOleContext = function() {
    var context = this.pendingHostedOleContext;
    var record;

    if (!context) {
      return;
    }

    record = this.resolveHostedRecord(context);
    if (!record) {
      return;
    }

    if (!context.title) {
      context.title = record.title || 'Hosted Module';
    }

    if (!context.moduleId) {
      context.moduleId = record.module || record.localModule || record.remoteModule || '';
    }

    if (!context.moduleEntry || context.moduleEntry === 'index.html') {
      context.moduleEntry = record.moduleEntry || record.launchUrl || '';
    }

    if (!context.moduleGuid) {
      context.moduleGuid = record.guid || '';
    }

    if ((!context.windowSize || !context.windowSize.length) && record.size) {
      context.windowSize = record.size.slice();
    }

    if ((!context.windowMinSize || !context.windowMinSize.length) && record.minSize) {
      context.windowMinSize = record.minSize.slice();
    }

    if ((!context.windowMaxSize || !context.windowMaxSize.length) && record.maxSize) {
      context.windowMaxSize = record.maxSize.slice();
    }
  };

  GameWizardPlugin.prototype.sendChildWindowCommand = function(windowId, action) {
    if (!windowId || !action) {
      return false;
    }

    try {
      window.localStorage.setItem(CHILD_WINDOW_COMMAND_KEY, JSON.stringify({
        windowId: windowId,
        action: action,
        ts: Date.now()
      }));
      return true;
    } catch (_) {
      return false;
    }
  };

  GameWizardPlugin.prototype.maybeAutoOpenHostedOleWindow = function() {
    var self = this;

    if (!this.pendingHostedOleContext || this.hostedOleLaunchRequested) {
      return;
    }

    if (window.GameLauncher && typeof window.GameLauncher.openHostedContext === 'function') {
      this.hostedOleLaunchRequested = true;
      window.GameLauncher.openHostedContext(this.pendingHostedOleContext);
      return;
    }

    if (this.hostedOleLaunchAttempts >= 20) {
      console.warn('[GameWizard] Hosted OLE reopen launcher unavailable');
      return;
    }

    this.hostedOleLaunchAttempts += 1;
    window.setTimeout(function() {
      self.maybeAutoOpenHostedOleWindow();
    }, 100);
  };

  GameWizardPlugin.prototype.onButton = function(id, windowId) {
    var targetWindowId = windowId;
    var buttonId = String(id);
    var activeChildWindowId = '';

    if (!targetWindowId && typeof id === 'string' && id) {
      targetWindowId = id;
    }

    if (!targetWindowId) {
      try {
        activeChildWindowId = window.localStorage.getItem(ACTIVE_CHILD_WINDOW_KEY) || '';
      } catch (_) {
        activeChildWindowId = '';
      }
      if (activeChildWindowId) {
        targetWindowId = activeChildWindowId;
      }
    }

    if (
      targetWindowId &&
      window.Asc &&
      window.Asc.plugin &&
      typeof window.Asc.plugin.executeMethod === 'function'
    ) {
      window.Asc.plugin.executeMethod('CloseWindow', [targetWindowId]);
      return;
    }

    if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.executeCommand === 'function') {
      window.Asc.plugin.executeCommand('close', '');
    }
  };

  var plugin = new GameWizardPlugin();

  function runWhenDomReady(callback) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', callback, { once: true });
      return;
    }
    callback();
  }

  if (window.Asc && window.Asc.plugin) {
    window.Asc.plugin.init = function() {
      runWhenDomReady(function() {
        plugin.init();
      });
    };

    window.Asc.plugin.button = function(id, windowId) {
      plugin.onButton(id, windowId);
    };

    window.Asc.plugin.onThemeChanged = function(theme) {
      if (window.Asc.plugin.onThemeChangedBase) {
        window.Asc.plugin.onThemeChangedBase(theme);
      }
      plugin.applyTheme(theme);
    };
  } else {
    runWhenDomReady(function() {
      plugin.init();
    });
  }
})(window);

