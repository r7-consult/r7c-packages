(function(window) {
  'use strict';

  var HOSTED_OLE_MARKER = '__data_processing_hosted_module__';
  var CHILD_WINDOW_COMMAND_KEY = 'data_processing_child_window_command';
  var ACTIVE_CHILD_WINDOW_KEY = 'data_processing_active_child_window_id';
  var r7cFlyoutLastSeenDateKey = 'r7c_flyout_last_seen_date';

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
    this.r7cFlyout = null;
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
      throw new Error('Data Processing shell modules are not loaded');
    }

    this.resizePanel();
    this.applyTheme(this.readTheme());
    window.GameWizardUI.setBranding(window.GameWizardConfig.branding || {});
    this.setupR7cFlyout();
    this.service = new window.GameWizardModuleService(window.GameWizardConfig);
    this.initialized = true;
    this.refreshModules('Обновление каталога решений...');
  };

  GameWizardPlugin.prototype.openExternalUrl = function(url) {
    var popup = null;
    var anchor = null;

    if (!url) {
      return;
    }

    try {
      if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.executeMethod === 'function') {
        window.Asc.plugin.executeMethod('OpenLink', [url]);
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
      return;
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
      return;
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
    } catch (_) {
    } finally {
      if (anchor && anchor.parentNode) {
        anchor.parentNode.removeChild(anchor);
      }
    }
  };

  GameWizardPlugin.prototype.getLocalDateStamp = function() {
    var date = new Date();
    var yyyy = date.getFullYear();
    var mm = String(date.getMonth() + 1).padStart(2, '0');
    var dd = String(date.getDate()).padStart(2, '0');
    return yyyy + '-' + mm + '-' + dd;
  };

  GameWizardPlugin.prototype.canShowR7cFlyoutToday = function() {
    try {
      return window.localStorage.getItem(r7cFlyoutLastSeenDateKey) !== this.getLocalDateStamp();
    } catch (_) {
      return true;
    }
  };

  GameWizardPlugin.prototype.hideR7cFlyoutForToday = function() {
    if (this.r7cFlyout) {
      this.r7cFlyout.classList.add('hidden');
    }

    try {
      window.localStorage.setItem(r7cFlyoutLastSeenDateKey, this.getLocalDateStamp());
    } catch (_) {}
  };

  GameWizardPlugin.prototype.setupR7cFlyout = function() {
    var self = this;
    var openTarget;

    this.r7cFlyout = document.getElementById('r7c-flyout');
    if (!this.r7cFlyout) {
      return;
    }

    if (!this.canShowR7cFlyoutToday()) {
      this.r7cFlyout.classList.add('hidden');
      return;
    }

    this.r7cFlyout.classList.remove('hidden');

    openTarget = function(event) {
      if (event) {
        event.preventDefault();
        event.stopPropagation();
      }
      self.openExternalUrl('https://data.slider-ai.ru/?utm_source=data_processing');
      self.hideR7cFlyoutForToday();
    };

    this.r7cFlyout.addEventListener('click', openTarget);
    this.r7cFlyout.addEventListener('keydown', function(event) {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        event.stopPropagation();
        openTarget();
      }
    });
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
      window.GameWizardUI.setMeta(message);
      window.GameWizardUI.setStatus(message);
    }
  };

  GameWizardPlugin.prototype.buildMeta = function(remote) {
    if (remote && remote.configured && remote.available) {
      return 'Локальные и удаленные решения';
    }

    if (remote && remote.configured && !remote.available) {
      return 'Локальный режим: GitHub недоступен';
    }

    return 'Локальный режим: GitHub не настроен';
  };

  GameWizardPlugin.prototype.buildIdleStatus = function(snapshot) {
    if (snapshot.remote && snapshot.remote.configured && snapshot.remote.available) {
      return 'Решений в списке: ' + snapshot.records.length;
    }

    if (snapshot.remote && snapshot.remote.configured && !snapshot.remote.available) {
      return 'GitHub недоступен, показаны локальные решения';
    }

    return 'Показаны локальные решения';
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
    this.setBusy(true, statusMessage || 'Обновление каталога решений...');

    this.service.refresh().then(function(snapshot) {
      self.applySnapshot(snapshot);
    }).catch(function(error) {
      var details = error && error.message ? error.message : 'неизвестная ошибка';
      window.GameWizardUI.setMeta('Ошибка обновления каталога');
      window.GameWizardUI.setStatus('Не удалось обновить каталог: ' + details);
      console.error('[GameWizard] Refresh failed', error);
    }).then(function() {
      self.setBusy(false);
    });
  };

  GameWizardPlugin.prototype.withMutation = function(label, operation, successMessage, pendingState) {
    var self = this;

    if (this.busy) {
      window.GameWizardUI.setStatus('Дождитесь окончания текущей операции');
      return;
    }

    if (pendingState && pendingState.recordId && pendingState.actionType) {
      window.GameWizardUI.setPendingAction(pendingState.recordId, pendingState.actionType, pendingState.busyLabel);
    }

    this.setBusy(true, label);

    operation().then(function(snapshot) {
      self.applySnapshot(snapshot, successMessage);
    }).catch(function(error) {
      var details = error && error.message ? error.message : 'неизвестная ошибка';
      window.GameWizardUI.setMeta('Ошибка операции');
      window.GameWizardUI.setStatus(label + ' Ошибка: ' + details);
      console.error('[GameWizard] Mutation failed', error);
    }).then(function() {
      window.GameWizardUI.setPendingAction();
      self.setBusy(false);
    });
  };

  GameWizardPlugin.prototype.findRecord = function(recordId) {
    var record = this.service.findRecord(recordId);
    if (!record) {
      window.GameWizardUI.setStatus('Решение не найдено');
      return null;
    }
    return record;
  };

  GameWizardPlugin.prototype.openModule = function(recordId) {
    var record = this.findRecord(recordId);
    var statusPrefix = '';

    if (!record || !record.canLaunch) {
      return;
    }

    try {
      this.hideR7cFlyoutForToday();
      window.GameLauncher.open(record);
      statusPrefix = record.launchMode === 'external-link' ? 'Открыта ссылка: ' : 'Открыто: ';
      window.GameWizardUI.setStatus(statusPrefix + record.title);
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

    this.withMutation('Установка ' + record.title + '...', this.service.install.bind(this.service, recordId), 'Установлено: ' + record.title, {
      recordId: recordId,
      actionType: 'install',
      busyLabel: 'Установка...'
    });
  };

  GameWizardPlugin.prototype.removeModule = function(recordId) {
    var record = this.findRecord(recordId);
    if (!record || !record.canRemove) {
      return;
    }

    this.withMutation('Удаление ' + record.title + '...', this.service.remove.bind(this.service, recordId), 'Удалено: ' + record.title, {
      recordId: recordId,
      actionType: 'remove',
      busyLabel: 'Удаление...'
    });
  };

  GameWizardPlugin.prototype.updateModule = function(recordId) {
    var record = this.findRecord(recordId);
    if (!record || !record.canUpdate) {
      return;
    }

    this.withMutation('Обновление ' + record.title + '...', this.service.update.bind(this.service, recordId), 'Обновлено: ' + record.title, {
      recordId: recordId,
      actionType: 'update',
      busyLabel: 'Обновление...'
    });
  };

  GameWizardPlugin.prototype.buildHostedOleLaunchContext = function(text) {
    var payload = parseHostedOlePayload(text);
    var pluginInfo;

    if (!payload || !payload.module) {
      return null;
    }

    pluginInfo = getPluginInfo();

    return {
      title: payload.title || payload.module.id || 'Модуль обработки данных',
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
      context.title = record.title || 'Модуль обработки данных';
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
    var targetWindowId = windowId || '';

    if (!targetWindowId && typeof id === 'string' && id && id !== '-1') {
      targetWindowId = id;
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
      return;
    }

    try {
      window.close();
    } catch (_) {}
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
