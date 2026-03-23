(function(window) {
  'use strict';

  function uniqueArray(items) {
    var seen = Object.create(null);
    var result = [];

    for (var i = 0; i < items.length; i += 1) {
      var value = items[i];
      if (!value || seen[value]) {
        continue;
      }
      seen[value] = true;
      result.push(value);
    }

    return result;
  }

  function createLookup(items) {
    var lookup = Object.create(null);

    if (!Array.isArray(items)) {
      return lookup;
    }

    for (var i = 0; i < items.length; i += 1) {
      if (items[i]) {
        lookup[items[i]] = true;
      }
    }

    return lookup;
  }

  function sortByDepthDesc(items) {
    return items.slice().sort(function(left, right) {
      return String(right || '').length - String(left || '').length;
    });
  }

  function normalizeSlashes(value) {
    return String(value || '').replace(/\\/g, '/');
  }

  function trimSlashes(value) {
    return normalizeSlashes(value).replace(/^\/+|\/+$/g, '');
  }

  function joinUrlPath() {
    var parts = [];

    for (var i = 0; i < arguments.length; i += 1) {
      var part = trimSlashes(arguments[i]);
      if (part) {
        parts.push(part);
      }
    }

    return parts.join('/');
  }

  function joinNativePath(base) {
    var path = String(base || '');
    var separator = path.indexOf('\\') !== -1 ? '\\' : '/';

    for (var i = 1; i < arguments.length; i += 1) {
      var segment = String(arguments[i] || '');
      if (!segment) {
        continue;
      }

      segment = segment.replace(/[\\/]+/g, separator).replace(new RegExp('^' + escapeRegExp(separator) + '+|' + escapeRegExp(separator) + '+$', 'g'), '');
      if (!segment) {
        continue;
      }

      if (!path) {
        path = segment;
        continue;
      }

      if (path.slice(-1) !== separator) {
        path += separator;
      }
      path += segment;
    }

    return path;
  }

  function dirnameNative(filePath) {
    return String(filePath || '').replace(/[\\/][^\\/]+$/, '');
  }

  function escapeRegExp(value) {
    return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function encodeUtf8ToBase64(text) {
    var bytes = new TextEncoder().encode(String(text == null ? '' : text));
    var binary = '';
    for (var i = 0; i < bytes.length; i += 1) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  function arrayBufferToBase64(buffer) {
    var bytes = new Uint8Array(buffer);
    var binary = '';
    for (var i = 0; i < bytes.length; i += 1) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  function parseJson(text, label) {
    try {
      return JSON.parse(text);
    } catch (error) {
      throw new Error('Не удалось разобрать JSON ' + label + ': ' + error.message);
    }
  }

  function compareVersions(a, b) {
    var left = String(a || '').trim();
    var right = String(b || '').trim();

    if (left === right) {
      return 0;
    }

    var leftParts = left.split('.');
    var rightParts = right.split('.');
    var numeric = true;

    for (var i = 0; i < leftParts.length; i += 1) {
      if (!/^\d+$/.test(leftParts[i])) {
        numeric = false;
        break;
      }
    }

    for (var j = 0; j < rightParts.length; j += 1) {
      if (!/^\d+$/.test(rightParts[j])) {
        numeric = false;
        break;
      }
    }

    if (!numeric) {
      return left < right ? -1 : 1;
    }

    var maxLength = Math.max(leftParts.length, rightParts.length);
    for (var index = 0; index < maxLength; index += 1) {
      var leftValue = index < leftParts.length ? parseInt(leftParts[index], 10) : 0;
      var rightValue = index < rightParts.length ? parseInt(rightParts[index], 10) : 0;

      if (leftValue < rightValue) {
        return -1;
      }
      if (leftValue > rightValue) {
        return 1;
      }
    }

    return 0;
  }

  function deriveIconText(title) {
    var normalized = String(title || '').trim();
    if (!normalized) {
      return 'GM';
    }

    if (/^\d/.test(normalized)) {
      return normalized.slice(0, 2);
    }

    var words = normalized.split(/\s+/).filter(Boolean);
    if (!words.length) {
      return normalized.slice(0, 2).toUpperCase();
    }

    if (words.length === 1) {
      return words[0].slice(0, 2).toUpperCase();
    }

    return (words[0].charAt(0) + words[1].charAt(0)).toUpperCase();
  }

  function pickLocalizedValue(raw, locale) {
    if (!raw) {
      return '';
    }

    if (typeof raw === 'string') {
      return raw;
    }

    if (locale && raw[locale]) {
      return raw[locale];
    }

    var shortLocale = locale ? locale.split('-')[0] : '';
    if (shortLocale && raw[shortLocale]) {
      return raw[shortLocale];
    }

    if (raw.ru) {
      return raw.ru;
    }

    if (raw.en) {
      return raw.en;
    }

    for (var key in raw) {
      if (Object.prototype.hasOwnProperty.call(raw, key) && raw[key]) {
        return raw[key];
      }
    }

    return '';
  }

  function pickVariation(config) {
    return config && Array.isArray(config.variations) && config.variations.length ? config.variations[0] : {};
  }

  function resolveAssetUrl(baseUrl, relativePath) {
    if (!relativePath) {
      return '';
    }

    try {
      return new URL(relativePath, baseUrl).toString();
    } catch (_) {
      return '';
    }
  }

  function pickIconPath(variation) {
    if (!variation || !Array.isArray(variation.icons) || !variation.icons.length) {
      return '';
    }

    for (var i = 0; i < variation.icons.length; i += 1) {
      if (variation.icons[i]) {
        return variation.icons[i];
      }
    }

    return '';
  }

  function resolveNativePluginDirectory() {
    var href = window.location.href || '';
    var protocol = '';
    var pathname = '';

    try {
      var url = new URL(href);
      protocol = url.protocol;
      pathname = decodeURIComponent(url.pathname || '');
    } catch (_) {
      pathname = decodeURIComponent(window.location.pathname || '');
    }

    if (!pathname) {
      return {
        protocol: protocol || 'unknown',
        href: href,
        directory: '',
        isNativePath: false
      };
    }

    if (/^\/[A-Za-z]:\//.test(pathname)) {
      pathname = pathname.slice(1).replace(/\//g, '\\');
    }

    var directory = pathname.replace(/[\\/][^\\/]+$/, '');
    var isNativePath = /^[A-Za-z]:\\/.test(directory) || directory.indexOf('/') === 0;

    return {
      protocol: protocol || 'unknown',
      href: href,
      directory: directory,
      isNativePath: isNativePath
    };
  }

  function buildFileUrl(filePath, isDirectory) {
    var normalized = normalizeSlashes(filePath);
    if (isDirectory && normalized.slice(-1) !== '/') {
      normalized += '/';
    }

    if (/^[A-Za-z]:\//.test(normalized)) {
      return 'file:///' + encodeURI(normalized);
    }

    if (normalized.indexOf('/') === 0) {
      return 'file://' + encodeURI(normalized);
    }

    return 'file:///' + encodeURI(normalized);
  }

  function xhrRequest(url, responseType) {
    return new Promise(function(resolve, reject) {
      var xhr = new XMLHttpRequest();

      try {
        xhr.open('GET', url, true);
      } catch (error) {
        reject(error);
        return;
      }

      if (responseType) {
        xhr.responseType = responseType;
      }

      xhr.onload = function() {
        if (this.status >= 200 && this.status < 300 || this.status === 0) {
          resolve(this.responseType ? this.response : this.responseText);
          return;
        }
        reject(new Error('HTTP ' + this.status + ' for ' + url));
      };

      xhr.onerror = function() {
        reject(new Error('XHR failed for ' + url));
      };

      xhr.send(null);
    });
  }

  function fetchText(url, headers) {
    if (typeof window.fetch === 'function' && /^https?:/i.test(url)) {
      return window.fetch(url, {
        method: 'GET',
        headers: headers || {}
      }).then(function(response) {
        if (!response.ok) {
          throw new Error('HTTP ' + response.status + ' for ' + url);
        }
        return response.text();
      });
    }

    return xhrRequest(url, 'text');
  }

  function fetchArrayBuffer(url, headers) {
    if (typeof window.fetch === 'function' && /^https?:/i.test(url)) {
      return window.fetch(url, {
        method: 'GET',
        headers: headers || {}
      }).then(function(response) {
        if (!response.ok) {
          throw new Error('HTTP ' + response.status + ' for ' + url);
        }
        return response.arrayBuffer();
      });
    }

    return xhrRequest(url, 'arraybuffer');
  }

  function parseDirectoryListing(html) {
    if (!html) {
      return [];
    }

    var doc = new DOMParser().parseFromString(String(html), 'text/html');
    var links = doc.querySelectorAll('a[href]');
    var entries = [];

    for (var i = 0; i < links.length; i += 1) {
      var rawHref = links[i].getAttribute('href') || '';
      var href = rawHref.split('#')[0].split('?')[0];

      if (!href || href === '/' || href === './' || href === '../') {
        continue;
      }

      href = decodeURIComponent(href).replace(/^\/+/, '');
      entries.push(href);
    }

    return uniqueArray(entries);
  }

  function parseCatalog(payload) {
    if (!payload || !Array.isArray(payload.modules)) {
      return [];
    }

    var result = [];
    for (var i = 0; i < payload.modules.length; i += 1) {
      var item = payload.modules[i];
      if (item && item.module) {
        result.push(String(item.module));
      }
    }
    return uniqueArray(result);
  }

  function deriveDirectoriesFromFiles(files) {
    var directories = [];

    for (var i = 0; i < files.length; i += 1) {
      var relativePath = normalizeSlashes(files[i]);
      if (!relativePath) {
        continue;
      }

      var parts = relativePath.split('/');
      while (parts.length > 1) {
        parts.pop();
        directories.push(parts.join('/'));
      }
    }

    return sortByDepthDesc(uniqueArray(directories.filter(Boolean)));
  }

  function parseGitHubRepository(repositoryUrl, fallbackBranch) {
    var raw = String(repositoryUrl || '').trim();
    if (!raw) {
      return null;
    }

    var normalized = raw.replace(/\.git$/i, '').replace(/\/+$/, '');
    var match = normalized.match(/^https:\/\/github\.com\/([^\/]+)\/([^\/]+?)(?:\/tree\/([^\/]+))?$/i);

    if (!match) {
      return null;
    }

    return {
      owner: match[1],
      repo: match[2],
      branch: match[3] || fallbackBranch || 'main'
    };
  }

  function isGenericMainIconPath(iconPath) {
    var normalized = normalizeSlashes(iconPath || '').toLowerCase();
    return normalized.indexOf('/resources/icons/main/') !== -1 || normalized.indexOf('resources/icons/main/') === 0;
  }

  function createStandaloneModuleRecord(options) {
    var config = options.config || {};
    var variation = pickVariation(config);
    var locale = options.locale;
    var title = pickLocalizedValue(config.nameLocale, locale) || config.name || options.module;
    var description = pickLocalizedValue(variation.descriptionLocale, locale) || variation.description || 'Модуль плагина';
    var managerWindow = config.managerWindow || {};
    var launchMode = managerWindow.launchMode || 'direct';
    var managerEntry = managerWindow.moduleEntry || variation.url || 'index.html';
    var launchUrl = launchMode === 'external-link'
      ? String(managerWindow.launchUrl || '')
      : (options.launchUrl || resolveAssetUrl(options.baseUrl, variation.url || 'index.html'));
    var iconPath = pickIconPath(variation);
    var resolvedIconUrl = resolveAssetUrl(options.baseUrl, iconPath);

    if (!resolvedIconUrl || isGenericMainIconPath(iconPath)) {
      resolvedIconUrl = resolveAssetUrl(options.baseUrl, 'resources/card-icon.svg');
    }

    return {
      module: options.module,
      catalogIndex: typeof options.catalogIndex === 'number' ? options.catalogIndex : 9999,
      guid: config.guid || ('module:' + options.module),
      title: title,
      description: description,
      version: String(config.version || '0.0.0'),
      entry: variation.url || 'index.html',
      moduleEntry: launchMode === 'external-link' ? '' : joinUrlPath(options.relativeBasePath || '', managerEntry),
      iconUrl: resolvedIconUrl,
      iconText: deriveIconText(title),
      launchMode: launchMode,
      actionLabel: managerWindow.actionLabel || '',
      badge: managerWindow.badge || '',
      size: Array.isArray(managerWindow.size) ? managerWindow.size.slice(0, 2) : (Array.isArray(variation.size) ? variation.size.slice(0, 2) : null),
      minSize: Array.isArray(managerWindow.minSize) ? managerWindow.minSize.slice(0, 2) : null,
      maxSize: Array.isArray(managerWindow.maxSize) ? managerWindow.maxSize.slice(0, 2) : null,
      variation: {
        description: variation.description || title,
        isViewer: variation.isViewer !== false,
        isVisual: variation.isVisual !== false,
        isModal: variation.isModal !== false,
        isResizable: variation.isResizable !== false,
        isInsideMode: !!variation.isInsideMode,
        isUpdateOleOnResize: !!variation.isUpdateOleOnResize,
        isDisplayedInViewer: !!variation.isDisplayedInViewer,
        initDataType: variation.initDataType || 'none',
        initData: variation.initData || '',
        buttons: Array.isArray(variation.buttons) ? variation.buttons.slice() : [],
        EditorsSupport: Array.isArray(variation.EditorsSupport) && variation.EditorsSupport.length ? variation.EditorsSupport.slice() : ['cell']
      },
      config: config,
      configUrl: options.configUrl || '',
      baseUrl: options.baseUrl || '',
      localModulePath: options.localModulePath || '',
      launchUrl: launchUrl,
      source: options.source
    };
  }

  function buildRecordKey(moduleInfo) {
    return moduleInfo.guid || ('module:' + moduleInfo.module);
  }

  function isHiddenModule(config, moduleName) {
    var hiddenModules = (window.GameWizardConfig && window.GameWizardConfig.hiddenModules) || [];
    if (hiddenModules.indexOf(moduleName) !== -1) {
      return true;
    }
    return /^[_\.]/.test(moduleName);
  }

  function sortRecords(records) {
    var order = {
      'local-only': 0,
      'update-available': 1,
      installed: 1,
      available: 2,
      'external-link': 3
    };

    return records.slice().sort(function(left, right) {
      var leftOrder = Object.prototype.hasOwnProperty.call(order, left.state) ? order[left.state] : 99;
      var rightOrder = Object.prototype.hasOwnProperty.call(order, right.state) ? order[right.state] : 99;
      var leftCatalogIndex = typeof left.catalogIndex === 'number' ? left.catalogIndex : 9999;
      var rightCatalogIndex = typeof right.catalogIndex === 'number' ? right.catalogIndex : 9999;

      if (leftOrder !== rightOrder) {
        return leftOrder - rightOrder;
      }

      if (leftCatalogIndex !== rightCatalogIndex) {
        return leftCatalogIndex - rightCatalogIndex;
      }

      return left.title.localeCompare(right.title, 'ru');
    });
  }

  function mergeModules(localModules, remoteResult, bundledAvailableModules) {
    var localMap = Object.create(null);
    var remoteMap = Object.create(null);
    var bundledMap = Object.create(null);
    var keys = [];
    var remoteAvailable = !!remoteResult.available;

    for (var i = 0; i < localModules.length; i += 1) {
      var localInfo = localModules[i];
      var localKey = buildRecordKey(localInfo);
      localMap[localKey] = localInfo;
      keys.push(localKey);
    }

    for (var j = 0; j < remoteResult.modules.length; j += 1) {
      var remoteInfo = remoteResult.modules[j];
      var remoteKey = buildRecordKey(remoteInfo);
      remoteMap[remoteKey] = remoteInfo;
      keys.push(remoteKey);
    }

    for (var bundledIndex = 0; bundledIndex < bundledAvailableModules.length; bundledIndex += 1) {
      var bundledInfo = bundledAvailableModules[bundledIndex];
      var bundledKey = buildRecordKey(bundledInfo);
      bundledMap[bundledKey] = bundledInfo;
      keys.push(bundledKey);
    }

    keys = uniqueArray(keys);

    var records = [];
    for (var index = 0; index < keys.length; index += 1) {
      var key = keys[index];
      var localEntry = localMap[key] || null;
      var remoteEntry = remoteMap[key] || null;
      var bundledEntry = bundledMap[key] || null;
      var installEntry = remoteEntry || bundledEntry || null;
      var state = 'available';
      var base = null;
      var launchMode = 'direct';
      var launchUrl = '';
      var isExternalLink = false;
      var hasBundledSource = false;

      if (localEntry && remoteEntry) {
        state = compareVersions(localEntry.version, remoteEntry.version) < 0 ? 'update-available' : 'installed';
      } else if (localEntry && !remoteEntry) {
        state = remoteAvailable ? 'local-only' : 'installed';
      } else if ((remoteEntry || bundledEntry) && !localEntry) {
        state = 'available';
      }

      base = localEntry || bundledEntry || remoteEntry;
      launchMode = localEntry ? localEntry.launchMode : (bundledEntry ? bundledEntry.launchMode : (remoteEntry ? remoteEntry.launchMode : 'direct'));
      launchUrl = localEntry && localEntry.launchUrl
        ? localEntry.launchUrl
        : (bundledEntry && bundledEntry.launchUrl ? bundledEntry.launchUrl : (remoteEntry ? remoteEntry.launchUrl : ''));
      isExternalLink = launchMode === 'external-link';
      hasBundledSource = !!bundledEntry || !!(localEntry && localEntry.isBundled);

      if (isExternalLink) {
        state = 'external-link';
      }

      records.push({
        id: key,
        guid: base.guid,
        module: localEntry ? localEntry.module : (bundledEntry ? bundledEntry.module : remoteEntry.module),
        localModule: localEntry ? localEntry.module : '',
        remoteModule: installEntry ? installEntry.module : '',
        title: localEntry ? localEntry.title : (bundledEntry ? bundledEntry.title : remoteEntry.title),
        description: localEntry ? localEntry.description : (bundledEntry ? bundledEntry.description : remoteEntry.description),
        localVersion: localEntry ? localEntry.version : '',
        remoteVersion: remoteEntry ? remoteEntry.version : (bundledEntry ? bundledEntry.version : ''),
        version: localEntry ? localEntry.version : (bundledEntry ? bundledEntry.version : remoteEntry.version),
        catalogIndex: Math.min(
          localEntry && typeof localEntry.catalogIndex === 'number' ? localEntry.catalogIndex : 9999,
          remoteEntry && typeof remoteEntry.catalogIndex === 'number' ? remoteEntry.catalogIndex : 9999,
          bundledEntry && typeof bundledEntry.catalogIndex === 'number' ? bundledEntry.catalogIndex : 9999
        ),
        state: state,
        iconUrl: localEntry && localEntry.iconUrl ? localEntry.iconUrl : (bundledEntry && bundledEntry.iconUrl ? bundledEntry.iconUrl : (remoteEntry ? remoteEntry.iconUrl : '')),
        iconText: localEntry ? localEntry.iconText : (bundledEntry ? bundledEntry.iconText : remoteEntry.iconText),
        launchUrl: launchUrl,
        moduleEntry: localEntry ? localEntry.moduleEntry : (bundledEntry ? bundledEntry.moduleEntry : (remoteEntry ? remoteEntry.moduleEntry : '')),
        launchMode: launchMode,
        actionLabel: localEntry && localEntry.actionLabel ? localEntry.actionLabel : (bundledEntry && bundledEntry.actionLabel ? bundledEntry.actionLabel : (remoteEntry ? remoteEntry.actionLabel : '')),
        size: localEntry && localEntry.size ? localEntry.size : (bundledEntry && bundledEntry.size ? bundledEntry.size : (remoteEntry ? remoteEntry.size : null)),
        minSize: localEntry && localEntry.minSize ? localEntry.minSize : (bundledEntry && bundledEntry.minSize ? bundledEntry.minSize : (remoteEntry ? remoteEntry.minSize : null)),
        maxSize: localEntry && localEntry.maxSize ? localEntry.maxSize : (bundledEntry && bundledEntry.maxSize ? bundledEntry.maxSize : (remoteEntry ? remoteEntry.maxSize : null)),
        hasLocal: !!localEntry,
        hasRemote: !!installEntry,
        hasRemoteSource: !!remoteEntry,
        hasBundledSource: hasBundledSource,
        isBundled: hasBundledSource,
        canLaunch: isExternalLink ? !!launchUrl : !!localEntry,
        canInstall: !isExternalLink && !localEntry && !!installEntry,
        canUpdate: !isExternalLink && !!localEntry && !!remoteEntry && compareVersions(localEntry.version, remoteEntry.version) < 0,
        canRemove: !isExternalLink && !!localEntry,
        badge: localEntry && localEntry.badge ? localEntry.badge : (remoteEntry && remoteEntry.badge ? remoteEntry.badge : (state === 'local-only' ? 'delete' : '')),
        remoteAvailable: remoteAvailable,
        localModulePath: localEntry ? localEntry.localModulePath : '',
        remoteConfigUrl: remoteEntry ? remoteEntry.configUrl : (bundledEntry ? bundledEntry.configUrl : '')
      });
    }

    return sortRecords(records);
  }

  function ModuleService(config) {
    this.config = config || window.GameWizardConfig || {};
    this.locale = window.Asc && window.Asc.plugin && window.Asc.plugin.info ? window.Asc.plugin.info.lang || 'ru-RU' : 'ru-RU';
    this.pathInfo = resolveNativePluginDirectory();
    this.pluginRootPath = this.pathInfo.directory || '';
    this.modulesRootPath = this.pluginRootPath ? joinNativePath(this.pluginRootPath, this.config.local.modulesRoot) : '';
    this.localCatalogPath = this.modulesRootPath ? joinNativePath(this.modulesRootPath, 'local-catalog.json') : '';
    this.bundledCatalogPath = this.modulesRootPath ? joinNativePath(this.modulesRootPath, 'catalog.json') : '';
    this.lastSnapshot = {
      records: [],
      remote: {
        configured: false,
        available: false,
        error: '',
        modules: []
      }
    };
  }

  ModuleService.prototype.hasDesktopBridge = function() {
    return !!(window.AscDesktopEditor && typeof window.AscDesktopEditor.execCommand === 'function');
  };

  ModuleService.prototype.execKznUnit = function(method, args) {
    if (!this.hasDesktopBridge()) {
      throw new Error('Desktop bridge недоступен');
    }

    var payload = {
      method: method,
      args: args || {}
    };

    var raw = window.AscDesktopEditor.execCommand('kzn_unit', JSON.stringify(payload));
    if (typeof raw !== 'string') {
      return raw;
    }

    try {
      return JSON.parse(raw);
    } catch (_) {
      return raw;
    }
  };

  ModuleService.prototype.exists = function(filePath) {
    if (!filePath) {
      return false;
    }

    if (this.hasDesktopBridge()) {
      var result = this.execKznUnit('exists', {
        file: encodeUtf8ToBase64(filePath)
      });
      return !!(result && result.data);
    }

    return false;
  };

  ModuleService.prototype.readLocalText = function(filePath, isDirectory) {
    if (!filePath) {
      return Promise.reject(new Error('Путь к локальному файлу не определен'));
    }

    return xhrRequest(buildFileUrl(filePath, !!isDirectory), 'text');
  };

  ModuleService.prototype.readLocalJson = function(filePath, label) {
    return this.readLocalText(filePath).then(function(text) {
      return parseJson(text, label);
    });
  };

  ModuleService.prototype.writeBase64File = function(filePath, base64Data) {
    if (!this.hasDesktopBridge()) {
      throw new Error('Desktop bridge недоступен для записи');
    }

    var result = this.execKznUnit('save_file', {
      file: encodeUtf8ToBase64(filePath),
      data: String(base64Data || '')
    });

    if (result && result.error) {
      throw new Error('save_file failed: ' + result.error);
    }
  };

  ModuleService.prototype.writeTextFile = function(filePath, text) {
    this.writeBase64File(filePath, encodeUtf8ToBase64(text));
  };

  ModuleService.prototype.writeBinaryFile = function(filePath, buffer) {
    this.writeBase64File(filePath, arrayBufferToBase64(buffer));
  };

  ModuleService.prototype.buildDirectoryVariants = function(directoryPath) {
    if (!directoryPath) {
      return [];
    }

    var trimmed = String(directoryPath).replace(/[\\/]+$/, '');
    if (!trimmed) {
      return [String(directoryPath)];
    }

    return uniqueArray([
      trimmed,
      trimmed + '\\',
      trimmed + '/'
    ]);
  };

  ModuleService.prototype.removeFile = function(filePath) {
    if (!filePath || !this.exists(filePath)) {
      return false;
    }

    if (window.AscDesktopEditor && typeof window.AscDesktopEditor.RemoveFile === 'function') {
      window.AscDesktopEditor.RemoveFile(filePath);
      return true;
    }

    throw new Error('AscDesktopEditor.RemoveFile недоступен');
  };

  ModuleService.prototype.removeDirectory = function(directoryPath) {
    var variants = this.buildDirectoryVariants(directoryPath);
    var existingVariants = [];
    var self = this;

    for (var index = 0; index < variants.length; index += 1) {
      if (this.exists(variants[index])) {
        existingVariants.push(variants[index]);
      }
    }

    if (!existingVariants.length) {
      return false;
    }

    if (window.AscDesktopEditor && typeof window.AscDesktopEditor.RemoveFile === 'function') {
      for (var directIndex = 0; directIndex < existingVariants.length; directIndex += 1) {
        try {
          window.AscDesktopEditor.RemoveFile(existingVariants[directIndex]);
        } catch (_) {
          // Best-effort fallback continues below.
        }
      }

      var removedAfterDirect = existingVariants.every(function(variant) {
        return !self.exists(variant);
      });
      if (removedAfterDirect) {
        return true;
      }
    }

    if (this.hasDesktopBridge()) {
      var methods = ['remove_dir', 'delete_dir', 'rmdir', 'remove_file', 'delete_file', 'delete', 'unlink'];
      var argNames = ['dir', 'path', 'file'];

      for (var methodIndex = 0; methodIndex < methods.length; methodIndex += 1) {
        for (var variantIndex = 0; variantIndex < existingVariants.length; variantIndex += 1) {
          for (var argIndex = 0; argIndex < argNames.length; argIndex += 1) {
            try {
              var args = {};
              args[argNames[argIndex]] = encodeUtf8ToBase64(existingVariants[variantIndex]);
              this.execKznUnit(methods[methodIndex], args);
            } catch (_) {
              // Continue probing best-effort directory removal methods.
            }

            var removed = existingVariants.every(function(variant) {
              return !self.exists(variant);
            });
            if (removed) {
              return true;
            }
          }
        }
      }
    }

    throw new Error('Не удалось удалить директорию через доступные bridge методы');
  };

  ModuleService.prototype.tryReadCatalogFile = function(filePath) {
    var self = this;

    if (!filePath || !this.exists(filePath)) {
      return Promise.resolve([]);
    }

    return this.readLocalJson(filePath, filePath).then(function(payload) {
      return parseCatalog(payload);
    }).catch(function() {
      return self.readLocalText(filePath).then(function(text) {
        return parseCatalog(parseJson(text, filePath));
      }).catch(function() {
        return [];
      });
    });
  };

  ModuleService.prototype.tryListModuleDirectories = function() {
    var self = this;

    if (!this.modulesRootPath) {
      return Promise.resolve([]);
    }

    return this.readLocalText(this.modulesRootPath, true).then(function(html) {
      return parseDirectoryListing(html).filter(function(entry) {
        var normalized = entry.replace(/\/+$/, '');
        return !!normalized && normalized.indexOf('.') !== 0;
      }).map(function(entry) {
        return entry.replace(/\/+$/, '');
      });
    }).catch(function() {
      return [];
    }).then(function(entries) {
      return entries.filter(function(entry) {
        return !isHiddenModule(self.config, entry);
      });
    });
  };

  ModuleService.prototype.scanLocalModuleTree = function(moduleName, relativePath) {
    var self = this;
    var modulePath = joinNativePath(this.modulesRootPath, moduleName, relativePath || '');

    return this.readLocalText(modulePath, true).then(function(html) {
      var entries = parseDirectoryListing(html);
      var tasks = [];
      var files = [];
      var directories = [];

      for (var i = 0; i < entries.length; i += 1) {
        var entry = entries[i];
        if (!entry || entry === '../') {
          continue;
        }

        var normalized = entry.replace(/^\/+/, '');
        var childRelative = normalizeSlashes(joinUrlPath(relativePath || '', normalized.replace(/\/+$/, '')));

        if (/\/$/.test(entry)) {
          directories.push(childRelative);
          tasks.push(self.scanLocalModuleTree(moduleName, childRelative).then(function(childTree) {
            for (var index = 0; index < childTree.files.length; index += 1) {
              files.push(childTree.files[index]);
            }
            for (var directoryIndex = 0; directoryIndex < childTree.directories.length; directoryIndex += 1) {
              directories.push(childTree.directories[directoryIndex]);
            }
          }));
        } else {
          files.push(childRelative);
        }
      }

      return Promise.all(tasks).then(function() {
        return {
          files: uniqueArray(files),
          directories: sortByDepthDesc(uniqueArray(directories))
        };
      });
    }).catch(function() {
      return {
        files: [],
        directories: []
      };
    });
  };

  ModuleService.prototype.listModuleFilesRecursively = function(moduleName, relativePath) {
    return this.scanLocalModuleTree(moduleName, relativePath).then(function(tree) {
      return tree.files;
    });
  };

  ModuleService.prototype.readLocalModules = function(candidateNames) {
    var self = this;
    var tasks = [];

    for (var i = 0; i < candidateNames.length; i += 1) {
      (function(moduleName, catalogIndex) {
        if (!moduleName || isHiddenModule(self.config, moduleName)) {
          return;
        }

        var moduleRootPath = joinNativePath(self.modulesRootPath, moduleName);
        var configPath = joinNativePath(moduleRootPath, 'config.json');

        if (!self.exists(configPath)) {
          return;
        }

        tasks.push(self.readLocalJson(configPath, configPath).then(function(config) {
          var baseUrl = new URL('modules/' + moduleName + '/', window.location.href).toString();
          return createStandaloneModuleRecord({
            module: moduleName,
            catalogIndex: catalogIndex,
            config: config,
            configUrl: new URL('config.json', baseUrl).toString(),
            baseUrl: baseUrl,
            relativeBasePath: joinUrlPath('modules', moduleName),
            launchUrl: resolveAssetUrl(baseUrl, pickVariation(config).url || 'index.html'),
            localModulePath: moduleRootPath,
            locale: self.locale,
            source: 'local'
          });
        }).catch(function(error) {
          console.warn('[GameWizard] Failed to read local module config', moduleName, error);
          return null;
        }));
      })(candidateNames[i], i);
    }

    return Promise.all(tasks).then(function(items) {
      return items.filter(Boolean);
    });
  };

  ModuleService.prototype.readRemoteModules = function() {
    var remoteConfig = this.config.remote || {};
    var repo = parseGitHubRepository(remoteConfig.repositoryUrl, remoteConfig.branch);
    var self = this;

    if (!repo) {
      return Promise.resolve({
        configured: false,
        available: false,
        error: '',
        modules: []
      });
    }

    var rawBaseUrl = 'https://raw.githubusercontent.com/' + repo.owner + '/' + repo.repo + '/' + repo.branch;
    var modulesRoot = trimSlashes(remoteConfig.modulesRoot || 'modules');
    var catalogUrl = rawBaseUrl + '/' + joinUrlPath(modulesRoot, 'catalog.json');

    return fetchText(catalogUrl).then(function(text) {
      var moduleNames = parseCatalog(parseJson(text, catalogUrl));
      var tasks = [];

      for (var i = 0; i < moduleNames.length; i += 1) {
        (function(moduleName, catalogIndex) {
          if (isHiddenModule(self.config, moduleName)) {
            return;
          }

          var baseUrl = rawBaseUrl + '/' + joinUrlPath(modulesRoot, moduleName) + '/';
          var configUrl = rawBaseUrl + '/' + joinUrlPath(modulesRoot, moduleName, 'config.json');

          tasks.push(fetchText(configUrl).then(function(configText) {
            var config = parseJson(configText, configUrl);
            return createStandaloneModuleRecord({
              module: moduleName,
              catalogIndex: catalogIndex,
              config: config,
              configUrl: configUrl,
              baseUrl: baseUrl,
              relativeBasePath: joinUrlPath(modulesRoot, moduleName),
              locale: self.locale,
              source: 'remote'
            });
          }));
        })(moduleNames[i], i);
      }

      return Promise.all(tasks).then(function(modules) {
        return {
          configured: true,
          available: true,
          error: '',
          repo: repo,
          rawBaseUrl: rawBaseUrl,
          modulesRoot: modulesRoot,
          modules: modules
        };
      });
    }).catch(function(error) {
      return {
        configured: true,
        available: false,
        error: error.message,
        repo: repo,
        rawBaseUrl: rawBaseUrl,
        modulesRoot: modulesRoot,
        modules: []
      };
    });
  };

  ModuleService.prototype.writeModuleManifest = function(moduleName, moduleRootPath, files, directories) {
    if (!moduleRootPath || !this.hasDesktopBridge()) {
      return;
    }

    var normalizedFiles = sortByDepthDesc(uniqueArray((files || []).map(function(filePath) {
      return normalizeSlashes(filePath);
    }).filter(Boolean)));
    var normalizedDirectories = sortByDepthDesc(uniqueArray((directories || []).map(function(directoryPath) {
      return normalizeSlashes(directoryPath);
    }).filter(Boolean)));
    var payload = {
      module: moduleName,
      files: normalizedFiles
    };

    if (normalizedDirectories.length) {
      payload.directories = normalizedDirectories;
    }

    this.writeTextFile(joinNativePath(moduleRootPath, 'module-manifest.json'), JSON.stringify(payload, null, 2));
  };

  ModuleService.prototype.ensureModuleManifest = function(record) {
    var self = this;

    if (!record || !record.localModulePath || !this.hasDesktopBridge()) {
      return Promise.resolve();
    }

    var manifestPath = joinNativePath(record.localModulePath, 'module-manifest.json');
    if (this.exists(manifestPath)) {
      return Promise.resolve();
    }

    return this.scanLocalModuleTree(record.localModule, '').then(function(tree) {
      if (!tree.files.length) {
        return;
      }

      self.writeModuleManifest(record.localModule, record.localModulePath, tree.files, tree.directories);
    }).catch(function(error) {
      console.warn('[GameWizard] Failed to backfill module manifest', record.localModule, error);
    });
  };

  ModuleService.prototype.persistLocalCatalog = function(moduleNames) {
    if (!this.localCatalogPath || !this.modulesRootPath || !this.hasDesktopBridge()) {
      return Promise.resolve();
    }

    var payload = {
      modules: uniqueArray(moduleNames).map(function(moduleName) {
        return { module: moduleName };
      })
    };

    try {
      this.writeTextFile(this.localCatalogPath, JSON.stringify(payload, null, 2));
      return Promise.resolve();
    } catch (error) {
      console.warn('[GameWizard] Failed to persist local catalog', error);
      return Promise.resolve();
    }
  };

  ModuleService.prototype.getKnownLocalModuleNames = function() {
    var names = [];
    var snapshot = this.lastSnapshot || {};

    if (Array.isArray(snapshot.localModules)) {
      for (var localIndex = 0; localIndex < snapshot.localModules.length; localIndex += 1) {
        if (snapshot.localModules[localIndex] && snapshot.localModules[localIndex].module) {
          names.push(snapshot.localModules[localIndex].module);
        }
      }
    }

    if (Array.isArray(snapshot.records)) {
      for (var recordIndex = 0; recordIndex < snapshot.records.length; recordIndex += 1) {
        if (snapshot.records[recordIndex] && snapshot.records[recordIndex].hasLocal && snapshot.records[recordIndex].localModule) {
          names.push(snapshot.records[recordIndex].localModule);
        }
      }
    }

    return uniqueArray(names);
  };

  ModuleService.prototype.enableLocalModule = function(moduleName) {
    return this.persistLocalCatalog(this.getKnownLocalModuleNames().concat([moduleName]));
  };

  ModuleService.prototype.disableLocalModule = function(moduleName) {
    return this.persistLocalCatalog(this.getKnownLocalModuleNames().filter(function(localModuleName) {
      return localModuleName && localModuleName !== moduleName;
    }));
  };

  ModuleService.prototype.refresh = function(extraCandidateNames) {
    var self = this;
    var forcedModules = Array.isArray(extraCandidateNames) ? uniqueArray(extraCandidateNames.filter(Boolean)) : [];
    var hasLocalCatalog = !!(this.localCatalogPath && this.exists(this.localCatalogPath));

    return Promise.all([
      this.tryReadCatalogFile(this.bundledCatalogPath),
      this.tryReadCatalogFile(this.localCatalogPath),
      this.tryListModuleDirectories()
    ]).then(function(results) {
      var bundledModules = results[0];
      var cachedModules = results[1];
      var scannedModules = results[2];
      var bundledLookup = createLookup(bundledModules);
      var nonBundledScannedModules = scannedModules.filter(function(moduleName) {
        return !bundledLookup[moduleName];
      });
      var enabledModules = hasLocalCatalog
        ? uniqueArray(cachedModules.concat(nonBundledScannedModules))
        : uniqueArray(bundledModules.concat(nonBundledScannedModules));
      var disabledBundledModules = bundledModules.filter(function(moduleName) {
        return enabledModules.indexOf(moduleName) === -1;
      });
      var candidates = uniqueArray(enabledModules.concat(disabledBundledModules, forcedModules));

      return Promise.all([
        self.readLocalModules(candidates),
        self.readRemoteModules()
      ]).then(function(data) {
        return {
          discoveredModules: data[0],
          remoteResult: data[1],
          bundledModules: bundledModules,
          enabledModules: enabledModules
        };
      });
    }).then(function(state) {
      var bundledLookup = createLookup(state.bundledModules);
      var enabledLookup = createLookup(state.enabledModules);
      var localModules = [];
      var bundledAvailableModules = [];

      for (var discoveredIndex = 0; discoveredIndex < state.discoveredModules.length; discoveredIndex += 1) {
        var discoveredModule = state.discoveredModules[discoveredIndex];
        if (!discoveredModule) {
          continue;
        }

        discoveredModule.isBundled = !!bundledLookup[discoveredModule.module];

        if (enabledLookup[discoveredModule.module]) {
          localModules.push(discoveredModule);
        } else if (discoveredModule.isBundled) {
          bundledAvailableModules.push(discoveredModule);
        }
      }

      return Promise.all(localModules.map(function(item) {
        return self.ensureModuleManifest(item);
      })).then(function() {
        var records = mergeModules(localModules, state.remoteResult, bundledAvailableModules);
        var mutationSupported = self.hasDesktopBridge() && !!self.modulesRootPath;

        for (var recordIndex = 0; recordIndex < records.length; recordIndex += 1) {
          records[recordIndex].canInstall = records[recordIndex].canInstall && mutationSupported;
          records[recordIndex].canUpdate = records[recordIndex].canUpdate && mutationSupported;
          records[recordIndex].canRemove = records[recordIndex].canRemove && mutationSupported;
        }

        self.lastSnapshot = {
          records: records,
          localModules: localModules,
          bundledModules: state.bundledModules,
          enabledModules: state.enabledModules,
          remote: state.remoteResult
        };

        return self.persistLocalCatalog(localModules.map(function(item) {
          return item.module;
        })).then(function() {
          return self.lastSnapshot;
        });
      });
    });
  };

  ModuleService.prototype.findRecord = function(recordId) {
    var records = this.lastSnapshot && Array.isArray(this.lastSnapshot.records) ? this.lastSnapshot.records : [];
    for (var i = 0; i < records.length; i += 1) {
      if (records[i].id === recordId) {
        return records[i];
      }
    }
    return null;
  };

  ModuleService.prototype.resolveGitHubContext = function() {
    var remoteConfig = this.config.remote || {};
    var repo = parseGitHubRepository(remoteConfig.repositoryUrl, remoteConfig.branch);
    if (!repo) {
      throw new Error('GitHub repository is not configured');
    }

    return {
      repo: repo,
      modulesRoot: trimSlashes(remoteConfig.modulesRoot || 'modules'),
      rawBaseUrl: 'https://raw.githubusercontent.com/' + repo.owner + '/' + repo.repo + '/' + repo.branch,
      apiBaseUrl: 'https://api.github.com/repos/' + repo.owner + '/' + repo.repo + '/contents/'
    };
  };

  ModuleService.prototype.fetchGitHubDirectory = function(relativePath) {
    var context = this.resolveGitHubContext();
    var requestPath = trimSlashes(relativePath);
    var url = context.apiBaseUrl + requestPath + '?ref=' + encodeURIComponent(context.repo.branch);
    return fetchText(url, {
      Accept: 'application/vnd.github+json'
    }).then(function(text) {
      return parseJson(text, url);
    });
  };

  ModuleService.prototype.collectRemoteFiles = function(moduleName) {
    var self = this;
    var context = this.resolveGitHubContext();
    var moduleRootRelative = joinUrlPath(context.modulesRoot, moduleName);
    var manifestUrl = context.rawBaseUrl + '/' + joinUrlPath(moduleRootRelative, 'module-manifest.json');

    function buildFilesFromManifest(payload) {
      var files = payload && Array.isArray(payload.files) ? payload.files.map(function(relativePath) {
        return normalizeSlashes(relativePath);
      }).filter(Boolean) : [];

      if (!files.length) {
        throw new Error('Remote module manifest is empty for ' + moduleName);
      }

      return files.map(function(relativePath) {
        return {
          path: joinUrlPath(moduleRootRelative, relativePath),
          relativePath: relativePath,
          downloadUrl: context.rawBaseUrl + '/' + joinUrlPath(moduleRootRelative, relativePath)
        };
      });
    }

    function walk(path) {
      return self.fetchGitHubDirectory(path).then(function(payload) {
        if (!Array.isArray(payload)) {
          throw new Error('Unexpected GitHub contents payload for ' + path);
        }

        var tasks = [];
        var files = [];

        for (var i = 0; i < payload.length; i += 1) {
          var item = payload[i];
          if (!item || !item.type || !item.path) {
            continue;
          }

          if (item.type === 'dir') {
            tasks.push(walk(item.path).then(function(childFiles) {
              for (var index = 0; index < childFiles.length; index += 1) {
                files.push(childFiles[index]);
              }
            }));
          } else if (item.type === 'file' && item.download_url) {
            files.push({
              path: item.path,
              relativePath: normalizeSlashes(item.path).slice(moduleRootRelative.length + 1),
              downloadUrl: item.download_url
            });
          }
        }

        return Promise.all(tasks).then(function() {
          return files;
        });
      });
    }

    return fetchText(manifestUrl).then(function(text) {
      return buildFilesFromManifest(parseJson(text, manifestUrl));
    }).catch(function() {
      return walk(moduleRootRelative).catch(function(error) {
        if (/HTTP 403 .*api\.github\.com/i.test(String(error && error.message || ''))) {
          throw new Error('GitHub API rate limit exceeded and remote module manifest is unavailable: ' + manifestUrl);
        }
        throw error;
      });
    });
  };

  ModuleService.prototype.installRemoteModuleFiles = function(moduleName) {
    var self = this;

    if (!moduleName) {
      return Promise.reject(new Error('Не указано имя модуля для установки'));
    }

    return this.collectRemoteFiles(moduleName).then(function(files) {
      if (!files.length) {
        throw new Error('Удаленный модуль не содержит файлов');
      }

      var moduleRootPath = joinNativePath(self.modulesRootPath, moduleName);
      var downloadedRelativePaths = [];
      var chain = Promise.resolve();

      files.forEach(function(file) {
        chain = chain.then(function() {
          return fetchArrayBuffer(file.downloadUrl).then(function(buffer) {
            var targetPath = joinNativePath(moduleRootPath, file.relativePath);
            self.writeBinaryFile(targetPath, buffer);
            downloadedRelativePaths.push(normalizeSlashes(file.relativePath));
          });
        });
      });

      return chain.then(function() {
        self.writeModuleManifest(
          moduleName,
          moduleRootPath,
          downloadedRelativePaths,
          deriveDirectoriesFromFiles(downloadedRelativePaths)
        );
      });
    });
  };

  ModuleService.prototype.install = function(recordId) {
    var self = this;
    var record = this.findRecord(recordId);

    if (!record || !record.hasRemote) {
      return Promise.reject(new Error('Модуль для установки не найден в удаленном каталоге'));
    }

    if (record.hasBundledSource && !record.hasLocal) {
      return this.enableLocalModule(record.remoteModule || record.module).then(function() {
        return self.refresh();
      });
    }

    return this.installRemoteModuleFiles(record.remoteModule).then(function() {
      return self.enableLocalModule(record.remoteModule);
    }).then(function() {
      return self.refresh([record.remoteModule]);
    });
  };

  ModuleService.prototype.buildLocalRemovalPlan = function(record) {
    var manifestPath = joinNativePath(record.localModulePath, 'module-manifest.json');

    if (this.exists(manifestPath)) {
      return this.readLocalJson(manifestPath, manifestPath).then(function(payload) {
        var files = payload && Array.isArray(payload.files) ? payload.files.map(function(file) {
          return normalizeSlashes(file);
        }).filter(Boolean) : [];
        var directories = payload && Array.isArray(payload.directories) ? payload.directories.map(function(directory) {
          return normalizeSlashes(directory);
        }).filter(Boolean) : deriveDirectoriesFromFiles(files);

        return {
          files: uniqueArray(files),
          directories: sortByDepthDesc(uniqueArray(directories))
        };
      }).catch(function() {
        return {
          files: [],
          directories: []
        };
      });
    }

    return this.scanLocalModuleTree(record.localModule, '').then(function(tree) {
      return {
        files: tree.files.filter(Boolean),
        directories: sortByDepthDesc(uniqueArray(tree.directories.concat(deriveDirectoriesFromFiles(tree.files))))
      };
    }).catch(function() {
      return {
        files: [],
        directories: []
      };
    });
  };

  ModuleService.prototype.remove = function(recordId) {
    var self = this;
    var record = this.findRecord(recordId);

    if (!record || !record.hasLocal) {
      return Promise.reject(new Error('Локальный модуль для удаления не найден'));
    }

    if (record.isBundled) {
      return this.disableLocalModule(record.localModule).then(function() {
        return self.refresh();
      });
    }

    return this.buildLocalRemovalPlan(record).then(function(plan) {
      var uniqueFiles = sortByDepthDesc(uniqueArray((plan.files || []).concat(['config.json', 'module-manifest.json'])));
      var uniqueDirectories = sortByDepthDesc(uniqueArray((plan.directories || []).concat(deriveDirectoriesFromFiles(uniqueFiles))));

      for (var i = 0; i < uniqueFiles.length; i += 1) {
        var filePath = joinNativePath(record.localModulePath, uniqueFiles[i]);
        try {
          self.removeFile(filePath);
        } catch (error) {
          console.warn('[GameWizard] Failed to remove file', filePath, error);
        }
      }

      for (var directoryIndex = 0; directoryIndex < uniqueDirectories.length; directoryIndex += 1) {
        var directoryPath = joinNativePath(record.localModulePath, uniqueDirectories[directoryIndex]);
        try {
          self.removeDirectory(directoryPath);
        } catch (directoryError) {
          console.warn('[GameWizard] Failed to remove directory', directoryPath, directoryError);
        }
      }

      try {
        self.removeDirectory(record.localModulePath);
      } catch (rootError) {
        console.warn('[GameWizard] Failed to remove module root', record.localModulePath, rootError);
      }
    }).then(function() {
      return self.disableLocalModule(record.localModule);
    }).then(function() {
      return self.refresh();
    });
  };

  ModuleService.prototype.update = function(recordId) {
    var self = this;
    var record = this.findRecord(recordId);

    if (record && record.isBundled && record.hasRemoteSource) {
      return this.installRemoteModuleFiles(record.remoteModule).then(function() {
        return self.enableLocalModule(record.remoteModule);
      }).then(function() {
        return self.refresh([record.remoteModule]);
      });
    }

    return this.remove(recordId).then(function() {
      return self.install(recordId);
    });
  };

  window.GameWizardModuleService = ModuleService;
})(window);



