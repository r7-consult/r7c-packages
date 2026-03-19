(function(window) {
  'use strict';

  function clearNode(node) {
    while (node.firstChild) {
      node.removeChild(node.firstChild);
    }
  }

  function createElement(tagName, className, text) {
    var node = document.createElement(tagName);
    if (className) {
      node.className = className;
    }
    if (text != null) {
      node.textContent = text;
    }
    return node;
  }

  function stopEvent(event) {
    event.preventDefault();
    event.stopPropagation();
  }

  function createActionButton(label, className, onClick) {
    var button = createElement('button', className, label);
    button.type = 'button';
    button.addEventListener('click', function(event) {
      stopEvent(event);
      onClick();
    });
    return button;
  }

  function createIcon(record) {
    var icon = createElement('span', 'module-card-icon');

    if (record.iconUrl) {
      var image = document.createElement('img');
      image.className = 'module-card-icon-image';
      image.src = record.iconUrl;
      image.alt = '';
      image.addEventListener('error', function() {
        if (image.parentNode) {
          image.parentNode.removeChild(image);
        }
        icon.textContent = record.iconText || 'GM';
      }, { once: true });
      icon.appendChild(image);
      return icon;
    }

    icon.textContent = record.iconText || 'GM';
    return icon;
  }

  function createMetaLine(record) {
    var meta = createElement('div', 'module-card-meta');
    var stateLabel = record.state === 'update-available'
      ? 'Доступно обновление'
      : record.state === 'available'
        ? 'Доступно к установке'
        : record.state === 'local-only'
          ? 'Локальный модуль'
          : 'Установлено';

    meta.appendChild(createElement('span', 'module-card-version', 'Версия: ' + (record.version || '0.0.0')));
    meta.appendChild(createElement('span', 'module-card-state', stateLabel));
    return meta;
  }

  function bindCardActivation(card, onActivate) {
    card.addEventListener('click', onActivate);
    card.addEventListener('keydown', function(event) {
      if (event.key === 'Enter' || event.key === ' ') {
        stopEvent(event);
        onActivate();
      }
    });
  }

  function createCard(record, handlers) {
    var card = createElement('article', 'module-card module-card--' + record.state);
    var actions = createElement('div', 'module-card-actions');
    var main = createElement('div', 'module-card-main');
    var title = createElement('div', 'module-card-title', record.title);
    var description = createElement('div', 'module-card-desc', record.description || 'Модуль плагина');

    card.setAttribute('data-record-id', record.id);
    card.appendChild(createIcon(record));

    if (record.badge) {
      card.appendChild(createElement('span', 'module-card-badge', record.badge));
    }

    main.appendChild(title);
    main.appendChild(description);
    main.appendChild(createMetaLine(record));
    card.appendChild(main);

    if (record.canInstall) {
      actions.appendChild(createActionButton('Установить', 'module-action module-action--install', function() {
        handlers.onInstall(record.id);
      }));
    }

    if (record.canUpdate) {
      actions.appendChild(createActionButton('Обновить', 'module-action module-action--secondary', function() {
        handlers.onUpdate(record.id);
      }));
    }

    if (record.canRemove) {
      actions.appendChild(createActionButton('Удалить', 'module-action module-action--danger', function() {
        handlers.onRemove(record.id);
      }));
    }

    if (!actions.childNodes.length) {
      actions.appendChild(createElement('span', 'module-card-action-hint', record.canLaunch ? 'Открыть' : '')); 
    }

    card.appendChild(actions);

    if (record.canLaunch) {
      card.classList.add('is-clickable');
      card.tabIndex = 0;
      card.setAttribute('role', 'button');
      bindCardActivation(card, function() {
        handlers.onOpen(record.id);
      });
    }

    return card;
  }

  function createEmptyState(text) {
    var empty = createElement('div', 'modules-empty');
    empty.appendChild(createElement('div', 'modules-empty-title', 'Игры не найдены'));
    empty.appendChild(createElement('div', 'modules-empty-text', text || 'В менеджере пока нет доступных модулей.'));
    return empty;
  }

  window.GameWizardUI = {
    renderModules: function(records, handlers) {
      var list = document.getElementById('games-list');
      if (!list) {
        return;
      }

      clearNode(list);

      if (!records || !records.length) {
        list.appendChild(createEmptyState('Проверьте локальную папку modules или настройку GitHub-каталога.'));
        return;
      }

      for (var i = 0; i < records.length; i += 1) {
        list.appendChild(createCard(records[i], handlers));
      }
    },

    setStatus: function(message) {
      var status = document.getElementById('gw-status');
      if (status) {
        status.textContent = message;
      }
    },

    setMeta: function(message) {
      var meta = document.getElementById('gw-main-meta');
      if (meta) {
        meta.textContent = message;
      }
    },

    setBranding: function(branding) {
      var title = branding && branding.title ? branding.title : 'Game Wizard';
      var subtitle = branding && branding.subtitle ? branding.subtitle : '';
      var titleNode = document.getElementById('gw-title');
      var subtitleNode = document.getElementById('gw-subtitle');

      if (titleNode) {
        titleNode.textContent = title;
      }
      if (subtitleNode) {
        subtitleNode.textContent = subtitle;
      }

      document.title = title;
    }
  };
})(window);
