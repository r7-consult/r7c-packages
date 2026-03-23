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

  function createActionButton(label, className, actionType, recordId, onClick) {
    var button = createElement('button', className, label);
    button.type = 'button';
    button.setAttribute('data-action-type', actionType || '');
    button.setAttribute('data-record-id', recordId || '');
    button.setAttribute('data-default-label', label || '');
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
        icon.textContent = record.iconText || 'DP';
      }, { once: true });
      icon.appendChild(image);
      return icon;
    }

    icon.textContent = record.iconText || 'DP';
    return icon;
  }

  function buildStateLabel(state) {
    if (state === 'update-available') {
      return 'Доступно обновление';
    }

    if (state === 'available') {
      return 'Доступно к установке';
    }

    if (state === 'local-only') {
      return 'Только локально';
    }

    if (state === 'external-link') {
      return 'Внешняя ссылка';
    }

    return 'Установлено';
  }

  function createMetaLine(record) {
    var meta = createElement('div', 'module-card-meta');
    var version = record.version || record.remoteVersion || record.localVersion || '';

    if (version) {
      meta.appendChild(createElement('span', 'module-card-version', 'Версия: ' + version));
    }

    meta.appendChild(createElement('span', 'module-card-state', buildStateLabel(record.state)));
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

  function eachRecordCard(recordId, callback) {
    var cards = document.querySelectorAll('.module-card[data-record-id]');
    var index;

    for (index = 0; index < cards.length; index += 1) {
      if (cards[index].getAttribute('data-record-id') === recordId) {
        callback(cards[index]);
      }
    }
  }

  function clearPendingState(rootNode) {
    var scope = rootNode || document;
    var cards = scope.querySelectorAll('.module-card.is-pending');
    var buttons = scope.querySelectorAll('.module-action[data-default-label]');
    var index;

    for (index = 0; index < cards.length; index += 1) {
      cards[index].classList.remove('is-pending');
      cards[index].removeAttribute('aria-busy');
    }

    for (index = 0; index < buttons.length; index += 1) {
      buttons[index].classList.remove('is-pending');
      buttons[index].disabled = false;
      buttons[index].textContent = buttons[index].getAttribute('data-default-label') || buttons[index].textContent;
    }
  }

  function applyPendingState(recordId, actionType, busyLabel) {
    eachRecordCard(recordId, function(card) {
      var buttons = card.querySelectorAll('.module-action');
      var targetButton = null;
      var index;

      card.classList.add('is-pending');
      card.setAttribute('aria-busy', 'true');

      for (index = 0; index < buttons.length; index += 1) {
        buttons[index].disabled = true;
        if (buttons[index].getAttribute('data-action-type') === actionType) {
          targetButton = buttons[index];
        }
      }

      if (targetButton) {
        targetButton.classList.add('is-pending');
        targetButton.textContent = busyLabel || targetButton.getAttribute('data-default-label') || targetButton.textContent;
      }
    });
  }

  function createCard(record, handlers) {
    var isExternalLink = record.launchMode === 'external-link' && !!record.launchUrl;
    var card = createElement(isExternalLink ? 'a' : 'article', 'module-card module-card--' + record.state);
    var actions = createElement('div', 'module-card-actions');
    var main = createElement('div', 'module-card-main');
    var title = createElement('div', 'module-card-title', record.title);
    var description = createElement('div', 'module-card-desc', record.description || 'Модуль обработки данных');
    var launchLabel = record.actionLabel || (record.launchMode === 'external-link' ? 'Перейти' : 'Открыть');

    card.setAttribute('data-record-id', record.id);

    if (isExternalLink) {
      card.href = record.launchUrl;
      card.target = '_blank';
      card.rel = 'noopener noreferrer';
    }

    card.appendChild(createIcon(record));

    if (record.badge) {
      card.appendChild(createElement('span', 'module-card-badge', record.badge));
    }

    main.appendChild(title);
    main.appendChild(description);
    main.appendChild(createMetaLine(record));
    card.appendChild(main);

    if (isExternalLink && record.canLaunch) {
      actions.appendChild(createElement('span', 'module-card-action-hint', launchLabel));
    }

    if (record.canInstall) {
      actions.appendChild(createActionButton('Установить', 'module-action module-action--install', 'install', record.id, function() {
        handlers.onInstall(record.id);
      }));
    }

    if (record.canUpdate) {
      actions.appendChild(createActionButton('Обновить', 'module-action module-action--secondary', 'update', record.id, function() {
        handlers.onUpdate(record.id);
      }));
    }

    if (record.canRemove) {
      actions.appendChild(createActionButton('Удалить', 'module-action module-action--danger', 'remove', record.id, function() {
        handlers.onRemove(record.id);
      }));
    }

    if (!actions.childNodes.length) {
      actions.appendChild(createElement('span', 'module-card-action-hint', record.canLaunch ? launchLabel : ''));
    }

    card.appendChild(actions);

    if (record.canLaunch && !isExternalLink) {
      card.classList.add('is-clickable');
      card.tabIndex = 0;
      card.setAttribute('role', 'button');
      bindCardActivation(card, function() {
        handlers.onOpen(record.id);
      });
    } else if (isExternalLink) {
      card.classList.add('is-clickable');
    }

    return card;
  }

  function createEmptyState(text) {
    var empty = createElement('div', 'modules-empty');
    empty.appendChild(createElement('div', 'modules-empty-title', 'Решения не найдены'));
    empty.appendChild(createElement('div', 'modules-empty-text', text || 'Проверьте локальную папку modules или удаленный каталог GitHub.'));
    return empty;
  }

  window.GameWizardUI = {
    renderModules: function(records, handlers) {
      var list = document.getElementById('games-list');
      var index;

      if (!list) {
        return;
      }

      clearNode(list);

      if (!records || !records.length) {
        list.appendChild(createEmptyState('Проверьте локальную папку modules или настройку GitHub-каталога.'));
        return;
      }

      for (index = 0; index < records.length; index += 1) {
        list.appendChild(createCard(records[index], handlers));
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

    setPendingAction: function(recordId, actionType, busyLabel) {
      var list = document.getElementById('games-list');
      if (!list) {
        return;
      }

      clearPendingState(list);

      if (!recordId || !actionType) {
        return;
      }

      applyPendingState(recordId, actionType, busyLabel);
    },

    setBranding: function(branding) {
      var title = branding && branding.title ? branding.title : 'Обработка данных';
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
