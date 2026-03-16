(function(window) {
  'use strict';

  function clearNode(node) {
    while (node.firstChild) {
      node.removeChild(node.firstChild);
    }
  }

  function createButton(game, onOpen) {
    const isDisabled = !!game.disabled;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'game-card' + (isDisabled ? ' is-disabled' : '');
    btn.setAttribute('data-game-id', game.id);
    btn.setAttribute(
      'aria-label',
      isDisabled ? ('Игра временно недоступна: ' + game.title) : ('Открыть игру ' + game.title)
    );

    if (isDisabled) {
      btn.disabled = true;
      btn.setAttribute('aria-disabled', 'true');
    }

    const icon = document.createElement('span');
    icon.className = 'game-card-icon';
    icon.textContent = game.icon || 'GM';

    const main = document.createElement('span');
    main.className = 'game-card-main';

    const title = document.createElement('span');
    title.className = 'game-card-title';
    title.textContent = game.title;

    const desc = document.createElement('span');
    desc.className = 'game-card-desc';
    desc.textContent = game.description || 'Запуск игры';

    const action = document.createElement('span');
    action.className = 'game-card-action';
    action.textContent = isDisabled ? 'Недоступно' : 'Открыть';

    main.appendChild(title);
    main.appendChild(desc);
    btn.appendChild(icon);
    btn.appendChild(main);
    btn.appendChild(action);

    if (!isDisabled) {
      btn.addEventListener('click', function() {
        onOpen(game.id);
      });
    }

    return btn;
  }

  window.GameWizardUI = {
    renderGames: function(games, onOpen) {
      const list = document.getElementById('games-list');
      if (!list) {
        return;
      }

      clearNode(list);
      for (let i = 0; i < games.length; i += 1) {
        list.appendChild(createButton(games[i], onOpen));
      }
    },

    setStatus: function(message) {
      const status = document.getElementById('gw-status');
      if (status) {
        status.textContent = message;
      }
    }
  };
})(window);
