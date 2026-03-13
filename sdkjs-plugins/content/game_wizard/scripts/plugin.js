(function(window) {
  'use strict';

  class GameWizardPlugin {
    constructor() {
      this.initialized = false;
      this.defaultStatus = 'Готово';
    }

    init() {
      if (this.initialized) {
        return;
      }

      this.resizePanel();
      this.bindUI();
      this.applyTheme(this.readTheme());
      this.initialized = true;
    }

    bindUI() {
      if (!window.GameRegistry || !window.GameWizardUI) {
        throw new Error('UI modules are not loaded');
      }

      const games = window.GameRegistry.list();
      window.GameWizardUI.renderGames(games, this.openGame.bind(this));
      window.GameWizardUI.setStatus(this.defaultStatus);
    }

    openGame(gameId) {
      const game = window.GameRegistry.getById(gameId);
      if (!game) {
        window.GameWizardUI.setStatus('Игра не найдена');
        return;
      }

      if (game.disabled) {
        window.GameWizardUI.setStatus('Временно недоступно: ' + game.title);
        return;
      }

      try {
        window.GameLauncher.open(game);
        window.GameWizardUI.setStatus('Открыто: ' + game.title);
      } catch (error) {
        const details = error && error.message ? error.message : 'неизвестная ошибка';
        window.GameWizardUI.setStatus('Ошибка запуска: ' + game.title + ' (' + details + ')');
        console.error('[GameWizard] Failed to open game', error);
      }
    }

    resizePanel() {
      if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.resizeWindow === 'function') {
        window.Asc.plugin.resizeWindow(420, 760, 340, 500, 760, 1200);
      }
    }

    readTheme() {
      if (window.Asc && window.Asc.plugin && window.Asc.plugin.theme) {
        return window.Asc.plugin.theme;
      }
      if (window.Asc && window.Asc.plugin && window.Asc.plugin.info && window.Asc.plugin.info.theme) {
        return window.Asc.plugin.info.theme;
      }
      return null;
    }

    applyTheme(theme) {
      const isDark = !!(theme && theme.type && String(theme.type).toLowerCase().indexOf('dark') !== -1);
      document.body.classList.toggle('theme-dark', isDark);
    }

    onButton(id, windowId) {
      if (
        windowId &&
        window.Asc &&
        window.Asc.plugin &&
        typeof window.Asc.plugin.executeMethod === 'function'
      ) {
        window.Asc.plugin.executeMethod('CloseWindow', [windowId]);
        return;
      }

      if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.executeCommand === 'function') {
        window.Asc.plugin.executeCommand('close', '');
      }
    }
  }

  const plugin = new GameWizardPlugin();

  function runWhenDomReady(callback) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', callback, { once: true });
      return;
    }
    callback();
  }

  function bootstrapStandalone() {
    runWhenDomReady(function() {
      try {
        plugin.init();
      } catch (error) {
        console.error('[GameWizard] Bootstrap failed', error);
      }
    });
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
    bootstrapStandalone();
  }
})(window);
