(function(window) {
  'use strict';

  window.GameWizardConfig = Object.freeze({
    branding: {
      title: 'Game Wizard',
      subtitle: 'Выберите игру для запуска'
    },
    local: {
      modulesRoot: 'modules',
      bundledCatalogPath: 'modules/catalog.json',
      localCatalogPath: 'modules/local-catalog.json'
    },
    remote: {
      provider: 'github',
      repositoryUrl: 'https://github.com/r7-consult/game_wizard',
      branch: 'main',
      modulesRoot: 'modules'
    },
    hiddenModules: ['_wrk_chess_optimized']
  });
})(window);
