(function(window) {
  'use strict';

  window.GameWizardConfig = Object.freeze({
    branding: {
      title: 'Обработка данных',
      subtitle: 'Каталог инструментов обработки данных'
    },
    local: {
      modulesRoot: 'modules',
      bundledCatalogPath: 'modules/catalog.json',
      localCatalogPath: 'modules/local-catalog.json'
    },
    remote: {
      provider: 'github',
      repositoryUrl: 'https://github.com/r7-consult/data_processing',
      branch: 'main',
      modulesRoot: 'modules'
    },
    hiddenModules: []
  });
})(window);
