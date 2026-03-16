(function(window) {
  'use strict';

  const GAMES = Object.freeze([
    {
      id: 'g2048',
      title: '2048',
      description: 'Пазл на объединение чисел',
      icon: '24',
      modulePath: 'modules/01_2048/index.html',
      size: [520, 820],
      minSize: [420, 620],
      maxSize: [900, 1100]
    },
    {
      id: 'klondike',
      title: 'Solitaire Klondike',
      description: 'Классический пасьянс',
      icon: 'KL',
      modulePath: 'modules/02_cards_Solitaire_Klondike/index.html',
      size: [1024, 820],
      minSize: [760, 620],
      maxSize: [1600, 1100]
    },
    {
      id: 'spider',
      title: 'Solitaire Spider',
      description: 'Пасьянс с уровнями сложности',
      icon: 'SP',
      modulePath: 'modules/03_cards_Solitaire_Spider/index.html',
      size: [1024, 820],
      minSize: [760, 620],
      maxSize: [1600, 1100]
    },
    {
      id: 'freecell',
      title: 'Solitaire FreeCell',
      description: 'Пасьянс со свободными ячейками',
      icon: 'FC',
      modulePath: 'modules/04_cards_Solitaire_FreeCell/index.html',
      size: [1024, 820],
      minSize: [760, 620],
      maxSize: [1600, 1100]
    },
    {
      id: 'chess',
      title: 'Chess',
      description: 'Шахматы против ИИ',
      icon: 'CH',
      modulePath: 'modules/_wrk_chess_optimized/index.html',
      size: [980, 760],
      minSize: [760, 560],
      maxSize: [1500, 1000],
      disabled: true
    }
  ]);

  function getById(id) {
    for (let i = 0; i < GAMES.length; i += 1) {
      if (GAMES[i].id === id) {
        return GAMES[i];
      }
    }
    return null;
  }

  window.GameRegistry = {
    list: function() {
      return GAMES.slice();
    },
    getById: getById
  };
})(window);
