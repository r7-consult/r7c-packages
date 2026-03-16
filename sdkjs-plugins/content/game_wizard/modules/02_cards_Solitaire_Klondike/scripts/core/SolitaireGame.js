/**
 * @summary Main game controller for Solitaire Klondike.
 * @description This class manages the entire game lifecycle, including the game engine,
 * UI rendering, and user input handling. It consolidates the logic from the original
 * SolitaireEngine and SolitaireManager classes into a single, cohesive unit.
 * @version 1.0.0
 * @author Roo
 * 
 * @knowledge_map
 * - Link to Architecture Document: _coding_standard/01_plugin_architecture_guide.md
 * - Link to API Patterns: _coding_standard/02_api_reference_patterns.md
 * 
 * @workflow
 * 1. Instantiated by main.js.
 * 2. `initialize()` is called to set up the DOM and event handlers.
 * 3. `startNewGame()` is called to create and render the initial game state.
 * 4. Responds to user interactions (clicks, drags) to update the game state.
 * 5. `updateUI()` is called after every state change to re-render the board.
 * 6. Checks for win conditions and handles game-over scenarios.
 */
class SolitaireGame {
    constructor(gameBoardId) {
        this.gameBoard = document.getElementById(gameBoardId);
        if (!this.gameBoard) {
            throw new Error(`Game board element with id '${gameBoardId}' not found.`);
        }

        // Game state and logic properties from SolitaireEngine
        this.SUITS = ['hearts', 'diamonds', 'clubs', 'spades'];
        this.RANKS = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
        this.SUIT_COLORS = { 'hearts': 'red', 'diamonds': 'red', 'clubs': 'black', 'spades': 'black' };
        
        this.stock = [];
        this.waste = [];
        this.foundations = [[], [], [], []];
        this.tableau = [[], [], [], [], [], [], []];
        
        this.moves = 0;
        this.score = 0;
        this.gameWon = false;
        this.previousStates = [];
        this.drawCount = 3;

        // UI and interaction properties from SolitaireManager
        this.stockElement = null;
        this.wasteElement = null;
        this.foundationElements = [];
        this.tableauElements = [];
        this.scoreElement = document.getElementById('current-score');
        this.movesElement = document.getElementById('moves-count');
        this.bestScoreElement = document.getElementById('best-score');
        
        this.draggedInfo = null;
        
        this.cardWidth = 70;
        this.cardHeight = 100;
        this.cardOverlap = 20;
        this.tableauFaceDownOverlap = 10;
        this.tableauFaceUpOverlap = 30;

        this.suitSymbols = { 'hearts': '♥', 'diamonds': '♦', 'clubs': '♣', 'spades': '♠' };
        
        this.bestScore = parseInt(localStorage.getItem('solitaire_bestScore') || '0');
    }

    // =========================================================================
    // INITIALIZATION AND SETUP
    // =========================================================================

    initialize() {
        this.createGameLayout();
        this.setupEventHandlers();
        if (this.bestScoreElement) {
            this.bestScoreElement.textContent = this.bestScore;
        }
        console.log('[SolitaireGame] Initialized successfully.');
    }

    createGameLayout() {
        this.gameBoard.innerHTML = '';
        const container = document.createElement('div');
        container.className = 'solitaire-container';

        const topRow = document.createElement('div');
        topRow.className = 'solitaire-top-row';

        this.stockElement = this.createPile('stock-pile', 'Stock');
        topRow.appendChild(this.stockElement);

        this.wasteElement = this.createPile('waste-pile', 'Waste');
        topRow.appendChild(this.wasteElement);

        const spacer = document.createElement('div');
        spacer.className = 'pile-spacer';
        topRow.appendChild(spacer);

        this.foundationElements = Array(4).fill(null).map((_, i) => {
            const foundation = this.createPile(`foundation-${i}`, `Foundation ${i + 1}`);
            foundation.dataset.foundationIndex = i;
            topRow.appendChild(foundation);
            return foundation;
        });

        container.appendChild(topRow);

        const tableauRow = document.createElement('div');
        tableauRow.className = 'solitaire-tableau-row';

        this.tableauElements = Array(7).fill(null).map((_, i) => {
            const column = this.createTableauColumn(i);
            tableauRow.appendChild(column);
            return column;
        });

        container.appendChild(tableauRow);
        this.gameBoard.appendChild(container);
    }
    
    createPile(className, label) {
        const pile = document.createElement('div');
        pile.className = `card-pile ${className}`;
        pile.dataset.label = label;
        const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder';
        pile.appendChild(placeholder);
        return pile;
    }

    createTableauColumn(index) {
        const column = document.createElement('div');
        column.className = 'tableau-column';
        column.dataset.columnIndex = index;
        const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder';
        column.appendChild(placeholder);
        return column;
    }

    setupEventHandlers() {
        this.stockElement.addEventListener('click', () => this.handleStockClick());
        
        const dropZones = [...this.foundationElements, ...this.tableauElements];
        dropZones.forEach(zone => {
            zone.addEventListener('dragover', e => this.handleDragOver(e));
            zone.addEventListener('dragleave', e => this.handleDragLeave(e));
            zone.addEventListener('drop', e => this.handleDrop(e));
        });
    }

    // =========================================================================
    // GAME LOGIC (from SolitaireEngine)
    // =========================================================================

    startNewGame() {
        this.stock = [];
        this.waste = [];
        this.foundations = [[], [], [], []];
        this.tableau = Array(7).fill(null).map(() => []);
        this.moves = 0;
        this.score = 0;
        this.gameWon = false;
        this.previousStates = [];

        const deck = this.shuffleDeck(this.createDeck());

        let cardIndex = 0;
        for (let i = 0; i < 7; i++) {
            for (let j = 0; j <= i; j++) {
                const card = deck[cardIndex++];
                card.faceUp = (j === i);
                this.tableau[i].push(card);
            }
        }
        this.stock = deck.slice(cardIndex);
        
        this.updateUI();
        console.log('[SolitaireGame] New game started.');
    }

    createDeck() {
        return this.SUITS.flatMap(suit => 
            this.RANKS.map(rank => ({
                suit,
                rank,
                color: this.SUIT_COLORS[suit],
                faceUp: false,
                id: `${rank}_${suit}`
            }))
        );
    }
    
    shuffleDeck(deck) {
        for (let i = deck.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [deck[i], deck[j]] = [deck[j], deck[i]];
        }
        return deck;
    }

    getRankValue(rank) {
        const values = { 'A': 1, 'J': 11, 'Q': 12, 'K': 13 };
        return values[rank] || parseInt(rank, 10);
    }

    handleStockClick() {
        this.saveState();
        if (this.stock.length > 0) {
            const cardsToDraw = Math.min(this.drawCount, this.stock.length);
            for (let i = 0; i < cardsToDraw; i++) {
                const card = this.stock.pop();
                card.faceUp = true;
                this.waste.push(card);
            }
        } else if (this.waste.length > 0) {
            this.stock = this.waste.reverse();
            this.stock.forEach(card => card.faceUp = false);
            this.waste = [];
        }
        this.moves++;
        this.updateUI();
    }
    
    move(card, from, to) {
        // This is a placeholder for a more robust move validation system
        // combining canPlaceOnTableau and canPlaceOnFoundation
        
        this.saveState();
        // ... logic to perform move ...
        this.moves++;
        this.score += 5; // Example scoring
        this.checkWinCondition();
        this.updateUI();
    }
    
    checkWinCondition() {
        const totalInFoundations = this.foundations.reduce((sum, f) => sum + f.length, 0);
        if (totalInFoundations === 52) {
            this.gameWon = true;
            this.score += 100;
            console.log('[SolitaireGame] Game Won!');
            this.showWinMessage();
        }
    }
    
    undo() {
        if (this.previousStates.length === 0) return;
        const prevState = JSON.parse(this.previousStates.pop());
        Object.assign(this, prevState);
        this.updateUI();
    }

    saveState() {
        const state = {
            stock: JSON.parse(JSON.stringify(this.stock)),
            waste: JSON.parse(JSON.stringify(this.waste)),
            foundations: JSON.parse(JSON.stringify(this.foundations)),
            tableau: JSON.parse(JSON.stringify(this.tableau)),
            moves: this.moves,
            score: this.score,
            gameWon: this.gameWon
        };
        this.previousStates.push(JSON.stringify(state));
        if (this.previousStates.length > 50) {
            this.previousStates.shift();
        }
    }

    // =========================================================================
    // UI RENDERING AND INTERACTION (from SolitaireManager)
    // =========================================================================

    updateUI() {
        this.renderStock();
        this.renderWaste();
        this.renderFoundations();
        this.renderTableau();
        this.updateStats();
    }

    renderStock() {
        this.stockElement.innerHTML = '';
        const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder';
        if (this.stock.length === 0) {
           placeholder.classList.add('recycle');
        }
        this.stockElement.appendChild(placeholder);

        if (this.stock.length > 0) {
            const topCard = this.createCardElement(this.stock[0]);
            this.stockElement.appendChild(topCard);
        }
    }
    
    renderWaste() {
        this.wasteElement.innerHTML = '';
         const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder';
        this.wasteElement.appendChild(placeholder);
        
        this.waste.slice(-this.drawCount).forEach((cardData, index) => {
            const cardElement = this.createCardElement(cardData);
            cardElement.style.left = `${index * (this.cardOverlap + 5)}px`;
            // Only the top waste card is draggable
            if (this.waste.length > 0 && cardData === this.waste[this.waste.length - 1]) {
                this.makeDraggable(cardElement, { type: 'waste', card: cardData });
            }
            this.wasteElement.appendChild(cardElement);
        });
    }

    renderFoundations() {
        this.foundations.forEach((foundation, index) => {
            const foundationElement = this.foundationElements[index];
            foundationElement.innerHTML = '';
             const placeholder = document.createElement('div');
            placeholder.className = 'card-placeholder';
            foundationElement.appendChild(placeholder);
            if (foundation.length > 0) {
                const topCard = foundation[foundation.length - 1];
                const cardElement = this.createCardElement(topCard);
                foundationElement.appendChild(cardElement);
            }
        });
    }

    renderTableau() {
        this.tableau.forEach((column, colIndex) => {
            const columnElement = this.tableauElements[colIndex];
            let currentTop = 0;
            columnElement.innerHTML = '';
             const placeholder = document.createElement('div');
            placeholder.className = 'card-placeholder';
            columnElement.appendChild(placeholder);
            column.forEach((cardData, cardIndex) => {
                const cardElement = this.createCardElement(cardData);
                cardElement.style.top = `${currentTop}px`;
                if (cardData.faceUp) {
                    this.makeDraggable(cardElement, { type: 'tableau', col: colIndex, index: cardIndex, card: cardData });
                }
                columnElement.appendChild(cardElement);
                currentTop += cardData.faceUp ? this.tableauFaceUpOverlap : this.tableauFaceDownOverlap;
            });
        });
    }

    createCardElement(cardData) {
        const cardDiv = document.createElement('div');
        cardDiv.className = 'playing-card';
        cardDiv.dataset.cardId = cardData.id;

        if (cardData.faceUp) {
            cardDiv.classList.add('face-up', `card-${cardData.color}`);
            cardDiv.innerHTML = `
                <span class="card-rank">${cardData.rank}</span>
                <span class="card-suit">${this.suitSymbols[cardData.suit]}</span>
                <span class="card-suit-center">${this.suitSymbols[cardData.suit]}</span>
            `;
        } else {
            cardDiv.classList.add('face-down');
        }
        return cardDiv;
    }

    updateStats() {
        if (this.scoreElement) this.scoreElement.textContent = this.score;
        if (this.movesElement) this.movesElement.textContent = this.moves;
        if (this.score > this.bestScore) {
            this.bestScore = this.score;
            localStorage.setItem('solitaire_bestScore', this.bestScore);
            if (this.bestScoreElement) this.bestScoreElement.textContent = this.bestScore;
        }
    }
    
    showWinMessage() {
        const overlay = document.createElement('div');
        overlay.className = 'win-overlay';
        overlay.innerHTML = `<div>Congratulations! You won!</div><button id="win-new-game">Play Again</button>`;
        this.gameBoard.appendChild(overlay);
        document.getElementById('win-new-game').addEventListener('click', () => {
            overlay.remove();
            this.startNewGame();
        });
    }

    // =========================================================================
    // DRAG AND DROP HANDLING
    // =========================================================================

    makeDraggable(element, info) {
        element.draggable = true;
        element.addEventListener('dragstart', e => this.handleDragStart(e, info));
    }

    handleDragStart(e, info) {
        this.draggedInfo = info;
        e.dataTransfer.setData('text/plain', JSON.stringify(info));
        e.dataTransfer.effectAllowed = 'move';
    }

    handleDragOver(e) {
        e.preventDefault();
        e.currentTarget.classList.add('drag-over');
    }

    handleDragLeave(e) {
        e.currentTarget.classList.remove('drag-over');
    }
    
    handleDrop(e) {
        e.preventDefault();
        const targetElement = e.currentTarget;
        targetElement.classList.remove('drag-over');

        if (!this.draggedInfo) return;

        const { type: sourceType, card: draggedCard, col: fromCol, index: fromIndex } = this.draggedInfo;

        // Determine drop target
        const isFoundation = targetElement.classList.contains('card-pile') && targetElement.dataset.foundationIndex !== undefined;
        const isTableau = targetElement.classList.contains('tableau-column');

        if (isFoundation) {
            const foundationIndex = parseInt(targetElement.dataset.foundationIndex, 10);
            this.tryMoveToFoundation(draggedCard, sourceType, fromCol, foundationIndex);
        } else if (isTableau) {
            const toCol = parseInt(targetElement.dataset.columnIndex, 10);
            this.tryMoveToTableau(draggedCard, sourceType, fromCol, fromIndex, toCol);
        }

        this.draggedInfo = null;
    }

    tryMoveToFoundation(card, sourceType, fromCol, foundationIndex) {
        const foundation = this.foundations[foundationIndex];
        if (!this.canPlaceOnFoundation(card, foundation)) {
            return;
        }
        this.saveState();

        // Remove card from source
        if (sourceType === 'tableau') {
            // Only top card can go to foundation
            const fromColumn = this.tableau[fromCol];
            if (fromColumn.length === 0 || fromColumn[fromColumn.length - 1] !== card) {
                return; // invalid drag
            }
            fromColumn.pop();
            if (fromColumn.length > 0) {
                fromColumn[fromColumn.length - 1].faceUp = true;
            }
        } else if (sourceType === 'waste') {
            if (this.waste.length === 0 || this.waste[this.waste.length - 1] !== card) {
                return; // invalid drag
            }
            this.waste.pop();
        }

        // Add to foundation
        foundation.push(card);

        this.moves++;
        this.score += 10;
        this.checkWinCondition();
        this.updateUI();
    }

    tryMoveToTableau(card, sourceType, fromCol, fromIndex, toCol) {
        if (!this.canPlaceOnTableau(card, this.tableau[toCol])) {
            return;
        }
        this.saveState();
        
        // Remove card(s) from source
        let cardsToMove = [];
        if (sourceType === 'tableau') {
            // Default to top card if index wasn't set
            if (typeof fromIndex !== 'number') {
                fromIndex = this.tableau[fromCol].length - 1;
            }
            cardsToMove = this.tableau[fromCol].splice(fromIndex);
            if (this.tableau[fromCol].length > 0) {
                this.tableau[fromCol][this.tableau[fromCol].length - 1].faceUp = true;
            }
        } else if (sourceType === 'waste') {
            if (this.waste.length === 0 || this.waste[this.waste.length - 1] !== card) {
                return; // invalid drag
            }
            cardsToMove = [this.waste.pop()];
        }

        // Add card(s) to destination
        this.tableau[toCol].push(...cardsToMove);
        
        this.moves++;
        this.score += 5;
        this.updateUI();
    }

    canPlaceOnTableau(card, targetColumn) {
        if (targetColumn.length === 0) {
            return this.getRankValue(card.rank) === 13; // King
        }
        const targetCard = targetColumn[targetColumn.length - 1];
        return card.color !== targetCard.color && this.getRankValue(card.rank) === this.getRankValue(targetCard.rank) - 1;
    }

    canPlaceOnFoundation(card, foundation) {
        if (foundation.length === 0) {
            return this.getRankValue(card.rank) === 1; // Ace
        }
        const topCard = foundation[foundation.length - 1];
        return card.suit === topCard.suit && this.getRankValue(card.rank) === this.getRankValue(topCard.rank) + 1;
    }
}
