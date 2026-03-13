/**
 * Spider Solitaire Manager - UI and State Management
 * 
 * Responsibilities:
 * - UI rendering and updates for 10-column layout
 * - Connecting Spider game engine to UI
 * - Card display and animations
 * - Drag and drop handling for Spider rules
 * - Score and completed suits display
 * 
 * NO GAME LOGIC - Uses SpiderEngine for all logic
 */

class SpiderManager {
    constructor() {
        console.log('[SpiderManager] Initializing manager');
        
        // Game engine instance
        this.engine = null;
        
        // DOM elements
        this.gameBoard = null;
        this.stockElement = null;
        this.tableauElements = [];
        this.scoreElement = null;
        this.movesElement = null;
        this.suitsElement = null; // Display completed suits count
        this.dealsElement = null; // Display remaining deals
        
        // Drag state
        this.draggedCards = [];
        this.dragSource = null;
        
        // Card dimensions (smaller for 10 columns)
        this.cardWidth = 60;
        this.cardHeight = 85;
        this.cardOverlap = 18; // Legacy fallback
        this.faceDownOverlap = 8; // Closed cards: tighter stacking
        this.faceUpOverlap = 24; // Open cards: larger spacing for readability
        
        // Card symbols
        this.suitSymbols = {
            'hearts': '♥',
            'diamonds': '♦',
            'clubs': '♣',
            'spades': '♠'
        };
        
        // Difficulty level
        this.difficulty = 1; // Default to 1 suit
        
        // Best score from localStorage
        this.bestScore = parseInt(localStorage.getItem('spider_bestScore') || '0');
        
        console.log('[SpiderManager] Manager initialized');
    }
    
    /**
     * Initialize the game
     */
    initialize(difficulty = 1) {
        console.log('[SpiderManager] Initializing game with difficulty:', difficulty);
        
        this.difficulty = difficulty;
        
        // Get DOM elements
        this.gameBoard = document.getElementById('game-board');
        this.scoreElement = document.getElementById('current-score');
        this.movesElement = document.getElementById('moves-count');
        this.bestScoreElement = document.getElementById('best-score');
        this.suitsElement = document.getElementById('suits-completed');
        this.dealsElement = document.getElementById('deals-remaining');
        
        if (!this.gameBoard) {
            console.error('[SpiderManager] Game board element not found!');
            return false;
        }
        
        // Create game engine with selected difficulty
        if (typeof SpiderEngine === 'undefined' && typeof window.SpiderEngine !== 'undefined') {
            this.engine = new window.SpiderEngine(difficulty);
        } else {
            this.engine = new SpiderEngine(difficulty);
        }
        
        // Create game layout
        this.createGameLayout();
        
        // Setup event handlers
        this.setupEventHandlers();
        
        // Display best score
        if (this.bestScoreElement) {
            this.bestScoreElement.textContent = this.bestScore;
        }
        
        console.log('[SpiderManager] Initialization complete');
        return true;
    }
    
    /**
     * Create the game layout for Spider (10 columns)
     */
    createGameLayout() {
        console.log('[SpiderManager] Creating Spider game layout');
        
        // Clear board
        this.gameBoard.innerHTML = '';
        
        // Create main container
        const container = document.createElement('div');
        container.className = 'spider-container';
        
        // Create top row (stock and info)
        const topRow = document.createElement('div');
        topRow.className = 'spider-top-row';
        
        // Create stock pile
        this.stockElement = this.createStockPile();
        topRow.appendChild(this.stockElement);
        
        // Create completed suits display
        const suitsDisplay = document.createElement('div');
        suitsDisplay.className = 'suits-display';
        suitsDisplay.innerHTML = `
            <div class="info-label">Suits Completed</div>
            <div class="info-value" id="suits-completed">0 / 8</div>
        `;
        topRow.appendChild(suitsDisplay);
        
        // Create deals remaining display
        const dealsDisplay = document.createElement('div');
        dealsDisplay.className = 'deals-display';
        dealsDisplay.innerHTML = `
            <div class="info-label">Deals Left</div>
            <div class="info-value" id="deals-remaining">5</div>
        `;
        topRow.appendChild(dealsDisplay);
        
        container.appendChild(topRow);
        
        // Create tableau (10 columns)
        const tableauRow = document.createElement('div');
        tableauRow.className = 'spider-tableau-row';
        
        this.tableauElements = [];
        for (let i = 0; i < 10; i++) {
            const column = this.createTableauColumn(i);
            this.tableauElements.push(column);
            tableauRow.appendChild(column);
        }
        
        container.appendChild(tableauRow);
        this.gameBoard.appendChild(container);
        
        // Update element references
        this.suitsElement = document.getElementById('suits-completed');
        this.dealsElement = document.getElementById('deals-remaining');
        
        console.log('[SpiderManager] Spider game layout created');
    }
    
    /**
     * Create stock pile element
     */
    createStockPile() {
        const stockContainer = document.createElement('div');
        stockContainer.className = 'spider-stock-container';
        
        const pile = document.createElement('div');
        pile.className = 'card-pile stock-pile';
        pile.dataset.label = 'Stock';
        pile.title = 'Click to deal 10 cards';
        
        // Add placeholder
        const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder';
        pile.appendChild(placeholder);
        
        stockContainer.appendChild(pile);
        
        return pile;
    }
    
    /**
     * Create a tableau column
     */
    createTableauColumn(index) {
        const column = document.createElement('div');
        column.className = 'spider-tableau-column';
        column.dataset.columnIndex = index;
        
        // Add placeholder
        const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder';
        column.appendChild(placeholder);
        
        return column;
    }
    
    /**
     * Setup event handlers
     */
    setupEventHandlers() {
        console.log('[SpiderManager] Setting up event handlers');
        
        // Stock pile click handler (deals 10 cards)
        if (this.stockElement) {
            this.stockElement.addEventListener('click', () => this.handleStockClick());
        }
        
        // Setup drag and drop
        this.setupDragAndDrop();
        
        console.log('[SpiderManager] Event handlers setup complete');
    }
    
    /**
     * Setup drag and drop functionality
     */
    setupDragAndDrop() {
        console.log('[SpiderManager] Setting up drag and drop');
        
        // Make tableau columns drop zones
        this.tableauElements.forEach((column, index) => {
            column.addEventListener('dragover', (e) => this.handleDragOver(e));
            column.addEventListener('drop', (e) => this.handleDropOnTableau(e, index));
            column.addEventListener('dragleave', (e) => this.handleDragLeave(e));
            column.addEventListener('dragenter', (e) => e.preventDefault());
        });
        
        // Add document-level drop handler to prevent default browser behavior
        document.addEventListener('dragover', (e) => e.preventDefault());
        document.addEventListener('drop', (e) => e.preventDefault());
    }
    
    /**
     * Start new game
     */
    startNewGame() {
        console.log('[SpiderManager] Starting new game');
        
        if (!this.engine) {
            console.error('[SpiderManager] Engine not initialized!');
            return;
        }
        
        const gameState = this.engine.startNewGame();
        this.renderGame(gameState);
        this.updateStats(gameState);
    }
    
    /**
     * Render the entire game
     */
    renderGame(gameState) {
        console.log('[SpiderManager] Rendering game');
        
        // Render stock
        this.renderStock(gameState.stock);
        
        // Render tableau (10 columns)
        gameState.tableau.forEach((column, index) => {
            this.renderTableauColumn(column, index);
        });
    }
    
    /**
     * Create a card element
     */
    createCardElement(card, index = 0) {
        const cardDiv = document.createElement('div');
        cardDiv.className = 'playing-card';
        cardDiv.dataset.cardId = card.id;
        cardDiv.dataset.suit = card.suit;
        cardDiv.dataset.rank = card.rank;
        cardDiv.dataset.color = card.color;
        
        if (card.faceUp) {
            cardDiv.classList.add('face-up');
            cardDiv.classList.add(`card-${card.color}`);
            
            // Add rank and suit
            const rankSpan = document.createElement('span');
            rankSpan.className = 'card-rank';
            rankSpan.textContent = card.rank;
            cardDiv.appendChild(rankSpan);
            
            const suitSpan = document.createElement('span');
            suitSpan.className = 'card-suit';
            suitSpan.textContent = this.suitSymbols[card.suit];
            cardDiv.appendChild(suitSpan);
            
            // Add center suit
            const centerSuit = document.createElement('span');
            centerSuit.className = 'card-suit-center';
            centerSuit.textContent = this.suitSymbols[card.suit];
            cardDiv.appendChild(centerSuit);
        } else {
            cardDiv.classList.add('face-down');
        }
        
        // Set z-index for stacking
        cardDiv.style.zIndex = index;
        
        return cardDiv;
    }
    
    /**
     * Render stock pile
     */
    renderStock(stock) {
        // Clear existing cards
        const cards = this.stockElement.querySelectorAll('.playing-card');
        cards.forEach(card => card.remove());
        
        if (stock.length > 0) {
            // Show stack of cards
            const numToShow = Math.min(5, Math.floor(stock.length / 10));
            for (let i = 0; i < numToShow; i++) {
                const cardElement = document.createElement('div');
                cardElement.className = 'playing-card face-down stock-card';
                cardElement.style.top = `${i * 2}px`;
                cardElement.style.left = `${i * 2}px`;
                this.stockElement.appendChild(cardElement);
            }
            
            // Add count indicator
            const countBadge = document.createElement('span');
            countBadge.className = 'card-count';
            countBadge.textContent = stock.length;
            this.stockElement.appendChild(countBadge);
        }
    }
    
    /**
     * Render tableau column
     */
    renderTableauColumn(column, columnIndex) {
        const columnElement = this.tableauElements[columnIndex];
        const faceDownStep = this.getCssOverlap('--card-overlap-face-down', this.faceDownOverlap);
        const faceUpStep = this.getCssOverlap('--card-overlap-face-up', this.faceUpOverlap);
        let currentTop = 0;
        
        // Clear existing cards
        const cards = columnElement.querySelectorAll('.playing-card');
        cards.forEach(card => card.remove());
        
        // Render all cards in column
        column.forEach((card, cardIndex) => {
            const cardElement = this.createCardElement(card, cardIndex);
            cardElement.classList.add('tableau-card');
            
            // Closed cards stack tighter, open cards have larger spacing.
            cardElement.style.top = `${currentTop}px`;
            
            // Check if this card starts a valid sequence
            if (card.faceUp) {
                // Find the longest valid sequence starting from this card
                const sequenceLength = this.getValidSequenceLength(column, cardIndex);
                
                if (sequenceLength > 0) {
                    cardElement.draggable = true;
                    cardElement.classList.add('draggable');
                    this.setupCardDrag(cardElement, 'tableau', columnIndex, cardIndex);
                    
                    // Add visual indicator for sequence
                    if (sequenceLength > 1) {
                        cardElement.classList.add('sequence-start');
                        cardElement.dataset.sequenceLength = sequenceLength;
                    }
                }
            }
            
            columnElement.appendChild(cardElement);
            currentTop += card.faceUp ? faceUpStep : faceDownStep;
        });
    }

    getCssOverlap(variableName, fallbackValue) {
        const rawValue = getComputedStyle(document.documentElement)
            .getPropertyValue(variableName)
            .trim();
        const parsedValue = parseFloat(rawValue);
        return Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : fallbackValue;
    }
    
    /**
     * Get the length of valid sequence starting from index
     */
    getValidSequenceLength(column, startIndex) {
        if (startIndex >= column.length) return 0;
        if (!column[startIndex].faceUp) return 0;
        
        let length = 1;
        const suit = column[startIndex].suit;
        let expectedValue = this.engine.getRankValue(column[startIndex].rank);
        
        for (let i = startIndex + 1; i < column.length; i++) {
            if (!column[i].faceUp) break;
            if (column[i].suit !== suit) break;
            
            expectedValue--;
            if (this.engine.getRankValue(column[i].rank) !== expectedValue) break;
            
            length++;
        }
        
        return length;
    }
    
    /**
     * Setup drag handlers for a card
     */
    setupCardDrag(cardElement, source, columnIndex, cardIndex) {
        cardElement.addEventListener('dragstart', (e) => {
            this.handleDragStart(e, source, columnIndex, cardIndex);
        });
        
        cardElement.addEventListener('dragend', (e) => {
            this.handleDragEnd(e);
        });
        
        // Add double-click for auto-move
        cardElement.addEventListener('dblclick', (e) => {
            this.handleCardDoubleClick(e, source, columnIndex, cardIndex);
        });
    }
    
    /**
     * Handle drag start
     */
    handleDragStart(e, source, columnIndex, cardIndex) {
        console.log('[SpiderManager] Drag started', { source, columnIndex, cardIndex });
        
        this.dragSource = { source, columnIndex, cardIndex };
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'card');
        e.target.classList.add('dragging');
        this.clearDropZoneHighlights();
        
        // Highlight the sequence being dragged
        const sequenceLength = parseInt(e.target.dataset.sequenceLength) || 1;
        if (sequenceLength > 1) {
            const column = this.tableauElements[columnIndex];
            const cards = column.querySelectorAll('.playing-card');
            for (let i = cardIndex; i < cardIndex + sequenceLength && i < cards.length; i++) {
                cards[i].classList.add('dragging-sequence');
            }
        }
        
        // Store drag data globally as backup
        window.currentDragSource = this.dragSource;
    }
    
    /**
     * Handle drag end
     */
    handleDragEnd(e) {
        console.log('[SpiderManager] Drag ended');
        
        e.target.classList.remove('dragging');
        
        // Remove sequence highlighting
        document.querySelectorAll('.dragging-sequence').forEach(el => {
            el.classList.remove('dragging-sequence');
        });
        
        this.clearDropZoneHighlights();
        
        this.dragSource = null;
        window.currentDragSource = null;
    }
    
    /**
     * Handle drag over (allow drop)
     */
    handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        this.clearDropZoneHighlights(e.currentTarget);
        e.currentTarget.classList.add('drop-zone-highlight');
    }
    
    /**
     * Handle drag leave
     */
    handleDragLeave(e) {
        if (!e.currentTarget.contains(e.relatedTarget)) {
            e.currentTarget.classList.remove('drop-zone-highlight');
        }
    }
    
    /**
     * Handle drop on tableau
     */
    handleDropOnTableau(e, tableauIndex) {
        e.preventDefault();
        e.stopPropagation();
        this.clearDropZoneHighlights();
        
        // Use backup source if main source is lost
        const source = this.dragSource || window.currentDragSource;
        console.log('[SpiderManager] Drop on tableau', tableauIndex, 'from', source);
        
        if (!source) {
            console.log('[SpiderManager] No drag source found');
            return;
        }
        
        let success = false;
        
        if (source.source === 'tableau' && source.columnIndex !== tableauIndex) {
            console.log(`[SpiderManager] Moving tableau ${source.columnIndex}[${source.cardIndex}] to ${tableauIndex}`);
            success = this.engine.moveTableauToTableau(
                source.columnIndex,
                source.cardIndex,
                tableauIndex
            );
        }
        
        if (success) {
            console.log('[SpiderManager] Move successful, updating game');
            this.updateGame();
        } else {
            console.log('[SpiderManager] Move failed');
        }
        
        // Clear backup
        window.currentDragSource = null;
    }

    clearDropZoneHighlights(exceptElement = null) {
        document.querySelectorAll('.drop-zone-highlight').forEach((el) => {
            if (el !== exceptElement) {
                el.classList.remove('drop-zone-highlight');
            }
        });
    }
    
    /**
     * Handle card double-click (try to find best move)
     */
    handleCardDoubleClick(e, source, columnIndex, cardIndex) {
        console.log('[SpiderManager] Card double-clicked');
        e.stopPropagation();
        
        // Try to find a valid move for this card/sequence
        if (source === 'tableau') {
            // Try each column as target
            for (let toCol = 0; toCol < 10; toCol++) {
                if (toCol !== columnIndex) {
                    if (this.engine.moveTableauToTableau(columnIndex, cardIndex, toCol)) {
                        this.updateGame();
                        return;
                    }
                }
            }
        }
    }
    
    /**
     * Handle stock click (deal 10 cards)
     */
    handleStockClick() {
        console.log('[SpiderManager] Stock clicked');
        
        if (this.engine.dealFromStock()) {
            this.updateGame();
            
            // Add animation for dealing
            this.animateDealing();
        } else {
            console.log('[SpiderManager] Cannot deal - empty columns or not enough cards');
            this.showMessage('Fill all columns before dealing');
        }
    }
    
    /**
     * Animate dealing cards
     */
    animateDealing() {
        // Simple animation - cards appear with slight delay
        this.tableauElements.forEach((column, index) => {
            setTimeout(() => {
                const cards = column.querySelectorAll('.playing-card');
                if (cards.length > 0) {
                    const lastCard = cards[cards.length - 1];
                    lastCard.classList.add('card-dealt');
                    setTimeout(() => lastCard.classList.remove('card-dealt'), 500);
                }
            }, index * 50);
        });
    }
    
    /**
     * Update game after a move
     */
    updateGame() {
        const gameState = this.engine.getGameState();
        this.renderGame(gameState);
        this.updateStats(gameState);
        
        if (gameState.gameWon) {
            this.showWinMessage();
        }
    }
    
    /**
     * Update statistics display
     */
    updateStats(gameState) {
        if (this.scoreElement) {
            this.scoreElement.textContent = gameState.score;
        }
        
        if (this.movesElement) {
            this.movesElement.textContent = gameState.moves;
        }
        
        if (this.suitsElement) {
            this.suitsElement.textContent = `${gameState.completedSuits.length} / 8`;
        }
        
        if (this.dealsElement) {
            this.dealsElement.textContent = gameState.dealsRemaining;
        }
        
        // Update best score if needed
        if (gameState.score > this.bestScore) {
            this.bestScore = gameState.score;
            localStorage.setItem('spider_bestScore', this.bestScore.toString());
            
            if (this.bestScoreElement) {
                this.bestScoreElement.textContent = this.bestScore;
            }
        }
        
        // Update undo button state
        const undoBtn = document.getElementById('btn-undo');
        if (undoBtn) {
            undoBtn.disabled = !gameState.canUndo;
        }
    }
    
    /**
     * Show message to user
     */
    showMessage(message) {
        const notification = document.createElement('div');
        notification.className = 'spider-notification';
        notification.textContent = message;
        document.body.appendChild(notification);
        
        setTimeout(() => notification.remove(), 3000);
    }
    
    /**
     * Show win message
     */
    showWinMessage() {
        console.log('[SpiderManager] Showing win message');
        
        const message = window.app?.state?.strings?.youWin || 'You Win!';
        const score = this.engine.score;
        
        // Create overlay
        const overlay = document.createElement('div');
        overlay.className = 'plugin-modal';
        overlay.innerHTML = `
            <div class="modal-title">${message} 🎉</div>
            <div class="modal-score">Score: ${score}</div>
            <div class="modal-score">Moves: ${this.engine.moves}</div>
            <div class="modal-score">Difficulty: ${this.getDifficultyName()}</div>
            <button class="btn-plugin primary" onclick="window.spiderManager.startNewGame(); this.parentElement.remove();">
                New Game
            </button>
        `;
        
        document.body.appendChild(overlay);
    }
    
    /**
     * Get difficulty name for display
     */
    getDifficultyName() {
        switch(this.difficulty) {
            case 1: return 'One Suit';
            case 2: return 'Two Suits';
            case 4: return 'Four Suits';
            default: return 'Unknown';
        }
    }
    
    /**
     * Undo last move
     */
    undo() {
        if (this.engine.undo()) {
            this.updateGame();
        }
    }
    
    /**
     * Get hint
     */
    getHint() {
        const hint = this.engine.getHint();
        if (hint) {
            console.log('[SpiderManager] Hint:', hint);
            // Highlight suggested move
            if (hint.type === 'move_sequence') {
                const fromColumn = this.tableauElements[hint.from];
                const cards = fromColumn.querySelectorAll('.playing-card');
                if (cards[hint.fromIndex]) {
                    cards[hint.fromIndex].classList.add('hint-highlight');
                    setTimeout(() => cards[hint.fromIndex].classList.remove('hint-highlight'), 2000);
                }
            } else if (hint.type === 'deal_from_stock') {
                this.stockElement.classList.add('hint-highlight');
                setTimeout(() => this.stockElement.classList.remove('hint-highlight'), 2000);
            }
        } else {
            console.log('[SpiderManager] No hints available');
        }
        return hint;
    }
    
    /**
     * Get current game state
     * Required by app.js for UI updates
     */
    getGameState() {
        if (!this.engine) {
            return {
                score: 0,
                moves: 0,
                canUndo: false,
                gameWon: false
            };
        }
        return this.engine.getGameState();
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpiderManager;
}

// Also expose to global window object for browser usage
if (typeof window !== 'undefined') {
    window.SpiderManager = SpiderManager;
    window.spiderManager = null; // Will be initialized by app.js
}
