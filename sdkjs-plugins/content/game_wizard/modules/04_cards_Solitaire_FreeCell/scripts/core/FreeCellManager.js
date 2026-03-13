/**
 * FreeCell Solitaire Manager - UI and State Management
 * 
 * Responsibilities:
 * - UI rendering and updates for FreeCell layout
 * - Connecting FreeCell game engine to UI
 * - Card display and animations
 * - Drag and drop handling for FreeCell rules
 * - Free cells and foundations display
 * 
 * NO GAME LOGIC - Uses FreeCellEngine for all logic
 */

class FreeCellManager {
    constructor() {
        console.log('[FreeCellManager] Initializing manager');
        
        // Game engine instance
        this.engine = null;
        
        // DOM elements
        this.gameBoard = null;
        this.freeCellElements = [];
        this.foundationElements = [];
        this.tableauElements = [];
        this.scoreElement = null;
        this.movesElement = null;
        
        // Drag state
        this.draggedCards = [];
        this.dragSource = null;
        
        // Card dimensions
        this.cardWidth = 70;
        this.cardHeight = 100;
        this.cardOverlap = 28; // Larger step: less overlap, cards are easier to read
        
        // Card suits and symbols
        this.SUITS = ['hearts', 'diamonds', 'clubs', 'spades'];
        this.suitSymbols = {
            'hearts': '♥',
            'diamonds': '♦',
            'clubs': '♣',
            'spades': '♠'
        };
        
        // Best score from localStorage
        this.bestScore = parseInt(localStorage.getItem('freecell_bestScore') || '0');
        
        console.log('[FreeCellManager] Manager initialized');
    }
    
    /**
     * Initialize the game
     */
    initialize() {
        console.log('[FreeCellManager] Initializing game');
        
        // Get DOM elements
        this.gameBoard = document.getElementById('game-board');
        this.scoreElement = document.getElementById('current-score');
        this.movesElement = document.getElementById('moves-count');
        this.bestScoreElement = document.getElementById('best-score');
        
        if (!this.gameBoard) {
            console.error('[FreeCellManager] Game board element not found!');
            return false;
        }
        
        // Create game engine
        if (typeof FreeCellEngine === 'undefined' && typeof window.FreeCellEngine !== 'undefined') {
            this.engine = new window.FreeCellEngine();
        } else {
            this.engine = new FreeCellEngine();
        }
        
        // Create game layout
        this.createGameLayout();
        
        // Setup event handlers
        this.setupEventHandlers();
        
        // Display best score
        if (this.bestScoreElement) {
            this.bestScoreElement.textContent = this.bestScore;
        }
        
        console.log('[FreeCellManager] Initialization complete');
        return true;
    }
    
    /**
     * Create the game layout for FreeCell
     */
    createGameLayout() {
        console.log('[FreeCellManager] Creating FreeCell game layout');
        
        // Clear board
        this.gameBoard.innerHTML = '';
        
        // Create main container
        const container = document.createElement('div');
        container.className = 'freecell-container';
        
        // Create top row (free cells and foundations)
        const topRow = document.createElement('div');
        topRow.className = 'freecell-top-row';
        
        // Create free cells container
        const freeCellsContainer = document.createElement('div');
        freeCellsContainer.className = 'freecells-container';
        
        this.freeCellElements = [];
        for (let i = 0; i < 4; i++) {
            const freeCell = this.createFreeCell(i);
            this.freeCellElements.push(freeCell);
            freeCellsContainer.appendChild(freeCell);
        }
        
        // Create foundations container
        const foundationsContainer = document.createElement('div');
        foundationsContainer.className = 'foundations-container';
        
        this.foundationElements = [];
        for (let i = 0; i < 4; i++) {
            const foundation = this.createFoundation(i);
            this.foundationElements.push(foundation);
            foundationsContainer.appendChild(foundation);
        }
        
        topRow.appendChild(freeCellsContainer);
        topRow.appendChild(foundationsContainer);
        container.appendChild(topRow);
        
        // Create tableau (8 columns)
        const tableauRow = document.createElement('div');
        tableauRow.className = 'freecell-tableau-row';
        
        this.tableauElements = [];
        for (let i = 0; i < 8; i++) {
            const column = this.createTableauColumn(i);
            this.tableauElements.push(column);
            tableauRow.appendChild(column);
        }
        
        container.appendChild(tableauRow);
        this.gameBoard.appendChild(container);
        
        console.log('[FreeCellManager] FreeCell game layout created');
    }
    
    /**
     * Create a free cell element
     */
    createFreeCell(index) {
        const cell = document.createElement('div');
        cell.className = 'free-cell';
        cell.dataset.cellIndex = index;
        cell.dataset.label = `F${index + 1}`;
        
        // Add placeholder
        const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder';
        cell.appendChild(placeholder);
        
        return cell;
    }
    
    /**
     * Create a foundation pile element
     */
    createFoundation(index) {
        const foundation = document.createElement('div');
        foundation.className = 'foundation-pile';
        foundation.dataset.foundationIndex = index;
        foundation.dataset.suit = this.SUITS[index] || '';
        
        // Add placeholder with suit symbol
        const placeholder = document.createElement('div');
        placeholder.className = 'card-placeholder foundation-placeholder';
        if (this.SUITS[index]) {
            placeholder.innerHTML = `<span class="placeholder-suit">${this.suitSymbols[this.SUITS[index]]}</span>`;
        }
        foundation.appendChild(placeholder);
        
        return foundation;
    }
    
    /**
     * Create a tableau column
     */
    createTableauColumn(index) {
        const column = document.createElement('div');
        column.className = 'freecell-tableau-column';
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
        console.log('[FreeCellManager] Setting up event handlers');
        
        // Setup drag and drop
        this.setupDragAndDrop();
        
        console.log('[FreeCellManager] Event handlers setup complete');
    }
    
    /**
     * Setup drag and drop functionality
     */
    setupDragAndDrop() {
        console.log('[FreeCellManager] Setting up drag and drop');
        
        // Make tableau columns drop zones
        this.tableauElements.forEach((column, index) => {
            column.addEventListener('dragover', (e) => this.handleDragOver(e));
            column.addEventListener('drop', (e) => this.handleDropOnTableau(e, index));
            column.addEventListener('dragleave', (e) => this.handleDragLeave(e));
            column.addEventListener('dragenter', (e) => e.preventDefault());
        });
        
        // Make free cells drop zones
        this.freeCellElements.forEach((cell, index) => {
            cell.addEventListener('dragover', (e) => this.handleDragOver(e));
            cell.addEventListener('drop', (e) => this.handleDropOnFreeCell(e, index));
            cell.addEventListener('dragleave', (e) => this.handleDragLeave(e));
            cell.addEventListener('dragenter', (e) => e.preventDefault());
        });
        
        // Make foundations drop zones
        this.foundationElements.forEach((foundation, index) => {
            foundation.addEventListener('dragover', (e) => this.handleDragOver(e));
            foundation.addEventListener('drop', (e) => this.handleDropOnFoundation(e, index));
            foundation.addEventListener('dragleave', (e) => this.handleDragLeave(e));
            foundation.addEventListener('dragenter', (e) => e.preventDefault());
        });
        
        // Add document-level drop handler to prevent default browser behavior
        document.addEventListener('dragover', (e) => e.preventDefault());
        document.addEventListener('drop', (e) => e.preventDefault());
    }
    
    /**
     * Start new game
     */
    startNewGame() {
        console.log('[FreeCellManager] Starting new game');
        
        if (!this.engine) {
            console.error('[FreeCellManager] Engine not initialized!');
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
        console.log('[FreeCellManager] Rendering game');
        
        // Render free cells
        this.renderFreeCells(gameState.freeCells);
        
        // Render foundations
        this.renderFoundations(gameState.foundations);
        
        // Render tableau (8 columns)
        gameState.tableau.forEach((column, index) => {
            this.renderTableauColumn(column, index);
        });
    }
    
    /**
     * Create a card element
     */
    createCardElement(card, index = 0) {
        const cardDiv = document.createElement('div');
        cardDiv.className = 'playing-card face-up';
        cardDiv.dataset.cardId = card.id;
        cardDiv.dataset.suit = card.suit;
        cardDiv.dataset.rank = card.rank;
        cardDiv.dataset.color = card.color;
        
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
        
        // Set z-index for stacking
        cardDiv.style.zIndex = index;
        
        return cardDiv;
    }
    
    /**
     * Render free cells
     */
    renderFreeCells(freeCells) {
        freeCells.forEach((card, index) => {
            const cellElement = this.freeCellElements[index];
            
            // Clear existing cards
            const existingCards = cellElement.querySelectorAll('.playing-card');
            existingCards.forEach(c => c.remove());
            
            if (card) {
                const cardElement = this.createCardElement(card);
                cardElement.draggable = true;
                cardElement.classList.add('freecell-card');
                this.setupCardDrag(cardElement, 'freecell', index, 0);
                cellElement.appendChild(cardElement);
            }
        });
    }
    
    /**
     * Render foundations
     */
    renderFoundations(foundations) {
        foundations.forEach((foundation, index) => {
            const foundationElement = this.foundationElements[index];
            
            // Clear existing cards
            const existingCards = foundationElement.querySelectorAll('.playing-card');
            existingCards.forEach(c => c.remove());
            
            // Show only the top card
            if (foundation.length > 0) {
                const topCard = foundation[foundation.length - 1];
                const cardElement = this.createCardElement(topCard);
                cardElement.classList.add('foundation-card');
                
                // Foundation cards are not draggable in standard FreeCell
                cardElement.draggable = false;
                
                foundationElement.appendChild(cardElement);
                
                // Update foundation suit indicator
                foundationElement.dataset.suit = topCard.suit;
            }
        });
    }
    
    /**
     * Render tableau column
     */
    renderTableauColumn(column, columnIndex) {
        const columnElement = this.tableauElements[columnIndex];
        const overlapStep = this.getCssOverlap('--card-overlap', this.cardOverlap);
        
        // Clear existing cards
        const cards = columnElement.querySelectorAll('.playing-card');
        cards.forEach(card => card.remove());
        
        // Render all cards in column
        column.forEach((card, cardIndex) => {
            const cardElement = this.createCardElement(card, cardIndex);
            cardElement.classList.add('tableau-card');
            
            // Stack cards vertically
            cardElement.style.top = `${cardIndex * overlapStep}px`;
            
            // Check if this card starts a valid sequence
            const sequenceLength = this.getValidSequenceLength(column, cardIndex);
            
            // Let the engine validate concrete move limits on drop target.
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
            
            columnElement.appendChild(cardElement);
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

        const tail = column.slice(startIndex);
        if (tail.length === 0) return 0;
        if (tail.length === 1) return 1;

        for (let i = 0; i < tail.length - 1; i++) {
            const currentCard = tail[i];
            const nextCard = tail[i + 1];

            if (currentCard.color === nextCard.color) {
                return 0;
            }

            if (this.engine.getRankValue(nextCard.rank) !== this.engine.getRankValue(currentCard.rank) - 1) {
                return 0;
            }
        }

        return tail.length;
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
        console.log('[FreeCellManager] Drag started', { source, columnIndex, cardIndex });

        const sequenceLength = parseInt(e.target.dataset.sequenceLength, 10) || 1;
        this.dragSource = { source, columnIndex, cardIndex, sequenceLength };
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'card');
        e.target.classList.add('dragging');
        this.clearDropZoneHighlights();
        
        // Highlight the sequence being dragged
        if (source === 'tableau') {
            if (sequenceLength > 1) {
                const column = this.tableauElements[columnIndex];
                const cards = column.querySelectorAll('.playing-card');
                for (let i = cardIndex; i < cardIndex + sequenceLength && i < cards.length; i++) {
                    cards[i].classList.add('dragging-sequence');
                }
            }
        }
        
        // Store drag data globally as backup
        window.currentDragSource = this.dragSource;
    }
    
    /**
     * Handle drag end
     */
    handleDragEnd(e) {
        console.log('[FreeCellManager] Drag ended');
        
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
        console.log('[FreeCellManager] Drop on tableau', tableauIndex, 'from', source);
        
        if (!source) {
            console.log('[FreeCellManager] No drag source found');
            return;
        }
        
        let success = false;
        
        if (source.source === 'tableau' && source.columnIndex !== tableauIndex) {
            console.log(`[FreeCellManager] Moving tableau ${source.columnIndex}[${source.cardIndex}] to ${tableauIndex}`);
            success = this.engine.moveTableauToTableau(
                source.columnIndex,
                source.cardIndex,
                tableauIndex,
                source.sequenceLength || 1
            );
        } else if (source.source === 'freecell') {
            console.log(`[FreeCellManager] Moving free cell ${source.columnIndex} to tableau ${tableauIndex}`);
            success = this.engine.moveFromFreeCell(
                source.columnIndex,
                'tableau',
                tableauIndex
            );
        }
        
        if (success) {
            console.log('[FreeCellManager] Move successful, updating game');
            this.updateGame();
        } else {
            console.log('[FreeCellManager] Move failed');
        }
        
        // Clear backup
        window.currentDragSource = null;
    }
    
    /**
     * Handle drop on free cell
     */
    handleDropOnFreeCell(e, freeCellIndex) {
        e.preventDefault();
        e.stopPropagation();
        this.clearDropZoneHighlights();
        
        const source = this.dragSource || window.currentDragSource;
        console.log('[FreeCellManager] Drop on free cell', freeCellIndex, 'from', source);
        
        if (!source) return;
        
        let success = false;
        
        if (source.source === 'tableau') {
            // Only allow moving the top card to free cell
            const column = this.engine.tableau[source.columnIndex];
            if (source.cardIndex === column.length - 1) {
                success = this.engine.moveToFreeCell('tableau', source.columnIndex, freeCellIndex);
            }
        }
        
        if (success) {
            console.log('[FreeCellManager] Move to free cell successful');
            this.updateGame();
        }
        
        window.currentDragSource = null;
    }
    
    /**
     * Handle drop on foundation
     */
    handleDropOnFoundation(e, foundationIndex) {
        e.preventDefault();
        e.stopPropagation();
        this.clearDropZoneHighlights();
        
        const source = this.dragSource || window.currentDragSource;
        console.log('[FreeCellManager] Drop on foundation', foundationIndex, 'from', source);
        
        if (!source) return;
        
        let success = false;
        
        if (source.source === 'tableau') {
            // Only allow moving the top card to foundation
            const column = this.engine.tableau[source.columnIndex];
            if (source.cardIndex === column.length - 1) {
                success = this.engine.moveToFoundation('tableau', source.columnIndex, foundationIndex);
            }
        } else if (source.source === 'freecell') {
            success = this.engine.moveToFoundation('freecell', source.columnIndex, foundationIndex);
        }
        
        if (success) {
            console.log('[FreeCellManager] Move to foundation successful');
            this.updateGame();
        }
        
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
        console.log('[FreeCellManager] Card double-clicked');
        e.stopPropagation();
        
        if (source === 'tableau') {
            // Try to move to foundation first
            for (let f = 0; f < 4; f++) {
                if (this.engine.moveToFoundation('tableau', columnIndex, f)) {
                    this.updateGame();
                    return;
                }
            }
            
            // Try to move to free cell
            for (let fc = 0; fc < 4; fc++) {
                if (this.engine.moveToFreeCell('tableau', columnIndex, fc)) {
                    this.updateGame();
                    return;
                }
            }
            
            // Try to move to another tableau column
            for (let toCol = 0; toCol < 8; toCol++) {
                if (toCol !== columnIndex) {
                    if (this.engine.moveTableauToTableau(columnIndex, cardIndex, toCol)) {
                        this.updateGame();
                        return;
                    }
                }
            }
        } else if (source === 'freecell') {
            // Try to move to foundation first
            for (let f = 0; f < 4; f++) {
                if (this.engine.moveToFoundation('freecell', columnIndex, f)) {
                    this.updateGame();
                    return;
                }
            }
            
            // Try to move to tableau
            for (let col = 0; col < 8; col++) {
                if (this.engine.moveFromFreeCell(columnIndex, 'tableau', col)) {
                    this.updateGame();
                    return;
                }
            }
        }
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
        
        // Update best score if needed
        if (gameState.score > this.bestScore) {
            this.bestScore = gameState.score;
            localStorage.setItem('freecell_bestScore', this.bestScore.toString());
            
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
     * Show win message
     */
    showWinMessage() {
        console.log('[FreeCellManager] Showing win message');
        
        const message = window.app?.state?.strings?.youWin || 'You Win!';
        const score = this.engine.score;
        
        // Create overlay
        const overlay = document.createElement('div');
        overlay.className = 'plugin-modal';
        overlay.innerHTML = `
            <div class="modal-title">${message} 🎉</div>
            <div class="modal-score">Score: ${score}</div>
            <div class="modal-score">Moves: ${this.engine.moves}</div>
            <button class="btn-plugin primary" onclick="window.freeCellManager.startNewGame(); this.parentElement.remove();">
                New Game
            </button>
        `;
        
        document.body.appendChild(overlay);
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
            console.log('[FreeCellManager] Hint:', hint);
            // Could add visual highlighting here
        } else {
            console.log('[FreeCellManager] No hints available');
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
    module.exports = FreeCellManager;
}

// Also expose to global window object for browser usage
if (typeof window !== 'undefined') {
    window.FreeCellManager = FreeCellManager;
    window.freeCellManager = null; // Will be initialized by app.js
}

