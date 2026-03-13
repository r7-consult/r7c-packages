/**
 * Advanced 2048 UI Renderer
 * Handles all UI rendering with animations, responsive design, and accessibility
 */

class Game2048UIRenderer {
    constructor(container, gridSize = 4) {
        this.container = container;
        this.gridSize = gridSize;
        this.tileElements = new Map();
        this.animationQueue = [];
        this.isAnimating = false;
        
        this.init();
    }
    
    /**
     * Initialize the UI renderer
     */
    init() {
        if (!this.container) {
            console.error('Container element not found');
            return;
        }
        
        this.createBoardStructure();
        this.setupAnimationSystem();
        this.addAccessibilityFeatures();
    }
    
    /**
     * Create the board HTML structure
     */
    createBoardStructure() {
        // Clear existing content safely
        while (this.container.firstChild) {
            this.container.removeChild(this.container.firstChild);
        }
        
        // Create main game container
        const gameContainer = document.createElement('div');
        gameContainer.className = 'game-2048-container';
        gameContainer.setAttribute('role', 'application');
        gameContainer.setAttribute('aria-label', '2048 game board');
        
        // Create header
        const header = this.createHeader();
        gameContainer.appendChild(header);
        
        // Create controls
        const controls = this.createControls();
        gameContainer.appendChild(controls);
        
        // Create game board
        const board = this.createBoard();
        gameContainer.appendChild(board);
        
        // Create message overlay
        const message = this.createMessageOverlay();
        gameContainer.appendChild(message);
        
        // Create loading overlay
        const loading = this.createLoadingOverlay();
        gameContainer.appendChild(loading);
        
        this.container.appendChild(gameContainer);
        
        // Store references
        this.elements = {
            container: gameContainer,
            board: board.querySelector('.game-2048-board'),
            gridContainer: board.querySelector('.game-2048-grid-container'),
            tileContainer: board.querySelector('.game-2048-tile-container'),
            currentScore: header.querySelector('#current-score'),
            bestScore: header.querySelector('#best-score'),
            message: message,
            messageText: message.querySelector('.message-text'),
            loading: loading
        };
    }
    
    /**
     * Create header with scores
     */
    createHeader() {
        const header = document.createElement('div');
        header.className = 'game-2048-header';
        
        const title = document.createElement('h1');
        title.className = 'game-2048-title';
        title.textContent = '2048';
        
        const scores = document.createElement('div');
        scores.className = 'game-2048-scores';
        
        scores.innerHTML = `
            <div class="score-container">
                <div class="score-label">SCORE</div>
                <div class="score-value" id="current-score">0</div>
            </div>
            <div class="score-container">
                <div class="score-label">BEST</div>
                <div class="score-value" id="best-score">0</div>
            </div>
        `;
        
        header.appendChild(title);
        header.appendChild(scores);
        
        return header;
    }
    
    /**
     * Create control buttons
     */
    createControls() {
        const controls = document.createElement('div');
        controls.className = 'game-2048-controls';
        
        const newGameBtn = document.createElement('button');
        newGameBtn.className = 'btn-2048 btn-new-game';
        newGameBtn.id = 'btn-new-game';
        newGameBtn.textContent = 'New Game';
        newGameBtn.setAttribute('aria-label', 'Start a new game');
        
        const undoBtn = document.createElement('button');
        undoBtn.className = 'btn-2048 btn-undo';
        undoBtn.id = 'btn-undo';
        undoBtn.textContent = 'Undo';
        undoBtn.disabled = true;
        undoBtn.setAttribute('aria-label', 'Undo last move');
        
        controls.appendChild(newGameBtn);
        controls.appendChild(undoBtn);
        
        return controls;
    }
    
    /**
     * Create game board
     */
    createBoard() {
        const boardWrapper = document.createElement('div');
        boardWrapper.className = 'game-2048-board-wrapper';
        
        const board = document.createElement('div');
        board.className = 'game-2048-board';
        board.setAttribute('tabindex', '0');
        board.setAttribute('role', 'grid');
        board.setAttribute('aria-label', `${this.gridSize} by ${this.gridSize} game grid`);
        
        // Create grid container
        const gridContainer = document.createElement('div');
        gridContainer.className = 'game-2048-grid-container';
        gridContainer.setAttribute('data-size', this.gridSize);
        
        // Create grid cells
        for (let i = 0; i < this.gridSize * this.gridSize; i++) {
            const cell = document.createElement('div');
            cell.className = 'grid-cell';
            cell.setAttribute('role', 'gridcell');
            gridContainer.appendChild(cell);
        }
        
        // Create tile container
        const tileContainer = document.createElement('div');
        tileContainer.className = 'game-2048-tile-container';
        
        board.appendChild(gridContainer);
        board.appendChild(tileContainer);
        boardWrapper.appendChild(board);
        
        return boardWrapper;
    }
    
    /**
     * Create message overlay
     */
    createMessageOverlay() {
        const message = document.createElement('div');
        message.className = 'game-2048-message';
        message.setAttribute('role', 'alert');
        message.setAttribute('aria-live', 'assertive');
        
        const messageText = document.createElement('p');
        messageText.className = 'message-text';
        
        const keepPlayingBtn = document.createElement('button');
        keepPlayingBtn.className = 'btn-2048 btn-keep-playing';
        keepPlayingBtn.textContent = 'Keep Playing';
        keepPlayingBtn.style.display = 'none';
        
        const tryAgainBtn = document.createElement('button');
        tryAgainBtn.className = 'btn-2048 btn-try-again';
        tryAgainBtn.textContent = 'Try Again';
        tryAgainBtn.style.display = 'none';
        
        message.appendChild(messageText);
        message.appendChild(keepPlayingBtn);
        message.appendChild(tryAgainBtn);
        
        return message;
    }
    
    /**
     * Create loading overlay
     */
    createLoadingOverlay() {
        const loading = document.createElement('div');
        loading.className = 'loading-overlay';
        loading.setAttribute('role', 'status');
        loading.setAttribute('aria-label', 'Loading');
        
        const spinner = document.createElement('div');
        spinner.className = 'loading-spinner';
        
        loading.appendChild(spinner);
        
        return loading;
    }
    
    /**
     * Update the grid with new state
     */
    updateGrid(gridState, previousState = null) {
        if (!gridState || !this.elements.tileContainer) return;
        
        // Calculate moves for animation
        const moves = this.calculateMoves(previousState, gridState);
        
        // Queue animations
        this.queueAnimations(moves);
        
        // Process animation queue
        this.processAnimationQueue();
    }
    
    /**
     * Calculate tile moves for animation
     */
    calculateMoves(previousState, currentState) {
        const moves = [];
        
        if (!previousState) {
            // Initial render - just create tiles
            for (let row = 0; row < this.gridSize; row++) {
                for (let col = 0; col < this.gridSize; col++) {
                    const value = currentState[row][col];
                    if (value > 0) {
                        moves.push({
                            type: 'appear',
                            row,
                            col,
                            value
                        });
                    }
                }
            }
        } else {
            // Calculate actual moves
            // This is simplified - real implementation would track tile movements
            for (let row = 0; row < this.gridSize; row++) {
                for (let col = 0; col < this.gridSize; col++) {
                    const prevValue = previousState[row][col];
                    const currValue = currentState[row][col];
                    
                    if (currValue > 0 && currValue !== prevValue) {
                        if (currValue > prevValue) {
                            moves.push({
                                type: 'merge',
                                row,
                                col,
                                value: currValue
                            });
                        } else {
                            moves.push({
                                type: 'move',
                                row,
                                col,
                                value: currValue
                            });
                        }
                    } else if (currValue > 0) {
                        moves.push({
                            type: 'static',
                            row,
                            col,
                            value: currValue
                        });
                    }
                }
            }
        }
        
        return moves;
    }
    
    /**
     * Queue animations
     */
    queueAnimations(moves) {
        this.animationQueue = moves;
    }
    
    /**
     * Process animation queue
     */
    processAnimationQueue() {
        if (this.isAnimating || this.animationQueue.length === 0) return;
        
        this.isAnimating = true;
        
        // Clear existing tiles
        this.clearTiles();
        
        // Process all moves
        for (const move of this.animationQueue) {
            this.animateTile(move);
        }
        
        // Clear queue
        this.animationQueue = [];
        
        // Reset animation flag after animations complete
        setTimeout(() => {
            this.isAnimating = false;
        }, 150);
    }
    
    /**
     * Animate a single tile
     */
    animateTile(move) {
        const tile = this.createTileElement(move.value, move.row, move.col);
        
        if (move.type === 'appear') {
            tile.classList.add('tile-new');
        } else if (move.type === 'merge') {
            tile.classList.add('tile-merged');
        }
        
        this.elements.tileContainer.appendChild(tile);
        this.tileElements.set(`${move.row}-${move.col}`, tile);
    }
    
    /**
     * Create tile element
     */
    createTileElement(value, row, col) {
        const tile = document.createElement('div');
        tile.className = `tile tile-${value}`;
        tile.textContent = value;
        tile.setAttribute('role', 'gridcell');
        tile.setAttribute('aria-label', `Tile ${value} at row ${row + 1}, column ${col + 1}`);
        
        // Position tile
        const size = 100 / this.gridSize;
        tile.style.width = `calc(${size}% - ${10 * (this.gridSize - 1) / this.gridSize}px)`;
        tile.style.height = `calc(${size}% - ${10 * (this.gridSize - 1) / this.gridSize}px)`;
        tile.style.left = `calc(${col * size}% + ${10 * col / this.gridSize}px)`;
        tile.style.top = `calc(${row * size}% + ${10 * row / this.gridSize}px)`;
        
        // Add super class for large numbers
        if (value > 2048) {
            tile.classList.add('tile-super');
        }
        
        return tile;
    }
    
    /**
     * Clear all tiles
     */
    clearTiles() {
        if (!this.elements.tileContainer) return;
        
        // Remove all tile elements
        while (this.elements.tileContainer.firstChild) {
            this.elements.tileContainer.removeChild(this.elements.tileContainer.firstChild);
        }
        
        this.tileElements.clear();
    }
    
    /**
     * Update score displays
     */
    updateScore(score, bestScore, addition = 0) {
        if (this.elements.currentScore) {
            this.elements.currentScore.textContent = score;
            
            // Show score addition animation
            if (addition > 0) {
                this.showScoreAddition(addition);
            }
        }
        
        if (this.elements.bestScore) {
            this.elements.bestScore.textContent = bestScore;
        }
    }
    
    /**
     * Show score addition animation
     */
    showScoreAddition(points) {
        const addition = document.createElement('div');
        addition.className = 'score-addition';
        addition.textContent = `+${points}`;
        
        const scoreContainer = this.elements.currentScore.parentElement;
        scoreContainer.appendChild(addition);
        
        // Remove after animation
        setTimeout(() => {
            addition.remove();
        }, 600);
    }
    
    /**
     * Show game message
     */
    showMessage(type, text) {
        if (!this.elements.message) return;
        
        this.elements.messageText.textContent = text;
        this.elements.message.classList.add('show');
        
        // Show appropriate buttons
        const keepPlayingBtn = this.elements.message.querySelector('.btn-keep-playing');
        const tryAgainBtn = this.elements.message.querySelector('.btn-try-again');
        
        if (type === 'win') {
            keepPlayingBtn.style.display = 'inline-block';
            tryAgainBtn.style.display = 'inline-block';
        } else if (type === 'lose') {
            keepPlayingBtn.style.display = 'none';
            tryAgainBtn.style.display = 'inline-block';
        }
    }
    
    /**
     * Hide game message
     */
    hideMessage() {
        if (this.elements.message) {
            this.elements.message.classList.remove('show');
        }
    }
    
    /**
     * Show loading state
     */
    showLoading() {
        if (this.elements.loading) {
            this.elements.loading.classList.add('show');
        }
    }
    
    /**
     * Hide loading state
     */
    hideLoading() {
        if (this.elements.loading) {
            this.elements.loading.classList.remove('show');
        }
    }
    
    /**
     * Setup animation system
     */
    setupAnimationSystem() {
        // Check for reduced motion preference
        this.prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
        
        if (this.prefersReducedMotion) {
            // Disable animations
            document.documentElement.style.setProperty('--animation-fast', '0ms');
            document.documentElement.style.setProperty('--animation-normal', '0ms');
            document.documentElement.style.setProperty('--animation-slow', '0ms');
        }
    }
    
    /**
     * Add accessibility features
     */
    addAccessibilityFeatures() {
        // Add live region for announcements
        const liveRegion = document.createElement('div');
        liveRegion.className = 'sr-only';
        liveRegion.setAttribute('role', 'status');
        liveRegion.setAttribute('aria-live', 'polite');
        liveRegion.id = 'game-announcer';
        
        this.container.appendChild(liveRegion);
        this.announcer = liveRegion;
    }
    
    /**
     * Announce to screen readers
     */
    announce(message) {
        if (this.announcer) {
            this.announcer.textContent = message;
        }
    }
    
    /**
     * Cleanup
     */
    cleanup() {
        // Clear tiles
        this.clearTiles();
        
        // Remove event listeners
        this.tileElements.clear();
        
        // Clear animation queue
        this.animationQueue = [];
        
        // Remove elements
        if (this.container) {
            while (this.container.firstChild) {
                this.container.removeChild(this.container.firstChild);
            }
        }
    }
}

// Export for use
window.Game2048UIRenderer = Game2048UIRenderer;