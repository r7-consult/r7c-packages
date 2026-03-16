/**
 * Game 2048 Manager - UI and State Management
 * 
 * Responsibilities:
 * - UI rendering and updates
 * - Connecting game engine to UI
 * - Score display management
 * - Game state persistence
 * 
 * NO GAME LOGIC - Uses Game2048Engine for all logic
 */
class Game2048Manager {
    constructor() {
        console.log('[Game2048Manager] Initializing manager');
        
        // Game engine instance
        this.engine = null;
        
        // DOM elements
        this.boardElement = null;
        this.scoreElement = null;
        this.bestScoreElement = null;
        
        // Tile colors configuration
        this.tileColors = {
            2: { bg: '#eee4da', text: '#776e65' },
            4: { bg: '#ede0c8', text: '#776e65' },
            8: { bg: '#f2b179', text: '#f9f6f2' },
            16: { bg: '#f59563', text: '#f9f6f2' },
            32: { bg: '#f67c5f', text: '#f9f6f2' },
            64: { bg: '#f65e3b', text: '#f9f6f2' },
            128: { bg: '#edcf72', text: '#f9f6f2' },
            256: { bg: '#edcc61', text: '#f9f6f2' },
            512: { bg: '#edc850', text: '#f9f6f2' },
            1024: { bg: '#edc53f', text: '#f9f6f2' },
            2048: { bg: '#edc22e', text: '#f9f6f2' },
            4096: { bg: '#3c3a32', text: '#f9f6f2' },
            8192: { bg: '#3c3a32', text: '#f9f6f2' }
        };
        
        // Best score from localStorage
        this.bestScore = parseInt(localStorage.getItem('2048_bestScore') || '0');
    }
    
    /**
     * Initialize the game
     */
    initialize() {
        console.log('[Game2048Manager] Initializing game');
        
        // Get DOM elements
        this.boardElement = document.getElementById('game-board');
        this.scoreElement = document.getElementById('current-score');
        this.bestScoreElement = document.getElementById('best-score');
        
        if (!this.boardElement) {
            console.error('[Game2048Manager] Board element not found!');
            return false;
        }
        
        // Create game engine
        this.engine = new Game2048Engine();
        
        // Display best score
        if (this.bestScoreElement) {
            this.bestScoreElement.textContent = this.bestScore;
        }
        
        console.log('[Game2048Manager] Initialization complete');
        return true;
    }
    
    /**
     * Start new game
     */
    startNewGame() {
        console.log('[Game2048Manager] Starting new game');
        
        if (!this.engine) {
            console.error('[Game2048Manager] Engine not initialized!');
            return;
        }
        
        const gameState = this.engine.startNewGame();
        this.renderBoard(gameState);
        this.updateScore(gameState.score);
    }
    
    /**
     * Handle move
     */
    makeMove(direction) {
        if (!this.engine) {
            console.error('[Game2048Manager] Engine not initialized!');
            return;
        }
        
        const moved = this.engine.move(direction);
        
        if (moved) {
            const gameState = this.engine.getGameState();
            this.renderBoard(gameState);
            this.updateScore(gameState.score);
            
            // Check for game over or win
            if (gameState.gameWon) {
                const message = window.app?.state?.strings?.youWin || 'You Win! 🎉';
                this.showMessage(message + ' 🎉');
            } else if (gameState.gameOver) {
                const message = window.app?.state?.strings?.gameOver || 'Game Over!';
                this.showMessage(message + ' 😢');
            }
        }
    }
    
    /**
     * Undo last move
     */
    undo() {
        if (!this.engine) return;
        
        const undone = this.engine.undo();
        if (undone) {
            const gameState = this.engine.getGameState();
            this.renderBoard(gameState);
            this.updateScore(gameState.score);
        }
    }
    
    /**
     * Render the game board
     */
    renderBoard(gameState) {
        if (!this.boardElement) return;
        
        console.log('[Game2048Manager] Rendering board');
        
        // Clear current board
        this.boardElement.innerHTML = '';
        
        // Create board container with plugin styles
        const boardContainer = document.createElement('div');
        boardContainer.className = 'game-board-container';
        
        // Create grid background
        const gridContainer = document.createElement('div');
        gridContainer.className = 'board-grid';
        
        // Add grid cells
        for (let i = 0; i < 25; i++) {
            const cell = document.createElement('div');
            cell.className = 'grid-cell';
            gridContainer.appendChild(cell);
        }
        
        boardContainer.appendChild(gridContainer);
        
        // Create tile container
        const tileContainer = document.createElement('div');
        tileContainer.className = 'tile-container';
        
        // Render tiles
        for (let row = 0; row < 5; row++) {
            for (let col = 0; col < 5; col++) {
                const value = gameState.board[row][col];
                
                if (value > 0) {
                    const tile = this.createTile(value, row, col);
                    tileContainer.appendChild(tile);
                }
            }
        }
        
        boardContainer.appendChild(tileContainer);
        this.boardElement.appendChild(boardContainer);
        
        // Update undo button state
        const undoBtn = document.getElementById('btn-undo');
        if (undoBtn) {
            undoBtn.disabled = !gameState.canUndo;
        }
    }
    
    /**
     * Create a tile element
     */
    createTile(value, row, col) {
        const tile = document.createElement('div');
        tile.className = 'game-tile tile-' + value;
        tile.textContent = value;
        
        // Position calculation for plugin grid (56px tiles + 6px gaps)
        const size = 54;
        const gap = 6;
        const left = col * (size + gap);
        const top = row * (size + gap);
        
        // Position the tile
        tile.style.left = left + 'px';
        tile.style.top = top + 'px';
        
        return tile;
    }
    
    /**
     * Update score display
     */
    updateScore(score) {
        if (this.scoreElement) {
            this.scoreElement.textContent = score;
        }
        
        // Update best score if needed
        if (score > this.bestScore) {
            this.bestScore = score;
            localStorage.setItem('2048_bestScore', score.toString());
            
            if (this.bestScoreElement) {
                this.bestScoreElement.textContent = score;
            }
        }
    }
    
    /**
     * Show message overlay
     */
    showMessage(message) {
        console.log('[Game2048Manager] Showing message:', message);
        
        // Create plugin-style message overlay
        const overlay = document.createElement('div');
        overlay.className = 'plugin-modal';
        
        const score = this.engine ? this.engine.score : 0;
        
        const playAgainText = window.app?.state?.strings?.playAgain || 'Play Again';
        const scoreText = window.app?.state?.strings?.score || 'Score';
        
        overlay.innerHTML = `
            <div class="modal-title">${message}</div>
            <div class="modal-score">${scoreText}: ${score}</div>
            <button class="btn-plugin primary" onclick="window.gameManager.startNewGame(); this.parentElement.remove();">
                ${playAgainText}
            </button>
        `;
        
        document.body.appendChild(overlay);
    }
    
    /**
     * Get current game state
     */
    getGameState() {
        return this.engine ? this.engine.getGameState() : null;
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = Game2048Manager;
}

// Also expose to global window object for browser usage
if (typeof window !== 'undefined') {
    window.Game2048Manager = Game2048Manager;
}