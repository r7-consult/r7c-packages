/**
 * Game 2048 Engine - Pure game logic
 * 
 * Responsibilities:
 * - Board state management
 * - Move validation and execution
 * - Tile merging logic
 * - Score calculation
 * - Game over/win detection
 * 
 * NO UI OPERATIONS - Pure logic only
 */
class Game2048Engine {
    constructor() {
        console.log('[Game2048Engine] Initializing engine');
        
        // Fixed 5x5 board configuration
        this.SIZE = 5;
        this.board = [];
        this.score = 0;
        this.previousState = null;
        this.gameOver = false;
        this.gameWon = false;
        this.winValue = 2048;
        
        this.initializeBoard();
    }
    
    /**
     * Initialize empty board
     */
    initializeBoard() {
        console.log('[Game2048Engine] Creating 5x5 board');
        this.board = [];
        for (let i = 0; i < this.SIZE; i++) {
            this.board[i] = new Array(this.SIZE).fill(0);
        }
        this.score = 0;
        this.gameOver = false;
        this.gameWon = false;
        console.log('[Game2048Engine] Board initialized:', this.board);
    }
    
    /**
     * Start new game
     */
    startNewGame() {
        console.log('[Game2048Engine] Starting new game');
        this.initializeBoard();
        this.addRandomTile();
        this.addRandomTile();
        console.log('[Game2048Engine] Game started with 2 tiles');
        return this.getGameState();
    }
    
    /**
     * Add random tile (2 or 4)
     */
    addRandomTile() {
        const emptyCells = [];
        for (let row = 0; row < this.SIZE; row++) {
            for (let col = 0; col < this.SIZE; col++) {
                if (this.board[row][col] === 0) {
                    emptyCells.push({row, col});
                }
            }
        }
        
        if (emptyCells.length > 0) {
            const randomCell = emptyCells[Math.floor(Math.random() * emptyCells.length)];
            const value = Math.random() < 0.9 ? 2 : 4;
            this.board[randomCell.row][randomCell.col] = value;
            console.log(`[Game2048Engine] Added tile ${value} at [${randomCell.row},${randomCell.col}]`);
            return {row: randomCell.row, col: randomCell.col, value};
        }
        return null;
    }
    
    /**
     * Execute move in given direction
     */
    move(direction) {
        console.log(`[Game2048Engine] Attempting move: ${direction}`);
        
        // Save state for undo
        this.previousState = {
            board: this.board.map(row => [...row]),
            score: this.score,
            gameOver: this.gameOver,
            gameWon: this.gameWon
        };
        
        let moved = false;
        const originalBoard = this.board.map(row => [...row]);
        
        switch(direction) {
            case 'UP':
                moved = this.moveUp();
                break;
            case 'DOWN':
                moved = this.moveDown();
                break;
            case 'LEFT':
                moved = this.moveLeft();
                break;
            case 'RIGHT':
                moved = this.moveRight();
                break;
        }
        
        if (moved) {
            this.addRandomTile();
            this.checkGameState();
            console.log(`[Game2048Engine] Move successful, score: ${this.score}`);
        } else {
            // Restore previous state if no move
            this.previousState = null;
            console.log('[Game2048Engine] Move had no effect');
        }
        
        return moved;
    }
    
    /**
     * Move tiles up
     */
    moveUp() {
        let moved = false;
        for (let col = 0; col < this.SIZE; col++) {
            const column = [];
            for (let row = 0; row < this.SIZE; row++) {
                if (this.board[row][col] !== 0) {
                    column.push(this.board[row][col]);
                }
            }
            
            const merged = this.mergeTiles(column);
            
            for (let row = 0; row < this.SIZE; row++) {
                const newValue = row < merged.length ? merged[row] : 0;
                if (this.board[row][col] !== newValue) {
                    moved = true;
                    this.board[row][col] = newValue;
                }
            }
        }
        return moved;
    }
    
    /**
     * Move tiles down
     */
    moveDown() {
        let moved = false;
        for (let col = 0; col < this.SIZE; col++) {
            const column = [];
            for (let row = this.SIZE - 1; row >= 0; row--) {
                if (this.board[row][col] !== 0) {
                    column.push(this.board[row][col]);
                }
            }
            
            const merged = this.mergeTiles(column);
            
            for (let row = 0; row < this.SIZE; row++) {
                const newValue = row < merged.length ? merged[row] : 0;
                const boardRow = this.SIZE - 1 - row;
                if (this.board[boardRow][col] !== newValue) {
                    moved = true;
                    this.board[boardRow][col] = newValue;
                }
            }
        }
        return moved;
    }
    
    /**
     * Move tiles left
     */
    moveLeft() {
        let moved = false;
        for (let row = 0; row < this.SIZE; row++) {
            const line = [];
            for (let col = 0; col < this.SIZE; col++) {
                if (this.board[row][col] !== 0) {
                    line.push(this.board[row][col]);
                }
            }
            
            const merged = this.mergeTiles(line);
            
            for (let col = 0; col < this.SIZE; col++) {
                const newValue = col < merged.length ? merged[col] : 0;
                if (this.board[row][col] !== newValue) {
                    moved = true;
                    this.board[row][col] = newValue;
                }
            }
        }
        return moved;
    }
    
    /**
     * Move tiles right
     */
    moveRight() {
        let moved = false;
        for (let row = 0; row < this.SIZE; row++) {
            const line = [];
            for (let col = this.SIZE - 1; col >= 0; col--) {
                if (this.board[row][col] !== 0) {
                    line.push(this.board[row][col]);
                }
            }
            
            const merged = this.mergeTiles(line);
            
            for (let col = 0; col < this.SIZE; col++) {
                const newValue = col < merged.length ? merged[col] : 0;
                const boardCol = this.SIZE - 1 - col;
                if (this.board[row][boardCol] !== newValue) {
                    moved = true;
                    this.board[row][boardCol] = newValue;
                }
            }
        }
        return moved;
    }
    
    /**
     * Merge tiles in a line
     */
    mergeTiles(line) {
        const result = [];
        let i = 0;
        
        while (i < line.length) {
            if (i < line.length - 1 && line[i] === line[i + 1]) {
                // Merge tiles
                const mergedValue = line[i] * 2;
                result.push(mergedValue);
                this.score += mergedValue;
                
                // Check for win condition
                if (mergedValue >= this.winValue && !this.gameWon) {
                    this.gameWon = true;
                    console.log('[Game2048Engine] WIN! Reached 2048!');
                }
                
                i += 2;
            } else {
                result.push(line[i]);
                i++;
            }
        }
        
        return result;
    }
    
    /**
     * Check if game is over
     */
    checkGameState() {
        // Check for empty cells
        for (let row = 0; row < this.SIZE; row++) {
            for (let col = 0; col < this.SIZE; col++) {
                if (this.board[row][col] === 0) {
                    return; // Game continues
                }
            }
        }
        
        // Check for possible merges
        for (let row = 0; row < this.SIZE; row++) {
            for (let col = 0; col < this.SIZE; col++) {
                const current = this.board[row][col];
                
                // Check right neighbor
                if (col < this.SIZE - 1 && this.board[row][col + 1] === current) {
                    return; // Merge possible
                }
                
                // Check bottom neighbor
                if (row < this.SIZE - 1 && this.board[row + 1][col] === current) {
                    return; // Merge possible
                }
            }
        }
        
        // No moves possible
        this.gameOver = true;
        console.log('[Game2048Engine] GAME OVER! No moves remaining');
    }
    
    /**
     * Undo last move
     */
    undo() {
        if (this.previousState) {
            console.log('[Game2048Engine] Undoing last move');
            this.board = this.previousState.board;
            this.score = this.previousState.score;
            this.gameOver = this.previousState.gameOver;
            this.gameWon = this.previousState.gameWon;
            this.previousState = null;
            return true;
        }
        console.log('[Game2048Engine] No move to undo');
        return false;
    }
    
    /**
     * Get current game state
     */
    getGameState() {
        return {
            board: this.board.map(row => [...row]),
            score: this.score,
            gameOver: this.gameOver,
            gameWon: this.gameWon,
            canUndo: this.previousState !== null
        };
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = Game2048Engine;
}

// Also expose to global window object for browser usage
if (typeof window !== 'undefined') {
    window.Game2048Engine = Game2048Engine;
}