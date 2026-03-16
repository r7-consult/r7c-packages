/**
 * Simple Chess Board Renderer
 * A lightweight board renderer that works with the new architecture
 */

class SimpleBoardRenderer {
    constructor(containerElement, gameManager, aiOpponent) {
        this.container = containerElement;
        this.gameManager = gameManager;
        this.aiOpponent = aiOpponent;
        this.selectedSquare = null;
        this.possibleMoves = [];
        this.isPlayerTurn = true;
        
        // Piece symbols for display
        this.pieceSymbols = {
            0: '',   // Empty
            1: '♙',  // White pawn
            2: '♖',  // White rook
            3: '♘',  // White knight
            4: '♗',  // White bishop
            5: '♕',  // White queen
            6: '♔',  // White king
            7: '♟',  // Black pawn
            8: '♜',  // Black rook
            9: '♞',  // Black knight
            10: '♝', // Black bishop
            11: '♛', // Black queen
            12: '♚'  // Black king
        };
    }

    /**
     * Initialize the board
     */
    async initialize() {
        console.log('SimpleBoardRenderer.initialize() called');
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'SimpleBoardRenderer initializing');
        
        if (!this.container) {
            console.error('SimpleBoardRenderer: No container element!');
            window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'No container element for board');
            return;
        }
        
        console.log('Creating board in container:', this.container);
        this.createBoard();
        this.setupEventListeners();
        
        // Listen to game state changes
        if (this.gameManager) {
            this.gameManager.addEventListener('gameStateChanged', (gameState) => {
                this.updateBoard(gameState);
            });
            
            this.gameManager.addEventListener('moveAttempted', (moveData) => {
                if (!moveData.success) {
                    this.showError('Invalid move!');
                }
            });
        }
        
        // Start new game
        if (this.gameManager) {
            await this.gameManager.startNewGame();
        }
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'SimpleBoardRenderer initialized');
    }

    /**
     * Create the board HTML structure
     */
    createBoard() {
        if (!this.container) return;
        
        // Clear container
        this.container.innerHTML = '';
        
        // Create board grid
        const board = document.createElement('div');
        board.className = 'chess-grid';
        
        // Create 64 squares
        for (let row = 0; row < 8; row++) {
            for (let col = 0; col < 8; col++) {
                const square = document.createElement('div');
                square.dataset.row = row;
                square.dataset.col = col;
                square.className = (row + col) % 2 === 0
                    ? 'chess-square light'
                    : 'chess-square dark';
                
                // Add hover effect
                square.addEventListener('mouseenter', () => {
                    if (this.isPlayerTurn) {
                        square.classList.add('square-hover');
                    }
                });
                
                square.addEventListener('mouseleave', () => {
                    square.classList.remove('square-hover');
                });
                
                board.appendChild(square);
            }
        }
        
        this.container.appendChild(board);
        this.boardElement = board;
    }

    /**
     * Setup event listeners for piece movement
     */
    setupEventListeners() {
        if (!this.boardElement) return;
        
        this.boardElement.addEventListener('click', async (event) => {
            console.log('Board clicked!', event.target);
            
            const square = event.target.closest('.chess-square');
            if (!square) {
                console.log('No chess square found');
                return;
            }
            
            if (!this.isPlayerTurn) {
                console.log('Not player turn');
                return;
            }
            
            const row = parseInt(square.dataset.row);
            const col = parseInt(square.dataset.col);
            console.log(`Square clicked: (${row}, ${col})`);
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                `Square clicked at (${row},${col}), playerTurn: ${this.isPlayerTurn}`);
            
            if (this.selectedSquare) {
                // Try to move to this square
                await this.tryMove(this.selectedSquare, { row, col });
                this.clearSelection();
            } else {
                // Select this square if it has a player's piece
                this.selectSquare({ row, col });
            }
        });
    }

    /**
     * Select a square
     */
    selectSquare(position) {
        const gameState = this.gameManager?.getCurrentGameState();
        if (!gameState) return;
        
        const piece = gameState.board[position.row][position.col];
        
        // Check if it's a player's piece (white pieces: 1-6)
        if (piece >= 1 && piece <= 6) {
            this.selectedSquare = position;
            this.highlightSquare(position, '#ffff00');
            
            // Show possible moves (simplified - just highlight empty squares for now)
            this.showPossibleMoves(position, gameState.board);
        }
    }

    /**
     * Show possible moves for a piece
     */
    showPossibleMoves(from, board) {
        const moves = [];
        const gameState = this.gameManager?.getCurrentGameState();
        if (!gameState) return;
        
        // Create a move validator to check legal moves
        if (!window.ChessMoveValidator) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'ChessMoveValidator not available, using basic validation');
            return;
        }
        const validator = new window.ChessMoveValidator();
        
        // Check all possible destination squares
        for (let row = 0; row < 8; row++) {
            for (let col = 0; col < 8; col++) {
                const move = { from: from, to: { row, col } };
                
                // Check if this move is valid
                if (validator.isValidMove(board, move, gameState)) {
                    moves.push({ row, col });
                    // Highlight with different colors for captures vs normal moves
                    const target = board[row][col];
                    const color = target !== 0 ? 'rgba(255, 0, 0, 0.3)' : 'rgba(0, 255, 0, 0.3)';
                    this.highlightSquare({ row, col }, color);
                }
            }
        }
        
        this.possibleMoves = moves;
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            `Found ${moves.length} legal moves for piece at (${from.row},${from.col})`);
    }

    /**
     * Try to make a move
     */
    async tryMove(from, to) {
        if (!this.gameManager || !this.isPlayerTurn) return;
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            `Attempting move from (${from.row},${from.col}) to (${to.row},${to.col})`);
        
        try {
            this.isPlayerTurn = false; // Disable input while processing
            
            const success = await this.gameManager.makeMove({
                from: { row: from.row, col: from.col },
                to: { row: to.row, col: to.col }
            });
            
            if (success) {
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Player move successful, waiting for AI...');
                
                // AI will respond automatically through the event system
                // Re-enable input after AI moves (handled in updateBoard)
            } else {
                this.isPlayerTurn = true; // Re-enable if move failed
                this.showError('Invalid move!');
            }
        } catch (error) {
            this.isPlayerTurn = true;
            window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'Move failed', error);
            this.showError('Move failed: ' + error.message);
        }
    }

    /**
     * Clear selection and highlights
     */
    clearSelection() {
        this.selectedSquare = null;
        this.possibleMoves = [];
        
        // Clear all highlights
        const squares = this.boardElement.querySelectorAll('.chess-square');
        squares.forEach(square => {
            square.classList.remove('square-hover');
            square.style.background = '';
        });
    }

    /**
     * Highlight a square
     */
    highlightSquare(position, color) {
        const index = position.row * 8 + position.col;
        const square = this.boardElement.children[index];
        if (square) {
            square.style.background = color;
        }
    }

    /**
     * Update board display
     */
    updateBoard(gameState) {
        if (!gameState || !gameState.board) return;
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            `Updating board, turn: ${gameState.turn}`);
        
        // Update each square
        for (let row = 0; row < 8; row++) {
            for (let col = 0; col < 8; col++) {
                const index = row * 8 + col;
                const square = this.boardElement.children[index];
                if (square) {
                    const piece = gameState.board[row][col];
                    square.textContent = this.pieceSymbols[piece] || '';
                    if (piece >= 1 && piece <= 6) {
                        square.dataset.pieceColor = 'white';
                    } else if (piece >= 7 && piece <= 12) {
                        square.dataset.pieceColor = 'black';
                    } else {
                        delete square.dataset.pieceColor;
                    }
                }
            }
        }
        
        // Update turn indicator
        this.updateTurnIndicator(gameState.turn);
        
        // Re-enable player input if it's their turn
        this.isPlayerTurn = (gameState.turn === 'white');
        
        // Check for game end
        if (gameState.status === window.ChessConstants.GAME_STATE.FINISHED) {
            this.showGameEnd(gameState);
        }
    }

    /**
     * Update turn indicator
     */
    updateTurnIndicator(turn) {
        const statusText = document.getElementById('game-status-text');
        if (statusText) {
            statusText.textContent = turn === 'white' ? 'Your turn' : 'Computer thinking...';
        }
        
        // Update player highlights
        const whitePlayer = document.getElementById('white-player');
        const blackPlayer = document.getElementById('black-player');
        
        if (whitePlayer) {
            whitePlayer.classList.toggle('active', turn === 'white');
        }
        if (blackPlayer) {
            blackPlayer.classList.toggle('active', turn === 'black');
        }
    }

    /**
     * Show error message
     */
    showError(message) {
        // Simple alert for now
        console.error('Chess Error:', message);
        
        // Flash the board red briefly
        this.boardElement.style.border = '2px solid #d32f2f';
        setTimeout(() => {
            this.boardElement.style.border = '';
        }, 500);
    }

    /**
     * Show game end
     */
    showGameEnd(gameState) {
        const message = gameState.winner ? 
            `Game Over! ${gameState.winner === 'white' ? 'You' : 'Computer'} won!` :
            'Game Over! Draw!';
        
        const statusText = document.getElementById('game-status-text');
        if (statusText) {
            statusText.textContent = message;
        }
        
        this.isPlayerTurn = false;
    }

    /**
     * Cleanup
     */
    async cleanup() {
        this.selectedSquare = null;
        this.possibleMoves = [];
        this.isPlayerTurn = true;
        
        if (this.container) {
            this.container.innerHTML = '';
        }
    }
}

// Export
window.SimpleBoardRenderer = SimpleBoardRenderer;
