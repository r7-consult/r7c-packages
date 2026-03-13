/**
 * Chess Engine Core Logic
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/CODING_STANDARD.md#state-management
 */

class ChessEngine {
    constructor() {
        this.isInitialized = false;
        this.board = [];
        this.moveValidator = null;
        this.gameState = {
            turn: 'white',
            castlingRights: {
                whiteKingside: true,
                whiteQueenside: true,
                blackKingside: true,
                blackQueenside: true
            },
            enPassantTarget: null,
            halfmoveClock: 0,
            fullmoveNumber: 1,
            history: [],
            status: window.ChessConstants.GAME_STATE.WAITING
        };
        
        // Piece definitions
        this.pieces = {
            EMPTY: 0,
            WHITE_PAWN: 1,
            WHITE_ROOK: 2,
            WHITE_KNIGHT: 3,
            WHITE_BISHOP: 4,
            WHITE_QUEEN: 5,
            WHITE_KING: 6,
            BLACK_PAWN: 7,
            BLACK_ROOK: 8,
            BLACK_KNIGHT: 9,
            BLACK_BISHOP: 10,
            BLACK_QUEEN: 11,
            BLACK_KING: 12
        };

        // Piece values for evaluation
        this.pieceValues = {
            [this.pieces.WHITE_PAWN]: 100,
            [this.pieces.WHITE_ROOK]: 500,
            [this.pieces.WHITE_KNIGHT]: 300,
            [this.pieces.WHITE_BISHOP]: 300,
            [this.pieces.WHITE_QUEEN]: 900,
            [this.pieces.WHITE_KING]: 0,
            [this.pieces.BLACK_PAWN]: -100,
            [this.pieces.BLACK_ROOK]: -500,
            [this.pieces.BLACK_KNIGHT]: -300,
            [this.pieces.BLACK_BISHOP]: -300,
            [this.pieces.BLACK_QUEEN]: -900,
            [this.pieces.BLACK_KING]: 0
        };
    }

    /**
     * Initialize chess engine
     */
    async initialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'ChessEngine initializing');

        try {
            // Initialize move validator
            if (window.ChessMoveValidator) {
                this.moveValidator = new window.ChessMoveValidator();
            }
            
            this.initializeBoard();
            this.gameState.status = window.ChessConstants.GAME_STATE.WAITING;
            this.isInitialized = true;
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'ChessEngine initialized');

        } catch (error) {
            throw new window.ChessErrors.ChessInitializationError(
                'Chess engine initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Initialize board to starting position
     */
    initializeBoard() {
        // Create 8x8 board
        this.board = Array(8).fill().map(() => Array(8).fill(this.pieces.EMPTY));
        
        // Set up initial position
        // Black pieces (top of board)
        this.board[0] = [
            this.pieces.BLACK_ROOK, this.pieces.BLACK_KNIGHT, this.pieces.BLACK_BISHOP, this.pieces.BLACK_QUEEN,
            this.pieces.BLACK_KING, this.pieces.BLACK_BISHOP, this.pieces.BLACK_KNIGHT, this.pieces.BLACK_ROOK
        ];
        this.board[1] = Array(8).fill(this.pieces.BLACK_PAWN);
        
        // Empty squares (middle of board)
        for (let row = 2; row <= 5; row++) {
            this.board[row] = Array(8).fill(this.pieces.EMPTY);
        }
        
        // White pieces (bottom of board)
        this.board[6] = Array(8).fill(this.pieces.WHITE_PAWN);
        this.board[7] = [
            this.pieces.WHITE_ROOK, this.pieces.WHITE_KNIGHT, this.pieces.WHITE_BISHOP, this.pieces.WHITE_QUEEN,
            this.pieces.WHITE_KING, this.pieces.WHITE_BISHOP, this.pieces.WHITE_KNIGHT, this.pieces.WHITE_ROOK
        ];
        
        // Reset game state
        this.gameState = {
            turn: 'white',
            castlingRights: {
                whiteKingside: true,
                whiteQueenside: true,
                blackKingside: true,
                blackQueenside: true
            },
            enPassantTarget: null,
            halfmoveClock: 0,
            fullmoveNumber: 1,
            history: [],
            status: window.ChessConstants.GAME_STATE.PLAYING
        };

        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Board initialized to starting position');
    }

    /**
     * Load board from FEN string
     */
    loadFromFEN(fen) {
        try {
            const parts = fen.trim().split(' ');
            if (parts.length !== 6) {
                throw new Error('Invalid FEN string format');
            }

            // Parse board position
            this.parseFENBoard(parts[0]);
            
            // Parse game state
            this.gameState.turn = parts[1] === 'w' ? 'white' : 'black';
            this.parseFENCastling(parts[2]);
            this.gameState.enPassantTarget = parts[3] === '-' ? null : parts[3];
            this.gameState.halfmoveClock = parseInt(parts[4]);
            this.gameState.fullmoveNumber = parseInt(parts[5]);
            
            window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'Board loaded from FEN', { fen });

        } catch (error) {
            throw new window.ChessErrors.ChessValidationError(
                'Invalid FEN string',
                { fen, originalError: error }
            );
        }
    }

    /**
     * Parse FEN board section
     */
    parseFENBoard(boardFEN) {
        const pieceMap = {
            'p': this.pieces.BLACK_PAWN, 'r': this.pieces.BLACK_ROOK, 'n': this.pieces.BLACK_KNIGHT,
            'b': this.pieces.BLACK_BISHOP, 'q': this.pieces.BLACK_QUEEN, 'k': this.pieces.BLACK_KING,
            'P': this.pieces.WHITE_PAWN, 'R': this.pieces.WHITE_ROOK, 'N': this.pieces.WHITE_KNIGHT,
            'B': this.pieces.WHITE_BISHOP, 'Q': this.pieces.WHITE_QUEEN, 'K': this.pieces.WHITE_KING
        };

        this.board = [];
        const ranks = boardFEN.split('/');
        
        for (const rank of ranks) {
            const row = [];
            for (const char of rank) {
                if (char >= '1' && char <= '8') {
                    // Empty squares
                    const emptyCount = parseInt(char);
                    for (let i = 0; i < emptyCount; i++) {
                        row.push(this.pieces.EMPTY);
                    }
                } else if (pieceMap[char]) {
                    row.push(pieceMap[char]);
                }
            }
            this.board.push(row);
        }
    }

    /**
     * Parse FEN castling rights
     */
    parseFENCastling(castlingFEN) {
        this.gameState.castlingRights = {
            whiteKingside: castlingFEN.includes('K'),
            whiteQueenside: castlingFEN.includes('Q'),
            blackKingside: castlingFEN.includes('k'),
            blackQueenside: castlingFEN.includes('q')
        };
    }

    /**
     * Convert current position to FEN string
     */
    toFEN() {
        try {
            const boardFEN = this.getBoardFEN();
            const turn = this.gameState.turn === 'white' ? 'w' : 'b';
            const castling = this.getCastlingFEN();
            const enPassant = this.gameState.enPassantTarget || '-';
            const halfmove = this.gameState.halfmoveClock;
            const fullmove = this.gameState.fullmoveNumber;

            return `${boardFEN} ${turn} ${castling} ${enPassant} ${halfmove} ${fullmove}`;

        } catch (error) {
            throw new window.ChessErrors.ChessValidationError(
                'FEN generation failed',
                { gameState: this.gameState, originalError: error }
            );
        }
    }

    /**
     * Get board section of FEN
     */
    getBoardFEN() {
        const pieceMap = {
            [this.pieces.BLACK_PAWN]: 'p', [this.pieces.BLACK_ROOK]: 'r', [this.pieces.BLACK_KNIGHT]: 'n',
            [this.pieces.BLACK_BISHOP]: 'b', [this.pieces.BLACK_QUEEN]: 'q', [this.pieces.BLACK_KING]: 'k',
            [this.pieces.WHITE_PAWN]: 'P', [this.pieces.WHITE_ROOK]: 'R', [this.pieces.WHITE_KNIGHT]: 'N',
            [this.pieces.WHITE_BISHOP]: 'B', [this.pieces.WHITE_QUEEN]: 'Q', [this.pieces.WHITE_KING]: 'K'
        };

        const ranks = [];
        for (const row of this.board) {
            let rankString = '';
            let emptyCount = 0;

            for (const piece of row) {
                if (piece === this.pieces.EMPTY) {
                    emptyCount++;
                } else {
                    if (emptyCount > 0) {
                        rankString += emptyCount.toString();
                        emptyCount = 0;
                    }
                    rankString += pieceMap[piece];
                }
            }

            if (emptyCount > 0) {
                rankString += emptyCount.toString();
            }

            ranks.push(rankString);
        }

        return ranks.join('/');
    }

    /**
     * Get castling section of FEN
     */
    getCastlingFEN() {
        let castling = '';
        
        if (this.gameState.castlingRights.whiteKingside) castling += 'K';
        if (this.gameState.castlingRights.whiteQueenside) castling += 'Q';
        if (this.gameState.castlingRights.blackKingside) castling += 'k';
        if (this.gameState.castlingRights.blackQueenside) castling += 'q';
        
        return castling || '-';
    }

    /**
     * Make a move on the board
     */
    makeMove(move) {
        try {
            if (!this.isValidMove(move)) {
                throw new Error('Invalid move');
            }

            // Store move in history
            const moveRecord = {
                ...move,
                capturedPiece: this.board[move.to.row][move.to.col],
                boardStateBefore: this.getBoardCopy(),
                gameStateBefore: { ...this.gameState },
                timestamp: Date.now()
            };

            this.gameState.history.push(moveRecord);

            // Execute the move
            this.executeMoveOnBoard(move);

            // Update game state
            this.updateGameStateAfterMove(move);

            // Check for game end conditions
            this.checkGameEndConditions();

            window.ChessDebug?.logMove(move, this.gameState);

            return true;

        } catch (error) {
            throw new window.ChessErrors.ChessValidationError(
                'Move execution failed',
                { move, originalError: error }
            );
        }
    }

    /**
     * Execute move on board (low-level board manipulation)
     */
    executeMoveOnBoard(move) {
        const piece = this.board[move.from.row][move.from.col];
        
        // Move piece
        this.board[move.to.row][move.to.col] = piece;
        this.board[move.from.row][move.from.col] = this.pieces.EMPTY;

        // Handle special moves
        if (move.type === window.ChessConstants.MOVE_TYPE.CASTLE) {
            this.executeCastle(move);
        } else if (move.type === window.ChessConstants.MOVE_TYPE.EN_PASSANT) {
            this.executeEnPassant(move);
        } else if (move.type === window.ChessConstants.MOVE_TYPE.PROMOTION) {
            this.executePromotion(move);
        }
    }

    /**
     * Execute castling move
     */
    executeCastle(move) {
        const isKingside = move.to.col > move.from.col;
        const row = move.from.row;
        
        if (isKingside) {
            // Move rook from h-file to f-file
            this.board[row][5] = this.board[row][7];
            this.board[row][7] = this.pieces.EMPTY;
        } else {
            // Move rook from a-file to d-file
            this.board[row][3] = this.board[row][0];
            this.board[row][0] = this.pieces.EMPTY;
        }
    }

    /**
     * Execute en passant capture
     */
    executeEnPassant(move) {
        const capturedPawnRow = this.gameState.turn === 'white' ? move.to.row + 1 : move.to.row - 1;
        this.board[capturedPawnRow][move.to.col] = this.pieces.EMPTY;
    }

    /**
     * Execute pawn promotion
     */
    executePromotion(move) {
        const isWhite = this.gameState.turn === 'white';
        const promotionPiece = move.promoteTo || 'queen';
        
        const pieceMap = {
            queen: isWhite ? this.pieces.WHITE_QUEEN : this.pieces.BLACK_QUEEN,
            rook: isWhite ? this.pieces.WHITE_ROOK : this.pieces.BLACK_ROOK,
            bishop: isWhite ? this.pieces.WHITE_BISHOP : this.pieces.BLACK_BISHOP,
            knight: isWhite ? this.pieces.WHITE_KNIGHT : this.pieces.BLACK_KNIGHT
        };
        
        this.board[move.to.row][move.to.col] = pieceMap[promotionPiece];
    }

    /**
     * Update game state after move
     */
    updateGameStateAfterMove(move) {
        // Switch turn
        this.gameState.turn = this.gameState.turn === 'white' ? 'black' : 'white';
        
        // Update move counters
        if (this.gameState.turn === 'white') {
            this.gameState.fullmoveNumber++;
        }
        
        // Update halfmove clock
        const piece = this.board[move.to.row][move.to.col];
        const isPawnMove = piece === this.pieces.WHITE_PAWN || piece === this.pieces.BLACK_PAWN;
        const isCapture = move.capturedPiece !== this.pieces.EMPTY;
        
        if (isPawnMove || isCapture) {
            this.gameState.halfmoveClock = 0;
        } else {
            this.gameState.halfmoveClock++;
        }
        
        // Update castling rights
        this.updateCastlingRights(move);
        
        // Update en passant target
        this.updateEnPassantTarget(move);
    }

    /**
     * Update castling rights after move
     */
    updateCastlingRights(move) {
        const piece = this.board[move.to.row][move.to.col];
        
        // King moves lose all castling rights for that side
        if (piece === this.pieces.WHITE_KING) {
            this.gameState.castlingRights.whiteKingside = false;
            this.gameState.castlingRights.whiteQueenside = false;
        } else if (piece === this.pieces.BLACK_KING) {
            this.gameState.castlingRights.blackKingside = false;
            this.gameState.castlingRights.blackQueenside = false;
        }
        
        // Rook moves lose castling rights for that side
        if (move.from.row === 7 && move.from.col === 7) {
            this.gameState.castlingRights.whiteKingside = false;
        } else if (move.from.row === 7 && move.from.col === 0) {
            this.gameState.castlingRights.whiteQueenside = false;
        } else if (move.from.row === 0 && move.from.col === 7) {
            this.gameState.castlingRights.blackKingside = false;
        } else if (move.from.row === 0 && move.from.col === 0) {
            this.gameState.castlingRights.blackQueenside = false;
        }
    }

    /**
     * Update en passant target square
     */
    updateEnPassantTarget(move) {
        const piece = this.board[move.to.row][move.to.col];
        const isPawnMove = piece === this.pieces.WHITE_PAWN || piece === this.pieces.BLACK_PAWN;
        const isTwoSquareMove = Math.abs(move.to.row - move.from.row) === 2;
        
        if (isPawnMove && isTwoSquareMove) {
            const targetRow = (move.from.row + move.to.row) / 2;
            this.gameState.enPassantTarget = this.coordinateToNotation(targetRow, move.from.col);
        } else {
            this.gameState.enPassantTarget = null;
        }
    }

    /**
     * Basic move validation
     */
    isValidMove(move) {
        // Use the move validator if available
        if (this.moveValidator) {
            return this.moveValidator.isValidMove(this.board, move, this.gameState);
        }
        
        // Fallback to basic validation
        // Basic bounds checking
        if (!this.isInBounds(move.from) || !this.isInBounds(move.to)) {
            return false;
        }

        const piece = this.board[move.from.row][move.from.col];
        
        // Check if there's a piece at the from square
        if (piece === this.pieces.EMPTY) {
            return false;
        }
        
        // Check if it's the correct player's turn
        const isWhitePiece = piece >= this.pieces.WHITE_PAWN && piece <= this.pieces.WHITE_KING;
        if ((this.gameState.turn === 'white') !== isWhitePiece) {
            return false;
        }
        
        // Basic validation - just check it's not the same square
        return !(move.from.row === move.to.row && move.from.col === move.to.col);
    }

    /**
     * Check if coordinates are within board bounds
     */
    isInBounds(pos) {
        return pos.row >= 0 && pos.row < 8 && pos.col >= 0 && pos.col < 8;
    }

    /**
     * Convert coordinates to chess notation
     */
    coordinateToNotation(row, col) {
        const files = 'abcdefgh';
        const ranks = '87654321';
        return files[col] + ranks[row];
    }

    /**
     * Convert chess notation to coordinates
     */
    notationToCoordinate(notation) {
        const files = 'abcdefgh';
        const ranks = '87654321';
        
        return {
            row: ranks.indexOf(notation[1]),
            col: files.indexOf(notation[0])
        };
    }

    /**
     * Get copy of current board
     */
    getBoardCopy() {
        return this.board.map(row => [...row]);
    }

    /**
     * Get current game state
     */
    getGameState() {
        return { ...this.gameState, board: this.getBoardCopy() };
    }

    /**
     * Check for game end conditions
     */
    checkGameEndConditions() {
        // Basic implementation - can be enhanced
        if (this.gameState.halfmoveClock >= 100) {
            this.gameState.status = window.ChessConstants.GAME_STATE.FINISHED;
            return 'draw_50_moves';
        }
        
        // Additional end conditions would be checked here
        return null;
    }

    /**
     * Cleanup chess engine
     */
    async cleanup() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'ChessEngine cleanup');
        
        this.board = [];
        this.gameState = null;
        this.isInitialized = false;
    }
}

// Export chess engine
window.ChessEngine = ChessEngine;