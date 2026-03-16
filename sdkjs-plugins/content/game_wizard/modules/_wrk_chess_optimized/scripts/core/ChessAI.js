/**
 * Chess AI Engine - Offline Computer Opponent
 * Following OnlyOffice Plugin Development Standards
 * 
 * Implements simplified AI for chess moves
 */

class ChessAI {
    constructor(difficulty = 'medium') {
        this.difficulty = difficulty;
        this.maxDepth = this.getDifficultyDepth(difficulty);
        this.isThinking = false;
        
        // Map numeric piece codes to the chess engine's piece system
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
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            `Chess AI initialized with difficulty: ${difficulty} (depth: ${this.maxDepth})`);
    }

    /**
     * Get search depth based on difficulty
     */
    getDifficultyDepth(difficulty) {
        switch (difficulty) {
            case 'easy': return 1;
            case 'medium': return 2;
            case 'hard': return 3;
            case 'expert': return 4;
            default: return 2;
        }
    }

    /**
     * Get the best move for the AI
     */
    async getBestMove(gameState, isWhite = false) {
        if (this.isThinking) {
            return null;
        }

        this.isThinking = true;
        
        try {
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                `AI thinking... (${this.difficulty} difficulty, playing as ${isWhite ? 'white' : 'black'})`);

            // Get all possible moves
            const possibleMoves = this.getAllPossibleMoves(gameState, isWhite);
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                `Found ${possibleMoves.length} possible moves`);
            
            if (possibleMoves.length === 0) {
                return null; // No moves available
            }

            // For easy difficulty, sometimes make random moves
            if (this.difficulty === 'easy' && Math.random() < 0.5) {
                const randomMove = possibleMoves[Math.floor(Math.random() * possibleMoves.length)];
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                    `AI (easy) selected random move: ${this.moveToNotation(randomMove)}`);
                return randomMove;
            }

            // Evaluate all moves and pick the best one
            let bestMove = null;
            let bestScore = isWhite ? -Infinity : Infinity;

            for (const move of possibleMoves) {
                const score = this.evaluateMove(gameState.board, move);
                
                // For black (minimizing), we want the lowest score
                // For white (maximizing), we want the highest score
                if ((isWhite && score > bestScore) || (!isWhite && score < bestScore)) {
                    bestScore = score;
                    bestMove = move;
                }
            }

            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                `AI selected move: ${this.moveToNotation(bestMove)} (score: ${bestScore})`);

            return bestMove;

        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessEngineError(
                    'AI move calculation failed',
                    { originalError: error }
                )
            );
            return null;
        } finally {
            this.isThinking = false;
        }
    }

    /**
     * Get all possible moves for a player
     */
    getAllPossibleMoves(gameState, isWhite) {
        const moves = [];
        const board = gameState.board;
        
        if (!board) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'No board in game state');
            return moves;
        }
        
        for (let row = 0; row < 8; row++) {
            for (let col = 0; col < 8; col++) {
                const piece = board[row][col];
                if (piece !== this.pieces.EMPTY && this.isPieceColor(piece, isWhite)) {
                    const pieceMoves = this.getPieceValidMoves(board, row, col, piece);
                    moves.push(...pieceMoves);
                }
            }
        }

        return moves;
    }

    /**
     * Get valid moves for a specific piece
     */
    getPieceValidMoves(board, row, col, piece) {
        const moves = [];
        
        // Determine piece type
        const isWhite = piece >= this.pieces.WHITE_PAWN && piece <= this.pieces.WHITE_KING;
        const pieceType = isWhite ? piece : piece - 6; // Normalize black pieces to white piece types
        
        switch (pieceType) {
            case this.pieces.WHITE_PAWN:
                moves.push(...this.getPawnMoves(board, row, col, piece));
                break;
            case this.pieces.WHITE_ROOK:
                moves.push(...this.getRookMoves(board, row, col, piece));
                break;
            case this.pieces.WHITE_KNIGHT:
                moves.push(...this.getKnightMoves(board, row, col, piece));
                break;
            case this.pieces.WHITE_BISHOP:
                moves.push(...this.getBishopMoves(board, row, col, piece));
                break;
            case this.pieces.WHITE_QUEEN:
                moves.push(...this.getQueenMoves(board, row, col, piece));
                break;
            case this.pieces.WHITE_KING:
                moves.push(...this.getKingMoves(board, row, col, piece));
                break;
        }

        return moves;
    }

    /**
     * Get pawn moves
     */
    getPawnMoves(board, row, col, piece) {
        const moves = [];
        const isWhite = piece === this.pieces.WHITE_PAWN;
        const direction = isWhite ? -1 : 1;
        const startRow = isWhite ? 6 : 1;

        // Forward move
        const newRow = row + direction;
        if (this.isValidSquare(newRow, col) && board[newRow][col] === this.pieces.EMPTY) {
            moves.push({ from: { row, col }, to: { row: newRow, col } });
            
            // Double move from start
            if (row === startRow) {
                const doubleRow = row + 2 * direction;
                if (board[doubleRow][col] === this.pieces.EMPTY) {
                    moves.push({ from: { row, col }, to: { row: doubleRow, col } });
                }
            }
        }

        // Capture moves
        for (const captureCol of [col - 1, col + 1]) {
            if (this.isValidSquare(newRow, captureCol)) {
                const target = board[newRow][captureCol];
                if (target !== this.pieces.EMPTY && this.isPieceColor(target, !isWhite)) {
                    moves.push({ from: { row, col }, to: { row: newRow, col: captureCol } });
                }
            }
        }

        return moves;
    }

    /**
     * Get rook moves
     */
    getRookMoves(board, row, col, piece) {
        const moves = [];
        const directions = [[0, 1], [0, -1], [1, 0], [-1, 0]];
        const isWhite = this.isPieceWhite(piece);
        
        for (const [dRow, dCol] of directions) {
            for (let i = 1; i < 8; i++) {
                const newRow = row + i * dRow;
                const newCol = col + i * dCol;
                
                if (!this.isValidSquare(newRow, newCol)) break;
                
                const target = board[newRow][newCol];
                if (target === this.pieces.EMPTY) {
                    moves.push({ from: { row, col }, to: { row: newRow, col: newCol } });
                } else {
                    if (this.isPieceColor(target, !isWhite)) {
                        moves.push({ from: { row, col }, to: { row: newRow, col: newCol } });
                    }
                    break;
                }
            }
        }

        return moves;
    }

    /**
     * Get knight moves
     */
    getKnightMoves(board, row, col, piece) {
        const moves = [];
        const knightMoves = [
            [-2, -1], [-2, 1], [-1, -2], [-1, 2],
            [1, -2], [1, 2], [2, -1], [2, 1]
        ];
        const isWhite = this.isPieceWhite(piece);

        for (const [dRow, dCol] of knightMoves) {
            const newRow = row + dRow;
            const newCol = col + dCol;
            
            if (this.isValidSquare(newRow, newCol)) {
                const target = board[newRow][newCol];
                if (target === this.pieces.EMPTY || this.isPieceColor(target, !isWhite)) {
                    moves.push({ from: { row, col }, to: { row: newRow, col: newCol } });
                }
            }
        }

        return moves;
    }

    /**
     * Get bishop moves
     */
    getBishopMoves(board, row, col, piece) {
        const moves = [];
        const directions = [[1, 1], [1, -1], [-1, 1], [-1, -1]];
        const isWhite = this.isPieceWhite(piece);
        
        for (const [dRow, dCol] of directions) {
            for (let i = 1; i < 8; i++) {
                const newRow = row + i * dRow;
                const newCol = col + i * dCol;
                
                if (!this.isValidSquare(newRow, newCol)) break;
                
                const target = board[newRow][newCol];
                if (target === this.pieces.EMPTY) {
                    moves.push({ from: { row, col }, to: { row: newRow, col: newCol } });
                } else {
                    if (this.isPieceColor(target, !isWhite)) {
                        moves.push({ from: { row, col }, to: { row: newRow, col: newCol } });
                    }
                    break;
                }
            }
        }

        return moves;
    }

    /**
     * Get queen moves (combination of rook and bishop)
     */
    getQueenMoves(board, row, col, piece) {
        return [
            ...this.getRookMoves(board, row, col, piece),
            ...this.getBishopMoves(board, row, col, piece)
        ];
    }

    /**
     * Get king moves
     */
    getKingMoves(board, row, col, piece) {
        const moves = [];
        const kingMoves = [
            [-1, -1], [-1, 0], [-1, 1],
            [0, -1],           [0, 1],
            [1, -1],  [1, 0],  [1, 1]
        ];
        const isWhite = this.isPieceWhite(piece);

        for (const [dRow, dCol] of kingMoves) {
            const newRow = row + dRow;
            const newCol = col + dCol;
            
            if (this.isValidSquare(newRow, newCol)) {
                const target = board[newRow][newCol];
                if (target === this.pieces.EMPTY || this.isPieceColor(target, !isWhite)) {
                    moves.push({ from: { row, col }, to: { row: newRow, col: newCol } });
                }
            }
        }

        return moves;
    }

    /**
     * Evaluate a move (simple evaluation)
     */
    evaluateMove(board, move) {
        let score = 0;

        // Value of captured piece
        const targetPiece = board[move.to.row][move.to.col];
        if (targetPiece !== this.pieces.EMPTY) {
            score += Math.abs(this.pieceValues[targetPiece]) * 10; // Prioritize captures
        }

        // Central control bonus
        const centerDistance = Math.abs(3.5 - move.to.row) + Math.abs(3.5 - move.to.col);
        score += (7 - centerDistance) * 2; // Prefer center squares

        // Add some randomness for variety
        score += Math.random() * 5;

        return score;
    }

    /**
     * Helper functions
     */
    isValidSquare(row, col) {
        return row >= 0 && row < 8 && col >= 0 && col < 8;
    }

    isPieceWhite(piece) {
        return piece >= this.pieces.WHITE_PAWN && piece <= this.pieces.WHITE_KING;
    }

    isPieceColor(piece, isWhite) {
        return this.isPieceWhite(piece) === isWhite;
    }

    moveToNotation(move) {
        if (!move) return 'null';
        const files = 'abcdefgh';
        const ranks = '87654321';
        return files[move.from.col] + ranks[move.from.row] + '-' + 
               files[move.to.col] + ranks[move.to.row];
    }

    /**
     * Change AI difficulty
     */
    setDifficulty(difficulty) {
        this.difficulty = difficulty;
        this.maxDepth = this.getDifficultyDepth(difficulty);
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            `AI difficulty changed to: ${difficulty} (depth: ${this.maxDepth})`);
    }

    /**
     * Check if AI is currently thinking
     */
    isAIThinking() {
        return this.isThinking;
    }
}

window.ChessAI = ChessAI;