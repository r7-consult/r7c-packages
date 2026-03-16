/**
 * Chess Move Validator
 * Validates chess moves according to standard rules
 */

class ChessMoveValidator {
    constructor() {
        // Piece constants matching ChessEngine
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
    }

    /**
     * Validate if a move is legal
     */
    isValidMove(board, move, gameState) {
        // Basic bounds checking
        if (!this.isInBounds(move.from) || !this.isInBounds(move.to)) {
            return false;
        }

        // Same square check
        if (move.from.row === move.to.row && move.from.col === move.to.col) {
            return false;
        }

        const piece = board[move.from.row][move.from.col];
        const targetPiece = board[move.to.row][move.to.col];
        
        // Check if there's a piece at the from square
        if (piece === this.pieces.EMPTY) {
            return false;
        }
        
        // Check if it's the correct player's turn
        const isWhitePiece = piece >= this.pieces.WHITE_PAWN && piece <= this.pieces.WHITE_KING;
        if ((gameState.turn === 'white') !== isWhitePiece) {
            return false;
        }
        
        // Can't capture own piece
        if (targetPiece !== this.pieces.EMPTY) {
            const isTargetWhite = targetPiece >= this.pieces.WHITE_PAWN && targetPiece <= this.pieces.WHITE_KING;
            if (isWhitePiece === isTargetWhite) {
                return false;
            }
        }
        
        // Check piece-specific movement rules
        switch (piece) {
            case this.pieces.WHITE_PAWN:
            case this.pieces.BLACK_PAWN:
                return this.isValidPawnMove(board, move, piece);
            
            case this.pieces.WHITE_ROOK:
            case this.pieces.BLACK_ROOK:
                return this.isValidRookMove(board, move);
            
            case this.pieces.WHITE_KNIGHT:
            case this.pieces.BLACK_KNIGHT:
                return this.isValidKnightMove(move);
            
            case this.pieces.WHITE_BISHOP:
            case this.pieces.BLACK_BISHOP:
                return this.isValidBishopMove(board, move);
            
            case this.pieces.WHITE_QUEEN:
            case this.pieces.BLACK_QUEEN:
                return this.isValidQueenMove(board, move);
            
            case this.pieces.WHITE_KING:
            case this.pieces.BLACK_KING:
                return this.isValidKingMove(move);
            
            default:
                return false;
        }
    }

    /**
     * Validate pawn move
     */
    isValidPawnMove(board, move, piece) {
        const isWhite = piece === this.pieces.WHITE_PAWN;
        const direction = isWhite ? -1 : 1;
        const startRow = isWhite ? 6 : 1;
        
        const rowDiff = move.to.row - move.from.row;
        const colDiff = Math.abs(move.to.col - move.from.col);
        const targetPiece = board[move.to.row][move.to.col];
        
        // Forward move
        if (colDiff === 0) {
            // Single step forward
            if (rowDiff === direction && targetPiece === this.pieces.EMPTY) {
                return true;
            }
            // Double step from start
            if (move.from.row === startRow && rowDiff === 2 * direction && 
                targetPiece === this.pieces.EMPTY &&
                board[move.from.row + direction][move.from.col] === this.pieces.EMPTY) {
                return true;
            }
        }
        // Diagonal capture
        else if (colDiff === 1 && rowDiff === direction) {
            if (targetPiece !== this.pieces.EMPTY) {
                return true;
            }
        }
        
        return false;
    }

    /**
     * Validate rook move
     */
    isValidRookMove(board, move) {
        // Rook moves horizontally or vertically
        if (move.from.row !== move.to.row && move.from.col !== move.to.col) {
            return false;
        }
        
        // Check path is clear
        return this.isPathClear(board, move);
    }

    /**
     * Validate knight move
     */
    isValidKnightMove(move) {
        const rowDiff = Math.abs(move.to.row - move.from.row);
        const colDiff = Math.abs(move.to.col - move.from.col);
        
        // Knight moves in L shape
        return (rowDiff === 2 && colDiff === 1) || (rowDiff === 1 && colDiff === 2);
    }

    /**
     * Validate bishop move
     */
    isValidBishopMove(board, move) {
        const rowDiff = Math.abs(move.to.row - move.from.row);
        const colDiff = Math.abs(move.to.col - move.from.col);
        
        // Bishop moves diagonally
        if (rowDiff !== colDiff) {
            return false;
        }
        
        // Check path is clear
        return this.isPathClear(board, move);
    }

    /**
     * Validate queen move
     */
    isValidQueenMove(board, move) {
        // Queen moves like rook or bishop
        return this.isValidRookMove(board, move) || this.isValidBishopMove(board, move);
    }

    /**
     * Validate king move
     */
    isValidKingMove(move) {
        const rowDiff = Math.abs(move.to.row - move.from.row);
        const colDiff = Math.abs(move.to.col - move.from.col);
        
        // King moves one square in any direction
        return rowDiff <= 1 && colDiff <= 1;
    }

    /**
     * Check if path is clear for sliding pieces
     */
    isPathClear(board, move) {
        const rowStep = move.to.row > move.from.row ? 1 : move.to.row < move.from.row ? -1 : 0;
        const colStep = move.to.col > move.from.col ? 1 : move.to.col < move.from.col ? -1 : 0;
        
        let currentRow = move.from.row + rowStep;
        let currentCol = move.from.col + colStep;
        
        while (currentRow !== move.to.row || currentCol !== move.to.col) {
            if (board[currentRow][currentCol] !== this.pieces.EMPTY) {
                return false;
            }
            currentRow += rowStep;
            currentCol += colStep;
        }
        
        return true;
    }

    /**
     * Check if coordinates are within board bounds
     */
    isInBounds(pos) {
        return pos.row >= 0 && pos.row < 8 && pos.col >= 0 && pos.col < 8;
    }
}

// Export
window.ChessMoveValidator = ChessMoveValidator;