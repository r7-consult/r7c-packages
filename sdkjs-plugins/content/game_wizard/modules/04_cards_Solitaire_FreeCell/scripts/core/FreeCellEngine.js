/**
 * FreeCell Solitaire Engine - Pure game logic
 * 
 * FreeCell Rules:
 * - Uses 1 deck (52 cards)
 * - 8 tableau columns
 * - 4 free cells (temporary storage)
 * - 4 foundation piles (build A→K by suit)
 * - All cards dealt face-up
 * - Build down alternating colors on tableau
 * - Only one card can be moved at a time (or sequences using free cells)
 * 
 * NO UI OPERATIONS - Pure logic only
 */

class FreeCellEngine {
    constructor() {
        console.log('[FreeCellEngine] Initializing engine');
        
        // Card properties
        this.SUITS = ['hearts', 'diamonds', 'clubs', 'spades'];
        this.RANKS = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
        
        // Game areas
        this.freeCells = [null, null, null, null]; // 4 free cells
        this.foundations = [[], [], [], []]; // 4 foundation piles
        this.tableau = [[], [], [], [], [], [], [], []]; // 8 columns
        
        // Game state
        this.moves = 0;
        this.score = 0;
        this.gameWon = false;
        this.previousStates = []; // For undo functionality
        
        console.log('[FreeCellEngine] Engine initialized');
    }
    
    /**
     * Create a standard deck (52 cards)
     */
    createDeck() {
        console.log('[FreeCellEngine] Creating deck');
        const deck = [];
        
        for (const suit of this.SUITS) {
            for (const rank of this.RANKS) {
                deck.push({
                    suit: suit,
                    rank: rank,
                    color: (suit === 'hearts' || suit === 'diamonds') ? 'red' : 'black',
                    faceUp: true, // All cards face-up in FreeCell
                    id: `${rank}_${suit}`
                });
            }
        }
        
        console.log(`[FreeCellEngine] Created deck with ${deck.length} cards`);
        return deck;
    }
    
    /**
     * Shuffle deck using Fisher-Yates algorithm
     */
    shuffleDeck(deck) {
        console.log('[FreeCellEngine] Shuffling deck');
        const shuffled = [...deck];
        
        for (let i = shuffled.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
        }
        
        return shuffled;
    }
    
    /**
     * Start a new game
     */
    startNewGame() {
        console.log('[FreeCellEngine] Starting new game');
        
        // Reset game state
        this.freeCells = [null, null, null, null];
        this.foundations = [[], [], [], []];
        this.tableau = [[], [], [], [], [], [], [], []];
        this.moves = 0;
        this.score = 0;
        this.gameWon = false;
        this.previousStates = [];
        
        // Create and shuffle deck
        const deck = this.shuffleDeck(this.createDeck());
        
        // Deal all cards to tableau
        // FreeCell deal pattern: first 4 columns get 7 cards, last 4 get 6 cards
        let cardIndex = 0;
        for (let row = 0; row < 7; row++) {
            for (let col = 0; col < 8; col++) {
                // First 4 columns get 7 cards, last 4 get 6
                if (row < 6 || col < 4) {
                    if (cardIndex < deck.length) {
                        this.tableau[col].push(deck[cardIndex++]);
                    }
                }
            }
        }
        
        console.log('[FreeCellEngine] Game started - All cards dealt');
        return this.getGameState();
    }
    
    /**
     * Get rank value for comparison
     */
    getRankValue(rank) {
        const values = {
            'A': 1, '2': 2, '3': 3, '4': 4, '5': 5,
            '6': 6, '7': 7, '8': 8, '9': 9, '10': 10,
            'J': 11, 'Q': 12, 'K': 13
        };
        return values[rank];
    }
    
    /**
     * Check if a card can be placed on tableau column
     */
    canPlaceOnTableau(card, targetColumn) {
        if (targetColumn.length === 0) {
            // Any card can be placed on empty column
            return true;
        }
        
        const targetCard = targetColumn[targetColumn.length - 1];
        
        // Must be opposite color and one rank lower
        const oppositeColor = card.color !== targetCard.color;
        const validRank = this.getRankValue(card.rank) === this.getRankValue(targetCard.rank) - 1;
        
        console.log(`[FreeCellEngine] Can place ${card.rank} of ${card.suit} on ${targetCard.rank} of ${targetCard.suit}? ${oppositeColor && validRank}`);
        
        return oppositeColor && validRank;
    }
    
    /**
     * Check if a card can be placed on foundation
     */
    canPlaceOnFoundation(card, foundationIndex) {
        const foundation = this.foundations[foundationIndex];
        
        if (foundation.length === 0) {
            // Only Ace can start a foundation
            return card.rank === 'A';
        }
        
        const topCard = foundation[foundation.length - 1];
        
        // Must be same suit and one rank higher
        const sameSuit = card.suit === topCard.suit;
        const validRank = this.getRankValue(card.rank) === this.getRankValue(topCard.rank) + 1;
        
        return sameSuit && validRank;
    }
    
    /**
     * Get the maximum number of cards that can be moved as a sequence.
     * If target column is empty, it cannot be used as temporary storage.
     */
    getMaxSequenceLength(targetColumnIndex = null) {
        const emptyFreeCells = this.freeCells.filter(cell => cell == null).length;
        let emptyColumns = this.tableau.filter(col => col.length === 0).length;

        if (
            Number.isInteger(targetColumnIndex) &&
            targetColumnIndex >= 0 &&
            targetColumnIndex < this.tableau.length &&
            this.tableau[targetColumnIndex].length === 0
        ) {
            emptyColumns = Math.max(0, emptyColumns - 1);
        }
        
        return (emptyFreeCells + 1) * Math.pow(2, emptyColumns);
    }
    
    /**
     * Check if a sequence of cards can be moved
     */
    canMoveSequence(cards, targetColumnIndex = null, enforceCapacityLimit = false) {
        if (cards.length === 0) return false;
        if (cards.length === 1) return true;
        
        // Check if cards form a valid sequence (alternating colors, descending)
        for (let i = 0; i < cards.length - 1; i++) {
            const card1 = cards[i];
            const card2 = cards[i + 1];
            
            // Must be opposite colors
            if (card1.color === card2.color) return false;
            
            // Must be descending by one rank
            if (this.getRankValue(card2.rank) !== this.getRankValue(card1.rank) - 1) {
                return false;
            }
        }
        
        if (!enforceCapacityLimit) {
            return true;
        }

        // Optional classic FreeCell capacity limit.
        return cards.length <= this.getMaxSequenceLength(targetColumnIndex);
    }
    
    /**
     * Move card from tableau to tableau
     */
    moveTableauToTableau(fromCol, fromIndex, toCol, moveCount = null) {
        console.log(`[FreeCellEngine] Moving from tableau ${fromCol}[${fromIndex}] to ${toCol}`);
        
        const sourceColumn = this.tableau[fromCol];
        const targetColumn = this.tableau[toCol];
        
        if (fromIndex >= sourceColumn.length) {
            console.log(`[FreeCellEngine] Invalid fromIndex: ${fromIndex}`);
            return false;
        }
        
        // Get cards to move
        const safeMoveCount = Number.isInteger(moveCount) && moveCount > 0
            ? moveCount
            : sourceColumn.length - fromIndex;
        const cardsToMove = sourceColumn.slice(fromIndex, fromIndex + safeMoveCount);
        if (cardsToMove.length === 0) {
            console.log('[FreeCellEngine] No cards selected for move');
            return false;
        }
        
        // Check if it's a valid sequence to move
        if (!this.canMoveSequence(cardsToMove, toCol)) {
            console.log(`[FreeCellEngine] Cannot move sequence: invalid order or rule mismatch`);
            return false;
        }
        
        const card = cardsToMove[0];
        
        if (!this.canPlaceOnTableau(card, targetColumn)) {
            console.log(`[FreeCellEngine] Cannot place on target column`);
            return false;
        }
        
        this.saveState();
        
        // Move the cards
        sourceColumn.splice(fromIndex, cardsToMove.length);
        targetColumn.push(...cardsToMove);
        
        this.moves++;
        this.score++;
        
        // Try auto-move to foundations
        this.tryAutoMoveToFoundations();
        
        this.checkWinCondition();
        
        console.log(`[FreeCellEngine] Move successful`);
        return true;
    }
    
    /**
     * Move card to free cell
     */
    moveToFreeCell(source, sourceIndex, freeCellIndex) {
        console.log(`[FreeCellEngine] Moving to free cell ${freeCellIndex}`);
        
        if (this.freeCells[freeCellIndex] !== null) {
            console.log(`[FreeCellEngine] Free cell ${freeCellIndex} is occupied`);
            return false;
        }
        
        let card = null;
        
        if (source === 'tableau') {
            const column = this.tableau[sourceIndex];
            if (column.length === 0) return false;
            
            // Can only move the top card to free cell
            const lastIndex = column.length - 1;
            card = column[lastIndex];
            
            this.saveState();
            column.pop();
        } else {
            return false;
        }
        
        this.freeCells[freeCellIndex] = card;
        this.moves++;
        
        console.log(`[FreeCellEngine] Moved ${card.rank} of ${card.suit} to free cell ${freeCellIndex}`);
        return true;
    }
    
    /**
     * Move card from free cell
     */
    moveFromFreeCell(freeCellIndex, target, targetIndex) {
        console.log(`[FreeCellEngine] Moving from free cell ${freeCellIndex}`);
        
        const card = this.freeCells[freeCellIndex];
        if (!card) {
            console.log(`[FreeCellEngine] Free cell ${freeCellIndex} is empty`);
            return false;
        }
        
        if (target === 'tableau') {
            const targetColumn = this.tableau[targetIndex];
            
            if (!this.canPlaceOnTableau(card, targetColumn)) {
                console.log(`[FreeCellEngine] Cannot place on tableau ${targetIndex}`);
                return false;
            }
            
            this.saveState();
            this.freeCells[freeCellIndex] = null;
            targetColumn.push(card);
            
        } else if (target === 'foundation') {
            if (!this.canPlaceOnFoundation(card, targetIndex)) {
                console.log(`[FreeCellEngine] Cannot place on foundation ${targetIndex}`);
                return false;
            }
            
            this.saveState();
            this.freeCells[freeCellIndex] = null;
            this.foundations[targetIndex].push(card);
            this.score += 10; // Bonus for foundation move
            
        } else {
            return false;
        }
        
        this.moves++;
        
        // Try auto-move to foundations
        this.tryAutoMoveToFoundations();
        
        this.checkWinCondition();
        
        console.log(`[FreeCellEngine] Move successful`);
        return true;
    }
    
    /**
     * Move card to foundation
     */
    moveToFoundation(source, sourceIndex, foundationIndex) {
        console.log(`[FreeCellEngine] Moving to foundation ${foundationIndex}`);
        
        let card = null;
        
        if (source === 'tableau') {
            const column = this.tableau[sourceIndex];
            if (column.length === 0) return false;
            
            card = column[column.length - 1];
            
            if (!this.canPlaceOnFoundation(card, foundationIndex)) {
                console.log(`[FreeCellEngine] Cannot place on foundation`);
                return false;
            }
            
            this.saveState();
            column.pop();
            
        } else if (source === 'freecell') {
            card = this.freeCells[sourceIndex];
            if (!card) return false;
            
            if (!this.canPlaceOnFoundation(card, foundationIndex)) {
                console.log(`[FreeCellEngine] Cannot place on foundation`);
                return false;
            }
            
            this.saveState();
            this.freeCells[sourceIndex] = null;
            
        } else {
            return false;
        }
        
        this.foundations[foundationIndex].push(card);
        this.moves++;
        this.score += 10; // Bonus for foundation move
        
        // Try auto-move more cards
        this.tryAutoMoveToFoundations();
        
        this.checkWinCondition();
        
        console.log(`[FreeCellEngine] Moved ${card.rank} of ${card.suit} to foundation`);
        return true;
    }
    
    /**
     * Try to auto-move cards to foundations
     */
    tryAutoMoveToFoundations() {
        console.log('[FreeCellEngine] Checking for auto-moves to foundations');
        
        let moved = true;
        while (moved) {
            moved = false;
            
            // Check tableau columns
            for (let col = 0; col < 8; col++) {
                const column = this.tableau[col];
                if (column.length > 0) {
                    const card = column[column.length - 1];
                    
                    // Try each foundation
                    for (let f = 0; f < 4; f++) {
                        if (this.canPlaceOnFoundation(card, f) && this.isSafeToAutoMove(card)) {
                            column.pop();
                            this.foundations[f].push(card);
                            this.score += 10;
                            moved = true;
                            console.log(`[FreeCellEngine] Auto-moved ${card.rank} of ${card.suit} to foundation`);
                            break;
                        }
                    }
                }
            }
            
            // Check free cells
            for (let fc = 0; fc < 4; fc++) {
                const card = this.freeCells[fc];
                if (card) {
                    // Try each foundation
                    for (let f = 0; f < 4; f++) {
                        if (this.canPlaceOnFoundation(card, f) && this.isSafeToAutoMove(card)) {
                            this.freeCells[fc] = null;
                            this.foundations[f].push(card);
                            this.score += 10;
                            moved = true;
                            console.log(`[FreeCellEngine] Auto-moved ${card.rank} of ${card.suit} from free cell to foundation`);
                            break;
                        }
                    }
                }
            }
        }
    }
    
    /**
     * Check if it's safe to auto-move a card to foundation
     */
    isSafeToAutoMove(card) {
        // A card is safe to auto-move if all cards of opposite color
        // and one rank lower are already in foundations
        
        const rankValue = this.getRankValue(card.rank);
        
        // Aces and Twos are always safe
        if (rankValue <= 2) return true;
        
        // Check foundations for opposite color cards
        const oppositeColors = card.color === 'red' ? ['clubs', 'spades'] : ['hearts', 'diamonds'];
        
        for (const suit of oppositeColors) {
            const foundationIndex = this.SUITS.indexOf(suit);
            const foundation = this.foundations[foundationIndex];
            
            // Check if the foundation has at least (rankValue - 1) cards
            if (foundation.length < rankValue - 1) {
                return false;
            }
        }
        
        return true;
    }
    
    /**
     * Check if the game is won
     */
    checkWinCondition() {
        // Game is won when all 52 cards are in foundations
        const totalInFoundations = this.foundations.reduce((sum, f) => sum + f.length, 0);
        this.gameWon = totalInFoundations === 52;
        
        if (this.gameWon) {
            console.log('[FreeCellEngine] Game Won!');
            this.score += 500; // Win bonus
        }
        
        return this.gameWon;
    }
    
    /**
     * Save current state for undo
     */
    saveState() {
        const state = {
            freeCells: this.freeCells.map(c => c ? {...c} : null),
            foundations: this.foundations.map(f => f.map(c => ({...c}))),
            tableau: this.tableau.map(col => col.map(c => ({...c}))),
            moves: this.moves,
            score: this.score
        };
        
        this.previousStates.push(JSON.stringify(state));
        
        // Keep only last 50 states to prevent memory issues
        if (this.previousStates.length > 50) {
            this.previousStates.shift();
        }
    }
    
    /**
     * Undo last move
     */
    undo() {
        if (this.previousStates.length === 0) {
            console.log('[FreeCellEngine] No moves to undo');
            return false;
        }
        
        const state = JSON.parse(this.previousStates.pop());
        
        this.freeCells = state.freeCells;
        this.foundations = state.foundations;
        this.tableau = state.tableau;
        this.moves = state.moves;
        this.score = state.score;
        
        console.log('[FreeCellEngine] Move undone');
        return true;
    }
    
    /**
     * Get current game state
     */
    getGameState() {
        return {
            freeCells: this.freeCells,
            foundations: this.foundations,
            tableau: this.tableau,
            moves: this.moves,
            score: this.score,
            gameWon: this.gameWon,
            canUndo: this.previousStates.length > 0
        };
    }
    
    /**
     * Get hint for next possible move
     */
    getHint() {
        // Check for moves to foundations first (highest priority)
        for (let col = 0; col < 8; col++) {
            const column = this.tableau[col];
            if (column.length > 0) {
                const card = column[column.length - 1];
                for (let f = 0; f < 4; f++) {
                    if (this.canPlaceOnFoundation(card, f)) {
                        return { type: 'move_to_foundation', from: 'tableau', fromIndex: col, to: f };
                    }
                }
            }
        }
        
        // Check free cells to foundations
        for (let fc = 0; fc < 4; fc++) {
            const card = this.freeCells[fc];
            if (card) {
                for (let f = 0; f < 4; f++) {
                    if (this.canPlaceOnFoundation(card, f)) {
                        return { type: 'move_to_foundation', from: 'freecell', fromIndex: fc, to: f };
                    }
                }
            }
        }
        
        // Check tableau to tableau moves
        for (let fromCol = 0; fromCol < 8; fromCol++) {
            const column = this.tableau[fromCol];
            if (column.length > 0) {
                // Try to find sequences to move
                for (let i = column.length - 1; i >= 0; i--) {
                    const sequence = column.slice(i);
                    for (let toCol = 0; toCol < 8; toCol++) {
                        if (
                            fromCol !== toCol &&
                            this.canPlaceOnTableau(sequence[0], this.tableau[toCol]) &&
                            this.canMoveSequence(sequence, toCol)
                        ) {
                                return { 
                                    type: 'move_sequence', 
                                    from: fromCol, 
                                    fromIndex: i, 
                                    to: toCol,
                                    cards: sequence.length
                                };
                        }
                    }
                }
            }
        }
        
        // Check moves to free cells
        for (let fc = 0; fc < 4; fc++) {
            if (this.freeCells[fc] === null) {
                for (let col = 0; col < 8; col++) {
                    const column = this.tableau[col];
                    if (column.length > 0) {
                        return { type: 'move_to_freecell', from: col, to: fc };
                    }
                }
            }
        }
        
        return null; // No hints available
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = FreeCellEngine;
}

// Also expose to global window object for browser usage
if (typeof window !== 'undefined') {
    window.FreeCellEngine = FreeCellEngine;
}
