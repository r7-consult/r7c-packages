/**
 * Spider Solitaire Engine - Pure game logic
 * 
 * Spider Solitaire Rules:
 * - Uses 2 decks (104 cards)
 * - 10 tableau columns
 * - Build down by suit only (K to A)
 * - Complete suits are automatically removed
 * - Stock deals 10 cards at once (one to each column)
 * - Can move sequences of same suit together
 * - Any card can be placed on empty column
 * 
 * NO UI OPERATIONS - Pure logic only
 */

class SpiderEngine {
    constructor(difficulty = 1) {
        console.log('[SpiderEngine] Initializing engine with difficulty:', difficulty);
        
        // Difficulty levels:
        // 1 = One suit (easiest - all spades)
        // 2 = Two suits (medium - spades and hearts)
        // 4 = Four suits (hardest - all suits)
        this.difficulty = difficulty;
        
        // Card suits based on difficulty
        this.SUITS = this.getSuitsByDifficulty(difficulty);
        this.RANKS = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
        
        // Game areas
        this.stock = [];        // Remaining cards to deal
        this.tableau = [[], [], [], [], [], [], [], [], [], []]; // 10 columns
        this.completedSuits = []; // Removed complete suits
        
        // Game state
        this.moves = 0;
        this.score = 500; // Starting score (decreases with moves)
        this.gameWon = false;
        this.previousStates = []; // For undo functionality
        this.dealsRemaining = 5; // Stock can be dealt 5 times
        
        console.log('[SpiderEngine] Engine initialized');
    }
    
    /**
     * Get suits based on difficulty level
     */
    getSuitsByDifficulty(difficulty) {
        switch(difficulty) {
            case 1: // One suit - all spades
                return ['spades'];
            case 2: // Two suits - spades and hearts
                return ['spades', 'hearts'];
            case 4: // Four suits - all suits
                return ['spades', 'hearts', 'clubs', 'diamonds'];
            default:
                return ['spades'];
        }
    }
    
    /**
     * Create a double deck (104 cards) for Spider
     */
    createDeck() {
        console.log('[SpiderEngine] Creating double deck');
        const deck = [];
        
        // Create 8 full suits (104 cards total)
        const suitsNeeded = 8 / this.SUITS.length; // How many of each suit
        
        for (let i = 0; i < suitsNeeded; i++) {
            for (const suit of this.SUITS) {
                for (const rank of this.RANKS) {
                    deck.push({
                        suit: suit,
                        rank: rank,
                        color: (suit === 'hearts' || suit === 'diamonds') ? 'red' : 'black',
                        faceUp: false,
                        id: `${rank}_${suit}_${i}` // Include deck number to make unique
                    });
                }
            }
        }
        
        console.log(`[SpiderEngine] Created deck with ${deck.length} cards`);
        return deck;
    }
    
    /**
     * Shuffle deck using Fisher-Yates algorithm
     */
    shuffleDeck(deck) {
        console.log('[SpiderEngine] Shuffling deck');
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
        console.log('[SpiderEngine] Starting new game');
        
        // Reset game state
        this.stock = [];
        this.tableau = [[], [], [], [], [], [], [], [], [], []];
        this.completedSuits = [];
        this.moves = 0;
        this.score = 500;
        this.gameWon = false;
        this.previousStates = [];
        this.dealsRemaining = 5;
        
        // Create and shuffle deck
        const deck = this.shuffleDeck(this.createDeck());
        
        // Deal initial tableau
        // Spider initial deal: first 4 columns get 6 cards, next 6 columns get 5 cards
        let cardIndex = 0;
        
        // First 4 columns: 6 cards each
        for (let col = 0; col < 4; col++) {
            for (let row = 0; row < 6; row++) {
                const card = deck[cardIndex++];
                card.faceUp = (row === 5); // Only last card face up
                this.tableau[col].push(card);
            }
        }
        
        // Next 6 columns: 5 cards each
        for (let col = 4; col < 10; col++) {
            for (let row = 0; row < 5; row++) {
                const card = deck[cardIndex++];
                card.faceUp = (row === 4); // Only last card face up
                this.tableau[col].push(card);
            }
        }
        
        // Remaining cards go to stock (50 cards)
        while (cardIndex < deck.length) {
            this.stock.push(deck[cardIndex++]);
        }
        
        console.log(`[SpiderEngine] Game started - Stock: ${this.stock.length}, Tableau dealt`);
        return this.getGameState();
    }
    
    /**
     * Deal from stock (10 cards, one to each column)
     */
    dealFromStock() {
        console.log('[SpiderEngine] Dealing from stock');
        
        // Check if stock has enough cards (need 10)
        if (this.stock.length < 10) {
            console.log('[SpiderEngine] Not enough cards in stock');
            return false;
        }
        
        // Check if any column is empty (can't deal with empty columns)
        for (let col = 0; col < 10; col++) {
            if (this.tableau[col].length === 0) {
                console.log('[SpiderEngine] Cannot deal with empty columns');
                return false;
            }
        }
        
        this.saveState();
        
        // Deal one card to each column
        for (let col = 0; col < 10; col++) {
            const card = this.stock.pop();
            card.faceUp = true;
            this.tableau[col].push(card);
        }
        
        this.dealsRemaining--;
        this.moves++;
        this.score = Math.max(0, this.score - 1);
        
        console.log(`[SpiderEngine] Dealt 10 cards, ${this.stock.length} remaining in stock`);
        
        // Check for completed suits after dealing
        this.checkAndRemoveCompleteSuits();
        
        return true;
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
     * Check if a card sequence can be moved (must be same suit and descending)
     */
    isValidSequence(cards) {
        if (cards.length === 0) return false;
        if (cards.length === 1) return true;
        
        const suit = cards[0].suit;
        let expectedValue = this.getRankValue(cards[0].rank);
        
        for (let i = 1; i < cards.length; i++) {
            expectedValue--;
            if (cards[i].suit !== suit || this.getRankValue(cards[i].rank) !== expectedValue) {
                return false;
            }
        }
        
        return true;
    }
    
    /**
     * Check if a card can be placed on tableau column
     */
    canPlaceOnTableau(card, targetColumn) {
        if (targetColumn.length === 0) {
            // Any card can be placed on empty column in Spider
            return true;
        }
        
        const targetCard = targetColumn[targetColumn.length - 1];
        if (!targetCard.faceUp) {
            return false;
        }
        
        // Must be one rank lower (build down)
        const validRank = this.getRankValue(card.rank) === this.getRankValue(targetCard.rank) - 1;
        
        console.log(`[SpiderEngine] Can place ${card.rank} of ${card.suit} on ${targetCard.rank} of ${targetCard.suit}? ${validRank}`);
        
        return validRank;
    }
    
    /**
     * Move card(s) between tableau columns
     */
    moveTableauToTableau(fromCol, fromIndex, toCol) {
        console.log(`[SpiderEngine] Moving from tableau ${fromCol}[${fromIndex}] to ${toCol}`);
        
        const sourceColumn = this.tableau[fromCol];
        const targetColumn = this.tableau[toCol];
        
        if (fromIndex >= sourceColumn.length) {
            console.log(`[SpiderEngine] Invalid fromIndex: ${fromIndex} >= ${sourceColumn.length}`);
            return false;
        }
        
        // Get cards to move
        const cardsToMove = sourceColumn.slice(fromIndex);
        
        // All cards must be face up
        if (!cardsToMove.every(card => card.faceUp)) {
            console.log(`[SpiderEngine] Not all cards are face up`);
            return false;
        }
        
        // Check if it's a valid sequence (same suit, descending)
        if (!this.isValidSequence(cardsToMove)) {
            console.log(`[SpiderEngine] Not a valid sequence to move`);
            return false;
        }
        
        const card = cardsToMove[0];
        console.log(`[SpiderEngine] Attempting to move ${card.rank} of ${card.suit}`);
        
        if (!this.canPlaceOnTableau(card, targetColumn)) {
            console.log(`[SpiderEngine] Move not valid according to rules`);
            return false;
        }
        
        this.saveState();
        
        // Move all cards from the selected position
        sourceColumn.splice(fromIndex);
        targetColumn.push(...cardsToMove);
        
        // Flip the new top card if needed
        if (sourceColumn.length > 0 && !sourceColumn[sourceColumn.length - 1].faceUp) {
            sourceColumn[sourceColumn.length - 1].faceUp = true;
            this.score += 5; // Bonus for revealing a card
        }
        
        this.moves++;
        this.score = Math.max(0, this.score - 1);
        
        // Check for completed suits after move
        this.checkAndRemoveCompleteSuits();
        
        this.checkWinCondition();
        
        console.log(`[SpiderEngine] Move successful`);
        return true;
    }
    
    /**
     * Check and remove complete suits (K to A of same suit)
     */
    checkAndRemoveCompleteSuits() {
        console.log('[SpiderEngine] Checking for complete suits');
        
        for (let col = 0; col < 10; col++) {
            const column = this.tableau[col];
            
            if (column.length < 13) continue; // Need at least 13 cards for a complete suit
            
            // Check if top 13 cards form a complete suit (K to A)
            const topCards = column.slice(-13);
            
            // All must be face up and same suit
            if (!topCards.every(card => card.faceUp)) continue;
            
            const suit = topCards[0].suit;
            if (!topCards.every(card => card.suit === suit)) continue;
            
            // Check if it's K through A in order
            let valid = true;
            for (let i = 0; i < 13; i++) {
                if (this.getRankValue(topCards[i].rank) !== 13 - i) {
                    valid = false;
                    break;
                }
            }
            
            if (valid) {
                console.log(`[SpiderEngine] Found complete suit of ${suit} in column ${col}`);
                
                // Remove the complete suit
                column.splice(-13);
                this.completedSuits.push({
                    suit: suit,
                    cards: topCards,
                    completedAt: new Date().toISOString()
                });
                
                // Flip new top card if needed
                if (column.length > 0 && !column[column.length - 1].faceUp) {
                    column[column.length - 1].faceUp = true;
                }
                
                // Score bonus for completing a suit
                this.score += 100;
                
                console.log(`[SpiderEngine] Removed complete suit, ${this.completedSuits.length} suits completed`);
                
                // Check win condition
                this.checkWinCondition();
                
                // Recursively check in case there's another complete suit
                this.checkAndRemoveCompleteSuits();
                return;
            }
        }
    }
    
    /**
     * Check if the game is won (8 complete suits removed)
     */
    checkWinCondition() {
        this.gameWon = this.completedSuits.length === 8;
        
        if (this.gameWon) {
            console.log('[SpiderEngine] Game Won!');
            this.score += 500; // Win bonus
        }
        
        return this.gameWon;
    }
    
    /**
     * Save current state for undo
     */
    saveState() {
        const state = {
            stock: this.stock.map(c => ({...c})),
            tableau: this.tableau.map(col => col.map(c => ({...c}))),
            completedSuits: this.completedSuits.map(s => ({...s})),
            moves: this.moves,
            score: this.score,
            dealsRemaining: this.dealsRemaining
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
            console.log('[SpiderEngine] No moves to undo');
            return false;
        }
        
        const state = JSON.parse(this.previousStates.pop());
        
        this.stock = state.stock;
        this.tableau = state.tableau;
        this.completedSuits = state.completedSuits;
        this.moves = state.moves;
        this.score = state.score;
        this.dealsRemaining = state.dealsRemaining;
        
        console.log('[SpiderEngine] Move undone');
        return true;
    }
    
    /**
     * Get current game state
     */
    getGameState() {
        return {
            stock: this.stock,
            tableau: this.tableau,
            completedSuits: this.completedSuits,
            moves: this.moves,
            score: this.score,
            gameWon: this.gameWon,
            canUndo: this.previousStates.length > 0,
            dealsRemaining: this.dealsRemaining,
            difficulty: this.difficulty
        };
    }
    
    /**
     * Get hint for next possible move
     */
    getHint() {
        // Check for moveable sequences
        for (let fromCol = 0; fromCol < 10; fromCol++) {
            const column = this.tableau[fromCol];
            
            // Find sequences in this column
            for (let i = 0; i < column.length; i++) {
                if (!column[i].faceUp) continue;
                
                // Check if cards from i to end form a valid sequence
                const sequence = column.slice(i);
                if (this.isValidSequence(sequence)) {
                    // Try to place this sequence somewhere
                    for (let toCol = 0; toCol < 10; toCol++) {
                        if (fromCol !== toCol && this.canPlaceOnTableau(sequence[0], this.tableau[toCol])) {
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
        
        // If stock has cards and we can deal
        if (this.stock.length >= 10) {
            let canDeal = true;
            for (let col = 0; col < 10; col++) {
                if (this.tableau[col].length === 0) {
                    canDeal = false;
                    break;
                }
            }
            if (canDeal) {
                return { type: 'deal_from_stock' };
            }
        }
        
        return null; // No hints available
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpiderEngine;
}

// Also expose to global window object for browser usage
if (typeof window !== 'undefined') {
    window.SpiderEngine = SpiderEngine;
}