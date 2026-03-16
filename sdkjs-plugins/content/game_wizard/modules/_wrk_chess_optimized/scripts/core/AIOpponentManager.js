/**
 * Chess AI Opponent Manager
 * Following OnlyOffice Plugin Development Standards
 * 
 * Manages AI opponent and offline gameplay
 * Based on: _coding_standard/02_api_reference_patterns.md#single-player-patterns
 */

class ChessAIOpponentManager {
    constructor() {
        this.isInitialized = false;
        this.gameMode = 'vs_computer';
        this.aiEngine = null;
        this.aiDifficulty = 'medium';
        this.playerColor = 'white'; // Player is white by default
        this.eventListeners = new Map();
        this.isAITurn = false;
        this.aiThinkingDelay = 1000; // 1 second delay for AI moves
        this.gameManager = null;
        this.isAIEnabled = true;
    }

    /**
     * Initialize AI opponent manager
     */
    async initialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'AI Opponent Manager initializing');

        try {
            // Create AI engine
            this.aiEngine = new window.ChessAI(this.aiDifficulty);
            
            // Setup player as white by default
            this.setupGameMode();
            
            this.isInitialized = true;
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                `AI Opponent Manager initialized - Player: ${this.playerColor}, AI: ${this.getAIColor()}, Difficulty: ${this.aiDifficulty}`);

            // Notify listeners
            this.triggerEvent('initialized', {
                gameMode: this.gameMode,
                playerColor: this.playerColor,
                aiDifficulty: this.aiDifficulty
            });

            return true;

        } catch (error) {
            throw new window.ChessErrors.ChessInitializationError(
                'AI Opponent Manager initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Setup game mode configuration
     */
    setupGameMode() {
        // Player gets white pieces (goes first)
        this.playerColor = 'white';
        this.isAITurn = false; // Player moves first
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION,
            'Game mode setup: Player vs Computer');
    }

    /**
     * Handle player move - check if AI should respond
     */
    async onPlayerMove(move, gameState) {
        if (!this.isInitialized || !this.aiEngine) {
            return;
        }

        // Check if it's now AI's turn
        const currentTurn = gameState.turn;
        const aiColor = this.getAIColor();
        
        if (currentTurn === aiColor && !this.isAITurn) {
            this.isAITurn = true;
            
            // Trigger AI thinking event
            this.triggerEvent('aiThinking', {
                aiColor: aiColor,
                difficulty: this.aiDifficulty
            });

            // Delay AI move for better UX
            setTimeout(async () => {
                await this.makeAIMove(gameState);
            }, this.aiThinkingDelay);
        }
    }

    /**
     * Make AI move
     */
    async makeAIMove(gameState) {
        if (!this.aiEngine || !this.isAITurn) {
            return;
        }

        try {
            const aiColor = this.getAIColor();
            const isWhite = aiColor === 'white';
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION,
                `AI (${aiColor}) is thinking...`);

            // Get AI move
            const aiMove = await this.aiEngine.getBestMove(gameState, isWhite);
            
            if (aiMove) {
                // Notify game manager to make the move
                this.triggerEvent('aiMove', {
                    move: aiMove,
                    aiColor: aiColor,
                    notation: this.aiEngine.moveToNotation(aiMove)
                });

                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION,
                    `AI (${aiColor}) played: ${this.aiEngine.moveToNotation(aiMove)}`);
            } else {
                // AI has no moves (checkmate or stalemate)
                this.triggerEvent('aiNoMoves', {
                    aiColor: aiColor,
                    gameState: gameState
                });
            }

        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessEngineError(
                    'AI move generation failed',
                    { aiColor: this.getAIColor(), originalError: error }
                )
            );
        } finally {
            this.isAITurn = false;
        }
    }

    /**
     * Get AI color
     */
    getAIColor() {
        return this.playerColor === 'white' ? 'black' : 'white';
    }

    /**
     * Set AI difficulty
     */
    setDifficulty(difficulty) {
        this.aiDifficulty = difficulty;
        
        if (this.aiEngine) {
            this.aiEngine.setDifficulty(difficulty);
        }

        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION,
            `AI difficulty changed to: ${difficulty}`);

        this.triggerEvent('difficultyChanged', {
            difficulty: difficulty,
            depth: this.aiEngine ? this.aiEngine.maxDepth : 0
        });
    }

    /**
     * Change player color (and AI color accordingly)
     */
    setPlayerColor(color) {
        this.playerColor = color;
        this.isAITurn = (color === 'black'); // If player is black, AI (white) goes first
        
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION,
            `Player color changed to: ${color}, AI color: ${this.getAIColor()}`);

        this.triggerEvent('colorChanged', {
            playerColor: this.playerColor,
            aiColor: this.getAIColor(),
            aiGoesFirst: this.isAITurn
        });

        // If AI should go first, trigger AI move
        if (this.isAITurn && this.gameManager) {
            setTimeout(() => {
                this.makeAIMove(this.gameManager.getCurrentGameState());
            }, this.aiThinkingDelay);
        }
    }

    /**
     * Check if it's player's turn
     */
    isPlayerTurn(currentTurn) {
        return currentTurn === this.playerColor;
    }

    /**
     * Check if it's AI's turn
     */
    isAIsTurn(currentTurn) {
        return currentTurn === this.getAIColor();
    }

    /**
     * Get current game configuration
     */
    getGameConfiguration() {
        return {
            gameMode: this.gameMode,
            playerColor: this.playerColor,
            aiColor: this.getAIColor(),
            aiDifficulty: this.aiDifficulty,
            isAIThinking: this.aiEngine ? this.aiEngine.isAIThinking() : false
        };
    }

    /**
     * Get player information for display
     */
    getPlayerInfo() {
        return {
            white: {
                name: this.playerColor === 'white' ? 'You' : `Computer (${this.aiDifficulty})`,
                type: this.playerColor === 'white' ? 'human' : 'computer',
                isPlayer: this.playerColor === 'white'
            },
            black: {
                name: this.playerColor === 'black' ? 'You' : `Computer (${this.aiDifficulty})`,
                type: this.playerColor === 'black' ? 'human' : 'computer', 
                isPlayer: this.playerColor === 'black'
            }
        };
    }

    /**
     * Start new game
     */
    async startNewGame(config = {}) {
        // Reset AI state
        this.isAITurn = false;
        
        // Apply any configuration
        if (config.playerColor) {
            this.setPlayerColor(config.playerColor);
        }
        
        if (config.difficulty) {
            this.setDifficulty(config.difficulty);
        }

        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION,
            'New game started', this.getGameConfiguration());

        this.triggerEvent('newGameStarted', this.getGameConfiguration());
    }

    /**
     * Event handling
     */
    addEventListener(eventType, handler) {
        if (!this.eventListeners.has(eventType)) {
            this.eventListeners.set(eventType, []);
        }
        this.eventListeners.get(eventType).push(handler);
    }

    removeEventListener(eventType, handler) {
        if (this.eventListeners.has(eventType)) {
            const handlers = this.eventListeners.get(eventType);
            const index = handlers.indexOf(handler);
            if (index > -1) {
                handlers.splice(index, 1);
            }
        }
    }

    triggerEvent(eventType, data) {
        if (this.eventListeners.has(eventType)) {
            this.eventListeners.get(eventType).forEach(handler => {
                try {
                    handler(data);
                } catch (error) {
                    window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION,
                        `Event handler error for ${eventType}`, error);
                }
            });
        }
    }

    /**
     * Set game manager reference
     */
    setGameManager(gameManager) {
        this.gameManager = gameManager;
    }

    /**
     * Cleanup
     */
    async cleanup() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'AI Opponent Manager cleaning up');

        this.isInitialized = false;
        this.isAITurn = false;
        this.aiEngine = null;
        this.gameManager = null;
        this.eventListeners.clear();

        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'AI Opponent Manager cleanup completed');
    }
}

// Make it available globally and maintain backward compatibility
window.ChessCollaborationManager = ChessAIOpponentManager;
window.ChessAIOpponentManager = ChessAIOpponentManager;