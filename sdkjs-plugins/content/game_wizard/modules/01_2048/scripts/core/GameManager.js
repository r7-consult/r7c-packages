/**
 * Chess Game Manager
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/01_plugin_architecture_guide.md#advanced-patterns
 */

class ChessGameManager {
    constructor(chessEngine) {
        this.chessEngine = chessEngine;
        this.isInitialized = false;
        this.eventListeners = new Map();
        this.currentGameId = null;
        this.aiOpponent = null; // AI opponent manager
        this.players = {
            white: null,
            black: null
        };
        this.gameSettings = {
            allowUndo: true,
            showPossibleMoves: true,
            autoSave: true
        };
    }

    /**
     * Initialize game manager
     */
    async initialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'GameManager initializing');

        try {
            if (!this.chessEngine || !this.chessEngine.isInitialized) {
                throw new Error('Chess engine not available or not initialized');
            }

            // Setup event listeners
            this.setupEventListeners();
            
            // Load game settings
            await this.loadGameSettings();
            
            this.isInitialized = true;
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'GameManager initialized');

        } catch (error) {
            throw new window.ChessErrors.ChessInitializationError(
                'Game manager initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Set AI opponent manager
     */
    setAIOpponent(aiOpponent) {
        this.aiOpponent = aiOpponent;
        
        if (this.aiOpponent) {
            // Set up AI event listeners
            this.aiOpponent.addEventListener('aiMove', (data) => {
                this.handleAIMove(data);
            });
            
            this.aiOpponent.addEventListener('aiThinking', (data) => {
                this.notifyGameStateChanged({
                    ...this.getCurrentGameState(),
                    aiThinking: true,
                    currentPlayer: data.aiColor
                });
            });
            
            // Give AI opponent access to game manager
            this.aiOpponent.setGameManager(this);
        }
    }

    /**
     * Setup internal event listeners
     */
    setupEventListeners() {
        // Listen for chess engine events if supported
        if (this.chessEngine.addEventListener) {
            this.chessEngine.addEventListener('moveExecuted', (moveData) => {
                this.handleMoveExecuted(moveData);
            });
            
            this.chessEngine.addEventListener('gameEnded', (endData) => {
                this.handleGameEnded(endData);
            });
        }
    }

    /**
     * Start a new chess game
     */
    async startNewGame(gameOptions = {}) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Starting new chess game', gameOptions);

        try {
            // Generate new game ID
            this.currentGameId = this.generateGameId();
            
            // Initialize chess engine with starting position
            await this.chessEngine.initialize();
            
            // Setup players
            this.setupPlayers(gameOptions.players);
            
            // Load custom starting position if provided
            if (gameOptions.startingFEN) {
                this.chessEngine.loadFromFEN(gameOptions.startingFEN);
            }
            
            // Apply game settings
            this.applyGameSettings(gameOptions.settings);
            
            // Mark game as active
            const gameState = this.chessEngine.getGameState();
            gameState.status = window.ChessConstants.GAME_STATE.PLAYING;
            
            // Notify listeners
            this.notifyGameStateChanged(gameState);
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'New chess game started', { gameId: this.currentGameId });

            return this.currentGameId;

        } catch (error) {
            throw new window.ChessErrors.ChessInitializationError(
                'Failed to start new game',
                { gameOptions, originalError: error }
            );
        }
    }

    /**
     * Load existing game state
     */
    async loadGameState(gameData) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Loading game state', { hasGameData: !!gameData });

        try {
            if (!gameData || !gameData.fen) {
                throw new Error('Invalid game data');
            }

            // Load board position
            this.chessEngine.loadFromFEN(gameData.fen);
            
            // Restore game metadata
            if (gameData.gameId) {
                this.currentGameId = gameData.gameId;
            }
            
            if (gameData.players) {
                this.players = { ...gameData.players };
            }
            
            if (gameData.settings) {
                this.gameSettings = { ...this.gameSettings, ...gameData.settings };
            }
            
            // Notify listeners
            const gameState = this.chessEngine.getGameState();
            this.notifyGameStateChanged(gameState);
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'Game state loaded successfully');

        } catch (error) {
            throw new window.ChessErrors.ChessPersistenceError(
                'Failed to load game state',
                { gameData, originalError: error }
            );
        }
    }

    /**
     * Make a move in the game
     */
    async makeMove(move) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Attempting move', move);

        try {
            // Validate move format
            if (!this.isValidMoveFormat(move)) {
                throw new Error('Invalid move format');
            }

            // Check if it's the current player's turn
            const gameState = this.chessEngine.getGameState();
            if (!this.canPlayerMove(gameState.turn)) {
                throw new Error('Not your turn');
            }

            // Attempt move on chess engine
            const success = await this.chessEngine.makeMove(move);
            
            if (success) {
                // Auto-save if enabled
                if (this.gameSettings.autoSave) {
                    await this.autoSave();
                }
                
                // Notify successful move
                this.notifyMoveAttempted({
                    success: true,
                    move,
                    gameState: this.chessEngine.getGameState()
                });
                
                // Update game state
                const newGameState = this.chessEngine.getGameState();
                this.notifyGameStateChanged(newGameState);
                
                // Notify AI opponent about the move
                if (this.aiOpponent) {
                    await this.aiOpponent.onPlayerMove(move, newGameState);
                }
                
                return true;
            } else {
                throw new Error('Move rejected by chess engine');
            }

        } catch (error) {
            // Notify failed move
            this.notifyMoveAttempted({
                success: false,
                move,
                error: error.message
            });
            
            throw new window.ChessErrors.ChessValidationError(
                'Move failed',
                { move, originalError: error }
            );
        }
    }

    /**
     * Handle AI move
     */
    async handleAIMove(data) {
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Processing AI move', data);

        try {
            // Make AI move directly on chess engine (bypass player validation)
            const success = await this.chessEngine.makeMove(data.move);
            
            if (success) {
                // Auto-save if enabled
                if (this.gameSettings.autoSave) {
                    await this.autoSave();
                }
                
                // Notify AI move
                this.notifyMoveAttempted({
                    success: true,
                    move: data.move,
                    isAIMove: true,
                    aiColor: data.aiColor,
                    notation: data.notation,
                    gameState: this.chessEngine.getGameState()
                });
                
                // Update game state
                const newGameState = this.chessEngine.getGameState();
                this.notifyGameStateChanged({
                    ...newGameState,
                    aiThinking: false,
                    lastMove: {
                        ...data.move,
                        isAI: true,
                        notation: data.notation
                    }
                });
                
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                    `AI move completed: ${data.notation}`);
                
            } else {
                window.ChessErrorHandler?.handleError(
                    new window.ChessErrors.ChessEngineError(
                        'AI move was rejected by chess engine',
                        { aiMove: data.move, aiColor: data.aiColor }
                    )
                );
            }

        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessEngineError(
                    'AI move processing failed',
                    { aiMove: data.move, aiColor: data.aiColor, originalError: error }
                )
            );
        }
    }

    /**
     * Undo last move
     */
    async undoLastMove() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Attempting to undo last move');

        try {
            if (!this.gameSettings.allowUndo) {
                throw new Error('Undo not allowed in current game');
            }

            const gameState = this.chessEngine.getGameState();
            const history = gameState.history;
            
            if (!history || history.length === 0) {
                throw new Error('No moves to undo');
            }

            // Get last move
            const lastMove = history[history.length - 1];
            
            // Restore previous board state
            if (lastMove.boardStateBefore && lastMove.gameStateBefore) {
                this.chessEngine.board = lastMove.boardStateBefore;
                this.chessEngine.gameState = { ...lastMove.gameStateBefore };
                
                // Remove the undone move from history
                this.chessEngine.gameState.history = history.slice(0, -1);
                
                // Notify state change
                this.notifyGameStateChanged(this.chessEngine.getGameState());
                
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                    'Move undone successfully', { undoneMoveNotation: lastMove.notation });
                
                return true;
            } else {
                throw new Error('Cannot restore previous game state');
            }

        } catch (error) {
            throw new window.ChessErrors.ChessValidationError(
                'Undo failed',
                { originalError: error }
            );
        }
    }

    /**
     * Resign current game
     */
    async resignGame() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Player resigning');

        try {
            const gameState = this.chessEngine.getGameState();
            const resigningPlayer = gameState.turn;
            const winner = resigningPlayer === 'white' ? 'black' : 'white';
            
            // Update game state
            gameState.status = window.ChessConstants.GAME_STATE.FINISHED;
            gameState.endCondition = window.ChessConstants.END_CONDITION.RESIGNATION;
            gameState.winner = winner;
            gameState.endTime = new Date().toISOString();
            
            // Notify game ended
            this.notifyGameEnded({
                condition: gameState.endCondition,
                winner: winner,
                resignedPlayer: resigningPlayer
            });
            
            // Update state
            this.notifyGameStateChanged(gameState);
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'Game resigned', { winner, resignedPlayer: resigningPlayer });

        } catch (error) {
            throw new window.ChessErrors.ChessValidationError(
                'Resignation failed',
                { originalError: error }
            );
        }
    }

    /**
     * Get current game state
     */
    getCurrentGameState() {
        if (!this.chessEngine || !this.chessEngine.isInitialized) {
            return null;
        }
        
        const engineState = this.chessEngine.getGameState();
        
        return {
            ...engineState,
            gameId: this.currentGameId,
            players: this.players,
            settings: this.gameSettings
        };
    }

    /**
     * Check if game has unsaved changes
     */
    hasUnsavedChanges() {
        const gameState = this.getCurrentGameState();
        if (!gameState || !gameState.history) {
            return false;
        }
        
        // Check if there are moves since last save
        return gameState.history.length > 0 && !this.lastSaveState;
    }

    /**
     * Auto-save game state
     */
    async autoSave() {
        try {
            const gameState = this.getCurrentGameState();
            if (!gameState) return;
            
            const saveData = {
                gameId: this.currentGameId,
                fen: this.chessEngine.toFEN(),
                players: this.players,
                settings: this.gameSettings,
                lastSaved: new Date().toISOString()
            };
            
            // Save to localStorage
            const storageKey = `${window.ChessConstants.STORAGE.GAME_STATE}_${this.currentGameId}`;
            localStorage.setItem(storageKey, JSON.stringify(saveData));
            
            this.lastSaveState = { ...saveData };
            
            window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'Game auto-saved', { gameId: this.currentGameId });

        } catch (error) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'Auto-save failed', error);
        }
    }

    /**
     * Handle player joining
     */
    onPlayerJoined(playerData) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Player joined', playerData);

        if (playerData.color && ['white', 'black'].includes(playerData.color)) {
            this.players[playerData.color] = {
                name: playerData.name || 'Anonymous',
                id: playerData.id,
                joinedAt: new Date().toISOString()
            };
            
            // Notify state change
            this.notifyGameStateChanged(this.getCurrentGameState());
        }
    }

    /**
     * Handle received move from collaboration
     */
    onMoveReceived(moveData) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Move received from collaboration', moveData);

        // Validate and apply the move
        this.makeMove(moveData.move).catch(error => {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessCollaborationError(
                    'Collaborative move failed',
                    { moveData, originalError: error }
                )
            );
        });
    }

    /**
     * Validate move format
     */
    isValidMoveFormat(move) {
        return move &&
               move.from && typeof move.from.row === 'number' && typeof move.from.col === 'number' &&
               move.to && typeof move.to.row === 'number' && typeof move.to.col === 'number';
    }

    /**
     * Check if player can make a move
     */
    canPlayerMove(turn) {
        // In offline AI mode, only allow player moves on player's turn
        if (this.aiOpponent) {
            return this.aiOpponent.isPlayerTurn(turn);
        }
        
        // Default: allow all moves
        return true;
    }

    /**
     * Setup players
     */
    setupPlayers(playersConfig = {}) {
        this.players = {
            white: playersConfig.white || { name: 'Player 1', type: 'human' },
            black: playersConfig.black || { name: 'Player 2', type: 'human' }
        };
    }

    /**
     * Apply game settings
     */
    applyGameSettings(settings = {}) {
        this.gameSettings = {
            ...this.gameSettings,
            ...settings
        };
    }

    /**
     * Load game settings from storage
     */
    async loadGameSettings() {
        try {
            const stored = localStorage.getItem(window.ChessConstants.STORAGE.PLAYER_PREFERENCES);
            if (stored) {
                const preferences = JSON.parse(stored);
                if (preferences.gameSettings) {
                    this.gameSettings = { ...this.gameSettings, ...preferences.gameSettings };
                }
            }
        } catch (error) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                'Failed to load game settings', error);
        }
    }

    /**
     * Generate unique game ID
     */
    generateGameId() {
        return 'chess_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    /**
     * Event listener management
     */
    addEventListener(event, callback) {
        if (!this.eventListeners.has(event)) {
            this.eventListeners.set(event, []);
        }
        this.eventListeners.get(event).push(callback);
    }

    removeEventListener(event, callback) {
        if (this.eventListeners.has(event)) {
            const listeners = this.eventListeners.get(event);
            const index = listeners.indexOf(callback);
            if (index > -1) {
                listeners.splice(index, 1);
            }
        }
    }

    /**
     * Notify game state changed
     */
    notifyGameStateChanged(gameState) {
        const listeners = this.eventListeners.get('gameStateChanged') || [];
        listeners.forEach(callback => {
            try {
                callback(gameState);
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                    'Game state listener error', error);
            }
        });
    }

    /**
     * Notify move attempted
     */
    notifyMoveAttempted(moveData) {
        const listeners = this.eventListeners.get('moveAttempted') || [];
        listeners.forEach(callback => {
            try {
                callback(moveData);
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                    'Move attempt listener error', error);
            }
        });
    }

    /**
     * Notify game ended
     */
    notifyGameEnded(endData) {
        const listeners = this.eventListeners.get('gameEnded') || [];
        listeners.forEach(callback => {
            try {
                callback(endData);
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
                    'Game end listener error', error);
            }
        });
    }

    /**
     * Handle move executed by engine
     */
    handleMoveExecuted(moveData) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Move executed by engine', moveData);
        
        // Auto-save if enabled
        if (this.gameSettings.autoSave) {
            this.autoSave();
        }
    }

    /**
     * Handle game ended by engine
     */
    handleGameEnded(endData) {
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'Game ended', endData);
        
        this.notifyGameEnded(endData);
    }

    /**
     * Cleanup game manager
     */
    async cleanup() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.CHESS_ENGINE, 
            'GameManager cleanup');

        // Auto-save current state
        if (this.hasUnsavedChanges()) {
            await this.autoSave();
        }
        
        // Clear event listeners
        this.eventListeners.clear();
        
        // Reset state
        this.currentGameId = null;
        this.players = { white: null, black: null };
        this.isInitialized = false;
    }
}

// Export game manager
window.ChessGameManager = ChessGameManager;