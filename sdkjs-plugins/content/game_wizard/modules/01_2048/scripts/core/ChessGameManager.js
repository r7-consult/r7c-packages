/**
 * @fileoverview Chess Game Manager Implementation
 * @description Concrete implementation of chess game using GameManagerBase template method pattern
 * @see {@link _coding_standard/01_plugin_architecture_guide.md} Architecture Guide
 * @see {@link https://api.onlyoffice.com/plugin/basic} OnlyOffice Plugin API
 * @author Casual Games Plugin Development Team
 * @version 1.0.0
 * @since 1.0.0
 */

// =============================================================================
// 1. IMPORTS AND DEPENDENCIES
// =============================================================================
// GameManagerBase will be loaded before this file
// ChessEngine and related chess components will be loaded separately

// =============================================================================
// 2. CONSTANTS AND CONFIGURATION
// =============================================================================
const CHESS_CONSTANTS = Object.freeze({
    GAME_STATE: {
        NOT_STARTED: 'not_started',
        PLAYING: 'playing', 
        PAUSED: 'paused',
        FINISHED: 'finished'
    },
    END_CONDITION: {
        CHECKMATE: 'checkmate',
        STALEMATE: 'stalemate', 
        RESIGNATION: 'resignation',
        TIMEOUT: 'timeout',
        DRAW: 'draw'
    },
    STORAGE: {
        GAME_STATE: 'chess_game_state',
        PLAYER_PREFERENCES: 'chess_player_preferences'
    }
});

// =============================================================================
// 3. CHESS GAME MANAGER IMPLEMENTATION
// =============================================================================

/**
 * Chess Game Manager - Concrete implementation of GameManagerBase
 * Handles chess-specific game logic while following the abstract template
 * 
 * CRITICAL ONLYOFFICE API INTEGRATION POINTS:
 * - Uses OnlyOffice collaboration system for multiplayer games
 * - Integrates with OnlyOffice document storage for game state persistence  
 * - Follows OnlyOffice plugin lifecycle (init, button, onThemeChanged)
 * - Supports OnlyOffice theming and UI patterns
 */
class ChessGameManager extends GameManagerBase {
    #chessEngine = null;
    #aiOpponent = null;
    #players = { white: null, black: null };
    #chessSettings = Object.freeze({
        allowUndo: true,
        showPossibleMoves: true,
        autoSave: true,
        enableAI: false
    });
    
    constructor(config = {}) {
        // Call parent constructor with chess-specific configuration
        super('chess', {
            minPlayers: 1, // Allow single player with AI
            maxPlayers: 2, // Chess is always 2-player
            allowUndo: true,
            autoSave: true,
            supportsAI: true,
            supportsCollaboration: true,
            ...config
        });
        
        this.#chessSettings = Object.freeze({
            ...this.#chessSettings,
            ...config.chessSettings
        });
        
        window.debug?.debug('ChessGameManager', 'Chess game manager constructed', {
            config: this.config,
            chessSettings: this.#chessSettings
        });
    }

    // =============================================================================
    // ABSTRACT METHOD IMPLEMENTATIONS (REQUIRED BY GAMEMANAGERBASE)
    // =============================================================================
    
    /**
     * @override
     * Validate chess-specific configuration
     */
    async validateGameConfiguration() {
        window.debug?.debug('ChessGameManager', 'Validating chess configuration');
        
        // Validate chess engine availability
        if (!window.ChessEngine) {
            throw new Error('ChessEngine not available - ensure chess.js is loaded');
        }
        
        // Validate board size (always 8x8 for chess)
        if (this.config.boardSize && this.config.boardSize !== 8) {
            throw new Error('Chess requires 8x8 board size');
        }
        
        // Validate player configuration for chess
        if (this.config.minPlayers < 1 || this.config.maxPlayers !== 2) {
            throw new Error('Chess supports 1-2 players only');
        }
        
        window.debug?.info('ChessGameManager', 'Chess configuration validated');
    }
    
    /**
     * @override  
     * Initialize chess engine and related components
     */
    async initializeGameEngine() {
        window.debug?.debug('ChessGameManager', 'Initializing chess engine');
        
        try {
            // Create chess engine instance
            this.#chessEngine = new ChessEngine();
            await this.#chessEngine.initialize();
            
            // Set the engine in the base class for access
            this._setGameEngine(this.#chessEngine);
            
            // Initialize AI opponent if enabled
            if (this.#chessSettings.enableAI && window.ChessAI) {
                this.#aiOpponent = new ChessAI();
                await this.#aiOpponent.initialize();
                this.#setupAIEventHandlers();
            }
            
            // CRITICAL ONLYOFFICE INTEGRATION:
            // Chess engine should integrate with OnlyOffice API for board rendering
            // and move synchronization across collaborative sessions
            if (window.Asc && window.Asc.plugin) {
                window.debug?.info('ChessGameManager', 'OnlyOffice API available for integration');
            }
            
            window.debug?.info('ChessGameManager', 'Chess engine initialized', {
                engineReady: this.#chessEngine.isInitialized,
                aiEnabled: !!this.#aiOpponent
            });
            
        } catch (error) {
            throw new Error(`Chess engine initialization failed: ${error.message}`);
        }
    }
    
    /**
     * @override
     * Create chess-specific user interface  
     */
    async createGameInterface() {
        window.debug?.debug('ChessGameManager', 'Creating chess interface');
        
        try {
            // CRITICAL ONLYOFFICE UI INTEGRATION:
            // UI should follow OnlyOffice plugin patterns and theming
            // Use OnlyOffice's theme system for consistent appearance
            
            // Create chess board container
            const gameContainer = document.createElement('div');
            gameContainer.id = 'chess-game-container';
            gameContainer.className = 'chess-game-main';
            
            // Create chess board area
            const boardArea = document.createElement('div'); 
            boardArea.id = 'chess-board-area';
            boardArea.className = 'chess-board-container';
            
            // Create game controls area
            const controlsArea = document.createElement('div');
            controlsArea.id = 'chess-controls-area'; 
            controlsArea.className = 'chess-controls-container';
            
            // Add chess-specific controls
            this.#createChessControls(controlsArea);
            
            // Assemble interface
            gameContainer.appendChild(boardArea);
            gameContainer.appendChild(controlsArea);
            
            // Add to page
            const targetContainer = document.getElementById('game-specific-area') || document.body;
            targetContainer.appendChild(gameContainer);
            
            // Initialize board renderer if available
            if (window.ChessBoardRenderer) {
                this.boardRenderer = new ChessBoardRenderer(boardArea, {
                    showCoordinates: true,
                    enableDragDrop: true,
                    highlightMoves: this.#chessSettings.showPossibleMoves
                });
                
                await this.boardRenderer.initialize();
            }
            
            window.debug?.info('ChessGameManager', 'Chess interface created');
            
        } catch (error) {
            throw new Error(`Chess interface creation failed: ${error.message}`);
        }
    }
    
    /**
     * @override
     * Setup chess-specific event handlers
     */
    async setupEventHandlers() {
        window.debug?.debug('ChessGameManager', 'Setting up chess event handlers');
        
        // CRITICAL ONLYOFFICE EVENT INTEGRATION:
        // Events should integrate with OnlyOffice button handlers and lifecycle
        
        // Chess engine events
        if (this.#chessEngine && this.#chessEngine.addEventListener) {
            this.#chessEngine.addEventListener('moveExecuted', (moveData) => {
                this.#handleChessMoveExecuted(moveData);
            });
            
            this.#chessEngine.addEventListener('gameEnded', (endData) => {
                this.#handleChessGameEnded(endData);  
            });
            
            this.#chessEngine.addEventListener('check', (checkData) => {
                this.#handleCheck(checkData);
            });
        }
        
        // Board renderer events
        if (this.boardRenderer) {
            this.boardRenderer.addEventListener('squareClicked', (data) => {
                this.#handleSquareClick(data);
            });
            
            this.boardRenderer.addEventListener('pieceDragged', (data) => {
                this.#handlePieceDrag(data);
            });
        }
        
        // Game manager events (using base class system)
        this.addEventListener('gameStateChanged', (gameState) => {
            this.#updateChessUI(gameState);
        });
        
        window.debug?.info('ChessGameManager', 'Chess event handlers setup complete');
    }
    
    /**
     * @override  
     * Process a chess move
     */
    async makeMove(moveData) {
        window.debug?.debug('ChessGameManager', 'Processing chess move', moveData);
        
        try {
            // Validate move format for chess
            if (!this.#isValidChessMoveFormat(moveData)) {
                throw new Error('Invalid chess move format');
            }
            
            // Check if it's the current player's turn
            const gameState = this.getCurrentGameState();
            if (!this.#canPlayerMakeChessMove(gameState, moveData)) {
                throw new Error('Not your turn or invalid player');
            }
            
            // Attempt move on chess engine
            const moveResult = await this.#chessEngine.makeMove(moveData);
            
            if (moveResult.success) {
                // Update game state using base class method
                this.updateGameState({
                    lastMove: moveResult.move,
                    currentPlayer: moveResult.nextPlayer,
                    boardState: moveResult.boardState,
                    history: [...(gameState.history || []), moveResult.move]
                }, 'CHESS_MOVE');
                
                // Notify move completed
                this.notifyGameEvent('moveAttempted', {
                    success: true,
                    move: moveResult.move,
                    notation: moveResult.notation,
                    gameState: this.getCurrentGameState()
                });
                
                // Trigger AI opponent if applicable
                if (this.#aiOpponent && this.#shouldTriggerAI(moveResult.nextPlayer)) {
                    await this.#aiOpponent.onPlayerMove(moveResult.move, this.getCurrentGameState());
                }
                
                return true;
                
            } else {
                throw new Error(moveResult.error || 'Move rejected by chess engine');
            }
            
        } catch (error) {
            // Notify failed move
            this.notifyGameEvent('moveAttempted', {
                success: false,
                move: moveData,
                error: error.message
            });
            
            window.debug?.warn('ChessGameManager', 'Chess move failed', {
                move: moveData,
                error: error.message
            });
            
            throw error;
        }
    }

    // =============================================================================
    // CHESS-SPECIFIC METHODS
    // =============================================================================
    
    /**
     * Start new chess game with specific options
     */
    async startNewChessGame(gameOptions = {}) {
        window.debug?.info('ChessGameManager', 'Starting new chess game', gameOptions);
        
        // Use base class method with chess-specific setup
        const gameId = await this.startNewGame({
            ...gameOptions,
            settings: {
                ...this.#chessSettings,
                ...gameOptions.settings
            }
        });
        
        // Chess-specific initialization
        if (gameOptions.startingFEN) {
            await this.#chessEngine.loadFromFEN(gameOptions.startingFEN);
        } else {
            await this.#chessEngine.resetToStartingPosition();
        }
        
        // Setup chess players
        this.#setupChessPlayers(gameOptions.players);
        
        // Update UI
        if (this.boardRenderer) {
            await this.boardRenderer.updateBoard(this.#chessEngine.getBoardState());
        }
        
        return gameId;
    }
    
    /**
     * Load chess game from FEN notation
     */
    async loadChessGameFromFEN(fen) {
        window.debug?.debug('ChessGameManager', 'Loading chess game from FEN', { fen });
        
        try {
            await this.#chessEngine.loadFromFEN(fen);
            
            // Update game state
            this.updateGameState({
                boardState: this.#chessEngine.getBoardState(),
                currentPlayer: this.#chessEngine.getCurrentPlayer(),
                status: 'playing'
            }, 'LOAD_FROM_FEN');
            
            // Update UI
            if (this.boardRenderer) {
                await this.boardRenderer.updateBoard(this.#chessEngine.getBoardState());
            }
            
            return true;
            
        } catch (error) {
            throw new Error(`Failed to load chess game from FEN: ${error.message}`);
        }
    }
    
    /**
     * Resign chess game
     */
    async resignChessGame() {
        window.debug?.info('ChessGameManager', 'Player resigning chess game');
        
        const gameState = this.getCurrentGameState();
        const resigningPlayer = gameState.currentPlayer;
        const winner = resigningPlayer === 'white' ? 'black' : 'white';
        
        // Update game state
        this.updateGameState({
            status: CHESS_CONSTANTS.GAME_STATE.FINISHED,
            endCondition: CHESS_CONSTANTS.END_CONDITION.RESIGNATION,
            winner: winner,
            resignedPlayer: resigningPlayer,
            endTime: new Date().toISOString()
        }, 'RESIGNATION');
        
        // Notify game ended
        this.notifyGameEvent('gameEnded', {
            condition: CHESS_CONSTANTS.END_CONDITION.RESIGNATION,
            winner: winner,
            resignedPlayer: resigningPlayer
        });
    }
    
    /**
     * Undo last chess move
     */
    async undoLastChessMove() {
        window.debug?.debug('ChessGameManager', 'Undoing last chess move');
        
        if (!this.#chessSettings.allowUndo) {
            throw new Error('Undo not allowed in current game');
        }
        
        try {
            const undoResult = await this.#chessEngine.undoLastMove();
            
            if (undoResult.success) {
                this.updateGameState({
                    boardState: undoResult.boardState,
                    currentPlayer: undoResult.currentPlayer,
                    history: undoResult.history
                }, 'UNDO_MOVE');
                
                // Update UI
                if (this.boardRenderer) {
                    await this.boardRenderer.updateBoard(undoResult.boardState);
                }
                
                return true;
            } else {
                throw new Error(undoResult.error || 'Undo failed');
            }
            
        } catch (error) {
            throw new Error(`Chess move undo failed: ${error.message}`);
        }
    }

    // =============================================================================
    // CHESS AI INTEGRATION
    // =============================================================================
    
    /**
     * Setup AI opponent event handlers
     */
    #setupAIEventHandlers() {
        if (!this.#aiOpponent) return;
        
        this.#aiOpponent.addEventListener('aiMove', async (data) => {
            await this.#handleAIMove(data);
        });
        
        this.#aiOpponent.addEventListener('aiThinking', (data) => {
            this.updateGameState({
                aiThinking: true,
                aiColor: data.aiColor
            }, 'AI_THINKING');
        });
    }
    
    /**
     * Handle AI move execution
     */
    async #handleAIMove(aiMoveData) {
        window.debug?.info('ChessGameManager', 'Processing AI move', aiMoveData);
        
        try {
            // Execute AI move directly through chess engine
            const moveResult = await this.#chessEngine.makeMove(aiMoveData.move);
            
            if (moveResult.success) {
                this.updateGameState({
                    lastMove: moveResult.move,
                    currentPlayer: moveResult.nextPlayer,
                    boardState: moveResult.boardState,
                    aiThinking: false,
                    history: [...(this.getCurrentGameState().history || []), moveResult.move]
                }, 'AI_MOVE');
                
                // Update UI
                if (this.boardRenderer) {
                    await this.boardRenderer.updateBoard(moveResult.boardState);
                    this.boardRenderer.highlightMove(moveResult.move);
                }
                
                // Notify AI move completed
                this.notifyGameEvent('moveAttempted', {
                    success: true,
                    move: moveResult.move,
                    isAIMove: true,
                    aiColor: aiMoveData.aiColor,
                    notation: moveResult.notation
                });
                
            } else {
                throw new Error('AI move rejected by chess engine');
            }
            
        } catch (error) {
            window.debug?.error('ChessGameManager', 'AI move processing failed', error);
            
            // Reset AI thinking state
            this.updateGameState({
                aiThinking: false
            }, 'AI_MOVE_FAILED');
        }
    }

    // =============================================================================
    // CHESS EVENT HANDLERS
    // =============================================================================
    
    /**
     * Handle chess move executed by engine
     */
    #handleChessMoveExecuted(moveData) {
        window.debug?.debug('ChessGameManager', 'Chess move executed by engine', moveData);
        
        // This is called when engine processes moves successfully
        // Update any chess-specific state tracking here
    }
    
    /**
     * Handle chess game ended by engine
     */
    #handleChessGameEnded(endData) {
        window.debug?.info('ChessGameManager', 'Chess game ended by engine', endData);
        
        this.updateGameState({
            status: CHESS_CONSTANTS.GAME_STATE.FINISHED,
            endCondition: endData.condition,
            winner: endData.winner,
            endTime: new Date().toISOString()
        }, 'GAME_ENDED');
        
        this.notifyGameEvent('gameEnded', endData);
    }
    
    /**
     * Handle check situation
     */
    #handleCheck(checkData) {
        window.debug?.info('ChessGameManager', 'Check detected', checkData);
        
        this.updateGameState({
            inCheck: true,
            checkedPlayer: checkData.player
        }, 'CHECK');
        
        // Update UI to show check
        if (this.boardRenderer) {
            this.boardRenderer.highlightCheck(checkData.kingSquare);
        }
    }

    // =============================================================================
    // CHESS UI HELPERS
    // =============================================================================
    
    /**
     * Create chess-specific controls
     */
    #createChessControls(container) {
        // New Game button
        const newGameBtn = document.createElement('button');
        newGameBtn.textContent = 'New Game';
        newGameBtn.onclick = () => this.startNewChessGame();
        
        // Undo button
        const undoBtn = document.createElement('button');
        undoBtn.textContent = 'Undo';
        undoBtn.onclick = () => this.undoLastChessMove().catch(err => {
            window.debug?.warn('ChessGameManager', 'Undo failed', err);
        });
        
        // Resign button
        const resignBtn = document.createElement('button');
        resignBtn.textContent = 'Resign';
        resignBtn.onclick = () => this.resignChessGame();
        
        container.appendChild(newGameBtn);
        container.appendChild(undoBtn);
        container.appendChild(resignBtn);
    }
    
    /**
     * Update chess UI based on game state
     */
    #updateChessUI(gameState) {
        if (this.boardRenderer && gameState.boardState) {
            this.boardRenderer.updateBoard(gameState.boardState);
        }
        
        // Update other chess-specific UI elements
    }

    // =============================================================================
    // CHESS VALIDATION HELPERS  
    // =============================================================================
    
    /**
     * Validate chess move format
     */
    #isValidChessMoveFormat(move) {
        return move &&
               move.from && 
               typeof move.from.row === 'number' && 
               typeof move.from.col === 'number' &&
               move.to && 
               typeof move.to.row === 'number' && 
               typeof move.to.col === 'number';
    }
    
    /**
     * Check if player can make chess move
     */
    #canPlayerMakeChessMove(gameState, moveData) {
        // Add chess-specific turn validation logic here
        return gameState.status === 'playing';
    }
    
    /**
     * Setup chess players
     */
    #setupChessPlayers(playersConfig) {
        this.#players = {
            white: playersConfig?.white || { name: 'White Player', type: 'human' },
            black: playersConfig?.black || { name: 'Black Player', type: 'human' }
        };
    }
    
    /**
     * Check if AI should be triggered for next move
     */
    #shouldTriggerAI(nextPlayer) {
        return this.#aiOpponent && this.#players[nextPlayer]?.type === 'ai';
    }
    
    /**
     * Handle square click on chess board
     */
    #handleSquareClick(data) {
        // Implement chess move selection logic
        window.debug?.debug('ChessGameManager', 'Square clicked', data);
    }
    
    /**
     * Handle piece drag on chess board
     */
    #handlePieceDrag(data) {
        // Implement chess piece drag logic
        window.debug?.debug('ChessGameManager', 'Piece dragged', data);
    }
}

// =============================================================================
// 4. CHESS GAME REGISTRATION
// =============================================================================

/**
 * Register chess game type with the registry
 * This happens when the script loads
 */
document.addEventListener('DOMContentLoaded', () => {
    // Wait for registry to be available
    if (window.gameTypeRegistry) {
        window.gameTypeRegistry.registerGameType('chess', ChessGameManager, {
            name: 'Chess',
            description: 'Classic chess game with AI and collaboration support',
            category: 'strategy',
            difficulty: 'hard',
            minPlayers: 1,
            maxPlayers: 2,
            supportsAI: true,
            supportsCollaboration: true,
            requiresOnlyOfficeAPI: true,
            iconPath: '../../resources/icons/main/light/icon.png',
            estimatedLoadTime: 2000,
            memoryUsage: 'medium'
        });
        
        window.debug?.info('ChessGameManager', 'Chess game type registered successfully');
    } else {
        window.debug?.warn('ChessGameManager', 'Game registry not available for chess registration');
    }
});

// =============================================================================
// 5. EXPORTS
// =============================================================================
window.ChessGameManager = ChessGameManager;
window.CHESS_CONSTANTS = CHESS_CONSTANTS;
