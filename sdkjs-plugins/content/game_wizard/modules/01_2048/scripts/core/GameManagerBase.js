/**
 * @fileoverview Abstract Game Manager Base Class
 * @description Provides common game management functionality for casual games using Template Method Pattern
 * @see {@link _coding_standard/01_plugin_architecture_guide.md} Architecture Guide
 * @see {@link https://api.onlyoffice.com/plugin/basic} OnlyOffice Plugin API
 * @author Casual Games Plugin Development Team
 * @version 1.0.0
 * @since 1.0.0
 */

// =============================================================================
// 1. IMPORTS AND DEPENDENCIES  
// =============================================================================
// OnlyOffice API integration will be handled by concrete implementations

// =============================================================================
// 2. CONSTANTS AND CONFIGURATION
// =============================================================================
const DEFAULT_GAME_STATE = Object.freeze({
    initialized: false,
    status: 'not_started', // not_started, playing, paused, finished
    currentPlayer: null,
    players: [],
    settings: {},
    history: []
});

const GAME_EVENTS = Object.freeze({
    STATE_CHANGED: 'gameStateChanged',
    MOVE_ATTEMPTED: 'moveAttempted', 
    GAME_STARTED: 'gameStarted',
    GAME_ENDED: 'gameEnded',
    PLAYER_JOINED: 'playerJoined'
});

// =============================================================================
// 3. ABSTRACT BASE CLASS DEFINITION
// =============================================================================

/**
 * Abstract base class defining game management algorithm
 * Derived classes implement specific game logic
 * 
 * CRITICAL ONLYOFFICE INTEGRATION NOTES:
 * - Concrete implementations must handle OnlyOffice API calls
 * - Event system must integrate with OnlyOffice plugin lifecycle  
 * - Auto-save should use OnlyOffice document storage when available
 * - Collaboration features should leverage OnlyOffice's real-time sync
 */
class GameManagerBase {
    #initialized = false;
    #gameType = 'unknown';
    #gameState = Object.freeze({ ...DEFAULT_GAME_STATE });
    #eventListeners = new Map();
    #currentGameId = null;
    #gameEngine = null; // Game-specific engine instance
    
    constructor(gameType, config = {}) {
        if (this.constructor === GameManagerBase) {
            throw new Error('GameManagerBase is abstract and cannot be instantiated directly');
        }
        
        this.#gameType = gameType;
        this.config = Object.freeze({
            allowUndo: true,
            autoSave: true,
            maxPlayers: 2,
            minPlayers: 1,
            ...config
        });
        
        // CRITICAL: Connect to debug system for OnlyOffice troubleshooting
        window.debug?.info('GameManagerBase', `Initializing ${gameType} game manager`);
    }

    // =============================================================================
    // TEMPLATE METHOD - DEFINES ALGORITHM STEPS  
    // =============================================================================
    
    /**
     * Template method - defines game initialization algorithm
     * NEVER override this method in derived classes
     * 
     * ONLYOFFICE INTEGRATION: This method should be called after OnlyOffice 
     * plugin.init() completes successfully
     */
    async initializeGame() {
        if (this.#initialized) return;
        
        try {
            window.debug?.info('GameManagerBase', `Starting ${this.#gameType} initialization`);
            
            // Step 1: Validate configuration (Abstract - MUST implement)
            await this.validateGameConfiguration();
            
            // Step 2: Initialize game engine (Abstract - MUST implement)  
            await this.initializeGameEngine();
            
            // Step 3: Create game interface (Abstract - MUST implement)
            await this.createGameInterface();
            
            // Step 4: Setup event handlers (Abstract - MUST implement)
            await this.setupEventHandlers();
            
            // Step 5: Connect to collaboration system (Virtual - CAN override)
            await this.initializeCollaboration();
            
            // Step 6: Finalize setup (Virtual - CAN override)
            await this.finalizeGameSetup();
            
            this.#initialized = true;
            this.onGameInitializationComplete(); // Hook method
            
            window.debug?.info('GameManagerBase', `${this.#gameType} initialization completed`);
            
        } catch (error) {
            await this.handleGameInitializationError(error);
            throw error;
        }
    }

    // =============================================================================
    // ABSTRACT METHODS - MUST BE IMPLEMENTED BY CONCRETE CLASSES
    // =============================================================================
    
    /**
     * Validate game-specific configuration
     * @throws {ValidationError} When configuration is invalid
     */
    async validateGameConfiguration() {
        throw new Error(`Must implement validateGameConfiguration() for game type: ${this.#gameType}`);
    }
    
    /**
     * Initialize game-specific engine
     * ONLYOFFICE NOTE: Engine should integrate with OnlyOffice API for document manipulation
     */
    async initializeGameEngine() {
        throw new Error(`Must implement initializeGameEngine() for game type: ${this.#gameType}`);
    }
    
    /**
     * Create game-specific user interface
     * ONLYOFFICE NOTE: UI should follow OnlyOffice plugin UI patterns and theming
     */
    async createGameInterface() {
        throw new Error(`Must implement createGameInterface() for game type: ${this.#gameType}`);
    }
    
    /**
     * Setup game-specific event handlers
     * ONLYOFFICE NOTE: Should integrate with OnlyOffice button handlers and theme changes
     */
    async setupEventHandlers() {
        throw new Error(`Must implement setupEventHandlers() for game type: ${this.#gameType}`);
    }
    
    /**
     * Process a game move/action
     * @param {Object} moveData - Game-specific move data
     * @returns {Promise<boolean>} Success status
     */
    async makeMove(moveData) {
        throw new Error(`Must implement makeMove() for game type: ${this.#gameType}`);
    }

    // =============================================================================
    // VIRTUAL METHODS - CAN BE OVERRIDDEN BY CONCRETE CLASSES
    // =============================================================================
    
    /**
     * Initialize collaboration features
     * ONLYOFFICE NOTE: Should use OnlyOffice collaboration API when available
     */
    async initializeCollaboration() {
        // Default implementation - no collaboration
        window.debug?.debug('GameManagerBase', 'No collaboration system initialized');
    }
    
    /**
     * Finalize game setup after all core components initialized
     */
    async finalizeGameSetup() {
        // Default implementation - load saved game state
        await this.loadGameSettings();
    }
    
    /**
     * Handle initialization errors
     */
    async handleGameInitializationError(error) {
        window.debug?.error('GameManagerBase', 'Game initialization failed', {
            gameType: this.#gameType,
            error: error.message
        });
    }

    // =============================================================================
    // COMMON GAME MANAGEMENT METHODS
    // =============================================================================
    
    /**
     * Start a new game with specified options
     * @param {Object} gameOptions - Game configuration options
     * @returns {Promise<string>} Game ID
     */
    async startNewGame(gameOptions = {}) {
        window.debug?.debug('GameManagerBase', 'Starting new game', { 
            gameType: this.#gameType, 
            options: gameOptions 
        });

        try {
            // Generate new game ID
            this.#currentGameId = this.generateGameId();
            
            // Reset game state
            this.#gameState = Object.freeze({
                ...DEFAULT_GAME_STATE,
                status: 'playing',
                gameId: this.#currentGameId,
                gameType: this.#gameType,
                startTime: new Date().toISOString(),
                settings: { ...this.config, ...gameOptions.settings }
            });
            
            // Setup players
            await this.setupPlayers(gameOptions.players);
            
            // Initialize game engine with new game
            if (this.#gameEngine && this.#gameEngine.startNewGame) {
                await this.#gameEngine.startNewGame(gameOptions);
            }
            
            // Notify listeners
            this.notifyGameEvent(GAME_EVENTS.GAME_STARTED, {
                gameId: this.#currentGameId,
                gameType: this.#gameType
            });
            
            this.notifyGameEvent(GAME_EVENTS.STATE_CHANGED, this.#gameState);
            
            window.debug?.info('GameManagerBase', 'New game started', { 
                gameId: this.#currentGameId 
            });

            return this.#currentGameId;

        } catch (error) {
            throw new Error(`Failed to start new ${this.#gameType} game: ${error.message}`);
        }
    }
    
    /**
     * Get current game state (immutable)
     */
    getCurrentGameState() {
        return this.#gameState;
    }
    
    /**
     * Update game state immutably
     * @param {Object} updates - State updates
     * @param {string} actionType - Action that triggered the update
     */
    updateGameState(updates, actionType = 'UPDATE') {
        const oldState = this.#gameState;
        
        this.#gameState = Object.freeze({
            ...oldState,
            ...updates,
            lastUpdated: new Date().toISOString(),
            lastAction: actionType
        });
        
        // Auto-save if enabled
        // ONLYOFFICE NOTE: This should save to OnlyOffice document when possible
        if (this.config.autoSave) {
            this.autoSave().catch(error => {
                window.debug?.warn('GameManagerBase', 'Auto-save failed', error);
            });
        }
        
        // Notify listeners
        this.notifyGameEvent(GAME_EVENTS.STATE_CHANGED, this.#gameState);
        
        window.debug?.debug('GameManagerBase', 'Game state updated', {
            actionType,
            changes: this.calculateStateChanges(oldState, this.#gameState)
        });
    }

    // =============================================================================
    // EVENT SYSTEM
    // =============================================================================
    
    /**
     * Add event listener
     */
    addEventListener(event, callback) {
        if (!this.#eventListeners.has(event)) {
            this.#eventListeners.set(event, []);
        }
        this.#eventListeners.get(event).push(callback);
        
        window.debug?.debug('GameManagerBase', `Event listener added: ${event}`);
    }
    
    /**
     * Remove event listener
     */
    removeEventListener(event, callback) {
        if (this.#eventListeners.has(event)) {
            const listeners = this.#eventListeners.get(event);
            const index = listeners.indexOf(callback);
            if (index > -1) {
                listeners.splice(index, 1);
            }
        }
    }
    
    /**
     * Notify event listeners with error isolation
     */
    notifyGameEvent(event, data) {
        const listeners = this.#eventListeners.get(event) || [];
        
        listeners.forEach(callback => {
            try {
                callback(data);
            } catch (error) {
                window.debug?.error('GameManagerBase', `Event listener error: ${event}`, error);
                // Continue processing other listeners
            }
        });
    }

    // =============================================================================
    // UTILITY METHODS
    // =============================================================================
    
    /**
     * Generate unique game ID
     */
    generateGameId() {
        return `${this.#gameType}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }
    
    /**
     * Setup players for the game
     */
    async setupPlayers(playersConfig = []) {
        const players = Array.isArray(playersConfig) ? playersConfig : [];
        
        // Validate player count
        if (players.length < this.config.minPlayers || players.length > this.config.maxPlayers) {
            throw new Error(`Invalid player count. Requires ${this.config.minPlayers}-${this.config.maxPlayers} players`);
        }
        
        this.updateGameState({ players }, 'SETUP_PLAYERS');
    }
    
    /**
     * Auto-save game state
     * ONLYOFFICE NOTE: Should integrate with OnlyOffice document storage
     */
    async autoSave() {
        try {
            const saveData = {
                gameId: this.#currentGameId,
                gameType: this.#gameType,
                gameState: this.#gameState,
                savedAt: new Date().toISOString()
            };
            
            // Save to localStorage as fallback
            const storageKey = `casual_game_${this.#currentGameId}`;
            localStorage.setItem(storageKey, JSON.stringify(saveData));
            
            window.debug?.debug('GameManagerBase', 'Game auto-saved', { 
                gameId: this.#currentGameId 
            });

        } catch (error) {
            window.debug?.warn('GameManagerBase', 'Auto-save failed', error);
        }
    }
    
    /**
     * Load game settings from storage
     */
    async loadGameSettings() {
        try {
            const storageKey = `casual_game_settings_${this.#gameType}`;
            const stored = localStorage.getItem(storageKey);
            
            if (stored) {
                const settings = JSON.parse(stored);
                window.debug?.debug('GameManagerBase', 'Game settings loaded', settings);
            }
            
        } catch (error) {
            window.debug?.warn('GameManagerBase', 'Failed to load game settings', error);
        }
    }
    
    /**
     * Calculate differences between two states
     */
    calculateStateChanges(oldState, newState) {
        const changes = {};
        
        for (const [key, value] of Object.entries(newState)) {
            if (oldState[key] !== value) {
                changes[key] = {
                    from: oldState[key],
                    to: value
                };
            }
        }
        
        return changes;
    }
    
    // =============================================================================
    // HOOK METHODS - CALLED AT SPECIFIC POINTS
    // =============================================================================
    
    /**
     * Called when game initialization completes successfully
     */
    onGameInitializationComplete() {
        window.debug?.info('GameManagerBase', `${this.#gameType} game manager ready`);
    }
    
    // =============================================================================
    // CLEANUP
    // =============================================================================
    
    /**
     * Cleanup game manager resources
     */
    async cleanup() {
        window.debug?.debug('GameManagerBase', `Cleaning up ${this.#gameType} game manager`);

        // Auto-save current state
        if (this.#currentGameId && this.config.autoSave) {
            await this.autoSave();
        }
        
        // Clear event listeners
        this.#eventListeners.clear();
        
        // Reset state
        this.#currentGameId = null;
        this.#gameState = Object.freeze({ ...DEFAULT_GAME_STATE });
        this.#initialized = false;
        
        // Cleanup game engine if available
        if (this.#gameEngine && typeof this.#gameEngine.cleanup === 'function') {
            await this.#gameEngine.cleanup();
        }
    }
    
    // =============================================================================
    // PROTECTED METHODS FOR CONCRETE CLASSES
    // =============================================================================
    
    /**
     * Set the game engine instance (for concrete classes)
     */
    _setGameEngine(engine) {
        this.#gameEngine = engine;
    }
    
    /**
     * Get the game engine instance (for concrete classes)
     */
    _getGameEngine() {
        return this.#gameEngine;
    }
    
    /**
     * Check if game is initialized (for concrete classes)
     */
    _isInitialized() {
        return this.#initialized;
    }
    
    /**
     * Get game type (for concrete classes)
     */
    _getGameType() {
        return this.#gameType;
    }
}

// =============================================================================
// 4. EXPORTS
// =============================================================================
window.GameManagerBase = GameManagerBase;
window.GAME_EVENTS = GAME_EVENTS;