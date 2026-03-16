/**
 * Chess Plugin Lifecycle Manager
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/CODING_STANDARD.md#template-method-pattern-for-plugin-lifecycle
 */

/**
 * Base Plugin Lifecycle Manager implementing Template Method Pattern
 */
class PluginLifecycleManagerBase {
    constructor() {
        this.isInitialized = false;
        this.initializationStartTime = null;
        this.components = new Map();
        this.lifecycleHooks = new Map();
    }

    /**
     * Template Method: Main initialization flow
     * This method defines the algorithm structure and calls hook methods
     */
    async initialize(data = null) {
        if (this.isInitialized) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                'Plugin already initialized, skipping');
            return;
        }

        this.initializationStartTime = Date.now();
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Starting plugin initialization');

        try {
            // Template Method steps
            await this.preInitialize();
            await this.initializeServices();
            await this.initializeUI();
            await this.processInitialData(data);
            await this.postInitialize();
            
            this.isInitialized = true;
            const duration = Date.now() - this.initializationStartTime;
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                `Plugin initialized successfully in ${duration}ms`);

            this.triggerLifecycleHook('initialized', { duration });

        } catch (error) {
            const lifecycleError = new window.ChessErrors.ChessInitializationError(
                'Plugin initialization failed',
                { originalError: error, initTime: this.initializationStartTime }
            );
            window.ChessErrorHandler?.handleError(lifecycleError);
            throw lifecycleError;
        }
    }

    /**
     * Template Method: Plugin cleanup flow
     */
    async cleanup() {
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Starting plugin cleanup');

        try {
            await this.preCleanup();
            await this.cleanupComponents();
            await this.cleanupServices();
            await this.postCleanup();

            this.isInitialized = false;
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                'Plugin cleanup completed');

        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessInitializationError(
                    'Plugin cleanup failed',
                    { originalError: error }
                )
            );
        }
    }

    /**
     * Hook methods to be implemented by subclasses
     */
    async preInitialize() {
        // Override in subclass
    }

    async initializeServices() {
        // Override in subclass
    }

    async initializeUI() {
        // Override in subclass
    }

    async processInitialData(data) {
        // Override in subclass
    }

    async postInitialize() {
        // Override in subclass
    }

    async preCleanup() {
        // Override in subclass
    }

    async cleanupComponents() {
        // Override in subclass
    }

    async cleanupServices() {
        // Override in subclass
    }

    async postCleanup() {
        // Override in subclass
    }

    /**
     * Component registration and management
     */
    registerComponent(name, component) {
        if (this.components.has(name)) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                `Component '${name}' already registered, replacing`);
        }
        
        this.components.set(name, component);
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            `Component registered: ${name}`);
    }

    getComponent(name) {
        return this.components.get(name);
    }

    /**
     * Lifecycle hook system
     */
    addLifecycleHook(event, callback) {
        if (!this.lifecycleHooks.has(event)) {
            this.lifecycleHooks.set(event, []);
        }
        this.lifecycleHooks.get(event).push(callback);
    }

    triggerLifecycleHook(event, data = null) {
        const hooks = this.lifecycleHooks.get(event) || [];
        hooks.forEach(callback => {
            try {
                callback(data);
            } catch (error) {
                window.ChessErrorHandler?.handleError(
                    new window.ChessErrors.ChessInitializationError(
                        `Lifecycle hook '${event}' failed`,
                        { event, originalError: error }
                    )
                );
            }
        });
    }
}

/**
 * Chess-specific Plugin Lifecycle Manager
 */
class ChessPluginLifecycleManager extends PluginLifecycleManagerBase {
    constructor() {
        super();
        this.chessEngine = null;
        this.gameManager = null;
        this.collaborationManager = null;
        this.uiController = null;
        this.themeManager = null;
        this.boardRenderer = null;
    }

    /**
     * Pre-initialization setup
     */
    async preInitialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Chess pre-initialization starting');

        // Initialize error handler with user notification
        if (!window.ChessErrorHandler.isInitialized) {
            window.ChessErrorHandler.initialize(this.showUserNotification.bind(this));
        }

        // Setup performance monitoring
        this.setupPerformanceMonitoring();
        
        window.ChessDebug?.logLifecycleEvent('pre_initialize_complete');
    }

    /**
     * Initialize core services
     */
    async initializeServices() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Initializing chess services');

        try {
            // Initialize chess engine
            if (window.ChessEngine) {
                this.chessEngine = new window.ChessEngine();
                this.registerComponent('chessEngine', this.chessEngine);
                await this.chessEngine.initialize();
            }

            // Initialize game manager
            if (window.ChessGameManager) {
                this.gameManager = new window.ChessGameManager(this.chessEngine);
                this.registerComponent('gameManager', this.gameManager);
                await this.gameManager.initialize();
            }

            // Initialize AI opponent manager (replaces collaboration manager)
            if (window.ChessAIOpponentManager) {
                this.aiOpponent = new window.ChessAIOpponentManager();
                this.registerComponent('aiOpponent', this.aiOpponent);
                await this.aiOpponent.initialize();
                
                // Connect AI opponent to game manager
                if (this.gameManager) {
                    this.gameManager.setAIOpponent(this.aiOpponent);
                }
            }

            window.ChessDebug?.logLifecycleEvent('services_initialized', {
                components: Array.from(this.components.keys())
            });

        } catch (error) {
            throw new window.ChessErrors.ChessInitializationError(
                'Service initialization failed',
                { phase: 'services', originalError: error }
            );
        }
    }

    /**
     * Initialize UI components
     */
    async initializeUI() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Initializing chess UI');

        try {
            // Initialize theme manager
            if (window.ChessThemeManager) {
                this.themeManager = new window.ChessThemeManager();
                this.registerComponent('themeManager', this.themeManager);
                await this.themeManager.initialize();
            }

            // Initialize board renderer - use SimpleBoardRenderer for AI gameplay
            const chessContainer = document.getElementById('chess');
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'Looking for chess container', { 
                    found: !!chessContainer,
                    id: chessContainer?.id,
                    className: chessContainer?.className 
                });
            
            if (window.SimpleBoardRenderer && chessContainer) {
                this.boardRenderer = new window.SimpleBoardRenderer(
                    chessContainer,
                    this.gameManager,
                    this.aiOpponent
                );
                this.registerComponent('boardRenderer', this.boardRenderer);
                await this.boardRenderer.initialize();
            } else if (window.ChessBoardRenderer) {
                // Fallback to legacy renderer
                this.boardRenderer = new window.ChessBoardRenderer(
                    document.getElementById('chess'),
                    this.themeManager
                );
                this.registerComponent('boardRenderer', this.boardRenderer);
                await this.boardRenderer.initialize();
            }

            // Initialize UI controller with AI opponent
            if (window.ChessUIController) {
                this.uiController = new window.ChessUIController(
                    this.gameManager,
                    this.boardRenderer,
                    this.aiOpponent
                );
                this.registerComponent('uiController', this.uiController);
                await this.uiController.initialize();
            }

            window.ChessDebug?.logLifecycleEvent('ui_initialized');

        } catch (error) {
            throw new window.ChessErrors.ChessRenderingError(
                'UI initialization failed',
                { phase: 'ui', originalError: error }
            );
        }
    }

    /**
     * Process initial data from OnlyOffice
     */
    async processInitialData(data) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Processing initial data', { dataType: typeof data, hasData: !!data });

        try {
            if (data && this.gameManager) {
                // Handle OLE object data for chess game restoration
                const gameData = this.parseInitialData(data);
                if (gameData) {
                    await this.gameManager.loadGameState(gameData);
                    window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                        'Game state restored from initial data');
                }
            }

            // Start new game if no data provided
            if (!data && this.gameManager) {
                await this.gameManager.startNewGame();
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                    'New chess game started');
            }

        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessInitializationError(
                    'Initial data processing failed',
                    { data, originalError: error }
                )
            );
            
            // Fallback: start new game
            if (this.gameManager) {
                await this.gameManager.startNewGame();
            }
        }
    }

    /**
     * Post-initialization setup
     */
    async postInitialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Chess post-initialization');

        try {
            // Connect components
            this.wireComponentEvents();

            // Setup collaboration if available
            if (this.collaborationManager && window.Asc?.plugin?.info?.isCoAuthoringEnable) {
                await this.collaborationManager.joinSession();
            }

            // Apply current theme
            if (this.themeManager) {
                await this.themeManager.applyCurrentTheme();
            }

            // Update UI to show initial state
            if (this.uiController) {
                this.uiController.updateDisplay();
            }
            
            // Update player display for AI game
            this.updatePlayerDisplay();

            window.ChessDebug?.logLifecycleEvent('post_initialize_complete');

        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessInitializationError(
                    'Post-initialization failed',
                    { originalError: error }
                )
            );
        }
    }

    /**
     * Setup component event wiring
     */
    wireComponentEvents() {
        // Game manager events
        if (this.gameManager && this.uiController) {
            this.gameManager.addEventListener('gameStateChanged', 
                this.uiController.onGameStateChanged.bind(this.uiController));
            this.gameManager.addEventListener('moveAttempted', 
                this.uiController.onMoveAttempted.bind(this.uiController));
        }

        // AI opponent events
        if (this.aiOpponent && this.gameManager) {
            this.aiOpponent.addEventListener('newGameStarted', 
                this.gameManager.onPlayerJoined.bind(this.gameManager));
            this.aiOpponent.addEventListener('difficultyChanged', 
                (data) => this.handleDifficultyChanged(data));
        }
        
        // Setup UI control handlers
        this.setupUIControls();

        // Theme change events
        if (this.themeManager && this.boardRenderer) {
            this.themeManager.addEventListener('themeChanged', 
                this.boardRenderer.onThemeChanged.bind(this.boardRenderer));
        }
    }

    /**
     * Component cleanup
     */
    async cleanupComponents() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
            'Cleaning up chess components');

        const componentsToCleanup = [
            'uiController',
            'boardRenderer', 
            'aiOpponent',
            'gameManager',
            'chessEngine',
            'themeManager'
        ];

        for (const componentName of componentsToCleanup) {
            const component = this.getComponent(componentName);
            if (component && typeof component.cleanup === 'function') {
                try {
                    await component.cleanup();
                    window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                        `Component cleaned up: ${componentName}`);
                } catch (error) {
                    window.ChessErrorHandler?.handleError(
                        new window.ChessErrors.ChessInitializationError(
                            `Component cleanup failed: ${componentName}`,
                            { component: componentName, originalError: error }
                        )
                    );
                }
            }
        }

        this.components.clear();
    }

    /**
     * Parse initial data for game restoration
     */
    parseInitialData(data) {
        try {
            if (typeof data === 'string' && data.trim()) {
                return JSON.parse(data);
            }
            return null;
        } catch (error) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                'Failed to parse initial data', { data, error: error.message });
            return null;
        }
    }

    /**
     * Performance monitoring setup
     */
    setupPerformanceMonitoring() {
        if (window.performance && window.performance.mark) {
            window.performance.mark('chess-plugin-init-start');
            
            this.addLifecycleHook('initialized', () => {
                window.performance.mark('chess-plugin-init-end');
                window.performance.measure('chess-plugin-initialization', 
                    'chess-plugin-init-start', 'chess-plugin-init-end');
            });
        }
    }

    /**
     * User notification callback for error handler
     */
    showUserNotification(message, category) {
        if (this.uiController && this.uiController.showNotification) {
            this.uiController.showNotification(message, category);
        } else {
            // Fallback notification
            console.warn('Chess Plugin:', message);
        }
    }

    /**
     * Setup UI control handlers
     */
    setupUIControls() {
        // AI difficulty selector
        const difficultySelect = document.getElementById('ai-difficulty');
        if (difficultySelect && this.aiOpponent) {
            difficultySelect.addEventListener('change', (event) => {
                this.aiOpponent.setDifficulty(event.target.value);
                this.updatePlayerDisplay();
            });
        }

        // New game button
        const newGameBtn = document.getElementById('new-game-btn');
        if (newGameBtn && this.gameManager) {
            newGameBtn.addEventListener('click', async () => {
                try {
                    await this.gameManager.startNewGame();
                    if (this.aiOpponent) {
                        await this.aiOpponent.startNewGame();
                    }
                } catch (error) {
                    this.showUserNotification('Failed to start new game: ' + error.message, 'error');
                }
            });
        }

        // Undo button
        const undoBtn = document.getElementById('undo-btn');
        if (undoBtn && this.gameManager) {
            undoBtn.addEventListener('click', async () => {
                try {
                    await this.gameManager.undoLastMove();
                } catch (error) {
                    this.showUserNotification('Cannot undo: ' + error.message, 'warning');
                }
            });
        }

        // Resign button
        const resignBtn = document.getElementById('resign-btn');
        if (resignBtn && this.gameManager) {
            resignBtn.addEventListener('click', async () => {
                try {
                    await this.gameManager.resignGame();
                } catch (error) {
                    this.showUserNotification('Failed to resign: ' + error.message, 'error');
                }
            });
        }
    }

    /**
     * Handle AI difficulty change
     */
    handleDifficultyChanged(data) {
        this.updatePlayerDisplay();
        this.showUserNotification(`AI difficulty changed to ${data.difficulty}`, 'info');
    }

    /**
     * Update player display in UI
     */
    updatePlayerDisplay() {
        if (!this.aiOpponent) return;
        
        const playerInfo = this.aiOpponent.getPlayerInfo();
        
        // Update white player display
        const whitePlayerEl = document.getElementById('white-player');
        if (whitePlayerEl) {
            whitePlayerEl.innerHTML = `<span>White: ${playerInfo.white.name}</span>`;
            if (playerInfo.white.isPlayer) {
                whitePlayerEl.classList.add('active');
            } else {
                whitePlayerEl.classList.remove('active');
            }
        }
        
        // Update black player display
        const blackPlayerEl = document.getElementById('black-player');
        if (blackPlayerEl) {
            blackPlayerEl.innerHTML = `<span>Black: ${playerInfo.black.name}</span>`;
            if (playerInfo.black.isPlayer) {
                blackPlayerEl.classList.add('active');
            } else {
                blackPlayerEl.classList.remove('active');
            }
        }

        // Update difficulty selector
        const difficultySelect = document.getElementById('ai-difficulty');
        if (difficultySelect) {
            difficultySelect.value = this.aiOpponent.aiDifficulty;
        }
    }
}

// Export classes
window.ChessLifecycleClasses = {
    PluginLifecycleManagerBase,
    ChessPluginLifecycleManager
};

// Create global lifecycle manager instance
window.ChessLifecycleManager = new ChessPluginLifecycleManager();