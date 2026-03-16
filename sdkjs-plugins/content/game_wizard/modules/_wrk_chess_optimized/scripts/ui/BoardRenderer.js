/**
 * Chess Board Renderer
 * Following OnlyOffice Plugin Development Standards
 * 
 * Bridges between new architecture and existing chess.js rendering
 * Based on: _coding_standard/02_api_reference_patterns.md#performance-optimization-patterns
 */

class ChessBoardRenderer {
    constructor(containerElement, themeManager) {
        this.container = containerElement;
        this.themeManager = themeManager;
        this.isInitialized = false;
        this.legacyBoard = null; // Reference to existing CChessBoard instance
        this.currentTheme = 'light';
    }

    /**
     * Initialize board renderer
     */
    async initialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'BoardRenderer initializing');

        try {
            // Wait for legacy chess.js to be available
            await this.waitForLegacyChess();
            
            // Initialize legacy board if available
            if (window.ChessBoard && window.ChessBoard.isReady && window.ChessBoard.isReady()) {
                this.legacyBoard = window.g_board;
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Using ready chess board instance');
                
                // Initialize with starting position
                try {
                    if (this.legacyBoard.draw) {
                        this.legacyBoard.draw(true);
                    }
                } catch (error) {
                    window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Failed to draw chess board', error);
                }
                
            } else if (window.g_board && window.CChessBoard) {
                this.legacyBoard = window.g_board;
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Using existing chess board instance (not fully ready)');
                
                // Try to initialize it
                setTimeout(() => {
                    if (this.legacyBoard && this.legacyBoard.draw) {
                        try {
                            this.legacyBoard.draw(true);
                        } catch (error) {
                            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                                'Failed to draw chess board after delay', error);
                        }
                    }
                }, 500);
                
            } else if (window.CChessBoard && this.container) {
                // Create new instance if needed
                try {
                    this.legacyBoard = new window.CChessBoard(this.container.id);
                    window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Created new chess board instance');
                } catch (error) {
                    window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Failed to create new chess board instance', error);
                    this.createBasicBoard();
                }
            } else {
                // Create minimal board placeholder
                this.createBasicBoard();
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Created basic board placeholder - legacy components not available');
            }
            
            // Apply current theme
            if (this.themeManager) {
                this.currentTheme = this.themeManager.getCurrentTheme();
                await this.updateTheme(this.currentTheme);
            }
            
            // Setup event listeners
            this.setupEventListeners();
            
            this.isInitialized = true;
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'BoardRenderer initialized');

        } catch (error) {
            throw new window.ChessErrors.ChessRenderingError(
                'Board renderer initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Wait for legacy chess.js components to be available
     */
    async waitForLegacyChess() {
        return new Promise((resolve, reject) => {
            let attempts = 0;
            const maxAttempts = 30; // 3 seconds maximum wait
            
            const checkAvailability = () => {
                // Check for both the class and if ChessBoard wrapper is ready
                if (window.CChessBoard && window.ChessBoard && window.ChessBoard.isReady && window.ChessBoard.isReady()) {
                    window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Legacy chess components found and ready');
                    resolve();
                    return;
                }
                
                // Also check if we have the basic components even if not fully ready
                if (window.CChessBoard && window.ChessBoard && attempts > 15) {
                    window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Legacy chess components found but not fully ready');
                    resolve();
                    return;
                }
                
                attempts++;
                if (attempts >= maxAttempts) {
                    window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Legacy chess components not available, proceeding without them');
                    // Don't reject - just resolve and handle gracefully
                    resolve();
                    return;
                }
                
                setTimeout(checkAvailability, 100);
            };
            
            checkAvailability();
        });
    }

    /**
     * Create basic chess board placeholder when legacy components aren't available
     */
    createBasicBoard() {
        if (!this.container) {
            return;
        }

        // Create basic board structure
        this.container.innerHTML = `
            <div class="chess-board-placeholder" style="
                width: 400px;
                height: 400px;
                border: 2px solid var(--border-color);
                background: linear-gradient(45deg, var(--board-light-square) 25%, transparent 25%), 
                           linear-gradient(-45deg, var(--board-light-square) 25%, transparent 25%), 
                           linear-gradient(45deg, transparent 75%, var(--board-light-square) 75%), 
                           linear-gradient(-45deg, transparent 75%, var(--board-light-square) 75%);
                background-size: 50px 50px;
                background-position: 0 0, 0 25px, 25px -25px, -25px 0px;
                background-color: var(--board-dark-square);
                display: flex;
                align-items: center;
                justify-content: center;
                margin: 0 auto;
                border-radius: var(--border-radius);
                box-shadow: var(--shadow-lg);
            ">
                <div style="
                    background: var(--bg-primary);
                    padding: var(--spacing-md);
                    border-radius: var(--border-radius);
                    box-shadow: var(--shadow-md);
                    text-align: center;
                    color: var(--text-primary);
                    font-size: var(--font-size-base);
                ">
                    <div style="margin-bottom: var(--spacing-sm);">♔ Chess Board ♕</div>
                    <div id="chess-status" style="font-size: var(--font-size-sm); color: var(--text-secondary);">
                        Chess engine ready - basic mode
                    </div>
                    <div style="font-size: var(--font-size-xs); color: var(--text-muted); margin-top: var(--spacing-xs);">
                        Full chess board loading...
                    </div>
                </div>
            </div>
        `;

        // Create mock legacy board interface
        this.legacyBoard = {
            draw: () => {},
            init: (data) => {
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Basic board received init data', { hasData: !!data });
            },
            getResult: () => ({ image: '' }),
            getData: () => 'rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR',
            fen: 'rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR'
        };
        
        // Try to upgrade to real board after a delay
        this.tryUpgradeBoard();
    }

    /**
     * Try to upgrade from basic board to real chess board
     */
    tryUpgradeBoard() {
        let attempts = 0;
        const maxAttempts = 20; // 10 seconds
        
        const checkForUpgrade = () => {
            if (window.ChessBoard && window.ChessBoard.isReady && window.ChessBoard.isReady()) {
                // Upgrade to real board
                this.legacyBoard = window.g_board;
                
                // Clear the placeholder
                if (this.container) {
                    this.container.innerHTML = '';
                }
                
                // Initialize the real board
                try {
                    if (this.legacyBoard.draw) {
                        this.legacyBoard.draw(true);
                    }
                    
                    window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Successfully upgraded to real chess board');
                    return;
                } catch (error) {
                    window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Failed to upgrade to real chess board', error);
                }
            }
            
            attempts++;
            if (attempts < maxAttempts) {
                setTimeout(checkForUpgrade, 500);
            } else {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Giving up on chess board upgrade - staying with basic board');
            }
        };
        
        // Start checking after a short delay
        setTimeout(checkForUpgrade, 1000);
    }

    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Listen for theme changes
        if (this.themeManager) {
            this.themeManager.addEventListener('themeChanged', (themeData) => {
                this.onThemeChanged(themeData);
            });
        }
        
        // Handle window resize
        window.addEventListener('resize', () => {
            this.handleResize();
        });
    }

    /**
     * Update theme
     */
    async updateTheme(theme) {
        this.currentTheme = theme;
        
        if (this.legacyBoard && this.legacyBoard.draw) {
            // Force redraw with new theme
            try {
                this.legacyBoard.draw(true);
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Board redrawn with new theme', { theme });
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Board redraw failed', error);
            }
        }
    }

    /**
     * Handle theme change event
     */
    onThemeChanged(themeData) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'Board renderer received theme change', themeData);
        
        this.updateTheme(themeData.theme);
    }

    /**
     * Handle window/plugin resize
     */
    handleResize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'Board renderer handling resize');
            
        if (this.legacyBoard && this.legacyBoard.resize) {
            // Use debouncing to avoid excessive redraws
            if (this.resizeTimeout) {
                clearTimeout(this.resizeTimeout);
            }
            
            this.resizeTimeout = setTimeout(() => {
                try {
                    this.legacyBoard.resize();
                    window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Board resized successfully');
                } catch (error) {
                    window.ChessErrorHandler?.handleError(
                        new window.ChessErrors.ChessRenderingError(
                            'Board resize failed',
                            { originalError: error }
                        )
                    );
                }
                this.resizeTimeout = null;
            }, 150); // 150ms debounce
        }
    }

    /**
     * Force redraw of the board
     */
    forceRedraw() {
        if (this.legacyBoard && this.legacyBoard.draw) {
            try {
                this.legacyBoard.draw(true);
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Board force redrawn');
            } catch (error) {
                throw new window.ChessErrors.ChessRenderingError(
                    'Board force redraw failed',
                    { originalError: error }
                );
            }
        }
    }

    /**
     * Cancel current selection (called from UI controller)
     */
    cancelSelection() {
        if (this.legacyBoard) {
            try {
                // Reset any selection state in legacy board
                if (this.legacyBoard.track !== undefined) {
                    this.legacyBoard.track = "";
                }
                
                // Force redraw to clear selections
                this.forceRedraw();
                
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Board selection cancelled');
                    
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Failed to cancel selection', error);
            }
        }
    }

    /**
     * Get current board data (for saving)
     */
    getBoardData() {
        if (this.legacyBoard && this.legacyBoard.getData) {
            try {
                return this.legacyBoard.getData();
            } catch (error) {
                window.ChessErrorHandler?.handleError(
                    new window.ChessErrors.ChessPersistenceError(
                        'Failed to get board data',
                        { originalError: error }
                    )
                );
                return null;
            }
        }
        return null;
    }

    /**
     * Get board result image (for OLE object)
     */
    getBoardResult(width, height) {
        if (this.legacyBoard && this.legacyBoard.getResult) {
            try {
                return this.legacyBoard.getResult(width, height);
            } catch (error) {
                window.ChessErrorHandler?.handleError(
                    new window.ChessErrors.ChessRenderingError(
                        'Failed to get board result',
                        { width, height, originalError: error }
                    )
                );
                return null;
            }
        }
        return null;
    }

    /**
     * Initialize board with data (from FEN or saved state)
     */
    initializeWithData(data) {
        if (this.legacyBoard && this.legacyBoard.init) {
            try {
                this.legacyBoard.init(data);
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Board initialized with data', { hasData: !!data });
            } catch (error) {
                window.ChessErrorHandler?.handleError(
                    new window.ChessErrors.ChessInitializationError(
                        'Failed to initialize board with data',
                        { data, originalError: error }
                    )
                );
            }
        }
    }

    /**
     * Update board position from game state
     */
    updateFromGameState(gameState) {
        if (!gameState || !this.legacyBoard) {
            return;
        }

        try {
            // If we have a FEN string, use it
            if (gameState.fen) {
                this.initializeWithData(gameState.fen);
            } else if (gameState.board) {
                // Convert internal board representation if needed
                // This would require implementing a conversion method
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Board update from internal representation not yet implemented');
            }
        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessRenderingError(
                    'Failed to update board from game state',
                    { gameState, originalError: error }
                )
            );
        }
    }

    /**
     * Get performance metrics
     */
    getPerformanceMetrics() {
        return {
            isInitialized: this.isInitialized,
            currentTheme: this.currentTheme,
            hasLegacyBoard: !!this.legacyBoard,
            containerSize: this.container ? {
                width: this.container.offsetWidth,
                height: this.container.offsetHeight
            } : null
        };
    }

    /**
     * Cleanup board renderer
     */
    async cleanup() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'BoardRenderer cleanup');

        // Clear resize timeout
        if (this.resizeTimeout) {
            clearTimeout(this.resizeTimeout);
            this.resizeTimeout = null;
        }

        // Remove event listeners
        if (this.themeManager) {
            // Remove theme change listener (would need to track the specific callback)
        }

        // Clear references
        this.legacyBoard = null;
        this.container = null;
        this.themeManager = null;
        this.isInitialized = false;
    }
}

// Export board renderer
window.ChessBoardRenderer = ChessBoardRenderer;