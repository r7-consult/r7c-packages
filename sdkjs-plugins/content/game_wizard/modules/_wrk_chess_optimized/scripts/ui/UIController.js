/**
 * Chess Plugin UI Controller
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/02_api_reference_patterns.md#performance-optimization-patterns
 */

class ChessUIController {
    constructor(gameManager, boardRenderer, aiOpponent) {
        this.gameManager = gameManager;
        this.boardRenderer = boardRenderer;
        this.aiOpponent = aiOpponent;
        this.isInitialized = false;
        
        // UI elements
        this.elements = {
            connectionStatus: null,
            gameStatusText: null,
            whitePlayer: null,
            blackPlayer: null,
            moveHistory: null,
            undoBtn: null,
            resignBtn: null,
            newGameBtn: null,
            loadingOverlay: null,
            notificationContainer: null
        };
        
        // State tracking
        this.isLoading = false;
        this.activeNotifications = new Map();
        this.notificationIdCounter = 0;
    }

    /**
     * Initialize UI controller
     */
    async initialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'UIController initializing');

        try {
            // Cache UI elements
            this.cacheUIElements();
            
            // Setup event listeners
            this.setupEventListeners();
            
            // Initialize loading state
            this.updateLoadingState(false);
            
            // Set initial connection status
            this.updateConnectionStatus('connected');
            
            this.isInitialized = true;
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'UIController initialized');

        } catch (error) {
            throw new window.ChessErrors.ChessRenderingError(
                'UI controller initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Cache UI elements for performance
     */
    cacheUIElements() {
        const elementIds = {
            connectionStatus: 'connection-status',
            gameStatusText: 'game-status-text', 
            whitePlayer: 'white-player',
            blackPlayer: 'black-player',
            moveHistory: 'move-history',
            undoBtn: 'undo-btn',
            resignBtn: 'resign-btn',
            newGameBtn: 'new-game-btn',
            loadingOverlay: 'loading-overlay',
            notificationContainer: 'notification-container'
        };

        for (const [key, id] of Object.entries(elementIds)) {
            this.elements[key] = document.getElementById(id);
            if (!this.elements[key]) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    `UI element not found: ${id}`);
            }
        }
    }

    /**
     * Setup UI event listeners
     */
    setupEventListeners() {
        // Button event listeners
        if (this.elements.undoBtn) {
            this.elements.undoBtn.addEventListener('click', () => {
                this.handleUndoClick();
            });
        }

        if (this.elements.resignBtn) {
            this.elements.resignBtn.addEventListener('click', () => {
                this.handleResignClick();
            });
        }

        if (this.elements.newGameBtn) {
            this.elements.newGameBtn.addEventListener('click', () => {
                this.handleNewGameClick();
            });
        }

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            this.handleKeyboardInput(e);
        });

        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'UI event listeners setup complete');
    }

    /**
     * Handle game state changes
     */
    onGameStateChanged(gameState) {
        window.ChessDebug?.logUIEvent('game_state_changed', null, {
            turn: gameState.turn,
            moveCount: gameState.moveHistory?.length || 0
        });

        try {
            // Update game status
            this.updateGameStatus(gameState);
            
            // Update player information
            this.updatePlayerInfo(gameState);
            
            // Update move history
            this.updateMoveHistory(gameState);
            
            // Update button states
            this.updateButtonStates(gameState);

        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessRenderingError(
                    'Game state update failed',
                    { gameState, originalError: error }
                )
            );
        }
    }

    /**
     * Handle move attempts
     */
    onMoveAttempted(moveData) {
        window.ChessDebug?.logUIEvent('move_attempted', null, moveData);

        if (moveData.success) {
            this.showNotification('Move successful', 'success', 2000);
        } else {
            this.showNotification(moveData.error || 'Invalid move', 'error', 3000);
        }
    }

    /**
     * Update game status display
     */
    updateGameStatus(gameState) {
        if (!this.elements.gameStatusText) return;

        let statusText = 'Ready';
        let statusClass = 'connected';

        switch (gameState.status) {
            case window.ChessConstants.GAME_STATE.WAITING:
                statusText = 'Waiting for players...';
                statusClass = 'waiting';
                break;
            case window.ChessConstants.GAME_STATE.PLAYING:
                statusText = `${gameState.turn === 'white' ? 'White' : 'Black'} to move`;
                statusClass = 'connected';
                break;
            case window.ChessConstants.GAME_STATE.FINISHED:
                statusText = `Game Over - ${gameState.winner || 'Draw'}`;
                statusClass = 'waiting';
                break;
            case window.ChessConstants.GAME_STATE.ERROR:
                statusText = 'Game Error';
                statusClass = 'error';
                break;
        }

        this.elements.gameStatusText.textContent = statusText;
        this.updateConnectionStatus(statusClass);
    }

    /**
     * Update connection status indicator
     */
    updateConnectionStatus(status) {
        if (!this.elements.connectionStatus) return;

        // Remove existing status classes
        this.elements.connectionStatus.classList.remove('waiting', 'error');
        
        // Add new status class
        if (status !== 'connected') {
            this.elements.connectionStatus.classList.add(status);
        }
    }

    /**
     * Update player information
     */
    updatePlayerInfo(gameState) {
        // Update white player
        if (this.elements.whitePlayer) {
            const whitePlayer = gameState.players?.white || { name: 'Waiting...' };
            this.elements.whitePlayer.innerHTML = `<span>White: ${whitePlayer.name}</span>`;
            
            if (gameState.turn === 'white' && gameState.status === window.ChessConstants.GAME_STATE.PLAYING) {
                this.elements.whitePlayer.classList.add('active');
            } else {
                this.elements.whitePlayer.classList.remove('active');
            }
        }

        // Update black player
        if (this.elements.blackPlayer) {
            const blackPlayer = gameState.players?.black || { name: 'Waiting...' };
            this.elements.blackPlayer.innerHTML = `<span>Black: ${blackPlayer.name}</span>`;
            
            if (gameState.turn === 'black' && gameState.status === window.ChessConstants.GAME_STATE.PLAYING) {
                this.elements.blackPlayer.classList.add('active');
            } else {
                this.elements.blackPlayer.classList.remove('active');
            }
        }
    }

    /**
     * Update move history display
     */
    updateMoveHistory(gameState) {
        if (!this.elements.moveHistory) return;

        const moveHistory = gameState.moveHistory || [];
        
        if (moveHistory.length === 0) {
            this.elements.moveHistory.textContent = 'No moves yet';
        } else {
            // Show last few moves
            const lastMoves = moveHistory.slice(-3);
            const moveText = lastMoves
                .map((move, index) => `${Math.floor((moveHistory.length - lastMoves.length + index) / 2) + 1}. ${move.notation}`)
                .join(' ');
            this.elements.moveHistory.textContent = moveText;
        }
    }

    /**
     * Update button states based on game state
     */
    updateButtonStates(gameState) {
        // Undo button
        if (this.elements.undoBtn) {
            const canUndo = gameState.moveHistory && 
                           gameState.moveHistory.length > 0 && 
                           gameState.status === window.ChessConstants.GAME_STATE.PLAYING;
            this.elements.undoBtn.disabled = !canUndo;
        }

        // Resign button
        if (this.elements.resignBtn) {
            const canResign = gameState.status === window.ChessConstants.GAME_STATE.PLAYING;
            this.elements.resignBtn.disabled = !canResign;
        }

        // New game button is always enabled
        if (this.elements.newGameBtn) {
            this.elements.newGameBtn.disabled = false;
        }
    }

    /**
     * Handle undo button click
     */
    async handleUndoClick() {
        window.ChessDebug?.logUIEvent('undo_button_clicked', this.elements.undoBtn);

        try {
            if (this.gameManager && this.gameManager.undoLastMove) {
                await this.gameManager.undoLastMove();
            }
        } catch (error) {
            this.showNotification('Cannot undo move', 'error');
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessValidationError(
                    'Undo operation failed',
                    { originalError: error }
                )
            );
        }
    }

    /**
     * Handle resign button click
     */
    async handleResignClick() {
        window.ChessDebug?.logUIEvent('resign_button_clicked', this.elements.resignBtn);

        const confirmed = await this.showConfirmation(
            'Are you sure you want to resign?',
            () => this.performResignation(),
            () => {} // Cancel - do nothing
        );
    }

    /**
     * Perform resignation
     */
    async performResignation() {
        try {
            if (this.gameManager && this.gameManager.resignGame) {
                await this.gameManager.resignGame();
                this.showNotification('You have resigned', 'warning');
            }
        } catch (error) {
            this.showNotification('Failed to resign', 'error');
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessValidationError(
                    'Resignation failed',
                    { originalError: error }
                )
            );
        }
    }

    /**
     * Handle new game button click
     */
    async handleNewGameClick() {
        window.ChessDebug?.logUIEvent('new_game_button_clicked', this.elements.newGameBtn);

        try {
            this.updateLoadingState(true, 'Starting new game...');
            
            if (this.gameManager && this.gameManager.startNewGame) {
                await this.gameManager.startNewGame();
                this.showNotification('New game started', 'success');
            }
        } catch (error) {
            this.showNotification('Failed to start new game', 'error');
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessInitializationError(
                    'New game creation failed',
                    { originalError: error }
                )
            );
        } finally {
            this.updateLoadingState(false);
        }
    }

    /**
     * Handle keyboard input
     */
    handleKeyboardInput(event) {
        // Only handle keyboard shortcuts if not typing in an input field
        if (event.target.tagName === 'INPUT' || event.target.tagName === 'TEXTAREA') {
            return;
        }

        switch (event.key) {
            case 'u':
            case 'U':
                if (event.ctrlKey) {
                    event.preventDefault();
                    this.handleUndoClick();
                }
                break;
            case 'n':
            case 'N':
                if (event.ctrlKey) {
                    event.preventDefault();
                    this.handleNewGameClick();
                }
                break;
            case 'Escape':
                // Cancel any active operations
                if (this.boardRenderer && this.boardRenderer.cancelSelection) {
                    this.boardRenderer.cancelSelection();
                }
                break;
        }
    }

    /**
     * Update loading state
     */
    updateLoadingState(isLoading, message = 'Loading...') {
        this.isLoading = isLoading;

        if (this.elements.loadingOverlay) {
            if (isLoading) {
                this.elements.loadingOverlay.style.display = 'flex';
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Loading state activated', { message });
            } else {
                this.elements.loadingOverlay.style.display = 'none';
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Loading state deactivated');
            }
        }
    }

    /**
     * Show notification to user
     */
    showNotification(message, type = 'info', duration = 5000) {
        if (!this.elements.notificationContainer) {
            console.log(`Chess Plugin: ${message}`);
            return;
        }

        const notificationId = ++this.notificationIdCounter;
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.textContent = message;
        notification.id = `notification-${notificationId}`;

        // Add click to dismiss
        notification.addEventListener('click', () => {
            this.dismissNotification(notificationId);
        });

        this.elements.notificationContainer.appendChild(notification);
        this.activeNotifications.set(notificationId, notification);

        // Auto-dismiss after duration
        setTimeout(() => {
            this.dismissNotification(notificationId);
        }, duration);

        window.ChessDebug?.logUIEvent('notification_shown', notification, { message, type, duration });
    }

    /**
     * Dismiss notification
     */
    dismissNotification(notificationId) {
        const notification = this.activeNotifications.get(notificationId);
        if (notification && notification.parentNode) {
            notification.style.animation = 'slideOut 0.3s ease-out';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
                this.activeNotifications.delete(notificationId);
            }, 300);
        }
    }

    /**
     * Show confirmation dialog
     */
    async showConfirmation(message, onConfirm, onCancel) {
        // For now, use browser confirm - can be enhanced with custom modal later
        const confirmed = confirm(message);
        if (confirmed) {
            onConfirm();
        } else {
            onCancel();
        }
        return confirmed;
    }

    /**
     * Update display (called by lifecycle manager)
     */
    updateDisplay() {
        if (this.gameManager && this.gameManager.getCurrentGameState) {
            const gameState = this.gameManager.getCurrentGameState();
            if (gameState) {
                this.onGameStateChanged(gameState);
            }
        }
    }

    /**
     * Handle external interaction (called by OnlyOffice)
     */
    handleExternalInteraction() {
        window.ChessDebug?.logUIEvent('external_interaction', null);
        
        // Cancel any active selections or operations
        if (this.boardRenderer && this.boardRenderer.cancelSelection) {
            this.boardRenderer.cancelSelection();
        }
    }

    /**
     * Handle plugin resize
     */
    handleResize() {
        window.ChessDebug?.logUIEvent('plugin_resize', null);
        
        // Trigger board redraw if needed
        if (this.boardRenderer && this.boardRenderer.handleResize) {
            this.boardRenderer.handleResize();
        }
    }

    /**
     * Open settings (placeholder for future implementation)
     */
    openSettings() {
        this.showNotification('Settings functionality coming soon', 'info');
    }

    /**
     * Cleanup UI controller
     */
    async cleanup() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'UIController cleanup');

        // Clear active notifications
        this.activeNotifications.forEach((notification, id) => {
            this.dismissNotification(id);
        });

        // Remove event listeners
        if (this.elements.undoBtn) {
            this.elements.undoBtn.replaceWith(this.elements.undoBtn.cloneNode(true));
        }
        if (this.elements.resignBtn) {
            this.elements.resignBtn.replaceWith(this.elements.resignBtn.cloneNode(true));
        }
        if (this.elements.newGameBtn) {
            this.elements.newGameBtn.replaceWith(this.elements.newGameBtn.cloneNode(true));
        }

        this.isInitialized = false;
    }
}

// Export UI controller
window.ChessUIController = ChessUIController;