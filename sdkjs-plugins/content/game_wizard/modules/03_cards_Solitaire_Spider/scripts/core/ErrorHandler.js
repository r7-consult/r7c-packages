/**
 * Chess Plugin Error Handling System
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/02_api_reference_patterns.md#error-handling-patterns
 */

/**
 * Base Chess Plugin Error Class
 */
class ChessPluginError extends Error {
    constructor(message, category, context = {}) {
        super(message);
        this.name = 'ChessPluginError';
        this.category = category;
        this.context = context;
        this.timestamp = new Date().toISOString();
        this.pluginVersion = window.ChessConstants?.PLUGIN?.VERSION || 'unknown';
    }

    toLogObject() {
        return {
            name: this.name,
            message: this.message,
            category: this.category,
            context: this.context,
            timestamp: this.timestamp,
            pluginVersion: this.pluginVersion,
            stack: this.stack
        };
    }
}

/**
 * Chess-specific Error Classes
 */
class ChessInitializationError extends ChessPluginError {
    constructor(message, context = {}) {
        super(message, window.ChessConstants.ERROR_TYPES.INITIALIZATION, context);
        this.name = 'ChessInitializationError';
    }
}

class ChessValidationError extends ChessPluginError {
    constructor(message, context = {}) {
        super(message, window.ChessConstants.ERROR_TYPES.VALIDATION, context);
        this.name = 'ChessValidationError';
    }
}

class ChessCollaborationError extends ChessPluginError {
    constructor(message, context = {}) {
        super(message, window.ChessConstants.ERROR_TYPES.COLLABORATION, context);
        this.name = 'ChessCollaborationError';
    }
}

class ChessRenderingError extends ChessPluginError {
    constructor(message, context = {}) {
        super(message, window.ChessConstants.ERROR_TYPES.RENDERING, context);
        this.name = 'ChessRenderingError';
    }
}

class ChessPersistenceError extends ChessPluginError {
    constructor(message, context = {}) {
        super(message, window.ChessConstants.ERROR_TYPES.PERSISTENCE, context);
        this.name = 'ChessPersistenceError';
    }
}

/**
 * Centralized Error Handler
 */
class ChessErrorHandler {
    constructor() {
        this.errorQueue = [];
        this.maxErrorQueueSize = 50;
        this.isInitialized = false;
        this.userNotificationCallback = null;
    }

    /**
     * Initialize error handler with user notification callback
     */
    initialize(userNotificationCallback = null) {
        this.userNotificationCallback = userNotificationCallback;
        this.isInitialized = true;

        // Set up global error handlers
        window.addEventListener('error', (event) => {
            this.handleError(new ChessPluginError(
                event.message || 'Unknown global error',
                window.ChessConstants.ERROR_TYPES.INITIALIZATION,
                {
                    filename: event.filename,
                    lineno: event.lineno,
                    colno: event.colno
                }
            ));
        });

        window.addEventListener('unhandledrejection', (event) => {
            this.handleError(new ChessPluginError(
                'Unhandled promise rejection: ' + (event.reason?.message || event.reason),
                window.ChessConstants.ERROR_TYPES.INITIALIZATION,
                { reason: event.reason }
            ));
        });
    }

    /**
     * Main error handling method
     */
    handleError(error, shouldNotifyUser = true) {
        if (!this.isInitialized) {
            console.error('ChessErrorHandler not initialized:', error);
            return;
        }

        // Add to error queue
        this.addToErrorQueue(error);

        // Log to console and debug system
        this.logError(error);

        // Notify user if appropriate
        if (shouldNotifyUser && this.shouldNotifyUser(error)) {
            this.notifyUser(error);
        }

        // Trigger error recovery if needed
        this.attemptRecovery(error);
    }

    /**
     * Add error to internal queue for debugging
     */
    addToErrorQueue(error) {
        this.errorQueue.push(error);
        
        // Maintain queue size
        if (this.errorQueue.length > this.maxErrorQueueSize) {
            this.errorQueue.shift();
        }
    }

    /**
     * Log error to debug system
     */
    logError(error) {
        if (window.ChessDebug) {
            window.ChessDebug.error('ErrorHandler', error.toLogObject ? error.toLogObject() : error);
        } else {
            console.error('Chess Plugin Error:', error);
        }
    }

    /**
     * Determine if user should be notified
     */
    shouldNotifyUser(error) {
        // Don't spam user with multiple similar errors
        const recentSimilarErrors = this.errorQueue
            .slice(-5)
            .filter(e => e.category === error.category && e.name === error.name);
        
        return recentSimilarErrors.length <= 2;
    }

    /**
     * Notify user with friendly message
     */
    notifyUser(error) {
        if (this.userNotificationCallback) {
            const userFriendlyMessage = this.getUserFriendlyMessage(error);
            this.userNotificationCallback(userFriendlyMessage, error.category);
        }
    }

    /**
     * Convert technical error to user-friendly message
     */
    getUserFriendlyMessage(error) {
        const messages = {
            [window.ChessConstants.ERROR_TYPES.INITIALIZATION]: 
                'The chess game failed to start properly. Please try refreshing the plugin.',
            [window.ChessConstants.ERROR_TYPES.VALIDATION]: 
                'Invalid move detected. Please check your selection and try again.',
            [window.ChessConstants.ERROR_TYPES.COLLABORATION]: 
                'Connection issue with other players. Attempting to reconnect...',
            [window.ChessConstants.ERROR_TYPES.RENDERING]: 
                'Display issue detected. The board may not appear correctly.',
            [window.ChessConstants.ERROR_TYPES.PERSISTENCE]: 
                'Unable to save game progress. Changes may be lost.'
        };

        return messages[error.category] || 'An unexpected error occurred. Please try again.';
    }

    /**
     * Attempt automatic recovery for certain error types
     */
    attemptRecovery(error) {
        switch (error.category) {
            case window.ChessConstants.ERROR_TYPES.RENDERING:
                this.attemptRenderRecovery();
                break;
            case window.ChessConstants.ERROR_TYPES.COLLABORATION:
                this.attemptCollaborationRecovery();
                break;
            case window.ChessConstants.ERROR_TYPES.PERSISTENCE:
                this.attemptStorageRecovery();
                break;
        }
    }

    /**
     * Recovery methods
     */
    attemptRenderRecovery() {
        setTimeout(() => {
            if (window.ChessBoardRenderer) {
                try {
                    window.ChessBoardRenderer.forceRedraw();
                } catch (e) {
                    this.handleError(new ChessRenderingError('Render recovery failed', { originalError: e }));
                }
            }
        }, 100);
    }

    attemptCollaborationRecovery() {
        if (window.ChessCollaborationManager) {
            try {
                window.ChessCollaborationManager.reconnect();
            } catch (e) {
                this.handleError(new ChessCollaborationError('Collaboration recovery failed', { originalError: e }));
            }
        }
    }

    attemptStorageRecovery() {
        try {
            // Clear potentially corrupted storage
            localStorage.removeItem(window.ChessConstants.STORAGE.GAME_STATE);
            localStorage.removeItem(window.ChessConstants.STORAGE.COLLABORATION_DATA);
        } catch (e) {
            // Storage might be completely unavailable
            this.handleError(new ChessPersistenceError('Storage recovery failed', { originalError: e }));
        }
    }

    /**
     * Get recent errors for debugging
     */
    getRecentErrors(count = 10) {
        return this.errorQueue.slice(-count);
    }

    /**
     * Clear error queue
     */
    clearErrorQueue() {
        this.errorQueue = [];
    }

    /**
     * Create validation error with context
     */
    createValidationError(field, value, expectedType) {
        return new ChessValidationError(
            `Invalid ${field}: expected ${expectedType}, got ${typeof value}`,
            { field, value, expectedType }
        );
    }

    /**
     * Wrap async operations with error handling
     */
    async wrapAsync(operation, errorContext = {}) {
        try {
            return await operation();
        } catch (error) {
            const wrappedError = error instanceof ChessPluginError 
                ? error 
                : new ChessPluginError(error.message, 'unknown', { ...errorContext, originalError: error });
            
            this.handleError(wrappedError);
            throw wrappedError;
        }
    }
}

// Export error classes and handler
window.ChessErrors = {
    ChessPluginError,
    ChessInitializationError,
    ChessValidationError,
    ChessCollaborationError,
    ChessRenderingError,
    ChessPersistenceError,
    ChessErrorHandler
};

// Create global error handler instance
window.ChessErrorHandler = new ChessErrorHandler();