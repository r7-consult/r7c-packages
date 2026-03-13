/**
 * Game Debug System
 * Following OnlyOffice Plugin Development Standards
 * 
 * Generic debug system for casual games plugin
 */

class GameDebugSystem {
    constructor() {
        this.isEnabled = false;
        this.logLevel = 'INFO';
        this.categories = new Set();
        this.logHistory = [];
        this.maxHistorySize = 1000;
        this.startTime = Date.now();
        
        // Log levels in order of severity
        this.logLevels = {
            ERROR: 0,
            WARN: 1,
            INFO: 2,
            DEBUG: 3,
            TRACE: 4
        };
        
        this.initialize();
    }

    /**
     * Initialize debug system
     * Enable debugging if URL contains debug=true
     */
    initialize() {
        const urlParams = new URLSearchParams(window.location.search);
        this.isEnabled = urlParams.has('debug') || urlParams.has('debuggame');
        
        if (this.isEnabled) {
            this.logLevel = urlParams.get('loglevel') || 'DEBUG';
            console.log('%c[Game Debug] Debug mode enabled', 
                'color: #4CAF50; font-weight: bold;');
            this.setupPerformanceMonitoring();
        }

        // Enable all categories by default
        if (window.GameConstants && window.GameConstants.DEBUG_CATEGORIES) {
            Object.values(window.GameConstants.DEBUG_CATEGORIES).forEach(category => {
                this.categories.add(category);
            });
        }
    }

    /**
     * Setup performance monitoring
     */
    setupPerformanceMonitoring() {
        // Monitor long-running operations
        this.originalSetTimeout = window.setTimeout;
        
        window.setTimeout = (callback, delay, ...args) => {
            if (window.GameConstants && delay > window.GameConstants.PERFORMANCE.LOAD_WARNING_MS) {
                this.warn('PERFORMANCE', `Long timeout scheduled: ${delay}ms`);
            }
            return this.originalSetTimeout(callback, delay, ...args);
        };
    }

    /**
     * Check if logging is enabled for this level and category
     */
    shouldLog(level, category) {
        if (!this.isEnabled) return false;
        if (category && !this.categories.has(category)) return false;
        return this.logLevels[level] <= this.logLevels[this.logLevel];
    }

    /**
     * Format log message with timestamp and context
     */
    formatMessage(level, category, message, data = null) {
        const timestamp = Date.now() - this.startTime;
        const timeStr = `+${timestamp}ms`;
        const prefix = `[Game ${level}${category ? ':' + category : ''}] ${timeStr}`;
        
        return {
            prefix,
            message,
            data,
            timestamp: Date.now(),
            level,
            category
        };
    }

    /**
     * Add log entry to history
     */
    addToHistory(logEntry) {
        this.logHistory.push(logEntry);
        if (this.logHistory.length > this.maxHistorySize) {
            this.logHistory.shift();
        }
    }

    /**
     * Generic log method
     */
    log(level, category, message, data = null) {
        if (!this.shouldLog(level, category)) return;

        const logEntry = this.formatMessage(level, category, message, data);
        this.addToHistory(logEntry);

        const style = this.getConsoleStyle(level);
        
        if (data) {
            console.groupCollapsed(`%c${logEntry.prefix} ${message}`, style);
            console.log('Data:', data);
            console.trace('Stack trace');
            console.groupEnd();
        } else {
            console.log(`%c${logEntry.prefix} ${message}`, style);
        }
    }

    /**
     * Get console style for log level
     */
    getConsoleStyle(level) {
        const styles = {
            ERROR: 'color: #f44336; font-weight: bold;',
            WARN: 'color: #ff9800; font-weight: bold;',
            INFO: 'color: #2196f3;',
            DEBUG: 'color: #4caf50;',
            TRACE: 'color: #9e9e9e;'
        };
        return styles[level] || '';
    }

    /**
     * Specific log level methods
     */
    error(category, message, data = null) {
        this.log('ERROR', category, message, data);
    }

    warn(category, message, data = null) {
        this.log('WARN', category, message, data);
    }

    info(category, message, data = null) {
        this.log('INFO', category, message, data);
    }

    debug(category, message, data = null) {
        this.log('DEBUG', category, message, data);
    }

    trace(category, message, data = null) {
        this.log('TRACE', category, message, data);
    }

    /**
     * Performance timing methods
     */
    time(label, category = 'PERFORMANCE') {
        if (!this.isEnabled) return;
        
        const key = `${category}:${label}`;
        console.time(key);
        this.debug(category, `Timer started: ${label}`);
    }

    timeEnd(label, category = 'PERFORMANCE') {
        if (!this.isEnabled) return;
        
        const key = `${category}:${label}`;
        console.timeEnd(key);
        this.debug(category, `Timer ended: ${label}`);
    }

    /**
     * Memory usage logging
     */
    logMemoryUsage(category = 'PERFORMANCE') {
        if (!this.isEnabled || !window.performance.memory) return;
        
        const memory = window.performance.memory;
        this.debug(category, 'Memory usage', {
            used: `${Math.round(memory.usedJSHeapSize / 1024 / 1024)}MB`,
            total: `${Math.round(memory.totalJSHeapSize / 1024 / 1024)}MB`,
            limit: `${Math.round(memory.jsHeapSizeLimit / 1024 / 1024)}MB`
        });
    }

    /**
     * Game-specific logging methods
     */
    logGameMove(move, gameState) {
        this.debug('game_engine', 'Move made', {
            move,
            gameState: gameState ? {
                score: gameState.score,
                moveCount: gameState.movesCount || 0,
                gameOver: gameState.gameOver,
                gameWon: gameState.gameWon
            } : null
        });
    }

    logUIEvent(event, element, data = null) {
        this.trace('ui_events', `UI Event: ${event}`, {
            element: element ? (element.tagName || element.constructor?.name || 'unknown') : 'none',
            data
        });
    }

    logAPICall(method, params, duration = null) {
        this.debug('api_calls', `API: ${method}`, {
            params,
            duration: duration ? `${duration}ms` : 'pending'
        });
    }

    logLifecycleEvent(event, context = null) {
        this.info('lifecycle', `Lifecycle: ${event}`, context);
    }

    /**
     * Debug utilities
     */
    dumpGameState(gameState) {
        if (!this.isEnabled) return;
        
        console.group('%c[Game Debug] Game State Dump', 'color: #9c27b0; font-weight: bold;');
        console.log('Game State:', gameState);
        console.groupEnd();
    }

    /**
     * Export debug data
     */
    exportDebugData() {
        return {
            isEnabled: this.isEnabled,
            logLevel: this.logLevel,
            categories: Array.from(this.categories),
            logHistory: this.logHistory.slice(-100), // Last 100 entries
            memoryUsage: window.performance.memory ? {
                used: window.performance.memory.usedJSHeapSize,
                total: window.performance.memory.totalJSHeapSize,
                limit: window.performance.memory.jsHeapSizeLimit
            } : null,
            timestamp: new Date().toISOString()
        };
    }

    /**
     * Enable/disable specific categories
     */
    enableCategory(category) {
        this.categories.add(category);
        this.debug('DEBUG_SYSTEM', `Category enabled: ${category}`);
    }

    disableCategory(category) {
        this.categories.delete(category);
        this.debug('DEBUG_SYSTEM', `Category disabled: ${category}`);
    }

    /**
     * Clear log history
     */
    clearHistory() {
        this.logHistory = [];
        console.clear();
        this.info('DEBUG_SYSTEM', 'Debug history cleared');
    }

    /**
     * Get filtered log history
     */
    getLogHistory(level = null, category = null, limit = 50) {
        let filtered = this.logHistory;
        
        if (level) {
            filtered = filtered.filter(entry => entry.level === level);
        }
        
        if (category) {
            filtered = filtered.filter(entry => entry.category === category);
        }
        
        return filtered.slice(-limit);
    }
}

// Create global debug instance with multiple aliases for compatibility
window.GameDebug = new GameDebugSystem();
window.debug = window.GameDebug; // Short alias
window.ChessDebug = window.GameDebug; // Backward compatibility