/**
 * Input Validation Utilities
 * Provides secure validation for game data and user inputs
 */

class InputValidator {
    /**
     * Validate and sanitize game state object
     */
    static validateGameState(state) {
        if (!state || typeof state !== 'object') {
            throw new Error('Invalid game state: must be an object');
        }
        
        const validated = {};
        
        // Validate grid
        if (state.grid) {
            validated.grid = this.validateGrid(state.grid);
        }
        
        // Validate scores
        validated.score = this.validateNumber(state.score, 0, Number.MAX_SAFE_INTEGER, 0);
        validated.bestScore = this.validateNumber(state.bestScore, 0, Number.MAX_SAFE_INTEGER, 0);
        
        // Validate flags
        validated.gameOver = Boolean(state.gameOver);
        validated.gameWon = Boolean(state.gameWon);
        
        // Validate grid size
        validated.gridSize = this.validateNumber(state.gridSize, 3, 8, 4);
        
        // Validate winning tile
        validated.winningTile = this.validateNumber(state.winningTile, 8, 131072, 2048);
        
        // Validate move history (with size limit)
        if (state.moveHistory && Array.isArray(state.moveHistory)) {
            validated.moveHistory = this.validateMoveHistory(state.moveHistory);
        } else {
            validated.moveHistory = [];
        }
        
        // Validate moves count
        validated.movesCount = this.validateNumber(state.movesCount, 0, Number.MAX_SAFE_INTEGER, 0);
        
        return validated;
    }
    
    /**
     * Validate grid structure
     */
    static validateGrid(grid) {
        if (!Array.isArray(grid)) {
            throw new Error('Grid must be an array');
        }
        
        const size = grid.length;
        if (size < 3 || size > 8) {
            throw new Error('Grid size must be between 3 and 8');
        }
        
        const validatedGrid = [];
        
        for (let row = 0; row < size; row++) {
            if (!Array.isArray(grid[row]) || grid[row].length !== size) {
                throw new Error(`Invalid row ${row}: must be array of length ${size}`);
            }
            
            validatedGrid[row] = [];
            
            for (let col = 0; col < size; col++) {
                const value = grid[row][col];
                
                // Validate tile value (must be 0 or power of 2)
                if (value === 0 || value === null || value === undefined) {
                    validatedGrid[row][col] = 0;
                } else {
                    const num = Number(value);
                    
                    if (!Number.isInteger(num) || num < 0 || num > 131072) {
                        throw new Error(`Invalid tile value at [${row}][${col}]: ${value}`);
                    }
                    
                    // Check if power of 2
                    if (num > 0 && (num & (num - 1)) !== 0) {
                        throw new Error(`Tile value must be power of 2: ${num}`);
                    }
                    
                    validatedGrid[row][col] = num;
                }
            }
        }
        
        return validatedGrid;
    }
    
    /**
     * Validate number within range
     */
    static validateNumber(value, min, max, defaultValue) {
        const num = Number(value);
        
        if (isNaN(num) || !isFinite(num)) {
            return defaultValue;
        }
        
        if (num < min) return min;
        if (num > max) return max;
        
        return Math.floor(num);
    }
    
    /**
     * Validate move direction
     */
    static validateDirection(direction) {
        const validDirections = ['UP', 'DOWN', 'LEFT', 'RIGHT'];
        
        if (typeof direction !== 'string') {
            throw new Error('Direction must be a string');
        }
        
        const upperDirection = direction.toUpperCase().trim();
        
        if (!validDirections.includes(upperDirection)) {
            throw new Error(`Invalid direction: ${direction}. Must be one of: ${validDirections.join(', ')}`);
        }
        
        return upperDirection;
    }
    
    /**
     * Validate move history (limit size to prevent memory issues)
     */
    static validateMoveHistory(history, maxSize = 100) {
        if (!Array.isArray(history)) {
            return [];
        }
        
        // Limit history size
        const limitedHistory = history.slice(-maxSize);
        
        const validated = [];
        
        for (const move of limitedHistory) {
            if (move && typeof move === 'object') {
                try {
                    const validMove = {
                        direction: this.validateDirection(move.direction),
                        timestamp: this.validateTimestamp(move.timestamp),
                        score: this.validateNumber(move.score, 0, Number.MAX_SAFE_INTEGER, 0)
                    };
                    
                    // Only store essential data to save memory
                    if (move.previousState && move.previousState.grid) {
                        validMove.previousState = {
                            grid: this.validateGrid(move.previousState.grid),
                            score: this.validateNumber(move.previousState.score, 0, Number.MAX_SAFE_INTEGER, 0)
                        };
                    }
                    
                    validated.push(validMove);
                } catch (error) {
                    window.debug?.warn('InputValidator', 'Invalid move in history', error);
                    // Skip invalid moves
                }
            }
        }
        
        return validated;
    }
    
    /**
     * Validate timestamp
     */
    static validateTimestamp(timestamp) {
        const ts = Number(timestamp);
        
        if (isNaN(ts) || ts < 0 || ts > Date.now() + 86400000) { // Max 1 day in future
            return Date.now();
        }
        
        return Math.floor(ts);
    }
    
    /**
     * Sanitize string input
     */
    static sanitizeString(input, maxLength = 100) {
        if (typeof input !== 'string') {
            return '';
        }
        
        // Remove control characters and limit length
        return input
            .replace(/[\x00-\x1F\x7F]/g, '') // Remove control characters
            .substring(0, maxLength)
            .trim();
    }
    
    /**
     * Validate plugin data from OnlyOffice
     */
    static validatePluginData(data) {
        if (!data) {
            return null;
        }
        
        try {
            // First, check if it's a string that needs parsing
            let parsed = data;
            
            if (typeof data === 'string') {
                // Limit size to prevent DOS
                if (data.length > 1000000) { // 1MB limit
                    throw new Error('Data too large');
                }
                
                parsed = JSON.parse(data);
            }
            
            // Validate the parsed data
            if (parsed.gameType) {
                parsed.gameType = this.sanitizeString(parsed.gameType, 50);
            }
            
            if (parsed.gameState) {
                parsed.gameState = this.validateGameState(parsed.gameState);
            }
            
            return parsed;
            
        } catch (error) {
            window.debug?.error('InputValidator', 'Failed to validate plugin data', error);
            return null;
        }
    }
    
    /**
     * Validate configuration object
     */
    static validateConfig(config) {
        if (!config || typeof config !== 'object') {
            return {};
        }
        
        const validated = {};
        
        // Validate boolean flags
        validated.allowUndo = Boolean(config.allowUndo);
        validated.autoSave = Boolean(config.autoSave);
        validated.supportsAI = Boolean(config.supportsAI);
        validated.supportsCollaboration = Boolean(config.supportsCollaboration);
        
        // Validate numbers
        validated.gridSize = this.validateNumber(config.gridSize, 3, 8, 4);
        validated.maxPlayers = this.validateNumber(config.maxPlayers, 1, 4, 1);
        validated.minPlayers = this.validateNumber(config.minPlayers, 1, 4, 1);
        
        return validated;
    }
}

// Export for use
window.InputValidator = InputValidator;