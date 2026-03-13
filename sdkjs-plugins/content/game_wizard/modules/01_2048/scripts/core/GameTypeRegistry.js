/**
 * @fileoverview Game Type Registry
 * @description Central registry for managing different game type implementations
 * @see {@link _coding_standard/01_plugin_architecture_guide.md} Architecture Guide - Registry Pattern
 * @see {@link _coding_standard/02_api_reference_patterns.md} API Patterns
 * @author Casual Games Plugin Development Team
 * @version 1.0.0
 * @since 1.0.0
 */

// =============================================================================
// 1. IMPORTS AND DEPENDENCIES
// =============================================================================
// GameManagerBase will be imported by concrete implementations

// =============================================================================
// 2. CONSTANTS AND CONFIGURATION  
// =============================================================================
const GAME_CATEGORIES = Object.freeze({
    STRATEGY: 'strategy',
    PUZZLE: 'puzzle', 
    CARD: 'card',
    ARCADE: 'arcade',
    EDUCATIONAL: 'educational',
    GENERAL: 'general'
});

const DIFFICULTY_LEVELS = Object.freeze({
    EASY: 'easy',
    MEDIUM: 'medium',
    HARD: 'hard'
});

// =============================================================================
// 3. GAME TYPE REGISTRY CLASS
// =============================================================================

/**
 * Central registry for managing game types
 * Follows Singleton pattern to ensure single source of truth
 * 
 * ONLYOFFICE INTEGRATION NOTES:
 * - Game types should integrate with OnlyOffice plugin lifecycle
 * - Metadata used for plugin UI generation and game selection
 * - Registry enables dynamic loading of game implementations
 */
class GameTypeRegistry {
    #gameTypes = new Map();
    #defaultGameType = 'chess'; // Fallback to chess implementation
    #initialized = false;
    
    constructor() {
        window.debug?.debug('GameTypeRegistry', 'Initializing game type registry');
    }

    // =============================================================================
    // GAME TYPE REGISTRATION
    // =============================================================================
    
    /**
     * Register a new game type with metadata and validation
     * @param {string} gameType - Unique game type identifier
     * @param {Class} gameManagerClass - Game manager implementation class (extends GameManagerBase)
     * @param {Object} metadata - Game metadata for UI and configuration
     * @throws {ValidationError} When registration data is invalid
     */
    registerGameType(gameType, gameManagerClass, metadata = {}) {
        // Validate game type identifier
        if (!gameType || typeof gameType !== 'string' || gameType.trim().length === 0) {
            throw new Error(`Game type must be a non-empty string, got: ${gameType}`);
        }
        
        // Validate game manager class
        if (!gameManagerClass || typeof gameManagerClass !== 'function') {
            throw new Error(`Game manager class must be a constructor function for: ${gameType}`);
        }
        
        // Check if already registered
        if (this.#gameTypes.has(gameType)) {
            window.debug?.warn('GameTypeRegistry', `Game type '${gameType}' already registered, overriding`);
        }
        
        // Create validated game registration
        const gameRegistration = Object.freeze({
            gameType: gameType.toLowerCase().trim(),
            gameManagerClass,
            metadata: Object.freeze({
                // Basic metadata
                name: metadata.name || this.#formatGameTypeName(gameType),
                description: metadata.description || `${gameType} game`,
                version: metadata.version || '1.0.0',
                
                // Game classification
                category: this.#validateCategory(metadata.category),
                difficulty: this.#validateDifficulty(metadata.difficulty),
                
                // Player configuration
                minPlayers: Math.max(1, parseInt(metadata.minPlayers) || 1),
                maxPlayers: Math.max(1, parseInt(metadata.maxPlayers) || 2),
                
                // Features
                supportsAI: Boolean(metadata.supportsAI),
                supportsCollaboration: Boolean(metadata.supportsCollaboration),
                requiresOnlyOfficeAPI: Boolean(metadata.requiresOnlyOfficeAPI),
                
                // UI metadata
                iconPath: metadata.iconPath || '../../resources/icons/main/light/icon.png',
                thumbnailPath: metadata.thumbnailPath || null,
                
                // Technical metadata
                estimatedLoadTime: parseInt(metadata.estimatedLoadTime) || 1000,
                memoryUsage: metadata.memoryUsage || 'low',
                
                // Extended metadata
                ...metadata
            }),
            registeredAt: new Date().toISOString(),
            registeredBy: 'GameTypeRegistry'
        });
        
        // Store registration
        this.#gameTypes.set(gameType.toLowerCase().trim(), gameRegistration);
        
        window.debug?.info('GameTypeRegistry', 'Registered game type', {
            gameType: gameType.toLowerCase().trim(),
            managerClass: gameManagerClass.name,
            metadata: {
                name: gameRegistration.metadata.name,
                category: gameRegistration.metadata.category,
                players: `${gameRegistration.metadata.minPlayers}-${gameRegistration.metadata.maxPlayers}`,
                features: {
                    ai: gameRegistration.metadata.supportsAI,
                    collaboration: gameRegistration.metadata.supportsCollaboration
                }
            }
        });
        
        return true;
    }
    
    /**
     * Unregister a game type
     */
    unregisterGameType(gameType) {
        const normalizedType = gameType.toLowerCase().trim();
        const existed = this.#gameTypes.delete(normalizedType);
        
        if (existed) {
            window.debug?.info('GameTypeRegistry', `Unregistered game type: ${normalizedType}`);
        } else {
            window.debug?.warn('GameTypeRegistry', `Attempted to unregister non-existent game type: ${normalizedType}`);
        }
        
        return existed;
    }

    // =============================================================================
    // GAME TYPE QUERY METHODS
    // =============================================================================
    
    /**
     * Get all available game types
     * @returns {string[]} Array of registered game type identifiers
     */
    getAvailableGameTypes() {
        return Array.from(this.#gameTypes.keys()).sort();
    }
    
    /**
     * Get game types by category
     */
    getGameTypesByCategory(category) {
        const normalizedCategory = category.toLowerCase();
        return Array.from(this.#gameTypes.values())
            .filter(registration => registration.metadata.category === normalizedCategory)
            .map(registration => registration.gameType)
            .sort();
    }
    
    /**
     * Get game metadata
     */
    getGameMetadata(gameType) {
        const registration = this.#gameTypes.get(gameType.toLowerCase().trim());
        return registration ? registration.metadata : null;
    }
    
    /**
     * Get all game metadata for UI generation
     */
    getAllGameMetadata() {
        const metadata = [];
        
        for (const [gameType, registration] of this.#gameTypes.entries()) {
            metadata.push({
                gameType,
                ...registration.metadata,
                registeredAt: registration.registeredAt
            });
        }
        
        return metadata.sort((a, b) => a.name.localeCompare(b.name));
    }
    
    /**
     * Check if game type is registered
     */
    isGameTypeRegistered(gameType) {
        return this.#gameTypes.has(gameType.toLowerCase().trim());
    }

    // =============================================================================
    // GAME MANAGER CREATION
    // =============================================================================
    
    /**
     * Create game manager instance for specified game type
     * @param {string} gameType - Game type identifier
     * @param {Object} config - Game-specific configuration
     * @returns {GameManagerBase} Game manager instance
     * @throws {Error} When game type not registered or creation fails
     */
    createGameManager(gameType, config = {}) {
        const normalizedType = gameType.toLowerCase().trim();
        const registration = this.#gameTypes.get(normalizedType);
        
        if (!registration) {
            const available = this.getAvailableGameTypes();
            throw new Error(
                `Game type '${gameType}' not registered. Available types: ${available.join(', ')}`
            );
        }
        
        try {
            window.debug?.debug('GameTypeRegistry', 'Creating game manager', {
                gameType: normalizedType,
                managerClass: registration.gameManagerClass.name,
                config
            });
            
            // Create instance with merged configuration
            const mergedConfig = {
                // Default configuration from metadata
                minPlayers: registration.metadata.minPlayers,
                maxPlayers: registration.metadata.maxPlayers,
                supportsAI: registration.metadata.supportsAI,
                supportsCollaboration: registration.metadata.supportsCollaboration,
                
                // User-provided configuration
                ...config
            };
            
            const gameManager = new registration.gameManagerClass(mergedConfig);
            
            // Validate that created instance extends GameManagerBase
            // NOTE: In production environment, this check ensures proper inheritance
            if (typeof gameManager.initializeGame !== 'function') {
                throw new Error(
                    `Game manager for '${gameType}' does not implement required GameManagerBase interface`
                );
            }
            
            window.debug?.info('GameTypeRegistry', 'Game manager created successfully', {
                gameType: normalizedType,
                managerClass: registration.gameManagerClass.name
            });
            
            return gameManager;
            
        } catch (error) {
            window.debug?.error('GameTypeRegistry', 'Failed to create game manager', {
                gameType: normalizedType,
                error: error.message
            });
            
            throw new Error(
                `Failed to create game manager for '${gameType}': ${error.message}`
            );
        }
    }
    
    // =============================================================================
    // REGISTRY MANAGEMENT
    // =============================================================================
    
    /**
     * Initialize registry with default game types
     * ONLYOFFICE NOTE: Should be called after plugin initialization completes
     */
    async initialize() {
        if (this.#initialized) return;
        
        try {
            window.debug?.info('GameTypeRegistry', 'Initializing registry');
            
            // Registry is now ready for game type registration
            // Individual games will register themselves when their scripts load
            
            this.#initialized = true;
            
            window.debug?.info('GameTypeRegistry', 'Registry initialization completed', {
                registeredTypes: this.getAvailableGameTypes().length
            });
            
        } catch (error) {
            window.debug?.error('GameTypeRegistry', 'Registry initialization failed', error);
            throw error;
        }
    }
    
    /**
     * Get registry statistics
     */
    getRegistryStats() {
        const stats = {
            totalGames: this.#gameTypes.size,
            categories: {},
            difficulties: {},
            features: {
                aiSupport: 0,
                collaboration: 0,
                onlyOfficeAPI: 0
            }
        };
        
        for (const registration of this.#gameTypes.values()) {
            const meta = registration.metadata;
            
            // Count by category
            stats.categories[meta.category] = (stats.categories[meta.category] || 0) + 1;
            
            // Count by difficulty
            stats.difficulties[meta.difficulty] = (stats.difficulties[meta.difficulty] || 0) + 1;
            
            // Count features
            if (meta.supportsAI) stats.features.aiSupport++;
            if (meta.supportsCollaboration) stats.features.collaboration++;
            if (meta.requiresOnlyOfficeAPI) stats.features.onlyOfficeAPI++;
        }
        
        return stats;
    }
    
    /**
     * Clear all registered game types (for testing)
     */
    clearRegistry() {
        window.debug?.warn('GameTypeRegistry', 'Clearing all registered game types');
        this.#gameTypes.clear();
    }

    // =============================================================================
    // PRIVATE VALIDATION METHODS
    // =============================================================================
    
    /**
     * Validate and normalize game category
     */
    #validateCategory(category) {
        if (!category) return GAME_CATEGORIES.GENERAL;
        
        const normalizedCategory = category.toLowerCase();
        const validCategories = Object.values(GAME_CATEGORIES);
        
        return validCategories.includes(normalizedCategory) 
            ? normalizedCategory 
            : GAME_CATEGORIES.GENERAL;
    }
    
    /**
     * Validate and normalize difficulty level
     */
    #validateDifficulty(difficulty) {
        if (!difficulty) return DIFFICULTY_LEVELS.MEDIUM;
        
        const normalizedDifficulty = difficulty.toLowerCase();
        const validDifficulties = Object.values(DIFFICULTY_LEVELS);
        
        return validDifficulties.includes(normalizedDifficulty)
            ? normalizedDifficulty
            : DIFFICULTY_LEVELS.MEDIUM;
    }
    
    /**
     * Format game type name for display
     */
    #formatGameTypeName(gameType) {
        return gameType
            .split(/[\s_-]+/)
            .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
            .join(' ');
    }
}

// =============================================================================
// 4. SINGLETON INSTANCE AND GLOBAL ACCESS
// =============================================================================

/**
 * Global singleton instance for game type registry
 * ONLYOFFICE INTEGRATION: Accessible globally for plugin components
 */
const gameTypeRegistry = new GameTypeRegistry();

// Make registry globally accessible
window.GameTypeRegistry = GameTypeRegistry;
window.gameTypeRegistry = gameTypeRegistry;
window.GAME_CATEGORIES = GAME_CATEGORIES;
window.DIFFICULTY_LEVELS = DIFFICULTY_LEVELS;

// =============================================================================
// 5. INITIALIZATION
// =============================================================================

// Initialization is handled by unified-init.js

// =============================================================================
// 6. EXPORTS  
// =============================================================================
window.gameTypeRegistry = gameTypeRegistry;
