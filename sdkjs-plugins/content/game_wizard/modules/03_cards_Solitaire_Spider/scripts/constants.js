/**
 * Casual Games Plugin Constants
 * Following OnlyOffice Plugin Development Standards
 * 
 * Game-agnostic constants for the casual games framework
 */

window.GameConstants = {
    // Plugin Configuration
    PLUGIN: {
        NAME: '2048 Game',
        VERSION: '1.0.0',
        GUID: 'asc.{FFE1F462-1EA2-4391-990D-4CC84940B754}',
        OLE_DATA_TYPE: 'ole'
    },

    // UI Configuration
    UI: {
        ANIMATION_DURATION: 150,
        THEMES: {
            LIGHT: 'light',
            DARK: 'dark'
        }
    },

    // Game States
    GAME_STATE: {
        WAITING: 'waiting',
        PLAYING: 'playing',
        PAUSED: 'paused',
        FINISHED: 'finished',
        ERROR: 'error'
    },

    // Error Categories
    ERROR_TYPES: {
        INITIALIZATION: 'initialization_error',
        VALIDATION: 'validation_error',
        RENDERING: 'rendering_error',
        PERSISTENCE: 'persistence_error'
    },

    // Debug Categories
    DEBUG_CATEGORIES: {
        LIFECYCLE: 'lifecycle',
        GAME_ENGINE: 'game_engine',
        UI_EVENTS: 'ui_events',
        API_CALLS: 'api_calls'
    },

    // Performance Thresholds
    PERFORMANCE: {
        LOAD_WARNING_MS: 1000,
        MOVE_WARNING_MS: 500,
        RENDER_WARNING_MS: 100,
        MAX_HISTORY_SIZE: 1000
    },

    // OnlyOffice Integration
    ONLYOFFICE: {
        BUTTON_IDS: {
            CANCEL: -1,
            OK: 0,
            SETTINGS: 1
        },
        COMMAND_TYPES: {
            INSERT: 'insert',
            CLOSE: 'close',
            RESIZE: 'resize'
        }
    },

    // Storage Keys
    STORAGE: {
        GAME_STATE: 'game_state',
        PLAYER_PREFERENCES: 'player_prefs',
        HIGH_SCORES: 'high_scores'
    },

    // Network Configuration
    NETWORK: {
        TIMEOUT_MS: 5000,
        RETRY_ATTEMPTS: 3,
        HEARTBEAT_INTERVAL_MS: 30000
    }
};

// Freeze constants to prevent modification
Object.freeze(window.GameConstants);

// Maintain backward compatibility temporarily
window.ChessConstants = window.GameConstants;