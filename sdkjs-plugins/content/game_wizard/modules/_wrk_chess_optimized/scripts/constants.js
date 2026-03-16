/**
 * Chess Plugin Constants
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/CODING_STANDARD.md
 */

window.ChessConstants = {
    // Plugin Configuration
    PLUGIN: {
        NAME: 'Chess',
        VERSION: '1.0.1',
        GUID: 'asc.{FFE1F462-1EA2-4391-990D-4CC84940B754}',
        OLE_DATA_TYPE: 'ole'
    },

    // UI Configuration
    UI: {
        BOARD_SIZE: 480,
        CELL_SIZE: 60,
        ANIMATION_DURATION: 300,
        THEMES: {
            LIGHT: 'light',
            DARK: 'dark'
        }
    },

    // Chess Game States
    GAME_STATE: {
        WAITING: 'waiting',
        PLAYING: 'playing',
        PAUSED: 'paused',
        FINISHED: 'finished',
        ERROR: 'error'
    },

    // Player Colors
    PLAYER: {
        WHITE: 'white',
        BLACK: 'black',
        SPECTATOR: 'spectator'
    },

    // Chess Piece Types
    PIECES: {
        PAWN: 'pawn',
        ROOK: 'rook',
        KNIGHT: 'knight',
        BISHOP: 'bishop',
        QUEEN: 'queen',
        KING: 'king'
    },

    // Move Types
    MOVE_TYPE: {
        NORMAL: 'normal',
        CASTLE: 'castle',
        EN_PASSANT: 'en_passant',
        PROMOTION: 'promotion'
    },

    // Game End Conditions
    END_CONDITION: {
        CHECKMATE: 'checkmate',
        STALEMATE: 'stalemate',
        DRAW: 'draw',
        RESIGNATION: 'resignation',
        TIMEOUT: 'timeout'
    },

    // Collaboration Events
    COLLAB_EVENTS: {
        PLAYER_JOINED: 'player_joined',
        PLAYER_LEFT: 'player_left',
        MOVE_MADE: 'move_made',
        GAME_STATE_CHANGED: 'game_state_changed',
        CHAT_MESSAGE: 'chat_message'
    },

    // Error Categories
    ERROR_TYPES: {
        INITIALIZATION: 'initialization_error',
        VALIDATION: 'validation_error',
        COLLABORATION: 'collaboration_error',
        RENDERING: 'rendering_error',
        PERSISTENCE: 'persistence_error'
    },

    // Debug Categories
    DEBUG_CATEGORIES: {
        LIFECYCLE: 'lifecycle',
        CHESS_ENGINE: 'chess_engine',
        COLLABORATION: 'collaboration',
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
        GAME_STATE: 'chess_game_state',
        PLAYER_PREFERENCES: 'chess_player_prefs',
        COLLABORATION_DATA: 'chess_collab_data'
    },

    // Network Configuration
    NETWORK: {
        TIMEOUT_MS: 5000,
        RETRY_ATTEMPTS: 3,
        HEARTBEAT_INTERVAL_MS: 30000
    }
};

// Freeze constants to prevent modification
Object.freeze(window.ChessConstants);