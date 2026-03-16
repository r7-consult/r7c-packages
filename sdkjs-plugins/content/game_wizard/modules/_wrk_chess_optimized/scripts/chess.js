/**
 * Chess.js - Stub file for backward compatibility
 * The actual chess functionality is now in the modular architecture
 */

// Create minimal global objects for backward compatibility
window.CChessBoard = null;
window.g_board = null;

// Create ChessBoard wrapper to indicate the new system is ready
window.ChessBoard = {
    isReady: function() {
        // Return true when the new architecture is loaded
        return !!(window.ChessLifecycleManager && window.SimpleBoardRenderer);
    },
    
    init: function(data) {
        console.log('Legacy chess.js init called - using new architecture');
        // The new architecture handles initialization
        return true;
    },
    
    instance: {
        initialized: false
    }
};

console.log('Chess.js stub loaded - using new modular architecture');