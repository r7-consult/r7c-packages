/**
 * @summary Main entry point for the FreeCell Solitaire plugin.
 * @description Initializes the game and handles OnlyOffice plugin integration.
 * @version 1.0.0
 * @author Roo
 * 
 * @knowledge_map
 * - Link to Plugin Lifecycle: _coding_standard/01_plugin_architecture_guide.md#plugin-lifecycle
 */
(function() {
    'use strict';

    let game;

    // A flag to ensure initialization happens only once.
    let isInitialized = false;

    /**
     * Initializes the game instance and UI.
     * This function is called once the DOM is fully loaded.
     */
    function initializeGame() {
        if (isInitialized) {
            console.warn('[Main] Game is already initialized.');
            return;
        }

        try {
            // Create FreeCell game manager
            game = new FreeCellManager();
            game.initialize();
            game.startNewGame();
            
            // Make globally accessible for debugging
            window.freeCellManager = game;
            
            // Setup global event listeners for buttons
            document.getElementById('new-game-btn').addEventListener('click', () => game.startNewGame());
            document.getElementById('btn-undo').addEventListener('click', () => game.undo());
            
            // Add hint button if available
            const hintBtn = document.getElementById('btn-hint');
            if (hintBtn) {
                hintBtn.addEventListener('click', () => game.getHint());
            }

            isInitialized = true;
            console.log('[Main] FreeCell game initialized and started.');

        } catch (error) {
            console.error('[Main] Failed to initialize FreeCell game:', error);
            // Optionally, display a user-friendly error message on the screen
        }
    }

    // Standard DOMContentLoaded listener to start the game
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeGame);
    } else {
        // The DOM is already ready
        initializeGame();
    }

    // =========================================================================
    // ONLYOFFICE PLUGIN INTEGRATION
    // =========================================================================

    /**
     * Checks if the script is running inside an OnlyOffice environment.
     * @returns {boolean} True if in OnlyOffice, otherwise false.
     */
    function isOnlyOfficeEnvironment() {
        return window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod;
    }

    function closeCurrentPluginWindow() {
        if (!window.Asc || !window.Asc.plugin) {
            return;
        }

        const windowId = window.Asc.plugin.windowID;
        if (windowId && typeof window.Asc.plugin.executeMethod === 'function') {
            window.Asc.plugin.executeMethod('CloseWindow', [windowId]);
            return;
        }

        if (typeof window.Asc.plugin.executeCommand === 'function') {
            window.Asc.plugin.executeCommand('close', '');
        }
    }

    if (isOnlyOfficeEnvironment()) {
        /**
         * The main entry point for the OnlyOffice plugin system.
         * @param {string} data - The data passed from the editor.
         */
        window.Asc.plugin.init = function(data) {
            console.log('[Main] OnlyOffice Plugin Init:', data);
            
            // The DOM is already loaded at this point, so we can ensure the game is initialized
            if (!isInitialized) {
                initializeGame();
            }

            // Here you could add logic to load a saved game state from `data` if implemented
        };

        /**
         * Handler for the plugin's UI buttons (e.g., Close).
         * @param {string} id - The ID of the button pressed.
         */
        window.Asc.plugin.button = function(id) {
            // The default config.json has one button which will have an id of 0.
            // A close button is implicitly added by the editor, which returns -1.
            if (id === 0 || id === -1) {
                closeCurrentPluginWindow();
            }
        };

        /**
         * Handler for theme changes in the OnlyOffice editor.
         * @param {object} theme - The new theme object.
         */
        window.Asc.plugin.onThemeChanged = function(theme) {
             // You can use the ThemeManager or simple class toggling here
            if (theme.type.includes('dark')) {
                document.documentElement.setAttribute('data-theme', 'dark');
            } else {
                document.documentElement.setAttribute('data-theme', 'light');
            }
        };
        
         /**
         * If the plugin is resizable, this handler will be called on window resize.
         * You can use this to re-render or adjust the UI.
         */
        window.Asc.plugin.onWindowResize = function() {
            // For a canvas-based game, you might call a `game.resize(newWidth, newHeight)` method.
            // For a DOM-based game, CSS should handle most of this, but you can add JS hooks here if needed.
            console.log('[Main] Plugin window resized.');
        }

    } else {
        console.warn('[Main] Not running in OnlyOffice environment. Plugin-specific features will be disabled.');
    }

})();
