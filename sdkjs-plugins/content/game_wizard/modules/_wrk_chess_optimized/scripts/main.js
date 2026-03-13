/**
 * Chess Plugin OnlyOffice Integration Layer
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/02_api_reference_patterns.md#essential-plugin-object-methods
 */

(function() {
    'use strict';

    /**
     * OnlyOffice Plugin API Implementation
     * This module acts as the bridge between OnlyOffice and the chess plugin architecture
     */

    // Wait for OnlyOffice API to be available
    function waitForOnlyOfficeAPI() {
        return new Promise((resolve) => {
            if (window.Asc && window.Asc.plugin) {
                resolve();
            } else {
                const checkAPI = setInterval(() => {
                    if (window.Asc && window.Asc.plugin) {
                        clearInterval(checkAPI);
                        resolve();
                    }
                }, 50);
                
                // Timeout after 10 seconds
                setTimeout(() => {
                    clearInterval(checkAPI);
                    console.warn('OnlyOffice API not available, plugin will run in standalone mode');
                    resolve();
                }, 10000);
            }
        });
    }

    /**
     * Plugin initialization - called by OnlyOffice
     * @param {string} data - Initial data based on initDataType configuration
     */
    function initializePlugin(data) {
        window.ChessDebug?.logLifecycleEvent('onlyoffice_init_called', { 
            dataType: typeof data, 
            hasData: !!data,
            dataLength: data?.length || 0
        });

        // Prevent multiple initializations
        if (window.isChessPluginInitialized) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                'Plugin already initialized via OnlyOffice, skipping');
            return;
        }

        window.isChessPluginInitialized = true;

        // Initialize the chess plugin using the lifecycle manager
        window.ChessErrorHandler.wrapAsync(async () => {
            if (!window.ChessLifecycleManager) {
                throw new window.ChessErrors.ChessInitializationError(
                    'ChessLifecycleManager not available'
                );
            }

            await window.ChessLifecycleManager.initialize(data);
            
            // Also initialize legacy chess board if available
            if (window.ChessBoard && window.ChessBoard.instance && window.ChessBoard.init) {
                try {
                    window.ChessBoard.init(data);
                    window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                        'Legacy chess board initialized with data');
                } catch (error) {
                    window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                        'Failed to initialize legacy chess board', error);
                }
            }
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                'Chess plugin successfully initialized via OnlyOffice');

        }, { phase: 'onlyoffice_init' }).catch(error => {
            window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                'OnlyOffice initialization failed', error);
        });
    }

    /**
     * Button click handler - called by OnlyOffice
     * @param {number} id - Button ID from config.json
     */
    function handleButtonClick(id) {
        window.ChessDebug?.logUIEvent('onlyoffice_button_clicked', { buttonId: id });

        try {
            // In OnlyOffice: 0 = OK button, non-zero (usually 1) = Cancel button
            if (id === 0) {
                // OK button
                handleOkButton();
            } else {
                // Cancel button or any other button
                handleCancelButton();
            }
        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessInitializationError(
                    `Button handler failed for ID: ${id}`,
                    { buttonId: id, originalError: error }
                )
            );
        }
    };

    /**
     * External mouse up handler - called by OnlyOffice
     * Used for cleanup when plugin loses focus
     */
    function handleExternalMouseUp() {
        window.ChessDebug?.logLifecycleEvent('onlyoffice_external_mouse_up');

        try {
            // Cleanup any active operations
            if (window.ChessLifecycleManager) {
                const gameManager = window.ChessLifecycleManager.getComponent('gameManager');
                if (gameManager && gameManager.cancelActiveOperation) {
                    gameManager.cancelActiveOperation();
                }

                const uiController = window.ChessLifecycleManager.getComponent('uiController');
                if (uiController && uiController.handleExternalInteraction) {
                    uiController.handleExternalInteraction();
                }
            }

            // Handle legacy chess board cleanup
            if (window.ChessBoard && window.ChessBoard.handleExternalMouseUp) {
                window.ChessBoard.handleExternalMouseUp();
            }
        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessInitializationError(
                    'External mouse up handler failed',
                    { originalError: error }
                )
            );
        }
    }

    /**
     * Handle Cancel button click
     */
    function handleCancelButton() {
        window.ChessDebug?.logUIEvent('cancel_button_clicked');

        // Cancel any ongoing operations
        const gameManager = window.ChessLifecycleManager?.getComponent('gameManager');
        if (gameManager && gameManager.hasUnsavedChanges && gameManager.hasUnsavedChanges()) {
            // Show confirmation dialog for unsaved changes
            showUnsavedChangesConfirmation(() => {
                closePlugin();
            });
        } else {
            closePlugin();
        }
    }

    /**
     * Handle OK button click - save game state and close
     */
    function handleOkButton() {
        window.ChessDebug?.logUIEvent('ok_button_clicked');

        window.ChessErrorHandler.wrapAsync(async () => {
            const gameManager = window.ChessLifecycleManager?.getComponent('gameManager');
            
            if (!gameManager) {
                throw new window.ChessErrors.ChessInitializationError(
                    'Game manager not available for saving'
                );
            }

            // Get current game state
            const gameState = await gameManager.getCurrentGameState();
            
            if (!gameState) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.API_CALLS, 
                    'No game state to save');
                closePlugin();
                return;
            }

            // Save game state to document as OLE object
            await saveGameStateToDocument(gameState);
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.API_CALLS, 
                'Game state saved successfully');

            closePlugin();

        }, { phase: 'save_game_state' }).catch(error => {
            window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.API_CALLS, 
                'Failed to save game state', error);
            
            // Show error to user but still allow closing
            showSaveErrorConfirmation(() => {
                closePlugin();
            });
        });
    }

    /**
     * Handle Settings button click
     */
    function handleSettingsButton() {
        window.ChessDebug?.logUIEvent('settings_button_clicked');

        try {
            const uiController = window.ChessLifecycleManager?.getComponent('uiController');
            if (uiController && uiController.openSettings) {
                uiController.openSettings();
            } else {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Settings functionality not available');
            }
        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessInitializationError(
                    'Settings handler failed',
                    { originalError: error }
                )
            );
        }
    }

    /**
     * Save game state to document using OnlyOffice API
     */
    async function saveGameStateToDocument(gameState) {
        return new Promise((resolve, reject) => {
            try {
                // Get plugin info
                const pluginInfo = window.Asc.plugin.info;
                
                // Determine method based on whether we're creating or editing
                const method = pluginInfo.objectId ? 'EditOleObject' : 'AddOleObject';
                
                // Get image and data from legacy chess board if available
                let boardImage = '';
                let boardData = JSON.stringify(gameState);
                
                if (window.ChessBoard && window.ChessBoard.instance) {
                    try {
                        const widthPix = (pluginInfo.mmToPx * (pluginInfo.width || 70)) >> 0;
                        const heightPix = (pluginInfo.mmToPx * (pluginInfo.height || 70)) >> 0;
                        const result = window.ChessBoard.getResult(widthPix, heightPix);
                        if (result && result.image) {
                            boardImage = result.image;
                        }
                        boardData = window.ChessBoard.getData() || boardData;
                    } catch (error) {
                        window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.API_CALLS, 
                            'Failed to get legacy board data, using gameState', error);
                    }
                }
                
                const oleParams = {
                    guid: pluginInfo.guid,
                    widthPix: (pluginInfo.mmToPx * (pluginInfo.width || 70)) >> 0,
                    heightPix: (pluginInfo.mmToPx * (pluginInfo.height || 70)) >> 0,
                    width: pluginInfo.width || 70,
                    height: pluginInfo.height || 70,
                    imgSrc: boardImage,
                    data: boardData,
                    objectId: pluginInfo.objectId,
                    resize: pluginInfo.resize
                };
                
                window.Asc.plugin.executeMethod(method, [oleParams], function(result) {
                    if (result && result.error) {
                        reject(new window.ChessErrors.ChessPersistenceError(
                            `Failed to ${method}`,
                            { result, method, dataLength: boardData.length }
                        ));
                    } else {
                        window.ChessDebug?.logAPICall(method, { 
                            dataLength: boardData.length,
                            hasImage: !!boardImage,
                            success: true 
                        });
                        resolve(result);
                    }
                });

            } catch (error) {
                reject(new window.ChessErrors.ChessPersistenceError(
                    'Game state serialization failed',
                    { originalError: error }
                ));
            }
        });
    }

    /**
     * Close plugin safely
     */
    function closePlugin() {
        window.ChessDebug?.logLifecycleEvent('plugin_closing');

        // Check if we're in OnlyOffice or standalone mode
        const isOnlyOffice = !!(window.Asc && window.Asc.plugin);
        const closeOnlyOfficeWindow = () => {
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
        };

        // Cleanup components
        if (window.ChessLifecycleManager) {
            window.ChessLifecycleManager.cleanup().then(() => {
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                    'Plugin cleanup completed');
                
                // Close the plugin window or reset in standalone
                if (isOnlyOffice) {
                    closeOnlyOfficeWindow();
                } else {
                    closeStandalonePlugin();
                }
                
            }).catch(error => {
                window.ChessDebug?.error(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                    'Plugin cleanup failed', error);
                
                // Force close even if cleanup fails
                if (isOnlyOffice) {
                    closeOnlyOfficeWindow();
                } else {
                    closeStandalonePlugin();
                }
            });
        } else {
            // Direct close if no lifecycle manager
            if (isOnlyOffice) {
                closeOnlyOfficeWindow();
            } else {
                closeStandalonePlugin();
            }
        }
    }
    
    /**
     * Close plugin in standalone mode
     */
    function closeStandalonePlugin() {
        // In standalone mode, we can't actually close the window
        // But we can reset the game or show a message
        const gameManager = window.ChessLifecycleManager?.getComponent('gameManager');
        if (gameManager && gameManager.startNewGame) {
            gameManager.startNewGame();
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                'Game reset (standalone mode - cannot close window)');
        }
        
        // Show notification
        const statusText = document.getElementById('game-status-text');
        if (statusText) {
            statusText.textContent = 'Game reset. Close browser tab to exit.';
        }
    }

    /**
     * Show unsaved changes confirmation
     */
    function showUnsavedChangesConfirmation(onConfirm) {
        const uiController = window.ChessLifecycleManager?.getComponent('uiController');
        if (uiController && uiController.showConfirmation) {
            uiController.showConfirmation(
                'You have unsaved changes. Are you sure you want to close?',
                onConfirm,
                () => {} // Cancel - do nothing
            );
        } else {
            // Fallback to browser confirm
            if (confirm('You have unsaved changes. Are you sure you want to close?')) {
                onConfirm();
            }
        }
    }

    /**
     * Show save error confirmation
     */
    function showSaveErrorConfirmation(onConfirm) {
        const uiController = window.ChessLifecycleManager?.getComponent('uiController');
        if (uiController && uiController.showConfirmation) {
            uiController.showConfirmation(
                'Failed to save game state. Close anyway?',
                onConfirm,
                () => {} // Cancel - do nothing
            );
        } else {
            // Fallback to browser confirm
            if (confirm('Failed to save game state. Close anyway?')) {
                onConfirm();
            }
        }
    }

    /**
     * Theme change handler - called by OnlyOffice when theme changes
     */
    if (window.Asc && window.Asc.plugin && window.Asc.plugin.onThemeChanged) {
        window.Asc.plugin.onThemeChanged = function(theme) {
            window.ChessDebug?.logLifecycleEvent('onlyoffice_theme_changed', { theme });

            try {
                const themeManager = window.ChessLifecycleManager?.getComponent('themeManager');
                if (themeManager && themeManager.handleOnlyOfficeThemeChange) {
                    themeManager.handleOnlyOfficeThemeChange(theme);
                }
            } catch (error) {
                window.ChessErrorHandler?.handleError(
                    new window.ChessErrors.ChessRenderingError(
                        'Theme change handler failed',
                        { theme, originalError: error }
                    )
                );
            }
        };
    }

    /**
     * Plugin resize handler - called when OnlyOffice resizes the plugin
     */
    if (window.Asc && window.Asc.plugin && window.Asc.plugin.onResize) {
        window.Asc.plugin.onResize = function() {
            window.ChessDebug?.logLifecycleEvent('onlyoffice_resize');

            try {
                const boardRenderer = window.ChessLifecycleManager?.getComponent('boardRenderer');
                if (boardRenderer && boardRenderer.handleResize) {
                    boardRenderer.handleResize();
                }

                const uiController = window.ChessLifecycleManager?.getComponent('uiController');
                if (uiController && uiController.handleResize) {
                    uiController.handleResize();
                }
            } catch (error) {
                window.ChessErrorHandler?.handleError(
                    new window.ChessErrors.ChessRenderingError(
                        'Resize handler failed',
                        { originalError: error }
                    )
                );
            }
        };
    }

    /**
     * Setup buttons for standalone mode
     */
    function setupStandaloneButtons() {
        // Setup Cancel button
        const cancelBtn = document.getElementById('cancel-btn');
        if (cancelBtn) {
            cancelBtn.addEventListener('click', () => {
                handleCancelButton();
            });
        }
        
        // Setup Close button
        const closeBtn = document.getElementById('close-btn');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => {
                handleCancelButton(); // Same as cancel in standalone
            });
        }
        
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'Standalone buttons initialized');
    }

    // Initialize the plugin when DOM is ready
    async function initializeChessPlugin() {
        try {
            // Wait for OnlyOffice API
            await waitForOnlyOfficeAPI();
            
            // Set up OnlyOffice handlers if API is available
            if (window.Asc && window.Asc.plugin) {
                window.Asc.plugin.init = initializePlugin;
                window.Asc.plugin.button = handleButtonClick;
                
                // Resize window if needed (optional)
                // window.Asc.plugin.resizeWindow(undefined, undefined, 800, 600, 0, 0);
                
                // Set up theme change handler if available
                if (window.Asc.plugin.onThemeChanged) {
                    window.Asc.plugin.onThemeChanged = function(theme) {
                        window.ChessDebug?.logLifecycleEvent('onlyoffice_theme_changed', { theme });
                        
                        try {
                            const themeManager = window.ChessLifecycleManager?.getComponent('themeManager');
                            if (themeManager && themeManager.handleOnlyOfficeThemeChange) {
                                themeManager.handleOnlyOfficeThemeChange(theme);
                            }
                        } catch (error) {
                            window.ChessErrorHandler?.handleError(
                                new window.ChessErrors.ChessRenderingError(
                                    'Theme change handler failed',
                                    { theme, originalError: error }
                                )
                            );
                        }
                    };
                }
                
                // Set up resize handler if available
                if (window.Asc.plugin.onResize) {
                    window.Asc.plugin.onResize = function() {
                        window.ChessDebug?.logLifecycleEvent('onlyoffice_resize');
                        
                        try {
                            const boardRenderer = window.ChessLifecycleManager?.getComponent('boardRenderer');
                            if (boardRenderer && boardRenderer.handleResize) {
                                boardRenderer.handleResize();
                            }
                            
                            const uiController = window.ChessLifecycleManager?.getComponent('uiController');
                            if (uiController && uiController.handleResize) {
                                uiController.handleResize();
                            }
                        } catch (error) {
                            window.ChessErrorHandler?.handleError(
                                new window.ChessErrors.ChessRenderingError(
                                    'Resize handler failed',
                                    { originalError: error }
                                )
                            );
                        }
                    };
                }
                
                // Set up external mouse up handler
                window.Asc.plugin.onExternalMouseUp = handleExternalMouseUp;
                
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                    'OnlyOffice integration layer loaded with API');
            } else {
                // Initialize in standalone mode
                initializePlugin(null);
                setupStandaloneButtons();
                window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.LIFECYCLE, 
                    'Chess plugin loaded in standalone mode');
            }
        } catch (error) {
            console.error('Failed to initialize chess plugin:', error);
        }
    }

    // Initialize when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeChessPlugin);
    } else {
        initializeChessPlugin();
    }

})();
