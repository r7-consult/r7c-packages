/**
 * Single Application Entry Point - Clean Architecture
 * This is the ONLY initialization system for the 2048 game plugin
 * 
 * Responsibilities:
 * - Single initialization point
 * - Event delegation (all events)
 * - State management
 * - Environment detection
 * - Modal resizing
 */

(function() {
    'use strict';

    class App {
        constructor() {
            // Single source of truth for application state
            this.state = {
                initialized: false,
                gameManager: null,
                uiController: null,
                isOnlyOffice: false,
                config: {
                    gridSize: 5,  // Fixed 5x5 board
                    autoSave: true,
                    allowUndo: true
                }
            };
            
            // Prevent multiple instances
            if (window.app) {
                console.warn('App already instantiated');
                return window.app;
            }
        }
        
        /**
         * Single initialization method - called ONCE
         */
        async init() {
            // Prevent multiple initializations
            if (this.state.initialized) {
                console.log('✅ App already initialized');
                return;
            }
            
            console.log('🚀 Starting clean app initialization...');
            
            
            try {
                // Step 1: Detect environment
                
                this.detectEnvironment();
                
                // Step 2: Setup OnlyOffice if needed
                if (this.state.isOnlyOffice) {
                    
                    this.setupOnlyOffice();
                } else {
                    
                }
                
                // Step 3: Setup single event delegation system
                
                this.setupEventDelegation();
                
                // Step 4: Initialize game registry
                
                await this.initializeRegistry();
                
                // Step 5: Create game
                
                await this.createGame();
                
                // Step 6: Initialize UI
                
                this.initializeUI();
                
                this.state.initialized = true;
                console.log('✅ App initialization complete!');
                
                
            } catch (error) {
                console.error('❌ App initialization failed:', error);
                this.showError('Failed to initialize game: ' + error.message);
            }
        }
        
        /**
         * Detect if running in OnlyOffice environment
         */
        detectEnvironment() {
            this.state.isOnlyOffice = !!(
                window.Asc && 
                window.Asc.plugin && 
                (typeof window.Asc.plugin.executeMethod === 'function' || 
                 typeof window.AscDesktopEditor !== 'undefined')
            );
            
            console.log(`📍 Environment: ${this.state.isOnlyOffice ? 'OnlyOffice' : 'Standalone'}`);
        }
        
        /**
         * Setup OnlyOffice plugin handlers
         */
        setupOnlyOffice() {
            console.log('🔌 Setting up OnlyOffice handlers...');
            
            // OnlyOffice init handler
            window.Asc.plugin.init = (data) => {
                console.log('📨 OnlyOffice init called with data:', data);
                if (data && this.state.gameManager) {
                    try {
                        // Load saved game state
                        const gameData = JSON.parse(data);
                        this.state.gameManager.loadGameState(gameData);
                    } catch (error) {
                        console.warn('Could not load saved game state:', error);
                    }
                }
            };
            
            // OnlyOffice button handler
            window.Asc.plugin.button = (id) => {
                console.log('🔘 OnlyOffice button clicked:', id);
                if (id === -1 || id === 0) {
                    // Close without saving
                    this.closePlugin();
                }
            };
            
            // Set modal size for 5x5 board
            this.setModalSize();
        }
        
        /**
         * Setup single event delegation system
         * ALL events are handled here - no other event handlers allowed
         */
        setupEventDelegation() {
            console.log('🎯 Setting up single event delegation system...');
            
            // Single click handler for entire document
            document.addEventListener('click', this.handleClick.bind(this), false);
            console.log('🔗 Click event listener attached');
            
            // Single change handler for entire document
            document.addEventListener('change', this.handleChange.bind(this), false);
            console.log('🔗 Change event listener attached');
            
            
            // Keyboard handler for game controls
            document.addEventListener('keydown', this.handleKeyboard.bind(this));
            console.log('🔗 Keydown event listener attached');
            
        }
        
        /**
         * Global click handler - routes all clicks
         */
        handleClick(event) {
            const target = event.target;
            const id = target.id;
            
            // Route clicks based on element ID
            const handlers = {
                'new-game-btn': () => this.newGame(),
                'btn-new-game': () => this.newGame(), // Handle both IDs for now
                'undo-btn': () => this.undo(),
                'btn-undo': () => this.undo()
            };
            
            const handler = handlers[id];
            if (handler) {
                event.preventDefault();
                event.stopPropagation();
                console.log(`🔘 Handling click on: ${id}`);
                handler();
            }
        }
        
        /**
         * Global change handler - routes all change events
         */
        handleChange(event) {
            // No longer handling grid size changes as it's fixed at 5x5
        }
        
        /**
         * Keyboard handler for game controls
         */
        handleKeyboard(event) {
            // Check if an input is focused
            const activeElement = document.activeElement;
            if (activeElement && (
                activeElement.tagName === 'INPUT' || 
                activeElement.tagName === 'TEXTAREA' ||
                activeElement.contentEditable === 'true'
            )) {
                return;
            }
            
            // Map keycodes to directions
            const keyMap = {
                37: 'LEFT',  // Arrow Left
                38: 'UP',    // Arrow Up
                39: 'RIGHT', // Arrow Right
                40: 'DOWN',  // Arrow Down
                65: 'LEFT',  // A
                87: 'UP',    // W
                68: 'RIGHT', // D
                83: 'DOWN'   // S
            };
            
            const direction = keyMap[event.keyCode];
            if (direction && this.state.gameManager) {
                event.preventDefault();
                this.makeMove(direction);
            }
        }
        
        /**
         * Initialize game registry
         */
        async initializeRegistry() {
            console.log('📚 Initializing game registry...');
            
            if (!window.GameTypeRegistry) {
                throw new Error('GameTypeRegistry not found');
            }
            
            if (!window.gameTypeRegistry) {
                window.gameTypeRegistry = new window.GameTypeRegistry();
            }
            
            await window.gameTypeRegistry.initialize();
            
            // Register 2048 game
            if (!window.gameTypeRegistry.isGameTypeRegistered('2048')) {
                if (window.Game2048Manager) {
                    window.gameTypeRegistry.registerGameType('2048', window.Game2048Manager, {
                        name: '2048',
                        description: 'Slide tiles to create 2048',
                        category: 'puzzle',
                        minPlayers: 1,
                        maxPlayers: 1
                    });
                }
            }
        }
        
        /**
         * Create game instance
         */
        async createGame() {
            console.log('🎮 Creating game...');
            
            // Create game manager
            this.state.gameManager = new window.Game2048Manager();
            
            // Make globally accessible for debug messages
            window.gameManager = this.state.gameManager;
            
            // Initialize game
            const initialized = this.state.gameManager.initialize();
            if (!initialized) {
                throw new Error('Failed to initialize game manager');
            }
            
            // Start new game
            this.state.gameManager.startNewGame();
            
            console.log('✅ Game created and started');
        }
        
        /**
         * Initialize UI
         */
        initializeUI() {
            console.log('🎨 Initializing UI...');
            
            // Create star effects
            this.createStars();
            
            // Set language for buttons
            this.setLanguage();
            
            // Update UI with initial game state
            if (this.state.gameManager) {
                const gameState = this.state.gameManager.getGameState();
                this.updateUI(gameState);
            }
        }
        
        /**
         * Detect and set language based on office settings or browser
         */
        setLanguage() {
            // Try to get language from OnlyOffice settings
            let lang = 'en';
            
            if (this.state.isOnlyOffice && window.Asc?.plugin?.info?.lang) {
                lang = window.Asc.plugin.info.lang;
            } else if (navigator.language) {
                lang = navigator.language.substring(0, 2);
            }
            
            console.log('🌍 Detected language:', lang);
            
            // Language strings
            const translations = {
                en: {
                    newGame: 'New Game',
                    undo: 'Undo',
                    score: 'Score',
                    best: 'Best',
                    playAgain: 'Play Again',
                    gameOver: 'Game Over!',
                    youWin: 'You Win!'
                },
                ru: {
                    newGame: 'Новая игра',
                    undo: 'Отменить',
                    score: 'Счёт',
                    best: 'Рекорд',
                    playAgain: 'Играть снова',
                    gameOver: 'Игра окончена!',
                    youWin: 'Вы выиграли!'
                },
                es: {
                    newGame: 'Nuevo Juego',
                    undo: 'Deshacer',
                    score: 'Puntos',
                    best: 'Mejor',
                    playAgain: 'Jugar de Nuevo',
                    gameOver: '¡Juego Terminado!',
                    youWin: '¡Ganaste!'
                },
                de: {
                    newGame: 'Neues Spiel',
                    undo: 'Rückgängig',
                    score: 'Punkte',
                    best: 'Beste',
                    playAgain: 'Nochmal Spielen',
                    gameOver: 'Spiel Vorbei!',
                    youWin: 'Du Gewinnst!'
                },
                fr: {
                    newGame: 'Nouveau Jeu',
                    undo: 'Annuler',
                    score: 'Score',
                    best: 'Meilleur',
                    playAgain: 'Rejouer',
                    gameOver: 'Jeu Terminé!',
                    youWin: 'Vous Gagnez!'
                },
                zh: {
                    newGame: '新游戏',
                    undo: '撤销',
                    score: '分数',
                    best: '最高分',
                    playAgain: '再玩一次',
                    gameOver: '游戏结束！',
                    youWin: '你赢了！'
                }
            };
            
            // Default to English if language not supported
            const strings = translations[lang] || translations['en'];
            
            // Update button texts
            const newGameBtn = document.getElementById('new-game-btn');
            if (newGameBtn) newGameBtn.textContent = strings.newGame;
            
            const undoBtn = document.getElementById('btn-undo');
            if (undoBtn) undoBtn.textContent = strings.undo;
            
            // Update score labels
            const scoreLabels = document.querySelectorAll('.score-label');
            if (scoreLabels[0]) scoreLabels[0].textContent = strings.score;
            if (scoreLabels[1]) scoreLabels[1].textContent = strings.best;
            
            // Store strings for later use
            this.state.strings = strings;
        }
        
        /**
         * Create star effects for background
         */
        createStars() {
            const starsContainer = document.getElementById('stars');
            if (!starsContainer) return;
            
            // Create 50 stars
            for (let i = 0; i < 50; i++) {
                const star = document.createElement('div');
                star.className = 'star';
                star.style.left = Math.random() * 100 + '%';
                star.style.top = Math.random() * 100 + '%';
                star.style.animationDelay = Math.random() * 4 + 's';
                starsContainer.appendChild(star);
            }
        }
        
        /**
         * Start new game
         */
        async newGame() {
            try {
                console.log('🎮 Starting new game...');
                if (!this.state.gameManager) {
                    await this.createGame();
                } else {
                    this.state.gameManager.startNewGame();
                }
                
                // Update UI after new game
                const gameState = this.state.gameManager.getGameState();
                this.updateUI(gameState);
                
                console.log('✅ New game started');
            } catch (error) {
                console.error('❌ Failed to start new game:', error);
                this.showError('Failed to start new game');
            }
        }
        
        /**
         * Undo last move
         */
        async undo() {
            try {
                if (this.state.gameManager) {
                    this.state.gameManager.undo();
                    const gameState = this.state.gameManager.getGameState();
                    this.updateUI(gameState);
                    console.log('↩️ Move undone');
                }
            } catch (error) {
                console.error('❌ Failed to undo:', error);
            }
        }
        
        /**
         * Make a move in the game
         */
        async makeMove(direction) {
            try {
                if (this.state.gameManager) {
                    this.state.gameManager.makeMove(direction);
                    const gameState = this.state.gameManager.getGameState();
                    this.updateUI(gameState);
                    console.log(`➡️ Moved ${direction}`);
                }
            } catch (error) {
                console.error('❌ Failed to make move:', error);
            }
        }
        
        
        /**
         * Set modal window size for OnlyOffice
         */
        setModalSize() {
            if (!this.state.isOnlyOffice || !window.Asc?.plugin?.resizeWindow) {
                return;
            }
            
            // Slider AI design size: 320px board + UI elements
            const width = 420;  // Board width + logo + astronaut mascot
            const height = 520; // Board height + header + controls + padding
            
            try {
                window.Asc.plugin.resizeWindow(
                    width,      // width
                    height,     // height  
                    width,      // minWidth (same as width - no resizing)
                    height,     // minHeight (same as height - no resizing)
                    width,      // maxWidth (same as width - no resizing)
                    height      // maxHeight (same as height - no resizing)
                );
                console.log(`📐 Modal set to static size: ${width}x${height} (accommodates all board sizes)`);
            } catch (error) {
                console.warn('Could not resize modal:', error);
            }
        }
        
        /**
         * Update UI with game state
         */
        updateUI(gameState) {
            if (!gameState) return;
            
            console.log('[App] Updating UI with game state');
            
            // The Game2048Manager handles all rendering internally
            // We just need to update button states
            const undoBtn = document.getElementById('btn-undo');
            if (undoBtn) {
                undoBtn.disabled = !gameState.canUndo;
            }
        }
            
        
        
        closePlugin() {
            const ascPlugin = window.Asc && window.Asc.plugin;
            if (!ascPlugin) {
                return;
            }

            const windowId = ascPlugin.windowID;
            if (windowId && typeof ascPlugin.executeMethod === 'function') {
                ascPlugin.executeMethod('CloseWindow', [windowId]);
                return;
            }

            if (typeof ascPlugin.executeCommand === 'function') {
                ascPlugin.executeCommand('close', '');
            }
        }
        
        /**
         * Show error message
         */
        showError(message) {
            console.error('ERROR:', message);
            
            // Create error notification
            const notification = document.createElement('div');
            notification.className = 'notification error';
            notification.textContent = message;
            notification.style.cssText = `
                position: fixed;
                top: 20px;
                right: 20px;
                background: #f44336;
                color: white;
                padding: 15px 20px;
                border-radius: 4px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.2);
                z-index: 10000;
                animation: slideIn 0.3s ease-out;
            `;
            
            document.body.appendChild(notification);
            
            // Remove after 5 seconds
            setTimeout(() => {
                notification.remove();
            }, 5000);
        }
    }
    
    // ============================================================
    // SINGLE GLOBAL INSTANCE
    // ============================================================
    
    window.app = new App();
    
    
    // ============================================================
    // SINGLE DOM READY HANDLER
    // ============================================================
    
    if (document.readyState === 'loading') {
        
        document.addEventListener('DOMContentLoaded', () => {
            
            window.app.init();
        });
    } else {
        // DOM already loaded
        
        window.app.init();
    }
    
})();
