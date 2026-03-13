/**
 * Accessibility Utilities for 2048 Game
 * Provides ARIA labels, keyboard navigation, and screen reader support
 */

class GameAccessibility {
    constructor() {
        this.announcer = null;
        this.lastAnnouncement = '';
        this.init();
    }
    
    /**
     * Initialize accessibility features
     */
    init() {
        // Create screen reader announcer
        this.createAnnouncer();
        
        // Add keyboard instructions
        this.addKeyboardInstructions();
        
        // Set up focus management
        this.setupFocusManagement();
    }
    
    /**
     * Create live region for screen reader announcements
     */
    createAnnouncer() {
        this.announcer = document.getElementById('game-announcer');
        
        if (!this.announcer) {
            this.announcer = document.createElement('div');
            this.announcer.id = 'game-announcer';
            this.announcer.className = 'sr-only';
            this.announcer.setAttribute('role', 'status');
            this.announcer.setAttribute('aria-live', 'polite');
            this.announcer.setAttribute('aria-atomic', 'true');
            
            // Screen reader only styles
            this.announcer.style.cssText = `
                position: absolute;
                left: -10000px;
                width: 1px;
                height: 1px;
                overflow: hidden;
            `;
            
            document.body.appendChild(this.announcer);
        }
    }
    
    /**
     * Announce message to screen readers
     */
    announce(message, priority = 'polite') {
        if (!this.announcer || !message) return;
        
        // Avoid duplicate announcements
        if (message === this.lastAnnouncement && priority !== 'assertive') {
            return;
        }
        
        this.lastAnnouncement = message;
        
        // Set priority
        this.announcer.setAttribute('aria-live', priority);
        
        // Clear and set new message (forces re-announcement)
        this.announcer.textContent = '';
        setTimeout(() => {
            this.announcer.textContent = message;
        }, 100);
    }
    
    /**
     * Add keyboard instructions
     */
    addKeyboardInstructions() {
        const gameArea = document.getElementById('game-area');
        if (!gameArea) return;
        
        // Add ARIA labels
        gameArea.setAttribute('role', 'application');
        gameArea.setAttribute('aria-label', '2048 game board. Use arrow keys to move tiles.');
        gameArea.setAttribute('tabindex', '0');
        
        // Add keyboard help
        const helpText = document.createElement('div');
        helpText.id = 'keyboard-help';
        helpText.className = 'sr-only';
        helpText.textContent = 'Game controls: Use arrow keys to slide tiles. Press H for help, N for new game, U for undo.';
        
        if (!document.getElementById('keyboard-help')) {
            gameArea.parentNode?.insertBefore(helpText, gameArea);
        }
    }
    
    /**
     * Setup focus management
     */
    setupFocusManagement() {
        const gameArea = document.getElementById('game-area');
        if (!gameArea) return;
        
        // Focus game area on click
        gameArea.addEventListener('click', () => {
            gameArea.focus();
        });
        
        // Add focus indicator
        gameArea.addEventListener('focus', () => {
            gameArea.style.outline = '2px solid #2196F3';
            gameArea.style.outlineOffset = '2px';
            this.announce('Game board focused. Use arrow keys to play.');
        });
        
        gameArea.addEventListener('blur', () => {
            gameArea.style.outline = 'none';
        });
    }
    
    /**
     * Announce game move
     */
    announceMove(direction, result) {
        if (!result) {
            this.announce(`Cannot move ${direction}`, 'polite');
            return;
        }
        
        let message = `Moved ${direction}. `;
        
        if (result.score !== undefined) {
            message += `Score: ${result.score}. `;
        }
        
        if (result.newTile) {
            message += `New tile: ${result.newTile.value} at row ${result.newTile.row + 1}, column ${result.newTile.col + 1}. `;
        }
        
        if (result.gameWon) {
            message = 'Congratulations! You won! You created a 2048 tile!';
            this.announce(message, 'assertive');
            return;
        }
        
        if (result.gameOver) {
            message = 'Game over. No more moves available.';
            this.announce(message, 'assertive');
            return;
        }
        
        this.announce(message, 'polite');
    }
    
    /**
     * Announce game state
     */
    announceGameState(state) {
        if (!state) return;
        
        let message = `Game state: Score ${state.score}. `;
        
        if (state.bestScore) {
            message += `Best score: ${state.bestScore}. `;
        }
        
        if (state.movesCount) {
            message += `${state.movesCount} moves made. `;
        }
        
        if (state.hasUndo) {
            message += 'Undo available. ';
        }
        
        this.announce(message, 'polite');
    }
    
    /**
     * Add ARIA labels to buttons
     */
    enhanceButtons() {
        const buttons = {
            'new-game-btn': 'Start a new game',
            'undo-btn': 'Undo last move',
            'btn-new-game': 'Start a new 2048 game',
            'btn-undo': 'Undo the last move'
        };
        
        for (const [id, label] of Object.entries(buttons)) {
            const button = document.getElementById(id);
            if (button) {
                button.setAttribute('aria-label', label);
                
                // Add keyboard hint
                if (id.includes('new')) {
                    button.setAttribute('title', `${label} (Keyboard: N)`);
                } else if (id.includes('undo')) {
                    button.setAttribute('title', `${label} (Keyboard: U)`);
                }
            }
        }
    }
    
    /**
     * Setup keyboard shortcuts
     */
    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (event) => {
            // Skip if typing in input field
            if (event.target.tagName === 'INPUT' || event.target.tagName === 'TEXTAREA') {
                return;
            }
            
            switch(event.key.toLowerCase()) {
                case 'h':
                    this.announceHelp();
                    break;
                case 'n':
                    document.getElementById('new-game-btn')?.click();
                    this.announce('Starting new game', 'polite');
                    break;
                case 'u':
                    if (!document.getElementById('undo-btn')?.disabled) {
                        document.getElementById('undo-btn')?.click();
                        this.announce('Move undone', 'polite');
                    }
                    break;
                case 's':
                    // Announce current score
                    const scoreElement = document.getElementById('current-score');
                    if (scoreElement) {
                        this.announce(`Current score: ${scoreElement.textContent}`, 'polite');
                    }
                    break;
            }
        });
    }
    
    /**
     * Announce help information
     */
    announceHelp() {
        const helpMessage = `
            2048 Game Help:
            Arrow keys: Move tiles in that direction.
            N key: Start new game.
            U key: Undo last move.
            S key: Hear current score.
            H key: Hear this help message.
            Goal: Combine tiles to create a 2048 tile.
        `;
        
        this.announce(helpMessage, 'assertive');
    }
    
    /**
     * Describe board state for screen readers
     */
    describeBoardState(grid) {
        if (!grid || !Array.isArray(grid)) return;
        
        let description = 'Current board: ';
        let tileCount = 0;
        let maxTile = 0;
        
        for (let row = 0; row < grid.length; row++) {
            for (let col = 0; col < grid[row].length; col++) {
                const value = grid[row][col];
                if (value > 0) {
                    tileCount++;
                    maxTile = Math.max(maxTile, value);
                }
            }
        }
        
        description += `${tileCount} tiles on board. `;
        description += `Highest tile: ${maxTile}. `;
        
        this.announce(description, 'polite');
    }
}

// Create global instance
window.gameAccessibility = new GameAccessibility();