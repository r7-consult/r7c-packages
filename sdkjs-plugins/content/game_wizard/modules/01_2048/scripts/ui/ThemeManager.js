/**
 * Chess Plugin Theme Manager
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/_prompt_CODE_ONLYOFFICE_ASSISTANT.md#enhanced-css-with-theme-support
 */

class ChessThemeManager {
    constructor() {
        this.currentTheme = 'light';
        this.isInitialized = false;
        this.eventListeners = new Map();
        this.onlyOfficeThemeMapping = {
            'theme-light': 'light',
            'theme-dark': 'dark',
            'theme-contrast-dark': 'dark'
        };
    }

    /**
     * Initialize theme manager
     */
    async initialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'ThemeManager initializing');

        try {
            // Detect initial theme
            await this.detectInitialTheme();
            
            // Apply initial theme
            await this.applyCurrentTheme();
            
            // Setup theme detection observers
            this.setupThemeObservers();
            
            this.isInitialized = true;
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'ThemeManager initialized', { currentTheme: this.currentTheme });

        } catch (error) {
            throw new window.ChessErrors.ChessRenderingError(
                'Theme manager initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Detect initial theme from various sources
     */
    async detectInitialTheme() {
        // 1. Check OnlyOffice theme information
        if (window.Asc && window.Asc.plugin && window.Asc.plugin.info && window.Asc.plugin.info.theme) {
            const onlyOfficeTheme = window.Asc.plugin.info.theme;
            const mappedTheme = this.mapOnlyOfficeTheme(onlyOfficeTheme);
            if (mappedTheme) {
                this.currentTheme = mappedTheme;
                window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Theme detected from OnlyOffice', { onlyOfficeTheme, mappedTheme });
                return;
            }
        }

        // 2. Check system preference
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            this.currentTheme = 'dark';
            window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'Dark theme detected from system preference');
            return;
        }

        // 3. Check localStorage preference
        const savedTheme = localStorage.getItem(window.ChessConstants.STORAGE.PLAYER_PREFERENCES);
        if (savedTheme) {
            try {
                const preferences = JSON.parse(savedTheme);
                if (preferences.theme && ['light', 'dark'].includes(preferences.theme)) {
                    this.currentTheme = preferences.theme;
                    window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                        'Theme detected from localStorage', { theme: preferences.theme });
                    return;
                }
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Failed to parse saved theme preference', error);
            }
        }

        // 4. Default to light theme
        this.currentTheme = 'light';
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'Using default light theme');
    }

    /**
     * Map OnlyOffice theme to chess theme
     */
    mapOnlyOfficeTheme(onlyOfficeTheme) {
        if (typeof onlyOfficeTheme === 'string') {
            return this.onlyOfficeThemeMapping[onlyOfficeTheme] || null;
        }

        if (onlyOfficeTheme && typeof onlyOfficeTheme === 'object') {
            // OnlyOffice might provide theme object with type property
            if (onlyOfficeTheme.type) {
                return this.onlyOfficeThemeMapping[onlyOfficeTheme.type] || null;
            }
            
            // Check for dark theme indicators in theme object
            if (onlyOfficeTheme.isDark || 
                onlyOfficeTheme.name?.toLowerCase().includes('dark') ||
                onlyOfficeTheme.id?.toLowerCase().includes('dark')) {
                return 'dark';
            }
        }

        return null;
    }

    /**
     * Setup theme change observers
     */
    setupThemeObservers() {
        // Watch for system theme changes
        if (window.matchMedia) {
            const darkModeQuery = window.matchMedia('(prefers-color-scheme: dark)');
            darkModeQuery.addEventListener('change', (e) => {
                // Only apply system theme if no OnlyOffice theme is set
                if (!this.hasOnlyOfficeTheme()) {
                    const newTheme = e.matches ? 'dark' : 'light';
                    this.setTheme(newTheme);
                }
            });
        }

        // Watch for manual theme attribute changes on document
        if (window.MutationObserver) {
            const observer = new MutationObserver((mutations) => {
                mutations.forEach((mutation) => {
                    if (mutation.type === 'attributes' && 
                        mutation.attributeName === 'data-theme') {
                        const newTheme = document.documentElement.getAttribute('data-theme');
                        if (newTheme && newTheme !== this.currentTheme) {
                            this.currentTheme = newTheme;
                            this.notifyThemeChanged(newTheme);
                        }
                    }
                });
            });

            observer.observe(document.documentElement, {
                attributes: true,
                attributeFilter: ['data-theme']
            });
        }
    }

    /**
     * Check if OnlyOffice theme is available
     */
    hasOnlyOfficeTheme() {
        return window.Asc && 
               window.Asc.plugin && 
               window.Asc.plugin.info && 
               window.Asc.plugin.info.theme;
    }

    /**
     * Apply current theme to the document
     */
    async applyCurrentTheme() {
        return this.applyTheme(this.currentTheme);
    }

    /**
     * Apply specific theme
     */
    async applyTheme(theme) {
        if (!['light', 'dark'].includes(theme)) {
            throw new window.ChessErrors.ChessValidationError(
                `Invalid theme: ${theme}`,
                { validThemes: ['light', 'dark'] }
            );
        }

        try {
            // Set theme attribute on document
            document.documentElement.setAttribute('data-theme', theme);
            
            // Update board colors if board renderer is available
            if (window.ChessLifecycleManager) {
                const boardRenderer = window.ChessLifecycleManager.getComponent('boardRenderer');
                if (boardRenderer && boardRenderer.updateTheme) {
                    await boardRenderer.updateTheme(theme);
                }
            }

            // Save theme preference
            await this.saveThemePreference(theme);

            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'Theme applied successfully', { theme });

        } catch (error) {
            throw new window.ChessErrors.ChessRenderingError(
                `Failed to apply theme: ${theme}`,
                { theme, originalError: error }
            );
        }
    }

    /**
     * Set theme and notify listeners
     */
    async setTheme(theme) {
        if (theme === this.currentTheme) {
            return; // No change needed
        }

        const previousTheme = this.currentTheme;
        
        try {
            await this.applyTheme(theme);
            this.currentTheme = theme;
            this.notifyThemeChanged(theme, previousTheme);
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'Theme changed', { from: previousTheme, to: theme });

        } catch (error) {
            window.ChessErrorHandler?.handleError(error);
            throw error;
        }
    }

    /**
     * Toggle between light and dark themes
     */
    async toggleTheme() {
        const newTheme = this.currentTheme === 'light' ? 'dark' : 'light';
        await this.setTheme(newTheme);
    }

    /**
     * Handle OnlyOffice theme change event
     */
    async handleOnlyOfficeThemeChange(onlyOfficeTheme) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'OnlyOffice theme change received', onlyOfficeTheme);

        try {
            const mappedTheme = this.mapOnlyOfficeTheme(onlyOfficeTheme);
            if (mappedTheme && mappedTheme !== this.currentTheme) {
                await this.setTheme(mappedTheme);
            }
        } catch (error) {
            window.ChessErrorHandler?.handleError(
                new window.ChessErrors.ChessRenderingError(
                    'OnlyOffice theme change handling failed',
                    { onlyOfficeTheme, originalError: error }
                )
            );
        }
    }

    /**
     * Save theme preference to localStorage
     */
    async saveThemePreference(theme) {
        try {
            const storageKey = window.ChessConstants.STORAGE.PLAYER_PREFERENCES;
            let preferences = {};
            
            // Load existing preferences
            const existingPrefs = localStorage.getItem(storageKey);
            if (existingPrefs) {
                preferences = JSON.parse(existingPrefs);
            }
            
            // Update theme preference
            preferences.theme = theme;
            preferences.lastUpdated = new Date().toISOString();
            
            // Save back to localStorage
            localStorage.setItem(storageKey, JSON.stringify(preferences));
            
        } catch (error) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                'Failed to save theme preference', error);
            // Don't throw - theme preference saving is not critical
        }
    }

    /**
     * Get current theme
     */
    getCurrentTheme() {
        return this.currentTheme;
    }

    /**
     * Check if current theme is dark
     */
    isDarkTheme() {
        return this.currentTheme === 'dark';
    }

    /**
     * Add event listener for theme changes
     */
    addEventListener(event, callback) {
        if (event === 'themeChanged') {
            if (!this.eventListeners.has(event)) {
                this.eventListeners.set(event, []);
            }
            this.eventListeners.get(event).push(callback);
        }
    }

    /**
     * Remove event listener
     */
    removeEventListener(event, callback) {
        if (this.eventListeners.has(event)) {
            const listeners = this.eventListeners.get(event);
            const index = listeners.indexOf(callback);
            if (index > -1) {
                listeners.splice(index, 1);
            }
        }
    }

    /**
     * Notify listeners of theme change
     */
    notifyThemeChanged(newTheme, previousTheme = null) {
        const listeners = this.eventListeners.get('themeChanged') || [];
        listeners.forEach(callback => {
            try {
                callback({ 
                    theme: newTheme, 
                    previousTheme,
                    isDark: newTheme === 'dark'
                });
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
                    'Theme change listener error', error);
            }
        });
    }

    /**
     * Get theme colors for current theme
     */
    getThemeColors() {
        const isDark = this.isDarkTheme();
        
        return {
            // Background colors
            bgPrimary: isDark ? '#2b2b2b' : '#ffffff',
            bgSecondary: isDark ? '#3c3c3c' : '#f8f9fa',
            bgTertiary: isDark ? '#4a4a4a' : '#e9ecef',
            bgBoard: isDark ? '#363636' : '#f5f5f5',
            
            // Text colors
            textPrimary: isDark ? '#f8f9fa' : '#212529',
            textSecondary: isDark ? '#adb5bd' : '#6c757d',
            textMuted: isDark ? '#6c757d' : '#adb5bd',
            
            // Chess board colors
            boardLightSquare: isDark ? '#7d7d7d' : '#f0d9b5',
            boardDarkSquare: isDark ? '#4a4a4a' : '#b58863',
            boardHighlight: isDark ? '#ffdd44' : '#ffff00',
            boardSelected: isDark ? '#7fd64d' : '#9fd64d',
            
            // UI colors
            borderColor: isDark ? '#495057' : '#dee2e6',
            accentColor: isDark ? '#42A5F5' : '#2196F3',
            accentHover: isDark ? '#2196F3' : '#1976D2'
        };
    }

    /**
     * Cleanup theme manager
     */
    async cleanup() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.UI_EVENTS, 
            'ThemeManager cleanup');
        
        // Clear event listeners
        this.eventListeners.clear();
        
        // Reset theme to light
        if (this.currentTheme !== 'light') {
            document.documentElement.setAttribute('data-theme', 'light');
        }
        
        this.isInitialized = false;
    }
}

// Export theme manager
window.ChessThemeManager = ChessThemeManager;