/**
 * @fileoverview OnlyOffice Theme Manager for UI SDK
 * @description Manages automatic theme detection and switching for OnlyOffice integration
 * @see {@link https://api.onlyoffice.com/plugin/basic} OnlyOffice Plugin API
 * @see {@link /CODE_STANDARD.MD} Plugin Architecture Guide
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

/**
 * OnlyOffice Theme Manager
 * Handles automatic theme detection, CSS variable management, and component theme updates
 */
class OnlyOfficeThemeManager {
    #pluginApi = null;
    #sdk = null;
    #currentTheme = null;
    #isInitialized = false;
    #config = {};
    #originalThemeHandler = null;
    #cssVariables = new Map();
    #componentThemes = new Map();

    /**
     * Creates a new OnlyOffice Theme Manager
     * @param {Object} options - Theme manager options
     * @param {Object} options.pluginApi - OnlyOffice plugin API instance
     * @param {Object} options.sdk - SDK instance
     * @param {Function} [options.onThemeChange] - Theme change callback
     * @param {boolean} [options.autoApply=true] - Auto-apply theme changes
     * @param {boolean} [options.customVariables=true] - Enable custom CSS variables
     */
    constructor(options = {}) {
        this.#pluginApi = options.pluginApi;
        this.#sdk = options.sdk;
        this.#config = {
            autoApply: true,
            customVariables: true,
            debugMode: false,
            onThemeChange: options.onThemeChange,
            ...options.config
        };

        this.#setupDefaultTheme();
    }

    /**
     * Initializes theme management
     * @returns {Promise<void>}
     */
    async initialize() {
        if (this.#isInitialized) {
            console.warn('[Theme Manager] Already initialized');
            return;
        }

        try {
            this.#log('info', 'Initializing OnlyOffice theme management...');

            // Setup theme detection
            this.#setupThemeDetection();

            // Setup CSS variable system
            if (this.#config.customVariables) {
                this.#setupCSSVariables();
            }

            // Register component themes
            this.#registerComponentThemes();

            // Detect current theme
            await this.#detectCurrentTheme();

            this.#isInitialized = true;
            this.#log('info', 'Theme management initialized successfully');

            // Emit theme ready event
            this.#sdk?.getEventSystem()?.emit('theme:manager:ready', {
                currentTheme: this.#currentTheme,
                timestamp: new Date().toISOString()
            });

        } catch (error) {
            this.#log('error', 'Failed to initialize theme management:', error);
            throw error;
        }
    }

    /**
     * Gets current theme information
     * @returns {Object|null} Current theme object
     */
    getCurrentTheme() {
        return this.#currentTheme ? { ...this.#currentTheme } : null;
    }

    /**
     * Applies theme to all registered components
     * @param {Object} theme - Theme object
     */
    applyTheme(theme) {
        if (!theme) {
            this.#log('warn', 'No theme provided to apply');
            return;
        }

        this.#log('info', `Applying theme: ${theme.type || 'unknown'}`);

        try {
            // Update current theme
            this.#currentTheme = { ...theme };

            // Update CSS variables
            if (this.#config.customVariables) {
                this.#updateCSSVariables(theme);
            }

            // Apply theme to components
            this.#applyThemeToComponents(theme);

            // Call custom theme change handler
            if (this.#config.onThemeChange) {
                this.#config.onThemeChange(theme);
            }

            // Emit theme change event
            this.#sdk?.getEventSystem()?.emit('theme:changed', {
                theme: { ...theme },
                timestamp: new Date().toISOString()
            });

        } catch (error) {
            this.#log('error', 'Failed to apply theme:', error);
        }
    }

    /**
     * Registers a component for theme updates
     * @param {string} componentName - Component name
     * @param {Object} themeConfig - Component theme configuration
     */
    registerComponentTheme(componentName, themeConfig) {
        this.#componentThemes.set(componentName, themeConfig);
        this.#log('info', `Registered theme for component: ${componentName}`);

        // Apply current theme if available
        if (this.#currentTheme) {
            this.#applyThemeToComponent(componentName, themeConfig, this.#currentTheme);
        }
    }

    /**
     * Unregisters a component from theme updates
     * @param {string} componentName - Component name
     */
    unregisterComponentTheme(componentName) {
        this.#componentThemes.delete(componentName);
        this.#log('info', `Unregistered theme for component: ${componentName}`);
    }

    /**
     * Forces theme detection and update
     * @returns {Promise<Object>} Detected theme
     */
    async detectTheme() {
        return this.#detectCurrentTheme();
    }

    /**
     * Creates CSS variables for theme colors
     * @param {Object} theme - Theme object
     * @returns {Object} CSS variables object
     */
    createCSSVariables(theme) {
        if (!theme) return {};

        const variables = {};

        // Standard OnlyOffice theme properties
        const standardProps = [
            'background-normal',
            'background-toolbar',
            'background-tabs',
            'text-normal',
            'text-normal-pressed',
            'border-toolbar',
            'border-toolbar-button-hover',
            'highlight-button-hover',
            'highlight-button-pressed'
        ];

        standardProps.forEach(prop => {
            if (theme[prop]) {
                variables[`--onlyoffice-${prop}`] = theme[prop];
            }
        });

        // Add semantic variables for common use cases
        variables['--background-primary'] = theme['background-normal'] || '#ffffff';
        variables['--background-secondary'] = theme['background-toolbar'] || '#f5f5f5';
        variables['--text-primary'] = theme['text-normal'] || '#333333';
        variables['--text-secondary'] = theme['text-normal-pressed'] || '#666666';
        variables['--border-primary'] = theme['border-toolbar'] || '#d1d5db';
        variables['--border-hover'] = theme['border-toolbar-button-hover'] || '#9ca3af';

        // Determine theme type for conditional styling
        const isDark = this.#isDarkTheme(theme);
        variables['--theme-type'] = isDark ? 'dark' : 'light';
        variables['--is-dark-theme'] = isDark ? '1' : '0';

        return variables;
    }

    /**
     * Destroys theme manager and cleans up
     */
    destroy() {
        this.#log('info', 'Destroying theme manager...');

        // Restore original theme handler
        if (this.#originalThemeHandler && this.#pluginApi) {
            this.#pluginApi.onThemeChanged = this.#originalThemeHandler;
        }

        // Clear CSS variables
        this.#clearCSSVariables();

        // Clear component themes
        this.#componentThemes.clear();

        this.#pluginApi = null;
        this.#sdk = null;
        this.#currentTheme = null;
        this.#isInitialized = false;
    }

    /**
     * Sets up default theme
     * @private
     */
    #setupDefaultTheme() {
        this.#currentTheme = {
            type: 'light',
            'background-normal': '#ffffff',
            'background-toolbar': '#f5f5f5',
            'background-tabs': '#e5e5e5',
            'text-normal': '#333333',
            'text-normal-pressed': '#666666',
            'border-toolbar': '#d1d5db',
            'border-toolbar-button-hover': '#9ca3af',
            'highlight-button-hover': 'rgba(0, 0, 0, 0.05)',
            'highlight-button-pressed': 'rgba(0, 0, 0, 0.1)'
        };
    }

    /**
     * Sets up theme detection from OnlyOffice
     * @private
     */
    #setupThemeDetection() {
        if (!this.#pluginApi) {
            this.#log('warn', 'Plugin API not available for theme detection');
            return;
        }

        // Store original theme handler if it exists
        this.#originalThemeHandler = this.#pluginApi.onThemeChanged;

        // Setup theme change handler
        this.#pluginApi.onThemeChanged = (theme) => {
            this.#log('info', 'OnlyOffice theme change detected');

            // Call original handler first
            if (this.#originalThemeHandler) {
                this.#originalThemeHandler(theme);
            }

            // Apply our theme handling
            if (this.#config.autoApply) {
                this.applyTheme(theme);
            }
        };

        // Call onThemeChangedBase if available
        if (this.#pluginApi.onThemeChangedBase) {
            const originalBase = this.#pluginApi.onThemeChangedBase;
            this.#pluginApi.onThemeChangedBase = (theme) => {
                if (originalBase) {
                    originalBase(theme);
                }
                
                // Ensure our theme is applied after base theme
                if (this.#config.autoApply) {
                    setTimeout(() => this.applyTheme(theme), 0);
                }
            };
        }
    }

    /**
     * Sets up CSS variable system
     * @private
     */
    #setupCSSVariables() {
        // Create or get theme style element
        let themeStyle = document.getElementById('onlyoffice-sdk-theme');
        if (!themeStyle) {
            themeStyle = document.createElement('style');
            themeStyle.id = 'onlyoffice-sdk-theme';
            themeStyle.type = 'text/css';
            document.head.appendChild(themeStyle);
        }

        this.#themeStyle = themeStyle;
        this.#log('info', 'CSS variables system initialized');
    }

    /**
     * Registers built-in component themes
     * @private
     */
    #registerComponentThemes() {
        // Register themes for all SDK components
        const componentThemes = {
            'tree-view': this.#getTreeViewTheme(),
            'modal': this.#getModalTheme(),
            'data-grid': this.#getDataGridTheme(),
            'form-components': this.#getFormTheme(),
            'search-box': this.#getSearchBoxTheme(),
            'split-panel': this.#getSplitPanelTheme(),
            'tab-container': this.#getTabContainerTheme(),
            'toolbar': this.#getToolbarTheme(),
            'context-menu': this.#getContextMenuTheme(),
            'tooltip': this.#getTooltipTheme(),
            'progress-bar': this.#getProgressBarTheme()
        };

        Object.entries(componentThemes).forEach(([name, config]) => {
            this.registerComponentTheme(name, config);
        });
    }

    /**
     * Detects current theme from OnlyOffice
     * @returns {Promise<Object>} Detected theme
     * @private
     */
    async #detectCurrentTheme() {
        try {
            // Try to get theme from OnlyOffice plugin
            if (this.#pluginApi && this.#pluginApi.theme) {
                const detectedTheme = { ...this.#pluginApi.theme };
                this.#currentTheme = detectedTheme;
                this.#log('info', 'Theme detected from OnlyOffice API');
                
                if (this.#config.autoApply) {
                    this.applyTheme(detectedTheme);
                }
                
                return detectedTheme;
            }

            // Fallback: Try to detect from DOM
            const detectedTheme = this.#detectThemeFromDOM();
            if (detectedTheme) {
                this.#currentTheme = detectedTheme;
                this.#log('info', 'Theme detected from DOM');
                
                if (this.#config.autoApply) {
                    this.applyTheme(detectedTheme);
                }
                
                return detectedTheme;
            }

            // Use default theme
            this.#log('info', 'Using default theme');
            if (this.#config.autoApply) {
                this.applyTheme(this.#currentTheme);
            }
            
            return this.#currentTheme;

        } catch (error) {
            this.#log('error', 'Failed to detect theme:', error);
            return this.#currentTheme;
        }
    }

    /**
     * Detects theme from DOM elements
     * @returns {Object|null} Detected theme
     * @private
     */
    #detectThemeFromDOM() {
        try {
            // Look for OnlyOffice editor elements with computed styles
            const editorElement = document.querySelector('[data-cke-editor-name]') || 
                                 document.querySelector('.ace_editor') ||
                                 document.body;

            if (!editorElement) return null;

            const computedStyle = window.getComputedStyle(editorElement);
            const backgroundColor = computedStyle.backgroundColor;
            const color = computedStyle.color;

            // Determine if it's a dark theme based on background color
            const isDark = this.#isColorDark(backgroundColor);

            return {
                type: isDark ? 'dark' : 'light',
                'background-normal': backgroundColor || (isDark ? '#1e1e1e' : '#ffffff'),
                'text-normal': color || (isDark ? '#ffffff' : '#333333'),
                'background-toolbar': isDark ? '#2d2d2d' : '#f5f5f5',
                'border-toolbar': isDark ? '#3e3e3e' : '#d1d5db'
            };

        } catch (error) {
            this.#log('error', 'Failed to detect theme from DOM:', error);
            return null;
        }
    }

    /**
     * Updates CSS variables based on theme
     * @param {Object} theme - Theme object
     * @private
     */
    #updateCSSVariables(theme) {
        if (!this.#themeStyle) return;

        const variables = this.createCSSVariables(theme);
        this.#cssVariables = new Map(Object.entries(variables));

        // Generate CSS rules with sanitization
        const cssRules = [':root {'];
        this.#cssVariables.forEach((value, variable) => {
            // SECURITY FIX: Sanitize CSS variable names and values
            const sanitizedVariable = this.#sanitizeCSSVariableName(variable);
            const sanitizedValue = this.#sanitizeCSSValue(value);
            
            if (sanitizedVariable && sanitizedValue) {
                cssRules.push(`  ${sanitizedVariable}: ${sanitizedValue};`);
            }
        });
        cssRules.push('}');

        // Add component-specific theme classes
        cssRules.push(this.#generateComponentThemeCSS(theme));

        this.#themeStyle.textContent = cssRules.join('\n');
        this.#log('debug', 'CSS variables updated');
    }

    /**
     * Applies theme to all registered components
     * @param {Object} theme - Theme object
     * @private
     */
    #applyThemeToComponents(theme) {
        this.#componentThemes.forEach((config, componentName) => {
            this.#applyThemeToComponent(componentName, config, theme);
        });
    }

    /**
     * Applies theme to specific component
     * @param {string} componentName - Component name
     * @param {Object} config - Component theme config
     * @param {Object} theme - Theme object
     * @private
     */
    #applyThemeToComponent(componentName, config, theme) {
        try {
            // Apply CSS class based on theme type
            const themeClass = theme.type === 'dark' ? 'dark-theme' : 'light-theme';
            const elements = document.querySelectorAll(config.selector || `.onlyoffice-${componentName}`);
            
            elements.forEach(element => {
                element.classList.remove('dark-theme', 'light-theme');
                element.classList.add(themeClass);
            });

            // Apply custom theme function if provided
            if (config.applyTheme && typeof config.applyTheme === 'function') {
                config.applyTheme(theme);
            }

        } catch (error) {
            this.#log('error', `Failed to apply theme to ${componentName}:`, error);
        }
    }

    /**
     * Generates component-specific theme CSS
     * @param {Object} theme - Theme object
     * @returns {string} CSS rules
     * @private
     */
    #generateComponentThemeCSS(theme) {
        const isDark = this.#isDarkTheme(theme);
        const themeClass = isDark ? '.dark-theme' : '.light-theme';

        // SECURITY FIX: Sanitize CSS values to prevent injection
        const sanitizedValues = {
            background: this.#sanitizeCSSValue(theme['background-normal']) || (isDark ? '#1e1e1e' : '#ffffff'),
            text: this.#sanitizeCSSValue(theme['text-normal']) || (isDark ? '#ffffff' : '#333333'),
            border: this.#sanitizeCSSValue(theme['border-toolbar']) || (isDark ? '#3e3e3e' : '#d1d5db'),
            hover: this.#sanitizeCSSValue(theme['highlight-button-hover']) || (isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.05)')
        };

        return `
/* Component theme classes */
${themeClass} {
    --component-background: ${sanitizedValues.background};
    --component-text: ${sanitizedValues.text};
    --component-border: ${sanitizedValues.border};
    --component-hover: ${sanitizedValues.hover};
}`;
    }

    /**
     * Determines if a theme is dark
     * @param {Object} theme - Theme object
     * @returns {boolean} True if dark theme
     * @private
     */
    #isDarkTheme(theme) {
        if (theme.type === 'dark') return true;
        if (theme.type === 'light') return false;

        // Analyze background color
        const bg = theme['background-normal'];
        if (bg) {
            return this.#isColorDark(bg);
        }

        return false;
    }

    /**
     * Determines if a color is dark
     * @param {string} color - Color string
     * @returns {boolean} True if dark color
     * @private
     */
    #isColorDark(color) {
        try {
            // Convert color to RGB values
            const rgb = this.#parseColor(color);
            if (!rgb) return false;

            // Calculate relative luminance
            const luminance = (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b) / 255;
            return luminance < 0.5;

        } catch (error) {
            return false;
        }
    }

    /**
     * Parses color string to RGB values
     * @param {string} color - Color string
     * @returns {Object|null} RGB object
     * @private
     */
    #parseColor(color) {
        if (!color) return null;

        // Handle hex colors
        if (color.startsWith('#')) {
            const hex = color.slice(1);
            if (hex.length === 3) {
                return {
                    r: parseInt(hex[0] + hex[0], 16),
                    g: parseInt(hex[1] + hex[1], 16),
                    b: parseInt(hex[2] + hex[2], 16)
                };
            } else if (hex.length === 6) {
                return {
                    r: parseInt(hex.slice(0, 2), 16),
                    g: parseInt(hex.slice(2, 4), 16),
                    b: parseInt(hex.slice(4, 6), 16)
                };
            }
        }

        // Handle rgb colors
        const rgbMatch = color.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
        if (rgbMatch) {
            return {
                r: parseInt(rgbMatch[1]),
                g: parseInt(rgbMatch[2]),
                b: parseInt(rgbMatch[3])
            };
        }

        return null;
    }

    /**
     * Clears all CSS variables
     * @private
     */
    #clearCSSVariables() {
        if (this.#themeStyle) {
            this.#themeStyle.textContent = '';
        }
        this.#cssVariables.clear();
    }

    /**
     * Component theme configurations
     * @private
     */
    #getTreeViewTheme() {
        return {
            selector: '.onlyoffice-tree-view',
            applyTheme: (theme) => {
                // Custom tree view theme logic
            }
        };
    }

    #getModalTheme() {
        return {
            selector: '.onlyoffice-modal',
            applyTheme: (theme) => {
                // Custom modal theme logic
            }
        };
    }

    #getDataGridTheme() {
        return {
            selector: '.onlyoffice-datagrid',
            applyTheme: (theme) => {
                // Custom data grid theme logic
            }
        };
    }

    #getFormTheme() {
        return {
            selector: '.onlyoffice-form',
            applyTheme: (theme) => {
                // Custom form theme logic
            }
        };
    }

    #getSearchBoxTheme() {
        return {
            selector: '.onlyoffice-searchbox',
            applyTheme: (theme) => {
                // Custom search box theme logic
            }
        };
    }

    #getSplitPanelTheme() {
        return {
            selector: '.onlyoffice-split-panel',
            applyTheme: (theme) => {
                // Custom split panel theme logic
            }
        };
    }

    #getTabContainerTheme() {
        return {
            selector: '.onlyoffice-tab-container',
            applyTheme: (theme) => {
                // Custom tab container theme logic
            }
        };
    }

    #getToolbarTheme() {
        return {
            selector: '.onlyoffice-toolbar',
            applyTheme: (theme) => {
                // Custom toolbar theme logic
            }
        };
    }

    #getContextMenuTheme() {
        return {
            selector: '.onlyoffice-context-menu',
            applyTheme: (theme) => {
                // Custom context menu theme logic
            }
        };
    }

    #getTooltipTheme() {
        return {
            selector: '.onlyoffice-tooltip',
            applyTheme: (theme) => {
                // Custom tooltip theme logic
            }
        };
    }

    #getProgressBarTheme() {
        return {
            selector: '.onlyoffice-progress-bar',
            applyTheme: (theme) => {
                // Custom progress bar theme logic
            }
        };
    }

    /**
     * Sanitizes CSS variable names to prevent injection attacks
     * @param {string} variable - CSS variable name to sanitize
     * @returns {string|null} Sanitized variable name or null if invalid
     * @private
     */
    #sanitizeCSSVariableName(variable) {
        if (!variable || typeof variable !== 'string') {
            return null;
        }

        const cleaned = variable.trim();
        
        // CSS variables must start with -- and contain only safe characters
        if (!/^--[a-zA-Z0-9\-]+$/.test(cleaned)) {
            console.warn('[Theme Manager] Rejected unsafe CSS variable name:', variable);
            return null;
        }

        // Limit length
        if (cleaned.length > 50) {
            console.warn('[Theme Manager] Rejected overly long CSS variable name:', variable);
            return null;
        }

        return cleaned;
    }

    /**
     * Sanitizes CSS values to prevent injection attacks
     * @param {string} value - CSS value to sanitize
     * @returns {string|null} Sanitized value or null if invalid
     * @private
     */
    #sanitizeCSSValue(value) {
        if (!value || typeof value !== 'string') {
            return null;
        }

        // Remove any potentially dangerous characters
        const cleaned = value.trim();
        
        // Allow only safe CSS characters: alphanumeric, #, %, (, ), comma, space, dot, minus
        if (!/^[a-zA-Z0-9#%(),.\ \-rgba]+$/.test(cleaned)) {
            console.warn('[Theme Manager] Rejected unsafe CSS value:', value);
            return null;
        }

        // Block javascript: and data: URLs
        if (/javascript:|data:/i.test(cleaned)) {
            console.warn('[Theme Manager] Rejected CSS value with URL:', value);
            return null;
        }

        // Block CSS expressions and imports
        if (/expression\s*\(|@import|url\s*\(/i.test(cleaned)) {
            console.warn('[Theme Manager] Rejected CSS value with expression/import:', value);
            return null;
        }

        // Limit length to prevent abuse
        if (cleaned.length > 100) {
            console.warn('[Theme Manager] Rejected overly long CSS value:', value);
            return null;
        }

        return cleaned;
    }

    /**
     * Logs messages with theme manager prefix
     * @param {string} level - Log level
     * @param {string} message - Message
     * @param {...any} args - Additional arguments
     * @private
     */
    #log(level, message, ...args) {
        if (this.#config.debugMode || level === 'error') {
            const prefix = '[Theme Manager]';
            console[level](prefix, message, ...args);
        }
    }
}

// Export OnlyOffice Theme Manager
if (typeof window !== 'undefined') {
    window.OnlyOfficeThemeManager = OnlyOfficeThemeManager;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = OnlyOfficeThemeManager;
}

export { OnlyOfficeThemeManager };
export default OnlyOfficeThemeManager;