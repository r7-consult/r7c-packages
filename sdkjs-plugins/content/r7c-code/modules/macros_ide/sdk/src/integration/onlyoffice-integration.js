/**
 * @fileoverview OnlyOffice Integration Layer for UI SDK
 * @description Integrates SDK components with OnlyOffice Plugin API
 * @see {@link https://api.onlyoffice.com/plugin/basic} OnlyOffice Plugin API
 * @see {@link /CODE_STANDARD.MD} Plugin Architecture Guide
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

/**
 * OnlyOffice Plugin Integration Manager
 * Provides seamless integration between SDK components and OnlyOffice Plugin API
 * SECURITY: Now uses robust API guard for safe API access
 */
class OnlyOfficeIntegration {
    #sdk = null;
    #apiGuard = null;
    #editorType = null;
    #isInitialized = false;
    #config = {};
    #eventListeners = new Map();
    #themeManager = null;
    #gracefulDegradation = false;

    /**
     * Creates a new OnlyOffice Integration instance
     * @param {Object} options - Integration options
     * @param {Object} options.sdk - SDK instance
     * @param {Object} [options.config={}] - Integration configuration
     * @param {boolean} [options.autoTheme=true] - Auto-detect and apply themes
     * @param {boolean} [options.crossEditorCompat=true] - Enable cross-editor compatibility
     * @param {Function} [options.onPluginInit] - Plugin initialization callback
     * @param {Function} [options.onThemeChange] - Theme change callback
     */
    constructor(options = {}) {
        this.#sdk = options.sdk;
        this.#config = {
            autoTheme: true,
            crossEditorCompat: true,
            debugMode: false,
            performanceOptimization: true,
            gracefulDegradation: true,
            apiTimeout: 30000,
            maxRetries: 5,
            ...options.config
        };

        // SECURITY FIX: Use robust API guard instead of direct API access
        // Create mock API guard immediately for safety
        this.#createMockAPIGuard();
    }

    /**
     * Initializes OnlyOffice integration
     * @returns {Promise<void>}
     */
    async initialize() {
        if (this.#isInitialized) {
            console.warn('[OnlyOffice Integration] Already initialized');
            return;
        }

        try {
            this.#log('info', 'Initializing OnlyOffice integration...');

            // ROBUSTNESS FIX: Initialize proper API guard
            await this.#initializeAPIGuard();
            
            // Wait for API to be available
            const apiAvailable = await this.#apiGuard.initialize();
            
            if (!apiAvailable) {
                if (this.#config.gracefulDegradation) {
                    this.#gracefulDegradation = true;
                    this.#log('warn', 'OnlyOffice API not available - running in degraded mode');
                } else {
                    throw new Error('OnlyOffice API not available and graceful degradation disabled');
                }
            }

            // Get editor information from API guard
            this.#editorType = this.#apiGuard.getEditorType() || 'unknown';
            this.#log('info', `Detected editor type: ${this.#editorType}`);

            // Setup plugin lifecycle hooks
            await this.#setupPluginLifecycle();

            // Initialize theme management
            if (this.#config.autoTheme) {
                await this.#initializeThemeManagement();
            }

            // Setup cross-editor compatibility
            if (this.#config.crossEditorCompat) {
                this.#setupCrossEditorCompatibility();
            }

            // Setup performance optimizations
            if (this.#config.performanceOptimization) {
                this.#setupPerformanceOptimizations();
            }

            this.#isInitialized = true;
            this.#log('info', 'OnlyOffice integration initialized successfully');

            // Emit integration ready event
            this.#sdk?.getEventSystem()?.emit('onlyoffice:integration:ready', {
                editorType: this.#editorType,
                apiVersion: this.#apiGuard.getAPIVersion(),
                apiAvailable: apiAvailable,
                gracefulDegradation: this.#gracefulDegradation,
                capabilities: this.#apiGuard.getCapabilities(),
                timestamp: new Date().toISOString()
            });

        } catch (error) {
            this.#log('error', 'Failed to initialize OnlyOffice integration:', error);
            
            // Try graceful degradation on error
            if (this.#config.gracefulDegradation && !this.#gracefulDegradation) {
                this.#gracefulDegradation = true;
                this.#log('warn', 'Falling back to graceful degradation mode');
                return this.initialize(); // Retry with degradation
            }
            
            throw error;
        }
    }

    /**
     * Gets the current editor type
     * @returns {string} Editor type (word, cell, slide, unknown)
     */
    getEditorType() {
        return this.#editorType || 'unknown';
    }

    /**
     * Gets OnlyOffice API instance
     * @returns {Object|null} OnlyOffice API instance
     */
    getPluginAPI() {
        return this.#pluginApi;
    }

    /**
     * Executes OnlyOffice plugin command
     * @param {Function} command - Command function to execute
     * @param {boolean} [async=true] - Whether command is async
     * @param {boolean} [silent=false] - Whether to suppress errors
     * @param {Function} [callback] - Callback function
     * @returns {Promise<any>} Command result
     */
    async executeCommand(command, async = true, silent = false, callback) {
        if (!this.#pluginApi) {
            throw new Error('OnlyOffice Plugin API not available');
        }

        return new Promise((resolve, reject) => {
            try {
                this.#pluginApi.callCommand(command, async, silent, (result) => {
                    if (callback) {
                        callback(result);
                    }
                    resolve(result);
                });
            } catch (error) {
                this.#log('error', 'Failed to execute plugin command:', error);
                reject(error);
            }
        });
    }

    /**
     * Gets current document selection information
     * @returns {Promise<Object>} Selection information
     */
    async getSelection() {
        // ROBUSTNESS FIX: Use API guard for safe access
        return this.#apiGuard.safeAPICall(async () => {
            if (this.#gracefulDegradation) {
                return { text: '', range: null, type: this.#editorType };
            }
            
            const selection = await this.#apiGuard.getSelection();
            return selection || { text: '', range: null, type: this.#editorType };
        }, { text: '', range: null, type: this.#editorType || 'unknown' });
    }

    /**
     * Inserts content at current position
     * @param {Object} content - Content to insert
     * @param {string} content.type - Content type (text, table, image, etc.)
     * @param {any} content.data - Content data
     * @returns {Promise<boolean>} Success status
     */
    async insertContent(content) {
        // ROBUSTNESS FIX: Use API guard for safe content insertion
        return this.#apiGuard.safeAPICall(async () => {
            if (this.#gracefulDegradation) {
                console.warn('[OnlyOffice Integration] Content insertion disabled in degraded mode');
                return { success: false, error: 'API not available' };
            }
            
            const result = await this.#apiGuard.insertContent(content);
            return { success: result, data: content };
        }, { success: false, error: 'API not available' });
    }

    /**
     * Registers for plugin events
     * @param {string} eventType - Event type
     * @param {Function} handler - Event handler
     */
    onPluginEvent(eventType, handler) {
        if (!this.#eventListeners.has(eventType)) {
            this.#eventListeners.set(eventType, []);
        }
        
        this.#eventListeners.get(eventType).push(handler);

        // Register with OnlyOffice API if available
        if (this.#pluginApi && this.#pluginApi.attachEvent) {
            this.#pluginApi.attachEvent(eventType, handler);
        }
    }

    /**
     * Unregisters from plugin events
     * @param {string} eventType - Event type
     * @param {Function} handler - Event handler
     */
    offPluginEvent(eventType, handler) {
        const handlers = this.#eventListeners.get(eventType);
        if (handlers) {
            const index = handlers.indexOf(handler);
            if (index > -1) {
                handlers.splice(index, 1);
            }
        }

        // Unregister from OnlyOffice API if available
        if (this.#pluginApi && this.#pluginApi.detachEvent) {
            this.#pluginApi.detachEvent(eventType, handler);
        }
    }

    /**
     * Shows OnlyOffice notification
     * @param {string} message - Notification message
     * @param {string} [type='info'] - Notification type
     */
    showNotification(message, type = 'info') {
        if (this.#pluginApi && this.#pluginApi.executeMethod) {
            this.#pluginApi.executeMethod('ShowNotification', [message, type]);
        } else {
            // Fallback to console
            console[type === 'error' ? 'error' : 'log'](`[OnlyOffice] ${message}`);
        }
    }

    /**
     * Gets current theme information
     * @returns {Object} Theme information
     */
    getCurrentTheme() {
        return this.#themeManager ? this.#themeManager.getCurrentTheme() : null;
    }

    /**
     * Destroys the integration and cleans up resources
     */
    destroy() {
        this.#log('info', 'Destroying OnlyOffice integration...');

        // Remove event listeners
        this.#eventListeners.forEach((handlers, eventType) => {
            handlers.forEach(handler => {
                if (this.#pluginApi && this.#pluginApi.detachEvent) {
                    this.#pluginApi.detachEvent(eventType, handler);
                }
            });
        });
        this.#eventListeners.clear();

        // Cleanup theme manager
        if (this.#themeManager && this.#themeManager.destroy) {
            this.#themeManager.destroy();
        }

        this.#sdk = null;
        this.#pluginApi = null;
        this.#themeManager = null;
        this.#isInitialized = false;
    }

    /**
     * Validates OnlyOffice environment
     * @private
     */
    #validateOnlyOfficeEnvironment() {
        if (typeof window === 'undefined') {
            throw new Error('OnlyOffice Integration requires browser environment');
        }

        if (!window.Asc || !window.Asc.plugin) {
            this.#log('warn', 'OnlyOffice Plugin API not found - running in standalone mode');
            return false;
        }

        return true;
    }

    /**
     * Sets up plugin API reference
     * @private
     */
    #setupPluginAPI() {
        if (window.Asc && window.Asc.plugin) {
            this.#pluginApi = window.Asc.plugin;
            this.#log('info', 'OnlyOffice Plugin API connected');
        }
    }

    /**
     * Detects the current editor type
     * @returns {string} Editor type
     * @private
     */
    #detectEditorType() {
        if (!this.#pluginApi) {
            return 'unknown';
        }

        // Try to detect based on available API methods
        try {
            if (typeof Api !== 'undefined') {
                if (Api.GetActiveSheet) return 'cell';
                if (Api.GetDocument) return 'word';
                if (Api.GetPresentation) return 'slide';
            }
        } catch (error) {
            this.#log('warn', 'Failed to detect editor type:', error);
        }

        return 'unknown';
    }

    /**
     * Sets up plugin lifecycle hooks
     * @private
     */
    #setupPluginLifecycle() {
        if (!this.#pluginApi) return;

        // Store original handlers
        const originalInit = this.#pluginApi.init;
        const originalButton = this.#pluginApi.button;

        // Wrap init handler
        this.#pluginApi.init = (...args) => {
            this.#log('info', 'Plugin initialization detected');
            
            if (originalInit) {
                originalInit.apply(this.#pluginApi, args);
            }

            // Emit SDK event
            this.#sdk?.getEventSystem()?.emit('onlyoffice:plugin:init', {
                args,
                timestamp: new Date().toISOString()
            });

            // Initialize SDK if not already done
            if (this.#config.onPluginInit) {
                this.#config.onPluginInit(args);
            }
        };

        // Wrap button handler
        this.#pluginApi.button = (...args) => {
            this.#log('info', 'Plugin button action detected');
            
            if (originalButton) {
                originalButton.apply(this.#pluginApi, args);
            }

            // Emit SDK event
            this.#sdk?.getEventSystem()?.emit('onlyoffice:plugin:button', {
                args,
                timestamp: new Date().toISOString()
            });
        };
    }

    /**
     * Initializes theme management
     * @private
     */
    async #initializeThemeManagement() {
        try {
            const { OnlyOfficeThemeManager } = await import('./onlyoffice-theme-manager.js');
            
            this.#themeManager = new OnlyOfficeThemeManager({
                pluginApi: this.#pluginApi,
                sdk: this.#sdk,
                onThemeChange: this.#config.onThemeChange
            });

            await this.#themeManager.initialize();
            this.#log('info', 'Theme management initialized');

        } catch (error) {
            this.#log('error', 'Failed to initialize theme management:', error);
        }
    }

    /**
     * Sets up cross-editor compatibility features
     * @private
     */
    #setupCrossEditorCompatibility() {
        // Add cross-editor utility methods to SDK
        if (this.#sdk) {
            this.#sdk.editorUtils = {
                getEditorType: () => this.#editorType,
                isWordEditor: () => this.#editorType === 'word',
                isCellEditor: () => this.#editorType === 'cell',
                isSlideEditor: () => this.#editorType === 'slide',
                getSelection: () => this.getSelection(),
                insertContent: (content) => this.insertContent(content)
            };
        }

        this.#log('info', 'Cross-editor compatibility enabled');
    }

    /**
     * Sets up performance optimizations
     * @private
     */
    #setupPerformanceOptimizations() {
        // Implement lazy loading for heavy components
        if (this.#sdk && this.#sdk.createComponent) {
            const originalCreateComponent = this.#sdk.createComponent.bind(this.#sdk);
            
            this.#sdk.createComponent = async (componentType, options = {}) => {
                // Add performance timing
                const startTime = performance.now();
                
                const component = await originalCreateComponent(componentType, options);
                
                const endTime = performance.now();
                this.#log('debug', `Component ${componentType} created in ${endTime - startTime}ms`);
                
                return component;
            };
        }

        this.#log('info', 'Performance optimizations enabled');
    }

    /**
     * Gets OnlyOffice API version
     * @returns {string} API version
     * @private
     */
    #getApiVersion() {
        try {
            if (this.#pluginApi && this.#pluginApi.version) {
                return this.#pluginApi.version;
            }
            return 'unknown';
        } catch (error) {
            return 'unknown';
        }
    }

    /**
     * Serializes selection object for cross-editor compatibility
     * @param {Object} selection - Selection object
     * @returns {Object} Serialized selection
     * @private
     */
    #serializeSelection(selection) {
        try {
            switch (this.#editorType) {
                case 'cell':
                    return {
                        type: 'cell',
                        address: selection.GetAddress ? selection.GetAddress(true, true, "xlA1", false) : null,
                        value: selection.GetValue ? selection.GetValue() : null,
                        row: selection.GetRow ? selection.GetRow() : null,
                        col: selection.GetCol ? selection.GetCol() : null
                    };
                case 'word':
                    return {
                        type: 'word',
                        text: selection.GetText ? selection.GetText() : null,
                        range: selection.GetRange ? selection.GetRange() : null
                    };
                case 'slide':
                    return {
                        type: 'slide',
                        slideIndex: selection.GetSlideIndex ? selection.GetSlideIndex() : null
                    };
                default:
                    return { type: 'unknown', raw: selection };
            }
        } catch (error) {
            return { type: 'error', error: error.message };
        }
    }

    /**
     * Inserts content in Word editor
     * @param {Object} api - Word API
     * @param {string} type - Content type
     * @param {any} data - Content data
     * @returns {Object} Result
     * @private
     */
    #insertWordContent(api, type, data) {
        const document = api.GetDocument();
        if (!document) throw new Error('Document not available');

        switch (type) {
            case 'text':
                const paragraph = api.CreateParagraph();
                paragraph.AddText(data);
                document.InsertContent([paragraph]);
                return { success: true };
            default:
                throw new Error(`Unsupported content type: ${type}`);
        }
    }

    /**
     * Inserts content in Cell editor
     * @param {Object} api - Cell API
     * @param {string} type - Content type
     * @param {any} data - Content data
     * @returns {Object} Result
     * @private
     */
    #insertCellContent(api, type, data) {
        const sheet = api.GetActiveSheet();
        if (!sheet) throw new Error('Sheet not available');

        switch (type) {
            case 'text':
            case 'value':
                const activeCell = sheet.GetActiveCell();
                if (activeCell) {
                    activeCell.SetValue(data);
                    return { success: true };
                }
                throw new Error('No active cell');
            default:
                throw new Error(`Unsupported content type: ${type}`);
        }
    }

    /**
     * Inserts content in Slide editor
     * @param {Object} api - Slide API
     * @param {string} type - Content type
     * @param {any} data - Content data
     * @returns {Object} Result
     * @private
     */
    #insertSlideContent(api, type, data) {
        const presentation = api.GetPresentation();
        if (!presentation) throw new Error('Presentation not available');

        switch (type) {
            case 'text':
                const slide = presentation.GetCurrentSlide();
                if (slide) {
                    // Add text box or other slide content
                    return { success: true };
                }
                throw new Error('No current slide');
            default:
                throw new Error(`Unsupported content type: ${type}`);
        }
    }

    /**
     * Initializes the API guard
     * @private
     */
    async #initializeAPIGuard() {
        try {
            // Import API guard dynamically to avoid circular dependencies
            const { default: OnlyOfficeAPIGuard } = await import('./onlyoffice-api-guard.js');
            this.#apiGuard = new OnlyOfficeAPIGuard({
                maxRetries: this.#config.maxRetries,
                timeout: this.#config.apiTimeout,
                gracefulDegradation: this.#config.gracefulDegradation
            });
        } catch (error) {
            console.error('[OnlyOffice Integration] Failed to load API guard:', error);
            // Create a mock API guard for graceful degradation
            this.#createMockAPIGuard();
        }
    }

    /**
     * Creates a mock API guard for graceful degradation
     * @private
     */
    #createMockAPIGuard() {
        this.#apiGuard = {
            initialize: async () => false,
            isAPIAvailable: () => false,
            getAPIVersion: () => 'unknown',
            getEditorType: () => 'unknown',
            hasCapability: () => false,
            getCapabilities: () => [],
            safeAPICall: async (_, fallback) => fallback,
            registerEventHandler: () => {},
            unregisterEventHandler: () => {},
            getPluginInfo: () => ({ guid: 'unknown', version: '1.0.0', variations: [], lang: 'en' }),
            getSelection: async () => null,
            insertContent: async () => false,
            destroy: () => {}
        };
        this.#gracefulDegradation = true;
    }

    /**
     * Logs messages with integration prefix
     * @param {string} level - Log level
     * @param {string} message - Message
     * @param {...any} args - Additional arguments
     * @private
     */
    #log(level, message, ...args) {
        if (this.#config.debugMode || level === 'error') {
            const prefix = '[OnlyOffice Integration]';
            console[level](prefix, message, ...args);
        }
    }
}

// Export OnlyOffice Integration
if (typeof window !== 'undefined') {
    window.OnlyOfficeIntegration = OnlyOfficeIntegration;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = OnlyOfficeIntegration;
}

export default OnlyOfficeIntegration;