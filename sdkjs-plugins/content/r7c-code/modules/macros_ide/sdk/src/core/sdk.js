/**
 * @fileoverview OnlyOffice UI SDK Core Framework
 * @description Main SDK class for OnlyOffice JavaScript macros UI development
 * @see {@link https://api.onlyoffice.com/plugin/basic} OnlyOffice Plugin API
 * @see {@link /CODE_STANDARD.MD} Plugin Architecture Guide
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

/**
 * OnlyOffice UI SDK Core Class
 * Main entry point for the SDK providing component management and configuration
 */
const { getMergeConfig } = require('../utils/merge-config-loader.js');
const mergeConfig = getMergeConfig('../utils/merge-config.js');
class OnlyOfficeUISDK {
    #initialized = false;
    #components = new Map();
    #eventSystem = null;
    #stateManager = null;
    #lazyLoader = null;
    #componentClasses = new Map();
    #config = {
        theme: 'auto',
        debug: false,
        performance: {
            lazyLoad: true,
            virtualScrolling: true,
            bundleOptimization: true,
            preloadComponents: ['tree-view', 'modal', 'tooltip'],
            heavyComponents: ['data-grid', 'monaco-editor', 'chat-interface']
        },
        compatibility: {
            legacySupport: true,
            browserSupport: ['chrome', 'firefox', 'safari', 'edge']
        }
    };

    /**
     * Creates a new OnlyOffice UI SDK instance
     * @param {Object} options - SDK configuration options
     * @param {string} [options.theme='auto'] - Theme mode (auto, light, dark)
     * @param {boolean} [options.debug=false] - Debug mode
     * @param {Object} [options.performance] - Performance settings
     * @param {Object} [options.compatibility] - Compatibility settings
     */
    constructor(options = {}) {
        this.#config = mergeConfig(this.#config, options);
        this.#setupDebugMode();
        this.#logSDKInfo();
    }

    /**
     * Initializes the SDK with core systems
     * @param {string} containerId - Main container ID for SDK
     * @returns {Promise<void>}
     * @throws {Error} When initialization fails
     */
    async initialize(containerId = 'onlyoffice-sdk-container') {
        if (this.#initialized) {
            this.#log('warn', 'SDK already initialized');
            return;
        }

        try {
            this.#log('info', 'Initializing OnlyOffice UI SDK...');

            // Initialize core systems
            await this.#initializeCoreSystem();
            
            // Initialize event system
            await this.#initializeEventSystem();
            
            // Initialize state manager
            await this.#initializeStateManager();
            
            // Setup theme management
            await this.#setupThemeManagement();
            
            // Initialize component registry
            await this.#initializeComponentRegistry();
            
            // Setup OnlyOffice plugin integration
            await this.#setupOnlyOfficeIntegration();

            this.#initialized = true;
            this.#log('info', 'SDK initialized successfully');
            
            // Emit initialization complete event
            this.#eventSystem?.emit('sdk:initialized', {
                timestamp: new Date().toISOString(),
                config: this.#config
            });

        } catch (error) {
            this.#log('error', 'SDK initialization failed:', error);
            throw new Error(`OnlyOffice UI SDK initialization failed: ${error.message}`);
        }
    }

    /**
     * Registers a component with the SDK
     * @param {string} name - Component name
     * @param {Function} ComponentClass - Component class constructor
     * @param {Object} options - Component options
     * @returns {Promise<void>}
     */
    async registerComponent(name, ComponentClass, options = {}) {
        if (!this.#initialized) {
            throw new Error('SDK must be initialized before registering components');
        }

        try {
            const component = new ComponentClass({
                sdk: this,
                eventSystem: this.#eventSystem,
                stateManager: this.#stateManager,
                ...options
            });

            this.#components.set(name, component);
            this.#log('info', `Component registered: ${name}`);

            // Emit component registered event
            this.#eventSystem?.emit('component:registered', {
                name,
                component,
                timestamp: new Date().toISOString()
            });

        } catch (error) {
            this.#log('error', `Failed to register component ${name}:`, error);
            throw error;
        }
    }

    /**
     * Gets a registered component by name
     * @param {string} name - Component name
     * @returns {Object|null} Component instance or null if not found
     */
    getComponent(name) {
        return this.#components.get(name) || null;
    }

    /**
     * Creates a component instance
     * @param {string} componentType - Type of component to create
     * @param {Object} options - Component options
     * @returns {Promise<Object>} Component instance
     */
    async createComponent(componentType, options = {}) {
        if (!this.#initialized) {
            throw new Error('SDK must be initialized before creating components');
        }

        let ComponentClass = this.#getComponentClass(componentType);
        
        // If component not found and lazy loading is enabled, try lazy loading
        if (!ComponentClass && this.#config.performance.lazyLoad && this.#lazyLoader) {
            try {
                ComponentClass = await this.#lazyLoader.loadComponent(componentType);
                // Cache the loaded component class
                this.#registerComponentClass(componentType, ComponentClass);
            } catch (error) {
                this.#log('error', `Failed to lazy load component ${componentType}:`, error);
                throw new Error(`Component ${componentType} not available: ${error.message}`);
            }
        }

        if (!ComponentClass) {
            throw new Error(`Unknown component type: ${componentType}`);
        }

        try {
            const component = new ComponentClass({
                sdk: this,
                eventSystem: this.#eventSystem,
                stateManager: this.#stateManager,
                ...options
            });

            await component.initialize();
            this.#log('info', `Component created: ${componentType}`);
            
            return component;
        } catch (error) {
            this.#log('error', `Failed to create component ${componentType}:`, error);
            throw error;
        }
    }

    /**
     * Gets the event system instance
     * @returns {Object} Event system instance
     */
    getEventSystem() {
        return this.#eventSystem;
    }

    /**
     * Gets the state manager instance
     * @returns {Object} State manager instance
     */
    getStateManager() {
        return this.#stateManager;
    }

    /**
     * Gets SDK configuration
     * @returns {Object} SDK configuration
     */
    getConfig() {
        return { ...this.#config };
    }

    /**
     * Updates SDK configuration
     * @param {Object} updates - Configuration updates
     */
    updateConfig(updates) {
        this.#config = mergeConfig(this.#config, updates);
        this.#eventSystem?.emit('config:updated', {
            config: this.#config,
            timestamp: new Date().toISOString()
        });
    }

    /**
     * Gets the lazy loader instance
     * @returns {Object|null} Lazy loader instance or null if disabled
     */
    getLazyLoader() {
        return this.#lazyLoader;
    }

    /**
     * Preloads components for better performance
     * @param {Array<string>} componentNames - Components to preload
     * @returns {Promise<void>}
     */
    async preloadComponents(componentNames) {
        if (!this.#lazyLoader) {
            this.#log('warn', 'Lazy loading is disabled, cannot preload components');
            return;
        }
        
        return this.#lazyLoader.preloadComponents(componentNames);
    }

    /**
     * Gets component loading status
     * @returns {Object} Loading status information
     */
    getComponentStatus() {
        const status = {
            immediate: Array.from(this.#componentClasses.keys()),
            lazy: this.#lazyLoader ? this.#lazyLoader.getLoadingStatus() : { cached: [], loading: [], available: [] }
        };
        
        return status;
    }

    /**
     * Creates a lazy component factory
     * @param {string} componentType - Component type
     * @param {Object} defaultOptions - Default options
     * @returns {Function} Component factory function
     */
    createLazyComponentFactory(componentType, defaultOptions = {}) {
        if (!this.#lazyLoader) {
            throw new Error('Lazy loading is disabled');
        }
        
        return this.#lazyLoader.createLazyComponent(componentType, {
            sdk: this,
            eventSystem: this.#eventSystem,
            stateManager: this.#stateManager,
            ...defaultOptions
        });
    }

    /**
     * Destroys the SDK and cleans up resources
     * @returns {Promise<void>}
     */
    async destroy() {
        if (!this.#initialized) {
            return;
        }

        try {
            this.#log('info', 'Destroying SDK...');

            // Destroy all components
            for (const [name, component] of this.#components) {
                if (component.destroy) {
                    await component.destroy();
                }
                this.#log('info', `Component destroyed: ${name}`);
            }

            // Clean up core systems
            if (this.#eventSystem?.destroy) {
                await this.#eventSystem.destroy();
            }
            if (this.#stateManager?.destroy) {
                await this.#stateManager.destroy();
            }
            if (this.#lazyLoader?.destroy) {
                await this.#lazyLoader.destroy();
            }

            // Reset state
            this.#components.clear();
            this.#componentClasses.clear();
            this.#eventSystem = null;
            this.#stateManager = null;
            this.#lazyLoader = null;
            this.#initialized = false;

            this.#log('info', 'SDK destroyed successfully');

        } catch (error) {
            this.#log('error', 'Error during SDK destruction:', error);
            throw error;
        }
    }

    /**
     * Initializes core system
     * @private
     */
    async #initializeCoreSystem() {
        // Import core modules
        const { ComponentBase } = await import('./component.js');
        
        // Setup base component class
        this.ComponentBase = ComponentBase;
        
        // Initialize lazy loader if performance.lazyLoad is enabled
        if (this.#config.performance.lazyLoad) {
            const LazyLoader = (await import('./lazy-loader.js')).default;
            this.#lazyLoader = new LazyLoader({
                eventSystem: null, // Will be set after event system initialization
                enableCache: true,
                timeout: 10000,
                onLoadStart: (data) => this.#log('debug', `Loading component: ${data.componentName}`),
                onLoadComplete: (data) => this.#log('info', `Component loaded: ${data.componentName}`),
                onLoadError: (data) => this.#log('error', `Failed to load component: ${data.componentName}`, data.error)
            });
            this.#log('info', 'Lazy loader initialized');
        }
        
        this.#log('info', 'Core system initialized');
    }

    /**
     * Initializes event system
     * @private
     */
    async #initializeEventSystem() {
        const { EventSystem } = await import('./event-system.js');
        this.#eventSystem = new EventSystem(this.#config);
        await this.#eventSystem.initialize();
        
        // Connect lazy loader to event system
        if (this.#lazyLoader) {
            this.#lazyLoader.eventSystem = this.#eventSystem;
        }
        
        this.#log('info', 'Event system initialized');
    }

    /**
     * Initializes state manager
     * @private
     */
    async #initializeStateManager() {
        const { StateManager } = await import('./state-manager.js');
        this.#stateManager = new StateManager(this.#config);
        await this.#stateManager.initialize();
        this.#log('info', 'State manager initialized');
    }

    /**
     * Sets up theme management
     * @private
     */
    async #setupThemeManagement() {
        const { ThemeManager } = await import('../layout/theme-manager.js');
        this.#themeManager = new ThemeManager(this.#config);
        
        // Setup OnlyOffice theme integration
        if (typeof window !== 'undefined' && window.Asc?.plugin?.onThemeChanged) {
            const originalOnThemeChanged = window.Asc.plugin.onThemeChanged;
            window.Asc.plugin.onThemeChanged = (theme) => {
                if (originalOnThemeChanged) {
                    originalOnThemeChanged(theme);
                }
                this.#themeManager.updateTheme(theme);
                this.#eventSystem?.emit('theme:changed', theme);
            };
        }

        this.#log('info', 'Theme management initialized');
    }

    /**
     * Initializes component registry
     * @private
     */
    async #initializeComponentRegistry() {
        // Load built-in components
        await this.#loadBuiltinComponents();
        
        // Preload commonly used components if lazy loading is enabled
        if (this.#config.performance.lazyLoad && this.#lazyLoader && this.#config.performance.preloadComponents) {
            this.#log('info', 'Preloading commonly used components...');
            await this.#lazyLoader.preloadComponents(this.#config.performance.preloadComponents);
        }
        
        this.#log('info', 'Component registry initialized');
    }

    /**
     * Sets up OnlyOffice plugin integration
     * @private
     */
    async #setupOnlyOfficeIntegration() {
        if (typeof window !== 'undefined' && window.Asc?.plugin) {
            // Integration with OnlyOffice plugin lifecycle
            const originalInit = window.Asc.plugin.init;
            window.Asc.plugin.init = async () => {
                if (originalInit) {
                    await originalInit();
                }
                this.#eventSystem?.emit('onlyoffice:plugin:initialized');
            };
        }
        this.#log('info', 'OnlyOffice integration setup complete');
    }

    /**
     * Loads built-in components
     * @private
     */
    async #loadBuiltinComponents() {
        const allComponents = [
            'tree-view',
            'modal', 
            'split-panel',
            'tab-container',
            'toolbar',
            'context-menu',
            'tooltip',
            'progress-bar',
            'data-grid',
            'form-components',
            'search-box',
            'monaco-editor',
            'chat-interface'
        ];

        const heavyComponents = new Set(this.#config.performance.heavyComponents);

        for (const type of allComponents) {
            try {
                if (this.#config.performance.lazyLoad && heavyComponents.has(type)) {
                    // Register heavy components for lazy loading only
                    this.#log('debug', `Registering ${type} for lazy loading`);
                    this.#registerLazyComponent(type);
                } else {
                    // Load light components immediately
                    const module = await import(`../components/${type}/${type}.js`);
                    this.#registerComponentClass(type, module.default || module[this.#toPascalCase(type)]);
                    this.#log('debug', `Loaded component: ${type}`);
                }
            } catch (error) {
                this.#log('warn', `Failed to load component ${type}:`, error);
            }
        }
    }

    /**
     * Registers a component class
     * @param {string} type - Component type
     * @param {Function} ComponentClass - Component class
     * @private
     */
    #registerComponentClass(type, ComponentClass) {
        this.#componentClasses.set(type, ComponentClass);
    }

    /**
     * Registers a component for lazy loading
     * @param {string} type - Component type
     * @private
     */
    #registerLazyComponent(type) {
        // Mark component as available for lazy loading
        // The actual loading will be handled by the LazyLoader
        this.#log('debug', `Component ${type} registered for lazy loading`);
    }

    /**
     * Gets a component class by type
     * @param {string} type - Component type
     * @returns {Function|null} Component class or null
     * @private
     */
    #getComponentClass(type) {
        return this.#componentClasses?.get(type) || null;
    }

    /**
     * Converts kebab-case to PascalCase
     * @param {string} str - String to convert
     * @returns {string} PascalCase string
     * @private
     */
    #toPascalCase(str) {
        return str.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())
                  .replace(/^[a-z]/, letter => letter.toUpperCase());
    }

    /**
     * Merges configuration objects
     * @param {Object} target - Target configuration
     * @param {Object} source - Source configuration
     * @returns {Object} Merged configuration
     * @private
     */
    /**
     * Sets up debug mode
     * @private
     */
    #setupDebugMode() {
        if (this.#config.debug && typeof window !== 'undefined') {
            window.OnlyOfficeSDK_Debug = {
                sdk: this,
                getComponents: () => Array.from(this.#components.keys()),
                getConfig: () => this.#config,
                getEventSystem: () => this.#eventSystem,
                getStateManager: () => this.#stateManager
            };
        }
    }

    /**
     * Logs SDK information
     * @private
     */
    #logSDKInfo() {
        if (this.#config.debug) {
            console.log('%c OnlyOffice UI SDK v1.0.0 ', 'background: #4CAF50; color: white; padding: 2px 8px; border-radius: 4px;');
            console.log('Configuration:', this.#config);
        }
    }

    /**
     * Logs messages with SDK prefix
     * @param {string} level - Log level
     * @param {string} message - Message to log
     * @param {...any} args - Additional arguments
     * @private
     */
    #log(level, message, ...args) {
        if (this.#config.debug) {
            const prefix = `[OnlyOffice SDK]`;
            console[level](prefix, message, ...args);
        }
    }
}

// Export SDK class
if (typeof window !== 'undefined') {
    window.OnlyOfficeUISDK = OnlyOfficeUISDK;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = OnlyOfficeUISDK;
}

export default OnlyOfficeUISDK;
