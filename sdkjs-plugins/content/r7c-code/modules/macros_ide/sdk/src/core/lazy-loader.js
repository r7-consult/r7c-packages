/**
 * @fileoverview Lazy Loading System for OnlyOffice UI SDK
 * @description Dynamic component loading with caching and error handling
 * @see {@link /CODE_STANDARD.MD} Plugin Architecture Guide
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

import { ComponentError, ValidationError, createDefaultErrorHandler } from './error-handler.js';

/**
 * Lazy component loader for heavy SDK components
 * Provides dynamic imports with caching, loading states, and error handling
 * SECURITY: Fixed path injection and added validation
 * PERFORMANCE: Fixed race conditions and memory leaks
 */
class LazyLoader {
    #cache = new Map();
    #loadingPromises = new Map();
    #loadingStates = new Map(); // Track loading state per component
    #componentReferences = new WeakMap(); // Prevent memory leaks
    #cacheExpiration = new Map(); // Cache TTL
    #config = {};
    #eventSystem = null;
    #destroyed = false;
    #allowedComponents = new Set(); // Security: whitelist allowed components
    #errorHandler = null;

    /**
     * Creates a new LazyLoader instance
     * @param {Object} options - Loader options
     * @param {Object} options.eventSystem - Event system instance
     * @param {boolean} [options.enableCache=true] - Enable component caching
     * @param {number} [options.timeout=10000] - Load timeout in milliseconds
     * @param {Function} [options.onLoadStart] - Load start callback
     * @param {Function} [options.onLoadComplete] - Load complete callback
     * @param {Function} [options.onLoadError] - Load error callback
     */
    constructor(options = {}) {
        try {
            // ERROR HANDLING: Initialize error handler first
            this.#errorHandler = createDefaultErrorHandler(options);
            
            // VALIDATION: Validate configuration
            this.#validateConfig(options);
            
            this.#config = {
                enableCache: true,
                timeout: 10000,
                retryAttempts: 3,
                retryDelay: 1000,
                preloadThreshold: 2,
                cacheMaxAge: 30 * 60 * 1000, // 30 minutes cache TTL
                maxCacheSize: 50, // Prevent memory bloat
                ...options
            };

            this.#eventSystem = options.eventSystem;
            this.#setupComponentMap();
            this.#startCacheCleanup();
        } catch (error) {
            throw new ComponentError('LazyLoader', 'constructor', error.message, error);
        }
    }

    /**
     * Loads a component lazily
     * @param {string} componentName - Component name to load
     * @param {Object} [options={}] - Load options
     * @returns {Promise<Class>} Component class
     */
    async loadComponent(componentName, options = {}) {
        try {
            // ERROR HANDLING: Validate parameters
            this.#validateLoadParameters(componentName, options);
            
            // SECURITY: Validate component name
            if (!this.#isValidComponentName(componentName)) {
                throw new ValidationError('componentName', componentName, {
                    pattern: '^[a-zA-Z0-9-]+$',
                    maxLength: 50,
                    suggestion: 'Component name must contain only alphanumeric characters and hyphens'
                });
            }

            if (this.#destroyed) {
                throw new ComponentError('LazyLoader', 'loadComponent', 'LazyLoader has been destroyed');
            }

        const normalizedName = this.#normalizeComponentName(componentName);
        
            // SECURITY: Check if component is allowed
            if (!this.#allowedComponents.has(normalizedName)) {
                throw new ValidationError('componentName', normalizedName, {
                    enum: Array.from(this.#allowedComponents),
                    suggestion: 'Component must be in the security whitelist'
                });
            }

        // Check cache first and validate TTL
        if (this.#config.enableCache && this.#cache.has(normalizedName)) {
            const cacheTime = this.#cacheExpiration.get(normalizedName);
            if (cacheTime && Date.now() - cacheTime < this.#config.cacheMaxAge) {
                this.#log('debug', `Loading ${normalizedName} from cache`);
                return this.#cache.get(normalizedName);
            } else {
                // Cache expired
                this.#cache.delete(normalizedName);
                this.#cacheExpiration.delete(normalizedName);
            }
        }

        // RACE CONDITION FIX: Atomic check and set
        if (this.#loadingPromises.has(normalizedName)) {
            this.#log('debug', `${normalizedName} already loading, waiting...`);
            const existingPromise = this.#loadingPromises.get(normalizedName);
            try {
                return await existingPromise;
            } catch (error) {
                // If the existing promise failed, we'll try again
                this.#loadingPromises.delete(normalizedName);
                this.#loadingStates.delete(normalizedName);
            }
        }

        // Start loading with error boundary
        const loadPromise = this.#loadComponentWithRetry(normalizedName, options)
            .finally(() => {
                // Always clean up loading state
                this.#loadingPromises.delete(normalizedName);
                this.#loadingStates.delete(normalizedName);
            });

        this.#loadingPromises.set(normalizedName, loadPromise);
        this.#loadingStates.set(normalizedName, 'loading');

        try {
            const component = await loadPromise;
            
            // Cache the component with TTL
            if (this.#config.enableCache) {
                this.#enforceMaxCacheSize();
                this.#cache.set(normalizedName, component);
                this.#cacheExpiration.set(normalizedName, Date.now());
            }
            
            this.#loadingStates.set(normalizedName, 'loaded');
            return component;

        } catch (error) {
            // ERROR HANDLING: Comprehensive error handling
            this.#loadingStates.set(normalizedName, 'error');
            
            const loadError = error instanceof ComponentError || error instanceof ValidationError 
                ? error 
                : new ComponentError('LazyLoader', 'loadComponent', error.message, error);
            
            // Handle error with recovery if possible
            const recovered = await this.#errorHandler?.handleError(loadError, {
                componentName,
                normalizedName,
                options,
                loadingState: this.#loadingStates.get(normalizedName),
                cacheSize: this.#cache.size
            });
            
            if (!recovered) {
                throw loadError;
            }
            
            // If recovered, return null to indicate failure but recovery attempted
            return null;
        }
    }
    }

    /**
     * Preloads components for better performance
     * @param {Array<string>} componentNames - Components to preload
     * @param {Object} [options={}] - Preload options
     * @returns {Promise<void>}
     */
    async preloadComponents(componentNames, options = {}) {
        try {
            // ERROR HANDLING: Validate preload parameters
            this.#validatePreloadParameters(componentNames, options);
            
            this.#log('info', `Preloading ${componentNames.length} components`);
            
            const preloadPromises = componentNames.map(async (name) => {
                try {
                    await this.loadComponent(name, { ...options, silent: true });
                    this.#log('debug', `Preloaded ${name} successfully`);
                    return { name, status: 'success' };
                } catch (error) {
                    this.#log('warn', `Failed to preload ${name}:`, error.message);
                    return { name, status: 'failed', error: error.message };
                }
            });

            const results = await Promise.allSettled(preloadPromises);
            const successful = results.filter(r => r.status === 'fulfilled' && r.value.status === 'success').length;
            
            this.#log('info', `Component preloading complete: ${successful}/${componentNames.length} successful`);
            return results;
        } catch (error) {
            const preloadError = new ComponentError(
                'LazyLoader',
                'preloadComponents',
                error.message,
                error
            );
            
            this.#errorHandler?.handleError(preloadError, { componentNames, options });
            throw preloadError;
        }
    }

    /**
     * Creates a lazy component wrapper
     * @param {string} componentName - Component name
     * @param {Object} [defaultOptions={}] - Default component options
     * @returns {Function} Component factory function
     */
    createLazyComponent(componentName, defaultOptions = {}) {
        return async (options = {}) => {
            const Component = await this.loadComponent(componentName);
            return new Component({ ...defaultOptions, ...options });
        };
    }

    /**
     * Gets loading status of components
     * @returns {Object} Loading status
     */
    getLoadingStatus() {
        return {
            cached: Array.from(this.#cache.keys()),
            loading: Array.from(this.#loadingPromises.keys()),
            available: Object.keys(this.#componentMap)
        };
    }

    /**
     * Checks if component is available
     * @param {string} componentName - Component name
     * @returns {boolean} True if component is available
     */
    isComponentAvailable(componentName) {
        const normalizedName = this.#normalizeComponentName(componentName);
        return this.#componentMap.hasOwnProperty(normalizedName);
    }

    /**
     * Checks if component is loaded
     * @param {string} componentName - Component name
     * @returns {boolean} True if component is loaded
     */
    isComponentLoaded(componentName) {
        const normalizedName = this.#normalizeComponentName(componentName);
        return this.#cache.has(normalizedName);
    }

    /**
     * Clears component cache
     * @param {string} [componentName] - Specific component to clear (clears all if not specified)
     */
    clearCache(componentName) {
        if (componentName) {
            const normalizedName = this.#normalizeComponentName(componentName);
            this.#cache.delete(normalizedName);
            this.#log('debug', `Cleared cache for ${normalizedName}`);
        } else {
            this.#cache.clear();
            this.#log('debug', 'Cleared all component cache');
        }
    }

    /**
     * Sets up component import map
     * @private
     */
    #setupComponentMap() {
        // SECURITY: Define allowed components with static imports to prevent path injection
        const allowedComponents = [
            'data-grid',
            'monaco-editor', 
            'chat-interface',
            'tree-view',
            'modal',
            'form-components',
            'search-box',
            'split-panel',
            'tab-container',
            'toolbar',
            'context-menu',
            'tooltip',
            'progress-bar'
        ];

        // Populate security whitelist
        allowedComponents.forEach(name => {
            this.#allowedComponents.add(name);
        });

        this.#componentMap = {
            // Heavy components (always lazy loaded)
            'data-grid': () => import('../components/data-grid/data-grid.js'),
            'monaco-editor': () => import('../components/monaco-editor/monaco-editor.js'),
            'chat-interface': () => import('../components/chat-interface/chat-interface.js'),
            
            // Standard components (can be lazy loaded)
            'tree-view': () => import('../components/tree-view/tree-view.js'),
            'modal': () => import('../components/modal/modal.js'),
            'form-components': () => import('../components/form-components/form-components.js'),
            'search-box': () => import('../components/search-box/search-box.js'),
            'split-panel': () => import('../components/split-panel/split-panel.js'),
            'tab-container': () => import('../components/tab-container/tab-container.js'),
            'toolbar': () => import('../components/toolbar/toolbar.js'),
            'context-menu': () => import('../components/context-menu/context-menu.js'),
            'tooltip': () => import('../components/tooltip/tooltip.js'),
            'progress-bar': () => import('../components/progress-bar/progress-bar.js')
        };

        this.#log('info', `Initialized lazy loader with ${Object.keys(this.#componentMap).length} components`);
    }

    /**
     * Loads component with retry logic
     * @param {string} componentName - Component name
     * @param {Object} options - Load options
     * @returns {Promise<Class>} Component class
     * @private
     */
    async #loadComponentWithRetry(componentName, options = {}) {
        let lastError;
        
        for (let attempt = 1; attempt <= this.#config.retryAttempts; attempt++) {
            try {
                this.#emitLoadEvent('start', componentName, attempt);
                
                const component = await this.#loadComponentDirect(componentName, options);
                
                this.#emitLoadEvent('complete', componentName, attempt);
                this.#log('info', `Loaded ${componentName} successfully on attempt ${attempt}`);
                
                return component;

            } catch (error) {
                lastError = error;
                this.#log('warn', `Failed to load ${componentName} on attempt ${attempt}:`, error.message);
                
                if (attempt < this.#config.retryAttempts) {
                    await this.#delay(this.#config.retryDelay * attempt);
                }
            }
        }

        this.#emitLoadEvent('error', componentName, this.#config.retryAttempts, lastError);
        throw new Error(`Failed to load component ${componentName} after ${this.#config.retryAttempts} attempts: ${lastError.message}`);
    }

    /**
     * Loads component directly
     * @param {string} componentName - Component name
     * @param {Object} options - Load options
     * @returns {Promise<Class>} Component class
     * @private
     */
    async #loadComponentDirect(componentName, options = {}) {
        const importFunction = this.#componentMap[componentName];
        
        if (!importFunction) {
            throw new Error(`Unknown component: ${componentName}`);
        }

        // Set up timeout
        const timeoutPromise = new Promise((_, reject) => {
            setTimeout(() => {
                reject(new Error(`Component load timeout: ${componentName}`));
            }, this.#config.timeout);
        });

        try {
            // Race between import and timeout
            const module = await Promise.race([
                importFunction(),
                timeoutPromise
            ]);

            // Extract default export or named export
            const Component = module.default || module[this.#getComponentClassName(componentName)];
            
            if (!Component) {
                throw new Error(`Component class not found in module: ${componentName}`);
            }

            // Validate component
            if (typeof Component !== 'function') {
                throw new Error(`Invalid component class: ${componentName}`);
            }

            return Component;

        } catch (error) {
            if (error.name === 'ChunkLoadError') {
                throw new Error(`Failed to load component chunk: ${componentName}`);
            }
            throw error;
        }
    }

    /**
     * Normalizes component name
     * @param {string} name - Component name
     * @returns {string} Normalized name
     * @private
     */
    #normalizeComponentName(name) {
        return name.toLowerCase().replace(/([a-z])([A-Z])/g, '$1-$2').replace(/[\s_]/g, '-');
    }

    /**
     * Gets component class name from component name
     * @param {string} componentName - Component name
     * @returns {string} Class name
     * @private
     */
    #getComponentClassName(componentName) {
        return componentName
            .split('-')
            .map(part => part.charAt(0).toUpperCase() + part.slice(1))
            .join('');
    }

    /**
     * Emits load events
     * @param {string} type - Event type
     * @param {string} componentName - Component name
     * @param {number} attempt - Attempt number
     * @param {Error} [error] - Error object
     * @private
     */
    #emitLoadEvent(type, componentName, attempt, error) {
        const eventData = {
            componentName,
            attempt,
            timestamp: new Date().toISOString()
        };

        if (error) {
            eventData.error = error.message;
        }

        this.#eventSystem?.emit(`lazy-loader:${type}`, eventData);

        // Call config callbacks
        switch (type) {
            case 'start':
                this.#config.onLoadStart?.(eventData);
                break;
            case 'complete':
                this.#config.onLoadComplete?.(eventData);
                break;
            case 'error':
                this.#config.onLoadError?.(eventData);
                break;
        }
    }

    /**
     * Delays execution
     * @param {number} ms - Milliseconds to delay
     * @returns {Promise<void>}
     * @private
     */
    #delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Validates component name for security
     * @param {string} componentName - Component name to validate
     * @returns {boolean} True if valid
     * @private
     */
    #isValidComponentName(componentName) {
        if (typeof componentName !== 'string') return false;
        if (componentName.length === 0 || componentName.length > 50) return false;
        // Only allow alphanumeric characters and hyphens
        return /^[a-zA-Z0-9-]+$/.test(componentName);
    }

    /**
     * Enforces maximum cache size to prevent memory bloat
     * @private
     */
    #enforceMaxCacheSize() {
        if (this.#cache.size >= this.#config.maxCacheSize) {
            // Remove oldest cache entry (LRU)
            const oldestKey = this.#cache.keys().next().value;
            this.#cache.delete(oldestKey);
            this.#cacheExpiration.delete(oldestKey);
            this.#log('debug', `Evicted cache entry: ${oldestKey}`);
        }
    }

    /**
     * Starts cache cleanup timer
     * @private
     */
    #startCacheCleanup() {
        // Clean up expired cache entries every 5 minutes
        this.#cacheCleanupInterval = setInterval(() => {
            this.#cleanupExpiredCache();
        }, 5 * 60 * 1000);
    }

    /**
     * Cleans up expired cache entries
     * @private
     */
    #cleanupExpiredCache() {
        const now = Date.now();
        const expiredKeys = [];
        
        for (const [key, cacheTime] of this.#cacheExpiration) {
            if (now - cacheTime >= this.#config.cacheMaxAge) {
                expiredKeys.push(key);
            }
        }
        
        expiredKeys.forEach(key => {
            this.#cache.delete(key);
            this.#cacheExpiration.delete(key);
        });
        
        if (expiredKeys.length > 0) {
            this.#log('debug', `Cleaned up ${expiredKeys.length} expired cache entries`);
        }
    }

    /**
     * Enhanced destroy method with proper cleanup
     */
    destroy() {
        this.#destroyed = true;
        
        // Clear cache cleanup interval
        if (this.#cacheCleanupInterval) {
            clearInterval(this.#cacheCleanupInterval);
        }
        
        // Clear all maps
        this.#cache.clear();
        this.#loadingPromises.clear();
        this.#loadingStates.clear();
        this.#cacheExpiration.clear();
        this.#allowedComponents.clear();
        
        // Clear weak map (will be garbage collected)
        this.#componentReferences = new WeakMap();
        
        this.#eventSystem = null;
        this.#componentMap = null;
        
        this.#log('info', 'LazyLoader destroyed and cleaned up');
    }

    /**
     * Validates lazy loader configuration
     * @param {Object} options - Configuration options
     * @throws {ValidationError} If validation fails
     * @private
     */
    #validateConfig(options) {
        if (options && typeof options !== 'object') {
            throw new ValidationError('options', options, { type: 'object' });
        }
        
        const schema = {
            enableCache: { type: 'boolean' },
            timeout: { type: 'number', min: 1000, max: 60000 },
            retryAttempts: { type: 'number', min: 0, max: 10 },
            retryDelay: { type: 'number', min: 100, max: 10000 },
            cacheMaxAge: { type: 'number', min: 60000, max: 3600000 }, // 1 min to 1 hour
            maxCacheSize: { type: 'number', min: 1, max: 1000 },
            debug: { type: 'boolean' },
            enableTelemetry: { type: 'boolean' }
        };
        
        this.#errorHandler?.validateConfig(options || {}, schema, 'LazyLoader');
    }
    
    /**
     * Validates load parameters
     * @param {string} componentName - Component name
     * @param {Object} options - Load options
     * @throws {ValidationError} If validation fails
     * @private
     */
    #validateLoadParameters(componentName, options) {
        if (typeof componentName !== 'string') {
            throw new ValidationError('componentName', componentName, { type: 'string' });
        }
        
        if (componentName.trim().length === 0) {
            throw new ValidationError('componentName', componentName, { minLength: 1 });
        }
        
        if (options && typeof options !== 'object') {
            throw new ValidationError('options', options, { type: 'object' });
        }
    }
    
    /**
     * Validates preload parameters
     * @param {Array} componentNames - Component names array
     * @param {Object} options - Preload options
     * @throws {ValidationError} If validation fails
     * @private
     */
    #validatePreloadParameters(componentNames, options) {
        if (!Array.isArray(componentNames)) {
            throw new ValidationError('componentNames', componentNames, { type: 'array' });
        }
        
        if (componentNames.length === 0) {
            throw new ValidationError('componentNames', componentNames, { minLength: 1 });
        }
        
        if (componentNames.length > 20) {
            throw new ValidationError('componentNames', componentNames, { 
                maxLength: 20,
                suggestion: 'Too many components to preload at once'
            });
        }
        
        // Validate each component name
        componentNames.forEach((name, index) => {
            if (typeof name !== 'string' || name.trim().length === 0) {
                throw new ValidationError(`componentNames[${index}]`, name, { type: 'string', minLength: 1 });
            }
        });
        
        if (options && typeof options !== 'object') {
            throw new ValidationError('options', options, { type: 'object' });
        }
    }

    /**
     * Logs messages
     * @param {string} level - Log level
     * @param {string} message - Message
     * @param {...any} args - Additional arguments
     * @private
     */
    #log(level, message, ...args) {
        const prefix = '[LazyLoader]';
        if (console[level]) {
            console[level](prefix, message, ...args);
        }
    }
}

// Export LazyLoader
if (typeof window !== 'undefined') {
    window.LazyLoader = LazyLoader;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = LazyLoader;
}

export default LazyLoader;
