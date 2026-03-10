/*
 * TASK-058: OpenAI Request/Response Logging to External Service
 * Logging AI Provider - Decorator pattern wrapper for AI providers
 *
 * Wraps any AIProvider instance and logs all requests/responses to
 * external logging service. Uses decorator pattern for clean separation
 * of concerns without modifying existing provider code.
 *
 * External Dependencies:
 * - AIProvider: ../ai-provider.js (base class)
 * - ExternalLogService: ../services/external-log-service.js
 * - AISecurityFilter: ../utils/security-filter.js
 *
 * Architecture:
 * - Decorator pattern (wraps existing provider)
 * - Transparent pass-through of all methods
 * - Intercepts sendMessage() for logging
 * - Non-blocking (logging failures don't affect AI calls)
 *
 * @task_id TASK-058
 * @coding_standard Adheres to: .memory_bank/guides/coding_standards.md
 * @history
 *  - 2025-10-14: Created by Claude - TASK-058: Logging provider decorator (METHOD 3)
 */

// =============================================================================
// 1. IMPORTS AND DEPENDENCIES
// =============================================================================
// Requires: ai-provider.js, external-log-service.js, security-filter.js

// =============================================================================
// 2. LOGGING AI PROVIDER CLASS (DECORATOR)
// =============================================================================

/**
 * LoggingAIProvider - Decorator that adds logging to any AIProvider
 *
 * Wraps an existing AIProvider instance and intercepts sendMessage() calls
 * to log requests and responses to external service.
 *
 * Usage:
 *   const baseProvider = new OpenAIProvider();
 *   const loggingProvider = new LoggingAIProvider(baseProvider);
 *   // Use loggingProvider instead of baseProvider - all calls logged
 *
 * Design:
 * - Transparent wrapper (same interface as wrapped provider)
 * - Non-blocking logging (failures don't break AI functionality)
 * - Security filtering (API keys removed before logging)
 * - Fire-and-forget (no await on logging)
 */
class LoggingAIProvider {
    #wrappedProvider;
    #logService;
    #securityFilter;
    #enabled = true;

    /**
     * Creates a logging wrapper around an AI provider
     *
     * @param {AIProvider} provider - The AI provider to wrap
     * @param {ExternalLogService} logService - Optional custom log service
     */
    constructor(provider, logService = null) {
        if (!provider) {
            throw new Error('LoggingAIProvider requires a provider to wrap');
        }

        this.#wrappedProvider = provider;

        // Initialize log service
        if (logService) {
            this.#logService = logService;
        } else if (typeof window !== 'undefined' && window.ExternalLogService) {
            this.#logService = new window.ExternalLogService();
        } else {
            console.warn('[LoggingAIProvider] ExternalLogService not available, logging disabled');
            this.#enabled = false;
        }

        // Initialize security filter
        if (typeof window !== 'undefined' && window.AISecurityFilter) {
            this.#securityFilter = window.AISecurityFilter;
        } else {
            console.warn('[LoggingAIProvider] AISecurityFilter not available, logging disabled');
            this.#enabled = false;
        }
    }

    // =========================================================================
    // 3. CORE INTERCEPTION - SENDMESSAGE WITH LOGGING
    // =========================================================================

    /**
     * Sends message to AI provider with request/response logging
     * Intercepts AIProvider.sendMessage() to add logging
     *
     * @param {Array} messages - Array of message objects {role, content}
     * @param {Object} options - Additional options
     * @returns {Promise<string>} AI response text
     */
    async sendMessage(messages, options = {}) {
        const startTime = Date.now();
        let response = null;
        let error = null;

        try {
            // Call wrapped provider
            response = await this.#wrappedProvider.sendMessage(messages, options);

            // Log successful interaction
            this.#logInteraction(messages, options, response, startTime);

            return response;

        } catch (err) {
            error = err;

            // Log failed interaction
            this.#logInteraction(messages, options, null, startTime, error);

            // Re-throw error (logging must not suppress errors)
            throw err;
        }
    }

    /**
     * Logs OpenAI interaction to external service
     * @private
     * @param {Array} messages - Request messages
     * @param {Object} options - Request options
     * @param {string} response - Response text (null if error)
     * @param {number} startTime - Request start timestamp
     * @param {Error} error - Error object (null if success)
     */
    #logInteraction(messages, options, response, startTime, error = null) {
        if (!this.#enabled || !this.#logService || !this.#securityFilter) {
            return;
        }

        try {
            const duration = Date.now() - startTime;

            // Prepare request data
            const requestData = {
                messages: messages,
                options: options,
                timestamp: new Date(startTime).toISOString(),
                provider: this.#wrappedProvider.name || 'unknown'
            };

            // Prepare response data
            const responseData = error ? {
                error: true,
                errorMessage: error.message,
                errorType: error.constructor.name,
                duration: duration
            } : {
                error: false,
                result: response,
                duration: duration,
                responseLength: response ? response.length : 0
            };

            // Sanitize sensitive data
            const interaction = this.#securityFilter.sanitizeInteraction({
                ask: requestData,
                answ: responseData
            });

            // Send to external log service (fire-and-forget)
            this.#logService.logOpenAIInteraction(interaction);

            try {
                const userMessage = Array.isArray(interaction.ask?.messages)
                    ? interaction.ask.messages.find((entry) => entry && entry.role === 'user')
                    : null;
                const preview = (userMessage?.content || '').slice(0, 160);
                const summary = {
                    provider: interaction.ask?.provider || 'unknown',
                    durationMs: responseData.duration,
                    error: responseData.error || false,
                    promptPreview: preview,
                    responsePreview: (responseData.result || '').slice(0, 160)
                };
                console.info('[LLM]', summary);
            } catch (consoleError) {
                console.warn('[LoggingAIProvider] Failed to emit console summary:', consoleError);
            }

            if (window.debug) {
                window.debug.debug('LoggingAIProvider', 'Interaction logged', {
                    duration,
                    error: !!error
                });
            }

        } catch (logError) {
            // Never throw errors from logging - just warn
            console.error('[LoggingAIProvider] Failed to log interaction:', logError);
        }
    }

    // =========================================================================
    // 4. TRANSPARENT PASS-THROUGH OF ALL OTHER METHODS
    // =========================================================================

    /**
     * Converts VBA to JavaScript (pass-through to wrapped provider)
     * @param {string} vbaCode - VBA code to convert
     * @param {Object} options - Conversion options
     * @returns {Promise<string>} Converted JavaScript code
     */
    async convertVBAToJS(vbaCode, options = {}) {
        return await this.#wrappedProvider.convertVBAToJS(vbaCode, options);
    }

    // =========================================================================
    // 5. PROPERTY FORWARDING
    // =========================================================================

    // Forward all property access to wrapped provider
    get name() { return this.#wrappedProvider.name; }
    get baseUrl() { return this.#wrappedProvider.baseUrl; }
    get apiVersion() { return this.#wrappedProvider.apiVersion; }
    get timeout() { return this.#wrappedProvider.timeout; }
    get models() { return this.#wrappedProvider.models; }
    get defaultModel() { return this.#wrappedProvider.defaultModel; }

    setApiKey(apiKey) { return this.#wrappedProvider.setApiKey(apiKey); }
    getApiKey() { return this.#wrappedProvider.getApiKey(); }
    hasApiKey() { return this.#wrappedProvider.hasApiKey(); }

    // =========================================================================
    // 6. LOGGING CONTROL METHODS
    // =========================================================================

    /**
     * Enables logging
     */
    enableLogging() {
        if (this.#logService && this.#securityFilter) {
            this.#enabled = true;
            window.debug?.info('LoggingAIProvider', 'Logging enabled');
        } else {
            console.warn('[LoggingAIProvider] Cannot enable logging - dependencies not available');
        }
    }

    /**
     * Disables logging
     */
    disableLogging() {
        this.#enabled = false;
        window.debug?.info('LoggingAIProvider', 'Logging disabled');
    }

    /**
     * Checks if logging is enabled
     * @returns {boolean} True if logging is enabled
     */
    isLoggingEnabled() {
        return this.#enabled;
    }

    /**
     * Gets the wrapped provider instance
     * @returns {AIProvider} The wrapped provider
     */
    getWrappedProvider() {
        return this.#wrappedProvider;
    }

    /**
     * Gets the log service instance
     * @returns {ExternalLogService} The log service
     */
    getLogService() {
        return this.#logService;
    }
}

// =============================================================================
// 7. MODULE EXPORTS
// =============================================================================

// Global export for legacy compatibility
if (typeof window !== 'undefined') {
    window.LoggingAIProvider = LoggingAIProvider;
}

// ES6 module export (if supported)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = LoggingAIProvider;
}
