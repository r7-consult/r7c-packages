/*
 * ADR References:
 * - ADR-009-ai-vba-converter
 * - ADR-011-openai-integration-patterns
 * - ADR-032-plugin-ai-chat-and-settings
 * - ADR-035-ai-agent-architecture
 *
 * TASK-051: AI-Powered VBA Conversion - Phase 1 (Base Provider)
 * TASK-050: Rebuild AI Provider Using Official SDK Patterns
 * TASK-055: OpenAI-only support (removed Claude and Gemini)
 * ADR-009: AI-Powered VBA Conversion Architecture
 *
 * Base AI provider class for unified interface to OpenAI.
 * Simplified from AI Plugin to focus on text completion for VBA conversion.
 *
 * External Dependencies:
 * - Fetch API: https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API
 * - AI Storage: ./storage.js
 * - Fetch Polyfill: ./fetch-polyfill.js
 * - AI Errors: ./errors.js (TASK-050)
 *
 * Architecture References:
 * - ADR-008: System Role Transformation Pattern
 * - AI Plugin: scripts/engine/providers/provider.js
 * - OpenAI SDK: https://github.com/openai/openai-node (v6.3.0)
 *
 * @task_id TASK-051, TASK-050, TASK-055
 * @adr_ref ADR-009
 * @coding_standard Adheres to: .memory_bank/guides/coding_standards.md
 * @history
 *  - 2025-10-13: Created by Dev-Agent - TASK-051: Simplified base provider for VBA conversion
 *  - 2025-10-14: Updated by Dev-Agent - TASK-050: Added error hierarchy and request ID tracking from official SDKs
 *  - 2025-10-14: Updated by Claude - TASK-055: Removed Claude and Gemini support, OpenAI-only
 *  - 2025-10-15: Updated by Claude - q2.log Issue #3: Added AscSimpleRequest for Desktop CORS bypass (file:// origins)
 *  - 2025-10-15: Updated by Claude - q2.log Issue #4: Fixed JSON body formatting with null check (line 240)
 *  - 2025-10-15: Updated by Claude - q2.log Issue #5: Changed default model to gpt-4 (API key access issue)
 *  - 2025-10-15: Fixed by Dev - q2.log JSON parse: disable AscSimpleRequest and use fetch polyfill
 */

// =============================================================================
// 1. IMPORTS AND DEPENDENCIES
// =============================================================================
// Requires: errors.js, storage.js, fetch-polyfill.js to be loaded first

// =============================================================================
// 2. CONSTANTS AND CONFIGURATION
// =============================================================================

const DEFAULT_TIMEOUT = 30000; // 30 seconds
const DEFAULT_MAX_TOKENS = 4096;

// Temporary mitigation flag for Desktop mode
// AscSimpleRequest path led to OpenAI "could not parse JSON body" (tmp/q2.log)
// Route all requests through fetch/fetch-polyfill instead.
const USE_ASC_SIMPLE_REQUEST = false;  // TASK-068: Disable buggy AscSimpleRequest, use fetch polyfill

// =============================================================================
// 3. DESKTOP MODE DETECTION (CORS Bypass)
// =============================================================================

function isFileProtocolContext() {
	if (window.location && window.location.protocol === 'file:') {
		return true;
	}
	if (window.document && window.document.currentScript &&
	    typeof window.document.currentScript.src === 'string' &&
	    window.document.currentScript.src.indexOf('file:///') === 0) {
		return true;
	}
	return false;
}

function hasDesktopBridge() {
	return !!(
		typeof window.AscDesktopEditor !== 'undefined' ||
		typeof window.AscSimpleRequest !== 'undefined' ||
		(window.Asc && window.Asc.plugin)
	);
}

/**
 * Detect OnlyOffice Desktop mode (file:// origins that need CORS bypass)
 * Pattern from AI Plugin: scripts/engine/storage.js:41-59
 */
const isLocalDesktop = (function() {
	const userAgent = (window.navigator && typeof window.navigator.userAgent === 'string')
		? window.navigator.userAgent.toLowerCase()
		: '';
	const hasDesktopUserAgent = userAgent.indexOf('ascdesktopeditor') >= 0;

	if (!isFileProtocolContext()) {
		return false;
	}

	// Linux desktop builds are not guaranteed to expose the same user-agent
	// token as Windows/macOS, so also trust the host bridge objects.
	return hasDesktopUserAgent || hasDesktopBridge();
})();

/**
 * Check if we should use AscSimpleRequest for non-streamed requests
 * This bypasses browser CORS by routing through OnlyOffice application layer
 */
const isLocalDesktopForNotStreamedRequests = (function() {
	if (isLocalDesktop)
		return true;
	if (window.location && window.location.protocol === "onlyoffice:")
		return true;  // OnlyOffice custom protocol
	return false;
})();

function isDesktopLikeRuntime() {
	return isLocalDesktopForNotStreamedRequests === true ||
		isLocalDesktop === true ||
		(window.AIFetchPolyfill && window.AIFetchPolyfill.isDesktopMode === true) ||
		(hasDesktopBridge() && isFileProtocolContext());
}

// Log Desktop mode detection for debugging
if (isLocalDesktopForNotStreamedRequests) {
	if (USE_ASC_SIMPLE_REQUEST) {
		console.log('[AI Engine] Desktop mode: using AscSimpleRequest for CORS bypass');
	} else {
		console.log('[AI Engine] Desktop mode: using fetch() with polyfill (AscSimpleRequest disabled)');
	}
} else {
	console.log('[AI Engine] Online mode detected - using standard fetch() API');
}

// =============================================================================
// 4. BASE PROVIDER CLASS
// =============================================================================

/**
 * AIProvider - Base class for AI service providers
 *
 * Provides unified interface for OpenAI API.
 * Subclasses override specific methods for provider-specific behavior.
 */
class AIProvider {
	/**
	 * Create AI provider instance
	 * @param {string} name - Provider name (OpenAI)
	 * @param {string} baseUrl - Base API URL
	 * @param {string} apiVersion - API version (v1, v1beta, etc.)
	 */
	constructor(name, baseUrl, apiVersion = 'v1') {
		this.name = name;
		this.baseUrl = baseUrl;
		this.apiVersion = apiVersion;
		this.apiKey = null;
		this.timeout = DEFAULT_TIMEOUT;
	}

	// =========================================================================
	// 5. API KEY MANAGEMENT
	// =========================================================================

	/**
	 * Set API key for this provider
	 * @param {string} apiKey - API key
	 */
	setApiKey(apiKey) {
		this.apiKey = apiKey;
	}

	/**
	 * Get API key from storage or instance
	 * @returns {string|null} API key or null
	 */
	getApiKey() {
		if (this.apiKey) return this.apiKey;

		// Try to load from storage
		if (window.AIStorage) {
			const storageKey = this.name.toLowerCase().replace('-', '');
			return window.AIStorage.getApiKey(storageKey);
		}

		return null;
	}

	/**
	 * Check if provider has valid API key
	 * @returns {boolean} True if API key is configured
	 */
	hasApiKey() {
		return this.getApiKey() !== null;
	}

	// =========================================================================
	// 6. REQUEST BUILDING
	// =========================================================================

	/**
	 * Build full API endpoint URL
	 * @param {string} endpoint - Endpoint path (e.g., '/chat/completions')
	 * @returns {string} Full URL
	 */
	buildUrl(endpoint) {
		const base = (this.baseUrl || '').replace(/\/+$/, '');
		const rawEndpoint = (typeof endpoint === 'string') ? endpoint : '';
		const endpointPath = rawEndpoint
			? (rawEndpoint.indexOf('/') === 0 ? rawEndpoint : `/${rawEndpoint}`)
			: '';
		const version = (this.apiVersion == null)
			? ''
			: String(this.apiVersion).replace(/^\/+|\/+$/g, '');

		if (!version) {
			return `${base}${endpointPath}`;
		}

		return `${base}/${version}${endpointPath}`;
	}

	/**
	 * Build request headers
	 * @returns {Object} Headers object
	 */
    buildHeaders() {
        const headers = {
            'Content-Type': 'application/json; charset=utf-8',
            'Accept': 'application/json'
        };

		const apiKey = this.getApiKey();
		if (apiKey) {
			headers['Authorization'] = `Bearer ${apiKey}`;
		}

		return headers;
	}

	/**
	 * Build chat completion request body
	 * @param {Array} messages - Array of message objects {role, content}
	 * @param {Object} options - Additional options (model, temperature, etc.)
	 * @returns {Object} Request body
	 */
	buildChatRequest(messages, options = {}) {
		let defaultMaxTokens = DEFAULT_MAX_TOKENS;
		let defaultTemperature = 0.7;
		try {
			if (typeof window !== 'undefined' && window.AIConfiguration) {
				const maxFromConfig = window.AIConfiguration.get?.('defaultMaxTokens');
				if (typeof maxFromConfig === 'number' && Number.isFinite(maxFromConfig) && maxFromConfig > 0) {
					defaultMaxTokens = maxFromConfig;
				}
				const tempFromConfig = window.AIConfiguration.get?.('defaultTemperature');
				if (typeof tempFromConfig === 'number' && Number.isFinite(tempFromConfig)) {
					defaultTemperature = tempFromConfig;
				}
			}
		} catch (_) {}

		const base = {
			model: options.model || 'gpt-4',
			messages: messages,
			max_tokens: options.max_tokens || defaultMaxTokens,
			temperature: typeof options.temperature === 'number'
				? options.temperature
				: defaultTemperature
		};

		return {
			...base,
			...options
		};
	}

	// =========================================================================
	// 7. API COMMUNICATION
	// =========================================================================

	_createAbortContext(externalSignal) {
		const ctx = {
			controller: new AbortController(),
			abortReason: 'timeout',
			timeoutId: null,
			detachPluginAbort: null,
			detachExternalAbort: null
		};
		ctx.timeoutId = setTimeout(() => {
			ctx.abortReason = 'timeout';
			ctx.controller.abort();
		}, this.timeout);
		const pluginSignal = window.AbortUtils?.getPluginAbortSignal?.() || null;
		ctx.detachPluginAbort = window.AbortUtils?.linkAbortController
			? window.AbortUtils.linkAbortController(ctx.controller, pluginSignal, {
				onAbort: () => {
					ctx.abortReason = 'shutdown';
				}
			})
			: null;
		ctx.detachExternalAbort = window.AbortUtils?.linkAbortController && externalSignal
			? window.AbortUtils.linkAbortController(ctx.controller, externalSignal, {
				onAbort: () => {
					if (ctx.abortReason !== 'shutdown') {
						ctx.abortReason = 'external';
					}
				}
			})
			: null;
		return ctx;
	}

	_throwTransportError(error, abortReason) {
		if (error && error.name === 'AbortError') {
			if (abortReason === 'shutdown') {
				throw new window.AIErrors.APIConnectionError(`[${this.name}] Request aborted during plugin shutdown.`);
			}
			if (abortReason === 'external') {
				throw new window.AIErrors.APIConnectionError(`[${this.name}] Request aborted.`);
			}
			throw new window.AIErrors.APIConnectionError(`[${this.name}] Request timeout after ${this.timeout}ms. Please check your internet connection.`);
		}

		if (error instanceof window.AIErrors.AIError) {
			throw error;
		}

		throw new window.AIErrors.APIConnectionError(
			`[${this.name}] Network error: ${error.message}. Please check your internet connection.`
		);
	}

	/**
	 * Send HTTP request to AI provider with Desktop/Online mode support
	 *
	 * **Desktop Mode (file:// origins)**:
	 * - Uses window.AscSimpleRequest.createRequest() to bypass browser CORS
	 * - Requests route through OnlyOffice application layer, not browser
	 * - Authorization headers work because they bypass browser security sandbox
	 * - Pattern from AI Plugin: scripts/engine/engine.js:197-211
	 *
	 * **Online Mode (http:// origins)**:
	 * - Uses standard fetch() API with polyfill fallback
	 * - Standard browser CORS applies
	 * - fetch-polyfill.js provides XMLHttpRequest compatibility for older Desktop versions
	 *
	 * @param {string} url - Full URL
	 * @param {Object} body - Request body
	 * @returns {Promise<Object>} Response data
	 * @throws {AuthenticationError} - 401 Invalid API key
	 * @throws {RateLimitError} - 429 Rate limit or quota exceeded
	 * @throws {InvalidRequestError} - 400 Bad request
	 * @throws {PermissionDeniedError} - 403 Forbidden
	 * @throws {NotFoundError} - 404 Not found
	 * @throws {InternalServerError} - 500+ Server errors
	 * @throws {APIConnectionError} - Network or timeout errors
	 */
	async sendRequest(url, body, options = {}) {
		if (!this.hasApiKey()) {
			throw new window.AIErrors.AuthenticationError(
				`[${this.name}] API key not configured. Please set your API key in settings.`
			);
		}

		// Optional advanced logging (per-connection)
		const advancedLogging = this.__aiAdvancedLogging === true;
		if (advancedLogging) {
			try {
				const headers = this.buildHeaders();
				const safeHeaders = { ...headers };
				if (safeHeaders.Authorization) {
					safeHeaders.Authorization = '***';
				}
				let bodySummary = null;
				if (body && typeof body === 'object') {
					const messages = Array.isArray(body.messages)
						? body.messages.map((m) => ({
								role: m && m.role,
								length: typeof m?.content === 'string' ? m.content.length : 0
						  }))
						: undefined;
					bodySummary = {
						model: body.model,
						temperature: body.temperature,
						max_tokens: body.max_tokens,
						messages
					};
				}
				const payload = {
					provider: this.name,
					url,
					connectionId: this.__aiConnectionId || null,
					headers: safeHeaders,
					bodySummary
				};
				const tag = this.name === 'GitHub' ? 'GitHubAI' : 'AIProvider';
				if (window.debug && typeof window.debug.debug === 'function') {
					window.debug.debug(tag, 'advanced-request', payload);
				} else {
					console.debug(`[${tag}] advanced-request`, payload);
				}
			} catch (_) {}
		}

		// ===================================================================
		// DESKTOP MODE: Use AscSimpleRequest to bypass CORS
		// ===================================================================
		if (USE_ASC_SIMPLE_REQUEST && isLocalDesktopForNotStreamedRequests && typeof window.AscSimpleRequest !== 'undefined') {
			return new Promise((resolve, reject) => {
				const timeoutId = setTimeout(() => {
					reject(new window.AIErrors.APIConnectionError(
						`[${this.name}] Request timeout after ${this.timeout}ms. Please check your internet connection.`
					));
				}, this.timeout);

				// Debug logging to see exact request being sent
				console.log('[AI Engine] AscSimpleRequest - URL:', url);
				console.log('[AI Engine] AscSimpleRequest - Headers:', this.buildHeaders());
				console.log('[AI Engine] AscSimpleRequest - Body object:', body);
				const bodyString = body ? JSON.stringify(body) : "";
				console.log('[AI Engine] AscSimpleRequest - Body string:', bodyString);
				console.log('[AI Engine] AscSimpleRequest - Body string length:', bodyString.length);

				try {
					window.AscSimpleRequest.createRequest({
						url: url,
						method: 'POST',
						headers: this.buildHeaders(),
						body: bodyString,
						complete: (e, status) => {
							clearTimeout(timeoutId);
                    try {
                        const rawText = e && e.responseText != null ? String(e.responseText) : '';
                        const data = JSON.parse(rawText);

								// Check for API error responses
								if (data.error) {
									const errorMessage = data.error.message || data.error.code || 'Unknown API error';
									const errorType = data.error.type || 'api_error';

									// Map error types to specific error classes
									if (errorType === 'insufficient_quota' || status === 429) {
										reject(new window.AIErrors.RateLimitError(
											`[${this.name}] ${errorMessage}`,
											null,
											errorType
										));
									} else if (status === 401) {
										reject(new window.AIErrors.AuthenticationError(
											`[${this.name}] ${errorMessage}`
										));
									} else if (status === 400) {
										reject(new window.AIErrors.InvalidRequestError(
											`[${this.name}] ${errorMessage}`
										));
									} else {
										reject(new window.AIErrors.AIError(
											`[${this.name}] ${errorMessage}`,
											status
										));
									}
									return;
								}

								resolve(data);
                    } catch (parseError) {
                        console.error('[AI Engine][AscSimpleRequest] JSON parse failed. Status:', status, 'Raw:', e && e.responseText);
                        reject(new window.AIErrors.APIConnectionError(
                            `[${this.name}] Failed to parse response: ${parseError.message}`
                        ));
                    }
						},
						error: (e, status, error) => {
							clearTimeout(timeoutId);
							const statusCode = (e.statusCode === -102) ? 404 : e.statusCode;
							reject(new window.AIErrors.APIConnectionError(
								`[${this.name}] Request failed (status ${statusCode}): ${error || 'Internal error'}`
							));
						}
					});
				} catch (error) {
					clearTimeout(timeoutId);
					reject(new window.AIErrors.APIConnectionError(
						`[${this.name}] Failed to create request: ${error.message}`
					));
				}
			});
		}

			// ===================================================================
			// ONLINE MODE: Use standard fetch() API
			// ===================================================================
			const abortContext = this._createAbortContext(options && options.signal);
			const controller = abortContext.controller;

			try {
				const response = await fetch(url, {
					method: 'POST',
					headers: this.buildHeaders(),
					body: JSON.stringify(body),
					signal: controller.signal
				});

				clearTimeout(abortContext.timeoutId);
				try { abortContext.detachPluginAbort?.(); } catch (_) {}
				try { abortContext.detachExternalAbort?.(); } catch (_) {}

			// Extract request ID from response headers (OpenAI SDK pattern)
			const requestId = response.headers.get('x-request-id') ||
			                 response.headers.get('cf-ray') ||
			                 response.headers.get('request-id') ||
			                 null;

			if (!response.ok) {
				const errorText = await response.text();
				let errorData = null;

				try {
					errorData = JSON.parse(errorText);
				} catch (e) {
					// Not JSON, use raw text
				}

				// Extract error message (support multiple API formats)
				const errorMessage = errorData?.error?.message ||
				                    errorData?.message ||
				                    errorText ||
				                    'Unknown error';

				// Throw specific error types based on HTTP status code
				// Pattern from OpenAI SDK: https://github.com/openai/openai-node/blob/main/src/error.ts
				switch (response.status) {
					case 401:
						throw new window.AIErrors.AuthenticationError(
							`[${this.name}] Authentication failed: ${errorMessage}`,
							requestId
						);

					case 403:
						throw new window.AIErrors.PermissionDeniedError(
							`[${this.name}] Permission denied: ${errorMessage}`,
							requestId
						);

					case 404:
						throw new window.AIErrors.NotFoundError(
							`[${this.name}] Resource not found: ${errorMessage}`,
							requestId
						);

					case 409:
						throw new window.AIErrors.ConflictError(
							`[${this.name}] Conflict: ${errorMessage}`,
							requestId
						);

					case 422:
						throw new window.AIErrors.UnprocessableEntityError(
							`[${this.name}] Unprocessable entity: ${errorMessage}`,
							requestId
						);

					case 429:
						// Determine rate limit type from error data
						const errorType = errorData?.error?.type || 'rate_limit';
						throw new window.AIErrors.RateLimitError(
							`[${this.name}] Rate limit exceeded: ${errorMessage}`,
							requestId,
							errorType
						);

					case 400:
						throw new window.AIErrors.InvalidRequestError(
							`[${this.name}] Invalid request: ${errorMessage}`,
							requestId
						);

					case 500:
					case 502:
					case 503:
					case 504:
						throw new window.AIErrors.InternalServerError(
							`[${this.name}] Service error: ${errorMessage}`,
							response.status,
							requestId
						);

					default:
						throw new window.AIErrors.AIError(
							`[${this.name}] API error (${response.status}): ${errorMessage}`,
							response.status,
							requestId
						);
				}
			}

				return await response.json();
			} catch (error) {
				clearTimeout(abortContext.timeoutId);
				try { abortContext.detachPluginAbort?.(); } catch (_) {}
				try { abortContext.detachExternalAbort?.(); } catch (_) {}
				this._throwTransportError(error, abortContext.abortReason);
			}
		}

	/**
	 * Send chat completion request
	 * @param {Array} messages - Array of message objects {role, content}
	 * @param {Object} options - Additional options
	 * @returns {Promise<string>} AI response text
	 */
	    async sendMessage(messages, options = {}) {
	        const externalSignal = options && options.signal ? options.signal : null;
	        const safeOptions = { ...options };
	        if (Object.prototype.hasOwnProperty.call(safeOptions, 'signal')) {
	            delete safeOptions.signal;
	        }
	        const isDesktop = isDesktopLikeRuntime();
	        const useResponses = (typeof window !== 'undefined' && window.AIConfiguration)
	            ? window.AIConfiguration.getUseResponsesAPI()
	            : true;
        const desktopStrategy = (typeof window !== 'undefined' && window.AIConfiguration)
            ? window.AIConfiguration.getResponsesDesktopStrategy()
            : 'fallback';
        const enableStreaming = (typeof window !== 'undefined' && window.AIConfiguration)
            ? window.AIConfiguration.getEnableStreaming()
            : false;

        // Determine effective model and whether it is a Codex-family model
        let effectiveModel = (options && options.model) || ((typeof window !== 'undefined' && window.AIConfiguration)
            ? window.AIConfiguration.getDefaultModel('openai')
            : 'gpt-4o-mini');
        try {
            if (typeof window !== 'undefined' && window.AIModelNormalizer) {
                const normalized = window.AIModelNormalizer.normalize(effectiveModel, 'openai');
                if (normalized && normalized.canonical) effectiveModel = normalized.canonical;
            }
        } catch (_) {}
        const isCodexModel = typeof effectiveModel === 'string' && /codex/i.test(effectiveModel);

        // TASK-072: Route based on configuration and Desktop mode, not model type
        // Responses API has limited parameter support but is required for certain models
        const canUseResponses = useResponses && (!isDesktop || desktopStrategy === 'nonstream');
        const endpoint = canUseResponses ? '/responses' : '/chat/completions';
        const url = this.buildUrl(endpoint);

	        let requestBody;
	        if (endpoint === '/responses' && typeof window !== 'undefined' && window.ResponsesCompat) {
	            const opts = { ...safeOptions };
	            if (!isDesktop) {
	                opts.stream = enableStreaming === true;
	            } else {
	                opts.stream = false;
	            }
            // Ensure a model is present
            if (!opts.model) {
                opts.model = effectiveModel;
            }
	            requestBody = window.ResponsesCompat.buildPayload(messages, opts);
	        } else {
	            // Pass endpoint marker so provider overrides can adjust params
	            const opts = { ...safeOptions, _endpoint: endpoint };
	            if (!opts.model) {
	                opts.model = effectiveModel;
	            }
	            requestBody = this.buildChatRequest(messages, opts);
	        }

        // Streaming path for Responses API (online only)
	        if (endpoint === '/responses' && !isDesktop && enableStreaming === true && typeof window !== 'undefined' && window.ResponsesCompat) {
	            try {
	                return await this._sendResponsesStreaming(url, requestBody, null, externalSignal);
	            } catch (err) {
	                // Fallback to Chat Completions on Responses API failure
	                const url2 = this.buildUrl('/chat/completions');
	                const body2 = this.buildChatRequest(messages, { ...safeOptions, _endpoint: '/chat/completions' });
	                const resp2 = await this.sendRequest(url2, body2, { signal: externalSignal });
	                return this.extractResponse(resp2);
	            }
	        }

	        try {
	            const response = await this.sendRequest(url, requestBody, { signal: externalSignal });
	            if (endpoint === '/responses' && typeof window !== 'undefined' && window.ResponsesCompat) {
	                return window.ResponsesCompat.extractText(response);
	            }
	            return this.extractResponse(response);
	        } catch (err) {
	            if (endpoint === '/responses' && typeof window !== 'undefined' && window.ResponsesCompat) {
	                // Retry with conservative Responses payload before falling back to Chat Completions
	                try {
	                    const fallbackBody = window.ResponsesCompat.buildFallbackPayload(messages, { ...(safeOptions || {}), stream: false });
	                    const resp = await this.sendRequest(url, fallbackBody, { signal: externalSignal });
	                    return window.ResponsesCompat.extractText(resp);
	                } catch (err2) {
	                    // Fall through to Chat Completions
	                }
	                // Fallback to Chat Completions API
	                try {
	                    const url2 = this.buildUrl('/chat/completions');
	                    const body2 = this.buildChatRequest(messages, { ...safeOptions, _endpoint: '/chat/completions' });
	                    const resp2 = await this.sendRequest(url2, body2, { signal: externalSignal });
	                    return this.extractResponse(resp2);
	                } catch (err3) {
	                    throw err3;
	                }
	            }
	            throw err;
	        }
	    }

    /**
     * Streaming fetch for OpenAI Responses API using SSE
     * Aggregates text deltas and returns final text. Emits deltas via optional callback.
     * @private
     * @param {string} url
     * @param {Object} body
     * @param {Function} [onDelta] - Optional callback(deltaText)
     * @returns {Promise<string>} Final aggregated text
     */
		    async _sendResponsesStreaming(url, body, onDelta, externalSignal, returnMeta = false) {
		        const abortContext = this._createAbortContext(externalSignal);
		        const controller = abortContext.controller;
            let reader = null;
            const acc = window.ResponsesCompat.createAccumulator();
		        try {
		            const response = await fetch(url, {
	                method: 'POST',
	                headers: this.buildHeaders(),
	                body: JSON.stringify(body),
	                signal: controller.signal
	            });

            // Handle non-OK HTTP
            if (!response.ok) {
                const errorText = await response.text();
                let errorData = null;
                try { errorData = JSON.parse(errorText); } catch (_) {}
                const errorMessage = errorData?.error?.message || errorData?.message || errorText || 'Unknown error';
                switch (response.status) {
                    case 401: throw new window.AIErrors.AuthenticationError(`[${this.name}] Authentication failed: ${errorMessage}`);
                    case 403: throw new window.AIErrors.PermissionDeniedError(`[${this.name}] Permission denied: ${errorMessage}`);
                    case 404: throw new window.AIErrors.NotFoundError(`[${this.name}] Resource not found: ${errorMessage}`);
                    case 409: throw new window.AIErrors.ConflictError(`[${this.name}] Conflict: ${errorMessage}`);
                    case 422: throw new window.AIErrors.UnprocessableEntityError(`[${this.name}] Unprocessable entity: ${errorMessage}`);
                    case 429: throw new window.AIErrors.RateLimitError(`[${this.name}] Rate limit exceeded: ${errorMessage}`);
                    case 500: case 502: case 503: case 504:
                        throw new window.AIErrors.InternalServerError(`[${this.name}] Service error: ${errorMessage}`, response.status);
                    default:
                        throw new window.AIErrors.AIError(`[${this.name}] API error (${response.status}): ${errorMessage}`, response.status);
                }
            }

            // Desktop/polyfilled runtimes can return a fully buffered response
            // without a readable stream. In that case degrade to non-streaming
            // extraction instead of hard-failing on response.body.getReader().
            if (!response.body || typeof response.body.getReader !== 'function') {
                const payload = await response.json();
                acc.text = window.ResponsesCompat.extractText(payload);
                acc.responseId = typeof payload?.id === 'string' ? payload.id : '';
                acc.model = typeof payload?.model === 'string' ? payload.model : '';
                if (Array.isArray(payload?.output) && typeof window.ResponsesCompat.mapToolCall === 'function') {
                    payload.output.forEach((item) => {
                        const toolCall = window.ResponsesCompat.mapToolCall(item);
                        if (toolCall) acc.toolCalls.push(toolCall);
                    });
                }
                clearTimeout(abortContext.timeoutId);
                try { abortContext.detachPluginAbort?.(); } catch (_) {}
                try { abortContext.detachExternalAbort?.(); } catch (_) {}
                return returnMeta ? acc : (acc.text || '');
            }

            // Read SSE stream
            reader = response.body.getReader();
            const decoder = new TextDecoder();
            let buffer = '';

            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                buffer += decoder.decode(value, { stream: true });

                // Split on double newlines which separate events
                let idx;
                while ((idx = buffer.indexOf('\n\n')) !== -1) {
                    const rawEvent = buffer.slice(0, idx);
                    buffer = buffer.slice(idx + 2);

                    // Process lines that begin with 'data: '
                    const lines = rawEvent.split('\n');
                    for (const line of lines) {
                        if (!line.startsWith('data: ')) continue;
                        const data = line.slice(6).trim();
                        if (!data) continue;
                        if (data === '[DONE]') {
                            // Stream finished
                            break;
                        }
                        try {
                            const evt = JSON.parse(data);
                            // Emit delta text chunks if present
                            if (evt && evt.type === 'response.output_text.delta' && typeof evt.delta === 'string') {
                                if (typeof onDelta === 'function') onDelta(evt.delta);
                            }
                            window.ResponsesCompat.accumulateEvent(evt, acc);
                        } catch (e) {
                            // Ignore malformed line
                        }
                    }
                }
            }

	            clearTimeout(abortContext.timeoutId);
	            try { abortContext.detachPluginAbort?.(); } catch (_) {}
	            try { abortContext.detachExternalAbort?.(); } catch (_) {}
	            try { await reader.releaseLock?.(); } catch (_) {}
	            return returnMeta ? acc : (acc.text || '');
	        } catch (error) {
	            clearTimeout(abortContext.timeoutId);
	            try { abortContext.detachPluginAbort?.(); } catch (_) {}
	            try { abortContext.detachExternalAbort?.(); } catch (_) {}
	            try { await reader?.cancel(); } catch (_) {}
	            this._throwTransportError(error, abortContext.abortReason);
	        }
	    }

    /**
     * Public streaming API. Falls back to non-stream when not available.
     * @param {Array} messages
     * @param {Object} options
     * @param {Function} onDelta - Receives text delta chunks
     * @returns {Promise<string>} Final text
     */
	    async sendMessageStream(messages, options = {}, onDelta) {
	        const externalSignal = options && options.signal ? options.signal : null;
	        const safeOptions = { ...options };
	        if (Object.prototype.hasOwnProperty.call(safeOptions, 'signal')) {
	            delete safeOptions.signal;
	        }
	        const isDesktop = isDesktopLikeRuntime();
	        const useResponses = (typeof window !== 'undefined' && window.AIConfiguration)
	            ? window.AIConfiguration.getUseResponsesAPI()
	            : true;
        const enableStreaming = (typeof window !== 'undefined' && window.AIConfiguration)
            ? window.AIConfiguration.getEnableStreaming()
            : false;

	        if (!isDesktop && useResponses && enableStreaming && typeof window !== 'undefined' && window.ResponsesCompat) {
	            const url = this.buildUrl('/responses');
	            const opts = { ...safeOptions, stream: true };
	            if (!opts.model && typeof window !== 'undefined' && window.AIConfiguration) {
	                opts.model = window.AIConfiguration.getDefaultModel('openai');
	            }
	            const requestBody = window.ResponsesCompat.buildPayload(messages, opts);
	            const returnMeta = safeOptions && safeOptions.returnMeta === true;
	            return await this._sendResponsesStreaming(url, requestBody, onDelta, externalSignal, returnMeta);
	        }
	        // Fallback to non-stream
	        return await this.sendMessage(messages, { ...safeOptions, signal: externalSignal });
	    }

	// =========================================================================
	// 8. RESPONSE PROCESSING
	// =========================================================================

	/**
	 * Extract text content from AI response
	 * @param {Object} response - Raw API response
	 * @returns {string} Extracted text
	 */
	extractResponse(response) {
		// Check for API error responses first
		if (response.error) {
			const errorMessage = response.error.message || response.error.code || 'Unknown API error';
			throw new Error(`[${this.name}] API Error: ${errorMessage}`);
		}

		// Standard OpenAI-compatible response format
		if (response.choices && response.choices.length > 0) {
			const choice = response.choices[0];

			if (choice.message && choice.message.content) {
				return choice.message.content.trim();
			}

			if (choice.text) {
				return choice.text.trim();
			}
		}

		// Fallback for other formats
		if (response.content) {
			return response.content.trim();
		}

		// More helpful error message with response details
		console.error(`[${this.name}] Unexpected response format. Response:`, response);
		throw new Error(`[${this.name}] Unexpected response format. Check console for details.`);
	}

	// =========================================================================
	// 9. HELPER METHODS
	// =========================================================================

	/**
	 * Create chat messages array from system prompt and user input
	 * @param {string} systemPrompt - System role instruction
	 * @param {string} userInput - User's input text
	 * @returns {Array} Messages array
	 */
	createMessages(systemPrompt, userInput) {
		const messages = [];

		if (systemPrompt) {
			messages.push({
				role: 'system',
				content: systemPrompt
			});
		}

		messages.push({
			role: 'user',
			content: userInput
		});

		return messages;
	}

	/**
	 * Extract system message from messages array
	 * @param {Array} messages - Messages array
	 * @param {boolean} remove - If true, remove system message from array
	 * @returns {string} System message content or empty string
	 */
	extractSystemMessage(messages, remove = false) {
		const systemIndex = messages.findIndex(m => m.role === 'system');

		if (systemIndex === -1) {
			return '';
		}

		const systemMessage = messages[systemIndex].content;

		if (remove) {
			messages.splice(systemIndex, 1);
		}

		return systemMessage;
	}
}

// =============================================================================
// 9A. OPENAI-COMPATIBLE CHAT PROVIDER BASE
// =============================================================================

	class OpenAICompatibleChatProvider extends AIProvider {
		static normalizeBaseUrl(options, fallbackUrl) {
			const opts = options || {};
			const baseFromOptions = opts.baseUrl || opts.baseURL;
			if (typeof baseFromOptions === 'string' && baseFromOptions.trim()) {
				return baseFromOptions.trim().replace(/\/+$/, '');
			}
			return fallbackUrl;
		}

		constructor(name, baseUrl, apiVersion = 'v1', options = {}) {
			super(name, baseUrl, apiVersion);

			const opts = options || {};
		this.fallbackMaxTokens = (typeof opts.fallbackMaxTokens === 'number' && Number.isFinite(opts.fallbackMaxTokens) && opts.fallbackMaxTokens > 0)
			? opts.fallbackMaxTokens
			: DEFAULT_MAX_TOKENS;

		if (opts.timeout) {
			this.timeout = opts.timeout;
		} else if (typeof window !== 'undefined' && window.AIConfiguration) {
			this.timeout = window.AIConfiguration.getTimeout();
		}

		try {
			if (typeof window !== 'undefined' && window.AILogger) {
				window.AILogger.info(this.name, `Provider initialized with base URL: ${this.baseUrl}, timeout: ${this.timeout}ms`);
			}
		} catch (_) {}
	}

	resolveDefaultModelFromConfig(providerId, fallbackModel) {
		try {
			if (typeof window !== 'undefined' &&
				window.AIConfiguration &&
				typeof window.AIConfiguration.getDefaultModel === 'function') {
				const configured = window.AIConfiguration.getDefaultModel(providerId);
				if (configured && typeof configured === 'string') return configured;
			}
		} catch (_) {}
		return fallbackModel;
	}

	getDefaultModel() {
		return this.defaultModel;
	}

	getAvailableModels() {
		const source = this.models || {};
		return Object.keys(source).map((id) => {
			const info = source[id] || {};
			return {
				id,
				name: info.name || id,
				maxTokens: typeof info.maxTokens === 'number' ? info.maxTokens : this.fallbackMaxTokens
			};
		});
	}

	buildChatRequest(messages, options = {}) {
		const model = options.model || this.getDefaultModel();
		const body = super.buildChatRequest(messages, { ...options, model });
		return body;
	}

		async sendMessage(messages, options = {}) {
			const externalSignal = options && options.signal ? options.signal : null;
			const safeOptions = { ...options };
			if (Object.prototype.hasOwnProperty.call(safeOptions, 'signal')) {
				delete safeOptions.signal;
			}
			const model = safeOptions.model || this.getDefaultModel();
			const url = this.buildUrl('/chat/completions');
			const body = this.buildChatRequest(messages, { ...safeOptions, model });
			const response = await this.sendRequest(url, body, { signal: externalSignal });
			return this.extractResponse(response);
		}

	async sendMessageStream(messages, options = {}, onDelta) {
		const text = await this.sendMessage(messages, options);
		if (typeof onDelta === 'function' && text) {
			try { onDelta(text); } catch (_) {}
		}
		return text;
	}
}

// =============================================================================
// 10. MODULE EXPORTS
// =============================================================================

if (typeof window !== 'undefined') {
	window.AIProvider = AIProvider;
	window.OpenAICompatibleChatProvider = OpenAICompatibleChatProvider;
	window.isLocalDesktop = isLocalDesktop;  // Export for debugging
	window.isLocalDesktopForNotStreamedRequests = isLocalDesktopForNotStreamedRequests;  // Export for debugging
}
