/**
 * @fileoverview OnlyOffice UI SDK ChatInterface Component
 * @description Advanced chat interface with provider management, conversation threading, and real-time messaging
 * @see {@link https://api.onlyoffice.com/plugin/basic} OnlyOffice Plugin API
 * @see {@link /CODE_STANDARD.MD} Plugin Architecture Guide
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 * @based-on AI Plugin patterns from r7-plugins-content analysis
 */

import { ErrorHandler, ComponentError, ValidationError } from '../../core/error-handler.js';

/**
 * ChatInterface Component for advanced chat functionality
 * Provides AI chat interface, provider management, conversation threading, and real-time messaging
 */
class ChatInterface {
    #element = null;
    #config = {};
    #state = {
        isConnected: false,
        currentProvider: null,
        conversations: new Map(),
        activeConversation: null,
        messages: [],
        isTyping: false,
        isLoading: false,
        providers: []
    };
    #eventSystem = null;
    #stateManager = null;
    #messageHistory = [];
    #typingTimer = null;
    #autoScroll = true;
    #errorHandler = null;

    /**
     * Creates a new ChatInterface instance
     * @param {Object} options - Component options
     * @param {HTMLElement|string} options.container - Container element or selector
     * @param {Array} [options.providers=[]] - Available chat providers
     * @param {string} [options.defaultProvider] - Default provider ID
     * @param {boolean} [options.multiConversation=true] - Enable multiple conversations
     * @param {boolean} [options.showTypingIndicator=true] - Show typing indicators
     * @param {boolean} [options.enableMarkdown=true] - Enable markdown rendering
     * @param {boolean} [options.enableCodeHighlight=true] - Enable code syntax highlighting
     * @param {number} [options.maxMessages=1000] - Maximum messages per conversation
     * @param {number} [options.maxConversations=50] - Maximum number of conversations
     * @param {Function} [options.onMessage] - Message callback function
     * @param {Function} [options.onProviderChange] - Provider change callback
     * @param {Function} [options.onConversationChange] - Conversation change callback
     * @param {Object} options.eventSystem - Event system instance
     * @param {Object} options.stateManager - State manager instance
     */
    constructor(options = {}) {
        try {
            // ERROR HANDLING: Initialize error handler first
            this.#errorHandler = new ErrorHandler({
                logLevel: options.debug ? 'debug' : 'error',
                enableRecovery: true
            });
            
            // VALIDATION: Validate constructor options
            this.#validateConstructorOptions(options);
            
            this.#config = {
                multiConversation: true,
                showTypingIndicator: true,
                enableMarkdown: true,
                enableCodeHighlight: true,
                maxMessages: 1000,
                maxConversations: 50,
                autoReconnect: true,
                messageBubbles: true,
                timestamps: true,
                avatars: true,
                searchable: true,
                exportable: true,
                providers: [],
                ...options
            };

            this.#eventSystem = options.eventSystem;
            this.#stateManager = options.stateManager;
            this.#state.providers = this.#config.providers;

            this.#setupContainer(options.container);
            this.#bindEvents();
        } catch (error) {
            throw new ComponentError('ChatInterface', 'constructor', error.message, error);
        }
    }

    /**
     * Initializes the ChatInterface component
     * @returns {Promise<void>}
     */
    async initialize() {
        try {
            this.#createChatInterfaceStructure();
            this.#setupEventListeners();
            
            // Initialize first conversation if multiConversation enabled
            if (this.#config.multiConversation) {
                this.createConversation('General', true);
            }

            // Set default provider if specified
            if (this.#config.defaultProvider) {
                this.setProvider(this.#config.defaultProvider);
            }

            this.#eventSystem?.emit('chat:initialized', {
                component: this,
                timestamp: new Date().toISOString()
            });

        } catch (error) {
            console.error('ChatInterface initialization failed:', error);
            throw error;
        }
    }

    /**
     * Sets available providers
     * @param {Array} providers - Array of provider objects
     * @param {string} providers[].id - Provider ID
     * @param {string} providers[].name - Provider display name
     * @param {string} [providers[].icon] - Provider icon
     * @param {Object} [providers[].config] - Provider configuration
     * @param {boolean} [providers[].enabled=true] - Whether provider is enabled
     */
    setProviders(providers) {
        this.#state.providers = providers.map(provider => ({
            id: provider.id,
            name: provider.name,
            icon: provider.icon || '🤖',
            config: provider.config || {},
            enabled: provider.enabled !== false,
            status: 'disconnected',
            ...provider
        }));

        this.#renderProviders();
        
        this.#eventSystem?.emit('chat:providers:updated', {
            providers: this.#state.providers,
            component: this
        });
    }

    /**
     * Sets the active provider
     * @param {string} providerId - Provider ID
     * @returns {Promise<boolean>} Success status
     */
    async setProvider(providerId) {
        const provider = this.#state.providers.find(p => p.id === providerId);
        if (!provider) {
            console.warn(`Provider ${providerId} not found`);
            return false;
        }

        if (!provider.enabled) {
            console.warn(`Provider ${providerId} is disabled`);
            return false;
        }

        try {
            // Disconnect current provider
            if (this.#state.currentProvider) {
                await this.#disconnectProvider(this.#state.currentProvider);
            }

            // Connect new provider
            this.#state.currentProvider = provider;
            await this.#connectProvider(provider);

            this.#updateProviderDisplay();
            
            this.#eventSystem?.emit('chat:provider:changed', {
                provider,
                previousProvider: this.#state.currentProvider,
                component: this
            });

            if (this.#config.onProviderChange) {
                this.#config.onProviderChange(provider);
            }

            return true;

        } catch (error) {
            console.error(`Failed to set provider ${providerId}:`, error);
            return false;
        }
    }

    /**
     * Creates a new conversation
     * @param {string} [title='New Conversation'] - Conversation title
     * @param {boolean} [activate=true] - Whether to activate the conversation
     * @returns {string} Conversation ID
     */
    createConversation(title = 'New Conversation', activate = true) {
        if (!this.#config.multiConversation) {
            console.warn('Multi-conversation mode is disabled');
            return null;
        }

        const conversationId = this.#generateId();
        const conversation = {
            id: conversationId,
            title,
            messages: [],
            createdAt: new Date(),
            updatedAt: new Date(),
            participants: [],
            metadata: {}
        };

        this.#state.conversations.set(conversationId, conversation);

        // Limit number of conversations
        if (this.#state.conversations.size > this.#config.maxConversations) {
            const oldestId = Array.from(this.#state.conversations.keys())[0];
            this.#state.conversations.delete(oldestId);
        }

        this.#renderConversations();

        if (activate) {
            this.setActiveConversation(conversationId);
        }

        this.#eventSystem?.emit('chat:conversation:created', {
            conversation,
            component: this
        });

        return conversationId;
    }

    /**
     * Sets the active conversation
     * @param {string} conversationId - Conversation ID
     */
    setActiveConversation(conversationId) {
        const conversation = this.#state.conversations.get(conversationId);
        if (!conversation) {
            console.warn(`Conversation ${conversationId} not found`);
            return;
        }

        this.#state.activeConversation = conversationId;
        this.#state.messages = conversation.messages;
        
        this.#renderMessages();
        this.#updateConversationDisplay();

        this.#eventSystem?.emit('chat:conversation:changed', {
            conversation,
            conversationId,
            component: this
        });

        if (this.#config.onConversationChange) {
            this.#config.onConversationChange(conversation);
        }
    }

    /**
     * Sends a message
     * @param {string|Object} message - Message content or message object
     * @param {string} [type='user'] - Message type (user, assistant, system)
     * @param {Object} [metadata={}] - Additional message metadata
     * @returns {Promise<string>} Message ID
     */
    async sendMessage(message, type = 'user', metadata = {}) {
        try {
            // ERROR HANDLING: Validate message parameters
            this.#validateMessageParameters(message, type, metadata);
            
            const messageObj = typeof message === 'string' ? 
                { content: message, type, metadata } : 
                { type, metadata, ...message };

            // SECURITY: Sanitize message content
            if (messageObj.content) {
                messageObj.content = this.#sanitizeMessageContent(messageObj.content);
            }

            const messageId = this.#generateId();
            const fullMessage = {
                id: messageId,
                content: messageObj.content,
                type: messageObj.type,
                timestamp: new Date(),
                sender: type === 'user' ? 'user' : (this.#state.currentProvider?.name || 'Assistant'),
                senderId: type === 'user' ? 'user' : this.#state.currentProvider?.id,
                metadata: messageObj.metadata,
                status: 'sending'
            };

        // Add to current conversation
        this.#addMessageToConversation(fullMessage);

        try {
            // Emit message event
            this.#eventSystem?.emit('chat:message:sending', {
                message: fullMessage,
                component: this
            });

            // Call message handler
            if (this.#config.onMessage) {
                const response = await this.#config.onMessage(fullMessage, this.#state.currentProvider);
                
                if (response && type === 'user') {
                    // Add assistant response
                    await this.sendMessage(response, 'assistant');
                }
            }

            // Update message status
            fullMessage.status = 'sent';
            this.#updateMessage(messageId, { status: 'sent' });

            this.#eventSystem?.emit('chat:message:sent', {
                message: fullMessage,
                component: this
            });

            return messageId;

        } catch (error) {
            // ERROR HANDLING: Comprehensive message error handling
            const messageError = new ComponentError(
                'ChatInterface',
                'sendMessage',
                `Failed to send message: ${error.message}`,
                error
            );
            
            // Update message status if it exists
            if (fullMessage) {
                fullMessage.status = 'error';
                fullMessage.error = error.message;
                this.#updateMessage(messageId, { status: 'error', error: error.message });
            }
            
            // Handle error with recovery if possible
            const recovered = await this.#errorHandler?.handleError(messageError, {
                message: messageObj,
                type,
                metadata,
                provider: this.#state.currentProvider?.id
            });
            
            // Emit error event
            this.#eventSystem?.emit('chat:message:error', {
                message: fullMessage,
                error: messageError,
                component: this,
                recovered
            });
            
            if (!recovered) {
                throw messageError;
            }
            
            return null; // Return null if recovered but message couldn't be sent
        }
    }

    /**
     * Shows typing indicator
     * @param {string} [sender] - Sender name
     * @param {number} [duration=0] - Duration in ms (0 = indefinite)
     */
    showTypingIndicator(sender, duration = 0) {
        if (!this.#config.showTypingIndicator) return;

        this.#state.isTyping = true;
        this.#renderTypingIndicator(sender);

        if (duration > 0) {
            setTimeout(() => {
                this.hideTypingIndicator();
            }, duration);
        }
    }

    /**
     * Hides typing indicator
     */
    hideTypingIndicator() {
        this.#state.isTyping = false;
        this.#removeTypingIndicator();
    }

    /**
     * Clears current conversation
     */
    clearConversation() {
        if (this.#state.activeConversation) {
            const conversation = this.#state.conversations.get(this.#state.activeConversation);
            if (conversation) {
                conversation.messages = [];
                this.#state.messages = [];
                this.#renderMessages();
                
                this.#eventSystem?.emit('chat:conversation:cleared', {
                    conversationId: this.#state.activeConversation,
                    component: this
                });
            }
        }
    }

    /**
     * Exports conversation
     * @param {string} [format='json'] - Export format (json, txt, html)
     * @param {string} [conversationId] - Conversation ID (current if not specified)
     * @returns {string} Exported data
     */
    exportConversation(format = 'json', conversationId) {
        const convId = conversationId || this.#state.activeConversation;
        const conversation = this.#state.conversations.get(convId);
        
        if (!conversation) {
            throw new Error('No conversation to export');
        }

        switch (format.toLowerCase()) {
            case 'json':
                return JSON.stringify(conversation, null, 2);
            case 'txt':
                return this.#exportAsText(conversation);
            case 'html':
                return this.#exportAsHTML(conversation);
            default:
                throw new Error(`Unsupported export format: ${format}`);
        }
    }

    /**
     * Searches messages
     * @param {string} query - Search query
     * @param {Object} [options={}] - Search options
     * @returns {Array} Matching messages
     */
    searchMessages(query, options = {}) {
        try {
            // ERROR HANDLING: Validate search parameters
            this.#validateSearchParameters(query, options);
            
            const messages = this.#state.messages;
            const { caseSensitive = false, includeMetadata = false } = options;
            
            // SECURITY: Sanitize search query
            const sanitizedQuery = this.#sanitizeSearchQuery(query);
            const searchTerm = caseSensitive ? sanitizedQuery : sanitizedQuery.toLowerCase();
            
            return messages.filter(message => {
                try {
                    const content = caseSensitive ? message.content : message.content.toLowerCase();
                    let matches = content.includes(searchTerm);
                    
                    if (!matches && includeMetadata && message.metadata) {
                        const metadataStr = JSON.stringify(message.metadata);
                        const metadataSearch = caseSensitive ? metadataStr : metadataStr.toLowerCase();
                        matches = metadataSearch.includes(searchTerm);
                    }
                    
                    return matches;
                } catch (error) {
                    // Log search error but don't fail entire search
                    console.warn('[ChatInterface] Error searching message:', message.id, error);
                    return false;
                }
            });
        } catch (error) {
            const searchError = new ComponentError(
                'ChatInterface',
                'searchMessages',
                `Search failed: ${error.message}`,
                error
            );
            
            this.#errorHandler?.handleError(searchError, { query, options });
            return []; // Return empty array on search failure
        }
    }

    /**
     * Gets conversation history
     * @param {string} [conversationId] - Conversation ID (current if not specified)
     * @returns {Array} Message history
     */
    getConversationHistory(conversationId) {
        const convId = conversationId || this.#state.activeConversation;
        const conversation = this.#state.conversations.get(convId);
        return conversation ? [...conversation.messages] : [];
    }

    /**
     * Gets current state
     * @returns {Object} Current state
     */
    getState() {
        return {
            isConnected: this.#state.isConnected,
            currentProvider: this.#state.currentProvider ? { ...this.#state.currentProvider } : null,
            activeConversation: this.#state.activeConversation,
            conversationCount: this.#state.conversations.size,
            messageCount: this.#state.messages.length,
            isTyping: this.#state.isTyping,
            isLoading: this.#state.isLoading
        };
    }

    /**
     * Destroys the component and cleans up resources
     * @returns {Promise<void>}
     */
    async destroy() {
        this.#removeEventListeners();
        
        // Disconnect current provider
        if (this.#state.currentProvider) {
            await this.#disconnectProvider(this.#state.currentProvider);
        }

        if (this.#element) {
            this.#element.innerHTML = '';
            this.#element.classList.remove('onlyoffice-chat-interface');
        }

        this.#element = null;
        this.#eventSystem = null;
        this.#stateManager = null;
    }

    /**
     * Sets up container element
     * @param {HTMLElement|string} container - Container element or selector
     * @private
     */
    #setupContainer(container) {
        if (typeof container === 'string') {
            this.#element = document.querySelector(container);
        } else if (container instanceof HTMLElement) {
            this.#element = container;
        } else {
            throw new Error('Invalid container provided to ChatInterface');
        }

        if (!this.#element) {
            throw new Error('ChatInterface container not found');
        }

        this.#element.classList.add('onlyoffice-chat-interface');
    }

    /**
     * Creates the ChatInterface DOM structure
     * @private
     */
    #createChatInterfaceStructure() {
        const providerSection = this.#config.providers.length > 0 ? `
            <div class="chat-header">
                <div class="chat-provider-selector">
                    <select class="provider-select" id="provider-select">
                        <option value="">Select Provider...</option>
                    </select>
                    <div class="provider-status" id="provider-status">
                        <span class="status-indicator disconnected"></span>
                        <span class="status-text">Disconnected</span>
                    </div>
                </div>
                <div class="chat-actions">
                    <button class="chat-action-btn" id="new-conversation" title="New Conversation">➕</button>
                    <button class="chat-action-btn" id="clear-chat" title="Clear Chat">🗑️</button>
                    <button class="chat-action-btn" id="export-chat" title="Export Chat">💾</button>
                </div>
            </div>
        ` : '';

        const conversationsSection = this.#config.multiConversation ? `
            <div class="chat-conversations">
                <div class="conversations-header">Conversations</div>
                <div class="conversations-list" id="conversations-list"></div>
            </div>
        ` : '';

        this.#element.innerHTML = `
            <div class="chat-container">
                ${providerSection}
                ${conversationsSection}
                <div class="chat-main">
                    <div class="chat-messages" id="chat-messages">
                        <div class="welcome-message">
                            <div class="welcome-icon">💬</div>
                            <div class="welcome-text">Start a conversation</div>
                        </div>
                    </div>
                    <div class="chat-input-container">
                        <div class="chat-input-wrapper">
                            <textarea 
                                id="chat-input" 
                                class="chat-input" 
                                placeholder="Type your message..." 
                                rows="1"
                            ></textarea>
                            <button class="chat-send-btn" id="chat-send-btn" disabled>
                                <span class="send-icon">📤</span>
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        `;

        // Get references to key elements
        this.#chatMessages = this.#element.querySelector('#chat-messages');
        this.#chatInput = this.#element.querySelector('#chat-input');
        this.#sendButton = this.#element.querySelector('#chat-send-btn');
        this.#providerSelect = this.#element.querySelector('#provider-select');
        this.#providerStatus = this.#element.querySelector('#provider-status');
        this.#conversationsList = this.#element.querySelector('#conversations-list');

        // Initialize providers
        if (this.#state.providers.length > 0) {
            this.#renderProviders();
        }
    }

    /**
     * Sets up event listeners
     * @private
     */
    #setupEventListeners() {
        // Chat input events
        if (this.#chatInput) {
            this.#chatInput.addEventListener('input', this.#handleInputChange.bind(this));
            this.#chatInput.addEventListener('keydown', this.#handleInputKeyDown.bind(this));
        }

        // Send button
        if (this.#sendButton) {
            this.#sendButton.addEventListener('click', this.#handleSendMessage.bind(this));
        }

        // Provider selection
        if (this.#providerSelect) {
            this.#providerSelect.addEventListener('change', this.#handleProviderChange.bind(this));
        }

        // Action buttons
        const newConvBtn = this.#element.querySelector('#new-conversation');
        if (newConvBtn) {
            newConvBtn.addEventListener('click', () => this.createConversation());
        }

        const clearBtn = this.#element.querySelector('#clear-chat');
        if (clearBtn) {
            clearBtn.addEventListener('click', () => this.clearConversation());
        }

        const exportBtn = this.#element.querySelector('#export-chat');
        if (exportBtn) {
            exportBtn.addEventListener('click', this.#handleExportChat.bind(this));
        }
    }

    /**
     * Removes event listeners
     * @private
     */
    #removeEventListeners() {
        // Event listeners are removed when element is destroyed
    }

    /**
     * Renders available providers
     * @private
     */
    #renderProviders() {
        if (!this.#providerSelect) return;

        const options = this.#state.providers.map(provider => 
            `<option value="${provider.id}" ${!provider.enabled ? 'disabled' : ''}>
                ${provider.icon} ${provider.name}
            </option>`
        ).join('');

        this.#providerSelect.innerHTML = `
            <option value="">Select Provider...</option>
            ${options}
        `;
    }

    /**
     * Renders conversations list
     * @private
     */
    #renderConversations() {
        if (!this.#conversationsList) return;

        const conversations = Array.from(this.#state.conversations.values())
            .sort((a, b) => b.updatedAt - a.updatedAt);

        const conversationsHtml = conversations.map(conv => `
            <div class="conversation-item ${conv.id === this.#state.activeConversation ? 'active' : ''}" 
                 data-conversation-id="${conv.id}">
                <div class="conversation-title">${conv.title}</div>
                <div class="conversation-info">
                    <span class="message-count">${conv.messages.length} messages</span>
                    <span class="conversation-time">${this.#formatTime(conv.updatedAt)}</span>
                </div>
            </div>
        `).join('');

        this.#conversationsList.innerHTML = conversationsHtml;

        // Add click handlers
        this.#conversationsList.querySelectorAll('.conversation-item').forEach(item => {
            item.addEventListener('click', () => {
                const convId = item.dataset.conversationId;
                this.setActiveConversation(convId);
            });
        });
    }

    /**
     * Renders messages
     * @private
     */
    #renderMessages() {
        if (!this.#chatMessages) return;

        if (this.#state.messages.length === 0) {
            this.#chatMessages.innerHTML = `
                <div class="welcome-message">
                    <div class="welcome-icon">💬</div>
                    <div class="welcome-text">Start a conversation</div>
                </div>
            `;
            return;
        }

        const messagesHtml = this.#state.messages.map(message => 
            this.#createMessageHTML(message)
        ).join('');

        this.#chatMessages.innerHTML = messagesHtml;

        if (this.#autoScroll) {
            this.#scrollToBottom();
        }
    }

    /**
     * Creates HTML for a single message
     * @param {Object} message - Message object
     * @returns {string} Message HTML
     * @private
     */
    #createMessageHTML(message) {
        const isUser = message.type === 'user';
        const statusIcon = this.#getStatusIcon(message.status);
        const avatar = this.#config.avatars ? 
            `<div class="message-avatar">${isUser ? '👤' : (this.#state.currentProvider?.icon || '🤖')}</div>` : '';
        
        const timestamp = this.#config.timestamps ? 
            `<div class="message-timestamp">${this.#formatTime(message.timestamp)}</div>` : '';

        const content = this.#config.enableMarkdown ? 
            this.#renderMarkdown(message.content) : 
            this.#escapeHTML(message.content);

        return `
            <div class="chat-message ${isUser ? 'user' : 'assistant'} ${message.status || ''}" 
                 data-message-id="${message.id}">
                ${avatar}
                <div class="message-content">
                    <div class="message-header">
                        <span class="message-sender">${message.sender}</span>
                        ${timestamp}
                        <span class="message-status">${statusIcon}</span>
                    </div>
                    <div class="message-text">${content}</div>
                    ${message.metadata && Object.keys(message.metadata).length > 0 ? 
                        `<div class="message-metadata">${JSON.stringify(message.metadata)}</div>` : ''}
                </div>
            </div>
        `;
    }

    /**
     * Renders typing indicator
     * @param {string} sender - Sender name
     * @private
     */
    #renderTypingIndicator(sender) {
        if (!this.#chatMessages) return;

        const existingIndicator = this.#chatMessages.querySelector('.typing-indicator');
        if (existingIndicator) {
            existingIndicator.remove();
        }

        const indicatorHTML = `
            <div class="typing-indicator">
                <div class="message-avatar">🤖</div>
                <div class="message-content">
                    <div class="message-header">
                        <span class="message-sender">${sender || 'Assistant'}</span>
                    </div>
                    <div class="typing-animation">
                        <span></span>
                        <span></span>
                        <span></span>
                    </div>
                </div>
            </div>
        `;

        this.#chatMessages.insertAdjacentHTML('beforeend', indicatorHTML);
        this.#scrollToBottom();
    }

    /**
     * Removes typing indicator
     * @private
     */
    #removeTypingIndicator() {
        const indicator = this.#chatMessages?.querySelector('.typing-indicator');
        if (indicator) {
            indicator.remove();
        }
    }

    /**
     * Event handlers
     * @private
     */
    #handleInputChange(event) {
        const value = event.target.value.trim();
        this.#sendButton.disabled = value.length === 0;
    }

    #handleInputKeyDown(event) {
        if (event.key === 'Enter' && !event.shiftKey) {
            event.preventDefault();
            this.#handleSendMessage();
        }
    }

    async #handleSendMessage() {
        const message = this.#chatInput.value.trim();
        if (!message) return;

        try {
            this.#chatInput.value = '';
            this.#sendButton.disabled = true;
            
            await this.sendMessage(message, 'user');
        } catch (error) {
            console.error('Failed to send message:', error);
        }
    }

    async #handleProviderChange(event) {
        const providerId = event.target.value;
        if (providerId) {
            await this.setProvider(providerId);
        }
    }

    #handleExportChat() {
        try {
            const exported = this.exportConversation('json');
            const blob = new Blob([exported], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            
            const a = document.createElement('a');
            a.href = url;
            a.download = `chat-${this.#state.activeConversation || 'export'}-${Date.now()}.json`;
            a.click();
            
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Failed to export chat:', error);
        }
    }

    /**
     * Utility methods
     * @private
     */
    #addMessageToConversation(message) {
        this.#state.messages.push(message);
        
        // Update conversation
        if (this.#state.activeConversation) {
            const conversation = this.#state.conversations.get(this.#state.activeConversation);
            if (conversation) {
                conversation.messages.push(message);
                conversation.updatedAt = new Date();
                
                // Limit messages
                if (conversation.messages.length > this.#config.maxMessages) {
                    conversation.messages = conversation.messages.slice(-this.#config.maxMessages);
                }
            }
        }

        this.#renderMessages();
        this.#renderConversations();
    }

    #updateMessage(messageId, updates) {
        const messageIndex = this.#state.messages.findIndex(m => m.id === messageId);
        if (messageIndex > -1) {
            Object.assign(this.#state.messages[messageIndex], updates);
            this.#renderMessages();
        }
    }

    async #connectProvider(provider) {
        // Simulate connection
        provider.status = 'connecting';
        this.#updateProviderDisplay();
        
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        provider.status = 'connected';
        this.#state.isConnected = true;
        this.#updateProviderDisplay();
    }

    async #disconnectProvider(provider) {
        provider.status = 'disconnected';
        this.#state.isConnected = false;
        this.#updateProviderDisplay();
    }

    #updateProviderDisplay() {
        if (!this.#providerStatus) return;

        const provider = this.#state.currentProvider;
        if (provider) {
            const statusIndicator = this.#providerStatus.querySelector('.status-indicator');
            const statusText = this.#providerStatus.querySelector('.status-text');
            
            statusIndicator.className = `status-indicator ${provider.status}`;
            statusText.textContent = `${provider.name} - ${provider.status}`;
            
            this.#providerSelect.value = provider.id;
        }
    }

    #updateConversationDisplay() {
        if (this.#conversationsList) {
            this.#renderConversations();
        }
    }

    #scrollToBottom() {
        if (this.#chatMessages) {
            this.#chatMessages.scrollTop = this.#chatMessages.scrollHeight;
        }
    }

    #generateId() {
        return `msg_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }

    #formatTime(date) {
        return new Date(date).toLocaleTimeString([], { 
            hour: '2-digit', 
            minute: '2-digit' 
        });
    }

    #getStatusIcon(status) {
        switch (status) {
            case 'sending': return '⏳';
            case 'sent': return '✓';
            case 'error': return '❌';
            default: return '';
        }
    }

    #renderMarkdown(text) {
        // SECURITY FIX: Sanitize input before markdown processing
        if (!text || typeof text !== 'string') {
            return '';
        }
        
        // First escape HTML to prevent XSS
        const escaped = this.#escapeHTML(text);
        
        // Then apply safe markdown rendering
        return escaped
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\*(.*?)\*/g, '<em>$1</em>')
            .replace(/`(.*?)`/g, '<code>$1</code>')
            .replace(/\n/g, '<br>');
    }

    #escapeHTML(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    #exportAsText(conversation) {
        const header = `Conversation: ${conversation.title}\nCreated: ${conversation.createdAt}\nMessages: ${conversation.messages.length}\n\n`;
        const messages = conversation.messages.map(msg => 
            `[${this.#formatTime(msg.timestamp)}] ${msg.sender}: ${msg.content}`
        ).join('\n');
        
        return header + messages;
    }

    #exportAsHTML(conversation) {
        const messagesHtml = conversation.messages.map(msg => 
            this.#createMessageHTML(msg)
        ).join('');
        
        return `
            <!DOCTYPE html>
            <html>
            <head>
                <title>${conversation.title}</title>
                <style>
                    /* Chat styles would go here */
                </style>
            </head>
            <body>
                <h1>${conversation.title}</h1>
                <div class="chat-messages">${messagesHtml}</div>
            </body>
            </html>
        `;
    }

    /**
     * Validates constructor options
     * @param {Object} options - Constructor options
     * @throws {ValidationError} If validation fails
     * @private
     */
    #validateConstructorOptions(options) {
        if (!options || typeof options !== 'object') {
            throw new ValidationError('options', options, { type: 'object' });
        }
        
        const schema = {
            container: { required: true },
            maxMessages: { type: 'number', min: 10, max: 10000 },
            maxConversations: { type: 'number', min: 1, max: 100 },
            multiConversation: { type: 'boolean' },
            enableMarkdown: { type: 'boolean' },
            showTypingIndicator: { type: 'boolean' },
            providers: { type: 'object' }
        };
        
        this.#errorHandler?.validateConfig(options, schema, 'ChatInterface');
    }
    
    /**
     * Validates message parameters
     * @param {string|Object} message - Message content or object
     * @param {string} type - Message type
     * @param {Object} metadata - Message metadata
     * @throws {ValidationError} If validation fails
     * @private
     */
    #validateMessageParameters(message, type, metadata) {
        if (!message) {
            throw new ValidationError('message', message, { required: true });
        }
        
        if (typeof message === 'string') {
            if (message.trim().length === 0) {
                throw new ValidationError('message', message, { minLength: 1 });
            }
            if (message.length > 10000) {
                throw new ValidationError('message', message, { maxLength: 10000 });
            }
        } else if (typeof message === 'object') {
            if (!message.content || typeof message.content !== 'string') {
                throw new ValidationError('message.content', message.content, { type: 'string', required: true });
            }
        } else {
            throw new ValidationError('message', message, { type: 'string or object' });
        }
        
        const validTypes = ['user', 'assistant', 'system'];
        if (!validTypes.includes(type)) {
            throw new ValidationError('type', type, { enum: validTypes });
        }
        
        if (metadata && typeof metadata !== 'object') {
            throw new ValidationError('metadata', metadata, { type: 'object' });
        }
    }
    
    /**
     * Validates search parameters
     * @param {string} query - Search query
     * @param {Object} options - Search options
     * @throws {ValidationError} If validation fails
     * @private
     */
    #validateSearchParameters(query, options) {
        if (typeof query !== 'string') {
            throw new ValidationError('query', query, { type: 'string' });
        }
        
        if (query.trim().length === 0) {
            throw new ValidationError('query', query, { minLength: 1 });
        }
        
        if (query.length > 1000) {
            throw new ValidationError('query', query, { maxLength: 1000 });
        }
        
        if (options && typeof options !== 'object') {
            throw new ValidationError('options', options, { type: 'object' });
        }
    }
    
    /**
     * Sanitizes message content to prevent injection attacks
     * @param {string} content - Message content
     * @returns {string} Sanitized content
     * @private
     */
    #sanitizeMessageContent(content) {
        if (typeof content !== 'string') {
            return '';
        }
        
        // Remove potential script tags and dangerous HTML
        return content
            .replace(/<script[^>]*>.*?<\/script>/gi, '')
            .replace(/<iframe[^>]*>.*?<\/iframe>/gi, '')
            .replace(/javascript:/gi, '')
            .replace(/on\w+\s*=/gi, '')
            .trim();
    }
    
    /**
     * Sanitizes search query to prevent injection
     * @param {string} query - Search query
     * @returns {string} Sanitized query
     * @private
     */
    #sanitizeSearchQuery(query) {
        if (typeof query !== 'string') {
            return '';
        }
        
        // Remove potentially dangerous characters for search
        return query
            .replace(/[<>"'&]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    /**
     * Binds component events
     * @private
     */
    #bindEvents() {
        // No additional event binding needed here
    }
}

// Export ChatInterface component
if (typeof window !== 'undefined') {
    window.ChatInterface = ChatInterface;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = ChatInterface;
}

export default ChatInterface;