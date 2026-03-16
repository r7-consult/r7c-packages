/**
 * Chess Collaboration Manager
 * Following OnlyOffice Plugin Development Standards
 * 
 * Based on: _coding_standard/02_api_reference_patterns.md#onlyoffice-specific-patterns
 */

class ChessCollaborationManager {
    constructor() {
        this.isInitialized = false;
        this.eventListeners = new Map();
        this.sessionId = null;
        this.currentUser = null;
        this.connectedUsers = new Map();
        this.isOnlyOfficeEnabled = false;
        this.heartbeatInterval = null;
    }

    /**
     * Initialize collaboration manager
     */
    async initialize() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'CollaborationManager initializing');

        try {
            // Check if OnlyOffice collaboration is available
            this.checkOnlyOfficeCollaboration();
            
            // Setup user identification
            await this.setupUserIdentification();
            
            // Initialize collaboration features if available
            if (this.isOnlyOfficeEnabled) {
                await this.initializeOnlyOfficeCollaboration();
            } else {
                await this.initializeFallbackCollaboration();
            }
            
            this.isInitialized = true;
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                'CollaborationManager initialized', { 
                    onlyOfficeEnabled: this.isOnlyOfficeEnabled,
                    userId: this.currentUser?.id 
                });

        } catch (error) {
            throw new window.ChessErrors.ChessCollaborationError(
                'Collaboration manager initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Check if OnlyOffice collaboration features are available
     */
    checkOnlyOfficeCollaboration() {
        this.isOnlyOfficeEnabled = !!(
            window.Asc && 
            window.Asc.plugin && 
            window.Asc.plugin.info && 
            window.Asc.plugin.info.isCoAuthoringEnable
        );

        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'OnlyOffice collaboration check', { 
                available: this.isOnlyOfficeEnabled,
                hasAsc: !!window.Asc,
                hasPlugin: !!(window.Asc && window.Asc.plugin),
                hasInfo: !!(window.Asc && window.Asc.plugin && window.Asc.plugin.info),
                isCoAuthoringEnabled: !!(window.Asc && window.Asc.plugin && window.Asc.plugin.info && window.Asc.plugin.info.isCoAuthoringEnable)
            });
    }

    /**
     * Setup user identification
     */
    async setupUserIdentification() {
        try {
            // Try to get user info from OnlyOffice
            if (window.Asc && window.Asc.plugin && window.Asc.plugin.info) {
                const pluginInfo = window.Asc.plugin.info;
                this.currentUser = {
                    id: pluginInfo.userId || this.generateUserId(),
                    name: pluginInfo.userName || 'Anonymous',
                    color: null, // Will be assigned when joining game
                    isHost: pluginInfo.isHost || false
                };
            } else {
                // Fallback user identification
                this.currentUser = {
                    id: this.generateUserId(),
                    name: 'Anonymous',
                    color: null,
                    isHost: true // Default to host in standalone mode
                };
            }

            window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                'User identification setup', this.currentUser);

        } catch (error) {
            throw new window.ChessErrors.ChessCollaborationError(
                'User identification setup failed',
                { originalError: error }
            );
        }
    }

    /**
     * Initialize OnlyOffice-based collaboration
     */
    async initializeOnlyOfficeCollaboration() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Initializing OnlyOffice collaboration');

        try {
            // Setup OnlyOffice collaboration handlers
            this.setupOnlyOfficeHandlers();
            
            // Generate session ID based on document
            this.sessionId = this.generateSessionId();
            
            // Start heartbeat to maintain connection
            this.startHeartbeat();
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                'OnlyOffice collaboration initialized', { sessionId: this.sessionId });

        } catch (error) {
            throw new window.ChessErrors.ChessCollaborationError(
                'OnlyOffice collaboration initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Initialize fallback collaboration (local storage based)
     */
    async initializeFallbackCollaboration() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Initializing fallback collaboration');

        try {
            // Use localStorage for basic collaboration simulation
            this.sessionId = 'local_' + Date.now();
            
            // Setup storage-based communication
            this.setupStorageHandlers();
            
            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                'Fallback collaboration initialized', { sessionId: this.sessionId });

        } catch (error) {
            throw new window.ChessErrors.ChessCollaborationError(
                'Fallback collaboration initialization failed',
                { originalError: error }
            );
        }
    }

    /**
     * Setup OnlyOffice collaboration handlers
     */
    setupOnlyOfficeHandlers() {
        // This would integrate with OnlyOffice's collaboration API
        // For now, we'll use a simplified approach
        
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'OnlyOffice collaboration handlers setup complete');
    }

    /**
     * Setup localStorage-based handlers for fallback
     */
    setupStorageHandlers() {
        // Listen for storage changes (simple cross-tab communication)
        window.addEventListener('storage', (event) => {
            this.handleStorageChange(event);
        });

        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Storage-based collaboration handlers setup complete');
    }

    /**
     * Join collaboration session
     */
    async joinSession() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Joining collaboration session', { sessionId: this.sessionId });

        try {
            // Assign player color if not already assigned
            if (!this.currentUser.color) {
                this.currentUser.color = await this.assignPlayerColor();
            }

            // Add user to connected users
            this.connectedUsers.set(this.currentUser.id, this.currentUser);

            // Broadcast join event
            await this.broadcastEvent(window.ChessConstants.COLLAB_EVENTS.PLAYER_JOINED, {
                user: this.currentUser,
                timestamp: Date.now()
            });

            // Notify listeners
            this.notifyPlayerJoined(this.currentUser);

            window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                'Successfully joined collaboration session', { 
                    userId: this.currentUser.id, 
                    color: this.currentUser.color 
                });

        } catch (error) {
            throw new window.ChessErrors.ChessCollaborationError(
                'Failed to join collaboration session',
                { sessionId: this.sessionId, originalError: error }
            );
        }
    }

    /**
     * Assign player color
     */
    async assignPlayerColor() {
        // Simple color assignment logic
        const connectedColors = Array.from(this.connectedUsers.values())
            .map(user => user.color)
            .filter(color => color);

        if (!connectedColors.includes('white')) {
            return 'white';
        } else if (!connectedColors.includes('black')) {
            return 'black';
        } else {
            return 'spectator';
        }
    }

    /**
     * Send move to other players
     */
    async sendMove(move) {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Sending move to collaborators', move);

        try {
            const moveData = {
                move,
                playerId: this.currentUser.id,
                playerColor: this.currentUser.color,
                timestamp: Date.now(),
                sessionId: this.sessionId
            };

            await this.broadcastEvent(window.ChessConstants.COLLAB_EVENTS.MOVE_MADE, moveData);

        } catch (error) {
            throw new window.ChessErrors.ChessCollaborationError(
                'Failed to send move',
                { move, originalError: error }
            );
        }
    }

    /**
     * Broadcast event to other users
     */
    async broadcastEvent(eventType, data) {
        if (this.isOnlyOfficeEnabled) {
            await this.broadcastViaOnlyOffice(eventType, data);
        } else {
            await this.broadcastViaStorage(eventType, data);
        }
    }

    /**
     * Broadcast via OnlyOffice collaboration
     */
    async broadcastViaOnlyOffice(eventType, data) {
        // This would use OnlyOffice's collaboration API
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Broadcasting via OnlyOffice', { eventType, data });
        
        // Placeholder - would implement actual OnlyOffice collaboration
    }

    /**
     * Broadcast via localStorage (fallback)
     */
    async broadcastViaStorage(eventType, data) {
        const collaborationData = {
            eventType,
            data,
            sessionId: this.sessionId,
            senderId: this.currentUser.id,
            timestamp: Date.now()
        };

        const storageKey = `${window.ChessConstants.STORAGE.COLLABORATION_DATA}_${this.sessionId}`;
        localStorage.setItem(storageKey, JSON.stringify(collaborationData));

        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Broadcasted via storage', { eventType, storageKey });
    }

    /**
     * Handle storage change events
     */
    handleStorageChange(event) {
        if (!event.key || !event.key.includes(window.ChessConstants.STORAGE.COLLABORATION_DATA)) {
            return;
        }

        try {
            const collaborationData = JSON.parse(event.newValue);
            
            // Ignore our own events
            if (collaborationData.senderId === this.currentUser.id) {
                return;
            }

            // Process the collaboration event
            this.processCollaborationEvent(collaborationData);

        } catch (error) {
            window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                'Failed to process storage change', error);
        }
    }

    /**
     * Process collaboration event
     */
    processCollaborationEvent(collaborationData) {
        const { eventType, data } = collaborationData;

        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Processing collaboration event', { eventType, data });

        switch (eventType) {
            case window.ChessConstants.COLLAB_EVENTS.PLAYER_JOINED:
                this.handlePlayerJoined(data);
                break;
            case window.ChessConstants.COLLAB_EVENTS.PLAYER_LEFT:
                this.handlePlayerLeft(data);
                break;
            case window.ChessConstants.COLLAB_EVENTS.MOVE_MADE:
                this.handleMoveReceived(data);
                break;
            case window.ChessConstants.COLLAB_EVENTS.GAME_STATE_CHANGED:
                this.handleGameStateChanged(data);
                break;
            default:
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                    'Unknown collaboration event', { eventType });
        }
    }

    /**
     * Handle player joined event
     */
    handlePlayerJoined(data) {
        const { user } = data;
        this.connectedUsers.set(user.id, user);
        this.notifyPlayerJoined(user);
    }

    /**
     * Handle player left event
     */
    handlePlayerLeft(data) {
        const { user } = data;
        this.connectedUsers.delete(user.id);
        this.notifyPlayerLeft(user);
    }

    /**
     * Handle move received event
     */
    handleMoveReceived(data) {
        this.notifyMoveReceived(data);
    }

    /**
     * Handle game state change event
     */
    handleGameStateChanged(data) {
        // Process game state changes from other players
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Game state changed by collaborator', data);
    }

    /**
     * Reconnect to collaboration session
     */
    async reconnect() {
        window.ChessDebug?.info(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'Attempting to reconnect to collaboration session');

        try {
            // Re-initialize collaboration
            await this.initialize();
            
            // Rejoin session
            await this.joinSession();

        } catch (error) {
            throw new window.ChessErrors.ChessCollaborationError(
                'Reconnection failed',
                { originalError: error }
            );
        }
    }

    /**
     * Start heartbeat to maintain connection
     */
    startHeartbeat() {
        if (this.heartbeatInterval) {
            clearInterval(this.heartbeatInterval);
        }

        this.heartbeatInterval = setInterval(async () => {
            try {
                await this.sendHeartbeat();
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                    'Heartbeat failed', error);
            }
        }, window.ChessConstants.NETWORK.HEARTBEAT_INTERVAL_MS);
    }

    /**
     * Send heartbeat
     */
    async sendHeartbeat() {
        await this.broadcastEvent('heartbeat', {
            userId: this.currentUser.id,
            timestamp: Date.now()
        });
    }

    /**
     * Generate unique user ID
     */
    generateUserId() {
        return 'user_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    /**
     * Generate session ID
     */
    generateSessionId() {
        if (this.isOnlyOfficeEnabled && window.Asc && window.Asc.plugin && window.Asc.plugin.info) {
            // Use document-based session ID for OnlyOffice
            return 'chess_' + (window.Asc.plugin.info.documentId || Date.now());
        } else {
            return 'chess_local_' + Date.now();
        }
    }

    /**
     * Event listener management
     */
    addEventListener(event, callback) {
        if (!this.eventListeners.has(event)) {
            this.eventListeners.set(event, []);
        }
        this.eventListeners.get(event).push(callback);
    }

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
     * Notify player joined
     */
    notifyPlayerJoined(user) {
        const listeners = this.eventListeners.get('playerJoined') || [];
        listeners.forEach(callback => {
            try {
                callback(user);
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                    'Player joined listener error', error);
            }
        });
    }

    /**
     * Notify player left
     */
    notifyPlayerLeft(user) {
        const listeners = this.eventListeners.get('playerLeft') || [];
        listeners.forEach(callback => {
            try {
                callback(user);
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                    'Player left listener error', error);
            }
        });
    }

    /**
     * Notify move received
     */
    notifyMoveReceived(moveData) {
        const listeners = this.eventListeners.get('moveReceived') || [];
        listeners.forEach(callback => {
            try {
                callback(moveData);
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                    'Move received listener error', error);
            }
        });
    }

    /**
     * Get connected users
     */
    getConnectedUsers() {
        return Array.from(this.connectedUsers.values());
    }

    /**
     * Get current user
     */
    getCurrentUser() {
        return this.currentUser;
    }

    /**
     * Cleanup collaboration manager
     */
    async cleanup() {
        window.ChessDebug?.debug(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
            'CollaborationManager cleanup');

        // Stop heartbeat
        if (this.heartbeatInterval) {
            clearInterval(this.heartbeatInterval);
            this.heartbeatInterval = null;
        }

        // Broadcast leave event
        if (this.currentUser) {
            try {
                await this.broadcastEvent(window.ChessConstants.COLLAB_EVENTS.PLAYER_LEFT, {
                    user: this.currentUser,
                    timestamp: Date.now()
                });
            } catch (error) {
                window.ChessDebug?.warn(window.ChessConstants.DEBUG_CATEGORIES.COLLABORATION, 
                    'Failed to broadcast leave event', error);
            }
        }

        // Clear event listeners
        this.eventListeners.clear();

        // Reset state
        this.connectedUsers.clear();
        this.currentUser = null;
        this.sessionId = null;
        this.isInitialized = false;
    }
}

// Export collaboration manager
window.ChessCollaborationManager = ChessCollaborationManager;