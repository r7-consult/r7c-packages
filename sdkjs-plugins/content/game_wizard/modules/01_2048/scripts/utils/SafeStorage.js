/**
 * Safe Storage Utility
 * Provides safe localStorage access with fallback to memory storage
 * Handles disabled localStorage, quota exceeded, and other errors
 */

class SafeStorage {
    constructor(prefix = 'game_') {
        this.prefix = prefix;
        this.memoryStorage = new Map();
        this.isLocalStorageAvailable = this.checkLocalStorageAvailable();
        
        if (!this.isLocalStorageAvailable) {
            window.debug?.warn('SafeStorage', 'localStorage not available, using memory storage');
        }
    }
    
    /**
     * Check if localStorage is available and working
     */
    checkLocalStorageAvailable() {
        try {
            const testKey = '__storage_test__';
            localStorage.setItem(testKey, 'test');
            localStorage.removeItem(testKey);
            return true;
        } catch (e) {
            return false;
        }
    }
    
    /**
     * Set item in storage with error handling
     */
    setItem(key, value) {
        const fullKey = this.prefix + key;
        const stringValue = typeof value === 'string' ? value : JSON.stringify(value);
        
        // Always update memory storage
        this.memoryStorage.set(fullKey, stringValue);
        
        if (this.isLocalStorageAvailable) {
            try {
                localStorage.setItem(fullKey, stringValue);
                return true;
            } catch (error) {
                // Handle quota exceeded or other errors
                if (error.name === 'QuotaExceededError') {
                    window.debug?.warn('SafeStorage', 'localStorage quota exceeded', { key });
                    
                    // Try to clear old data and retry
                    this.clearOldData();
                    
                    try {
                        localStorage.setItem(fullKey, stringValue);
                        return true;
                    } catch (retryError) {
                        window.debug?.error('SafeStorage', 'Failed to store after cleanup', { key, error: retryError });
                    }
                } else {
                    window.debug?.error('SafeStorage', 'Failed to set item', { key, error });
                }
            }
        }
        
        // Data is still in memory storage
        return false;
    }
    
    /**
     * Get item from storage with fallback
     */
    getItem(key) {
        const fullKey = this.prefix + key;
        
        if (this.isLocalStorageAvailable) {
            try {
                const value = localStorage.getItem(fullKey);
                if (value !== null) {
                    return value;
                }
            } catch (error) {
                window.debug?.warn('SafeStorage', 'Failed to get item from localStorage', { key, error });
            }
        }
        
        // Fallback to memory storage
        return this.memoryStorage.get(fullKey) || null;
    }
    
    /**
     * Remove item from storage
     */
    removeItem(key) {
        const fullKey = this.prefix + key;
        
        this.memoryStorage.delete(fullKey);
        
        if (this.isLocalStorageAvailable) {
            try {
                localStorage.removeItem(fullKey);
            } catch (error) {
                window.debug?.warn('SafeStorage', 'Failed to remove item from localStorage', { key, error });
            }
        }
    }
    
    /**
     * Clear all items with this prefix
     */
    clear() {
        // Clear memory storage
        for (const key of this.memoryStorage.keys()) {
            if (key.startsWith(this.prefix)) {
                this.memoryStorage.delete(key);
            }
        }
        
        // Clear localStorage
        if (this.isLocalStorageAvailable) {
            try {
                const keysToRemove = [];
                for (let i = 0; i < localStorage.length; i++) {
                    const key = localStorage.key(i);
                    if (key && key.startsWith(this.prefix)) {
                        keysToRemove.push(key);
                    }
                }
                
                keysToRemove.forEach(key => {
                    try {
                        localStorage.removeItem(key);
                    } catch (error) {
                        window.debug?.warn('SafeStorage', 'Failed to remove key', { key, error });
                    }
                });
            } catch (error) {
                window.debug?.error('SafeStorage', 'Failed to clear localStorage', { error });
            }
        }
    }
    
    /**
     * Clear old data to free up space
     */
    clearOldData() {
        if (!this.isLocalStorageAvailable) return;
        
        try {
            const now = Date.now();
            const maxAge = 7 * 24 * 60 * 60 * 1000; // 7 days
            
            const keysToRemove = [];
            
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                if (!key || !key.startsWith(this.prefix)) continue;
                
                try {
                    const value = localStorage.getItem(key);
                    if (value) {
                        const data = JSON.parse(value);
                        if (data.timestamp && (now - data.timestamp) > maxAge) {
                            keysToRemove.push(key);
                        }
                    }
                } catch (parseError) {
                    // If we can't parse it, it might be old format - remove it
                    keysToRemove.push(key);
                }
            }
            
            keysToRemove.forEach(key => {
                try {
                    localStorage.removeItem(key);
                } catch (error) {
                    window.debug?.warn('SafeStorage', 'Failed to remove old key', { key, error });
                }
            });
            
            window.debug?.info('SafeStorage', `Cleared ${keysToRemove.length} old items`);
            
        } catch (error) {
            window.debug?.error('SafeStorage', 'Failed to clear old data', { error });
        }
    }
    
    /**
     * Get storage size info
     */
    getStorageInfo() {
        const info = {
            memoryItemCount: this.memoryStorage.size,
            localStorageAvailable: this.isLocalStorageAvailable,
            localStorageItemCount: 0,
            estimatedSize: 0
        };
        
        if (this.isLocalStorageAvailable) {
            try {
                let totalSize = 0;
                let itemCount = 0;
                
                for (let i = 0; i < localStorage.length; i++) {
                    const key = localStorage.key(i);
                    if (key && key.startsWith(this.prefix)) {
                        itemCount++;
                        const value = localStorage.getItem(key);
                        if (value) {
                            totalSize += key.length + value.length;
                        }
                    }
                }
                
                info.localStorageItemCount = itemCount;
                info.estimatedSize = totalSize;
                
            } catch (error) {
                window.debug?.warn('SafeStorage', 'Failed to get storage info', { error });
            }
        }
        
        return info;
    }
}

// Create global instance for game storage
window.gameStorage = new SafeStorage('game2048_');