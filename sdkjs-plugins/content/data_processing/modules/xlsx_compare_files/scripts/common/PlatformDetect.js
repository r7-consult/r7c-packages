/**
 * ═══════════════════════════════════════════════════════════════════════════════════
 * 🖥️ PLATFORM DETECTION UTILITY
 * ═══════════════════════════════════════════════════════════════════════════════════
 * 
 * @fileOverview    Detect platform/OS for conditional UI behavior
 * @version         1.0.0
 * @since           2025-12-17
 * 
 * @description
 * Provides platform detection for adapting UI components that behave differently
 * across platforms (Windows, Linux, macOS).
 * 
 * Known platform-specific issues:
 * - Linux/Wayland: Native <select> popups may not appear in iframes
 * - Linux: AscDesktopEditor.OpenFilenameDialog may not work
 * 
 * ═══════════════════════════════════════════════════════════════════════════════════
 */

export const PlatformDetect = {
    
    /** @type {string|null} Cached platform value */
    _cachedPlatform: null,
    
    /**
     * Get current platform
     * @returns {'windows'|'linux'|'macos'|'unknown'}
     */
    getPlatform() {
        if (this._cachedPlatform) {
            return this._cachedPlatform;
        }
        
        const userAgent = navigator.userAgent.toLowerCase();
        const platform = navigator.platform?.toLowerCase() || '';
        
        if (platform.includes('win') || userAgent.includes('windows')) {
            this._cachedPlatform = 'windows';
        } else if (platform.includes('linux') || userAgent.includes('linux')) {
            this._cachedPlatform = 'linux';
        } else if (platform.includes('mac') || userAgent.includes('macintosh')) {
            this._cachedPlatform = 'macos';
        } else {
            this._cachedPlatform = 'unknown';
        }
        
        return this._cachedPlatform;
    },
    
    /**
     * Check if running on Windows
     * @returns {boolean}
     */
    isWindows() {
        return this.getPlatform() === 'windows';
    },
    
    /**
     * Check if running on Linux
     * @returns {boolean}
     */
    isLinux() {
        return this.getPlatform() === 'linux';
    },
    
    /**
     * Check if running on macOS
     * @returns {boolean}
     */
    isMacOS() {
        return this.getPlatform() === 'macos';
    },
    
    /**
     * Check if native <select> popups are reliable
     * On Linux/Wayland they often fail inside iframes
     * @returns {boolean}
     */
    hasReliableNativeSelects() {
        // Native selects are unreliable on Linux
        return !this.isLinux();
    },

    /**
     * Check if this UI is running inside an iframe.
     * OnlyOffice/R7 plugin UIs are typically iframe-based, which can break native popups.
     * @returns {boolean}
     */
    isInIframe() {
        if (typeof window === 'undefined') {
            return false;
        }
        try {
            return window.self !== window.top;
        } catch (e) {
            // Cross-origin protection: assume iframe-like restrictions apply.
            return true;
        }
    },

    /**
     * Decide whether to prefer CustomSelectOverlay over native <select> popups.
     * @returns {boolean}
     */
    shouldUseCustomSelectOverlay() {
        // Even on Windows, native <select> popups can be unreliable inside the plugin host.
        return this.hasOnlyOfficePluginApi() || this.isInIframe() || !this.hasReliableNativeSelects();
    },

    /**
     * Check if ONLYOFFICE/R7 plugin API is available.
     * @returns {boolean}
     */
    hasOnlyOfficePluginApi() {
        if (typeof window === 'undefined') {
            return false;
        }
        return !!(window.Asc && window.Asc.plugin);
    },
    
    /**
     * Check if AscDesktopEditor file dialogs are available and reliable
     * @returns {boolean}
     */
    hasReliableFileDialogs() {
        // Check if API exists
        if (typeof AscDesktopEditor === 'undefined') {
            return false;
        }
        
        if (typeof AscDesktopEditor.OpenFilenameDialog !== 'function') {
            return false;
        }
        
        // On Linux, file dialogs may be unreliable
        // Return true for Windows/macOS, conditional for Linux
        return !this.isLinux();
    },
    
    /**
     * Check if we're running inside ONLYOFFICE Desktop Editor
     * @returns {boolean}
     */
    isDesktopEditor() {
        return typeof AscDesktopEditor !== 'undefined';
    },
    
    /**
     * Get platform info object
     * @returns {Object}
     */
    getInfo() {
        return {
            platform: this.getPlatform(),
            userAgent: navigator.userAgent,
            isDesktopEditor: this.isDesktopEditor(),
            isInIframe: this.isInIframe(),
            hasOnlyOfficePluginApi: this.hasOnlyOfficePluginApi(),
            hasReliableNativeSelects: this.hasReliableNativeSelects(),
            shouldUseCustomSelectOverlay: this.shouldUseCustomSelectOverlay(),
            hasReliableFileDialogs: this.hasReliableFileDialogs()
        };
    }
};

// Export default for convenience
export default PlatformDetect;
