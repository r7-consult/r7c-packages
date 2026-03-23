/**
 * ═══════════════════════════════════════════════════════════════════════════════════
 * 🎯 CUSTOM SELECT OVERLAY
 * ═══════════════════════════════════════════════════════════════════════════════════
 * 
 * @fileOverview    Custom dropdown overlay for <select> elements
 * @version         1.0.0
 * @since           2025-12-17
 * 
 * @description
 * Replaces native <select> popups with custom in-DOM dropdown overlays.
 * Solves issues with native selects on Linux/Wayland where popups fail to appear
 * or are clipped when contained by iframes/modals.
 * 
 * Based on ADR-023 pattern from Macros IDE.
 * 
 * Usage:
 *   import { CustomSelectOverlay } from './CustomSelectOverlay.js';
 *   
 *   // Initialize for specific selects
 *   CustomSelectOverlay.attach('#my-select');
 *   CustomSelectOverlay.attach(document.getElementById('another-select'));
 *   
 *   // Or attach to all selects in a container
 *   CustomSelectOverlay.attachAll('.my-modal');
 * 
 * ═══════════════════════════════════════════════════════════════════════════════════
 */

export class CustomSelectOverlay {
    
    /** @type {WeakSet<HTMLSelectElement>} Track attached selects to prevent double-binding */
    static #attachedSelects = new WeakSet();
    
    /** @type {HTMLElement|null} Currently open picker */
    static #currentPicker = null;
    
    /** @type {Function|null} Cleanup function for current picker */
    static #currentCleanup = null;
    
    /** @type {HTMLSelectElement|null} The select element for which picker is currently open */
    static #currentSelectEl = null;

    /** @type {Map<HTMLElement, { observer: MutationObserver, selectSelector: string }>} */
    static #autoAttachObservers = new Map();
    
    /**
     * Attach custom overlay to a select element
     * @param {string|HTMLSelectElement} selectOrSelector - Select element or CSS selector
     * @returns {boolean} - True if successfully attached
     */
    static attach(selectOrSelector) {
        const selectEl = typeof selectOrSelector === 'string' 
            ? document.querySelector(selectOrSelector)
            : selectOrSelector;
            
        if (!selectEl || selectEl.tagName !== 'SELECT') {
            console.warn('[CustomSelectOverlay] Invalid select element:', selectOrSelector);
            return false;
        }
        
        // Prevent double-binding
        if (this.#attachedSelects.has(selectEl)) {
            return true;
        }
        
        this.#attachedSelects.add(selectEl);
        
        // Block native popup
        selectEl.addEventListener('mousedown', (e) => {
            if (!selectEl.disabled) {
                e.preventDefault();
                this.#openPicker(selectEl, false);
            }
        });
        
        // Keyboard support
        selectEl.addEventListener('keydown', (e) => {
            if (selectEl.disabled) return;
            
            if (e.key === 'Enter' || e.key === ' ' || e.key === 'ArrowDown') {
                e.preventDefault();
                this.#openPicker(selectEl, true);
            }
        });
        
        return true;
    }
    
    /**
     * Attach custom overlay to all selects within a container
     * @param {string|HTMLElement} containerOrSelector - Container element or CSS selector
     * @param {string} [selectSelector='select'] - Optional selector to filter selects
     * @returns {number} - Number of selects attached
     */
    static attachAll(containerOrSelector, selectSelector = 'select') {
        const container = typeof containerOrSelector === 'string'
            ? document.querySelector(containerOrSelector)
            : containerOrSelector;
            
        if (!container) {
            console.warn('[CustomSelectOverlay] Container not found:', containerOrSelector);
            return 0;
        }
        
        const selects = container.querySelectorAll(selectSelector);
        let count = 0;
        
        selects.forEach(select => {
            if (this.attach(select)) {
                count++;
            }
        });
        
        return count;
    }

    /**
     * Enable automatic attachment for selects added dynamically to the DOM.
     * - Attaches to existing selects immediately.
     * - Observes the container for new selects and attaches when they appear.
     *
     * @param {string|HTMLElement} [containerOrSelector=document.body] - Container element or selector
     * @param {string} [selectSelector='select:not([multiple])'] - Select selector to bind (exclude multi-selects by default)
     * @returns {Function} cleanup function to disable the observer
     */
    static enableAutoAttach(containerOrSelector = document.body, selectSelector = 'select:not([multiple])') {
        const container = typeof containerOrSelector === 'string'
            ? document.querySelector(containerOrSelector)
            : containerOrSelector;

        if (!container) {
            console.warn('[CustomSelectOverlay] Container not found for auto-attach:', containerOrSelector);
            return () => {};
        }

        // Attach to existing selects first
        this.attachAll(container, selectSelector);

        const existing = this.#autoAttachObservers.get(container);
        if (existing && existing.selectSelector === selectSelector) {
            return () => {
                existing.observer.disconnect();
                this.#autoAttachObservers.delete(container);
            };
        }

        if (existing) {
            existing.observer.disconnect();
            this.#autoAttachObservers.delete(container);
        }

        const observer = new MutationObserver((mutations) => {
            for (const mutation of mutations) {
                for (const addedNode of mutation.addedNodes) {
                    if (!addedNode || addedNode.nodeType !== 1) continue; // ELEMENT_NODE
                    const element = /** @type {HTMLElement} */ (addedNode);

                    // If the node itself is a matching select, attach.
                    if (element.tagName === 'SELECT' && typeof element.matches === 'function' && element.matches(selectSelector)) {
                        this.attach(/** @type {HTMLSelectElement} */ (element));
                    }

                    // Attach to any matching selects inside the added subtree.
                    if (typeof element.querySelectorAll === 'function') {
                        element.querySelectorAll(selectSelector).forEach((sel) => {
                            if (sel && sel.tagName === 'SELECT') {
                                this.attach(/** @type {HTMLSelectElement} */ (sel));
                            }
                        });
                    }
                }
            }
        });

        observer.observe(container, { childList: true, subtree: true });
        this.#autoAttachObservers.set(container, { observer, selectSelector });

        return () => {
            observer.disconnect();
            this.#autoAttachObservers.delete(container);
        };
    }

    /**
     * Disable all auto-attach observers (useful for plugin teardown when iframe is hidden).
     * @returns {number} Number of observers disconnected
     */
    static disableAutoAttachAll() {
        let count = 0;
        try {
            for (const { observer } of this.#autoAttachObservers.values()) {
                try {
                    observer.disconnect();
                } catch (_) {}
                count++;
            }
        } finally {
            this.#autoAttachObservers.clear();
        }
        return count;
    }
    
    /**
     * Close any currently open picker
     */
    static closeCurrent() {
        if (this.#currentCleanup) {
            this.#currentCleanup();
            this.#currentCleanup = null;
        }
        if (this.#currentPicker) {
            this.#currentPicker.remove();
            this.#currentPicker = null;
        }
        this.#currentSelectEl = null;
    }
    
    /**
     * Open custom picker for a select
     * @param {HTMLSelectElement} selectEl 
     * @param {boolean} byKeyboard 
     */
    static #openPicker(selectEl, byKeyboard) {
        // If picker is already open for this select, close it (toggle behavior)
        if (this.#currentPicker && this.#currentSelectEl === selectEl) {
            this.closeCurrent();
            return;
        }
        
        // Close any existing picker for a different select
        this.closeCurrent();
        
        // Create overlay container
        const picker = document.createElement('div');
        picker.className = 'custom-select-overlay';
        picker.id = `picker-${Date.now()}`;
        picker.setAttribute('role', 'listbox');
        picker.setAttribute('tabindex', '-1');

        try {
            const safe = (v) => String(v || '').replace(/[^A-Za-z0-9_]/g, '_');
            const base = selectEl.getAttribute('data-testid')
                || (selectEl.id ? ('select-' + selectEl.id) : '')
                || (selectEl.name ? ('select-' + selectEl.name) : '')
                || 'select';
            picker.setAttribute('data-testid', 'custom-select-overlay-' + safe(base));
            picker.dataset.testidBase = safe(base);
        } catch (_e) {}
        
        // Build options list
        const options = Array.from(selectEl.options);
        options.forEach((opt, index) => {
            const isEmptyPlaceholder = !opt.value && !opt.textContent?.trim();
            if (opt.disabled && !opt.value) return; // Skip disabled placeholder
            if (isEmptyPlaceholder) return; // Не показываем пустой option
            
            const item = document.createElement('div');
            item.className = 'custom-select-item';
            item.setAttribute('role', 'option');
            item.setAttribute('data-value', opt.value);
            item.setAttribute('data-index', index);
            item.textContent = opt.textContent;

            try {
                const safe = (v) => String(v || '').replace(/[^A-Za-z0-9_]/g, '_');
                const base = picker.dataset.testidBase || 'select';
                const vSafe = safe(opt.value || opt.textContent || index);
                item.setAttribute('data-testid', `custom-select-item-${base}-${vSafe}`);
            } catch (_e) {}
            
            if (opt.selected) {
                item.classList.add('selected');
                item.setAttribute('aria-selected', 'true');
            }
            
            if (opt.disabled) {
                item.classList.add('disabled');
                item.setAttribute('aria-disabled', 'true');
            }
            
            item.addEventListener('click', (e) => {
                e.stopPropagation();
                if (!opt.disabled) {
                    this.#selectOption(selectEl, index);
                    this.closeCurrent();
                }
            });
            
            picker.appendChild(item);
        });
        
        // Add to DOM first (hidden) so we can measure dimensions
        picker.style.visibility = 'hidden';
        picker.style.position = 'fixed';
        document.body.appendChild(picker);
        
        // Now position the picker (after it's in DOM and has dimensions)
        this.#positionPicker(picker, selectEl);
        
        // Show the picker
        picker.style.visibility = 'visible';
        
        this.#currentPicker = picker;
        this.#currentSelectEl = selectEl;
        
        // Focus management
        if (byKeyboard) {
            picker.focus();
        }
        
        // Keyboard navigation within picker
        const handleKeydown = (e) => {
            const items = picker.querySelectorAll('.custom-select-item:not(.disabled)');
            const currentIndex = Array.from(items).findIndex(el => el.classList.contains('highlighted'));
            
            switch (e.key) {
                case 'ArrowDown':
                    e.preventDefault();
                    this.#highlightItem(items, currentIndex + 1);
                    break;
                case 'ArrowUp':
                    e.preventDefault();
                    this.#highlightItem(items, currentIndex - 1);
                    break;
                case 'Enter':
                case ' ':
                    e.preventDefault();
                    const highlighted = picker.querySelector('.custom-select-item.highlighted');
                    if (highlighted) {
                        const idx = parseInt(highlighted.dataset.index);
                        this.#selectOption(selectEl, idx);
                    }
                    this.closeCurrent();
                    selectEl.focus();
                    break;
                case 'Escape':
                    e.preventDefault();
                    this.closeCurrent();
                    selectEl.focus();
                    break;
                case 'Tab':
                    this.closeCurrent();
                    break;
            }
        };
        
        // Close on outside click
        const handleClickOutside = (e) => {
            if (!picker.contains(e.target) && e.target !== selectEl) {
                this.closeCurrent();
            }
        };
        
        // Attach event listeners
        picker.addEventListener('keydown', handleKeydown);
        document.addEventListener('mousedown', handleClickOutside);
        
        // Cleanup function
        this.#currentCleanup = () => {
            picker.removeEventListener('keydown', handleKeydown);
            document.removeEventListener('mousedown', handleClickOutside);
        };
        
        // Highlight currently selected item
        const selectedItem = picker.querySelector('.custom-select-item.selected');
        if (selectedItem) {
            selectedItem.classList.add('highlighted');
            selectedItem.scrollIntoView({ block: 'nearest' });
        }
    }
    
    /**
     * Position picker relative to select element
     * @param {HTMLElement} picker 
     * @param {HTMLSelectElement} selectEl 
     */
    static #positionPicker(picker, selectEl) {
        const rect = selectEl.getBoundingClientRect();
        const viewportHeight = window.innerHeight;
        const viewportWidth = window.innerWidth;
        
        // Calculate available space
        const spaceBelow = viewportHeight - rect.bottom;
        const spaceAbove = rect.top;
        
        // Get actual picker height (it's already in DOM but hidden)
        const pickerHeight = Math.min(300, picker.scrollHeight);
        
        // Default: position below, aligned to left edge of select
        let top = rect.bottom + 2;
        let maxHeight = Math.min(300, spaceBelow - 10);
        
        // If not enough space below and more space above, position above
        if (spaceBelow < pickerHeight + 10 && spaceAbove > spaceBelow) {
            maxHeight = Math.min(300, spaceAbove - 10);
            // Position above: top of select minus picker height
            top = rect.top - Math.min(pickerHeight, maxHeight) - 2;
        }
        
        // Left edge aligned with select, but don't go off-screen
        let left = rect.left;
        if (left < 5) left = 5;
        
        // Width: match select width, but allow natural expansion for content
        // Don't force minWidth larger than select itself
        const minWidth = rect.width;
        
        // Check if picker would go off right edge
        const actualPickerWidth = Math.max(picker.scrollWidth, minWidth);
        if (left + actualPickerWidth > viewportWidth - 5) {
            // Align to right edge of select instead
            left = rect.right - actualPickerWidth;
            if (left < 5) left = 5;
        }
        
        Object.assign(picker.style, {
            position: 'fixed',
            top: `${Math.max(5, top)}px`,
            left: `${left}px`,
            minWidth: `${minWidth}px`,
            maxHeight: `${maxHeight}px`,
            overflowY: 'auto',
            zIndex: '100000'
        });
    }
    
    /**
     * Highlight an item in the picker
     * @param {NodeList} items 
     * @param {number} index 
     */
    static #highlightItem(items, index) {
        if (!items.length) return;
        
        // Wrap around
        if (index < 0) index = items.length - 1;
        if (index >= items.length) index = 0;
        
        items.forEach(el => el.classList.remove('highlighted'));
        items[index].classList.add('highlighted');
        items[index].scrollIntoView({ block: 'nearest' });
    }
    
    /**
     * Select an option and update the native select
     * @param {HTMLSelectElement} selectEl 
     * @param {number} index 
     */
    static #selectOption(selectEl, index) {
        selectEl.selectedIndex = index;
        
        // Dispatch change event
        const event = new Event('change', { bubbles: true });
        selectEl.dispatchEvent(event);
    }
    
    /**
     * Inject default styles for the overlay
     * Call this once on page load
     */
    static injectStyles() {
        if (document.getElementById('custom-select-overlay-styles')) {
            return; // Already injected
        }
        
        const style = document.createElement('style');
        style.id = 'custom-select-overlay-styles';
        style.textContent = `
            .custom-select-overlay {
                background: var(--bg-content, #ffffff);
                border: 1px solid var(--border-color, #e1e1e1);
                border-radius: 3px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
                /* Match select font-family from fonts.css and connection-manager.css */
                font-family: 'Nunito', "Helvetica Neue", Helvetica, Arial, sans-serif;
                font-size: 11px;
            }
            
            .custom-select-item {
                /* Match select padding: 5px 8px from connection-manager.css */
                padding: 5px 8px;
                cursor: pointer;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                /* Match select color: var(--sql-text-primary) = #000000 */
                color: var(--sql-text-primary, #000000);
                /* Match typical select option height */
                line-height: 1.4;
                box-sizing: border-box;
            }
            
            .custom-select-item:hover,
            .custom-select-item.highlighted {
                background: var(--bg-hover, #e8e8e8);
            }
            
            .custom-select-item.selected {
                background: var(--bg-selected, #0030b2);
                color: var(--text-selected, #ffffff);
            }
            
            .custom-select-item.selected:hover,
            .custom-select-item.selected.highlighted {
                background: var(--bg-selected-hover, #00a4f6);
            }
            
            .custom-select-item.disabled {
                opacity: 0.5;
                cursor: not-allowed;
            }
        `;
        
        document.head.appendChild(style);
    }
}

// Auto-inject styles when module loads
if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => CustomSelectOverlay.injectStyles());
    } else {
        CustomSelectOverlay.injectStyles();
    }
}

// Export to window for compatibility with non-module scripts
if (typeof window !== 'undefined') {
    window.CustomSelectOverlay = CustomSelectOverlay;
}

// Register global cleanup (important when plugin host hides iframe on close)
try {
    if (typeof window !== 'undefined' &&
        window.PluginCleanup &&
        typeof window.PluginCleanup.register === 'function') {
        window.PluginCleanup.register('custom-select-overlay', () => {
            try {
                CustomSelectOverlay.closeCurrent();
            } catch (_) {}
            try {
                CustomSelectOverlay.disableAutoAttachAll();
            } catch (_) {}
            try {
                document.querySelectorAll('.custom-select-overlay').forEach((el) => el.remove());
            } catch (_) {}
        });
    }
} catch (_) {}
