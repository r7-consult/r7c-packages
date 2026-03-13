/**
 * Secure DOM Manipulation Utilities
 * Prevents XSS attacks by avoiding innerHTML and providing safe alternatives
 */

class SecureDOM {
    /**
     * Safely set text content (prevents XSS)
     */
    static setText(element, text) {
        if (!element) return;
        element.textContent = String(text);
    }
    
    /**
     * Safely create element with text
     */
    static createElement(tagName, options = {}) {
        const element = document.createElement(tagName);
        
        if (options.text) {
            element.textContent = String(options.text);
        }
        
        if (options.className) {
            element.className = String(options.className);
        }
        
        if (options.id) {
            element.id = String(options.id);
        }
        
        if (options.attributes) {
            for (const [key, value] of Object.entries(options.attributes)) {
                element.setAttribute(key, String(value));
            }
        }
        
        if (options.styles) {
            Object.assign(element.style, options.styles);
        }
        
        if (options.dataset) {
            for (const [key, value] of Object.entries(options.dataset)) {
                element.dataset[key] = String(value);
            }
        }
        
        return element;
    }
    
    /**
     * Safely build complex DOM structure
     */
    static buildDOM(structure) {
        if (!structure) return null;
        
        const element = this.createElement(structure.tag, {
            text: structure.text,
            className: structure.className,
            id: structure.id,
            attributes: structure.attributes,
            styles: structure.styles,
            dataset: structure.dataset
        });
        
        if (structure.children && Array.isArray(structure.children)) {
            for (const child of structure.children) {
                const childElement = this.buildDOM(child);
                if (childElement) {
                    element.appendChild(childElement);
                }
            }
        }
        
        return element;
    }
    
    /**
     * Sanitize HTML string (basic sanitization)
     */
    static sanitizeHTML(html) {
        const div = document.createElement('div');
        div.textContent = html;
        return div.innerHTML;
    }
    
    /**
     * Clear element safely
     */
    static clearElement(element) {
        if (!element) return;
        
        // Remove event listeners from children
        const children = element.querySelectorAll('*');
        for (const child of children) {
            // Clone node to remove all event listeners
            const clone = child.cloneNode(false);
            if (child.parentNode) {
                child.parentNode.replaceChild(clone, child);
            }
        }
        
        // Clear content
        while (element.firstChild) {
            element.removeChild(element.firstChild);
        }
    }
    
    /**
     * Safe append multiple elements
     */
    static appendChildren(parent, children) {
        if (!parent || !children) return;
        
        const fragment = document.createDocumentFragment();
        
        for (const child of children) {
            if (child instanceof Node) {
                fragment.appendChild(child);
            } else if (typeof child === 'string') {
                fragment.appendChild(document.createTextNode(child));
            }
        }
        
        parent.appendChild(fragment);
    }
    
    /**
     * Escape HTML entities
     */
    static escapeHTML(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }
    
    /**
     * Create safe template
     */
    static template(strings, ...values) {
        let result = strings[0];
        
        for (let i = 0; i < values.length; i++) {
            result += this.escapeHTML(String(values[i]));
            result += strings[i + 1];
        }
        
        return result;
    }
}

// Export for use
window.SecureDOM = SecureDOM;