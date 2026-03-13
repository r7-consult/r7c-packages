/**
 * Fallback for OnlyOffice plugin scripts when running standalone
 */

// Create fallback for missing OnlyOffice plugin API
if (typeof window.Asc === 'undefined') {
    console.log('📌 Creating OnlyOffice API fallback for standalone mode');
    
    window.Asc = {
        plugin: {
            init: function() {
                console.log('Fallback: plugin.init called');
            },
            button: function(id) {
                console.log('Fallback: plugin.button called with id:', id);
            },
            executeCommand: function(cmd, params) {
                console.log('Fallback: executeCommand:', cmd, params);
            },
            executeMethod: function(method, params, callback) {
                console.log('Fallback: executeMethod:', method, params);
                if (callback) callback({});
            },
            onExternalMouseUp: function() {
                console.log('Fallback: onExternalMouseUp');
            },
            info: {
                guid: 'standalone-mode',
                data: null,
                objectId: null,
                width: 70,
                height: 70,
                mmToPx: 3.78,
                resize: false
            },
            resizeWindow: function(w, h) {
                console.log('Fallback: resizeWindow', w, h);
            }
        }
    };
}

// Ensure these functions exist even if plugin files fail to load
if (typeof window.Common === 'undefined') {
    window.Common = {
        Gateway: {
            PluginInfo: {}
        }
    };
}