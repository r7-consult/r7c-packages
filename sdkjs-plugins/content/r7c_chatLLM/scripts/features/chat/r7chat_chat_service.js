(function (window) {
    'use strict';

    var root = window.R7Chat = window.R7Chat || {};
    root.features = root.features || {};

    function getService() {
        return root.services && root.services.chat ? root.services.chat : null;
    }

    root.features.chat = {
        request: function () {
            var service = getService();
            if (!service || typeof service.request !== 'function') {
                return Promise.reject(new Error('Chat service is unavailable'));
            }
            return service.request.apply(null, arguments);
        },
        stop: function () {
            var service = getService();
            if (!service || typeof service.stop !== 'function') return;
            return service.stop.apply(null, arguments);
        }
    };
})(window);
