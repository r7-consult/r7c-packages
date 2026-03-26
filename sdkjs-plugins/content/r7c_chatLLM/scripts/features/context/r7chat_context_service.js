(function (window) {
    'use strict';

    var root = window.R7Chat = window.R7Chat || {};
    root.features = root.features || {};

    function getService() {
        return root.services && root.services.context ? root.services.context : null;
    }

    root.features.context = {
        getWorkbookSheets: function (options) {
            var service = getService();
            if (!service || typeof service.getWorkbookSheets !== 'function') {
                return Promise.reject(new Error('Context service is unavailable'));
            }
            return service.getWorkbookSheets(options);
        },
        getActiveSheet: function (options) {
            var service = getService();
            if (!service || typeof service.getActiveSheet !== 'function') {
                return Promise.reject(new Error('Context service is unavailable'));
            }
            return service.getActiveSheet(options);
        }
    };
})(window);
