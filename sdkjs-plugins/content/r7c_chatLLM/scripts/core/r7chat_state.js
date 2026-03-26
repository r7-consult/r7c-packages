(function (window) {
    'use strict';

    const root = window.R7Chat = window.R7Chat || {};
    const state = root.state = root.state || {};

    state.metrics = state.metrics || {};
    state.metrics.ui = state.metrics.ui || {
        contextRenderCalls: 0,
        traceRenderCalls: 0,
        contextRenderScheduled: 0,
        traceRenderScheduled: 0
    };
    state.metrics.contextCache = state.metrics.contextCache || {
        workbookHits: 0,
        workbookMisses: 0,
        activeHits: 0,
        activeMisses: 0,
        invalidations: 0
    };
    state.metrics.planner = state.metrics.planner || {
        requests: 0,
        fastPathRuns: 0
    };

    function bumpMetric(path, delta) {
        const value = typeof delta === 'number' ? delta : 1;
        if (!path || typeof path !== 'string') return;
        const parts = path.split('.');
        let current = state.metrics;
        for (let i = 0; i < parts.length - 1; i++) {
            const key = parts[i];
            if (!current[key] || typeof current[key] !== 'object') {
                current[key] = {};
            }
            current = current[key];
        }
        const last = parts[parts.length - 1];
        const prev = Number(current[last] || 0);
        current[last] = prev + value;
    }

    function resetRunMetrics() {
        state.metrics.ui.contextRenderCalls = 0;
        state.metrics.ui.traceRenderCalls = 0;
        state.metrics.ui.contextRenderScheduled = 0;
        state.metrics.ui.traceRenderScheduled = 0;
        state.metrics.contextCache.workbookHits = 0;
        state.metrics.contextCache.workbookMisses = 0;
        state.metrics.contextCache.activeHits = 0;
        state.metrics.contextCache.activeMisses = 0;
        state.metrics.contextCache.invalidations = 0;
        state.metrics.planner.requests = 0;
        state.metrics.planner.fastPathRuns = 0;
    }

    function snapshotRunMetrics() {
        return JSON.parse(JSON.stringify(state.metrics));
    }

    root.bumpMetric = root.bumpMetric || bumpMetric;
    root.resetRunMetrics = root.resetRunMetrics || resetRunMetrics;
    root.snapshotRunMetrics = root.snapshotRunMetrics || snapshotRunMetrics;
})(window);
