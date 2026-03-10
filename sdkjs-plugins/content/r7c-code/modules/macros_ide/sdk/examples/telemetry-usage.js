/**
 * @fileoverview Telemetry Usage Examples for OnlyOffice UI SDK
 * @description Demonstrates how to use the telemetry system in production
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

import { TelemetryIntegration, telemetry } from '../src/core/telemetry-integration.js';
import { TelemetryEventType } from '../src/core/telemetry.js';

/**
 * Example: Basic telemetry setup
 */
function setupBasicTelemetry() {
    // Initialize telemetry with basic configuration
    const integration = TelemetryIntegration.initialize({
        enabled: true,
        endpoint: 'https://your-analytics-endpoint.com/events',
        apiKey: 'your-api-key',
        debug: process.env.NODE_ENV === 'development',
        
        // Consent function - implement your consent logic
        consent: () => {
            // Check if user has given consent for analytics
            return localStorage.getItem('analytics-consent') === 'true';
        },
        
        // Sampling rates to control data volume
        sampling: {
            performance: 1.0,    // Track all performance events
            usage: 1.0,          // Track all usage events  
            error: 1.0,          // Track all errors
            api: 0.1,            // Sample 10% of API calls
            security: 1.0        // Track all security events
        }
    });

    // Setup automatic page tracking
    TelemetryIntegration.setupPageTracking();
    
    console.log('Telemetry initialized');
    return integration;
}

/**
 * Example: Advanced telemetry setup with custom reporters
 */
function setupAdvancedTelemetry() {
    const integration = TelemetryIntegration.initialize({
        enabled: true,
        debug: true,
        
        // Custom consent logic
        consent: () => {
            // Check multiple consent sources
            return (
                localStorage.getItem('telemetry-consent') === 'true' ||
                document.cookie.includes('analytics=accepted') ||
                window.userConsent?.analytics === true
            );
        },
        
        // Custom headers for authentication
        headers: {
            'X-Application': 'OnlyOffice-SDK',
            'X-Version': '1.0.0'
        }
    });

    const telemetryInstance = TelemetryIntegration.getInstance();

    // Add custom data collector
    telemetryInstance.addCollector('user', () => ({
        userId: getCurrentUserId(),
        sessionLength: getSessionLength(),
        featureFlags: getActiveFeatureFlags()
    }));

    // Add custom event processor for data enrichment
    telemetryInstance.addProcessor('enrich', (event) => ({
        ...event,
        buildVersion: getBuildVersion(),
        deploymentEnvironment: getDeploymentEnvironment(),
        userTier: getUserTier()
    }));

    return integration;
}

/**
 * Example: Integrating telemetry with SDK components
 */
function integrateWithSDKComponents(eventSystem, lazyLoader, apiGuard) {
    // Integrate with event system
    TelemetryIntegration.integrateEventSystem(eventSystem, {
        trackAllEvents: true  // Track all event emissions
    });

    // Integrate with lazy loader
    TelemetryIntegration.integrateLazyLoader(lazyLoader);

    // Integrate with API guard
    TelemetryIntegration.integrateAPIGuard(apiGuard);

    console.log('SDK components integrated with telemetry');
}

/**
 * Example: Creating telemetry-aware components
 */
class ExampleComponent {
    constructor(options) {
        this.options = options;
        
        // Track component creation
        telemetry.track(
            TelemetryEventType.COMPONENT,
            'component_created',
            {
                component: 'ExampleComponent',
                options: Object.keys(options)
            }
        );
    }

    async initialize() {
        // Time the initialization operation
        return telemetry.time(
            'component_initialize',
            async () => {
                // Simulate async initialization
                await new Promise(resolve => setTimeout(resolve, 100));
                
                telemetry.track(
                    TelemetryEventType.COMPONENT,
                    'component_initialized',
                    { component: 'ExampleComponent' }
                );
                
                return true;
            },
            { component: 'ExampleComponent' }
        );
    }

    performAction(actionType, data) {
        try {
            // Track user action
            telemetry.track(
                TelemetryEventType.USER_ACTION,
                'action_performed',
                {
                    actionType,
                    component: 'ExampleComponent',
                    dataSize: JSON.stringify(data).length
                }
            );

            // Simulate action processing
            const result = this.processAction(actionType, data);
            
            return result;
        } catch (error) {
            // Track errors automatically
            telemetry.error(error, {
                operation: 'performAction',
                component: 'ExampleComponent',
                actionType
            });
            
            throw error;
        }
    }

    processAction(actionType, data) {
        // Simulate processing
        if (actionType === 'error') {
            throw new Error('Simulated error');
        }
        
        return { success: true, processed: data };
    }
}

/**
 * Example: Using component wrapper for automatic telemetry
 */
function createTelemetryWrappedComponent() {
    // Wrap component class for automatic telemetry
    const TelemetryExampleComponent = TelemetryIntegration.wrapComponent(
        ExampleComponent,
        'WrappedExampleComponent'
    );

    return TelemetryExampleComponent;
}

/**
 * Example: Performance monitoring decorator
 */
class PerformanceMonitoredComponent {
    constructor(options) {
        this.options = options;
    }

    // Use decorator for automatic performance monitoring
    @TelemetryIntegration.performanceMonitor('database_query', { type: 'read' })
    async fetchData(query) {
        // Simulate database query
        await new Promise(resolve => setTimeout(resolve, Math.random() * 1000));
        return { data: 'example data', query };
    }

    @TelemetryIntegration.performanceMonitor('ui_render')
    render() {
        // Simulate rendering
        const startTime = performance.now();
        
        // Rendering logic here
        for (let i = 0; i < 10000; i++) {
            // Simulate work
        }
        
        return { rendered: true, elements: 10000 };
    }
}

/**
 * Example: Manual tracking for specific scenarios
 */
function manualTrackingExamples() {
    // Track feature usage
    telemetry.track(
        TelemetryEventType.USAGE,
        'feature_used',
        {
            feature: 'advanced_search',
            query_complexity: 'high',
            results_count: 42
        }
    );

    // Track security events
    telemetry.track(
        TelemetryEventType.SECURITY,
        'security_event',
        {
            event: 'invalid_token',
            severity: 'medium',
            source: 'api_authentication'
        }
    );

    // Track API performance
    const apiStartTime = performance.now();
    
    // Simulate API call
    setTimeout(() => {
        const duration = performance.now() - apiStartTime;
        
        telemetry.performance(
            'api_call',
            duration,
            {
                endpoint: '/api/documents',
                method: 'GET',
                status_code: 200,
                response_size: 1024
            }
        );
    }, 150);
}

/**
 * Example: Error tracking with context
 */
function errorTrackingExamples() {
    try {
        // Simulate operation that might fail
        throw new Error('Network connection failed');
    } catch (error) {
        // Track error with rich context
        telemetry.error(error, {
            operation: 'document_save',
            user_action: 'save_button_click',
            document_size: 1024000,
            network_status: 'offline',
            retry_attempt: 2
        });
    }

    // Track handled errors
    const result = handleRiskyOperation();
    if (!result.success) {
        telemetry.track(
            TelemetryEventType.ERROR,
            'operation_failed',
            {
                operation: 'risky_operation',
                error_code: result.errorCode,
                error_message: result.errorMessage,
                recoverable: result.canRetry
            }
        );
    }
}

/**
 * Example: Custom user action tracking
 */
function setupUserActionTracking() {
    // Track button clicks
    document.addEventListener('click', TelemetryIntegration.trackUserAction(
        'button_click',
        (event) => ({
            button_text: event.target.textContent,
            button_id: event.target.id,
            position: {
                x: event.clientX,
                y: event.clientY
            }
        })
    ));

    // Track form submissions
    document.addEventListener('submit', TelemetryIntegration.trackUserAction(
        'form_submit',
        (event) => ({
            form_id: event.target.id,
            form_fields: Array.from(event.target.elements)
                .filter(el => el.name)
                .map(el => ({ name: el.name, type: el.type }))
        })
    ));

    // Track search queries
    document.addEventListener('input', (event) => {
        if (event.target.classList.contains('search-input')) {
            // Debounce search tracking
            clearTimeout(event.target.searchTimeout);
            event.target.searchTimeout = setTimeout(() => {
                telemetry.track(
                    TelemetryEventType.USER_ACTION,
                    'search_query',
                    {
                        query_length: event.target.value.length,
                        has_filters: document.querySelectorAll('.filter.active').length > 0
                    }
                );
            }, 500);
        }
    });
}

/**
 * Example: Monitoring dashboard data
 */
function displayTelemetryDashboard() {
    const dashboardData = telemetry.dashboard();
    
    console.log('=== Telemetry Dashboard ===');
    console.log('Session:', dashboardData.session);
    console.log('Summary:', dashboardData.summary);
    console.log('Active Integrations:', dashboardData.integrations);
    
    // Update UI dashboard (if you have one)
    updateDashboardUI(dashboardData);
}

/**
 * Example: Cleanup on page unload
 */
function setupCleanup() {
    window.addEventListener('beforeunload', () => {
        // Flush any remaining telemetry data
        const telemetryInstance = TelemetryIntegration.getInstance();
        if (telemetryInstance) {
            telemetryInstance.flush();
        }
    });

    // Cleanup when SDK is destroyed
    window.addEventListener('sdk-destroy', () => {
        telemetry.cleanup();
    });
}

/**
 * Example: A/B testing with telemetry
 */
function trackABTestVariant(testName, variant, userGroup) {
    telemetry.track(
        TelemetryEventType.USAGE,
        'ab_test_exposure',
        {
            test_name: testName,
            variant: variant,
            user_group: userGroup,
            timestamp: Date.now()
        }
    );
}

/**
 * Helper functions (would be implemented based on your app)
 */
function getCurrentUserId() {
    return localStorage.getItem('userId') || 'anonymous';
}

function getSessionLength() {
    const sessionStart = localStorage.getItem('sessionStart');
    return sessionStart ? Date.now() - parseInt(sessionStart) : 0;
}

function getActiveFeatureFlags() {
    return ['feature_a', 'feature_b']; // Your feature flag logic
}

function getBuildVersion() {
    return '1.0.0'; // Your build version
}

function getDeploymentEnvironment() {
    return process.env.NODE_ENV || 'production';
}

function getUserTier() {
    return localStorage.getItem('userTier') || 'free';
}

function handleRiskyOperation() {
    // Simulate operation that might fail
    if (Math.random() < 0.3) {
        return {
            success: false,
            errorCode: 'NETWORK_ERROR',
            errorMessage: 'Failed to connect to server',
            canRetry: true
        };
    }
    
    return { success: true };
}

function updateDashboardUI(data) {
    // Update your telemetry dashboard UI
    console.log('Dashboard UI updated with:', data);
}

// Export examples for testing
export {
    setupBasicTelemetry,
    setupAdvancedTelemetry,
    integrateWithSDKComponents,
    ExampleComponent,
    createTelemetryWrappedComponent,
    PerformanceMonitoredComponent,
    manualTrackingExamples,
    errorTrackingExamples,
    setupUserActionTracking,
    displayTelemetryDashboard,
    setupCleanup,
    trackABTestVariant
};

// Auto-setup if running in browser
if (typeof window !== 'undefined') {
    // Example initialization (customize for your needs)
    document.addEventListener('DOMContentLoaded', () => {
        const integration = setupBasicTelemetry();
        setupUserActionTracking();
        setupCleanup();
        
        // Display dashboard data every 30 seconds in debug mode
        if (process.env.NODE_ENV === 'development') {
            setInterval(displayTelemetryDashboard, 30000);
        }
    });
}