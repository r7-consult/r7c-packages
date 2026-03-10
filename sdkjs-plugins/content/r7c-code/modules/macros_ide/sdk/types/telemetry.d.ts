/**
 * @fileoverview Telemetry System Type Definitions
 * @description Detailed types for the telemetry and monitoring system
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

export type TelemetryEventType = 
  | 'error' 
  | 'performance' 
  | 'usage' 
  | 'security' 
  | 'api' 
  | 'component' 
  | 'user_action';

export interface TelemetryConfig {
  enabled?: boolean;
  endpoint?: string;
  apiKey?: string;
  bufferSize?: number;
  flushInterval?: number;
  maxRetries?: number;
  debug?: boolean;
  sampling?: TelemetrySampling;
  excludeEvents?: string[];
  consent?: () => boolean;
  headers?: Record<string, string>;
}

export interface TelemetrySampling {
  performance?: number;
  usage?: number;
  error?: number;
  api?: number;
  security?: number;
  component?: number;
  user_action?: number;
}

export interface TelemetryEvent {
  readonly id: string;
  readonly type: TelemetryEventType;
  readonly event: string;
  readonly data: Record<string, any>;
  readonly context: Record<string, any>;
  readonly timestamp: number;
  readonly sessionId: string;
  readonly environment?: EnvironmentInfo;
  readonly sdk?: {
    version: string;
    sessionId: string;
    timestamp: number;
  };
}

export interface EnvironmentInfo {
  browser?: {
    userAgent: string;
    language: string;
    platform: string;
    cookieEnabled: boolean;
    onLine: boolean;
  };
  screen?: {
    width: number;
    height: number;
    availWidth: number;
    availHeight: number;
    colorDepth: number;
  };
  timing?: {
    navigationStart: number;
    loadEventEnd: number;
    domContentLoadedEventEnd: number;
    pageLoadTime: number;
  };
}

export interface TelemetryMetrics {
  session: {
    id: string;
    duration: number;
    startTime: number;
  };
  events: Record<string, number>;
  buffer: {
    size: number;
    maxSize: number;
  };
  enabled: boolean;
}

export interface TelemetryCollector {
  (): Record<string, any>;
}

export interface TelemetryProcessor {
  (event: TelemetryEvent): TelemetryEvent | null;
}

export interface TelemetryReporter {
  send(events: TelemetryEvent[]): Promise<void>;
}

export interface HTTPReporterOptions {
  method?: string;
  headers?: Record<string, string>;
  timeout?: number;
}

export interface ConsoleReporterOptions {
  grouped?: boolean;
  detailed?: boolean;
}

export declare class Telemetry {
  constructor(options?: TelemetryConfig);
  
  readonly enabled: boolean;
  
  enable(): boolean;
  disable(): void;
  
  track(
    type: TelemetryEventType, 
    event: string, 
    data?: Record<string, any>, 
    context?: Record<string, any>
  ): void;
  
  trackError(error: Error, context?: Record<string, any>): void;
  trackPerformance(operation: string, duration: number, metadata?: Record<string, any>): void;
  trackAPI(method: string, duration?: number, success?: boolean, metadata?: Record<string, any>): void;
  trackComponent(component: string, action: string, metadata?: Record<string, any>): void;
  trackUserAction(action: string, target?: string, metadata?: Record<string, any>): void;
  trackSecurity(
    event: string, 
    severity: 'low' | 'medium' | 'high' | 'critical', 
    details?: Record<string, any>
  ): void;
  
  timeOperation<T>(
    operation: string, 
    fn: () => Promise<T>, 
    metadata?: Record<string, any>
  ): Promise<T>;
  
  mark(name: string, metadata?: Record<string, any>): void;
  measure(name: string, startMark: string, endMark?: string, metadata?: Record<string, any>): void;
  
  getMetrics(): TelemetryMetrics;
  flush(): Promise<boolean>;
  
  addCollector(name: string, collector: TelemetryCollector): void;
  addProcessor(name: string, processor: TelemetryProcessor): void;
  addReporter(reporter: TelemetryReporter): void;
  
  destroy(): void;
}

export declare class HTTPReporter implements TelemetryReporter {
  constructor(endpoint: string, options?: HTTPReporterOptions);
  send(events: TelemetryEvent[]): Promise<void>;
}

export declare class ConsoleReporter implements TelemetryReporter {
  constructor(options?: ConsoleReporterOptions);
  send(events: TelemetryEvent[]): Promise<void>;
}

export interface TelemetryIntegrationConfig extends TelemetryConfig {
  consent?: () => boolean;
}

export interface EventSystemIntegrationOptions {
  trackAllEvents?: boolean;
}

export interface TelemetryDashboardData {
  session: TelemetryMetrics['session'];
  summary: {
    totalEvents: number;
    errorCount: number;
    performanceEvents: number;
    userActions: number;
  };
  events: Record<string, number>;
  buffer: TelemetryMetrics['buffer'];
  integrations: string[];
}

export declare class TelemetryIntegration {
  static initialize(config?: TelemetryIntegrationConfig): TelemetryIntegration;
  static getInstance(): Telemetry | null;
  
  static integrateEventSystem(
    eventSystem: any, 
    options?: EventSystemIntegrationOptions
  ): void;
  
  static integrateLazyLoader(lazyLoader: any): void;
  static integrateAPIGuard(apiGuard: any): void;
  
  static wrapComponent<T extends new (...args: any[]) => any>(
    ComponentClass: T, 
    componentName: string
  ): T;
  
  static performanceMonitor(
    operationName: string, 
    metadata?: Record<string, any>
  ): MethodDecorator;
  
  static trackUserAction(
    action: string, 
    dataExtractor?: (event: Event) => Record<string, any>
  ): (event: Event) => void;
  
  static setupPageTracking(): void;
  static getDashboardData(): TelemetryDashboardData;
  static cleanup(): void;
}

// Convenience export for direct telemetry access
export interface TelemetryConvenience {
  init(config: TelemetryIntegrationConfig): TelemetryIntegration;
  track(type: TelemetryEventType, event: string, data?: Record<string, any>): void;
  error(error: Error, context?: Record<string, any>): void;
  performance(operation: string, duration: number, metadata?: Record<string, any>): void;
  time<T>(operation: string, fn: () => Promise<T>, metadata?: Record<string, any>): Promise<T>;
  dashboard(): TelemetryDashboardData;
  cleanup(): void;
}

export declare const telemetry: TelemetryConvenience;

export { TelemetryEventType as EventType };
export default Telemetry;