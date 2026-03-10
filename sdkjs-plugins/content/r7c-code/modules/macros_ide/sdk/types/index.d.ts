/**
 * @fileoverview TypeScript Definitions for OnlyOffice UI SDK
 * @description Complete type definitions for all SDK components
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

declare module '@onlyoffice/ui-sdk' {
  // Core Types
  export interface SDKConfig {
    debug?: boolean;
    theme?: ThemeConfig;
    errorHandler?: ErrorHandlerConfig;
    telemetry?: TelemetryConfig;
    components?: ComponentConfig;
  }

  export interface ThemeConfig {
    type?: 'light' | 'dark' | 'auto';
    customTheme?: Record<string, string>;
    autoDetect?: boolean;
  }

  export interface ErrorHandlerConfig {
    enabled?: boolean;
    logErrors?: boolean;
    throwOnCritical?: boolean;
    customHandler?: (error: Error, context: Record<string, any>) => void;
  }

  export interface ComponentConfig {
    lazyLoad?: boolean;
    cacheTTL?: number;
    maxCacheSize?: number;
  }

  // Component Base Types
  export interface ComponentOptions {
    container?: HTMLElement | string;
    eventSystem?: EventSystem;
    theme?: ThemeConfig;
    config?: Record<string, any>;
  }

  export interface Component {
    readonly id: string;
    readonly initialized: boolean;
    initialize(): Promise<void>;
    destroy(): Promise<void>;
    on(event: string, handler: EventHandler): () => void;
    emit(event: string, data?: any): void;
  }

  // Event System Types
  export type EventHandler = (data: any, event: string) => void;

  export interface EventSystemOptions {
    debug?: boolean;
    maxListeners?: number;
    throttleEvents?: string[];
    debounceEvents?: string[];
  }

  export interface EventSystem {
    on(event: string, handler: EventHandler, options?: EventOptions): () => void;
    once(event: string, handler: EventHandler, options?: EventOptions): () => void;
    emit(event: string, data?: any, options?: EmitOptions): void;
    off(event: string, handler?: EventHandler): void;
    removeComponent(component: Component): void;
    destroy(): void;
  }

  export interface EventOptions {
    once?: boolean;
    priority?: 'high' | 'normal' | 'low';
    component?: Component;
  }

  export interface EmitOptions {
    async?: boolean;
    timeout?: number;
    throttle?: number;
    debounce?: number;
  }

  // Lazy Loader Types
  export interface LazyLoaderOptions {
    cacheTTL?: number;
    maxCacheSize?: number;
    allowedComponents?: string[];
    debug?: boolean;
  }

  export interface LazyLoader {
    loadComponent(name: string, options?: ComponentOptions): Promise<Component>;
    preloadComponent(name: string): Promise<void>;
    clearCache(): void;
    getCacheStatus(): CacheStatus;
  }

  export interface CacheStatus {
    size: number;
    maxSize: number;
    components: string[];
    memory: {
      used: number;
      available: number;
    };
  }

  // OnlyOffice Integration Types
  export interface OnlyOfficeAPI {
    executeMethod(method: string, args?: any[], callback?: Function): Promise<any>;
    callCommand(command: Function, isNoCalc?: boolean): void;
    resizeWindow(width: number, height: number): void;
  }

  export interface OnlyOfficeAPIGuardOptions {
    retryAttempts?: number;
    retryDelay?: number;
    timeout?: number;
    fallbackMode?: boolean;
    debug?: boolean;
  }

  export interface OnlyOfficeAPIGuard {
    readonly isAvailable: boolean;
    readonly version: string | null;
    checkAvailability(): Promise<boolean>;
    safeAPICall<T>(apiCall: () => Promise<T>, fallback?: T): Promise<T>;
    withFallback<T>(operation: () => Promise<T>, fallback: T): Promise<T>;
  }

  export interface OnlyOfficeTheme {
    type: 'light' | 'dark';
    'background-normal': string;
    'background-toolbar': string;
    'background-tab-underline': string;
    'text-normal': string;
    'text-normal-pressed': string;
    'highlight-button-tab': string;
    'border-toolbar': string;
    'border-divider': string;
    'icon-normal': string;
    'icon-normal-pressed': string;
  }

  // Error Handling Types
  export class SDKError extends Error {
    readonly code: string;
    readonly context: Record<string, any>;
    readonly timestamp: number;
    readonly recoverable: boolean;
    
    constructor(message: string, code?: string, context?: Record<string, any>);
  }

  export class ComponentError extends SDKError {
    readonly component: string;
    
    constructor(component: string, message: string, context?: Record<string, any>);
  }

  export class ConfigError extends SDKError {
    readonly field: string;
    readonly expected: string;
    readonly actual: any;
    
    constructor(context: string, message: string, config?: any);
  }

  export class APIError extends SDKError {
    readonly method: string;
    readonly endpoint?: string;
    
    constructor(method: string, message: string, context?: Record<string, any>);
  }

  export interface ErrorHandler {
    handleError(error: Error, context?: Record<string, any>): void;
    validateConfig(config: any, schema: ValidationSchema, context?: string): void;
    createRecoveryStrategy(error: Error): RecoveryStrategy | null;
  }

  export interface ValidationSchema {
    [field: string]: ValidationRule;
  }

  export interface ValidationRule {
    type: 'string' | 'number' | 'boolean' | 'object' | 'array' | 'function';
    required?: boolean;
    validator?: (value: any) => boolean;
    transform?: (value: any) => any;
  }

  export interface RecoveryStrategy {
    canRecover: boolean;
    action: () => Promise<void>;
    fallback?: () => any;
  }

  // Telemetry Types
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
  }

  export interface TelemetryEvent {
    readonly id: string;
    readonly type: TelemetryEventType;
    readonly event: string;
    readonly data: Record<string, any>;
    readonly context: Record<string, any>;
    readonly timestamp: number;
    readonly sessionId: string;
  }

  export interface Telemetry {
    readonly enabled: boolean;
    enable(): boolean;
    disable(): void;
    track(type: TelemetryEventType, event: string, data?: Record<string, any>, context?: Record<string, any>): void;
    trackError(error: Error, context?: Record<string, any>): void;
    trackPerformance(operation: string, duration: number, metadata?: Record<string, any>): void;
    trackAPI(method: string, duration?: number, success?: boolean, metadata?: Record<string, any>): void;
    trackComponent(component: string, action: string, metadata?: Record<string, any>): void;
    trackUserAction(action: string, target?: string, metadata?: Record<string, any>): void;
    trackSecurity(event: string, severity: 'low' | 'medium' | 'high' | 'critical', details?: Record<string, any>): void;
    timeOperation<T>(operation: string, fn: () => Promise<T>, metadata?: Record<string, any>): Promise<T>;
    mark(name: string, metadata?: Record<string, any>): void;
    measure(name: string, startMark: string, endMark?: string, metadata?: Record<string, any>): void;
    getMetrics(): TelemetryMetrics;
    flush(): Promise<boolean>;
    addCollector(name: string, collector: () => Record<string, any>): void;
    addProcessor(name: string, processor: (event: TelemetryEvent) => TelemetryEvent | null): void;
    addReporter(reporter: TelemetryReporter): void;
    destroy(): void;
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

  export interface TelemetryReporter {
    send(events: TelemetryEvent[]): Promise<void>;
  }

  export interface TelemetryIntegration {
    initialize(config: TelemetryConfig): TelemetryIntegration;
    getInstance(): Telemetry | null;
    integrateEventSystem(eventSystem: EventSystem, options?: EventSystemIntegrationOptions): void;
    integrateLazyLoader(lazyLoader: LazyLoader): void;
    integrateAPIGuard(apiGuard: OnlyOfficeAPIGuard): void;
    wrapComponent<T extends Component>(ComponentClass: new (...args: any[]) => T, componentName: string): new (...args: any[]) => T;
    performanceMonitor(operationName: string, metadata?: Record<string, any>): MethodDecorator;
    trackUserAction(action: string, dataExtractor?: (event: Event) => Record<string, any>): (event: Event) => void;
    setupPageTracking(): void;
    getDashboardData(): TelemetryDashboardData;
    cleanup(): void;
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

  // Chat Interface Types
  export interface ChatInterfaceOptions extends ComponentOptions {
    apiKey?: string;
    endpoint?: string;
    model?: string;
    maxTokens?: number;
    temperature?: number;
    systemPrompt?: string;
    allowMarkdown?: boolean;
    autoScroll?: boolean;
    maxMessages?: number;
  }

  export interface ChatMessage {
    id: string;
    role: 'user' | 'assistant' | 'system';
    content: string;
    timestamp: number;
    metadata?: Record<string, any>;
  }

  export interface ChatInterface extends Component {
    sendMessage(content: string, options?: SendMessageOptions): Promise<ChatMessage>;
    getMessages(): ChatMessage[];
    clearMessages(): void;
    setSystemPrompt(prompt: string): void;
    getConversationContext(): string;
  }

  export interface SendMessageOptions {
    role?: 'user' | 'system';
    metadata?: Record<string, any>;
    stream?: boolean;
    onToken?: (token: string) => void;
  }

  // File Manager Types
  export interface FileManagerOptions extends ComponentOptions {
    allowedTypes?: string[];
    maxFileSize?: number;
    allowMultiple?: boolean;
    showPreview?: boolean;
    enableUpload?: boolean;
    uploadEndpoint?: string;
  }

  export interface FileItem {
    id: string;
    name: string;
    type: string;
    size: number;
    lastModified: number;
    url?: string;
    preview?: string;
    metadata?: Record<string, any>;
  }

  export interface FileManager extends Component {
    getFiles(): FileItem[];
    addFile(file: File | FileItem): Promise<FileItem>;
    removeFile(id: string): Promise<void>;
    selectFile(id: string): void;
    getSelectedFiles(): FileItem[];
    uploadFiles(files: File[]): Promise<FileItem[]>;
  }

  // Document Viewer Types
  export interface DocumentViewerOptions extends ComponentOptions {
    documentUrl?: string;
    documentType?: 'pdf' | 'docx' | 'xlsx' | 'pptx' | 'txt';
    enableAnnotations?: boolean;
    readonly?: boolean;
    zoom?: number;
    fitToWidth?: boolean;
  }

  export interface DocumentViewer extends Component {
    loadDocument(url: string, type?: string): Promise<void>;
    getDocument(): Document | null;
    setZoom(level: number): void;
    fitToWidth(): void;
    fitToHeight(): void;
    goToPage(page: number): void;
    getCurrentPage(): number;
    getTotalPages(): number;
  }

  // Main SDK Class
  export interface OnlyOfficeUISDK {
    readonly version: string;
    readonly initialized: boolean;
    readonly config: SDKConfig;
    
    initialize(config?: SDKConfig): Promise<void>;
    destroy(): Promise<void>;
    
    // Component access
    createComponent<T extends Component>(type: string, options?: ComponentOptions): Promise<T>;
    getComponent<T extends Component>(id: string): T | null;
    destroyComponent(id: string): Promise<void>;
    
    // System access
    readonly eventSystem: EventSystem;
    readonly lazyLoader: LazyLoader;
    readonly apiGuard: OnlyOfficeAPIGuard;
    readonly errorHandler: ErrorHandler;
    readonly telemetry: Telemetry;
    
    // Utilities
    waitForOnlyOffice(): Promise<OnlyOfficeAPI>;
    getTheme(): OnlyOfficeTheme;
    setTheme(theme: Partial<OnlyOfficeTheme>): void;
  }

  // Main SDK factory function
  export function createSDK(config?: SDKConfig): Promise<OnlyOfficeUISDK>;
  
  // Component factories
  export function createChatInterface(options?: ChatInterfaceOptions): Promise<ChatInterface>;
  export function createFileManager(options?: FileManagerOptions): Promise<FileManager>;
  export function createDocumentViewer(options?: DocumentViewerOptions): Promise<DocumentViewer>;
  
  // Utility exports
  export const SDK: OnlyOfficeUISDK;
  export const EventTypes: Record<string, string>;
  export const ComponentTypes: Record<string, string>;
  export const TelemetryEventTypes: Record<TelemetryEventType, TelemetryEventType>;
}

// Global type augmentations for OnlyOffice API
declare global {
  interface Window {
    Asc?: {
      plugin?: {
        executeMethod: (method: string, args?: any[], callback?: Function) => Promise<any>;
        callCommand: (command: Function, isNoCalc?: boolean) => void;
        resizeWindow: (width: number, height: number) => void;
        info: {
          documentType: 'word' | 'cell' | 'slide';
          editorType: 'desktop' | 'mobile' | 'embedded';
          version: string;
        };
        theme: OnlyOfficeTheme;
        onDocumentStateChange?: (data: any) => void;
        onThemeChanged?: (theme: OnlyOfficeTheme) => void;
        init?: () => void;
        button?: (id: number, text?: string) => void;
      };
    };
    
    Api?: {
      // Word API
      GetDocument?: () => any;
      CreateDocument?: () => any;
      
      // Cell API  
      GetActiveSheet?: () => any;
      GetWorksheet?: (index: number) => any;
      
      // Slide API
      GetPresentation?: () => any;
      CreatePresentation?: () => any;
    };
  }
}

export {};
