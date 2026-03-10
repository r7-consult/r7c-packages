/**
 * @fileoverview Core System Type Definitions
 * @description Types for event system, lazy loader, and OnlyOffice integration
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

// Event System Types
export type EventHandler = (data: any, event: string) => void;
export type EventListener = (data: any) => void;

export interface EventOptions {
  once?: boolean;
  priority?: 'high' | 'normal' | 'low';
  component?: any;
  passive?: boolean;
  signal?: AbortSignal;
}

export interface EmitOptions {
  async?: boolean;
  timeout?: number;
  throttle?: number;
  debounce?: number;
  requireListener?: boolean;
}

export interface EventSystemOptions {
  debug?: boolean;
  maxListeners?: number;
  throttleEvents?: string[];
  debounceEvents?: string[];
  warningThreshold?: number;
  enableMetrics?: boolean;
}

export interface EventMetrics {
  eventsEmitted: number;
  listenersRegistered: number;
  errorsOccurred: number;
  averageEmitTime: number;
  mostFrequentEvents: Array<{ event: string; count: number }>;
  memoryUsage: {
    listeners: number;
    components: number;
    throttledEvents: number;
    debouncedEvents: number;
  };
}

export declare class EventSystem {
  constructor(options?: EventSystemOptions);
  
  on(event: string, handler: EventHandler, options?: EventOptions): () => void;
  once(event: string, handler: EventHandler, options?: EventOptions): () => void;
  off(event: string, handler?: EventHandler): void;
  
  emit(event: string, data?: any, options?: EmitOptions): void;
  emitAsync(event: string, data?: any, timeout?: number): Promise<void>;
  
  removeComponent(component: any): void;
  removeAllListeners(event?: string): void;
  
  getListeners(event: string): EventHandler[];
  getEventNames(): string[];
  getListenerCount(event?: string): number;
  
  setMaxListeners(max: number): void;
  getMaxListeners(): number;
  
  enableDebug(): void;
  disableDebug(): void;
  
  getMetrics(): EventMetrics;
  resetMetrics(): void;
  
  destroy(): void;
}

// Lazy Loader Types
export interface LazyLoaderOptions {
  cacheTTL?: number;
  maxCacheSize?: number;
  allowedComponents?: string[];
  debug?: boolean;
  preloadComponents?: string[];
  loadTimeout?: number;
  retryAttempts?: number;
  retryDelay?: number;
}

export interface ComponentLoader {
  (): Promise<any>;
}

export interface ComponentOptions {
  container?: HTMLElement | string;
  eventSystem?: EventSystem;
  theme?: any;
  config?: Record<string, any>;
  [key: string]: any;
}

export interface CacheEntry {
  component: any;
  timestamp: number;
  accessCount: number;
  lastAccessed: number;
  size: number;
}

export interface CacheStatus {
  size: number;
  maxSize: number;
  components: string[];
  memory: {
    used: number;
    available: number;
    entries: Array<{
      name: string;
      size: number;
      age: number;
      accessCount: number;
    }>;
  };
  statistics: {
    hits: number;
    misses: number;
    evictions: number;
    totalLoads: number;
    averageLoadTime: number;
  };
}

export interface LoadingState {
  isLoading: boolean;
  progress: number;
  error: Error | null;
  retryCount: number;
  startTime: number;
}

export declare class LazyLoader {
  constructor(options?: LazyLoaderOptions);
  
  registerComponent(name: string, loader: ComponentLoader): void;
  unregisterComponent(name: string): void;
  
  loadComponent(name: string, options?: ComponentOptions): Promise<any>;
  preloadComponent(name: string): Promise<void>;
  preloadComponents(names: string[]): Promise<void>;
  
  isComponentRegistered(name: string): boolean;
  isComponentLoaded(name: string): boolean;
  isComponentLoading(name: string): boolean;
  
  getLoadingState(name: string): LoadingState | null;
  getRegisteredComponents(): string[];
  getLoadedComponents(): string[];
  
  clearCache(componentName?: string): void;
  evictComponent(componentName: string): void;
  
  getCacheStatus(): CacheStatus;
  optimizeCache(): void;
  
  enableDebug(): void;
  disableDebug(): void;
  
  destroy(): void;
}

// OnlyOffice Integration Types
export interface OnlyOfficeAPI {
  executeMethod(method: string, args?: any[], callback?: Function): Promise<any>;
  callCommand(command: Function, isNoCalc?: boolean): void;
  resizeWindow(width: number, height: number): void;
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

export interface OnlyOfficeAPIGuardOptions {
  retryAttempts?: number;
  retryDelay?: number;
  timeout?: number;
  fallbackMode?: boolean;
  debug?: boolean;
  healthCheckInterval?: number;
  enableHealthCheck?: boolean;
}

export interface APIHealthStatus {
  available: boolean;
  version: string | null;
  lastCheck: number;
  consecutiveFailures: number;
  responseTime: number;
  features: {
    executeMethod: boolean;
    callCommand: boolean;
    resizeWindow: boolean;
    theme: boolean;
  };
}

export declare class OnlyOfficeAPIGuard {
  constructor(options?: OnlyOfficeAPIGuardOptions);
  
  readonly isAvailable: boolean;
  readonly version: string | null;
  readonly healthStatus: APIHealthStatus;
  
  checkAvailability(): Promise<boolean>;
  waitForAvailability(timeout?: number): Promise<boolean>;
  
  safeAPICall<T>(apiCall: () => Promise<T>, fallback?: T): Promise<T>;
  withFallback<T>(operation: () => Promise<T>, fallback: T): Promise<T>;
  
  executeMethod(method: string, args?: any[], callback?: Function): Promise<any>;
  callCommand(command: Function, isNoCalc?: boolean): Promise<void>;
  resizeWindow(width: number, height: number): Promise<void>;
  
  getTheme(): OnlyOfficeTheme | null;
  setTheme(theme: Partial<OnlyOfficeTheme>): Promise<void>;
  
  onThemeChange(callback: (theme: OnlyOfficeTheme) => void): () => void;
  onDocumentStateChange(callback: (data: any) => void): () => void;
  
  startHealthCheck(): void;
  stopHealthCheck(): void;
  
  enableDebug(): void;
  disableDebug(): void;
  
  destroy(): void;
}

// Theme Manager Types
export interface ThemeManagerOptions {
  autoDetect?: boolean;
  syncWithOnlyOffice?: boolean;
  customThemes?: Record<string, Partial<OnlyOfficeTheme>>;
  debug?: boolean;
  sanitization?: {
    enabled?: boolean;
    allowedProperties?: string[];
    maxValueLength?: number;
  };
}

export interface ThemeValidationResult {
  valid: boolean;
  errors: string[];
  sanitized?: OnlyOfficeTheme;
}

export declare class OnlyOfficeThemeManager {
  constructor(options?: ThemeManagerOptions);
  
  getCurrentTheme(): OnlyOfficeTheme | null;
  setTheme(theme: Partial<OnlyOfficeTheme>): Promise<void>;
  
  registerCustomTheme(name: string, theme: Partial<OnlyOfficeTheme>): void;
  getCustomTheme(name: string): Partial<OnlyOfficeTheme> | null;
  getAvailableThemes(): string[];
  
  applyCustomTheme(name: string): Promise<void>;
  resetToDefault(): Promise<void>;
  
  validateTheme(theme: Partial<OnlyOfficeTheme>): ThemeValidationResult;
  sanitizeTheme(theme: Partial<OnlyOfficeTheme>): OnlyOfficeTheme;
  
  onThemeChange(callback: (theme: OnlyOfficeTheme) => void): () => void;
  
  enableAutoDetect(): void;
  disableAutoDetect(): void;
  
  enableDebug(): void;
  disableDebug(): void;
  
  destroy(): void;
}

// Base Component Types
export interface BaseComponentOptions {
  container?: HTMLElement | string;
  eventSystem?: EventSystem;
  theme?: OnlyOfficeTheme;
  config?: Record<string, any>;
  lazyLoad?: boolean;
  telemetry?: any;
  errorHandler?: any;
}

export interface ComponentEvents {
  'component:created': { component: any };
  'component:initialized': { component: any };
  'component:destroyed': { component: any };
  'component:error': { component: any; error: Error };
  'component:config-changed': { component: any; config: any };
  'component:theme-changed': { component: any; theme: OnlyOfficeTheme };
}

export interface ComponentMetadata {
  id: string;
  type: string;
  version: string;
  created: number;
  initialized: boolean;
  destroyed: boolean;
  container: HTMLElement | null;
  config: Record<string, any>;
}

export interface BaseComponent {
  readonly id: string;
  readonly type: string;
  readonly initialized: boolean;
  readonly destroyed: boolean;
  readonly metadata: ComponentMetadata;
  
  initialize(): Promise<void>;
  destroy(): Promise<void>;
  
  on(event: string, handler: EventHandler): () => void;
  once(event: string, handler: EventHandler): () => void;
  emit(event: string, data?: any): void;
  off(event: string, handler?: EventHandler): void;
  
  getConfig(): Record<string, any>;
  setConfig(config: Record<string, any>): void;
  updateConfig(config: Partial<Record<string, any>>): void;
  
  getContainer(): HTMLElement | null;
  setContainer(container: HTMLElement | string): void;
  
  getTheme(): OnlyOfficeTheme | null;
  setTheme(theme: Partial<OnlyOfficeTheme>): void;
  
  getElement(): HTMLElement | null;
  render(): HTMLElement;
  
  show(): void;
  hide(): void;
  isVisible(): boolean;
  
  focus(): void;
  blur(): void;
  isFocused(): boolean;
  
  resize(width?: number, height?: number): void;
  refresh(): void;
  
  validateConfig(config: any): void;
  handleError(error: Error, context?: Record<string, any>): void;
}

// Factory functions
export function createEventSystem(options?: EventSystemOptions): EventSystem;
export function createLazyLoader(options?: LazyLoaderOptions): LazyLoader;
export function createAPIGuard(options?: OnlyOfficeAPIGuardOptions): OnlyOfficeAPIGuard;
export function createThemeManager(options?: ThemeManagerOptions): OnlyOfficeThemeManager;

// Utility functions
export function isEventSystem(obj: any): obj is EventSystem;
export function isLazyLoader(obj: any): obj is LazyLoader;
export function isAPIGuard(obj: any): obj is OnlyOfficeAPIGuard;
export function isThemeManager(obj: any): obj is OnlyOfficeThemeManager;
export function isBaseComponent(obj: any): obj is BaseComponent;

// Constants
export const DEFAULT_EVENT_SYSTEM_OPTIONS: Required<EventSystemOptions>;
export const DEFAULT_LAZY_LOADER_OPTIONS: Required<LazyLoaderOptions>;
export const DEFAULT_API_GUARD_OPTIONS: Required<OnlyOfficeAPIGuardOptions>;
export const DEFAULT_THEME_MANAGER_OPTIONS: Required<ThemeManagerOptions>;

export default {
  EventSystem,
  LazyLoader,
  OnlyOfficeAPIGuard,
  OnlyOfficeThemeManager,
  createEventSystem,
  createLazyLoader,
  createAPIGuard,
  createThemeManager
};