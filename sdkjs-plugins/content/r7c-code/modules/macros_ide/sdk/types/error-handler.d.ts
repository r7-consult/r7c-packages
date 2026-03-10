/**
 * @fileoverview Error Handler Type Definitions
 * @description Comprehensive error handling and validation types
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

export interface ErrorHandlerConfig {
  enabled?: boolean;
  logErrors?: boolean;
  throwOnCritical?: boolean;
  customHandler?: (error: Error, context: Record<string, any>) => void;
  maxStackTraceDepth?: number;
  errorCodes?: Record<string, string>;
}

export interface ValidationSchema {
  [field: string]: ValidationRule;
}

export interface ValidationRule {
  type: 'string' | 'number' | 'boolean' | 'object' | 'array' | 'function';
  required?: boolean;
  validator?: (value: any) => boolean;
  transform?: (value: any) => any;
  min?: number;
  max?: number;
  pattern?: RegExp;
  enum?: any[];
  properties?: ValidationSchema; // For object types
  items?: ValidationRule; // For array types
}

export interface RecoveryStrategy {
  canRecover: boolean;
  action: () => Promise<void>;
  fallback?: () => any;
  retryCount?: number;
  maxRetries?: number;
}

export interface ErrorContext {
  component?: string;
  operation?: string;
  user?: string;
  timestamp?: number;
  sessionId?: string;
  stack?: string;
  data?: Record<string, any>;
}

export declare class SDKError extends Error {
  readonly code: string;
  readonly context: Record<string, any>;
  readonly timestamp: number;
  readonly recoverable: boolean;
  readonly severity: 'low' | 'medium' | 'high' | 'critical';
  
  constructor(
    message: string, 
    code?: string, 
    context?: Record<string, any>, 
    recoverable?: boolean
  );
  
  toJSON(): {
    name: string;
    message: string;
    code: string;
    context: Record<string, any>;
    timestamp: number;
    recoverable: boolean;
    severity: string;
    stack?: string;
  };
}

export declare class ComponentError extends SDKError {
  readonly component: string;
  
  constructor(
    component: string, 
    message: string, 
    context?: Record<string, any>
  );
}

export declare class ConfigError extends SDKError {
  readonly field: string;
  readonly expected: string;
  readonly actual: any;
  
  constructor(
    context: string, 
    message: string, 
    config?: any
  );
}

export declare class APIError extends SDKError {
  readonly method: string;
  readonly endpoint?: string;
  readonly statusCode?: number;
  readonly response?: any;
  
  constructor(
    method: string, 
    message: string, 
    context?: Record<string, any>
  );
}

export declare class ValidationError extends SDKError {
  readonly field: string;
  readonly rule: ValidationRule;
  readonly value: any;
  
  constructor(
    field: string, 
    rule: ValidationRule, 
    value: any, 
    context?: string
  );
}

export declare class SecurityError extends SDKError {
  readonly securityLevel: 'low' | 'medium' | 'high' | 'critical';
  readonly attackVector?: string;
  
  constructor(
    message: string, 
    securityLevel: 'low' | 'medium' | 'high' | 'critical', 
    context?: Record<string, any>
  );
}

export declare class NetworkError extends SDKError {
  readonly url?: string;
  readonly status?: number;
  readonly timeout?: boolean;
  
  constructor(
    message: string, 
    context?: Record<string, any>
  );
}

export interface ErrorHandlerOptions {
  config?: ErrorHandlerConfig;
  recoveryStrategies?: Map<string, (error: Error) => RecoveryStrategy>;
  customValidators?: Map<string, (value: any) => boolean>;
}

export declare class ErrorHandler {
  constructor(options?: ErrorHandlerOptions);
  
  handleError(error: Error, context?: Record<string, any>): void;
  
  validateConfig(
    config: any, 
    schema: ValidationSchema, 
    context?: string
  ): void;
  
  validateField(
    config: any, 
    field: string, 
    rules: ValidationRule, 
    context?: string
  ): any;
  
  createRecoveryStrategy(error: Error): RecoveryStrategy | null;
  
  addRecoveryStrategy(
    errorType: string, 
    strategy: (error: Error) => RecoveryStrategy
  ): void;
  
  addCustomValidator(
    name: string, 
    validator: (value: any) => boolean
  ): void;
  
  formatError(error: Error, context?: Record<string, any>): string;
  
  isRecoverable(error: Error): boolean;
  isCritical(error: Error): boolean;
  
  logError(error: Error, context?: Record<string, any>): void;
  
  sanitizeStack(stack: string, maxDepth?: number): string;
  
  getErrorCode(error: Error): string;
  getErrorSeverity(error: Error): 'low' | 'medium' | 'high' | 'critical';
  
  createError(
    type: 'sdk' | 'component' | 'config' | 'api' | 'validation' | 'security' | 'network',
    message: string,
    context?: Record<string, any>
  ): SDKError;
}

export interface ErrorReporter {
  report(error: Error, context?: Record<string, any>): Promise<void>;
}

export interface ErrorMetrics {
  totalErrors: number;
  errorsByType: Record<string, number>;
  errorsBySeverity: Record<string, number>;
  recoveryAttempts: number;
  successfulRecoveries: number;
  recentErrors: Array<{
    error: SDKError;
    timestamp: number;
    recovered: boolean;
  }>;
}

export interface ErrorHandlerEvents {
  'error:occurred': { error: Error; context: Record<string, any> };
  'error:recovered': { error: Error; strategy: RecoveryStrategy };
  'error:critical': { error: Error; context: Record<string, any> };
  'validation:failed': { field: string; value: any; rule: ValidationRule };
  'recovery:attempted': { error: Error; strategy: RecoveryStrategy };
  'recovery:succeeded': { error: Error; strategy: RecoveryStrategy };
  'recovery:failed': { error: Error; strategy: RecoveryStrategy; reason: string };
}

// Type guards
export function isSDKError(error: any): error is SDKError;
export function isComponentError(error: any): error is ComponentError;
export function isConfigError(error: any): error is ConfigError;
export function isAPIError(error: any): error is APIError;
export function isValidationError(error: any): error is ValidationError;
export function isSecurityError(error: any): error is SecurityError;
export function isNetworkError(error: any): error is NetworkError;

// Utility functions
export function createErrorHandler(options?: ErrorHandlerOptions): ErrorHandler;
export function createValidationSchema(schema: Record<string, Partial<ValidationRule>>): ValidationSchema;
export function createRecoveryStrategy(
  canRecover: boolean,
  action: () => Promise<void>,
  fallback?: () => any
): RecoveryStrategy;

// Error creation shortcuts
export function createSDKError(
  message: string, 
  code?: string, 
  context?: Record<string, any>
): SDKError;

export function createComponentError(
  component: string, 
  message: string, 
  context?: Record<string, any>
): ComponentError;

export function createConfigError(
  context: string, 
  message: string, 
  config?: any
): ConfigError;

export function createAPIError(
  method: string, 
  message: string, 
  context?: Record<string, any>
): APIError;

export function createValidationError(
  field: string, 
  rule: ValidationRule, 
  value: any, 
  context?: string
): ValidationError;

export function createSecurityError(
  message: string, 
  securityLevel: 'low' | 'medium' | 'high' | 'critical', 
  context?: Record<string, any>
): SecurityError;

export function createNetworkError(
  message: string, 
  context?: Record<string, any>
): NetworkError;

export default ErrorHandler;