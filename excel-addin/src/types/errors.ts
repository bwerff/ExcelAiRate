/**
 * Custom error types for Excel Add-in
 */

export class ExcelAIError extends Error {
  constructor(
    message: string,
    public code: string,
    public details?: unknown
  ) {
    super(message);
    this.name = 'ExcelAIError';
  }
}

export class AuthenticationError extends ExcelAIError {
  constructor(message: string, details?: unknown) {
    super(message, 'AUTH_ERROR', details);
    this.name = 'AuthenticationError';
  }
}

export class APIError extends ExcelAIError {
  constructor(message: string, public statusCode?: number, details?: unknown) {
    super(message, 'API_ERROR', details);
    this.name = 'APIError';
  }
}

export class ValidationError extends ExcelAIError {
  constructor(message: string, details?: unknown) {
    super(message, 'VALIDATION_ERROR', details);
    this.name = 'ValidationError';
  }
}

export class ExcelOperationError extends ExcelAIError {
  constructor(message: string, details?: unknown) {
    super(message, 'EXCEL_OPERATION_ERROR', details);
    this.name = 'ExcelOperationError';
  }
}

export class NetworkError extends ExcelAIError {
  constructor(message: string, details?: unknown) {
    super(message, 'NETWORK_ERROR', details);
    this.name = 'NetworkError';
  }
}

export class UsageLimitError extends ExcelAIError {
  constructor(message: string, public currentUsage: number, public limit: number) {
    super(message, 'USAGE_LIMIT_ERROR', { currentUsage, limit });
    this.name = 'UsageLimitError';
  }
}

export function isExcelAIError(error: unknown): error is ExcelAIError {
  return error instanceof ExcelAIError;
}

export function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  if (typeof error === 'string') {
    return error;
  }
  return 'An unknown error occurred';
}