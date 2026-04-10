class WorkbookAgentServiceError extends Error {
  readonly code: string;
  readonly statusCode: number;
  readonly retryable: boolean;

  constructor(input: { code: string; message: string; statusCode: number; retryable: boolean }) {
    super(input.message);
    this.name = "WorkbookAgentServiceError";
    this.code = input.code;
    this.statusCode = input.statusCode;
    this.retryable = input.retryable;
  }
}

export function isWorkbookAgentServiceError(value: unknown): value is Error & {
  readonly code: string;
  readonly statusCode: number;
  readonly retryable: boolean;
} {
  return value instanceof WorkbookAgentServiceError;
}

export function createWorkbookAgentServiceError(input: {
  code: string;
  message: string;
  statusCode: number;
  retryable: boolean;
}): Error & {
  readonly code: string;
  readonly statusCode: number;
  readonly retryable: boolean;
} {
  return new WorkbookAgentServiceError(input);
}
