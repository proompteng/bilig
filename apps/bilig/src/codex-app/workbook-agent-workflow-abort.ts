function createWorkflowAbortError(): Error {
  const error = new Error('Workbook agent workflow cancelled.')
  error.name = 'AbortError'
  return error
}

export function isWorkflowAbortError(error: unknown): boolean {
  return error instanceof Error && error.name === 'AbortError'
}

export function throwIfWorkflowCancelled(signal?: AbortSignal): void {
  if (!signal?.aborted) {
    return
  }
  if (signal.reason instanceof Error) {
    throw signal.reason
  }
  throw createWorkflowAbortError()
}
