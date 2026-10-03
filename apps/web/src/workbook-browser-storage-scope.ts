export interface WorkbookBrowserStorageScope {
  readonly documentId: string
  readonly userId: string
}

export function scopedWorkbookStorageKey(prefix: string, scope: WorkbookBrowserStorageScope): string {
  return `${prefix}${encodeURIComponent(scope.documentId)}:${encodeURIComponent(scope.userId)}`
}
