import { queries } from '@bilig/zero-sync'

interface ZeroWorkbookSyncSourceLike {
  materialize(query: unknown): unknown
}

interface LiveView<T> {
  readonly data: T
  addListener(listener: (value: T) => void): () => void
  destroy(): void
}

export interface WorkbookRevisionState {
  headRevision: number
  calculatedRevision: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeRevision(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isLiveView<T>(value: unknown): value is LiveView<T> {
  return isRecord(value) && 'data' in value && typeof value['addListener'] === 'function' && typeof value['destroy'] === 'function'
}

function assertLiveView<T>(value: unknown): LiveView<T> {
  if (!isLiveView<T>(value)) {
    throw new Error('Zero workbook sync source returned an invalid workbook revision view')
  }
  return value
}

function normalizeWorkbookRevisionState(value: unknown): WorkbookRevisionState | null {
  if (!isRecord(value) || !isSafeNonNegativeRevision(value['headRevision']) || !isSafeNonNegativeRevision(value['calculatedRevision'])) {
    return null
  }
  if (value['calculatedRevision'] > value['headRevision']) {
    return null
  }
  return {
    headRevision: value['headRevision'],
    calculatedRevision: value['calculatedRevision'],
  }
}

export class ZeroWorkbookRevisionSync {
  private readonly workbookView: LiveView<unknown>
  private readonly cleanup: (() => void)[]

  constructor(input: {
    zero: ZeroWorkbookSyncSourceLike
    documentId: string
    onRevisionState?: (revisionState: WorkbookRevisionState | null) => void
  }) {
    this.workbookView = assertLiveView<unknown>(input.zero.materialize(queries.workbook.get({ documentId: input.documentId })))
    const publishRevisionState = (value: unknown) => {
      input.onRevisionState?.(normalizeWorkbookRevisionState(value))
    }
    publishRevisionState(this.workbookView.data)
    this.cleanup = [
      this.workbookView.addListener((value) => {
        publishRevisionState(value)
      }),
    ]
  }

  dispose(): void {
    this.cleanup.forEach((cleanup) => cleanup())
    this.workbookView.destroy()
  }
}
