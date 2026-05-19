import type { SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { detachTrackedIndexChanges, hasDeferredTrackedIndexChanges } from './tracked-cell-index-changes.js'
import { captureTrackedEngineEvent, type CoreTrackedEngineEvent, type TrackedEngineEvent } from './tracked-engine-event-refs.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import type { WorkPaperCellChange } from './work-paper-types.js'

interface EngineTrackedEventSubscription {
  subscribeTracked(listener: (event: CoreTrackedEngineEvent) => void): () => void
}

function hasTrackedEngineSubscription(value: unknown): value is EngineTrackedEventSubscription {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'subscribeTracked') === 'function'
}

export class WorkPaperEngineEventTracker {
  private trackedEvents: TrackedEngineEvent[] = []
  private pendingLazyTrackedChanges: WorkPaperCellChange[][] = []
  private captureEnabled = true
  private retainedEventIndicesDepth = 0
  private unsubscribe: (() => void) | null = null

  get hasTrackedEvents(): boolean {
    return this.trackedEvents.length > 0
  }

  get hasPendingLazyChanges(): boolean {
    return this.pendingLazyTrackedChanges.length > 0
  }

  get isCaptureEnabled(): boolean {
    return this.captureEnabled
  }

  setCaptureEnabled(enabled: boolean): void {
    this.captureEnabled = enabled
  }

  attach(engine: SpreadsheetEngine): void {
    this.unsubscribe?.()
    this.trackedEvents = []
    if (!hasTrackedEngineSubscription(engine.events)) {
      throw new WorkPaperOperationError('Tracked engine event subscription is unavailable')
    }
    this.unsubscribe = engine.events.subscribeTracked((event) => {
      if (!this.captureEnabled) {
        return
      }
      this.trackedEvents.push(
        captureTrackedEngineEvent(event, {
          borrowChangedCellIndexViews: this.retainedEventIndicesDepth > 0,
          cloneChangedCellIndices: this.retainedEventIndicesDepth === 0,
        }),
      )
    })
  }

  detach(): void {
    this.unsubscribe?.()
    this.unsubscribe = null
    this.clearEvents()
  }

  dispose(): void {
    this.materializePendingLazyChanges()
    this.detach()
  }

  clearEvents(): void {
    this.trackedEvents = []
  }

  drain(): TrackedEngineEvent[] {
    const events = this.trackedEvents
    this.trackedEvents = []
    return events
  }

  withRetainedIndices<T>(callback: () => T): T {
    this.retainedEventIndicesDepth += 1
    try {
      return callback()
    } finally {
      this.retainedEventIndicesDepth -= 1
    }
  }

  withCaptureDisabled<T>(callback: () => T): T {
    const previous = this.captureEnabled
    this.captureEnabled = false
    this.trackedEvents = []
    try {
      return callback()
    } finally {
      this.captureEnabled = previous
      this.trackedEvents = []
    }
  }

  trackLazyChanges(changes: WorkPaperCellChange[]): void {
    if (hasDeferredTrackedIndexChanges(changes)) {
      this.pendingLazyTrackedChanges.push(changes)
    }
  }

  materializePendingLazyChanges(options: { readonly preservePositions?: boolean } = {}): void {
    if (this.pendingLazyTrackedChanges.length === 0) {
      return
    }
    const pending = this.pendingLazyTrackedChanges
    this.pendingLazyTrackedChanges = []
    for (let index = 0; index < pending.length; index += 1) {
      detachTrackedIndexChanges(
        pending[index]!,
        options.preservePositions === undefined ? {} : { preservePositions: options.preservePositions },
      )
    }
  }
}
