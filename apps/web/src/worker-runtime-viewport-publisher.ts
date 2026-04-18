import type { WorkbookLocalViewportBase } from '@bilig/storage-browser'
import {
  ValueTag,
  formulaLooksDateLike,
  isDateLikeHeaderValue,
  isLikelyExcelDateSerialValue,
  type CellSnapshot,
  type CellStyleRecord,
  type EngineEvent,
  type RecalcMetrics,
} from '@bilig/protocol'
import { encodeViewportPatch, type ViewportPatch, type ViewportPatchSubscription } from '@bilig/worker-transport'
import {
  normalizeViewport,
  viewportPatchMayBeImpacted,
  type SheetViewportImpact,
  type ViewportSubscriptionState,
  type WorkerEngine,
} from './worker-runtime-support.js'
import { DEFAULT_STYLE_ID, buildViewportPatchFromEngine, buildViewportPatchFromLocalBase } from './worker-runtime-viewport.js'

function hasFormulaErrorCells(base: WorkbookLocalViewportBase): boolean {
  return base.cells.some((cell) => cell.snapshot.formula !== undefined && cell.snapshot.value.tag === ValueTag.Error)
}

function hasUnformattedDateSerialCells(base: WorkbookLocalViewportBase): boolean {
  const cellsByPosition = new Map(base.cells.map((cell) => [`${cell.row}:${cell.col}`, cell.snapshot]))
  return base.cells.some((cell) => {
    if (cell.snapshot.format !== undefined || !isLikelyExcelDateSerialValue(cell.snapshot.value)) {
      return false
    }
    if (formulaLooksDateLike(cell.snapshot.formula)) {
      return true
    }
    const header = cellsByPosition.get(`${cell.row - 1}:${cell.col}`)
    return header !== undefined && isDateLikeHeaderValue(header.value)
  })
}

export function createEmptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  }
}

export class WorkerViewportPatchPublisher {
  private readonly viewportSubscriptions = new Set<ViewportSubscriptionState>()
  private readonly viewportSubscriptionsBySheet = new Map<string, Set<ViewportSubscriptionState>>()
  private readonly formatIds = new Map<string, number>([['', 0]])
  private readonly styles = new Map<string, CellStyleRecord>([[DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }]])
  private nextFormatId = 1

  constructor(
    private readonly options: {
      buildPatch: (
        state: ViewportSubscriptionState,
        event: EngineEvent | null,
        metrics: RecalcMetrics,
        sheetImpact: SheetViewportImpact | null,
      ) => ViewportPatch
      canReadLocalProjectionForViewport: () => boolean
      getCurrentMetrics: () => RecalcMetrics
      getProjectionEngine: () => WorkerEngine
      hasProjectionEngine: () => boolean
      readLocalViewport: (sheetName: string, viewport: ViewportPatchSubscription) => WorkbookLocalViewportBase | null
      scheduleProjectionEngineMaterialization: () => void
    },
  ) {}

  reset(): void {
    this.viewportSubscriptions.clear()
    this.viewportSubscriptionsBySheet.clear()
    this.formatIds.clear()
    this.formatIds.set('', 0)
    this.styles.clear()
    this.styles.set(DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID })
    this.nextFormatId = 1
  }

  subscribe(subscription: ViewportPatchSubscription, listener: (patch: Uint8Array) => void): () => void {
    const state: ViewportSubscriptionState = {
      subscription: normalizeViewport(subscription),
      listener,
      nextVersion: 1,
      knownStyleIds: new Set(),
      lastStyleSignatures: new Map<string, string>(),
      lastCellSignatures: new Map<string, string>(),
      lastColumnSignatures: new Map<number, string>(),
      lastRowSignatures: new Map<number, string>(),
    }
    listener(encodeViewportPatch(this.options.buildPatch(state, null, this.options.getCurrentMetrics(), null)))
    if (!this.options.hasProjectionEngine() && this.options.canReadLocalProjectionForViewport()) {
      this.options.scheduleProjectionEngineMaterialization()
    }
    this.viewportSubscriptions.add(state)
    this.addViewportSubscription(state)
    return () => {
      this.viewportSubscriptions.delete(state)
      this.removeViewportSubscription(state)
    }
  }

  buildPatch(
    state: ViewportSubscriptionState,
    event: EngineEvent | null,
    metrics: RecalcMetrics = this.options.getCurrentMetrics(),
    sheetImpact: SheetViewportImpact | null = null,
  ): ViewportPatch {
    if ((event === null || event.invalidation === 'full') && this.options.canReadLocalProjectionForViewport()) {
      const localBase = this.options.readLocalViewport(state.subscription.sheetName, state.subscription)
      if (
        localBase &&
        (!this.options.hasProjectionEngine() || (!hasFormulaErrorCells(localBase) && !hasUnformattedDateSerialCells(localBase)))
      ) {
        return buildViewportPatchFromLocalBase({
          state,
          metrics,
          base: localBase,
          getFormatId: (format) => this.getFormatId(format),
        })
      }
    }

    return buildViewportPatchFromEngine({
      state,
      event,
      metrics,
      sheetImpact,
      engine: this.options.getProjectionEngine(),
      emptyCellSnapshot: (sheetName, address) => createEmptyCellSnapshot(sheetName, address),
      getStyleRecord: (styleId) => this.getStyleRecord(styleId),
      getFormatId: (format) => this.getFormatId(format),
    })
  }

  broadcast(input: {
    event: EngineEvent | null
    impactsBySheet: ReadonlyMap<string, SheetViewportImpact> | null
    metrics?: RecalcMetrics
  }): void {
    const metrics = input.metrics ?? this.options.getCurrentMetrics()
    const impactedSheets = input.impactsBySheet === null ? null : new Set(input.impactsBySheet.keys())
    for (const subscription of this.getViewportSubscriptionsForEvent(input.event, impactedSheets)) {
      const sheetImpact = input.impactsBySheet?.get(subscription.subscription.sheetName) ?? null
      if (input.event !== null && !viewportPatchMayBeImpacted(subscription.subscription, input.event, sheetImpact, impactedSheets)) {
        continue
      }
      const patch = this.options.buildPatch(subscription, input.event, metrics, sheetImpact)
      if (patch.cells.length === 0 && patch.columns.length === 0 && patch.rows.length === 0) {
        continue
      }
      subscription.listener(encodeViewportPatch(patch))
    }
  }

  private addViewportSubscription(state: ViewportSubscriptionState): void {
    const subscriptions = this.viewportSubscriptionsBySheet.get(state.subscription.sheetName) ?? new Set()
    subscriptions.add(state)
    this.viewportSubscriptionsBySheet.set(state.subscription.sheetName, subscriptions)
  }

  private removeViewportSubscription(state: ViewportSubscriptionState): void {
    const subscriptions = this.viewportSubscriptionsBySheet.get(state.subscription.sheetName)
    if (!subscriptions) {
      return
    }
    subscriptions.delete(state)
    if (subscriptions.size === 0) {
      this.viewportSubscriptionsBySheet.delete(state.subscription.sheetName)
    }
  }

  private getViewportSubscriptionsForEvent(
    event: EngineEvent | null,
    impactedSheets: ReadonlySet<string> | null,
  ): Iterable<ViewportSubscriptionState> {
    if (event === null || event.invalidation === 'full' || impactedSheets === null) {
      return this.viewportSubscriptions
    }

    const subscriptions = new Set<ViewportSubscriptionState>()
    impactedSheets.forEach((sheetName) => {
      this.viewportSubscriptionsBySheet.get(sheetName)?.forEach((subscription) => {
        subscriptions.add(subscription)
      })
    })
    return subscriptions
  }

  private getStyleRecord(styleId: string): CellStyleRecord {
    const existing = this.styles.get(styleId)
    if (existing) {
      return existing
    }
    const resolved = this.options.getProjectionEngine().getCellStyle(styleId) ?? {
      id: DEFAULT_STYLE_ID,
    }
    this.styles.set(resolved.id, resolved)
    return resolved
  }

  private getFormatId(format: string | undefined): number {
    const key = format ?? ''
    const existing = this.formatIds.get(key)
    if (existing !== undefined) {
      return existing
    }
    const nextId = this.nextFormatId++
    this.formatIds.set(key, nextId)
    return nextId
  }
}
