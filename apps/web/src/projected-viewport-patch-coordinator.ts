import type { Viewport } from '@bilig/protocol'
import { decodeViewportPatch, type ViewportPatch, type WorkerEngineClient } from '@bilig/worker-transport'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import type { ProjectedViewportCellCache } from './projected-viewport-cell-cache.js'
import type { ProjectedViewportAxisStore } from './projected-viewport-axis-store.js'
import { applyProjectedViewportPatch, type ProjectedViewportPatchApplicationResult } from './projected-viewport-patch-application.js'

type CellItem = readonly [number, number]
export type ProjectedViewportPatchApplied = Pick<
  ProjectedViewportPatchApplicationResult,
  'damage' | 'axisChanged' | 'columnsChanged' | 'rowsChanged' | 'freezeChanged'
>
interface ProjectedViewportSubscriptionOptions {
  readonly initialPatch?: 'full' | 'none'
}

export class ProjectedViewportPatchCoordinator {
  constructor(
    private readonly options: {
      client?: WorkerEngineClient
      cellCache: ProjectedViewportCellCache
      axisStore: ProjectedViewportAxisStore
      onViewportPatchApplied?:
        | ((patch: ViewportPatch, result: ProjectedViewportPatchApplied, options: ProjectedViewportSubscriptionOptions) => void)
        | undefined
    },
  ) {}

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: CellItem }[]) => void,
    options: ProjectedViewportSubscriptionOptions = {},
  ): () => void {
    if (!this.options.client) {
      throw new Error('Worker viewport subscriptions require a worker engine client')
    }
    const stopTrackingViewport = this.options.cellCache.trackViewport(sheetName, viewport)
    let frameHandle: number | null = null
    let timeoutHandle: ReturnType<typeof setTimeout> | null = null
    const pendingDamage = new Map<string, { cell: CellItem }>()
    let pendingStructuralChange = false
    let lastAppliedPatchVersion = 0
    const scheduleFlush = () => {
      if (frameHandle !== null || timeoutHandle !== null) {
        return
      }
      const flush = () => {
        frameHandle = null
        timeoutHandle = null
        if (!pendingStructuralChange && pendingDamage.size === 0) {
          return
        }
        const damage = pendingDamage.size === 0 ? undefined : [...pendingDamage.values()]
        pendingDamage.clear()
        pendingStructuralChange = false
        listener(damage)
      }
      if (typeof window === 'undefined') {
        timeoutHandle = setTimeout(flush, 0)
        return
      }
      frameHandle = window.requestAnimationFrame(flush)
    }
    const applyPatchBytes = (bytes: Uint8Array) => {
      const patch = decodeViewportPatch(bytes)
      if (patch.version <= lastAppliedPatchVersion) {
        return
      }
      lastAppliedPatchVersion = patch.version
      const result = this.applyViewportPatchDetailed(patch)
      this.options.onViewportPatchApplied?.(patch, result, options)
      for (const entry of result.damage) {
        pendingDamage.set(`${entry.cell[0]}:${entry.cell[1]}`, entry)
      }
      pendingStructuralChange = pendingStructuralChange || result.axisChanged || result.freezeChanged
    }
    const subscription =
      options.initialPatch === 'none' ? ({ sheetName, ...viewport, initialPatch: 'none' } as const) : ({ sheetName, ...viewport } as const)
    const unsubscribe = this.options.client.subscribeViewportPatches(subscription, (bytes: Uint8Array) => {
      applyPatchBytes(bytes)
      scheduleFlush()
    })
    return () => {
      if (frameHandle !== null) {
        window.cancelAnimationFrame(frameHandle)
        frameHandle = null
      }
      if (timeoutHandle !== null) {
        clearTimeout(timeoutHandle)
        timeoutHandle = null
      }
      unsubscribe()
      stopTrackingViewport()
    }
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    return this.applyViewportPatchDetailed(patch).damage
  }

  applyViewportPatchDetailed(patch: ViewportPatch): ProjectedViewportPatchApplied {
    const result = applyProjectedViewportPatch({
      state: {
        ...this.options.cellCache.getPatchState(),
        ...this.options.axisStore.getPatchState(),
      },
      patch,
      touchCellKey: (key) => this.options.cellCache.touchCellKey(key),
    })
    getWorkbookScrollPerfCollector()?.noteViewportPatch({
      full: patch.full,
      damageCount: result.damage.length,
    })
    return this.options.cellCache.applyPatchResult(patch.viewport.sheetName, result)
  }
}
