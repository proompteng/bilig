import type { Viewport } from '@bilig/protocol'
import { decodeViewportPatch, type ViewportPatch, type WorkerEngineClient } from '@bilig/worker-transport'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import type { ProjectedViewportCellCache } from './projected-viewport-cell-cache.js'
import type { ProjectedViewportAxisStore } from './projected-viewport-axis-store.js'
import { applyProjectedViewportPatch } from './projected-viewport-patch-application.js'

type CellItem = readonly [number, number]

export class ProjectedViewportPatchCoordinator {
  constructor(
    private readonly options: {
      client?: WorkerEngineClient
      cellCache: ProjectedViewportCellCache
      axisStore: ProjectedViewportAxisStore
    },
  ) {}

  subscribeViewport(sheetName: string, viewport: Viewport, listener: (damage?: readonly { cell: CellItem }[]) => void): () => void {
    if (!this.options.client) {
      throw new Error('Worker viewport subscriptions require a worker engine client')
    }
    const stopTrackingViewport = this.options.cellCache.trackViewport(sheetName, viewport)
    const unsubscribe = this.options.client.subscribeViewportPatches({ sheetName, ...viewport }, (bytes: Uint8Array) => {
      listener(this.applyViewportPatch(decodeViewportPatch(bytes)))
    })
    return () => {
      unsubscribe()
      stopTrackingViewport()
    }
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
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
