import type { Viewport } from '@bilig/protocol'
import { decodeViewportPatch, type ViewportPatch, type WorkerEngineClient } from '@bilig/worker-transport'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import type { ProjectedViewportCellCache } from './projected-viewport-cell-cache.js'
import type { ProjectedViewportAxisStore } from './projected-viewport-axis-store.js'
import { applyProjectedViewportPatch, type ProjectedViewportPatchApplicationResult } from './projected-viewport-patch-application.js'

type CellItem = readonly [number, number]
const ACTIVE_SCROLL_PATCH_DEFER_MS = 96

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
    let frameHandle: number | null = null
    let timeoutHandle: ReturnType<typeof setTimeout> | null = null
    const pendingDamage = new Map<string, { cell: CellItem }>()
    const pendingPatchBytes: Uint8Array[] = []
    let pendingStructuralChange = false
    let patchTimeoutHandle: ReturnType<typeof setTimeout> | null = null
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
      const result = this.applyViewportPatchDetailed(decodeViewportPatch(bytes))
      for (const entry of result.damage) {
        pendingDamage.set(`${entry.cell[0]}:${entry.cell[1]}`, entry)
      }
      pendingStructuralChange = pendingStructuralChange || result.axisChanged || result.freezeChanged
    }
    const schedulePatchApply = () => {
      if (patchTimeoutHandle !== null) {
        return
      }
      const collector = getWorkbookScrollPerfCollector()
      const delay = collector?.isGridScrollRecentlyActive(ACTIVE_SCROLL_PATCH_DEFER_MS) ? ACTIVE_SCROLL_PATCH_DEFER_MS : 0
      patchTimeoutHandle = setTimeout(() => {
        patchTimeoutHandle = null
        if (collector?.isGridScrollRecentlyActive(ACTIVE_SCROLL_PATCH_DEFER_MS) && pendingPatchBytes.length > 0) {
          schedulePatchApply()
          return
        }
        const patches = pendingPatchBytes.splice(0)
        for (const patchBytes of patches) {
          applyPatchBytes(patchBytes)
        }
        scheduleFlush()
      }, delay)
    }
    const unsubscribe = this.options.client.subscribeViewportPatches({ sheetName, ...viewport }, (bytes: Uint8Array) => {
      pendingPatchBytes.push(bytes)
      schedulePatchApply()
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
      if (patchTimeoutHandle !== null) {
        clearTimeout(patchTimeoutHandle)
        patchTimeoutHandle = null
      }
      unsubscribe()
      stopTrackingViewport()
    }
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    return this.applyViewportPatchDetailed(patch).damage
  }

  applyViewportPatchDetailed(
    patch: ViewportPatch,
  ): Pick<ProjectedViewportPatchApplicationResult, 'damage' | 'axisChanged' | 'columnsChanged' | 'rowsChanged' | 'freezeChanged'> {
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
