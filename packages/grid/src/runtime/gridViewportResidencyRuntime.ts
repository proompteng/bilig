import { formatAddress } from '@bilig/formula'
import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { VisibleRegionState } from '../gridPointer.js'
import type { Item, Rectangle } from '../gridTypes.js'
import { collectViewportItems } from '../gridViewportItems.js'
import { sameViewportBounds } from '../gridViewportController.js'
import { viewportFromVisibleRegion } from '../useGridCameraState.js'
import { resolveResidentViewport } from '../workbookGridViewport.js'

export interface GridResidentHeaderRegion {
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  readonly tx: number
  readonly ty: number
  readonly freezeRows: number
  readonly freezeCols: number
}

export interface GridViewportResidencyState {
  readonly viewport: Viewport
  readonly residentViewport: Viewport
  readonly renderTileViewport: Viewport
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: GridResidentHeaderRegion
  readonly sceneRevision: number
  readonly visibleAddresses: readonly string[]
  readonly visibleItems: readonly Item[]
}

export interface GridViewportResidencyRuntimeInput {
  readonly freezeCols: number
  readonly freezeRows: number
  readonly visibleRegion: VisibleRegionState
}

export interface GridViewportResidencyInvalidationInput {
  readonly engine: GridEngineLike
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly visibleAddresses: readonly string[]
}

interface RuntimeConnection<Identity> {
  readonly identity: Identity
  readonly unsubscribe: (() => void) | undefined
}

interface LocalSceneInvalidationConnectionIdentity {
  readonly engine: GridEngineLike
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly visibleAddresses: readonly string[]
}

interface GridViewportResidentCache {
  readonly freezeCols: number
  readonly freezeRows: number
  readonly renderTileViewport: Viewport
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: GridResidentHeaderRegion
  readonly residentViewport: Viewport
  readonly visibleAddresses: readonly string[]
  readonly visibleItems: readonly Item[]
}

export class GridViewportResidencyRuntime {
  private residentCache: GridViewportResidentCache | null = null
  private residentViewport: Viewport | null = null
  private localSceneInvalidationConnection: RuntimeConnection<LocalSceneInvalidationConnectionIdentity> | null = null
  private readonly sceneRevisionListeners = new Set<() => void>()
  private sceneRevision = 0

  resolve(input: GridViewportResidencyRuntimeInput): GridViewportResidencyState {
    const viewport = viewportFromVisibleRegion(input.visibleRegion)
    const nextResidentViewport = resolveResidentViewport(viewport)
    if (!this.residentViewport || !sameViewportBounds(this.residentViewport, nextResidentViewport)) {
      this.residentViewport = nextResidentViewport
    }
    const residentViewport = this.residentViewport
    const residentCache = this.resolveResidentCache(input, residentViewport)

    return {
      renderTileViewport: residentCache.renderTileViewport,
      residentHeaderItems: residentCache.residentHeaderItems,
      residentHeaderRegion: residentCache.residentHeaderRegion,
      residentViewport,
      sceneRevision: this.sceneRevision,
      viewport,
      visibleAddresses: residentCache.visibleAddresses,
      visibleItems: residentCache.visibleItems,
    }
  }

  invalidateScene(): number {
    this.sceneRevision += 1
    this.emitSceneRevision()
    return this.sceneRevision
  }

  snapshotSceneRevision(): number {
    return this.sceneRevision
  }

  subscribeSceneRevision(listener: () => void): () => void {
    this.sceneRevisionListeners.add(listener)
    return () => {
      this.sceneRevisionListeners.delete(listener)
    }
  }

  connectLocalSceneInvalidation(input: GridViewportResidencyInvalidationInput, listener?: () => void): (() => void) | undefined {
    if (input.shouldUseRemoteRenderTileSource || input.visibleAddresses.length === 0) {
      return undefined
    }
    const invalidate = () => {
      this.invalidateScene()
      listener?.()
    }
    const unsubscribeCells = input.engine.subscribeCells(input.sheetName, input.visibleAddresses, invalidate)
    const unsubscribeMerges = input.engine.subscribeSheetChannel?.(input.sheetName, 'merges', invalidate)
    if (!unsubscribeMerges) {
      return unsubscribeCells
    }
    return () => {
      unsubscribeCells()
      unsubscribeMerges()
    }
  }

  syncLocalSceneInvalidation(input: GridViewportResidencyInvalidationInput): void {
    const identity: LocalSceneInvalidationConnectionIdentity = {
      engine: input.engine,
      sheetName: input.sheetName,
      shouldUseRemoteRenderTileSource: input.shouldUseRemoteRenderTileSource,
      visibleAddresses: input.visibleAddresses,
    }
    if (
      this.localSceneInvalidationConnection &&
      sameLocalSceneInvalidationConnectionIdentity(this.localSceneInvalidationConnection.identity, identity)
    ) {
      return
    }
    this.localSceneInvalidationConnection?.unsubscribe?.()
    this.localSceneInvalidationConnection = {
      identity,
      unsubscribe: this.connectLocalSceneInvalidation(input),
    }
  }

  disconnectLocalSceneInvalidation(): void {
    this.localSceneInvalidationConnection?.unsubscribe?.()
    this.localSceneInvalidationConnection = null
  }

  private resolveResidentCache(input: GridViewportResidencyRuntimeInput, residentViewport: Viewport): GridViewportResidentCache {
    const current = this.residentCache
    if (
      current?.residentViewport === residentViewport &&
      current.freezeCols === input.freezeCols &&
      current.freezeRows === input.freezeRows
    ) {
      return current
    }

    const visibleItems = collectViewportItems(residentViewport, {
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
    })
    const next: GridViewportResidentCache = {
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      renderTileViewport: {
        rowStart: input.freezeRows > 0 ? 0 : residentViewport.rowStart,
        rowEnd: residentViewport.rowEnd,
        colStart: input.freezeCols > 0 ? 0 : residentViewport.colStart,
        colEnd: residentViewport.colEnd,
      },
      residentHeaderItems: visibleItems,
      residentHeaderRegion: {
        range: {
          x: residentViewport.colStart,
          y: residentViewport.rowStart,
          width: residentViewport.colEnd - residentViewport.colStart + 1,
          height: residentViewport.rowEnd - residentViewport.rowStart + 1,
        },
        tx: 0,
        ty: 0,
        freezeRows: input.freezeRows,
        freezeCols: input.freezeCols,
      },
      residentViewport,
      visibleAddresses: visibleItems.map(([col, row]) => formatAddress(row, col)),
      visibleItems,
    }
    this.residentCache = next
    return next
  }

  private emitSceneRevision(): void {
    this.sceneRevisionListeners.forEach((listener) => {
      listener()
    })
  }
}

function sameLocalSceneInvalidationConnectionIdentity(
  left: LocalSceneInvalidationConnectionIdentity,
  right: LocalSceneInvalidationConnectionIdentity,
): boolean {
  return (
    left.engine === right.engine &&
    left.sheetName === right.sheetName &&
    left.shouldUseRemoteRenderTileSource === right.shouldUseRemoteRenderTileSource &&
    sameStringListIdentity(left.visibleAddresses, right.visibleAddresses)
  )
}

function sameStringListIdentity(left: readonly string[], right: readonly string[]): boolean {
  if (left === right) {
    return true
  }
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}
