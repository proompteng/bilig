import { formatAddress } from '@bilig/formula'
import type { Viewport } from '@bilig/protocol'
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
  readonly sceneRevision: number
  readonly visibleRegion: VisibleRegionState
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
      sceneRevision: input.sceneRevision,
      viewport,
      visibleAddresses: residentCache.visibleAddresses,
      visibleItems: residentCache.visibleItems,
    }
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
}
