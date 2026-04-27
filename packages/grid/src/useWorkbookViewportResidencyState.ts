import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { formatAddress } from '@bilig/formula'
import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { Item, Rectangle } from './gridTypes.js'
import { collectViewportItems } from './gridViewportItems.js'
import { sameViewportBounds } from './gridViewportController.js'
import { resolveResidentViewport } from './workbookGridViewport.js'
import { viewportFromVisibleRegion } from './useGridCameraState.js'

export interface WorkbookResidentHeaderRegion {
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  readonly tx: number
  readonly ty: number
  readonly freezeRows: number
  readonly freezeCols: number
}

export interface WorkbookViewportResidencyState {
  readonly viewport: Viewport
  readonly residentViewport: Viewport
  readonly renderTileViewport: Viewport
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: WorkbookResidentHeaderRegion
  readonly sceneRevision: number
  readonly visibleAddresses: readonly string[]
  readonly visibleItems: readonly Item[]
}

export function useWorkbookViewportResidencyState(input: {
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly visibleRegion: VisibleRegionState
}): WorkbookViewportResidencyState {
  const { engine, freezeCols, freezeRows, sheetName, shouldUseRemoteRenderTileSource, visibleRegion } = input
  const [sceneRevision, setSceneRevision] = useState(0)
  const viewport = useMemo<Viewport>(() => viewportFromVisibleRegion(visibleRegion), [visibleRegion])
  const residentViewportRef = useRef<Viewport>(resolveResidentViewport(viewport))
  const nextResidentViewport = resolveResidentViewport(viewport)
  if (!sameViewportBounds(residentViewportRef.current, nextResidentViewport)) {
    residentViewportRef.current = nextResidentViewport
  }
  const residentViewport = residentViewportRef.current
  const renderTileViewport = useMemo<Viewport>(
    () => ({
      rowStart: freezeRows > 0 ? 0 : residentViewport.rowStart,
      rowEnd: residentViewport.rowEnd,
      colStart: freezeCols > 0 ? 0 : residentViewport.colStart,
      colEnd: residentViewport.colEnd,
    }),
    [freezeCols, freezeRows, residentViewport.colEnd, residentViewport.colStart, residentViewport.rowEnd, residentViewport.rowStart],
  )
  const visibleItems = useMemo(() => {
    return collectViewportItems(residentViewport, { freezeRows, freezeCols })
  }, [freezeCols, freezeRows, residentViewport])
  const visibleAddresses = useMemo(() => visibleItems.map(([col, row]) => formatAddress(row, col)), [visibleItems])
  const residentHeaderRegion = useMemo<WorkbookResidentHeaderRegion>(
    () => ({
      range: {
        x: residentViewport.colStart,
        y: residentViewport.rowStart,
        width: residentViewport.colEnd - residentViewport.colStart + 1,
        height: residentViewport.rowEnd - residentViewport.rowStart + 1,
      },
      tx: 0,
      ty: 0,
      freezeRows,
      freezeCols,
    }),
    [freezeCols, freezeRows, residentViewport.colEnd, residentViewport.colStart, residentViewport.rowEnd, residentViewport.rowStart],
  )
  const invalidateScene = useCallback(() => {
    setSceneRevision((current) => current + 1)
  }, [])

  useEffect(() => {
    if (shouldUseRemoteRenderTileSource) {
      return
    }
    return engine.subscribeCells(sheetName, visibleAddresses, invalidateScene)
  }, [engine, invalidateScene, sheetName, shouldUseRemoteRenderTileSource, visibleAddresses])

  return {
    renderTileViewport,
    residentHeaderItems: visibleItems,
    residentHeaderRegion,
    residentViewport,
    sceneRevision,
    viewport,
    visibleAddresses,
    visibleItems,
  }
}
