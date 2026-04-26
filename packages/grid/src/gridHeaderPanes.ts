import type { Viewport } from '@bilig/protocol'
import type { GridGpuRect, GridGpuScene } from './gridGpuScene.js'
import type { GridTextItem, GridTextScene } from './gridTextScene.js'
import type { GridMetrics } from './gridMetrics.js'
import type { WorkbookRenderPaneState } from './renderer-v2/pane-scene-types.js'
import { packGridScenePacketV2, type GridScenePacketPaneId } from './renderer-v2/scene-packet-v2.js'

export interface GridHeaderPaneState extends WorkbookRenderPaneState {
  readonly paneId: 'top-frozen' | 'top-body' | 'left-frozen' | 'left-body'
}

interface ClipRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

function clipGpuRect(rect: GridGpuRect, clip: ClipRect): GridGpuRect | null {
  const left = Math.max(rect.x, clip.x)
  const top = Math.max(rect.y, clip.y)
  const right = Math.min(rect.x + rect.width, clip.x + clip.width)
  const bottom = Math.min(rect.y + rect.height, clip.y + clip.height)
  if (right <= left || bottom <= top) {
    return null
  }
  return {
    x: left - clip.x,
    y: top - clip.y,
    width: right - left,
    height: bottom - top,
    color: rect.color,
  }
}

function clipTextItem(item: GridTextItem, clip: ClipRect): GridTextItem | null {
  const left = Math.max(item.x, clip.x)
  const top = Math.max(item.y, clip.y)
  const right = Math.min(item.x + item.width, clip.x + clip.width)
  const bottom = Math.min(item.y + item.height, clip.y + clip.height)
  if (right <= left || bottom <= top) {
    return null
  }
  return {
    ...item,
    x: item.x - clip.x,
    y: item.y - clip.y,
    clipInsetTop: item.clipInsetTop + Math.max(0, clip.y - item.y),
    clipInsetRight: item.clipInsetRight + Math.max(0, item.x + item.width - (clip.x + clip.width)),
    clipInsetBottom: item.clipInsetBottom + Math.max(0, item.y + item.height - (clip.y + clip.height)),
    clipInsetLeft: item.clipInsetLeft + Math.max(0, clip.x - item.x),
  }
}

function clipGpuScene(scene: GridGpuScene, clip: ClipRect): GridGpuScene {
  return {
    fillRects: scene.fillRects.flatMap((rect) => {
      const clipped = clipGpuRect(rect, clip)
      return clipped ? [clipped] : []
    }),
    borderRects: scene.borderRects.flatMap((rect) => {
      const clipped = clipGpuRect(rect, clip)
      return clipped ? [clipped] : []
    }),
  }
}

function clipTextScene(scene: GridTextScene, clip: ClipRect): GridTextScene {
  return {
    items: scene.items.flatMap((item) => {
      const clipped = clipTextItem(item, clip)
      return clipped ? [clipped] : []
    }),
  }
}

export function buildHeaderPaneStates(input: {
  readonly sheetName?: string | undefined
  readonly residentViewport?: Viewport | undefined
  readonly freezeRows?: number | undefined
  readonly freezeCols?: number | undefined
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
  readonly hostWidth: number
  readonly hostHeight: number
  readonly gridMetrics: GridMetrics
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly residentBodyWidth: number
  readonly residentBodyHeight: number
}): readonly GridHeaderPaneState[] {
  const {
    gpuScene,
    textScene,
    hostWidth,
    hostHeight,
    gridMetrics,
    frozenColumnWidth,
    frozenRowHeight,
    residentBodyWidth,
    residentBodyHeight,
  } = input
  const sheetName = input.sheetName ?? 'Sheet1'
  const residentViewport = input.residentViewport ?? { colEnd: 0, colStart: 0, rowEnd: 0, rowStart: 0 }
  const freezeRows = Math.max(0, input.freezeRows ?? 0)
  const freezeCols = Math.max(0, input.freezeCols ?? 0)
  const bodyFrameWidth = Math.max(0, hostWidth - gridMetrics.rowMarkerWidth - frozenColumnWidth)
  const bodyFrameHeight = Math.max(0, hostHeight - gridMetrics.headerHeight - frozenRowHeight)
  const panes: GridHeaderPaneState[] = []

  if (frozenColumnWidth > 0) {
    const clip: ClipRect = {
      x: gridMetrics.rowMarkerWidth,
      y: 0,
      width: frozenColumnWidth,
      height: gridMetrics.headerHeight,
    }
    panes.push({
      generation: 0,
      paneId: 'top-frozen',
      frame: {
        x: clip.x,
        y: clip.y,
        width: clip.width,
        height: clip.height,
      },
      surfaceSize: {
        width: clip.width,
        height: clip.height,
      },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: false, y: false },
      packedScene: packHeaderScenePacket({
        clip,
        generation: 0,
        gpuScene,
        paneId: 'top-frozen',
        sheetName,
        textScene,
        viewport: { colEnd: Math.max(0, freezeCols - 1), colStart: 0, rowEnd: 0, rowStart: 0 },
      }),
    })
  }

  if (bodyFrameWidth > 0 && residentBodyWidth > 0) {
    const clip: ClipRect = {
      x: gridMetrics.rowMarkerWidth + frozenColumnWidth,
      y: 0,
      width: residentBodyWidth,
      height: gridMetrics.headerHeight,
    }
    panes.push({
      generation: 0,
      paneId: 'top-body',
      frame: {
        x: clip.x,
        y: clip.y,
        width: bodyFrameWidth,
        height: clip.height,
      },
      surfaceSize: {
        width: residentBodyWidth,
        height: clip.height,
      },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: true, y: false },
      packedScene: packHeaderScenePacket({
        clip,
        generation: 0,
        gpuScene,
        paneId: 'top-body',
        sheetName,
        textScene,
        viewport: { ...residentViewport, rowEnd: 0, rowStart: 0 },
      }),
    })
  }

  if (frozenRowHeight > 0) {
    const clip: ClipRect = {
      x: 0,
      y: gridMetrics.headerHeight,
      width: gridMetrics.rowMarkerWidth,
      height: frozenRowHeight,
    }
    panes.push({
      generation: 0,
      paneId: 'left-frozen',
      frame: {
        x: clip.x,
        y: clip.y,
        width: clip.width,
        height: clip.height,
      },
      surfaceSize: {
        width: clip.width,
        height: clip.height,
      },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: false, y: false },
      packedScene: packHeaderScenePacket({
        clip,
        generation: 0,
        gpuScene,
        paneId: 'left-frozen',
        sheetName,
        textScene,
        viewport: { colEnd: 0, colStart: 0, rowEnd: Math.max(0, freezeRows - 1), rowStart: 0 },
      }),
    })
  }

  if (bodyFrameHeight > 0 && residentBodyHeight > 0) {
    const clip: ClipRect = {
      x: 0,
      y: gridMetrics.headerHeight + frozenRowHeight,
      width: gridMetrics.rowMarkerWidth,
      height: residentBodyHeight,
    }
    panes.push({
      generation: 0,
      paneId: 'left-body',
      frame: {
        x: clip.x,
        y: clip.y,
        width: clip.width,
        height: bodyFrameHeight,
      },
      surfaceSize: {
        width: clip.width,
        height: residentBodyHeight,
      },
      contentOffset: { x: 0, y: 0 },
      scrollAxes: { x: false, y: true },
      packedScene: packHeaderScenePacket({
        clip,
        generation: 0,
        gpuScene,
        paneId: 'left-body',
        sheetName,
        textScene,
        viewport: { ...residentViewport, colEnd: 0, colStart: 0 },
      }),
    })
  }

  return panes
}

function packHeaderScenePacket(input: {
  readonly generation: number
  readonly sheetName: string
  readonly paneId: GridScenePacketPaneId
  readonly viewport: Viewport
  readonly clip: ClipRect
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}) {
  const gpuScene = clipGpuScene(input.gpuScene, input.clip)
  const textScene = clipTextScene(input.textScene, input.clip)
  return packGridScenePacketV2({
    generation: input.generation,
    gpuScene,
    paneId: input.paneId,
    sheetName: input.sheetName,
    surfaceSize: { height: input.clip.height, width: input.clip.width },
    textScene,
    viewport: normalizeViewport(input.viewport),
  })
}

function normalizeViewport(viewport: Viewport): Viewport {
  const rowStart = Math.max(0, Math.min(viewport.rowStart, viewport.rowEnd))
  const colStart = Math.max(0, Math.min(viewport.colStart, viewport.colEnd))
  return {
    colEnd: Math.max(colStart, viewport.colEnd),
    colStart,
    rowEnd: Math.max(rowStart, viewport.rowEnd),
    rowStart,
  }
}
