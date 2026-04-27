import type { Viewport } from '@bilig/protocol'
import type { GridGpuRect, GridGpuScene } from './gridGpuScene.js'
import type { GridTextItem, GridTextScene } from './gridTextScene.js'
import type { GridMetrics } from './gridMetrics.js'
import type { TextQuadRun } from './renderer-v2/line-text-quad-buffer.js'
import { packGridRectBufferV3 } from './renderer-v3/rect-instance-buffer.js'

export interface GridHeaderPaneState {
  readonly paneId: 'corner-header' | 'top-frozen' | 'top-body' | 'left-frozen' | 'left-body'
  readonly frame: ClipRect
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
  readonly scrollAxes: {
    readonly x: boolean
    readonly y: boolean
  }
  readonly rects: Float32Array
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly fillRectCount: number
  readonly borderRectCount: number
  readonly rectSignature: string
  readonly textRuns: readonly TextQuadRun[]
  readonly textCount: number
  readonly textSignature: string
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
  const bodyFrameWidth = Math.max(0, hostWidth - gridMetrics.rowMarkerWidth - frozenColumnWidth)
  const bodyFrameHeight = Math.max(0, hostHeight - gridMetrics.headerHeight - frozenRowHeight)
  const panes: GridHeaderPaneState[] = []

  if (gridMetrics.rowMarkerWidth > 0 && gridMetrics.headerHeight > 0) {
    const clip: ClipRect = {
      x: 0,
      y: 0,
      width: gridMetrics.rowMarkerWidth,
      height: gridMetrics.headerHeight,
    }
    panes.push({
      paneId: 'corner-header',
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
      ...packHeaderPaneBatch({
        clip,
        gpuScene,
        paneId: 'corner-header',
        sheetName,
        textScene,
      }),
    })
  }

  if (frozenColumnWidth > 0) {
    const clip: ClipRect = {
      x: gridMetrics.rowMarkerWidth,
      y: 0,
      width: frozenColumnWidth,
      height: gridMetrics.headerHeight,
    }
    panes.push({
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
      ...packHeaderPaneBatch({
        clip,
        gpuScene,
        paneId: 'top-frozen',
        sheetName,
        textScene,
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
      ...packHeaderPaneBatch({
        clip,
        gpuScene,
        paneId: 'top-body',
        sheetName,
        textScene,
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
      ...packHeaderPaneBatch({
        clip,
        gpuScene,
        paneId: 'left-frozen',
        sheetName,
        textScene,
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
      ...packHeaderPaneBatch({
        clip,
        gpuScene,
        paneId: 'left-body',
        sheetName,
        textScene,
      }),
    })
  }

  return panes
}

function packHeaderPaneBatch(input: {
  readonly sheetName: string
  readonly paneId: GridHeaderPaneState['paneId']
  readonly clip: ClipRect
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}) {
  const gpuScene = clipGpuScene(input.gpuScene, input.clip)
  const textScene = clipTextScene(input.textScene, input.clip)
  const surfaceSize = { height: input.clip.height, width: input.clip.width }
  const textRuns = textScene.items.map(mapHeaderTextRun)
  return {
    ...packGridRectBufferV3(gpuScene, surfaceSize),
    sheetName: input.sheetName,
    textCount: textRuns.length,
    textRuns,
    textSignature: resolveHeaderTextSignature(textRuns),
  }
}

function mapHeaderTextRun(item: GridTextItem): TextQuadRun {
  return {
    align: item.align,
    clipHeight: Math.max(0, item.height - item.clipInsetTop - item.clipInsetBottom),
    clipWidth: Math.max(0, item.width - item.clipInsetLeft - item.clipInsetRight),
    clipX: item.x + item.clipInsetLeft,
    clipY: item.y + item.clipInsetTop,
    color: item.color,
    font: item.font,
    fontSize: item.fontSize,
    height: item.height,
    strike: item.strike,
    text: item.text,
    underline: item.underline,
    width: item.width,
    wrap: item.wrap,
    x: item.x,
    y: item.y,
  }
}

function resolveHeaderTextSignature(textRuns: readonly TextQuadRun[]): string {
  let hash = 2_166_136_261
  hash = mixNumber(hash, textRuns.length)
  for (const run of textRuns) {
    hash = mixString(hash, run.text)
    hash = mixNumber(hash, run.x)
    hash = mixNumber(hash, run.y)
    hash = mixNumber(hash, run.width ?? 0)
    hash = mixNumber(hash, run.height ?? 0)
    hash = mixNumber(hash, run.clipX ?? 0)
    hash = mixNumber(hash, run.clipY ?? 0)
    hash = mixNumber(hash, run.clipWidth ?? 0)
    hash = mixNumber(hash, run.clipHeight ?? 0)
    hash = mixString(hash, run.align ?? '')
    hash = mixNumber(hash, run.wrap ? 1 : 0)
    hash = mixString(hash, run.font ?? '')
    hash = mixNumber(hash, run.fontSize ?? 0)
    hash = mixString(hash, run.color ?? '')
    hash = mixNumber(hash, run.underline ? 1 : 0)
    hash = mixNumber(hash, run.strike ? 1 : 0)
  }
  return hash.toString(36)
}

function mixString(hash: number, value: string): number {
  let next = hash
  for (let index = 0; index < value.length; index += 1) {
    next = mixInteger(next, value.charCodeAt(index))
  }
  return next
}

function mixNumber(hash: number, value: number): number {
  return mixInteger(hash, Math.round(value * 1_000))
}

function mixInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}
