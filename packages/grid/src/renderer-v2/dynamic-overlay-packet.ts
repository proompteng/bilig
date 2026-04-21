import type { GridGeometrySnapshot } from '../gridGeometry.js'
import { parseGpuColor, type GridGpuRect, type GridGpuScene } from '../gridGpuScene.js'
import type { Rectangle } from '../gridTypes.js'
import { workbookThemeColors } from '../workbookTheme.js'

export interface DynamicGridOverlayPacket {
  readonly gpuScene: GridGpuScene
  readonly textScene: { readonly items: readonly [] }
}

export function buildDynamicGridOverlayPacket(input: {
  readonly geometry: GridGeometrySnapshot
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly hoveredCell?: readonly [number, number] | null | undefined
  readonly showFillHandle: boolean
  readonly resizeGuideColumn?: number | null | undefined
  readonly resizeGuideRow?: number | null | undefined
}): DynamicGridOverlayPacket {
  const fillRects: GridGpuRect[] = []
  const borderRects: GridGpuRect[] = []
  appendSelectionOverlay({
    borderRects,
    fillRects,
    geometry: input.geometry,
    selectionRange: input.selectionRange,
    showFillHandle: input.showFillHandle,
  })
  appendHoverOverlay({
    fillRects,
    geometry: input.geometry,
    hoveredCell: input.hoveredCell ?? null,
    selectionRange: input.selectionRange,
  })
  appendResizeGuides({
    borderRects,
    geometry: input.geometry,
    resizeGuideColumn: input.resizeGuideColumn ?? null,
    resizeGuideRow: input.resizeGuideRow ?? null,
  })
  appendFrozenSeparators({ borderRects, geometry: input.geometry })
  return {
    gpuScene: {
      borderRects,
      fillRects,
    },
    textScene: { items: [] },
  }
}

function appendHoverOverlay(input: {
  readonly geometry: GridGeometrySnapshot
  readonly hoveredCell: readonly [number, number] | null
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly fillRects: GridGpuRect[]
}): void {
  if (!input.hoveredCell) {
    return
  }
  const [col, row] = input.hoveredCell
  if (
    input.selectionRange &&
    col >= input.selectionRange.x &&
    col < input.selectionRange.x + input.selectionRange.width &&
    row >= input.selectionRange.y &&
    row < input.selectionRange.y + input.selectionRange.height
  ) {
    return
  }
  const rect = input.geometry.cellScreenRect(col, row)
  if (!rect) {
    return
  }
  input.fillRects.push({
    x: rect.x + 1,
    y: rect.y + 1,
    width: Math.max(0, rect.width - 2),
    height: Math.max(0, rect.height - 2),
    color: parseGpuColor('rgba(31, 122, 67, 0.05)'),
  })
}

function appendSelectionOverlay(input: {
  readonly geometry: GridGeometrySnapshot
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  if (!input.selectionRange) {
    return
  }
  const fillColor = parseGpuColor('rgba(33, 86, 58, 0.08)')
  const borderColor = parseGpuColor(workbookThemeColors.accent)
  for (const rect of input.geometry.rangeScreenRects(input.selectionRange)) {
    input.fillRects.push({ ...rect, color: fillColor })
    appendBorderRects(input.borderRects, rect, borderColor, 1)
  }
  if (input.showFillHandle) {
    const handle = input.geometry.fillHandleScreenRect(input.selectionRange)
    if (handle) {
      input.fillRects.push({ ...handle, color: borderColor })
      appendBorderRects(input.borderRects, handle, parseGpuColor(workbookThemeColors.surface), 1)
    }
  }
}

function appendResizeGuides(input: {
  readonly geometry: GridGeometrySnapshot
  readonly resizeGuideColumn: number | null
  readonly resizeGuideRow: number | null
  readonly borderRects: GridGpuRect[]
}): void {
  const color = parseGpuColor('rgba(33, 86, 58, 0.72)')
  const glowColor = parseGpuColor('rgba(191, 213, 196, 0.28)')
  if (input.resizeGuideColumn !== null) {
    const rect = input.geometry.columnHeaderScreenRect(input.resizeGuideColumn)
    if (rect) {
      const x = rect.x + rect.width - 1
      input.borderRects.push({
        x: x - 1,
        y: 0,
        width: 3,
        height: input.geometry.camera.bodyViewportHeight + input.geometry.camera.frozenHeight + rect.height,
        color: glowColor,
      })
      input.borderRects.push({
        x,
        y: 0,
        width: 1,
        height: input.geometry.camera.bodyViewportHeight + input.geometry.camera.frozenHeight + rect.height,
        color,
      })
    }
  }
  if (input.resizeGuideRow !== null) {
    const rect = input.geometry.rowHeaderScreenRect(input.resizeGuideRow)
    if (rect) {
      const y = rect.y + rect.height - 1
      input.borderRects.push({
        x: 0,
        y: y - 1,
        width: input.geometry.camera.bodyViewportWidth + input.geometry.camera.frozenWidth + rect.width,
        height: 3,
        color: glowColor,
      })
      input.borderRects.push({
        x: 0,
        y,
        width: input.geometry.camera.bodyViewportWidth + input.geometry.camera.frozenWidth + rect.width,
        height: 1,
        color,
      })
    }
  }
}

function appendFrozenSeparators(input: { readonly geometry: GridGeometrySnapshot; readonly borderRects: GridGpuRect[] }): void {
  const color = parseGpuColor(workbookThemeColors.border)
  const hostWidth =
    input.geometry.camera.bodyViewportWidth +
    input.geometry.camera.frozenWidth +
    (input.geometry.camera.panes.find((pane) => pane.kind === 'row-header-body')?.frame.width ?? 0)
  const hostHeight =
    input.geometry.camera.bodyViewportHeight +
    input.geometry.camera.frozenHeight +
    (input.geometry.camera.panes.find((pane) => pane.kind === 'column-header-body')?.frame.height ?? 0)
  if (input.geometry.camera.frozenWidth > 0) {
    const x = input.geometry.camera.panes.find((pane) => pane.kind === 'body')?.frame.x ?? 0
    input.borderRects.push({ x: x - 1, y: 0, width: 1, height: hostHeight, color })
  }
  if (input.geometry.camera.frozenHeight > 0) {
    const y = input.geometry.camera.panes.find((pane) => pane.kind === 'body')?.frame.y ?? 0
    input.borderRects.push({ x: 0, y: y - 1, width: hostWidth, height: 1, color })
  }
}

function appendBorderRects(target: GridGpuRect[], rect: Rectangle, color: GridGpuRect['color'], thickness: number): void {
  target.push(
    { x: rect.x, y: rect.y, width: rect.width, height: thickness, color },
    { x: rect.x, y: rect.y + rect.height - thickness, width: rect.width, height: thickness, color },
    { x: rect.x, y: rect.y, width: thickness, height: rect.height, color },
    { x: rect.x + rect.width - thickness, y: rect.y, width: thickness, height: rect.height, color },
  )
}
