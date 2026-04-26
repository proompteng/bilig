import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridGeometrySnapshot, GridPaneKind } from '../gridGeometry.js'
import { parseGpuColor, type GridGpuRect } from '../gridGpuScene.js'
import type { HeaderSelection } from '../gridPointer.js'
import type { CompactSelectionState, GridSelection, Item, Rectangle } from '../gridTypes.js'
import { workbookThemeColors } from '../workbookTheme.js'
import { packGridScenePacketV2, type GridScenePacketV2 } from './scene-packet-v2.js'

export interface DynamicGridOverlayPacket {
  readonly packedScene: GridScenePacketV2
}

export function buildDynamicGridOverlayPacket(input: {
  readonly geometry: GridGeometrySnapshot
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly gridSelection?: GridSelection | null | undefined
  readonly selectedCell?: Item | null | undefined
  readonly hoveredCell?: readonly [number, number] | null | undefined
  readonly showFillHandle: boolean
  readonly activeHeaderDrag?: HeaderSelection | null | undefined
  readonly resizeGuideColumn?: number | null | undefined
  readonly resizeGuideRow?: number | null | undefined
}): DynamicGridOverlayPacket {
  const fillRects: GridGpuRect[] = []
  const borderRects: GridGpuRect[] = []
  appendAxisSelectionOverlay({
    borderRects,
    fillRects,
    geometry: input.geometry,
    gridSelection: input.gridSelection ?? null,
    selectedCell: input.selectedCell ?? null,
    selectionRange: input.selectionRange,
  })
  appendSelectionOverlay({
    borderRects,
    fillRects,
    geometry: input.geometry,
    gridSelection: input.gridSelection ?? null,
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
  appendHeaderDragGuides({
    activeHeaderDrag: input.activeHeaderDrag ?? null,
    borderRects,
    fillRects,
    geometry: input.geometry,
    gridSelection: input.gridSelection ?? null,
  })
  appendFrozenSeparators({ borderRects, geometry: input.geometry })
  const gpuScene = {
    borderRects,
    fillRects,
  }
  const surfaceSize = resolveOverlaySurfaceSize(input.geometry)
  return {
    packedScene: packGridScenePacketV2({
      cameraSeq: input.geometry.camera.seq,
      generatedAt: input.geometry.camera.updatedAt,
      generation: input.geometry.camera.seq,
      gpuScene,
      paneId: 'overlay',
      requestSeq: input.geometry.camera.seq,
      sheetName: input.geometry.camera.sheetName,
      surfaceSize,
      textScene: { items: [] },
      viewport: { colStart: 0, colEnd: 0, rowStart: 0, rowEnd: 0 },
    }),
  }
}

function resolveOverlaySurfaceSize(geometry: GridGeometrySnapshot): { readonly width: number; readonly height: number } {
  return geometry.camera.panes.reduce(
    (size, pane) => ({
      height: Math.max(size.height, pane.frame.y + pane.frame.height),
      width: Math.max(size.width, pane.frame.x + pane.frame.width),
    }),
    { height: 0, width: 0 },
  )
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
  readonly gridSelection: GridSelection | null
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  if (!input.selectionRange) {
    return
  }
  const hasAxisSelection = (input.gridSelection?.columns.length ?? 0) > 0 || (input.gridSelection?.rows.length ?? 0) > 0
  const fillColor = parseGpuColor('rgba(33, 86, 58, 0.08)')
  const borderColor = parseGpuColor(workbookThemeColors.accent)
  if (hasAxisSelection) {
    const activeCell = input.gridSelection?.current?.cell ?? [input.selectionRange.x, input.selectionRange.y]
    for (const activeRect of input.geometry.rangeScreenRects({ x: activeCell[0], y: activeCell[1], width: 1, height: 1 })) {
      appendBorderRects(input.borderRects, activeRect, borderColor, 1)
    }
    return
  }
  for (const rect of input.geometry.rangeScreenRects(input.selectionRange)) {
    input.fillRects.push({ ...rect, color: fillColor })
    appendBorderRects(input.borderRects, rect, borderColor, 1)
  }
  if (input.showFillHandle) {
    const handle = input.geometry.fillHandleScreenRect(input.selectionRange)
    if (handle) {
      input.fillRects.push({ ...handle, color: borderColor })
      appendSolidRectAsBorderStrips(input.borderRects, handle, borderColor)
      appendBorderRects(input.borderRects, handle, parseGpuColor(workbookThemeColors.surface), 1)
    }
  }
}

function appendAxisSelectionOverlay(input: {
  readonly geometry: GridGeometrySnapshot
  readonly gridSelection: GridSelection | null
  readonly selectedCell: Item | null
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  const headerFillColor = parseGpuColor(workbookThemeColors.accentSoft)
  const bodyFillColor = parseGpuColor('rgba(33, 86, 58, 0.08)')
  const columnRanges = resolveSelectedAxisRanges({
    axis: input.gridSelection?.columns ?? null,
    fallbackIndex: input.selectedCell?.[0] ?? null,
    fallbackRange: input.selectionRange
      ? { start: input.selectionRange.x, endExclusive: input.selectionRange.x + input.selectionRange.width }
      : null,
  })
  const rowRanges = resolveSelectedAxisRanges({
    axis: input.gridSelection?.rows ?? null,
    fallbackIndex: input.selectedCell?.[1] ?? null,
    fallbackRange: input.selectionRange
      ? { start: input.selectionRange.y, endExclusive: input.selectionRange.y + input.selectionRange.height }
      : null,
  })

  appendSelectedColumnHeaderFills({ color: headerFillColor, fillRects: input.fillRects, geometry: input.geometry, ranges: columnRanges })
  appendSelectedRowHeaderFills({ color: headerFillColor, fillRects: input.fillRects, geometry: input.geometry, ranges: rowRanges })

  if (input.gridSelection && input.gridSelection.columns.length > 0) {
    for (const range of input.gridSelection.columns.ranges) {
      appendRangeFills(input.fillRects, input.geometry, { x: range[0], y: 0, width: range[1] - range[0], height: MAX_ROWS }, bodyFillColor)
    }
  }
  if (input.gridSelection && input.gridSelection.rows.length > 0) {
    for (const range of input.gridSelection.rows.ranges) {
      appendRangeFills(input.fillRects, input.geometry, { x: 0, y: range[0], width: MAX_COLS, height: range[1] - range[0] }, bodyFillColor)
    }
  }
}

function appendSelectedColumnHeaderFills(input: {
  readonly geometry: GridGeometrySnapshot
  readonly ranges: readonly AxisSelectionRange[]
  readonly color: GridGpuRect['color']
  readonly fillRects: GridGpuRect[]
}): void {
  if (input.ranges.length === 0) {
    return
  }
  const clipFrozen = paneFrame(input.geometry, 'column-header-frozen')
  const clipBody = paneFrame(input.geometry, 'column-header-body')
  for (const index of visibleColumnIndexes(input.geometry)) {
    if (!isIndexSelected(index, input.ranges)) {
      continue
    }
    const rect = input.geometry.columnHeaderScreenRect(index)
    const clip = index < input.geometry.camera.frozenColumnCount ? clipFrozen : clipBody
    const clipped = rect && clip ? clipRect(rect, clip) : null
    if (clipped) {
      input.fillRects.push({ ...insetRect(clipped, 1, 1), color: input.color })
    }
  }
}

function appendSelectedRowHeaderFills(input: {
  readonly geometry: GridGeometrySnapshot
  readonly ranges: readonly AxisSelectionRange[]
  readonly color: GridGpuRect['color']
  readonly fillRects: GridGpuRect[]
}): void {
  if (input.ranges.length === 0) {
    return
  }
  const clipFrozen = paneFrame(input.geometry, 'row-header-frozen')
  const clipBody = paneFrame(input.geometry, 'row-header-body')
  for (const index of visibleRowIndexes(input.geometry)) {
    if (!isIndexSelected(index, input.ranges)) {
      continue
    }
    const rect = input.geometry.rowHeaderScreenRect(index)
    const clip = index < input.geometry.camera.frozenRowCount ? clipFrozen : clipBody
    const clipped = rect && clip ? clipRect(rect, clip) : null
    if (clipped) {
      input.fillRects.push({ ...insetRect(clipped, 1, 1), color: input.color })
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
    const rect = input.geometry.resizeGuideScreenRect({ kind: 'column', index: input.resizeGuideColumn })
    if (rect) {
      input.borderRects.push({
        x: rect.x - 1,
        y: rect.y,
        width: 3,
        height: rect.height,
        color: glowColor,
      })
      input.borderRects.push({
        ...rect,
        color,
      })
    }
  }
  if (input.resizeGuideRow !== null) {
    const rect = input.geometry.resizeGuideScreenRect({ kind: 'row', index: input.resizeGuideRow })
    if (rect) {
      input.borderRects.push({
        x: rect.x,
        y: rect.y - 1,
        width: rect.width,
        height: 3,
        color: glowColor,
      })
      input.borderRects.push({
        ...rect,
        color,
      })
    }
  }
}

function appendHeaderDragGuides(input: {
  readonly geometry: GridGeometrySnapshot
  readonly gridSelection: GridSelection | null
  readonly activeHeaderDrag: HeaderSelection | null
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  if (!input.activeHeaderDrag || !input.gridSelection) {
    return
  }
  const color = parseGpuColor('rgba(33, 86, 58, 0.72)')
  const host = hostRect(input.geometry)
  if (input.activeHeaderDrag.kind === 'column' && input.gridSelection.columns.length > 0) {
    const start = input.gridSelection.columns.first()
    const end = input.gridSelection.columns.last()
    if (start === undefined || end === undefined) {
      return
    }
    const startRect = input.geometry.columnHeaderScreenRect(start)
    const endRect = input.geometry.columnHeaderScreenRect(end)
    if (startRect && endRect) {
      input.borderRects.push(
        { x: startRect.x, y: 0, width: 1, height: host.height, color },
        { x: endRect.x + endRect.width - 1, y: 0, width: 1, height: host.height, color },
      )
    }
    const activeRect = input.geometry.columnHeaderScreenRect(input.activeHeaderDrag.index)
    if (activeRect) {
      input.fillRects.push({ x: activeRect.x, y: Math.max(0, activeRect.height - 3), width: activeRect.width, height: 3, color })
    }
  }
  if (input.activeHeaderDrag.kind === 'row' && input.gridSelection.rows.length > 0) {
    const start = input.gridSelection.rows.first()
    const end = input.gridSelection.rows.last()
    if (start === undefined || end === undefined) {
      return
    }
    const startRect = input.geometry.rowHeaderScreenRect(start)
    const endRect = input.geometry.rowHeaderScreenRect(end)
    if (startRect && endRect) {
      input.borderRects.push(
        { x: 0, y: startRect.y, width: host.width, height: 1, color },
        { x: 0, y: endRect.y + endRect.height - 1, width: host.width, height: 1, color },
      )
    }
    const activeRect = input.geometry.rowHeaderScreenRect(input.activeHeaderDrag.index)
    if (activeRect) {
      input.fillRects.push({ x: Math.max(0, activeRect.width - 3), y: activeRect.y, width: 3, height: activeRect.height, color })
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

interface AxisSelectionRange {
  readonly start: number
  readonly endExclusive: number
}

function resolveSelectedAxisRanges(input: {
  readonly axis: CompactSelectionState | null
  readonly fallbackIndex: number | null
  readonly fallbackRange: AxisSelectionRange | null
}): readonly AxisSelectionRange[] {
  if (input.axis && input.axis.length > 0) {
    return input.axis.ranges.map(([start, endExclusive]) => ({ start, endExclusive }))
  }
  if (input.fallbackRange) {
    return [input.fallbackRange]
  }
  return input.fallbackIndex === null ? [] : [{ start: input.fallbackIndex, endExclusive: input.fallbackIndex + 1 }]
}

function visibleColumnIndexes(geometry: GridGeometrySnapshot): readonly number[] {
  const indexes = new Set<number>()
  for (let index = 0; index < geometry.camera.frozenColumnCount; index += 1) {
    indexes.add(index)
  }
  const bodyRange = geometry.columns.visibleRangeForWorldRect(geometry.camera.bodyWorldX, geometry.camera.bodyViewportWidth)
  for (let index = bodyRange.startIndex; index < bodyRange.endIndexExclusive; index += 1) {
    if (!geometry.columns.isHidden(index) && geometry.columns.sizeOf(index) > 0) {
      indexes.add(index)
    }
  }
  return [...indexes].toSorted((left, right) => left - right)
}

function visibleRowIndexes(geometry: GridGeometrySnapshot): readonly number[] {
  const indexes = new Set<number>()
  for (let index = 0; index < geometry.camera.frozenRowCount; index += 1) {
    indexes.add(index)
  }
  const bodyRange = geometry.rows.visibleRangeForWorldRect(geometry.camera.bodyWorldY, geometry.camera.bodyViewportHeight)
  for (let index = bodyRange.startIndex; index < bodyRange.endIndexExclusive; index += 1) {
    if (!geometry.rows.isHidden(index) && geometry.rows.sizeOf(index) > 0) {
      indexes.add(index)
    }
  }
  return [...indexes].toSorted((left, right) => left - right)
}

function appendRangeFills(
  target: GridGpuRect[],
  geometry: GridGeometrySnapshot,
  range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  color: GridGpuRect['color'],
): void {
  for (const rect of geometry.rangeScreenRects(range)) {
    target.push({ ...insetRect(rect, 1, 1), color })
  }
}

function isIndexSelected(index: number, ranges: readonly AxisSelectionRange[]): boolean {
  return ranges.some((range) => index >= range.start && index < range.endExclusive)
}

function paneFrame(geometry: GridGeometrySnapshot, kind: GridPaneKind): Rectangle | null {
  return geometry.camera.panes.find((pane) => pane.kind === kind)?.frame ?? null
}

function hostRect(geometry: GridGeometrySnapshot): Rectangle {
  return geometry.camera.panes.reduce(
    (current, pane) => ({
      x: 0,
      y: 0,
      width: Math.max(current.width, pane.frame.x + pane.frame.width),
      height: Math.max(current.height, pane.frame.y + pane.frame.height),
    }),
    { x: 0, y: 0, width: 0, height: 0 },
  )
}

function clipRect(target: Rectangle, clip: Rectangle): Rectangle | null {
  const x = Math.max(target.x, clip.x)
  const y = Math.max(target.y, clip.y)
  const right = Math.min(target.x + target.width, clip.x + clip.width)
  const bottom = Math.min(target.y + target.height, clip.y + clip.height)
  return right <= x || bottom <= y ? null : { x, y, width: right - x, height: bottom - y }
}

function insetRect(rect: Rectangle, insetX: number, insetY: number): Rectangle {
  return {
    x: rect.x + insetX,
    y: rect.y + insetY,
    width: Math.max(0, rect.width - insetX * 2),
    height: Math.max(0, rect.height - insetY * 2),
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

function appendSolidRectAsBorderStrips(target: GridGpuRect[], rect: Rectangle, color: GridGpuRect['color']): void {
  const width = Math.max(1, Math.ceil(rect.width))
  for (let offset = 0; offset < width; offset += 1) {
    target.push({
      x: rect.x + offset,
      y: rect.y,
      width: 1,
      height: rect.height,
      color,
    })
  }
}
