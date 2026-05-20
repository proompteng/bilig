import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridGeometrySnapshot, GridPaneKind } from '../gridGeometry.js'
import { parseGpuColor, type GridGpuRect } from '../gridGpuPrimitives.js'
import type { HeaderSelection } from '../gridPointer.js'
import { splitSelectionFillRangeAroundActiveCell } from '../gridSelectionFillRanges.js'
import type { CompactSelectionState, GridSelection, Item, Rectangle } from '../gridTypes.js'
import { workbookThemeColors } from '../workbookTheme.js'
import { GRID_RECT_FLOAT_COUNT_V3, GRID_RECT_INSTANCE_FLOAT_COUNT_V3, packGridRectBufferV3 } from './rect-instance-buffer.js'

export const DYNAMIC_OVERLAY_RECT_FLOAT_COUNT_V3 = GRID_RECT_FLOAT_COUNT_V3
export const DYNAMIC_OVERLAY_RECT_INSTANCE_FLOAT_COUNT_V3 = GRID_RECT_INSTANCE_FLOAT_COUNT_V3

export interface DynamicGridOverlayBatchV3 {
  readonly seq: number
  readonly cameraSeq: number
  readonly generatedAt: number
  readonly sheetName: string
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly rects: Float32Array
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly fillRectCount: number
  readonly borderRectCount: number
  readonly rectSignature: string
}

export interface DynamicGridPreviewRectV3 {
  readonly role: 'target' | 'source'
  readonly bounds: Rectangle
}

export type DynamicGridSelectionOverlayModeV3 = 'all' | 'fills-only'

interface BorderSides {
  readonly bottom: boolean
  readonly left: boolean
  readonly right: boolean
  readonly top: boolean
}

export function buildDynamicGridOverlayBatchV3(input: {
  readonly geometry: GridGeometrySnapshot
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly fillPreviewRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null | undefined
  readonly previewRects?: readonly DynamicGridPreviewRectV3[] | undefined
  readonly gridSelection?: GridSelection | null | undefined
  readonly selectedCell?: Item | null | undefined
  readonly hoveredCell?: readonly [number, number] | null | undefined
  readonly showFillHandle: boolean
  readonly activeHeaderDrag?: HeaderSelection | null | undefined
  readonly showHoverOverlay?: boolean | undefined
  readonly selectionOverlayMode?: DynamicGridSelectionOverlayModeV3 | undefined
  readonly showSelectionOverlay?: boolean | undefined
  readonly resizeGuideColumn?: number | null | undefined
  readonly resizeGuideColumnWidth?: number | null | undefined
  readonly resizeGuideRow?: number | null | undefined
  readonly resizeGuideRowHeight?: number | null | undefined
}): DynamicGridOverlayBatchV3 {
  const fillRects: GridGpuRect[] = []
  const borderRects: GridGpuRect[] = []
  if (input.showSelectionOverlay !== false) {
    const selectionOverlayMode = input.selectionOverlayMode ?? 'all'
    const showSelectionChrome = selectionOverlayMode === 'all'
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
      showFillHandle: showSelectionChrome && input.showFillHandle,
      showSelectionChrome,
    })
  }
  appendFillPreviewOverlay({
    borderRects,
    fillPreviewRange: input.fillPreviewRange ?? null,
    geometry: input.geometry,
  })
  appendPreviewRects({
    borderRects,
    fillRects,
    previewRects: input.previewRects ?? [],
  })
  if (input.showHoverOverlay !== false) {
    appendHoverOverlay({
      fillRects,
      geometry: input.geometry,
      hoveredCell: input.hoveredCell ?? null,
      selectionRange: input.selectionRange,
    })
  }
  appendResizeGuides({
    borderRects,
    geometry: input.geometry,
    resizeGuideColumn: input.resizeGuideColumn ?? null,
    resizeGuideColumnWidth: input.resizeGuideColumnWidth ?? null,
    resizeGuideRow: input.resizeGuideRow ?? null,
    resizeGuideRowHeight: input.resizeGuideRowHeight ?? null,
  })
  appendHeaderDragGuides({
    activeHeaderDrag: input.activeHeaderDrag ?? null,
    borderRects,
    fillRects,
    geometry: input.geometry,
    gridSelection: input.gridSelection ?? null,
  })
  appendFrozenSeparators({ borderRects, geometry: input.geometry })

  const surfaceSize = resolveOverlaySurfaceSize(input.geometry)
  const rectBuffer = packGridRectBufferV3({ borderRects, fillRects }, surfaceSize)
  return {
    ...rectBuffer,
    cameraSeq: input.geometry.camera.seq,
    generatedAt: input.geometry.camera.updatedAt,
    seq: input.geometry.camera.seq,
    sheetName: input.geometry.camera.sheetName,
    surfaceSize,
  }
}

function appendFillPreviewOverlay(input: {
  readonly geometry: GridGeometrySnapshot
  readonly fillPreviewRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly borderRects: GridGpuRect[]
}): void {
  if (!input.fillPreviewRange) {
    return
  }
  const color = parseGpuColor(workbookThemeColors.textMuted)
  for (const rect of input.geometry.rangeScreenRects(input.fillPreviewRange)) {
    appendBorderRects(input.borderRects, rect, color, 1)
  }
}

function appendPreviewRects(input: {
  readonly previewRects: readonly DynamicGridPreviewRectV3[]
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  for (const preview of input.previewRects) {
    const isTarget = preview.role === 'target'
    input.fillRects.push({
      ...preview.bounds,
      color: parseGpuColor(isTarget ? 'rgba(56, 189, 248, 0.08)' : 'rgba(148, 163, 184, 0.06)'),
    })
    appendBorderRects(
      input.borderRects,
      preview.bounds,
      parseGpuColor(isTarget ? 'rgba(14, 116, 144, 0.9)' : 'rgba(100, 116, 139, 0.9)'),
      1,
    )
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
  readonly showSelectionChrome: boolean
  readonly fillRects: GridGpuRect[]
  readonly borderRects: GridGpuRect[]
}): void {
  if (!input.selectionRange) {
    return
  }
  const hasAxisSelection = (input.gridSelection?.columns.length ?? 0) > 0 || (input.gridSelection?.rows.length ?? 0) > 0
  const borderColor = parseGpuColor(workbookThemeColors.accent)
  if (hasAxisSelection) {
    if (!input.showSelectionChrome) {
      return
    }
    const activeCell = input.gridSelection?.current?.cell ?? [input.selectionRange.x, input.selectionRange.y]
    for (const activeRect of input.geometry.rangeScreenRects({ x: activeCell[0], y: activeCell[1], width: 1, height: 1 })) {
      appendBorderRects(input.borderRects, activeRect, borderColor, 2)
    }
    return
  }
  const isMultiCellSelection = input.selectionRange.width > 1 || input.selectionRange.height > 1
  const activeCell = input.gridSelection?.current?.cell ?? null
  if (isMultiCellSelection) {
    appendSelectionFillRects({
      activeCell,
      color: parseGpuColor(workbookThemeColors.selectionFill),
      fillRects: input.fillRects,
      geometry: input.geometry,
      range: input.selectionRange,
    })
    if (input.showSelectionChrome) {
      for (const rect of input.geometry.rangeScreenRects(input.selectionRange)) {
        appendBorderRects(input.borderRects, rect, borderColor, 1)
      }
    }
  } else if (input.showSelectionChrome) {
    for (const rect of input.geometry.rangeScreenRects(input.selectionRange)) {
      appendBorderRects(input.borderRects, rect, borderColor, 2)
    }
  }
  if (input.showSelectionChrome && activeCell && isMultiCellSelection && cellInRange(activeCell, input.selectionRange)) {
    for (const activeRect of input.geometry.rangeScreenRects({ x: activeCell[0], y: activeCell[1], width: 1, height: 1 })) {
      appendBorderRects(input.borderRects, activeRect, borderColor, 2)
    }
  }
  if (input.showFillHandle) {
    const handle = input.geometry.fillHandleScreenRect(input.selectionRange)
    if (handle) {
      input.fillRects.push({ ...handle, color: borderColor })
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

  const selectionFillColor = parseGpuColor(workbookThemeColors.selectionFill)
  if ((input.gridSelection?.columns.length ?? 0) > 0) {
    appendAxisBodySelectionFills({
      activeCell: input.gridSelection?.current?.cell ?? input.selectedCell,
      color: selectionFillColor,
      fillRects: input.fillRects,
      geometry: input.geometry,
      ranges: columnRanges,
      axis: 'column',
    })
  }
  if ((input.gridSelection?.rows.length ?? 0) > 0) {
    appendAxisBodySelectionFills({
      activeCell: input.gridSelection?.current?.cell ?? input.selectedCell,
      color: selectionFillColor,
      fillRects: input.fillRects,
      geometry: input.geometry,
      ranges: rowRanges,
      axis: 'row',
    })
  }
}

function appendSelectionFillRects(input: {
  readonly geometry: GridGeometrySnapshot
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  readonly activeCell?: Item | null | undefined
  readonly color: GridGpuRect['color']
  readonly fillRects: GridGpuRect[]
}): void {
  for (const fillRange of splitSelectionFillRangeAroundActiveCell(input.range, input.activeCell)) {
    for (const rect of input.geometry.rangeScreenRects(fillRange)) {
      const fill = insetRect(rect, 1, 1)
      if (fill.width > 0 && fill.height > 0) {
        input.fillRects.push({ ...fill, color: input.color })
      }
    }
  }
}

function appendAxisBodySelectionFills(input: {
  readonly geometry: GridGeometrySnapshot
  readonly ranges: readonly AxisSelectionRange[]
  readonly axis: 'column' | 'row'
  readonly activeCell?: Item | null | undefined
  readonly color: GridGpuRect['color']
  readonly fillRects: GridGpuRect[]
}): void {
  for (const range of input.ranges) {
    const start = Math.max(0, range.start)
    const endExclusive = Math.max(start + 1, Math.min(input.axis === 'column' ? MAX_COLS : MAX_ROWS, range.endExclusive))
    const selectionRange =
      input.axis === 'column'
        ? { x: start, y: 0, width: endExclusive - start, height: MAX_ROWS }
        : { x: 0, y: start, width: MAX_COLS, height: endExclusive - start }
    appendSelectionFillRects({
      activeCell: input.activeCell,
      color: input.color,
      fillRects: input.fillRects,
      geometry: input.geometry,
      range: selectionRange,
    })
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
  readonly resizeGuideColumnWidth: number | null
  readonly resizeGuideRow: number | null
  readonly resizeGuideRowHeight: number | null
  readonly borderRects: GridGpuRect[]
}): void {
  const color = parseGpuColor('rgba(33, 86, 58, 0.72)')
  const glowColor = parseGpuColor('rgba(191, 213, 196, 0.28)')
  if (input.resizeGuideColumn !== null) {
    const rect = resolveColumnResizeGuideRect(input.geometry, input.resizeGuideColumn, input.resizeGuideColumnWidth)
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
    const rect = resolveRowResizeGuideRect(input.geometry, input.resizeGuideRow, input.resizeGuideRowHeight)
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

function resolveColumnResizeGuideRect(geometry: GridGeometrySnapshot, columnIndex: number, previewWidth: number | null): Rectangle | null {
  if (previewWidth === null) {
    return geometry.resizeGuideScreenRect({ kind: 'column', index: columnIndex })
  }
  const defaultRect = geometry.resizeGuideScreenRect({ kind: 'column', index: columnIndex })
  const headerRect = geometry.columnHeaderScreenRect(columnIndex)
  if (!defaultRect || !headerRect) {
    return null
  }
  const surfaceSize = resolveOverlaySurfaceSize(geometry)
  return {
    height: surfaceSize.height,
    width: defaultRect.width,
    x: headerRect.x + Math.max(0, previewWidth) - 1,
    y: defaultRect.y,
  }
}

function resolveRowResizeGuideRect(geometry: GridGeometrySnapshot, rowIndex: number, previewHeight: number | null): Rectangle | null {
  if (previewHeight === null) {
    return geometry.resizeGuideScreenRect({ kind: 'row', index: rowIndex })
  }
  const defaultRect = geometry.resizeGuideScreenRect({ kind: 'row', index: rowIndex })
  const headerRect = geometry.rowHeaderScreenRect(rowIndex)
  if (!defaultRect || !headerRect) {
    return null
  }
  const surfaceSize = resolveOverlaySurfaceSize(geometry)
  return {
    height: defaultRect.height,
    width: surfaceSize.width,
    x: defaultRect.x,
    y: headerRect.y + Math.max(0, previewHeight) - 1,
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
  const indexes: number[] = []
  for (let index = 0; index < geometry.camera.frozenColumnCount; index += 1) {
    indexes.push(index)
  }
  const bodyRange = geometry.columns.visibleRangeForWorldRect(geometry.camera.bodyWorldX, geometry.camera.bodyViewportWidth)
  for (let index = bodyRange.startIndex; index < bodyRange.endIndexExclusive; index += 1) {
    if (index < geometry.camera.frozenColumnCount) {
      continue
    }
    if (!geometry.columns.isHidden(index) && geometry.columns.sizeOf(index) > 0) {
      indexes.push(index)
    }
  }
  return indexes
}

function visibleRowIndexes(geometry: GridGeometrySnapshot): readonly number[] {
  const indexes: number[] = []
  for (let index = 0; index < geometry.camera.frozenRowCount; index += 1) {
    indexes.push(index)
  }
  const bodyRange = geometry.rows.visibleRangeForWorldRect(geometry.camera.bodyWorldY, geometry.camera.bodyViewportHeight)
  for (let index = bodyRange.startIndex; index < bodyRange.endIndexExclusive; index += 1) {
    if (index < geometry.camera.frozenRowCount) {
      continue
    }
    if (!geometry.rows.isHidden(index) && geometry.rows.sizeOf(index) > 0) {
      indexes.push(index)
    }
  }
  return indexes
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

function cellInRange(cell: Item, range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>): boolean {
  return cell[0] >= range.x && cell[0] < range.x + range.width && cell[1] >= range.y && cell[1] < range.y + range.height
}

function appendBorderRects(target: GridGpuRect[], rect: Rectangle, color: GridGpuRect['color'], thickness: number): void {
  appendBorderRectsForSides(target, rect, color, thickness, { bottom: true, left: true, right: true, top: true })
}

function appendBorderRectsForSides(
  target: GridGpuRect[],
  rect: Rectangle,
  color: GridGpuRect['color'],
  thickness: number,
  sides: BorderSides,
): void {
  const nextRects: GridGpuRect[] = []
  if (sides.top) {
    nextRects.push({ x: rect.x, y: rect.y, width: rect.width, height: thickness, color })
  }
  if (sides.bottom) {
    nextRects.push({ x: rect.x, y: rect.y + rect.height - thickness, width: rect.width, height: thickness, color })
  }
  if (sides.left) {
    nextRects.push({ x: rect.x, y: rect.y, width: thickness, height: rect.height, color })
  }
  if (sides.right) {
    nextRects.push({ x: rect.x + rect.width - thickness, y: rect.y, width: thickness, height: rect.height, color })
  }
  target.push(...nextRects.filter((candidate) => candidate.width > 0 && candidate.height > 0))
}
