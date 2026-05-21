import { useEffect, useMemo, useState, type CSSProperties } from 'react'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import { splitSelectionFillRangeAroundActiveCell } from './gridSelectionFillRanges.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { workbookThemeColors } from './workbookTheme.js'

type VisualRectRole = 'selection-fill' | 'selection-border' | 'active-border' | 'fill-handle' | 'header-fill' | 'hover-fill'

export interface GridSelectionVisualRect {
  readonly role: VisualRectRole
  readonly key: string
  readonly bounds: Rectangle
  readonly strokeWidth?: number | undefined
}

export interface GridSelectionVisualOverlayProps {
  readonly geometry?: GridGeometrySnapshot | null | undefined
  readonly getGeometrySnapshot?: (() => GridGeometrySnapshot | null) | undefined
  readonly gridSelection: GridSelection
  readonly hoverCell?: Item | null | undefined
  readonly selectedCell: Item
  readonly selectionChromeMode?: 'visible' | 'geometry-only' | 'chrome-only' | undefined
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
  readonly scrollTransformStore?: WorkbookGridScrollStore | undefined
}

export function GridSelectionVisualOverlay(props: GridSelectionVisualOverlayProps) {
  const {
    geometry: staticGeometry,
    getGeometrySnapshot,
    gridSelection,
    hoverCell,
    scrollTransformStore,
    selectedCell,
    selectionChromeMode = 'visible',
    selectionRange,
    showFillHandle,
  } = props
  const [scrollVersion, setScrollVersion] = useState(0)
  useEffect(() => {
    if (!scrollTransformStore) {
      return
    }
    return scrollTransformStore.subscribe(() => {
      setScrollVersion((version) => version + 1)
    })
  }, [scrollTransformStore])
  const geometry = useMemo(() => {
    void scrollVersion
    return getGeometrySnapshot?.() ?? staticGeometry ?? null
  }, [getGeometrySnapshot, staticGeometry, scrollVersion])
  const rects = useMemo(
    () =>
      geometry
        ? buildGridSelectionVisualRects({
            geometry,
            gridSelection,
            hoverCell: hoverCell ?? null,
            selectedCell,
            selectionRange,
            showFillHandle,
          })
        : [],
    [geometry, gridSelection, hoverCell, selectedCell, selectionRange, showFillHandle],
  )

  if (rects.length === 0) {
    return null
  }

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-20 overflow-hidden"
      data-testid="grid-selection-visual-overlay"
    >
      {rects.map((rect) => (
        <div
          className={classNameForRole(rect.role)}
          data-grid-selection-visual-key={rect.key}
          data-grid-selection-visual-role={rect.role}
          key={keyForRect(rect)}
          style={styleForRect(rect, shouldHideVisualRect(rect.role, selectionChromeMode))}
        />
      ))}
    </div>
  )
}

function shouldHideVisualRect(role: VisualRectRole, mode: NonNullable<GridSelectionVisualOverlayProps['selectionChromeMode']>): boolean {
  if (mode === 'geometry-only') {
    return role !== 'hover-fill'
  }
  if (mode === 'chrome-only') {
    return role === 'selection-fill' || role === 'header-fill' || role === 'hover-fill'
  }
  return false
}

export function buildGridSelectionVisualRects(input: {
  readonly geometry: GridGeometrySnapshot
  readonly gridSelection: GridSelection
  readonly hoverCell?: Item | null | undefined
  readonly selectedCell: Item
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
}): readonly GridSelectionVisualRect[] {
  const rects: GridSelectionVisualRect[] = []
  appendHoverVisualRects(rects, input)
  appendAxisSelectionVisualRects(rects, input)
  appendBodySelectionVisualRects(rects, input)
  return rects
}

function appendHoverVisualRects(
  rects: GridSelectionVisualRect[],
  input: {
    readonly geometry: GridGeometrySnapshot
    readonly hoverCell?: Item | null | undefined
    readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  },
): void {
  const hoverCell = input.hoverCell ?? null
  if (!hoverCell) {
    return
  }
  if (input.selectionRange && cellInRange(hoverCell, input.selectionRange)) {
    return
  }
  let segmentIndex = 0
  for (const bounds of input.geometry.rangeScreenRects({ x: hoverCell[0], y: hoverCell[1], width: 1, height: 1 })) {
    appendInsetRect(rects, 'hover-fill', `hover-fill:cell:${hoverCell[0]}:${hoverCell[1]}:${segmentIndex}`, bounds, 1, 1)
    segmentIndex += 1
  }
}

function appendBodySelectionVisualRects(
  rects: GridSelectionVisualRect[],
  input: {
    readonly geometry: GridGeometrySnapshot
    readonly gridSelection: GridSelection
    readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
    readonly showFillHandle: boolean
  },
): void {
  if (!input.selectionRange) {
    return
  }

  const hasAxisSelection = input.gridSelection.columns.length > 0 || input.gridSelection.rows.length > 0
  if (hasAxisSelection) {
    const activeCell = input.gridSelection.current?.cell ?? [input.selectionRange.x, input.selectionRange.y]
    appendCellBorderRects(rects, input.geometry, activeCell, 'active-border', `active-border:cell:${activeCell[0]}:${activeCell[1]}`)
    return
  }

  const isMultiCellSelection = input.selectionRange.width > 1 || input.selectionRange.height > 1
  const activeCell = input.gridSelection.current?.cell ?? null
  if (isMultiCellSelection) {
    let fillIndex = 0
    for (const fillRange of splitSelectionFillRangeAroundActiveCell(input.selectionRange, activeCell)) {
      let segmentIndex = 0
      for (const bounds of input.geometry.rangeScreenRects(fillRange)) {
        appendInsetRect(rects, 'selection-fill', stableRangeKey('selection-fill:range', fillRange, fillIndex, segmentIndex), bounds, 1, 1)
        segmentIndex += 1
      }
      fillIndex += 1
    }
  }

  if (isMultiCellSelection) {
    let segmentIndex = 0
    for (const bounds of input.geometry.rangeScreenRects(input.selectionRange)) {
      rects.push({
        role: 'selection-border',
        key: stableRangeKey('selection-border:range', input.selectionRange, 0, segmentIndex),
        bounds,
        strokeWidth: 2,
      })
      segmentIndex += 1
    }
  } else {
    appendCellBorderRects(
      rects,
      input.geometry,
      [input.selectionRange.x, input.selectionRange.y],
      'active-border',
      `active-border:cell:${input.selectionRange.x}:${input.selectionRange.y}`,
    )
  }

  if (activeCell && isMultiCellSelection && cellInRange(activeCell, input.selectionRange)) {
    appendCellBorderRects(rects, input.geometry, activeCell, 'active-border', `active-border:cell:${activeCell[0]}:${activeCell[1]}`, {
      strokeWidth: 2,
    })
  }

  if (input.showFillHandle) {
    const handle = input.geometry.fillHandleScreenRect(input.selectionRange)
    if (handle) {
      rects.push({ role: 'fill-handle', key: stableRangeKey('fill-handle:range', input.selectionRange, 0, 0), bounds: handle })
    }
  }
}

function appendAxisSelectionVisualRects(
  rects: GridSelectionVisualRect[],
  input: {
    readonly geometry: GridGeometrySnapshot
    readonly gridSelection: GridSelection
    readonly selectedCell: Item
    readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  },
): void {
  const columnRanges = resolveSelectedAxisRanges({
    axisRanges: input.gridSelection.columns.ranges,
    fallbackIndex: input.selectedCell[0],
    fallbackRange: input.selectionRange
      ? { start: input.selectionRange.x, endExclusive: input.selectionRange.x + input.selectionRange.width }
      : null,
  })
  const rowRanges = resolveSelectedAxisRanges({
    axisRanges: input.gridSelection.rows.ranges,
    fallbackIndex: input.selectedCell[1],
    fallbackRange: input.selectionRange
      ? { start: input.selectionRange.y, endExclusive: input.selectionRange.y + input.selectionRange.height }
      : null,
  })

  for (const index of visibleColumnIndexes(input.geometry)) {
    if (!isIndexSelected(index, columnRanges)) {
      continue
    }
    const bounds = input.geometry.columnHeaderScreenRect(index)
    const clip =
      index < input.geometry.camera.frozenColumnCount
        ? paneFrame(input.geometry, 'column-header-frozen')
        : paneFrame(input.geometry, 'column-header-body')
    const clipped = bounds && clip ? clipRect(bounds, clip) : null
    if (clipped) {
      appendInsetRect(rects, 'header-fill', `header-fill:column:${index}`, clipped, 1, 1)
    }
  }
  for (const index of visibleRowIndexes(input.geometry)) {
    if (!isIndexSelected(index, rowRanges)) {
      continue
    }
    const bounds = input.geometry.rowHeaderScreenRect(index)
    const clip =
      index < input.geometry.camera.frozenRowCount
        ? paneFrame(input.geometry, 'row-header-frozen')
        : paneFrame(input.geometry, 'row-header-body')
    const clipped = bounds && clip ? clipRect(bounds, clip) : null
    if (clipped) {
      appendInsetRect(rects, 'header-fill', `header-fill:row:${index}`, clipped, 1, 1)
    }
  }

  const activeCell = input.gridSelection.current?.cell ?? input.selectedCell
  if (input.gridSelection.columns.length > 0) {
    for (const range of columnRanges) {
      const start = Math.max(0, range.start)
      const endExclusive = Math.max(start + 1, Math.min(MAX_COLS, range.endExclusive))
      let fillIndex = 0
      for (const fillRange of splitSelectionFillRangeAroundActiveCell(
        { x: start, y: 0, width: endExclusive - start, height: MAX_ROWS },
        activeCell,
      )) {
        let segmentIndex = 0
        for (const bounds of input.geometry.rangeScreenRects(fillRange)) {
          appendInsetRect(
            rects,
            'selection-fill',
            stableRangeKey(`selection-fill:columns:${start}:${endExclusive}`, fillRange, fillIndex, segmentIndex),
            bounds,
            1,
            1,
          )
          segmentIndex += 1
        }
        fillIndex += 1
      }
    }
  }
  if (input.gridSelection.rows.length > 0) {
    for (const range of rowRanges) {
      const start = Math.max(0, range.start)
      const endExclusive = Math.max(start + 1, Math.min(MAX_ROWS, range.endExclusive))
      let fillIndex = 0
      for (const fillRange of splitSelectionFillRangeAroundActiveCell(
        { x: 0, y: start, width: MAX_COLS, height: endExclusive - start },
        activeCell,
      )) {
        let segmentIndex = 0
        for (const bounds of input.geometry.rangeScreenRects(fillRange)) {
          appendInsetRect(
            rects,
            'selection-fill',
            stableRangeKey(`selection-fill:rows:${start}:${endExclusive}`, fillRange, fillIndex, segmentIndex),
            bounds,
            1,
            1,
          )
          segmentIndex += 1
        }
        fillIndex += 1
      }
    }
  }
}

function appendCellBorderRects(
  rects: GridSelectionVisualRect[],
  geometry: GridGeometrySnapshot,
  cell: readonly [number, number],
  role: VisualRectRole,
  keyPrefix: string,
  options?: {
    readonly strokeWidth?: number | undefined
  },
): void {
  let segmentIndex = 0
  for (const bounds of geometry.rangeScreenRects({ x: cell[0], y: cell[1], width: 1, height: 1 })) {
    rects.push({
      role,
      key: `${keyPrefix}:${segmentIndex}`,
      bounds,
      strokeWidth: options?.strokeWidth,
    })
    segmentIndex += 1
  }
}

function appendInsetRect(
  rects: GridSelectionVisualRect[],
  role: VisualRectRole,
  key: string,
  bounds: Rectangle,
  insetX: number,
  insetY: number,
): void {
  const insetBounds = {
    x: bounds.x + insetX,
    y: bounds.y + insetY,
    width: Math.max(0, bounds.width - insetX * 2),
    height: Math.max(0, bounds.height - insetY * 2),
  }
  if (insetBounds.width > 0 && insetBounds.height > 0) {
    rects.push({ role, key, bounds: insetBounds })
  }
}

function stableRangeKey(
  prefix: string,
  range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>,
  fillIndex: number,
  segmentIndex: number,
): string {
  return `${prefix}:${range.x}:${range.y}:${range.width}:${range.height}:${fillIndex}:${segmentIndex}`
}

interface AxisSelectionRange {
  readonly start: number
  readonly endExclusive: number
}

function resolveSelectedAxisRanges(input: {
  readonly axisRanges: readonly (readonly [number, number])[]
  readonly fallbackIndex: number
  readonly fallbackRange: AxisSelectionRange | null
}): readonly AxisSelectionRange[] {
  if (input.axisRanges.length > 0) {
    return input.axisRanges.map(([start, endExclusive]) => ({ start, endExclusive }))
  }
  if (input.fallbackRange) {
    return [input.fallbackRange]
  }
  return [{ start: input.fallbackIndex, endExclusive: input.fallbackIndex + 1 }]
}

function visibleColumnIndexes(geometry: GridGeometrySnapshot): readonly number[] {
  const indexes: number[] = []
  for (let index = 0; index < geometry.camera.frozenColumnCount; index += 1) {
    if (!geometry.columns.isHidden(index) && geometry.columns.sizeOf(index) > 0) {
      indexes.push(index)
    }
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
    if (!geometry.rows.isHidden(index) && geometry.rows.sizeOf(index) > 0) {
      indexes.push(index)
    }
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

function paneFrame(geometry: GridGeometrySnapshot, kind: GridGeometrySnapshot['camera']['panes'][number]['kind']): Rectangle | null {
  return geometry.camera.panes.find((pane) => pane.kind === kind)?.frame ?? null
}

function clipRect(rect: Rectangle, clip: Rectangle): Rectangle | null {
  const x = Math.max(rect.x, clip.x)
  const y = Math.max(rect.y, clip.y)
  const right = Math.min(rect.x + rect.width, clip.x + clip.width)
  const bottom = Math.min(rect.y + rect.height, clip.y + clip.height)
  if (right <= x || bottom <= y) {
    return null
  }
  return { x, y, width: right - x, height: bottom - y }
}

function cellInRange(cell: Item, range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>): boolean {
  return cell[0] >= range.x && cell[0] < range.x + range.width && cell[1] >= range.y && cell[1] < range.y + range.height
}

function classNameForRole(role: VisualRectRole): string {
  switch (role) {
    case 'selection-fill':
      return 'absolute box-border'
    case 'header-fill':
      return 'absolute box-border'
    case 'hover-fill':
      return 'absolute box-border'
    case 'selection-border':
      return 'absolute box-border'
    case 'active-border':
      return 'absolute box-border'
    case 'fill-handle':
      return 'absolute box-border rounded-[1px] border border-[var(--wb-surface)] bg-[var(--wb-accent)]'
  }
}

function keyForRect(rect: GridSelectionVisualRect): string {
  return `${rect.role}:${rect.key}`
}

function styleForRect(rect: GridSelectionVisualRect, geometryOnly = false): CSSProperties {
  const base = {
    height: rect.bounds.height,
    left: rect.bounds.x,
    opacity: geometryOnly ? 0 : undefined,
    top: rect.bounds.y,
    width: rect.bounds.width,
  }
  if (rect.role === 'selection-border' || rect.role === 'active-border') {
    const strokeWidth = rect.strokeWidth ?? (rect.role === 'active-border' ? 2 : 1)
    return {
      ...base,
      backgroundColor: 'transparent',
      borderBottomWidth: strokeWidth,
      borderColor: workbookThemeColors.selectionAccent,
      borderLeftWidth: strokeWidth,
      borderRightWidth: strokeWidth,
      borderStyle: 'solid',
      borderTopWidth: strokeWidth,
      boxSizing: 'border-box',
    }
  }
  if (rect.role === 'fill-handle') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.selectionAccent,
      boxShadow: `0 0 0 1px ${workbookThemeColors.selectionAccent}`,
    }
  }
  if (rect.role === 'header-fill') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.selectionHeaderFill,
    }
  }
  if (rect.role === 'hover-fill') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.hoverFill,
    }
  }
  return {
    ...base,
    backgroundColor: workbookThemeColors.selectionFill,
  }
}
