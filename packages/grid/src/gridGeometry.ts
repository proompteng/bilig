import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridMetrics } from './gridMetrics.js'
import type { Rectangle } from './gridTypes.js'
import { createGridAxisWorldIndexFromRecords, type GridAxisAnchor, type GridAxisWorldIndex } from './gridAxisWorldIndex.js'

export type GridPaneKind =
  | 'corner-header'
  | 'column-header-frozen'
  | 'column-header-body'
  | 'row-header-frozen'
  | 'row-header-body'
  | 'frozen-cells'
  | 'frozen-rows'
  | 'frozen-columns'
  | 'body'

export interface GridPaneGeometry {
  readonly kind: GridPaneKind
  readonly frame: Rectangle
  readonly scrollAxes: {
    readonly x: boolean
    readonly y: boolean
  }
}

export interface GridCameraSnapshotV2 {
  readonly seq: number
  readonly sheetName: string
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly bodyScrollX: number
  readonly bodyScrollY: number
  readonly bodyWorldX: number
  readonly bodyWorldY: number
  readonly bodyViewportWidth: number
  readonly bodyViewportHeight: number
  readonly frozenColumnCount: number
  readonly frozenRowCount: number
  readonly frozenWidth: number
  readonly frozenHeight: number
  readonly columnAnchor: GridAxisAnchor
  readonly rowAnchor: GridAxisAnchor
  readonly dpr: number
  readonly updatedAt: number
  readonly velocityX: number
  readonly velocityY: number
  readonly axisVersionX: number
  readonly axisVersionY: number
  readonly panes: readonly GridPaneGeometry[]
}

export interface GridGeometrySnapshot {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  cellWorldRect(col: number, row: number): Rectangle | null
  cellScreenRect(col: number, row: number): Rectangle | null
  cellScreenRectForPane(col: number, row: number, paneKind: GridPaneKind): Rectangle | null
  rangeWorldRects(range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>): readonly Rectangle[]
  rangeScreenRects(range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>): readonly Rectangle[]
  fillHandleScreenRect(range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>): Rectangle | null
  editorScreenRect(col: number, row: number): Rectangle | null
  resizeGuideScreenRect(input: GridResizeGuideState): Rectangle | null
  columnHeaderScreenRect(col: number): Rectangle | null
  rowHeaderScreenRect(row: number): Rectangle | null
  hitTestScreenPoint(point: { readonly x: number; readonly y: number }): { readonly col: number; readonly row: number } | null
}

export type GridResizeGuideState = { readonly kind: 'column'; readonly index: number } | { readonly kind: 'row'; readonly index: number }

export function createGridGeometrySnapshot(input: {
  readonly seq?: number
  readonly sheetName: string
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly hostWidth: number
  readonly hostHeight: number
  readonly dpr: number
  readonly gridMetrics: GridMetrics
  readonly columnWidths?: Readonly<Record<number, number>> | undefined
  readonly rowHeights?: Readonly<Record<number, number>> | undefined
  readonly hiddenColumns?: Readonly<Record<number, true>> | undefined
  readonly hiddenRows?: Readonly<Record<number, true>> | undefined
  readonly freezeRows?: number | undefined
  readonly freezeCols?: number | undefined
  readonly previousCamera?: GridCameraSnapshotV2 | null | undefined
  readonly updatedAt?: number | undefined
}): GridGeometrySnapshot {
  const columns = createGridAxisWorldIndexFromRecords({
    axisLength: MAX_COLS,
    defaultSize: input.gridMetrics.columnWidth,
    hidden: input.hiddenColumns,
    sizes: input.columnWidths,
  })
  const rows = createGridAxisWorldIndexFromRecords({
    axisLength: MAX_ROWS,
    defaultSize: input.gridMetrics.rowHeight,
    hidden: input.hiddenRows,
    sizes: input.rowHeights,
  })
  return createGridGeometrySnapshotFromAxes({
    ...input,
    columns,
    rows,
  })
}

export function createGridGeometrySnapshotFromAxes(input: {
  readonly seq?: number
  readonly sheetName: string
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly hostWidth: number
  readonly hostHeight: number
  readonly dpr: number
  readonly gridMetrics: GridMetrics
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly freezeRows?: number | undefined
  readonly freezeCols?: number | undefined
  readonly previousCamera?: GridCameraSnapshotV2 | null | undefined
  readonly updatedAt?: number | undefined
}): GridGeometrySnapshot {
  const frozenColumnCount = Math.max(0, Math.min(MAX_COLS, input.freezeCols ?? 0))
  const frozenRowCount = Math.max(0, Math.min(MAX_ROWS, input.freezeRows ?? 0))
  const frozenWidth = input.columns.span(0, frozenColumnCount)
  const frozenHeight = input.rows.span(0, frozenRowCount)
  const bodyPaneX = input.gridMetrics.rowMarkerWidth + frozenWidth
  const bodyPaneY = input.gridMetrics.headerHeight + frozenHeight
  const bodyViewportWidth = Math.max(0, input.hostWidth - bodyPaneX)
  const bodyViewportHeight = Math.max(0, input.hostHeight - bodyPaneY)
  const maxScrollX = Math.max(0, Math.max(0, input.columns.totalSize - frozenWidth) - bodyViewportWidth)
  const maxScrollY = Math.max(0, Math.max(0, input.rows.totalSize - frozenHeight) - bodyViewportHeight)
  const bodyScrollX = clamp(input.scrollLeft, 0, maxScrollX)
  const bodyScrollY = clamp(input.scrollTop, 0, maxScrollY)
  const bodyWorldX = frozenWidth + bodyScrollX
  const bodyWorldY = frozenHeight + bodyScrollY
  const updatedAt = input.updatedAt ?? (typeof performance === 'undefined' ? 0 : performance.now())
  const previous = input.previousCamera
  const elapsed = previous ? Math.max(1, updatedAt - previous.updatedAt) : 0
  const velocityX = previous ? ((bodyScrollX - previous.bodyScrollX) / elapsed) * 1000 : 0
  const velocityY = previous ? ((bodyScrollY - previous.bodyScrollY) / elapsed) * 1000 : 0
  const panes = createPaneGeometries({
    bodyPaneX,
    bodyPaneY,
    bodyViewportHeight,
    bodyViewportWidth,
    frozenHeight,
    frozenWidth,
    headerHeight: input.gridMetrics.headerHeight,
    hostHeight: input.hostHeight,
    hostWidth: input.hostWidth,
    rowHeaderWidth: input.gridMetrics.rowMarkerWidth,
  })
  const camera: GridCameraSnapshotV2 = {
    axisVersionX: input.columns.version,
    axisVersionY: input.rows.version,
    bodyScrollX,
    bodyScrollY,
    bodyViewportHeight,
    bodyViewportWidth,
    bodyWorldX,
    bodyWorldY,
    columnAnchor: input.columns.anchorAt(bodyWorldX),
    dpr: input.dpr,
    frozenColumnCount,
    frozenHeight,
    frozenRowCount,
    frozenWidth,
    panes,
    rowAnchor: input.rows.anchorAt(bodyWorldY),
    scrollLeft: input.scrollLeft,
    scrollTop: input.scrollTop,
    seq: input.seq ?? (previous?.seq ?? 0) + 1,
    sheetName: input.sheetName,
    updatedAt,
    velocityX,
    velocityY,
  }

  return {
    camera,
    columns: input.columns,
    rows: input.rows,
    cellScreenRectForPane: (col, row, paneKind) =>
      resolveCellScreenRectForPane({ camera, col, columns: input.columns, paneKind, row, rows: input.rows }),
    cellScreenRect: (col, row) => resolveCellScreenRect({ camera, col, columns: input.columns, row, rows: input.rows }),
    cellWorldRect: (col, row) => resolveCellWorldRect({ col, columns: input.columns, row, rows: input.rows }),
    columnHeaderScreenRect: (col) => resolveColumnHeaderScreenRect({ camera, col, columns: input.columns }),
    editorScreenRect: (col, row) => resolveEditorScreenRect({ camera, col, columns: input.columns, row, rows: input.rows }),
    fillHandleScreenRect: (range) => resolveFillHandleScreenRect({ camera, columns: input.columns, range, rows: input.rows }),
    hitTestScreenPoint: (point) => hitTestScreenPoint({ camera, columns: input.columns, point, rows: input.rows }),
    rangeWorldRects: (range) => resolveRangeWorldRects({ columns: input.columns, range, rows: input.rows }),
    rangeScreenRects: (range) => resolveRangeScreenRects({ camera, columns: input.columns, range, rows: input.rows }),
    resizeGuideScreenRect: (state) => resolveResizeGuideScreenRect({ camera, columns: input.columns, rows: input.rows, state }),
    rowHeaderScreenRect: (row) => resolveRowHeaderScreenRect({ camera, row, rows: input.rows }),
  }
}

function createPaneGeometries(input: {
  readonly bodyPaneX: number
  readonly bodyPaneY: number
  readonly bodyViewportWidth: number
  readonly bodyViewportHeight: number
  readonly frozenWidth: number
  readonly frozenHeight: number
  readonly headerHeight: number
  readonly hostWidth: number
  readonly hostHeight: number
  readonly rowHeaderWidth: number
}): readonly GridPaneGeometry[] {
  return [
    { kind: 'corner-header', frame: rect(0, 0, input.rowHeaderWidth, input.headerHeight), scrollAxes: { x: false, y: false } },
    {
      kind: 'column-header-frozen',
      frame: rect(input.rowHeaderWidth, 0, input.frozenWidth, input.headerHeight),
      scrollAxes: { x: false, y: false },
    },
    {
      kind: 'column-header-body',
      frame: rect(input.bodyPaneX, 0, input.bodyViewportWidth, input.headerHeight),
      scrollAxes: { x: true, y: false },
    },
    {
      kind: 'row-header-frozen',
      frame: rect(0, input.headerHeight, input.rowHeaderWidth, input.frozenHeight),
      scrollAxes: { x: false, y: false },
    },
    {
      kind: 'row-header-body',
      frame: rect(0, input.bodyPaneY, input.rowHeaderWidth, input.bodyViewportHeight),
      scrollAxes: { x: false, y: true },
    },
    {
      kind: 'frozen-cells',
      frame: rect(input.rowHeaderWidth, input.headerHeight, input.frozenWidth, input.frozenHeight),
      scrollAxes: { x: false, y: false },
    },
    {
      kind: 'frozen-rows',
      frame: rect(input.bodyPaneX, input.headerHeight, input.bodyViewportWidth, input.frozenHeight),
      scrollAxes: { x: true, y: false },
    },
    {
      kind: 'frozen-columns',
      frame: rect(input.rowHeaderWidth, input.bodyPaneY, input.frozenWidth, input.bodyViewportHeight),
      scrollAxes: { x: false, y: true },
    },
    {
      kind: 'body',
      frame: rect(input.bodyPaneX, input.bodyPaneY, input.bodyViewportWidth, input.bodyViewportHeight),
      scrollAxes: { x: true, y: true },
    },
  ]
}

function resolveCellWorldRect(input: {
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly col: number
  readonly row: number
}): Rectangle | null {
  const width = input.columns.sizeOf(input.col)
  const height = input.rows.sizeOf(input.row)
  if (width <= 0 || height <= 0 || input.columns.isHidden(input.col) || input.rows.isHidden(input.row)) {
    return null
  }
  return rect(input.columns.offsetOf(input.col), input.rows.offsetOf(input.row), width, height)
}

function resolveCellScreenRect(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly col: number
  readonly row: number
}): Rectangle | null {
  const world = resolveCellWorldRect(input)
  if (!world) {
    return null
  }
  const rowHeaderWidth = getBodyPaneX(input.camera) - input.camera.frozenWidth
  const columnHeaderHeight = getBodyPaneY(input.camera) - input.camera.frozenHeight
  const x =
    input.col < input.camera.frozenColumnCount ? rowHeaderWidth + world.x : getBodyPaneX(input.camera) + world.x - input.camera.bodyWorldX
  const y =
    input.row < input.camera.frozenRowCount ? columnHeaderHeight + world.y : getBodyPaneY(input.camera) + world.y - input.camera.bodyWorldY
  return rect(x, y, world.width, world.height)
}

function resolveCellScreenRectForPane(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly col: number
  readonly row: number
  readonly paneKind: GridPaneKind
}): Rectangle | null {
  if (!cellBelongsToPane(input.camera, input.col, input.row, input.paneKind)) {
    return null
  }
  const screenRect = resolveCellScreenRect(input)
  const pane = input.camera.panes.find((candidate) => candidate.kind === input.paneKind)
  return screenRect && pane ? clipRect(screenRect, pane.frame) : null
}

function cellBelongsToPane(camera: GridCameraSnapshotV2, col: number, row: number, paneKind: GridPaneKind): boolean {
  const frozenCol = col < camera.frozenColumnCount
  const frozenRow = row < camera.frozenRowCount
  switch (paneKind) {
    case 'frozen-cells':
      return frozenCol && frozenRow
    case 'frozen-rows':
      return !frozenCol && frozenRow
    case 'frozen-columns':
      return frozenCol && !frozenRow
    case 'body':
      return !frozenCol && !frozenRow
    case 'corner-header':
    case 'column-header-frozen':
    case 'column-header-body':
    case 'row-header-frozen':
    case 'row-header-body':
      return false
  }
}

function resolveRangeWorldRects(input: {
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
}): readonly Rectangle[] {
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, input.range.x))
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, input.range.y))
  const colEndExclusive = Math.max(colStart + 1, Math.min(MAX_COLS, input.range.x + input.range.width))
  const rowEndExclusive = Math.max(rowStart + 1, Math.min(MAX_ROWS, input.range.y + input.range.height))
  const width = input.columns.span(colStart, colEndExclusive)
  const height = input.rows.span(rowStart, rowEndExclusive)
  return width <= 0 || height <= 0 ? [] : [rect(input.columns.offsetOf(colStart), input.rows.offsetOf(rowStart), width, height)]
}

function resolveRangeScreenRects(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
}): readonly Rectangle[] {
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, input.range.x))
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, input.range.y))
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, input.range.x + input.range.width - 1))
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, input.range.y + input.range.height - 1))
  const colSegments = splitAxisRange(colStart, colEnd, input.camera.frozenColumnCount)
  const rowSegments = splitAxisRange(rowStart, rowEnd, input.camera.frozenRowCount)
  const rects: Rectangle[] = []
  for (const colSegment of colSegments) {
    for (const rowSegment of rowSegments) {
      const width = input.columns.span(colSegment.start, colSegment.end + 1)
      const height = input.rows.span(rowSegment.start, rowSegment.end + 1)
      if (width <= 0 || height <= 0) {
        continue
      }
      const worldX = input.columns.offsetOf(colSegment.start)
      const worldY = input.rows.offsetOf(rowSegment.start)
      const x = colSegment.frozen
        ? getBodyPaneX(input.camera) - input.camera.frozenWidth + worldX
        : getBodyPaneX(input.camera) + worldX - input.camera.bodyWorldX
      const y = rowSegment.frozen
        ? getBodyPaneY(input.camera) - input.camera.frozenHeight + worldY
        : getBodyPaneY(input.camera) + worldY - input.camera.bodyWorldY
      const clipped = clipRect(rect(x, y, width, height), resolvePaneClip(input.camera, colSegment.frozen, rowSegment.frozen))
      if (clipped) {
        rects.push(clipped)
      }
    }
  }
  return rects
}

function resolveEditorScreenRect(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly col: number
  readonly row: number
}): Rectangle | null {
  return resolveCellScreenRectForPane({
    camera: input.camera,
    col: input.col,
    columns: input.columns,
    paneKind: resolveCellPaneKind(input.camera, input.col, input.row),
    row: input.row,
    rows: input.rows,
  })
}

function resolveCellPaneKind(camera: GridCameraSnapshotV2, col: number, row: number): GridPaneKind {
  const frozenCol = col < camera.frozenColumnCount
  const frozenRow = row < camera.frozenRowCount
  if (frozenCol && frozenRow) {
    return 'frozen-cells'
  }
  if (frozenCol) {
    return 'frozen-columns'
  }
  if (frozenRow) {
    return 'frozen-rows'
  }
  return 'body'
}

function resolveResizeGuideScreenRect(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly state: GridResizeGuideState
}): Rectangle | null {
  if (input.state.kind === 'column') {
    const headerRect = resolveColumnHeaderScreenRect({ camera: input.camera, col: input.state.index, columns: input.columns })
    if (!headerRect) {
      return null
    }
    return rect(headerRect.x + headerRect.width - 1, 0, 1, getHostHeight(input.camera))
  }
  const headerRect = resolveRowHeaderScreenRect({ camera: input.camera, row: input.state.index, rows: input.rows })
  if (!headerRect) {
    return null
  }
  return rect(0, headerRect.y + headerRect.height - 1, getHostWidth(input.camera), 1)
}

function resolveFillHandleScreenRect(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
}): Rectangle | null {
  const rects = resolveRangeScreenRects(input)
  const anchor = rects.toSorted(
    (left, right) => right.y + right.height - (left.y + left.height) || right.x + right.width - (left.x + left.width),
  )[0]
  if (!anchor) {
    return null
  }
  const size = 8
  const handle = rect(anchor.x + anchor.width - size / 2, anchor.y + anchor.height - size / 2, size, size)
  return clipRect(handle, rect(0, 0, getHostWidth(input.camera), getHostHeight(input.camera)))
}

function splitAxisRange(
  start: number,
  end: number,
  frozenCount: number,
): readonly { readonly start: number; readonly end: number; readonly frozen: boolean }[] {
  const segments: Array<{ readonly start: number; readonly end: number; readonly frozen: boolean }> = []
  if (start < frozenCount) {
    segments.push({ start, end: Math.min(end, frozenCount - 1), frozen: true })
  }
  if (end >= frozenCount) {
    segments.push({ start: Math.max(start, frozenCount), end, frozen: false })
  }
  return segments
}

function resolvePaneClip(camera: GridCameraSnapshotV2, frozenCols: boolean, frozenRows: boolean): Rectangle {
  const bodyX = getBodyPaneX(camera)
  const bodyY = getBodyPaneY(camera)
  if (frozenCols && frozenRows) {
    return rect(bodyX - camera.frozenWidth, bodyY - camera.frozenHeight, camera.frozenWidth, camera.frozenHeight)
  }
  if (frozenCols) {
    return rect(bodyX - camera.frozenWidth, bodyY, camera.frozenWidth, camera.bodyViewportHeight)
  }
  if (frozenRows) {
    return rect(bodyX, bodyY - camera.frozenHeight, camera.bodyViewportWidth, camera.frozenHeight)
  }
  return rect(bodyX, bodyY, camera.bodyViewportWidth, camera.bodyViewportHeight)
}

function resolveColumnHeaderScreenRect(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly col: number
}): Rectangle | null {
  const width = input.columns.sizeOf(input.col)
  if (width <= 0 || input.columns.isHidden(input.col)) {
    return null
  }
  const worldX = input.columns.offsetOf(input.col)
  const rowHeaderWidth = getBodyPaneX(input.camera) - input.camera.frozenWidth
  const x =
    input.col < input.camera.frozenColumnCount ? rowHeaderWidth + worldX : getBodyPaneX(input.camera) + worldX - input.camera.bodyWorldX
  return rect(x, 0, width, getBodyPaneY(input.camera) - input.camera.frozenHeight)
}

function resolveRowHeaderScreenRect(input: {
  readonly camera: GridCameraSnapshotV2
  readonly rows: GridAxisWorldIndex
  readonly row: number
}): Rectangle | null {
  const height = input.rows.sizeOf(input.row)
  if (height <= 0 || input.rows.isHidden(input.row)) {
    return null
  }
  const worldY = input.rows.offsetOf(input.row)
  const columnHeaderHeight = getBodyPaneY(input.camera) - input.camera.frozenHeight
  const y =
    input.row < input.camera.frozenRowCount ? columnHeaderHeight + worldY : getBodyPaneY(input.camera) + worldY - input.camera.bodyWorldY
  return rect(0, y, getBodyPaneX(input.camera) - input.camera.frozenWidth, height)
}

function hitTestScreenPoint(input: {
  readonly camera: GridCameraSnapshotV2
  readonly columns: GridAxisWorldIndex
  readonly rows: GridAxisWorldIndex
  readonly point: { readonly x: number; readonly y: number }
}): { readonly col: number; readonly row: number } | null {
  const rowHeaderWidth = getBodyPaneX(input.camera) - input.camera.frozenWidth
  const columnHeaderHeight = getBodyPaneY(input.camera) - input.camera.frozenHeight
  const inFrozenColumns = input.point.x >= rowHeaderWidth && input.point.x < getBodyPaneX(input.camera)
  const inBodyColumns = input.point.x >= getBodyPaneX(input.camera)
  const inFrozenRows = input.point.y >= columnHeaderHeight && input.point.y < getBodyPaneY(input.camera)
  const inBodyRows = input.point.y >= getBodyPaneY(input.camera)
  if ((!inFrozenColumns && !inBodyColumns) || (!inFrozenRows && !inBodyRows)) {
    return null
  }
  const worldX = inFrozenColumns ? input.point.x - rowHeaderWidth : input.camera.bodyWorldX + input.point.x - getBodyPaneX(input.camera)
  const worldY = inFrozenRows ? input.point.y - columnHeaderHeight : input.camera.bodyWorldY + input.point.y - getBodyPaneY(input.camera)
  const col = input.columns.hitTest(worldX)
  const row = input.rows.hitTest(worldY)
  return col === null || row === null ? null : { col, row }
}

function getBodyPaneX(camera: GridCameraSnapshotV2): number {
  const bodyPane = camera.panes.find((pane) => pane.kind === 'body')
  return bodyPane?.frame.x ?? camera.frozenWidth
}

function getBodyPaneY(camera: GridCameraSnapshotV2): number {
  const bodyPane = camera.panes.find((pane) => pane.kind === 'body')
  return bodyPane?.frame.y ?? camera.frozenHeight
}

function getHostWidth(camera: GridCameraSnapshotV2): number {
  return camera.panes.reduce((max, pane) => Math.max(max, pane.frame.x + pane.frame.width), 0)
}

function getHostHeight(camera: GridCameraSnapshotV2): number {
  return camera.panes.reduce((max, pane) => Math.max(max, pane.frame.y + pane.frame.height), 0)
}

function rect(x: number, y: number, width: number, height: number): Rectangle {
  return { height: Math.max(0, height), width: Math.max(0, width), x, y }
}

function clipRect(target: Rectangle, clip: Rectangle): Rectangle | null {
  const x0 = Math.max(target.x, clip.x)
  const y0 = Math.max(target.y, clip.y)
  const x1 = Math.min(target.x + target.width, clip.x + clip.width)
  const y1 = Math.min(target.y + target.height, clip.y + clip.height)
  if (x1 <= x0 || y1 <= y0) {
    return null
  }
  return rect(x0, y0, x1 - x0, y1 - y0)
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value))
}
