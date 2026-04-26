import { formatAddress } from '@bilig/formula'
import type { WorkbookLocalViewportBase } from '@bilig/storage-browser'
import {
  ValueTag,
  formulaLooksDateLike,
  formatCellDisplayValue,
  isDateLikeHeaderValue,
  isLikelyExcelDateSerialValue,
  type CellSnapshot,
  type CellStyleRecord,
  type EngineEvent,
  type RecalcMetrics,
} from '@bilig/protocol'
import type { ViewportPatch, ViewportPatchedCell } from '@bilig/worker-transport'
import {
  buildAxisPatches,
  collectViewportCells,
  indexAxisEntries,
  styleSignature,
  type SheetViewportImpact,
  type ViewportSubscriptionState,
  type WorkerEngine,
} from './worker-runtime-support.js'

const PRODUCT_COLUMN_WIDTH = 104
const PRODUCT_ROW_HEIGHT = 22
export const DEFAULT_STYLE_ID = 'style-0'
export const MIN_COLUMN_WIDTH = 44
export const MAX_COLUMN_WIDTH = 480
export const AUTOFIT_PADDING = 28
export const AUTOFIT_CHAR_WIDTH = 8

interface PatchedCellContext {
  readonly state: ViewportSubscriptionState
  readonly styles: CellStyleRecord[]
  readonly cells: ViewportPatchedCell[]
  readonly getStyleRecord: (styleId: string) => CellStyleRecord
  readonly getFormatId: (format: string | undefined) => number
}

function snapshotValueSignature(snapshot: CellSnapshot): string {
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return `n:${snapshot.value.value}`
    case ValueTag.Boolean:
      return `b:${snapshot.value.value ? 1 : 0}`
    case ValueTag.String:
      return `s:${snapshot.value.stringId}:${snapshot.value.value}`
    case ValueTag.Error:
      return `e:${snapshot.value.code}`
    case ValueTag.Empty:
      return 'empty'
  }
}

function toEditorText(snapshot: CellSnapshot): string {
  if (snapshot.formula) {
    return `=${snapshot.formula}`
  }
  if (snapshot.input === null || snapshot.input === undefined) {
    return toDisplayText(snapshot)
  }
  if (typeof snapshot.input === 'boolean') {
    return snapshot.input ? 'TRUE' : 'FALSE'
  }
  return String(snapshot.input)
}

function toDisplayText(snapshot: CellSnapshot): string {
  return formatCellDisplayValue(snapshot.value, snapshot.format)
}

function inferLocalProjectionFormat(
  snapshot: CellSnapshot,
  row: number,
  col: number,
  cellsByPosition: ReadonlyMap<string, CellSnapshot>,
): string | undefined {
  if (snapshot.format !== undefined || !isLikelyExcelDateSerialValue(snapshot.value)) {
    return snapshot.format
  }
  if (formulaLooksDateLike(snapshot.formula)) {
    return 'date:short'
  }
  const header = cellsByPosition.get(`${row - 1}:${col}`)
  return header !== undefined && isDateLikeHeaderValue(header.value) ? 'date:short' : snapshot.format
}

function buildPatchedCellSignature(
  snapshot: CellSnapshot,
  displayText: string,
  copyText: string,
  editorText: string,
  formatId: number,
  styleId: string,
): string {
  return [
    snapshot.version,
    snapshot.flags,
    snapshot.formula ?? '',
    snapshot.input ?? '',
    snapshot.format ?? '',
    snapshot.styleId ?? '',
    formatId,
    styleId,
    snapshotValueSignature(snapshot),
    displayText,
    copyText,
    editorText,
  ].join('|')
}

function appendPatchedCell(context: PatchedCellContext, row: number, col: number, snapshot: CellSnapshot, force: boolean): void {
  const key = `${snapshot.sheetName}!${snapshot.address}`
  const formatId = context.getFormatId(snapshot.format)
  const style = context.getStyleRecord(snapshot.styleId ?? DEFAULT_STYLE_ID)
  const nextStyleSignature = styleSignature(style)
  const previousStyleSignature = context.state.lastStyleSignatures.get(style.id)
  if (force || previousStyleSignature !== nextStyleSignature || !context.state.knownStyleIds.has(style.id)) {
    context.state.knownStyleIds.add(style.id)
    context.state.lastStyleSignatures.set(style.id, nextStyleSignature)
    context.styles.push(style)
  }
  const editorText = toEditorText(snapshot)
  const displayText = toDisplayText(snapshot)
  const copyText = snapshot.formula ? editorText : displayText
  const signature = buildPatchedCellSignature(snapshot, displayText, copyText, editorText, formatId, style.id)
  if (force || context.state.lastCellSignatures.get(key) !== signature) {
    context.cells.push({
      row,
      col,
      snapshot,
      displayText,
      copyText,
      editorText,
      formatId,
      styleId: style.id,
    })
  }
  context.state.lastCellSignatures.set(key, signature)
}

export function buildViewportPatchFromEngine(input: {
  readonly state: ViewportSubscriptionState
  readonly event: EngineEvent | null
  readonly metrics: RecalcMetrics
  readonly sheetImpact: SheetViewportImpact | null
  readonly engine: WorkerEngine
  readonly emptyCellSnapshot: (sheetName: string, address: string) => CellSnapshot
  readonly getStyleRecord: (styleId: string) => CellStyleRecord
  readonly getFormatId: (format: string | undefined) => number
}): ViewportPatch {
  const { state, event, metrics, sheetImpact, engine } = input
  const viewport = state.subscription
  const hasSheet = engine.workbook.getSheet(viewport.sheetName) !== undefined
  const styles: CellStyleRecord[] = []
  const cells: ViewportPatchedCell[] = []
  const full = event === null || event.invalidation === 'full'
  const invalidatedRanges = sheetImpact?.invalidatedRanges ?? []
  const invalidatedRows = sheetImpact?.invalidatedRows ?? []
  const invalidatedColumns = sheetImpact?.invalidatedColumns ?? []
  const context: PatchedCellContext = {
    state,
    styles,
    cells,
    getStyleRecord: input.getStyleRecord,
    getFormatId: input.getFormatId,
  }

  if (full) {
    state.lastCellSignatures.clear()
    state.lastStyleSignatures.clear()
    for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
      for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
        const address = formatAddress(row, col)
        const snapshot = hasSheet ? engine.getCell(viewport.sheetName, address) : input.emptyCellSnapshot(viewport.sheetName, address)
        appendPatchedCell(context, row, col, snapshot, true)
      }
    }
  } else {
    const targetCells = collectViewportCells(viewport, sheetImpact?.changedCells ?? null, invalidatedRanges)
    for (const cell of targetCells) {
      appendPatchedCell(
        context,
        cell.row,
        cell.col,
        hasSheet ? engine.getCell(viewport.sheetName, cell.address) : input.emptyCellSnapshot(viewport.sheetName, cell.address),
        false,
      )
    }
  }

  const { patches: columns, signatures: columnSignatures } = buildAxisPatches(
    viewport.colStart,
    viewport.colEnd,
    indexAxisEntries(engine.getColumnAxisEntries(viewport.sheetName)),
    PRODUCT_COLUMN_WIDTH,
    state.lastColumnSignatures,
    full,
    invalidatedColumns,
  )
  const { patches: rows, signatures: rowSignatures } = buildAxisPatches(
    viewport.rowStart,
    viewport.rowEnd,
    indexAxisEntries(engine.getRowAxisEntries(viewport.sheetName)),
    PRODUCT_ROW_HEIGHT,
    state.lastRowSignatures,
    full,
    invalidatedRows,
  )
  state.lastColumnSignatures = columnSignatures
  state.lastRowSignatures = rowSignatures

  return {
    version: state.nextVersion++,
    full,
    freezeRows: engine.getFreezePane(viewport.sheetName)?.rows ?? 0,
    freezeCols: engine.getFreezePane(viewport.sheetName)?.cols ?? 0,
    viewport,
    metrics: { ...metrics },
    styles,
    cells,
    columns,
    rows,
  }
}

export function buildViewportPatchFromLocalBase(input: {
  readonly state: ViewportSubscriptionState
  readonly metrics: RecalcMetrics
  readonly base: WorkbookLocalViewportBase
  readonly getFormatId: (format: string | undefined) => number
}): ViewportPatch {
  const { state, metrics, base } = input
  const viewport = state.subscription
  state.lastCellSignatures.clear()
  state.lastStyleSignatures.clear()

  const styles = [...base.styles]
  const cellsByPosition = new Map(base.cells.map((cell) => [`${cell.row}:${cell.col}`, cell.snapshot]))
  styles.forEach((style) => {
    state.knownStyleIds.add(style.id)
    state.lastStyleSignatures.set(style.id, styleSignature(style))
  })

  const cells: ViewportPatchedCell[] = []
  for (const cell of base.cells) {
    const inferredFormat = inferLocalProjectionFormat(cell.snapshot, cell.row, cell.col, cellsByPosition)
    const snapshot: CellSnapshot =
      inferredFormat === undefined || inferredFormat === cell.snapshot.format ? cell.snapshot : { ...cell.snapshot, format: inferredFormat }
    const editorText = toEditorText(snapshot)
    const displayText = toDisplayText(snapshot)
    const copyText = snapshot.formula ? editorText : displayText
    const formatId = input.getFormatId(snapshot.format)
    const styleId = snapshot.styleId ?? DEFAULT_STYLE_ID
    cells.push({
      row: cell.row,
      col: cell.col,
      snapshot,
      displayText,
      copyText,
      editorText,
      formatId,
      styleId,
    })
    state.lastCellSignatures.set(
      `${snapshot.sheetName}!${snapshot.address}`,
      buildPatchedCellSignature(snapshot, displayText, copyText, editorText, formatId, styleId),
    )
  }

  const { patches: columns, signatures: columnSignatures } = buildAxisPatches(
    viewport.colStart,
    viewport.colEnd,
    indexAxisEntries(base.columnAxisEntries),
    PRODUCT_COLUMN_WIDTH,
    state.lastColumnSignatures,
    true,
  )
  const { patches: rows, signatures: rowSignatures } = buildAxisPatches(
    viewport.rowStart,
    viewport.rowEnd,
    indexAxisEntries(base.rowAxisEntries),
    PRODUCT_ROW_HEIGHT,
    state.lastRowSignatures,
    true,
  )
  state.lastColumnSignatures = columnSignatures
  state.lastRowSignatures = rowSignatures

  return {
    version: state.nextVersion++,
    full: true,
    freezeRows: base.freezeRows,
    freezeCols: base.freezeCols,
    viewport,
    metrics: { ...metrics },
    styles,
    cells,
    columns,
    rows,
  }
}
