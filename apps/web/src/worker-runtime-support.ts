import type { CommitOp, EngineReplicaSnapshot } from '@bilig/core'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  type CellNumberFormatInput,
  type CellRangeRef,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  type CellSnapshot,
  type EngineEvent,
  type LiteralInput,
  MAX_COLS,
  MAX_ROWS,
  type RecalcMetrics,
  type SyncState,
  type WorkbookAxisEntrySnapshot,
  type WorkbookFreezePaneSnapshot,
  type WorkbookSnapshot,
} from '@bilig/protocol'
import type { ViewportAxisPatch, ViewportPatchSubscription } from '@bilig/worker-transport'
import type { PendingWorkbookMutation } from './workbook-sync.js'

export interface WorkerSheet {
  name: string
  order: number
  grid: {
    forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void): void
  }
}

export interface WorkerWorkbook {
  workbookName: string
  cellStore: {
    sheetIds: Uint16Array
    rows: Uint32Array
    cols: Uint16Array
  }
  sheetsByName: Map<string, WorkerSheet>
  getSheet(sheetName: string): WorkerSheet | undefined
  getSheetNameById(sheetId: number): string
  getQualifiedAddress(cellIndex: number): string
}

export interface WorkerEngine {
  workbook: WorkerWorkbook
  ready(): Promise<void>
  createSheet(name: string): void
  subscribe(listener: (event: EngineEvent) => void): () => void
  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void
  getLastMetrics(): RecalcMetrics
  getSyncState(): SyncState
  getCell(sheetName: string, address: string): CellSnapshot
  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined
  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void
  clearRangeNumberFormat(range: CellRangeRef): void
  clearRange(range: CellRangeRef): void
  setCellValue(sheetName: string, address: string, value: LiteralInput): unknown
  setCellFormula(sheetName: string, address: string, formula: string): unknown
  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void
  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void
  clearCell(sheetName: string, address: string): void
  renderCommit(ops: CommitOp[]): void
  fillRange(source: CellRangeRef, target: CellRangeRef): void
  copyRange(source: CellRangeRef, target: CellRangeRef): void
  moveRange(source: CellRangeRef, target: CellRangeRef): void
  insertRows(sheetName: string, start: number, count: number): void
  deleteRows(sheetName: string, start: number, count: number): void
  insertColumns(sheetName: string, start: number, count: number): void
  deleteColumns(sheetName: string, start: number, count: number): void
  updateRowMetadata(sheetName: string, start: number, count: number, size: number | null, hidden: boolean | null): unknown
  updateColumnMetadata(sheetName: string, start: number, count: number, size: number | null, hidden: boolean | null): unknown
  setFreezePane(sheetName: string, rows: number, cols: number): unknown
  getFreezePane(sheetName: string): WorkbookFreezePaneSnapshot | undefined
  exportSnapshot(): WorkbookSnapshot
  exportReplicaSnapshot(): EngineReplicaSnapshot
  importSnapshot(snapshot: WorkbookSnapshot): void
  importReplicaSnapshot(snapshot: EngineReplicaSnapshot): void
  getColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[]
  getRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[]
}

export interface PersistedWorkbookState {
  snapshot: WorkbookSnapshot
  replica: EngineReplicaSnapshot
}

export interface PersistedPendingMutationState {
  pendingMutations: PendingWorkbookMutation[]
}

export interface ViewportSubscriptionState {
  subscription: ViewportPatchSubscription
  listener: (patch: Uint8Array) => void
  nextVersion: number
  knownStyleIds: Set<string>
  lastStyleSignatures: Map<string, string>
  lastCellSignatures: Map<string, string>
  lastColumnSignatures: Map<number, string>
  lastRowSignatures: Map<number, string>
}

export interface ViewportCellPosition {
  address: string
  row: number
  col: number
}

export interface ChangedSheetCells {
  addresses: Set<string>
  positions: ViewportCellPosition[]
}

export interface NormalizedRangeImpact {
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

export interface SheetViewportImpact {
  changedCells: ChangedSheetCells | null
  invalidatedRanges: NormalizedRangeImpact[]
  invalidatedRows: { sheetName: string; startIndex: number; endIndex: number }[]
  invalidatedColumns: { sheetName: string; startIndex: number; endIndex: number }[]
}

export function styleSignature(style: CellStyleRecord): string {
  const fill = style.fill?.backgroundColor ?? ''
  const font = style.font
  const alignment = style.alignment
  const borders = style.borders
  return [
    fill,
    font?.family ?? '',
    font?.size ?? '',
    font?.bold ? 1 : 0,
    font?.italic ? 1 : 0,
    font?.underline ? 1 : 0,
    font?.color ?? '',
    alignment?.horizontal ?? '',
    alignment?.vertical ?? '',
    alignment?.wrap ? 1 : 0,
    alignment?.indent ?? '',
    borders?.top ? `${borders.top.style}:${borders.top.weight}:${borders.top.color}` : '',
    borders?.right ? `${borders.right.style}:${borders.right.weight}:${borders.right.color}` : '',
    borders?.bottom ? `${borders.bottom.style}:${borders.bottom.weight}:${borders.bottom.color}` : '',
    borders?.left ? `${borders.left.style}:${borders.left.weight}:${borders.left.color}` : '',
  ].join('|')
}

export function buildAxisPatches(
  start: number,
  end: number,
  entries: Map<number, WorkbookAxisEntrySnapshot>,
  defaultSize: number,
  previous: Map<number, string>,
  full: boolean,
  invalidatedAxes: readonly { startIndex: number; endIndex: number }[] = [],
): { patches: ViewportAxisPatch[]; signatures: Map<number, string> } {
  if (!full && invalidatedAxes.length === 0) {
    return { patches: [], signatures: previous }
  }
  const signatures = full ? new Map<number, string>() : new Map(previous)
  const patches: ViewportAxisPatch[] = []
  const indices = full ? collectAxisIndices(start, end, null) : collectAxisIndices(start, end, invalidatedAxes)
  for (const index of indices) {
    const entry = entries.get(index)
    const size = entry?.size ?? defaultSize
    const hidden = entry?.hidden ?? false
    const signature = `${size}:${hidden ? 1 : 0}`
    signatures.set(index, signature)
    if (full || previous.get(index) !== signature) {
      patches.push({ index, size, hidden })
    }
  }
  return { patches, signatures }
}

export function indexAxisEntries(entries: readonly WorkbookAxisEntrySnapshot[]): Map<number, WorkbookAxisEntrySnapshot> {
  return new Map(entries.map((entry) => [entry.index, entry]))
}

export function collectViewportCells(
  viewport: ViewportPatchSubscription,
  changedCells: ChangedSheetCells | null,
  invalidatedRanges: readonly NormalizedRangeImpact[],
): ViewportCellPosition[] {
  const positions: ViewportCellPosition[] = []
  const seen = new Set<string>()

  changedCells?.positions.forEach((cell) => {
    if (cell.row < viewport.rowStart || cell.row > viewport.rowEnd || cell.col < viewport.colStart || cell.col > viewport.colEnd) {
      return
    }
    const key = `${cell.row}:${cell.col}`
    if (seen.has(key)) {
      return
    }
    seen.add(key)
    positions.push(cell)
  })

  for (let index = 0; index < invalidatedRanges.length; index += 1) {
    const range = invalidatedRanges[index]!
    const rowStart = Math.max(viewport.rowStart, range.rowStart)
    const rowEnd = Math.min(viewport.rowEnd, range.rowEnd)
    const colStart = Math.max(viewport.colStart, range.colStart)
    const colEnd = Math.min(viewport.colEnd, range.colEnd)
    if (rowStart > rowEnd || colStart > colEnd) {
      continue
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        const key = `${row}:${col}`
        if (seen.has(key)) {
          continue
        }
        seen.add(key)
        positions.push({ address: formatAddress(row, col), row, col })
      }
    }
  }
  return positions
}

function collectAxisIndices(
  start: number,
  end: number,
  invalidatedAxes: readonly { startIndex: number; endIndex: number }[] | null,
): number[] {
  if (invalidatedAxes === null) {
    const indices: number[] = []
    for (let index = start; index <= end; index += 1) {
      indices.push(index)
    }
    return indices
  }

  const indices = new Set<number>()
  for (let axisIndex = 0; axisIndex < invalidatedAxes.length; axisIndex += 1) {
    const axis = invalidatedAxes[axisIndex]!
    const clampedStart = Math.max(start, axis.startIndex)
    const clampedEnd = Math.min(end, axis.endIndex)
    if (clampedStart > clampedEnd) {
      continue
    }
    for (let index = clampedStart; index <= clampedEnd; index += 1) {
      indices.add(index)
    }
  }
  return Array.from(indices).toSorted((left, right) => left - right)
}

export function normalizeViewport(subscription: ViewportPatchSubscription): ViewportPatchSubscription {
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, subscription.rowStart))
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, subscription.rowEnd))
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, subscription.colStart))
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, subscription.colEnd))
  return {
    sheetName: subscription.sheetName,
    rowStart,
    rowEnd,
    colStart,
    colEnd,
    ...(subscription.initialPatch === 'none' ? { initialPatch: 'none' as const } : {}),
  }
}

export function collectChangedCellsBySheet(
  engine: WorkerEngine,
  changedCellIndices: readonly number[] | Uint32Array,
): Map<string, ChangedSheetCells> {
  const changedBySheet = new Map<string, ChangedSheetCells>()
  const workbook = engine.workbook
  const { cellStore } = workbook

  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined || sheetId === 0) {
      continue
    }
    const sheetName = workbook.getSheetNameById(sheetId)
    if (sheetName.length === 0) {
      continue
    }
    const row = cellStore.rows[cellIndex] ?? 0
    const col = cellStore.cols[cellIndex] ?? 0
    const address = formatAddress(row, col)
    const sheetCells = changedBySheet.get(sheetName) ?? {
      addresses: new Set<string>(),
      positions: [],
    }
    if (!sheetCells.addresses.has(address)) {
      sheetCells.addresses.add(address)
      sheetCells.positions.push({ address, row, col })
    }
    changedBySheet.set(sheetName, sheetCells)
  }

  return changedBySheet
}

export function collectSheetViewportImpacts(engine: WorkerEngine, event: EngineEvent): Map<string, SheetViewportImpact> | null {
  const changedCellsBySheet = event.invalidation !== 'full' ? collectChangedCellsBySheet(engine, event.changedCellIndices) : null
  const impactsBySheet = new Map<string, SheetViewportImpact>()

  changedCellsBySheet?.forEach((changedCells, sheetName) => {
    impactsBySheet.set(sheetName, {
      changedCells,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
    })
  })

  event.invalidatedRanges.forEach((range) => {
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    const impact = impactsBySheet.get(range.sheetName) ?? {
      changedCells: null,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
    }
    impact.invalidatedRanges.push({
      rowStart: Math.min(start.row, end.row),
      rowEnd: Math.max(start.row, end.row),
      colStart: Math.min(start.col, end.col),
      colEnd: Math.max(start.col, end.col),
    })
    impactsBySheet.set(range.sheetName, impact)
  })

  event.invalidatedRows.forEach((entry) => {
    const impact = impactsBySheet.get(entry.sheetName) ?? {
      changedCells: null,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
    }
    impact.invalidatedRows.push(entry)
    impactsBySheet.set(entry.sheetName, impact)
  })

  event.invalidatedColumns.forEach((entry) => {
    const impact = impactsBySheet.get(entry.sheetName) ?? {
      changedCells: null,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
    }
    impact.invalidatedColumns.push(entry)
    impactsBySheet.set(entry.sheetName, impact)
  })

  return impactsBySheet.size > 0 ? impactsBySheet : null
}

export function viewportPatchMayBeImpacted(
  viewport: ViewportPatchSubscription,
  event: EngineEvent,
  sheetImpact: SheetViewportImpact | null,
  impactedSheets: ReadonlySet<string> | null,
): boolean {
  if (event.invalidation === 'full') {
    return impactedSheets === null || impactedSheets.has(viewport.sheetName)
  }

  if (impactedSheets !== null && !impactedSheets.has(viewport.sheetName)) {
    return false
  }

  const changedCells = sheetImpact?.changedCells
  if (changedCells) {
    for (const parsed of changedCells.positions) {
      if (
        parsed.row >= viewport.rowStart &&
        parsed.row <= viewport.rowEnd &&
        parsed.col >= viewport.colStart &&
        parsed.col <= viewport.colEnd
      ) {
        return true
      }
    }
  }

  for (let index = 0; index < (sheetImpact?.invalidatedRanges.length ?? 0); index += 1) {
    const range = sheetImpact!.invalidatedRanges[index]!
    if (
      range.rowStart <= viewport.rowEnd &&
      range.rowEnd >= viewport.rowStart &&
      range.colStart <= viewport.colEnd &&
      range.colEnd >= viewport.colStart
    ) {
      return true
    }
  }

  for (let index = 0; index < (sheetImpact?.invalidatedRows.length ?? 0); index += 1) {
    const rowInvalidation = sheetImpact!.invalidatedRows[index]!
    if (rowInvalidation.startIndex <= viewport.rowEnd && rowInvalidation.endIndex >= viewport.rowStart) {
      return true
    }
  }

  for (let index = 0; index < (sheetImpact?.invalidatedColumns.length ?? 0); index += 1) {
    const columnInvalidation = sheetImpact!.invalidatedColumns[index]!
    if (columnInvalidation.startIndex <= viewport.colEnd && columnInvalidation.endIndex >= viewport.colStart) {
      return true
    }
  }

  return false
}
