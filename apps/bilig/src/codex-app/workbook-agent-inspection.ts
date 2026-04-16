import { formatAddress, parseCellAddress } from '@bilig/formula'
import {
  formatErrorCode,
  ValueTag,
  type CellNumberFormatRecord,
  type CellRangeRef,
  type CellStyleRecord,
  type WorkbookCommentThreadSnapshot,
  type WorkbookConditionalFormatSnapshot,
  type WorkbookDataValidationSnapshot,
  type WorkbookImageSnapshot,
  type WorkbookNoteSnapshot,
  type WorkbookRangeProtectionSnapshot,
  type WorkbookSheetProtectionSnapshot,
  type WorkbookShapeSnapshot,
} from '@bilig/protocol'
import type { JsonValue } from '@bilig/agent-api'
import type { WorkbookAgentUiContext, WorkbookViewport } from '@bilig/contracts'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'

function formatColumnLabel(index: number): string {
  return formatAddress(0, index).replace(/[0-9]/g, '')
}

function serializeCellValue(value: { tag: ValueTag; value?: number | boolean | string; code?: number }): JsonValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null
    case ValueTag.Error:
      return typeof value.code === 'number' ? formatErrorCode(value.code) : '#ERROR!'
    default:
      return null
  }
}

function viewportToRange(sheetName: string, viewport: WorkbookViewport): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(viewport.rowStart, viewport.colStart),
    endAddress: formatAddress(viewport.rowEnd, viewport.colEnd),
  }
}

function normalizeRange(range: CellRangeRef): CellRangeRef & {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    ...range,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

function rangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  const leftBounds = normalizeRange(left)
  const rightBounds = normalizeRange(right)
  return !(
    leftBounds.sheetName !== rightBounds.sheetName ||
    leftBounds.endRow < rightBounds.startRow ||
    rightBounds.endRow < leftBounds.startRow ||
    leftBounds.endCol < rightBounds.startCol ||
    rightBounds.endCol < leftBounds.startCol
  )
}

function collectIntersectingDataValidations(runtime: WorkbookRuntime, range: CellRangeRef): readonly WorkbookDataValidationSnapshot[] {
  return runtime.engine
    .getDataValidations(range.sheetName)
    .filter((validation) => rangesIntersect(validation.range, range))
    .map((validation) => structuredClone(validation))
}

function collectIntersectingConditionalFormats(
  runtime: WorkbookRuntime,
  range: CellRangeRef,
): readonly WorkbookConditionalFormatSnapshot[] {
  return runtime.engine
    .getConditionalFormats(range.sheetName)
    .filter((format) => rangesIntersect(format.range, range))
    .map((format) => structuredClone(format))
}

function collectIntersectingCommentThreads(runtime: WorkbookRuntime, range: CellRangeRef): readonly WorkbookCommentThreadSnapshot[] {
  return runtime.engine
    .getCommentThreads(range.sheetName)
    .filter((thread) =>
      rangesIntersect(range, {
        sheetName: thread.sheetName,
        startAddress: thread.address,
        endAddress: thread.address,
      }),
    )
    .map((thread) => structuredClone(thread))
}

function collectIntersectingNotes(runtime: WorkbookRuntime, range: CellRangeRef): readonly WorkbookNoteSnapshot[] {
  return runtime.engine
    .getNotes(range.sheetName)
    .filter((note) =>
      rangesIntersect(range, {
        sheetName: note.sheetName,
        startAddress: note.address,
        endAddress: note.address,
      }),
    )
    .map((note) => structuredClone(note))
}

function collectIntersectingImages(runtime: WorkbookRuntime, range: CellRangeRef): readonly WorkbookImageSnapshot[] {
  return runtime.engine
    .getImages()
    .filter((image) => {
      const anchor = parseCellAddress(image.address, image.sheetName)
      return rangesIntersect(range, {
        sheetName: image.sheetName,
        startAddress: image.address,
        endAddress: formatAddress(anchor.row + Math.max(0, image.rows - 1), anchor.col + Math.max(0, image.cols - 1)),
      })
    })
    .map((image) => structuredClone(image))
}

function collectIntersectingShapes(runtime: WorkbookRuntime, range: CellRangeRef): readonly WorkbookShapeSnapshot[] {
  return runtime.engine
    .getShapes()
    .filter((shape) => {
      const anchor = parseCellAddress(shape.address, shape.sheetName)
      return rangesIntersect(range, {
        sheetName: shape.sheetName,
        startAddress: shape.address,
        endAddress: formatAddress(anchor.row + Math.max(0, shape.rows - 1), anchor.col + Math.max(0, shape.cols - 1)),
      })
    })
    .map((shape) => structuredClone(shape))
}

function collectIntersectingRangeProtections(runtime: WorkbookRuntime, range: CellRangeRef): readonly WorkbookRangeProtectionSnapshot[] {
  return runtime.engine
    .getRangeProtections(range.sheetName)
    .filter((protection) => rangesIntersect(protection.range, range))
    .map((protection) => structuredClone(protection))
}

function sheetProtection(runtime: WorkbookRuntime, sheetName: string): WorkbookSheetProtectionSnapshot | null {
  return runtime.engine.getSheetProtection(sheetName) ?? null
}

function isFormulaHidden(runtime: WorkbookRuntime, sheetName: string, address: string): boolean {
  if (runtime.engine.getSheetProtection(sheetName)?.hideFormulas === true) {
    return true
  }
  return runtime.engine.getRangeProtections(sheetName).some(
    (protection) =>
      protection.hideFormulas === true &&
      rangesIntersect(protection.range, {
        sheetName,
        startAddress: address,
        endAddress: address,
      }),
  )
}

function summarizeWindowAxisState(input: { readonly range: ReturnType<typeof normalizeRange>; readonly runtime: WorkbookRuntime }) {
  const rowEntries = input.runtime.engine.getRowAxisEntries(input.range.sheetName)
  const columnEntries = input.runtime.engine.getColumnAxisEntries(input.range.sheetName)
  return {
    freezePane: input.runtime.engine.getFreezePane(input.range.sheetName) ?? null,
    hiddenRows: rowEntries
      .filter((entry) => entry.hidden === true && entry.index >= input.range.startRow && entry.index <= input.range.endRow)
      .map((entry) => ({
        rowNumber: entry.index + 1,
      })),
    hiddenColumns: columnEntries
      .filter((entry) => entry.hidden === true && entry.index >= input.range.startCol && entry.index <= input.range.endCol)
      .map((entry) => ({
        columnIndex: entry.index,
        columnLabel: formatColumnLabel(entry.index),
      })),
    resizedRows: rowEntries
      .filter((entry) => typeof entry.size === 'number' && entry.index >= input.range.startRow && entry.index <= input.range.endRow)
      .map((entry) => ({
        rowNumber: entry.index + 1,
        size: entry.size!,
      })),
    resizedColumns: columnEntries
      .filter((entry) => typeof entry.size === 'number' && entry.index >= input.range.startCol && entry.index <= input.range.endCol)
      .map((entry) => ({
        columnIndex: entry.index,
        columnLabel: formatColumnLabel(entry.index),
        size: entry.size!,
      })),
  }
}

function summarizeSelection(context: WorkbookAgentUiContext) {
  const selectionRange = normalizeRange({
    sheetName: context.selection.sheetName,
    startAddress: context.selection.range?.startAddress ?? context.selection.address,
    endAddress: context.selection.range?.endAddress ?? context.selection.address,
  })
  return {
    ...context.selection,
    kind: selectionRange.startAddress === selectionRange.endAddress ? 'cell' : 'range',
    startAddress: selectionRange.startAddress,
    endAddress: selectionRange.endAddress,
    rowCount: selectionRange.endRow - selectionRange.startRow + 1,
    columnCount: selectionRange.endCol - selectionRange.startCol + 1,
    cellCount: (selectionRange.endRow - selectionRange.startRow + 1) * (selectionRange.endCol - selectionRange.startCol + 1),
  }
}

function collectRangeFormattingCatalog(input: {
  readonly runtime: WorkbookRuntime
  readonly styleIds: ReadonlySet<string>
  readonly numberFormatIds: ReadonlySet<string>
}): {
  readonly styles: readonly CellStyleRecord[]
  readonly numberFormats: readonly CellNumberFormatRecord[]
} {
  return {
    styles: [...input.styleIds]
      .map((styleId) => input.runtime.engine.getCellStyle(styleId))
      .filter((style): style is CellStyleRecord => style !== undefined)
      .toSorted((left, right) => left.id.localeCompare(right.id)),
    numberFormats: [...input.numberFormatIds]
      .map((formatId) => input.runtime.engine.getCellNumberFormat(formatId))
      .filter((format): format is CellNumberFormatRecord => format !== undefined)
      .toSorted((left, right) => left.id.localeCompare(right.id)),
  }
}

export function inspectWorkbookContext(runtime: WorkbookRuntime, context: WorkbookAgentUiContext | null): string {
  if (!context) {
    return 'No browser view context is attached to this chat session yet.'
  }
  const selection = summarizeSelection(context)
  const visibleRange = normalizeRange(viewportToRange(context.selection.sheetName, context.viewport))
  return JSON.stringify(
    {
      selection,
      visibleRange: {
        sheetName: visibleRange.sheetName,
        startAddress: visibleRange.startAddress,
        endAddress: visibleRange.endAddress,
        rowCount: visibleRange.endRow - visibleRange.startRow + 1,
        columnCount: visibleRange.endCol - visibleRange.startCol + 1,
        cellCount: (visibleRange.endRow - visibleRange.startRow + 1) * (visibleRange.endCol - visibleRange.startCol + 1),
      },
      sheetState: summarizeWindowAxisState({
        runtime,
        range: visibleRange,
      }),
    },
    null,
    2,
  )
}

export function inspectWorkbookRange(
  runtime: WorkbookRuntime,
  range: CellRangeRef,
): {
  readonly range: CellRangeRef
  readonly sheetState: ReturnType<typeof summarizeWindowAxisState>
  readonly sheetProtection: WorkbookSheetProtectionSnapshot | null
  readonly rangeProtections: readonly WorkbookRangeProtectionSnapshot[]
  readonly dataValidations: readonly WorkbookDataValidationSnapshot[]
  readonly conditionalFormats: readonly WorkbookConditionalFormatSnapshot[]
  readonly commentThreads: readonly WorkbookCommentThreadSnapshot[]
  readonly notes: readonly WorkbookNoteSnapshot[]
  readonly images: readonly WorkbookImageSnapshot[]
  readonly shapes: readonly WorkbookShapeSnapshot[]
  readonly styles: readonly CellStyleRecord[]
  readonly numberFormats: readonly CellNumberFormatRecord[]
  readonly rows: readonly JsonValue[]
} {
  const normalizedRange = normalizeRange(range)
  const styleIds = new Set<string>()
  const numberFormatIds = new Set<string>()
  const rows: JsonValue[] = []
  for (let row = normalizedRange.startRow; row <= normalizedRange.endRow; row += 1) {
    const rowEntries: JsonValue[] = []
    for (let col = normalizedRange.startCol; col <= normalizedRange.endCol; col += 1) {
      const cell = runtime.engine.getCell(normalizedRange.sheetName, formatAddress(row, col))
      if (cell.styleId) {
        styleIds.add(cell.styleId)
      }
      if (cell.numberFormatId) {
        numberFormatIds.add(cell.numberFormatId)
      }
      rowEntries.push({
        address: cell.address,
        input: cell.input ?? null,
        value: serializeCellValue(cell.value),
        formula:
          cell.formula !== undefined && !isFormulaHidden(runtime, normalizedRange.sheetName, cell.address) ? `=${cell.formula}` : null,
        displayFormat: cell.format ?? null,
        styleId: cell.styleId ?? null,
        numberFormatId: cell.numberFormatId ?? null,
      })
    }
    rows.push(rowEntries)
  }

  return {
    range: {
      sheetName: normalizedRange.sheetName,
      startAddress: normalizedRange.startAddress,
      endAddress: normalizedRange.endAddress,
    },
    sheetState: summarizeWindowAxisState({
      runtime,
      range: normalizedRange,
    }),
    sheetProtection: sheetProtection(runtime, normalizedRange.sheetName),
    rangeProtections: collectIntersectingRangeProtections(runtime, normalizedRange),
    dataValidations: collectIntersectingDataValidations(runtime, normalizedRange),
    conditionalFormats: collectIntersectingConditionalFormats(runtime, normalizedRange),
    commentThreads: collectIntersectingCommentThreads(runtime, normalizedRange),
    notes: collectIntersectingNotes(runtime, normalizedRange),
    images: collectIntersectingImages(runtime, normalizedRange),
    shapes: collectIntersectingShapes(runtime, normalizedRange),
    ...collectRangeFormattingCatalog({
      runtime,
      styleIds,
      numberFormatIds,
    }),
    rows,
  }
}

export function inspectWorkbookCell(
  runtime: WorkbookRuntime,
  target: {
    sheetName: string
    address: string
  },
): {
  readonly sheetName: string
  readonly address: string
  readonly input: unknown
  readonly value: JsonValue
  readonly formula: string | null
  readonly displayFormat: string | null
  readonly styleId: string | null
  readonly style: CellStyleRecord | null
  readonly numberFormatId: string | null
  readonly numberFormat: CellNumberFormatRecord | null
  readonly version: number
  readonly inCycle: boolean
  readonly mode: unknown
  readonly topoRank: number | null
  readonly directPrecedents: readonly string[]
  readonly directDependents: readonly string[]
  readonly sheetProtection: WorkbookSheetProtectionSnapshot | null
  readonly rangeProtections: readonly WorkbookRangeProtectionSnapshot[]
  readonly dataValidations: readonly WorkbookDataValidationSnapshot[]
  readonly conditionalFormats: readonly WorkbookConditionalFormatSnapshot[]
  readonly commentThreads: readonly WorkbookCommentThreadSnapshot[]
  readonly notes: readonly WorkbookNoteSnapshot[]
  readonly images: readonly WorkbookImageSnapshot[]
  readonly shapes: readonly WorkbookShapeSnapshot[]
} {
  const snapshot = runtime.engine.getCell(target.sheetName, target.address)
  const cell = runtime.engine.explainCell(target.sheetName, target.address)
  return {
    sheetName: cell.sheetName,
    address: cell.address,
    input: snapshot.input ?? null,
    value: serializeCellValue(cell.value),
    formula: cell.formula !== undefined && !isFormulaHidden(runtime, target.sheetName, target.address) ? `=${cell.formula}` : null,
    displayFormat: cell.format ?? null,
    styleId: cell.styleId ?? null,
    style: runtime.engine.getCellStyle(cell.styleId) ?? null,
    numberFormatId: cell.numberFormatId ?? null,
    numberFormat: runtime.engine.getCellNumberFormat(cell.numberFormatId) ?? null,
    version: cell.version,
    inCycle: cell.inCycle,
    mode: cell.mode ?? null,
    topoRank: cell.topoRank ?? null,
    directPrecedents: [...cell.directPrecedents],
    directDependents: [...cell.directDependents],
    sheetProtection: sheetProtection(runtime, target.sheetName),
    rangeProtections: collectIntersectingRangeProtections(runtime, {
      sheetName: target.sheetName,
      startAddress: target.address,
      endAddress: target.address,
    }),
    dataValidations: collectIntersectingDataValidations(runtime, {
      sheetName: target.sheetName,
      startAddress: target.address,
      endAddress: target.address,
    }),
    conditionalFormats: collectIntersectingConditionalFormats(runtime, {
      sheetName: target.sheetName,
      startAddress: target.address,
      endAddress: target.address,
    }),
    commentThreads: collectIntersectingCommentThreads(runtime, {
      sheetName: target.sheetName,
      startAddress: target.address,
      endAddress: target.address,
    }),
    notes: collectIntersectingNotes(runtime, {
      sheetName: target.sheetName,
      startAddress: target.address,
      endAddress: target.address,
    }),
    images: collectIntersectingImages(runtime, {
      sheetName: target.sheetName,
      startAddress: target.address,
      endAddress: target.address,
    }),
    shapes: collectIntersectingShapes(runtime, {
      sheetName: target.sheetName,
      startAddress: target.address,
      endAddress: target.address,
    }),
  }
}
