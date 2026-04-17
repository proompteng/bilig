import { deepStrictEqual } from 'node:assert'
import * as fc from 'fast-check'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellNumberFormatInput, CellRangeRef, CellStylePatch, LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { EngineMutationError } from '../engine/errors.js'

export const engineFuzzSheetName = 'Sheet1'

type ComparableRangeRef = {
  sheetName: string
  startAddress: string
  endAddress: string
}

export type EngineSeedName =
  | 'blank'
  | 'formula-graph'
  | 'cross-sheet-graph'
  | 'sparse-format'
  | 'annotation-rich'
  | 'named-structures'
  | 'validation-filter-sort'
  | 'structural-metadata'
  | 'pivot-analytics'

export const engineSeedNames = [
  'blank',
  'formula-graph',
  'cross-sheet-graph',
  'sparse-format',
  'annotation-rich',
  'named-structures',
  'validation-filter-sort',
  'structural-metadata',
  'pivot-analytics',
] as const satisfies readonly EngineSeedName[]

export const engineSeedNameArbitrary = fc.constantFrom<EngineSeedName>(...engineSeedNames)

export type CoreAction =
  | { kind: 'values'; range: CellRangeRef; values: LiteralInput[][] }
  | { kind: 'formula'; address: string; formula: string }
  | { kind: 'style'; range: CellRangeRef; patch: CellStylePatch }
  | { kind: 'clearStyle'; range: CellRangeRef }
  | { kind: 'format'; range: CellRangeRef; format: CellNumberFormatInput }
  | { kind: 'clearFormat'; range: CellRangeRef }
  | { kind: 'clear'; range: CellRangeRef }
  | { kind: 'fill'; source: CellRangeRef; target: CellRangeRef }
  | { kind: 'copy'; source: CellRangeRef; target: CellRangeRef }
  | { kind: 'move'; source: CellRangeRef; target: CellRangeRef }
  | { kind: 'insertRows'; start: number; count: number }
  | { kind: 'deleteRows'; start: number; count: number }
  | { kind: 'insertColumns'; start: number; count: number }
  | { kind: 'deleteColumns'; start: number; count: number }

export type EngineReplayCommand = CoreAction | { kind: 'undo' } | { kind: 'redo' }

function assertNever(value: never): never {
  throw new Error(`Unexpected replay command: ${String(value)}`)
}

export function toRangeRef(targetSheetName: string, startRow: number, startCol: number, endRow: number, endCol: number): CellRangeRef {
  return {
    sheetName: targetSheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  }
}

function buildValueMatrix(height: number, width: number, values: readonly LiteralInput[]): LiteralInput[][] {
  const rows: LiteralInput[][] = []
  let offset = 0
  for (let row = 0; row < height; row += 1) {
    const nextRow: LiteralInput[] = []
    for (let col = 0; col < width; col += 1) {
      nextRow.push(values[offset] ?? null)
      offset += 1
    }
    rows.push(nextRow)
  }
  return rows
}

export function assertSnapshotInvariants(snapshot: WorkbookSnapshot): void {
  const sheetNames = snapshot.sheets.map((sheet) => sheet.name)
  if (new Set(sheetNames).size !== sheetNames.length) {
    throw new Error(`Duplicate sheet names in snapshot: ${sheetNames.join(', ')}`)
  }
  snapshot.sheets.forEach((sheet) => {
    const addresses = sheet.cells.map((cell) => cell.address)
    if (new Set(addresses).size !== addresses.length) {
      throw new Error(`Duplicate cell addresses in snapshot for ${sheet.name}`)
    }
  })
}

export function normalizeSnapshotForSemanticComparison(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const clone = structuredClone(snapshot)
  if (clone.workbook.metadata) {
    if (clone.workbook.metadata.properties) {
      clone.workbook.metadata.properties = clone.workbook.metadata.properties.toSorted((left, right) => left.key.localeCompare(right.key))
    }
    if (clone.workbook.metadata.definedNames) {
      clone.workbook.metadata.definedNames = clone.workbook.metadata.definedNames.toSorted((left, right) =>
        left.name.localeCompare(right.name),
      )
    }
    if (clone.workbook.metadata.tables) {
      clone.workbook.metadata.tables = clone.workbook.metadata.tables.toSorted((left, right) => left.name.localeCompare(right.name))
    }
    if (clone.workbook.metadata.styles) {
      clone.workbook.metadata.styles = clone.workbook.metadata.styles.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.formats) {
      clone.workbook.metadata.formats = clone.workbook.metadata.formats.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.pivots) {
      clone.workbook.metadata.pivots = clone.workbook.metadata.pivots.toSorted((left, right) =>
        `${left.sheetName}!${left.address}:${left.name}`.localeCompare(`${right.sheetName}!${right.address}:${right.name}`),
      )
    }
    if (clone.workbook.metadata.charts) {
      clone.workbook.metadata.charts = clone.workbook.metadata.charts.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.images) {
      clone.workbook.metadata.images = clone.workbook.metadata.images.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (clone.workbook.metadata.shapes) {
      clone.workbook.metadata.shapes = clone.workbook.metadata.shapes.toSorted((left, right) => left.id.localeCompare(right.id))
    }
  }
  clone.sheets = clone.sheets.map((sheet) => {
    const cells = sheet.cells.filter((cell) => cell.formula !== undefined || cell.format !== undefined || cell.value !== null)
    if (!sheet.metadata) {
      return {
        ...sheet,
        cells,
      }
    }
    const metadata = { ...sheet.metadata }
    if (metadata.rows) {
      metadata.rows = metadata.rows.map(({ id: _id, ...rest }) => rest).toSorted((left, right) => left.index - right.index)
    }
    if (metadata.columns) {
      metadata.columns = metadata.columns.map(({ id: _id, ...rest }) => rest).toSorted((left, right) => left.index - right.index)
    }
    if (metadata.rowMetadata) {
      metadata.rowMetadata = metadata.rowMetadata.toSorted((left, right) => left.start - right.start || left.count - right.count)
    }
    if (metadata.columnMetadata) {
      metadata.columnMetadata = metadata.columnMetadata.toSorted((left, right) => left.start - right.start || left.count - right.count)
    }
    if (metadata.styleRanges) {
      metadata.styleRanges = normalizeRangeRecords(metadata.styleRanges, 'styleId')
    }
    if (metadata.formatRanges) {
      metadata.formatRanges = normalizeRangeRecords(metadata.formatRanges, 'formatId')
    }
    if (metadata.filters) {
      metadata.filters = metadata.filters.toSorted(compareRangeRefs)
    }
    if (metadata.sorts) {
      metadata.sorts = metadata.sorts.toSorted((left, right) => {
        const rangeComparison = compareRangeRefs(left.range, right.range)
        if (rangeComparison !== 0) {
          return rangeComparison
        }
        return JSON.stringify(left.keys).localeCompare(JSON.stringify(right.keys))
      })
    }
    if (metadata.validations) {
      metadata.validations = metadata.validations.toSorted((left, right) => {
        const rangeComparison = compareRangeRefs(left.range, right.range)
        if (rangeComparison !== 0) {
          return rangeComparison
        }
        return JSON.stringify(left.rule).localeCompare(JSON.stringify(right.rule))
      })
    }
    if (metadata.conditionalFormats) {
      metadata.conditionalFormats = metadata.conditionalFormats.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (metadata.protectedRanges) {
      metadata.protectedRanges = metadata.protectedRanges.toSorted((left, right) => left.id.localeCompare(right.id))
    }
    if (metadata.commentThreads) {
      metadata.commentThreads = metadata.commentThreads.toSorted((left, right) =>
        `${left.sheetName}!${left.address}:${left.threadId}`.localeCompare(`${right.sheetName}!${right.address}:${right.threadId}`),
      )
    }
    if (metadata.notes) {
      metadata.notes = metadata.notes.toSorted((left, right) =>
        `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`),
      )
    }
    return {
      ...sheet,
      cells,
      metadata,
    }
  })
  return clone
}

function compareRangeRefs(left: ComparableRangeRef, right: ComparableRangeRef): number {
  const leftStart = parseCellAddress(left.startAddress, left.sheetName)
  const leftEnd = parseCellAddress(left.endAddress, left.sheetName)
  const rightStart = parseCellAddress(right.startAddress, right.sheetName)
  const rightEnd = parseCellAddress(right.endAddress, right.sheetName)
  return (
    left.sheetName.localeCompare(right.sheetName) ||
    leftStart.row - rightStart.row ||
    leftStart.col - rightStart.col ||
    leftEnd.row - rightEnd.row ||
    leftEnd.col - rightEnd.col
  )
}

function normalizeRangeRecords<TRecord extends { range: ComparableRangeRef } & Record<TKey, string>, TKey extends keyof TRecord & string>(
  records: readonly TRecord[],
  idKey: TKey,
): TRecord[] {
  const sorted = [...records].toSorted((left, right) => {
    const idComparison = left[idKey].localeCompare(right[idKey])
    if (idComparison !== 0) {
      return idComparison
    }
    return compareRangeRefs(left.range, right.range)
  })
  const grouped = new Map<string, TRecord[]>()
  for (const record of sorted) {
    const groupKey = record[idKey]
    const group = grouped.get(groupKey) ?? []
    grouped.set(groupKey, insertOrMergeRangeRecord(group, record))
  }
  return [...grouped.entries()]
    .toSorted(([left], [right]) => left.localeCompare(right))
    .flatMap(([_groupKey, group]) => group.toSorted((left, right) => compareRangeRefs(left.range, right.range)))
}

function insertOrMergeRangeRecord<TRecord extends { range: ComparableRangeRef }>(group: readonly TRecord[], record: TRecord): TRecord[] {
  const nextGroup = [...group]
  nextGroup.push(record)
  let changed = true
  while (changed) {
    changed = false
    for (let leftIndex = 0; leftIndex < nextGroup.length; leftIndex += 1) {
      for (let rightIndex = leftIndex + 1; rightIndex < nextGroup.length; rightIndex += 1) {
        const left = nextGroup[leftIndex]
        const right = nextGroup[rightIndex]
        if (!left || !right) {
          continue
        }
        const merged = mergeRanges(left.range, right.range)
        if (!merged) {
          continue
        }
        nextGroup[leftIndex] = { ...left, range: merged }
        nextGroup.splice(rightIndex, 1)
        changed = true
        break
      }
      if (changed) {
        break
      }
    }
  }
  return nextGroup
}

function mergeRanges(left: ComparableRangeRef, right: ComparableRangeRef): ComparableRangeRef | null {
  if (left.sheetName !== right.sheetName) {
    return null
  }
  const leftStart = parseCellAddress(left.startAddress, left.sheetName)
  const leftEnd = parseCellAddress(left.endAddress, left.sheetName)
  const rightStart = parseCellAddress(right.startAddress, right.sheetName)
  const rightEnd = parseCellAddress(right.endAddress, right.sheetName)

  const sameRows = leftStart.row === rightStart.row && leftEnd.row === rightEnd.row
  const horizontallyAdjacent = sameRows && Math.max(leftStart.col, rightStart.col) <= Math.min(leftEnd.col, rightEnd.col) + 1
  if (horizontallyAdjacent) {
    return {
      sheetName: left.sheetName,
      startAddress: formatAddress(leftStart.row, Math.min(leftStart.col, rightStart.col)),
      endAddress: formatAddress(leftEnd.row, Math.max(leftEnd.col, rightEnd.col)),
    }
  }

  const sameCols = leftStart.col === rightStart.col && leftEnd.col === rightEnd.col
  const verticallyAdjacent = sameCols && Math.max(leftStart.row, rightStart.row) <= Math.min(leftEnd.row, rightEnd.row) + 1
  if (verticallyAdjacent) {
    return {
      sheetName: left.sheetName,
      startAddress: formatAddress(Math.min(leftStart.row, rightStart.row), leftStart.col),
      endAddress: formatAddress(Math.max(leftEnd.row, rightEnd.row), leftEnd.col),
    }
  }

  return null
}

const literalInputArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -10_000, max: 10_000 }),
  fc.boolean(),
  fc.constantFrom('north', 'south', 'ready', 'done'),
  fc.constant(null),
)

const rangeSeedArbitrary = fc.record({
  startRow: fc.integer({ min: 0, max: 5 }),
  startCol: fc.integer({ min: 0, max: 5 }),
  height: fc.integer({ min: 1, max: 2 }),
  width: fc.integer({ min: 1, max: 2 }),
})

export const rangeArbitrary = rangeSeedArbitrary.map((value) =>
  toRangeRef(engineFuzzSheetName, value.startRow, value.startCol, value.startRow + value.height - 1, value.startCol + value.width - 1),
)

const sameSizedRangePairArbitrary = rangeSeedArbitrary.chain((source) =>
  fc
    .record({
      targetStartRow: fc.integer({ min: 0, max: 6 - source.height }),
      targetStartCol: fc.integer({ min: 0, max: 6 - source.width }),
    })
    .filter(({ targetStartRow, targetStartCol }) => targetStartRow !== source.startRow || targetStartCol !== source.startCol)
    .map(({ targetStartRow, targetStartCol }) => ({
      source: toRangeRef(
        engineFuzzSheetName,
        source.startRow,
        source.startCol,
        source.startRow + source.height - 1,
        source.startCol + source.width - 1,
      ),
      target: toRangeRef(
        engineFuzzSheetName,
        targetStartRow,
        targetStartCol,
        targetStartRow + source.height - 1,
        targetStartCol + source.width - 1,
      ),
    })),
)

const formulaArbitrary = fc
  .tuple(
    fc.constantFrom('A1', 'B2', 'C3', 'D4', 'E5', 'F6'),
    fc.constantFrom('+', '-', '*', '/'),
    fc.constantFrom('A1', 'B2', 'C3', 'D4', 'E5', 'F6'),
  )
  .map(([left, operator, right]) => `${left}${operator}${right}`)

const stylePatchArbitrary = fc.constantFrom<CellStylePatch>(
  { fill: { backgroundColor: '#dbeafe' } },
  { font: { bold: true } },
  { alignment: { horizontal: 'right', wrap: true } },
)

const formatInputArbitrary = fc.constantFrom<CellNumberFormatInput>(
  '0.00',
  { kind: 'currency', currency: 'USD', decimals: 2 },
  { kind: 'percent', decimals: 1 },
  { kind: 'text' },
)

export const valuesActionArbitrary = rangeSeedArbitrary.chain((range) =>
  fc
    .array(literalInputArbitrary, {
      minLength: range.height * range.width,
      maxLength: range.height * range.width,
    })
    .map((values) => ({
      kind: 'values' as const,
      range: toRangeRef(
        engineFuzzSheetName,
        range.startRow,
        range.startCol,
        range.startRow + range.height - 1,
        range.startCol + range.width - 1,
      ),
      values: buildValueMatrix(range.height, range.width, values),
    })),
)

export const formulaActionArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 5 }),
    col: fc.integer({ min: 0, max: 5 }),
    formula: formulaArbitrary,
  })
  .map(({ row, col, formula }) => ({
    kind: 'formula' as const,
    address: formatAddress(row, col),
    formula,
  }))

export const styleActionArbitrary = fc
  .record({ range: rangeArbitrary, patch: stylePatchArbitrary })
  .map(({ range, patch }) => ({ kind: 'style' as const, range, patch }))

export const clearStyleActionArbitrary = rangeArbitrary.map((range) => ({
  kind: 'clearStyle' as const,
  range,
}))

export const formatActionArbitrary = fc
  .record({ range: rangeArbitrary, format: formatInputArbitrary })
  .map(({ range, format }) => ({ kind: 'format' as const, range, format }))

export const clearFormatActionArbitrary = rangeArbitrary.map((range) => ({
  kind: 'clearFormat' as const,
  range,
}))

export const clearActionArbitrary = rangeArbitrary.map((range) => ({
  kind: 'clear' as const,
  range,
}))

export const fillActionArbitrary = sameSizedRangePairArbitrary.map(({ source, target }) => ({
  kind: 'fill' as const,
  source,
  target,
}))

export const copyActionArbitrary = sameSizedRangePairArbitrary.map(({ source, target }) => ({
  kind: 'copy' as const,
  source,
  target,
}))

export const moveActionArbitrary = sameSizedRangePairArbitrary.map(({ source, target }) => ({
  kind: 'move' as const,
  source,
  target,
}))

const axisMutationArbitrary = fc.record({
  start: fc.integer({ min: 0, max: 4 }),
  count: fc.integer({ min: 1, max: 2 }),
})

export const insertRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'insertRows' as const,
  start,
  count,
}))

export const deleteRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'deleteRows' as const,
  start,
  count,
}))

export const insertColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'insertColumns' as const,
  start,
  count,
}))

export const deleteColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'deleteColumns' as const,
  start,
  count,
}))

export const corePreparationActionArbitrary = fc.oneof<CoreAction>(
  valuesActionArbitrary,
  formulaActionArbitrary,
  styleActionArbitrary,
  formatActionArbitrary,
  clearActionArbitrary,
)

export const coreStructuralActionArbitrary = fc.oneof<CoreAction>(
  insertRowsActionArbitrary,
  deleteRowsActionArbitrary,
  insertColumnsActionArbitrary,
  deleteColumnsActionArbitrary,
)

export const coreReplicaActionArbitrary = fc.oneof<CoreAction>(
  valuesActionArbitrary,
  formulaActionArbitrary,
  styleActionArbitrary,
  formatActionArbitrary,
  clearActionArbitrary,
  fillActionArbitrary,
  copyActionArbitrary,
  moveActionArbitrary,
  insertRowsActionArbitrary,
  deleteRowsActionArbitrary,
  insertColumnsActionArbitrary,
  deleteColumnsActionArbitrary,
)

async function populateSeed(engine: SpreadsheetEngine, seedName: EngineSeedName): Promise<void> {
  engine.createSheet(engineFuzzSheetName)

  switch (seedName) {
    case 'blank':
      return
    case 'formula-graph':
      engine.setRangeValues(
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'B2',
        },
        [
          [4, 2],
          [3, 7],
        ],
      )
      engine.setCellFormula(engineFuzzSheetName, 'C1', 'A1+B2')
      engine.setCellFormula(engineFuzzSheetName, 'D2', 'C1*A2')
      engine.setCellFormula(engineFuzzSheetName, 'E3', 'SUM(A1:B2)')
      return
    case 'cross-sheet-graph':
      engine.createSheet('Summary')
      engine.setRangeValues(
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'B3',
        },
        [
          ['Region', 'Sales'],
          ['East', 12],
          ['West', 18],
        ],
      )
      engine.setRangeValues(
        {
          sheetName: 'Summary',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        [
          ['Tax', 'Multiplier'],
          [0.08, 2],
        ],
      )
      engine.setDefinedName('TaxCell', {
        kind: 'cell-ref',
        sheetName: 'Summary',
        address: 'A2',
      })
      engine.setCellFormula(engineFuzzSheetName, 'C2', 'B2*Summary!B2')
      engine.setCellFormula(engineFuzzSheetName, 'C3', 'B3*(1+TaxCell)')
      engine.setCellFormula('Summary', 'D1', 'SUM(Sheet1!B2:B3)')
      return
    case 'sparse-format':
      engine.setRangeNumberFormat({ sheetName: engineFuzzSheetName, startAddress: 'A1', endAddress: 'A1' }, '0.00')
      engine.setRangeStyle(
        { sheetName: engineFuzzSheetName, startAddress: 'C3', endAddress: 'C3' },
        { fill: { backgroundColor: '#dbeafe' } },
      )
      engine.setRangeStyle(
        { sheetName: engineFuzzSheetName, startAddress: 'E2', endAddress: 'F2' },
        { alignment: { horizontal: 'center', wrap: true } },
      )
      return
    case 'annotation-rich':
      engine.setRangeValues(
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'C4',
        },
        [
          ['Owner', 'Status', 'Amount'],
          ['Greg', 'Draft', 10],
          ['Ada', 'Final', 24],
          ['Lin', 'Review', 17],
        ],
      )
      engine.setConditionalFormat({
        id: 'cf-annotation-seed',
        range: {
          sheetName: engineFuzzSheetName,
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'textContains',
          text: 'Final',
        },
        style: {
          fill: { backgroundColor: '#dcfce7' },
          font: { bold: true },
        },
      })
      engine.setCommentThread({
        threadId: 'thread-annotation-seed',
        sheetName: engineFuzzSheetName,
        address: 'C2',
        comments: [{ id: 'comment-1', body: 'Verify reviewed total.' }],
      })
      engine.setNote({
        sheetName: engineFuzzSheetName,
        address: 'A4',
        text: 'Backfill owner before close.',
      })
      engine.setChart({
        id: 'chart-annotation-seed',
        sheetName: engineFuzzSheetName,
        address: 'E2',
        source: {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'C4',
        },
        chartType: 'column',
        firstRowAsHeaders: true,
        firstColumnAsLabels: true,
        title: 'Pipeline',
        legendPosition: 'right',
        rows: 6,
        cols: 4,
      })
      engine.setImage({
        id: 'image-annotation-seed',
        sheetName: engineFuzzSheetName,
        address: 'H2',
        sourceUrl: 'https://example.invalid/pipeline.png',
        rows: 3,
        cols: 2,
        altText: 'Pipeline thumbnail',
      })
      engine.setShape({
        id: 'shape-annotation-seed',
        sheetName: engineFuzzSheetName,
        address: 'H6',
        shapeType: 'roundedRectangle',
        rows: 2,
        cols: 2,
        text: 'Review',
        fillColor: '#e0f2fe',
        strokeColor: '#0284c7',
      })
      return
    case 'named-structures':
      engine.setRangeValues(
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'B3',
        },
        [
          ['Qty', 'Amount'],
          [1, 10],
          [2, 20],
        ],
      )
      engine.setDefinedName('TaxRate', 0.085)
      engine.setDefinedName('SalesRange', {
        kind: 'range-ref',
        sheetName: engineFuzzSheetName,
        startAddress: 'A1',
        endAddress: 'B3',
      })
      engine.setTable({
        name: 'Sales',
        sheetName: engineFuzzSheetName,
        startAddress: 'A1',
        endAddress: 'B3',
        columnNames: ['Qty', 'Amount'],
        headerRow: true,
        totalsRow: false,
      })
      engine.setCellFormula(engineFuzzSheetName, 'C1', 'SUM(SalesRange)')
      return
    case 'validation-filter-sort':
      engine.setRangeValues(
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'C3',
        },
        [
          ['Id', 'Status', 'Amount'],
          [1, 'Draft', 10],
          [2, 'Final', 20],
        ],
      )
      engine.setFilter(engineFuzzSheetName, {
        sheetName: engineFuzzSheetName,
        startAddress: 'A1',
        endAddress: 'C3',
      })
      engine.setSort(
        engineFuzzSheetName,
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'C3',
        },
        [{ keyAddress: 'B1', direction: 'asc' }],
      )
      engine.setDataValidation({
        range: {
          sheetName: engineFuzzSheetName,
          startAddress: 'B2',
          endAddress: 'B3',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
      })
      return
    case 'structural-metadata':
      engine.insertRows(engineFuzzSheetName, 0, 1)
      engine.insertColumns(engineFuzzSheetName, 0, 1)
      engine.updateRowMetadata(engineFuzzSheetName, 0, 1, 28, false)
      engine.updateColumnMetadata(engineFuzzSheetName, 0, 1, 120, false)
      engine.setFreezePane(engineFuzzSheetName, 1, 1)
      engine.setRangeValues(
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'B2',
          endAddress: 'C3',
        },
        [
          [1, 2],
          [3, 4],
        ],
      )
      return
    case 'pivot-analytics':
      engine.createSheet('Pivot')
      engine.setRangeValues(
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'C5',
        },
        [
          ['Region', 'Quarter', 'Sales'],
          ['East', 'Q1', 10],
          ['West', 'Q1', 6],
          ['East', 'Q2', 12],
          ['West', 'Q2', 9],
        ],
      )
      engine.setTable({
        name: 'QuarterlySales',
        sheetName: engineFuzzSheetName,
        startAddress: 'A1',
        endAddress: 'C5',
        columnNames: ['Region', 'Quarter', 'Sales'],
        headerRow: true,
        totalsRow: false,
      })
      engine.setFreezePane(engineFuzzSheetName, 1, 0)
      engine.setFilter(engineFuzzSheetName, {
        sheetName: engineFuzzSheetName,
        startAddress: 'A1',
        endAddress: 'C5',
      })
      engine.setSort(
        engineFuzzSheetName,
        {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'C5',
        },
        [{ keyAddress: 'C1', direction: 'desc' }],
      )
      engine.setPivotTable('Pivot', 'B2', {
        name: 'QuarterlyPivot',
        source: {
          sheetName: engineFuzzSheetName,
          startAddress: 'A1',
          endAddress: 'C5',
        },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
      })
      engine.setCellFormula('Pivot', 'F2', 'SUM(Sheet1!C2:C5)')
      return
  }
}

export async function createEngineSeedSnapshot(seedName: EngineSeedName, workbookName: string): Promise<WorkbookSnapshot> {
  const engine = new SpreadsheetEngine({
    workbookName,
    replicaId: `seed-${seedName}`,
  })
  await engine.ready()
  await populateSeed(engine, seedName)
  return engine.exportSnapshot()
}

export async function createSeededEngine(seedName: EngineSeedName, workbookName: string, replicaId: string): Promise<SpreadsheetEngine> {
  const snapshot = await createEngineSeedSnapshot(seedName, workbookName)
  const engine = new SpreadsheetEngine({ workbookName, replicaId })
  await engine.ready()
  engine.importSnapshot(snapshot)
  return engine
}

export function applyCoreAction(engine: SpreadsheetEngine, action: CoreAction): void {
  switch (action.kind) {
    case 'values':
      engine.setRangeValues(action.range, action.values)
      break
    case 'formula':
      engine.setCellFormula(engineFuzzSheetName, action.address, action.formula)
      break
    case 'style':
      engine.setRangeStyle(action.range, action.patch)
      break
    case 'clearStyle':
      engine.clearRangeStyle(action.range)
      break
    case 'format':
      engine.setRangeNumberFormat(action.range, action.format)
      break
    case 'clearFormat':
      engine.clearRangeNumberFormat(action.range)
      break
    case 'clear':
      engine.clearRange(action.range)
      break
    case 'fill':
      engine.fillRange(action.source, action.target)
      break
    case 'copy':
      engine.copyRange(action.source, action.target)
      break
    case 'move':
      engine.moveRange(action.source, action.target)
      break
    case 'insertRows':
      engine.insertRows(engineFuzzSheetName, action.start, action.count)
      break
    case 'deleteRows':
      engine.deleteRows(engineFuzzSheetName, action.start, action.count)
      break
    case 'insertColumns':
      engine.insertColumns(engineFuzzSheetName, action.start, action.count)
      break
    case 'deleteColumns':
      engine.deleteColumns(engineFuzzSheetName, action.start, action.count)
      break
  }
}

export function applyActionAndCaptureResult(
  engine: SpreadsheetEngine,
  action: CoreAction,
): { accepted: boolean; before: WorkbookSnapshot; after: WorkbookSnapshot } {
  const before = engine.exportSnapshot()
  try {
    applyCoreAction(engine, action)
    const after = engine.exportSnapshot()
    return {
      accepted: hasSemanticChange(before, after),
      before,
      after,
    }
  } catch (error) {
    if (!(error instanceof EngineMutationError)) {
      throw error
    }
    const after = engine.exportSnapshot()
    deepStrictEqual(after, before)
    return {
      accepted: false,
      before,
      after,
    }
  }
}

export async function exportReplaySnapshot(initialSnapshot: WorkbookSnapshot, actions: readonly CoreAction[]): Promise<WorkbookSnapshot> {
  const replay = new SpreadsheetEngine({
    workbookName: initialSnapshot.workbook.name,
    replicaId: `replay-${initialSnapshot.workbook.name}`,
  })
  await replay.ready()
  replay.importSnapshot(structuredClone(initialSnapshot))
  for (const action of actions) {
    applyCoreAction(replay, action)
  }
  return replay.exportSnapshot()
}

function hasSemanticChange(before: WorkbookSnapshot, after: WorkbookSnapshot): boolean {
  try {
    deepStrictEqual(after, before)
    return false
  } catch {
    return true
  }
}

export function applyReplayCommand(engine: SpreadsheetEngine, command: EngineReplayCommand): void {
  switch (command.kind) {
    case 'undo':
      if (!engine.undo()) {
        throw new Error('Expected undo() to succeed while replaying fixture')
      }
      return
    case 'redo':
      if (!engine.redo()) {
        throw new Error('Expected redo() to succeed while replaying fixture')
      }
      return
    case 'values':
    case 'formula':
    case 'style':
    case 'clearStyle':
    case 'format':
    case 'clearFormat':
    case 'clear':
    case 'fill':
    case 'copy':
    case 'move':
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns':
      applyCoreAction(engine, command)
      return
  }
  return assertNever(command)
}
