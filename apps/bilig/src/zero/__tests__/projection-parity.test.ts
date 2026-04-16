import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { SpreadsheetEngine } from '@bilig/core'
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
  WorkbookCalculationSettingsSnapshot,
  WorkbookDefinedNameValueSnapshot,
} from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import {
  buildWorkbookSourceProjection,
  buildWorkbookSourceProjectionFromEngine,
  type WorkbookProjectionOptions,
  type WorkbookSourceProjection,
} from '../projection.js'

type ProjectionParityAction =
  | { kind: 'value'; address: string; value: LiteralInput }
  | { kind: 'formula'; address: string; formula: string }
  | { kind: 'style'; range: CellRangeRef; patch: CellStylePatch }
  | { kind: 'format'; range: CellRangeRef; format: CellNumberFormatInput }
  | {
      kind: 'rowMetadata'
      start: number
      count: number
      size: number | null
      hidden: boolean | null
    }
  | {
      kind: 'columnMetadata'
      start: number
      count: number
      size: number | null
      hidden: boolean | null
    }
  | { kind: 'freezePane'; rows: number; cols: number }
  | { kind: 'definedName'; name: string; value: WorkbookDefinedNameValueSnapshot }
  | { kind: 'workbookMetadata'; key: string; value: LiteralInput }
  | { kind: 'calculationSettings'; settings: WorkbookCalculationSettingsSnapshot }
  | { kind: 'insertRows'; start: number; count: number }
  | { kind: 'deleteRows'; start: number; count: number }
  | { kind: 'insertColumns'; start: number; count: number }
  | { kind: 'deleteColumns'; start: number; count: number }

const sheetName = 'Sheet1'
const projectionOptions: WorkbookProjectionOptions = {
  revision: 9,
  calculatedRevision: 9,
  ownerUserId: 'owner-1',
  updatedBy: 'user-1',
  updatedAt: '2026-04-09T12:00:00.000Z',
}

const addressArbitrary = fc.constantFrom('A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3')
const rangeArbitrary = fc.constantFrom<CellRangeRef>(
  { sheetName, startAddress: 'A1', endAddress: 'A1' },
  { sheetName, startAddress: 'A1', endAddress: 'B2' },
  { sheetName, startAddress: 'B2', endAddress: 'C3' },
  { sheetName, startAddress: 'A2', endAddress: 'C2' },
)
const literalInputArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -50, max: 50 }),
  fc.boolean(),
  fc.constantFrom('north', 'south', 'ready', 'hold'),
  fc.constant(null),
)
const formulaArbitrary = fc.constantFrom('1', 'A1+1', 'A1+B1', 'SUM(A1:B2)', 'A1&B1')
const stylePatchArbitrary = fc.constantFrom<CellStylePatch>(
  { fill: { backgroundColor: '#ffcc00' } },
  { font: { bold: true } },
  { alignment: { horizontal: 'center' } },
)
const numberFormatArbitrary = fc.constantFrom<CellNumberFormatInput>(
  '0.00',
  { kind: 'number', code: '$#,##0.00' },
  { kind: 'percent', code: '0.0%' },
)
const metadataCountArbitrary = fc.integer({ min: 1, max: 2 })
const rowMetadataArbitrary = fc
  .record({
    start: fc.integer({ min: 0, max: 3 }),
    count: metadataCountArbitrary,
    size: fc.option(fc.integer({ min: 18, max: 60 }), { nil: null }),
    hidden: fc.option(fc.boolean(), { nil: null }),
  })
  .map((action) => Object.assign({ kind: `rowMetadata` as const }, action))
const columnMetadataArbitrary = fc
  .record({
    start: fc.integer({ min: 0, max: 3 }),
    count: metadataCountArbitrary,
    size: fc.option(fc.integer({ min: 60, max: 180 }), { nil: null }),
    hidden: fc.option(fc.boolean(), { nil: null }),
  })
  .map((action) => Object.assign({ kind: `columnMetadata` as const }, action))

const projectionParityActionArbitrary = fc.oneof<ProjectionParityAction>(
  fc.record({ address: addressArbitrary, value: literalInputArbitrary }).map((action) => Object.assign({ kind: `value` as const }, action)),
  fc.record({ address: addressArbitrary, formula: formulaArbitrary }).map((action) => Object.assign({ kind: `formula` as const }, action)),
  fc.record({ range: rangeArbitrary, patch: stylePatchArbitrary }).map((action) => Object.assign({ kind: `style` as const }, action)),
  fc.record({ range: rangeArbitrary, format: numberFormatArbitrary }).map((action) => Object.assign({ kind: `format` as const }, action)),
  rowMetadataArbitrary,
  columnMetadataArbitrary,
  fc
    .record({
      rows: fc.integer({ min: 0, max: 2 }),
      cols: fc.integer({ min: 0, max: 2 }),
    })
    .map((action) => Object.assign({ kind: `freezePane` as const }, action)),
  fc
    .record({
      name: fc.constantFrom('TaxRate', 'Region', 'StatusFlag'),
      value: fc.oneof<WorkbookDefinedNameValueSnapshot>(
        fc.integer({ min: 1, max: 9 }),
        fc.constant('west'),
        fc.constant(true),
        fc.constant({ kind: 'formula', formula: 'A1+1' }),
      ),
    })
    .map((action) => Object.assign({ kind: `definedName` as const }, action)),
  fc
    .record({
      key: fc.constantFrom('theme', 'ownerRegion', 'syncMode'),
      value: literalInputArbitrary,
    })
    .map((action) => Object.assign({ kind: `workbookMetadata` as const }, action)),
  fc
    .record({
      settings: fc.record({
        mode: fc.constantFrom<WorkbookCalculationSettingsSnapshot['mode']>('automatic', 'manual'),
        compatibilityMode: fc.constantFrom<WorkbookCalculationSettingsSnapshot['compatibilityMode']>('excel-modern', 'odf-1.4'),
      }),
    })
    .map((action) => Object.assign({ kind: `calculationSettings` as const }, action)),
  fc
    .record({
      start: fc.integer({ min: 0, max: 2 }),
      count: fc.integer({ min: 1, max: 1 }),
    })
    .map((action) => Object.assign({ kind: `insertRows` as const }, action)),
  fc
    .record({
      start: fc.integer({ min: 0, max: 2 }),
      count: fc.integer({ min: 1, max: 1 }),
    })
    .map((action) => Object.assign({ kind: `deleteRows` as const }, action)),
  fc
    .record({
      start: fc.integer({ min: 0, max: 2 }),
      count: fc.integer({ min: 1, max: 1 }),
    })
    .map((action) => Object.assign({ kind: `insertColumns` as const }, action)),
  fc
    .record({
      start: fc.integer({ min: 0, max: 2 }),
      count: fc.integer({ min: 1, max: 1 }),
    })
    .map((action) => Object.assign({ kind: `deleteColumns` as const }, action)),
)

describe('projection parity', () => {
  it('should build the same workbook source projection from the engine and exported snapshot', async () => {
    // Arrange
    const engine = new SpreadsheetEngine({
      workbookName: 'projection-parity-doc',
      replicaId: 'projection-parity-doc',
    })
    await engine.ready()
    engine.setCellValue(sheetName, 'A1', 7)
    engine.setCellFormula(sheetName, 'B2', 'A1+1')
    engine.setRangeStyle({ sheetName, startAddress: 'B2', endAddress: 'C3' }, { fill: { backgroundColor: '#abcdef' } })
    engine.setRangeNumberFormat({ sheetName, startAddress: 'B2', endAddress: 'C3' }, { kind: 'number', code: '$#,##0.00' })
    engine.updateRowMetadata(sheetName, 1, 1, 28, false)
    engine.updateColumnMetadata(sheetName, 2, 1, 120, false)
    engine.setDefinedName('TaxRate', 0.07)
    engine.setWorkbookMetadata('theme', 'classic')

    // Act
    const fromEngine = normalizeProjection(buildWorkbookSourceProjectionFromEngine('doc-1', engine, projectionOptions))
    const fromSnapshot = normalizeProjection(buildWorkbookSourceProjection('doc-1', engine.exportSnapshot(), projectionOptions))

    // Assert
    expect(fromEngine).toEqual(fromSnapshot)
  })

  it('should preserve engine and snapshot projection parity across mixed workbook mutations', async () => {
    await runProperty({
      suite: 'bilig/projection-source-parity',
      arbitrary: fc.array(projectionParityActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        // Arrange
        const engine = new SpreadsheetEngine({
          workbookName: 'projection-parity-fuzz',
          replicaId: 'projection-parity-fuzz',
        })
        await engine.ready()

        // Act
        for (const action of actions) {
          applyProjectionParityAction(engine, action)
        }

        const fromEngine = normalizeProjection(buildWorkbookSourceProjectionFromEngine('doc-1', engine, projectionOptions))
        const fromSnapshot = normalizeProjection(buildWorkbookSourceProjection('doc-1', engine.exportSnapshot(), projectionOptions))

        // Assert
        expect(fromEngine).toEqual(fromSnapshot)
      },
    })
  })
})

// Helpers

function applyProjectionParityAction(engine: SpreadsheetEngine, action: ProjectionParityAction): void {
  switch (action.kind) {
    case 'value':
      engine.setCellValue(sheetName, action.address, action.value)
      return
    case 'formula':
      engine.setCellFormula(sheetName, action.address, action.formula)
      return
    case 'style':
      engine.setRangeStyle(action.range, action.patch)
      return
    case 'format':
      engine.setRangeNumberFormat(action.range, action.format)
      return
    case 'rowMetadata':
      engine.updateRowMetadata(sheetName, action.start, action.count, action.size, action.hidden)
      return
    case 'columnMetadata':
      engine.updateColumnMetadata(sheetName, action.start, action.count, action.size, action.hidden)
      return
    case 'freezePane':
      engine.setFreezePane(sheetName, action.rows, action.cols)
      return
    case 'definedName':
      engine.setDefinedName(action.name, action.value)
      return
    case 'workbookMetadata':
      engine.setWorkbookMetadata(action.key, action.value)
      return
    case 'calculationSettings':
      engine.setCalculationSettings(action.settings)
      return
    case 'insertRows':
      engine.insertRows(sheetName, action.start, action.count)
      return
    case 'deleteRows':
      engine.deleteRows(sheetName, action.start, action.count)
      return
    case 'insertColumns':
      engine.insertColumns(sheetName, action.start, action.count)
      return
    case 'deleteColumns':
      engine.deleteColumns(sheetName, action.start, action.count)
      return
  }
}

function normalizeProjection(projection: WorkbookSourceProjection): WorkbookSourceProjection {
  return {
    workbook: projection.workbook,
    sheets: [...projection.sheets].toSorted((left, right) => compareTuples([left.sheetId], [right.sheetId])),
    cells: [...projection.cells].toSorted((left, right) =>
      compareTuples([left.sheetName, left.rowNum, left.colNum], [right.sheetName, right.rowNum, right.colNum]),
    ),
    rowMetadata: [...projection.rowMetadata].toSorted((left, right) =>
      compareTuples([left.sheetName, left.startIndex, left.count], [right.sheetName, right.startIndex, right.count]),
    ),
    columnMetadata: [...projection.columnMetadata].toSorted((left, right) =>
      compareTuples([left.sheetName, left.startIndex, left.count], [right.sheetName, right.startIndex, right.count]),
    ),
    definedNames: [...projection.definedNames].toSorted((left, right) => compareTuples([left.name], [right.name])),
    workbookMetadataEntries: [...projection.workbookMetadataEntries].toSorted((left, right) => compareTuples([left.key], [right.key])),
    calculationSettings: projection.calculationSettings,
    styles: [...projection.styles].toSorted((left, right) => compareTuples([left.id], [right.id])),
    numberFormats: [...projection.numberFormats].toSorted((left, right) => compareTuples([left.id], [right.id])),
  }
}

function compareTuples(left: readonly unknown[], right: readonly unknown[]): number {
  for (let index = 0; index < Math.min(left.length, right.length); index += 1) {
    const leftValue = String(left[index])
    const rightValue = String(right[index])
    if (leftValue < rightValue) {
      return -1
    }
    if (leftValue > rightValue) {
      return 1
    }
  }
  return left.length - right.length
}
