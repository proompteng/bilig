import * as fc from 'fast-check'
import type { SpreadsheetEngine } from '@bilig/core'
import type { CellNumberFormatInput, CellRangeRef, CellStylePatch, LiteralInput } from '@bilig/protocol'
import {
  buildWorkbookSourceProjection,
  buildWorkbookSourceProjectionFromEngine,
  type WorkbookProjectionOptions,
  type WorkbookSourceProjection,
} from '../projection.js'

export type ProjectionAction =
  | { kind: 'value'; address: string; value: LiteralInput }
  | { kind: 'formula'; address: string; formula: string }
  | { kind: 'style'; range: CellRangeRef; patch: CellStylePatch }
  | { kind: 'format'; range: CellRangeRef; format: CellNumberFormatInput }
  | { kind: 'insertRows'; start: number; count: number }
  | { kind: 'deleteRows'; start: number; count: number }
  | { kind: 'insertColumns'; start: number; count: number }
  | { kind: 'deleteColumns'; start: number; count: number }

const projectionSheetName = 'Sheet1'

const projectionOptions: WorkbookProjectionOptions = {
  revision: 9,
  calculatedRevision: 9,
  ownerUserId: 'owner-1',
  updatedBy: 'user-1',
  updatedAt: '2026-04-13T08:00:00.000Z',
}

const addressArbitrary = fc.constantFrom('A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3')
const rangeArbitrary = fc.constantFrom<CellRangeRef>(
  { sheetName: projectionSheetName, startAddress: 'A1', endAddress: 'A1' },
  { sheetName: projectionSheetName, startAddress: 'A1', endAddress: 'B2' },
  { sheetName: projectionSheetName, startAddress: 'B2', endAddress: 'C3' },
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

export const projectionActionArbitrary = fc.oneof<ProjectionAction>(
  fc.record({ address: addressArbitrary, value: literalInputArbitrary }).map((action) => ({
    kind: 'value' as const,
    address: action.address,
    value: action.value,
  })),
  fc.record({ address: addressArbitrary, formula: formulaArbitrary }).map((action) => ({
    kind: 'formula' as const,
    address: action.address,
    formula: action.formula,
  })),
  fc.record({ range: rangeArbitrary, patch: stylePatchArbitrary }).map((action) => ({
    kind: 'style' as const,
    range: action.range,
    patch: action.patch,
  })),
  fc.record({ range: rangeArbitrary, format: numberFormatArbitrary }).map((action) => ({
    kind: 'format' as const,
    range: action.range,
    format: action.format,
  })),
  structuralActionArbitrary('insertRows'),
  structuralActionArbitrary('deleteRows'),
  structuralActionArbitrary('insertColumns'),
  structuralActionArbitrary('deleteColumns'),
)

function structuralActionArbitrary(
  kind: Extract<ProjectionAction['kind'], 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns'>,
): fc.Arbitrary<ProjectionAction> {
  return fc
    .record({
      start: fc.integer({ min: 0, max: 2 }),
      count: fc.integer({ min: 1, max: 1 }),
    })
    .map((action) => Object.assign({ kind }, action) as ProjectionAction)
}

export function applyProjectionAction(engine: SpreadsheetEngine, action: ProjectionAction): void {
  switch (action.kind) {
    case 'value':
      engine.setCellValue(projectionSheetName, action.address, action.value)
      return
    case 'formula':
      engine.setCellFormula(projectionSheetName, action.address, action.formula)
      return
    case 'style':
      engine.setRangeStyle(action.range, action.patch)
      return
    case 'format':
      engine.setRangeNumberFormat(action.range, action.format)
      return
    case 'insertRows':
      engine.insertRows(projectionSheetName, action.start, action.count)
      return
    case 'deleteRows':
      engine.deleteRows(projectionSheetName, action.start, action.count)
      return
    case 'insertColumns':
      engine.insertColumns(projectionSheetName, action.start, action.count)
      return
    case 'deleteColumns':
      engine.deleteColumns(projectionSheetName, action.start, action.count)
      return
  }
}

function normalizeProjection(projection: WorkbookSourceProjection): WorkbookSourceProjection {
  return {
    ...projection,
    sheets: projection.sheets.toSorted((left, right) => left.sheetId - right.sheetId),
    cells: projection.cells.toSorted(compareProjectionKey),
    rowMetadata: projection.rowMetadata.toSorted(compareAxisKey),
    columnMetadata: projection.columnMetadata.toSorted(compareAxisKey),
    definedNames: projection.definedNames.toSorted((left, right) => left.name.localeCompare(right.name)),
    workbookMetadataEntries: projection.workbookMetadataEntries.toSorted((left, right) => left.key.localeCompare(right.key)),
    styles: projection.styles.toSorted((left, right) => left.id.localeCompare(right.id)),
    numberFormats: projection.numberFormats.toSorted((left, right) => left.id.localeCompare(right.id)),
  }
}

export function projectProjectionFromEngine(engine: SpreadsheetEngine): WorkbookSourceProjection {
  return normalizeProjection(buildWorkbookSourceProjectionFromEngine('doc-1', engine, projectionOptions))
}

export function projectProjectionFromSnapshot(engine: SpreadsheetEngine): WorkbookSourceProjection {
  return normalizeProjection(buildWorkbookSourceProjection('doc-1', engine.exportSnapshot(), projectionOptions))
}

function compareProjectionKey(left: WorkbookSourceProjection['cells'][number], right: WorkbookSourceProjection['cells'][number]): number {
  return left.sheetName.localeCompare(right.sheetName) || left.rowNum - right.rowNum || left.colNum - right.colNum
}

function compareAxisKey(
  left: WorkbookSourceProjection['rowMetadata'][number],
  right: WorkbookSourceProjection['rowMetadata'][number],
): number {
  return left.sheetName.localeCompare(right.sheetName) || left.startIndex - right.startIndex || left.count - right.count
}
