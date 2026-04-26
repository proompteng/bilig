import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { parseCsv } from '../csv.js'
import { runProperty } from '@bilig/test-fuzz'

describe('engine import/export fuzz', () => {
  it('preserves CSV-representable single-sheet semantics across export and import', async () => {
    await runProperty({
      suite: 'core/import-export/csv-roundtrip-parity',
      arbitrary: csvSheetModelArbitrary(),
      predicate: async (model) => {
        // Arrange
        const original = await createEngineFromCsvModel('fuzz-csv-original', model)

        // Act
        const exported = original.exportSheetCsv(sheetName)
        const restored = new SpreadsheetEngine({
          workbookName: 'fuzz-csv-restored',
          replicaId: 'fuzz-csv-restored',
        })
        await restored.ready()
        restored.createSheet(sheetName)
        restored.importSheetCsv(sheetName, exported)
        const reexported = restored.exportSheetCsv(sheetName)

        // Assert
        expect(reexported).toBe(exported)
        const bounds = inferCsvBounds(exported)
        expect(projectCsvSemanticSheet(restored, bounds)).toEqual(projectCsvSemanticSheet(original, bounds))
      },
    })
  })
})

// Helpers

const sheetName = 'Sheet1'

type CsvValueCellModel = {
  kind: 'value'
  value: boolean | number | string
}

type CsvFormulaCellModel = {
  kind: 'formula'
  formula: string
}

type CsvCellModel = CsvValueCellModel | CsvFormulaCellModel | { kind: 'empty' }

type CsvSheetModel = {
  rowCount: number
  colCount: number
  cells: CsvCellModel[][]
}

type CsvBounds = {
  rowCount: number
  colCount: number
}

type CsvSemanticCell = {
  formula: string | null
  value:
    | null
    | boolean
    | number
    | string
    | {
        error: number
      }
}

function csvSheetModelArbitrary(): fc.Arbitrary<CsvSheetModel> {
  return fc.oneof(
    fc.constant({
      rowCount: 0,
      colCount: 0,
      cells: [],
    } satisfies CsvSheetModel),
    fc
      .record({
        rowCount: fc.integer({ min: 1, max: 5 }),
        colCount: fc.integer({ min: 1, max: 5 }),
      })
      .chain(({ rowCount, colCount }) => {
        const positions = Array.from({ length: rowCount * colCount }, (_entry, index) => index)
        return fc
          .tuple(
            fc.array(csvLiteralCellModelArbitrary(), {
              minLength: positions.length,
              maxLength: positions.length,
            }),
            fc.subarray(positions, { maxLength: Math.min(positions.length, 6) }),
          )
          .chain(([baseCells, formulaTargets]) => {
            const formulaTargetSet = new Set(formulaTargets)
            const referenceAddresses = positions
              .filter((index) => !formulaTargetSet.has(index))
              .map((index) => addressForIndex(index, colCount))
            const formulaCellsArbitrary =
              referenceAddresses.length === 0
                ? fc.constant([] as CsvFormulaCellModel[])
                : fc.array(csvFormulaCellModelArbitrary(referenceAddresses), {
                    minLength: formulaTargets.length,
                    maxLength: formulaTargets.length,
                  })

            return formulaCellsArbitrary.map((formulaCells) => {
              const cells = Array.from({ length: rowCount }, () => Array.from<CsvCellModel>({ length: colCount }))
              const formulaByIndex = new Map<number, CsvFormulaCellModel>()
              formulaTargets.forEach((target, index) => {
                const formulaCell = formulaCells[index]
                if (formulaCell) {
                  formulaByIndex.set(target, formulaCell)
                }
              })
              positions.forEach((index) => {
                const row = Math.floor(index / colCount)
                const col = index % colCount
                const rowCells = cells[row]
                if (!rowCells) {
                  throw new Error(`Missing generated row for CSV model at ${row}`)
                }
                rowCells[col] = formulaByIndex.get(index) ?? baseCells[index] ?? { kind: 'empty' }
              })
              return {
                rowCount,
                colCount,
                cells,
              }
            })
          })
      }),
  )
}

function csvLiteralCellModelArbitrary(): fc.Arbitrary<CsvCellModel> {
  return fc.oneof(
    fc.constant({ kind: 'empty' } satisfies CsvCellModel),
    fc.integer().map((value) => ({ kind: 'value', value }) satisfies CsvCellModel),
    fc.boolean().map((value) => ({ kind: 'value', value }) satisfies CsvCellModel),
    fc.string().map((value) => ({ kind: 'value', value: `text:${value}` }) satisfies CsvCellModel),
  )
}

function csvFormulaCellModelArbitrary(referenceAddresses: readonly string[]): fc.Arbitrary<CsvFormulaCellModel> {
  const referenceArbitrary = fc.constantFrom(...referenceAddresses)
  return fc.oneof(
    fc.tuple(referenceArbitrary, fc.constantFrom('+', '-', '*', '/'), referenceArbitrary).map(
      ([left, operator, right]) =>
        ({
          kind: 'formula',
          formula: `${left}${operator}${right}`,
        }) satisfies CsvFormulaCellModel,
    ),
    fc.tuple(referenceArbitrary, referenceArbitrary).map(([left, right]) => {
      const leftAddress = parseCellAddress(left, sheetName)
      const rightAddress = parseCellAddress(right, sheetName)
      const start = formatAddress(Math.min(leftAddress.row, rightAddress.row), Math.min(leftAddress.col, rightAddress.col))
      const end = formatAddress(Math.max(leftAddress.row, rightAddress.row), Math.max(leftAddress.col, rightAddress.col))
      return {
        kind: 'formula',
        formula: `SUM(${start}:${end})`,
      } satisfies CsvFormulaCellModel
    }),
    referenceArbitrary.map(
      (reference) =>
        ({
          kind: 'formula',
          formula: `IF(${reference}>0,"text:yes","text:no")`,
        }) satisfies CsvFormulaCellModel,
    ),
  )
}

function addressForIndex(index: number, colCount: number): string {
  return formatAddress(Math.floor(index / colCount), index % colCount)
}

async function createEngineFromCsvModel(workbookName: string, model: CsvSheetModel): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName,
    replicaId: workbookName,
  })
  await engine.ready()
  engine.createSheet(sheetName)
  let hasFormula = false
  for (let row = 0; row < model.rowCount; row += 1) {
    for (let col = 0; col < model.colCount; col += 1) {
      const cell = model.cells[row]?.[col] ?? { kind: 'empty' }
      const address = formatAddress(row, col)
      if (cell.kind === 'value') {
        engine.setRangeValues({ sheetName, startAddress: address, endAddress: address }, [[cell.value]])
      } else if (cell.kind === 'formula') {
        engine.setCellFormula(sheetName, address, cell.formula)
        hasFormula = true
      }
    }
  }
  if (hasFormula) {
    engine.recalculateNow()
    engine.recalculateNow()
  }
  return engine
}

function inferCsvBounds(csv: string): CsvBounds {
  const rows = parseCsv(csv)
  return {
    rowCount: rows.length,
    colCount: rows.reduce((max, row) => Math.max(max, row.length), 0),
  }
}

function projectCsvSemanticSheet(engine: SpreadsheetEngine, bounds: CsvBounds): CsvSemanticCell[][] {
  return Array.from({ length: bounds.rowCount }, (_rowEntry, row) =>
    Array.from({ length: bounds.colCount }, (_colEntry, col) => {
      const cell = engine.getCell(sheetName, formatAddress(row, col))
      return {
        formula: cell.formula ?? null,
        value: simplifyCellValue(cell.value),
      }
    }),
  )
}

function simplifyCellValue(value: CellValue): CsvSemanticCell['value'] {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Number:
      return value.value
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return { error: value.code }
  }
}
