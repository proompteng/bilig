import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { ErrorCode, FormulaMode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { fastPathFormulaArbitrary } from '../../../formula/src/__tests__/formula-fuzz-helpers.js'
import { runProperty } from '@bilig/test-fuzz'

const sheetName = 'Sheet1'
const lookupKeys = ['apple', 'banana', 'pear', 'plum', 'quince', 'rice'] as const
const lookupFormulaKindArbitrary = fc.constantFrom('VLOOKUP', 'HLOOKUP', 'XLOOKUP' as const)
const lookupReturnValueArbitrary = fc.oneof(fc.integer({ min: -100, max: 100 }), fc.constant(null))

interface LookupFormulaCase {
  readonly formula: string
  readonly range: { startAddress: string; endAddress: string }
  readonly values: LiteralInput[][]
  readonly expected: CellValue
}

const fastPathLookupFormulaCaseArbitrary = fc
  .uniqueArray(fc.constantFrom(...lookupKeys), { minLength: 2, maxLength: 5 })
  .chain((keys) =>
    fc.record({
      kind: lookupFormulaKindArbitrary,
      keys: fc.constant(keys),
      returns: fc.array(lookupReturnValueArbitrary, {
        minLength: keys.length,
        maxLength: keys.length,
      }),
      query: fc.oneof(fc.constantFrom(...keys), fc.constant('missing')),
    }),
  )
  .map(({ kind, keys, returns, query }) => buildLookupFormulaCase(kind, keys, returns, query))

function buildNumericGrid(rows: number, cols: number): number[][] {
  return Array.from({ length: rows }, (_rowValue, row) => Array.from({ length: cols }, (_colValue, col) => row * cols + col + 1))
}

describe('formula runtime differential fuzz', () => {
  it('keeps generated fast-path formulas in JS and wasm parity', async () => {
    await runProperty({
      suite: 'core/formula-runtime/generated-differential',
      arbitrary: fastPathFormulaArbitrary,
      predicate: async (formula) => {
        const engine = new SpreadsheetEngine({
          workbookName: `fuzz-formula-diff-${formula.length}`,
          replicaId: 'fuzz-formula-diff',
        })
        await engine.ready()
        engine.createSheet(sheetName)
        engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'F6' }, buildNumericGrid(6, 6))
        engine.setCellFormula(sheetName, 'G1', formula)

        const explanation = engine.explainCell(sheetName, 'G1')
        expect(explanation.mode).toBe(FormulaMode.WasmFastPath)

        const differential = engine.recalculateDifferential()
        expect(differential.drift).toEqual([])

        const snapshot = engine.exportSnapshot()
        const restored = new SpreadsheetEngine({
          workbookName: snapshot.workbook.name,
          replicaId: 'fuzz-formula-diff-restored',
        })
        await restored.ready()
        restored.importSnapshot(snapshot)

        expect(restored.getCellValue(sheetName, 'G1')).toEqual(engine.getCellValue(sheetName, 'G1'))
      },
    })
  })

  it('keeps exact lookup fast-path formulas in JS and wasm parity', async () => {
    await runProperty({
      suite: 'core/formula-runtime/generated-lookup-differential',
      arbitrary: fastPathLookupFormulaCaseArbitrary,
      parameters: { numRuns: 75, interruptAfterTimeLimit: 15_000 },
      predicate: async (lookupCase) => {
        const engine = new SpreadsheetEngine({
          workbookName: `fuzz-lookup-diff-${lookupCase.formula.length}`,
          replicaId: 'fuzz-lookup-diff',
        })
        await engine.ready()
        engine.createSheet(sheetName)
        engine.setRangeValues({ sheetName, ...lookupCase.range }, lookupCase.values)
        engine.setCellFormula(sheetName, 'G1', lookupCase.formula)

        expect(engine.explainCell(sheetName, 'G1').mode).toBe(FormulaMode.WasmFastPath)
        expectCellSemantics(engine.getCellValue(sheetName, 'G1'), lookupCase.expected)

        const differential = engine.recalculateDifferential()
        expect(differential.drift).toEqual([])

        const snapshot = engine.exportSnapshot()
        const restored = new SpreadsheetEngine({
          workbookName: snapshot.workbook.name,
          replicaId: 'fuzz-lookup-diff-restored',
        })
        await restored.ready()
        restored.importSnapshot(snapshot)

        expectCellSemantics(restored.getCellValue(sheetName, 'G1'), lookupCase.expected)
      },
    })
  })
})

function expectCellSemantics(actual: CellValue, expected: CellValue): void {
  if (expected.tag === ValueTag.String) {
    expect(actual).toMatchObject({ tag: ValueTag.String, value: expected.value })
    return
  }
  expect(actual).toEqual(expected)
}

function buildLookupFormulaCase(
  kind: 'VLOOKUP' | 'HLOOKUP' | 'XLOOKUP',
  keys: readonly (typeof lookupKeys)[number][],
  returns: readonly (number | null)[],
  query: (typeof lookupKeys)[number] | 'missing',
): LookupFormulaCase {
  const matchIndex = keys.findIndex((key) => key === query)
  const expected = matchIndex === -1 ? lookupMissValue(kind) : lookupReturnValue(returns[matchIndex] ?? null)
  const quotedQuery = quoteFormulaString(query)
  if (kind === 'HLOOKUP') {
    const endAddress = `${columnName(keys.length - 1)}2`
    const keyRow: LiteralInput[] = [...keys]
    return {
      formula: `HLOOKUP(${quotedQuery},A1:${endAddress},2,FALSE)`,
      range: { startAddress: 'A1', endAddress },
      values: [keyRow, [...returns]],
      expected,
    }
  }
  const endAddress = `B${keys.length}`
  const values = keys.map((key, index) => [key, returns[index] ?? null])
  return {
    formula:
      kind === 'VLOOKUP'
        ? `VLOOKUP(${quotedQuery},A1:${endAddress},2,FALSE)`
        : `XLOOKUP(${quotedQuery},A1:A${keys.length},B1:B${keys.length},"missing")`,
    range: { startAddress: 'A1', endAddress },
    values,
    expected,
  }
}

function lookupMissValue(kind: 'VLOOKUP' | 'HLOOKUP' | 'XLOOKUP'): CellValue {
  return kind === 'XLOOKUP' ? stringValue('missing') : { tag: ValueTag.Error, code: ErrorCode.NA }
}

function lookupReturnValue(value: number | null): CellValue {
  return { tag: ValueTag.Number, value: value ?? 0 }
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function quoteFormulaString(value: string): string {
  return `"${value.replace(/"/g, '""')}"`
}

function columnName(index: number): string {
  return String.fromCharCode('A'.charCodeAt(0) + index)
}
