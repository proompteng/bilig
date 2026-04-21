import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { FormulaNode } from '@bilig/formula'
import { SpreadsheetEngine } from '../engine.js'
import { EngineFormulaEvaluationError } from '../engine/errors.js'
import type { RuntimeFormula } from '../engine/runtime-state.js'
import type { EngineFormulaEvaluationService } from '../engine/services/formula-evaluation-service.js'
import type { EngineMutationSupportService } from '../engine/services/mutation-support-service.js'

function isEngineFormulaEvaluationService(value: unknown): value is EngineFormulaEvaluationService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'evaluateUnsupportedFormula') === 'function' &&
    typeof Reflect.get(value, 'resolveStructuredReference') === 'function' &&
    typeof Reflect.get(value, 'resolveSpillReference') === 'function'
  )
}

function isEngineMutationSupportService(value: unknown): value is EngineMutationSupportService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'clearOwnedSpill') === 'function'
}

function isRuntimeFormulaTable(value: unknown): value is { get(cellIndex: number): RuntimeFormula | undefined } {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'get') === 'function'
}

function getEvaluationService(engine: SpreadsheetEngine): EngineFormulaEvaluationService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const evaluation = Reflect.get(runtime, 'evaluation')
  if (!isEngineFormulaEvaluationService(evaluation)) {
    throw new TypeError('Expected engine formula evaluation service')
  }
  return evaluation
}

function getMutationSupportService(engine: SpreadsheetEngine): EngineMutationSupportService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const support = Reflect.get(runtime, 'support')
  if (!isEngineMutationSupportService(support)) {
    throw new TypeError('Expected engine mutation support service')
  }
  return support
}

function getInternalFormulaStore(engine: SpreadsheetEngine): { get(cellIndex: number): RuntimeFormula | undefined } {
  const formulas = Reflect.get(engine, 'formulas')
  if (!isRuntimeFormulaTable(formulas)) {
    throw new TypeError('Expected internal formulas store')
  }
  return formulas
}

function readRuntimeDirectLookupKind(engine: SpreadsheetEngine, sheetName: string, address: string): string | undefined {
  const formulas = getInternalFormulaStore(engine)
  const cellIndex = engine.workbook.getCellIndex(sheetName, address)
  if (cellIndex === undefined) {
    throw new Error(`expected runtime formula at ${sheetName}!${address}`)
  }
  const runtimeFormula = formulas.get(cellIndex)
  if (typeof runtimeFormula !== 'object' || runtimeFormula === null) {
    throw new Error(`expected runtime formula at ${sheetName}!${address}`)
  }
  const directLookup = Reflect.get(runtimeFormula, 'directLookup')
  if (typeof directLookup !== 'object' || directLookup === null) {
    return undefined
  }
  const kind = Reflect.get(directLookup, 'kind')
  return typeof kind === 'string' ? kind : undefined
}

describe('EngineFormulaEvaluationService', () => {
  it('re-evaluates JS indirection spills through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-indirect' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellFormula('Sheet1', 'G1', 'INDIRECT("B1:B2")')

    const g1Index = engine.workbook.getCellIndex('Sheet1', 'G1')
    expect(g1Index).toBeDefined()

    Effect.runSync(getMutationSupportService(engine).clearOwnedSpill(g1Index!))

    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toBeUndefined()

    Effect.runSync(getEvaluationService(engine).evaluateUnsupportedFormula(g1Index!))

    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'G1', rows: 2, cols: 1 }])
    expect(Effect.runSync(getEvaluationService(engine).resolveSpillReference('Sheet1', undefined, 'G1'))).toEqual({
      kind: 'RangeRef',
      refKind: 'cells',
      sheetName: 'Sheet1',
      start: 'G1',
      end: 'G2',
    } satisfies FormulaNode)
  })

  it('resolves structured references to table body rows through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-structured-ref' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Amount', 'Total'],
      headerRow: true,
      totalsRow: false,
    })

    const resolved = Effect.runSync(getEvaluationService(engine).resolveStructuredReference('Sales', 'Amount'))

    expect(resolved).toEqual({
      kind: 'RangeRef',
      refKind: 'cells',
      sheetName: 'Sheet1',
      start: 'A2',
      end: 'A3',
    } satisfies FormulaNode)

    expect(Effect.runSync(getEvaluationService(engine).resolveStructuredReference('Missing', 'Amount'))).toBeUndefined()
    expect(Effect.runSync(getEvaluationService(engine).resolveStructuredReference('Sales', 'Missing'))).toBeUndefined()

    engine.setTable({
      name: 'HeaderOnly',
      sheetName: 'Sheet1',
      startAddress: 'D1',
      endAddress: 'D1',
      columnNames: ['Amount'],
      headerRow: true,
      totalsRow: false,
    })
    expect(Effect.runSync(getEvaluationService(engine).resolveStructuredReference('HeaderOnly', 'Amount'))).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Ref,
    } satisfies FormulaNode)

    expect(Effect.runSync(getEvaluationService(engine).resolveSpillReference('Sheet1', undefined, 'Z1'))).toBeUndefined()
  })

  it('resolves MULTIPLE.OPERATIONS through reference replacements and missing formula cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-multiple-operations' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellValue('Sheet1', 'A2', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')

    const evaluation = getEvaluationService(engine)
    expect(
      Effect.runSync(
        evaluation.resolveMultipleOperations({
          formulaSheetName: 'Sheet1',
          formulaAddress: 'C1',
          rowCellSheetName: 'Sheet1',
          rowCellAddress: 'A1',
          rowReplacementSheetName: 'Sheet1',
          rowReplacementAddress: 'A2',
          columnCellSheetName: 'Sheet1',
          columnCellAddress: 'B1',
          columnReplacementSheetName: 'Sheet1',
          columnReplacementAddress: 'B2',
        }),
      ),
    ).toEqual({ tag: ValueTag.Number, value: 30 })

    expect(
      Effect.runSync(
        evaluation.resolveMultipleOperations({
          formulaSheetName: 'Sheet1',
          formulaAddress: 'Z99',
          rowCellSheetName: 'Sheet1',
          rowCellAddress: 'A1',
          rowReplacementSheetName: 'Sheet1',
          rowReplacementAddress: 'A2',
        }),
      ),
    ).toEqual({ tag: ValueTag.Empty })
  })

  it('returns empty results for non-formula cells and evaluates literal MATCH through the lookup resolver', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'evaluation-lookup-resolver',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'apple')
    engine.setCellValue('Sheet1', 'A2', 'pear')
    engine.setCellValue('Sheet1', 'A3', 'plum')
    engine.setCellFormula('Sheet1', 'B1', 'MATCH("pear",A1:A3,0)')

    const evaluation = getEvaluationService(engine)
    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()

    expect(Effect.runSync(evaluation.evaluateDirectLookupFormula(a1Index!))).toBeUndefined()
    expect(Effect.runSync(evaluation.evaluateUnsupportedFormula(a1Index!))).toEqual([])

    Effect.runSync(evaluation.evaluateUnsupportedFormula(b1Index!))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('evaluates full-column MATCH formulas through the generic vector lookup fallback', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-full-column-match' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'A3', 5)
    engine.setCellValue('Sheet1', 'A4', 'apple')
    engine.setCellValue('Sheet1', 'A5', 'pear')
    engine.setCellFormula('Sheet1', 'F1', 'MATCH(4,A:A,1)')
    engine.setCellFormula('Sheet1', 'F2', 'MATCH("pear",A:A,0)')

    const evaluation = getEvaluationService(engine)
    const f1Index = engine.workbook.getCellIndex('Sheet1', 'F1')
    const f2Index = engine.workbook.getCellIndex('Sheet1', 'F2')
    expect(f1Index).toBeDefined()
    expect(f2Index).toBeDefined()

    Effect.runSync(evaluation.evaluateUnsupportedFormula(f1Index!))
    Effect.runSync(evaluation.evaluateUnsupportedFormula(f2Index!))

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('treats missing external sheets as #REF! during JS evaluation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-missing-external-sheet' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'Sheet2!B1*2')
    engine.setCellFormula('Sheet1', 'A2', 'SUM(Sheet2!A1:A2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
  })

  it('wraps workbook access failures from structured, spill, and multiple-operations helpers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-error-wrappers' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const evaluation = getEvaluationService(engine)

    const getTableSpy = vi.spyOn(engine.workbook, 'getTable').mockImplementation(() => {
      throw new Error('structured explode')
    })
    const structured = Effect.runSync(Effect.either(evaluation.resolveStructuredReference('Sales', 'Amount')))
    expect(structured._tag).toBe('Left')
    expect(structured.left).toBeInstanceOf(EngineFormulaEvaluationError)
    expect(structured.left.message).toContain('structured explode')
    getTableSpy.mockRestore()

    const getSpillSpy = vi.spyOn(engine.workbook, 'getSpill').mockImplementation(() => {
      throw new Error('spill explode')
    })
    const spill = Effect.runSync(Effect.either(evaluation.resolveSpillReference('Sheet1', undefined, 'A1')))
    expect(spill._tag).toBe('Left')
    expect(spill.left).toBeInstanceOf(EngineFormulaEvaluationError)
    expect(spill.left.message).toContain('spill explode')
    getSpillSpy.mockRestore()

    const getCellIndexSpy = vi.spyOn(engine.workbook, 'getCellIndex').mockImplementation(() => {
      throw new Error('multiple operations explode')
    })
    const multipleOperations = Effect.runSync(
      Effect.either(
        evaluation.resolveMultipleOperations({
          formulaSheetName: 'Sheet1',
          formulaAddress: 'A1',
          rowCellSheetName: 'Sheet1',
          rowCellAddress: 'A1',
          rowReplacementSheetName: 'Sheet1',
          rowReplacementAddress: 'A2',
        }),
      ),
    )
    expect(multipleOperations._tag).toBe('Left')
    expect(multipleOperations.left).toBeInstanceOf(EngineFormulaEvaluationError)
    expect(multipleOperations.left.message).toContain('multiple operations explode')
    getCellIndexSpy.mockRestore()
  })

  it('wraps direct-lookup and unsupported-formula evaluation failures', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'evaluation-top-level-wrapper-errors',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'MATCH(2,A1:A3,0)')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A3)')

    const evaluation = getEvaluationService(engine)
    const formulas = getInternalFormulaStore(engine)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const c1Index = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(b1Index).toBeDefined()
    expect(c1Index).toBeDefined()

    const getSpy = vi.spyOn(formulas, 'get').mockImplementation(() => {
      throw new Error('formula explode')
    })
    const directLookup = Effect.runSync(Effect.either(evaluation.evaluateDirectLookupFormula(b1Index!)))
    expect(directLookup._tag).toBe('Left')
    expect(directLookup.left).toBeInstanceOf(EngineFormulaEvaluationError)
    expect(directLookup.left.message).toContain('formula explode')
    getSpy.mockRestore()

    const getSheetNameByIdSpy = vi.spyOn(engine.workbook, 'getSheetNameById').mockImplementation(() => {
      throw new Error('unsupported explode')
    })
    const unsupported = Effect.runSync(Effect.either(evaluation.evaluateUnsupportedFormula(c1Index!)))
    expect(unsupported._tag).toBe('Left')
    expect(unsupported.left).toBeInstanceOf(EngineFormulaEvaluationError)
    expect(unsupported.left.message).toContain('unsupported explode')
    getSheetNameByIdSpy.mockRestore()
  })

  it('evaluates direct exact lookup formulas across uniform, text, and mixed columns', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'evaluation-direct-exact-service',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 1)
    engine.setCellValue('Sheet1', 'C1', 'pear')
    engine.setCellValue('Sheet1', 'C2', 'apple')
    engine.setCellValue('Sheet1', 'C3', 'pear')
    engine.setCellValue('Sheet1', 'E2', 'pear')
    engine.setCellValue('Sheet1', 'E3', false)
    engine.setCellValue('Sheet1', 'D1', 2.5)
    engine.setCellValue('Sheet1', 'D2', 2)
    engine.setCellValue('Sheet1', 'D3', false)
    engine.setCellValue('Sheet1', 'D4', false)

    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D1,A1:A3,0)')
    engine.setCellFormula('Sheet1', 'F2', 'MATCH(D2,B1:B3,0)')
    engine.setCellFormula('Sheet1', 'F3', 'MATCH(D3,C1:C3,0)')
    engine.setCellFormula('Sheet1', 'F4', 'MATCH(D4,E1:E3,0)')

    const evaluation = getEvaluationService(engine)
    const f1Index = engine.workbook.getCellIndex('Sheet1', 'F1')
    const f2Index = engine.workbook.getCellIndex('Sheet1', 'F2')
    const f3Index = engine.workbook.getCellIndex('Sheet1', 'F3')
    const f4Index = engine.workbook.getCellIndex('Sheet1', 'F4')
    expect(f1Index).toBeDefined()
    expect(f2Index).toBeDefined()
    expect(f3Index).toBeDefined()
    expect(f4Index).toBeDefined()

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!))
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    engine.setCellValue('Sheet1', 'D1', 2)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!))
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[4], [5], [6]])
    engine.setCellValue('Sheet1', 'D1', 5)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!))
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!))
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B3' }, [[6], [5], [4]])
    engine.setCellValue('Sheet1', 'D2', 5)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!))
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 2 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!))
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f4Index!))
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 3 })
  })

  it('evaluates direct approximate lookup formulas across uniform, refreshed, and text columns', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-approx-service' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 1)
    engine.setCellValue('Sheet1', 'C1', 'apple')
    engine.setCellValue('Sheet1', 'C2', 'banana')
    engine.setCellValue('Sheet1', 'C3', 'pear')
    engine.setCellValue('Sheet1', 'D1', true)
    engine.setCellValue('Sheet1', 'D2', 2.5)
    engine.setCellValue('Sheet1', 'D3', 'peach')
    engine.setCellValue('Sheet1', 'D4', 5)

    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D1,A1:A3,1)')
    engine.setCellFormula('Sheet1', 'F2', 'MATCH(D2,B1:B3,-1)')
    engine.setCellFormula('Sheet1', 'F3', 'MATCH(D3,C1:C3,1)')
    engine.setCellFormula('Sheet1', 'F4', 'MATCH(D4,C1:C3,1)')

    const evaluation = getEvaluationService(engine)
    const f1Index = engine.workbook.getCellIndex('Sheet1', 'F1')
    const f2Index = engine.workbook.getCellIndex('Sheet1', 'F2')
    const f3Index = engine.workbook.getCellIndex('Sheet1', 'F3')
    const f4Index = engine.workbook.getCellIndex('Sheet1', 'F4')
    expect(f1Index).toBeDefined()
    expect(f2Index).toBeDefined()
    expect(f3Index).toBeDefined()
    expect(f4Index).toBeDefined()

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!))
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 1 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!))
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 1 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!))
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 2 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f4Index!))
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellValue('Sheet1', 'A3', 5)
    engine.setCellValue('Sheet1', 'D1', 4.5)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!))
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B3' }, [[6], [5], [4]])
    engine.setCellValue('Sheet1', 'D2', 4.5)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!))
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValue('Sheet1', 'C2', 'blueberry')
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!))
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('refreshes direct uniform numeric lookup descriptors across exact and approximate branches', async () => {
    const exactEngine = new SpreadsheetEngine({
      workbookName: 'evaluation-direct-exact-uniform-refresh',
      useColumnIndex: true,
    })
    await exactEngine.ready()
    exactEngine.createSheet('Sheet1')
    exactEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[1], [2], [3], [4]])
    exactEngine.setCellValue('Sheet1', 'B1', 3)
    exactEngine.setCellFormula('Sheet1', 'C1', 'MATCH(B1,A1:A4,0)')

    const exactEvaluation = getEvaluationService(exactEngine)
    const exactIndex = exactEngine.workbook.getCellIndex('Sheet1', 'C1')
    expect(exactIndex).toBeDefined()
    expect(readRuntimeDirectLookupKind(exactEngine, 'Sheet1', 'C1')).toBe('exact-uniform-numeric')

    Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))
    expect(exactEngine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })

    exactEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[12], [22], [32], [42]])
    exactEngine.setCellValue('Sheet1', 'B1', 32)
    Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))
    expect(exactEngine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(readRuntimeDirectLookupKind(exactEngine, 'Sheet1', 'C1')).toBe('exact-uniform-numeric')

    exactEngine.setCellValue('Sheet1', 'B1', '32')
    Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))
    expect(exactEngine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    exactEngine.setCellFormula('Sheet1', 'B1', '1/0')
    expect(Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))).toBeUndefined()

    const approximateEngine = new SpreadsheetEngine({
      workbookName: 'evaluation-direct-approx-uniform-refresh',
    })
    await approximateEngine.ready()
    approximateEngine.createSheet('Sheet1')
    approximateEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' }, [
      [10, 40],
      [20, 30],
      [30, 20],
      [40, 10],
    ])
    approximateEngine.setCellValue('Sheet1', 'D1', 35)
    approximateEngine.setCellValue('Sheet1', 'E1', 25)
    approximateEngine.setCellFormula('Sheet1', 'F1', 'MATCH(D1,A1:A4,1)')
    approximateEngine.setCellFormula('Sheet1', 'F2', 'MATCH(E1,B1:B4,-1)')

    const approximateEvaluation = getEvaluationService(approximateEngine)
    const ascendingIndex = approximateEngine.workbook.getCellIndex('Sheet1', 'F1')
    const descendingIndex = approximateEngine.workbook.getCellIndex('Sheet1', 'F2')
    expect(ascendingIndex).toBeDefined()
    expect(descendingIndex).toBeDefined()
    expect(readRuntimeDirectLookupKind(approximateEngine, 'Sheet1', 'F1')).toBe('approximate-uniform-numeric')
    expect(readRuntimeDirectLookupKind(approximateEngine, 'Sheet1', 'F2')).toBe('approximate-uniform-numeric')

    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    approximateEngine.setCellValue('Sheet1', 'D1', null)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    approximateEngine.setCellValue('Sheet1', 'D1', 50)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })

    approximateEngine.setCellValue('Sheet1', 'D1', 'pear')
    expect(Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))).toBeUndefined()

    approximateEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[100], [90], [80], [70]])
    approximateEngine.setCellValue('Sheet1', 'D1', 85)
    expect(Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))).toBeUndefined()

    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })

    approximateEngine.setCellValue('Sheet1', 'E1', 50)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    approximateEngine.setCellValue('Sheet1', 'E1', 5)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
  })

  it('evaluates direct aggregate formulas with progressive prefixes and coercion rules', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-aggregate-service' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A4', 'skip')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A4)')
    engine.setCellFormula('Sheet1', 'B2', 'AVERAGE(A1:A4)')
    engine.setCellFormula('Sheet1', 'B3', 'COUNT(A1:A4)')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A1)')
    engine.setCellFormula('Sheet1', 'C2', 'SUM(A1:A2)')
    engine.setCellFormula('Sheet1', 'C3', 'SUM(A1:A4)')
    engine.setCellFormula('Sheet1', 'C4', 'SUM(A2:A4)')
    engine.setCellFormula('Sheet1', 'C5', 'AVERAGE(A2:A4)')

    const evaluation = getEvaluationService(engine)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const b2Index = engine.workbook.getCellIndex('Sheet1', 'B2')
    const b3Index = engine.workbook.getCellIndex('Sheet1', 'B3')
    const c1Index = engine.workbook.getCellIndex('Sheet1', 'C1')
    const c2Index = engine.workbook.getCellIndex('Sheet1', 'C2')
    const c3Index = engine.workbook.getCellIndex('Sheet1', 'C3')
    const c4Index = engine.workbook.getCellIndex('Sheet1', 'C4')
    const c5Index = engine.workbook.getCellIndex('Sheet1', 'C5')
    expect(b1Index).toBeDefined()
    expect(b2Index).toBeDefined()
    expect(b3Index).toBeDefined()
    expect(c1Index).toBeDefined()
    expect(c2Index).toBeDefined()
    expect(c3Index).toBeDefined()
    expect(c4Index).toBeDefined()
    expect(c5Index).toBeDefined()

    Effect.runSync(evaluation.evaluateDirectLookupFormula(c1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c3Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c4Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c5Index!))
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C5')).toEqual({ tag: ValueTag.Number, value: 0.5 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(b1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b3Index!))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellFormula('Sheet1', 'A4', 'NA()')
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b3Index!))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('evaluates direct aggregate formulas with formula members from live cell state', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-aggregate-live' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'A2', 'A1*3')
    engine.setCellFormula('Sheet1', 'A3', 'IF(A1>0,NA(),"ok")')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'B2', 'COUNT(A1:A3)')
    engine.setCellFormula('Sheet1', 'B3', 'MIN(A1:A3)')
    engine.setCellFormula('Sheet1', 'B4', 'MAX(A1:A3)')

    const evaluation = getEvaluationService(engine)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const b2Index = engine.workbook.getCellIndex('Sheet1', 'B2')
    const b3Index = engine.workbook.getCellIndex('Sheet1', 'B3')
    const b4Index = engine.workbook.getCellIndex('Sheet1', 'B4')
    expect(b1Index).toBeDefined()
    expect(b2Index).toBeDefined()
    expect(b3Index).toBeDefined()
    expect(b4Index).toBeDefined()

    Effect.runSync(evaluation.evaluateDirectLookupFormula(b1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b3Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b4Index!))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 6 })
  })
})
