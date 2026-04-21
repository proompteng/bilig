import { afterEach, describe, expect, it, vi } from 'vitest'
import { utcDateToExcelSerial } from '@bilig/formula'
import { SpreadsheetEngine, type EngineSyncClient } from '../index.js'
import { ErrorCode, FormulaMode, Opcode, ValueTag, type EngineEvent } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'

type RuntimeFormulaWithDependencies = {
  dependencyIndices: Uint32Array
}

type RuntimeFormulaWithCompiled = {
  formulaSlotId: number
  planId: number
  compiled: {
    mode: number
    deps: string[]
    jsPlan: unknown[]
  }
  plan: {
    id: number
    source: string
    compiled: {
      mode: number
      deps: string[]
      jsPlan: unknown[]
    }
  }
  dependencyEntities: { ptr: number; len: number }
  runtimeProgram: Uint32Array
}

type RuntimeFormulaWithRanges = {
  rangeDependencies: Uint32Array
  runtimeProgram: Uint32Array
}

type RuntimeFormulaWithDirectLookup = {
  directLookup:
    | {
        kind: 'exact'
        operandCellIndex: number
        prepared: { sheetName: string; rowStart: number; rowEnd: number; col: number }
        searchMode: 1 | -1
      }
    | {
        kind: 'exact-uniform-numeric'
        operandCellIndex: number
        sheetName: string
        rowStart: number
        rowEnd: number
        col: number
        length: number
        searchMode: 1 | -1
      }
    | {
        kind: 'approximate'
        operandCellIndex: number
        prepared: { sheetName: string; rowStart: number; rowEnd: number; col: number }
        matchMode: 1 | -1
      }
    | {
        kind: 'approximate-uniform-numeric'
        operandCellIndex: number
        sheetName: string
        rowStart: number
        rowEnd: number
        col: number
        length: number
        matchMode: 1 | -1
      }
}

type RuntimeFormulaWithDirectCriteria = {
  directCriteria: {
    aggregateKind: 'count' | 'sum' | 'average' | 'min' | 'max'
    aggregateRange:
      | {
          sheetName: string
          rowStart: number
          rowEnd: number
          col: number
          length: number
        }
      | undefined
    criteriaPairs: Array<{
      range: {
        sheetName: string
        rowStart: number
        rowEnd: number
        col: number
        length: number
      }
      criterion:
        | {
            kind: 'literal'
            value: unknown
          }
        | {
            kind: 'cell'
            cellIndex: number
          }
    }>
  }
}

type RuntimeFormulaWithDirectAggregate = {
  directAggregate: {
    aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
    sheetName: string
    rowStart: number
    rowEnd: number
    col: number
    length: number
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function hasFormulaStore(value: unknown): value is { formulas: { get(cellIndex: number): unknown } } {
  return (
    isRecord(value) &&
    'formulas' in value &&
    isRecord(value.formulas) &&
    'get' in value.formulas &&
    typeof value.formulas.get === 'function'
  )
}

function readRuntimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  if (!hasFormulaStore(engine)) {
    throw new Error('SpreadsheetEngine test expected an internal formulas store')
  }
  return engine.formulas.get(cellIndex)
}

function readRuntimeTemplateId(engine: SpreadsheetEngine, cellIndex: number): number | undefined {
  const runtimeFormula = readRuntimeFormula(engine, cellIndex)
  if (!isRecord(runtimeFormula)) {
    return undefined
  }
  return typeof runtimeFormula.templateId === 'number' ? runtimeFormula.templateId : undefined
}

function readRuntimeDirectScalar(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const runtimeFormula = readRuntimeFormula(engine, cellIndex)
  return isRecord(runtimeFormula) ? runtimeFormula.directScalar : undefined
}

function isRuntimeFormulaWithDependencies(value: unknown): value is RuntimeFormulaWithDependencies {
  return isRecord(value) && value.dependencyIndices instanceof Uint32Array
}

function isRuntimeFormulaWithCompiled(value: unknown): value is RuntimeFormulaWithCompiled {
  return (
    isRecord(value) &&
    typeof value.formulaSlotId === 'number' &&
    typeof value.planId === 'number' &&
    isRecord(value.compiled) &&
    typeof value.compiled.mode === 'number' &&
    Array.isArray(value.compiled.deps) &&
    Array.isArray(value.compiled.jsPlan) &&
    isRecord(value.plan) &&
    typeof value.plan.id === 'number' &&
    typeof value.plan.source === 'string' &&
    isRecord(value.plan.compiled) &&
    typeof value.plan.compiled.mode === 'number' &&
    Array.isArray(value.plan.compiled.deps) &&
    Array.isArray(value.plan.compiled.jsPlan) &&
    isRecord(value.dependencyEntities) &&
    typeof value.dependencyEntities.ptr === 'number' &&
    typeof value.dependencyEntities.len === 'number' &&
    value.runtimeProgram instanceof Uint32Array
  )
}

function isRuntimeFormulaWithRanges(value: unknown): value is RuntimeFormulaWithRanges {
  return isRecord(value) && value.rangeDependencies instanceof Uint32Array && value.runtimeProgram instanceof Uint32Array
}

function isRuntimeFormulaWithDirectLookup(value: unknown): value is RuntimeFormulaWithDirectLookup {
  if (!isRecord(value) || !('directLookup' in value) || !isRecord(value.directLookup)) {
    return false
  }
  const directLookup = value.directLookup
  const prepared = directLookup.prepared
  if (typeof directLookup.operandCellIndex !== 'number') {
    return false
  }
  if (directLookup.kind === 'exact') {
    return (
      isRecord(prepared) &&
      typeof prepared.sheetName === 'string' &&
      typeof prepared.rowStart === 'number' &&
      typeof prepared.rowEnd === 'number' &&
      typeof prepared.col === 'number' &&
      (directLookup.searchMode === 1 || directLookup.searchMode === -1)
    )
  }
  if (directLookup.kind === 'exact-uniform-numeric') {
    return (
      typeof directLookup.sheetName === 'string' &&
      typeof directLookup.rowStart === 'number' &&
      typeof directLookup.rowEnd === 'number' &&
      typeof directLookup.col === 'number' &&
      typeof directLookup.length === 'number' &&
      (directLookup.searchMode === 1 || directLookup.searchMode === -1)
    )
  }
  if (directLookup.kind === 'approximate') {
    return (
      isRecord(prepared) &&
      typeof prepared.sheetName === 'string' &&
      typeof prepared.rowStart === 'number' &&
      typeof prepared.rowEnd === 'number' &&
      typeof prepared.col === 'number' &&
      (directLookup.matchMode === 1 || directLookup.matchMode === -1)
    )
  }
  if (directLookup.kind === 'approximate-uniform-numeric') {
    return (
      typeof directLookup.sheetName === 'string' &&
      typeof directLookup.rowStart === 'number' &&
      typeof directLookup.rowEnd === 'number' &&
      typeof directLookup.col === 'number' &&
      typeof directLookup.length === 'number' &&
      (directLookup.matchMode === 1 || directLookup.matchMode === -1)
    )
  }
  return false
}

function isRuntimeFormulaWithDirectCriteria(value: unknown): value is RuntimeFormulaWithDirectCriteria {
  if (!isRecord(value) || !('directCriteria' in value) || !isRecord(value.directCriteria)) {
    return false
  }
  const directCriteria = value.directCriteria
  return (
    (directCriteria.aggregateKind === 'count' ||
      directCriteria.aggregateKind === 'sum' ||
      directCriteria.aggregateKind === 'average' ||
      directCriteria.aggregateKind === 'min' ||
      directCriteria.aggregateKind === 'max') &&
    Array.isArray(directCriteria.criteriaPairs)
  )
}

function isRuntimeFormulaWithDirectAggregate(value: unknown): value is RuntimeFormulaWithDirectAggregate {
  if (!isRecord(value) || !('directAggregate' in value) || !isRecord(value.directAggregate)) {
    return false
  }
  const directAggregate = value.directAggregate
  return (
    (directAggregate.aggregateKind === 'sum' ||
      directAggregate.aggregateKind === 'average' ||
      directAggregate.aggregateKind === 'count' ||
      directAggregate.aggregateKind === 'min' ||
      directAggregate.aggregateKind === 'max') &&
    typeof directAggregate.sheetName === 'string' &&
    typeof directAggregate.rowStart === 'number' &&
    typeof directAggregate.rowEnd === 'number' &&
    typeof directAggregate.col === 'number' &&
    typeof directAggregate.length === 'number'
  )
}

function seedPivotSource(engine: SpreadsheetEngine): void {
  engine.createSheet('Data')
  engine.createSheet('Pivot')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
    ['Region', 'Sales'],
    ['East', 10],
    ['West', 7],
  ])
}

describe('SpreadsheetEngine', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('recalculates simple formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 20 })

    engine.setCellValue('Sheet1', 'A1', 12)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 24 })
  })

  it('recalculateDirty performs incremental recalculation from regions', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'D1', 'C1+10')

    expect(engine.getCellValue('Sheet1', 'B1').value).toBe(20)
    expect(engine.getCellValue('Sheet1', 'D1').value).toBe(15)

    // Update A1 and C1 without recalculating immediately.
    // SpreadsheetEngine methods mostly call applyBatch which always recalculates.
    // We'll use the internal workbook store to bypass immediate calc.
    const a1Index = engine.workbook.ensureCell('Sheet1', 'A1')
    const c1Index = engine.workbook.ensureCell('Sheet1', 'C1')

    engine.workbook.cellStore.setValue(a1Index, { tag: ValueTag.Number, value: 50 })
    engine.workbook.cellStore.setValue(c1Index, { tag: ValueTag.Number, value: 100 })

    // Verify values are updated but NOT recalculated (B1 and D1 should have old results)
    expect(engine.getCellValue('Sheet1', 'A1').value).toBe(50)
    expect(engine.getCellValue('Sheet1', 'B1').value).toBe(20) // Still old result
    expect(engine.getCellValue('Sheet1', 'C1').value).toBe(100)
    expect(engine.getCellValue('Sheet1', 'D1').value).toBe(15) // Still old result

    // Recalculate only A1's region
    const changed = engine.recalculateDirty([{ sheetName: 'Sheet1', rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 }])

    // Should contain B1 (0,1) because it depends on A1 (0,0)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')!
    expect([...changed]).toContain(b1Index)

    // D1 (0,3) should NOT be in changed because we didn't mark C1 (0,2) as dirty
    const d1Index = engine.workbook.getCellIndex('Sheet1', 'D1')!
    expect([...changed]).not.toContain(d1Index)

    expect(engine.getCellValue('Sheet1', 'B1').value).toBe(100) // 50 * 2
    expect(engine.getCellValue('Sheet1', 'D1').value).toBe(15) // Still old value
  })

  it('evaluates string concatenation and string comparisons on the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'hello')
    engine.setCellFormula('Sheet1', 'B1', 'A1&" world"')
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    engine.setCellFormula('Sheet1', 'C1', 'A1="HELLO"')
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    engine.setCellFormula('Sheet1', 'D1', '"b">"A"')
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'hello world',
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Boolean, value: true })
  })

  it('reports differential drift when restored recalculation produces different changed-cell sets', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Sheet1', 'A1', 12)

    const differential = engine.recalculateDifferential()

    expect(differential.drift).toEqual([])
    expect(differential.wasm).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          sheetName: 'Sheet1',
          address: 'B1',
          value: { tag: ValueTag.Number, value: 24 },
        }),
      ]),
    )
    expect(differential.js).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          sheetName: 'Sheet1',
          address: 'B1',
          value: { tag: ValueTag.Number, value: 24 },
        }),
      ]),
    )
  })

  it('relocates relative formulas when copying a range', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 5)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' },
    )

    expect(engine.getCell('Sheet1', 'B2').formula).toBe('A2*2')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('preserves absolute references when copying formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellFormula('Sheet1', 'B1', '$A1+A$1+$A$1')

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C2' },
    )

    expect(engine.getCell('Sheet1', 'C2').formula).toBe('$A2+B$1+$A$1')
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 16 })
  })

  it('moves a range and clears the source cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B2', 'left')
    engine.setCellValue('Sheet1', 'C2', 'right')

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C2' },
      { sheetName: 'Sheet1', startAddress: 'D4', endAddress: 'E4' },
    )

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'left',
      }),
    )
    expect(engine.getCellValue('Sheet1', 'E4')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'right',
      }),
    )
  })

  it('treats copying empty cells into tracked empty dependencies as a history no-op', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'copy-undo-empty-targets-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    seed.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const snapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'copy-undo-empty-targets' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    const beforeCopy = engine.exportSnapshot()

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E6' },
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
    )

    expect(engine.exportSnapshot()).toEqual(beforeCopy)
    expect(engine.undo()).toBe(false)
    expect(engine.exportSnapshot()).toEqual(beforeCopy)
  })

  it('treats clearing an already-empty tracked dependency cell as a no-op', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'clear-empty-dependency-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')

    const before = engine.exportSnapshot()
    engine.clearCell('Sheet1', 'D4')

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats sheet-id clear mutations on already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'clear-empty-dependency-noop-by-id' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    const before = engine.exportSnapshot()
    engine.clearCellAt(sheetId, 3, 3)

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats copying empty cells into empty targets as a no-op', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'copy-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const before = engine.exportSnapshot()

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E6' },
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
    )

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats filling empty cells into empty targets as a no-op', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'fill-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const before = engine.exportSnapshot()

    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E6' },
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
    )

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('undoes formula fills into blank targets without leaving explicit empty cells behind', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'fill-undo-blank-target-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    const initialSnapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'fill-undo-blank-target' })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula('Sheet1', 'A1', 'E5+A1')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'E5', endAddress: 'E5' },
    )
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
    )
    engine.insertColumns('Sheet1', 0, 1)

    let undoCount = 0
    while (engine.undo()) {
      undoCount += 1
      expect(undoCount).toBeLessThanOrEqual(16)
    }
    expect(undoCount).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('undoes formula creation on tracked dependency placeholders without exporting authored blanks', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'formula-undo-placeholder-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    const initialSnapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'formula-undo-placeholder' })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula('Sheet1', 'A1', 'C3+A1')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
    )
    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' })
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' }, [[null]])
    engine.setCellFormula('Sheet1', 'C3', 'A1+A1')

    let undoCount = 0
    while (engine.undo()) {
      undoCount += 1
      expect(undoCount).toBeLessThanOrEqual(16)
    }
    expect(undoCount).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('applies cell mutations by sheet id and returns inverse ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-mutation-refs' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    const undoOps = engine.applyCellMutationsAt([
      {
        sheetId,
        mutation: {
          kind: 'setCellValue',
          row: 0,
          col: 0,
          value: 10,
        },
      },
      {
        sheetId,
        mutation: {
          kind: 'setCellFormula',
          row: 0,
          col: 1,
          formula: 'A1*2',
        },
      },
      {
        sheetId,
        mutation: {
          kind: 'clearCell',
          row: 3,
          col: 3,
        },
      },
      {
        sheetId,
        mutation: {
          kind: 'setCellFormula',
          row: 0,
          col: 2,
          formula: 'SUM(',
        },
      },
    ])

    expect(undoOps).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ kind: 'clearCell', address: 'A1' }),
        expect.objectContaining({ kind: 'clearCell', address: 'B1' }),
        expect.objectContaining({ kind: 'clearCell', address: 'D4' }),
      ]),
    )
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCell('Sheet1', 'C1').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getLastMetrics().compileMs).toBeGreaterThanOrEqual(0)
  })

  it('supports direct cell mutations by coordinates', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-cell-mutations' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    expect(engine.setCellValueAt(sheetId, 1, 1, 5)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(engine.setCellFormulaAt(sheetId, 1, 2, 'B2*3')).toEqual({
      tag: ValueTag.Number,
      value: 15,
    })

    engine.clearCellAt(sheetId, 1, 1)

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(() => engine.setCellValueAt(999, 0, 0, 1)).toThrow('Unknown sheet id: 999')
  })

  it('moves overlapping ranges without losing cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'first')
    engine.setCellValue('Sheet1', 'B1', 'second')

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C1' },
    )

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'first',
      }),
    )
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'second',
      }),
    )
  })

  it('relocates formulas when filling down', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')

    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B3' },
    )

    expect(engine.getCell('Sheet1', 'B2').formula).toBe('A2*3')
    expect(engine.getCell('Sheet1', 'B3').formula).toBe('A3*3')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 12 })
  })

  it('validates bulk range helpers and no-ops empty fills', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    expect(() => engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [[1]])).toThrow(
      'setRangeValues requires a value matrix that exactly matches the target range',
    )
    expect(() => engine.setRangeFormulas({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [['A1']])).toThrow(
      'setRangeFormulas requires a formula matrix that exactly matches the target range',
    )

    engine.fillRange(
      { sheetName: 'Missing', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D2' },
    )
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Empty })

    expect(() =>
      engine.copyRange(
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C1' },
      ),
    ).toThrow('copyRange requires source and target dimensions to match exactly')
  })

  it('stores invalid formulas as #VALUE errors instead of throwing', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    expect(() => engine.setCellFormula('Sheet1', 'A1', '1+')).not.toThrow()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('spills sequence formulas through the runtime and recalculates downstream refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'A3*2')
    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 0 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'A1', rows: 3, cols: 1 }])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(restored.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(restored.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(restored.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'A1', rows: 3, cols: 1 }])
  })

  it('clears prior sequence spills when the owner becomes a scalar', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'A3*2')
    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')

    engine.setCellValue('Sheet1', 'A1', 7)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toBeUndefined()
  })

  it('blocks sequence spills when target cells are occupied', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A2', 99)
    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 99 })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toBeUndefined()
  })

  it('evaluates nested sequence aggregates on the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'C1', 'AVG(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'D1', 'MIN(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'E1', 'MAX(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'F1', 'COUNT(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'G1', 'COUNTA(SEQUENCE(A1,1,1,1))')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('evaluates unsupported wasm formulas through the JS runtime fallback', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 4)
    engine.setCellValue('Sheet1', 'A2', 6)
    engine.setCellFormula('Sheet1', 'B1', 'LEN(A1:A2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.JsOnly)
  })

  it('evaluates LET through the wasm fast path after rewrite-based lowering', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'LET(x,2,x+3)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates reference metadata functions through the JS runtime fallback', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Summary')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 4)
    engine.setCellValue('Sheet1', 'A2', 'text')
    engine.setCellFormula('Sheet2', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'C1', 'ROW()')
    engine.setCellFormula('Sheet1', 'D1', 'COLUMN()')
    engine.setCellFormula('Sheet1', 'E1', 'FORMULATEXT(Sheet2!B1)')
    engine.setCellFormula('Sheet1', 'F1', 'SHEET()')
    engine.setCellFormula('Sheet1', 'G1', 'SHEETS()')
    engine.setCellFormula('Sheet1', 'H1', 'CELL("address",B3)')
    engine.setCellFormula('Sheet1', 'I1', 'CELL("contents",A1)')
    engine.setCellFormula('Sheet1', 'J1', 'CELL("type",A2)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'E1')).toMatchObject({
      tag: ValueTag.String,
      value: '=A1*2',
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.String,
      value: '$B$3',
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.String,
      value: 'l',
    })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
  })

  it('routes TEXTSPLIT, EXPAND, and TRIMRANGE through wasm while keeping indirection helpers on JS', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 'red,blue|green')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'L2', 1)
    engine.setCellValue('Sheet1', 'M2', 2)
    engine.setCellValue('Sheet1', 'L3', 3)
    engine.setCellFormula('Sheet2', 'A1', 'B1*2')
    engine.setCellFormula('Sheet1', 'C1', 'TEXTSPLIT(A1,",","|")')
    engine.setCellFormula('Sheet1', 'E1', 'EXPAND(B1:B2,3,2,0)')
    engine.setCellFormula('Sheet1', 'G1', 'INDIRECT("B1:B2")')
    engine.setCellFormula('Sheet1', 'H1', 'INDIRECT("B2")')
    engine.setCellFormula('Sheet1', 'I1', 'FORMULA(Sheet2!A1)')
    engine.setCellFormula('Sheet1', 'K6', 'TRIMRANGE(K1:N4)')

    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.String,
      value: 'red',
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.String,
      value: 'blue',
    })
    expect(engine.getCellValue('Sheet1', 'C2')).toMatchObject({
      tag: ValueTag.String,
      value: 'green',
    })
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'E3')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.String,
      value: '=B1*2',
    })
    expect(engine.getCellValue('Sheet1', 'K6')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'L6')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'K7')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'L7')).toEqual({ tag: ValueTag.Empty })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'K6').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes DATEDIF and financial scalar helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'DATEDIF(DATE(2020,1,15),DATE(2021,3,20),"YM")')
    engine.setCellFormula('Sheet1', 'B1', 'DAYS360(DATE(2024,1,29),DATE(2024,3,31))')
    engine.setCellFormula('Sheet1', 'C1', 'DAYS360(DATE(2024,1,29),DATE(2024,3,31),TRUE)')
    engine.setCellFormula('Sheet1', 'D1', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)')
    engine.setCellFormula('Sheet1', 'E1', 'FVSCHEDULE(1000,0.09,0.11,0.1)')
    engine.setCellFormula('Sheet1', 'F1', 'DB(10000,1000,5,1)')
    engine.setCellFormula('Sheet1', 'G1', 'DDB(2400,300,10,2)')
    engine.setCellFormula('Sheet1', 'H1', 'VDB(2400,300,10,1,3)')
    engine.setCellFormula('Sheet1', 'I1', 'SLN(10000,1000,9)')
    engine.setCellFormula('Sheet1', 'J1', 'SYD(10000,1000,9,1)')
    engine.setCellFormula('Sheet1', 'K1', 'DISC(DATE(2023,1,1),DATE(2023,4,1),97,100,2)')
    engine.setCellFormula('Sheet1', 'L1', 'INTRATE(DATE(2023,1,1),DATE(2023,4,1),1000,1030,2)')
    engine.setCellFormula('Sheet1', 'M1', 'RECEIVED(DATE(2023,1,1),DATE(2023,4,1),1000,0.12,2)')
    engine.setCellFormula('Sheet1', 'N1', 'PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)')
    engine.setCellFormula('Sheet1', 'O1', 'YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)')
    engine.setCellFormula('Sheet1', 'P1', 'TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)')
    engine.setCellFormula('Sheet1', 'Q1', 'TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)')
    engine.setCellFormula('Sheet1', 'R1', 'TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)')
    engine.setCellFormula('Sheet1', 'S1', 'PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)')
    engine.setCellFormula('Sheet1', 'T1', 'YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)')
    engine.setCellFormula(
      'Sheet1',
      'U1',
      'ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)',
    )
    engine.setCellFormula('Sheet1', 'V1', 'ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)')
    engine.setCellFormula('Sheet1', 'W1', 'ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)')
    engine.setCellFormula('Sheet1', 'X1', 'ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)')
    engine.setCellFormula('Sheet1', 'Y1', 'EFFECT(0.12,12)')
    engine.setCellFormula('Sheet1', 'Z1', 'NOMINAL(0.12682503013196977,12)')
    engine.setCellFormula('Sheet1', 'AA1', 'PDURATION(0.1,100,121)')
    engine.setCellFormula('Sheet1', 'AB1', 'RRI(2,100,121)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.Number,
      value: 62,
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Number,
      value: 61,
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(182 / 365, 12),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1330.89, 12),
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3690, 12),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(384, 12),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(691.2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 1000 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 1800 })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(engine.getCellValue('Sheet1', 'L1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(engine.getCellValue('Sheet1', 'M1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1030.9278350515465, 12),
    })
    expect(engine.getCellValue('Sheet1', 'N1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.79583333333333, 12),
    })
    expect(engine.getCellValue('Sheet1', 'O1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.05282257198685834, 12),
    })
    expect(engine.getCellValue('Sheet1', 'P1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(98.45, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Q1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09141696292534264, 12),
    })
    expect(engine.getCellValue('Sheet1', 'R1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09415149356594302, 12),
    })
    expect(engine.getCellValue('Sheet1', 'S1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.98449887555694, 12),
    })
    expect(engine.getCellValue('Sheet1', 'T1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.060954333691538576, 12),
    })
    expect(engine.getCellValue('Sheet1', 'U1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(113.597717474079, 12),
    })
    expect(engine.getCellValue('Sheet1', 'V1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0772455415972989, 11),
    })
    expect(engine.getCellValue('Sheet1', 'W1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.8782860147213, 12),
    })
    expect(engine.getCellValue('Sheet1', 'X1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0451922356291692, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Y1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12682503013196977, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Z1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(engine.getCellValue('Sheet1', 'AA1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'AB1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1, 12),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'N1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'O1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'P1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'Q1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'R1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'S1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'T1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'U1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'V1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'W1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'X1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'Y1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'Z1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'AA1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'AB1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes cash-flow rate helpers through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    ;[-70000, 12000, 15000, 18000, 21000, 26000].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    ;[-120000, 39000, 30000, 21000, 37000, 46000].forEach((value, index) => {
      engine.setCellValue('Sheet1', `C${index + 1}`, value)
    })
    ;[-10000, 2750, 4250, 3250, 2750].forEach((value, index) => {
      engine.setCellValue('Sheet1', `E${index + 1}`, value)
    })
    ;[39448, 39508, 39751, 39859, 39904].forEach((value, index) => {
      engine.setCellValue('Sheet1', `F${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'H1', 'IRR(A1:A6)')
    engine.setCellFormula('Sheet1', 'I1', 'MIRR(C1:C6,10%,12%)')
    engine.setCellFormula('Sheet1', 'J1', 'XNPV(0.09,E1:E5,F1:F5)')
    engine.setCellFormula('Sheet1', 'K1', 'XIRR(E1:E5,F1:F5)')

    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.08663094803653162, 12),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1260941303659051, 12),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2086.647602031535, 9),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.37336253351883136, 12),
    })

    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes covariance and regression helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellValue('Sheet1', 'A2', 8)
    engine.setCellValue('Sheet1', 'A3', 11)
    engine.setCellValue('Sheet1', 'B1', 1)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 3)

    engine.setCellFormula('Sheet1', 'C1', 'CORREL(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'D1', 'COVAR(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'E1', 'COVARIANCE.P(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'F1', 'COVARIANCE.S(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'G1', 'PEARSON(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'H1', 'INTERCEPT(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'I1', 'SLOPE(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'J1', 'RSQ(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'K1', 'STEYX(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'L1', 'FORECAST(4,A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'M1', 'FORECAST.LINEAR(4,A1:A3,B1:B3)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 14 })

    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills TREND and GROWTH through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellValue('Sheet1', 'A2', 8)
    engine.setCellValue('Sheet1', 'A3', 11)
    engine.setCellValue('Sheet1', 'B1', 1)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 3)
    engine.setCellValue('Sheet1', 'D1', 4)
    engine.setCellValue('Sheet1', 'D2', 5)
    engine.setCellFormula('Sheet1', 'F1', 'TREND(A1:A3,B1:B3,D1:D2)')

    engine.setCellValue('Sheet1', 'H1', 2)
    engine.setCellValue('Sheet1', 'H2', 4)
    engine.setCellValue('Sheet1', 'H3', 8)
    engine.setCellValue('Sheet1', 'I1', 1)
    engine.setCellValue('Sheet1', 'I2', 2)
    engine.setCellValue('Sheet1', 'I3', 3)
    engine.setCellValue('Sheet1', 'K1', 4)
    engine.setCellValue('Sheet1', 'K2', 5)
    engine.setCellFormula('Sheet1', 'M1', 'GROWTH(H1:H3,I1:I3,K1:K2)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 17 })
    expect(engine.getCellValue('Sheet1', 'M1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(16, 12),
    })
    expect(engine.getCellValue('Sheet1', 'M2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(32, 12),
    })

    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills LINEST and LOGEST through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellValue('Sheet1', 'A2', 8)
    engine.setCellValue('Sheet1', 'A3', 11)
    engine.setCellValue('Sheet1', 'B1', 1)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 3)
    engine.setCellFormula('Sheet1', 'D1', 'LINEST(A1:A3,B1:B3)')

    engine.setCellValue('Sheet1', 'G1', 2)
    engine.setCellValue('Sheet1', 'G2', 4)
    engine.setCellValue('Sheet1', 'G3', 8)
    engine.setCellValue('Sheet1', 'H1', 1)
    engine.setCellValue('Sheet1', 'H2', 2)
    engine.setCellValue('Sheet1', 'H3', 3)
    engine.setCellFormula('Sheet1', 'J1', 'LOGEST(G1:G3,H1:H3)')

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })

    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes promoted scalar and reducer math helpers through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'A3', 4)
    engine.setCellFormula('Sheet1', 'B1', 'ACOSH(1)')
    engine.setCellFormula('Sheet1', 'C1', 'COT(1)')
    engine.setCellFormula('Sheet1', 'D1', 'SECH(0)')
    engine.setCellFormula('Sheet1', 'E1', 'EVEN(-3)')
    engine.setCellFormula('Sheet1', 'F1', 'ODD(-2)')
    engine.setCellFormula('Sheet1', 'G1', 'FACT(5)')
    engine.setCellFormula('Sheet1', 'H1', 'FACTDOUBLE(6)')
    engine.setCellFormula('Sheet1', 'I1', 'COMBIN(8,3)')
    engine.setCellFormula('Sheet1', 'J1', 'COMBINA(3,2)')
    engine.setCellFormula('Sheet1', 'K1', 'GCD(A1:A3)')
    engine.setCellFormula('Sheet1', 'L1', 'LCM(A1:A3)')
    engine.setCellFormula('Sheet1', 'M1', 'PRODUCT(A1:A3)')
    engine.setCellFormula('Sheet1', 'N1', 'QUOTIENT(7,3)')
    engine.setCellFormula('Sheet1', 'O1', 'GEOMEAN(A1:A3)')
    engine.setCellFormula('Sheet1', 'P1', 'HARMEAN(A1:A3)')
    engine.setCellFormula('Sheet1', 'Q1', 'SUMSQ(A1:A3)')
    engine.setCellFormula('Sheet1', 'R1', 'TRUNC(-3.98,1)')
    engine.setCellFormula('Sheet1', 'S1', 'FLOOR.MATH(-5.5,2)')
    engine.setCellFormula('Sheet1', 'T1', 'FLOOR.PRECISE(-5.5,2)')
    engine.setCellFormula('Sheet1', 'U1', 'CEILING.MATH(-5.5,2)')
    engine.setCellFormula('Sheet1', 'V1', 'CEILING.PRECISE(-5.5,2)')
    engine.setCellFormula('Sheet1', 'W1', 'ISO.CEILING(-5.5,2)')
    engine.setCellFormula('Sheet1', 'X1', 'MROUND(10,4)')
    engine.setCellFormula('Sheet1', 'Y1', 'SQRTPI(2)')
    engine.setCellFormula('Sheet1', 'Z1', 'PERMUT(5,3)')
    engine.setCellFormula('Sheet1', 'AA1', 'PERMUTATIONA(2,3)')
    engine.setCellFormula('Sheet1', 'AB1', 'SERIESSUM(2,1,2,1,2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6420926159343306, 12),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: -3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 120 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 48 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 56 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'O1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2.8844991406148166, 12),
    })
    expect(engine.getCellValue('Sheet1', 'P1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2.769230769230769, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Q1')).toEqual({ tag: ValueTag.Number, value: 29 })
    expect(engine.getCellValue('Sheet1', 'R1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.9, 12),
    })
    expect(engine.getCellValue('Sheet1', 'S1')).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(engine.getCellValue('Sheet1', 'T1')).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(engine.getCellValue('Sheet1', 'U1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'V1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'W1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'X1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'Y1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2.5066282746310002, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Z1')).toEqual({ tag: ValueTag.Number, value: 60 })
    expect(engine.getCellValue('Sheet1', 'AA1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'AB1')).toEqual({ tag: ValueTag.Number, value: 18 })

    for (const address of [
      'B1',
      'C1',
      'D1',
      'E1',
      'F1',
      'G1',
      'H1',
      'I1',
      'J1',
      'K1',
      'L1',
      'M1',
      'N1',
      'O1',
      'P1',
      'Q1',
      'R1',
      'S1',
      'T1',
      'U1',
      'V1',
      'W1',
      'X1',
      'Y1',
      'Z1',
      'AA1',
      'AB1',
    ]) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
  })

  it('routes coupon-date and bond pricing helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'COUPDAYBS(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'B1', 'COUPDAYS(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'C1', 'COUPDAYSNC(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'D1', 'COUPNCD(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'E1', 'COUPNUM(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'F1', 'COUPPCD(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'G1', 'PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)')
    engine.setCellFormula('Sheet1', 'H1', 'YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)')
    engine.setCellFormula('Sheet1', 'I1', 'DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)')
    engine.setCellFormula('Sheet1', 'J1', 'MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 180 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 110 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 39217 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 39036 })
    expect(engine.getCellValue('Sheet1', 'G1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(94.63436162132213, 12),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.065, 7),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(10.919145281591925, 12),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5.735669813918838, 12),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes annuity and cumulative loan helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'PV(0.1,2,-576.1904761904761)')
    engine.setCellFormula('Sheet1', 'B1', 'PMT(0.1,2,1000)')
    engine.setCellFormula('Sheet1', 'C1', 'NPER(0.1,-576.1904761904761,1000)')
    engine.setCellFormula('Sheet1', 'D1', 'RATE(48,-200,8000)')
    engine.setCellFormula('Sheet1', 'E1', 'IPMT(0.1,1,2,1000)')
    engine.setCellFormula('Sheet1', 'F1', 'PPMT(0.1,1,2,1000)')
    engine.setCellFormula('Sheet1', 'G1', 'ISPMT(0.1,1,2,1000)')
    engine.setCellFormula('Sheet1', 'H1', 'CUMIPMT(9%/12,30*12,125000,13,24,0)')
    engine.setCellFormula('Sheet1', 'I1', 'CUMPRINC(9%/12,30*12,125000,13,24,0)')
    engine.setCellFormula('Sheet1', 'J1', 'FV(0.1,2,-100,-1000)')
    engine.setCellFormula('Sheet1', 'K1', 'NPV(0.1,100,200,300)')

    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1000.0000000000006, 12),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-576.1904761904758, 12),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.9999999999999982, 12),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.007701472488246008, 12),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: -100 })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-476.1904761904758, 12),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: -50 })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-11135.232130750845, 12),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-934.1071234208765, 12),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1420, 12),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(481.5927873779113, 12),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes rank helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 20)
    engine.setCellValue('Sheet1', 'A4', 30)

    engine.setCellFormula('Sheet1', 'B1', 'RANK(20,A1:A4)')
    engine.setCellFormula('Sheet1', 'C1', 'RANK.EQ(20,A1:A4)')
    engine.setCellFormula('Sheet1', 'D1', 'RANK.AVG(20,A1:A4)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2.5 })

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes order-statistics helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 4, 7, 8, 9, 10, 12].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'B1', 'MEDIAN(A1:A8)')
    engine.setCellFormula('Sheet1', 'C1', 'SMALL(A1:A8,3)')
    engine.setCellFormula('Sheet1', 'D1', 'LARGE(A1:A8,2)')
    engine.setCellFormula('Sheet1', 'E1', 'PERCENTILE(A1:A8,0.25)')
    engine.setCellFormula('Sheet1', 'F1', 'PERCENTILE.INC(A1:A8,0.25)')
    engine.setCellFormula('Sheet1', 'G1', 'PERCENTILE.EXC(A1:A8,0.25)')
    engine.setCellFormula('Sheet1', 'H1', 'QUARTILE(A1:A8,1)')
    engine.setCellFormula('Sheet1', 'I1', 'QUARTILE.INC(A1:A8,1)')
    engine.setCellFormula('Sheet1', 'J1', 'QUARTILE.EXC(A1:A8,1)')
    engine.setCellFormula('Sheet1', 'K1', 'PERCENTRANK(A1:A8,8)')
    engine.setCellFormula('Sheet1', 'L1', 'PERCENTRANK.INC(A1:A8,8)')
    engine.setCellFormula('Sheet1', 'M1', 'PERCENTRANK.EXC(A1:A8,8)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 7.5 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 2.5 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 2.5 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 0.571 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 0.571 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 0.555 })

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills MODE.MULT and FREQUENCY through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 2, 3, 3, 4].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    ;[79, 85, 78, 85, 50, 81].forEach((value, index) => {
      engine.setCellValue('Sheet1', `C${index + 1}`, value)
    })
    ;[60, 80, 90].forEach((value, index) => {
      engine.setCellValue('Sheet1', `D${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'F1', 'MODE.MULT(A1:A6)')
    engine.setCellFormula('Sheet1', 'G1', 'FREQUENCY(C1:C6,D1:D3)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'G3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G4')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates database aggregation formulas through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[
      ['Age', 'Height', 'Yield'],
      [10, 100, 5],
      [12, 110, 7],
      [12, 120, 9],
      [15, 130, 11],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`, value)
      })
    })
    engine.setCellValue('Sheet1', 'E1', 'Age')
    engine.setCellValue('Sheet1', 'E2', 12)
    engine.setCellValue('Sheet1', 'F1', 'Age')
    engine.setCellValue('Sheet1', 'F2', 15)

    engine.setCellFormula('Sheet1', 'H1', 'DAVERAGE(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H2', 'DCOUNT(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H3', 'DCOUNTA(A1:C5,"Height",E1:E2)')
    engine.setCellFormula('Sheet1', 'H4', 'DGET(A1:C5,"Height",F1:F2)')
    engine.setCellFormula('Sheet1', 'H5', 'DPRODUCT(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H6', 'DSTDEV(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H7', 'DVARP(A1:C5,"Yield",E1:E2)')

    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'H3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'H4')).toEqual({ tag: ValueTag.Number, value: 130 })
    expect(engine.getCellValue('Sheet1', 'H5')).toEqual({ tag: ValueTag.Number, value: 63 })
    expect(engine.getCellValue('Sheet1', 'H6')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.SQRT2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'H7')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H4').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H6').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes PROB and TRIMMEAN through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 3, 4].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    ;[0.1, 0.2, 0.3, 0.4].forEach((value, index) => {
      engine.setCellValue('Sheet1', `B${index + 1}`, value)
    })
    ;[1, 2, 4, 7, 8, 9, 10, 12].forEach((value, index) => {
      engine.setCellValue('Sheet1', `D${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'F1', 'PROB(A1:A4,B1:B4,2,3)')
    engine.setCellFormula('Sheet1', 'G1', 'TRIMMEAN(D1:D8,0.25)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 0.5 })
    expect(engine.getCellValue('Sheet1', 'G1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(40 / 6, 12),
    })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes COUNTBLANK, ISOWEEKNUM, and TIMEVALUE through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 8)
    engine.setCellValue('Sheet1', 'A2', null)
    engine.setCellValue('Sheet1', 'B1', 'x')
    engine.setCellValue('Sheet1', 'B2', null)
    engine.setCellFormula('Sheet1', 'D1', 'COUNTBLANK(A1:B2)')
    engine.setCellFormula('Sheet1', 'E1', 'ISOWEEKNUM(DATE(2024,1,1))')
    engine.setCellFormula('Sheet1', 'F1', 'TIMEVALUE("1:30 PM")')

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5625, 12),
    })

    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes chi-square inverse functions and compatibility aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'CHIDIST(18.307,10)')
    engine.setCellFormula('Sheet1', 'B1', 'LEGACY.CHIDIST(18.307,10)')
    engine.setCellFormula('Sheet1', 'C1', 'CHISQDIST(18.307,10)')
    engine.setCellFormula('Sheet1', 'D1', 'CHIINV(0.050001,10)')
    engine.setCellFormula('Sheet1', 'E1', 'CHISQ.INV.RT(0.050001,10)')
    engine.setCellFormula('Sheet1', 'F1', 'CHISQINV(0.050001,10)')
    engine.setCellFormula('Sheet1', 'G1', 'LEGACY.CHIINV(0.050001,10)')
    engine.setCellFormula('Sheet1', 'H1', 'CHISQ.INV(0.93,1)')

    const a1 = engine.getCellValue('Sheet1', 'A1')
    expect(a1).toMatchObject({ tag: ValueTag.Number })
    expect(a1.value).toBeCloseTo(0.05000058909139826, 12)

    const b1 = engine.getCellValue('Sheet1', 'B1')
    expect(b1).toMatchObject({ tag: ValueTag.Number })
    expect(b1.value).toBeCloseTo(0.05000058909139826, 12)

    const c1 = engine.getCellValue('Sheet1', 'C1')
    expect(c1).toMatchObject({ tag: ValueTag.Number })
    expect(c1.value).toBeCloseTo(0.05000058909139826, 12)

    const d1 = engine.getCellValue('Sheet1', 'D1')
    expect(d1).toMatchObject({ tag: ValueTag.Number })
    expect(d1.value).toBeCloseTo(18.30697345696106, 12)

    const e1 = engine.getCellValue('Sheet1', 'E1')
    expect(e1).toMatchObject({ tag: ValueTag.Number })
    expect(e1.value).toBeCloseTo(18.30697345696106, 12)

    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(18.30697345696106, 12)

    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(18.30697345696106, 12)

    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(3.2830202867594993, 12)

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes chi-square test functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [
      [58, 35],
      [11, 25],
      [10, 23],
    ])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E3' }, [
      [45.35, 47.65],
      [17.56, 18.44],
      [16.09, 16.91],
    ])
    engine.setCellFormula('Sheet1', 'G1', 'CHISQ.TEST(A1:B3,D1:E3)')
    engine.setCellFormula('Sheet1', 'H1', 'CHITEST(A1:B3,D1:E3)')
    engine.setCellFormula('Sheet1', 'I1', 'LEGACY.CHITEST(A1:B3,D1:E3)')

    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.0003082, 7)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(0.0003082, 7)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(0.0003082, 7)

    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes beta and f distribution functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'BETA.DIST(2,8,10,TRUE,1,3)')
    engine.setCellFormula('Sheet1', 'B1', 'BETADIST(2,8,10,1,3)')
    engine.setCellFormula('Sheet1', 'C1', 'BETA.INV(0.6854705810117458,8,10,1,3)')
    engine.setCellFormula('Sheet1', 'D1', 'BETAINV(0.6854705810117458,8,10,1,3)')
    engine.setCellFormula('Sheet1', 'E1', 'F.DIST(15.2068649,6,4,TRUE)')
    engine.setCellFormula('Sheet1', 'F1', 'F.DIST.RT(15.2068649,6,4)')
    engine.setCellFormula('Sheet1', 'G1', 'FDIST(15.2068649,6,4)')
    engine.setCellFormula('Sheet1', 'H1', 'LEGACY.FDIST(15.2068649,6,4)')
    engine.setCellFormula('Sheet1', 'I1', 'F.INV(0.01,6,4)')
    engine.setCellFormula('Sheet1', 'J1', 'F.INV.RT(0.01,6,4)')
    engine.setCellFormula('Sheet1', 'K1', 'FINV(0.01,6,4)')
    engine.setCellFormula('Sheet1', 'L1', 'LEGACY.FINV(0.01,6,4)')

    const a1 = engine.getCellValue('Sheet1', 'A1')
    expect(a1).toMatchObject({ tag: ValueTag.Number })
    expect(a1.value).toBeCloseTo(0.6854705810117458, 10)
    const b1 = engine.getCellValue('Sheet1', 'B1')
    expect(b1).toMatchObject({ tag: ValueTag.Number })
    expect(b1.value).toBeCloseTo(0.6854705810117458, 10)
    const c1 = engine.getCellValue('Sheet1', 'C1')
    expect(c1).toMatchObject({ tag: ValueTag.Number })
    expect(c1.value).toBeCloseTo(2, 10)
    const d1 = engine.getCellValue('Sheet1', 'D1')
    expect(d1).toMatchObject({ tag: ValueTag.Number })
    expect(d1.value).toBeCloseTo(2, 10)
    const e1 = engine.getCellValue('Sheet1', 'E1')
    expect(e1).toMatchObject({ tag: ValueTag.Number })
    expect(e1.value).toBeCloseTo(0.99, 9)
    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(0.01, 9)
    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.01, 9)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(0.01, 9)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(0.10930991466299911, 8)
    const j1 = engine.getCellValue('Sheet1', 'J1')
    expect(j1).toMatchObject({ tag: ValueTag.Number })
    expect(j1.value).toBeCloseTo(15.206864870947697, 7)
    const k1 = engine.getCellValue('Sheet1', 'K1')
    expect(k1).toMatchObject({ tag: ValueTag.Number })
    expect(k1.value).toBeCloseTo(15.206864870947697, 7)
    const l1 = engine.getCellValue('Sheet1', 'L1')
    expect(l1).toMatchObject({ tag: ValueTag.Number })
    expect(l1.value).toBeCloseTo(15.206864870947697, 7)

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes student-t distribution functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[1], [2], [4]])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B3' }, [[1], [3], [3]])
    engine.setCellFormula('Sheet1', 'D1', 'T.DIST(1,1,TRUE)')
    engine.setCellFormula('Sheet1', 'E1', 'T.DIST.RT(1,1)')
    engine.setCellFormula('Sheet1', 'F1', 'T.DIST.2T(1,1)')
    engine.setCellFormula('Sheet1', 'G1', 'TDIST(1,1,2)')
    engine.setCellFormula('Sheet1', 'H1', 'T.INV(0.75,1)')
    engine.setCellFormula('Sheet1', 'I1', 'T.INV.2T(0.5,1)')
    engine.setCellFormula('Sheet1', 'J1', 'TINV(0.5,1)')
    engine.setCellFormula('Sheet1', 'K1', 'CONFIDENCE.T(0.5,2,4)')
    engine.setCellFormula('Sheet1', 'L1', 'T.TEST(A1:A3,B1:B3,2,1)')
    engine.setCellFormula('Sheet1', 'M1', 'TTEST(A1:A3,B1:B3,2,1)')

    const d1 = engine.getCellValue('Sheet1', 'D1')
    expect(d1).toMatchObject({ tag: ValueTag.Number })
    expect(d1.value).toBeCloseTo(0.75, 12)
    const e1 = engine.getCellValue('Sheet1', 'E1')
    expect(e1).toMatchObject({ tag: ValueTag.Number })
    expect(e1.value).toBeCloseTo(0.25, 12)
    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(0.5, 12)
    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.5, 12)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(1, 9)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(1, 9)
    const j1 = engine.getCellValue('Sheet1', 'J1')
    expect(j1).toMatchObject({ tag: ValueTag.Number })
    expect(j1.value).toBeCloseTo(1, 9)
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.764892328404345, 12),
    })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 1 })

    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes statistical scalar and dispersion builtins through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 3, 4, 5].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'C1', 'STANDARDIZE(1,0,1)')
    engine.setCellFormula('Sheet1', 'C2', 'STDEV(A1:A4)')
    engine.setCellFormula('Sheet1', 'C3', 'STDEVA(2,TRUE(),"skip")')
    engine.setCellFormula('Sheet1', 'C4', 'VAR(A1:A4)')
    engine.setCellFormula('Sheet1', 'C5', 'VARA(2,TRUE(),"skip")')
    engine.setCellFormula('Sheet1', 'C6', 'MODE(1,2,2,3)')
    engine.setCellFormula('Sheet1', 'C7', 'MODE.SNGL(1,2,2,3)')
    engine.setCellFormula('Sheet1', 'D1', 'SKEW(A1:A5)')
    engine.setCellFormula('Sheet1', 'D2', 'KURT(A1:A5)')
    engine.setCellFormula('Sheet1', 'D3', 'NORMDIST(1,0,1,TRUE)')
    engine.setCellFormula('Sheet1', 'D4', 'NORMINV(0.8413447460685429,0,1)')
    engine.setCellFormula('Sheet1', 'D5', 'NORMSDIST(1)')
    engine.setCellFormula('Sheet1', 'E1', 'CONFIDENCE.NORM(0.05,1,100)')
    engine.setCellFormula('Sheet1', 'E2', 'NORMSINV(0.001)')
    engine.setCellFormula('Sheet1', 'E3', 'LOGINV(0.5,0,1)')
    engine.setCellFormula('Sheet1', 'E4', 'LOGNORMDIST(1,0,1)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(5 / 3), 12),
    })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C4')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5 / 3, 12),
    })
    expect(engine.getCellValue('Sheet1', 'C5')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C7')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'D2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1.2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'D3')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(engine.getCellValue('Sheet1', 'D4')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    })
    expect(engine.getCellValue('Sheet1', 'D5')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1959963986120195, 12),
    })
    expect(engine.getCellValue('Sheet1', 'E2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.090232306167813, 8),
    })
    expect(engine.getCellValue('Sheet1', 'E3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E4')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 8),
    })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C4').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C5').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C6').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C7').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D4').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D5').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E4').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes gamma inverse functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'GAMMA.INV(0.08030139707139418,3,2)')
    engine.setCellFormula('Sheet1', 'B1', 'GAMMAINV(0.08030139707139418,3,2)')

    const a1 = engine.getCellValue('Sheet1', 'A1')
    expect(a1).toMatchObject({ tag: ValueTag.Number })
    expect(a1.value).toBeCloseTo(2, 10)
    const b1 = engine.getCellValue('Sheet1', 'B1')
    expect(b1).toMatchObject({ tag: ValueTag.Number })
    expect(b1.value).toBeCloseTo(2, 10)

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes f-test and z-test functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A5' }, [[6], [7], [9], [15], [21]])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B5' }, [[20], [28], [31], [38], [40]])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D5' }, [[1], [2], [3], [4], [5]])
    engine.setCellFormula('Sheet1', 'F1', 'F.TEST(A1:A5,B1:B5)')
    engine.setCellFormula('Sheet1', 'G1', 'FTEST(A1:A5,B1:B5)')
    engine.setCellFormula('Sheet1', 'H1', 'Z.TEST(D1:D5,2,1)')
    engine.setCellFormula('Sheet1', 'I1', 'ZTEST(D1:D5,2,1)')

    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(0.648317846786175, 12)
    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.648317846786175, 12)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(0.012673617875446075, 12)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(0.012673617875446075, 12)

    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes WORKDAY.INTL and NETWORKDAYS.INTL through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 46094)
    engine.setCellValue('Sheet1', 'A2', 46098)
    engine.setCellValue('Sheet1', 'B1', 46096)
    engine.setCellFormula('Sheet1', 'D1', 'WORKDAY.INTL(A1,1,7)')
    engine.setCellFormula('Sheet1', 'E1', 'WORKDAY.INTL(A1,2,7,B1)')
    engine.setCellFormula('Sheet1', 'F1', 'NETWORKDAYS.INTL(A1,A2,7)')
    engine.setCellFormula('Sheet1', 'G1', 'NETWORKDAYS.INTL(A1,A2,7,B1)')

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 46097 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 46099 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes NUMBERVALUE and VALUETOTEXT through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'NUMBERVALUE("2.500,27",",",".")')
    engine.setCellFormula('Sheet1', 'B1', 'VALUETOTEXT("alpha",1)')
    engine.setCellFormula('Sheet1', 'C1', 'TEXT(1234.567,"#,##0.00")')
    engine.setCellFormula('Sheet1', 'D1', 'TEXT(DATE(2024,3,5),"yyyy-mm-dd")')

    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({ tag: ValueTag.Number })
    expect(engine.getCellValue('Sheet1', 'A1').value).toBeCloseTo(2500.27, 12)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.String,
      value: '"alpha"',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.String,
      value: '1,234.57',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({
      tag: ValueTag.String,
      value: '2024-03-05',
      stringId: expect.any(Number),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes scalar text conversion helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'CHAR(65)')
    engine.setCellFormula('Sheet1', 'B1', 'CODE("A")')
    engine.setCellFormula('Sheet1', 'C1', 'UNICODE("A")')
    engine.setCellFormula('Sheet1', 'D1', 'UNICHAR(66)')
    engine.setCellFormula('Sheet1', 'E1', 'CLEAN(CHAR(97)&CHAR(1)&CHAR(98))')
    engine.setCellFormula('Sheet1', 'F1', 'ASC("ＡＢＣ　１２３")')
    engine.setCellFormula('Sheet1', 'G1', 'JIS("ABC 123")')
    engine.setCellFormula('Sheet1', 'H1', 'DBCS("ｶﾞｷﾞｸﾞｹﾞｺﾞ")')
    engine.setCellFormula('Sheet1', 'I1', 'BAHTTEXT(1234)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: 'A',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 65 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 65 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({
      tag: ValueTag.String,
      value: 'B',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.String,
      value: 'ab',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.String,
      value: 'ABC 123',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({
      tag: ValueTag.String,
      value: 'ＡＢＣ　１２３',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({
      tag: ValueTag.String,
      value: 'ガギグゲゴ',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({
      tag: ValueTag.String,
      value: 'หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน',
      stringId: expect.any(Number),
    })

    for (const address of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes PHONETIC through the wasm path and reads the top-left range member', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'カタカナ')
    engine.setCellFormula('Sheet1', 'B1', '1/0')
    engine.setCellFormula('Sheet1', 'C1', 'PHONETIC(A1:B1)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.String,
      value: 'カタカナ',
      stringId: expect.any(Number),
    })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes CHOOSE, TEXTBEFORE, TEXTAFTER, and TEXTJOIN through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'C1', 20)
    engine.setCellValue('Sheet1', 'B2', 30)
    engine.setCellValue('Sheet1', 'C2', 40)
    engine.setCellValue('Sheet1', 'D1', 100)
    engine.setCellValue('Sheet1', 'E1', 200)
    engine.setCellValue('Sheet1', 'D2', 300)
    engine.setCellValue('Sheet1', 'E2', 400)
    engine.setCellValue('Sheet1', 'G1', 'alpha')
    engine.setCellValue('Sheet1', 'G2', null)
    engine.setCellValue('Sheet1', 'G3', 'beta')
    engine.setCellFormula('Sheet1', 'A1', 'CHOOSE(2,"red","blue","green")')
    engine.setCellFormula('Sheet1', 'A2', 'TEXTBEFORE("alpha-beta","-")')
    engine.setCellFormula('Sheet1', 'A3', 'TEXTAFTER("alpha-beta","-")')
    engine.setCellFormula('Sheet1', 'A4', 'TEXTJOIN("-",TRUE,G1:G3)')
    engine.setCellFormula('Sheet1', 'H1', 'CHOOSE(1,B1:C2,D1:E2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: 'blue',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.String,
      value: 'alpha',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({
      tag: ValueTag.String,
      value: 'beta',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'A4')).toEqual({
      tag: ValueTag.String,
      value: 'alpha-beta',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Sheet1', 'I2')).toEqual({ tag: ValueTag.Number, value: 40 })

    for (const address of ['A1', 'A2', 'A3', 'A4', 'H1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes byte-oriented text builtins through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'LENB("é")')
    engine.setCellFormula('Sheet1', 'B1', 'LEFTB("abcdef",2)')
    engine.setCellFormula('Sheet1', 'C1', 'MIDB("abcdef",3,2)')
    engine.setCellFormula('Sheet1', 'D1', 'RIGHTB("abcdef",3)')
    engine.setCellFormula('Sheet1', 'E1', 'FINDB("d","abcdef",3)')
    engine.setCellFormula('Sheet1', 'F1', 'SEARCHB("ph","alphabet")')
    engine.setCellFormula('Sheet1', 'G1', 'REPLACEB("alphabet",3,2,"Z")')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.String,
      value: 'ab',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.String,
      value: 'cd',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({
      tag: ValueTag.String,
      value: 'def',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({
      tag: ValueTag.String,
      value: 'alZabet',
      stringId: expect.any(Number),
    })

    for (const address of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes ADDRESS and dollar-format helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'ADDRESS(12,3)')
    engine.setCellFormula('Sheet1', 'B1', 'DOLLAR(-1234.5,1)')
    engine.setCellFormula('Sheet1', 'C1', 'DOLLARDE(1.08,16)')
    engine.setCellFormula('Sheet1', 'D1', 'DOLLARFR(1.5,16)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: '$C$12',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.String,
      value: '-$1,234.5',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1.5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1.08 })

    for (const address of ['A1', 'B1', 'C1', 'D1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes bitwise and base-conversion helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'BITAND(6,3)')
    engine.setCellFormula('Sheet1', 'B1', 'BITOR(6,3)')
    engine.setCellFormula('Sheet1', 'C1', 'BITXOR(6,3)')
    engine.setCellFormula('Sheet1', 'D1', 'BITLSHIFT(1,4)')
    engine.setCellFormula('Sheet1', 'E1', 'BITRSHIFT(16,4)')
    engine.setCellFormula('Sheet1', 'F1', 'BASE(255,16,4)')
    engine.setCellFormula('Sheet1', 'G1', 'DECIMAL("00FF",16)')
    engine.setCellFormula('Sheet1', 'H1', 'BIN2DEC("1111111111")')
    engine.setCellFormula('Sheet1', 'I1', 'BIN2HEX("1111111111")')
    engine.setCellFormula('Sheet1', 'J1', 'BIN2OCT("1111111111")')
    engine.setCellFormula('Sheet1', 'K1', 'DEC2BIN(10,8)')
    engine.setCellFormula('Sheet1', 'L1', 'DEC2HEX(255,4)')
    engine.setCellFormula('Sheet1', 'M1', 'DEC2OCT(15,4)')
    engine.setCellFormula('Sheet1', 'N1', 'HEX2BIN("A",8)')
    engine.setCellFormula('Sheet1', 'O1', 'HEX2DEC("FFFFFFFFFF")')
    engine.setCellFormula('Sheet1', 'P1', 'HEX2OCT("F",4)')
    engine.setCellFormula('Sheet1', 'Q1', 'OCT2BIN("12",8)')
    engine.setCellFormula('Sheet1', 'R1', 'OCT2DEC("17")')
    engine.setCellFormula('Sheet1', 'S1', 'OCT2HEX("17",4)')
    engine.setCellFormula('Sheet1', 'T1', 'BESSELI(1.5,1)')
    engine.setCellFormula('Sheet1', 'U1', 'BESSELJ(1.9,2)')
    engine.setCellFormula('Sheet1', 'V1', 'BESSELK(1.5,1)')
    engine.setCellFormula('Sheet1', 'W1', 'BESSELY(2.5,1)')
    engine.setCellFormula('Sheet1', 'X1', 'CONVERT(6,"mi","km")')
    engine.setCellFormula('Sheet1', 'Y1', 'EUROCONVERT(1.2,"DEM","EUR")')
    engine.setCellFormula('Sheet1', 'Z1', 'EUROCONVERT(1,"FRF","DEM",TRUE,3)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 16 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.String,
      value: '00FF',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 255 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({
      tag: ValueTag.String,
      value: 'FFFFFFFFFF',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({
      tag: ValueTag.String,
      value: '7777777777',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({
      tag: ValueTag.String,
      value: '00FF',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({
      tag: ValueTag.String,
      value: '0017',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'O1')).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(engine.getCellValue('Sheet1', 'P1')).toEqual({
      tag: ValueTag.String,
      value: '0017',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'Q1')).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'R1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'S1')).toEqual({
      tag: ValueTag.String,
      value: '000F',
      stringId: expect.any(Number),
    })
    const besseli = engine.getCellValue('Sheet1', 'T1')
    expect(besseli).toMatchObject({ tag: ValueTag.Number })
    if (besseli.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELI cell to be numeric')
    }
    expect(besseli.value).toBeCloseTo(0.981666428, 7)
    const besselj = engine.getCellValue('Sheet1', 'U1')
    expect(besselj).toMatchObject({ tag: ValueTag.Number })
    if (besselj.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELJ cell to be numeric')
    }
    expect(besselj.value).toBeCloseTo(0.329925728, 7)
    const besselk = engine.getCellValue('Sheet1', 'V1')
    expect(besselk).toMatchObject({ tag: ValueTag.Number })
    if (besselk.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELK cell to be numeric')
    }
    expect(besselk.value).toBeCloseTo(0.277387804, 7)
    const bessely = engine.getCellValue('Sheet1', 'W1')
    expect(bessely).toMatchObject({ tag: ValueTag.Number })
    if (bessely.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELY cell to be numeric')
    }
    expect(bessely.value).toBeCloseTo(0.145918138, 7)
    expect(engine.getCellValue('Sheet1', 'X1')).toEqual({ tag: ValueTag.Number, value: 9.656064 })
    expect(engine.getCellValue('Sheet1', 'Y1')).toEqual({ tag: ValueTag.Number, value: 0.61 })
    const z1 = engine.getCellValue('Sheet1', 'Z1')
    expect(z1).toMatchObject({ tag: ValueTag.Number })
    if (z1.tag !== ValueTag.Number) {
      throw new Error('Expected EUROCONVERT triangulation cell to be numeric')
    }
    expect(z1.value).toBeCloseTo(0.29728616, 12)

    for (const address of [
      'A1',
      'B1',
      'C1',
      'D1',
      'E1',
      'F1',
      'G1',
      'H1',
      'I1',
      'J1',
      'K1',
      'L1',
      'M1',
      'N1',
      'O1',
      'P1',
      'Q1',
      'R1',
      'S1',
      'T1',
      'U1',
      'V1',
      'W1',
      'X1',
      'Y1',
      'Z1',
    ] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes USE.THE.COUNTIF through the wasm path as a COUNTIF alias', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', -2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'USE.THE.COUNTIF(A1:A3,">0")')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills FILTER with a computed comparison mask and UNIQUE through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'A3', 2)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellFormula('Sheet1', 'B1', 'FILTER(A1:A4,A1:A4>2)')
    engine.setCellFormula('Sheet1', 'C1', 'UNIQUE(A1:A4)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills FILTER through the wasm fast path when the include mask is a range', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'A3', 2)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellValue('Sheet1', 'B1', false)
    engine.setCellValue('Sheet1', 'B2', true)
    engine.setCellValue('Sheet1', 'B3', false)
    engine.setCellValue('Sheet1', 'B4', true)
    engine.setCellFormula('Sheet1', 'C1', 'FILTER(A1:A4,B1:B4)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates missing logical functions and lambda arrays', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'B1', 4)
    engine.setCellValue('Sheet1', 'B2', 5)
    engine.setCellValue('Sheet1', 'B3', 6)
    engine.setCellFormula('Sheet1', 'C1', 'IFS(A1>1,"big",TRUE(),"small")')
    engine.setCellFormula('Sheet1', 'D1', 'SWITCH(A1,1,"one","other")')
    engine.setCellFormula('Sheet1', 'E1', 'XOR(TRUE(),FALSE(),TRUE())')
    engine.setCellFormula('Sheet1', 'F1', 'LAMBDA(x,x+1)(4)')
    engine.setCellFormula('Sheet1', 'G1', 'MAP(A1:A3,LAMBDA(x,x*2))')
    engine.setCellFormula('Sheet1', 'H1', 'BYROW(A1:B3,LAMBDA(r,SUM(r)))')
    engine.setCellFormula('Sheet1', 'I1', 'BYCOL(A1:B3,LAMBDA(c,SUM(c)))')
    engine.setCellFormula('Sheet1', 'K1', 'REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))')
    engine.setCellFormula('Sheet1', 'L1', 'SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))')
    engine.setCellFormula('Sheet1', 'M1', 'MAKEARRAY(2,2,LAMBDA(r,c,r+c))')
    engine.setCellFormula('Sheet1', 'O1', 'BYROW(A1:B3,LAMBDA(r,AVERAGE(r)))')
    engine.setCellFormula('Sheet1', 'P1', 'BYCOL(A1:B3,LAMBDA(c,COUNTA(c)))')
    engine.setCellFormula('Sheet1', 'R1', 'REDUCE(1,A1:A3,LAMBDA(acc,x,acc*x))')
    engine.setCellFormula('Sheet1', 'S1', 'SCAN(1,A1:A3,LAMBDA(acc,x,acc*x))')

    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.String,
      value: 'small',
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.String,
      value: 'one',
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'G3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'H3')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'L2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'L3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'M2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'N2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'O1')).toEqual({ tag: ValueTag.Number, value: 2.5 })
    expect(engine.getCellValue('Sheet1', 'O2')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'O3')).toEqual({ tag: ValueTag.Number, value: 4.5 })
    expect(engine.getCellValue('Sheet1', 'P1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'Q1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'R1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'S1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'S2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'S3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'O1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'P1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'R1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'S1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates accelerated math builtins and JS matrix spills', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'B2', 4)
    engine.setCellValue('Sheet1', 'C1', 4)
    engine.setCellValue('Sheet1', 'C2', 5)
    engine.setCellValue('Sheet1', 'D1', 6)
    engine.setCellValue('Sheet1', 'D2', 7)
    engine.setCellValue('Sheet1', 'K1', Math.PI / 2)
    engine.setCellValue('Sheet1', 'K2', -3.98)
    engine.setCellFormula('Sheet1', 'E1', 'SIN(K1)')
    engine.setCellFormula('Sheet1', 'F1', 'TRUNC(K2,1)')
    engine.setCellFormula('Sheet1', 'G1', 'MUNIT(2)')
    engine.setCellFormula('Sheet1', 'I1', 'MMULT(A1:B2,C1:D2)')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: -3.9 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 19 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 27 })
    expect(engine.getCellValue('Sheet1', 'I2')).toEqual({ tag: ValueTag.Number, value: 28 })
    expect(engine.getCellValue('Sheet1', 'J2')).toEqual({ tag: ValueTag.Number, value: 40 })
  })

  it('supports cross-sheet references', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 4)
    engine.setCellFormula('Sheet2', 'B2', 'Sheet1!A1*3')
    expect(engine.getCellValue('Sheet2', 'B2')).toEqual({ tag: ValueTag.Number, value: 12 })
  })

  it('recalculates TODAY and NOW on the wasm path for each recalc-triggering batch', async () => {
    vi.useFakeTimers()
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    vi.setSystemTime(new Date('2026-03-19T15:45:30.000Z'))
    engine.setCellFormula('Sheet1', 'A1', 'TODAY()')
    engine.setCellFormula('Sheet1', 'B1', 'NOW()')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(utcDateToExcelSerial(new Date('2026-03-19T15:45:30.000Z'))),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Number,
      value: utcDateToExcelSerial(new Date('2026-03-19T15:45:30.000Z')),
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 0 })

    vi.setSystemTime(new Date('2026-03-20T01:02:03.000Z'))
    engine.setCellValue('Sheet1', 'C1', 1)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(utcDateToExcelSerial(new Date('2026-03-20T01:02:03.000Z'))),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Number,
      value: utcDateToExcelSerial(new Date('2026-03-20T01:02:03.000Z')),
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 0 })
  })

  it('recalculates RAND on the wasm path for each recalc-triggering batch', async () => {
    const randomSpy = vi.spyOn(Math, 'random')
    randomSpy.mockReturnValueOnce(0.125).mockReturnValueOnce(0.875)

    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'RAND()')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.125 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'B1', 1)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.875 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('evaluates row and column aggregate ranges on the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A3', 5)
    engine.setCellValue('Sheet1', 'B3', 7)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A:A)')
    engine.setCellFormula('Sheet1', 'C2', 'SUM(3:3)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for supported aggregate formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)+ABS(A1/2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)
  })

  it('deduplicates overlapped precedents when scalar refs and ranges touch the same cell', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)+A1')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getDependencies('Sheet1', 'B1').directPrecedents).toEqual(['Sheet1!A1', 'Sheet1!A2'])
  })

  it('uses the wasm fast path for IF branch formulas once comparison parity exists', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellValue('Sheet1', 'A2', 9)
    engine.setCellFormula('Sheet1', 'B1', 'IF(A1>0,A1*2,A2-1)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', 0)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity logical builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 0)
    engine.setCellFormula('Sheet1', 'B1', 'AND(A1,TRUE)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B2', 'OR(A2,FALSE)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B3', 'NOT(A2)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B11', 'AND(TRUE,A4)')
    expect(engine.getCellValue('Sheet1', 'B11')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B12', 'OR(A4,TRUE)')
    expect(engine.getCellValue('Sheet1', 'B12')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B13', 'NOT(2)')
    expect(engine.getCellValue('Sheet1', 'B13')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B14', 'IF(1/0,1,2)')
    expect(engine.getCellValue('Sheet1', 'B14')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A3', 'hello')
    engine.setCellFormula('Sheet1', 'B4', 'AND(A3,TRUE)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B5', 'IFERROR(A1/0,"fallback")')
    expect(engine.getCellValue('Sheet1', 'B5')).toMatchObject({
      tag: ValueTag.String,
      value: 'fallback',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B6', 'IFNA(NA(),"missing")')
    expect(engine.getCellValue('Sheet1', 'B6')).toMatchObject({
      tag: ValueTag.String,
      value: 'missing',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for INDEX VLOOKUP and HLOOKUP', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'pear')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'A2', 'apple')
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'D1', 'Q1')
    engine.setCellValue('Sheet1', 'E1', 'Q2')
    engine.setCellValue('Sheet1', 'F1', 'Q3')
    engine.setCellValue('Sheet1', 'D2', 100)
    engine.setCellValue('Sheet1', 'E2', 200)
    engine.setCellValue('Sheet1', 'F2', 300)
    engine.setCellFormula('Sheet1', 'H1', 'INDEX(A1:B2,2,2)')
    engine.setCellFormula('Sheet1', 'H2', 'VLOOKUP("apple",A1:B2,2,FALSE)')
    engine.setCellFormula('Sheet1', 'H3', 'HLOOKUP("Q3",D1:F2,2,FALSE)')

    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H3')).toEqual({ tag: ValueTag.Number, value: 300 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
  })

  it('uses the direct criteria path for conditional aggregates and keeps SUMPRODUCT on wasm', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellValue('Sheet1', 'A3', -1)
    engine.setCellValue('Sheet1', 'A4', 6)
    engine.setCellValue('Sheet1', 'B1', 'x')
    engine.setCellValue('Sheet1', 'B2', 'x')
    engine.setCellValue('Sheet1', 'B3', 'y')
    engine.setCellValue('Sheet1', 'B4', 'x')
    engine.setCellValue('Sheet1', 'C1', 10)
    engine.setCellValue('Sheet1', 'C2', 20)
    engine.setCellValue('Sheet1', 'C3', 30)
    engine.setCellValue('Sheet1', 'C4', 40)
    engine.setCellValue('Sheet1', 'D1', 1)
    engine.setCellValue('Sheet1', 'D2', 2)
    engine.setCellValue('Sheet1', 'D3', 3)
    engine.setCellValue('Sheet1', 'E1', 4)
    engine.setCellValue('Sheet1', 'E2', 5)
    engine.setCellValue('Sheet1', 'E3', 6)

    engine.setCellFormula('Sheet1', 'F1', 'COUNTIF(A1:A4,">0")')
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F2', 'COUNTIFS(A1:A4,">0",B1:B4,"x")')
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F2').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F3', 'SUMIF(A1:A4,">0",C1:C4)')
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F3').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F4', 'SUMIFS(C1:C4,A1:A4,">0",B1:B4,"x")')
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F4').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F5', 'AVERAGEIF(A1:A4,">0")')
    expect(engine.getCellValue('Sheet1', 'F5')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F5').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F6', 'AVERAGEIFS(C1:C4,A1:A4,">0",B1:B4,"x")')
    expect(engine.getCellValue('Sheet1', 'F6')).toEqual({
      tag: ValueTag.Number,
      value: (10 + 20 + 40) / 3,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F6').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F7', 'SUMPRODUCT(D1:D3,E1:E3)')
    expect(engine.getCellValue('Sheet1', 'F7')).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A3', 8)
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getCellValue('Sheet1', 'F5')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'F6')).toEqual({
      tag: ValueTag.Number,
      value: (10 + 20 + 40) / 3,
    })
    expect(engine.getLastMetrics().jsFormulaCount).toBe(0)
  })

  it('binds direct criteria descriptors for cell-driven criteria and handles min and max aggregates', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'criteria-cell-driven' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'x')
    engine.setCellValue('Sheet1', 'A2', 'y')
    engine.setCellValue('Sheet1', 'A3', 'x')
    engine.setCellValue('Sheet1', 'A4', 'y')
    engine.setCellValue('Sheet1', 'C1', 10)
    engine.setCellValue('Sheet1', 'C2', 20)
    engine.setCellValue('Sheet1', 'C3', 30)
    engine.setCellValue('Sheet1', 'C4', 40)
    engine.setCellValue('Sheet1', 'D1', 'x')
    engine.setCellFormula('Sheet1', 'F1', 'COUNTIF(A1:A4,D1)')
    engine.setCellFormula('Sheet1', 'F2', 'MINIFS(C1:C4,A1:A4,D1)')
    engine.setCellFormula('Sheet1', 'F3', 'MAXIFS(C1:C4,A1:A4,D1)')
    engine.setCellFormula('Sheet1', 'F4', 'AVERAGEIFS(C1:C4,A1:A4,D1)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F4').mode).toBe(FormulaMode.WasmFastPath)

    const countIndex = engine.workbook.getCellIndex('Sheet1', 'F1')
    const minIndex = engine.workbook.getCellIndex('Sheet1', 'F2')
    const maxIndex = engine.workbook.getCellIndex('Sheet1', 'F3')
    const averageIndex = engine.workbook.getCellIndex('Sheet1', 'F4')
    if (countIndex === undefined || minIndex === undefined || maxIndex === undefined || averageIndex === undefined) {
      throw new Error('expected direct criteria formulas to exist')
    }

    const countFormula = readRuntimeFormula(engine, countIndex)
    if (!isRuntimeFormulaWithDirectCriteria(countFormula)) {
      throw new Error('expected COUNTIF runtime formula to expose direct criteria metadata')
    }
    expect(countFormula.directCriteria.aggregateKind).toBe('count')
    expect(countFormula.directCriteria.aggregateRange).toBeUndefined()
    expect(countFormula.directCriteria.criteriaPairs).toHaveLength(1)
    expect(countFormula.directCriteria.criteriaPairs[0]?.criterion).toMatchObject({
      kind: 'cell',
    })

    const minFormula = readRuntimeFormula(engine, minIndex)
    if (!isRuntimeFormulaWithDirectCriteria(minFormula)) {
      throw new Error('expected MINIFS runtime formula to expose direct criteria metadata')
    }
    expect(minFormula.directCriteria.aggregateKind).toBe('min')
    expect(minFormula.directCriteria.aggregateRange).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 2,
      length: 4,
    })

    const maxFormula = readRuntimeFormula(engine, maxIndex)
    if (!isRuntimeFormulaWithDirectCriteria(maxFormula)) {
      throw new Error('expected MAXIFS runtime formula to expose direct criteria metadata')
    }
    expect(maxFormula.directCriteria.aggregateKind).toBe('max')

    const averageFormula = readRuntimeFormula(engine, averageIndex)
    if (!isRuntimeFormulaWithDirectCriteria(averageFormula)) {
      throw new Error('expected AVERAGEIFS runtime formula to expose direct criteria metadata')
    }
    expect(averageFormula.directCriteria.aggregateKind).toBe('average')

    engine.setCellValue('Sheet1', 'D1', 'y')
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 40 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 30 })

    engine.setCellValue('Sheet1', 'D1', 'z')
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
  })

  it('binds direct aggregate descriptors for bounded single-column SUM, AVERAGE, and COUNT', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'B2', 'AVERAGE(A1:A3)')
    engine.setCellFormula('Sheet1', 'B3', 'COUNT(A1:A3)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B2').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'B3').mode).toBe(FormulaMode.WasmFastPath)

    const sumIndex = engine.workbook.getCellIndex('Sheet1', 'B1')
    const averageIndex = engine.workbook.getCellIndex('Sheet1', 'B2')
    const countIndex = engine.workbook.getCellIndex('Sheet1', 'B3')
    if (sumIndex === undefined || averageIndex === undefined || countIndex === undefined) {
      throw new Error('expected direct aggregate formulas to exist')
    }

    const sumFormula = readRuntimeFormula(engine, sumIndex)
    const averageFormula = readRuntimeFormula(engine, averageIndex)
    const countFormula = readRuntimeFormula(engine, countIndex)
    if (!isRuntimeFormulaWithDirectAggregate(sumFormula)) {
      throw new Error('expected SUM runtime formula to expose direct aggregate metadata')
    }
    if (!isRuntimeFormulaWithDirectAggregate(averageFormula)) {
      throw new Error('expected AVERAGE runtime formula to expose direct aggregate metadata')
    }
    if (!isRuntimeFormulaWithDirectAggregate(countFormula)) {
      throw new Error('expected COUNT runtime formula to expose direct aggregate metadata')
    }
    expect(isRuntimeFormulaWithRanges(sumFormula)).toBe(true)
    expect(sumFormula.rangeDependencies).toHaveLength(0)
    expect(isRuntimeFormulaWithDependencies(sumFormula)).toBe(true)
    expect(sumFormula.dependencyIndices).toEqual(new Uint32Array())

    expect(sumFormula.directAggregate).toMatchObject({
      aggregateKind: 'sum',
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
      length: 3,
    })
    expect(averageFormula.directAggregate.aggregateKind).toBe('average')
    expect(countFormula.directAggregate.aggregateKind).toBe('count')

    engine.setCellValue('Sheet1', 'A1', 5)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('rebinds direct aggregate formulas when a formula appears inside a previously literal aggregate range', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-rebind-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })

    engine.setCellFormula('Sheet1', 'A2', 'A1*3')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })

    engine.setCellValue('Sheet1', 'A1', 4)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 16 })
  })

  it('keeps direct aggregate dependents current through coordinate mutation APIs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-coordinate-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 2)
    engine.setCellValueAt(sheetId, 1, 0, 3)
    engine.setCellFormulaAt(sheetId, 0, 1, 'SUM(A1:A2)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })

    engine.setCellFormulaAt(sheetId, 1, 0, 'A1*3')
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })

    engine.setCellValueAt(sheetId, 0, 0, 4)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 16 })
  })

  it('keeps direct aggregate formulas current through coordinate clears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-clear-at-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 2)
    engine.setCellValueAt(sheetId, 1, 0, 3)
    engine.setCellFormulaAt(sheetId, 0, 1, 'SUM(A1:A2)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })

    engine.clearCellAt(sheetId, 1, 0)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('keeps lookup formulas current through coordinate literal writes', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'lookup-coordinate-write-spec',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 10)
    engine.setCellValueAt(sheetId, 1, 0, 20)
    engine.setCellValueAt(sheetId, 2, 0, 30)
    engine.setCellValueAt(sheetId, 0, 3, 20)
    engine.setCellValueAt(sheetId, 1, 3, 25)
    engine.setCellFormulaAt(sheetId, 0, 4, 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormulaAt(sheetId, 0, 5, 'MATCH(D2,A1:A3,1)')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValueAt(sheetId, 1, 0, 25)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValueAt(sheetId, 1, 0, 20)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('keeps lookup formulas current through coordinate clears', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'lookup-coordinate-clear-spec',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 10)
    engine.setCellValueAt(sheetId, 1, 0, 20)
    engine.setCellValueAt(sheetId, 2, 0, 30)
    engine.setCellValueAt(sheetId, 0, 3, 20)
    engine.setCellValueAt(sheetId, 1, 3, 25)
    engine.setCellFormulaAt(sheetId, 0, 4, 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormulaAt(sheetId, 0, 5, 'MATCH(D2,A1:A3,1)')

    engine.clearCellAt(sheetId, 1, 0)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('uses the direct js path for exact MATCH and XMATCH while keeping XLOOKUP on wasm', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'apple')
    engine.setCellValue('Sheet1', 'A2', 'pear')
    engine.setCellValue('Sheet1', 'A3', 'pear')
    engine.setCellValue('Sheet1', 'A4', 'plum')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'B3', 30)
    engine.setCellValue('Sheet1', 'B4', 40)
    engine.setCellValue('Sheet1', 'C1', 1)
    engine.setCellValue('Sheet1', 'C2', 3)
    engine.setCellValue('Sheet1', 'C3', 5)

    engine.setCellFormula('Sheet1', 'D1', 'MATCH("pear",A1:A4,0)')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    engine.setCellFormula('Sheet1', 'D3', 'XMATCH("pear",A1:A4,0,-1)')
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    engine.setCellFormula('Sheet1', 'D4', 'XLOOKUP("pear",A1:A4,B1:B4)')
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D5', 'XLOOKUP("missing",A1:A4,B1:B4,"fallback")')
    expect(engine.getCellValue('Sheet1', 'D5')).toMatchObject({
      tag: ValueTag.String,
      value: 'fallback',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', 'pear')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 2 })
  })

  it('uses the direct js path for approximate sorted MATCH', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'approx-lookup' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'C1', 1)
    engine.setCellValue('Sheet1', 'C2', 3)
    engine.setCellValue('Sheet1', 'C3', 5)
    engine.setCellValue('Sheet1', 'D1', 4)

    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,C1:C3,1)')
    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    const operandCellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    const runtimeFormula = formulaCellIndex === undefined ? undefined : readRuntimeFormula(engine, formulaCellIndex)
    expect(isRuntimeFormulaWithDirectLookup(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.directLookup.kind).toBe('approximate-uniform-numeric')
    expect(runtimeFormula?.directLookup.operandCellIndex).toBe(operandCellIndex)
    expect(runtimeFormula?.directLookup).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 2,
      length: 3,
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'C2', 4)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })
  })

  it('skips dirtying approximate MATCH when an irrelevant high-side tail value changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'approx-lookup-tail' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 3, 4, 5].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    engine.setCellValue('Sheet1', 'D1', 2.5)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A5,1)')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValue('Sheet1', 'A5', 6)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
  })

  it('uses the direct indexed path for exact MATCH when column indexing is enabled', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'indexed-lookup', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'D1', 2)

    engine.setCellFormula('Sheet1', 'E1', '=MATCH(D1,A1:A3,0)')
    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    const operandCellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    const runtimeFormula = formulaCellIndex === undefined ? undefined : readRuntimeFormula(engine, formulaCellIndex)
    expect(isRuntimeFormulaWithDirectLookup(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.directLookup.kind).toBe('exact-uniform-numeric')
    expect(runtimeFormula?.directLookup.operandCellIndex).toBe(operandCellIndex)
    expect(runtimeFormula?.directLookup).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
      length: 3,
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', 10)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })

    engine.setCellValue('Sheet1', 'A2', 20)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    engine.setCellValue('Sheet1', 'D1', 3)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })
  })

  it('uses the direct indexed path for exact string MATCH and reverse XMATCH when column indexing is enabled', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'indexed-string-lookup',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'pear')
    engine.setCellValue('Sheet1', 'A2', 'apple')
    engine.setCellValue('Sheet1', 'A3', 'pear')

    engine.setCellFormula('Sheet1', 'B1', '=MATCH("APPLE",A1:A3,0)')
    engine.setCellFormula('Sheet1', 'B2', '=XMATCH("pear",A1:A3,0,-1)')

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'B2').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 3 })

    engine.setCellValue('Sheet1', 'A3', 'banana')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 2, wasmFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity info and date builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A3', 'hello')
    engine.setCellValue('Sheet1', 'A4', 45351)
    engine.setCellValue('Sheet1', 'A5', 45351.75)
    engine.setCellValue('Sheet1', 'A6', 60)
    engine.setCellValue('Sheet1', 'A7', 45322)
    engine.setCellValue('Sheet1', 'A8', 45337)
    engine.setCellValue('Sheet1', 'A9', 'bad')
    engine.setCellValue('Sheet1', 'A10', 0.5208333333333334)
    engine.setCellValue('Sheet1', 'A11', 0.5208449074074074)

    engine.setCellFormula('Sheet1', 'B1', 'ISBLANK()')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'ISNUMBER()')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'ISTEXT()')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B4', 'ISBLANK(A1)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B5', 'ISNUMBER(A1)')
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B6', 'ISTEXT(A3)')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B7', 'DATE(2024,2,29)')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 45351 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B8', 'YEAR(B7)')
    expect(engine.getCellValue('Sheet1', 'B8')).toEqual({ tag: ValueTag.Number, value: 2024 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B9', 'MONTH(A5)')
    expect(engine.getCellValue('Sheet1', 'B9')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B10', 'DAY(A6)')
    expect(engine.getCellValue('Sheet1', 'B10')).toEqual({ tag: ValueTag.Number, value: 29 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B11', 'EDATE(A7,1.9)')
    expect(engine.getCellValue('Sheet1', 'B11')).toEqual({ tag: ValueTag.Number, value: 45351 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B12', 'EOMONTH(A8,A2)')
    expect(engine.getCellValue('Sheet1', 'B12')).toEqual({ tag: ValueTag.Number, value: 45382 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B13', 'DATE(A9,2,29)')
    expect(engine.getCellValue('Sheet1', 'B13')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B14', 'EDATE(A9,1)')
    expect(engine.getCellValue('Sheet1', 'B14')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B15', 'EOMONTH(A9,1)')
    expect(engine.getCellValue('Sheet1', 'B15')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B16', 'TIME(12,30,0)')
    expect(engine.getCellValue('Sheet1', 'B16')).toEqual({
      tag: ValueTag.Number,
      value: 0.5208333333333334,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B17', 'HOUR(A10)')
    expect(engine.getCellValue('Sheet1', 'B17')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B18', 'MINUTE(A10)')
    expect(engine.getCellValue('Sheet1', 'B18')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B19', 'SECOND(A11)')
    expect(engine.getCellValue('Sheet1', 'B19')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B20', 'WEEKDAY(DATE(2026,3,15))')
    expect(engine.getCellValue('Sheet1', 'B20')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity information and threshold helpers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A3', 'alpha')

    engine.setCellFormula('Sheet1', 'B1', 'T(A3)')
    engine.setCellFormula('Sheet1', 'B2', 'N(A2)')
    engine.setCellFormula('Sheet1', 'B3', 'TYPE(A3)')
    engine.setCellFormula('Sheet1', 'B4', 'DELTA(4,4)')
    engine.setCellFormula('Sheet1', 'B5', 'GESTEP(-1)')
    engine.setCellFormula('Sheet1', 'B6', 'GAUSS(0)')
    engine.setCellFormula('Sheet1', 'B7', 'PHI(0)')

    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'alpha',
    })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'B6')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    })
    expect(engine.getCellValue('Sheet1', 'B7')).toMatchObject({
      tag: ValueTag.Number,
      value: 0.3989422804014327,
    })

    for (const address of ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for literal and dynamic VALUE coercion', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'VALUE("42")')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', '  -17.25e1  ')
    engine.setCellFormula('Sheet1', 'B2', 'VALUE(A1)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: -172.5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A2', 'not-a-number')
    engine.setCellFormula('Sheet1', 'B3', 'VALUE(A2)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity LEN builtin', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', null)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A3', 123.4)
    engine.setCellValue('Sheet1', 'A4', 'hello')
    engine.setCellFormula('Sheet1', 'A5', 'A3/0')

    engine.setCellFormula('Sheet1', 'B1', 'LEN(A1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'LEN(A2)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'LEN(A3)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B4', 'LEN(A4)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B5', 'LEN(A5)')
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for EXACT text equality', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Alpha')
    engine.setCellValue('Sheet1', 'A2', 'alpha')

    engine.setCellFormula('Sheet1', 'B1', 'EXACT(A1,A1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'EXACT(A1,A2)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for string literals, direct refs, and CONCAT', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'world')

    engine.setCellFormula('Sheet1', 'B1', '"hello"')
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'hello',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'A1')
    expect(engine.getCellValue('Sheet1', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'world',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'CONCAT("hi ",A1)')
    expect(engine.getCellValue('Sheet1', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'hi world',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for text slicing, casing, and search builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Alpha')
    engine.setCellValue('Sheet1', 'A2', '  alpha   beta  ')
    engine.setCellValue('Sheet1', 'A3', 'alpha')
    engine.setCellValue('Sheet1', 'A4', 'BETA')
    engine.setCellValue('Sheet1', 'A5', 'alphabet')

    engine.setCellFormula('Sheet1', 'B1', 'LEFT(A1,2)')
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'Al',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'RIGHT(A1)')
    expect(engine.getCellValue('Sheet1', 'B2')).toMatchObject({ tag: ValueTag.String, value: 'a' })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'MID(A1,2,3)')
    expect(engine.getCellValue('Sheet1', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'lph',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B4', 'TRIM(A2)')
    expect(engine.getCellValue('Sheet1', 'B4')).toMatchObject({
      tag: ValueTag.String,
      value: 'alpha beta',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B5', 'UPPER(A3)')
    expect(engine.getCellValue('Sheet1', 'B5')).toMatchObject({
      tag: ValueTag.String,
      value: 'ALPHA',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B6', 'LOWER(A4)')
    expect(engine.getCellValue('Sheet1', 'B6')).toMatchObject({
      tag: ValueTag.String,
      value: 'beta',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B7', 'FIND("ph",A5,3)')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B8', 'SEARCH("P*",A5)')
    expect(engine.getCellValue('Sheet1', 'B8')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity rounding builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 123.4)
    engine.setCellFormula('Sheet1', 'B1', 'ROUND(A1,-1)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 120 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B2', 'FLOOR(TRUE,0.5)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B3', 'CEILING(7,2)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B4', 'FLOOR(A1,0)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellValue('Sheet1', 'A2', 'oops')
    engine.setCellFormula('Sheet1', 'B5', 'ROUND(A2,1)')
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B6', 'INT(-3.1)')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B7', 'ROUNDUP(-3.141,2)')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: -3.15 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B8', 'ROUNDDOWN(-3.141,2)')
    expect(engine.getCellValue('Sheet1', 'B8')).toEqual({ tag: ValueTag.Number, value: -3.14 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)
  })

  it('preserves topo order across mixed wasm and js formula runs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 5)
    engine.setCellFormula('Sheet1', 'B2', 'A1+A2')
    engine.setCellFormula('Sheet1', 'D1', 'SUM(2:2)')

    engine.setCellValue('Sheet1', 'A1', 12)

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(2)
    expect(engine.getLastMetrics().jsFormulaCount).toBe(0)
  })

  it('rebinds formulas when a referenced sheet appears later', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'Sheet2!B1*2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.createSheet('Sheet2')
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet2', 'B1', 3)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('rebinds bounded cross-sheet ranges from #REF! back onto the wasm path when a sheet appears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'SUM(Sheet2!A1:A2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.createSheet('Sheet2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet2', 'A1', 2)
    engine.setCellValue('Sheet2', 'A2', 3)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('rebinds formulas to #REF! when a referenced sheet is deleted', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet2', 'B1', 3)
    engine.setCellFormula('Sheet1', 'A1', 'Sheet2!B1*2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.deleteSheet('Sheet2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('rebinds bounded cross-sheet ranges to #REF! when a referenced sheet is deleted', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet2', 'A1', 2)
    engine.setCellValue('Sheet2', 'A2', 3)
    engine.setCellFormula('Sheet1', 'A1', 'SUM(Sheet2!A1:A2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.deleteSheet('Sheet2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('renames sheets without breaking formulas, names, or sheet metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }, [
      [2, 3],
      [4, 5],
    ])
    engine.setCellFormula('Data', 'C1', 'A1+B1')
    engine.setCellFormula('Summary', 'A1', 'SUM(Data!A1:B2)')
    engine.setDefinedName('AnchorCell', { kind: 'cell-ref', sheetName: 'Data', address: 'A1' })
    engine.setDefinedName('SalesRange', '=Data!A1:B2')
    engine.setFreezePane('Data', 1, 0)
    engine.setFilter('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' })
    engine.setSort('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Q1', 'Q2'],
      headerRow: false,
      totalsRow: false,
    })
    engine.setPivotTable('Summary', 'D2', {
      name: 'SalesPivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' },
      groupBy: ['Q1'],
      values: [{ sourceColumn: 'Q2', summarizeBy: 'sum' }],
    })

    engine.renameSheet('Data', 'Revenue')

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toEqual(['Revenue', 'Summary'])
    expect(engine.getCell('Revenue', 'C1').formula).toBe('A1+B1')
    expect(engine.getCellValue('Revenue', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCell('Summary', 'A1').formula).toBe('SUM(Revenue!A1:B2)')
    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getDefinedName('AnchorCell')).toEqual({
      name: 'AnchorCell',
      value: { kind: 'cell-ref', sheetName: 'Revenue', address: 'A1' },
    })
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: '=Revenue!A1:B2',
    })
    expect(engine.getFreezePane('Revenue')).toEqual({ sheetName: 'Revenue', rows: 1, cols: 0 })
    expect(engine.getFreezePane('Data')).toBeUndefined()
    expect(engine.getFilters('Revenue')).toEqual([
      {
        sheetName: 'Revenue',
        range: { sheetName: 'Revenue', startAddress: 'A1', endAddress: 'B2' },
      },
    ])
    expect(engine.getSorts('Revenue')).toEqual([
      {
        sheetName: 'Revenue',
        range: { sheetName: 'Revenue', startAddress: 'A1', endAddress: 'B2' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])
    expect(engine.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Revenue',
        startAddress: 'A1',
        endAddress: 'B2',
        columnNames: ['Q1', 'Q2'],
        headerRow: false,
        totalsRow: false,
      },
    ])
    expect(engine.getPivotTables()).toEqual([
      {
        name: 'SalesPivot',
        sheetName: 'Summary',
        address: 'D2',
        source: { sheetName: 'Revenue', startAddress: 'A1', endAddress: 'B2' },
        groupBy: ['Q1'],
        values: [{ sourceColumn: 'Q2', summarizeBy: 'sum' }],
        rows: 1,
        cols: 1,
      },
    ])
  })

  it('clears reverse range edges when a range-backed formula is removed', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A:A)')

    expect(engine.getDependencies('Sheet1', 'A1').directDependents).toContain('Sheet1!B1')

    engine.clearCell('Sheet1', 'B1')

    expect(engine.getDependencies('Sheet1', 'A1').directDependents).toEqual([])
  })

  it('rebinds column and row range formulas when new cells materialize later', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A:A)')
    engine.setCellFormula('Sheet1', 'B3', 'SUM(2:2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 0 })

    engine.setCellValue('Sheet1', 'A4', 3)
    engine.setCellValue('Sheet1', 'C2', 5)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getDependencies('Sheet1', 'B1').directPrecedents).toEqual(['Sheet1!A1', 'Sheet1!A4'])
    expect(engine.getDependencies('Sheet1', 'B3').directPrecedents).toEqual(['Sheet1!C2'])

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(b1Index).toBeDefined()
    const runtimeFormula = b1Index === undefined ? undefined : readRuntimeFormula(engine, b1Index)
    expect(isRuntimeFormulaWithDependencies(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.dependencyIndices).toBeInstanceOf(Uint32Array)
  })

  it('converges under reordered replicated batches and restores replica state', async () => {
    const engineA = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const engineB = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([engineA.ready(), engineB.ready()])

    const outboundA: EngineOpBatch[] = []
    const outboundB: EngineOpBatch[] = []
    engineA.subscribeBatches((batch) => outboundA.push(batch))
    engineB.subscribeBatches((batch) => outboundB.push(batch))

    engineA.createSheet('Sheet1')
    engineB.createSheet('Sheet1')
    engineA.setCellValue('Sheet1', 'A1', 1)
    engineB.setCellValue('Sheet1', 'A1', 2)

    ;[...outboundB].toReversed().forEach((batch) => engineA.applyRemoteBatch(batch))
    ;[...outboundA].forEach((batch) => engineB.applyRemoteBatch(batch))

    expect(engineA.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engineB.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })

    const restored = new SpreadsheetEngine({ workbookName: 'restored', replicaId: 'b' })
    await restored.ready()
    restored.importSnapshot(engineB.exportSnapshot())
    restored.importReplicaSnapshot(engineB.exportReplicaSnapshot())

    const latestOutboundA = outboundA.at(-1)
    expect(latestOutboundA).toBeDefined()
    restored.applyRemoteBatch(latestOutboundA)
    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('ignores duplicate remote batches and stale cell replays behind sheet tombstones', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: EngineOpBatch[] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Sheet1')
    const createBatch = outbound.at(-1)
    expect(createBatch).toBeDefined()

    primary.setCellValue('Sheet1', 'A1', 7)
    const valueBatch = outbound.at(-1)
    expect(valueBatch).toBeDefined()

    replica.applyRemoteBatch(createBatch)
    expect(replica.applyRemoteBatch(valueBatch)).toBe(true)
    const versionBeforeDuplicate = replica.explainCell('Sheet1', 'A1').version

    expect(replica.applyRemoteBatch(valueBatch)).toBe(false)
    expect(replica.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(replica.explainCell('Sheet1', 'A1').version).toBe(versionBeforeDuplicate)

    primary.deleteSheet('Sheet1')
    const deleteBatch = outbound.at(-1)
    expect(deleteBatch).toBeDefined()
    replica.applyRemoteBatch(deleteBatch)
    expect(replica.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    const restored = new SpreadsheetEngine({ workbookName: 'restored', replicaId: 'b' })
    await restored.ready()
    restored.importSnapshot(replica.exportSnapshot())
    restored.importReplicaSnapshot(replica.exportReplicaSnapshot())

    restored.applyRemoteBatch(valueBatch)
    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
  })

  it('resolves workbook defined names through engine metadata and recalculates dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 100)
    engine.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')

    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    const changed: number[][] = []
    const unsubscribe = engine.subscribe((event) => {
      changed.push(Array.from(event.changedCellIndices))
    })

    const a2Index = engine.workbook.getCellIndex('Sheet1', 'A2')
    expect(a2Index).toBeDefined()

    engine.setDefinedName('TaxRate', 0.085)

    expect(engine.getDefinedName('taxrate')).toEqual({ name: 'TaxRate', value: 0.085 })
    expect(engine.getDefinedNames()).toEqual([{ name: 'TaxRate', value: 0.085 }])
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8.5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    expect(changed.at(-1)).toContain(a2Index!)

    engine.setDefinedName('TAXRATE', 0.09)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    expect(engine.deleteDefinedName('taxrate')).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    unsubscribe()
  })

  it('replicates defined-name batches and replays them through transaction history', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: EngineOpBatch[] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Sheet1')
    primary.setCellValue('Sheet1', 'A1', 100)
    primary.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')
    outbound.forEach((batch) => replica.applyRemoteBatch(batch))

    primary.setDefinedName('TaxRate', 0.08)
    const defineBatch = outbound.at(-1)
    expect(defineBatch?.ops).toEqual([{ kind: 'upsertDefinedName', name: 'TaxRate', value: 0.08 }])
    expect(primary.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8 })

    replica.applyRemoteBatch(defineBatch)
    expect(replica.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8 })

    expect(primary.undo()).toBe(true)
    expect(primary.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    expect(primary.redo()).toBe(true)
    expect(primary.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8 })

    primary.deleteDefinedName('taxrate')
    const deleteBatch = outbound.at(-1)
    expect(deleteBatch?.ops).toEqual([{ kind: 'deleteDefinedName', name: 'taxrate' }])

    replica.applyRemoteBatch(deleteBatch)
    expect(replica.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
  })

  it('emits lightweight tracked events for ordinary mutations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const tracked = vi.fn()

    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.setCellValue('Sheet1', 'A1', 7)

    expect(tracked).toHaveBeenCalledTimes(1)
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: new Uint32Array([0]),
        invalidatedRows: [],
        invalidatedColumns: [],
      }),
    )

    unsubscribe()
  })

  it('skips no-op defined-name, sort, and table writes and reports missing clears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const outbound: EngineOpBatch[] = []
    const unsubscribeBatches = engine.subscribeBatches((batch) => outbound.push(batch))

    expect(engine.deleteDefinedName('MissingRate')).toBe(false)
    expect(
      engine.clearSort('Sheet1', {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      }),
    ).toBe(false)
    expect(engine.getCellNumberFormat(undefined)).toMatchObject({
      kind: 'general',
      code: 'general',
    })
    expect(engine.getWorkbookMetadata('locale')).toBeUndefined()
    expect(engine.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
    })

    engine.setDefinedName('Rate', 0.1)
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'upsertDefinedName', name: 'Rate', value: 0.1 }])
    outbound.splice(0)

    engine.setDefinedName('Rate', 0.1)
    expect(outbound).toEqual([])

    const sortRange = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } as const
    const sortKeys = [{ keyAddress: 'B1', direction: 'asc' as const }]
    engine.setSort('Sheet1', sortRange, sortKeys)
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'setSort', sheetName: 'Sheet1', range: sortRange, keys: sortKeys }])
    outbound.splice(0)

    engine.setSort('Sheet1', sortRange, sortKeys)
    expect(outbound).toEqual([])

    engine.setWorkbookMetadata('locale', 'en-US')
    expect(engine.getWorkbookMetadata('locale')).toEqual({ key: 'locale', value: 'en-US' })
    expect(engine.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    outbound.splice(0)

    engine.setWorkbookMetadata('locale', 'en-US')
    expect(outbound).toEqual([])

    engine.setCalculationSettings({ mode: 'automatic' })
    expect(outbound).toEqual([])

    const table = {
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    } as const
    engine.setTable(table)
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'upsertTable', table }])
    expect(engine.getTable('Sales')).toEqual(table)
    expect(engine.getTables()).toEqual([table])
    outbound.splice(0)

    engine.setTable({ ...table, columnNames: [...table.columnNames] })
    expect(outbound).toEqual([])

    expect(engine.deleteTable('MissingTable')).toBe(false)

    engine.setSpillRange('Sheet1', 'D4', 2, 3)
    expect(engine.getSpillRanges()).toEqual([{ sheetName: 'Sheet1', address: 'D4', rows: 2, cols: 3 }])
    outbound.splice(0)

    engine.setSpillRange('Sheet1', 'D4', 2, 3)
    expect(outbound).toEqual([])
    expect(engine.deleteSpillRange('Sheet1', 'Z9')).toBe(false)
    expect(engine.deletePivotTable('Sheet1', 'A1')).toBe(false)

    engine.setPivotTable('Sheet1', 'F1', {
      name: 'SalesPivot',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(engine.getPivotTable('Sheet1', 'F1')).toMatchObject({
      name: 'SalesPivot',
      sheetName: 'Sheet1',
      address: 'F1',
    })
    expect(engine.getPivotTables()).toHaveLength(1)

    outbound.splice(0)
    engine.clearCell('Sheet1', 'A1')
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }])

    outbound.splice(0)
    unsubscribeBatches()
    engine.setWorkbookMetadata('timezone', 'UTC')
    expect(outbound).toEqual([])
  })

  it('reads and clears data validations through the direct engine helpers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'data-validation-helpers' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const range = {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'A2',
    } as const

    engine.setDataValidation({
      range,
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: true,
    })

    expect(engine.getDataValidation('Sheet1', range)).toEqual({
      range,
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: true,
    })
    expect(engine.clearDataValidation('Sheet1', range)).toBe(true)
    expect(engine.getDataValidation('Sheet1', range)).toBeUndefined()
    expect(engine.clearDataValidation('Sheet1', range)).toBe(false)
  })

  it('persists workbook defined names through snapshot roundtrip', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.setDefinedName('TaxRate', 0.085)
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 100)
    engine.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')

    const snapshot = engine.exportSnapshot()

    expect(snapshot.workbook.metadata?.definedNames).toEqual([{ name: 'TaxRate', value: 0.085 }])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getDefinedNames()).toEqual([{ name: 'TaxRate', value: 0.085 }])
    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8.5 })
    expect(restored.exportSnapshot().workbook.metadata?.definedNames).toEqual([{ name: 'TaxRate', value: 0.085 }])
  })

  it('supports explicit range-ref and formula defined-name metadata values', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[1], [2], [3]])
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setDefinedName('SalesRange', {
      kind: 'range-ref',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'A3',
    })
    engine.setDefinedName('TaxExpr', { kind: 'formula', formula: '=B1*0.1' })
    engine.setCellFormula('Sheet1', 'C1', 'SUM(SalesRange)+TaxExpr')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.definedNames).toEqual([
      {
        name: 'SalesRange',
        value: {
          kind: 'range-ref',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A3',
        },
      },
      { name: 'TaxExpr', value: { kind: 'formula', formula: '=B1*0.1' } },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)
    expect(restored.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })
  })

  it('treats invalid and cyclic formula-backed defined names as workbook metadata errors', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'defined-name-errors' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setDefinedName('BrokenExpr', { kind: 'formula', formula: '=1+' })
    engine.setDefinedName('LoopA', { kind: 'formula', formula: '=LoopB' })
    engine.setDefinedName('LoopB', { kind: 'formula', formula: '=LoopA' })

    engine.setCellFormula('Sheet1', 'A1', 'BrokenExpr')
    engine.setCellFormula('Sheet1', 'A2', 'LoopA')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('replicates structural workbook metadata through authoritative op batches', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: EngineOpBatch[] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Sheet1')
    outbound.splice(0).forEach((batch) => replica.applyRemoteBatch(batch))

    primary.setWorkbookMetadata('locale', 'en-US')
    primary.updateRowMetadata('Sheet1', 2, 3, 24, false)
    primary.updateColumnMetadata('Sheet1', 1, 2, 120, true)
    primary.setFreezePane('Sheet1', 1, 2)
    primary.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' })
    primary.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' }, [{ keyAddress: 'B1', direction: 'desc' }])
    primary.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C10',
      columnNames: ['Region', 'Product', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    primary.setSpillRange('Sheet1', 'E1', 2, 3)

    expect(outbound.at(0)?.ops).toEqual([{ kind: 'setWorkbookMetadata', key: 'locale', value: 'en-US' }])
    expect(outbound.at(1)?.ops).toEqual([
      {
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        start: 2,
        count: 3,
        size: 24,
        hidden: false,
      },
    ])
    expect(outbound.at(2)?.ops).toEqual([
      {
        kind: 'updateColumnMetadata',
        sheetName: 'Sheet1',
        start: 1,
        count: 2,
        size: 120,
        hidden: true,
      },
    ])
    expect(outbound.at(3)?.ops).toEqual([{ kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 2 }])
    expect(outbound.at(4)?.ops).toEqual([
      {
        kind: 'setFilter',
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
      },
    ])
    expect(outbound.at(5)?.ops).toEqual([
      {
        kind: 'setSort',
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
        keys: [{ keyAddress: 'B1', direction: 'desc' }],
      },
    ])
    expect(outbound.at(6)?.ops).toEqual([
      {
        kind: 'upsertTable',
        table: {
          name: 'Sales',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'C10',
          columnNames: ['Region', 'Product', 'Sales'],
          headerRow: true,
          totalsRow: false,
        },
      },
    ])
    expect(outbound.at(7)?.ops).toEqual([
      {
        kind: 'upsertSpillRange',
        sheetName: 'Sheet1',
        address: 'E1',
        rows: 2,
        cols: 3,
      },
    ])

    outbound.forEach((batch) => replica.applyRemoteBatch(batch))

    expect(replica.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(replica.getRowMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 2, count: 3, size: 24, hidden: false }])
    expect(replica.getColumnMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 1, count: 2, size: 120, hidden: true }])
    expect(replica.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 2 })
    expect(replica.getFilters('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
      },
    ])
    expect(replica.getSorts('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
        keys: [{ keyAddress: 'B1', direction: 'desc' }],
      },
    ])
    expect(replica.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C10',
        columnNames: ['Region', 'Product', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
    expect(replica.getSpillRanges()).toEqual([{ sheetName: 'Sheet1', address: 'E1', rows: 2, cols: 3 }])
  })

  it('undoes and redoes structural metadata through the transaction log', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setWorkbookMetadata('locale', 'en-US')
    engine.setFreezePane('Sheet1', 1, 1)
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B5',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })

    expect(engine.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(engine.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 1 })
    expect(engine.getTables()).toHaveLength(1)

    expect(engine.undo()).toBe(true)
    expect(engine.getTables()).toEqual([])

    expect(engine.undo()).toBe(true)
    expect(engine.getFreezePane('Sheet1')).toBeUndefined()

    expect(engine.undo()).toBe(true)
    expect(engine.getWorkbookMetadataEntries()).toEqual([])

    expect(engine.redo()).toBe(true)
    expect(engine.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])

    expect(engine.redo()).toBe(true)
    expect(engine.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 1 })

    expect(engine.redo()).toBe(true)
    expect(engine.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B5',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
  })

  it('persists expanded workbook metadata through snapshot roundtrip', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setWorkbookMetadata('locale', 'en-US')
    engine.updateRowMetadata('Sheet1', 2, 2, 24, false)
    engine.updateColumnMetadata('Sheet1', 1, 1, 140, null)
    engine.setFreezePane('Sheet1', 1, 2)
    engine.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' })
    engine.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C10',
      columnNames: ['Region', 'Product', 'Sales'],
      headerRow: true,
      totalsRow: true,
    })
    engine.setSpillRange('Sheet1', 'E1', 2, 2)

    const snapshot = engine.exportSnapshot()

    expect(snapshot.workbook.metadata?.properties).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(snapshot.workbook.metadata?.tables).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C10',
        columnNames: ['Region', 'Product', 'Sales'],
        headerRow: true,
        totalsRow: true,
      },
    ])
    expect(snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'E1', rows: 2, cols: 2 }])
    expect(snapshot.sheets.find((sheet) => sheet.name === 'Sheet1')?.metadata).toEqual({
      rows: [
        { id: 'row-1', index: 2, size: 24, hidden: false },
        { id: 'row-2', index: 3, size: 24, hidden: false },
      ],
      columns: [{ id: 'column-1', index: 1, size: 140 }],
      rowMetadata: [{ start: 2, count: 2, size: 24, hidden: false }],
      columnMetadata: [{ start: 1, count: 1, size: 140 }],
      freezePane: { rows: 1, cols: 2 },
      filters: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' }],
      sorts: [
        {
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
          keys: [{ keyAddress: 'B1', direction: 'asc' }],
        },
      ],
    })

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(restored.getRowAxisEntries('Sheet1')).toEqual([
      { id: 'row-1', index: 2, size: 24, hidden: false },
      { id: 'row-2', index: 3, size: 24, hidden: false },
    ])
    expect(restored.getColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-1', index: 1, size: 140 }])
    expect(restored.getRowMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 2, count: 2, size: 24, hidden: false }])
    expect(restored.getColumnMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 1, count: 1, size: 140, hidden: null }])
    expect(restored.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 2 })
    expect(restored.getFilters('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
      },
    ])
    expect(restored.getSorts('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])
    expect(restored.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C10',
        columnNames: ['Region', 'Product', 'Sales'],
        headerRow: true,
        totalsRow: true,
      },
    ])
    expect(restored.getSpillRanges()).toEqual([{ sheetName: 'Sheet1', address: 'E1', rows: 2, cols: 2 }])
    expect(restored.exportSnapshot()).toEqual(snapshot)
  })

  it('roundtrips structurally shifted range-backed defined names through snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'snapshot-structural-defined-range' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [
      ['Qty', 'Amount'],
      [1, 10],
      [2, 20],
    ])
    engine.setDefinedName('SalesRange', {
      kind: 'range-ref',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
    })
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula('Sheet1', 'C1', 'SUM(SalesRange)')
    engine.insertColumns('Sheet1', 0, 1)

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.definedNames).toEqual([
      {
        name: 'SalesRange',
        value: {
          kind: 'range-ref',
          sheetName: 'Sheet1',
          startAddress: 'B1',
          endAddress: 'C3',
        },
      },
    ])

    const restored = new SpreadsheetEngine({
      workbookName: 'snapshot-structural-defined-range-restored',
    })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'C3',
      },
    })
    expect(restored.getCell('Sheet1', 'D1').formula).toBe('SUM(SalesRange)')
    expect(restored.getCellValue('Sheet1', 'D1')).toMatchObject({ tag: 1, value: 33 })
    expect(restored.exportSnapshot()).toEqual(snapshot)
  })

  it('treats no-op structural, metadata, freeze, and filter updates as stable public operations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.updateRowMetadata('Sheet1', 1, 1, 25, false)
    engine.updateColumnMetadata('Sheet1', 0, 1, 90, true)
    engine.setFreezePane('Sheet1', 1, 2)
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B5' } as const
    engine.setFilter('Sheet1', range)

    const before = engine.exportSnapshot()

    engine.insertRows('Sheet1', 0, 0)
    engine.deleteRows('Sheet1', 0, 0)
    engine.moveRows('Sheet1', 1, 0, 2)
    engine.moveRows('Sheet1', 1, 1, 1)
    engine.insertColumns('Sheet1', 0, 0)
    engine.deleteColumns('Sheet1', 0, 0)
    engine.moveColumns('Sheet1', 0, 0, 1)
    engine.moveColumns('Sheet1', 0, 1, 0)
    engine.updateRowMetadata('Sheet1', 1, 1, 25, false)
    engine.updateColumnMetadata('Sheet1', 0, 1, 90, true)
    engine.updateRowMetadata('Sheet1', 5, 1, null, null)
    engine.updateColumnMetadata('Sheet1', 5, 1, null, null)
    engine.setFreezePane('Sheet1', 1, 2)
    engine.setFilter('Sheet1', range)

    expect(engine.exportSnapshot()).toEqual(before)
    expect(engine.clearFreezePane('Sheet1')).toBe(true)
    expect(engine.clearFreezePane('Sheet1')).toBe(false)
    expect(engine.clearFilter('Sheet1', range)).toBe(true)
    expect(engine.clearFilter('Sheet1', range)).toBe(false)
    expect(engine.getFreezePane('Sheet1')).toBeUndefined()
    expect(engine.getFilters('Sheet1')).toEqual([])
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 20 })
  })

  it('tracks structural row identities and rewrites formulas for row inserts and moves', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)')
    engine.updateRowMetadata('Sheet1', 1, 1, 30, false)

    const before = engine.getRowAxisEntries('Sheet1')
    expect(before).toEqual([{ id: 'row-1', index: 1, size: 30, hidden: false }])

    engine.insertRows('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('SUM(A1:A3)')
    expect(engine.getRowAxisEntries('Sheet1')).toEqual([
      { id: 'row-2', index: 1 },
      { id: 'row-1', index: 2, size: 30, hidden: false },
    ])

    engine.moveRows('Sheet1', 2, 1, 0)
    expect(engine.getRowAxisEntries('Sheet1')).toEqual([
      { id: 'row-1', index: 0, size: 30, hidden: false },
      { id: 'row-2', index: 2 },
    ])
  })

  it('keeps repeated direct aggregate families correct across structural row transforms', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-aggregate-rows' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.insertRows('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('SUM(A1:A1)')
    expect(engine.getCell('Sheet1', 'B3').formula).toBe('SUM(A1:A3)')
    expect(engine.getCell('Sheet1', 'B5').formula).toBe('SUM(A1:A5)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.deleteRows('Sheet1', 1, 1)

    for (let row = 1; row <= 4; row += 1) {
      expect(engine.getCell('Sheet1', `B${row}`).formula).toBe(`SUM(A1:A${row})`)
      expect(engine.getCellValue('Sheet1', `B${row}`)).toEqual({
        tag: ValueTag.Number,
        value: (row * (row + 1)) / 2,
      })
    }
  })

  it('rewrites metadata-backed ranges, names, freeze panes, and pivot sources across structural row edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 7],
      ['East', 5],
    ])
    engine.setDefinedName('SalesRange', '=Data!A1:B4')
    engine.setFreezePane('Data', 1, 0)
    engine.setFilter('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' })
    engine.setSort('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesPivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    engine.insertRows('Data', 0, 1)

    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: '=Data!A2:B5',
    })
    expect(engine.getFreezePane('Data')).toEqual({ sheetName: 'Data', rows: 2, cols: 0 })
    expect(engine.getFilters('Data')).toEqual([{ sheetName: 'Data', range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' } }])
    expect(engine.getSorts('Data')).toEqual([
      {
        sheetName: 'Data',
        range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' },
        keys: [{ keyAddress: 'B2', direction: 'asc' }],
      },
    ])
    expect(engine.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A2',
        endAddress: 'B5',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
    expect(engine.getPivotTables()).toEqual([
      {
        name: 'SalesPivot',
        sheetName: 'Pivot',
        address: 'B2',
        source: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
        rows: 3,
        cols: 2,
      },
    ])

    engine.deleteRows('Data', 0, 5)

    expect(engine.getDefinedName('SalesRange')).toEqual({ name: 'SalesRange', value: '=#REF!' })
    expect(engine.getFreezePane('Data')).toBeUndefined()
    expect(engine.getFilters('Data')).toEqual([])
    expect(engine.getSorts('Data')).toEqual([])
    expect(engine.getTables()).toEqual([])
    expect(engine.getPivotTables()).toEqual([])
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })
  })

  it('recalculates named, table, and direct cross-sheet formulas after structural row deletes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 7],
      ['North', 5],
    ])
    engine.setDefinedName('SalesRange', '=Data!B2:B4')
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula('Summary', 'A1', 'SUM(Data!B2:B4)')
    engine.setCellFormula('Summary', 'A2', 'SUM(Sales[Sales])')
    engine.setCellFormula('Summary', 'A3', 'SUM(SalesRange)')

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 22 })

    engine.deleteRows('Data', 2, 1)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('recalculates positional cross-sheet formulas after structural row moves', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A4' }, [[10], [20], [30], [40]])
    engine.setCellFormula('Summary', 'A1', 'INDEX(Data!A1:A4,1)')
    engine.setCellFormula('Summary', 'A2', 'INDEX(Data!A1:A4,2)')
    engine.setCellFormula('Summary', 'A3', 'SUM(Data!A1:A4)')

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 100 })

    engine.moveRows('Data', 2, 1, 0)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 100 })
  })

  it('undoes structural row deletes without losing cross-sheet formula correctness', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-delete-undo-cross-sheet' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A4' }, [[10], [20], [30], [40]])
    engine.setCellFormula('Summary', 'A1', 'SUM(Data!A1:A4)')
    engine.setCellFormula('Summary', 'A2', 'INDEX(Data!A1:A4,2)')

    engine.deleteRows('Data', 1, 1)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 80 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 30 })

    expect(engine.undo()).toBe(true)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Data', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
  })

  it('rewrites formulas for structural column inserts and roundtrips calc settings metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:B1)')
    engine.setCalculationSettings({ mode: 'manual', compatibilityMode: 'odf-1.4' })

    engine.insertColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(A1:C1)')

    engine.recalculateNow()
    expect(engine.exportSnapshot().workbook.metadata?.calculationSettings).toEqual({
      mode: 'manual',
      compatibilityMode: 'odf-1.4',
    })
    expect(engine.exportSnapshot().workbook.metadata?.volatileContext?.recalcEpoch).toBeGreaterThan(0)

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())
    expect(restored.getCalculationSettings()).toEqual({
      mode: 'manual',
      compatibilityMode: 'odf-1.4',
    })
  })

  it('rewrites formulas and axis identities for structural column deletes and moves', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'E1', 'SUM(A1:B1)')
    engine.updateColumnMetadata('Sheet1', 0, 1, 90, true)

    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-1', index: 0, size: 90, hidden: true }])

    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(A1:A1)')
    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([])

    engine.updateColumnMetadata('Sheet1', 1, 1, 110, false)
    engine.setCellFormula('Sheet1', 'D2', 'B1')
    engine.moveColumns('Sheet1', 1, 1, 0)

    expect(engine.getCell('Sheet1', 'D2').formula).toBe('A1')
    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-2', index: 0, size: 110, hidden: false }])
  })

  it('keeps simple cell-reference formula families correct across structural column transforms', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-simple-columns' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    engine.setCellFormula('Sheet1', 'D1', 'C1*2')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.insertColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'D1').formula).toBe('A1+C1')
    expect(engine.getCell('Sheet1', 'E1').formula).toBe('D1*2')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.deleteColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'C1').formula).toBe('A1+B1')
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('C1*2')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.moveColumns('Sheet1', 1, 1, 0)

    expect(engine.getCell('Sheet1', 'C1').formula).toBe('B1+A1')
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('C1*2')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('keeps repeated row-shifted formula families correct across structural column transforms', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-column-families' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 2)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*2`)
    }

    const planIdsBeforeInsert = new Map<string, number>()
    const templateIdsBeforeInsert = new Map<string, number | undefined>()
    const directScalarsBeforeInsert = new Map<string, unknown>()
    for (let row = 1; row <= 4; row += 1) {
      const cIndex = engine.workbook.getCellIndex('Sheet1', `C${row}`)
      const dIndex = engine.workbook.getCellIndex('Sheet1', `D${row}`)
      planIdsBeforeInsert.set(`C${row}`, readRuntimeFormula(engine, cIndex!)!.planId)
      planIdsBeforeInsert.set(`D${row}`, readRuntimeFormula(engine, dIndex!)!.planId)
      templateIdsBeforeInsert.set(`C${row}`, readRuntimeTemplateId(engine, cIndex!))
      templateIdsBeforeInsert.set(`D${row}`, readRuntimeTemplateId(engine, dIndex!))
      directScalarsBeforeInsert.set(`C${row}`, readRuntimeDirectScalar(engine, cIndex!))
      directScalarsBeforeInsert.set(`D${row}`, readRuntimeDirectScalar(engine, dIndex!))
    }

    engine.insertColumns('Sheet1', 1, 1)

    for (let row = 1; row <= 4; row += 1) {
      expect(engine.getCell('Sheet1', `D${row}`).formula).toBe(`A${row}+C${row}`)
      expect(engine.getCell('Sheet1', `E${row}`).formula).toBe(`D${row}*2`)
      expect(engine.getCellValue('Sheet1', `D${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 3,
      })
      expect(engine.getCellValue('Sheet1', `E${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 6,
      })
      const dIndex = engine.workbook.getCellIndex('Sheet1', `D${row}`)
      const eIndex = engine.workbook.getCellIndex('Sheet1', `E${row}`)
      expect(readRuntimeFormula(engine, dIndex!)?.planId).toBe(planIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeFormula(engine, eIndex!)?.planId).toBe(planIdsBeforeInsert.get(`D${row}`))
      expect(readRuntimeTemplateId(engine, dIndex!)).toBe(templateIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeTemplateId(engine, eIndex!)).toBe(templateIdsBeforeInsert.get(`D${row}`))
      expect(readRuntimeDirectScalar(engine, dIndex!)).toBe(directScalarsBeforeInsert.get(`C${row}`))
      expect(readRuntimeDirectScalar(engine, eIndex!)).toBe(directScalarsBeforeInsert.get(`D${row}`))
    }

    engine.deleteColumns('Sheet1', 1, 1)

    for (let row = 1; row <= 4; row += 1) {
      expect(engine.getCell('Sheet1', `C${row}`).formula).toBe(`A${row}+B${row}`)
      expect(engine.getCell('Sheet1', `D${row}`).formula).toBe(`C${row}*2`)
      expect(engine.getCellValue('Sheet1', `C${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 3,
      })
      expect(engine.getCellValue('Sheet1', `D${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 6,
      })
      const cIndex = engine.workbook.getCellIndex('Sheet1', `C${row}`)
      const dIndex = engine.workbook.getCellIndex('Sheet1', `D${row}`)
      expect(readRuntimeFormula(engine, cIndex!)?.planId).toBe(planIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeFormula(engine, dIndex!)?.planId).toBe(planIdsBeforeInsert.get(`D${row}`))
      expect(readRuntimeTemplateId(engine, cIndex!)).toBe(templateIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeTemplateId(engine, dIndex!)).toBe(templateIdsBeforeInsert.get(`D${row}`))
    }
  })

  it('routes multi-name scalar formulas through the wasm path once names exist', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'TaxRate+FeeRate')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    engine.setDefinedName('TaxRate', 0.085)
    engine.setDefinedName('FeeRate', 0.015)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.1 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('resolves named range formulas through workbook metadata and rebinds dependencies', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 12)
    engine.setCellValue('Sheet1', 'A3', 15)
    engine.setDefinedName('SalesRange', '=Sheet1!A1:A3')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(SalesRange)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 37 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A2', 20)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 45 })

    engine.setDefinedName('SalesRange', '=Sheet1!A1:A2')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 30 })
  })

  it('binds structured table references through table metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Region')
    engine.setCellValue('Sheet1', 'B1', 'Amount')
    engine.setCellValue('Sheet1', 'A2', 'North')
    engine.setCellValue('Sheet1', 'B2', 10)
    engine.setCellValue('Sheet1', 'A3', 'South')
    engine.setCellValue('Sheet1', 'B3', 12)
    engine.setCellValue('Sheet1', 'A4', 'West')
    engine.setCellValue('Sheet1', 'B4', 15)
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })

    engine.setCellFormula('Sheet1', 'C1', 'SUM(Sales[Amount])')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 37 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'B3', 20)
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 45 })

    engine.setCellFormula('Sheet1', 'D1', 'Sales[Amount]')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('rebinds spill-shape formulas when owner ranges appear and resize', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1#)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })

    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(2,1,1,1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
  })

  it('materializes pivot tables, refreshes aggregates, and roundtrips snapshot metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' }, [
      ['Region', 'Notes', 'Product', 'Sales'],
      ['East', 'priority', 'Widget', 10],
      ['West', 'priority', 'Widget', 7],
      ['East', 'priority', 'Gizmo', 5],
    ])

    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.getCellValue('Pivot', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'Region',
    })
    expect(engine.getCellValue('Pivot', 'C2')).toMatchObject({
      tag: ValueTag.String,
      value: 'SUM of Sales',
    })
    expect(engine.getCellValue('Pivot', 'D2')).toMatchObject({
      tag: ValueTag.String,
      value: 'Rows',
    })
    expect(engine.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Pivot', 'D3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Pivot', 'B4')).toMatchObject({
      tag: ValueTag.String,
      value: 'West',
    })
    expect(engine.getCellValue('Pivot', 'C4')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Pivot', 'D4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getPivotTables()).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'B2',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
        groupBy: ['Region'],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum' },
          { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
        ],
        rows: 3,
        cols: 3,
      },
    ])

    engine.setCellValue('Data', 'D3', 9)

    expect(engine.getCellValue('Pivot', 'C4')).toEqual({ tag: ValueTag.Number, value: 9 })

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.pivots).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'B2',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
        groupBy: ['Region'],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum' },
          { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
        ],
        rows: 3,
        cols: 3,
      },
    ])
    expect(snapshot.sheets.find((sheet) => sheet.name === 'Pivot')?.cells).toEqual([])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(restored.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(restored.getCellValue('Pivot', 'C4')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(restored.exportSnapshot().workbook.metadata?.pivots).toEqual(snapshot.workbook.metadata?.pivots)
  })

  it('evaluates GETPIVOTDATA against workbook pivot metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' }, [
      ['Region', 'Notes', 'Product', 'Sales'],
      ['East', 'priority', 'Widget', 10],
      ['West', 'priority', 'Widget', 7],
      ['East', 'priority', 'Gizmo', 5],
    ])

    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    engine.setCellFormula('Sheet1', 'A1', 'GETPIVOTDATA("Sales",Pivot!B2)')
    engine.setCellFormula('Sheet1', 'A2', 'GETPIVOTDATA("Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Sheet1', 'A3', 'GETPIVOTDATA("Rows",Pivot!B2,"Region","West")')
    engine.setCellFormula('Sheet1', 'A4', 'GETPIVOTDATA("Missing",Pivot!B2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'A4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.explainCell('Sheet1', 'A2').mode).toBe(FormulaMode.JsOnly)
  })

  it('evaluates GROUPBY, PIVOTBY, and MULTIPLE.OPERATIONS end to end', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D5' }, [
      ['Region', 'Product', 'Sales', 'Include'],
      ['East', 'Widget', 10, true],
      ['West', 'Widget', 7, true],
      ['East', 'Gizmo', 5, true],
      ['West', 'Gizmo', 4, true],
    ])
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P5', 'P2*P3+P4')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellValue('Sheet1', 'R2', 3)

    engine.setCellFormula('Sheet1', 'F1', 'GROUPBY(A1:A5,C1:C5,SUM,3,1)')
    engine.setCellFormula('Sheet1', 'J1', 'PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)')
    engine.setCellFormula('Sheet1', 'N1', 'MULTIPLE.OPERATIONS(P5,P3,Q4,P2,R2)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.String,
      value: 'Region',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({
      tag: ValueTag.String,
      value: 'Sales',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.String,
      value: 'East',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'G4')).toEqual({ tag: ValueTag.Number, value: 26 })

    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({
      tag: ValueTag.String,
      value: 'Region',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({
      tag: ValueTag.String,
      value: 'Widget',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'J2')).toEqual({
      tag: ValueTag.String,
      value: 'East',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'K2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'M4')).toEqual({ tag: ValueTag.Number, value: 26 })

    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({ tag: ValueTag.Number, value: 23 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'N1').mode).toBe(FormulaMode.JsOnly)
  })

  it('evaluates row-only MULTIPLE.OPERATIONS substitutions through the JS workbook path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P5', 'P2*P3+P4')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P5,P3,Q4)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.JsOnly)
  })

  it('returns literal, empty, and cycle states for MULTIPLE.OPERATIONS target cells', async () => {
    const literalEngine = new SpreadsheetEngine({ workbookName: 'multiple-ops-literal' })
    await literalEngine.ready()
    literalEngine.setCellValue('Sheet1', 'B5', 42)
    literalEngine.setCellValue('Sheet1', 'C4', 5)
    literalEngine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(B5,B3,C4)')
    expect(literalEngine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })

    const missingEngine = new SpreadsheetEngine({ workbookName: 'multiple-ops-missing' })
    await missingEngine.ready()
    missingEngine.setCellValue('Sheet1', 'C4', 5)
    missingEngine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(Z9,B3,C4)')
    expect(missingEngine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    const cycleEngine = new SpreadsheetEngine({ workbookName: 'multiple-ops-cycle' })
    await cycleEngine.ready()
    cycleEngine.setCellFormula('Sheet1', 'P2', 'P3')
    cycleEngine.setCellFormula('Sheet1', 'P3', 'P2')
    cycleEngine.setCellValue('Sheet1', 'Q4', 5)
    cycleEngine.setCellValue('Sheet1', 'R4', 9)
    cycleEngine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P2,Q4,R4)')
    expect(cycleEngine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('applies MULTIPLE.OPERATIONS replacements through ranged formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'multiple-ops-range' })
    await engine.ready()
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P6', 'SUM(P2:P4)')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellValue('Sheet1', 'R2', 3)
    engine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P6,P3,Q4,P2,R2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: 16,
    })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.JsOnly)
  })

  it('evaluates nested MULTIPLE.OPERATIONS formulas through the workbook callback path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'multiple-ops-nested' })
    await engine.ready()
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P5', 'P2*P3+P4')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellValue('Sheet1', 'R2', 3)
    engine.setCellFormula('Sheet1', 'P7', 'MULTIPLE.OPERATIONS(P5,P3,Q4)')
    engine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P7,P2,R2)')

    expect(engine.getCellValue('Sheet1', 'P7')).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.JsOnly)
  })

  it('undoes pivot deletion through the transaction log', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D3' }, [
      ['Region', 'Notes', 'Product', 'Sales'],
      ['East', 'priority', 'Widget', 10],
      ['West', 'priority', 'Widget', 7],
    ])

    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(engine.deletePivotTable('Pivot', 'B2')).toBe(true)
    expect(engine.getPivotTables()).toEqual([])
    expect(engine.getCellValue('Pivot', 'B3')).toEqual({ tag: ValueTag.Empty })

    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTables()).toHaveLength(1)
    expect(engine.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
  })

  it('returns #VALUE for pivots whose configured headers are missing', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 5],
    ])

    engine.setPivotTable('Pivot', 'A1', {
      name: 'BrokenPivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Missing'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getCellValue('Pivot', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('returns #REF for missing pivot source sheets and rebinds once source cells appear', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Pivot')

    engine.setPivotTable('Pivot', 'A1', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getCellValue('Pivot', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 5],
    ])

    expect(engine.getCellValue('Pivot', 'A1')).toMatchObject({
      tag: ValueTag.String,
      value: 'Region',
    })
    expect(engine.getCellValue('Pivot', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'SUM of Sales',
    })
    expect(engine.getCellValue('Pivot', 'A2')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Pivot', 'A3')).toMatchObject({
      tag: ValueTag.String,
      value: 'West',
    })
    expect(engine.getCellValue('Pivot', 'B3')).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('blocks overlapping pivot outputs and deletes pivots when users overwrite pivot cells', async () => {
    const blockedByValue = new SpreadsheetEngine({ workbookName: 'pivot-blocked-value' })
    await blockedByValue.ready()
    seedPivotSource(blockedByValue)
    blockedByValue.setCellValue('Pivot', 'C3', 99)
    blockedByValue.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedByValue.getCellValue('Pivot', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedByValue.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 99 })

    const blockedByFormula = new SpreadsheetEngine({ workbookName: 'pivot-blocked-formula' })
    await blockedByFormula.ready()
    seedPivotSource(blockedByFormula)
    blockedByFormula.setCellFormula('Pivot', 'C3', '1+1')
    blockedByFormula.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedByFormula.getCellValue('Pivot', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedByFormula.getCellValue('Pivot', 'C3')).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })

    const blockedBySpillChild = new SpreadsheetEngine({ workbookName: 'pivot-blocked-spill' })
    await blockedBySpillChild.ready()
    seedPivotSource(blockedBySpillChild)
    blockedBySpillChild.setCellFormula('Pivot', 'C2', 'SEQUENCE(2,1,1,1)')
    blockedBySpillChild.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedBySpillChild.getCellValue('Pivot', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedBySpillChild.getCellValue('Pivot', 'C3')).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })

    const blockedByPivotOwner = new SpreadsheetEngine({ workbookName: 'pivot-blocked-pivot' })
    await blockedByPivotOwner.ready()
    seedPivotSource(blockedByPivotOwner)
    blockedByPivotOwner.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    blockedByPivotOwner.setPivotTable('Pivot', 'A1', {
      name: 'Overlap',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedByPivotOwner.getCellValue('Pivot', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedByPivotOwner.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })

    const overwrittenPivot = new SpreadsheetEngine({ workbookName: 'pivot-overwrite' })
    await overwrittenPivot.ready()
    seedPivotSource(overwrittenPivot)
    overwrittenPivot.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(overwrittenPivot.getPivotTables()).toHaveLength(1)

    overwrittenPivot.setCellValue('Pivot', 'B3', 'manual')

    expect(overwrittenPivot.getPivotTables()).toEqual([])
    expect(overwrittenPivot.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(overwrittenPivot.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'manual',
    })
    expect(overwrittenPivot.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Empty })
  })

  it('explains missing cells and undoes table spill and pivot metadata changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')

    expect(engine.explainCell('Data', 'Z99')).toEqual({
      sheetName: 'Data',
      address: 'Z99',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
      inCycle: false,
      directPrecedents: [],
      directDependents: [],
    })

    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 5],
    ])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setSpillRange('Pivot', 'E1', 2, 2)
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: true,
    })
    engine.setSpillRange('Pivot', 'E1', 3, 1)
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Sales', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.getTable('Sales')).toMatchObject({ totalsRow: true })
    expect(engine.getSpillRanges()).toContainEqual({
      sheetName: 'Pivot',
      address: 'E1',
      rows: 3,
      cols: 1,
    })
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Sales', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.deleteTable('Sales')).toBe(true)
    expect(engine.deleteSpillRange('Pivot', 'E1')).toBe(true)
    expect(engine.deletePivotTable('Pivot', 'B2')).toBe(true)

    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Sales', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getSpillRanges()).toContainEqual({
      sheetName: 'Pivot',
      address: 'E1',
      rows: 3,
      cols: 1,
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getTable('Sales')).toMatchObject({ totalsRow: true })

    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getSpillRanges()).toContainEqual({
      sheetName: 'Pivot',
      address: 'E1',
      rows: 2,
      cols: 2,
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getTable('Sales')).toMatchObject({ totalsRow: false })
  })

  it('exports sparse high-row cells without truncating the sheet', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A10002', 7)

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets[0]?.cells).toContainEqual({ address: 'A10002', value: 7 })
  })

  it('roundtrips a single sheet through CSV with formulas and quoted strings', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Sheet1', 'A2', 'alpha,beta')

    const csv = engine.exportSheetCsv('Sheet1')
    expect(csv).toBe('12,=A1*2\n"alpha,beta",')

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSheetCsv('Sheet1', csv)

    expect(restored.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(restored.getCell('Sheet1', 'B1').formula).toBe('A1*2')
    expect(restored.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(restored.getCell('Sheet1', 'A2').value).toEqual({
      tag: ValueTag.String,
      value: 'alpha,beta',
      stringId: 1,
    })
  })

  it('recalculates range formulas over imported formula cells after CSV import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-range-recalc' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'SUM(B1:B1)')
    engine.setCellValue('Sheet1', 'B1', 'text:@4yt')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(B2:C4)')
    engine.setCellValue('Sheet1', 'A2', false)
    engine.setCellValue('Sheet1', 'C2', true)
    engine.setCellFormula('Sheet1', 'A3', 'SUM(B1:B1)')
    engine.setCellFormula('Sheet1', 'C3', 'B1+B1')
    engine.setCellFormula('Sheet1', 'A4', 'SUM(B1:B1)')
    engine.setCellValue('Sheet1', 'B4', 'text:"k')
    engine.setCellValue('Sheet1', 'C4', 'text:&Pr!}${')

    const csv = engine.exportSheetCsv('Sheet1')

    const restored = new SpreadsheetEngine({ workbookName: 'csv-range-recalc-restored' })
    await restored.ready()
    restored.importSheetCsv('Sheet1', csv)

    expect(restored.getCellValue('Sheet1', 'C3')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(restored.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('recalculates transitive range dependents when a downstream formula becomes an error', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'transitive-range-recalc' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'C1', 1_835_115_565)
    engine.setCellValue('Sheet1', 'D1', -24)
    engine.setCellFormula('Sheet1', 'A2', 'SUM(B1:E2)')
    engine.setCellFormula('Sheet1', 'B2', 'IF(E2>0,"text:yes","text:no")')

    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Number,
      value: 1_835_115_541,
    })

    engine.setCellValue('Sheet1', 'E2', 'text:ooe)ZL#<')

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('persists cell formats through imperative updates and snapshot roundtrip', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12)
    engine.setCellFormat('Sheet1', 'A1', 'currency-usd')

    expect(engine.getCell('Sheet1', 'A1').format).toBe('currency-usd')
    expect(engine.explainCell('Sheet1', 'A1').format).toBe('currency-usd')

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A1').format).toBe('currency-usd')
    expect(restored.exportSnapshot().sheets[0]?.cells).toContainEqual({
      address: 'A1',
      value: 12,
      format: 'currency-usd',
    })
  })

  it('includes format-only mutations in changed cell events', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12)

    const changed: Array<{ indices: number[]; cells: EngineEvent['changedCells'] }> = []
    const unsubscribe = engine.subscribe((event) => {
      changed.push({
        indices: Array.from(event.changedCellIndices),
        cells: event.changedCells,
      })
    })

    engine.setCellFormat('Sheet1', 'A1', 'currency-usd')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    expect(a1Index).toBeDefined()
    expect(sheetId).toBeDefined()
    expect(changed.at(-1)?.indices).toEqual([a1Index!])
    expect(changed.at(-1)?.cells).toEqual([
      {
        kind: 'cell',
        cellIndex: a1Index!,
        address: { sheet: sheetId!, row: 0, col: 0 },
        sheetName: 'Sheet1',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 12 },
      },
    ])

    unsubscribe()
  })

  it('persists pooled cell styles and style ranges without materializing empty cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
      {
        fill: { backgroundColor: '#ABCDEF' },
        font: { family: 'Fira Sans' },
      },
    )

    expect(engine.workbook.getCellIndex('Sheet1', 'B2')).toBeUndefined()
    const styled = engine.getCell('Sheet1', 'B2')
    expect(styled.styleId).toBeDefined()

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets[0]?.cells).toEqual([])
    expect(snapshot.workbook.metadata?.styles).toHaveLength(1)
    expect(snapshot.workbook.metadata?.styles?.[0]).toMatchObject({
      fill: { backgroundColor: '#abcdef' },
      font: { family: 'Fira Sans' },
    })
    expect(snapshot.sheets[0]?.metadata?.styleRanges).toEqual([
      {
        range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
        styleId: styled.styleId,
      },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getCell('Sheet1', 'C3').styleId).toBe(styled.styleId)
    expect(restored.getCellStyle(styled.styleId)).toMatchObject({
      fill: { backgroundColor: '#abcdef' },
      font: { family: 'Fira Sans' },
    })
  })

  it('emits full invalidation when importing a snapshot', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()

    const events: Array<{
      invalidation: 'cells' | 'full'
      changedCellIndices: number[]
      changedCells: number
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        changedCellIndices: Array.from(event.changedCellIndices),
        changedCells: event.changedCells.length,
      })
    })

    engine.importSnapshot({
      version: 1,
      workbook: { name: 'spec' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [{ address: 'A1', value: 12 }],
        },
      ],
    })

    expect(events.at(-1)).toEqual({
      invalidation: 'full',
      changedCellIndices: [],
      changedCells: 0,
    })
    unsubscribe()
  })

  it('emits targeted range invalidation for style-only edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const events: Array<{
      invalidation: 'cells' | 'full'
      invalidatedRanges: readonly { sheetName: string; startAddress: string; endAddress: string }[]
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        invalidatedRanges: event.invalidatedRanges,
      })
    })

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }, { fill: { backgroundColor: '#ABCDEF' } })

    expect(events.at(-1)).toEqual({
      invalidation: 'cells',
      invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }],
    })
    unsubscribe()
  })

  it('interns identical cell styles across ranges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const patch = {
      fill: { backgroundColor: '#ff0000' },
      font: { family: 'IBM Plex Sans' },
    } as const
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, patch)
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' }, patch)

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.styles).toHaveLength(1)
    expect(snapshot.sheets[0]?.metadata?.styleRanges).toHaveLength(2)
    expect(snapshot.sheets[0]?.metadata?.styleRanges?.[0]?.styleId).toBe(snapshot.sheets[0]?.metadata?.styleRanges?.[1]?.styleId)
  })

  it('merges and clears style fields independently', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      {
        fill: { backgroundColor: '#ff0000' },
        font: { family: 'Inter' },
      },
    )
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, { font: { family: 'IBM Plex Sans' } })

    const mergedStyle = engine.getCellStyle(engine.getCell('Sheet1', 'A1').styleId)
    expect(mergedStyle).toMatchObject({
      fill: { backgroundColor: '#ff0000' },
      font: { family: 'IBM Plex Sans' },
    })

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, ['fontFamily'])
    const clearedStyle = engine.getCellStyle(engine.getCell('Sheet1', 'A1').styleId)
    expect(clearedStyle).toMatchObject({
      fill: { backgroundColor: '#ff0000' },
    })
    expect(clearedStyle?.font).toBeUndefined()
  })

  it('notifies address listeners for style-only edits on empty cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    let notified = 0
    const unsubscribe = engine.subscribeCell('Sheet1', 'D4', () => {
      notified += 1
    })

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'D4', endAddress: 'D4' }, { fill: { backgroundColor: '#00ff00' } })

    expect(notified).toBe(1)
    unsubscribe()
  })

  it('persists pooled number formats and sparse format ranges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeNumberFormat(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B4' },
      { kind: 'accounting', currency: 'USD', decimals: 2, useGrouping: true },
    )

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.formats).toHaveLength(1)
    expect(snapshot.sheets[0]?.metadata?.formatRanges).toHaveLength(1)
    expect(engine.getCell('Sheet1', 'B3').format).toContain('accounting:USD:2')
  })

  it('clears number formats, clears sorts, and tracks existing watched cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)

    let notifications = 0
    const unsubscribe = engine.subscribeCells('Sheet1', ['A1', 'Z9'], () => {
      notifications += 1
    })

    engine.setCellValue('Sheet1', 'A1', 6)
    expect(notifications).toBe(1)
    unsubscribe()

    engine.setRangeNumberFormat(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' },
      { kind: 'currency', currency: 'USD', decimals: 2, useGrouping: true },
    )
    engine.clearRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' })
    expect(engine.getCell('Sheet1', 'B2').format).toBeUndefined()

    const sortRange = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } as const
    engine.setSort('Sheet1', sortRange, [{ keyAddress: 'A1', direction: 'asc' }])
    expect(engine.clearSort('Sheet1', sortRange)).toBe(true)
    expect(engine.getSorts('Sheet1')).toEqual([])
    expect(engine.getVolatileContext()).toEqual({ recalcEpoch: 0 })
  })

  it('emits targeted axis invalidation for column metadata edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const events: Array<{
      invalidation: 'cells' | 'full'
      invalidatedColumns: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        invalidatedColumns: event.invalidatedColumns,
      })
    })

    engine.updateColumnMetadata('Sheet1', 2, 2, 120, true)

    expect(events.at(-1)).toEqual({
      invalidation: 'cells',
      invalidatedColumns: [{ sheetName: 'Sheet1', startIndex: 2, endIndex: 3 }],
    })
    unsubscribe()
  })

  it('emits structural row invalidation without flooding changed cells for row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)')

    const events: Array<{
      invalidation: 'cells' | 'full'
      changedCells: number
      invalidatedRows: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        changedCells: event.changedCells.length,
        invalidatedRows: event.invalidatedRows,
      })
    })

    engine.insertRows('Sheet1', 1, 1)

    expect(events.at(-1)).toEqual({
      invalidation: 'cells',
      changedCells: 0,
      invalidatedRows: [{ sheetName: 'Sheet1', startIndex: 1, endIndex: 1 }],
    })
    unsubscribe()
  })

  it('merges advanced style fields including borders and font weight', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'C5' },
      {
        font: { bold: true, color: '#111827', size: 14 },
        alignment: { horizontal: 'right', wrap: true },
        borders: {
          bottom: { style: 'double', weight: 'medium', color: '#111827' },
        },
      },
    )

    const style = engine.getCellStyle(engine.getCell('Sheet1', 'C5').styleId)
    expect(style).toMatchObject({
      font: { bold: true, color: '#111827', size: 14 },
      alignment: { horizontal: 'right', wrap: true },
      borders: {
        bottom: { style: 'double', weight: 'medium', color: '#111827' },
      },
    })
  })

  it('removes style subfields through null patches and clearing all style fields', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D6', endAddress: 'D6' },
      {
        font: { family: 'Inter', bold: true },
        alignment: { horizontal: 'center', wrap: true, indent: 2 },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#111111' },
          right: { style: 'double', weight: 'medium', color: '#222222' },
        },
      },
    )
    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D6', endAddress: 'D6' },
      {
        alignment: { horizontal: null, wrap: null },
        borders: {
          top: null,
          right: { style: 'solid', weight: 'thin', color: null },
        },
      },
    )

    const partiallyCleared = engine.getCellStyle(engine.getCell('Sheet1', 'D6').styleId)
    expect(partiallyCleared).toMatchObject({
      font: { family: 'Inter', bold: true },
      alignment: { indent: 2 },
    })
    expect(partiallyCleared?.borders).toBeUndefined()

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'D6', endAddress: 'D6' })
    expect(engine.getCell('Sheet1', 'D6').styleId).toBeUndefined()
  })

  it('preserves sibling style fields when clearing only part of a section', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'E7', endAddress: 'E7' },
      {
        font: { family: 'Inter', bold: true },
        alignment: { horizontal: 'right', wrap: true },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#111111' },
          left: { style: 'double', weight: 'medium', color: '#222222' },
        },
      },
    )

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'E7', endAddress: 'E7' }, ['fontBold', 'alignmentWrap', 'borderTop'])

    const style = engine.getCellStyle(engine.getCell('Sheet1', 'E7').styleId)
    expect(style).toMatchObject({
      font: { family: 'Inter' },
      alignment: { horizontal: 'right' },
      borders: {
        left: { style: 'double', weight: 'medium', color: '#222222' },
      },
    })
  })

  it('replaces existing sheet contents on CSV import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'C3', 9)

    engine.importSheetCsv('Sheet1', '7,8')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Empty })
  })

  it('explains formula cells with mode, version, and dependencies', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const explanation = engine.explainCell('Sheet1', 'B1')

    expect(explanation.formula).toBe('A1*2')
    expect(explanation.mode).toBeDefined()
    expect(explanation.version).toBeGreaterThan(0)
    expect(explanation.directPrecedents).toEqual(['Sheet1!A1'])
    expect(explanation.directDependents).toEqual([])
    expect(explanation.inCycle).toBe(false)
  })

  it('stores runtime formula slot and compiled plan metadata separately', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'C3', 'A1*2')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'C3')
    expect(cellIndex).toBeDefined()

    const formulaId = engine.workbook.cellStore.formulaIds[cellIndex!]
    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(formulaId).toBeGreaterThan(0)
    expect(isRuntimeFormulaWithCompiled(runtimeFormula)).toBe(true)
    expect(runtimeFormula).toBeDefined()
    expect(runtimeFormula?.formulaSlotId).toBe(formulaId)
    expect(runtimeFormula?.planId).toBe(runtimeFormula?.plan.id)
    expect(runtimeFormula?.compiled).toBe(runtimeFormula?.plan.compiled)
    expect(runtimeFormula?.dependencyEntities.ptr).toBeGreaterThanOrEqual(0)
    expect(runtimeFormula?.runtimeProgram.length).toBeGreaterThan(0)
  })

  it('reuses one compiled plan for identical formula sources while keeping distinct slots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'shared-plan-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', '1+2')
    engine.setCellFormula('Sheet1', 'B1', '1+2')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()

    const leftFormula = readRuntimeFormula(engine, a1Index!)
    const rightFormula = readRuntimeFormula(engine, b1Index!)

    expect(isRuntimeFormulaWithCompiled(leftFormula)).toBe(true)
    expect(isRuntimeFormulaWithCompiled(rightFormula)).toBe(true)
    expect(leftFormula?.formulaSlotId).not.toBe(rightFormula?.formulaSlotId)
    expect(leftFormula?.planId).toBe(rightFormula?.planId)
    expect(leftFormula?.compiled).toBe(rightFormula?.compiled)
  })

  it('replaces direct lookup range dependencies with lookup-column subscribers and formula-cell deps only', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'lookup-subscriber-spec',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 30)
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellFormula('Sheet1', 'E1', 'XMATCH(D1,A1:A3,0)')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    expect(cellIndex).toBeDefined()
    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(isRuntimeFormulaWithRanges(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.rangeDependencies).toHaveLength(0)
    expect(runtimeFormula?.dependencyIndices).toEqual(Uint32Array.of(engine.workbook.getCellIndex('Sheet1', 'D1')!))
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getDependencies('Sheet1', 'D1').directDependents).toContain('Sheet1!E1')
    expect(engine.getDependencies('Sheet1', 'A2').directDependents).toEqual([])

    engine.setCellValue('Sheet1', 'A2', 25)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
  })

  it('patches runtime cell and range operands from packed symbolic binding buffers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'symbolic-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)+C1')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    const c1Index = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(cellIndex).toBeDefined()
    expect(c1Index).toBeDefined()

    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(isRuntimeFormulaWithRanges(runtimeFormula)).toBe(true)
    expect(runtimeFormula).toBeDefined()
    const pushCellOpcode = Number(Opcode.PushCell)
    const pushRangeOpcode = Number(Opcode.PushRange)
    const pushCell = runtimeFormula?.runtimeProgram.find((instruction) => instruction >>> 24 === pushCellOpcode)
    const pushRange = runtimeFormula?.runtimeProgram.find((instruction) => instruction >>> 24 === pushRangeOpcode)

    expect(pushCell).toBeDefined()
    expect(pushRange).toBeDefined()
    expect(runtimeFormula?.dependencyIndices).toEqual(Uint32Array.of(c1Index!))
    expect(pushCell! & 0x00ff_ffff).toBe(c1Index)
    expect(pushRange! & 0x00ff_ffff).toBe(runtimeFormula?.rangeDependencies[0])
  })

  it('keeps packed range entity ids stable across structural inserts when the range survives', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'symbolic-structural-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)+C1')

    const beforeCellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    expect(beforeCellIndex).toBeDefined()
    const beforeRuntimeFormula = readRuntimeFormula(engine, beforeCellIndex!)
    expect(isRuntimeFormulaWithRanges(beforeRuntimeFormula)).toBe(true)

    const beforeRangeIndex = beforeRuntimeFormula?.rangeDependencies[0]
    expect(beforeRangeIndex).toBeDefined()

    engine.insertRows('Sheet1', 0, 1)

    const afterCellIndex = engine.workbook.getCellIndex('Sheet1', 'D2')
    expect(afterCellIndex).toBeDefined()
    const afterRuntimeFormula = readRuntimeFormula(engine, afterCellIndex!)
    expect(isRuntimeFormulaWithRanges(afterRuntimeFormula)).toBe(true)
    expect(afterRuntimeFormula?.rangeDependencies[0]).toBe(beforeRangeIndex)
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('keeps packed range entity ids and plan ids stable across structural row deletes when the range survives', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'symbolic-structural-delete-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 10)
      engine.setCellValue('Sheet1', `C${row}`, row * 100)
    }
    engine.setCellFormula('Sheet1', 'D5', 'SUM(A1:B4)+C4')

    const beforeCellIndex = engine.workbook.getCellIndex('Sheet1', 'D5')
    expect(beforeCellIndex).toBeDefined()
    const beforeRuntimeFormula = readRuntimeFormula(engine, beforeCellIndex!)
    expect(isRuntimeFormulaWithRanges(beforeRuntimeFormula)).toBe(true)
    expect(isRuntimeFormulaWithCompiled(beforeRuntimeFormula)).toBe(true)

    const beforePlanId = beforeRuntimeFormula?.planId
    const beforeRangeIndex = beforeRuntimeFormula?.rangeDependencies[0]
    expect(beforePlanId).toBeDefined()
    expect(beforeRangeIndex).toBeDefined()

    engine.deleteRows('Sheet1', 1, 1)

    const afterCellIndex = engine.workbook.getCellIndex('Sheet1', 'D4')
    expect(afterCellIndex).toBeDefined()
    const afterRuntimeFormula = readRuntimeFormula(engine, afterCellIndex!)
    expect(isRuntimeFormulaWithRanges(afterRuntimeFormula)).toBe(true)
    expect(isRuntimeFormulaWithCompiled(afterRuntimeFormula)).toBe(true)
    expect(afterRuntimeFormula?.planId).toBe(beforePlanId)
    expect(afterRuntimeFormula?.rangeDependencies[0]).toBe(beforeRangeIndex)
    expect(engine.getCell('Sheet1', 'D4').formula).toBe('SUM(A1:B3)+C3')
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 488 })
  })

  it('tracks literal-backed ranges through range entities without inflating topo dependency cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'range-topology-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(cellIndex).toBeDefined()
    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(isRuntimeFormulaWithRanges(runtimeFormula)).toBe(true)
    expect(isRuntimeFormulaWithDirectAggregate(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.dependencyIndices).toEqual(new Uint32Array())
    expect(runtimeFormula?.rangeDependencies).toHaveLength(0)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })

    engine.setCellValue('Sheet1', 'A2', 4)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
  })

  it('assigns deterministic cycle group ids for cyclic formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cycle-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'B1+1')
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')

    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBeGreaterThanOrEqual(0)
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBe(engine.workbook.cellStore.cycleGroupIds[b1Index!])
  })

  it('assigns topo ranks through range-node dependents deterministically', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'topo-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)')

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const d1Index = engine.workbook.getCellIndex('Sheet1', 'D1')

    expect(b1Index).toBeDefined()
    expect(d1Index).toBeDefined()
    expect(engine.workbook.cellStore.topoRanks[b1Index!]).toBeLessThan(engine.workbook.cellStore.topoRanks[d1Index!])
  })

  it('notifies per-cell listeners only for the cells that changed', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    let a1Notifications = 0
    let b1Notifications = 0
    const unsubscribeA1 = engine.subscribeCell('Sheet1', 'A1', () => {
      a1Notifications += 1
    })
    const unsubscribeB1 = engine.subscribeCell('Sheet1', 'B1', () => {
      b1Notifications += 1
    })

    engine.setCellValue('Sheet1', 'A1', 1)
    expect(a1Notifications).toBe(1)
    expect(b1Notifications).toBe(0)

    engine.setCellValue('Sheet1', 'B1', 2)
    expect(a1Notifications).toBe(1)
    expect(b1Notifications).toBe(1)

    engine.setCellValue('Sheet1', 'C1', 3)
    expect(a1Notifications).toBe(1)
    expect(b1Notifications).toBe(1)

    unsubscribeA1()
    unsubscribeB1()
  })

  it('notifies grouped watched cells only when one of them changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    let notifications = 0
    const unsubscribe = engine.subscribeCells('Sheet1', ['A1', 'A2'], () => {
      notifications += 1
    })

    engine.setCellValue('Sheet1', 'B1', 5)
    expect(notifications).toBe(0)

    engine.setCellValue('Sheet1', 'A2', 8)
    expect(notifications).toBe(1)

    unsubscribe()
  })

  it('notifies watched cells when sheet deletion clears them', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)

    let notifications = 0
    const unsubscribe = engine.subscribeCell('Sheet1', 'A1', () => {
      notifications += 1
    })

    engine.deleteSheet('Sheet1')

    expect(notifications).toBe(1)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    unsubscribe()
  })

  it('tracks selection state inside the engine and notifies subscribers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()

    const seen: string[] = []
    const unsubscribe = engine.subscribeSelection(() => {
      const snapshot = engine.getSelectionState()
      seen.push(`${snapshot.sheetName}!${snapshot.address ?? 'null'}`)
    })

    engine.setSelection('Sheet2', 'B3')
    engine.setSelection('Sheet2', 'B3')
    engine.setSelection('Sheet1', 'A1')

    expect(engine.getSelectionState()).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
      anchorAddress: 'A1',
      range: { startAddress: 'A1', endAddress: 'A1' },
      editMode: 'idle',
    })
    expect(seen).toEqual(['Sheet2!B3', 'Sheet1!A1'])

    unsubscribe()
  })

  it('restores snapshots through transactions without emitting batches or undo history', async () => {
    const source = new SpreadsheetEngine({ workbookName: 'source' })
    await source.ready()
    source.createSheet('Sheet1')
    source.setCellValue('Sheet1', 'A1', 100)
    source.setDefinedName('TaxRate', 0.1)
    source.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    const outbound: EngineOpBatch[] = []
    restored.subscribeBatches((batch) => outbound.push(batch))

    restored.importSnapshot(source.exportSnapshot())

    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(restored.getDefinedNames()).toEqual([{ name: 'TaxRate', value: 0.1 }])
    expect(outbound).toEqual([])
    expect(restored.undo()).toBe(false)
  })

  it('supports range mutation helpers and undo/redo over the same local apply path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [
      [1, 2],
      [3, 4],
    ])
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 4 })

    engine.setRangeFormulas({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' }, [['SUM(A1:B1)'], ['SUM(A2:B2)']])
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 7 })

    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'B2' })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0 })

    engine.undo()
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 7 })

    engine.redo()
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0 })
  })

  it('captures undo ops for a local mutation and reapplies raw engine ops deterministically', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const { undoOps } = engine.captureUndoOps(() => {
      engine.setCellValue('Sheet1', 'A1', 'seed')
    })

    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({
      tag: ValueTag.String,
      value: 'seed',
    })
    expect(undoOps).not.toBeNull()

    const redoOps = engine.applyOps(undoOps ?? [], { captureUndo: true })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(redoOps).not.toBeNull()

    engine.applyOps(redoOps ?? [], { captureUndo: true })
    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({
      tag: ValueTag.String,
      value: 'seed',
    })
  })

  it('emits cell invalidation for local applyOps batches without captured undo', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const events: EngineEvent[] = []
    const unsubscribe = engine.subscribe((event) => {
      events.push(event)
    })

    engine.applyOps([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 7 }], {
      source: 'local',
    })

    unsubscribe()

    expect(events).toHaveLength(1)
    expect(events[0]).toMatchObject({
      kind: 'batch',
      invalidation: 'cells',
    })
    expect(events[0]?.changedCellIndices.length).toBeGreaterThan(0)
    expect(events[0]?.changedCells.map((change) => `${change.sheetName}!${change.a1}`)).toEqual(['Sheet1!A1', 'Sheet1!B1'])
  })

  it('applies coordinate-native cell mutations with formula recomputation and undo compatibility', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.setCellValueAt(sheetId, 0, 0, 2)
    engine.setCellFormulaAt(sheetId, 0, 1, 'A1*3')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })

    const undoOps = engine.applyCellMutationsAt(
      [
        { sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 5 } },
        { sheetId, mutation: { kind: 'clearCell', row: 1, col: 0 } },
      ],
      1,
    )

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(undoOps).not.toBeNull()

    engine.applyOps(undoOps ?? [], { captureUndo: true })

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
  })

  it('applies coordinate-native restore mutations without recording undo history', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'restore-cell-mutation-refs' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    const undoOps = engine.applyCellMutationsAtWithOptions(
      [
        { sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 5 } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' } },
      ],
      {
        captureUndo: false,
        potentialNewCells: 2,
        source: 'restore',
      },
    )

    expect(undoOps).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'full',
      }),
    )
    expect(engine.undo()).toBe(true)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()

    unsubscribe()
  })

  it('applies coordinate-native restore mutations without listeners', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'restore-cell-mutation-refs-no-listeners' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    const undoOps = engine.applyCellMutationsAtWithOptions([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 9 } }], {
      captureUndo: false,
      potentialNewCells: 1,
      source: 'restore',
    })

    expect(undoOps).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.undo()).toBe(true)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()
  })

  it('emits standard engine events for restore mutations without tracked listeners', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'restore-cell-mutation-general-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const listener = vi.fn()
    const unsubscribe = engine.subscribe(listener)

    const undoOps = engine.applyCellMutationsAtWithOptions([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 11 } }], {
      captureUndo: false,
      potentialNewCells: 1,
      source: 'restore',
    })

    expect(undoOps).toBeNull()
    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'full',
      }),
    )
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(engine.undo()).toBe(true)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()

    unsubscribe()
  })

  it('emits standard engine events for coordinate-native clear cell mutations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'clear-cell-at-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)
    const listener = vi.fn()
    const unsubscribe = engine.subscribe(listener)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    engine.clearCellAt(sheetId, 0, 0)

    expect(listener).toHaveBeenCalled()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    unsubscribe()
  })

  it('reads rectangular range values as a dense matrix without per-cell callers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }, [
      [11, 12],
      [13, 14],
    ])
    engine.setCellFormula('Sheet1', 'D2', 'SUM(B2:C2)')

    expect(
      engine.getRangeValues({
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'D3',
      }),
    ).toEqual([
      [{ tag: ValueTag.Empty }, { tag: ValueTag.Empty }, { tag: ValueTag.Empty }, { tag: ValueTag.Empty }],
      [
        { tag: ValueTag.Empty },
        { tag: ValueTag.Number, value: 11 },
        { tag: ValueTag.Number, value: 12 },
        { tag: ValueTag.Number, value: 23 },
      ],
      [{ tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 13 }, { tag: ValueTag.Number, value: 14 }, { tag: ValueTag.Empty }],
    ])
  })

  it('copies and fills rectangular ranges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [
      [1, 2],
      [3, 4],
    ])

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E2' },
    )
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 4 })

    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'D5' },
    )
    expect(engine.getCellValue('Sheet1', 'A4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D5')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('tracks sync client connection state and forwards local batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()

    const forwarded: EngineOpBatch[] = []
    let connected = false
    let disconnected = false

    await engine.connectSyncClient({
      async connect(this: void, handlers: Parameters<EngineSyncClient['connect']>[0]) {
        connected = true
        handlers.setState('behind')
        return {
          send(batch) {
            forwarded.push(batch)
          },
          async disconnect() {
            disconnected = true
          },
        }
      },
    })

    expect(connected).toBe(true)
    expect(engine.getSyncState()).toBe('behind')

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 9)

    expect(forwarded).toHaveLength(2)
    expect(forwarded[1]?.ops).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 9 }])

    await engine.disconnectSyncClient()
    expect(disconnected).toBe(true)
    expect(engine.getSyncState()).toBe('local-only')
  })

  it('disables sync when replica version tracking is off', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'local-only',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })

    await expect(
      engine.connectSyncClient({
        connect() {
          throw new Error('should not connect')
        },
      }),
    ).rejects.toThrow('Sync is unavailable when trackReplicaVersions is disabled; construct the engine with trackReplicaVersions enabled.')
  })
})
