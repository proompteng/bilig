import { compileFormula } from '@bilig/formula'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'
import { CellFlags } from '../cell-store.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectScalarDescriptor, RuntimeFormula } from '../engine/runtime-state.js'
import { createEngineCounters } from '../perf/engine-counters.js'
import { FormulaTable } from '../formula-table.js'
import { WorkbookStore } from '../workbook-store.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import { createOperationDirectFormulaDeltas } from '../engine/services/operation-direct-formula-deltas.js'

type DirectFormulaDeltaState = Parameters<typeof createOperationDirectFormulaDeltas>[0]['state']

const TEST_COMPILED_FORMULA = compileFormula('0')
const TEST_DIRECT_AGGREGATE: RuntimeDirectAggregateDescriptor = {
  regionId: 1,
  aggregateKind: 'sum',
  sheetName: 'Sheet1',
  rowStart: 0,
  rowEnd: 0,
  col: 0,
  colEnd: 0,
  length: 1,
}
const TEST_DIRECT_SCALAR: RuntimeDirectScalarDescriptor = {
  kind: 'abs',
  operand: { kind: 'literal-number', value: 1 },
}

describe('createOperationDirectFormulaDeltas', () => {
  it('should apply direct formula deltas through batched column updates', () => {
    // Arrange
    const { counters, formulas, helpers, sheet, workbook } = createHarness({
      canSkipTerminalFormulaColumnVersion: () => false,
    })
    const first = createCell(workbook, sheet.id, {
      row: 0,
      col: 0,
      value: 5,
      stringId: 7,
      error: ErrorCode.NA,
      flags: CellFlags.Materialized | CellFlags.SpillChild | CellFlags.PivotOutput,
      version: 3,
    })
    const second = createCell(workbook, sheet.id, {
      row: 1,
      col: 0,
      value: 7,
      stringId: 9,
      error: ErrorCode.VALUE,
      flags: CellFlags.Materialized | CellFlags.SpillChild,
      version: 4,
    })
    formulas.set(first, createRuntimeFormula(first, { directAggregate: TEST_DIRECT_AGGREGATE }))
    formulas.set(second, createRuntimeFormula(second, { directScalar: TEST_DIRECT_SCALAR }))

    const collection = new DirectFormulaIndexCollection()
    collection.addDelta(first, 2)
    collection.addDelta(second, -1)

    // Act
    const changed = helpers.tryApplyDirectFormulaDeltas(collection)

    // Assert
    expect(Array.from(changed ?? [])).toEqual([first, second])
    expect(workbook.cellStore.numbers[first]).toBe(7)
    expect(workbook.cellStore.numbers[second]).toBe(6)
    expect(workbook.cellStore.stringIds[first]).toBe(0)
    expect(workbook.cellStore.stringIds[second]).toBe(0)
    expect(workbook.cellStore.errors[first]).toBe(ErrorCode.None)
    expect(workbook.cellStore.errors[second]).toBe(ErrorCode.None)
    expect(workbook.cellStore.flags[first]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.flags[second]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.versions[first]).toBe(4)
    expect(workbook.cellStore.versions[second]).toBe(5)
    expect(sheet.columnVersions[0]).toBe(1)
    expect(counters.directAggregateDeltaApplications).toBe(1)
    expect(counters.directScalarDeltaApplications).toBe(1)
  })

  it('should validate complete direct formula deltas with one formula table read per cell', () => {
    // Arrange
    const { counters, formulas, helpers, sheet, workbook } = createHarness({
      canSkipTerminalFormulaColumnVersion: () => true,
    })
    const first = createCell(workbook, sheet.id, { row: 0, col: 0, value: 5, version: 3 })
    const second = createCell(workbook, sheet.id, { row: 1, col: 0, value: 7, version: 4 })
    formulas.set(first, createRuntimeFormula(first, { directAggregate: TEST_DIRECT_AGGREGATE }))
    formulas.set(second, createRuntimeFormula(second, { directScalar: TEST_DIRECT_SCALAR }))
    const getFormula = vi.spyOn(formulas, 'get')

    const collection = new DirectFormulaIndexCollection()
    collection.addDelta(first, 2)
    collection.addDelta(second, -1)

    // Act
    const changed = helpers.tryApplyDirectFormulaDeltas(collection)

    // Assert
    expect(Array.from(changed ?? [])).toEqual([first, second])
    expect(getFormula).toHaveBeenCalledTimes(2)
    expect(workbook.cellStore.numbers[first]).toBe(7)
    expect(workbook.cellStore.numbers[second]).toBe(6)
    expect(counters.directAggregateDeltaApplications).toBe(1)
    expect(counters.directScalarDeltaApplications).toBe(1)
  })

  it('should apply mixed scalar and aggregate constant deltas without materializing per-cell deltas', () => {
    // Arrange
    const { counters, formulas, helpers, sheet, workbook } = createHarness({
      canSkipTerminalFormulaColumnVersion: () => true,
    })
    const scalarCells: number[] = []
    const aggregateCells: number[] = []
    for (let row = 0; row < 24; row += 1) {
      const cellIndex = createCell(workbook, sheet.id, { row, col: 1, value: row, version: 1 })
      scalarCells.push(cellIndex)
      formulas.set(cellIndex, createRuntimeFormula(cellIndex, { directScalar: TEST_DIRECT_SCALAR }))
    }
    for (let row = 0; row < 24; row += 1) {
      const cellIndex = createCell(workbook, sheet.id, { row, col: 2, value: row * 10, version: 1 })
      aggregateCells.push(cellIndex)
      formulas.set(cellIndex, createRuntimeFormula(cellIndex, { directAggregate: TEST_DIRECT_AGGREGATE }))
    }
    const collection = new DirectFormulaIndexCollection()
    collection.appendConstantDelta(Uint32Array.from(scalarCells), 5, 'scalar')
    collection.appendConstantDelta(Uint32Array.from(aggregateCells), 5)

    // Act
    const changed = helpers.tryApplyDirectFormulaDeltas(collection)

    // Assert
    expect(collection.getConstantDelta()).toBe(5)
    expect(collection.getConstantScalarDelta()).toBeUndefined()
    expect(Array.from(changed ?? [])).toEqual([...scalarCells, ...aggregateCells])
    expect(workbook.cellStore.numbers[scalarCells[23]]).toBe(28)
    expect(workbook.cellStore.numbers[aggregateCells[23]]).toBe(235)
    expect(counters.directScalarDeltaApplications).toBe(24)
    expect(counters.directAggregateDeltaApplications).toBe(24)
  })

  it('should apply same-delta direct formulas through the batch writer', () => {
    // Arrange
    const { helpers, sheet, workbook } = createHarness()
    const first = createCell(workbook, sheet.id, {
      row: 0,
      col: 0,
      value: 11,
      stringId: 3,
      error: ErrorCode.NAME,
      flags: CellFlags.Materialized | CellFlags.SpillChild,
      version: 4,
    })
    const second = createCell(workbook, sheet.id, {
      row: 1,
      col: 0,
      value: 17,
      stringId: 5,
      error: ErrorCode.VALUE,
      flags: CellFlags.Materialized | CellFlags.PivotOutput,
      version: 7,
    })
    const third = createCell(workbook, sheet.id, {
      row: 2,
      col: 0,
      value: 23,
      version: 2,
    })

    // Act
    const applied = workbook.withBatchedColumnVersionUpdates(() =>
      helpers.applyDirectFormulaNumericDeltaBatch(Uint32Array.of(first, second, third), 6),
    )

    // Assert
    expect(applied).toBe(true)
    expect(workbook.cellStore.numbers[first]).toBe(17)
    expect(workbook.cellStore.numbers[second]).toBe(23)
    expect(workbook.cellStore.numbers[third]).toBe(29)
    expect(workbook.cellStore.stringIds[first]).toBe(0)
    expect(workbook.cellStore.stringIds[second]).toBe(0)
    expect(workbook.cellStore.errors[first]).toBe(ErrorCode.None)
    expect(workbook.cellStore.errors[second]).toBe(ErrorCode.None)
    expect(workbook.cellStore.flags[first]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.flags[second]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.versions[first]).toBe(5)
    expect(workbook.cellStore.versions[second]).toBe(8)
    expect(workbook.cellStore.versions[third]).toBe(3)
    expect(sheet.columnVersions[0]).toBe(1)
  })

  it('should reject batch direct formula deltas without partial writes', () => {
    // Arrange
    const { helpers, sheet, workbook } = createHarness()
    const numeric = createCell(workbook, sheet.id, {
      row: 0,
      col: 0,
      value: 11,
      version: 4,
    })
    const text = createCell(workbook, sheet.id, {
      row: 1,
      col: 0,
      tag: ValueTag.String,
      stringId: 9,
      version: 7,
    })

    // Act
    const applied = helpers.applyDirectFormulaNumericDeltaBatch(Uint32Array.of(numeric, text), 6)

    // Assert
    expect(applied).toBe(false)
    expect(workbook.cellStore.numbers[numeric]).toBe(11)
    expect(workbook.cellStore.versions[numeric]).toBe(4)
    expect(workbook.cellStore.stringIds[text]).toBe(9)
    expect(workbook.cellStore.versions[text]).toBe(7)
    expect(sheet.columnVersions[0] ?? 0).toBe(0)
  })

  it('should apply validated direct scalar deltas through terminal writes', () => {
    // Arrange
    const { counters, formulas, helpers, sheet, workbook } = createHarness({
      canSkipDirectFormulaColumnVersion: () => false,
    })
    const first = createCell(workbook, sheet.id, {
      row: 0,
      col: 0,
      value: 2,
      stringId: 3,
      error: ErrorCode.NAME,
      flags: CellFlags.Materialized | CellFlags.SpillChild | CellFlags.PivotOutput,
      version: 1,
    })
    const second = createCell(workbook, sheet.id, {
      row: 1,
      col: 1,
      value: 8,
      stringId: 4,
      error: ErrorCode.REF,
      flags: CellFlags.Materialized | CellFlags.PivotOutput,
      version: 6,
    })
    formulas.set(first, createRuntimeFormula(first, { directScalar: TEST_DIRECT_SCALAR }))
    formulas.set(second, createRuntimeFormula(second, { directScalar: TEST_DIRECT_SCALAR }))

    const collection = new DirectFormulaIndexCollection()
    const directScalarCells = Uint32Array.of(first, second)
    collection.appendConstantDelta(directScalarCells, 4, 'scalar')
    collection.markScalarDeltaCellsValidated()

    // Act
    const changed = helpers.tryApplyDirectScalarDeltas(collection)

    // Assert
    expect(changed).toBe(directScalarCells)
    expect(Array.from(changed ?? [])).toEqual([first, second])
    expect(workbook.cellStore.numbers[first]).toBe(6)
    expect(workbook.cellStore.numbers[second]).toBe(12)
    expect(workbook.cellStore.stringIds[first]).toBe(0)
    expect(workbook.cellStore.stringIds[second]).toBe(0)
    expect(workbook.cellStore.errors[first]).toBe(ErrorCode.None)
    expect(workbook.cellStore.errors[second]).toBe(ErrorCode.None)
    expect(workbook.cellStore.flags[first]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.flags[second]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.versions[first]).toBe(2)
    expect(workbook.cellStore.versions[second]).toBe(7)
    expect(sheet.columnVersions[0] ?? 0).toBe(0)
    expect(sheet.columnVersions[1] ?? 0).toBe(0)
    expect(counters.directScalarDeltaApplications).toBe(2)
  })

  it('should apply clean validated direct scalar deltas without slot cleanup work', () => {
    // Arrange
    const { counters, formulas, sheet, workbook, helpers } = createHarness({
      canSkipDirectFormulaColumnVersion: () => false,
    })
    const first = createCell(workbook, sheet.id, {
      row: 0,
      col: 0,
      value: 2,
      version: 1,
    })
    const second = createCell(workbook, sheet.id, {
      row: 1,
      col: 1,
      value: 8,
      version: 6,
    })
    formulas.set(first, createRuntimeFormula(first, { directScalar: TEST_DIRECT_SCALAR }))
    formulas.set(second, createRuntimeFormula(second, { directScalar: TEST_DIRECT_SCALAR }))

    const collection = new DirectFormulaIndexCollection()
    const directScalarCells = Uint32Array.of(first, second)
    collection.appendConstantDelta(directScalarCells, 4, 'scalar')
    collection.markScalarDeltaCellsValidated()
    collection.markScalarDeltaCellsCleanNumber()

    // Act
    const changed = helpers.tryApplyDirectScalarDeltas(collection)

    // Assert
    expect(changed).toBe(directScalarCells)
    expect(workbook.cellStore.numbers[first]).toBe(6)
    expect(workbook.cellStore.numbers[second]).toBe(12)
    expect(workbook.cellStore.stringIds[first]).toBe(0)
    expect(workbook.cellStore.stringIds[second]).toBe(0)
    expect(workbook.cellStore.errors[first]).toBe(ErrorCode.None)
    expect(workbook.cellStore.errors[second]).toBe(ErrorCode.None)
    expect(workbook.cellStore.flags[first]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.flags[second]).toBe(CellFlags.Materialized)
    expect(workbook.cellStore.versions[first]).toBe(2)
    expect(workbook.cellStore.versions[second]).toBe(7)
    expect(sheet.columnVersions[0] ?? 0).toBe(0)
    expect(sheet.columnVersions[1] ?? 0).toBe(0)
    expect(counters.directScalarDeltaApplications).toBe(2)
  })

  it('should trust linear direct scalar closures without revalidating formula table entries', () => {
    // Arrange
    const { counters, formulas, sheet, workbook, helpers } = createHarness()
    const first = createCell(workbook, sheet.id, {
      row: 0,
      col: 0,
      value: 2,
      version: 1,
    })
    const second = createCell(workbook, sheet.id, {
      row: 1,
      col: 0,
      value: 8,
      version: 6,
    })
    const getFormula = vi.spyOn(formulas, 'get')

    const collection = new DirectFormulaIndexCollection()
    const directScalarCells = Uint32Array.of(first, second)
    collection.appendConstantDelta(directScalarCells, 4, 'scalar')
    collection.markScalarDeltaCellsValidated()
    collection.markScalarDeltaCellsCleanNumber()
    collection.markScalarDeltaCellsTrustedDirectScalarFormulas()

    // Act
    const changed = helpers.tryApplyDirectScalarDeltas(collection)

    // Assert
    expect(changed).toBe(directScalarCells)
    expect(getFormula).not.toHaveBeenCalled()
    expect(workbook.cellStore.numbers[first]).toBe(6)
    expect(workbook.cellStore.numbers[second]).toBe(12)
    expect(counters.directScalarDeltaApplications).toBe(2)
  })

  it('should return undefined when a direct formula delta targets a cycle cell', () => {
    // Arrange
    const { counters, helpers, sheet, workbook } = createHarness()
    const cellIndex = createCell(workbook, sheet.id, {
      row: 0,
      col: 0,
      value: 5,
      flags: CellFlags.Materialized | CellFlags.InCycle,
      version: 7,
    })

    const collection = new DirectFormulaIndexCollection()
    collection.addDelta(cellIndex, 3)

    // Act
    const changed = helpers.tryApplyDirectFormulaDeltas(collection)

    // Assert
    expect(changed).toBeUndefined()
    expect(workbook.cellStore.numbers[cellIndex]).toBe(5)
    expect(workbook.cellStore.versions[cellIndex]).toBe(7)
    expect(sheet.columnVersions[0] ?? 0).toBe(0)
    expect(counters.directAggregateDeltaApplications).toBe(0)
    expect(counters.directScalarDeltaApplications).toBe(0)
  })

  it('should return false when a terminal direct formula write targets a non-numeric cell', () => {
    // Arrange
    const { helpers, sheet, workbook } = createHarness()
    const cellIndex = createCell(workbook, sheet.id, {
      row: 0,
      col: 1,
      tag: ValueTag.String,
      stringId: 11,
      version: 4,
    })

    // Act
    const applied = helpers.applyTerminalDirectFormulaNumericDelta(cellIndex, 3)

    // Assert
    expect(applied).toBe(false)
    expect(workbook.cellStore.tags[cellIndex]).toBe(ValueTag.String)
    expect(workbook.cellStore.stringIds[cellIndex]).toBe(11)
    expect(workbook.cellStore.versions[cellIndex]).toBe(4)
    expect(sheet.columnVersions[1] ?? 0).toBe(0)
  })
})

// Helpers

function createHarness(
  overrides: Partial<
    Pick<
      Parameters<typeof createOperationDirectFormulaDeltas>[0],
      'canSkipTerminalFormulaColumnVersion' | 'canSkipDirectFormulaColumnVersion'
    >
  > = {},
) {
  const counters = createEngineCounters()
  const workbook = new WorkbookStore('direct-formula-deltas', counters)
  const sheet = workbook.createSheet('Sheet1')
  const formulas = new FormulaTable<RuntimeFormula>(workbook.cellStore)

  const state: DirectFormulaDeltaState = {
    workbook,
    counters,
    formulas,
  }

  const helpers = createOperationDirectFormulaDeltas({
    state,
    canSkipTerminalFormulaColumnVersion: overrides.canSkipTerminalFormulaColumnVersion ?? (() => false),
    canSkipDirectFormulaColumnVersion: overrides.canSkipDirectFormulaColumnVersion ?? (() => false),
  })

  return { counters, formulas, helpers, sheet, workbook }
}

function createCell(
  workbook: WorkbookStore,
  sheetId: number,
  options: {
    row: number
    col: number
    value?: number
    tag?: ValueTag
    stringId?: number
    error?: ErrorCode
    flags?: number
    version?: number
  },
): number {
  const cellIndex = workbook.cellStore.allocate(sheetId, options.row, options.col)

  workbook.cellStore.tags[cellIndex] = options.tag ?? ValueTag.Number
  workbook.cellStore.numbers[cellIndex] = options.value ?? 0
  workbook.cellStore.stringIds[cellIndex] = options.stringId ?? 0
  workbook.cellStore.errors[cellIndex] = options.error ?? ErrorCode.None
  workbook.cellStore.flags[cellIndex] = options.flags ?? CellFlags.Materialized
  workbook.cellStore.versions[cellIndex] = options.version ?? 0

  return cellIndex
}

function createRuntimeFormula(
  cellIndex: number,
  overrides: {
    directAggregate?: RuntimeDirectAggregateDescriptor
    directScalar?: RuntimeDirectScalarDescriptor
  } = {},
): RuntimeFormula {
  return {
    cellIndex,
    formulaSlotId: 0,
    planId: 0,
    templateId: undefined,
    source: '0',
    compiled: TEST_COMPILED_FORMULA,
    plan: { id: 0, source: '0', compiled: TEST_COMPILED_FORMULA },
    dependencyIndices: new Uint32Array(0),
    dependencyEntities: { ptr: -1, len: 0, cap: 0 },
    rangeDependencies: new Uint32Array(0),
    graphRangeDependencies: new Uint32Array(0),
    runtimeProgram: TEST_COMPILED_FORMULA.program,
    constants: TEST_COMPILED_FORMULA.constants,
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: TEST_COMPILED_FORMULA.program.length,
    constNumberOffset: 0,
    constNumberLength: TEST_COMPILED_FORMULA.constants.length,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: overrides.directAggregate,
    directScalar: overrides.directScalar,
    directCriteria: undefined,
  }
}
