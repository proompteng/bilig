import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { compileCriteriaMatcher, indexToColumn } from '@bilig/formula'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import { operationServiceTestHooks, type EngineOperationService } from '../engine/services/operation-service.js'
import {
  approximateUniformLookupCurrentResult,
  approximateUniformLookupNumericResult,
  exactUniformLookupCurrentResult,
  exactUniformLookupNumericResult,
  normalizeApproximateNumericValue,
  normalizeApproximateTextValue,
  normalizeExactLookupKey,
} from '../engine/services/direct-lookup-helpers.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from '../cell-mutations-at.js'

function isEngineOperationService(value: unknown): value is EngineOperationService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'applyBatch') === 'function' && typeof Reflect.get(value, 'applyDerivedOp') === 'function'
}

function getOperationService(engine: SpreadsheetEngine): EngineOperationService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (!isEngineOperationService(operations)) {
    throw new TypeError('Expected engine operation service')
  }
  return operations
}

function hasRuntimeStateMetrics(value: unknown): value is {
  getLastMetrics(): unknown
  setLastMetrics(metrics: unknown): void
} {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const getLastMetrics = Reflect.get(value, 'getLastMetrics')
  const setLastMetrics = Reflect.get(value, 'setLastMetrics')
  return typeof getLastMetrics === 'function' && typeof setLastMetrics === 'function'
}

function getRuntimeState(engine: SpreadsheetEngine): {
  getLastMetrics(): unknown
  setLastMetrics(metrics: unknown): void
} {
  const state = Reflect.get(engine, 'state')
  if (!hasRuntimeStateMetrics(state)) {
    throw new TypeError('Expected runtime state metric accessors')
  }
  return state
}

function expectBatch<Batch>(batch: Batch | undefined): Batch {
  expect(batch).toBeDefined()
  return batch
}

function lookupTestString(id: number): string {
  return id === 7 ? 'needle' : 'fallback'
}

function getReplicaState(engine: SpreadsheetEngine) {
  const replicaState = Reflect.get(engine, 'replicaState')
  if (typeof replicaState !== 'object' || replicaState === null) {
    throw new TypeError('Expected engine replica state')
  }
  return replicaState
}

describe('EngineOperationService', () => {
  it('covers operation helper branches used by direct formulas and lookup tracking', () => {
    expect(operationServiceTestHooks.directAggregateNumericContribution({ tag: ValueTag.Number, value: 4 })).toBe(4)
    expect(operationServiceTestHooks.directAggregateNumericContribution({ tag: ValueTag.Boolean, value: true })).toBe(1)
    expect(operationServiceTestHooks.directAggregateNumericContribution({ tag: ValueTag.Boolean, value: false })).toBe(0)
    expect(operationServiceTestHooks.directAggregateNumericContribution({ tag: ValueTag.Empty })).toBe(0)
    expect(operationServiceTestHooks.directAggregateNumericContribution({ tag: ValueTag.String, value: 'x' })).toBe(0)
    expect(operationServiceTestHooks.directAggregateNumericContribution({ tag: ValueTag.Error, code: ErrorCode.VALUE })).toBeUndefined()

    expect(normalizeExactLookupKey({ tag: ValueTag.Empty }, lookupTestString)).toBe('e:')
    expect(normalizeExactLookupKey({ tag: ValueTag.Number, value: -0 }, lookupTestString)).toBe('n:0')
    expect(normalizeExactLookupKey({ tag: ValueTag.Boolean, value: true }, lookupTestString)).toBe('b:1')
    expect(normalizeExactLookupKey({ tag: ValueTag.String, value: 'local' }, lookupTestString, 7)).toBe('s:NEEDLE')
    expect(normalizeExactLookupKey({ tag: ValueTag.Error, code: ErrorCode.NA }, lookupTestString)).toBeUndefined()

    expect(normalizeApproximateNumericValue({ tag: ValueTag.Empty })).toBe(0)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Number, value: -0 })).toBe(0)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Boolean, value: false })).toBe(0)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.String, value: 'x' })).toBeUndefined()
    expect(normalizeApproximateTextValue({ tag: ValueTag.String, value: 'local' }, lookupTestString, 7)).toBe('NEEDLE')
    expect(normalizeApproximateTextValue({ tag: ValueTag.Empty }, lookupTestString)).toBe('')
    expect(normalizeApproximateTextValue({ tag: ValueTag.Number, value: 1 }, lookupTestString)).toBeUndefined()

    const exactLookup: Parameters<typeof exactUniformLookupNumericResult>[0] = {
      columnVersion: 1,
      kind: 'exact-uniform-numeric',
      length: 4,
      resultKind: 'row-number',
      rowEnd: 3,
      rowStart: 0,
      start: 10,
      step: 2,
      structureVersion: 1,
    }
    expect(exactUniformLookupNumericResult(exactLookup, 14)).toBe(3)
    expect(exactUniformLookupNumericResult(exactLookup, 15)).toBeUndefined()
    expect(exactUniformLookupCurrentResult(exactLookup, 15)).toEqual({ kind: 'error', code: ErrorCode.NA })
    expect(
      exactUniformLookupNumericResult(
        {
          ...exactLookup,
          tailPatch: { columnVersion: 2, newNumeric: 99, oldNumeric: 10, row: 0 },
        },
        99,
      ),
    ).toBe(1)

    const approximateLookup: Parameters<typeof approximateUniformLookupNumericResult>[0] = {
      columnVersion: 1,
      kind: 'approximate-uniform-numeric',
      length: 4,
      matchMode: 1,
      resultKind: 'row-number',
      rowEnd: 3,
      rowStart: 0,
      start: 10,
      step: 2,
      structureVersion: 1,
    }
    expect(approximateUniformLookupNumericResult(approximateLookup, 13)).toBe(2)
    expect(approximateUniformLookupNumericResult(approximateLookup, 9)).toBeUndefined()
    expect(approximateUniformLookupCurrentResult(approximateLookup, 9)).toEqual({
      kind: 'error',
      code: ErrorCode.NA,
    })

    const directCriteria: Parameters<typeof operationServiceTestHooks.directCriteriaTouchesPoint>[0] = {
      aggregateKind: 'sum',
      aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 4, col: 1, length: 5, regionId: 10 },
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 4, col: 2, length: 5, regionId: 11 },
          criterion: { kind: 'literal', matcher: compileCriteriaMatcher('x') },
        },
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 4, col: 3, length: 5, regionId: 12 },
          criterion: { cellIndex: 42, kind: 'cell' },
        },
      ],
    }
    expect(operationServiceTestHooks.directCriteriaTouchesPoint(directCriteria, { sheetName: 'Sheet1', row: 3, col: 1 })).toBe(true)
    expect(operationServiceTestHooks.directCriteriaTouchesPoint(directCriteria, { sheetName: 'Sheet1', row: 3, col: 3 })).toBe(true)
    expect(
      operationServiceTestHooks.directCriteriaTouchesPoint(directCriteria, { sheetName: 'Other', row: 3, col: 3, inputCellIndex: 42 }),
    ).toBe(true)
    expect(operationServiceTestHooks.directCriteriaTouchesPoint(directCriteria, { sheetName: 'Other', row: 3, col: 3 })).toBe(false)

    expect(Array.from(operationServiceTestHooks.mergeChangedCellIndices([], [3, 4]))).toEqual([3, 4])
    expect(Array.from(operationServiceTestHooks.mergeChangedCellIndices([3], [3]))).toEqual([3])
    expect(Array.from(operationServiceTestHooks.mergeChangedCellIndices([3], [4]))).toEqual([3, 4])
    expect(Array.from(operationServiceTestHooks.mergeChangedCellIndices([3, 4], [4, 5]))).toEqual([3, 4, 5])
    expect(Array.from(operationServiceTestHooks.composeSingleDisjointExplicitEventChanges(2, Uint32Array.of(5, 6)))).toEqual([2, 5, 6])
  })

  it('covers direct formula index collection delta materialization branches', () => {
    const collection = new DirectFormulaIndexCollection()

    collection.appendConstantDelta(Uint32Array.from([10, 11, 12]), 3, 'scalar')
    expect(collection.size).toBe(3)
    expect(collection.has(11)).toBe(true)
    expect(collection.hasDelta(12)).toBe(true)
    expect(collection.getDelta(10)).toBe(3)
    expect(collection.getDeltaAt(2)).toBe(3)
    expect(collection.getScalarDeltaAt(1)).toBe(3)
    expect(collection.getConstantScalarDelta()).toBe(3)
    expect(collection.hasCompleteDeltas()).toBe(true)
    expect(collection.hasCompleteScalarDeltas()).toBe(true)
    collection.markScalarDeltaCellsValidated()
    expect(collection.hasValidatedScalarDeltaCells()).toBe(true)

    collection.addScalarDelta(13, 3)
    collection.addDelta(11, 2)
    collection.addCurrentResult(12, { kind: 'number', value: 42 })
    expect(collection.getDelta(11)).toBe(5)
    expect(collection.getScalarDeltaAt(1)).toBeUndefined()
    expect(collection.getCurrentResult(12)).toEqual({ kind: 'number', value: 42 })
    expect(collection.getCurrentResultAt(2)).toEqual({ kind: 'number', value: 42 })
    expect(collection.getConstantScalarDelta()).toBeUndefined()
    expect(collection.hasCompleteScalarDeltas()).toBe(false)

    collection.markDirectFormulaInputCovered(101)
    collection.markDirectFormulaInputCovered(101)
    collection.markDirectRangeInputCovered(202)
    collection.markDirectRangeInputCovered(202)
    expect(collection.hasCoveredDirectFormulaInput(101)).toBe(true)
    expect(collection.hasCoveredDirectFormulaInput(102)).toBe(false)
    expect(collection.hasCoveredDirectRangeInput(202)).toBe(true)
    expect(collection.hasCoveredDirectRangeInput(203)).toBe(false)

    const cells: number[] = []
    const indexed: string[] = []
    collection.forEach((cellIndex) => cells.push(cellIndex))
    collection.forEachIndexed((cellIndex, index) => indexed.push(`${index}:${cellIndex}`))
    expect(cells).toEqual([10, 11, 12, 13])
    expect(indexed).toEqual(['0:10', '1:11', '2:12', '3:13'])

    const largeCollection = new DirectFormulaIndexCollection()
    largeCollection.appendDeltas(
      Uint32Array.from(Array.from({ length: 18 }, (_unused, index) => index + 1)),
      Array.from({ length: 18 }, (_unused, index) => index + 10),
      'scalar',
    )
    largeCollection.appendDeltas(Uint32Array.from([5, 30]), [4, 8])
    expect(largeCollection.has(30)).toBe(true)
    expect(largeCollection.getDelta(5)).toBe(18)
    expect(largeCollection.getDelta(30)).toBe(8)
    expect(largeCollection.getScalarDeltaAt(4)).toBeUndefined()
    expect(largeCollection.hasCompleteDeltas()).toBe(true)
  })

  it('applies remote rename batches through the service and keeps the selection on the renamed sheet', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'operation-rename', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'operation-rename', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: Parameters<SpreadsheetEngine['applyRemoteBatch']>[0][] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Old')
    const createdSheetBatch = expectBatch(outbound.at(-1))
    replica.applyRemoteBatch(createdSheetBatch)
    replica.setSelection('Old', 'B2')

    primary.renameSheet('Old', 'New')
    const renameBatch = expectBatch(outbound.at(-1))

    Effect.runSync(getOperationService(replica).applyBatch(renameBatch, 'remote'))

    expect(replica.getSelectionState()).toMatchObject({
      sheetName: 'New',
      address: 'B2',
      anchorAddress: 'B2',
    })
  })

  it('rejects stale remote cell replays behind sheet tombstones through the service', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'operation-tombstone', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'operation-tombstone', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: Parameters<SpreadsheetEngine['applyRemoteBatch']>[0][] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Sheet1')
    const createdSheetBatch = expectBatch(outbound.at(-1))

    primary.setCellValue('Sheet1', 'A1', 7)
    const valueBatch = expectBatch(outbound.at(-1))

    primary.deleteSheet('Sheet1')
    const deleteBatch = expectBatch(outbound.at(-1))

    replica.applyRemoteBatch(createdSheetBatch)
    replica.applyRemoteBatch(deleteBatch)

    const restored = new SpreadsheetEngine({ workbookName: 'restored', replicaId: 'b' })
    await restored.ready()
    restored.importSnapshot(replica.exportSnapshot())
    restored.importReplicaSnapshot(replica.exportReplicaSnapshot())

    Effect.runSync(getOperationService(restored).applyBatch(valueBatch, 'remote'))

    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
  })

  it('does not rewrite last metrics once per formula during snapshot restore', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-restore-metrics' })
    await engine.ready()

    const state = getRuntimeState(engine)
    const setLastMetricsSpy = vi.spyOn(state, 'setLastMetrics')
    setLastMetricsSpy.mockClear()

    engine.importSnapshot({
      version: 1,
      workbook: { name: 'operation-restore-metrics' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', formula: 'A1*2' },
            { address: 'C1', formula: 'B1+1' },
          ],
        },
      ],
    })

    expect(setLastMetricsSpy).toHaveBeenCalledTimes(2)
  })

  it('applies structural and metadata batches through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-structural-metadata' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)')

    const batch = createBatch(getReplicaState(engine), [
      { kind: 'insertColumns', sheetName: 'Sheet1', start: 1, count: 1 },
      {
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
        size: 24,
        hidden: false,
      },
      {
        kind: 'updateColumnMetadata',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
        size: 90,
        hidden: true,
      },
    ])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCell('Sheet1', 'E1').formula).toBe('SUM(A1:C1)')
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getRowMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 0, count: 1, size: 24, hidden: false }])
    expect(engine.getColumnMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 0, count: 1, size: 90, hidden: true }])
  })

  it('deletes structural columns through logical axis membership without survivor remaps', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-structural-delete-columns-axis-owned' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 2)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*2`)
    }

    engine.resetPerformanceCounters()
    engine.deleteColumns('Sheet1', 1, 1)

    expect(engine.getCellValue('Sheet1', `A${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount })
    expect(engine.getCell('Sheet1', 'B1').formula).toBe('A1+#REF!')
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('B1*2')
    expect(engine.getPerformanceCounters()).toMatchObject({
      structuralTransactions: 1,
      structuralPlannedCells: rowCount,
      structuralRemovedCells: rowCount,
      structuralSurvivorCellsRemapped: 0,
      sheetGridBlockScans: 0,
      axisMapSplices: 1,
    })
  })

  it('repairs topo ranks locally for acyclic formula rewrites without forcing a full rebuild', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-dynamic-topo' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet1', 'C1', 'B1+1')
    engine.setCellFormula('Sheet1', 'D1', 'C1+1')

    engine.resetPerformanceCounters()
    engine.setCellFormula('Sheet1', 'C1', 'B1+A1')

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const c1Index = engine.workbook.getCellIndex('Sheet1', 'C1')
    const d1Index = engine.workbook.getCellIndex('Sheet1', 'D1')

    expect(b1Index).toBeDefined()
    expect(c1Index).toBeDefined()
    expect(d1Index).toBeDefined()
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getPerformanceCounters().topoRebuilds).toBe(0)
    expect(engine.workbook.cellStore.topoRanks[b1Index!]).toBeLessThan(engine.workbook.cellStore.topoRanks[c1Index!])
    expect(engine.workbook.cellStore.topoRanks[c1Index!]).toBeLessThan(engine.workbook.cellStore.topoRanks[d1Index!])
  })

  it('uses a direct scalar delta root for same-topology formula replacements', async () => {
    const downstreamCount = 24
    const engine = new SpreadsheetEngine({ workbookName: 'operation-formula-rewrite-direct-delta-root' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = 2 + offset
      engine.setCellFormula('Sheet1', `${indexToColumn(col)}1`, `${indexToColumn(col - 1)}1+1`)
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(formulaCellIndex).toBeDefined()
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        cellIndex: formulaCellIndex,
        mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'A1*B1' },
      },
    ]
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 0))

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', `${indexToColumn(2 + downstreamCount)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 2 + downstreamCount,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0 })
    expect(engine.getPerformanceCounters().formulasParsed).toBe(1)
    expect(engine.getPerformanceCounters().formulasBound).toBe(1)
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(downstreamCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    const changedIndices = Array.from(tracked.mock.calls.at(-1)?.[0].changedCellIndices ?? [])
    expect(changedIndices[0]).toBe(formulaCellIndex)
    expect(changedIndices).toContain(engine.workbook.getCellIndex('Sheet1', `${indexToColumn(2 + downstreamCount)}1`))
    unsubscribe()
  })

  it('treats batched clears of already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-clear-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'D4' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats batched null writes to already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-null-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'D4', value: null }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats batched null writes to missing cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-null-missing-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'D4', value: null }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats mutation-ref clears of already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-mutation-clear-empty-noop',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'clearCell', row: 3, col: 3 },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)
    const before = engine.exportSnapshot()

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('prunes explicit null literal writes back out of tracked empty dependency cells', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-null-literal-empty-prune',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellValue', row: 3, col: 3, value: null },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)
    const before = engine.exportSnapshot()

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats mutation-ref null writes to missing cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-mutation-null-missing-noop',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellValue', row: 3, col: 3, value: null },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)
    const before = engine.exportSnapshot()

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('rebinds formulas over existing literal cells through mutation refs', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-mutation-formula-over-literal',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 0, formula: '2+2' },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 4 })
  })

  it('keeps lookup formulas current through generic batch literal writes', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-batch-lookup-write',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 30)
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellValue('Sheet1', 'D2', 25)
    engine.setCellFormula('Sheet1', 'E1', 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D2,A1:A3,1)')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A2', value: 25 }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('updates standalone direct lookup operands without dirty traversal', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-lookup-post-recalc' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[1], [2], [3]])
    engine.setCellValue('Sheet1', 'D1', 2.5)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A3,1)')
    engine.setCellValue('Sheet1', 'D2', 2)
    engine.setCellFormula('Sheet1', 'E2', 'MATCH(D2,A1:A3,0)')

    engine.setCellValue('Sheet1', 'D1', 3.5)
    engine.setCellValue('Sheet1', 'D2', 3)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(2)

    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'D1', 1.5)

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    expect(inputIndex).toBeDefined()
    expect(formulaIndex).toBeDefined()
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex!, formulaIndex!]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('skips direct lookup writeback when an operand edit leaves the result unchanged', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-lookup-no-result-change' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[1], [2], [3]])
    engine.setCellValue('Sheet1', 'D1', 2.1)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A3,1)')
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    engine.resetPerformanceCounters()

    engine.setCellValue('Sheet1', 'D1', 2.2)

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    expect(inputIndex).toBeDefined()
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      kernelSyncOnlyRecalcSkips: 1,
      directFormulaKernelSyncOnlyRecalcSkips: 0,
    })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex!]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('updates indexed exact lookup operands without dirty traversal or index rebuilds', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-direct-indexed-exact-lookup-post-recalc',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[1], [2], [3], [4]])
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A4,0)')
    engine.resetPerformanceCounters()

    engine.setCellValue('Sheet1', 'D1', 4)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      exactIndexBuilds: 0,
      lookupOwnerBuilds: 0,
      changedCellPayloadsBuilt: 0,
    })
  })

  it('updates indexed exact text lookup operands without dirty traversal or index rebuilds', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-direct-indexed-exact-text-lookup-post-recalc',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [['alpha'], ['bravo'], ['charlie'], ['delta']])
    engine.setCellValue('Sheet1', 'D1', 'alpha')
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A4,0)')
    engine.resetPerformanceCounters()

    engine.setCellValue('Sheet1', 'D1', 'delta')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      exactIndexBuilds: 0,
      lookupOwnerBuilds: 0,
      changedCellPayloadsBuilt: 0,
    })
  })

  it('updates non-uniform approximate lookup operands through prepared numeric vectors', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-approximate-duplicate-lookup-post-recalc' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, Math.ceil(row / 2))
    }
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},1)`)
    engine.resetPerformanceCounters()

    engine.setCellValue('Sheet1', 'D1', 11)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      approxIndexBuilds: 0,
      changedCellPayloadsBuilt: 0,
    })
  })

  it('skips dirty traversal for exact lookup column writes that cannot match the numeric operand', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-exact-lookup-column-no-impact',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    engine.resetPerformanceCounters()

    engine.setCellValue('Sheet1', `A${rowCount}`, rowCount + 1_000)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(rowCount / 2),
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${rowCount}`)
    expect(inputIndex).toBeDefined()
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex!]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('keeps exact lookup owners warm after skipped numeric column writes', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-exact-lookup-column-no-impact-owner',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)

    engine.setCellValue('Sheet1', `A${rowCount}`, rowCount + 1_000)
    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'D1', rowCount - 1)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: rowCount - 1 })
    expect(engine.getPerformanceCounters().lookupOwnerBuilds).toBe(0)
  })

  it('keeps exact uniform lookup tail writes correct for later operand edits', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-exact-lookup-tail-patch',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)

    engine.setCellValue('Sheet1', `A${rowCount}`, rowCount + 1_000)
    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'D1', rowCount)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    engine.setCellValue('Sheet1', 'D1', rowCount + 1_000)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: rowCount })
    expect(engine.getPerformanceCounters().lookupOwnerBuilds).toBe(0)
  })

  it('skips dirty traversal for approximate lookup tail writes that preserve the sorted match', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-approximate-lookup-tail-no-impact' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2) + 0.5)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},1)`)
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    engine.resetPerformanceCounters()

    engine.setCellValue('Sheet1', `A${rowCount}`, rowCount + 1)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${rowCount}`)
    expect(inputIndex).toBeDefined()
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex!]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('keeps approximate lookup owners warm after skipped sorted tail writes', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-approximate-lookup-tail-no-impact-owner' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2) + 0.5)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},1)`)

    engine.setCellValue('Sheet1', `A${rowCount}`, rowCount + 1)
    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'D1', rowCount - 0.5)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: rowCount - 1 })
    expect(engine.getPerformanceCounters().lookupOwnerBuilds).toBe(0)
  })

  it('keeps approximate uniform lookup tail writes correct for later operand edits', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-approximate-lookup-tail-patch' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2) + 0.5)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},1)`)

    engine.setCellValue('Sheet1', `A${rowCount}`, rowCount + 1)
    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'D1', rowCount + 0.5)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: rowCount - 1 })

    engine.setCellValue('Sheet1', 'D1', rowCount + 1)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: rowCount })
    expect(engine.getPerformanceCounters().lookupOwnerBuilds).toBe(0)
  })

  it('skips recalc for new numeric cells outside approximate lookup ranges', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-approximate-lookup-new-cell-outside-range',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2) + 0.5)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},1)`)
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', `A${rowCount + 1}`, rowCount + 1)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${rowCount + 1}`)
    expect(inputIndex).toBeDefined()
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex!]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('keeps aggregate formulas current through generic batch clears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-aggregate-clear' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A2' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('updates small direct aggregate fanout without dirty traversal', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-small-direct-aggregate-fanout' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 32; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A32)')

    const sheet = engine.workbook.getSheet('Sheet1')
    expect(sheet).toBeDefined()
    const formulaColumnVersion = sheet!.columnVersions[1] ?? 0
    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'A1', 10)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 537 })
    expect(sheet!.columnVersions[1] ?? 0).toBe(formulaColumnVersion)
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('updates large overlapping aggregate fanout with direct deltas', async () => {
    const rowCount = 256
    const engine = new SpreadsheetEngine({ workbookName: 'operation-large-direct-aggregate-fanout' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'A1', 10)

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount * (rowCount + 1)) / 2 + 9,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('deletes direct aggregate rows without region-query rebuilds or dirty traversal', async () => {
    const rowCount = 256
    const deletedRowIndex = 127
    const deletedValue = deletedRowIndex + 1
    const engine = new SpreadsheetEngine({ workbookName: 'operation-structural-delete-aggregate-kernel-sync-only' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.resetPerformanceCounters()
    engine.deleteRows('Sheet1', deletedRowIndex, 1)

    expect(engine.getCellValue('Sheet1', `B${rowCount - 1}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount * (rowCount + 1)) / 2 - deletedValue,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      kernelSyncOnlyRecalcSkips: 1,
      regionQueryIndexBuilds: 0,
      topoRebuilds: 0,
      wasmFullUploads: 0,
      structuralFormulaImpactCandidates: 0,
      structuralSurvivorCellsRemapped: 0,
    })
  })

  it('uses structural invalidation patches instead of per-cell payloads for direct aggregate row deletes', async () => {
    const rowCount = 256
    const deletedRowIndex = 127
    const engine = new SpreadsheetEngine({ workbookName: 'operation-structural-delete-aggregate-invalidation-patches' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.deleteRows('Sheet1', deletedRowIndex, 1)

    expect(tracked).toHaveBeenCalledTimes(1)
    const event = tracked.mock.calls[0]?.[0]
    expect(event?.patches).toEqual([
      {
        kind: 'row-invalidation',
        sheetName: 'Sheet1',
        startIndex: deletedRowIndex,
        endIndex: deletedRowIndex,
      },
    ])
    expect(engine.getPerformanceCounters()).toMatchObject({
      changedCellPayloadsBuilt: 0,
      kernelSyncOnlyRecalcSkips: 1,
    })
    unsubscribe()
  })

  it('updates sliding aggregate windows without building a region query index', async () => {
    const rowCount = 256
    const window = 32
    const engine = new SpreadsheetEngine({ workbookName: 'operation-sliding-aggregate-no-region-build' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      const endRow = Math.min(rowCount, row + window - 1)
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A${row}:A${endRow})`)
    }

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'A1', 99)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 626 })
    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)
  })

  it('returns typed changed cells for existing numeric direct aggregate mutations', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-typed-direct-aggregate',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 32; row += 1) {
      const endRow = Math.min(64, row + 31)
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A${row}:A${endRow})`)
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'B1')!

    engine.resetPerformanceCounters()
    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 99,
    })

    expect(result).toEqual({
      firstChangedCellIndex: inputIndex,
      secondChangedCellIndex: formulaIndex,
      secondChangedRow: 0,
      secondChangedCol: 1,
      secondChangedNumericValue: 626,
      changedCellCount: 2,
      explicitChangedCount: 1,
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 626 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('can suppress tracked events for direct existing numeric aggregate mutations', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-suppressed-tracked-event',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 32; row += 1) {
      const endRow = Math.min(64, row + 31)
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A${row}:A${endRow})`)
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'B1')!
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 99,
      emitTracked: false,
    })

    expect(result).toEqual({
      firstChangedCellIndex: inputIndex,
      secondChangedCellIndex: formulaIndex,
      secondChangedRow: 0,
      secondChangedCol: 1,
      secondChangedNumericValue: 626,
      changedCellCount: 2,
      explicitChangedCount: 1,
    })
    expect(tracked).not.toHaveBeenCalled()
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 626 })
    unsubscribe()
  })

  it('returns typed changed cells for trusted existing numeric leaf formula mutations', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-leaf-formula',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 100)
    engine.setCellValue('Sheet1', 'B1', 20)
    engine.setCellFormula('Sheet1', 'D1', 'A1+B1*2')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'D1')!

    engine.resetPerformanceCounters()
    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 101,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 100,
    })

    expect(result).toEqual({
      firstChangedCellIndex: inputIndex,
      secondChangedCellIndex: formulaIndex,
      secondChangedRow: 0,
      secondChangedCol: 3,
      secondChangedNumericValue: 141,
      changedCellCount: 2,
      explicitChangedCount: 1,
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 141 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 1 })
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
  })

  it('omits unchanged trusted leaf formulas from compact numeric mutation results', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-unchanged-leaf-formula',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'IF(A1>0,"yes","no")')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!

    engine.resetPerformanceCounters()
    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 2,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
    })

    expect(result).toEqual({
      firstChangedCellIndex: inputIndex,
      changedCellCount: 1,
      explicitChangedCount: 1,
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({ tag: ValueTag.String, value: 'yes' })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 1 })
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
  })

  it('rejects trusted direct aggregate numeric mutations when lookup dependents share the input column', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-aggregate-lookup-guard',
      trackReplicaVersions: false,
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 32; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A32)')
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A32,0)')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!

    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 99,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
    })

    expect(result).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 528 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('returns typed changed cells for trusted existing numeric direct scalar chains', async () => {
    const downstreamCount = 24
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-direct-scalar-chain',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = offset
      engine.setCellFormula('Sheet1', `${indexToColumn(col)}1`, `${indexToColumn(col - 1)}1+1`)
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const terminalIndex = engine.workbook.getCellIndex('Sheet1', `${indexToColumn(downstreamCount)}1`)!

    engine.resetPerformanceCounters()
    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 99,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
    })

    expect(result?.explicitChangedCount).toBe(1)
    expect(result?.changedCellIndices?.length).toBe(downstreamCount + 1)
    expect(result?.changedCellIndices?.[0]).toBe(inputIndex)
    expect(result?.changedCellIndices?.[downstreamCount]).toBe(terminalIndex)
    expect(engine.getCellValue('Sheet1', `${indexToColumn(downstreamCount)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 99 + downstreamCount,
    })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(downstreamCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('applies overlapping sliding aggregate deltas on the single-literal fast path', async () => {
    const rowCount = 96
    const window = 16
    const updateRow = 16
    const engine = new SpreadsheetEngine({ workbookName: 'operation-sliding-aggregate-direct-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      const endRow = Math.min(rowCount, row + window - 1)
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A${row}:A${endRow})`)
    }
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', `A${updateRow}`, 99)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 219 })
    expect(engine.getCellValue('Sheet1', `B${updateRow}`)).toEqual({ tag: ValueTag.Number, value: 459 })
    expect(engine.getCellValue('Sheet1', `B${updateRow + 1}`)).toEqual({ tag: ValueTag.Number, value: 392 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(window)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)

    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${updateRow}`)
    expect(inputIndex).toBeDefined()
    const event = tracked.mock.calls.at(-1)?.[0]
    expect(event).toEqual(expect.objectContaining({ explicitChangedCount: 1 }))
    const changedIndices = Array.from(event.changedCellIndices)
    const expectedFormulaIndices = Array.from({ length: window }, (_, index) => {
      const cellIndex = engine.workbook.getCellIndex('Sheet1', `B${index + 1}`)
      expect(cellIndex).toBeDefined()
      return cellIndex!
    })
    expect(changedIndices[0]).toBe(inputIndex)
    expect(changedIndices.slice(1).toSorted((left, right) => left - right)).toEqual(
      expectedFormulaIndices.toSorted((left, right) => left - right),
    )
    unsubscribe()
  })

  it('accumulates direct aggregate deltas across generic batch literal writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-aggregate-batch-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 32; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A32)')

    const batch = createBatch(getReplicaState(engine), [
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 10 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A2', value: 20 },
    ])

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 555 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('counts direct scalar generic batch updates without dirty traversal', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-batch-metrics' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 5 }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('updates dense same-column affine scalar batches without dirty traversal', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-affine-column-batch', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*2`)
    }
    const refs: EngineCellMutationRef[] = Array.from({ length: rowCount }, (_, index) => ({
      sheetId,
      cellIndex: engine.workbook.getCellIndex('Sheet1', `A${index + 1}`)!,
      mutation: { kind: 'setCellValue', row: index, col: 0, value: index * 3 },
    }))
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, null, 'local', 0))

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: (rowCount - 1) * 6 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(changed).toHaveLength(rowCount * 2)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(rowCount)
    unsubscribe()
  })

  it('updates descending dense affine scalar undo batches without dirty traversal', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-affine-column-undo-batch', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row * 3)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*2`)
    }
    const refs: EngineCellMutationRef[] = Array.from({ length: rowCount }, (_, index) => {
      const row = rowCount - index - 1
      return {
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 1 },
      }
    })
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, null, 'undo', 0))

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount * 2 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(changed).toHaveLength(rowCount * 2)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(rowCount)
    unsubscribe()
  })

  it('updates dense row-pair simple scalar batches without dirty traversal', async () => {
    const rowCount = 48
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-row-pair-batch', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row + 1)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `A${row}*B${row}`)
    }
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 0, value: row * 3 },
      })
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `B${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 1, value: row * 5 },
      })
    }
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, null, 'local', 0))

    expect(engine.getCellValue('Sheet1', `C${rowCount}`)).toEqual({ tag: ValueTag.Number, value: (rowCount - 1) * 8 })
    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 3 * ((rowCount - 1) * 5),
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount * 2)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(changed).toHaveLength(rowCount * 4)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(rowCount * 2)
    unsubscribe()
  })

  it('accumulates cell-by-cell direct scalar deltas across same-row batch writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-cell-product-batch-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    engine.setCellFormula('Sheet1', 'D1', 'A1*B1')

    const batch = createBatch(getReplicaState(engine), [
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 5 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 7 },
    ])

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 35 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(2)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('propagates simple direct scalar chains with numeric deltas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-chain-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet1', 'C1', 'B1+1')
    engine.setCellFormula('Sheet1', 'D1', 'C1+1')

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'A1', 5)

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(3)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('updates large direct scalar fanout with constant bulk deltas', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-bulk-deltas', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellFormula('Sheet1', `B${row}`, 'A1+1')
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const outputColumnVersionBefore = engine.workbook.getSheetById(sheetId)?.columnVersions[1] ?? 0
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.applyCellMutationsAtWithOptions(
      [
        {
          sheetId,
          cellIndex: engine.workbook.getCellIndex('Sheet1', 'A1')!,
          mutation: { kind: 'setCellValue', row: 0, col: 0, value: 5 },
        },
      ],
      { captureUndo: true, potentialNewCells: 0, source: 'local', returnUndoOps: false, reuseRefs: true },
    )

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.workbook.getSheetById(sheetId)?.columnVersions[1] ?? 0).toBe(outputColumnVersionBefore)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(1)
    unsubscribe()
  })

  it('keeps mixed direct scalar and aggregate fanout on constant delta storage', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-aggregate-mixed-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `=$A$1+${row}`)
      engine.setCellFormula('Sheet1', `C${row}`, `=SUM(A1:A${row})`)
    }
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'A1', 99)

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: 99 + rowCount })
    expect(engine.getCellValue('Sheet1', `C${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount * (rowCount + 1)) / 2 + 98,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)

    const event = tracked.mock.calls.at(-1)?.[0]
    const changedIndices = Array.from(event.changedCellIndices)
    expect(event).toEqual(expect.objectContaining({ explicitChangedCount: 1 }))
    expect(changedIndices[0]).toBe(engine.workbook.getCellIndex('Sheet1', 'A1'))
    expect(changedIndices).toContain(engine.workbook.getCellIndex('Sheet1', `B${rowCount}`))
    expect(changedIndices).toContain(engine.workbook.getCellIndex('Sheet1', `C${rowCount}`))
    unsubscribe()
  })

  it('updates copied SUMIF formulas from aggregate column writes with direct deltas', async () => {
    const formulaCount = 32
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-criteria-sum-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Group')
    engine.setCellValue('Sheet1', 'B1', 'Value')
    engine.setCellValue('Sheet1', 'D1', 'A')
    for (let row = 2; row <= 9; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row % 2 === 0 ? 'A' : 'B')
      engine.setCellValue('Sheet1', `B${row}`, row)
    }
    for (let index = 0; index < formulaCount; index += 1) {
      engine.setCellFormula('Sheet1', `${indexToColumn(4 + index)}1`, '=SUMIF(A2:A9,D1,B2:B9)')
    }

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'B2', 100)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 118 })
    expect(engine.getCellValue('Sheet1', `${indexToColumn(4 + formulaCount - 1)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 118,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(formulaCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('replaces existing formulas with generic batch literal writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-literal-over-formula' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 9 }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getCell('Sheet1', 'B1').formula).toBeUndefined()
  })

  it('treats generic batch clears of missing cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-clear-missing-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'G7' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('keeps lookup and aggregate dependents current through generic batch clears', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-batch-clear-lookup-and-aggregate',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 30)
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellValue('Sheet1', 'D2', 25)
    engine.setCellFormula('Sheet1', 'E1', 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D2,A1:A3,1)')
    engine.setCellFormula('Sheet1', 'G1', 'SUM(A1:A3)')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A2' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 40 })
  })

  it('applies local cell mutation refs through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-local-refs', replicaId: 'a' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellValue', row: 0, col: 0, value: 10 },
      },
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' },
      },
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'SUM(' },
      },
      {
        sheetId,
        mutation: { kind: 'clearCell', row: 3, col: 3 },
      },
      {
        sheetId,
        mutation: { kind: 'clearCell', row: 0, col: 0 },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 3))

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Empty })
  })

  it('rejects local cell mutation refs for unknown sheets', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-local-refs-missing',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const refs: EngineCellMutationRef[] = [
      {
        sheetId: 999,
        mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 },
      },
    ]
    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }])

    expect(() => Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))).toThrow('Unknown sheet id: 999')
  })
})
