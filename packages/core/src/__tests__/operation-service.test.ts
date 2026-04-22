import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineOperationService } from '../engine/services/operation-service.js'
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

function getReplicaState(engine: SpreadsheetEngine) {
  const replicaState = Reflect.get(engine, 'replicaState')
  if (typeof replicaState !== 'object' || replicaState === null) {
    throw new TypeError('Expected engine replica state')
  }
  return replicaState
}

describe('EngineOperationService', () => {
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

    expect(setLastMetricsSpy).toHaveBeenCalledTimes(3)
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

    engine.setCellValue('Sheet1', 'A1', 10)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 537 })
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

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 555 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
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
