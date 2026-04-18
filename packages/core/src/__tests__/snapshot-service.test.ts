import { Effect, Exit } from 'effect'
import { describe, expect, it } from 'vitest'
import { FormulaTable } from '../formula-table.js'
import { createReplicaState } from '../replica-state.js'
import { createEngineSnapshotService } from '../engine/services/snapshot-service.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { SpreadsheetEngine } from '../engine.js'
import { createEngineCounters } from '../perf/engine-counters.js'

describe('EngineSnapshotService', () => {
  it('normalizes thrown import failures into tagged snapshot errors', () => {
    const workbook = new WorkbookStore('book')
    const service = createEngineSnapshotService({
      state: {
        workbook,
        strings: new StringPool(),
        formulas: new FormulaTable(workbook.cellStore),
        replicaState: createReplicaState('replica'),
        entityVersions: new Map(),
        sheetDeleteVersions: new Map(),
      },
      getCellByIndex: () => {
        throw new Error('unused')
      },
      resetWorkbook: () => {
        throw new Error('broken')
      },
      executeRestoreTransaction: () => {},
    })

    const exit = Effect.runSyncExit(
      service.importWorkbook({
        version: 1,
        workbook: { name: 'book' },
        sheets: [],
      }),
    )

    expect(Exit.isFailure(exit)).toBe(true)
  })

  it('roundtrips replica tracking maps through the service boundary', () => {
    const workbook = new WorkbookStore('book')
    const state = {
      workbook,
      strings: new StringPool(),
      formulas: new FormulaTable(workbook.cellStore),
      replicaState: createReplicaState('replica'),
      entityVersions: new Map([['cell:1', { counter: 2, replicaId: 'replica', batchId: 'replica:2', opIndex: 0 }]]),
      sheetDeleteVersions: new Map([['Sheet1', { counter: 3, replicaId: 'replica', batchId: 'replica:3', opIndex: 0 }]]),
      counters: createEngineCounters(),
    }
    const service = createEngineSnapshotService({
      state,
      getCellByIndex: () => {
        throw new Error('unused')
      },
      resetWorkbook: () => {},
      executeRestoreTransaction: () => {},
    })

    const exported = Effect.runSync(service.exportReplica())
    state.entityVersions.clear()
    state.sheetDeleteVersions.clear()

    Effect.runSync(service.importReplica(exported))

    expect(state.entityVersions.get('cell:1')).toEqual(exported.entityVersions[0]?.order)
    expect(state.sheetDeleteVersions.get('Sheet1')).toEqual(exported.sheetDeleteVersions[0]?.order)
  })

  it('counts replayed restore ops during snapshot import', () => {
    const workbook = new WorkbookStore('book')
    const counters = createEngineCounters()
    const state = {
      workbook,
      strings: new StringPool(),
      formulas: new FormulaTable(workbook.cellStore),
      replicaState: createReplicaState('replica'),
      entityVersions: new Map(),
      sheetDeleteVersions: new Map(),
      counters,
    }
    const service = createEngineSnapshotService({
      state,
      getCellByIndex: () => {
        throw new Error('unused')
      },
      resetWorkbook: () => {},
      executeRestoreTransaction: () => {},
    })

    Effect.runSync(
      service.importWorkbook({
        version: 1,
        workbook: { name: 'book' },
        sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [{ address: 'A1', value: 1 }] }],
      }),
    )

    expect(counters.snapshotOpsReplayed).toBeGreaterThan(0)
  })

  it('preserves explicit authored blank cells through snapshot import', async () => {
    const snapshot = {
      version: 1 as const,
      workbook: { name: 'snapshot-authored-blank' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'D3', value: null },
            { address: 'E3', formula: 'D3+C3' },
          ],
        },
      ],
    }
    const restored = new SpreadsheetEngine({
      workbookName: snapshot.workbook.name,
      replicaId: 'snapshot-authored-blank-restore',
    })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.exportSnapshot()).toEqual(snapshot)
  })
})
