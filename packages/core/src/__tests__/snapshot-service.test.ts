import { Effect, Exit } from 'effect'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { FormulaTable } from '../formula-table.js'
import { createReplicaState } from '../replica-state.js'
import { createEngineSnapshotService } from '../engine/services/snapshot-service.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { SpreadsheetEngine } from '../engine.js'
import { createEngineCounters } from '../perf/engine-counters.js'
import { readRuntimeImage, readRuntimeSnapshot } from '../snapshot/runtime-image-codec.js'
import { CellFlags } from '../cell-store.js'

describe('EngineSnapshotService', () => {
  it('ignores missing runtime image carriers', () => {
    expect(readRuntimeImage(undefined)).toBeUndefined()
    expect(readRuntimeSnapshot(null)).toBeUndefined()
  })

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
      initializeCellFormulasAt: () => {
        throw new Error('unused')
      },
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
      initializeCellFormulasAt: () => {
        throw new Error('unused')
      },
    })

    const exported = Effect.runSync(service.exportReplica())
    state.entityVersions.clear()
    state.sheetDeleteVersions.clear()

    Effect.runSync(service.importReplica(exported))

    expect(state.entityVersions.get('cell:1')).toEqual(exported.entityVersions[0]?.order)
    expect(state.sheetDeleteVersions.get('Sheet1')).toEqual(exported.sheetDeleteVersions[0]?.order)
  })

  it('does not replay restore ops during raw snapshot import', () => {
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
      resetWorkbook: (workbookName) => {
        workbook.reset(workbookName)
      },
      initializeCellFormulasAt: () => {
        throw new Error('unexpected formula initialization for literal-only snapshot')
      },
    })

    Effect.runSync(
      service.importWorkbook({
        version: 1,
        workbook: { name: 'book' },
        sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [{ address: 'A1', value: 1 }] }],
      }),
    )

    expect(counters.snapshotOpsReplayed).toBe(0)
  })

  it('restores raw snapshots directly when coordinate formula initialization is available', () => {
    const workbook = new WorkbookStore('book')
    const strings = new StringPool()
    const counters = createEngineCounters()
    const state = {
      workbook,
      strings,
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
      resetWorkbook: (workbookName) => {
        workbook.reset(workbookName)
      },
      initializeCellFormulasAt: () => {
        throw new Error('unexpected formula initialization for literal-only snapshot')
      },
    })

    Effect.runSync(
      service.importWorkbook({
        version: 1,
        workbook: { name: 'direct-restore' },
        sheets: [
          {
            id: 9,
            name: 'Sheet1',
            order: 0,
            cells: [
              { address: 'A1', value: 12 },
              { address: 'B2', value: 'segment-1', format: '0.00' },
              { address: 'C3', value: null },
            ],
          },
        ],
      }),
    )

    const sheetId = workbook.getSheet('Sheet1')?.id
    expect(workbook.workbookName).toBe('direct-restore')
    expect(sheetId).toBe(9)
    const a1 = workbook.getCellIndexAt(sheetId!, 0, 0)
    const b2 = workbook.getCellIndexAt(sheetId!, 1, 1)
    const c3 = workbook.getCellIndexAt(sheetId!, 2, 2)
    expect(a1).toBeDefined()
    expect(b2).toBeDefined()
    expect(c3).toBeDefined()
    expect(workbook.cellStore.getValue(a1!, (id) => strings.get(id))).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(workbook.cellStore.getValue(b2!, (id) => strings.get(id))).toEqual({
      tag: ValueTag.String,
      value: 'segment-1',
      stringId: 1,
    })
    expect(workbook.getCellFormat(b2!)).toBe('0.00')
    expect(workbook.cellStore.flags[c3!] & CellFlags.AuthoredBlank).toBe(CellFlags.AuthoredBlank)
    expect(counters.snapshotOpsReplayed).toBe(0)
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

  it('restores structurally rewritten direct scalar refs as #REF errors', async () => {
    const source = new SpreadsheetEngine({
      workbookName: 'snapshot-direct-scalar-ref-rewrite-source',
      replicaId: 'snapshot-direct-scalar-ref-rewrite-source',
    })
    await source.ready()
    source.createSheet('Sheet1')
    source.setCellValue('Sheet1', 'A1', 1)
    source.setCellFormula('Sheet1', 'B2', 'A1*2')

    source.deleteRows('Sheet1', 0, 1)

    const snapshot = source.exportSnapshot()
    const restored = new SpreadsheetEngine({
      workbookName: 'snapshot-direct-scalar-ref-rewrite-restored',
      replicaId: 'snapshot-direct-scalar-ref-rewrite-restored',
    })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getCell('Sheet1', 'B1').formula).toBe('#REF!*2')
    expect(restored.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
  })

  it('attaches a runtime image to exported snapshots and restores without replay when reused in-process', async () => {
    const source = new SpreadsheetEngine({
      workbookName: 'snapshot-runtime-image-source',
      replicaId: 'snapshot-runtime-image-source',
    })
    await source.ready()
    source.createSheet('Sheet1')
    source.setCellValue('Sheet1', 'A1', 4)
    source.setCellValue('Sheet1', 'B1', 7)
    source.setCellFormula('Sheet1', 'C1', 'A1+B1')
    source.setCellFormula('Sheet1', 'D1', 'C1*2')

    const snapshot = source.exportSnapshot()
    const runtimeImage = readRuntimeImage(snapshot)
    expect(runtimeImage).toBeDefined()
    expect(runtimeImage?.templateBank.length).toBeGreaterThan(0)
    expect(runtimeImage?.formulaInstances.length).toBe(2)
    expect(runtimeImage?.sheetCells).toEqual([
      {
        sheetName: 'Sheet1',
        coords: [
          { row: 0, col: 0 },
          { row: 0, col: 1 },
          { row: 0, col: 2 },
          { row: 0, col: 3 },
        ],
      },
    ])

    const restored = new SpreadsheetEngine({
      workbookName: 'snapshot-runtime-image-restored',
      replicaId: 'snapshot-runtime-image-restored',
    })
    await restored.ready()
    restored.resetPerformanceCounters()
    restored.importSnapshot(snapshot)

    expect(restored.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(restored.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(restored.getPerformanceCounters().snapshotOpsReplayed).toBe(0)
    expect(restored.exportSnapshot()).toEqual(snapshot)
  })
})
