import { beforeEach, describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { WorkbookRuntimeManager, type WorkbookRuntime } from '../../workbook-runtime/runtime-manager.js'
import type { WorkbookChangeRecord } from '../workbook-change-store.js'
import type { Queryable } from '../store.js'

const runtimeStoreFns = vi.hoisted(() => ({
  acquireWorkbookMutationLock: vi.fn(async () => undefined),
}))

const mutationStoreFns = vi.hoisted(() => ({
  loadAppliedWorkbookClientMutation: vi.fn(async () => null),
  persistWorkbookMutation: vi.fn(async () => ({
    revision: 12,
    calculatedRevision: 12,
    updatedAt: '2026-05-16T00:00:00.000Z',
    projectionCommit: { kind: 'mock' },
  })),
}))

const changeStoreFns = vi.hoisted(() => ({
  createWorkbookChangeStoreConnection: vi.fn((db: unknown) => db),
  listWorkbookChangesAfterRevision: vi.fn(async () => []),
  loadLatestRedoableWorkbookChange: vi.fn(async () => null),
  loadLatestUndoableWorkbookChange: vi.fn(async () => null),
  loadWorkbookChange: vi.fn(async () => null),
}))

vi.mock('../workbook-runtime-store.js', () => ({
  acquireWorkbookMutationLock: runtimeStoreFns.acquireWorkbookMutationLock,
}))

vi.mock('../workbook-mutation-store.js', () => ({
  loadAppliedWorkbookClientMutation: mutationStoreFns.loadAppliedWorkbookClientMutation,
  persistWorkbookMutation: mutationStoreFns.persistWorkbookMutation,
}))

vi.mock('../workbook-change-store.js', () => ({
  createWorkbookChangeStoreConnection: changeStoreFns.createWorkbookChangeStoreConnection,
  listWorkbookChangesAfterRevision: changeStoreFns.listWorkbookChangesAfterRevision,
  loadLatestRedoableWorkbookChange: changeStoreFns.loadLatestRedoableWorkbookChange,
  loadLatestUndoableWorkbookChange: changeStoreFns.loadLatestUndoableWorkbookChange,
  loadWorkbookChange: changeStoreFns.loadWorkbookChange,
}))

import { handleServerMutator } from '../server-mutators.js'

describe('server mutator client mutation idempotency', () => {
  beforeEach(() => {
    runtimeStoreFns.acquireWorkbookMutationLock.mockClear()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockReset()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockResolvedValue(null)
    mutationStoreFns.persistWorkbookMutation.mockClear()
    changeStoreFns.loadLatestRedoableWorkbookChange.mockClear()
    changeStoreFns.loadLatestUndoableWorkbookChange.mockClear()
    changeStoreFns.listWorkbookChangesAfterRevision.mockClear()
    changeStoreFns.loadWorkbookChange.mockClear()
    changeStoreFns.createWorkbookChangeStoreConnection.mockClear()
  })

  it('treats a matching clientMutationId as a replay without mutating the workbook again', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockResolvedValueOnce({
      documentId: 'doc-1',
      clientMutationId: 'doc-1:pending:7',
      revision: 7,
      createdAt: '2026-05-16T12:00:00.000Z',
      payload: {
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'A1',
        value: 17,
      },
    })

    await handleServerMutator(
      createServerTransaction(db),
      'workbook.setCellValue',
      {
        documentId: 'doc-1',
        clientMutationId: 'doc-1:pending:7',
        sheetName: 'Sheet1',
        address: 'A1',
        value: 17,
      },
      runtimeManager,
    )

    expect(runtimeStoreFns.acquireWorkbookMutationLock).toHaveBeenCalledWith(db, 'doc-1')
    expect(mutationStoreFns.loadAppliedWorkbookClientMutation).toHaveBeenCalledWith(db, 'doc-1', 'doc-1:pending:7')
    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('rejects a reused clientMutationId with a different payload before touching runtime state', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockResolvedValueOnce({
      documentId: 'doc-1',
      clientMutationId: 'doc-1:pending:7',
      revision: 7,
      createdAt: '2026-05-16T12:00:00.000Z',
      payload: {
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'A1',
        value: 17,
      },
    })

    await expect(
      handleServerMutator(
        createServerTransaction(db),
        'workbook.setCellValue',
        {
          documentId: 'doc-1',
          clientMutationId: 'doc-1:pending:7',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 18,
        },
        runtimeManager,
      ),
    ).rejects.toThrow('was already applied with a different payload')

    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('acknowledges replayed revert mutations before rejected already-reverted target validation', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockResolvedValueOnce({
      documentId: 'doc-1',
      clientMutationId: 'doc-1:pending:8',
      revision: 8,
      createdAt: '2026-05-16T12:00:00.000Z',
      payload: {
        kind: 'revertChange',
        targetRevision: 3,
        targetSummary: 'Set A1',
        appliedBundle: {
          kind: 'engineOps',
          ops: [],
        },
      },
    })

    await handleServerMutator(
      createServerTransaction(db),
      'workbook.revertChange',
      {
        documentId: 'doc-1',
        clientMutationId: 'doc-1:pending:8',
        revision: 3,
      },
      runtimeManager,
    )

    expect(changeStoreFns.loadWorkbookChange).not.toHaveBeenCalled()
    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('acknowledges replayed undo-latest mutations before querying the current undo target', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockResolvedValueOnce({
      documentId: 'doc-1',
      clientMutationId: 'doc-1:pending:9',
      revision: 9,
      createdAt: '2026-05-16T12:00:00.000Z',
      payload: {
        kind: 'revertChange',
        targetRevision: 4,
        targetSummary: 'Set B1',
        appliedBundle: {
          kind: 'engineOps',
          ops: [],
        },
      },
    })

    await handleServerMutator(
      createServerTransaction(db),
      'workbook.undoLatestChange',
      {
        documentId: 'doc-1',
        clientMutationId: 'doc-1:pending:9',
      },
      runtimeManager,
    )

    expect(changeStoreFns.loadLatestUndoableWorkbookChange).not.toHaveBeenCalled()
    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('acknowledges replayed redo-latest mutations before querying the current redo target', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockResolvedValueOnce({
      documentId: 'doc-1',
      clientMutationId: 'doc-1:pending:11',
      revision: 11,
      createdAt: '2026-05-16T12:00:00.000Z',
      payload: {
        kind: 'redoChange',
        targetRevision: 4,
        targetSummary: 'Redo B1',
        appliedBundle: {
          kind: 'engineOps',
          ops: [],
        },
      },
    })

    await handleServerMutator(
      createServerTransaction(db),
      'workbook.redoLatestChange',
      {
        documentId: 'doc-1',
        clientMutationId: 'doc-1:pending:11',
      },
      runtimeManager,
    )

    expect(changeStoreFns.loadLatestRedoableWorkbookChange).not.toHaveBeenCalled()
    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('rejects replayed history clientMutationIds when the stored event kind does not match the requested command', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    mutationStoreFns.loadAppliedWorkbookClientMutation.mockResolvedValueOnce({
      documentId: 'doc-1',
      clientMutationId: 'doc-1:pending:10',
      revision: 10,
      createdAt: '2026-05-16T12:00:00.000Z',
      payload: {
        kind: 'redoChange',
        targetRevision: 4,
        targetSummary: 'Redo B1',
        appliedBundle: {
          kind: 'engineOps',
          ops: [],
        },
      },
    })

    await expect(
      handleServerMutator(
        createServerTransaction(db),
        'workbook.undoLatestChange',
        {
          documentId: 'doc-1',
          clientMutationId: 'doc-1:pending:10',
        },
        runtimeManager,
      ),
    ).rejects.toThrow('was already applied with a different payload')

    expect(changeStoreFns.loadLatestUndoableWorkbookChange).not.toHaveBeenCalled()
    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('resolves undo-latest targets only after the workbook mutation lock is held', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    let lockHeld = false
    runtimeStoreFns.acquireWorkbookMutationLock.mockImplementationOnce(async () => {
      lockHeld = true
    })
    changeStoreFns.loadLatestUndoableWorkbookChange.mockImplementationOnce(async () => {
      expect(lockHeld).toBe(true)
      return createWorkbookChangeRecord({
        revision: 4,
        summary: 'Set B1',
        sheetName: 'Sheet1',
        anchorAddress: 'B1',
      })
    })

    await expect(
      handleServerMutator(
        createServerTransaction(db),
        'workbook.undoLatestChange',
        {
          documentId: 'doc-1',
          clientMutationId: 'doc-1:pending:12',
        },
        runtimeManager,
      ),
    ).rejects.toThrow('Duplicate client mutations must not load runtime state')

    expect(mutationStoreFns.loadAppliedWorkbookClientMutation).toHaveBeenCalledWith(db, 'doc-1', 'doc-1:pending:12')
    expect(changeStoreFns.createWorkbookChangeStoreConnection).toHaveBeenCalledWith(
      expect.objectContaining({
        query: expect.any(Function),
        run: expect.any(Function),
      }),
    )
    expect(changeStoreFns.loadLatestUndoableWorkbookChange).toHaveBeenCalledWith(
      expect.objectContaining({
        query: expect.any(Function),
        run: expect.any(Function),
      }),
      {
        documentId: 'doc-1',
        actorUserId: 'system',
      },
    )
    expect(loadRuntime).toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('persists structural range scope in authored history undo events', async () => {
    const db = createQueryable()
    const { loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    const engine = new SpreadsheetEngine({ workbookName: 'doc-1' })
    await engine.ready()
    const applyOps = vi.spyOn(engine, 'applyOps').mockReturnValue([{ kind: 'insertRows', sheetName: 'Sheet1', start: 2, count: 2 }])
    loadRuntime.mockResolvedValueOnce(createWorkbookRuntime(engine))
    changeStoreFns.loadLatestUndoableWorkbookChange.mockResolvedValueOnce(
      createWorkbookChangeRecord({
        revision: 4,
        summary: 'Inserted rows 3:4 on Sheet1',
        sheetName: 'Sheet1',
        anchorAddress: 'A3',
        eventKind: 'insertRows',
        range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'rows' },
      }),
    )

    await handleServerMutator(
      createServerTransaction(db),
      'workbook.undoLatestChange',
      {
        documentId: 'doc-1',
        clientMutationId: 'doc-1:pending:scope',
      },
      runtimeManager,
    )

    expect(applyOps).toHaveBeenCalledWith([], { captureUndo: true })
    expect(mutationStoreFns.persistWorkbookMutation).toHaveBeenCalledWith(
      db,
      'doc-1',
      expect.objectContaining({
        eventPayload: expect.objectContaining({
          kind: 'revertChange',
          targetRevision: 4,
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A3',
            endAddress: 'A4',
            scope: 'rows',
          },
        }),
      }),
    )
  })

  it('validates explicit revert targets only after the workbook mutation lock is held', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    let lockHeld = false
    runtimeStoreFns.acquireWorkbookMutationLock.mockImplementationOnce(async () => {
      lockHeld = true
    })
    changeStoreFns.loadWorkbookChange.mockImplementationOnce(async () => {
      expect(lockHeld).toBe(true)
      return createWorkbookChangeRecord({
        revision: 3,
        summary: 'Set A1',
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        revertedByRevision: 8,
      })
    })

    await expect(
      handleServerMutator(
        createServerTransaction(db),
        'workbook.revertChange',
        {
          documentId: 'doc-1',
          clientMutationId: 'doc-1:pending:13',
          revision: 3,
        },
        runtimeManager,
      ),
    ).rejects.toThrow('Workbook change was already reverted in r8')

    expect(mutationStoreFns.loadAppliedWorkbookClientMutation).toHaveBeenCalledWith(db, 'doc-1', 'doc-1:pending:13')
    expect(changeStoreFns.createWorkbookChangeStoreConnection).toHaveBeenCalledWith(
      expect.objectContaining({
        query: expect.any(Function),
        run: expect.any(Function),
      }),
    )
    expect(changeStoreFns.loadWorkbookChange).toHaveBeenCalledWith(
      expect.objectContaining({
        query: expect.any(Function),
        run: expect.any(Function),
      }),
      'doc-1',
      3,
    )
    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('rejects explicit revert when a later active change overlaps the target range', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    changeStoreFns.loadWorkbookChange.mockResolvedValueOnce(
      createWorkbookChangeRecord({
        revision: 3,
        summary: 'Set A1',
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
      }),
    )
    changeStoreFns.listWorkbookChangesAfterRevision.mockResolvedValueOnce([
      createWorkbookChangeRecord({
        revision: 4,
        summary: 'Set A1 again',
        sheetName: 'Sheet1',
        anchorAddress: 'A1',
        actorUserId: 'morgan@example.com',
      }),
    ])

    await expect(
      handleServerMutator(
        createServerTransaction(db),
        'workbook.revertChange',
        {
          documentId: 'doc-1',
          clientMutationId: 'doc-1:pending:14',
          revision: 3,
        },
        runtimeManager,
      ),
    ).rejects.toThrow('Workbook change cannot be safely reverted after overlapping r4')

    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })

  it('rejects explicit revert when a later structural row change overlaps the target cell', async () => {
    const db = createQueryable()
    const { commitMutation, loadRuntime, runtimeManager } = createRuntimeManagerHarness()
    changeStoreFns.loadWorkbookChange.mockResolvedValueOnce(
      createWorkbookChangeRecord({
        revision: 5,
        summary: 'Set B3',
        sheetName: 'Sheet1',
        anchorAddress: 'B3',
        range: { sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'B3' },
      }),
    )
    changeStoreFns.listWorkbookChangesAfterRevision.mockResolvedValueOnce([
      createWorkbookChangeRecord({
        revision: 6,
        summary: 'Inserted rows 3:4',
        sheetName: 'Sheet1',
        anchorAddress: 'A3',
        actorUserId: 'morgan@example.com',
        eventKind: 'insertRows',
        range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'rows' },
      }),
    ])

    await expect(
      handleServerMutator(
        createServerTransaction(db),
        'workbook.revertChange',
        {
          documentId: 'doc-1',
          clientMutationId: 'doc-1:pending:15',
          revision: 5,
        },
        runtimeManager,
      ),
    ).rejects.toThrow('Workbook change cannot be safely reverted after overlapping r6')

    expect(loadRuntime).not.toHaveBeenCalled()
    expect(commitMutation).not.toHaveBeenCalled()
    expect(mutationStoreFns.persistWorkbookMutation).not.toHaveBeenCalled()
  })
})

function createQueryable(): Queryable {
  return {
    query: vi.fn(async () => ({ rows: [] })),
  }
}

function createServerTransaction(db: Queryable): unknown {
  return {
    run: vi.fn(async () => undefined),
    dbTransaction: {
      wrappedTransaction: db,
    },
  }
}

function createWorkbookChangeRecord(input: {
  readonly revision: number
  readonly summary: string
  readonly sheetName?: string | null
  readonly anchorAddress?: string | null
  readonly actorUserId?: string
  readonly revertedByRevision?: number | null
  readonly revertsRevision?: number | null
  readonly eventKind?: WorkbookChangeRecord['eventKind']
  readonly range?: WorkbookChangeRecord['range']
}): WorkbookChangeRecord {
  return {
    revision: input.revision,
    actorUserId: input.actorUserId ?? 'system',
    clientMutationId: null,
    eventKind: input.eventKind ?? 'setCellValue',
    summary: input.summary,
    sheetId: 1,
    sheetName: input.sheetName ?? null,
    anchorAddress: input.anchorAddress ?? null,
    range: input.range ?? null,
    rangeInvalid: false,
    undoBundle: {
      kind: 'engineOps',
      ops: [],
    },
    revertedByRevision: input.revertedByRevision ?? null,
    revertsRevision: input.revertsRevision ?? null,
    createdAtUnixMs: 1_768_348_800_000,
  }
}

function createWorkbookRuntime(engine: SpreadsheetEngine): WorkbookRuntime {
  const updatedAt = '2026-05-16T12:00:00.000Z'
  return {
    documentId: 'doc-1',
    engine,
    projection: {
      workbook: {
        id: 'doc-1',
        name: 'doc-1',
        ownerUserId: 'owner-1',
        headRevision: 4,
        calculatedRevision: 4,
        calcMode: 'automatic',
        compatibilityMode: 'excel-modern',
        recalcEpoch: 0,
        updatedAt,
      },
      sheets: [],
      cells: [],
      rowMetadata: [],
      columnMetadata: [],
      definedNames: [],
      workbookMetadataEntries: [],
      calculationSettings: {
        workbookId: 'doc-1',
        mode: 'automatic',
        recalcEpoch: 0,
      },
      styles: [],
      numberFormats: [],
    },
    headRevision: 4,
    calculatedRevision: 4,
    ownerUserId: 'owner-1',
  }
}

function createRuntimeManagerHarness() {
  const runtimeManager = new WorkbookRuntimeManager({
    loadMetadata: async () => ({
      headRevision: 0,
      calculatedRevision: 0,
      ownerUserId: 'owner-1',
    }),
    loadState: async () => {
      throw new Error('Duplicate client mutations must not load runtime state')
    },
    createEngine: async () => {
      throw new Error('Duplicate client mutations must not create workbook engines')
    },
  })
  const loadRuntime = vi
    .spyOn(runtimeManager, 'loadRuntime')
    .mockRejectedValue(new Error('Duplicate client mutations must not load runtime state'))
  const commitMutation = vi.spyOn(runtimeManager, 'commitMutation')
  return { runtimeManager, loadRuntime, commitMutation }
}
