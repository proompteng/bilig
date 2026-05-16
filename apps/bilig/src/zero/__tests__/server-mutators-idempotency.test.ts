import { beforeEach, describe, expect, it, vi } from 'vitest'
import { WorkbookRuntimeManager } from '../../workbook-runtime/runtime-manager.js'
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
    changeStoreFns.loadWorkbookChange.mockClear()
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
})

function createQueryable(): Queryable {
  return {
    query: vi.fn(async () => ({ rows: [] })),
  }
}

function createServerTransaction(db: Queryable): unknown {
  return {
    dbTransaction: {
      wrappedTransaction: db,
    },
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
