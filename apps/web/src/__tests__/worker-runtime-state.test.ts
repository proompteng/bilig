import { describe, expect, it } from 'vitest'
import {
  EMPTY_RUNTIME_METRICS,
  buildCachedWorkerRuntimeState,
  buildWorkerRuntimeStateFromBootstrap,
  cloneWorkerRuntimeState,
  listOrderedSheetNames,
  normalizeWorkerRuntimeStateSnapshot,
  withExternalSyncState,
} from '../worker-runtime-state.js'

describe('worker runtime state helpers', () => {
  it('orders sheet names by workbook order', () => {
    expect(
      listOrderedSheetNames({
        sheetsByName: new Map([
          ['c', { id: 3, name: 'Sheet3', order: 2 }],
          ['a', { id: 1, name: 'Sheet1', order: 0 }],
          ['b', { id: 2, name: 'Sheet2', order: 1 }],
        ]),
      }),
    ).toEqual(['Sheet1', 'Sheet2', 'Sheet3'])
  })

  it('clones runtime state and applies external sync overrides without mutating the cache copy', () => {
    const cachedState = cloneWorkerRuntimeState({
      workbookName: 'Book',
      sheetNames: ['Sheet1'],
      definedNames: [
        {
          name: 'TaxRate',
          value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'B2' },
        },
      ],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: 'local',
      localPersistenceMode: 'ephemeral',
    })

    const publicState = withExternalSyncState(cachedState, 'syncing')

    expect(publicState.syncState).toBe('syncing')
    expect(cachedState.syncState).toBe('local')
    expect(publicState.localHistoryState).toEqual({ canUndo: false, canRedo: false })
    expect(publicState.localPersistenceMode).toBe('ephemeral')
    expect(publicState.sheets).toEqual([{ id: 1, name: 'Sheet1', order: 0 }])
    expect(publicState.sheets).not.toBe(cachedState.sheets)
    expect(publicState.metrics).not.toBe(cachedState.metrics)
    expect(publicState.definedNames).not.toBe(cachedState.definedNames)
    expect(publicState.definedNames).toEqual(cachedState.definedNames)
  })

  it('derives sheet names from sheet identities when structured identities are present', () => {
    const state = cloneWorkerRuntimeState({
      workbookName: 'Book',
      sheets: [
        { id: 9, name: 'Archive', order: 1 },
        { id: 42, name: 'Actuals', order: 0 },
      ],
      sheetNames: ['StaleName'],
      definedNames: [],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: 'local',
    })

    expect(state.sheets).toEqual([
      { id: 42, name: 'Actuals', order: 0 },
      { id: 9, name: 'Archive', order: 1 },
    ])
    expect(state.sheetNames).toEqual(['Actuals', 'Archive'])
  })

  it('normalizes trusted runtime-state payloads at app boundaries', () => {
    const state = normalizeWorkerRuntimeStateSnapshot({
      workbookName: 'Book',
      sheets: [
        { id: 9, name: 'Archive', order: 1 },
        { id: 42, name: 'Actuals', order: 0 },
      ],
      sheetNames: ['StaleName'],
      definedNames: [],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: 'live',
      localHistoryState: { canUndo: true, canRedo: false },
      pendingMutationSummary: {
        activeCount: 1,
        failedCount: 1,
        firstFailed: {
          id: 'mutation-1',
          method: 'setCellValue',
          failureMessage: 'network unavailable',
          attemptCount: 2,
        },
      },
      localPersistenceMode: 'ephemeral',
    })

    expect(state?.sheetNames).toEqual(['Actuals', 'Archive'])
    expect(state?.pendingMutationSummary).toEqual({
      activeCount: 1,
      failedCount: 1,
      firstFailed: {
        id: 'mutation-1',
        method: 'setCellValue',
        failureMessage: 'network unavailable',
        attemptCount: 2,
      },
    })
  })

  it('rejects impossible runtime-state payloads at app boundaries', () => {
    const validState = {
      workbookName: 'Book',
      sheetNames: ['Sheet1'],
      definedNames: [],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: 'syncing',
      localHistoryState: { canUndo: false, canRedo: false },
    }

    expect(normalizeWorkerRuntimeStateSnapshot({ ...validState, syncState: 'offline' })).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        pendingMutationSummary: { activeCount: -1, failedCount: 0, firstFailed: null },
      }),
    ).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        pendingMutationSummary: { activeCount: 1, failedCount: 1, firstFailed: { id: 'm1', method: 'setCellValue' } },
      }),
    ).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        pendingMutationSummary: { activeCount: 1, failedCount: 2, firstFailed: null },
      }),
    ).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        pendingMutationSummary: {
          activeCount: 0,
          failedCount: 0,
          firstFailed: {
            id: 'm1',
            method: 'setCellValue',
            failureMessage: 'failed',
            attemptCount: 1,
          },
        },
      }),
    ).toBeNull()
    expect(normalizeWorkerRuntimeStateSnapshot({ ...validState, authoritativeRevision: -1 })).toBeNull()
    expect(normalizeWorkerRuntimeStateSnapshot({ ...validState, authoritativeRevision: Number.NaN })).toBeNull()
    expect(normalizeWorkerRuntimeStateSnapshot({ ...validState, metrics: { ...EMPTY_RUNTIME_METRICS, recalcMs: 'slow' } })).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({ ...validState, metrics: { ...EMPTY_RUNTIME_METRICS, jsFormulaCount: Number.NaN } }),
    ).toBeNull()
    expect(normalizeWorkerRuntimeStateSnapshot({ ...validState, sheetNames: ['Sheet1', 'Sheet1'] })).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        sheets: [
          { id: 1, name: 'Sheet1', order: 0 },
          { id: 1, name: 'Archive', order: 1 },
        ],
      }),
    ).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        sheets: [
          { id: 1, name: 'Sheet1', order: 0 },
          { id: 2, name: 'Sheet1', order: 1 },
        ],
      }),
    ).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        sheets: [{ id: 1, name: 'Sheet1', order: -1 }],
      }),
    ).toBeNull()
    expect(normalizeWorkerRuntimeStateSnapshot({ ...validState, definedNames: [{ name: '', value: 1 }] })).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        definedNames: [
          { name: 'Rate', value: 0.1 },
          { name: ' rate ', value: 0.2 },
        ],
      }),
    ).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        definedNames: [{ name: 'BadScalar', value: { kind: 'scalar', value: Number.NaN } }],
      }),
    ).toBeNull()
    expect(
      normalizeWorkerRuntimeStateSnapshot({
        ...validState,
        definedNames: [{ name: 'BrokenRange', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1' } }],
      }),
    ).toBeNull()
  })

  it('builds public cached runtime state without rebuilding sheet identities from names', () => {
    const cachedState = {
      workbookName: 'Book',
      sheets: [
        { id: 42, name: 'Actuals', order: 0 },
        { id: 9, name: 'Archive', order: 1 },
      ],
      sheetNames: ['StaleName'],
      definedNames: [],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: 'local',
    }

    const publicState = buildCachedWorkerRuntimeState({
      cachedState,
      externalSyncState: 'syncing',
      localHistoryState: { canUndo: true, canRedo: false },
      authoritativeRevision: 12,
      pendingMutationSummary: { activeCount: 0, failedCount: 0, firstFailed: null },
      localPersistenceMode: 'ephemeral',
    })

    expect(publicState.sheets).toEqual([
      { id: 42, name: 'Actuals', order: 0 },
      { id: 9, name: 'Archive', order: 1 },
    ])
    expect(publicState.sheetNames).toEqual(['Actuals', 'Archive'])
    expect(publicState.syncState).toBe('syncing')
    expect(publicState.localHistoryState).toEqual({ canUndo: true, canRedo: false })
    expect(publicState.authoritativeRevision).toBe(12)
    expect(publicState.localPersistenceMode).toBe('ephemeral')
  })

  it('builds bootstrap runtime state with empty metrics and syncing status', () => {
    expect(
      buildWorkerRuntimeStateFromBootstrap({
        workbookName: 'Book',
        sheetNames: ['Sheet1'],
        localPersistenceMode: 'ephemeral',
      }),
    ).toEqual({
      workbookName: 'Book',
      sheets: [{ id: 1, name: 'Sheet1', order: 0 }],
      sheetNames: ['Sheet1'],
      definedNames: [],
      metrics: EMPTY_RUNTIME_METRICS,
      syncState: 'syncing',
      localHistoryState: { canUndo: false, canRedo: false },
      localPersistenceMode: 'ephemeral',
    })
  })
})
