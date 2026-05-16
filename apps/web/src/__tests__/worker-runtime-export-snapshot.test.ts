import { describe, expect, it, vi } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'

import { exportWorkerRuntimeSnapshot } from '../worker-runtime-export-snapshot.js'

const snapshot = {
  workbookName: 'Snapshot',
  sheetNames: ['Sheet1'],
  sheets: [],
  definedNames: [],
} satisfies WorkbookSnapshot

describe('exportWorkerRuntimeSnapshot', () => {
  it('exports the installed projection snapshot when the projection engine is materialized', () => {
    const exportProjectionSnapshot = vi.fn(() => snapshot)

    expect(
      exportWorkerRuntimeSnapshot({
        exportProjectionSnapshot,
        getReadyAuthoritativeSnapshot: () => {
          throw new Error('should not read authoritative snapshot')
        },
        pendingMutationCount: 1,
      }),
    ).toBe(snapshot)
    expect(exportProjectionSnapshot).toHaveBeenCalledTimes(1)
  })

  it('exports a ready authoritative snapshot only when no local mutations are pending', () => {
    expect(
      exportWorkerRuntimeSnapshot({
        exportProjectionSnapshot: null,
        getReadyAuthoritativeSnapshot: () => snapshot,
        pendingMutationCount: 0,
      }),
    ).toBe(snapshot)
  })

  it('does not export an authoritative snapshot over pending local mutations', () => {
    expect(() =>
      exportWorkerRuntimeSnapshot({
        exportProjectionSnapshot: null,
        getReadyAuthoritativeSnapshot: () => snapshot,
        pendingMutationCount: 1,
      }),
    ).toThrow('Workbook worker runtime projection snapshot is not ready')
  })

  it('fails closed when neither projection nor authoritative snapshots are ready', () => {
    expect(() =>
      exportWorkerRuntimeSnapshot({
        exportProjectionSnapshot: null,
        getReadyAuthoritativeSnapshot: () => null,
        pendingMutationCount: 0,
      }),
    ).toThrow('Workbook worker runtime projection snapshot is not ready')
  })
})
