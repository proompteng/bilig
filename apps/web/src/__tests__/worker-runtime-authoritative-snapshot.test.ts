import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import type { PendingWorkbookMutation } from '../workbook-sync.js'
import { prepareAuthoritativeSnapshotProjection } from '../worker-runtime-authoritative-snapshot.js'

function buildSnapshot(value: number): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'authoritative-snapshot-doc' },
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        cells: [{ address: 'A1', value }],
      },
    ],
  }
}

function buildMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'authoritative-snapshot-doc:pending:1',
    localSeq: 1,
    baseRevision: 0,
    method: 'setCellValue',
    args: ['Sheet1', 'A1', 17],
    enqueuedAtUnixMs: 100,
    submittedAtUnixMs: null,
    lastAttemptedAtUnixMs: null,
    ackedAtUnixMs: null,
    rebasedAtUnixMs: null,
    failedAtUnixMs: null,
    attemptCount: 0,
    failureMessage: null,
    status: 'local',
    ...overrides,
  }
}

describe('worker runtime authoritative snapshot projection', () => {
  it('uses a single projection engine as authority during clean bootstrap installs', async () => {
    const importSnapshot = vi.spyOn(SpreadsheetEngine.prototype, 'importSnapshot')

    const prepared = await prepareAuthoritativeSnapshotProjection({
      documentId: 'clean-bootstrap-doc',
      replicaId: 'browser:test',
      snapshot: buildSnapshot(42),
      mode: 'bootstrap',
      pendingMutations: [],
    })

    expect(importSnapshot).toHaveBeenCalledTimes(1)
    expect(prepared.authoritativeEngine).toBeNull()
    expect(prepared.authoritativeReplica).not.toBeNull()
    expect(prepared.projectionOverlayScope).toBeNull()
    expect(prepared.shouldMarkPendingMutationsRebased).toBe(false)
    expect(prepared.projectionEngine.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })
  })

  it('keeps a dedicated authoritative engine and overlays pending edits during reconcile installs', async () => {
    const prepared = await prepareAuthoritativeSnapshotProjection({
      documentId: 'reconcile-doc',
      replicaId: 'browser:test',
      snapshot: buildSnapshot(5),
      mode: 'reconcile',
      pendingMutations: [buildMutation()],
    })

    expect(prepared.authoritativeEngine).not.toBeNull()
    expect(prepared.authoritativeEngine?.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(prepared.projectionEngine.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })
    expect(prepared.projectionOverlayScope).not.toBeNull()
    expect(prepared.shouldMarkPendingMutationsRebased).toBe(true)
  })
})
