import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { decodeViewportPatch } from '@bilig/worker-transport'
import type { PendingWorkbookMutation } from '../workbook-sync.js'
import { WorkbookWorkerRuntime } from '../worker-runtime.js'

function buildSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'bootstrap-doc' },
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        cells: [
          {
            address: 'A1',
            value: 42,
          },
        ],
      },
    ],
  }
}

function buildPendingMutation(overrides: Partial<PendingWorkbookMutation> = {}): PendingWorkbookMutation {
  return {
    id: 'bootstrap-doc:pending:7',
    localSeq: 7,
    baseRevision: 2,
    method: 'clearCell',
    args: ['Sheet1', 'D10'],
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

function snapshotWithCell(address: string, value: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'bootstrap-doc' },
    sheets: [
      {
        name: 'Sheet1',
        order: 0,
        cells: [
          {
            address,
            value,
          },
        ],
      },
    ],
  }
}

describe('worker runtime authoritative bootstrap', () => {
  it('imports a bootstrap authoritative snapshot once for the projection state', async () => {
    const importSnapshot = vi.spyOn(SpreadsheetEngine.prototype, 'importSnapshot')
    const runtime = new WorkbookWorkerRuntime()

    await runtime.bootstrap({
      documentId: 'single-import-bootstrap-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
        initialPatch: 'none',
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    await runtime.installAuthoritativeSnapshot({
      snapshot: buildSnapshot(),
      authoritativeRevision: 3,
      mode: 'bootstrap',
    })

    expect(importSnapshot).toHaveBeenCalledTimes(1)
    expect(runtime.getAuthoritativeRevision()).toBe(3)
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })
    expect(received).toHaveLength(1)
    expect(received[0]?.full).toBe(true)
    expect(received[0]?.authoritativeRevision).toBe(3)
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'A1')?.displayText).toBe('42')

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 4,
          clientMutationId: null,
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'A1',
            value: 84,
          },
        },
      ],
      4,
    )

    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 84,
    })
  })

  it('replays restored pending clears over stale authoritative bootstrap snapshots', async () => {
    const runtime = new WorkbookWorkerRuntime()

    await runtime.bootstrap({
      documentId: 'bootstrap-doc',
      replicaId: 'browser:test',
      persistState: true,
      mutationJournalEntries: [buildPendingMutation({ status: 'submitted', submittedAtUnixMs: 200 })],
      nextPendingMutationSeq: 8,
    })

    await runtime.installAuthoritativeSnapshot({
      snapshot: snapshotWithCell('D10', 'delete-clear-viewport-reload'),
      authoritativeRevision: 3,
      mode: 'bootstrap',
    })

    expect(runtime.getCell('Sheet1', 'D10').value).toEqual({ tag: ValueTag.Empty })
    expect(runtime.listPendingMutations().map((mutation) => mutation.id)).toEqual(['bootstrap-doc:pending:7'])

    const nextMutation = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'E10', 'after-reload'],
    })
    expect(nextMutation.id).toBe('bootstrap-doc:browser:test:pending:8')
  })

  it('honors restored mutation high-water marks even when no pending entries remain', async () => {
    const runtime = new WorkbookWorkerRuntime()

    await runtime.bootstrap({
      documentId: 'bootstrap-doc',
      replicaId: 'browser:test',
      persistState: true,
      mutationJournalEntries: [],
      nextPendingMutationSeq: 12,
    })

    const nextMutation = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'E10', 'after-acked-reload'],
    })

    expect(nextMutation.id).toBe('bootstrap-doc:browser:test:pending:12')
  })
})
