import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { decodeViewportPatch } from '@bilig/worker-transport'
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
})
