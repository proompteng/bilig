import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type EngineEvent } from '@bilig/protocol'
import type { AuthoritativeWorkbookEventRecord } from '@bilig/zero-sync'

import { applyAuthoritativeWorkbookEvents } from '../worker-runtime-authoritative-events.js'
import type { WorkerEngine } from '../worker-runtime-support.js'

function authoritativeEngine(): SpreadsheetEngine & WorkerEngine {
  const engine = new SpreadsheetEngine({ workbookName: 'authoritative-events' })
  engine.createSheet('Sheet1')
  engine.setCellValue('Sheet1', 'A1', 1)
  return engine as SpreadsheetEngine & WorkerEngine
}

describe('applyAuthoritativeWorkbookEvents', () => {
  it('validates and applies authoritative events while capturing persistence evidence', () => {
    const engine = authoritativeEngine()
    const event = {
      revision: 1,
      clientMutationId: 'pending-1',
      payload: {
        kind: 'setCellValue',
        sheetName: 'Sheet1',
        address: 'A1',
        value: 7,
      },
    } satisfies AuthoritativeWorkbookEventRecord

    const applied = applyAuthoritativeWorkbookEvents(engine, [event])

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(applied.absorbedMutationIds).toEqual(new Set(['pending-1']))
    expect(applied.payloads).toEqual([event.payload])
    expect(applied.previousSheets).toEqual([{ sheetId: engine.workbook.getSheet('Sheet1')?.id, name: 'Sheet1' }])
    expect(applied.authoritativeEngineEvents).toHaveLength(1)
    expect(applied.authoritativeEngineEvents[0]).toEqual(expect.objectContaining({ kind: 'batch' } satisfies Partial<EngineEvent>))
  })

  it('rejects malformed event batches before mutating the authoritative engine', () => {
    const engine = authoritativeEngine()

    expect(() =>
      applyAuthoritativeWorkbookEvents(engine, [
        {
          revision: 1,
          clientMutationId: 'pending-1',
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
          },
        },
      ]),
    ).toThrow('Invalid authoritative workbook event batch')
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
  })

  it('unsubscribes from authoritative engine events after applying the batch', () => {
    const engine = authoritativeEngine()
    const applied = applyAuthoritativeWorkbookEvents(engine, [
      {
        revision: 1,
        clientMutationId: null,
        payload: {
          kind: 'setCellValue',
          sheetName: 'Sheet1',
          address: 'A1',
          value: 7,
        },
      },
    ])

    engine.setCellValue('Sheet1', 'B1', 9)

    expect(applied.authoritativeEngineEvents).toHaveLength(1)
  })
})
