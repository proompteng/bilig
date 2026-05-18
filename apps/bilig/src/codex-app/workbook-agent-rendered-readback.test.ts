import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { WorkbookAgentRenderedContext } from '@bilig/contracts'
import { selectWorkbookRenderedReadback } from './workbook-agent-rendered-readback.js'

function renderedContext(input: {
  readonly capturedRevision?: number | null
  readonly batchId: number | null
  readonly value?: string
}): WorkbookAgentRenderedContext {
  return {
    capturedAtUnixMs: 1_000,
    capturedRevision: input.capturedRevision ?? null,
    batchId: input.batchId,
    selection: {
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      rowCount: 1,
      columnCount: 1,
      cellCount: 1,
      truncated: false,
      rows: [
        [
          {
            address: 'A1',
            input: input.value ?? 'ok',
            value: { tag: ValueTag.String, value: input.value ?? 'ok' },
            formula: null,
            displayFormat: input.value ?? 'ok',
            styleId: null,
            numberFormatId: null,
            style: null,
          },
        ],
      ],
    },
    visibleRange: null,
  }
}

describe('selectWorkbookRenderedReadback', () => {
  it('does not treat renderer batch ids as authoritative workbook revisions', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ batchId: 99 }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 1,
    })

    expect(proof.capturedBatchId).toBe(99)
    expect(proof.capturedRevision).toBeNull()
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
    expect(proof.incompleteReason).toContain('older than the requested verification revision')
  })

  it('requires capturedRevision to satisfy rendered freshness even when batch id is newer', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ capturedRevision: 3, batchId: 100 }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 4,
    })

    expect(proof.capturedBatchId).toBe(100)
    expect(proof.capturedRevision).toBe(3)
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
  })
})
