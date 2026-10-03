import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { cloneWorkPaperClipboardPayload, createWorkPaperClipboardPayload } from '../work-paper-clipboard.js'

describe('work paper clipboard helpers', () => {
  it('creates payloads with a copied source anchor', () => {
    const sourceAnchor = { sheet: 1, row: 2, col: 3 }
    const payload = createWorkPaperClipboardPayload({
      sourceAnchor,
      serialized: [[1, '=A1']],
      values: [[{ tag: ValueTag.Number, value: 1 }]],
    })

    expect(payload.sourceAnchor).toEqual(sourceAnchor)
    expect(payload.sourceAnchor).not.toBe(sourceAnchor)
  })

  it('deep-clones clipboard payloads for rebuild handoff', () => {
    const payload = createWorkPaperClipboardPayload({
      sourceAnchor: { sheet: 1, row: 2, col: 3 },
      serialized: [[1, '=A1']],
      values: [[{ tag: ValueTag.String, value: 'x', stringId: 4 }]],
    })
    const cloned = cloneWorkPaperClipboardPayload(payload)

    expect(cloned).toEqual(payload)
    expect(cloned).not.toBe(payload)
    expect(cloned?.sourceAnchor).not.toBe(payload.sourceAnchor)
    expect(cloned?.serialized).not.toBe(payload.serialized)
    expect(cloned?.serialized[0]).not.toBe(payload.serialized[0])
    expect(cloned?.values).not.toBe(payload.values)
    expect(cloned?.values[0]).not.toBe(payload.values[0])
    expect(cloned?.values[0]?.[0]).not.toBe(payload.values[0]?.[0])
    expect(cloneWorkPaperClipboardPayload(null)).toBeNull()
  })
})
