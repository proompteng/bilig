import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { parseClipboardHtml } from '../gridClipboard.js'

const invalidNumericEntityArbitrary = fc.oneof(
  fc.constant('#'),
  fc.constant('#x'),
  fc.constant('#xzz'),
  fc.constant('#99999999'),
  fc.integer({ min: 0x11_0000, max: 0x1f_ffff }).map((value) => `#x${value.toString(16)}`),
)

describe('grid clipboard fuzz', () => {
  it('should preserve malformed numeric html entities instead of throwing', async () => {
    await runProperty({
      suite: 'grid/clipboard/malformed-html-entities',
      arbitrary: invalidNumericEntityArbitrary,
      predicate: async (entity) => {
        const rawEntity = `&${entity};`
        const parsed = parseClipboardHtml(`<table><tr><td>${rawEntity}</td></tr></table>`)

        expect(parsed).toEqual([[rawEntity]])
      },
      parameters: { numRuns: 80 },
    })
  })
})
