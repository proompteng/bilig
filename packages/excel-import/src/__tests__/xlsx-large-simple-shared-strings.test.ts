import { describe, expect, it } from 'vitest'

import { readLargeSimpleReferencedSharedStringsFromChunks } from '../xlsx-large-simple-shared-strings.js'

const encoder = new TextEncoder()

describe('large simple shared string streaming', () => {
  it('discards unreferenced shared-string bodies while waiting for the closing tag', () => {
    const retainedBufferLengths: number[] = []
    const largeUnreferencedText = 'x'.repeat(100_000)
    const chunks = [
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>',
      largeUnreferencedText.slice(0, 40_000),
      largeUnreferencedText.slice(40_000, 80_000),
      largeUnreferencedText.slice(80_000),
      '</t></si><si><t>Keep me</t></si></sst>',
    ]

    const sharedStrings = readLargeSimpleReferencedSharedStringsFromChunks(
      (onChunk) => {
        for (const chunk of chunks) {
          onChunk(encoder.encode(chunk))
        }
        return true
      },
      new Set([1]),
      {
        onRetainedBufferLength: (length) => retainedBufferLengths.push(length),
      },
    )

    expect(sharedStrings?.[1]).toEqual({ text: 'Keep me', rich: false })
    expect(retainedBufferLengths.length).toBeGreaterThan(0)
    expect(Math.max(...retainedBufferLengths)).toBeLessThan(1024)
  })
})
