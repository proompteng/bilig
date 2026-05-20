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

  it('stores sparse high shared-string references without a giant array', () => {
    const sharedStrings = readLargeSimpleReferencedSharedStringsFromChunks(
      (onChunk) => {
        onChunk(
          encoder.encode(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Zero</t></si><si><t>One</t></si></sst>',
          ),
        )
        return true
      },
      new Set([1]),
    )

    expect(Array.isArray(sharedStrings)).toBe(true)

    const sparseSharedStrings = readLargeSimpleReferencedSharedStringsFromChunks(
      (onChunk) => {
        onChunk(encoder.encode('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'))
        for (let index = 0; index < 100; index += 1) {
          onChunk(encoder.encode(`<si><t>Skip ${String(index)}</t></si>`))
        }
        onChunk(encoder.encode('<si><t>Keep sparse</t></si></sst>'))
        return true
      },
      new Set([100]),
    )

    expect(Array.isArray(sparseSharedStrings)).toBe(false)
    expect(sparseSharedStrings?.length).toBe(101)
    expect(sparseSharedStrings?.[100]).toEqual({ text: 'Keep sparse', rich: false })
  })

  it('decodes rich shared-string text lazily', () => {
    const sharedStrings = readLargeSimpleReferencedSharedStringsFromChunks(
      (onChunk) => {
        onChunk(
          encoder.encode(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>Rich</t></r><r><t> Text</t></r></si></sst>',
          ),
        )
        return true
      },
      new Set([0]),
    )
    const entry = sharedStrings?.[0]

    expect(entry?.rich).toBe(true)
    expect(typeof Object.getOwnPropertyDescriptor(entry, 'text')?.get).toBe('function')
    expect(entry?.text).toBe('Rich Text')
    expect(entry?.text).toBe('Rich Text')
  })
})
