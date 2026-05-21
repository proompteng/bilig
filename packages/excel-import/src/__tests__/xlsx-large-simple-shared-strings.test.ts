import { describe, expect, it } from 'vitest'

import {
  collectReferencedLargeSimpleRichSharedStringIndexes,
  createLargeSimpleSharedStringSubset,
  readLargeSimpleSharedStrings,
  readLargeSimpleReferencedSharedStringsFromChunks,
} from '../xlsx-large-simple-shared-strings.js'
import { ImportedWorkbookStringPool } from '../xlsx-large-simple-string-pool.js'

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

  it('deduplicates repeated referenced shared-string text through the import string pool', () => {
    const pool = new ImportedWorkbookStringPool()
    const sharedStrings = readLargeSimpleReferencedSharedStringsFromChunks(
      (onChunk) => {
        onChunk(
          encoder.encode(
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
              '<si><t>Repeated vendor label</t></si>' +
              '<si><t>Repeated vendor label</t></si>' +
              '<si><r><t>Repeated vendor label</t></r></si>' +
              '</sst>',
          ),
        )
        return true
      },
      new Set([0, 1, 2]),
      { deduplicateText: true, stringPool: pool },
    )

    expect(sharedStrings?.[0]?.text).toBe('Repeated vendor label')
    expect(sharedStrings?.[1]?.text).toBe('Repeated vendor label')
    expect(sharedStrings?.[2]?.text).toBe('Repeated vendor label')
    expect(sharedStrings?.[1]).toBe(sharedStrings?.[0])
    expect(sharedStrings?.[2]).not.toBe(sharedStrings?.[0])
    expect(pool.count).toBe(1)
  })

  it('deduplicates fallback full shared-string table text with the same pool path', () => {
    const pool = new ImportedWorkbookStringPool()
    const sharedStrings = readLargeSimpleSharedStrings(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<si><t>Same</t></si>' +
        '<si><t>Same</t></si>' +
        '</sst>',
      { deduplicateText: true, stringPool: pool },
    )

    expect(sharedStrings).toEqual([
      { text: 'Same', rich: false },
      { text: 'Same', rich: false },
    ])
    expect(sharedStrings[1]).toBe(sharedStrings[0])
    expect(pool.count).toBe(1)
  })

  it('creates sparse sheet-scoped shared-string subsets without unrelated entries', () => {
    const subset = createLargeSimpleSharedStringSubset(
      [
        { text: 'Unused 0', rich: false },
        { text: 'Alpha', rich: false },
        { text: 'Unused 2', rich: false },
        { text: 'Beta', rich: false },
      ],
      new Set([1, 3]),
    )

    expect(Array.isArray(subset)).toBe(false)
    expect(subset?.length).toBe(4)
    expect(subset?.[0]).toBeUndefined()
    expect(subset?.[1]).toEqual({ text: 'Alpha', rich: false })
    expect(subset?.[3]).toEqual({ text: 'Beta', rich: false })
  })

  it('collects only rich shared-string indexes for sheet-scoped retention', () => {
    const richIndexes = collectReferencedLargeSimpleRichSharedStringIndexes(
      [
        { text: 'Plain A', rich: false },
        { text: 'Rich B', rich: true, xml: '<si><r><t>Rich B</t></r></si>' },
        { text: 'Plain C', rich: false },
        { text: 'Rich D', rich: true, xml: '<si><r><t>Rich D</t></r></si>' },
      ],
      new Set([0, 1, 3]),
    )

    expect(richIndexes).toEqual(new Set([1, 3]))
    expect(collectReferencedLargeSimpleRichSharedStringIndexes([{ text: 'Only', rich: false }], new Set([2]))).toBeNull()
  })
})
