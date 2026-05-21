import { describe, expect, it } from 'vitest'

import { readLargeSimpleWorkbookNumberFormatsFromChunks } from '../xlsx-large-simple-number-formats.js'
import { readLargeSimpleWorkbookStylesFromChunks } from '../xlsx-large-simple-styles.js'

const encoder = new TextEncoder()

describe('large simple styles streaming', () => {
  it('discards unneeded indexed style children while waiting for the closing tag', () => {
    const retainedBufferLengths: number[] = []
    const largeUnneededFillPayload = 'x'.repeat(100_000)
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<cellXfs count="2"><xf numFmtId="0" fillId="0" fontId="0"/><xf numFmtId="0" fillId="1" fontId="1" applyFill="1" applyFont="1"/></cellXfs>',
      '<fills count="2"><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill><unused>',
      largeUnneededFillPayload.slice(0, 40_000),
      largeUnneededFillPayload.slice(40_000, 80_000),
      largeUnneededFillPayload.slice(80_000),
      '</unused></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>',
      '<fonts count="2"><font/><font><b/><name val="Inter"/><sz val="11"/></font></fonts>',
      '</styleSheet>',
    ]

    const styles = readLargeSimpleWorkbookStylesFromChunks(
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

    expect(styles?.get(1)).toEqual({
      fill: { backgroundColor: '#ffcc00' },
      font: { bold: true, family: 'Inter', size: 11 },
    })
    expect(retainedBufferLengths.length).toBeGreaterThan(0)
    expect(Math.max(...retainedBufferLengths)).toBeLessThan(1024)
  })

  it('streams number formats without dropping the visual style on the same xf', () => {
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="00000"/></numFmts>',
      '<cellXfs count="2"><xf numFmtId="0" fillId="0" fontId="0"/><xf numFmtId="164" fillId="1" fontId="0" applyFill="1" applyNumberFormat="1"/></cellXfs>',
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>',
      '<fonts count="1"><font/></fonts>',
      '</styleSheet>',
    ]
    const readChunks = (onChunk: (chunk: Uint8Array) => void): boolean => {
      for (const chunk of chunks) {
        onChunk(encoder.encode(chunk))
      }
      return true
    }

    expect(readLargeSimpleWorkbookStylesFromChunks(readChunks, new Set([1]))?.get(1)).toEqual({
      fill: { backgroundColor: '#ffcc00' },
    })
    expect(readLargeSimpleWorkbookNumberFormatsFromChunks(readChunks, new Set([1]))?.get(1)).toBe('00000')
  })
})
