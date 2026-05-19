import { describe, expect, it } from 'vitest'

import { readLargeSimpleSheetHyperlinkRefsFromBytes } from '../xlsx-large-simple-hyperlinks.js'

const encoder = new TextEncoder()

describe('large simple hyperlink byte scan', () => {
  it('parses hyperlink refs and decoded display metadata from bytes', () => {
    const bytes = encoder.encode(
      [
        '<hyperlinks>',
        '<hyperlink ref="A1" r:id="rIdHyperlink1" tooltip="Open &amp; review" display="Report &quot;A&quot;"/>',
        "<hyperlink ref='B2' location='Summary!A1' display='Jump'/>",
        '</hyperlinks>',
      ].join(''),
    )

    expect(readLargeSimpleSheetHyperlinkRefsFromBytes(bytes, 0, bytes.byteLength)).toEqual([
      {
        ref: 'A1',
        relationshipId: 'rIdHyperlink1',
        tooltip: 'Open & review',
        display: 'Report "A"',
      },
      {
        ref: 'B2',
        location: 'Summary!A1',
        display: 'Jump',
      },
    ])
  })

  it('signals fallback when hyperlink range expansion would lose fidelity', () => {
    const bytes = encoder.encode('<hyperlinks><hyperlink ref="A1:A2000" location="Summary!A1"/></hyperlinks>')

    expect(readLargeSimpleSheetHyperlinkRefsFromBytes(bytes, 0, bytes.byteLength)).toBeNull()
  })
})
