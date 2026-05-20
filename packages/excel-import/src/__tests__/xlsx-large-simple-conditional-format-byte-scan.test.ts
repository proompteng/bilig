import { describe, expect, it } from 'vitest'

import { readLargeSimpleConditionalFormattingFromBytes } from '../xlsx-large-simple-conditional-format-byte-scan.js'

const encoder = new TextEncoder()

describe('large simple conditional format byte scan', () => {
  it('parses faithfully typed rules without retaining the full XML block', () => {
    const xml = [
      '<conditionalFormatting sqref="A1:A2 C1:C2">',
      '<cfRule type="cellIs" priority="1" operator="greaterThan" stopIfTrue="0"><formula>3</formula></cfRule>',
      '<cfRule type="expression" priority="2"><formula>LEN(B1)&gt;0</formula></cfRule>',
      '</conditionalFormatting>',
    ].join('')

    const scan = readLargeSimpleConditionalFormattingFromBytes('Data', encoder.encode(xml), 0, encoder.encode(xml).byteLength, 1)

    expect(scan).toEqual({
      ruleCount: 4,
      conditionalFormats: [
        {
          id: 'xlsx-cf:Data:A1:A2:1',
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A2' },
          rule: { kind: 'cellIs', operator: 'greaterThan', values: [3] },
          style: {},
          stopIfTrue: false,
          priority: 1,
        },
        {
          id: 'xlsx-cf:Data:C1:C2:2',
          range: { sheetName: 'Data', startAddress: 'C1', endAddress: 'C2' },
          rule: { kind: 'cellIs', operator: 'greaterThan', values: [3] },
          style: {},
          stopIfTrue: false,
          priority: 1,
        },
        {
          id: 'xlsx-cf:Data:A1:A2:3',
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A2' },
          rule: { kind: 'formula', formula: '=LEN(B1)>0' },
          style: {},
          priority: 2,
        },
        {
          id: 'xlsx-cf:Data:C1:C2:4',
          range: { sheetName: 'Data', startAddress: 'C1', endAddress: 'C2' },
          rule: { kind: 'formula', formula: '=LEN(B1)>0' },
          style: {},
          priority: 2,
        },
      ],
    })
  })

  it('parses literal formula entities and blank rules', () => {
    const xml = [
      '<conditionalFormatting sqref="$B$1">',
      '<cfRule type="cellIs" priority="3" operator="notEqual"><formula>&quot;Closed&quot;</formula></cfRule>',
      '<cfRule type="containsBlanks" priority="4" stopIfTrue="1"/>',
      '</conditionalFormatting>',
    ].join('')
    const bytes = encoder.encode(xml)

    const scan = readLargeSimpleConditionalFormattingFromBytes('Checks', bytes, 0, bytes.byteLength, 7)

    expect(scan).toEqual({
      ruleCount: 2,
      conditionalFormats: [
        {
          id: 'xlsx-cf:Checks:B1:B1:7',
          range: { sheetName: 'Checks', startAddress: 'B1', endAddress: 'B1' },
          rule: { kind: 'cellIs', operator: 'notEqual', values: ['Closed'] },
          style: {},
          priority: 3,
        },
        {
          id: 'xlsx-cf:Checks:B1:B1:8',
          range: { sheetName: 'Checks', startAddress: 'B1', endAddress: 'B1' },
          rule: { kind: 'blanks' },
          style: {},
          stopIfTrue: true,
          priority: 4,
        },
      ],
    })
  })

  it('keeps raw XML when exact artifacts are still required', () => {
    const xml = [
      '<conditionalFormatting sqref="A1">',
      '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule>',
      '</conditionalFormatting>',
    ].join('')
    const bytes = encoder.encode(xml)

    expect(readLargeSimpleConditionalFormattingFromBytes('Data', bytes, 0, bytes.byteLength, 1)).toEqual({
      ruleCount: 1,
      artifactXml: xml,
    })
  })
})
