import { describe, expect, it } from 'vitest'

import { readLargeSimpleAutoFiltersFromBytes } from '../xlsx-large-simple-autofilter-byte-scan.js'

const encoder = new TextEncoder()

describe('large simple autofilter byte scan', () => {
  it('parses value and custom filter criteria from bytes', () => {
    const bytes = encoder.encode(
      [
        '<autoFilter ref="A1:D6">',
        '<filterColumn colId="1" hiddenButton="1"><filters blank="0"><filter val="Finance &amp; Ops"/></filters></filterColumn>',
        '<filterColumn colId="2"><customFilters and="1"><customFilter operator="lessThan" val="0"/></customFilters></filterColumn>',
        '<filterColumn colId="3" showButton="0"><filters blank="1"/></filterColumn>',
        '</autoFilter>',
      ].join(''),
    )

    expect(readLargeSimpleAutoFiltersFromBytes('Ledger', bytes, 0, bytes.byteLength)).toEqual([
      {
        sheetName: 'Ledger',
        startAddress: 'A1',
        endAddress: 'D6',
        criteria: [
          {
            colId: 1,
            hiddenButton: true,
            filters: { blank: false, values: ['Finance & Ops'] },
          },
          {
            colId: 2,
            customFilters: { and: true, filters: [{ operator: 'lessThan', value: '0' }] },
          },
          {
            colId: 3,
            showButton: false,
            filters: { blank: true, values: [] },
          },
        ],
      },
    ])
  })

  it('parses range-only autofilters from self-closing tags', () => {
    const bytes = encoder.encode('<autoFilter ref="B2:C4"/>')

    expect(readLargeSimpleAutoFiltersFromBytes('Data', bytes, 0, bytes.byteLength)).toEqual([
      {
        sheetName: 'Data',
        startAddress: 'B2',
        endAddress: 'C4',
      },
    ])
  })
})
