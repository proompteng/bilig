import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX round-trip semantics', () => {
  it('preserves sort keys that point at a header row outside the sorted data range', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'header-sort-key' },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A1', value: 'Status' },
            { address: 'A2', value: 'Closed' },
            { address: 'A3', value: 'Open' },
            { address: 'A4', value: 'Pending' },
          ],
          metadata: {
            sorts: [
              {
                range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'A4' },
                keys: [{ keyAddress: 'A1', direction: 'asc' }],
              },
            ],
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const imported = importXlsx(exported, 'header-sort-key.xlsx')

    expect(sheetXml).toContain('<sortState ref="A2:A4">')
    expect(sheetXml).toContain('<sortCondition ref="A1:A4"/>')
    expect(imported.snapshot.sheets[0]?.metadata?.sorts).toEqual(snapshot.sheets[0]?.metadata?.sorts)
  })

  it('preserves custom row metadata on rows that have no cells', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'blank-row-metadata' },
      sheets: [
        {
          id: 1,
          name: 'Rows',
          order: 0,
          cells: [{ address: 'A3', value: 'visible data' }],
          metadata: {
            rows: [
              { id: 'row:0', index: 0, size: 44 },
              { id: 'row:1', index: 1, hidden: true },
              { id: 'row:2', index: 2, size: 30 },
            ],
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const importedRows = importXlsx(exported, 'blank-row-metadata.xlsx').snapshot.sheets[0]?.metadata?.rows

    expect(sheetXml).toContain('<row r="1" ht="44" customHeight="1"/>')
    expect(sheetXml).toContain('<row r="2" hidden="1"/>')
    expect(importedRows).toEqual(snapshot.sheets[0]?.metadata?.rows)
  })

  it('keeps populated cells when row metadata and conditional formats are both exported', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'row-metadata-conditional-format' },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A3', value: 'visible data' },
            { address: 'A5', value: 'other data' },
          ],
          metadata: {
            rows: [
              { id: 'row:0', index: 0, size: 44 },
              { id: 'row:1', index: 1, size: 55 },
            ],
            conditionalFormats: [
              {
                id: 'cf-visible-data',
                range: { sheetName: 'Data', startAddress: 'A3', endAddress: 'A3' },
                rule: { kind: 'textContains', text: 'visible' },
                style: { font: { bold: true } },
              },
            ],
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const imported = importXlsx(exported, 'row-metadata-conditional-format.xlsx')

    expect(sheetXml).toContain('<c r="A3"')
    expect(imported.snapshot.sheets[0]?.cells).toEqual(snapshot.sheets[0]?.cells)
  })

  it('keeps text-formatted strings after formatted blank cells are exported', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'text-after-blank-formatting',
        metadata: {
          styleArtifacts: {
            stylesXml: minimalStylesXml(),
          },
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A2', value: 'Name' },
            { address: 'B2', value: 'Application Type' },
            { address: 'A3', value: 'TBA17-40676P1', format: '@' },
            { address: 'B3', value: 'Building Application', format: '@' },
          ],
          metadata: {
            styleArtifacts: {
              cellStyleIndexes: [],
              blankCellAddresses: ['C2', 'D2', 'E2'],
            },
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const imported = importXlsx(exported, 'text-after-blank-formatting.xlsx')

    expect([...sheetXml.matchAll(/\br="A3"/gu)]).toHaveLength(1)
    expect(imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A3')).toMatchObject({
      value: 'TBA17-40676P1',
      format: '@',
    })
  })

  it('keeps rich text containing dollar apostrophe literals', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'rich-text-dollar-apostrophe' },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [{ address: 'A1', value: "$'000" }],
          metadata: {
            richTextArtifacts: {
              cells: [
                {
                  address: 'A1',
                  text: "$'000",
                  storage: 'sharedString',
                  xml: `<si><r><rPr><b/></rPr><t>$'000</t></r></si>`,
                },
              ],
            },
          },
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sharedStringsXml = strFromU8(unzipSync(exported)['xl/sharedStrings.xml'] ?? new Uint8Array())
    const imported = importXlsx(exported, 'rich-text-dollar-apostrophe.xlsx')

    expect(sharedStringsXml).toContain(`<t>$'000</t>`)
    expect(imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')).toEqual({ address: 'A1', value: "$'000" })
  })

  it('keeps earlier text values when duplicate blank cell XML follows them', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'duplicate-blank-cell',
      },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [{ address: 'A3', value: 'TBA17-40676P1', format: '@' }],
        },
      ],
    }
    const zip = unzipSync(exportXlsx(snapshot))
    const sheetPath = 'xl/worksheets/sheet1.xml'
    const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
    zip[sheetPath] = strToU8(sheetXml.replace(/(<row r="3"[^>]*>[\s\S]*?)<\/row>/u, '$1<c r="A3" s="4" t="z"></c></row>'))

    const imported = importXlsx(zipSync(zip), 'duplicate-blank-cell.xlsx')

    expect(imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A3')).toMatchObject({ value: 'TBA17-40676P1' })
  })
})

function minimalStylesXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>',
    '<borders count="1"><border/></borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    '</styleSheet>',
  ].join('')
}
