import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import { strFromU8, unzipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { readBenchToleranceMultiplier } from '../../../../scripts/bench-tolerance.js'
import { exportXlsx, importXlsx } from '../index.js'

describe('large simple XLSX export', () => {
  it('round-trips large value-heavy sheets without the style writer hot path', () => {
    const exported = exportXlsx(buildLargeSimpleSnapshot())
    const imported = importXlsx(exported, 'large-simple.xlsx')
    const sheet = imported.snapshot.sheets[0]
    const styleRange = sheet?.metadata?.styleRanges?.find((entry) => entry.range.startAddress === 'A1')
    const style = imported.snapshot.workbook.metadata?.styles?.find((entry) => entry.id === styleRange?.styleId)

    expect(sheet?.cells).toHaveLength(100_000)
    expect(sheet?.cells.find((cell) => cell.address === 'A2')).toMatchObject({ value: 50, format: '0.00' })
    expect(sheet?.metadata?.merges).toEqual([{ sheetName: 'Large', startAddress: 'A1', endAddress: 'B1' }])
    expect(style?.fill?.backgroundColor).toBe('#ffcc00')
    expect(style?.font?.bold).toBe(true)
  }, 15_000)

  it('does not let broad column metadata expand the exported worksheet scan range', () => {
    const start = performance.now()
    const exported = exportXlsx(buildBroadColumnMetadataSnapshot())
    const durationMs = performance.now() - start
    const workbook = XLSX.read(exported, { type: 'array', cellFormula: true, cellText: false, cellDates: false })

    expect(durationMs).toBeLessThan(1_500 * readBenchmarkTolerance())
    expect(workbook.Sheets['Wide']?.['!ref']).toBe('A3040')
  }, 15_000)

  it('exports sparse raw style artifacts without widening the SheetJS writer scan range', () => {
    const start = performance.now()
    const exported = exportXlsx(buildSparseStyleArtifactSnapshot())
    const durationMs = performance.now() - start
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(durationMs).toBeLessThan(3_000 * readBenchmarkTolerance())
    expect(sheetXml).toContain('<dimension ref="A1:CF65000"/>')
    expect(sheetXml).toContain('<c r="CF65000" s="1"/>')
  }, 15_000)

  it('exports formula-heavy metadata workbooks inside the production timeout budget', () => {
    const start = performance.now()
    const exported = exportXlsx(buildFormulaHeavyMetadataSnapshot())
    const durationMs = performance.now() - start
    const workbook = XLSX.read(exported, { type: 'array', cellFormula: true, cellText: false, cellDates: false })
    const imported = importXlsx(exported, 'issue-90-formula-heavy-export.xlsx')
    const sheet = imported.snapshot.sheets[0]

    expect(durationMs).toBeLessThan(6_000 * readBenchmarkTolerance())
    expect(sheet?.cells).toHaveLength(60_000)
    expect(sheet?.cells.find((cell) => cell.address === 'B100')).toMatchObject({
      value: 1189,
      formula: 'A100+1',
    })
    expect(workbook.Workbook?.Names).toHaveLength(40)
    expect(workbook.Sheets['Export']?.['!ref']).toBe('A1:L5000')
  }, 20_000)
})

function readBenchmarkTolerance(): number {
  return readBenchToleranceMultiplier(process.env)
}

function buildLargeSimpleSnapshot(): WorkbookSnapshot {
  const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
  for (let row = 0; row < 2_000; row += 1) {
    for (let column = 0; column < 50; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column })
      cells.push({
        address,
        value: row * 50 + column,
        ...(address === 'A2' ? { format: '0.00' } : {}),
      })
    }
  }
  return {
    version: 1,
    workbook: {
      name: 'large-simple',
      metadata: {
        styles: [
          {
            id: 'header',
            fill: { backgroundColor: '#ffcc00' },
            font: { bold: true },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Large',
        order: 0,
        cells,
        metadata: {
          columns: [{ id: 'column:0', index: 0, size: 120 }],
          rows: [{ id: 'row:0', index: 0, size: 28 }],
          merges: [{ sheetName: 'Large', startAddress: 'A1', endAddress: 'B1' }],
          styleRanges: [{ range: { sheetName: 'Large', startAddress: 'A1', endAddress: 'B1' }, styleId: 'header' }],
        },
      },
    ],
  }
}

function buildBroadColumnMetadataSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'wide-column-metadata' },
    sheets: [
      {
        id: 1,
        name: 'Wide',
        order: 0,
        cells: [{ address: 'A3040', value: 1 }],
        metadata: {
          columns: Array.from({ length: 16_384 }, (_entry, index) => ({
            id: `column:${String(index)}`,
            index,
            size: 64,
          })),
        },
      },
    ],
  }
}

function buildSparseStyleArtifactSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'sparse-style-artifacts',
      metadata: {
        styleArtifacts: {
          stylesXml: minimalStylesXml,
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Sparse',
        order: 0,
        cells: [{ address: 'A1', value: 'Header' }],
        metadata: {
          richTextArtifacts: {
            cells: [
              {
                address: 'A1',
                text: 'Header',
                storage: 'sharedString',
                xml: '<si><r><rPr><b/></rPr><t>Header</t></r></si>',
              },
            ],
          },
          styleArtifacts: {
            cellStyleIndexes: Array.from({ length: 65_000 }, (_entry, index) => ({
              address: `CF${String(index + 1)}`,
              styleIndex: 1,
            })),
          },
        },
      },
    ],
  }
}

function buildFormulaHeavyMetadataSnapshot(): WorkbookSnapshot {
  const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
  for (let row = 0; row < 5_000; row += 1) {
    for (let column = 0; column < 12; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column })
      const formula = row < 1_260 ? `A${String(row + 1)}+${String(column)}` : undefined
      cells.push({
        address,
        value: row * 12 + column,
        ...(formula !== undefined ? { formula } : {}),
      })
    }
  }

  return {
    version: 1,
    workbook: {
      name: 'issue-90-formula-heavy-export',
      metadata: {
        definedNames: Array.from({ length: 40 }, (_entry, index) => ({
          name: `Issue90Name_${String(index + 1)}`,
          formula: `Export!$A$${String(index + 1)}`,
        })),
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Export',
        order: 0,
        cells,
        metadata: {
          rows: Array.from({ length: 9_000 }, (_entry, index) => ({
            id: `row:${String(index)}`,
            index,
            size: 18,
          })),
          columns: Array.from({ length: 12 }, (_entry, index) => ({
            id: `column:${String(index)}`,
            index,
            size: 64,
          })),
          filters: [
            {
              range: {
                sheetName: 'Export',
                startAddress: 'A1',
                endAddress: 'L5000',
              },
            },
          ],
        },
      },
    ],
  }
}

const minimalStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')
