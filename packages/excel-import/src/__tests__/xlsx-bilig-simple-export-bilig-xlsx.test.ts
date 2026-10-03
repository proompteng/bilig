import { spawnSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

const repoRoot = fileURLToPath(new URL('../../../..', import.meta.url))

describe('@bilig/xlsx simple export boundary', () => {
  it('exports bordered generated snapshots without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/excel-import/src/index.ts'),
  import('./packages/xlsx/src/index.ts'),
])
  .then(([{ exportXlsx }, { readXlsxZipEntries }]) => {
    const snapshot = {
      version: 1,
      workbook: {
        name: 'bilig-xlsx-bordered-export',
        metadata: {
          styles: [
            {
              id: 'total-border',
              font: { bold: true },
              borders: {
                bottom: { style: 'solid', weight: 'thin', color: '#000000' },
              },
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Report',
          order: 0,
          cells: [
            { address: 'A1', value: 'Total' },
            { address: 'B1', formula: 'SUM(B2:B3)', value: 30 },
          ],
          metadata: {
            styleRanges: [
              {
                range: { sheetName: 'Report', startAddress: 'A1', endAddress: 'B1' },
                styleId: 'total-border',
              },
            ],
          },
        },
      ],
    }
    const exported = exportXlsx(snapshot)
    const zip = readXlsxZipEntries(exported)
    const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'])
    const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({ loaded, bytes: exported.byteLength, stylesXml, sheetXml }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isBiligXlsxExportOutput(output)).toBe(true)
    if (!isBiligXlsxExportOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.bytes).toBeGreaterThan(0)
    expect(output.stylesXml).toContain('<bottom style="thin"><color rgb="FF000000"/></bottom>')
    expect(output.stylesXml).toContain('applyBorder="1"')
    expect(output.sheetXml).toContain('<c r="A1" s="1"')
    expect(output.sheetXml).toContain('<c r="B1" s="1"')
  }, 15_000)

  it('exports sparse style artifacts with @bilig/xlsx without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/excel-import/src/index.ts'),
  import('./packages/xlsx/src/index.ts'),
])
  .then(([{ exportXlsx }, { readXlsxZipEntries }]) => {
    const minimalRawStylesXml = [
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
    const snapshot = {
      version: 1,
      workbook: {
        name: 'bilig-xlsx-sparse-style-artifacts',
        metadata: { styleArtifacts: { stylesXml: minimalRawStylesXml } },
      },
      sheets: [
        {
          id: 1,
          name: 'Sparse',
          order: 0,
          cells: [{ address: 'A1', value: 'Header' }],
          metadata: {
            styleArtifacts: {
              cellStyleIndexes: Array.from({ length: 65_000 }, (_entry, index) => ({
                address: 'CF' + String(index + 1),
                styleIndex: 1,
              })),
            },
          },
        },
      ],
    }
    const exported = exportXlsx(snapshot)
    const zip = readXlsxZipEntries(exported)
    const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'])
    const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({
      loaded,
      bytes: exported.byteLength,
      hasRawStyles: stylesXml === minimalRawStylesXml,
      hasTailStyle: sheetXml.includes('<c r="CF65000" s="1"/>'),
      hasDimension: sheetXml.includes('<dimension ref="A1:CF65000"/>'),
      hasSharedStrings: zip['xl/sharedStrings.xml'] !== undefined,
    }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isSparseStyleArtifactOutput(output)).toBe(true)
    if (!isSparseStyleArtifactOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.bytes).toBeGreaterThan(0)
    expect(output.hasRawStyles).toBe(true)
    expect(output.hasTailStyle).toBe(true)
    expect(output.hasDimension).toBe(true)
    expect(output.hasSharedStrings).toBe(false)
  }, 15_000)

  it('exports table metadata with @bilig/xlsx without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/excel-import/src/index.ts'),
  import('./packages/xlsx/src/index.ts'),
])
  .then(([{ exportXlsx }, { readXlsxZipEntries }]) => {
    const snapshot = {
      version: 1,
      workbook: {
        name: 'bilig-xlsx-table-export',
        metadata: {
          tables: [
            {
              name: 'SalesTable',
              sheetName: 'Data',
              startAddress: 'A1',
              endAddress: 'B3',
              columnNames: ['Name', 'Total'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A1', value: 'Name' },
            { address: 'B1', value: 'Total' },
            { address: 'A2', value: 'One' },
            { address: 'B2', value: 10 },
          ],
        },
      ],
    }
    const exported = exportXlsx(snapshot)
    const zip = readXlsxZipEntries(exported)
    const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const tableXml = new TextDecoder().decode(zip['xl/tables/table1.xml'] ?? new Uint8Array())
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({
      loaded,
      bytes: exported.byteLength,
      hasTablePart: tableXml.includes('displayName="SalesTable"'),
      hasSheetTableRef: sheetXml.includes('<tableParts count="1">'),
      hasRelationship: zip['xl/worksheets/_rels/sheet1.xml.rels'] !== undefined,
    }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isTableExportOutput(output)).toBe(true)
    if (!isTableExportOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.bytes).toBeGreaterThan(0)
    expect(output.hasTablePart).toBe(true)
    expect(output.hasSheetTableRef).toBe(true)
    expect(output.hasRelationship).toBe(true)
  }, 15_000)

  it('exports recalculated ExcelJS bridge workbooks without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/exceljs-formula-recalc/node_modules/exceljs'),
  import('./packages/exceljs-formula-recalc/src/index.ts'),
])
  .then(async ([exceljsModule, { recalculateExceljsWorkbook }]) => {
    const ExcelJS = exceljsModule.default.default ?? exceljsModule.default
    const workbook = new ExcelJS.Workbook()
    const inputs = workbook.addWorksheet('Inputs')
    inputs.getCell('A1').value = 'Metric'
    inputs.getCell('B1').value = 'Value'
    inputs.getCell('A2').value = 'Units'
    inputs.getCell('B2').value = 40
    inputs.getCell('A3').value = 'Price'
    inputs.getCell('B3').value = 1200
    const summary = workbook.addWorksheet('Summary')
    summary.getCell('A1').value = 'Metric'
    summary.getCell('B1').value = 'Value'
    summary.getCell('A2').value = 'Revenue'
    summary.getCell('B2').value = { formula: 'Inputs!B2*Inputs!B3', result: 48000 }

    const result = await recalculateExceljsWorkbook(workbook, {
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })
    const readbackCell = result.reads['Summary!B2']
    const readback = typeof readbackCell === 'object' && readbackCell !== null && 'value' in readbackCell ? readbackCell.value : null
    const cachedCell = workbook.getWorksheet('Summary')?.getCell('B2').value
    const cachedResult = typeof cachedCell === 'object' && cachedCell !== null && 'result' in cachedCell ? cachedCell.result : null
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({
      loaded,
      bytes: result.xlsx.byteLength,
      cachedResult,
      readback,
      workbookMutated: result.workbookMutated,
    }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isExceljsBridgeExportOutput(output)).toBe(true)
    if (!isExceljsBridgeExportOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.bytes).toBeGreaterThan(0)
    expect(output.cachedResult).toBe(72_000)
    expect(output.readback).toBe(72_000)
    expect(output.workbookMutated).toBe(true)
  }, 15_000)

  it('exports recalculated Bilig-generated workbook reports without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/workpaper/src/index.ts'),
  import('./packages/workpaper/src/xlsx.ts'),
  import('./packages/xlsx-formula-recalc/src/workbook-compatibility-report.ts'),
])
  .then(([{ WorkPaper }, { exportXlsx }, { buildWorkbookCompatibilityReport }]) => {
    const workbook = WorkPaper.buildFromSheets({
      Inputs: [['Metric', 'Value'], ['Units', 40], ['Price', 1200]],
      Summary: [['Metric', 'Value'], ['Revenue', '=Inputs!B2*Inputs!B3']],
    })
    const bytes = exportXlsx(workbook.exportSnapshot())
    workbook.dispose()

    const report = buildWorkbookCompatibilityReport(bytes, {
      fileName: 'pricing-risk.xlsx',
      inspectLimit: 'all',
    })
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({
      loaded,
      sheetCount: report.workbook.sheetCount,
      formulaCellCount: report.workbook.formulaCellCount,
      missingCachedFormulas: report.findings.missingCachedFormulaValues.count,
      verified: report.verified,
    }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isBiligGeneratedReportOutput(output)).toBe(true)
    if (!isBiligGeneratedReportOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.sheetCount).toBe(2)
    expect(output.formulaCellCount).toBe(1)
    expect(output.missingCachedFormulas).toBe(1)
    expect(output.verified).toBe(true)
  }, 15_000)

  it('exports rich text artifacts with @bilig/xlsx without loading SheetJS xlsx', () => {
    const script = `
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
Promise.all([
  import('./packages/excel-import/src/index.ts'),
  import('./packages/xlsx/src/index.ts'),
])
  .then(([{ exportXlsx }, { readXlsxZipEntries }]) => {
    const sharedRichStringXml = [
      '<si>',
      '<r><rPr><b/><sz val="10"/><color rgb="FF1F4E79"/><rFont val="Aptos"/></rPr><t>Important:</t></r>',
      '<r><rPr><i/><sz val="10"/><color rgb="FFC00000"/><rFont val="Aptos"/></rPr><t xml:space="preserve"> Before signing off</t></r>',
      '</si>',
    ].join('')
    const inlineRichStringXml = [
      '<is>',
      '<r><rPr><u/><sz val="11"/><color rgb="FF008000"/><rFont val="Aptos"/></rPr><t>Revenue</t></r>',
      '<r><rPr><sz val="11"/><rFont val="Aptos"/></rPr><t xml:space="preserve"> sensitivity</t></r>',
      '</is>',
    ].join('')
    const snapshot = {
      version: 1,
      workbook: { name: 'bilig-xlsx-rich-text-export' },
      sheets: [
        {
          id: 1,
          name: 'Labels',
          order: 0,
          cells: [
            { address: 'A1', value: 'Important: Before signing off' },
            { address: 'B1', value: 'Revenue sensitivity' },
          ],
          metadata: {
            richTextArtifacts: {
              cells: [
                {
                  address: 'A1',
                  text: 'Important: Before signing off',
                  storage: 'sharedString',
                  xml: sharedRichStringXml,
                },
                {
                  address: 'B1',
                  text: 'Revenue sensitivity',
                  storage: 'inlineString',
                  xml: inlineRichStringXml,
                },
              ],
            },
          },
        },
      ],
    }
    const exported = exportXlsx(snapshot)
    const zip = readXlsxZipEntries(exported)
    const sharedStringsXml = new TextDecoder().decode(zip['xl/sharedStrings.xml'])
    const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
    const loaded = Object.keys(requireForCache.cache).filter((path) =>
      /[\\\\/]node_modules[\\\\/](?:\\.pnpm[\\\\/]xlsx@[^\\\\/]+[\\\\/]node_modules[\\\\/]xlsx|xlsx)(?:[\\\\/]|$)/u.test(path)
    )
    process.stdout.write(JSON.stringify({
      loaded,
      bytes: exported.byteLength,
      hasSharedStringPart: sharedStringsXml.includes(sharedRichStringXml),
      hasSharedStringCell: sheetXml.includes('<c r="A1" t="s"><v>0</v></c>'),
      hasInlineStringCell: sheetXml.includes('<c r="B1" t="inlineStr">' + inlineRichStringXml + '</c>'),
    }) + '\\n')
  })
  .catch((error) => {
    console.error(error)
    process.exit(1)
  })
`
    const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(result.status, result.stderr).toBe(0)
    const output: unknown = JSON.parse(result.stdout)
    expect(isRichTextArtifactOutput(output)).toBe(true)
    if (!isRichTextArtifactOutput(output)) {
      throw new Error(`Unexpected child output: ${result.stdout}`)
    }
    expect(output.loaded).toEqual([])
    expect(output.bytes).toBeGreaterThan(0)
    expect(output.hasSharedStringPart).toBe(true)
    expect(output.hasSharedStringCell).toBe(true)
    expect(output.hasInlineStringCell).toBe(true)
  }, 15_000)
})

function isBiligXlsxExportOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly bytes: number
  readonly stylesXml: string
  readonly sheetXml: string
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['bytes'] === 'number' &&
    typeof value['stylesXml'] === 'string' &&
    typeof value['sheetXml'] === 'string'
  )
}

function isSparseStyleArtifactOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly bytes: number
  readonly hasRawStyles: boolean
  readonly hasTailStyle: boolean
  readonly hasDimension: boolean
  readonly hasSharedStrings: boolean
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['bytes'] === 'number' &&
    typeof value['hasRawStyles'] === 'boolean' &&
    typeof value['hasTailStyle'] === 'boolean' &&
    typeof value['hasDimension'] === 'boolean' &&
    typeof value['hasSharedStrings'] === 'boolean'
  )
}

function isRichTextArtifactOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly bytes: number
  readonly hasSharedStringPart: boolean
  readonly hasSharedStringCell: boolean
  readonly hasInlineStringCell: boolean
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['bytes'] === 'number' &&
    typeof value['hasSharedStringPart'] === 'boolean' &&
    typeof value['hasSharedStringCell'] === 'boolean' &&
    typeof value['hasInlineStringCell'] === 'boolean'
  )
}

function isTableExportOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly bytes: number
  readonly hasTablePart: boolean
  readonly hasSheetTableRef: boolean
  readonly hasRelationship: boolean
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['bytes'] === 'number' &&
    typeof value['hasTablePart'] === 'boolean' &&
    typeof value['hasSheetTableRef'] === 'boolean' &&
    typeof value['hasRelationship'] === 'boolean'
  )
}

function isExceljsBridgeExportOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly bytes: number
  readonly cachedResult: number
  readonly readback: number
  readonly workbookMutated: boolean
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['bytes'] === 'number' &&
    typeof value['cachedResult'] === 'number' &&
    typeof value['readback'] === 'number' &&
    typeof value['workbookMutated'] === 'boolean'
  )
}

function isBiligGeneratedReportOutput(value: unknown): value is {
  readonly loaded: readonly string[]
  readonly sheetCount: number
  readonly formulaCellCount: number
  readonly missingCachedFormulas: number
  readonly verified: boolean
} {
  if (!isRecord(value)) {
    return false
  }
  const loaded = value['loaded']
  return (
    Array.isArray(loaded) &&
    loaded.every((entry) => typeof entry === 'string') &&
    typeof value['sheetCount'] === 'number' &&
    typeof value['formulaCellCount'] === 'number' &&
    typeof value['missingCachedFormulas'] === 'number' &&
    typeof value['verified'] === 'boolean'
  )
}

function isRecord(value: unknown): value is Record<PropertyKey, unknown> {
  return typeof value === 'object' && value !== null
}
