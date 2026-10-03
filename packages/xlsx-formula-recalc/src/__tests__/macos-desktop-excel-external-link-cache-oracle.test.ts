import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelPackageOpenSaveOracle } from '@bilig/excel-fixtures'
import { ValueTag } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper, exportXlsx, recalculateXlsx } from '@bilig/workpaper/xlsx'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const externalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
const externalRangeAddresses = ['C1', 'C2', 'C3'] as const

const staleFormulaCacheValues = {
  C1: '120',
  C2: '40',
  C3: '60',
} as const

const updatedFormulaCacheValues = {
  C1: '180',
  C2: '60',
  C3: '80',
} as const

describe('macOS Desktop Excel external-link cache recalc oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel updated external-link caches after formula recalculation',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-xlsx-recalc-external-link-cache-oracle-')
      try {
        const linkedSourcePath = join(tempDir, 'rates.xlsx')
        const sourcePath = join(tempDir, 'external-link-cache.xlsx')
        writeFileSync(linkedSourcePath, buildExternalSourceWorkbook([20, 30, 40]))

        const sourceBytes = buildExternalLinkRangeCacheWorkbook(pathToFileURL(linkedSourcePath).href)
        expect(worksheetFormulaCacheValues(sourceBytes)).toEqual(staleFormulaCacheValues)
        writeFileSync(sourcePath, sourceBytes)

        const excelTruth = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          companionWorkbookPaths: [linkedSourcePath],
          saveWorkbook: true,
          timeoutMs: 120_000,
          updateLinks: 'external',
        })
        expect(excelTruth.excelVersion).toMatch(/^\d+\./u)
        expect(excelTruth.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'C1', value: { kind: 'number', value: 180 } },
          { address: 'C2', value: { kind: 'number', value: 60 } },
          { address: 'C3', value: { kind: 'number', value: 80 } },
        ])

        const excelUpdatedBytes = new Uint8Array(readFileSync(sourcePath))
        expect(worksheetFormulaCacheValues(excelUpdatedBytes)).toEqual(updatedFormulaCacheValues)

        const recalculated = recalculateXlsx(sourceBytes, {
          fileName: 'external-link-cache.xlsx',
          externalWorkbooks: [
            {
              fileName: 'rates.xlsx',
              target: pathToFileURL(linkedSourcePath).href,
              bytes: new Uint8Array(readFileSync(linkedSourcePath)),
            },
          ],
          reads: externalRangeAddresses.map((address) => `Model!${address}`),
        })
        expect(
          Object.fromEntries(externalRangeAddresses.map((address) => [address, numberCell(recalculated.reads[`Model!${address}`])])),
        ).toEqual({
          C1: 180,
          C2: 60,
          C3: 80,
        })
        expect(worksheetFormulaCacheValues(recalculated.xlsx)).toEqual(updatedFormulaCacheValues)
        expect(externalLinkCachePayload(unzipSync(recalculated.xlsx))).toEqual(externalLinkCachePayload(unzipSync(excelUpdatedBytes)))

        const recalculatedPath = join(tempDir, 'external-link-cache-recalculated.xlsx')
        writeFileSync(recalculatedPath, recalculated.xlsx)
        const excelReopened = runMacosExcelInspectionOracle({
          workbookPath: recalculatedPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          calculationPolicy: 'none',
          saveWorkbook: false,
          timeoutMs: 120_000,
          updateLinks: 'never',
        })
        expect(excelReopened.cells.map(({ address, value }) => ({ address, value }))).toEqual(
          excelTruth.cells.map(({ address, value }) => ({ address, value })),
        )
        expect(excelReopened.cells.map(({ address, formula }) => ({ address, formula }))).toEqual([
          { address: 'C1', formula: '=SUM(__bilig_ext_1_Rates!$B$2:$B$4)*B1' },
          { address: 'C2', formula: '=XLOOKUP("B",__bilig_ext_1_Rates!$A$2:$A$4,__bilig_ext_1_Rates!$B$2:$B$4)*B1' },
          { address: 'C3', formula: '=SUMIFS(__bilig_ext_1_Rates!$B$2:$B$4,__bilig_ext_1_Rates!$A$2:$A$4,"C")*B1' },
        ])
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel native changed-companion package save without explicit calculation',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-xlsx-recalc-external-link-native-update-')
      try {
        const linkedSourcePath = join(tempDir, 'rates.xlsx')
        const sourcePath = join(tempDir, 'external-link-native-update.xlsx')
        writeFileSync(linkedSourcePath, buildExternalSourceWorkbook([20, 30, 40]))

        const sourceBytes = buildExternalLinkRangeCacheWorkbook(pathToFileURL(linkedSourcePath).href)
        expect(worksheetFormulaCacheValues(sourceBytes)).toEqual(staleFormulaCacheValues)
        expect(externalLinkCacheValuesByAddress(sourceBytes, ['B2', 'B3', 'B4'])).toEqual({ B2: '10', B3: '20', B4: '30' })
        writeFileSync(sourcePath, sourceBytes)

        const excelNativeSave = runMacosExcelPackageOpenSaveOracle({
          workbookPath: sourcePath,
          companionWorkbookPaths: [linkedSourcePath],
          calculationPolicy: 'none',
          saveWorkbook: true,
          timeoutMs: 120_000,
          updateLinks: 'external',
        })
        expect(excelNativeSave.excelVersion).toMatch(/^\d+\./u)

        const excelUpdatedBytes = new Uint8Array(readFileSync(sourcePath))
        expect(worksheetFormulaCacheValues(excelUpdatedBytes)).toEqual(updatedFormulaCacheValues)
        expect(externalLinkCacheValuesByAddress(excelUpdatedBytes, ['B2', 'B3', 'B4'])).toEqual({
          B2: '20',
          B3: '30',
          B4: '40',
        })

        const recalculated = recalculateXlsx(sourceBytes, {
          fileName: 'external-link-native-update.xlsx',
          externalWorkbooks: [
            {
              fileName: 'rates.xlsx',
              target: pathToFileURL(linkedSourcePath).href,
              bytes: new Uint8Array(readFileSync(linkedSourcePath)),
            },
          ],
          reads: externalRangeAddresses.map((address) => `Model!${address}`),
        })

        expect(
          Object.fromEntries(externalRangeAddresses.map((address) => [address, numberCell(recalculated.reads[`Model!${address}`])])),
        ).toEqual({
          C1: 180,
          C2: 60,
          C3: 80,
        })
        expect(worksheetFormulaCacheValues(recalculated.xlsx)).toEqual(worksheetFormulaCacheValues(excelUpdatedBytes))
        expect(externalLinkCachePayload(unzipSync(recalculated.xlsx))).toEqual(externalLinkCachePayload(unzipSync(excelUpdatedBytes)))
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel sparse blank and error external-link cache recalculation',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-xlsx-recalc-external-link-sparse-cache-oracle-')
      try {
        const linkedSourcePath = join(tempDir, 'rates.xlsx')
        const sourcePath = join(tempDir, 'external-link-sparse-cache.xlsx')
        writeFileSync(linkedSourcePath, buildSparseExternalSourceWorkbook())

        const sourceBytes = buildSparseExternalLinkCacheWorkbook(pathToFileURL(linkedSourcePath).href)
        expect(worksheetFormulaCacheValuesFor(sourceBytes, ['C1', 'C2', 'C3'])).toEqual({ C1: '60', C2: '60', C3: '60' })
        writeFileSync(sourcePath, sourceBytes)

        const excelTruth = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: ['C1', 'C2', 'C3'],
          companionWorkbookPaths: [linkedSourcePath],
          saveWorkbook: true,
          timeoutMs: 120_000,
          updateLinks: 'external',
        })
        expect(excelTruth.excelVersion).toMatch(/^\d+\./u)
        expect(excelTruth.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'C1', value: { kind: 'number', value: 70 } },
          { address: 'C2', value: { kind: 'number', value: 99 } },
          { address: 'C3', value: { kind: 'number', value: 88 } },
        ])

        const excelUpdatedBytes = new Uint8Array(readFileSync(sourcePath))
        expect(worksheetFormulaCacheValuesFor(excelUpdatedBytes, ['C1', 'C2', 'C3'])).toEqual({ C1: '70', C2: '99', C3: '88' })

        const recalculated = recalculateXlsx(sourceBytes, {
          fileName: 'external-link-sparse-cache.xlsx',
          externalWorkbooks: [
            {
              fileName: 'rates.xlsx',
              target: pathToFileURL(linkedSourcePath).href,
              bytes: new Uint8Array(readFileSync(linkedSourcePath)),
            },
          ],
          reads: ['Model!C1', 'Model!C2', 'Model!C3'],
        })
        expect(numberCell(recalculated.reads['Model!C1'])).toBe(70)
        expect(numberCell(recalculated.reads['Model!C2'])).toBe(99)
        expect(numberCell(recalculated.reads['Model!C3'])).toBe(88)
        expect(worksheetFormulaCacheValuesFor(recalculated.xlsx, ['C1', 'C2', 'C3'])).toEqual({ C1: '70', C2: '99', C3: '88' })
        expect(externalLinkCachePayload(unzipSync(recalculated.xlsx))).toEqual(externalLinkCachePayload(unzipSync(excelUpdatedBytes)))
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    180_000,
  )
})

function buildExternalSourceWorkbook(rates: readonly [number, number, number]): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Rates: [
      ['SKU', 'Rate'],
      ['A', rates[0]],
      ['B', rates[1]],
      ['C', rates[2]],
    ],
  })
  try {
    return exportXlsx(workbook.exportSnapshot())
  } finally {
    workbook.dispose()
  }
}

function buildSparseExternalSourceWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Rates: [
      ['SKU', 'Rate'],
      ['A', 20],
      ['B', null],
      ['C', 50],
      ['D', 0],
      ['E', 0],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      xmlText(zip, 'xl/worksheets/sheet1.xml')
        .replace(/<c\b[^>]*\br=(["'])B5\1[^>]*>[\s\S]*?<\/c>/u, '<c r="B5" t="e"><v>#N/A</v></c>')
        .replace(/<c\b[^>]*\br=(["'])B6\1[^>]*>[\s\S]*?<\/c>/u, '<c r="B6" t="e"><v>#NULL!</v></c>'),
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function buildExternalLinkRangeCacheWorkbook(target: string): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [
      [null, 2, 120],
      [null, null, 40],
      [null, null, 60],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      xmlText(zip, 'xl/worksheets/sheet1.xml')
        .replace(/<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>120</v></c>')
        .replace(
          /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
          "<c r=\"C2\"><f>_xlfn.XLOOKUP(&quot;B&quot;,'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1</f><v>40</v></c>",
        )
        .replace(
          /<c\b[^>]*\br=(["'])C3\1[^>]*>[\s\S]*?<\/c>/u,
          "<c r=\"C3\"><f>SUMIFS('[1]Rates'!$B$2:$B$4,'[1]Rates'!$A$2:$A$4,&quot;C&quot;)*B1</f><v>60</v></c>",
        ),
    )
    zip['xl/workbook.xml'] = strToU8(
      ensureRelationshipNamespace(xmlText(zip, 'xl/workbook.xml')).replace(
        '</sheets>',
        '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
      ),
    )
    zip['xl/_rels/workbook.xml.rels'] = strToU8(
      xmlText(zip, 'xl/_rels/workbook.xml.rels').replace(
        '</Relationships>',
        '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
      ),
    )
    zip['xl/externalLinks/externalLink5.xml'] = strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
        '<externalBook r:id="rId1">',
        '<sheetNames><sheetName val="Rates"/></sheetNames>',
        '<sheetDataSet><sheetData sheetId="0">',
        '<row r="1"><cell r="A1" t="str"><v>SKU</v></cell><cell r="B1" t="str"><v>Rate</v></cell></row>',
        '<row r="2"><cell r="A2" t="str"><v>A</v></cell><cell r="B2"><v>10</v></cell></row>',
        '<row r="3"><cell r="A3" t="str"><v>B</v></cell><cell r="B3"><v>20</v></cell></row>',
        '<row r="4"><cell r="A4" t="str"><v>C</v></cell><cell r="B4"><v>30</v></cell></row>',
        '</sheetData></sheetDataSet>',
        '</externalBook>',
        '</externalLink>',
      ].join(''),
    )
    zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
        '</Relationships>',
    )
    zip['[Content_Types].xml'] = strToU8(
      upsertContentTypeOverride(xmlText(zip, '[Content_Types].xml'), {
        partName: '/xl/externalLinks/externalLink5.xml',
        contentType: externalLinkContentType,
      }),
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function buildSparseExternalLinkCacheWorkbook(target: string): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [
      [null, 1, 60],
      [null, null, 60],
      [null, null, 60],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      xmlText(zip, 'xl/worksheets/sheet1.xml')
        .replace(/<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>60</v></c>')
        .replace(/<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C2"><f>IFERROR(SUM(\'[1]Rates\'!$B$2:$B$5),99)</f><v>60</v></c>')
        .replace(/<c\b[^>]*\br=(["'])C3\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C3"><f>IFERROR(SUM(\'[1]Rates\'!$B$6),88)</f><v>60</v></c>'),
    )
    zip['xl/workbook.xml'] = strToU8(
      ensureRelationshipNamespace(xmlText(zip, 'xl/workbook.xml')).replace(
        '</sheets>',
        '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
      ),
    )
    zip['xl/_rels/workbook.xml.rels'] = strToU8(
      xmlText(zip, 'xl/_rels/workbook.xml.rels').replace(
        '</Relationships>',
        '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
      ),
    )
    zip['xl/externalLinks/externalLink5.xml'] = strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
        '<externalBook r:id="rId1">',
        '<sheetNames><sheetName val="Rates"/></sheetNames>',
        '<sheetDataSet><sheetData sheetId="0">',
        '<row r="2"><cell r="B2"><v>10</v></cell></row>',
        '<row r="3"><cell r="B3"><v>20</v></cell></row>',
        '<row r="4"><cell r="B4"><v>30</v></cell></row>',
        '<row r="5"><cell r="B5"><v>40</v></cell></row>',
        '<row r="6"><cell r="B6"><v>40</v></cell></row>',
        '</sheetData></sheetDataSet>',
        '</externalBook>',
        '</externalLink>',
      ].join(''),
    )
    zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
        '</Relationships>',
    )
    zip['[Content_Types].xml'] = strToU8(
      upsertContentTypeOverride(xmlText(zip, '[Content_Types].xml'), {
        partName: '/xl/externalLinks/externalLink5.xml',
        contentType: externalLinkContentType,
      }),
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function numberCell(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'tag' in value && value.tag === ValueTag.Number && 'value' in value) {
    return Number(value.value)
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}

function createExcelAccessibleTempDir(prefix: string): string {
  const root = join(tmpdir(), 'bilig-xlsx-recalc-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

function worksheetFormulaCacheValues(bytes: Uint8Array): Readonly<Record<string, string | null>> {
  return worksheetFormulaCacheValuesFor(bytes, externalRangeAddresses)
}

function worksheetFormulaCacheValuesFor(bytes: Uint8Array, addresses: readonly string[]): Readonly<Record<string, string | null>> {
  const sheetXml = xmlText(unzipSync(bytes), 'xl/worksheets/sheet1.xml')
  return Object.fromEntries(
    addresses.map((address) => {
      const cellXml = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)?.[0]
      if (!cellXml) {
        throw new Error(`Missing formula cell ${address}`)
      }
      return [address, /<v>([\s\S]*?)<\/v>/u.exec(cellXml)?.[1] ?? null]
    }),
  )
}

function externalLinkCacheValuesByAddress(bytes: Uint8Array, addresses: readonly string[]): Readonly<Record<string, string | null>> {
  const valuesByAddress = new Map(externalLinkCachePayload(unzipSync(bytes)).cells.map((cell) => [cell.address, cell.value]))
  return Object.fromEntries(addresses.map((address) => [address, valuesByAddress.get(address) ?? null]))
}

function externalLinkCachePayload(zip: Record<string, Uint8Array>): {
  readonly sheetNames: readonly string[]
  readonly cells: readonly {
    readonly sheetId: string
    readonly address: string
    readonly type: string
    readonly value: string
  }[]
} {
  const externalLinkXmls = Object.entries(zip)
    .filter(([path]) => /^xl\/externalLinks\/externalLink[^/]*\.xml$/u.test(path))
    .map(([, bytes]) => strFromU8(bytes))
  const sheetNames = externalLinkXmls.flatMap((xml) =>
    [...xml.matchAll(/<sheetName\b[^>]*\bval="([^"]+)"/gu)].flatMap((match) => match[1] ?? []),
  )
  const cells = externalLinkXmls.flatMap((xml) =>
    [...xml.matchAll(/<sheetData\b([^>]*)>([\s\S]*?)<\/sheetData>/gu)].flatMap((sheetDataMatch) => {
      const sheetId = readXmlAttribute(sheetDataMatch[1] ?? '', 'sheetId')
      return [...(sheetDataMatch[2] ?? '').matchAll(/<cell\b([^>]*)>[\s\S]*?<v>([\s\S]*?)<\/v>[\s\S]*?<\/cell>/gu)].map((cellMatch) => ({
        sheetId,
        address: readXmlAttribute(cellMatch[1] ?? '', 'r'),
        type: readXmlAttribute(cellMatch[1] ?? '', 't'),
        value: cellMatch[2] ?? '',
      }))
    }),
  )
  return {
    sheetNames: sheetNames.toSorted(),
    cells: cells.toSorted(
      (left, right) =>
        left.sheetId.localeCompare(right.sheetId) ||
        left.address.localeCompare(right.address) ||
        left.type.localeCompare(right.type) ||
        left.value.localeCompare(right.value),
    ),
  }
}

function readXmlAttribute(attributes: string, attributeName: string): string {
  return new RegExp(`\\b${attributeName}="([^"]*)"`, 'u').exec(attributes)?.[1] ?? ''
}

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<workbook\b([^>]*)>/u, `<workbook$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function upsertContentTypeOverride(
  contentTypesXml: string,
  input: {
    readonly partName: string
    readonly contentType: string
  },
): string {
  if (contentTypesXml.includes(`PartName="${input.partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${input.partName}" ContentType="${input.contentType}"/></Types>`)
}

function xmlText(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}
