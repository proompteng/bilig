import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { strToU8, zipSync } from 'fflate'
import { afterEach, describe, expect, it } from 'vitest'

import { inspectWorkbookFootprintIsolated } from '../public-workbook-corpus-footprint.ts'

const publicWorkbookCorpusFootprintWorkerScriptPath = fileURLToPath(
  new URL('../public-workbook-corpus-footprint-worker.ts', import.meta.url),
)

describe('public workbook corpus footprint worker', () => {
  const tempDirs: string[] = []

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('inspects large-simple XLSX footprints from a file path without requiring stdin workbook bytes', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'bilig-footprint-'))
    tempDirs.push(dir)
    const filePath = join(dir, 'large-simple.xlsx')
    writeFileSync(filePath, buildLargeSimpleWorkbook())

    const footprint = await inspectWorkbookFootprintIsolated({
      bytes: new Uint8Array(0),
      filePath,
      fileName: 'large-simple.xlsx',
      scriptPath: publicWorkbookCorpusFootprintWorkerScriptPath,
      options: {
        timeoutMs: 30_000,
        maxRssBytes: 256 * 1024 * 1024,
        rssCheckIntervalMs: 25,
      },
    })

    expect(footprint?.largeSimpleXlsxImport).toEqual({ eligible: true, blockers: [] })
    expect(footprint?.featureCounts).toMatchObject({
      sheetCount: 1,
      cellCount: 2,
      valueCellCount: 2,
      formulaCellCount: 0,
    })
    expect(footprint?.workbookMetadata.dimensions).toEqual([
      {
        sheetName: 'Data',
        rowCount: 1,
        columnCount: 2,
        nonEmptyCellCount: 2,
        usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 },
      },
    ])
  }, 30_000)
})

function buildLargeSimpleWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>',
    ),
  })
}
