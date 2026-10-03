import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { recalculateXlsxFileToFile } from '../index.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('@bilig/xlsx-formula-recalc native typed range aggregates', () => {
  it('evaluates SUM, AVERAGE, and COUNTA ranges through the native kernel', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-typed-aggregate-'))
    try {
      const sourcePath = join(tempDir, 'typed-aggregate.xlsx')
      const outputPath = join(tempDir, 'typed-aggregate.recalculated.xlsx')
      writeFileSync(sourcePath, buildTypedAggregateWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2', 'Data!A6', 'Data!A7'],
      })

      expect(result.reads['Data!A2']).toMatchObject({ value: 12 })
      expect(result.reads['Data!A6']).toMatchObject({ value: 16 / 3 })
      expect(result.reads['Data!A7']).toMatchObject({ value: 5 })
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(3)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(3)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(3)

      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>SUM(B2:B5)</f><v>12</v></c>')
      expect(sheetXml).toContain('<c r="A6"><f>AVERAGE(B2:B6)</f><v>5.333333333333333</v></c>')
      expect(sheetXml).toContain('<c r="A7"><f>COUNTA(B2:B7)</f><v>5</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

function buildTypedAggregateWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2:B7"/>
  <sheetData>
    <row r="2"><c r="A2"><f>SUM(B2:B5)</f><v>0</v></c><c r="B2"><v>5</v></c></row>
    <row r="3"><c r="B3" t="inlineStr"><is><t>ignored</t></is></c></row>
    <row r="4"/>
    <row r="5"><c r="B5"><v>7</v></c></row>
    <row r="6"><c r="A6"><f>AVERAGE(B2:B6)</f><v>0</v></c><c r="B6"><v>4</v></c></row>
    <row r="7"><c r="A7"><f>COUNTA(B2:B7)</f><v>0</v></c><c r="B7" t="inlineStr"><is><t>included</t></is></c></row>
  </sheetData>
</worksheet>`),
  })
}
