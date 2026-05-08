import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

type ImportedStyleWithAlignment = {
  readonly alignment?: unknown
}

describe('XLSX alignment roundtrip', () => {
  it('preserves cell alignment, wrapping, indentation, reading order, shrink-to-fit, and rotation', () => {
    const imported = importXlsx(buildAlignmentWorkbookBytes(), 'alignment-roundtrip.xlsx')
    const importedStyle = readSingleImportedStyle(imported.snapshot.workbook.metadata?.styles ?? [])

    expect(importedStyle?.alignment).toEqual({
      horizontal: 'centerContinuous',
      vertical: 'middle',
      wrap: true,
      indent: 2,
      shrinkToFit: true,
      readingOrder: 2,
      textRotation: 45,
    })

    const exported = exportXlsx(imported.snapshot)
    const exportedStylesXml = strFromU8(unzipSync(exported)['xl/styles.xml'] ?? new Uint8Array())

    expect(exportedStylesXml).toContain('horizontal="centerContinuous"')
    expect(exportedStylesXml).toContain('vertical="center"')
    expect(exportedStylesXml).toContain('wrapText="1"')
    expect(exportedStylesXml).toContain('indent="2"')
    expect(exportedStylesXml).toContain('shrinkToFit="1"')
    expect(exportedStylesXml).toContain('readingOrder="2"')
    expect(exportedStylesXml).toContain('textRotation="45"')

    const reimported = importXlsx(exported, 'alignment-roundtrip-exported.xlsx')
    const reimportedStyle = readSingleImportedStyle(reimported.snapshot.workbook.metadata?.styles ?? [])
    expect(reimportedStyle?.alignment).toEqual(importedStyle?.alignment)
  })

  it('preserves explicit general horizontal alignment with other alignment fields', () => {
    const imported = importXlsx(buildGeneralAlignmentWorkbookBytes(), 'general-alignment-roundtrip.xlsx')
    const importedStyle = readSingleImportedStyle(imported.snapshot.workbook.metadata?.styles ?? [])

    expect(importedStyle?.alignment).toEqual({
      horizontal: 'general',
      vertical: 'top',
      wrap: true,
    })

    const exported = exportXlsx(imported.snapshot)
    const exportedStylesXml = strFromU8(unzipSync(exported)['xl/styles.xml'] ?? new Uint8Array())

    expect(exportedStylesXml).toContain('horizontal="general"')
    expect(exportedStylesXml).toContain('vertical="top"')
    expect(exportedStylesXml).toContain('wrapText="1"')

    const reimported = importXlsx(exported, 'general-alignment-roundtrip-exported.xlsx')
    const reimportedStyle = readSingleImportedStyle(reimported.snapshot.workbook.metadata?.styles ?? [])
    expect(reimportedStyle?.alignment).toEqual(importedStyle?.alignment)
  })
})

function readSingleImportedStyle(styles: readonly ImportedStyleWithAlignment[]): ImportedStyleWithAlignment | undefined {
  return styles.find((style) => style.alignment !== undefined)
}

function buildAlignmentWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[123]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Aligned')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(alignmentWorksheetXml)
  zip['xl/styles.xml'] = strToU8(alignmentStylesXml)
  return zipSync(zip)
}

function buildGeneralAlignmentWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([['general']])
  XLSX.utils.book_append_sheet(workbook, sheet, 'General')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(alignmentWorksheetXml)
  zip['xl/styles.xml'] = strToU8(generalAlignmentStylesXml)
  return zipSync(zip)
}

const alignmentWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<dimension ref="A1:A1"/>',
  '<sheetData><row r="1"><c r="A1" s="1"><v>123</v></c></row></sheetData>',
  '</worksheet>',
].join('')

const alignmentStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2">',
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">',
  '<alignment horizontal="centerContinuous" vertical="center" wrapText="1" indent="2" shrinkToFit="1" readingOrder="2" textRotation="45"/>',
  '</xf>',
  '</cellXfs>',
  '</styleSheet>',
].join('')

const generalAlignmentStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2">',
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">',
  '<alignment horizontal="general" vertical="top" wrapText="1"/>',
  '</xf>',
  '</cellXfs>',
  '</styleSheet>',
].join('')
