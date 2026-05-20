import { describe, expect, it } from 'vitest'
import { strToU8, unzipSync, zipSync } from 'fflate'
import { readRuntimeImage } from '@bilig/core'

import { importXlsx, inspectXlsx, XlsxImportSizeLimitExceededError } from '../index.js'
import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { forEachInflatedXlsxZipEntryChunk, readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple XLSX import fast path', () => {
  it('imports simple OpenXML worksheets without SheetJS workbook materialization', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:E5"/>',
        '<sheetFormatPr defaultRowHeight="15"/>',
        '<cols><col min="1" max="2" width="10" customWidth="1"/></cols>',
        '<sheetData>',
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42.5</v></c></row>',
        '<row r="2" ht="20" customHeight="1"><c r="C2" t="b"><v>1</v></c></row>',
        '<row r="3"><c r="D3" t="inlineStr"><is><t>Inline text</t></is></c></row>',
        '<row r="4"><c r="E4" t="e"><v>#N/A</v></c></row>',
        '<row r="5"><c r="A5" t="s"><v>1</v></c></row>',
        '</sheetData>',
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'large-simple.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha &#8211;' },
      { address: 'B1', value: 42.5 },
      { address: 'C2', value: true },
      { address: 'D3', value: 'Inline text' },
      { address: 'E4', value: '#N/A' },
      { address: 'A5', value: 'Line\nBreak' },
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: 'Data', startAddress: 'A1', endAddress: 'B1' }])
    expect(imported?.snapshot.sheets[0]?.metadata).toMatchObject({
      columns: [
        { id: 'col:0', index: 0, size: 60 },
        { id: 'col:1', index: 1, size: 60 },
      ],
      columnMetadata: [{ start: 0, count: 2, size: 60, xlsxWidth: 10, customWidth: true }],
      rows: [{ id: 'row:1', index: 1, size: 27 }],
      rowMetadata: [{ start: 1, count: 1, size: 27, xlsxHeight: 20, customHeight: true }],
      sheetFormatPr: { defaultRowHeight: 15 },
    })
    expect(imported?.preview.sheets[0]).toMatchObject({
      rowCount: 5,
      columnCount: 5,
      nonEmptyCellCount: 6,
      previewRows: [
        ['Alpha &#8211;', '42.5', '', '', ''],
        ['', '', 'TRUE', '', ''],
        ['', '', '', 'Inline text', ''],
        ['', '', '', '', '#N/A'],
        ['Line\nBreak', '', '', '', ''],
      ],
    })
  })

  it('imports large value-only worksheets that omit sharedStrings.xml', () => {
    const rows: string[] = []
    for (let row = 1; row <= 2_000; row += 1) {
      rows.push(
        [
          `<row r="${String(row)}">`,
          `<c r="A${String(row)}"><v>${String(row)}</v></c>`,
          `<c r="B${String(row)}" t="inlineStr"><is><t>Row ${String(row)}</t></is></c>`,
          `<c r="C${String(row)}" t="b"><v>${row % 2 === 0 ? '1' : '0'}</v></c>`,
          '</row>',
        ].join(''),
      )
    }
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:C2000"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'numeric-only.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toHaveLength(6_000)
    expect(imported?.snapshot.sheets[0]?.cells.slice(0, 6)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 'Row 1' },
      { address: 'C1', value: false },
      { address: 'A2', value: 2 },
      { address: 'B2', value: 'Row 2' },
      { address: 'C2', value: true },
    ])
    expect(imported?.preview.sheets[0]).toMatchObject({
      rowCount: 2_000,
      columnCount: 3,
      nonEmptyCellCount: 6_000,
    })
    expect(readRuntimeImage(imported!.snapshot)?.sheetCells?.[0]).toMatchObject({
      sheetName: 'Data',
      coords: [],
      coordinateOrder: 'dense-row-major',
      dimensions: { width: 3, height: 2_000 },
      cellCount: 6_000,
    })
  })

  it('preflights public import materialization limits before building snapshot cell objects', () => {
    const rows: string[] = []
    for (let row = 1; row <= 4; row += 1) {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
    }
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A4"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '</worksheet>',
      ].join(''),
    })

    const inspection = inspectXlsx(bytes, 'limit-preflight.xlsx')
    expect(inspection?.stats.cellCount).toBe(4)
    expect(() =>
      importXlsx(bytes, 'limit-preflight.xlsx', {
        limits: { maxMaterializedCells: 3 },
      }),
    ).toThrow(XlsxImportSizeLimitExceededError)

    const imported = importXlsx(bytes, 'limit-preflight.xlsx', { limits: { maxMaterializedCells: 4 } })
    expect(imported.snapshot.sheets[0]?.cells).toHaveLength(4)
  })

  it('preflights formula-heavy imports before WorkPaper build can hit evaluation timeout', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1+1</f><v>2</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(() =>
      importXlsx(bytes, 'formula-limit-preflight.xlsx', {
        limits: { maxMaterializedFormulaCells: 0 },
      }),
    ).toThrow(XlsxImportSizeLimitExceededError)
  })

  it('preflights formula-heavy sheets even when workbook features require SheetJS fallback', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1" cm="1"><v>1</v></c><c r="B1"><f>A1+1</f><v>2</v></c></row></sheetData>',
        '<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>',
        '</worksheet>',
      ].join(''),
    })

    expect(inspectXlsx(bytes, 'fallback-preflight.xlsx')?.stats).toMatchObject({
      cellCount: 2,
      formulaCellCount: 1,
    })
    expect(() =>
      importXlsx(bytes, 'fallback-preflight.xlsx', {
        limits: { maxMaterializedFormulaCells: 0 },
      }),
    ).toThrow(XlsxImportSizeLimitExceededError)
  })

  it('falls back when shared string cells reference a missing sharedStrings part', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'missing-shared-strings.xlsx', unzipSync(bytes), { minByteLength: 0 })).toBeNull()
  })

  it('streams non-materialized large worksheets without inflating worksheet or shared string entries', () => {
    const rows: string[] = []
    for (let row = 1; row <= 128; row += 1) {
      rows.push(
        [
          `<row r="${String(row)}">`,
          `<c r="A${String(row)}" t="s"><v>0</v></c>`,
          `<c r="B${String(row)}"><f>A${String(row)}&amp;"!"</f><v>${String(row)}</v></c>`,
          '</row>',
        ].join(''),
      )
    }
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B128"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '<conditionalFormatting sqref="B1:B128"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>64</formula></cfRule></conditionalFormatting>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should not be inflated for non-materialized import')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'streamed-headless.xlsx', zip, { materializeCells: false, minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([])
    expect(imported?.stats).toMatchObject({
      sheetCount: 1,
      cellCount: 256,
      formulaCellCount: 128,
      valueCellCount: 256,
      mergeCount: 1,
      conditionalFormatCount: 1,
    })
    expect(imported?.stats.dimensions[0]).toMatchObject({
      sheetName: 'Data',
      rowCount: 128,
      columnCount: 2,
      nonEmptyCellCount: 256,
      usedRange: { startRow: 0, startColumn: 0, endRow: 127, endColumn: 1 },
    })
  })

  it('skips worksheet metadata XML retention in verifier-only headless mode', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B2"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row></sheetData>',
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '<conditionalFormatting sqref="B1:B2"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
        '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should not be inflated for headless verifier import')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'headless-verifier.xlsx', zip, {
      materializeCells: false,
      materializeMetadata: false,
      minByteLength: 0,
    })

    expect(imported?.snapshot.sheets[0]?.metadata).toBeUndefined()
    expect(imported?.stats.tableCount).toBe(1)
    expect(imported?.stats.mergeCount).toBe(1)
    expect(imported?.stats.conditionalFormatCount).toBe(1)
  })

  it('streams compressed worksheet zip entries across multiple compressed chunks', () => {
    const rows: string[] = []
    for (let row = 1; row <= 4_096; row += 1) {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row * 17)}</v></c></row>`)
    }
    const worksheetXml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<dimension ref="A1:A4096"/>',
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join('')
    const bytes = buildLargeSimpleWorkbook({ includeSharedStrings: false, worksheetXml })
    const zip = readXlsxZipEntriesLazy(bytes)
    let streamedByteLength = 0

    const streamed = forEachInflatedXlsxZipEntryChunk(
      zip,
      'xl/worksheets/sheet1.xml',
      (chunk) => {
        streamedByteLength += chunk.byteLength
      },
      { chunkSize: 128 },
    )

    expect(streamed).toBe(true)
    expect(streamedByteLength).toBe(strToU8(worksheetXml).byteLength)
  })

  it('streams materialized worksheet cells without inflating the lazy worksheet entry', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:C2"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>1+1</f><v>2</v></c></row>',
        '<row r="2"><c r="C2" t="inlineStr"><is><t>Inline</t></is></c></row></sheetData>',
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('materialized worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'streamed-materialized.xlsx', zip, { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha &#8211;' },
      { address: 'B1', value: 2, formula: '1+1' },
      { address: 'C2', value: 'Inline' },
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: 'Data', startAddress: 'A1', endAddress: 'B1' }])
  })

  it('streams only referenced shared strings for materialized imports', () => {
    const richStringXml = '<si><r><rPr><b/></rPr><t>Rich</t></r><r><t xml:space="preserve"> Value</t></r></si>'
    const bytes = buildLargeSimpleWorkbook({
      sharedStringsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
  <si><t>Unused 0</t></si>
  <si><t>Alpha</t></si>
  <si><t>Unused 2</t></si>
  ${richStringXml}
  <si><t>Unused 4</t></si>
</sst>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>3</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should be streamed instead of fully inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'streamed-shared-strings.xlsx', zip, { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha' },
      { address: 'B1', value: 'Rich Value' },
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.richTextArtifacts).toEqual({
      cells: [
        {
          address: 'B1',
          text: 'Rich Value',
          storage: 'sharedString',
          xml: richStringXml,
        },
      ],
    })
  })

  it('collects materialized shared-string references during the real worksheet stream', () => {
    const bytes = buildLargeSimpleWorkbook({
      sharedStringsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Unused</t></si>
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
</sst>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>2</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    const worksheetStreamCount = countLazyZipEntryStreams(zip, 'xl/worksheets/sheet1.xml')
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should be streamed instead of fully inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'single-pass-shared-strings.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha' },
      { address: 'B1', value: 'Beta' },
    ])
    expect(worksheetStreamCount()).toBe(1)
    expect(imported?.stats.phaseTelemetry.map((entry) => entry.phase)).toEqual([
      'zip-setup',
      'worksheet-scan',
      'metadata-parsing',
      'shared-string-resolution',
      'style-parsing',
      'zip-source-release',
      'public-snapshot-materialization',
    ])
    expect(imported?.stats.phaseTelemetry.every((entry) => Number.isInteger(entry.elapsedMs) && entry.elapsedMs >= 0)).toBe(true)
    expect(imported?.stats.phaseTelemetry.every((entry) => (entry.rssBytes ?? 0) > 0)).toBe(true)
  })

  it('imports workbook defined names without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      definedNamesXml:
        '<definedNames><definedName name="BrokenReference">#REF!</definedName><definedName name="ScalarLimit">42</definedName></definedNames>',
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>123</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'defined-names.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'BrokenReference', value: { kind: 'formula', formula: '=#REF!' } },
      { name: 'ScalarLimit', value: { kind: 'scalar', value: 42 } },
    ])
    expect(imported?.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 123 }])
  })

  it('imports rich shared strings and preserves their cell artifacts without falling back to SheetJS', () => {
    const richStringXml = '<si><r><rPr><b/></rPr><t>Rich</t></r><r><rPr><i/></rPr><t xml:space="preserve"> Text</t></r></si>'
    const bytes = buildLargeSimpleWorkbook({
      sharedStringsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">${richStringXml}</sst>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'rich-shared-string.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 'Rich Text' }])
    expect(imported?.snapshot.sheets[0]?.metadata?.richTextArtifacts).toEqual({
      cells: [
        {
          address: 'A1',
          text: 'Rich Text',
          storage: 'sharedString',
          xml: richStringXml,
        },
      ],
    })
  })

  it('imports simple hyperlink and drawing metadata without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      sheetRelationshipsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHyperlink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/report" TargetMode="External"/>
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<dimension ref="A1:B2"/>',
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Open report</t></is></c><c r="B1"><v>7</v></c></row></sheetData>',
        '<hyperlinks>',
        '<hyperlink ref="A1" r:id="rIdHyperlink1" tooltip="Open report" display="Report"/>',
        '<hyperlink ref="B2" location="Data!A1" display="Jump"/>',
        '</hyperlinks>',
        '<drawing r:id="rIdDrawing1"/>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        '[Content_Types].xml': [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="png" ContentType="image/png"/>',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
          '</Types>',
        ].join(''),
        'xl/drawings/drawing1.xml': [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
          '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>3</xdr:row></xdr:to>',
          '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>',
          '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImage1"/></xdr:blipFill>',
          '<xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
        ].join(''),
        'xl/drawings/_rels/drawing1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
        'xl/media/image1.png': Uint8Array.from([137, 80, 78, 71]),
      },
    })

    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('drawing metadata should use the streamed worksheet relationship id instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'hyperlinks-drawing.xlsx', zip, { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.metadata?.hyperlinks).toEqual([
      { sheetName: 'Data', address: 'A1', target: 'https://example.com/report', tooltip: 'Open report', display: 'Report' },
      { sheetName: 'Data', address: 'B2', target: '#Data!A1', display: 'Jump' },
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.drawingArtifacts).toEqual({ relationshipTarget: '../drawings/drawing1.xml' })
    expect(imported?.snapshot.workbook.metadata?.drawingArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/drawings/_rels/drawing1.xml.rels',
      'xl/drawings/drawing1.xml',
      'xl/media/image1.png',
    ])
  })

  it('imports print page setup and binary printer settings without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      sheetRelationshipsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPrinterSettings1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>
</Relationships>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"/>',
        '<pageSetup orientation="landscape" r:id="rIdPrinterSettings1"/>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'xl/printerSettings/printerSettings1.bin': Uint8Array.from([1, 2, 3, 4]),
      },
    })

    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('print metadata should use streamed typed records instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'printer-settings.xlsx', zip, { minByteLength: 0 })
    const metadata = imported?.snapshot.sheets[0]?.metadata

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 7 }])
    expect(metadata?.printPageSetup).toEqual({
      pageMarginsXml: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"/>',
      pageSetupXml: '<pageSetup orientation="landscape" r:id="rIdPrinterSettings1"/>',
    })
    expect(metadata?.printerSettings).toEqual([
      {
        relationshipTarget: '../printerSettings/printerSettings1.bin',
        storage: 'base64',
        dataBase64: 'AQIDBA==',
        byteLength: 4,
        pageSetupXml: '<pageSetup orientation="landscape" r:id="rIdPrinterSettings1"/>',
      },
    ])
  })

  it('preserves Power Pivot data model artifacts on the streaming path without releasing their lazy ZIP source', () => {
    const modelBytes = deterministicBytes(1_000_000)
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdModel" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData" Target="model/item.data"/>
</Relationships>`,
        'xl/model/item.data': modelBytes,
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'power-pivot-model.xlsx', readXlsxZipEntriesLazy(bytes), {
      minByteLength: 0,
      releaseZipSource: true,
    })
    const parts = imported?.snapshot.workbook.metadata?.dataModelArtifacts?.parts ?? []
    const modelPart = parts.find((part) => part.path === 'xl/model/item.data')

    expect(imported?.stats.cellCount).toBe(1)
    expect(parts.map((part) => part.path)).toEqual(['xl/model/item.data'])
    expect(modelPart?.byteLength).toBe(modelBytes.byteLength)
    expect(decodeBase64(modelPart?.dataBase64 ?? '')).toEqual(modelBytes)
  })

  it('preserves streamed cell metadata references without falling back to SheetJS', () => {
    const metadataXml =
      '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>'
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1" vm="1"><v>7</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdCellMetadata" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`,
        'xl/metadata.xml': metadataXml,
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'cell-metadata.xlsx', readXlsxZipEntriesLazy(bytes), { minByteLength: 0 })

    expect(imported?.stats.cellCount).toBe(1)
    expect(imported?.snapshot.workbook.metadata?.cellMetadata).toEqual({
      relationshipTarget: 'metadata.xml',
      metadataXml,
    })
    expect(imported?.snapshot.sheets[0]?.metadata?.cellMetadataRefs).toEqual([
      {
        address: 'A1',
        vm: '1',
        cellSignature: JSON.stringify({ value: 7, formula: null, format: null }),
      },
    ])
  })

  it('preserves slicer connection artifacts on the streaming path without inflating worksheet XML for metadata', () => {
    const sheetSlicerListExtXml =
      '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList><x14:slicer r:id="rIdSlicer"/></x14:slicerList></ext>'
    const workbookSlicerCachesExtXml =
      '<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:slicerCaches><x15:slicerCache r:id="rIdSlicerCache"/></x15:slicerCaches></ext>'
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        `<extLst>${sheetSlicerListExtXml}</extLst>`,
        '</worksheet>',
      ].join(''),
      sheetRelationshipsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlicer" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
</Relationships>`,
      extraEntries: {
        'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <extLst>${workbookSlicerCachesExtXml}</extLst>
</workbook>`,
        'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdSlicerCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
  <Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
</Relationships>`,
        'xl/connections.xml': '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>',
        'xl/slicerCaches/slicerCache1.xml':
          '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>',
        'xl/slicers/slicer1.xml': '<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Region"/>',
      },
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('slicer metadata should use streamed extLst records instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'slicer-connections.xlsx', zip, { minByteLength: 0 })
    const artifacts = imported?.snapshot.workbook.metadata?.slicerConnectionArtifacts

    expect(imported?.stats.cellCount).toBe(1)
    expect(artifacts?.workbookSlicerCachesExtXml).toBe(workbookSlicerCachesExtXml)
    expect(artifacts?.sheetArtifacts).toEqual([
      {
        sheetName: 'Data',
        sheetSlicerListExtXml,
        relationships: [
          {
            id: 'rIdSlicer',
            type: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
            target: '../slicers/slicer1.xml',
          },
        ],
      },
    ])
    expect(artifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/connections.xml',
      'xl/slicerCaches/slicerCache1.xml',
      'xl/slicers/slicer1.xml',
    ])
  })

  it('imports worksheet auto filters without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B3"/>',
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Status</t></is></c><c r="B1" t="inlineStr"><is><t>Total</t></is></c></row>',
        '<row r="2"><c r="A2" t="inlineStr"><is><t>Open</t></is></c><c r="B2"><v>10</v></c></row></sheetData>',
        '<autoFilter ref="A1:B3"><filterColumn colId="0"><filters><filter val="Open"/></filters></filterColumn></autoFilter>',
        '</worksheet>',
      ].join(''),
    })

    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('auto filter metadata should use streamed typed records instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'auto-filter.xlsx', zip, { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.metadata?.filters).toEqual([
      {
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B3',
        criteria: [{ colId: 0, filters: { values: ['Open'] } }],
      },
    ])
  })

  it('imports table metadata without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      sheetRelationshipsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<dimension ref="A1:B3"/>',
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Total</t></is></c></row>',
        '<row r="2"><c r="A2" t="inlineStr"><is><t>One</t></is></c><c r="B2"><v>10</v></c></row></sheetData>',
        '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'xl/tables/table1.xml': [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTable" displayName="SalesTable" ref="A1:B3" headerRowCount="1" totalsRowShown="0">',
          '<autoFilter ref="A1:B3"/>',
          '<tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Total"/></tableColumns>',
          '<tableStyleInfo name="TableStyleMedium4" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
          '</table>',
        ].join(''),
      },
    })

    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('table metadata should use streamed relationship ids instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'table.xlsx', zip, { minByteLength: 0 })

    expect(imported?.snapshot.workbook.metadata?.tables).toEqual([
      {
        name: 'SalesTable',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B3',
        columnNames: ['Name', 'Total'],
        headerRow: true,
        totalsRow: false,
        style: {
          name: 'TableStyleMedium4',
          showColumnStripes: false,
          showFirstColumn: false,
          showLastColumn: false,
          showRowStripes: true,
        },
      },
    ])
  })

  it('imports conditional formatting metadata without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A2"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData>',
        '<conditionalFormatting sqref="A1:A2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'xl/styles.xml': [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></dxf></dxfs>',
          '</styleSheet>',
        ].join(''),
      },
    })

    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('conditional formatting metadata should use streamed records instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'conditional-format.xlsx', zip, { minByteLength: 0 })
    const metadata = imported?.snapshot.sheets[0]?.metadata

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 7 },
      { address: 'A2', value: 1 },
    ])
    expect(metadata?.conditionalFormats).toEqual([
      {
        id: 'xlsx-cf:Data:A1:A2:1',
        range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A2' },
        rule: { kind: 'cellIs', operator: 'greaterThan', values: [3] },
        style: { fill: { backgroundColor: '#ffcc00' } },
        priority: 1,
      },
    ])
    expect(metadata?.conditionalFormatArtifacts?.xml).toContain('type="cellIs"')
  })

  it('preserves unsupported conditional formatting artifacts on the large simple path', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A2"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="A2"><v>7</v></c></row></sheetData>',
        '<conditionalFormatting sqref="A1:A2"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>',
        '</worksheet>',
      ].join(''),
    })

    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('unsupported conditional formatting artifacts should use streamed XML instead of inflating worksheet XML')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'duplicate-format.xlsx', zip, { minByteLength: 0 })
    const metadata = imported?.snapshot.sheets[0]?.metadata

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 7 },
      { address: 'A2', value: 7 },
    ])
    expect(metadata?.conditionalFormats).toBeUndefined()
    expect(metadata?.conditionalFormatArtifacts?.xml).toContain('type="duplicateValues"')
  })

  it('coalesces contiguous same-style cells into compact style ranges', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      includeStyles: true,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:C2"/>',
        '<sheetData>',
        '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>2</v></c><c r="C1"><v>3</v></c></row>',
        '<row r="2"><c r="A2" s="1"><v>4</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    })

    const zip = readXlsxZipEntriesLazy(bytes)
    const stylesStreamCount = countLazyZipEntryStreams(zip, 'xl/styles.xml')
    Object.defineProperty(zip, 'xl/styles.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('styles.xml should be streamed instead of fully inflated')
      },
    })
    const imported = tryImportLargeSimpleXlsx(bytes, 'coalesced-styles.xlsx', zip, { minByteLength: 0 })
    const styleRanges = imported?.snapshot.sheets[0]?.metadata?.styleRanges
    const styleId = styleRanges?.[0]?.styleId

    expect(styleRanges).toEqual([
      { range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B1' }, styleId },
      { range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'A2' }, styleId },
    ])
    expect(imported?.snapshot.workbook.metadata?.styles).toEqual([{ id: styleId, fill: { backgroundColor: '#ffcc00' } }])
    expect(stylesStreamCount()).toBe(2)
  })

  it('compacts broad style-only blank templates into style ranges without SheetJS fallback', () => {
    const rows: string[] = []
    for (let row = 1; row <= 3_001; row += 1) {
      const cells: string[] = []
      for (let column = 0; column < 20; column += 1) {
        const address = `${encodeColumnName(column)}${String(row)}`
        cells.push(address === 'A1' ? '<c r="A1" s="1"><v>123</v></c>' : `<c r="${address}" s="1"/>`)
      }
      rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
    }
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      includeStyles: true,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:T3001"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'styled-blank-template.xlsx', unzipSync(bytes), { minByteLength: 0 })
    const styleRanges = imported?.snapshot.sheets[0]?.metadata?.styleRanges
    const styleId = styleRanges?.[0]?.styleId

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 123 }])
    expect(styleRanges).toEqual([{ range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'T3001' }, styleId }])
  })

  it('imports simple formula cells with cached values without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'formula.xlsx', unzipSync(bytes), { minByteLength: 0 })?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 2, formula: '1+1' },
    ])
  })

  it('imports formula-only cells from the compact scanner and previews the formula text', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><f>SUM(B1:C1)</f></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'formula-only.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', formula: 'SUM(B1:C1)' }])
    expect(imported?.preview.sheets[0]?.previewRows[0]?.[0]).toBe('=SUM(B1:C1)')
  })

  it('scans namespaced worksheet cell tags without full SheetJS fallback', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<x:dimension ref="A1:B1"/>',
        '<x:sheetData><x:row r="1"><x:c r="A1"><x:v>3</x:v></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>Namespaced</x:t></x:is></x:c></x:row></x:sheetData>',
        '</x:worksheet>',
      ].join(''),
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'namespaced.xlsx', unzipSync(bytes), { minByteLength: 0 })?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 3 },
      { address: 'B1', value: 'Namespaced' },
    ])
  })

  it('expands shared formulas without falling back to SheetJS', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A2"/>',
        '<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>2</v></c></row>',
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>3</v></c><c r="B2"><v>2</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(
      tryImportLargeSimpleXlsx(bytes, 'shared-formula.xlsx', unzipSync(bytes), { minByteLength: 0 })?.snapshot.sheets[0]?.cells,
    ).toEqual([
      { address: 'A1', value: 2, formula: 'B1+1' },
      { address: 'A2', value: 3, formula: 'B2+1' },
      { address: 'B2', value: 2 },
    ])
  })

  it('falls back when array formulas require full formula artifact preservation', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A2"/>',
        '<sheetData><row r="1"><c r="A1"><f t="array" ref="A1:A2">B1:B2*2</f><v>2</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'array-formula.xlsx', unzipSync(bytes), { minByteLength: 0 })).toBeNull()
  })

  it('falls back when structured table references require formula translation', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><f>SUM(SalesTable[Total])</f><v>10</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'structured-formula.xlsx', unzipSync(bytes), { minByteLength: 0 })).toBeNull()
  })

  it('does not fall back solely because preserved chart package parts are present', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>123</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'xl/charts/chart1.xml': '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
      },
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'chart.xlsx', unzipSync(bytes), { minByteLength: 0 })?.stats.cellCount).toBe(1)
    expect(inspectXlsx(bytes, 'chart.xlsx')?.stats.cellCount).toBe(1)
  })

  it('streams large cached formula workbooks with calcChain below the old byte threshold', () => {
    const rows: string[] = []
    for (let row = 1; row <= 50_001; row += 1) {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><f>B${String(row)}+1</f><v>${String(row)}</v></c></row>`)
    }
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A50001"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'xl/calcChain.xml': '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>',
        'docProps/padding.bin': deterministicBytes(1_100_000),
      },
    })

    const imported = importXlsx(bytes, 'cached-formula-calcchain.xlsx')

    expect(imported.stats?.formulaCellCount).toBe(50_001)
    expect(imported.snapshot.sheets[0]?.cells).toHaveLength(50_001)
  })
})

function buildLargeSimpleWorkbook(input: {
  worksheetXml: string
  definedNamesXml?: string
  extraEntries?: Record<string, string | Uint8Array>
  includeSharedStrings?: boolean
  includeStyles?: boolean
  sharedStringsXml?: string
  sheetRelationshipsXml?: string
}): Uint8Array {
  const includeSharedStrings = input.includeSharedStrings ?? true
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  ${input.definedNamesXml ?? ''}
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  ${
    includeSharedStrings
      ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
      : ''
  }
</Relationships>`),
    ...(includeSharedStrings
      ? {
          'xl/sharedStrings.xml': strToU8(
            input.sharedStringsXml ??
              `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Alpha &amp;#8211;</t></si>
  <si><t>Line_x000a_Break</t></si>
</sst>`,
          ),
        }
      : {}),
    ...(input.sheetRelationshipsXml ? { 'xl/worksheets/_rels/sheet1.xml.rels': strToU8(input.sheetRelationshipsXml) } : {}),
    ...(input.includeStyles
      ? {
          'xl/styles.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0"/></cellXfs>
</styleSheet>`),
        }
      : {}),
    'xl/worksheets/sheet1.xml': strToU8(input.worksheetXml),
    ...Object.fromEntries(
      Object.entries(input.extraEntries ?? {}).map(([path, value]) => [path, typeof value === 'string' ? strToU8(value) : value]),
    ),
  })
}

function countLazyZipEntryStreams(zip: Record<string, Uint8Array>, path: string): () => number {
  const metadata = readLazyZipMetadata(zip)
  const entry = metadata?.entriesByPath.get(path)
  if (!metadata || !entry) {
    throw new Error(`Missing lazy ZIP metadata for ${path}`)
  }
  const source = metadata.source
  const localHeader = source.readRange(entry.localHeaderOffset, entry.localHeaderOffset + 30)
  const fileNameLength = readLittleEndianUint16(localHeader, 26)
  const extraFieldLength = readLittleEndianUint16(localHeader, 28)
  const dataStart = entry.localHeaderOffset + 30 + fileNameLength + extraFieldLength
  const dataEnd = dataStart + entry.compressedSize
  let streamCount = 0
  metadata.source = new Proxy(source, {
    get(target, property) {
      if (property === 'readRange') {
        return (start?: number, end?: number) => {
          if (start === dataStart && end === dataEnd) {
            streamCount += 1
          }
          return target.readRange(start ?? 0, end ?? target.byteLength)
        }
      }
      const value = Reflect.get(target, property, target)
      return typeof value === 'function' ? value.bind(target) : value
    },
  })
  return () => streamCount
}

function readLazyZipMetadata(zip: Record<string, Uint8Array>):
  | {
      source: XlsxLazyZipByteSource
      readonly entriesByPath: ReadonlyMap<
        string,
        {
          readonly localHeaderOffset: number
          readonly compressedSize: number
        }
      >
    }
  | undefined {
  for (const symbol of Object.getOwnPropertySymbols(zip)) {
    const value = Reflect.get(zip, symbol) as unknown
    if (isLazyZipMetadata(value)) {
      return value
    }
  }
  return undefined
}

function isLazyZipMetadata(value: unknown): value is {
  source: XlsxLazyZipByteSource
  readonly entriesByPath: ReadonlyMap<
    string,
    {
      readonly localHeaderOffset: number
      readonly compressedSize: number
    }
  >
} {
  return (
    typeof value === 'object' &&
    value !== null &&
    'source' in value &&
    isLazyZipByteSource(value.source) &&
    'entriesByPath' in value &&
    value.entriesByPath instanceof Map
  )
}

interface XlsxLazyZipByteSource {
  readonly byteLength: number
  readRange(start: number, end: number): Uint8Array
}

function isLazyZipByteSource(value: unknown): value is XlsxLazyZipByteSource {
  return (
    typeof value === 'object' &&
    value !== null &&
    'byteLength' in value &&
    typeof value.byteLength === 'number' &&
    'readRange' in value &&
    typeof value.readRange === 'function'
  )
}

function readLittleEndianUint16(source: Uint8Array, offset: number): number {
  return source[offset] | (source[offset + 1] << 8)
}

function deterministicBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  let state = 0x12345678
  for (let index = 0; index < length; index += 1) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    bytes[index] = (state >>> 24) & 0xff
  }
  return bytes
}

function decodeBase64(value: string): Uint8Array {
  return new Uint8Array(Buffer.from(value, 'base64'))
}

function encodeColumnName(index: number): string {
  let value = index + 1
  let output = ''
  while (value > 0) {
    value -= 1
    output = String.fromCharCode(65 + (value % 26)) + output
    value = Math.floor(value / 26)
  }
  return output
}
