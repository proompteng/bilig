import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { importXlsxFromZipByteSource } from '../xlsx-byte-source-import.js'
import { exportXlsx } from '../xlsx-export.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple XLSX lazy package artifacts', () => {
  it('imports from an internal byte source without attaching an in-memory untouched export copy', () => {
    const bytes = buildWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'docProps/padding.bin': deterministicBytes(1_200_000),
      },
    })
    const source = new CountingXlsxZipByteSource(bytes)

    const imported = importXlsxFromZipByteSource(source, 'byte-source.xlsx')

    expect(imported.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 7 }])
    expect(source.fullReadCount).toBe(0)
    expect(source.releaseCount).toBe(0)
    expect(exportXlsx(imported.snapshot)).toEqual(bytes)
    expect(source.fullReadCount).toBe(1)
    expect(source.releaseCount).toBe(0)
  })

  it('keeps pivot cache package XML lazy until export materialization needs it', () => {
    const pivotCacheRecordsXml = `<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><r>${'x'.repeat(
      400_000,
    )}</r></pivotCacheRecords>`
    const bytes = buildWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<pivotTableDefinition r:id="rIdPivot"/>',
        '</worksheet>',
      ].join(''),
      sheetRelationshipsXml: relationshipXml([
        {
          id: 'rIdPivot',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
          target: '../pivotTables/pivotTable1.xml',
        },
      ]),
      workbookRelationshipsXml: relationshipXml([
        {
          id: 'rId1',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          target: 'worksheets/sheet1.xml',
        },
        {
          id: 'rIdPivotCache',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
          target: 'pivotCache/pivotCacheDefinition1.xml',
        },
      ]),
      workbookExtraXml: '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache"/></pivotCaches>',
      extraEntries: {
        'xl/pivotTables/pivotTable1.xml':
          '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1"/>',
        'xl/pivotCache/pivotCacheDefinition1.xml':
          '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="1"/>',
        'xl/pivotCache/pivotCacheRecords1.xml': pivotCacheRecordsXml,
      },
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    const pivotRecordsStreams = countLazyZipEntryStreams(zip, 'xl/pivotCache/pivotCacheRecords1.xml')

    const imported = tryImportLargeSimpleXlsx({ byteLength: bytes.byteLength }, 'lazy-pivot.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })
    const pivotRecordsPart = imported?.snapshot.workbook.metadata?.pivotArtifacts?.parts.find(
      (part) => part.path === 'xl/pivotCache/pivotCacheRecords1.xml',
    )

    expect(imported?.stats.cellCount).toBe(1)
    expect(pivotRecordsStreams()).toBe(0)
    expect(pivotRecordsPart?.xml).toBe(pivotCacheRecordsXml)
    expect(pivotRecordsStreams()).toBe(1)
  })

  it('keeps drawing dependency binaries lazy until their preserved part is read', () => {
    const imageBytes = deterministicBytes(700_000)
    const bytes = buildWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>',
        '<drawing r:id="rIdDrawing"/>',
        '</worksheet>',
      ].join(''),
      sheetRelationshipsXml: relationshipXml([
        {
          id: 'rIdDrawing',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          target: '../drawings/drawing1.xml',
        },
      ]),
      extraEntries: {
        'xl/drawings/drawing1.xml': [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
          '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor>',
          '</xdr:wsDr>',
        ].join(''),
        'xl/drawings/_rels/drawing1.xml.rels': relationshipXml([
          {
            id: 'rIdImage',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            target: '../media/image1.png',
          },
        ]),
        'xl/media/image1.png': imageBytes,
      },
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    const imageStreams = countLazyZipEntryStreams(zip, 'xl/media/image1.png')

    const imported = tryImportLargeSimpleXlsx({ byteLength: bytes.byteLength }, 'lazy-drawing.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })
    const imagePart = imported?.snapshot.workbook.metadata?.drawingArtifacts?.parts.find((part) => part.path === 'xl/media/image1.png')

    expect(imported?.stats.cellCount).toBe(1)
    expect(imageStreams()).toBe(0)
    expect(decodeBase64(imagePart?.dataBase64 ?? '')).toEqual(imageBytes)
    expect(imageStreams()).toBe(1)
  })
})

function buildWorkbook(input: {
  readonly worksheetXml: string
  readonly sheetRelationshipsXml?: string
  readonly workbookRelationshipsXml?: string
  readonly workbookExtraXml?: string
  readonly extraEntries?: Readonly<Record<string, string | Uint8Array>>
}): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  ${input.workbookExtraXml ?? ''}
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(
      input.workbookRelationshipsXml ??
        relationshipXml([
          {
            id: 'rId1',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
            target: 'worksheets/sheet1.xml',
          },
        ]),
    ),
    ...(input.sheetRelationshipsXml ? { 'xl/worksheets/_rels/sheet1.xml.rels': strToU8(input.sheetRelationshipsXml) } : {}),
    'xl/worksheets/sheet1.xml': strToU8(input.worksheetXml),
    ...Object.fromEntries(
      Object.entries(input.extraEntries ?? {}).map(([path, value]) => [path, typeof value === 'string' ? strToU8(value) : value]),
    ),
  })
}

function relationshipXml(
  relationships: readonly {
    readonly id: string
    readonly type: string
    readonly target: string
  }[],
): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships
    .map((relationship) => `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`)
    .join('')}</Relationships>`
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

class CountingXlsxZipByteSource {
  readonly byteLength: number
  fullReadCount = 0
  releaseCount = 0

  constructor(private readonly bytes: Uint8Array) {
    this.byteLength = bytes.byteLength
  }

  readRange(start: number, end: number): Uint8Array {
    if (start === 0 && end === this.byteLength) {
      this.fullReadCount += 1
    }
    return this.bytes.subarray(start, end)
  }

  release(): void {
    this.releaseCount += 1
  }
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
