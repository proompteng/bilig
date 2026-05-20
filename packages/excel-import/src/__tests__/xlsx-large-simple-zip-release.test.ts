import { strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { importXlsx } from '../index.js'
import { exportXlsx } from '../xlsx-export.js'
import { tryInspectLargeSimpleXlsxHeadless } from '../xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { readLazyXlsxZipSourceByteLength, readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple XLSX import ZIP ownership', () => {
  it('records lazy ZIP source release before shared-string snapshot materialization', () => {
    let ownedBytes = buildSharedStringWorkbook()
    const zip = readXlsxZipEntriesLazy(ownedBytes)
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

    const imported = tryImportLargeSimpleXlsx({ byteLength: ownedBytes.byteLength }, 'zip-release.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
      releaseOwnedSourceBytes: () => {
        const ownedSourceBytesBeforeRelease = ownedBytes.byteLength
        ownedBytes = new Uint8Array(0)
        return {
          ownedSourceBytesBeforeRelease,
          ownedSourceBytesAfterRelease: ownedBytes.byteLength,
        }
      },
    })

    const phases = imported?.stats.phaseTelemetry.map((entry) => entry.phase) ?? []
    const releasePhase = imported?.stats.phaseTelemetry.find((entry) => entry.phase === 'zip-source-release')
    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha' },
      { address: 'B1', value: 'Beta' },
    ])
    expect(releasePhase).toMatchObject({
      zipSourceBytesBeforeRelease: releasePhase?.ownedSourceBytesBeforeRelease,
      zipSourceBytesAfterRelease: 0,
      ownedSourceBytesAfterRelease: 0,
    })
    expect(releasePhase?.ownedSourceBytesBeforeRelease).toBeGreaterThan(0)
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
    expect(ownedBytes.byteLength).toBe(0)
    expect(phases.indexOf('zip-source-release')).toBeGreaterThanOrEqual(0)
    expect(phases.indexOf('zip-source-release')).toBeLessThan(phases.indexOf('public-snapshot-materialization'))
  })

  it('records owned source release for the headless verifier path', () => {
    let ownedBytes = buildSharedStringWorkbook()
    const zip = readXlsxZipEntriesLazy(ownedBytes)

    const inspected = tryInspectLargeSimpleXlsxHeadless({ byteLength: ownedBytes.byteLength }, 'headless-zip-release.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
      releaseOwnedSourceBytes: () => {
        const ownedSourceBytesBeforeRelease = ownedBytes.byteLength
        ownedBytes = new Uint8Array(0)
        return {
          ownedSourceBytesBeforeRelease,
          ownedSourceBytesAfterRelease: ownedBytes.byteLength,
        }
      },
    })

    const releasePhase = inspected?.stats.phaseTelemetry.find((entry) => entry.phase === 'zip-source-release')
    expect(inspected?.stats.cellCount).toBe(2)
    expect(releasePhase?.zipSourceBytesAfterRelease).toBe(0)
    expect(releasePhase?.ownedSourceBytesBeforeRelease).toBeGreaterThan(0)
    expect(releasePhase?.ownedSourceBytesAfterRelease).toBe(0)
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
    expect(ownedBytes.byteLength).toBe(0)
  })

  it('releases public import source bytes instead of attaching an untouched export copy', () => {
    const bytes = buildSharedStringWorkbook({
      'customXml/item1.xml': strToU8('<root><value>preserved data model artifact</value></root>'),
      'docProps/padding.bin': deterministicBytes(1_200_000),
    })

    const imported = importXlsx(bytes, 'public-zip-release.xlsx')
    const releasePhase = imported.stats?.phaseTelemetry.find((entry) => entry.phase === 'zip-source-release')
    const exported = exportXlsx(imported.snapshot)
    const roundTripped = importXlsx(exported, 'roundtrip.xlsx')

    expect(imported.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha' },
      { address: 'B1', value: 'Beta' },
    ])
    expect(releasePhase).toMatchObject({
      zipSourceBytesAfterRelease: 0,
      ownedSourceBytesAfterRelease: 0,
    })
    expect(releasePhase?.ownedSourceBytesBeforeRelease).toBeGreaterThan(1_000_000)
    expect(imported.snapshot.workbook.metadata?.dataModelArtifacts?.parts[0]?.dataBase64).toBeTruthy()
    expect(unzipSync(exported)['docProps/padding.bin']).toBeUndefined()
    expect(unzipSync(exported)['customXml/item1.xml']).toBeDefined()
    expect(roundTripped.snapshot.sheets[0]?.cells.map(({ address, value }) => ({ address, value }))).toEqual(
      imported.snapshot.sheets[0]?.cells.map(({ address, value }) => ({ address, value })),
    )
  })
})

function buildSharedStringWorkbook(extraEntries: Readonly<Record<string, Uint8Array>> = {}): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
</sst>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B1"/>
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>
</worksheet>`),
    ...extraEntries,
  })
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
