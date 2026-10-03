import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { readRuntimeImage } from '@bilig/core'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import {
  createFileImportedXlsxSourceReader,
  exportXlsx,
  exportXlsxSourceLiteralPatches,
  exportXlsxSourceLiteralPatchesToFile,
  exportXlsxSourceLiteralPatchesToFileAsync,
  importXlsx,
  importXlsxFromZipByteSource,
} from '@bilig/excel-import'
import { WorkPaper } from '../index.js'
import { exportWorkPaperXlsxToFileAsync, importXlsxFile } from '../xlsx.js'

const sourcePreservingOutputCalculationSettings = Symbol.for('bilig.sourcePreservingXlsxOutputCalculationSettings')

function sourceWorkbookBytes(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Data',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 1 },
            { address: 'B1', row: 0, col: 1, formula: 'A1+1', value: 2 },
            { address: 'A2', row: 1, col: 0, value: 2 },
            { address: 'B2', row: 1, col: 1, formula: 'A2+1', value: 3 },
          ],
        },
      ],
    }),
  )
  zip['customXml/item1.xml'] = strToU8('<keep source="true"/>')
  return zipSync(zip)
}

function sourceWorkbookWithoutFormulaCachesBytes(): Uint8Array {
  const zip = unzipSync(sourceWorkbookBytes())
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(sheetXml.replace(/<v>2<\/v>/gu, '').replace(/<v>3<\/v>/gu, ''))
  return zipSync(zip)
}

function largeSourceWorkbookBytes(): Uint8Array {
  const zip = unzipSync(sourceWorkbookBytes())
  zip['docProps/padding.bin'] = deterministicBytes(2_000_000)
  return zipSync(zip)
}

function fallbackSourceWorkbookBytes(): Uint8Array {
  const zip = unzipSync(sourceWorkbookBytes())
  zip['xl/threadedComments/threadedComment1.xml'] = strToU8(
    '<threadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>',
  )
  zip['docProps/fallback-padding.bin'] = deterministicBytes(2_000_000)
  return zipSync(zip)
}

function deterministicBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  let state = 0x9e3779b9
  for (let index = 0; index < bytes.length; index += 1) {
    state ^= state << 13
    state ^= state >>> 17
    state ^= state << 5
    bytes[index] = state & 0xff
  }
  return bytes
}

function attachSourceBytesForTest(snapshot: object, bytes: Uint8Array): void {
  Object.defineProperty(snapshot, Symbol.for('bilig.importedXlsxSourceBytes'), {
    configurable: true,
    enumerable: false,
    value: bytes,
  })
}

function attachSourceReaderForTest(snapshot: object, bytes: Uint8Array): void {
  Object.defineProperty(snapshot, Symbol.for('bilig.importedXlsxSourceBytes'), {
    configurable: true,
    enumerable: false,
    value: {
      byteLength: bytes.byteLength,
      readBytes() {
        throw new Error('Source-preserving export should not materialize reader bytes')
      },
      readRange(start: number, end: number) {
        return bytes.subarray(start, end)
      },
      readRangeInto(start: number, end: number, target: Uint8Array) {
        const chunk = bytes.subarray(start, end)
        target.set(chunk)
        return target.subarray(0, chunk.byteLength)
      },
    },
  })
}

function attachSourcePatchesForTest(
  snapshot: object,
  patches: readonly {
    readonly sheetName: string
    readonly address: string
    readonly value: string | number | boolean | null
    readonly preserveFormula?: boolean
  }[],
): void {
  Object.defineProperty(snapshot, Symbol.for('bilig.importedXlsxSourceCellPatches'), {
    configurable: true,
    enumerable: false,
    value: patches.map((patch) => ({ kind: 'literal', ...patch })),
  })
}

function sourceReaderForTest(bytes: Uint8Array): {
  readonly byteLength: number
  readBytes(): Uint8Array
  readRange(start: number, end: number): Uint8Array
  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array
} {
  return {
    byteLength: bytes.byteLength,
    readBytes() {
      throw new Error('Source-preserving export should not materialize reader bytes')
    },
    readRange(start: number, end: number) {
      return bytes.subarray(start, end)
    },
    readRangeInto(start: number, end: number, target: Uint8Array) {
      const chunk = bytes.subarray(start, end)
      target.set(chunk)
      return target.subarray(0, chunk.byteLength)
    },
  }
}

function instrumentedSourceReaderForTest(bytes: Uint8Array): {
  readonly byteLength: number
  readonly fullRangeReadCount: () => number
  readonly releaseCount: () => number
  readBytes(): Uint8Array
  readRange(start: number, end: number): Uint8Array
  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array
  release(): void
} {
  let fullRangeReads = 0
  let releases = 0
  return {
    byteLength: bytes.byteLength,
    fullRangeReadCount() {
      return fullRangeReads
    },
    releaseCount() {
      return releases
    },
    readBytes() {
      fullRangeReads += 1
      return bytes
    },
    readRange(start: number, end: number) {
      if (start === 0 && end === bytes.byteLength) {
        fullRangeReads += 1
      }
      return bytes.subarray(start, end)
    },
    readRangeInto(start: number, end: number, target: Uint8Array) {
      if (start === 0 && end === bytes.byteLength) {
        fullRangeReads += 1
      }
      const chunk = bytes.subarray(start, end)
      target.set(chunk)
      return target.subarray(0, chunk.byteLength)
    },
    release() {
      releases += 1
    },
  }
}

describe('WorkPaper source-preserving XLSX export', () => {
  it('exports imported XLSX scalar edits by patching the source package without a runtime image snapshot', () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving.xlsx')
    attachSourceBytesForTest(imported.snapshot, sourceBytes)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      workbook.setCellContents({ sheet: 1, row: 0, col: 0 }, 5)

      const exportedSnapshot = workbook.exportSnapshot()
      expect(readRuntimeImage(exportedSnapshot)).toBeUndefined()

      const exportedZip = unzipSync(exportXlsx(exportedSnapshot))
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
      expect(exportedZip['xl/calcChain.xml']).toBeUndefined()

      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('r="A1"')
      expect(sheetXml).toContain('<v>5</v>')
      expect(sheetXml).toContain('<f>A1+1</f><v>2</v>')

      const workbookXml = strFromU8(exportedZip['xl/workbook.xml'] ?? new Uint8Array())
      expect(workbookXml).toContain('fullCalcOnLoad="1"')
      expect(workbookXml).toContain('forceFullCalc="1"')
    } finally {
      workbook.dispose()
    }
  })

  it('does not retain imported XLSX snapshot cell arrays after building the WorkPaper', () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-retention.xlsx')
    attachSourceReaderForTest(imported.snapshot, sourceBytes)
    const sourceSheet = imported.snapshot.sheets[0]
    if (sourceSheet === undefined) {
      throw new Error('Expected imported workbook to contain a sheet')
    }
    const originalCellCount = sourceSheet.cells.length
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      expect(sourceSheet.cells).toHaveLength(0)
      sourceSheet.cells.push({ address: 'Z99', value: 999 })

      const exportedSnapshot = workbook.exportSnapshot()
      expect(readRuntimeImage(exportedSnapshot)).toBeUndefined()
      expect(exportedSnapshot.sheets[0]?.cells).toHaveLength(originalCellCount)
      expect(exportedSnapshot.sheets[0]?.cells.some((cell) => cell.address === 'Z99')).toBe(false)
    } finally {
      workbook.dispose()
    }
  })

  it('preserves formula XML when source-preserving scalar edits patch recalculated caches', () => {
    const sourceBytes = sourceWorkbookWithoutFormulaCachesBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-without-formula-caches.xlsx')
    attachSourceBytesForTest(imported.snapshot, sourceBytes)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      workbook.setCellContents({ sheet: 1, row: 0, col: 0 }, 5)

      const exportedZip = unzipSync(exportXlsx(workbook.exportSnapshot()))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>5</v>')
      expect(cellXml(sheetXml, 'B1')).toContain('<f>A1+1</f>')
      expect(cellXml(sheetXml, 'B1')).toContain('<v>6</v>')
      expect(cellXml(sheetXml, 'B2')).toContain('<f>A2+1</f>')
      expect(cellXml(sheetXml, 'B2')).not.toContain('<v>')

      const workbookXml = strFromU8(exportedZip['xl/workbook.xml'] ?? new Uint8Array())
      expect(workbookXml).toContain('fullCalcOnLoad="1"')
      expect(workbookXml).toContain('forceFullCalc="1"')
    } finally {
      workbook.dispose()
    }
  })

  it('keeps source-preserving scalar patches across explicit rebuilds', () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-rebuild-formula-cache.xlsx')
    attachSourceBytesForTest(imported.snapshot, sourceBytes)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      workbook.setCellContents({ sheet: 1, row: 0, col: 0 }, 5)
      workbook.rebuildAndRecalculate()

      const minimalSnapshot = workbook.exportSourcePreservingXlsxSnapshot()
      expect(minimalSnapshot?.sheets[0]?.cells).toHaveLength(0)

      const exportedZip = unzipSync(exportXlsx(minimalSnapshot!))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>5</v>')
      expect(cellXml(sheetXml, 'B1')).toContain('<f>A1+1</f><v>2</v>')
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
    } finally {
      workbook.dispose()
    }
  })

  it('applies recalculated workbook calc settings to source-preserving patches', () => {
    const sourceBytes = setWorkbookCalcPr(sourceWorkbookBytes(), '<calcPr calcMode="manual" fullCalcOnLoad="0" forceFullCalc="1"/>')
    const imported = importXlsx(sourceBytes, 'source-preserving-calc-settings.xlsx')
    attachSourceBytesForTest(imported.snapshot, sourceBytes)
    attachSourcePatchesForTest(imported.snapshot, [{ sheetName: 'Data', address: 'A1', value: 5 }])
    imported.snapshot.workbook.metadata = {
      ...imported.snapshot.workbook.metadata,
      calculationSettings: { mode: 'manual' },
    }
    Object.defineProperty(imported.snapshot, sourcePreservingOutputCalculationSettings, {
      configurable: true,
      enumerable: false,
      value: { mode: 'manual' },
    })

    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const workbookXml = strFromU8(exportedZip['xl/workbook.xml'] ?? new Uint8Array())
    expect(readCalcPrAttribute(workbookXml, 'calcMode')).toBe('manual')
    expect(readCalcPrAttribute(workbookXml, 'fullCalcOnLoad')).toBeNull()
    expect(readCalcPrAttribute(workbookXml, 'forceFullCalc')).toBeNull()
  })

  it('exports source-preserving scalar edits from a range reader without materializing source bytes', () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-reader.xlsx')
    attachSourceReaderForTest(imported.snapshot, sourceBytes)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      workbook.setCellContents({ sheet: 1, row: 1, col: 0 }, 9)

      const exportedZip = unzipSync(exportXlsx(workbook.exportSnapshot()))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A2')).toContain('<v>9</v>')
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
    } finally {
      workbook.dispose()
    }
  })

  it('keeps byte-source XLSX imports range-patchable without full source reads', async () => {
    const sourceBytes = largeSourceWorkbookBytes()
    const source = instrumentedSourceReaderForTest(sourceBytes)
    const imported = importXlsxFromZipByteSource(source, 'source-preserving-byte-source.xlsx')
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    const fullRangeReadsAfterBuild = source.fullRangeReadCount()
    const directory = mkdtempSync(join(tmpdir(), 'bilig-workpaper-byte-source-patch-'))
    try {
      const outputPath = join(directory, 'patched.xlsx')
      workbook.setCellContents({ sheet: 1, row: 0, col: 0 }, 31)

      const result = await exportWorkPaperXlsxToFileAsync(workbook, outputPath)

      expect(result.bytesWritten).toBeGreaterThan(0)
      expect(source.fullRangeReadCount()).toBe(fullRangeReadsAfterBuild)
      const exportedZip = unzipSync(readFileSync(outputPath))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>31</v>')
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
    } finally {
      workbook.dispose()
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('keeps SheetJS fallback byte-source imports range-patchable without source rereads', async () => {
    const sourceBytes = fallbackSourceWorkbookBytes()
    const source = instrumentedSourceReaderForTest(sourceBytes)
    const imported = importXlsxFromZipByteSource(source, 'source-preserving-byte-source-fallback.xlsx', {
      allowLegacyLargeSheetJsFallback: true,
      limits: false,
    })
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    const fullRangeReadsAfterBuild = source.fullRangeReadCount()
    const directory = mkdtempSync(join(tmpdir(), 'bilig-workpaper-byte-source-fallback-patch-'))
    try {
      const outputPath = join(directory, 'patched.xlsx')
      workbook.setCellContents({ sheet: 1, row: 0, col: 0 }, 37)

      const result = await exportWorkPaperXlsxToFileAsync(workbook, outputPath)

      expect(result.bytesWritten).toBeGreaterThan(0)
      expect(source.fullRangeReadCount()).toBe(fullRangeReadsAfterBuild)
      const exportedZip = unzipSync(readFileSync(outputPath))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>37</v>')
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
      expect(exportedZip['xl/threadedComments/threadedComment1.xml']).toBeDefined()
    } finally {
      workbook.dispose()
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('imports XLSX files through the headless file-backed byte-source path', async () => {
    const sourceBytes = sourceWorkbookBytes()
    const directory = mkdtempSync(join(tmpdir(), 'bilig-workpaper-file-import-'))
    try {
      const inputPath = join(directory, 'source.xlsx')
      const outputPath = join(directory, 'patched.xlsx')
      writeFileSync(inputPath, sourceBytes)
      const imported = importXlsxFile(inputPath)
      const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
      try {
        workbook.setCellContents({ sheet: 1, row: 1, col: 0 }, 41)

        const result = await exportWorkPaperXlsxToFileAsync(workbook, outputPath)

        expect(result.bytesWritten).toBeGreaterThan(0)
        const exportedZip = unzipSync(readFileSync(outputPath))
        const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
        expect(cellXml(sheetXml, 'A2')).toContain('<v>41</v>')
        expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
      } finally {
        workbook.dispose()
      }
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('releases imported XLSX source readers when WorkPaper is disposed', () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-release.xlsx')
    const source = instrumentedSourceReaderForTest(sourceBytes)
    Object.defineProperty(imported.snapshot, Symbol.for('bilig.importedXlsxSourceBytes'), {
      configurable: true,
      enumerable: false,
      value: source,
    })
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)

    workbook.dispose()
    workbook.dispose()

    expect(source.releaseCount()).toBe(1)
  })

  it('exports WorkPaper XLSX scalar source patches from a minimal snapshot', () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-workpaper-helper.xlsx')
    attachSourceReaderForTest(imported.snapshot, sourceBytes)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      workbook.setCellContents({ sheet: 1, row: 0, col: 0 }, 7)
      const minimalSnapshot = workbook.exportSourcePreservingXlsxSnapshot()
      expect(minimalSnapshot?.sheets[0]?.cells).toHaveLength(0)

      const exportedZip = unzipSync(exportXlsx(minimalSnapshot!))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>7</v>')
    } finally {
      workbook.dispose()
    }
  })

  it('writes WorkPaper XLSX scalar source patches to a file without materializing source bytes', async () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-workpaper-file-helper.xlsx')
    attachSourceReaderForTest(imported.snapshot, sourceBytes)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)
    const directory = mkdtempSync(join(tmpdir(), 'bilig-workpaper-source-patch-async-'))
    try {
      const outputPath = join(directory, 'patched.xlsx')
      workbook.setCellContents({ sheet: 1, row: 1, col: 0 }, 19)

      const result = await exportWorkPaperXlsxToFileAsync(workbook, outputPath)

      expect(result.bytesWritten).toBeGreaterThan(0)
      const exportedZip = unzipSync(readFileSync(outputPath))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A2')).toContain('<v>19</v>')
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
    } finally {
      workbook.dispose()
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('writes WorkPaper source patches to a file without asking for the full export snapshot', async () => {
    const sourceBytes = sourceWorkbookBytes()
    const directory = mkdtempSync(join(tmpdir(), 'bilig-workpaper-source-patch-no-snapshot-'))
    try {
      const outputPath = join(directory, 'patched.xlsx')
      const sourcePreservingSnapshot = {
        version: 1,
        workbook: { name: 'Workbook' },
        sheets: [{ name: 'Data', order: 0, cells: [] }],
      }
      attachSourceReaderForTest(sourcePreservingSnapshot, sourceBytes)
      attachSourcePatchesForTest(sourcePreservingSnapshot, [{ sheetName: 'Data', address: 'A1', value: 23 }])

      const result = await exportWorkPaperXlsxToFileAsync(
        {
          exportSourcePreservingXlsxSnapshot: () => sourcePreservingSnapshot,
          exportSnapshot: () => {
            throw new Error('Full WorkPaper snapshot should not be exported for source-preserving XLSX file patches')
          },
        },
        outputPath,
      )

      expect(result.bytesWritten).toBeGreaterThan(0)
      const exportedZip = unzipSync(readFileSync(outputPath))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>23</v>')
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('keeps formula XML when source metadata patches cached formula values', async () => {
    const sourceBytes = sourceWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'source-preserving-formula-cache.xlsx')
    attachSourceReaderForTest(imported.snapshot, sourceBytes)
    attachSourcePatchesForTest(imported.snapshot, [
      { sheetName: 'Data', address: 'A1', value: 8 },
      { sheetName: 'Data', address: 'B1', value: 9, preserveFormula: true },
    ])
    const directory = mkdtempSync(join(tmpdir(), 'bilig-source-formula-cache-'))
    try {
      const outputPath = join(directory, 'patched.xlsx')

      await exportWorkPaperXlsxToFileAsync(
        {
          exportSourcePreservingXlsxSnapshot: () => imported.snapshot,
          exportSnapshot: () => {
            throw new Error('Full WorkPaper snapshot should not be exported for source-preserving formula cache patches')
          },
        },
        outputPath,
      )

      const exportedZip = unzipSync(readFileSync(outputPath))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>8</v>')
      expect(cellXml(sheetXml, 'B1')).toContain('<f>A1+1</f><v>9</v>')
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('exports literal XLSX source patches without importing or building a WorkPaper snapshot', () => {
    const sourceBytes = sourceWorkbookBytes()

    const exportedZip = unzipSync(
      exportXlsxSourceLiteralPatches({
        source: sourceReaderForTest(sourceBytes),
        sheetNames: ['Data'],
        patches: [{ sheetName: 'Data', address: 'A2', value: 11 }],
      }),
    )
    const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    expect(cellXml(sheetXml, 'A2')).toContain('<v>11</v>')
    expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
  })

  it('writes literal XLSX source patches to a file without materializing source bytes', () => {
    const sourceBytes = sourceWorkbookBytes()
    const directory = mkdtempSync(join(tmpdir(), 'bilig-source-patch-'))
    try {
      const outputPath = join(directory, 'patched.xlsx')

      const result = exportXlsxSourceLiteralPatchesToFile({
        source: sourceReaderForTest(sourceBytes),
        outputPath,
        sheetNames: ['Data'],
        patches: [{ sheetName: 'Data', address: 'A1', value: 13 }],
      })

      expect(result.bytesWritten).toBeGreaterThan(0)
      const exportedZip = unzipSync(readFileSync(outputPath))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A1')).toContain('<v>13</v>')
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  it('writes literal XLSX source patches through the async native file path', async () => {
    const sourceBytes = sourceWorkbookBytes()
    const directory = mkdtempSync(join(tmpdir(), 'bilig-source-patch-async-'))
    let source:
      | {
          readonly byteLength: number
          readBytes(): Uint8Array
          readRange?(start: number, end: number): Uint8Array
          readRangeInto?(start: number, end: number, target: Uint8Array): Uint8Array
          release?(): void
        }
      | undefined
    try {
      const sourcePath = join(directory, 'source.xlsx')
      const outputPath = join(directory, 'patched.xlsx')
      writeFileSync(sourcePath, sourceBytes)
      source = createFileImportedXlsxSourceReader(sourcePath)

      const result = await exportXlsxSourceLiteralPatchesToFileAsync({
        source,
        outputPath,
        sheetNames: ['Data'],
        patches: [{ sheetName: 'Data', address: 'A2', value: 17 }],
      })

      expect(result.bytesWritten).toBeGreaterThan(0)
      const exportedZip = unzipSync(readFileSync(outputPath))
      const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(cellXml(sheetXml, 'A2')).toContain('<v>17</v>')
      expect(strFromU8(exportedZip['customXml/item1.xml'] ?? new Uint8Array())).toBe('<keep source="true"/>')
    } finally {
      source?.release?.()
      rmSync(directory, { recursive: true, force: true })
    }
  })
})

function cellXml(sheetXml: string, address: string): string {
  const match = new RegExp(`<c\\b(?=[^>]*\\br="${address}")(?:[^>"']|"[^"]*"|'[^']*')*(?:/>|>[\\s\\S]*?</c>)`, 'u').exec(sheetXml)
  return match?.[0] ?? ''
}

function setWorkbookCalcPr(bytes: Uint8Array, calcPrXml: string): Uint8Array {
  const zip = unzipSync(bytes)
  const workbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  if (/<calcPr\b[\s\S]*?\/>/u.test(workbookXml)) {
    zip['xl/workbook.xml'] = strToU8(workbookXml.replace(/<calcPr\b[\s\S]*?\/>/u, calcPrXml))
    return zipSync(zip)
  }
  zip['xl/workbook.xml'] = strToU8(workbookXml.replace('</workbook>', `${calcPrXml}</workbook>`))
  return zipSync(zip)
}

function readCalcPrAttribute(workbookXml: string, name: string): string | null {
  const calcPr = /<calcPr\b([^>]*)\/?>/u.exec(workbookXml)?.[1] ?? ''
  const match = new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(calcPr)
  return match?.[1] ?? null
}
