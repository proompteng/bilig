import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from './xlsx-large-simple-worksheet-stream-scanner.js'
import {
  forEachInflatedXlsxZipEntryChunk,
  readLazyXlsxZipSourceByteLength,
  readXlsxZipEntryUncompressedSize,
  type XlsxZipEntries,
} from './xlsx-zip.js'

interface LargeSimpleStyleCoordinateWorksheetEntry {
  readonly path: string
}

interface LargeSimpleStyleCoordinateScannedWorksheet {
  readonly cellScan: ImportedWorksheetCellScan
}

const maxDimensionCellPreallocation = 16_000_000
const minXmlBytesPerPreallocatedCell = 16
const deferredStyleCoordinateWorksheetXmlThreshold = 16_000_000

export function maxPreallocatedWorksheetCells(zip: XlsxZipEntries, path: string): number {
  const uncompressedSize = readXlsxZipEntryUncompressedSize(zip, path)
  return uncompressedSize === undefined
    ? 0
    : Math.min(maxDimensionCellPreallocation, Math.floor(uncompressedSize / minXmlBytesPerPreallocatedCell))
}

export function shouldDeferLargeSimpleStyleCoordinates(
  zip: XlsxZipEntries,
  path: string,
  options: {
    readonly materializeCells: boolean
    readonly hasStyles: boolean
  },
): boolean {
  if (!options.materializeCells || !options.hasStyles || readLazyXlsxZipSourceByteLength(zip) === undefined) {
    return false
  }
  return (readXlsxZipEntryUncompressedSize(zip, path) ?? 0) >= deferredStyleCoordinateWorksheetXmlThreshold
}

export function releaseLargeSimpleStyleIndexes(scannedWorksheets: Iterable<LargeSimpleStyleCoordinateScannedWorksheet | undefined>): void {
  for (const scanned of scannedWorksheets) {
    scanned?.cellScan.styleIndexes.release()
  }
}

export function prepareLargeSimpleStyleIndexes<T extends LargeSimpleStyleCoordinateScannedWorksheet>(
  zip: XlsxZipEntries,
  worksheetEntries: readonly LargeSimpleStyleCoordinateWorksheetEntry[],
  scannedWorksheets: (T | undefined)[],
  stylesByIndex: ReadonlyMap<unknown, unknown>,
  options: {
    readonly hasSharedStrings: boolean
    readonly allowUnsupportedFormulaText?: boolean
    readonly allowUnsupportedCellMetadata?: boolean
  },
): boolean {
  if (stylesByIndex.size === 0) {
    releaseLargeSimpleStyleIndexes(scannedWorksheets)
    return true
  }
  const rescannedStyleIndexes = rescanDeferredLargeSimpleStyleIndexes(zip, worksheetEntries, scannedWorksheets, options)
  if (!rescannedStyleIndexes) {
    return false
  }
  for (const [index, styleIndexes] of rescannedStyleIndexes) {
    const scanned = scannedWorksheets[index]
    if (scanned) {
      scannedWorksheets[index] = { ...scanned, cellScan: { ...scanned.cellScan, styleIndexes } } as T
    }
  }
  return true
}

export function prepareLargeSimpleStyleIndexForWorksheet(
  zip: XlsxZipEntries,
  worksheetEntries: readonly LargeSimpleStyleCoordinateWorksheetEntry[],
  scanned: LargeSimpleStyleCoordinateScannedWorksheet,
  options: {
    readonly hasSharedStrings: boolean
    readonly allowUnsupportedFormulaText?: boolean
    readonly allowUnsupportedCellMetadata?: boolean
  },
): ImportedWorksheetCellScan['styleIndexes'] | null {
  if (scanned.cellScan.styleIndexes.hasCoordinateStorage) {
    return scanned.cellScan.styleIndexes
  }
  const entry = worksheetEntries[scanned.cellScan.sheetIndex]
  if (!entry) {
    return null
  }
  const streamed = parseLargeSimpleWorksheetCellsFromChunks(
    (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
    scanned.cellScan.sheetIndex,
    {
      hasSharedStrings: options.hasSharedStrings,
      retainCells: false,
      retainMetadataXml: false,
      retainStyleIndexes: true,
      retainStyleCoordinates: true,
      maxDimensionCellPreallocation: maxPreallocatedWorksheetCells(zip, entry.path),
      ...(options.allowUnsupportedFormulaText === undefined ? {} : { allowUnsupportedFormulaText: options.allowUnsupportedFormulaText }),
      ...(options.allowUnsupportedCellMetadata === undefined ? {} : { allowUnsupportedCellMetadata: options.allowUnsupportedCellMetadata }),
    },
  )
  if (!streamed) {
    return null
  }
  scanned.cellScan.styleIndexes.release()
  return streamed.cellScan.styleIndexes
}

export function rescanDeferredLargeSimpleStyleIndexes(
  zip: XlsxZipEntries,
  worksheetEntries: readonly LargeSimpleStyleCoordinateWorksheetEntry[],
  scannedWorksheets: readonly (LargeSimpleStyleCoordinateScannedWorksheet | undefined)[],
  options: {
    readonly hasSharedStrings: boolean
    readonly allowUnsupportedFormulaText?: boolean
    readonly allowUnsupportedCellMetadata?: boolean
  },
): Map<number, ImportedWorksheetCellScan['styleIndexes']> | null {
  const rescannedStyleIndexes = new Map<number, ImportedWorksheetCellScan['styleIndexes']>()
  for (const [index, scanned] of scannedWorksheets.entries()) {
    if (!scanned || scanned.cellScan.styleIndexes.hasCoordinateStorage) {
      continue
    }
    const entry = worksheetEntries[scanned.cellScan.sheetIndex]
    if (!entry) {
      return null
    }
    const styleIndexes = prepareLargeSimpleStyleIndexForWorksheet(zip, worksheetEntries, scanned, options)
    if (!styleIndexes) {
      return null
    }
    rescannedStyleIndexes.set(index, styleIndexes)
  }
  return rescannedStyleIndexes
}
