import {
  readLargeSimpleReferencedSharedStringsFromChunks,
  readLargeSimpleSharedStrings,
  type LargeSimpleSharedStringEntry,
} from './xlsx-large-simple-shared-strings.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from './xlsx-large-simple-worksheet-stream-scanner.js'
import { forEachInflatedXlsxZipEntryChunk, getZipText, type XlsxZipEntries } from './xlsx-zip.js'

const sharedStringsPath = 'xl/sharedStrings.xml'

export function readMaterializedLargeSimpleSharedStrings(
  zip: XlsxZipEntries,
  worksheetEntries: readonly { readonly path: string }[],
): LargeSimpleSharedStringEntry[] | null {
  const referencedIndexes = collectLargeSimpleSharedStringIndexes(zip, worksheetEntries)
  if (referencedIndexes) {
    const streamed = readLargeSimpleReferencedSharedStringsFromChunks(
      (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, sharedStringsPath, onChunk),
      referencedIndexes,
    )
    if (streamed) {
      return streamed
    }
  }
  const sharedStringsXml = getZipText(zip, sharedStringsPath)
  return sharedStringsXml ? readLargeSimpleSharedStrings(sharedStringsXml) : null
}

function collectLargeSimpleSharedStringIndexes(
  zip: XlsxZipEntries,
  worksheetEntries: readonly { readonly path: string }[],
): Set<number> | null {
  const referencedIndexes = new Set<number>()
  for (const [sheetIndex, entry] of worksheetEntries.entries()) {
    const streamed = parseLargeSimpleWorksheetCellsFromChunks(
      (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
      sheetIndex,
      {
        hasSharedStrings: true,
        retainCells: false,
        collectSharedStringIndexes: true,
        allowInlineStringsWithoutRetention: true,
      },
    )
    if (!streamed) {
      return null
    }
    for (const index of streamed.sharedStringIndexes) {
      referencedIndexes.add(index)
    }
  }
  return referencedIndexes
}
