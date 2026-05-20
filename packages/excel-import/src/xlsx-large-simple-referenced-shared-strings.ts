import {
  readLargeSimpleReferencedSharedStringsFromChunks,
  readLargeSimpleSharedStrings,
  type LargeSimpleSharedStrings,
} from './xlsx-large-simple-shared-strings.js'
import { forEachInflatedXlsxZipEntryChunk, getZipText, type XlsxZipEntries } from './xlsx-zip.js'

const sharedStringsPath = 'xl/sharedStrings.xml'

export function readReferencedLargeSimpleSharedStrings(
  zip: XlsxZipEntries,
  referencedIndexes: ReadonlySet<number>,
): LargeSimpleSharedStrings | null {
  const streamed = readLargeSimpleReferencedSharedStringsFromChunks(
    (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, sharedStringsPath, onChunk),
    referencedIndexes,
  )
  if (streamed) {
    return streamed
  }
  return readAllLargeSimpleSharedStrings(zip)
}

export function readAllLargeSimpleSharedStrings(zip: XlsxZipEntries): LargeSimpleSharedStrings | null {
  const sharedStringsXml = getZipText(zip, sharedStringsPath)
  return sharedStringsXml ? readLargeSimpleSharedStrings(sharedStringsXml) : null
}
