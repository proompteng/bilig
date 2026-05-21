import {
  readLargeSimpleReferencedSharedStringsFromChunks,
  readLargeSimpleSharedStrings,
  type LargeSimpleSharedStrings,
  type LargeSimpleReferencedSharedStringScanOptions,
} from './xlsx-large-simple-shared-strings.js'
import type { LargeSimpleSharedStringIndexSet } from './xlsx-large-simple-shared-string-indexes.js'
import { forEachInflatedXlsxZipEntryChunk, getZipText, type XlsxZipEntries } from './xlsx-zip.js'

const sharedStringsPath = 'xl/sharedStrings.xml'

export function readReferencedLargeSimpleSharedStrings(
  zip: XlsxZipEntries,
  referencedIndexes: LargeSimpleSharedStringIndexSet,
  options: LargeSimpleReferencedSharedStringScanOptions = {},
): LargeSimpleSharedStrings | null {
  const streamed = readLargeSimpleReferencedSharedStringsFromChunks(
    (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, sharedStringsPath, onChunk),
    referencedIndexes,
    options,
  )
  if (streamed) {
    return streamed
  }
  return readAllLargeSimpleSharedStrings(zip, options)
}

export function readAllLargeSimpleSharedStrings(
  zip: XlsxZipEntries,
  options: LargeSimpleReferencedSharedStringScanOptions = {},
): LargeSimpleSharedStrings | null {
  const sharedStringsXml = getZipText(zip, sharedStringsPath)
  return sharedStringsXml ? readLargeSimpleSharedStrings(sharedStringsXml, options) : null
}
