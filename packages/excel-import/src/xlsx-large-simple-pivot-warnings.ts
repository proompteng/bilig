import { getZipText, type XlsxZipEntries } from './xlsx-zip.js'

export function hasExternalLargeSimplePivotCaches(zip: XlsxZipEntries): boolean {
  return Object.keys(zip).some((path) => {
    if (!/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/u.test(path)) {
      return false
    }
    const xml = getZipText(zip, path)
    return typeof xml === 'string' && /<(?:[A-Za-z_][\w.-]*:)?cacheSource\b[^>]*\btype=(["'])external\1/u.test(xml)
  })
}
