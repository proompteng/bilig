import { strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookIgnoredErrorsSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

const ignoredErrorsElementPattern = /<(?:[A-Za-z_][\w.-]*:)?ignoredErrors\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?ignoredErrors>)/gu

const worksheetIgnoredErrorsTailElements = [
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
] as const

const deterministicZipOptions = { mtime: new Date(1980, 0, 1, 0, 0, 0) } as const

function setZipText(zip: Record<string, Uint8Array>, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function readIgnoredErrorsXml(sheetXml: string): string | undefined {
  ignoredErrorsElementPattern.lastIndex = 0
  const ignoredErrorsXml = ignoredErrorsElementPattern.exec(sheetXml)?.[0]
  return ignoredErrorsXml ? addMissingNamespaceDeclarations(sheetXml, ignoredErrorsXml) : undefined
}

function removeIgnoredErrorsXml(sheetXml: string): string {
  ignoredErrorsElementPattern.lastIndex = 0
  return sheetXml.replace(ignoredErrorsElementPattern, '')
}

function usedNamespacePrefixes(xml: string): Set<string> {
  const prefixes = new Set<string>()
  for (const match of xml.matchAll(/\b([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\b/gu)) {
    const prefix = match[1]
    if (prefix && prefix !== 'xml' && prefix !== 'xmlns') {
      prefixes.add(prefix)
    }
  }
  return prefixes
}

function worksheetNamespaceDeclaration(sheetXml: string, prefix: string): string | null {
  const worksheetOpening = /<worksheet\b([^>]*)>/u.exec(sheetXml)?.[1] ?? ''
  const declaration = new RegExp(`\\sxmlns:${prefix}=(["'])([\\s\\S]*?)\\1`, 'u').exec(worksheetOpening)
  return declaration ? `xmlns:${prefix}=${declaration[1]}${declaration[2] ?? ''}${declaration[1]}` : null
}

function addMissingNamespaceDeclarations(sheetXml: string, ignoredErrorsXml: string): string {
  const missingDeclarations = [...usedNamespacePrefixes(ignoredErrorsXml)].flatMap((prefix) => {
    if (new RegExp(`\\sxmlns:${prefix}=`, 'u').test(ignoredErrorsXml)) {
      return []
    }
    const declaration = worksheetNamespaceDeclaration(sheetXml, prefix)
    return declaration ? [declaration] : []
  })
  if (missingDeclarations.length === 0) {
    return ignoredErrorsXml
  }
  return ignoredErrorsXml.replace(
    /<((?:[A-Za-z_][\w.-]*:)?ignoredErrors)\b([^>]*?)(\/?)>/u,
    (_match, tagName: string, attributes: string, selfClosing: string) =>
      `<${tagName}${attributes} ${missingDeclarations.join(' ')}${selfClosing}>`,
  )
}

function insertIgnoredErrorsXml(sheetXml: string, ignoredErrorsXml: string): string {
  const withoutIgnoredErrors = removeIgnoredErrorsXml(sheetXml)
  let insertIndex = withoutIgnoredErrors.indexOf('</worksheet>')
  for (const elementName of worksheetIgnoredErrorsTailElements) {
    const elementIndex = withoutIgnoredErrors.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return withoutIgnoredErrors
  }
  return `${withoutIgnoredErrors.slice(0, insertIndex)}${ignoredErrorsXml}${withoutIgnoredErrors.slice(insertIndex)}`
}

export function readImportedWorkbookIgnoredErrors(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookIgnoredErrorsSnapshot> {
  const zip = readXlsxZipEntries(source)
  const ignoredErrorsBySheet = new Map<string, WorkbookIgnoredErrorsSnapshot>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !/<(?:[A-Za-z_][\w.-]*:)?ignoredErrors\b/u.test(sheetXml)) {
      return
    }
    const ignoredErrorsXml = readIgnoredErrorsXml(sheetXml)
    if (ignoredErrorsXml) {
      ignoredErrorsBySheet.set(sheetName, { xml: ignoredErrorsXml })
    }
  })

  return ignoredErrorsBySheet
}

export function addExportIgnoredErrorsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const zip = unzipSync(bytes)
  let changed = false

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const ignoredErrorsXml = sheet.metadata?.ignoredErrors?.xml
      const nextSheetXml = ignoredErrorsXml ? insertIgnoredErrorsXml(sheetXml, ignoredErrorsXml) : removeIgnoredErrorsXml(sheetXml)
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, sheetPath, nextSheetXml)
        changed = true
      }
    })

  return changed ? zipSync(zip, deterministicZipOptions) : bytes
}
