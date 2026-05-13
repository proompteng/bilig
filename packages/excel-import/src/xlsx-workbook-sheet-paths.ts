import { XMLParser } from 'fast-xml-parser'
import type * as XLSX from 'xlsx'

import { parseRelationships, resolveTargetPath } from './xlsx-pivot-artifacts.js'
import { normalizeZipPath } from './xlsx-zip.js'

interface WorkbookSheetEntry {
  name: string
  relationshipId: string
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: true,
})

const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function workbookRecord(workbook: XLSX.WorkBook): Record<string, unknown> | null {
  const value: unknown = workbook
  return isRecord(value) ? value : null
}

function workbookFiles(workbook: XLSX.WorkBook): unknown {
  return workbookRecord(workbook)?.['files']
}

function getFileText(files: unknown, path: string): string | null {
  if (!isRecord(files)) {
    return null
  }
  const file = files[normalizeZipPath(path)]
  if (!isRecord(file)) {
    return null
  }
  const content = file['content']
  if (typeof content === 'string') {
    return content
  }
  if (content instanceof ArrayBuffer) {
    return new TextDecoder().decode(content)
  }
  if (ArrayBuffer.isView(content)) {
    return new TextDecoder().decode(content)
  }
  return null
}

function readWorkbookSheetEntries(workbookXml: string | null): WorkbookSheetEntry[] {
  if (!workbookXml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const workbook = recordChild(parsed, 'workbook')
  const sheets = recordChild(workbook, 'sheets')
  return asArray(sheets?.['sheet']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string' || typeof entry['id'] !== 'string') {
      return []
    }
    return [{ name: entry['name'], relationshipId: entry['id'] }]
  })
}

export function workbookDirectorySheetPaths(workbook: XLSX.WorkBook): string[] {
  const directory = workbookRecord(workbook)?.['Directory']
  if (!isRecord(directory)) {
    return []
  }
  return asArray(directory['sheets']).flatMap((entry) => (typeof entry === 'string' ? [entry] : []))
}

export function workbookSheetPathsByName(workbook: XLSX.WorkBook): Map<string, string> {
  const files = workbookFiles(workbook)
  const workbookRelationships = parseRelationships(getFileText(files, workbookRelationshipsPath))
  const worksheetRelationshipsById = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type === worksheetRelationshipType || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))]),
  )
  const output = new Map<string, string>()
  for (const entry of readWorkbookSheetEntries(getFileText(files, workbookPath))) {
    const worksheetPath = worksheetRelationshipsById.get(entry.relationshipId)
    if (worksheetPath) {
      output.set(entry.name, worksheetPath)
    }
  }
  return output
}

export function workbookSheetPath(
  pathsByName: ReadonlyMap<string, string>,
  fallbackPaths: readonly string[],
  sheetName: string,
  sheetIndex: number,
): string | undefined {
  return pathsByName.get(sheetName) ?? fallbackPaths[sheetIndex]
}
