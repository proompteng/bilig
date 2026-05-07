import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type { WorkbookFreezePaneActivePane, WorkbookFreezePaneSnapshot, WorkbookSnapshot } from '@bilig/protocol'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

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

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getZipText(zip: ZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

function setZipText(zip: ZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function toPositiveInteger(value: unknown): number {
  const number = typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : Number.NaN
  return Number.isInteger(number) && number > 0 ? number : 0
}

function normalizeCellReference(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const normalized = value.trim().replaceAll('$', '').toUpperCase()
  if (!/^[A-Z]{1,3}[1-9][0-9]*$/u.test(normalized)) {
    return undefined
  }
  try {
    const decoded = XLSX.utils.decode_cell(normalized)
    return decoded.r >= 0 && decoded.c >= 0 ? XLSX.utils.encode_cell(decoded) : undefined
  } catch {
    return undefined
  }
}

function normalizeActivePane(value: unknown): WorkbookFreezePaneActivePane | undefined {
  switch (value) {
    case 'bottomRight':
    case 'bottomLeft':
    case 'topRight':
    case 'topLeft':
      return value
    default:
      return undefined
  }
}

function normalizeFreezePane(freezePane: WorkbookFreezePaneSnapshot | undefined): WorkbookFreezePaneSnapshot | null {
  if (!freezePane) {
    return null
  }
  const rows = toPositiveInteger(freezePane.rows)
  const cols = toPositiveInteger(freezePane.cols)
  if (rows <= 0 && cols <= 0) {
    return null
  }
  const normalized: WorkbookFreezePaneSnapshot = { rows, cols }
  const topLeftCell = normalizeCellReference(freezePane.topLeftCell)
  if (topLeftCell !== undefined) {
    normalized.topLeftCell = topLeftCell
  }
  const activePane = normalizeActivePane(freezePane.activePane)
  if (activePane !== undefined) {
    normalized.activePane = activePane
  }
  return normalized
}

function activePaneForFreeze(freezePane: WorkbookFreezePaneSnapshot): WorkbookFreezePaneActivePane {
  if (freezePane.rows > 0 && freezePane.cols > 0) {
    return 'bottomRight'
  }
  return freezePane.rows > 0 ? 'bottomLeft' : 'topRight'
}

function buildFreezePaneXml(freezePane: WorkbookFreezePaneSnapshot): string {
  const attributes: string[] = []
  if (freezePane.cols > 0) {
    attributes.push(`xSplit="${String(freezePane.cols)}"`)
  }
  if (freezePane.rows > 0) {
    attributes.push(`ySplit="${String(freezePane.rows)}"`)
  }
  const topLeftCell = freezePane.topLeftCell ?? XLSX.utils.encode_cell({ r: freezePane.rows, c: freezePane.cols })
  const activePane = freezePane.activePane ?? activePaneForFreeze(freezePane)
  attributes.push(`topLeftCell="${escapeXml(topLeftCell)}"`)
  attributes.push(`activePane="${activePane}"`)
  attributes.push('state="frozen"')
  return `<pane ${attributes.join(' ')}/><selection pane="${activePane}"/>`
}

function removeExistingPaneMarkup(sheetViewInnerXml: string): string {
  return sheetViewInnerXml
    .replace(/<pane\b[^>]*(?:\/>|>[\s\S]*?<\/pane>)/u, '')
    .replace(/<selection\b[^>]*(?:\/>|>[\s\S]*?<\/selection>)/gu, '')
}

function addFreezePaneToSheetView(sheetViewXml: string, freezePane: WorkbookFreezePaneSnapshot): string {
  const selfClosingMatch = /^<sheetView\b([^>]*)\/>$/u.exec(sheetViewXml)
  if (selfClosingMatch) {
    return `<sheetView${selfClosingMatch[1] ?? ''}>${buildFreezePaneXml(freezePane)}</sheetView>`
  }

  const expandedMatch = /^<sheetView\b([^>]*)>([\s\S]*?)<\/sheetView>$/u.exec(sheetViewXml)
  if (!expandedMatch) {
    return sheetViewXml
  }
  return `<sheetView${expandedMatch[1] ?? ''}>${buildFreezePaneXml(freezePane)}${removeExistingPaneMarkup(expandedMatch[2] ?? '')}</sheetView>`
}

function insertFreezePaneIntoWorksheet(sheetXml: string, freezePane: WorkbookFreezePaneSnapshot): string {
  const sheetViewMatch = /<sheetView\b[^>]*(?:\/>|>[\s\S]*?<\/sheetView>)/u.exec(sheetXml)
  if (sheetViewMatch) {
    return sheetXml.replace(sheetViewMatch[0], addFreezePaneToSheetView(sheetViewMatch[0], freezePane))
  }

  const sheetViews = `<sheetViews><sheetView workbookViewId="0">${buildFreezePaneXml(freezePane)}</sheetView></sheetViews>`
  const dimensionMatch = /<dimension\b[^>]*\/>/u.exec(sheetXml)
  if (dimensionMatch) {
    return sheetXml.replace(dimensionMatch[0], `${dimensionMatch[0]}${sheetViews}`)
  }
  return sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1>${sheetViews}`)
}

export function addExportFreezePanesToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => normalizeFreezePane(sheet.metadata?.freezePane) !== null)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const freezePane = normalizeFreezePane(sheet.metadata?.freezePane)
      if (!freezePane) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertFreezePaneIntoWorksheet(sheetXml, freezePane))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

function parseFreezePane(pane: Record<string, unknown>): WorkbookFreezePaneSnapshot | null {
  const state = pane['state']
  if (state !== undefined && state !== 'frozen' && state !== 'frozenSplit') {
    return null
  }
  const rows = toPositiveInteger(pane['ySplit'])
  const cols = toPositiveInteger(pane['xSplit'])
  if (rows <= 0 && cols <= 0) {
    return null
  }
  const freezePane: WorkbookFreezePaneSnapshot = { rows, cols }
  const topLeftCell = normalizeCellReference(pane['topLeftCell'])
  if (topLeftCell !== undefined) {
    freezePane.topLeftCell = topLeftCell
  }
  const activePane = normalizeActivePane(pane['activePane'])
  if (activePane !== undefined) {
    freezePane.activePane = activePane
  }
  return freezePane
}

export function readImportedWorkbookFreezePanes(bytes: Uint8Array, sheetNames: readonly string[]): Map<string, WorkbookFreezePaneSnapshot> {
  const zip = unzipSync(bytes)
  const freezePanesBySheet = new Map<string, WorkbookFreezePaneSnapshot>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !/<pane\b/u.test(sheetXml)) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const sheetViews = recordChild(recordChild(parsed, 'worksheet'), 'sheetViews')
    const sheetView = asArray(sheetViews?.['sheetView']).find(isRecord)
    const pane = sheetView ? recordChild(sheetView, 'pane') : null
    if (!pane) {
      return
    }
    const freezePane = parseFreezePane(pane)
    if (freezePane) {
      freezePanesBySheet.set(sheetName, freezePane)
    }
  })

  return freezePanesBySheet
}
