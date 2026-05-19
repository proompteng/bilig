import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type {
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
  WorkbookConditionalFormatRuleSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookSheetConditionalFormatArtifactsSnapshot,
  WorkbookSnapshot,
  WorkbookValidationComparisonOperator,
} from '@bilig/protocol'
import { applyExportWorksheetDimensionsToWorksheetXml } from './xlsx-dimensions.js'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

const worksheetConditionalFormattingTailElements = [
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'pivotTableDefinition',
] as const

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

function stringValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  return null
}

function numberValue(value: unknown): number | null {
  const raw = stringValue(value)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const number = Number(raw)
  return Number.isFinite(number) ? number : null
}

function booleanValue(value: unknown): boolean | undefined {
  if (value === true || value === '1' || value === 'true') {
    return true
  }
  if (value === false || value === '0' || value === 'false') {
    return false
  }
  return undefined
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

function hasWorksheetConditionalFormatting(sheetXml: string): boolean {
  return /<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b/u.test(sheetXml)
}

function worksheetConditionalFormattingRegex(): RegExp {
  const qualifiedName = '(?:[A-Za-z_][\\w.-]*:)?conditionalFormatting'
  return new RegExp(`<${qualifiedName}\\b[^>]*>[\\s\\S]*?<\\/${qualifiedName}>|<${qualifiedName}\\b[^>]*\\/>`, 'gu')
}

function extractWorksheetConditionalFormattingXml(sheetXml: string): string[] {
  return Array.from(sheetXml.matchAll(worksheetConditionalFormattingRegex()), (match) => match[0])
}

function removeWorksheetConditionalFormattingXml(sheetXml: string): string {
  return sheetXml.replace(worksheetConditionalFormattingRegex(), '')
}

function normalizeRgbColor(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const normalized = value.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return `#${normalized.toLowerCase()}`
  }
  if (/^[0-9a-fA-F]{8}$/.test(normalized)) {
    return `#${normalized.slice(2).toLowerCase()}`
  }
  return null
}

function toArgbColor(value: string): string | null {
  const normalized = normalizeRgbColor(value)
  return normalized ? `FF${normalized.slice(1).toUpperCase()}` : null
}

function rangeRefA1(range: CellRangeRef): string | null {
  try {
    const decoded = XLSX.utils.decode_range(`${range.startAddress}:${range.endAddress}`.replaceAll('$', ''))
    return XLSX.utils.encode_range(decoded)
  } catch {
    return null
  }
}

function parseSqrefRange(sheetName: string, ref: string): CellRangeRef | null {
  try {
    const decoded = XLSX.utils.decode_range(ref.replaceAll('$', ''))
    return {
      sheetName,
      startAddress: XLSX.utils.encode_cell(decoded.s),
      endAddress: XLSX.utils.encode_cell(decoded.e),
    }
  } catch {
    return null
  }
}

function formatLiteralFormula(value: LiteralInput): string | null {
  if (value === null) {
    return null
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? String(value) : null
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return `"${value.replaceAll('"', '""')}"`
}

function parseLiteralFormula(value: unknown): LiteralInput | undefined {
  const raw = stringValue(value)
  if (raw === null) {
    return undefined
  }
  const trimmed = raw.trim()
  if (trimmed.length === 0) {
    return ''
  }
  if (trimmed === 'TRUE') {
    return true
  }
  if (trimmed === 'FALSE') {
    return false
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replaceAll('""', '"')
  }
  const number = Number(trimmed)
  return Number.isFinite(number) ? number : trimmed
}

function parseComparisonOperator(value: unknown): WorkbookValidationComparisonOperator | null {
  switch (value) {
    case 'between':
    case 'notBetween':
    case 'equal':
    case 'notEqual':
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual':
      return value
    default:
      return null
  }
}

function buildDxfXml(style: CellStylePatch): string | null {
  const fontParts: string[] = []
  if (style.font?.bold === true) {
    fontParts.push('<b/>')
  }
  if (style.font?.italic === true) {
    fontParts.push('<i/>')
  }
  if (style.font?.underline === true) {
    fontParts.push('<u/>')
  }
  const fontColor = style.font?.color ? toArgbColor(style.font.color) : null
  if (fontColor) {
    fontParts.push(`<color rgb="${fontColor}"/>`)
  }

  const fillColor = style.fill?.backgroundColor ? toArgbColor(style.fill.backgroundColor) : null
  const parts = [
    ...(fontParts.length > 0 ? [`<font>${fontParts.join('')}</font>`] : []),
    ...(fillColor
      ? [`<fill><patternFill patternType="solid"><fgColor rgb="${fillColor}"/><bgColor indexed="64"/></patternFill></fill>`]
      : []),
  ]
  return parts.length > 0 ? `<dxf>${parts.join('')}</dxf>` : null
}

function readDxfStyle(dxf: unknown): CellStylePatch {
  if (!isRecord(dxf)) {
    return {}
  }

  const font = recordChild(dxf, 'font')
  const fontColor = normalizeRgbColor(recordChild(font, 'color')?.['rgb'])
  const fill = recordChild(dxf, 'fill')
  const patternFill = recordChild(fill, 'patternFill')
  const fillColor = normalizeRgbColor(recordChild(patternFill, 'fgColor')?.['rgb'] ?? recordChild(patternFill, 'bgColor')?.['rgb'])
  const style: CellStylePatch = {
    ...(fillColor ? { fill: { backgroundColor: fillColor } } : {}),
    ...(font && (font['b'] !== undefined || font['i'] !== undefined || font['u'] !== undefined || fontColor)
      ? {
          font: {
            ...(font['b'] !== undefined ? { bold: true } : {}),
            ...(font['i'] !== undefined ? { italic: true } : {}),
            ...(font['u'] !== undefined ? { underline: true } : {}),
            ...(fontColor ? { color: fontColor } : {}),
          },
        }
      : {}),
  }
  return style
}

function readDxfs(stylesXml: string | null): CellStylePatch[] {
  if (!stylesXml) {
    return []
  }
  const dxfsXml = extractStyleXmlElement(stylesXml, 'dxfs')
  if (!dxfsXml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(`<styleSheet>${dxfsXml}</styleSheet>`)
  return asArray(recordChild(recordChild(parsed, 'styleSheet'), 'dxfs')?.['dxf']).map(readDxfStyle)
}

function extractStyleXmlElement(stylesXml: string, elementName: string): string | null {
  const qualifiedName = `(?:[A-Za-z_][\\w.-]*:)?${elementName}`
  const expanded = new RegExp(`<${qualifiedName}\\b[^>]*>[\\s\\S]*?<\\/${qualifiedName}>`, 'u').exec(stylesXml)
  if (expanded) {
    return expanded[0]
  }
  const selfClosing = new RegExp(`<${qualifiedName}\\b[^>]*\\/>`, 'u').exec(stylesXml)
  return selfClosing?.[0] ?? null
}

function countExistingDxfs(stylesXml: string): number {
  const countMatch = /<dxfs\b[^>]*\bcount="(\d+)"/u.exec(stylesXml)
  if (countMatch) {
    const count = Number(countMatch[1])
    if (Number.isInteger(count) && count >= 0) {
      return count
    }
  }
  return (stylesXml.match(/<dxf\b/g) ?? []).length
}

function appendDxfs(stylesXml: string, dxfXmls: readonly string[]): string {
  if (dxfXmls.length === 0) {
    return stylesXml
  }
  const existingCount = countExistingDxfs(stylesXml)
  const nextCount = existingCount + dxfXmls.length
  const appended = dxfXmls.join('')
  if (/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/u.test(stylesXml)) {
    return stylesXml.replace(
      /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/u,
      (_match, existingInner: string) => `<dxfs count="${String(nextCount)}">${existingInner}${appended}</dxfs>`,
    )
  }
  if (/<dxfs\b[^>]*\/>/u.test(stylesXml)) {
    return stylesXml.replace(/<dxfs\b[^>]*\/>/u, `<dxfs count="${String(nextCount)}">${appended}</dxfs>`)
  }
  const insertAt = stylesXml.search(/<tableStyles\b|<colors\b|<extLst\b|<\/styleSheet>/u)
  if (insertAt < 0) {
    return stylesXml
  }
  return `${stylesXml.slice(0, insertAt)}<dxfs count="${String(nextCount)}">${appended}</dxfs>${stylesXml.slice(insertAt)}`
}

function insertWorksheetConditionalFormatting(sheetXml: string, conditionalFormattingXml: readonly string[]): string {
  if (conditionalFormattingXml.length === 0) {
    return sheetXml
  }
  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetConditionalFormattingTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  const insert = conditionalFormattingXml.join('')
  return `${sheetXml.slice(0, insertIndex)}${insert}${sheetXml.slice(insertIndex)}`
}

function buildFormulaRuleXml(rule: WorkbookConditionalFormatRuleSnapshot): string[] | null {
  switch (rule.kind) {
    case 'cellIs': {
      const formulas = rule.values.map(formatLiteralFormula)
      return formulas.every((formula): formula is string => formula !== null)
        ? formulas.map((formula) => `<formula>${escapeXml(formula)}</formula>`)
        : null
    }
    case 'formula':
      return [`<formula>${escapeXml(rule.formula.trim().replace(/^=/u, ''))}</formula>`]
    case 'textContains': {
      const text = rule.text.replaceAll('"', '""')
      return [`<formula>NOT(ISERROR(SEARCH(&quot;${escapeXml(text)}&quot;,A1)))</formula>`]
    }
    case 'blanks':
    case 'notBlanks':
      return []
  }
}

function conditionalFormatType(rule: WorkbookConditionalFormatRuleSnapshot): string {
  switch (rule.kind) {
    case 'cellIs':
      return 'cellIs'
    case 'formula':
      return 'expression'
    case 'textContains':
      return 'containsText'
    case 'blanks':
      return 'containsBlanks'
    case 'notBlanks':
      return 'notContainsBlanks'
  }
}

function buildConditionalFormattingXml(
  format: WorkbookConditionalFormatSnapshot,
  dxfId: number | undefined,
  fallbackPriority: number,
): string | null {
  const sqref = rangeRefA1(format.range)
  if (!sqref) {
    return null
  }
  const formulas = buildFormulaRuleXml(format.rule)
  if (!formulas) {
    return null
  }
  const attributes = [
    `type="${conditionalFormatType(format.rule)}"`,
    ...(dxfId !== undefined ? [`dxfId="${String(dxfId)}"`] : []),
    `priority="${String(format.priority ?? fallbackPriority)}"`,
    ...(format.rule.kind === 'cellIs' ? [`operator="${format.rule.operator}"`] : []),
    ...(format.rule.kind === 'textContains' ? [`operator="containsText" text="${escapeXml(format.rule.text)}"`] : []),
    ...(format.stopIfTrue === true ? ['stopIfTrue="1"'] : []),
  ]
  return `<conditionalFormatting sqref="${escapeXml(sqref)}"><cfRule ${attributes.join(' ')}>${formulas.join('')}</cfRule></conditionalFormatting>`
}

function parseConditionalFormatRule(rule: Record<string, unknown>): WorkbookConditionalFormatRuleSnapshot | null {
  switch (rule['type']) {
    case 'cellIs': {
      const operator = parseComparisonOperator(rule['operator'])
      const values = asArray(rule['formula']).flatMap((formula) => {
        const value = parseLiteralFormula(formula)
        return value === undefined ? [] : [value]
      })
      return operator && values.length > 0 ? { kind: 'cellIs', operator, values } : null
    }
    case 'expression': {
      const formula = stringValue(asArray(rule['formula'])[0])
      return formula !== null && formula.trim().length > 0
        ? { kind: 'formula', formula: formula.startsWith('=') ? formula : `=${formula}` }
        : null
    }
    case 'containsText': {
      const text = stringValue(rule['text'])
      return text !== null ? { kind: 'textContains', text } : null
    }
    case 'containsBlanks':
      return { kind: 'blanks' }
    case 'notContainsBlanks':
      return { kind: 'notBlanks' }
    default:
      return null
  }
}

export function addExportConditionalFormatsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const sheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  if (
    !sheets.some(
      (sheet) =>
        (sheet.metadata?.conditionalFormats ?? []).length > 0 || (sheet.metadata?.conditionalFormatArtifacts?.xml.trim().length ?? 0) > 0,
    )
  ) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const stylesXml = getZipText(zip, 'xl/styles.xml')
  if (!stylesXml) {
    return bytes
  }

  const existingDxfCount = countExistingDxfs(stylesXml)
  const dxfIdsByStyle = new Map<string, number>()
  const dxfXmls: string[] = []
  let priority = 1
  let changed = false
  sheets.forEach((sheet, sheetIndex) => {
    const conditionalFormats = sheet.metadata?.conditionalFormats ?? []
    const importedConditionalFormatXml = sheet.metadata?.conditionalFormatArtifacts?.xml.trim()
    const conditionalFormattingXml: string[] = importedConditionalFormatXml ? [importedConditionalFormatXml] : []
    if (!importedConditionalFormatXml) {
      for (const format of conditionalFormats) {
        const dxfXml = buildDxfXml(format.style)
        let dxfId: number | undefined
        if (dxfXml) {
          const cacheKey = JSON.stringify(format.style)
          const cached = dxfIdsByStyle.get(cacheKey)
          if (cached !== undefined) {
            dxfId = cached
          } else {
            dxfId = existingDxfCount + dxfXmls.length
            dxfIdsByStyle.set(cacheKey, dxfId)
            dxfXmls.push(dxfXml)
          }
        }
        const xml = buildConditionalFormattingXml(format, dxfId, priority)
        priority += 1
        if (xml) {
          conditionalFormattingXml.push(xml)
        }
      }
    }
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml) {
      return
    }
    const updatedSheetXml = applyExportWorksheetDimensionsToWorksheetXml(
      insertWorksheetConditionalFormatting(removeWorksheetConditionalFormattingXml(sheetXml), conditionalFormattingXml),
      sheet.metadata,
    )
    if (updatedSheetXml === sheetXml) {
      return
    }
    setZipText(zip, sheetPath, updatedSheetXml)
    changed = true
  })

  if (!changed) {
    return bytes
  }
  setZipText(zip, 'xl/styles.xml', appendDxfs(stylesXml, dxfXmls))
  return zipSync(zip)
}

export function readImportedWorkbookConditionalFormats(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookConditionalFormatSnapshot[]> {
  return readImportedWorkbookConditionalFormatsFromWorksheetPaths(
    source,
    sheetNames.map((sheetName, sheetIndex) => ({
      name: sheetName,
      path: `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`,
    })),
  )
}

export function readImportedWorkbookConditionalFormatsFromWorksheetPaths(
  source: XlsxZipSource,
  worksheets: readonly { readonly name: string; readonly path: string }[],
): Map<string, WorkbookConditionalFormatSnapshot[]> {
  const zip = readXlsxZipEntries(source)
  const conditionalFormatsBySheet = new Map<string, WorkbookConditionalFormatSnapshot[]>()

  worksheets.forEach(({ name: sheetName, path }) => {
    const sheetXml = getZipText(zip, path)
    if (!sheetXml) {
      return
    }
    const conditionalFormats = readImportedSheetConditionalFormatsFromWorksheetXml(zip, sheetName, sheetXml)
    if (conditionalFormats) {
      conditionalFormatsBySheet.set(sheetName, conditionalFormats)
    }
  })

  return conditionalFormatsBySheet
}

export function readImportedSheetConditionalFormatsFromWorksheetXml(
  source: XlsxZipSource,
  sheetName: string,
  sheetXml: string,
): WorkbookConditionalFormatSnapshot[] | undefined {
  if (!hasWorksheetConditionalFormatting(sheetXml)) {
    return undefined
  }
  return readImportedSheetConditionalFormatsFromElementXml(source, sheetName, extractWorksheetConditionalFormattingXml(sheetXml))
}

export function readImportedSheetConditionalFormatsFromElementXml(
  source: XlsxZipSource,
  sheetName: string,
  conditionalFormattingXml: readonly string[],
): WorkbookConditionalFormatSnapshot[] | undefined {
  if (conditionalFormattingXml.length === 0) {
    return undefined
  }
  const dxfs = readDxfs(getZipText(readXlsxZipEntries(source), 'xl/styles.xml'))
  const parsed: unknown = xmlParser.parse(`<worksheet>${conditionalFormattingXml.join('')}</worksheet>`)
  const conditionalFormats: WorkbookConditionalFormatSnapshot[] = []
  for (const conditionalFormatting of asArray(recordChild(parsed, 'worksheet')?.['conditionalFormatting'])) {
    if (!isRecord(conditionalFormatting) || typeof conditionalFormatting['sqref'] !== 'string') {
      continue
    }
    const ranges = conditionalFormatting['sqref'].split(/\s+/u).flatMap((ref) => {
      const range = parseSqrefRange(sheetName, ref)
      return range ? [range] : []
    })
    if (ranges.length === 0) {
      continue
    }
    for (const ruleEntry of asArray(conditionalFormatting['cfRule'])) {
      if (!isRecord(ruleEntry)) {
        continue
      }
      const rule = parseConditionalFormatRule(ruleEntry)
      if (!rule) {
        continue
      }
      const dxfId = numberValue(ruleEntry['dxfId'])
      const style = dxfId !== null ? (dxfs[dxfId] ?? {}) : {}
      const stopIfTrue = booleanValue(ruleEntry['stopIfTrue'])
      const priority = numberValue(ruleEntry['priority']) ?? undefined
      for (const range of ranges) {
        conditionalFormats.push({
          id: `xlsx-cf:${sheetName}:${range.startAddress}:${range.endAddress}:${String(conditionalFormats.length + 1)}`,
          range,
          rule,
          style,
          ...(stopIfTrue !== undefined ? { stopIfTrue } : {}),
          ...(priority !== undefined ? { priority } : {}),
        })
      }
    }
  }
  return conditionalFormats.length > 0 ? conditionalFormats : undefined
}

export function readImportedWorkbookConditionalFormatArtifacts(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetConditionalFormatArtifactsSnapshot> {
  return readImportedWorkbookConditionalFormatArtifactsFromWorksheetPaths(
    source,
    sheetNames.map((sheetName, sheetIndex) => ({
      name: sheetName,
      path: `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`,
    })),
  )
}

export function readImportedWorkbookConditionalFormatArtifactsFromWorksheetPaths(
  source: XlsxZipSource,
  worksheets: readonly { readonly name: string; readonly path: string }[],
): Map<string, WorkbookSheetConditionalFormatArtifactsSnapshot> {
  const zip = readXlsxZipEntries(source)
  const artifactsBySheet = new Map<string, WorkbookSheetConditionalFormatArtifactsSnapshot>()

  worksheets.forEach(({ name: sheetName, path }) => {
    const sheetXml = getZipText(zip, path)
    if (!sheetXml) {
      return
    }
    const artifacts = readImportedSheetConditionalFormatArtifactsFromWorksheetXml(sheetXml)
    if (artifacts) {
      artifactsBySheet.set(sheetName, artifacts)
    }
  })

  return artifactsBySheet
}

export function readImportedSheetConditionalFormatArtifactsFromWorksheetXml(
  sheetXml: string,
): WorkbookSheetConditionalFormatArtifactsSnapshot | undefined {
  return readImportedSheetConditionalFormatArtifactsFromElementXml(extractWorksheetConditionalFormattingXml(sheetXml))
}

export function readImportedSheetConditionalFormatArtifactsFromElementXml(
  conditionalFormattingXml: readonly string[],
): WorkbookSheetConditionalFormatArtifactsSnapshot | undefined {
  const xml = conditionalFormattingXml.filter(needsConditionalFormatArtifactXml).join('')
  return xml.length > 0 ? { xml } : undefined
}

export function needsConditionalFormatArtifactXml(conditionalFormattingXml: string): boolean {
  let parsed: unknown
  try {
    parsed = xmlParser.parse(`<worksheet>${conditionalFormattingXml}</worksheet>`)
  } catch {
    return true
  }
  const conditionalFormatting = recordChild(parsed, 'worksheet')?.['conditionalFormatting']
  if (!isRecord(conditionalFormatting) || typeof conditionalFormatting['sqref'] !== 'string') {
    return true
  }
  const supportedConditionalFormattingKeys = new Set(['sqref', 'cfRule'])
  if (Object.keys(conditionalFormatting).some((key) => !supportedConditionalFormattingKeys.has(key))) {
    return true
  }
  const rules = asArray(conditionalFormatting['cfRule'])
  return rules.length === 0 || rules.some((rule) => !isFaithfullyTypedConditionalFormatRule(rule))
}

function isFaithfullyTypedConditionalFormatRule(rule: unknown): boolean {
  if (!isRecord(rule) || rule['dxfId'] !== undefined || parseConditionalFormatRule(rule) === null) {
    return false
  }
  const type = rule['type']
  switch (type) {
    case 'cellIs':
      return hasOnlyConditionalFormatRuleKeys(rule, ['type', 'priority', 'stopIfTrue', 'operator', 'formula'])
    case 'expression':
      return hasOnlyConditionalFormatRuleKeys(rule, ['type', 'priority', 'stopIfTrue', 'formula'])
    case 'containsBlanks':
    case 'notContainsBlanks':
      return hasOnlyConditionalFormatRuleKeys(rule, ['type', 'priority', 'stopIfTrue'])
    default:
      return false
  }
}

function hasOnlyConditionalFormatRuleKeys(rule: Record<string, unknown>, keys: readonly string[]): boolean {
  const supported = new Set(keys)
  return Object.keys(rule).every((key) => supported.has(key))
}
