import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type { WorkbookCalculationSettingsSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { escapeXmlAttribute } from './xlsx-export-xml.js'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const workbookCalcPrTailElements = [
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
] as const

export const precisionAsDisplayedCalculationWarning = 'Precision-as-displayed calculation is not supported during XLSX import.'
export const manualCalculationModeWarning = 'Manual calculation mode is preserved during XLSX import; cached formula values may be stale.'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function stringValue(value: unknown): string | null {
  return typeof value === 'string' && value.trim().length > 0 ? value : null
}

function booleanAttributeValue(value: unknown): boolean | undefined {
  const raw = stringValue(value)
  if (raw === null) {
    return undefined
  }
  return raw === '1' || raw.toLowerCase() === 'true'
}

function finiteIntegerAttributeValue(value: unknown): number | undefined {
  const raw = stringValue(value)
  if (raw === null) {
    return undefined
  }
  const number = Number(raw)
  return Number.isSafeInteger(number) ? number : undefined
}

function finiteNumericStringAttributeValue(value: unknown): string | undefined {
  const raw = stringValue(value)
  if (raw === null) {
    return undefined
  }
  return Number.isFinite(Number(raw)) ? raw : undefined
}

function formatBooleanAttribute(value: boolean): string {
  return value ? '1' : '0'
}

function calcPrAttribute(name: string, value: string | undefined): string {
  return value === undefined ? '' : ` ${name}="${escapeXmlAttribute(value)}"`
}

function hasSemanticCalculationSettings(settings: WorkbookCalculationSettingsSnapshot): boolean {
  return (
    settings.mode === 'manual' ||
    settings.iterate !== undefined ||
    settings.iterateCount !== undefined ||
    settings.iterateDelta !== undefined ||
    settings.fullCalcOnLoad !== undefined ||
    settings.concurrentCalc !== undefined
  )
}

function buildWorkbookCalcPr(settings: WorkbookCalculationSettingsSnapshot): string | null {
  if (!hasSemanticCalculationSettings(settings)) {
    return null
  }
  const attributes = [
    calcPrAttribute('calcMode', settings.mode === 'manual' ? 'manual' : undefined),
    calcPrAttribute('iterate', typeof settings.iterate === 'boolean' ? formatBooleanAttribute(settings.iterate) : undefined),
    calcPrAttribute('iterateCount', Number.isSafeInteger(settings.iterateCount) ? String(settings.iterateCount) : undefined),
    calcPrAttribute('iterateDelta', typeof settings.iterateDelta === 'string' ? settings.iterateDelta : undefined),
    calcPrAttribute(
      'fullCalcOnLoad',
      typeof settings.fullCalcOnLoad === 'boolean' ? formatBooleanAttribute(settings.fullCalcOnLoad) : undefined,
    ),
    calcPrAttribute(
      'concurrentCalc',
      typeof settings.concurrentCalc === 'boolean' ? formatBooleanAttribute(settings.concurrentCalc) : undefined,
    ),
  ].join('')
  return attributes.length > 0 ? `<calcPr${attributes}/>` : null
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

function insertWorkbookCalcPr(workbookXml: string, calcPrXml: string): string {
  if (/<calcPr\b/u.test(workbookXml)) {
    return workbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/u, calcPrXml)
  }

  let insertIndex = workbookXml.indexOf('</workbook>')
  for (const elementName of workbookCalcPrTailElements) {
    const elementIndex = workbookXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return workbookXml
  }
  return `${workbookXml.slice(0, insertIndex)}${calcPrXml}${workbookXml.slice(insertIndex)}`
}

export function addExportCalculationSettingsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const calcPrXml = snapshot.workbook.metadata?.calculationSettings
    ? buildWorkbookCalcPr(snapshot.workbook.metadata.calculationSettings)
    : null
  if (!calcPrXml) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return bytes
  }
  setZipText(zip, 'xl/workbook.xml', insertWorkbookCalcPr(workbookXml, calcPrXml))
  return zipSync(zip)
}

export function readImportedWorkbookCalculationSettings(source: XlsxZipSource): WorkbookCalculationSettingsSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return undefined
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const calcPr = recordChild(recordChild(parsed, 'workbook'), 'calcPr')
  if (!calcPr) {
    return undefined
  }
  const mode = calcPr['calcMode'] === 'manual' ? 'manual' : 'automatic'
  const iterate = booleanAttributeValue(calcPr['iterate'])
  const iterateCount = finiteIntegerAttributeValue(calcPr['iterateCount'])
  const iterateDelta = finiteNumericStringAttributeValue(calcPr['iterateDelta'])
  const fullCalcOnLoad = booleanAttributeValue(calcPr['fullCalcOnLoad'])
  const concurrentCalc = booleanAttributeValue(calcPr['concurrentCalc'])
  const settings: WorkbookCalculationSettingsSnapshot = {
    mode,
    compatibilityMode: 'excel-modern',
    ...(iterate !== undefined ? { iterate } : {}),
    ...(iterateCount !== undefined ? { iterateCount } : {}),
    ...(iterateDelta !== undefined ? { iterateDelta } : {}),
    ...(fullCalcOnLoad !== undefined ? { fullCalcOnLoad } : {}),
    ...(concurrentCalc !== undefined ? { concurrentCalc } : {}),
  }
  return hasSemanticCalculationSettings(settings) ? settings : undefined
}

export function readImportedWorkbookCalculationWarnings(source: XlsxZipSource): string[] {
  const zip = readXlsxZipEntries(source)
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const calcPr = recordChild(recordChild(parsed, 'workbook'), 'calcPr')
  return [
    ...(calcPr?.['calcMode'] === 'manual' ? [manualCalculationModeWarning] : []),
    ...(calcPr?.['fullPrecision'] === '0' ? [precisionAsDisplayedCalculationWarning] : []),
  ]
}
