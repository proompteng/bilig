import { unzipSync, zipSync } from 'fflate'

import { addMissingFormattedCells } from './xlsx-cell-insertion.js'
import {
  addCustomNumberFormatsToStylesXml,
  customNumberFormatStartId,
  getZipText,
  repairLeadingZeroNumberFormatIds,
  setXmlAttribute,
  setZipText,
} from './xlsx-export-xml.js'
import {
  appendCustomCellXfsToStylesXml,
  readCellXfs,
  readXmlNonNegativeIntegerAttribute,
  worksheetCellElementPattern,
  worksheetCellOpeningTagPattern,
} from './xlsx-style-xml.js'

function styleXfWithNumberFormat(xf: string, numberFormatId: number): string {
  return xf.replace(/<(?:[A-Za-z_][\w.-]*:)?xf\b[^>]*(?:\/>|>)/u, (openingTag) =>
    setXmlAttribute(
      setXmlAttribute(setXmlAttribute(openingTag, 'numFmtId', String(numberFormatId)), 'applyNumberFormat', '1'),
      'xfId',
      '0',
    ),
  )
}

class ExportNumberFormatRegistry {
  private readonly baseXfs: readonly string[]
  private readonly numberFormatIdsByCode = new Map<string, number>()
  private readonly styleIndexesByKey = new Map<string, number>()
  private readonly addedFormatIdsByCode = new Map<string, number>()
  private readonly addedXfs: string[] = []
  private nextNumberFormatId: number

  constructor(stylesXml: string) {
    this.baseXfs = readCellXfs(stylesXml)
    const usedIds = [...stylesXml.matchAll(/\bnumFmtId="([0-9]+)"/gu)].map((match) => Number(match[1])).filter(Number.isSafeInteger)
    this.nextNumberFormatId = Math.max(customNumberFormatStartId, ...usedIds.map((id) => id + 1))
  }

  styleIndexFor(baseStyleIndex: number, formatCode: string): number {
    const key = `${String(baseStyleIndex)}\u0000${formatCode}`
    const existingStyleIndex = this.styleIndexesByKey.get(key)
    if (existingStyleIndex !== undefined) {
      return existingStyleIndex
    }
    const numberFormatId = this.numberFormatIdFor(formatCode)
    const baseXf = this.baseXfs[baseStyleIndex] ?? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    const styleIndex = this.baseXfs.length + this.addedXfs.length
    this.addedXfs.push(styleXfWithNumberFormat(baseXf, numberFormatId))
    this.styleIndexesByKey.set(key, styleIndex)
    return styleIndex
  }

  apply(stylesXml: string): string {
    return appendCustomCellXfsToStylesXml(addCustomNumberFormatsToStylesXml(stylesXml, this.addedFormatIdsByCode), this.addedXfs)
  }

  private numberFormatIdFor(formatCode: string): number {
    const existingId = this.numberFormatIdsByCode.get(formatCode)
    if (existingId !== undefined) {
      return existingId
    }
    const id = this.nextNumberFormatId
    this.nextNumberFormatId += 1
    this.numberFormatIdsByCode.set(formatCode, id)
    this.addedFormatIdsByCode.set(formatCode, id)
    return id
  }
}

function applyNumberFormatsToSheetXml(
  sheetXml: string,
  formats: ReadonlyMap<string, string>,
  registry: ExportNumberFormatRegistry,
): string {
  if (formats.size === 0) {
    return sheetXml
  }
  const handledAddresses = new Set<string>()
  let output = sheetXml.replace(worksheetCellElementPattern, (cellXml) => {
    const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
    const address = openingTag ? /\br="([^"]+)"/u.exec(openingTag)?.[1] : undefined
    const format = address ? formats.get(address) : undefined
    if (!openingTag || !address || !format) {
      return cellXml
    }
    handledAddresses.add(address)
    const baseStyleIndex = readXmlNonNegativeIntegerAttribute(openingTag, 's') ?? 0
    const styleIndex = registry.styleIndexFor(baseStyleIndex, format)
    return cellXml.replace(openingTag, setXmlAttribute(openingTag, 's', String(styleIndex)))
  })
  const missingCells = [...formats.entries()]
    .filter(([address]) => !handledAddresses.has(address))
    .map(([address, format]) => ({
      address,
      styleIndex: registry.styleIndexFor(0, format),
    }))
  if (missingCells.length > 0) {
    output = addMissingFormattedCells(output, missingCells)
  }
  return output
}

export function preserveSnapshotNumberFormats(bytes: Uint8Array, sheetFormats: readonly ReadonlyMap<string, string>[]): Uint8Array {
  if (sheetFormats.every((formats) => formats.size === 0)) {
    return repairLeadingZeroNumberFormatIds(bytes)
  }
  const zip = unzipSync(repairLeadingZeroNumberFormatIds(bytes))
  const stylesXml = getZipText(zip, 'xl/styles.xml')
  if (!stylesXml) {
    return zipSync(zip)
  }
  const registry = new ExportNumberFormatRegistry(stylesXml)
  sheetFormats.forEach((formats, sheetIndex) => {
    if (formats.size === 0) {
      return
    }
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml) {
      return
    }
    setZipText(zip, sheetPath, applyNumberFormatsToSheetXml(sheetXml, formats, registry))
  })
  setZipText(zip, 'xl/styles.xml', registry.apply(stylesXml))
  return zipSync(zip)
}
