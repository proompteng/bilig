import type { CellStyleFontSnapshot, CellStyleRecord } from '@bilig/protocol'

type ImportedCellStyle = Omit<CellStyleRecord, 'id'>

const elementTextCache = new Map<string, RegExp>()

export function readLargeSimpleWorkbookStyles(
  stylesXml: string | null,
  requiredStyleIndexes: ReadonlySet<number>,
): Map<number, ImportedCellStyle> | null {
  if (!stylesXml || requiredStyleIndexes.size === 0) {
    return new Map()
  }
  const fills = readFillStyles(stylesXml)
  const fonts = readFontStyles(stylesXml)
  const cellXfsXml = extractElementXml(stylesXml, 'cellXfs')
  if (!fills || !fonts || !cellXfsXml) {
    return null
  }
  const cellXfs = [
    ...cellXfsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?xf>)/gu),
  ]
  const styles = new Map<number, ImportedCellStyle>()
  for (const styleIndex of requiredStyleIndexes) {
    const xfXml = cellXfs[styleIndex]?.[0]
    if (!xfXml) {
      return null
    }
    const openingTag = /<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xfXml)?.[0]
    if (!openingTag) {
      return null
    }
    const numFmtId = readNonNegativeIntegerAttribute(openingTag, 'numFmtId')
    if (numFmtId !== null && numFmtId !== 0) {
      return null
    }
    const fillId = readNonNegativeIntegerAttribute(openingTag, 'fillId')
    const fontId = readNonNegativeIntegerAttribute(openingTag, 'fontId')
    const borderId = readNonNegativeIntegerAttribute(openingTag, 'borderId')
    if (isApplied(openingTag, 'applyBorder', borderId) || hasChildElement(xfXml, 'alignment') || hasChildElement(xfXml, 'protection')) {
      return null
    }
    const fill = isApplied(openingTag, 'applyFill', fillId) ? fills[fillId ?? -1] : undefined
    const font = isApplied(openingTag, 'applyFont', fontId) ? fonts[fontId ?? -1] : undefined
    if (fill === null || font === null) {
      return null
    }
    const style: ImportedCellStyle = {
      ...(fill ? { fill } : {}),
      ...(font ? { font } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(styleIndex, style)
    }
  }
  return styles
}

function readFillStyles(stylesXml: string): Array<ImportedCellStyle['fill'] | null> | null {
  const fillsXml = extractElementXml(stylesXml, 'fills')
  if (!fillsXml) {
    return []
  }
  return [...fillsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?fill\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?fill>/gu)].map((match) =>
    readFillStyle(match[0] ?? ''),
  )
}

function readFillStyle(fillXml: string): ImportedCellStyle['fill'] | null | undefined {
  const patternFill = extractElementXml(fillXml, 'patternFill')
  if (!patternFill) {
    return undefined
  }
  const openingTag = /<(?:[A-Za-z_][\w.-]*:)?patternFill\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(patternFill)?.[0]
  const patternType = openingTag ? readAttribute(openingTag, 'patternType') : undefined
  if (!patternType || patternType === 'none' || patternType === 'gray125') {
    return undefined
  }
  if (patternType !== 'solid') {
    return null
  }
  const color = readColor(patternFill, 'fgColor') ?? readColor(patternFill, 'bgColor')
  return color ? { backgroundColor: color } : undefined
}

function readFontStyles(stylesXml: string): Array<CellStyleFontSnapshot | null | undefined> | null {
  const fontsXml = extractElementXml(stylesXml, 'fonts')
  if (!fontsXml) {
    return []
  }
  return [...fontsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?font\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?font>/gu)].map((match) =>
    readFontStyle(match[0] ?? ''),
  )
}

function readFontStyle(fontXml: string): CellStyleFontSnapshot | null | undefined {
  const family = readElementValue(fontXml, 'name')
  const size = readElementNumberValue(fontXml, 'sz')
  const color = readColor(fontXml, 'color')
  const font: CellStyleFontSnapshot = {
    ...(family ? { family } : {}),
    ...(size ? { size } : {}),
    ...(hasBooleanElement(fontXml, 'b') ? { bold: true } : {}),
    ...(hasBooleanElement(fontXml, 'i') ? { italic: true } : {}),
    ...(hasBooleanElement(fontXml, 'u') ? { underline: true } : {}),
    ...(color ? { color } : {}),
  }
  return Object.keys(font).length > 0 ? font : undefined
}

function extractElementXml(xml: string, elementName: string): string | null {
  let pattern = elementTextCache.get(elementName)
  if (!pattern) {
    const qualifiedName = `(?:[A-Za-z_][\\w.-]*:)?${elementName}`
    pattern = new RegExp(`<${qualifiedName}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${qualifiedName}>)`, 'u')
    elementTextCache.set(elementName, pattern)
  }
  return pattern.exec(xml)?.[0] ?? null
}

function hasChildElement(xml: string, elementName: string): boolean {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b`, 'u').test(xml.replace(/^<[^>]*>/u, ''))
}

function readElementValue(xml: string, elementName: string): string | undefined {
  const elementXml = extractElementXml(xml, elementName)
  if (!elementXml) {
    return undefined
  }
  return readAttribute(elementXml, 'val')
}

function readElementNumberValue(xml: string, elementName: string): number | undefined {
  const value = readElementValue(xml, elementName)
  if (value === undefined) {
    return undefined
  }
  const number = Number(value)
  return Number.isFinite(number) && number > 0 ? number : undefined
}

function hasBooleanElement(xml: string, elementName: string): boolean {
  const elementXml = extractElementXml(xml, elementName)
  if (!elementXml) {
    return false
  }
  const value = readAttribute(elementXml, 'val')
  return value === undefined || value === '1' || value.toLocaleLowerCase('en-US') === 'true'
}

function readColor(xml: string, elementName: string): string | undefined {
  const elementXml = extractElementXml(xml, elementName)
  const rgb = elementXml ? readAttribute(elementXml, 'rgb') : undefined
  if (!rgb) {
    return undefined
  }
  const normalized = rgb.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{8}$/u.test(normalized)) {
    return `#${normalized.slice(2).toLocaleLowerCase('en-US')}`
  }
  if (/^[0-9a-fA-F]{6}$/u.test(normalized)) {
    return `#${normalized.toLocaleLowerCase('en-US')}`
  }
  return undefined
}

function isApplied(tag: string, attributeName: string, componentId: number | null): boolean {
  const value = readAttribute(tag, attributeName)
  if (value === '1' || value?.toLocaleLowerCase('en-US') === 'true') {
    return true
  }
  if (value === '0' || value?.toLocaleLowerCase('en-US') === 'false') {
    return false
  }
  return componentId !== null && componentId > 0
}

function readNonNegativeIntegerAttribute(tag: string, attributeName: string): number | null {
  const value = readAttribute(tag, attributeName)
  if (!value || !/^[0-9]+$/u.test(value)) {
    return null
  }
  const number = Number(value)
  return Number.isSafeInteger(number) ? number : null
}

function readAttribute(tag: string, attributeName: string): string | undefined {
  const pattern = new RegExp(`\\s${escapeRegExp(attributeName)}=(?:"([^"]*)"|'([^']*)')`, 'u')
  const match = pattern.exec(tag)
  return match?.[1] ?? match?.[2]
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
