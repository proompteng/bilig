import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

export const customNumberFormatStartId = 164

export function escapeXmlAttribute(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function addCustomNumberFormatsToStylesXml(stylesXml: string, formatIdsByCode: ReadonlyMap<string, number>): string {
  if (formatIdsByCode.size === 0) {
    return stylesXml
  }
  const numFmtEntries = [...formatIdsByCode.entries()]
    .map(([formatCode, id]) => `<numFmt numFmtId="${String(id)}" formatCode="${escapeXmlAttribute(formatCode)}"/>`)
    .join('')
  const selfClosingNumFmts = /<numFmts\b[^>]*\/>/u
  if (selfClosingNumFmts.test(stylesXml)) {
    return stylesXml.replace(selfClosingNumFmts, () => `<numFmts count="${String(formatIdsByCode.size)}">${numFmtEntries}</numFmts>`)
  }
  const existingNumFmts = /<numFmts count="([0-9]+)">/u.exec(stylesXml)
  if (existingNumFmts) {
    const count = Number(existingNumFmts[1])
    const nextCount = Number.isFinite(count) ? count + formatIdsByCode.size : formatIdsByCode.size
    return stylesXml
      .replace(/<numFmts count="[0-9]+">/u, () => `<numFmts count="${String(nextCount)}">`)
      .replace('</numFmts>', () => `${numFmtEntries}</numFmts>`)
  }
  const numFmtsXml = `<numFmts count="${String(formatIdsByCode.size)}">${numFmtEntries}</numFmts>`
  return stylesXml.replace(/<fonts\b/u, (match) => `${numFmtsXml}${match}`)
}

export function repairLeadingZeroNumberFormatIds(bytes: Uint8Array): Uint8Array {
  const zip = unzipSync(bytes)
  const styles = zip['xl/styles.xml']
  if (!styles) {
    return bytes
  }
  let stylesXml = strFromU8(styles)
  const leadingZeroFormatCodes = [...new Set([...stylesXml.matchAll(/\bnumFmtId="(0[0-9]+)"/gu)].map((match) => match[1]!))]
  if (leadingZeroFormatCodes.length === 0) {
    return bytes
  }
  const usedIds = new Set([...stylesXml.matchAll(/\bnumFmtId="([0-9]+)"/gu)].map((match) => Number(match[1])))
  const formatIdsByCode = new Map<string, number>()
  let nextId = customNumberFormatStartId
  for (const formatCode of leadingZeroFormatCodes) {
    while (usedIds.has(nextId)) {
      nextId += 1
    }
    formatIdsByCode.set(formatCode, nextId)
    usedIds.add(nextId)
  }
  for (const [formatCode, id] of formatIdsByCode.entries()) {
    stylesXml = stylesXml.replaceAll(`numFmtId="${formatCode}"`, `numFmtId="${String(id)}"`)
  }
  const customIds = [...formatIdsByCode.values()].map(String).join('|')
  const xfWithCustomNumberFormatPattern = new RegExp(`<xf\\b([^>]*)\\bnumFmtId="(${customIds})"([^>]*)/>`, 'gu')
  stylesXml = stylesXml.replace(xfWithCustomNumberFormatPattern, (tag: string, before: string, id: string, after: string) =>
    tag.includes('applyNumberFormat=') ? tag : `<xf${before} numFmtId="${id}"${after} applyNumberFormat="1"/>`,
  )
  stylesXml = addCustomNumberFormatsToStylesXml(stylesXml, formatIdsByCode)
  zip['xl/styles.xml'] = strToU8(stylesXml)
  return zipSync(zip)
}

export function getZipText(zip: Record<string, Uint8Array>, path: string): string | null {
  const file = zip[path]
  return file ? strFromU8(file) : null
}

export function setZipText(zip: Record<string, Uint8Array>, path: string, text: string): void {
  zip[path] = strToU8(text)
}

export function setXmlAttribute(tag: string, name: string, value: string): string {
  const attribute = `${name}="${escapeXmlAttribute(value)}"`
  const existingAttribute = new RegExp(`\\s${name}="[^"]*"`, 'u')
  if (existingAttribute.test(tag)) {
    return tag.replace(existingAttribute, ` ${attribute}`)
  }
  return tag.replace(/\/?>$/u, (ending) => ` ${attribute}${ending}`)
}

export function readXmlNumberAttribute(tag: string, name: string): number | null {
  const match = new RegExp(`\\s${name}="([0-9]+)"`, 'u').exec(tag)
  if (!match) {
    return null
  }
  const value = Number(match[1])
  return Number.isSafeInteger(value) && value >= 0 ? value : null
}
