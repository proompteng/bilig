export function readXmlAttribute(tag: string, name: string): string | null {
  const doubleQuoted = new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(tag)
  if (doubleQuoted) {
    return doubleQuoted[1] ?? null
  }
  const singleQuoted = new RegExp(`\\b${name}='([^']*)'`, 'u').exec(tag)
  return singleQuoted?.[1] ?? null
}

export function readXmlNumberAttribute(tag: string, name: string): number | null {
  const raw = readXmlAttribute(tag, name)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const value = Number(raw)
  return Number.isFinite(value) ? value : null
}

export function readXmlPositiveIntegerAttribute(tag: string, name: string): number | null {
  const value = readXmlNumberAttribute(tag, name)
  return value !== null && Number.isSafeInteger(value) && value > 0 ? value : null
}

export function readXmlNonNegativeIntegerAttribute(tag: string, name: string): number | null {
  const value = readXmlNumberAttribute(tag, name)
  return value !== null && Number.isSafeInteger(value) && value >= 0 ? value : null
}

export function readXmlOptionalBooleanAttribute(tag: string, name: string): boolean | null {
  const raw = readXmlAttribute(tag, name)
  if (raw === null) {
    return null
  }
  return raw === '1' || raw.toLowerCase() === 'true'
}

export const worksheetCellElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?c)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
export const worksheetCellOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u

export function readCellXfs(stylesXml: string): readonly string[] {
  const match = /<((?:[A-Za-z_][\w.-]*:)?cellXfs)\b[^>]*>([\s\S]*?)<\/\1>/u.exec(stylesXml)
  if (!match) {
    return []
  }
  const body = match[2] ?? ''
  const entries: string[] = []
  let cursor = 0
  const nextXf = /<(?:[A-Za-z_][\w.-]*:)?xf\b/gu
  while (cursor < body.length) {
    nextXf.lastIndex = cursor
    const startMatch = nextXf.exec(body)
    if (!startMatch) {
      break
    }
    const start = startMatch.index
    const openingEnd = body.indexOf('>', start)
    if (openingEnd < 0) {
      break
    }
    const tagName = /^<([^\s/>]+)/u.exec(body.slice(start, openingEnd + 1))?.[1]
    if (!tagName) {
      break
    }
    if (body[openingEnd - 1] === '/') {
      entries.push(body.slice(start, openingEnd + 1))
      cursor = openingEnd + 1
      continue
    }
    const closingTag = `</${tagName}>`
    const closingStart = body.indexOf(closingTag, openingEnd + 1)
    if (closingStart < 0) {
      break
    }
    entries.push(body.slice(start, closingStart + closingTag.length))
    cursor = closingStart + closingTag.length
  }
  return entries
}

export function updateXmlElementCount(openingAttributes: string, count: number): string {
  return /\scount="[^"]*"/u.test(openingAttributes)
    ? openingAttributes.replace(/\scount="[^"]*"/u, ` count="${String(count)}"`)
    : `${openingAttributes} count="${String(count)}"`
}

export function appendCustomCellXfsToStylesXml(stylesXml: string, xfs: readonly string[]): string {
  if (xfs.length === 0) {
    return stylesXml
  }
  return stylesXml.replace(
    /<((?:[A-Za-z_][\w.-]*:)?cellXfs)\b([^>]*)>([\s\S]*?)<\/\1>/u,
    (_match, tagName: string, attributes: string, body: string) => {
      const count = Array.from(body.matchAll(/<(?:[A-Za-z_][\w.-]*:)?xf\b/gu)).length + xfs.length
      return `<${tagName}${updateXmlElementCount(attributes, count)}>${body}${xfs.join('')}</${tagName}>`
    },
  )
}
