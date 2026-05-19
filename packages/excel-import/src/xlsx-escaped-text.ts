function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

export function decodeExcelEscapedText(value: string): string {
  const escapedUnderscore = '\uE000'
  return value
    .replace(/_x005F_/giu, escapedUnderscore)
    .replace(/_x([0-9a-fA-F]{4})_/gu, (_match, code: string) => {
      const codePoint = Number.parseInt(code, 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    })
    .replaceAll(escapedUnderscore, '_')
}

function escapeExcelEscapeMarkers(value: string): string {
  let output = ''
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]!
    if (character === '_' && /^_x[0-9a-fA-F]{4}_/u.test(value.slice(index, index + 7))) {
      output += '_x005F_'
    } else {
      output += character
    }
  }
  return output
}

export function encodeExcelEscapedText(value: string): string {
  let output = ''
  for (const character of escapeExcelEscapeMarkers(value)) {
    const codePoint = character.codePointAt(0)!
    output += codePoint <= 0x1f ? `_x${codePoint.toString(16).padStart(4, '0')}_` : character
  }
  return output
}
