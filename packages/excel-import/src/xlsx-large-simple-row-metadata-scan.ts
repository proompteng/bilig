const doubleQuote = 34
const singleQuote = 39

const rowMetadataAttributeNames = [
  'ht',
  'hidden',
  'customHeight',
  's',
  'customFormat',
  'outlineLevel',
  'collapsed',
  'thickTop',
  'thickBottom',
] as const

export function rowTagHasMetadataAttribute(bytes: Uint8Array, nameEnd: number, tagEnd: number): boolean {
  let index = nameEnd
  while (index < tagEnd) {
    while (index < tagEnd && isAsciiWhitespace(bytes[index] ?? 0)) {
      index += 1
    }
    const attributeNameStart = index
    while (index < tagEnd && isXmlNameByte(bytes[index] ?? 0)) {
      index += 1
    }
    const attributeNameEnd = index
    if (rowMetadataAttributeNameMatches(bytes, attributeNameStart, attributeNameEnd)) {
      return true
    }
    index = skipAsciiWhitespace(bytes, index, tagEnd)
    if (bytes[index] !== 61) {
      index += 1
      continue
    }
    index = skipAsciiWhitespace(bytes, index + 1, tagEnd)
    const quote = bytes[index]
    if (quote !== doubleQuote && quote !== singleQuote) {
      index += 1
      continue
    }
    index += 1
    while (index < tagEnd && bytes[index] !== quote) {
      index += 1
    }
    index += 1
  }
  return false
}

function rowMetadataAttributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number): boolean {
  for (const attributeName of rowMetadataAttributeNames) {
    if (attributeNameMatches(bytes, startIndex, endIndex, attributeName)) {
      return true
    }
  }
  return false
}

function attributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, attributeName: string): boolean {
  if (endIndex - startIndex !== attributeName.length) {
    return false
  }
  for (let index = 0; index < attributeName.length; index += 1) {
    if (bytes[startIndex + index] !== attributeName.charCodeAt(index)) {
      return false
    }
  }
  return true
}

function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

function isAsciiWhitespace(byte: number): boolean {
  return byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32
}

function isXmlNameByte(byte: number): boolean {
  return (
    (byte >= 65 && byte <= 90) ||
    (byte >= 97 && byte <= 122) ||
    (byte >= 48 && byte <= 57) ||
    byte === 45 ||
    byte === 46 ||
    byte === 58 ||
    byte === 95
  )
}
