export function shouldUseSharedStringlessFastPathBytes(bytes: Uint8Array): boolean {
  let valueCellCount = 0
  for (let index = 0; index < bytes.byteLength; index += 1) {
    if (bytes[index] !== 60) {
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      continue
    }
    if (tag.localName === 'c') {
      if (cellOpeningTagUsesSharedString(bytes, tag.endIndex)) {
        return false
      }
      continue
    }
    if (tag.localName === 'v' || tag.localName === 'is') {
      valueCellCount += 1
    }
  }
  return valueCellCount > 0
}

function readXmlTagName(bytes: Uint8Array, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = bytes[startIndex]
  if (first === undefined || first === 33 || first === 47 || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < bytes.byteLength && isXmlNameByte(bytes[index] ?? 0)) {
    if (bytes[index] === 58) {
      localNameStart = index + 1
    }
    index += 1
  }
  if (index === localNameStart) {
    return null
  }
  return {
    localName: asciiSlice(bytes, localNameStart, index),
    endIndex: index,
  }
}

function cellOpeningTagUsesSharedString(bytes: Uint8Array, startIndex: number): boolean {
  let index = startIndex
  let quote: number | null = null
  while (index < bytes.byteLength) {
    const byte = bytes[index] ?? 0
    if (quote !== null) {
      if (byte === quote) {
        quote = null
      }
      index += 1
      continue
    }
    if (byte === 34 || byte === 39) {
      quote = byte
      index += 1
      continue
    }
    if (byte === 62 || byte === 47) {
      return false
    }
    if (byte === 116 && isAttributeBoundary(bytes[index - 1] ?? 0)) {
      const next = skipAsciiWhitespace(bytes, index + 1)
      if (bytes[next] === 61) {
        const valueStart = skipAsciiWhitespace(bytes, next + 1)
        const valueQuote = bytes[valueStart]
        if ((valueQuote === 34 || valueQuote === 39) && bytes[valueStart + 1] === 115 && bytes[valueStart + 2] === valueQuote) {
          return true
        }
      }
    }
    index += 1
  }
  return false
}

function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number): number {
  let index = startIndex
  while (index < bytes.byteLength && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

function isAttributeBoundary(byte: number): boolean {
  return byte === 0 || isAsciiWhitespace(byte)
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

function asciiSlice(bytes: Uint8Array, startIndex: number, endIndex: number): string {
  let output = ''
  for (let index = startIndex; index < endIndex; index += 1) {
    output += String.fromCharCode(bytes[index] ?? 0)
  }
  return output
}
