import { readLargeSimpleDataValidationsFromBytes } from '../packages/excel-import/src/xlsx-large-simple-data-validation-byte-scan.js'

const lessThanByte = 0x3c
const greaterThanByte = 0x3e
const slashByte = 0x2f
const doubleQuoteByte = 0x22
const singleQuoteByte = 0x27
const whitespaceBytes = new Set([0x09, 0x0a, 0x0d, 0x20])
const dataValidationElementNameBytes = asciiBytes('dataValidation')
const closeDataValidationElementBytes = asciiBytes('</dataValidation>')
const dataValidationElementTailLength = dataValidationElementNameBytes.length + 2

export class WorksheetDataValidationSupportScanner {
  private buffer = new Uint8Array()
  private dataValidationCount = 0
  private unsupportedDataValidationCount = 0

  push(chunk: Uint8Array, final: boolean): void {
    this.buffer = concatBytes(this.buffer, chunk)
    this.scanBufferedDataValidations(final)
  }

  finish(): { readonly dataValidationCount: number; readonly unsupportedDataValidationCount: number } {
    this.push(new Uint8Array(), true)
    return {
      dataValidationCount: this.dataValidationCount,
      unsupportedDataValidationCount: this.unsupportedDataValidationCount,
    }
  }

  private scanBufferedDataValidations(final: boolean): void {
    let index = 0
    while (index < this.buffer.length) {
      const tagStart = indexOfElementCandidate(this.buffer, dataValidationElementNameBytes, index)
      if (tagStart < 0) {
        this.retainFrom(final ? this.buffer.length : Math.max(index, this.buffer.length - dataValidationElementTailLength))
        return
      }
      if (!isElementStartBytes(this.buffer, tagStart, dataValidationElementNameBytes)) {
        index = tagStart + 2
        continue
      }
      const openingEnd = findXmlTagEnd(this.buffer, tagStart + dataValidationElementNameBytes.length + 1)
      if (openingEnd < 0) {
        if (final) {
          this.unsupportedDataValidationCount += 1
          this.retainFrom(this.buffer.length)
        } else {
          this.retainFrom(tagStart)
        }
        return
      }
      const selfClosing = isSelfClosingXmlTag(this.buffer, openingEnd)
      const closeStart = selfClosing ? openingEnd : indexOfBytes(this.buffer, closeDataValidationElementBytes, openingEnd + 1)
      if (closeStart < 0) {
        if (final) {
          this.unsupportedDataValidationCount += 1
          this.retainFrom(this.buffer.length)
        } else {
          this.retainFrom(tagStart)
        }
        return
      }
      const elementEnd = selfClosing ? openingEnd + 1 : closeStart + closeDataValidationElementBytes.length
      this.scanDataValidation(tagStart, elementEnd)
      index = Math.max(elementEnd, openingEnd + 1)
    }
    this.retainFrom(this.buffer.length)
  }

  private scanDataValidation(start: number, end: number): void {
    const validations = readLargeSimpleDataValidationsFromBytes('Sheet1', this.buffer, start, end)
    if (validations === null) {
      this.unsupportedDataValidationCount += 1
    } else {
      this.dataValidationCount += validations.length
    }
  }

  private retainFrom(index: number): void {
    this.buffer = copyBytes(this.buffer.subarray(index))
  }
}

function indexOfElementCandidate(xml: Uint8Array, elementName: Uint8Array, start: number): number {
  const maxStart = xml.length - elementName.length - 1
  for (let index = start; index <= maxStart; index += 1) {
    if (xml[index] !== lessThanByte) {
      continue
    }
    let matches = true
    for (let nameIndex = 0; nameIndex < elementName.length; nameIndex += 1) {
      if (xml[index + nameIndex + 1] !== elementName[nameIndex]) {
        matches = false
        break
      }
    }
    if (matches) {
      return index
    }
  }
  return -1
}

function isElementStartBytes(xml: Uint8Array, index: number, elementName: Uint8Array): boolean {
  if (xml[index] !== lessThanByte) {
    return false
  }
  for (let nameIndex = 0; nameIndex < elementName.length; nameIndex += 1) {
    if (xml[index + nameIndex + 1] !== elementName[nameIndex]) {
      return false
    }
  }
  const next = xml[index + elementName.length + 1]
  return next === undefined || next === slashByte || next === greaterThanByte || whitespaceBytes.has(next)
}

function findXmlTagEnd(source: Uint8Array, start: number): number {
  let quote: number | null = null
  for (let index = start; index < source.length; index += 1) {
    const byte = source[index] ?? 0
    if (quote !== null) {
      if (byte === quote) {
        quote = null
      }
      continue
    }
    if (byte === doubleQuoteByte || byte === singleQuoteByte) {
      quote = byte
      continue
    }
    if (byte === greaterThanByte) {
      return index
    }
  }
  return -1
}

function isSelfClosingXmlTag(source: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && whitespaceBytes.has(source[index] ?? 0)) {
    index -= 1
  }
  return source[index] === slashByte
}

function indexOfBytes(source: Uint8Array, search: Uint8Array, start: number): number {
  const maxStart = source.length - search.length
  for (let index = start; index <= maxStart; index += 1) {
    let matches = true
    for (let searchIndex = 0; searchIndex < search.length; searchIndex += 1) {
      if (source[index + searchIndex] !== search[searchIndex]) {
        matches = false
        break
      }
    }
    if (matches) {
      return index
    }
  }
  return -1
}

function concatBytes(left: Uint8Array, right: Uint8Array): Uint8Array {
  if (left.length === 0) {
    return right
  }
  if (right.length === 0) {
    return left
  }
  const combined = new Uint8Array(left.length + right.length)
  combined.set(left, 0)
  combined.set(right, left.length)
  return combined
}

function copyBytes(bytes: Uint8Array): Uint8Array {
  if (bytes.length === 0) {
    return new Uint8Array()
  }
  const copy = new Uint8Array(bytes.length)
  copy.set(bytes)
  return copy
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (character) => character.charCodeAt(0))
}
