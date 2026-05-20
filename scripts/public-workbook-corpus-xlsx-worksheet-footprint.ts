import { Buffer } from 'node:buffer'
import { TextDecoder } from 'node:util'
import { createInflateRaw } from 'node:zlib'

import { forEachInflatedXlsxZipEntryChunk, type XlsxZipEntries } from '../packages/excel-import/src/xlsx-zip.js'
import { WorksheetDataValidationSupportScanner } from './public-workbook-corpus-xlsx-data-validation-footprint.ts'
import type { WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

export interface WorksheetFootprint {
  cellCount: number
  columnCount: number
  conditionalFormatCount: number
  dataValidationCount: number
  formulaCellCount: number
  largeSimpleUnsupportedElementCount: number
  largeSimpleUnsupportedDataValidationCount: number
  largeSimpleUnsupportedFormulaCellCount: number
  mergeCount: number
  rowCount: number
  sharedStringCellCount: number
  usedRange: WorkbookFootprint['workbookMetadata']['dimensions'][number]['usedRange']
  valueCellCount: number
  xmlCellCount: number
}

const decoder = new TextDecoder()
const lessThanByte = 0x3c
const greaterThanByte = 0x3e
const slashByte = 0x2f
const equalsByte = 0x3d
const doubleQuoteByte = 0x22
const singleQuoteByte = 0x27
const whitespaceBytes = new Set([0x09, 0x0a, 0x0d, 0x20])

const cElementNameBytes = asciiBytes('c')
const fElementNameBytes = asciiBytes('f')
const vElementNameBytes = asciiBytes('v')
const isElementNameBytes = asciiBytes('is')
const rAttributeNameBytes = asciiBytes('r')
const siAttributeNameBytes = asciiBytes('si')
const tAttributeNameBytes = asciiBytes('t')
const closeCellElementBytes = asciiBytes('</c>')
const closeFormulaElementBytes = asciiBytes('</f>')
const mergeCellElementNameBytes = asciiBytes('mergeCell')
const conditionalFormattingElementNameBytes = asciiBytes('conditionalFormatting')
const unsupportedLargeSimpleElementNameBytes = [asciiBytes('picture'), asciiBytes('sheetProtection')]
const countedElementTailLength =
  Math.max(
    mergeCellElementNameBytes.length,
    conditionalFormattingElementNameBytes.length,
    ...unsupportedLargeSimpleElementNameBytes.map((entry) => entry.length),
  ) + 2

export function addWorksheetFootprint(total: WorksheetFootprint, footprint: WorksheetFootprint): void {
  total.cellCount += footprint.cellCount
  total.formulaCellCount += footprint.formulaCellCount
  total.largeSimpleUnsupportedFormulaCellCount += footprint.largeSimpleUnsupportedFormulaCellCount
  total.valueCellCount += footprint.valueCellCount
  total.xmlCellCount += footprint.xmlCellCount
  total.sharedStringCellCount += footprint.sharedStringCellCount
  total.mergeCount += footprint.mergeCount
  total.conditionalFormatCount += footprint.conditionalFormatCount
  total.dataValidationCount += footprint.dataValidationCount
  total.largeSimpleUnsupportedElementCount += footprint.largeSimpleUnsupportedElementCount
  total.largeSimpleUnsupportedDataValidationCount += footprint.largeSimpleUnsupportedDataValidationCount
}

export function worksheetDimension(
  sheetName: string,
  footprint: WorksheetFootprint,
): WorkbookFootprint['workbookMetadata']['dimensions'][number] {
  return {
    sheetName,
    rowCount: footprint.rowCount,
    columnCount: footprint.columnCount,
    nonEmptyCellCount: footprint.cellCount,
    usedRange: footprint.usedRange,
  }
}

export function inspectWorksheetXmlBytes(xml: Uint8Array): WorksheetFootprint {
  const scanner = new WorksheetXmlByteFootprintScanner()
  scanner.push(xml)
  return scanner.finish()
}

export function inspectWorksheetZipEntryFromLazyZip(zip: XlsxZipEntries, path: string): WorksheetFootprint {
  const scanner = new WorksheetXmlByteFootprintScanner()
  const found = forEachInflatedXlsxZipEntryChunk(zip, path, (chunk) => scanner.push(chunk))
  return found ? scanner.finish() : emptyWorksheetFootprint()
}

export function inspectDeflatedWorksheetXmlBytes(compressed: Uint8Array): Promise<WorksheetFootprint> {
  const scanner = new WorksheetXmlByteFootprintScanner()
  const inflate = createInflateRaw()
  return new Promise<WorksheetFootprint>((resolvePromise, reject) => {
    inflate.on('data', (chunk: Uint8Array) => scanner.push(chunk))
    inflate.on('error', reject)
    inflate.on('end', () => {
      try {
        resolvePromise(scanner.finish())
      } catch (error) {
        reject(error)
      }
    })
    inflate.end(Buffer.from(compressed.buffer, compressed.byteOffset, compressed.byteLength))
  })
}

export function emptyWorksheetFootprint(): WorksheetFootprint {
  return {
    cellCount: 0,
    formulaCellCount: 0,
    valueCellCount: 0,
    rowCount: 0,
    columnCount: 0,
    usedRange: null,
    mergeCount: 0,
    conditionalFormatCount: 0,
    dataValidationCount: 0,
    xmlCellCount: 0,
    sharedStringCellCount: 0,
    largeSimpleUnsupportedElementCount: 0,
    largeSimpleUnsupportedDataValidationCount: 0,
    largeSimpleUnsupportedFormulaCellCount: 0,
  }
}

class WorksheetXmlByteFootprintScanner {
  private buffer = new Uint8Array()
  private readonly sharedFormulaIndexes = new Set<string>()
  private readonly counter = new WorksheetElementStartCounter()
  private readonly dataValidationScanner = new WorksheetDataValidationSupportScanner()
  private readonly footprint = emptyWorksheetFootprint()

  push(chunk: Uint8Array): void {
    this.counter.push(chunk, false)
    this.dataValidationScanner.push(chunk, false)
    this.buffer = concatBytes(this.buffer, chunk)
    this.scanBufferedCells(false)
  }

  finish(): WorksheetFootprint {
    this.counter.finish()
    const dataValidationSupport = this.dataValidationScanner.finish()
    this.scanBufferedCells(true)
    const counted = this.counter.counts()
    return {
      ...this.footprint,
      mergeCount: counted.mergeCount,
      conditionalFormatCount: counted.conditionalFormatCount,
      dataValidationCount: dataValidationSupport.dataValidationCount,
      largeSimpleUnsupportedDataValidationCount: dataValidationSupport.unsupportedDataValidationCount,
      largeSimpleUnsupportedElementCount: counted.largeSimpleUnsupportedElementCount,
    }
  }

  private scanBufferedCells(final: boolean): void {
    let index = 0
    while (index < this.buffer.length) {
      const tagStart = indexOfElementCandidate(this.buffer, cElementNameBytes, index)
      if (tagStart < 0) {
        this.retainFrom(final ? this.buffer.length : Math.max(index, this.buffer.length - 1))
        return
      }
      if (!isElementStartBytes(this.buffer, tagStart, cElementNameBytes)) {
        index = tagStart + 2
        continue
      }
      const openingEnd = indexOfByte(this.buffer, greaterThanByte, tagStart + 2)
      if (openingEnd < 0) {
        if (final) {
          this.retainFrom(this.buffer.length)
          return
        }
        this.retainFrom(tagStart)
        return
      }
      const selfClosing = this.buffer[openingEnd - 1] === slashByte
      const closeStart = selfClosing ? openingEnd : indexOfBytes(this.buffer, closeCellElementBytes, openingEnd + 1, this.buffer.length)
      if (closeStart < 0 && !final) {
        this.retainFrom(tagStart)
        return
      }
      const effectiveCloseStart = closeStart < 0 ? openingEnd : closeStart
      const cellEnd = selfClosing ? openingEnd + 1 : closeStart < 0 ? openingEnd : closeStart + closeCellElementBytes.length
      this.scanCell(tagStart, openingEnd, effectiveCloseStart, selfClosing)
      index = Math.max(cellEnd, openingEnd + 1)
    }
    this.retainFrom(this.buffer.length)
  }

  private scanCell(tagStart: number, openingEnd: number, closeStart: number, selfClosing: boolean): void {
    let hasFormula = false
    let hasValue = false
    this.footprint.xmlCellCount += 1
    if (readXmlAttributeInBytesRange(this.buffer, tagStart, openingEnd + 1, tAttributeNameBytes) === 's') {
      this.footprint.sharedStringCellCount += 1
    }
    if (!selfClosing && closeStart > openingEnd) {
      hasFormula = containsElementStartBytes(this.buffer, openingEnd + 1, closeStart, fElementNameBytes)
      hasValue =
        containsElementStartBytes(this.buffer, openingEnd + 1, closeStart, vElementNameBytes) ||
        containsElementStartBytes(this.buffer, openingEnd + 1, closeStart, isElementNameBytes)
    }
    if (!hasFormula && !hasValue) {
      return
    }
    this.footprint.cellCount += 1
    if (hasFormula) {
      this.footprint.formulaCellCount += 1
      if (isUnsupportedLargeSimpleFormulaCell(this.buffer, openingEnd + 1, closeStart, this.sharedFormulaIndexes)) {
        this.footprint.largeSimpleUnsupportedFormulaCellCount += 1
      }
    }
    if (hasValue) {
      this.footprint.valueCellCount += 1
    }
    const address = readXmlAttributeInBytesRange(this.buffer, tagStart, openingEnd + 1, rAttributeNameBytes)
    const decoded = address ? decodeCellAddress(address) : null
    if (decoded) {
      this.footprint.rowCount = Math.max(this.footprint.rowCount, decoded.row + 1)
      this.footprint.columnCount = Math.max(this.footprint.columnCount, decoded.column + 1)
      this.footprint.usedRange = expandUsedRange(this.footprint.usedRange, decoded.row, decoded.column)
    }
  }

  private retainFrom(index: number): void {
    this.buffer = copyBytes(this.buffer.subarray(index))
  }
}

class WorksheetElementStartCounter {
  private tail = new Uint8Array()
  private mergeCellCount = 0
  private conditionalFormattingCount = 0
  private largeSimpleUnsupportedElementCount = 0

  push(chunk: Uint8Array, final: boolean): void {
    const buffer = concatBytes(this.tail, chunk)
    const scanEnd = final ? buffer.length : Math.max(0, buffer.length - countedElementTailLength)
    this.mergeCellCount += countElementStartsBytes(buffer, mergeCellElementNameBytes, 0, scanEnd)
    this.conditionalFormattingCount += countElementStartsBytes(buffer, conditionalFormattingElementNameBytes, 0, scanEnd)
    for (const elementName of unsupportedLargeSimpleElementNameBytes) {
      this.largeSimpleUnsupportedElementCount += countElementStartsBytes(buffer, elementName, 0, scanEnd)
    }
    this.tail = copyBytes(buffer.subarray(scanEnd))
  }

  finish(): void {
    this.push(new Uint8Array(), true)
  }

  counts(): Pick<WorksheetFootprint, 'conditionalFormatCount' | 'largeSimpleUnsupportedElementCount' | 'mergeCount'> {
    return {
      mergeCount: this.mergeCellCount,
      conditionalFormatCount: this.conditionalFormattingCount,
      largeSimpleUnsupportedElementCount: this.largeSimpleUnsupportedElementCount,
    }
  }
}

function countElementStartsBytes(xml: Uint8Array, elementName: Uint8Array, start = 0, end = xml.length): number {
  let count = 0
  let index = start
  while (index < end) {
    const tagStart = indexOfElementCandidate(xml, elementName, index)
    if (tagStart < 0 || tagStart >= end) {
      return count
    }
    if (isElementStartBytes(xml, tagStart, elementName)) {
      count += 1
    }
    index = tagStart + elementName.length + 1
  }
  return count
}

function containsElementStartBytes(xml: Uint8Array, start: number, end: number, elementName: Uint8Array): boolean {
  let index = start
  while (index < end) {
    const tagStart = indexOfElementCandidate(xml, elementName, index, end)
    if (tagStart < 0) {
      return false
    }
    if (isElementStartBytes(xml, tagStart, elementName)) {
      return true
    }
    index = tagStart + elementName.length + 1
  }
  return false
}

function isUnsupportedLargeSimpleFormulaCell(xml: Uint8Array, start: number, end: number, sharedFormulaIndexes: Set<string>): boolean {
  const formulaStart = indexOfElementCandidate(xml, fElementNameBytes, start, end)
  if (formulaStart < 0 || !isElementStartBytes(xml, formulaStart, fElementNameBytes)) {
    return false
  }
  const openingEnd = indexOfByte(xml, greaterThanByte, formulaStart, end)
  if (openingEnd < 0) {
    return true
  }
  const formulaType = readXmlAttributeInBytesRange(xml, formulaStart, openingEnd + 1, tAttributeNameBytes)
  const sharedFormulaIndex =
    formulaType === 'shared' ? readXmlAttributeInBytesRange(xml, formulaStart, openingEnd + 1, siAttributeNameBytes) : null
  if (formulaType === 'array' || formulaType === 'dataTable') {
    return true
  }
  if (xml[openingEnd - 1] === slashByte) {
    return formulaType !== 'shared' || sharedFormulaIndex === null || !sharedFormulaIndexes.has(sharedFormulaIndex)
  }
  const formulaEnd = indexOfBytes(xml, closeFormulaElementBytes, openingEnd + 1, end)
  if (formulaEnd < 0) {
    return true
  }
  const formula = decodeXmlEntities(decodeBytes(xml.subarray(openingEnd + 1, formulaEnd))).trim()
  if (formula.length === 0) {
    return formulaType !== 'shared' || sharedFormulaIndex === null || !sharedFormulaIndexes.has(sharedFormulaIndex)
  }
  const unsupported =
    formula.length === 0 ||
    /(?:^|[=,+(*/\s])'?\[[^\]]+\]/u.test(formula) ||
    /\[[#@\w]/u.test(formula) ||
    /(?:^|[^A-Z0-9_.])(?:NOW|RAND|RANDBETWEEN|TODAY)\s*\(/iu.test(formula)
  if (!unsupported && sharedFormulaIndex !== null) {
    sharedFormulaIndexes.add(sharedFormulaIndex)
  }
  return unsupported
}

function indexOfElementCandidate(xml: Uint8Array, elementName: Uint8Array, start: number, end = xml.length): number {
  const maxStart = end - elementName.length - 1
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
  return next === undefined || isElementBoundaryByte(next)
}

function isElementBoundaryByte(value: number): boolean {
  return value === slashByte || value === greaterThanByte || whitespaceBytes.has(value)
}

function readXmlAttributeInBytesRange(source: Uint8Array, start: number, end: number, name: Uint8Array): string | null {
  for (let index = start; index < end; index += 1) {
    if (!whitespaceBytes.has(source[index])) {
      continue
    }
    let nameMatches = true
    for (let nameIndex = 0; nameIndex < name.length; nameIndex += 1) {
      if (source[index + nameIndex + 1] !== name[nameIndex]) {
        nameMatches = false
        break
      }
    }
    if (!nameMatches || source[index + name.length + 1] !== equalsByte) {
      continue
    }
    const quote = source[index + name.length + 2]
    if (quote !== doubleQuoteByte && quote !== singleQuoteByte) {
      continue
    }
    const valueStart = index + name.length + 3
    const valueEnd = indexOfByte(source, quote, valueStart, end)
    if (valueEnd < 0) {
      return null
    }
    return decodeXmlEntities(decodeBytes(source.subarray(valueStart, valueEnd)))
  }
  return null
}

function indexOfByte(source: Uint8Array, byte: number, start: number, end = source.length): number {
  for (let index = start; index < end; index += 1) {
    if (source[index] === byte) {
      return index
    }
  }
  return -1
}

function indexOfBytes(source: Uint8Array, search: Uint8Array, start: number, end: number): number {
  const maxStart = end - search.length
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

function decodeCellAddress(address: string): { readonly row: number; readonly column: number } | null {
  const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/iu.exec(address)
  if (!match) {
    return null
  }
  let column = 0
  for (const character of match[1].toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  return Number.isSafeInteger(row) ? { row: row - 1, column: column - 1 } : null
}

function expandUsedRange(current: WorksheetFootprint['usedRange'], row: number, column: number): WorksheetFootprint['usedRange'] {
  return current
    ? {
        startRow: Math.min(current.startRow, row),
        startColumn: Math.min(current.startColumn, column),
        endRow: Math.max(current.endRow, row),
        endColumn: Math.max(current.endColumn, column),
      }
    : { startRow: row, startColumn: column, endRow: row, endColumn: column }
}

function decodeXmlEntities(value: string): string {
  return value.replace(/&(#x?[0-9a-f]+|amp|apos|gt|lt|quot);/giu, (entity, raw: string) => {
    switch (raw) {
      case 'amp':
        return '&'
      case 'apos':
        return "'"
      case 'gt':
        return '>'
      case 'lt':
        return '<'
      case 'quot':
        return '"'
      default: {
        const radix = raw.toLowerCase().startsWith('#x') ? 16 : 10
        const digits = raw.replace(/^#x?/iu, '')
        const codePoint = Number.parseInt(digits, radix)
        return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity
      }
    }
  })
}

function decodeBytes(bytes: Uint8Array): string {
  return decoder.decode(bytes)
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (character) => character.charCodeAt(0))
}
