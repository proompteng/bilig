import type { WorkbookRichTextCellSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { toDisplayText } from './workbook-import-helpers.js'
import { growUint32Array } from './xlsx-large-simple-array-storage.js'
import { ImportedWorkbookArenaBase } from './xlsx-large-simple-arena-base.js'
import { noPoolId, valueKindEmpty, valueKindSharedStringRef, valueKindString } from './xlsx-large-simple-arena-constants.js'
import { encodeCellAddress, isPreviewCell, maxSpreadsheetColumnCount, packArenaCellAddress } from './xlsx-large-simple-arena-helpers.js'
import type {
  ImportedWorkbookArenaSnapshot,
  ImportedWorksheetArenaCellInput,
  ImportedWorksheetArenaSharedStringCellInput,
} from './xlsx-large-simple-arena-types.js'
import { createLazyWorkbookRichTextCells } from './xlsx-large-simple-lazy-rich-text-cells.js'
import { createLazyWorkbookSheetCells } from './xlsx-large-simple-lazy-sheet-cells.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import { decodeCellAddress } from './xlsx-large-simple-xml-byte-utils.js'
export type { ImportedWorksheetCellScan } from './xlsx-large-simple-cell-scan.js'
export type {
  ImportedWorkbookArenaDedupeMode,
  ImportedWorkbookArenaOptions,
  ImportedWorkbookArenaSnapshot,
  ImportedWorksheetArenaCellInput,
  ImportedWorksheetArenaSharedStringCellInput,
} from './xlsx-large-simple-arena-types.js'
export { ImportedWorksheetStyleIndexArena } from './xlsx-large-simple-style-index-arena.js'

const lazyRichTextCellThreshold = 10_000

type WorkbookSheetCells = WorkbookSnapshot['sheets'][number]['cells']
type WorkbookSheetCell = WorkbookSheetCells[number]

export class ImportedWorkbookArena extends ImportedWorkbookArenaBase {
  addCell(input: ImportedWorksheetArenaCellInput): number {
    this.ensureCapacity(this.length + 1)
    const index = this.appendCell(input.sheetIndex, input.row, input.column)
    this.addValue(index, input.value)
    if (isPreviewCell(input.row, input.column) && input.value !== undefined) {
      const previewValue = this.materializeValue(index)
      if (previewValue !== undefined) {
        this.setPreviewValue(input.row, input.column, previewValue)
      }
    }
    return index
  }

  addSharedStringCell(input: ImportedWorksheetArenaSharedStringCellInput): number {
    this.ensureCapacity(this.length + 1)
    const index = this.appendCell(input.sheetIndex, input.row, input.column)
    this.valueKinds[index] = valueKindSharedStringRef
    this.storeSharedStringIndex(index, input.sharedStringIndex)
    return index
  }

  setFormula(cellIndex: number, formula: string): void {
    const formulaId = this.internFormula(formula)
    this.ensureFormulaIdStorage()[cellIndex] = formulaId
    const row = this.rowAt(cellIndex)
    const column = this.columnAt(cellIndex)
    if (isPreviewCell(row, column) && !this.hasPreviewValue(row, column)) {
      this.setPreviewValue(row, column, `=${this.formulas[formulaId] ?? formula}`)
    }
  }

  materializeSheetCells(sheetIndex: number): WorkbookSnapshot['sheets'][number]['cells'] {
    if (!this.hasCellsForSheet(sheetIndex)) {
      return []
    }
    const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
    cells.length = this.countMaterializedSheetCells(sheetIndex)
    let outputIndex = 0
    for (let index = 0; index < this.length; index += 1) {
      if (!this.cellBelongsToSheet(index, sheetIndex)) {
        continue
      }
      const cell = this.materializeCellAtArenaIndex(index)
      if (!cell) {
        continue
      }
      cells[outputIndex] = cell
      outputIndex += 1
    }
    return cells
  }

  materializeSheetCellsByAddress(sheetIndex: number, addresses: ReadonlySet<string>): ReadonlyMap<string, WorkbookSheetCell> {
    const requested = new Map<number, string>()
    for (const address of addresses) {
      const decoded = decodeCellAddress(address)
      if (!decoded || decoded.column >= maxSpreadsheetColumnCount) {
        continue
      }
      requested.set(packArenaCellAddress(decoded.row, decoded.column), encodeCellAddress(decoded.row, decoded.column))
    }
    const cellsByAddress = new Map<string, WorkbookSheetCell>()
    if (requested.size === 0 || !this.hasCellsForSheet(sheetIndex)) {
      return cellsByAddress
    }
    for (let index = 0; index < this.length; index += 1) {
      if (!this.cellBelongsToSheet(index, sheetIndex)) {
        continue
      }
      const address = requested.get(packArenaCellAddress(this.rowAt(index), this.columnAt(index)))
      if (!address || cellsByAddress.has(address)) {
        continue
      }
      const cell = this.materializeCellAtArenaIndex(index)
      if (cell) {
        cellsByAddress.set(address, cell)
      }
      if (cellsByAddress.size >= requested.size) {
        break
      }
    }
    return cellsByAddress
  }

  createLazySheetCells(sheetIndex: number): WorkbookSheetCells {
    const arenaIndexes = this.lazySheetCellIndexes(sheetIndex)
    const cellCount = typeof arenaIndexes === 'number' ? arenaIndexes : arenaIndexes.length
    return createLazyWorkbookSheetCells(cellCount, (index) => {
      if (!Number.isInteger(index) || index < 0 || index >= cellCount) {
        return undefined
      }
      return this.materializeCellAtArenaIndex(typeof arenaIndexes === 'number' ? index : (arenaIndexes[index] ?? -1))
    })
  }

  readPreviewText(row: number, column: number): string {
    return toDisplayText(this.readPreviewValue(row, column))
  }

  collectSharedStringIndexes(output: Set<number>): void {
    for (let index = 0; index < this.length; index += 1) {
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindSharedStringRef) {
        continue
      }
      const sharedStringIndex = this.sharedStringIndexAt(index)
      if (sharedStringIndex !== noPoolId) {
        output.add(sharedStringIndex)
      }
    }
  }

  resolveSharedStrings(sharedStrings: LargeSimpleSharedStrings): WorkbookRichTextCellSnapshot[] | null {
    const richTextCells: WorkbookRichTextCellSnapshot[] = []
    for (let index = 0; index < this.length; index += 1) {
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindSharedStringRef) {
        continue
      }
      const sharedStringIndex = this.sharedStringIndexAt(index)
      const entry = sharedStringIndex === noPoolId ? undefined : sharedStrings[sharedStringIndex]
      if (!entry) {
        return null
      }
      this.valueKinds[index] = valueKindString
      this.stringValueCount += 1
      const stringIds = this.ensureStringIdStorage()
      stringIds[index] = this.internString(entry.text)
      const row = this.rowAt(index)
      const column = this.columnAt(index)
      if (isPreviewCell(row, column)) {
        this.setPreviewValue(row, column, entry.text)
      }
      if (entry.rich) {
        const text = this.strings[stringIds[index] ?? noPoolId] ?? entry.text
        richTextCells.push({
          address: encodeCellAddress(row, column),
          text,
          storage: 'sharedString',
          xml: entry.xml ?? '',
        })
      }
    }
    return richTextCells
  }

  retainSharedStringReferences(
    sharedStrings: LargeSimpleSharedStrings,
    options: { readonly lazyRichTextCellThreshold?: number } = {},
  ): WorkbookRichTextCellSnapshot[] | null {
    let richTextCellIndexes: Uint32Array<ArrayBuffer> | undefined
    let richTextCellCount = 0
    this.sharedStrings = sharedStrings
    for (let index = 0; index < this.length; index += 1) {
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindSharedStringRef) {
        continue
      }
      const sharedStringIndex = this.sharedStringIndexAt(index)
      const entry = sharedStringIndex === noPoolId ? undefined : sharedStrings[sharedStringIndex]
      if (!entry) {
        return null
      }
      const row = this.rowAt(index)
      const column = this.columnAt(index)
      if (isPreviewCell(row, column)) {
        this.setPreviewValue(row, column, entry.text)
      }
      if (entry.rich) {
        if (!richTextCellIndexes) {
          richTextCellIndexes = new Uint32Array(1024)
        } else if (richTextCellCount >= richTextCellIndexes.length) {
          richTextCellIndexes = growUint32Array(richTextCellIndexes, richTextCellIndexes.length * 2)
        }
        richTextCellIndexes[richTextCellCount] = index
        richTextCellCount += 1
      }
    }
    const threshold = Math.max(0, Math.trunc(options.lazyRichTextCellThreshold ?? lazyRichTextCellThreshold))
    if (richTextCellCount <= threshold) {
      const richTextCells: WorkbookRichTextCellSnapshot[] = []
      for (let index = 0; index < richTextCellCount; index += 1) {
        richTextCells.push(this.materializeSharedStringRichTextCell(richTextCellIndexes?.[index] ?? -1))
      }
      return richTextCells
    }
    const lazyIndexes = richTextCellIndexes?.slice(0, richTextCellCount) ?? new Uint32Array(0)
    return createLazyWorkbookRichTextCells(lazyIndexes.length, (index) =>
      this.materializeSharedStringRichTextCell(lazyIndexes[index] ?? -1),
    )
  }

  resolveSharedStringsExcept(sharedStrings: LargeSimpleSharedStrings, retainedSharedStringIndexes: ReadonlySet<number>): boolean {
    for (let index = 0; index < this.length; index += 1) {
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindSharedStringRef) {
        continue
      }
      const sharedStringIndex = this.sharedStringIndexAt(index)
      const entry = sharedStringIndex === noPoolId ? undefined : sharedStrings[sharedStringIndex]
      if (!entry) {
        return false
      }
      const row = this.rowAt(index)
      const column = this.columnAt(index)
      if (retainedSharedStringIndexes.has(sharedStringIndex)) {
        if (isPreviewCell(row, column)) {
          this.setPreviewValue(row, column, entry.text)
        }
        continue
      }
      this.valueKinds[index] = valueKindString
      this.sharedStringRefCount = Math.max(0, this.sharedStringRefCount - 1)
      this.stringValueCount += 1
      this.ensureStringIdStorage()[index] = this.internString(entry.text)
      if (isPreviewCell(row, column)) {
        this.setPreviewValue(row, column, entry.text)
      }
    }
    return true
  }

  snapshot(): ImportedWorkbookArenaSnapshot {
    this.materializeCoordinateStorage(this.length)
    const stringIds = this.snapshotStringIds()
    return {
      sheetIndex: this.sheetIndex,
      ...(this.sheetIndexes ? { sheetIndexes: this.sheetIndexes.subarray(0, this.length) } : {}),
      rows: this.rows.subarray(0, this.length),
      columns: this.columns.subarray(0, this.length),
      valueKinds: this.valueKinds.subarray(0, this.length),
      ...(this.numberValues ? { numberValues: this.numberValues.subarray(0, this.length) } : {}),
      ...(this.tinyIntegerValues ? { tinyIntegerValues: this.tinyIntegerValues.subarray(0, this.length) } : {}),
      ...(this.smallIntegerValues ? { smallIntegerValues: this.smallIntegerValues.subarray(0, this.length) } : {}),
      ...(this.sparseSmallIntegerCount > 0 && this.sparseSmallIntegerCellIndexes
        ? { sparseSmallIntegerCellIndexes: this.sparseSmallIntegerCellIndexes.subarray(0, this.sparseSmallIntegerCount) }
        : {}),
      ...(this.sparseSmallIntegerCount > 0 && this.sparseSmallIntegerValues
        ? { sparseSmallIntegerValues: this.sparseSmallIntegerValues.subarray(0, this.sparseSmallIntegerCount) }
        : {}),
      ...(this.integerValues ? { integerValues: this.integerValues.subarray(0, this.length) } : {}),
      ...(this.sparseIntegerCount > 0 && this.sparseIntegerCellIndexes
        ? { sparseIntegerCellIndexes: this.sparseIntegerCellIndexes.subarray(0, this.sparseIntegerCount) }
        : {}),
      ...(this.sparseIntegerCount > 0 && this.sparseIntegerValues
        ? { sparseIntegerValues: this.sparseIntegerValues.subarray(0, this.sparseIntegerCount) }
        : {}),
      ...(stringIds ? { stringIds } : {}),
      ...(this.booleanValues ? { booleanValues: this.booleanValues.subarray(0, this.length) } : {}),
      ...(this.formulaIds ? { formulaIds: this.formulaIds.subarray(0, this.length) } : {}),
      strings: this.strings,
      formulas: this.formulas,
    }
  }

  release(): void {
    this.sheetIndex = null
    this.sheetIndexes = undefined
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
    this.valueKinds = new Uint8Array(0)
    this.numberValues = undefined
    this.tinyIntegerValues = undefined
    this.smallIntegerValues = undefined
    this.sparseSmallIntegerCellIndexes = undefined
    this.sparseSmallIntegerValues = undefined
    this.sparseSmallIntegerCount = 0
    this.integerValues = undefined
    this.sparseIntegerCellIndexes = undefined
    this.sparseIntegerValues = undefined
    this.sparseIntegerCount = 0
    this.stringIds = undefined
    this.sparseStringCellIndexes = undefined
    this.sparseStringIds = undefined
    this.booleanValues = undefined
    this.formulaIds = undefined
    this.length = 0
    this.denseRowMajorWidth = null
    this.linearCellIndexes = undefined
    this.linearRowMajorWidth = null
    this.strings.length = 0
    this.stringIdsByValue.clear()
    this.stringDedupeKeys.length = 0
    this.stringDedupeEvictionIndex = 0
    this.formulas.length = 0
    this.formulaIdsByValue.clear()
    this.formulaDedupeKeys.length = 0
    this.formulaDedupeEvictionIndex = 0
    this.sharedStrings = undefined
    this.stringValueCount = 0
    this.sharedStringRefCount = 0
    this.sharedStringRefsInNumberValues = false
    this.previewValues.fill(undefined)
    this.previewValueSet.fill(0)
  }

  releaseMaterializationScratch(): void {
    this.compactSparseStringIds()
    this.stringIdsByValue.clear()
    this.stringDedupeKeys.length = 0
    this.stringDedupeEvictionIndex = 0
    this.formulaIdsByValue.clear()
    this.formulaDedupeKeys.length = 0
    this.formulaDedupeEvictionIndex = 0
    this.previewValues.fill(undefined)
    this.previewValueSet.fill(0)
  }

  private materializeCellAtArenaIndex(
    index: number,
    options: { readonly includeCoordinates?: boolean } = {},
  ): WorkbookSheetCell | undefined {
    if (index < 0 || index >= this.length) {
      return undefined
    }
    const value = this.materializeValue(index)
    const formulaId = this.formulaIds?.[index] ?? noPoolId
    const formula = formulaId === noPoolId ? undefined : this.formulas[formulaId]
    if (value === undefined && formula === undefined) {
      return undefined
    }
    const row = this.rowAt(index)
    const col = this.columnAt(index)
    const cell: WorkbookSheetCell = {
      address: encodeCellAddress(row, col),
      ...(options.includeCoordinates ? { row, col } : {}),
    }
    if (value !== undefined) {
      cell.value = value
    }
    if (formula !== undefined) {
      cell.formula = formula
    }
    return cell
  }

  private countMaterializedSheetCells(sheetIndex: number): number {
    let count = 0
    for (let index = 0; index < this.length; index += 1) {
      if (!this.cellBelongsToSheet(index, sheetIndex)) {
        continue
      }
      const formulaId = this.formulaIds?.[index] ?? noPoolId
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindEmpty || formulaId !== noPoolId) {
        count += 1
      }
    }
    return count
  }

  private materializeSharedStringRichTextCell(index: number): WorkbookRichTextCellSnapshot {
    const sharedStringIndex = this.sharedStringIndexAt(index)
    const entry = sharedStringIndex === noPoolId ? undefined : this.sharedStrings?.[sharedStringIndex]
    return {
      address: encodeCellAddress(this.rowAt(index), this.columnAt(index)),
      text: entry?.text ?? '',
      storage: 'sharedString',
      xml: entry?.xml ?? '',
    }
  }

  private lazySheetCellIndexes(sheetIndex: number): Uint32Array | number {
    const count = this.countMaterializedSheetCells(sheetIndex)
    return this.sheetIndexes === undefined && this.sheetIndex === sheetIndex && count === this.length
      ? count
      : this.materializedSheetCellIndexesWithCount(sheetIndex, count)
  }

  private materializedSheetCellIndexesWithCount(sheetIndex: number, count: number): Uint32Array {
    const output = new Uint32Array(count)
    let outputIndex = 0
    for (let index = 0; index < this.length; index += 1) {
      if (!this.cellBelongsToSheet(index, sheetIndex)) {
        continue
      }
      const formulaId = this.formulaIds?.[index] ?? noPoolId
      if ((this.valueKinds[index] ?? valueKindEmpty) !== valueKindEmpty || formulaId !== noPoolId) {
        output[outputIndex] = index
        outputIndex += 1
      }
    }
    return output
  }

  sheetCellsAreDenseRowMajor(sheetIndex: number, width: number, height: number): boolean {
    if (this.denseRowMajorWidth !== null && this.sheetIndexes === undefined && this.sheetIndex === sheetIndex) {
      return this.denseRowMajorWidth === width && this.length === width * height
    }
    if (width <= 0 || height <= 0 || this.countMaterializedSheetCells(sheetIndex) !== width * height) {
      return false
    }
    let expected = 0
    for (let index = 0; index < this.length; index += 1) {
      if (!this.cellBelongsToSheet(index, sheetIndex)) {
        continue
      }
      const formulaId = this.formulaIds?.[index] ?? noPoolId
      if ((this.valueKinds[index] ?? valueKindEmpty) === valueKindEmpty && formulaId === noPoolId) {
        continue
      }
      if (this.rowAt(index) !== Math.floor(expected / width) || this.columnAt(index) !== expected % width) {
        return false
      }
      expected += 1
    }
    return expected === width * height
  }
}
