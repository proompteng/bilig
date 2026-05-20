import type { LiteralInput, WorkbookRichTextCellSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { toDisplayText } from './workbook-import-helpers.js'
import { filledUint32Array, growFloat64Array, growUint8Array, growUint16Array, growUint32Array } from './xlsx-large-simple-array-storage.js'
import { createLazyWorkbookRichTextCells } from './xlsx-large-simple-lazy-rich-text-cells.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import type { ImportedWorksheetStyleIndexArena } from './xlsx-large-simple-style-index-arena.js'
export { ImportedWorksheetStyleIndexArena } from './xlsx-large-simple-style-index-arena.js'

const initialCellCapacity = 1024
const noPoolId = 0xffffffff
const valueKindEmpty = 0
const valueKindNumber = 1
const valueKindString = 2
const valueKindBoolean = 3
const valueKindNull = 4
const valueKindSharedStringRef = 5
const previewRowCount = 8
const previewColumnCount = 6
const previewCellCount = previewRowCount * previewColumnCount
const lazyRichTextCellThreshold = 10_000

export interface ImportedWorksheetArenaCellInput {
  readonly sheetIndex: number
  readonly row: number
  readonly column: number
  readonly value: LiteralInput | undefined
}

export interface ImportedWorksheetArenaSharedStringCellInput {
  readonly sheetIndex: number
  readonly row: number
  readonly column: number
  readonly sharedStringIndex: number
}

export interface ImportedWorkbookArenaSnapshot {
  readonly sheetIndex: number | null
  readonly sheetIndexes?: Uint32Array
  readonly rows: Uint32Array
  readonly columns: Uint16Array
  readonly valueKinds: Uint8Array
  readonly numberValues?: Float64Array
  readonly stringIds?: Uint32Array
  readonly booleanValues?: Uint8Array
  readonly formulaIds?: Uint32Array
  readonly strings: readonly string[]
  readonly formulas: readonly string[]
}

type WorkbookSheetCells = WorkbookSnapshot['sheets'][number]['cells']
type WorkbookSheetCell = WorkbookSheetCells[number]

const lazyCellsBrand = Symbol('bilig.lazyImportedXlsxCells')

export interface ImportedWorkbookLazySheetCells extends Array<WorkbookSheetCell> {
  readonly [lazyCellsBrand]: true
}

export interface ImportedWorkbookArenaOptions {
  readonly deduplicateStrings?: ImportedWorkbookArenaDedupeMode
  readonly deduplicateFormulas?: ImportedWorkbookArenaDedupeMode
  readonly dedupeMaxEntries?: number
}

export type ImportedWorkbookArenaDedupeMode = boolean | 'bounded'

export class ImportedWorkbookArena {
  private sheetIndex: number | null = null
  private sheetIndexes: Uint32Array<ArrayBuffer> | undefined
  private rows: Uint32Array<ArrayBuffer> = new Uint32Array(initialCellCapacity)
  private columns: Uint16Array<ArrayBuffer> = new Uint16Array(initialCellCapacity)
  private valueKinds: Uint8Array<ArrayBuffer> = new Uint8Array(initialCellCapacity)
  private numberValues: Float64Array<ArrayBuffer> | undefined
  private stringIds: Uint32Array<ArrayBuffer> | undefined
  private booleanValues: Uint8Array<ArrayBuffer> | undefined
  private formulaIds: Uint32Array<ArrayBuffer> | undefined
  private length = 0
  private denseRowMajorWidth: number | null = null
  private readonly strings: string[] = []
  private readonly stringIdsByValue = new Map<string, number>()
  private readonly formulas: string[] = []
  private readonly formulaIdsByValue = new Map<string, number>()
  private sharedStrings: LargeSimpleSharedStrings | undefined
  private stringValueCount = 0
  private sharedStringRefCount = 0
  private sharedStringRefsInNumberValues = false
  private readonly previewValues: (LiteralInput | undefined)[] = Array.from({ length: previewCellCount })
  private readonly previewValueSet = new Uint8Array(previewCellCount)
  private readonly stringDedupeMode: ImportedWorkbookArenaDedupeMode
  private readonly formulaDedupeMode: ImportedWorkbookArenaDedupeMode
  private readonly dedupeMaxEntries: number
  private readonly stringDedupeKeys: string[] = []
  private stringDedupeEvictionIndex = 0
  private readonly formulaDedupeKeys: string[] = []
  private formulaDedupeEvictionIndex = 0

  constructor(
    private readonly stringPool?: ImportedWorkbookStringPool,
    options: ImportedWorkbookArenaOptions = {},
  ) {
    this.stringDedupeMode = options.deduplicateStrings ?? true
    this.formulaDedupeMode = options.deduplicateFormulas ?? true
    this.dedupeMaxEntries = Math.max(0, Math.trunc(options.dedupeMaxEntries ?? 8192))
  }

  get cellCount(): number {
    return this.length
  }

  reserveCellCapacity(capacity: number): void {
    if (!Number.isSafeInteger(capacity) || capacity <= this.valueKinds.length) {
      return
    }
    this.resizeStorage(capacity)
  }

  reserveDenseRowMajorCellCapacity(sheetIndex: number, width: number, height: number): void {
    const capacity = width * height
    if (!Number.isSafeInteger(capacity) || width <= 0 || height <= 0) {
      return
    }
    if (this.length !== 0 || this.sheetIndexes || (this.sheetIndex !== null && this.sheetIndex !== sheetIndex)) {
      this.reserveCellCapacity(capacity)
      return
    }
    this.sheetIndex = sheetIndex
    this.denseRowMajorWidth = width
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
    this.reserveCellCapacity(capacity)
  }

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

  createLazySheetCells(sheetIndex: number): WorkbookSheetCells {
    const arenaIndexes = this.lazySheetCellIndexes(sheetIndex)
    const cellCount = typeof arenaIndexes === 'number' ? arenaIndexes : arenaIndexes.length
    const materialize = (index: number): WorkbookSheetCell | undefined => {
      if (!Number.isInteger(index) || index < 0 || index >= cellCount) {
        return undefined
      }
      return this.materializeCellAtArenaIndex(typeof arenaIndexes === 'number' ? index : (arenaIndexes[index] ?? -1))
    }
    const iterate = function* (): IterableIterator<WorkbookSheetCell> {
      for (let index = 0; index < cellCount; index += 1) {
        const cell = materialize(index)
        if (cell) {
          yield cell
        }
      }
    }
    const target: ImportedWorkbookLazySheetCells = Object.assign([], { [lazyCellsBrand]: true as const })
    let proxy: WorkbookSheetCells
    proxy = new Proxy<ImportedWorkbookLazySheetCells>(target, {
      get: (_target, property) => {
        if (property === lazyCellsBrand) {
          return true
        }
        if (property === 'length') {
          return cellCount
        }
        if (property === Symbol.iterator || property === 'values') {
          return iterate
        }
        if (property === 'entries') {
          return function* entries(): IterableIterator<[number, WorkbookSheetCell]> {
            for (let index = 0; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell) {
                yield [index, cell]
              }
            }
          }
        }
        if (property === 'keys') {
          return function* keys(): IterableIterator<number> {
            for (let index = 0; index < cellCount; index += 1) {
              yield index
            }
          }
        }
        if (property === 'at') {
          return (index: number) => materialize(index < 0 ? cellCount + index : index)
        }
        if (property === 'forEach') {
          return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => void, thisArg?: unknown) => {
            for (let index = 0; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell) {
                callback.call(thisArg, cell, index, proxy)
              }
            }
          }
        }
        if (property === 'map') {
          return <T>(callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => T, thisArg?: unknown): T[] => {
            const output: T[] = []
            for (let index = 0; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell) {
                output.push(callback.call(thisArg, cell, index, proxy))
              }
            }
            return output
          }
        }
        if (property === 'filter') {
          return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
            const output: WorkbookSheetCell[] = []
            for (let index = 0; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell && callback.call(thisArg, cell, index, proxy)) {
                output.push(cell)
              }
            }
            return output
          }
        }
        if (property === 'reduce') {
          return function reduce(
            callback: (previous: unknown, cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => unknown,
            initialValue?: unknown,
          ): unknown {
            let index = 0
            let accumulator = initialValue
            if (arguments.length < 2) {
              const first = materialize(0)
              if (!first) {
                throw new TypeError('Reduce of empty array with no initial value')
              }
              accumulator = first
              index = 1
            }
            for (; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell) {
                accumulator = callback(accumulator, cell, index, proxy)
              }
            }
            return accumulator
          }
        }
        if (property === 'some') {
          return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
            for (let index = 0; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell && callback.call(thisArg, cell, index, proxy)) {
                return true
              }
            }
            return false
          }
        }
        if (property === 'every') {
          return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
            for (let index = 0; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell && !callback.call(thisArg, cell, index, proxy)) {
                return false
              }
            }
            return true
          }
        }
        if (property === 'find') {
          return (callback: (cell: WorkbookSheetCell, index: number, cells: WorkbookSheetCells) => boolean, thisArg?: unknown) => {
            for (let index = 0; index < cellCount; index += 1) {
              const cell = materialize(index)
              if (cell && callback.call(thisArg, cell, index, proxy)) {
                return cell
              }
            }
            return undefined
          }
        }
        if (property === 'slice') {
          return (start?: number, end?: number) => {
            const length = cellCount
            const from = normalizeSliceIndex(start ?? 0, length)
            const to = normalizeSliceIndex(end ?? length, length)
            const output: WorkbookSheetCell[] = []
            for (let index = from; index < to; index += 1) {
              const cell = materialize(index)
              if (cell) {
                output.push(cell)
              }
            }
            return output
          }
        }
        if (property === 'toJSON' || property === 'toArray') {
          return () => Array.from(iterate())
        }
        if (typeof property === 'string' && isArrayIndexProperty(property)) {
          return materialize(Number(property))
        }
        return Reflect.get(Array.prototype, property)
      },
      has: (_target, property) => property === 'length' || (typeof property === 'string' && isArrayIndexProperty(property)),
      getOwnPropertyDescriptor: (_target, property) => {
        if (property === 'length') {
          return { configurable: true, enumerable: false, value: cellCount }
        }
        if (typeof property === 'string' && isArrayIndexProperty(property)) {
          const value = materialize(Number(property))
          return value === undefined ? undefined : { configurable: true, enumerable: true, value }
        }
        return undefined
      },
    })
    return proxy
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

  snapshot(): ImportedWorkbookArenaSnapshot {
    this.materializeCoordinateStorage(this.length)
    return {
      sheetIndex: this.sheetIndex,
      ...(this.sheetIndexes ? { sheetIndexes: this.sheetIndexes.subarray(0, this.length) } : {}),
      rows: this.rows.subarray(0, this.length),
      columns: this.columns.subarray(0, this.length),
      valueKinds: this.valueKinds.subarray(0, this.length),
      ...(this.numberValues ? { numberValues: this.numberValues.subarray(0, this.length) } : {}),
      ...(this.stringIds ? { stringIds: this.stringIds.subarray(0, this.length) } : {}),
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
    this.stringIds = undefined
    this.booleanValues = undefined
    this.formulaIds = undefined
    this.length = 0
    this.denseRowMajorWidth = null
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
    this.stringIdsByValue.clear()
    this.stringDedupeKeys.length = 0
    this.stringDedupeEvictionIndex = 0
    this.formulaIdsByValue.clear()
    this.formulaDedupeKeys.length = 0
    this.formulaDedupeEvictionIndex = 0
    this.previewValues.fill(undefined)
    this.previewValueSet.fill(0)
  }

  private appendCell(sheetIndex: number, row: number, column: number): number {
    const index = this.length
    this.length += 1
    this.recordSheetIndex(index, sheetIndex)
    if (this.denseRowMajorWidth !== null) {
      const expectedRow = Math.floor(index / this.denseRowMajorWidth)
      if (
        this.sheetIndexes === undefined &&
        this.sheetIndex === sheetIndex &&
        row === expectedRow &&
        column === index % this.denseRowMajorWidth
      ) {
        return index
      }
      this.materializeCoordinateStorage(index)
    }
    this.rows[index] = row
    this.columns[index] = column
    return index
  }

  private recordSheetIndex(index: number, sheetIndex: number): void {
    if (this.sheetIndexes) {
      this.sheetIndexes[index] = sheetIndex
      return
    }
    if (this.sheetIndex === null) {
      this.sheetIndex = sheetIndex
      return
    }
    if (this.sheetIndex === sheetIndex) {
      return
    }
    const sheetIndexes = new Uint32Array(this.valueKinds.length)
    sheetIndexes.fill(this.sheetIndex, 0, index)
    sheetIndexes[index] = sheetIndex
    this.sheetIndexes = sheetIndexes
  }

  private addValue(index: number, value: LiteralInput | undefined): void {
    if (value === undefined) {
      this.valueKinds[index] = valueKindEmpty
      return
    }
    if (value === null) {
      this.valueKinds[index] = valueKindNull
      return
    }
    if (typeof value === 'number') {
      this.valueKinds[index] = valueKindNumber
      this.ensureNumberValueStorage()[index] = value
      return
    }
    if (typeof value === 'boolean') {
      this.valueKinds[index] = valueKindBoolean
      this.ensureBooleanValueStorage()[index] = value ? 1 : 0
      return
    }
    this.valueKinds[index] = valueKindString
    this.stringValueCount += 1
    this.ensureStringIdStorage()[index] = this.internString(value)
  }

  private materializeValue(index: number): LiteralInput | undefined {
    const valueKind = this.valueKinds[index] ?? valueKindEmpty
    switch (valueKind) {
      case valueKindNumber:
        return this.numberValues?.[index]
      case valueKindString: {
        const stringId = this.stringIds?.[index] ?? noPoolId
        return stringId === noPoolId ? undefined : this.strings[stringId]
      }
      case valueKindSharedStringRef: {
        const sharedStringIndex = this.sharedStringIndexAt(index)
        return sharedStringIndex === noPoolId ? undefined : this.sharedStrings?.[sharedStringIndex]?.text
      }
      case valueKindBoolean:
        return (this.booleanValues?.[index] ?? 0) === 1
      case valueKindNull:
        return null
      default:
        return undefined
    }
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

  private hasCellsForSheet(sheetIndex: number): boolean {
    return this.sheetIndexes !== undefined || this.sheetIndex === sheetIndex
  }

  private cellBelongsToSheet(index: number, sheetIndex: number): boolean {
    return this.sheetIndexes ? this.sheetIndexes[index] === sheetIndex : this.sheetIndex === sheetIndex
  }

  private ensureStringIdStorage(): Uint32Array<ArrayBuffer> {
    if (this.stringIds) {
      return this.stringIds
    }
    this.stringIds = filledUint32Array(this.valueKinds.length, noPoolId)
    return this.stringIds
  }

  private ensureNumberValueStorage(): Float64Array<ArrayBuffer> {
    if (this.numberValues) {
      return this.numberValues
    }
    this.numberValues = new Float64Array(this.valueKinds.length)
    this.numberValues.fill(Number.NaN)
    this.moveSharedStringIndexesToNumberValues()
    return this.numberValues
  }

  private ensureFormulaIdStorage(): Uint32Array<ArrayBuffer> {
    if (this.formulaIds) {
      return this.formulaIds
    }
    this.formulaIds = filledUint32Array(this.valueKinds.length, noPoolId)
    return this.formulaIds
  }

  private ensureBooleanValueStorage(): Uint8Array<ArrayBuffer> {
    if (this.booleanValues) {
      return this.booleanValues
    }
    this.booleanValues = new Uint8Array(this.valueKinds.length)
    return this.booleanValues
  }

  private storeSharedStringIndex(index: number, sharedStringIndex: number): void {
    this.sharedStringRefCount += 1
    if (this.numberValues) {
      this.sharedStringRefsInNumberValues = true
      this.numberValues[index] = sharedStringIndex
      return
    }
    this.ensureStringIdStorage()[index] = sharedStringIndex
  }

  private sharedStringIndexAt(index: number): number {
    if (this.sharedStringRefsInNumberValues) {
      const value = this.numberValues?.[index]
      if (value !== undefined && !Number.isNaN(value)) {
        return Math.trunc(value)
      }
    }
    return this.stringIds?.[index] ?? noPoolId
  }

  private moveSharedStringIndexesToNumberValues(): void {
    if (!this.numberValues || !this.stringIds || this.sharedStringRefCount === 0 || this.stringValueCount > 0) {
      return
    }
    for (let index = 0; index < this.length; index += 1) {
      if ((this.valueKinds[index] ?? valueKindEmpty) === valueKindSharedStringRef) {
        this.numberValues[index] = this.stringIds[index] ?? noPoolId
      }
    }
    this.stringIds = undefined
    this.sharedStringRefsInNumberValues = true
  }

  private internString(value: string): number {
    const interned = this.internValue(value, this.stringDedupeMode)
    if (interned === null) {
      const next = this.strings.length
      this.strings.push(value)
      return next
    }
    const existing = this.stringIdsByValue.get(interned)
    if (existing !== undefined) {
      return existing
    }
    const next = this.strings.length
    this.strings.push(interned)
    this.stringIdsByValue.set(interned, next)
    if (this.stringDedupeMode === 'bounded') {
      this.rememberBoundedDedupeKey(this.stringIdsByValue, this.stringDedupeKeys, 'string', interned)
    }
    return next
  }

  private internFormula(value: string): number {
    const interned = this.internValue(value, this.formulaDedupeMode)
    if (interned === null) {
      const next = this.formulas.length
      this.formulas.push(value)
      return next
    }
    const existing = this.formulaIdsByValue.get(interned)
    if (existing !== undefined) {
      return existing
    }
    const next = this.formulas.length
    this.formulas.push(interned)
    this.formulaIdsByValue.set(interned, next)
    if (this.formulaDedupeMode === 'bounded') {
      this.rememberBoundedDedupeKey(this.formulaIdsByValue, this.formulaDedupeKeys, 'formula', interned)
    }
    return next
  }

  private internValue(value: string, mode: ImportedWorkbookArenaDedupeMode): string | null {
    if (mode === false) {
      return null
    }
    if (mode === 'bounded') {
      return this.stringPool?.internBounded(value, this.dedupeMaxEntries) ?? value
    }
    return this.stringPool?.intern(value) ?? value
  }

  private rememberBoundedDedupeKey(map: Map<string, number>, keys: string[], kind: 'string' | 'formula', key: string): void {
    keys.push(key)
    let evictionIndex = kind === 'string' ? this.stringDedupeEvictionIndex : this.formulaDedupeEvictionIndex
    while (keys.length - evictionIndex > this.dedupeMaxEntries) {
      const evicted = keys[evictionIndex]
      evictionIndex += 1
      if (evicted !== undefined) {
        map.delete(evicted)
      }
    }
    if (evictionIndex > this.dedupeMaxEntries && evictionIndex * 2 > keys.length) {
      keys.splice(0, evictionIndex)
      evictionIndex = 0
    }
    if (kind === 'string') {
      this.stringDedupeEvictionIndex = evictionIndex
    } else {
      this.formulaDedupeEvictionIndex = evictionIndex
    }
  }

  private rowAt(index: number): number {
    return this.denseRowMajorWidth === null ? (this.rows[index] ?? 0) : Math.floor(index / this.denseRowMajorWidth)
  }

  private columnAt(index: number): number {
    return this.denseRowMajorWidth === null ? (this.columns[index] ?? 0) : index % this.denseRowMajorWidth
  }

  private materializeCoordinateStorage(upToIndex: number): void {
    const width = this.denseRowMajorWidth
    if (width === null) {
      return
    }
    this.rows = new Uint32Array(this.valueKinds.length)
    this.columns = new Uint16Array(this.valueKinds.length)
    for (let index = 0; index < upToIndex; index += 1) {
      this.rows[index] = Math.floor(index / width)
      this.columns[index] = index % width
    }
    this.denseRowMajorWidth = null
  }

  private setPreviewValue(row: number, column: number, value: LiteralInput): void {
    const index = previewIndex(row, column)
    if (index === -1) {
      return
    }
    this.previewValues[index] = value
    this.previewValueSet[index] = 1
  }

  private hasPreviewValue(row: number, column: number): boolean {
    const index = previewIndex(row, column)
    return index !== -1 && this.previewValueSet[index] === 1
  }

  private readPreviewValue(row: number, column: number): LiteralInput | undefined {
    const index = previewIndex(row, column)
    return index === -1 || this.previewValueSet[index] !== 1 ? undefined : this.previewValues[index]
  }

  private ensureCapacity(nextLength: number): void {
    if (nextLength <= this.valueKinds.length) {
      return
    }
    let nextCapacity = this.valueKinds.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.resizeStorage(nextCapacity)
  }

  private resizeStorage(nextCapacity: number): void {
    if (this.sheetIndexes) {
      this.sheetIndexes = growUint32Array(this.sheetIndexes, nextCapacity)
    }
    if (this.denseRowMajorWidth === null) {
      this.rows = growUint32Array(this.rows, nextCapacity)
      this.columns = growUint16Array(this.columns, nextCapacity)
    }
    this.valueKinds = growUint8Array(this.valueKinds, nextCapacity)
    if (this.numberValues) {
      this.numberValues = growFloat64Array(this.numberValues, nextCapacity)
    }
    if (this.stringIds) {
      this.stringIds = growUint32Array(this.stringIds, nextCapacity, noPoolId)
    }
    if (this.booleanValues) {
      this.booleanValues = growUint8Array(this.booleanValues, nextCapacity)
    }
    if (this.formulaIds) {
      this.formulaIds = growUint32Array(this.formulaIds, nextCapacity, noPoolId)
    }
  }
}

export interface ImportedWorksheetCellScan {
  readonly arena: ImportedWorkbookArena
  readonly sheetIndex: number
  readonly richTextCells: WorkbookRichTextCellSnapshot[]
  readonly styleIndexes: ImportedWorksheetStyleIndexArena
  readonly blankStyleCellCount: number
  readonly cellCount: number
  readonly valueCellCount: number
  readonly formulaCellCount: number
  readonly mergeCount?: number
  readonly conditionalFormatCount?: number
  readonly dataValidationCount?: number
  readonly tableCount?: number
  readonly rowCount: number
  readonly columnCount: number
  readonly usedRange: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  } | null
}

function isPreviewCell(row: number, column: number): boolean {
  return row >= 0 && row < previewRowCount && column >= 0 && column < previewColumnCount
}

function previewIndex(row: number, column: number): number {
  return isPreviewCell(row, column) ? row * previewColumnCount + column : -1
}

function isArrayIndexProperty(property: string): boolean {
  if (property.length === 0 || !/^(?:0|[1-9][0-9]*)$/u.test(property)) {
    return false
  }
  const index = Number(property)
  return Number.isSafeInteger(index) && index >= 0 && index < 2 ** 32 - 1
}

function normalizeSliceIndex(index: number, length: number): number {
  const integer = Math.trunc(index)
  if (integer < 0) {
    return Math.max(length + integer, 0)
  }
  return Math.min(integer, length)
}

function encodeCellAddress(row: number, column: number): string {
  let value = column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(row + 1)}`
}
