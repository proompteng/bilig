import type { LiteralInput } from '@bilig/protocol'
import {
  filledUint32Array,
  growFloat64Array,
  growInt16Array,
  growInt32Array,
  growInt8Array,
  growUint8Array,
  growUint16Array,
  growUint32Array,
} from './xlsx-large-simple-array-storage.js'
import {
  binarySearchUint32,
  binarySearchUint32Prefix,
  canStoreInt8Number,
  canStoreInt16Number,
  canStoreInt32Number,
  canStoreLinearCoordinate,
  previewCellCount,
  previewIndex,
} from './xlsx-large-simple-arena-helpers.js'
import {
  initialCellCapacity,
  initialSparseIntegerCapacity,
  noPoolId,
  valueKindBoolean,
  valueKindEmpty,
  valueKindInteger,
  valueKindNull,
  valueKindNumber,
  valueKindSharedStringRef,
  valueKindSmallInteger,
  valueKindString,
  valueKindTinyInteger,
} from './xlsx-large-simple-arena-constants.js'
import type { ImportedWorkbookArenaDedupeMode, ImportedWorkbookArenaOptions } from './xlsx-large-simple-arena-types.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'

export abstract class ImportedWorkbookArenaBase {
  protected sheetIndex: number | null = null
  protected sheetIndexes: Uint32Array<ArrayBuffer> | undefined
  protected rows: Uint32Array<ArrayBuffer> = new Uint32Array(initialCellCapacity)
  protected columns: Uint16Array<ArrayBuffer> = new Uint16Array(initialCellCapacity)
  protected valueKinds: Uint8Array<ArrayBuffer> = new Uint8Array(initialCellCapacity)
  protected numberValues: Float64Array<ArrayBuffer> | undefined
  protected tinyIntegerValues: Int8Array<ArrayBuffer> | undefined
  protected smallIntegerValues: Int16Array<ArrayBuffer> | undefined
  protected sparseSmallIntegerCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected sparseSmallIntegerValues: Int16Array<ArrayBuffer> | undefined
  protected sparseSmallIntegerCount = 0
  protected integerValues: Int32Array<ArrayBuffer> | undefined
  protected sparseIntegerCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected sparseIntegerValues: Int32Array<ArrayBuffer> | undefined
  protected sparseIntegerCount = 0
  protected stringIds: Uint32Array<ArrayBuffer> | undefined
  protected sparseStringCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected sparseStringIds: Uint32Array<ArrayBuffer> | undefined
  protected booleanValues: Uint8Array<ArrayBuffer> | undefined
  protected formulaIds: Uint32Array<ArrayBuffer> | undefined
  protected length = 0
  protected denseRowMajorWidth: number | null = null
  protected linearCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected linearRowMajorWidth: number | null = null
  protected readonly strings: string[] = []
  protected readonly stringIdsByValue = new Map<string, number>()
  protected readonly formulas: string[] = []
  protected readonly formulaIdsByValue = new Map<string, number>()
  protected sharedStrings: LargeSimpleSharedStrings | undefined
  protected stringValueCount = 0
  protected sharedStringRefCount = 0
  protected sharedStringRefsInNumberValues = false
  protected readonly previewValues: (LiteralInput | undefined)[] = Array.from({ length: previewCellCount })
  protected readonly previewValueSet = new Uint8Array(previewCellCount)
  protected readonly stringDedupeMode: ImportedWorkbookArenaDedupeMode
  protected readonly formulaDedupeMode: ImportedWorkbookArenaDedupeMode
  protected readonly dedupeMaxEntries: number
  protected readonly stringDedupeKeys: string[] = []
  protected stringDedupeEvictionIndex = 0
  protected readonly formulaDedupeKeys: string[] = []
  protected formulaDedupeEvictionIndex = 0

  constructor(
    protected readonly stringPool?: ImportedWorkbookStringPool,
    options: ImportedWorkbookArenaOptions = {},
  ) {
    this.stringDedupeMode = options.deduplicateStrings ?? true
    this.formulaDedupeMode = options.deduplicateFormulas ?? true
    this.dedupeMaxEntries = Math.max(0, Math.trunc(options.dedupeMaxEntries ?? 8192))
  }

  get cellCount(): number {
    return this.length
  }

  retainedStorageByteLength(): number {
    return (
      (this.sheetIndexes?.byteLength ?? 0) +
      this.rows.byteLength +
      this.columns.byteLength +
      this.valueKinds.byteLength +
      (this.numberValues?.byteLength ?? 0) +
      (this.tinyIntegerValues?.byteLength ?? 0) +
      (this.smallIntegerValues?.byteLength ?? 0) +
      (this.sparseSmallIntegerCellIndexes?.byteLength ?? 0) +
      (this.sparseSmallIntegerValues?.byteLength ?? 0) +
      (this.integerValues?.byteLength ?? 0) +
      (this.sparseIntegerCellIndexes?.byteLength ?? 0) +
      (this.sparseIntegerValues?.byteLength ?? 0) +
      (this.stringIds?.byteLength ?? 0) +
      (this.sparseStringCellIndexes?.byteLength ?? 0) +
      (this.sparseStringIds?.byteLength ?? 0) +
      (this.booleanValues?.byteLength ?? 0) +
      (this.formulaIds?.byteLength ?? 0) +
      (this.linearCellIndexes?.byteLength ?? 0)
    )
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

  protected appendCell(sheetIndex: number, row: number, column: number): number {
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
      if (canStoreLinearCoordinate(this.denseRowMajorWidth, row, column)) {
        this.materializeLinearCoordinateStorage(index)
      } else {
        this.materializeCoordinateStorage(index)
      }
    }
    if (this.linearCellIndexes && this.linearRowMajorWidth !== null) {
      if (canStoreLinearCoordinate(this.linearRowMajorWidth, row, column)) {
        this.linearCellIndexes[index] = row * this.linearRowMajorWidth + column
        return index
      }
      this.materializeCoordinateStorage(index)
    }
    this.rows[index] = row
    this.columns[index] = column
    return index
  }

  protected recordSheetIndex(index: number, sheetIndex: number): void {
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

  protected addValue(index: number, value: LiteralInput | undefined): void {
    if (value === undefined) {
      this.valueKinds[index] = valueKindEmpty
      return
    }
    if (value === null) {
      this.valueKinds[index] = valueKindNull
      return
    }
    if (typeof value === 'number') {
      if (canStoreInt8Number(value)) {
        this.valueKinds[index] = valueKindTinyInteger
        this.ensureTinyIntegerValueStorage()[index] = value
        return
      }
      if (canStoreInt16Number(value)) {
        this.valueKinds[index] = valueKindSmallInteger
        this.storeSmallIntegerValue(index, value)
        return
      }
      if (canStoreInt32Number(value)) {
        this.valueKinds[index] = valueKindInteger
        this.storeIntegerValue(index, value)
        return
      }
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

  protected materializeValue(index: number): LiteralInput | undefined {
    const valueKind = this.valueKinds[index] ?? valueKindEmpty
    switch (valueKind) {
      case valueKindNumber:
        return this.numberValues?.[index]
      case valueKindTinyInteger:
        return this.tinyIntegerValues?.[index]
      case valueKindSmallInteger:
        return this.smallIntegerValueAt(index)
      case valueKindInteger:
        return this.integerValueAt(index)
      case valueKindString: {
        const stringId = this.stringIdAt(index)
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

  protected hasCellsForSheet(sheetIndex: number): boolean {
    return this.sheetIndexes !== undefined || this.sheetIndex === sheetIndex
  }

  protected cellBelongsToSheet(index: number, sheetIndex: number): boolean {
    return this.sheetIndexes ? this.sheetIndexes[index] === sheetIndex : this.sheetIndex === sheetIndex
  }

  protected ensureStringIdStorage(): Uint32Array<ArrayBuffer> {
    if (this.stringIds) {
      return this.stringIds
    }
    if (this.sparseStringCellIndexes && this.sparseStringIds) {
      const output = filledUint32Array(this.valueKinds.length, noPoolId)
      for (let index = 0; index < this.sparseStringCellIndexes.length; index += 1) {
        const cellIndex = this.sparseStringCellIndexes[index] ?? -1
        if (cellIndex >= 0 && cellIndex < output.length) {
          output[cellIndex] = this.sparseStringIds[index] ?? noPoolId
        }
      }
      this.sparseStringCellIndexes = undefined
      this.sparseStringIds = undefined
      this.stringIds = output
      return this.stringIds
    }
    this.stringIds = filledUint32Array(this.valueKinds.length, noPoolId)
    return this.stringIds
  }

  protected ensureNumberValueStorage(): Float64Array<ArrayBuffer> {
    if (this.numberValues) {
      return this.numberValues
    }
    this.numberValues = new Float64Array(this.valueKinds.length)
    this.numberValues.fill(Number.NaN)
    this.moveSharedStringIndexesToNumberValues()
    return this.numberValues
  }

  protected ensureTinyIntegerValueStorage(): Int8Array<ArrayBuffer> {
    if (this.tinyIntegerValues) {
      return this.tinyIntegerValues
    }
    this.tinyIntegerValues = new Int8Array(this.valueKinds.length)
    return this.tinyIntegerValues
  }

  protected storeSmallIntegerValue(index: number, value: number): void {
    if (this.smallIntegerValues) {
      this.smallIntegerValues[index] = value
      return
    }
    if (this.sparseSmallIntegerCount >= this.sparseIntegerDenseThreshold()) {
      this.ensureSmallIntegerValueStorage()[index] = value
      return
    }
    if (!this.sparseSmallIntegerCellIndexes || !this.sparseSmallIntegerValues) {
      this.sparseSmallIntegerCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseSmallIntegerValues = new Int16Array(initialSparseIntegerCapacity)
    } else if (this.sparseSmallIntegerCount >= this.sparseSmallIntegerCellIndexes.length) {
      const nextCapacity = this.sparseSmallIntegerCellIndexes.length * 2
      this.sparseSmallIntegerCellIndexes = growUint32Array(this.sparseSmallIntegerCellIndexes, nextCapacity)
      this.sparseSmallIntegerValues = growInt16Array(this.sparseSmallIntegerValues, nextCapacity)
    }
    this.sparseSmallIntegerCellIndexes[this.sparseSmallIntegerCount] = index
    this.sparseSmallIntegerValues[this.sparseSmallIntegerCount] = value
    this.sparseSmallIntegerCount += 1
  }

  protected ensureSmallIntegerValueStorage(): Int16Array<ArrayBuffer> {
    if (this.smallIntegerValues) {
      return this.smallIntegerValues
    }
    this.smallIntegerValues = new Int16Array(this.valueKinds.length)
    const sparseIndexes = this.sparseSmallIntegerCellIndexes
    const sparseValues = this.sparseSmallIntegerValues
    if (sparseIndexes && sparseValues) {
      for (let offset = 0; offset < this.sparseSmallIntegerCount; offset += 1) {
        const cellIndex = sparseIndexes[offset] ?? -1
        if (cellIndex >= 0 && cellIndex < this.smallIntegerValues.length) {
          this.smallIntegerValues[cellIndex] = sparseValues[offset] ?? 0
        }
      }
    }
    this.sparseSmallIntegerCellIndexes = undefined
    this.sparseSmallIntegerValues = undefined
    this.sparseSmallIntegerCount = 0
    return this.smallIntegerValues
  }

  protected smallIntegerValueAt(index: number): number | undefined {
    if (this.smallIntegerValues) {
      return this.smallIntegerValues[index]
    }
    const sparseIndexes = this.sparseSmallIntegerCellIndexes
    const sparseValues = this.sparseSmallIntegerValues
    if (!sparseIndexes || !sparseValues || this.sparseSmallIntegerCount === 0) {
      return undefined
    }
    const offset = binarySearchUint32Prefix(sparseIndexes, this.sparseSmallIntegerCount, index)
    return offset === -1 ? undefined : sparseValues[offset]
  }

  protected storeIntegerValue(index: number, value: number): void {
    if (this.integerValues) {
      this.integerValues[index] = value
      return
    }
    if (this.sparseIntegerCount >= this.sparseIntegerDenseThreshold()) {
      this.ensureIntegerValueStorage()[index] = value
      return
    }
    if (!this.sparseIntegerCellIndexes || !this.sparseIntegerValues) {
      this.sparseIntegerCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseIntegerValues = new Int32Array(initialSparseIntegerCapacity)
    } else if (this.sparseIntegerCount >= this.sparseIntegerCellIndexes.length) {
      const nextCapacity = this.sparseIntegerCellIndexes.length * 2
      this.sparseIntegerCellIndexes = growUint32Array(this.sparseIntegerCellIndexes, nextCapacity)
      this.sparseIntegerValues = growInt32Array(this.sparseIntegerValues, nextCapacity)
    }
    this.sparseIntegerCellIndexes[this.sparseIntegerCount] = index
    this.sparseIntegerValues[this.sparseIntegerCount] = value
    this.sparseIntegerCount += 1
  }

  protected ensureIntegerValueStorage(): Int32Array<ArrayBuffer> {
    if (this.integerValues) {
      return this.integerValues
    }
    this.integerValues = new Int32Array(this.valueKinds.length)
    const sparseIndexes = this.sparseIntegerCellIndexes
    const sparseValues = this.sparseIntegerValues
    if (sparseIndexes && sparseValues) {
      for (let offset = 0; offset < this.sparseIntegerCount; offset += 1) {
        const cellIndex = sparseIndexes[offset] ?? -1
        if (cellIndex >= 0 && cellIndex < this.integerValues.length) {
          this.integerValues[cellIndex] = sparseValues[offset] ?? 0
        }
      }
    }
    this.sparseIntegerCellIndexes = undefined
    this.sparseIntegerValues = undefined
    this.sparseIntegerCount = 0
    return this.integerValues
  }

  protected integerValueAt(index: number): number | undefined {
    if (this.integerValues) {
      return this.integerValues[index]
    }
    const sparseIndexes = this.sparseIntegerCellIndexes
    const sparseValues = this.sparseIntegerValues
    if (!sparseIndexes || !sparseValues || this.sparseIntegerCount === 0) {
      return undefined
    }
    const offset = binarySearchUint32Prefix(sparseIndexes, this.sparseIntegerCount, index)
    return offset === -1 ? undefined : sparseValues[offset]
  }

  protected sparseIntegerDenseThreshold(): number {
    return Math.max(initialSparseIntegerCapacity, this.valueKinds.length >>> 2)
  }

  protected ensureFormulaIdStorage(): Uint32Array<ArrayBuffer> {
    if (this.formulaIds) {
      return this.formulaIds
    }
    this.formulaIds = filledUint32Array(this.valueKinds.length, noPoolId)
    return this.formulaIds
  }

  protected ensureBooleanValueStorage(): Uint8Array<ArrayBuffer> {
    if (this.booleanValues) {
      return this.booleanValues
    }
    this.booleanValues = new Uint8Array(this.valueKinds.length)
    return this.booleanValues
  }

  protected storeSharedStringIndex(index: number, sharedStringIndex: number): void {
    this.sharedStringRefCount += 1
    if (this.numberValues) {
      this.sharedStringRefsInNumberValues = true
      this.numberValues[index] = sharedStringIndex
      return
    }
    this.ensureStringIdStorage()[index] = sharedStringIndex
  }

  protected sharedStringIndexAt(index: number): number {
    if (this.sharedStringRefsInNumberValues) {
      const value = this.numberValues?.[index]
      if (value !== undefined && !Number.isNaN(value)) {
        return Math.trunc(value)
      }
    }
    return this.stringIdAt(index)
  }

  protected moveSharedStringIndexesToNumberValues(): void {
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

  protected stringIdAt(index: number): number {
    if (this.stringIds) {
      return this.stringIds[index] ?? noPoolId
    }
    const sparseIndexes = this.sparseStringCellIndexes
    const sparseIds = this.sparseStringIds
    if (!sparseIndexes || !sparseIds) {
      return noPoolId
    }
    const offset = binarySearchUint32(sparseIndexes, index)
    return offset === -1 ? noPoolId : (sparseIds[offset] ?? noPoolId)
  }

  protected internString(value: string): number {
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

  protected internFormula(value: string): number {
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

  protected internValue(value: string, mode: ImportedWorkbookArenaDedupeMode): string | null {
    if (mode === false) {
      return null
    }
    if (mode === 'bounded') {
      return this.stringPool?.internBounded(value, this.dedupeMaxEntries) ?? value
    }
    return this.stringPool?.intern(value) ?? value
  }

  protected rememberBoundedDedupeKey(map: Map<string, number>, keys: string[], kind: 'string' | 'formula', key: string): void {
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

  protected rowAt(index: number): number {
    if (this.denseRowMajorWidth !== null) {
      return Math.floor(index / this.denseRowMajorWidth)
    }
    if (this.linearCellIndexes && this.linearRowMajorWidth !== null) {
      return Math.floor((this.linearCellIndexes[index] ?? 0) / this.linearRowMajorWidth)
    }
    return this.rows[index] ?? 0
  }

  protected columnAt(index: number): number {
    if (this.denseRowMajorWidth !== null) {
      return index % this.denseRowMajorWidth
    }
    if (this.linearCellIndexes && this.linearRowMajorWidth !== null) {
      return (this.linearCellIndexes[index] ?? 0) % this.linearRowMajorWidth
    }
    return this.columns[index] ?? 0
  }

  protected materializeCoordinateStorage(upToIndex: number): void {
    const width = this.denseRowMajorWidth ?? this.linearRowMajorWidth
    if (width === null) {
      return
    }
    this.rows = new Uint32Array(this.valueKinds.length)
    this.columns = new Uint16Array(this.valueKinds.length)
    const linearCellIndexes = this.linearCellIndexes
    if (linearCellIndexes) {
      for (let index = 0; index < upToIndex; index += 1) {
        const linearCellIndex = linearCellIndexes[index] ?? 0
        this.rows[index] = Math.floor(linearCellIndex / width)
        this.columns[index] = linearCellIndex % width
      }
    } else {
      for (let index = 0; index < upToIndex; index += 1) {
        this.rows[index] = Math.floor(index / width)
        this.columns[index] = index % width
      }
    }
    this.denseRowMajorWidth = null
    this.linearCellIndexes = undefined
    this.linearRowMajorWidth = null
  }

  protected materializeLinearCoordinateStorage(upToIndex: number): void {
    const width = this.denseRowMajorWidth
    if (width === null) {
      return
    }
    this.linearCellIndexes = new Uint32Array(this.valueKinds.length)
    for (let index = 0; index < upToIndex; index += 1) {
      this.linearCellIndexes[index] = index
    }
    this.linearRowMajorWidth = width
    this.denseRowMajorWidth = null
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
  }

  protected setPreviewValue(row: number, column: number, value: LiteralInput): void {
    const index = previewIndex(row, column)
    if (index === -1) {
      return
    }
    this.previewValues[index] = value
    this.previewValueSet[index] = 1
  }

  protected hasPreviewValue(row: number, column: number): boolean {
    const index = previewIndex(row, column)
    return index !== -1 && this.previewValueSet[index] === 1
  }

  protected readPreviewValue(row: number, column: number): LiteralInput | undefined {
    const index = previewIndex(row, column)
    return index === -1 || this.previewValueSet[index] !== 1 ? undefined : this.previewValues[index]
  }

  protected ensureCapacity(nextLength: number): void {
    if (nextLength <= this.valueKinds.length) {
      return
    }
    let nextCapacity = this.valueKinds.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.resizeStorage(nextCapacity)
  }

  protected resizeStorage(nextCapacity: number): void {
    if (this.sparseStringCellIndexes || this.sparseStringIds) {
      this.ensureStringIdStorage()
    }
    if (this.sheetIndexes) {
      this.sheetIndexes = growUint32Array(this.sheetIndexes, nextCapacity)
    }
    if (this.linearCellIndexes) {
      this.linearCellIndexes = growUint32Array(this.linearCellIndexes, nextCapacity)
    } else if (this.denseRowMajorWidth === null) {
      this.rows = growUint32Array(this.rows, nextCapacity)
      this.columns = growUint16Array(this.columns, nextCapacity)
    }
    this.valueKinds = growUint8Array(this.valueKinds, nextCapacity)
    if (this.numberValues) {
      this.numberValues = growFloat64Array(this.numberValues, nextCapacity)
    }
    if (this.tinyIntegerValues) {
      this.tinyIntegerValues = growInt8Array(this.tinyIntegerValues, nextCapacity)
    }
    if (this.smallIntegerValues) {
      this.smallIntegerValues = growInt16Array(this.smallIntegerValues, nextCapacity)
    }
    if (this.integerValues) {
      this.integerValues = growInt32Array(this.integerValues, nextCapacity)
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

  protected compactSparseStringIds(): void {
    const denseStringIds = this.stringIds
    if (!denseStringIds || this.length === 0) {
      return
    }
    let retainedCount = 0
    for (let index = 0; index < this.length; index += 1) {
      if ((denseStringIds[index] ?? noPoolId) !== noPoolId) {
        retainedCount += 1
      }
    }
    if (retainedCount === 0) {
      this.stringIds = undefined
      this.sparseStringCellIndexes = undefined
      this.sparseStringIds = undefined
      return
    }
    if (retainedCount * 2 >= this.length) {
      return
    }
    const indexes = new Uint32Array(retainedCount)
    const ids = new Uint32Array(retainedCount)
    let outputIndex = 0
    for (let index = 0; index < this.length; index += 1) {
      const stringId = denseStringIds[index] ?? noPoolId
      if (stringId === noPoolId) {
        continue
      }
      indexes[outputIndex] = index
      ids[outputIndex] = stringId
      outputIndex += 1
    }
    this.stringIds = undefined
    this.sparseStringCellIndexes = indexes
    this.sparseStringIds = ids
  }

  protected snapshotStringIds(): Uint32Array | undefined {
    if (this.stringIds) {
      return this.stringIds.subarray(0, this.length)
    }
    const sparseIndexes = this.sparseStringCellIndexes
    const sparseIds = this.sparseStringIds
    if (!sparseIndexes || !sparseIds || sparseIndexes.length === 0) {
      return undefined
    }
    const output = filledUint32Array(this.length, noPoolId)
    for (let index = 0; index < sparseIndexes.length; index += 1) {
      const cellIndex = sparseIndexes[index] ?? -1
      if (cellIndex >= 0 && cellIndex < output.length) {
        output[cellIndex] = sparseIds[index] ?? noPoolId
      }
    }
    return output
  }
}
