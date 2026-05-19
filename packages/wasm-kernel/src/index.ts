import type * as Fs from 'node:fs'
import type * as FsPromises from 'node:fs/promises'
import type * as NodeUrl from 'node:url'
import { isRawKernelExports, type RawKernelExports } from './raw-kernel-exports.js'

type TypedArrayValue = Uint8Array | Uint16Array | Uint32Array | Float64Array

const ARRAY_BUFFER_CLASS_ID = 1
const UINT8_ARRAY_CLASS_ID = 4
const FLOAT64_ARRAY_CLASS_ID = 5
const UINT16_ARRAY_CLASS_ID = 6
const UINT32_ARRAY_CLASS_ID = 7

interface LoweredArraySpec<T extends TypedArrayValue> {
  align: number
  classId: number
  ctor: {
    new (buffer: ArrayBufferLike, byteOffset: number, length: number): T
  }
}

interface EnsureWasmBinaryPathForNodeOptions {
  readonly importMetaUrl: string
  readonly existsSync: (path: string) => boolean
  readonly fileURLToPath: (url: URL) => string
}

export function ensureWasmBinaryPathForNode(options: EnsureWasmBinaryPathForNodeOptions): string {
  const wasmPath = options.fileURLToPath(new URL('../build/release.wasm', options.importMetaUrl))
  if (options.existsSync(wasmPath)) {
    return wasmPath
  }
  throw new Error(`Unable to locate wasm kernel binary at '${wasmPath}'. Run 'pnpm wasm:build' before using @bilig/wasm-kernel.`)
}

export interface SpreadsheetKernel {
  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void
  ensureCellCapacity(nextCapacity: number): void
  ensureFormulaCapacity(nextCapacity: number): void
  ensureConstantCapacity(nextCapacity: number): void
  ensureRangeCapacity(nextCapacity: number): void
  ensureMemberCapacity(nextCapacity: number): void
  uploadPrograms(programs: Uint32Array, offsets: Uint32Array, lengths: Uint32Array, targets: Uint32Array): void
  uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void
  uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void
  uploadRangeShapes(rowCounts: Uint32Array, colCounts: Uint32Array): void
  uploadVolatileNowSerial(nowSerial: number): void
  uploadVolatileRandomValues(values: Float64Array): void
  uploadStringLengths(lengths: Uint32Array): void
  uploadStrings(offsets: Uint32Array, lengths: Uint32Array, data: Uint16Array): void
  writeCells(tags: Uint8Array, numbers: Float64Array, stringIds: Uint32Array, errors: Uint16Array): void
  evalDirectScalarValueBatch(
    operators: Uint8Array,
    leftBatchRefs: Uint32Array,
    leftTags: Uint8Array,
    leftValues: Float64Array,
    leftErrors: Uint16Array,
    rightBatchRefs: Uint32Array,
    rightTags: Uint8Array,
    rightValues: Float64Array,
    rightErrors: Uint16Array,
    resultOffsets: Float64Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void
  evalDenseNumericRowAggregateBatch(
    aggregateKind: number,
    values: Float64Array,
    rowCount: number,
    prefixColCount: number,
    startColOffset: number,
    aggregateColCount: number,
    resultOffset: number,
    outNumbers: Float64Array,
  ): void
  evalBatch(cellIndices: Uint32Array): void
  materializePivotTable(
    sourceRangeIndex: number,
    sourceWidth: number,
    groupByColumnIndices: Uint32Array,
    valueColumnIndices: Uint32Array,
    valueAggregations: Uint8Array,
  ): {
    rows: number
    cols: number
    tags: Uint8Array
    numbers: Float64Array
    stringIds: Uint32Array
    errors: Uint16Array
  }
  readTags(): Uint8Array
  readNumbers(): Float64Array
  readStringIds(): Uint32Array
  readErrors(): Uint16Array
  readProgramOffsets(): Uint32Array
  readProgramLengths(): Uint32Array
  readConstantOffsets(): Uint32Array
  readConstantLengths(): Uint32Array
  readConstants(): Float64Array
  readRangeOffsets(): Uint32Array
  readRangeLengths(): Uint32Array
  readRangeMembers(): Uint32Array
  readOutputStrings(): string[]
  readSpillRows(): Uint32Array
  readSpillCols(): Uint32Array
  readSpillOffsets(): Uint32Array
  readSpillLengths(): Uint32Array
  readSpillTags(): Uint8Array
  readSpillNumbers(): Float64Array
  getSpillValueCount(): number
  getCellCapacity(): number
  getFormulaCapacity(): number
  getConstantCapacity(): number
  getRangeCapacity(): number
  getMemberCapacity(): number
}

const uint8Spec: LoweredArraySpec<Uint8Array> = {
  align: 0,
  classId: UINT8_ARRAY_CLASS_ID,
  ctor: Uint8Array,
}

const uint16Spec: LoweredArraySpec<Uint16Array> = {
  align: 1,
  classId: UINT16_ARRAY_CLASS_ID,
  ctor: Uint16Array,
}

const uint32Spec: LoweredArraySpec<Uint32Array> = {
  align: 2,
  classId: UINT32_ARRAY_CLASS_ID,
  ctor: Uint32Array,
}

const float64Spec: LoweredArraySpec<Float64Array> = {
  align: 3,
  classId: FLOAT64_ARRAY_CLASS_ID,
  ctor: Float64Array,
}

class RawKernelBridge {
  private dataView: DataView

  constructor(private readonly raw: RawKernelExports) {
    this.dataView = new DataView(raw.memory.buffer)
  }

  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void {
    this.raw.init(cellCapacity, formulaCapacity, constantCapacity, rangeCapacity, memberCapacity)
  }

  ensureCellCapacity(nextCapacity: number): void {
    this.raw.ensureCellCapacity(nextCapacity)
  }

  ensureFormulaCapacity(nextCapacity: number): void {
    this.raw.ensureFormulaCapacity(nextCapacity)
  }

  ensureConstantCapacity(nextCapacity: number): void {
    this.raw.ensureConstantCapacity(nextCapacity)
  }

  ensureRangeCapacity(nextCapacity: number): void {
    this.raw.ensureRangeCapacity(nextCapacity)
  }

  ensureMemberCapacity(nextCapacity: number): void {
    this.raw.ensureMemberCapacity(nextCapacity)
  }

  uploadPrograms(programs: Uint32Array, offsets: Uint32Array, lengths: Uint32Array, targets: Uint32Array): void {
    const programsPtr = this.lowerTypedArray(programs, uint32Spec)
    const offsetsPtr = this.lowerTypedArray(offsets, uint32Spec)
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec)
    const targetsPtr = this.lowerTypedArray(targets, uint32Spec)
    try {
      this.raw.uploadPrograms(programsPtr, offsetsPtr, lengthsPtr, targetsPtr)
    } finally {
      this.raw.__unpin(programsPtr)
      this.raw.__unpin(offsetsPtr)
      this.raw.__unpin(lengthsPtr)
      this.raw.__unpin(targetsPtr)
    }
  }

  uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void {
    const constantsPtr = this.lowerTypedArray(constants, float64Spec)
    const offsetsPtr = this.lowerTypedArray(offsets, uint32Spec)
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec)
    try {
      this.raw.uploadConstants(constantsPtr, offsetsPtr, lengthsPtr)
    } finally {
      this.raw.__unpin(constantsPtr)
      this.raw.__unpin(offsetsPtr)
      this.raw.__unpin(lengthsPtr)
    }
  }

  uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void {
    const membersPtr = this.lowerTypedArray(members, uint32Spec)
    const offsetsPtr = this.lowerTypedArray(offsets, uint32Spec)
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec)
    try {
      this.raw.uploadRangeMembers(membersPtr, offsetsPtr, lengthsPtr)
    } finally {
      this.raw.__unpin(membersPtr)
      this.raw.__unpin(offsetsPtr)
      this.raw.__unpin(lengthsPtr)
    }
  }

  uploadRangeShapes(rowCounts: Uint32Array, colCounts: Uint32Array): void {
    const rowCountsPtr = this.lowerTypedArray(rowCounts, uint32Spec)
    const colCountsPtr = this.lowerTypedArray(colCounts, uint32Spec)
    try {
      this.raw.uploadRangeShapes(rowCountsPtr, colCountsPtr)
    } finally {
      this.raw.__unpin(rowCountsPtr)
      this.raw.__unpin(colCountsPtr)
    }
  }

  uploadVolatileNowSerial(nowSerial: number): void {
    this.raw.uploadVolatileNowSerial(nowSerial)
  }

  uploadVolatileRandomValues(values: Float64Array): void {
    const valuesPtr = this.lowerTypedArray(values, float64Spec)
    try {
      this.raw.uploadVolatileRandomValues(valuesPtr)
    } finally {
      this.raw.__unpin(valuesPtr)
    }
  }

  uploadStringLengths(lengths: Uint32Array): void {
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec)
    try {
      this.raw.uploadStringLengths(lengthsPtr)
    } finally {
      this.raw.__unpin(lengthsPtr)
    }
  }

  uploadStrings(offsets: Uint32Array, lengths: Uint32Array, data: Uint16Array): void {
    const offsetsPtr = this.lowerTypedArray(offsets, uint32Spec)
    const lengthsPtr = this.lowerTypedArray(lengths, uint32Spec)
    const dataPtr = this.lowerTypedArray(data, uint16Spec)
    try {
      this.raw.uploadStrings(offsetsPtr, lengthsPtr, dataPtr)
    } finally {
      this.raw.__unpin(offsetsPtr)
      this.raw.__unpin(lengthsPtr)
      this.raw.__unpin(dataPtr)
    }
  }

  writeCells(tags: Uint8Array, numbers: Float64Array, stringIds: Uint32Array, errors: Uint16Array): void {
    const tagsPtr = this.lowerTypedArray(tags, uint8Spec)
    const numbersPtr = this.lowerTypedArray(numbers, float64Spec)
    const stringIdsPtr = this.lowerTypedArray(stringIds, uint32Spec)
    const errorsPtr = this.lowerTypedArray(errors, uint16Spec)
    try {
      this.raw.writeCells(tagsPtr, numbersPtr, stringIdsPtr, errorsPtr)
    } finally {
      this.raw.__unpin(tagsPtr)
      this.raw.__unpin(numbersPtr)
      this.raw.__unpin(stringIdsPtr)
      this.raw.__unpin(errorsPtr)
    }
  }

  evalDirectScalarValueBatch(
    operators: Uint8Array,
    leftBatchRefs: Uint32Array,
    leftTags: Uint8Array,
    leftValues: Float64Array,
    leftErrors: Uint16Array,
    rightBatchRefs: Uint32Array,
    rightTags: Uint8Array,
    rightValues: Float64Array,
    rightErrors: Uint16Array,
    resultOffsets: Float64Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    const operatorsPtr = this.lowerTypedArray(operators, uint8Spec)
    const leftBatchRefsPtr = this.lowerTypedArray(leftBatchRefs, uint32Spec)
    const leftTagsPtr = this.lowerTypedArray(leftTags, uint8Spec)
    const leftValuesPtr = this.lowerTypedArray(leftValues, float64Spec)
    const leftErrorsPtr = this.lowerTypedArray(leftErrors, uint16Spec)
    const rightBatchRefsPtr = this.lowerTypedArray(rightBatchRefs, uint32Spec)
    const rightTagsPtr = this.lowerTypedArray(rightTags, uint8Spec)
    const rightValuesPtr = this.lowerTypedArray(rightValues, float64Spec)
    const rightErrorsPtr = this.lowerTypedArray(rightErrors, uint16Spec)
    const resultOffsetsPtr = this.lowerTypedArray(resultOffsets, float64Spec)
    const outTagsPtr = this.lowerTypedArray(outTags, uint8Spec)
    const outNumbersPtr = this.lowerTypedArray(outNumbers, float64Spec)
    const outErrorsPtr = this.lowerTypedArray(outErrors, uint16Spec)
    try {
      this.raw.evalDirectScalarValueBatch(
        operatorsPtr,
        leftBatchRefsPtr,
        leftTagsPtr,
        leftValuesPtr,
        leftErrorsPtr,
        rightBatchRefsPtr,
        rightTagsPtr,
        rightValuesPtr,
        rightErrorsPtr,
        resultOffsetsPtr,
        outTagsPtr,
        outNumbersPtr,
        outErrorsPtr,
      )
      this.copyLoweredTypedArray(outTagsPtr, outTags, uint8Spec)
      this.copyLoweredTypedArray(outNumbersPtr, outNumbers, float64Spec)
      this.copyLoweredTypedArray(outErrorsPtr, outErrors, uint16Spec)
    } finally {
      this.raw.__unpin(operatorsPtr)
      this.raw.__unpin(leftBatchRefsPtr)
      this.raw.__unpin(leftTagsPtr)
      this.raw.__unpin(leftValuesPtr)
      this.raw.__unpin(leftErrorsPtr)
      this.raw.__unpin(rightBatchRefsPtr)
      this.raw.__unpin(rightTagsPtr)
      this.raw.__unpin(rightValuesPtr)
      this.raw.__unpin(rightErrorsPtr)
      this.raw.__unpin(resultOffsetsPtr)
      this.raw.__unpin(outTagsPtr)
      this.raw.__unpin(outNumbersPtr)
      this.raw.__unpin(outErrorsPtr)
    }
  }

  evalDenseNumericRowAggregateBatch(
    aggregateKind: number,
    values: Float64Array,
    rowCount: number,
    prefixColCount: number,
    startColOffset: number,
    aggregateColCount: number,
    resultOffset: number,
    outNumbers: Float64Array,
  ): void {
    const valuesPtr = this.lowerTypedArray(values, float64Spec)
    const outNumbersPtr = this.lowerTypedArray(outNumbers, float64Spec)
    try {
      this.raw.evalDenseNumericRowAggregateBatch(
        aggregateKind,
        valuesPtr,
        rowCount,
        prefixColCount,
        startColOffset,
        aggregateColCount,
        resultOffset,
        outNumbersPtr,
      )
      this.copyLoweredTypedArray(outNumbersPtr, outNumbers, float64Spec)
    } finally {
      this.raw.__unpin(valuesPtr)
      this.raw.__unpin(outNumbersPtr)
    }
  }

  evalBatch(cellIndices: Uint32Array): void {
    const cellIndicesPtr = this.lowerTypedArray(cellIndices, uint32Spec)
    try {
      this.raw.evalBatch(cellIndicesPtr)
    } finally {
      this.raw.__unpin(cellIndicesPtr)
    }
  }

  materializePivotTable(
    sourceRangeIndex: number,
    sourceWidth: number,
    groupByColumnIndices: Uint32Array,
    valueColumnIndices: Uint32Array,
    valueAggregations: Uint8Array,
  ): void {
    const groupByPtr = this.lowerTypedArray(groupByColumnIndices, uint32Spec)
    const valueColsPtr = this.lowerTypedArray(valueColumnIndices, uint32Spec)
    const valueAggsPtr = this.lowerTypedArray(valueAggregations, uint8Spec)
    try {
      this.raw.materializePivotTable(
        sourceRangeIndex,
        sourceWidth,
        groupByColumnIndices.length,
        groupByPtr,
        valueColumnIndices.length,
        valueColsPtr,
        valueAggsPtr,
      )
    } finally {
      this.raw.__unpin(groupByPtr)
      this.raw.__unpin(valueColsPtr)
      this.raw.__unpin(valueAggsPtr)
    }
  }

  private lowerTypedArray<T extends TypedArrayValue>(values: T, spec: LoweredArraySpec<T>): number {
    const byteLength = values.length << spec.align
    const bufferPtr = this.raw.__pin(this.raw.__new(byteLength, ARRAY_BUFFER_CLASS_ID))
    const headerPtr = this.raw.__pin(this.raw.__new(12, spec.classId))
    try {
      this.setUint32(headerPtr, bufferPtr)
      this.setUint32(headerPtr + 4, bufferPtr)
      this.setUint32(headerPtr + 8, byteLength)
      new spec.ctor(this.raw.memory.buffer, bufferPtr, values.length).set(values)
      return headerPtr
    } finally {
      this.raw.__unpin(bufferPtr)
    }
  }

  private copyLoweredTypedArray<T extends TypedArrayValue>(pointer: number, target: T, spec: LoweredArraySpec<T>): void {
    target.set(new spec.ctor(this.raw.memory.buffer, this.getUint32(pointer + 4), target.length))
  }

  private setUint32(pointer: number, value: number): void {
    try {
      this.dataView.setUint32(pointer, value, true)
    } catch {
      this.dataView = new DataView(this.raw.memory.buffer)
      this.dataView.setUint32(pointer, value, true)
    }
  }

  private getUint32(pointer: number): number {
    try {
      return this.dataView.getUint32(pointer, true)
    } catch {
      this.dataView = new DataView(this.raw.memory.buffer)
      return this.dataView.getUint32(pointer, true)
    }
  }
}

class KernelHandle implements SpreadsheetKernel {
  private readonly bridge: RawKernelBridge
  private tags = new Uint8Array()
  private numbers = new Float64Array()
  private stringIds = new Uint32Array()
  private errors = new Uint16Array()
  private programOffsets = new Uint32Array()
  private programLengths = new Uint32Array()
  private constantOffsets = new Uint32Array()
  private constantLengths = new Uint32Array()
  private constants = new Float64Array()
  private rangeOffsets = new Uint32Array()
  private rangeLengths = new Uint32Array()
  private rangeMembers = new Uint32Array()
  private spillRows = new Uint32Array()
  private spillCols = new Uint32Array()
  private spillOffsets = new Uint32Array()
  private spillLengths = new Uint32Array()
  private spillTags = new Uint8Array()
  private spillNumbers = new Float64Array()

  constructor(private readonly raw: RawKernelExports) {
    this.bridge = new RawKernelBridge(raw)
    this.refreshViews()
  }

  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void {
    this.bridge.init(cellCapacity, formulaCapacity, constantCapacity, rangeCapacity, memberCapacity)
    this.refreshViews()
  }

  ensureCellCapacity(nextCapacity: number): void {
    this.bridge.ensureCellCapacity(nextCapacity)
    this.refreshViews()
  }

  ensureFormulaCapacity(nextCapacity: number): void {
    this.bridge.ensureFormulaCapacity(nextCapacity)
    this.refreshViews()
  }

  ensureConstantCapacity(nextCapacity: number): void {
    this.bridge.ensureConstantCapacity(nextCapacity)
    this.refreshViews()
  }

  ensureRangeCapacity(nextCapacity: number): void {
    this.bridge.ensureRangeCapacity(nextCapacity)
    this.refreshViews()
  }

  ensureMemberCapacity(nextCapacity: number): void {
    this.bridge.ensureMemberCapacity(nextCapacity)
    this.refreshViews()
  }

  uploadPrograms(programs: Uint32Array, offsets: Uint32Array, lengths: Uint32Array, targets: Uint32Array): void {
    this.bridge.uploadPrograms(programs, offsets, lengths, targets)
    this.refreshViews()
  }

  uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void {
    this.bridge.uploadConstants(constants, offsets, lengths)
    this.refreshViews()
  }

  uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void {
    this.bridge.uploadRangeMembers(members, offsets, lengths)
    this.refreshViews()
  }

  uploadRangeShapes(rowCounts: Uint32Array, colCounts: Uint32Array): void {
    this.bridge.uploadRangeShapes(rowCounts, colCounts)
    this.refreshViews()
  }

  uploadVolatileNowSerial(nowSerial: number): void {
    this.bridge.uploadVolatileNowSerial(nowSerial)
  }

  uploadVolatileRandomValues(values: Float64Array): void {
    this.bridge.uploadVolatileRandomValues(values)
  }

  uploadStringLengths(lengths: Uint32Array): void {
    this.bridge.uploadStringLengths(lengths)
    this.refreshViews()
  }

  uploadStrings(offsets: Uint32Array, lengths: Uint32Array, data: Uint16Array): void {
    this.bridge.uploadStrings(offsets, lengths, data)
    this.refreshViews()
  }

  writeCells(tags: Uint8Array, numbers: Float64Array, stringIds: Uint32Array, errors: Uint16Array): void {
    this.bridge.writeCells(tags, numbers, stringIds, errors)
    this.refreshViews()
  }

  evalDirectScalarValueBatch(
    operators: Uint8Array,
    leftBatchRefs: Uint32Array,
    leftTags: Uint8Array,
    leftValues: Float64Array,
    leftErrors: Uint16Array,
    rightBatchRefs: Uint32Array,
    rightTags: Uint8Array,
    rightValues: Float64Array,
    rightErrors: Uint16Array,
    resultOffsets: Float64Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    this.bridge.evalDirectScalarValueBatch(
      operators,
      leftBatchRefs,
      leftTags,
      leftValues,
      leftErrors,
      rightBatchRefs,
      rightTags,
      rightValues,
      rightErrors,
      resultOffsets,
      outTags,
      outNumbers,
      outErrors,
    )
  }

  evalDenseNumericRowAggregateBatch(
    aggregateKind: number,
    values: Float64Array,
    rowCount: number,
    prefixColCount: number,
    startColOffset: number,
    aggregateColCount: number,
    resultOffset: number,
    outNumbers: Float64Array,
  ): void {
    this.bridge.evalDenseNumericRowAggregateBatch(
      aggregateKind,
      values,
      rowCount,
      prefixColCount,
      startColOffset,
      aggregateColCount,
      resultOffset,
      outNumbers,
    )
  }

  evalBatch(cellIndices: Uint32Array): void {
    this.bridge.evalBatch(cellIndices)
    this.refreshViews()
  }

  materializePivotTable(
    sourceRangeIndex: number,
    sourceWidth: number,
    groupByColumnIndices: Uint32Array,
    valueColumnIndices: Uint32Array,
    valueAggregations: Uint8Array,
  ): {
    rows: number
    cols: number
    tags: Uint8Array
    numbers: Float64Array
    stringIds: Uint32Array
    errors: Uint16Array
  } {
    this.bridge.materializePivotTable(sourceRangeIndex, sourceWidth, groupByColumnIndices, valueColumnIndices, valueAggregations)
    const rows = this.raw.pivotResultRows.value
    const cols = this.raw.pivotResultCols.value
    const size = rows * cols
    const memory = this.raw.memory.buffer
    return {
      rows,
      cols,
      tags: new Uint8Array(memory, this.raw.getPivotResultTagsPtr(), size),
      numbers: new Float64Array(memory, this.raw.getPivotResultNumbersPtr(), size),
      stringIds: new Uint32Array(memory, this.raw.getPivotResultStringIdsPtr(), size),
      errors: new Uint16Array(memory, this.raw.getPivotResultErrorsPtr(), size),
    }
  }

  readTags(): Uint8Array {
    return this.tags
  }

  readNumbers(): Float64Array {
    return this.numbers
  }

  readStringIds(): Uint32Array {
    return this.stringIds
  }

  readErrors(): Uint16Array {
    return this.errors
  }

  readProgramOffsets(): Uint32Array {
    return this.programOffsets
  }

  readProgramLengths(): Uint32Array {
    return this.programLengths
  }

  readConstantOffsets(): Uint32Array {
    return this.constantOffsets
  }

  readConstantLengths(): Uint32Array {
    return this.constantLengths
  }

  readConstants(): Float64Array {
    return this.constants
  }

  readRangeOffsets(): Uint32Array {
    return this.rangeOffsets
  }

  readRangeLengths(): Uint32Array {
    return this.rangeLengths
  }

  readRangeMembers(): Uint32Array {
    return this.rangeMembers
  }

  readOutputStrings(): string[] {
    const memory = this.raw.memory.buffer
    const count = this.raw.getOutputStringCount()
    if (count === 0) return []

    const lengths = new Uint32Array(memory, this.raw.getOutputStringLengthsPtr(), count)
    const offsets = new Uint32Array(memory, this.raw.getOutputStringOffsetsPtr(), count)
    const data = new Uint16Array(memory, this.raw.getOutputStringDataPtr(), this.raw.getOutputStringDataLength())

    const strings: string[] = []
    for (let index = 0; index < count; index += 1) {
      const length = lengths[index]!
      const offset = offsets[index]!
      let str = ''
      for (let charIndex = 0; charIndex < length; charIndex += 1) {
        str += String.fromCharCode(data[offset + charIndex]!)
      }
      strings.push(str)
    }
    return strings
  }

  readSpillRows(): Uint32Array {
    return this.spillRows
  }

  readSpillCols(): Uint32Array {
    return this.spillCols
  }

  readSpillOffsets(): Uint32Array {
    return this.spillOffsets
  }

  readSpillLengths(): Uint32Array {
    return this.spillLengths
  }

  readSpillTags(): Uint8Array {
    return this.spillTags
  }

  readSpillNumbers(): Float64Array {
    return this.spillNumbers
  }

  getSpillValueCount(): number {
    return this.raw.getSpillResultValueCount()
  }

  getCellCapacity(): number {
    return this.raw.getCellCapacity()
  }

  getFormulaCapacity(): number {
    return this.raw.getFormulaCapacity()
  }

  getConstantCapacity(): number {
    return this.raw.getConstantCapacity()
  }

  getRangeCapacity(): number {
    return this.raw.getRangeCapacity()
  }

  getMemberCapacity(): number {
    return this.raw.getMemberCapacity()
  }

  private refreshViews(): void {
    const memory = this.raw.memory.buffer
    this.tags = new Uint8Array(memory, this.raw.getTagsPtr(), this.raw.getCellCapacity())
    this.numbers = new Float64Array(memory, this.raw.getNumbersPtr(), this.raw.getCellCapacity())
    this.stringIds = new Uint32Array(memory, this.raw.getStringIdsPtr(), this.raw.getCellCapacity())
    this.errors = new Uint16Array(memory, this.raw.getErrorsPtr(), this.raw.getCellCapacity())
    this.programOffsets = new Uint32Array(memory, this.raw.getProgramOffsetsPtr(), this.raw.getFormulaCapacity())
    this.programLengths = new Uint32Array(memory, this.raw.getProgramLengthsPtr(), this.raw.getFormulaCapacity())
    this.constantOffsets = new Uint32Array(memory, this.raw.getConstantOffsetsPtr(), this.raw.getFormulaCapacity())
    this.constantLengths = new Uint32Array(memory, this.raw.getConstantLengthsPtr(), this.raw.getFormulaCapacity())
    this.constants = new Float64Array(memory, this.raw.getConstantArenaPtr(), this.raw.getConstantCapacity())
    this.rangeOffsets = new Uint32Array(memory, this.raw.getRangeOffsetsPtr(), this.raw.getRangeCapacity())
    this.rangeLengths = new Uint32Array(memory, this.raw.getRangeLengthsPtr(), this.raw.getRangeCapacity())
    this.rangeMembers = new Uint32Array(memory, this.raw.getRangeMembersPtr(), this.raw.getMemberCapacity())
    this.spillRows = new Uint32Array(memory, this.raw.getSpillResultRowsPtr(), this.raw.getCellCapacity())
    this.spillCols = new Uint32Array(memory, this.raw.getSpillResultColsPtr(), this.raw.getCellCapacity())
    this.spillOffsets = new Uint32Array(memory, this.raw.getSpillResultOffsetsPtr(), this.raw.getCellCapacity())
    this.spillLengths = new Uint32Array(memory, this.raw.getSpillResultLengthsPtr(), this.raw.getCellCapacity())
    this.spillTags = new Uint8Array(memory, this.raw.getSpillResultTagsPtr(), this.raw.getSpillResultValueCount())
    this.spillNumbers = new Float64Array(memory, this.raw.getSpillResultNumbersPtr(), this.raw.getSpillResultValueCount())
  }
}

function isNodeLike(): boolean {
  return typeof process !== 'undefined' && process.versions != null && process.versions.node != null
}

async function loadWasmModule(): Promise<WebAssembly.WebAssemblyInstantiatedSource> {
  const imports = {
    env: {
      abort(_message: number, _fileName: number, lineNumber: number, columnNumber: number) {
        throw new Error(`AssemblyScript abort at ${lineNumber}:${columnNumber}`)
      },
    },
  }

  if (isNodeLike()) {
    const fsPromises = process.getBuiltinModule('fs/promises') as typeof FsPromises | undefined
    const fs = process.getBuiltinModule('fs') as typeof Fs | undefined
    const url = process.getBuiltinModule('node:url') as typeof NodeUrl | undefined
    if (!fsPromises) {
      throw new Error('Node fs/promises module is unavailable')
    }
    if (!fs) {
      throw new Error('Node fs module is unavailable')
    }
    if (!url) {
      throw new Error('Node url module is unavailable')
    }
    const wasmPath = ensureWasmBinaryPathForNode({
      importMetaUrl: import.meta.url,
      existsSync: fs.existsSync,
      fileURLToPath: url.fileURLToPath,
    })
    const { readFile } = fsPromises
    const bytes = await readFile(wasmPath)
    return WebAssembly.instantiate(bytes, imports)
  }

  const wasmUrl = new URL('../build/release.wasm', import.meta.url)
  const response = await fetch(wasmUrl)
  if (!response.ok) {
    throw new Error(`Failed to load wasm kernel: ${response.status} ${response.statusText}`)
  }
  const bytes = await response.arrayBuffer()
  return WebAssembly.instantiate(bytes, imports)
}

function loadWasmModuleSync(): WebAssembly.WebAssemblyInstantiatedSource {
  if (!isNodeLike()) {
    throw new Error('Synchronous wasm kernel loading is only supported in Node-like runtimes')
  }
  const imports = {
    env: {
      abort(_message: number, _fileName: number, lineNumber: number, columnNumber: number) {
        throw new Error(`AssemblyScript abort at ${lineNumber}:${columnNumber}`)
      },
    },
  }
  const fs = process.getBuiltinModule('fs') as typeof Fs | undefined
  const url = process.getBuiltinModule('node:url') as typeof NodeUrl | undefined
  if (!fs) {
    throw new Error('Node fs module is unavailable')
  }
  if (!url) {
    throw new Error('Node url module is unavailable')
  }
  const wasmPath = ensureWasmBinaryPathForNode({
    importMetaUrl: import.meta.url,
    existsSync: fs.existsSync,
    fileURLToPath: url.fileURLToPath,
  })
  const bytes = fs.readFileSync(wasmPath)
  const module = new WebAssembly.Module(bytes)
  const instance = new WebAssembly.Instance(module, imports)
  return { module, instance }
}

export async function createKernel(): Promise<SpreadsheetKernel> {
  const { instance } = await loadWasmModule()
  if (!isRawKernelExports(instance.exports)) {
    throw new Error('WASM exports did not match the kernel contract')
  }
  return new KernelHandle(instance.exports)
}

export function createKernelSync(): SpreadsheetKernel {
  const { instance } = loadWasmModuleSync()
  if (!isRawKernelExports(instance.exports)) {
    throw new Error('WASM exports did not match the kernel contract')
  }
  return new KernelHandle(instance.exports)
}
