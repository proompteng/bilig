import type * as Fs from 'node:fs'
import type * as FsPromises from 'node:fs/promises'
import type * as NodeUrl from 'node:url'
import {
  evalAnchoredPrefixAggregateBatchRaw,
  evalDenseCellRangeAggregateBatchRaw,
  evalDenseNumericRowAggregateBatchRaw,
} from './raw-kernel-direct-aggregate-bridge.js'
import {
  evalDirectCriteriaMatchedAggregateBatchRaw,
  evalDirectCriteriaPredicateAggregateBatchRaw,
} from './raw-kernel-direct-criteria-aggregate-bridge.js'
import { evalUniformNumericLookupBatchRaw } from './raw-kernel-direct-lookup-bridge.js'
import {
  evalDenseDirectScalarRowChainDivideStoreTargetBatchRaw,
  evalDenseDirectScalarRowChainStoreTargetBatchRaw,
  evalDirectConditionalPickBatchRaw,
  evalDirectScalarStoreTargetBatchRaw,
  evalDirectScalarValueBatchRaw,
} from './raw-kernel-direct-scalar-bridge.js'
import { materializePivotTableRaw } from './raw-kernel-pivot-bridge.js'
import { isRawKernelExports, type RawKernelExports } from './raw-kernel-exports.js'
import { evalBatchRaw } from './raw-kernel-vm-bridge.js'
import { float64ArraySpec, RawKernelArrayBridge, uint16ArraySpec, uint32ArraySpec, uint8ArraySpec } from './raw-kernel-array-bridge.js'
import type { SpreadsheetKernel } from './kernel-types.js'
import { RawWorksheetImportStorage, type WorksheetImportStorage } from './worksheet-import-storage.js'

export type { SpreadsheetKernel } from './kernel-types.js'
export type { WorksheetImportStorage, WorksheetImportStorageSnapshot } from './worksheet-import-storage.js'

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

class RawKernelBridge {
  private readonly arrayBridge: RawKernelArrayBridge

  constructor(private readonly raw: RawKernelExports) {
    this.arrayBridge = new RawKernelArrayBridge(raw)
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
    const [programsPtr, offsetsPtr, lengthsPtr, targetsPtr] = this.arrayBridge.lowerTypedArrays([
      [programs, uint32ArraySpec],
      [offsets, uint32ArraySpec],
      [lengths, uint32ArraySpec],
      [targets, uint32ArraySpec],
    ])
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
    const [constantsPtr, offsetsPtr, lengthsPtr] = this.arrayBridge.lowerTypedArrays([
      [constants, float64ArraySpec],
      [offsets, uint32ArraySpec],
      [lengths, uint32ArraySpec],
    ])
    try {
      this.raw.uploadConstants(constantsPtr, offsetsPtr, lengthsPtr)
    } finally {
      this.raw.__unpin(constantsPtr)
      this.raw.__unpin(offsetsPtr)
      this.raw.__unpin(lengthsPtr)
    }
  }

  uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void {
    const [membersPtr, offsetsPtr, lengthsPtr] = this.arrayBridge.lowerTypedArrays([
      [members, uint32ArraySpec],
      [offsets, uint32ArraySpec],
      [lengths, uint32ArraySpec],
    ])
    try {
      this.raw.uploadRangeMembers(membersPtr, offsetsPtr, lengthsPtr)
    } finally {
      this.raw.__unpin(membersPtr)
      this.raw.__unpin(offsetsPtr)
      this.raw.__unpin(lengthsPtr)
    }
  }

  uploadRangeShapes(rowCounts: Uint32Array, colCounts: Uint32Array): void {
    const [rowCountsPtr, colCountsPtr] = this.arrayBridge.lowerTypedArrays([
      [rowCounts, uint32ArraySpec],
      [colCounts, uint32ArraySpec],
    ])
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
    const [valuesPtr] = this.arrayBridge.lowerTypedArrays([[values, float64ArraySpec]])
    try {
      this.raw.uploadVolatileRandomValues(valuesPtr)
    } finally {
      this.raw.__unpin(valuesPtr)
    }
  }

  uploadStringLengths(lengths: Uint32Array): void {
    const [lengthsPtr] = this.arrayBridge.lowerTypedArrays([[lengths, uint32ArraySpec]])
    try {
      this.raw.uploadStringLengths(lengthsPtr)
    } finally {
      this.raw.__unpin(lengthsPtr)
    }
  }

  uploadStrings(offsets: Uint32Array, lengths: Uint32Array, data: Uint16Array): void {
    const [offsetsPtr, lengthsPtr, dataPtr] = this.arrayBridge.lowerTypedArrays([
      [offsets, uint32ArraySpec],
      [lengths, uint32ArraySpec],
      [data, uint16ArraySpec],
    ])
    try {
      this.raw.uploadStrings(offsetsPtr, lengthsPtr, dataPtr)
    } finally {
      this.raw.__unpin(offsetsPtr)
      this.raw.__unpin(lengthsPtr)
      this.raw.__unpin(dataPtr)
    }
  }

  writeCells(tags: Uint8Array, numbers: Float64Array, stringIds: Uint32Array, errors: Uint16Array): void {
    const [tagsPtr, numbersPtr, stringIdsPtr, errorsPtr] = this.arrayBridge.lowerTypedArrays([
      [tags, uint8ArraySpec],
      [numbers, float64ArraySpec],
      [stringIds, uint32ArraySpec],
      [errors, uint16ArraySpec],
    ])
    try {
      this.raw.writeCells(tagsPtr, numbersPtr, stringIdsPtr, errorsPtr)
    } finally {
      this.raw.__unpin(tagsPtr)
      this.raw.__unpin(numbersPtr)
      this.raw.__unpin(stringIdsPtr)
      this.raw.__unpin(errorsPtr)
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
    evalDenseNumericRowAggregateBatchRaw(
      this.raw,
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

  evalDenseCellRangeAggregateBatch(
    aggregateKind: number,
    tags: Uint8Array,
    numbers: Float64Array,
    errors: Uint16Array,
    rowCount: number,
    cellCount: number,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    evalDenseCellRangeAggregateBatchRaw(this.raw, aggregateKind, tags, numbers, errors, rowCount, cellCount, outTags, outNumbers, outErrors)
  }

  evalAnchoredPrefixAggregateBatch(
    aggregateKind: number,
    tags: Uint8Array,
    numbers: Float64Array,
    errors: Uint16Array,
    rowCount: number,
    colCount: number,
    formulaRowEnds: Uint32Array,
    resultOffsets: Float64Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    evalAnchoredPrefixAggregateBatchRaw(
      this.raw,
      aggregateKind,
      tags,
      numbers,
      errors,
      rowCount,
      colCount,
      formulaRowEnds,
      resultOffsets,
      outTags,
      outNumbers,
      outErrors,
    )
  }

  evalDirectCriteriaMatchedAggregateBatch(...args: Parameters<SpreadsheetKernel['evalDirectCriteriaMatchedAggregateBatch']>): void {
    evalDirectCriteriaMatchedAggregateBatchRaw(this.raw, ...args)
  }

  evalDirectCriteriaPredicateAggregateBatch(...args: Parameters<SpreadsheetKernel['evalDirectCriteriaPredicateAggregateBatch']>): void {
    evalDirectCriteriaPredicateAggregateBatchRaw(this.raw, ...args)
  }

  evalUniformNumericLookupBatch(
    kinds: Uint8Array,
    matchModes: Uint8Array,
    starts: Float64Array,
    steps: Float64Array,
    lengths: Uint32Array,
    repeatedRunLengths: Uint32Array,
    lookupTags: Uint8Array,
    lookupNumbers: Float64Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    evalUniformNumericLookupBatchRaw(
      this.raw,
      kinds,
      matchModes,
      starts,
      steps,
      lengths,
      repeatedRunLengths,
      lookupTags,
      lookupNumbers,
      outTags,
      outNumbers,
      outErrors,
    )
  }

  evalBatch(cellIndices: Uint32Array): void {
    evalBatchRaw(this.raw, cellIndices)
  }

  materializePivotTable(
    sourceRangeIndex: number,
    sourceWidth: number,
    groupByColumnIndices: Uint32Array,
    valueColumnIndices: Uint32Array,
    valueAggregations: Uint8Array,
  ): void {
    materializePivotTableRaw(this.raw, sourceRangeIndex, sourceWidth, groupByColumnIndices, valueColumnIndices, valueAggregations)
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
    evalDirectScalarValueBatchRaw(
      this.raw,
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

  evalDirectScalarStoreTargetBatch(
    targets: Uint32Array,
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
  ): void {
    evalDirectScalarStoreTargetBatchRaw(
      this.raw,
      targets,
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
    )
    this.refreshViews()
  }

  evalDirectConditionalPickBatch(...args: Parameters<SpreadsheetKernel['evalDirectConditionalPickBatch']>): void {
    evalDirectConditionalPickBatchRaw(this.raw, ...args)
  }

  evalDenseDirectScalarRowChainStoreTargetBatch(
    ...args: Parameters<SpreadsheetKernel['evalDenseDirectScalarRowChainStoreTargetBatch']>
  ): void {
    evalDenseDirectScalarRowChainStoreTargetBatchRaw(this.raw, ...args)
    this.refreshViews()
  }

  evalDenseDirectScalarRowChainDivideStoreTargetBatch(
    ...args: Parameters<SpreadsheetKernel['evalDenseDirectScalarRowChainDivideStoreTargetBatch']>
  ): void {
    evalDenseDirectScalarRowChainDivideStoreTargetBatchRaw(this.raw, ...args)
    this.refreshViews()
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

  evalDenseCellRangeAggregateBatch(
    aggregateKind: number,
    tags: Uint8Array,
    numbers: Float64Array,
    errors: Uint16Array,
    rowCount: number,
    cellCount: number,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    this.bridge.evalDenseCellRangeAggregateBatch(aggregateKind, tags, numbers, errors, rowCount, cellCount, outTags, outNumbers, outErrors)
  }

  evalAnchoredPrefixAggregateBatch(
    aggregateKind: number,
    tags: Uint8Array,
    numbers: Float64Array,
    errors: Uint16Array,
    rowCount: number,
    colCount: number,
    formulaRowEnds: Uint32Array,
    resultOffsets: Float64Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    this.bridge.evalAnchoredPrefixAggregateBatch(
      aggregateKind,
      tags,
      numbers,
      errors,
      rowCount,
      colCount,
      formulaRowEnds,
      resultOffsets,
      outTags,
      outNumbers,
      outErrors,
    )
  }

  evalDirectCriteriaMatchedAggregateBatch(...args: Parameters<SpreadsheetKernel['evalDirectCriteriaMatchedAggregateBatch']>): void {
    this.bridge.evalDirectCriteriaMatchedAggregateBatch(...args)
    this.refreshViews()
  }

  evalDirectCriteriaPredicateAggregateBatch(...args: Parameters<SpreadsheetKernel['evalDirectCriteriaPredicateAggregateBatch']>): void {
    this.bridge.evalDirectCriteriaPredicateAggregateBatch(...args)
    this.refreshViews()
  }

  evalUniformNumericLookupBatch(
    kinds: Uint8Array,
    matchModes: Uint8Array,
    starts: Float64Array,
    steps: Float64Array,
    lengths: Uint32Array,
    repeatedRunLengths: Uint32Array,
    lookupTags: Uint8Array,
    lookupNumbers: Float64Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void {
    this.bridge.evalUniformNumericLookupBatch(
      kinds,
      matchModes,
      starts,
      steps,
      lengths,
      repeatedRunLengths,
      lookupTags,
      lookupNumbers,
      outTags,
      outNumbers,
      outErrors,
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
    this.ensureViewsFresh()
    return this.tags
  }

  readNumbers(): Float64Array {
    this.ensureViewsFresh()
    return this.numbers
  }

  readStringIds(): Uint32Array {
    this.ensureViewsFresh()
    return this.stringIds
  }

  readErrors(): Uint16Array {
    this.ensureViewsFresh()
    return this.errors
  }

  readProgramOffsets(): Uint32Array {
    this.ensureViewsFresh()
    return this.programOffsets
  }

  readProgramLengths(): Uint32Array {
    this.ensureViewsFresh()
    return this.programLengths
  }

  readConstantOffsets(): Uint32Array {
    this.ensureViewsFresh()
    return this.constantOffsets
  }

  readConstantLengths(): Uint32Array {
    this.ensureViewsFresh()
    return this.constantLengths
  }

  readConstants(): Float64Array {
    this.ensureViewsFresh()
    return this.constants
  }

  readRangeOffsets(): Uint32Array {
    this.ensureViewsFresh()
    return this.rangeOffsets
  }

  readRangeLengths(): Uint32Array {
    this.ensureViewsFresh()
    return this.rangeLengths
  }

  readRangeMembers(): Uint32Array {
    this.ensureViewsFresh()
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
    this.ensureViewsFresh()
    return this.spillRows
  }

  readSpillCols(): Uint32Array {
    this.ensureViewsFresh()
    return this.spillCols
  }

  readSpillOffsets(): Uint32Array {
    this.ensureViewsFresh()
    return this.spillOffsets
  }

  readSpillLengths(): Uint32Array {
    this.ensureViewsFresh()
    return this.spillLengths
  }

  readSpillTags(): Uint8Array {
    this.ensureViewsFresh()
    return this.spillTags
  }

  readSpillNumbers(): Float64Array {
    this.ensureViewsFresh()
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

  private ensureViewsFresh(): void {
    if (this.tags.buffer !== this.raw.memory.buffer) {
      this.refreshViews()
    }
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

export function createWorksheetImportStorageSync(): WorksheetImportStorage {
  const { instance } = loadWasmModuleSync()
  if (!isRawKernelExports(instance.exports)) {
    throw new Error('WASM exports did not match the kernel contract')
  }
  return new RawWorksheetImportStorage(instance.exports)
}
