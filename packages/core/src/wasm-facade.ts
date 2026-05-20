import { type CellValue, type ErrorCode, ValueTag } from '@bilig/protocol'
import { createKernel, createKernelSync, type SpreadsheetKernel } from '@bilig/wasm-kernel'
import type { CellStore } from './cell-store.js'
import type { StringPool } from './string-pool.js'

export interface WasmFormulaUploadLayout {
  targets: Uint32Array
  programs: Uint32Array
  programOffsets: Uint32Array
  programLengths: Uint32Array
  constants: Float64Array
  constantOffsets: Uint32Array
  constantLengths: Uint32Array
}

export interface WasmRangeUploadLayout {
  members: Uint32Array
  offsets: Uint32Array
  lengths: Uint32Array
  rowCounts: Uint32Array
  colCounts: Uint32Array
}

export interface WasmSpillResult {
  rows: number
  cols: number
  values: CellValue[]
}

export interface WasmDirectScalarValueBatchLayout {
  operators: Uint8Array
  leftBatchRefs: Uint32Array
  leftTags: Uint8Array
  leftValues: Float64Array
  leftErrors: Uint16Array
  rightBatchRefs: Uint32Array
  rightTags: Uint8Array
  rightValues: Float64Array
  rightErrors: Uint16Array
  resultOffsets: Float64Array
  outTags: Uint8Array
  outNumbers: Float64Array
  outErrors: Uint16Array
}

export interface WasmDirectScalarStoreTargetBatchLayout {
  targets: Uint32Array
  operators: Uint8Array
  leftBatchRefs: Uint32Array
  leftTags: Uint8Array
  leftValues: Float64Array
  leftErrors: Uint16Array
  rightBatchRefs: Uint32Array
  rightTags: Uint8Array
  rightValues: Float64Array
  rightErrors: Uint16Array
  resultOffsets: Float64Array
}

export interface WasmDenseDirectScalarRowChainStoreTargetBatchLayout {
  leftValues: Float64Array
  rightValues: Float64Array
  firstTargets: Uint32Array
  secondTargets: Uint32Array
  rowCount: number
  firstFormulaCode: number
  secondFormulaScale: number
  secondFormulaOffset: number
}

export interface WasmDenseNumericRowAggregateBatchLayout {
  aggregateKind: number
  values: Float64Array
  rowCount: number
  prefixColCount: number
  startColOffset: number
  aggregateColCount: number
  resultOffset: number
  outNumbers: Float64Array
}

export interface WasmAnchoredPrefixAggregateBatchLayout {
  aggregateKind: number
  tags: Uint8Array
  numbers: Float64Array
  errors: Uint16Array
  rowCount: number
  colCount: number
  formulaRowEnds: Uint32Array
  resultOffsets: Float64Array
  outTags: Uint8Array
  outNumbers: Float64Array
  outErrors: Uint16Array
}

export interface WasmDirectCriteriaMatchedAggregateBatchLayout {
  aggregateKinds: Uint8Array
  matchStarts: Uint32Array
  matchLengths: Uint32Array
  matchedRows: Uint32Array
  aggregateTags: Uint8Array
  aggregateNumbers: Float64Array
  aggregateErrors: Uint16Array
  outTags: Uint8Array
  outNumbers: Float64Array
  outErrors: Uint16Array
}

export interface WasmDirectCriteriaPredicateAggregateBatchLayout {
  aggregateKind: number
  rowCount: number
  criteriaOps: Uint8Array
  criteriaKinds: Uint8Array
  criteriaValues: Float64Array
  criteriaStringIds: Uint32Array
  criteriaTags: Uint8Array
  criteriaNumbers: Float64Array
  criteriaStringIdsByRow: Uint32Array
  aggregateTags: Uint8Array
  aggregateNumbers: Float64Array
  aggregateErrors: Uint16Array
  outTags: Uint8Array
  outNumbers: Float64Array
  outErrors: Uint16Array
}

export interface WasmUniformNumericLookupBatchLayout {
  kinds: Uint8Array
  matchModes: Uint8Array
  starts: Float64Array
  steps: Float64Array
  lengths: Uint32Array
  repeatedRunLengths: Uint32Array
  lookupTags: Uint8Array
  lookupNumbers: Float64Array
  outTags: Uint8Array
  outNumbers: Float64Array
  outErrors: Uint16Array
}

const OUTPUT_STRING_BASE = 2147483648

function decodeSpillValue(tag: number, rawValue: number, strings: StringPool, outputStrings: readonly string[]): CellValue {
  switch (tag) {
    case 1:
      return { tag: ValueTag.Number, value: rawValue }
    case 2:
      return { tag: ValueTag.Boolean, value: rawValue !== 0 }
    case 3: {
      const outputIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : -1
      return {
        tag: ValueTag.String,
        value: outputIndex >= 0 ? (outputStrings[outputIndex] ?? '') : strings.get(rawValue),
        stringId: 0,
      }
    }
    case 4:
      return { tag: ValueTag.Error, code: rawValue as ErrorCode }
    default:
      return { tag: ValueTag.Empty }
  }
}

export class WasmKernelFacade {
  private kernel: SpreadsheetKernel | null = null
  private initPromise: Promise<void> | null = null
  private uploadedStringPoolSize = 0
  private storeSynced = false

  get ready(): boolean {
    return this.kernel !== null
  }

  async init(): Promise<void> {
    if (this.kernel) {
      return
    }
    if (this.initPromise) return this.initPromise
    this.initPromise = (async () => {
      try {
        const kernel = await createKernel()
        if (this.kernel === null) {
          this.kernel = kernel
          kernel.init(64, 64, 64, 64, 64)
          this.storeSynced = false
        }
      } catch {
        if (this.kernel === null) {
          this.kernel = null
        }
      }
    })()
    return this.initPromise
  }

  initSyncIfPossible(): boolean {
    if (this.kernel) {
      return true
    }
    try {
      const kernel = createKernelSync()
      if (this.kernel === null) {
        this.kernel = kernel
        kernel.init(64, 64, 64, 64, 64)
        this.storeSynced = false
      }
      this.initPromise ??= Promise.resolve()
      return true
    } catch {
      return false
    }
  }

  ensureCapacity(
    cellCapacity: number,
    formulaCapacity: number,
    constantCapacity: number,
    rangeCapacity = this.kernel?.getRangeCapacity() ?? 64,
    memberCapacity = this.kernel?.getMemberCapacity() ?? 64,
  ): void {
    this.kernel?.ensureCellCapacity(cellCapacity)
    this.kernel?.ensureFormulaCapacity(formulaCapacity)
    this.kernel?.ensureConstantCapacity(constantCapacity)
    this.kernel?.ensureRangeCapacity(rangeCapacity)
    this.kernel?.ensureMemberCapacity(memberCapacity)
  }

  uploadFormulas(layout: WasmFormulaUploadLayout): void {
    if (!this.kernel) return
    const wasmFormulaCount = layout.targets.length

    this.ensureCapacity(this.kernel.getCellCapacity(), Math.max(wasmFormulaCount, 1), Math.max(layout.constants.length, 1))
    this.kernel.uploadPrograms(layout.programs, layout.programOffsets, layout.programLengths, layout.targets)
    this.kernel.uploadConstants(layout.constants, layout.constantOffsets, layout.constantLengths)
  }

  uploadRanges(layout: WasmRangeUploadLayout): void {
    if (!this.kernel) return
    const rangeCapacity = Math.max(layout.offsets.length, 1)
    const memberCapacity = Math.max(layout.members.length, 1)

    this.ensureCapacity(
      this.kernel.getCellCapacity(),
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      rangeCapacity,
      memberCapacity,
    )
    this.kernel.uploadRangeMembers(layout.members, layout.offsets, layout.lengths)
    this.kernel.uploadRangeShapes(layout.rowCounts, layout.colCounts)
  }

  uploadVolatileNowSerial(nowSerial: number): void {
    this.kernel?.uploadVolatileNowSerial(nowSerial)
  }

  uploadVolatileRandomValues(values: Float64Array): void {
    this.kernel?.uploadVolatileRandomValues(values)
  }

  syncStringPool(layout: { offsets: Uint32Array; lengths: Uint32Array; data: Uint16Array }): void {
    if (!this.kernel) return
    if (layout.lengths.length === this.uploadedStringPoolSize) {
      return
    }
    this.kernel.uploadStrings(layout.offsets, layout.lengths, layout.data)
    this.uploadedStringPoolSize = layout.lengths.length
  }

  resetStoreState(): void {
    if (!this.kernel) {
      return
    }
    this.storeSynced = false
    this.kernel.readTags().fill(0)
    this.kernel.readNumbers().fill(0)
    this.kernel.readStringIds().fill(0)
    this.kernel.readErrors().fill(0)
    this.kernel.readRangeOffsets().fill(0)
    this.kernel.readRangeLengths().fill(0)
    this.kernel.readRangeMembers().fill(0)
    this.kernel.readProgramOffsets().fill(0)
    this.kernel.readProgramLengths().fill(0)
    this.kernel.readConstantOffsets().fill(0)
    this.kernel.readConstantLengths().fill(0)
    this.kernel.readConstants().fill(0)
  }

  syncFromStore(store: CellStore, changedCellIndices?: readonly number[] | Uint32Array): void {
    if (!this.kernel) return
    this.ensureCapacity(
      store.size,
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      this.kernel.getRangeCapacity(),
      this.kernel.getMemberCapacity(),
    )
    if (changedCellIndices === undefined || !this.storeSynced) {
      this.kernel.writeCells(
        store.tags.slice(0, store.size),
        store.numbers.slice(0, store.size),
        store.stringIds.slice(0, store.size),
        store.errors.slice(0, store.size),
      )
      this.storeSynced = true
      return
    }
    if (changedCellIndices.length === 0) {
      return
    }

    const tags = this.kernel.readTags()
    const numbers = this.kernel.readNumbers()
    const stringIds = this.kernel.readStringIds()
    const errors = this.kernel.readErrors()
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const cellIndex = changedCellIndices[index]!
      if (cellIndex >= store.size) {
        continue
      }
      tags[cellIndex] = store.tags[cellIndex]!
      numbers[cellIndex] = store.numbers[cellIndex]!
      stringIds[cellIndex] = store.stringIds[cellIndex]!
      errors[cellIndex] = store.errors[cellIndex]!
    }
  }

  evalBatch(cellIndices: Uint32Array): void {
    this.kernel?.evalBatch(cellIndices)
  }

  evalDirectScalarValueBatch(layout: WasmDirectScalarValueBatchLayout): boolean {
    if (!this.kernel) {
      return false
    }
    this.kernel.evalDirectScalarValueBatch(
      layout.operators,
      layout.leftBatchRefs,
      layout.leftTags,
      layout.leftValues,
      layout.leftErrors,
      layout.rightBatchRefs,
      layout.rightTags,
      layout.rightValues,
      layout.rightErrors,
      layout.resultOffsets,
      layout.outTags,
      layout.outNumbers,
      layout.outErrors,
    )
    return true
  }

  evalDirectScalarStoreTargetBatch(layout: WasmDirectScalarStoreTargetBatchLayout, cellCapacity: number): boolean {
    if (!this.kernel) {
      return false
    }
    this.ensureCapacity(
      cellCapacity,
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      this.kernel.getRangeCapacity(),
      this.kernel.getMemberCapacity(),
    )
    this.kernel.evalDirectScalarStoreTargetBatch(
      layout.targets,
      layout.operators,
      layout.leftBatchRefs,
      layout.leftTags,
      layout.leftValues,
      layout.leftErrors,
      layout.rightBatchRefs,
      layout.rightTags,
      layout.rightValues,
      layout.rightErrors,
      layout.resultOffsets,
    )
    return true
  }

  evalDenseDirectScalarRowChainStoreTargetBatch(
    layout: WasmDenseDirectScalarRowChainStoreTargetBatchLayout,
    cellCapacity: number,
  ): boolean {
    if (!this.kernel) {
      return false
    }
    this.ensureCapacity(
      cellCapacity,
      this.kernel.getFormulaCapacity(),
      this.kernel.getConstantCapacity(),
      this.kernel.getRangeCapacity(),
      this.kernel.getMemberCapacity(),
    )
    this.kernel.evalDenseDirectScalarRowChainStoreTargetBatch(
      layout.leftValues,
      layout.rightValues,
      layout.firstTargets,
      layout.secondTargets,
      layout.rowCount,
      layout.firstFormulaCode,
      layout.secondFormulaScale,
      layout.secondFormulaOffset,
    )
    return true
  }

  evalDenseNumericRowAggregateBatch(layout: WasmDenseNumericRowAggregateBatchLayout): boolean {
    if (!this.kernel) {
      return false
    }
    this.kernel.evalDenseNumericRowAggregateBatch(
      layout.aggregateKind,
      layout.values,
      layout.rowCount,
      layout.prefixColCount,
      layout.startColOffset,
      layout.aggregateColCount,
      layout.resultOffset,
      layout.outNumbers,
    )
    return true
  }

  evalAnchoredPrefixAggregateBatch(layout: WasmAnchoredPrefixAggregateBatchLayout): boolean {
    if (!this.kernel) {
      return false
    }
    this.kernel.evalAnchoredPrefixAggregateBatch(
      layout.aggregateKind,
      layout.tags,
      layout.numbers,
      layout.errors,
      layout.rowCount,
      layout.colCount,
      layout.formulaRowEnds,
      layout.resultOffsets,
      layout.outTags,
      layout.outNumbers,
      layout.outErrors,
    )
    return true
  }

  evalDirectCriteriaMatchedAggregateBatch(layout: WasmDirectCriteriaMatchedAggregateBatchLayout): boolean {
    if (!this.kernel) {
      return false
    }
    this.kernel.evalDirectCriteriaMatchedAggregateBatch(
      layout.aggregateKinds,
      layout.matchStarts,
      layout.matchLengths,
      layout.matchedRows,
      layout.aggregateTags,
      layout.aggregateNumbers,
      layout.aggregateErrors,
      layout.outTags,
      layout.outNumbers,
      layout.outErrors,
    )
    return true
  }

  evalDirectCriteriaPredicateAggregateBatch(layout: WasmDirectCriteriaPredicateAggregateBatchLayout): boolean {
    if (!this.kernel) {
      return false
    }
    this.kernel.evalDirectCriteriaPredicateAggregateBatch(
      layout.aggregateKind,
      layout.rowCount,
      layout.criteriaOps,
      layout.criteriaKinds,
      layout.criteriaValues,
      layout.criteriaStringIds,
      layout.criteriaTags,
      layout.criteriaNumbers,
      layout.criteriaStringIdsByRow,
      layout.aggregateTags,
      layout.aggregateNumbers,
      layout.aggregateErrors,
      layout.outTags,
      layout.outNumbers,
      layout.outErrors,
    )
    return true
  }

  evalUniformNumericLookupBatch(layout: WasmUniformNumericLookupBatchLayout): boolean {
    if (!this.kernel) {
      return false
    }
    this.kernel.evalUniformNumericLookupBatch(
      layout.kinds,
      layout.matchModes,
      layout.starts,
      layout.steps,
      layout.lengths,
      layout.repeatedRunLengths,
      layout.lookupTags,
      layout.lookupNumbers,
      layout.outTags,
      layout.outNumbers,
      layout.outErrors,
    )
    return true
  }

  readSpill(cellIndex: number, strings: StringPool): WasmSpillResult | undefined {
    if (!this.kernel) {
      return undefined
    }
    const rows = this.kernel.readSpillRows()[cellIndex] ?? 0
    const cols = this.kernel.readSpillCols()[cellIndex] ?? 0
    if (rows === 0 || cols === 0) {
      return undefined
    }
    const offset = this.kernel.readSpillOffsets()[cellIndex] ?? 0
    const length = this.kernel.readSpillLengths()[cellIndex] ?? 0
    const spillTags = this.kernel.readSpillTags()
    const spillValues = this.kernel.readSpillNumbers()
    const outputStrings = this.kernel.readOutputStrings()
    return {
      rows,
      cols,
      values: Array.from({ length }, (_, index) =>
        decodeSpillValue(spillTags[offset + index] ?? ValueTag.Empty, spillValues[offset + index] ?? 0, strings, outputStrings),
      ),
    }
  }

  materializePivotTable(
    sourceRangeIndex: number,
    sourceWidth: number,
    groupByColumnIndices: Uint32Array,
    valueColumnIndices: Uint32Array,
    valueAggregations: Uint8Array,
  ):
    | {
        rows: number
        cols: number
        tags: Uint8Array
        numbers: Float64Array
        stringIds: Uint32Array
        errors: Uint16Array
      }
    | undefined {
    return this.kernel?.materializePivotTable(sourceRangeIndex, sourceWidth, groupByColumnIndices, valueColumnIndices, valueAggregations)
  }

  syncToStore(
    store: CellStore,
    changedCellIndices: Uint32Array,
    strings: StringPool,
    onCellValueWritten?: (cellIndex: number) => void,
  ): void {
    if (!this.kernel) return

    const newStrings = this.kernel.readOutputStrings()
    const tags = this.kernel.readTags()
    const numbers = this.kernel.readNumbers()
    const stringIds = this.kernel.readStringIds()
    const errors = this.kernel.readErrors()
    const storeTags = store.tags
    const storeNumbers = store.numbers
    const storeStringIds = store.stringIds
    const storeErrors = store.errors
    const storeVersions = store.versions

    if (newStrings.length === 0) {
      for (let index = 0; index < changedCellIndices.length; index += 1) {
        const cellIndex = changedCellIndices[index]!
        if (cellIndex >= store.size) {
          continue
        }

        const nextTag = tags[cellIndex]!
        const nextNumber = numbers[cellIndex]!
        const nextStringId = stringIds[cellIndex]!
        const nextError = errors[cellIndex]!

        if (
          storeTags[cellIndex] !== nextTag ||
          storeNumbers[cellIndex] !== nextNumber ||
          storeStringIds[cellIndex] !== nextStringId ||
          storeErrors[cellIndex] !== nextError
        ) {
          storeVersions[cellIndex] = (storeVersions[cellIndex] ?? 0) + 1
          onCellValueWritten?.(cellIndex)
        }

        storeTags[cellIndex] = nextTag
        storeNumbers[cellIndex] = nextNumber
        storeStringIds[cellIndex] = nextStringId
        storeErrors[cellIndex] = nextError
      }
      return
    }

    const outputStringIdMap = new Map<number, number>()
    for (let index = 0; index < newStrings.length; index += 1) {
      outputStringIdMap.set((index | 0x80000000) >>> 0, strings.intern(newStrings[index]!))
    }

    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const cellIndex = changedCellIndices[index]!
      if (cellIndex >= store.size) {
        continue
      }

      const nextTag = tags[cellIndex]!
      const nextNumber = numbers[cellIndex]!
      let nextStringId = stringIds[cellIndex]!
      if ((nextStringId & 0x80000000) !== 0) {
        nextStringId = outputStringIdMap.get(nextStringId) ?? 0
      }
      const nextError = errors[cellIndex]!

      if (
        storeTags[cellIndex] !== nextTag ||
        storeNumbers[cellIndex] !== nextNumber ||
        storeStringIds[cellIndex] !== nextStringId ||
        storeErrors[cellIndex] !== nextError
      ) {
        storeVersions[cellIndex] = (storeVersions[cellIndex] ?? 0) + 1
        onCellValueWritten?.(cellIndex)
      }

      storeTags[cellIndex] = nextTag
      storeNumbers[cellIndex] = nextNumber
      storeStringIds[cellIndex] = nextStringId
      storeErrors[cellIndex] = nextError
    }
  }

  get tags(): Uint8Array {
    return this.kernel?.readTags() ?? new Uint8Array()
  }

  get numbers(): Float64Array {
    return this.kernel?.readNumbers() ?? new Float64Array()
  }

  get stringIds(): Uint32Array {
    return this.kernel?.readStringIds() ?? new Uint32Array()
  }

  get errors(): Uint16Array {
    return this.kernel?.readErrors() ?? new Uint16Array()
  }

  get programOffsets(): Uint32Array {
    return this.kernel?.readProgramOffsets() ?? new Uint32Array()
  }

  get programLengths(): Uint32Array {
    return this.kernel?.readProgramLengths() ?? new Uint32Array()
  }

  get constantOffsets(): Uint32Array {
    return this.kernel?.readConstantOffsets() ?? new Uint32Array()
  }

  get constantLengths(): Uint32Array {
    return this.kernel?.readConstantLengths() ?? new Uint32Array()
  }

  get constants(): Float64Array {
    return this.kernel?.readConstants() ?? new Float64Array()
  }

  get rangeOffsets(): Uint32Array {
    return this.kernel?.readRangeOffsets() ?? new Uint32Array()
  }

  get rangeLengths(): Uint32Array {
    return this.kernel?.readRangeLengths() ?? new Uint32Array()
  }

  get rangeMembers(): Uint32Array {
    return this.kernel?.readRangeMembers() ?? new Uint32Array()
  }
}
