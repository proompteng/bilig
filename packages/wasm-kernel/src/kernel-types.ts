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
  ): void
  evalDenseDirectScalarRowChainStoreTargetBatch(
    leftValues: Float64Array,
    rightValues: Float64Array,
    firstTargets: Uint32Array,
    secondTargets: Uint32Array,
    rowCount: number,
    firstFormulaCode: number,
    secondFormulaScale: number,
    secondFormulaOffset: number,
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
  ): void
  evalDirectCriteriaMatchedAggregateBatch(
    aggregateKinds: Uint8Array,
    matchStarts: Uint32Array,
    matchLengths: Uint32Array,
    matchedRows: Uint32Array,
    aggregateTags: Uint8Array,
    aggregateNumbers: Float64Array,
    aggregateErrors: Uint16Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void
  evalDirectCriteriaPredicateAggregateBatch(
    aggregateKind: number,
    rowCount: number,
    criteriaOps: Uint8Array,
    criteriaKinds: Uint8Array,
    criteriaValues: Float64Array,
    criteriaStringIds: Uint32Array,
    criteriaTags: Uint8Array,
    criteriaNumbers: Float64Array,
    criteriaStringIdsByRow: Uint32Array,
    aggregateTags: Uint8Array,
    aggregateNumbers: Float64Array,
    aggregateErrors: Uint16Array,
    outTags: Uint8Array,
    outNumbers: Float64Array,
    outErrors: Uint16Array,
  ): void
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
