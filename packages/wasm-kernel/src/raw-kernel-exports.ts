export interface RawKernelExports {
  memory: WebAssembly.Memory
  __new(size: number, id: number): number
  __pin(pointer: number): number
  __unpin(pointer: number): void
  init(cellCapacity: number, formulaCapacity: number, constantCapacity: number, rangeCapacity: number, memberCapacity: number): void
  ensureCellCapacity(nextCapacity: number): void
  ensureFormulaCapacity(nextCapacity: number): void
  ensureConstantCapacity(nextCapacity: number): void
  ensureRangeCapacity(nextCapacity: number): void
  ensureMemberCapacity(nextCapacity: number): void
  uploadPrograms(programs: number, offsets: number, lengths: number, targets: number): void
  uploadConstants(constants: number, offsets: number, lengths: number): void
  uploadRangeMembers(members: number, offsets: number, lengths: number): void
  uploadRangeShapes(rowCounts: number, colCounts: number): void
  uploadVolatileNowSerial(nowSerial: number): void
  uploadVolatileRandomValues(values: number): void
  uploadStringLengths(lengths: number): void
  uploadStrings(offsets: number, lengths: number, data: number): void
  writeCells(tags: number, numbers: number, stringIds: number, errors: number): void
  evalDirectScalarValueBatch(
    operators: number,
    leftBatchRefs: number,
    leftTags: number,
    leftValues: number,
    leftErrors: number,
    rightBatchRefs: number,
    rightTags: number,
    rightValues: number,
    rightErrors: number,
    resultOffsets: number,
    outTags: number,
    outNumbers: number,
    outErrors: number,
  ): void
  evalDirectScalarStoreTargetBatch(
    targets: number,
    operators: number,
    leftBatchRefs: number,
    leftTags: number,
    leftValues: number,
    leftErrors: number,
    rightBatchRefs: number,
    rightTags: number,
    rightValues: number,
    rightErrors: number,
    resultOffsets: number,
  ): void
  evalDenseDirectScalarRowChainStoreTargetBatch(
    leftValues: number,
    rightValues: number,
    firstTargets: number,
    secondTargets: number,
    rowCount: number,
    firstFormulaCode: number,
    secondFormulaScale: number,
    secondFormulaOffset: number,
  ): void
  evalDenseNumericRowAggregateBatch(
    aggregateKind: number,
    values: number,
    rowCount: number,
    prefixColCount: number,
    startColOffset: number,
    aggregateColCount: number,
    resultOffset: number,
    outNumbers: number,
  ): void
  evalAnchoredPrefixAggregateBatch(
    aggregateKind: number,
    tags: number,
    numbers: number,
    errors: number,
    rowCount: number,
    colCount: number,
    formulaRowEnds: number,
    resultOffsets: number,
    outTags: number,
    outNumbers: number,
    outErrors: number,
  ): void
  evalDirectCriteriaMatchedAggregateBatch(
    aggregateKinds: number,
    matchStarts: number,
    matchLengths: number,
    matchedRows: number,
    aggregateTags: number,
    aggregateNumbers: number,
    aggregateErrors: number,
    outTags: number,
    outNumbers: number,
    outErrors: number,
  ): void
  evalDirectCriteriaPredicateAggregateBatch(
    aggregateKind: number,
    rowCount: number,
    criteriaOps: number,
    criteriaKinds: number,
    criteriaValues: number,
    criteriaStringIds: number,
    criteriaTags: number,
    criteriaNumbers: number,
    criteriaStringIdsByRow: number,
    aggregateTags: number,
    aggregateNumbers: number,
    aggregateErrors: number,
    outTags: number,
    outNumbers: number,
    outErrors: number,
  ): void
  evalUniformNumericLookupBatch(
    kinds: number,
    matchModes: number,
    starts: number,
    steps: number,
    lengths: number,
    repeatedRunLengths: number,
    lookupTags: number,
    lookupNumbers: number,
    outTags: number,
    outNumbers: number,
    outErrors: number,
  ): void
  evalBatch(cellIndices: number): void
  materializePivotTable(
    sourceRangeIndex: number,
    sourceWidth: number,
    groupByCount: number,
    groupByColumnIndices: number,
    valueCount: number,
    valueColumnIndices: number,
    valueAggregations: number,
  ): void
  getPivotResultTagsPtr(): number
  getPivotResultNumbersPtr(): number
  getPivotResultStringIdsPtr(): number
  getPivotResultErrorsPtr(): number
  pivotResultRows: { value: number }
  pivotResultCols: { value: number }
  getTagsPtr(): number
  getNumbersPtr(): number
  getStringIdsPtr(): number
  getErrorsPtr(): number
  getProgramOffsetsPtr(): number
  getProgramLengthsPtr(): number
  getConstantOffsetsPtr(): number
  getConstantLengthsPtr(): number
  getConstantArenaPtr(): number
  getRangeOffsetsPtr(): number
  getRangeLengthsPtr(): number
  getRangeMembersPtr(): number
  getOutputStringLengthsPtr(): number
  getOutputStringOffsetsPtr(): number
  getOutputStringDataPtr(): number
  getOutputStringCount(): number
  getOutputStringDataLength(): number
  getSpillResultRowsPtr(): number
  getSpillResultColsPtr(): number
  getSpillResultOffsetsPtr(): number
  getSpillResultLengthsPtr(): number
  getSpillResultTagsPtr(): number
  getSpillResultNumbersPtr(): number
  getSpillResultValueCount(): number
  getCellCapacity(): number
  getFormulaCapacity(): number
  getConstantCapacity(): number
  getRangeCapacity(): number
  getMemberCapacity(): number
}

export function isRawKernelExports(value: unknown): value is RawKernelExports {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const requiredKeys = [
    'memory',
    '__new',
    '__pin',
    '__unpin',
    'init',
    'ensureCellCapacity',
    'ensureFormulaCapacity',
    'ensureConstantCapacity',
    'ensureRangeCapacity',
    'ensureMemberCapacity',
    'uploadPrograms',
    'uploadConstants',
    'uploadRangeMembers',
    'uploadRangeShapes',
    'uploadVolatileNowSerial',
    'uploadVolatileRandomValues',
    'uploadStringLengths',
    'uploadStrings',
    'writeCells',
    'evalDirectScalarValueBatch',
    'evalDirectScalarStoreTargetBatch',
    'evalDenseDirectScalarRowChainStoreTargetBatch',
    'evalDenseNumericRowAggregateBatch',
    'evalAnchoredPrefixAggregateBatch',
    'evalDirectCriteriaMatchedAggregateBatch',
    'evalDirectCriteriaPredicateAggregateBatch',
    'evalUniformNumericLookupBatch',
    'evalBatch',
    'materializePivotTable',
    'getPivotResultTagsPtr',
    'getPivotResultNumbersPtr',
    'getPivotResultStringIdsPtr',
    'getPivotResultErrorsPtr',
    'pivotResultRows',
    'pivotResultCols',
    'getTagsPtr',
    'getNumbersPtr',
    'getStringIdsPtr',
    'getErrorsPtr',
    'getProgramOffsetsPtr',
    'getProgramLengthsPtr',
    'getConstantOffsetsPtr',
    'getConstantLengthsPtr',
    'getConstantArenaPtr',
    'getRangeOffsetsPtr',
    'getRangeLengthsPtr',
    'getRangeMembersPtr',
    'getOutputStringLengthsPtr',
    'getOutputStringOffsetsPtr',
    'getOutputStringDataPtr',
    'getOutputStringCount',
    'getOutputStringDataLength',
    'getSpillResultRowsPtr',
    'getSpillResultColsPtr',
    'getSpillResultOffsetsPtr',
    'getSpillResultLengthsPtr',
    'getSpillResultTagsPtr',
    'getSpillResultNumbersPtr',
    'getSpillResultValueCount',
    'getCellCapacity',
    'getFormulaCapacity',
    'getConstantCapacity',
    'getRangeCapacity',
    'getMemberCapacity',
  ] as const
  return requiredKeys.every((key) => key in value)
}
