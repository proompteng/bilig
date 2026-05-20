import type { ErrorCode } from '@bilig/protocol'
import type { U32 } from '../runtime-state.js'

export type DirectScalarCurrentOperand = { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode }

export class PendingNumericCellValues {
  private readonly values: number[] = []
  private readonly assigned: boolean[] = []

  get(cellIndex: number): number | undefined {
    return this.assigned[cellIndex] === true ? this.values[cellIndex] : undefined
  }

  has(cellIndex: number): boolean {
    return this.assigned[cellIndex] === true
  }

  set(cellIndex: number, value: number): void {
    this.assigned[cellIndex] = true
    this.values[cellIndex] = value
  }
}

export class DirectFormulaIndexCollection {
  private readonly cellIndices: number[] = []
  private sharedCellIndices: readonly number[] | U32 | undefined
  private indexByCell: Map<number, number> | undefined
  private deltas: number[] | undefined
  private deltaAssigned: boolean[] | undefined
  private scalarDeltaAssigned: boolean[] | undefined
  private currentResults: DirectScalarCurrentOperand[] | undefined
  private currentResultAssigned: boolean[] | undefined
  private directFormulaCoveredInputCellIndices: number[] | undefined
  private directFormulaCoveredInputCellSet: Set<number> | undefined
  private directRangeCoveredInputCellIndices: number[] | undefined
  private directRangeCoveredInputCellSet: Set<number> | undefined
  private deltaCount = 0
  private scalarDeltaCount = 0
  private constantDelta: number | undefined
  private validatedScalarDeltaSize = -1
  private cleanScalarDeltaSize = -1
  private trustedDirectScalarDeltaSize = -1
  private linearHasProbeCount = 0

  get size(): number {
    return this.sharedCellIndices?.length ?? this.cellIndices.length
  }

  add(cellIndex: number): void {
    this.validatedScalarDeltaSize = -1
    this.cleanScalarDeltaSize = -1
    this.trustedDirectScalarDeltaSize = -1
    this.materializeConstantDeltas()
    if (this.indexByCell) {
      if (this.indexByCell.has(cellIndex)) {
        return
      }
      this.materializeSharedCellIndices()
      this.indexByCell.set(cellIndex, this.cellIndices.length)
      this.cellIndices.push(cellIndex)
      return
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return
      }
    }
    this.materializeSharedCellIndices()
    this.cellIndices.push(cellIndex)
    if (this.size > 16) {
      this.materializeIndexByCell()
    }
  }

  has(cellIndex: number): boolean {
    if (!this.indexByCell && this.size > 16) {
      if (this.sharedCellIndices !== undefined && this.linearHasProbeCount < 4) {
        this.linearHasProbeCount += 1
        for (let index = 0; index < this.sharedCellIndices.length; index += 1) {
          if (this.sharedCellIndices[index] === cellIndex) {
            return true
          }
        }
        return false
      }
      this.materializeIndexByCell()
    }
    if (this.indexByCell) {
      return this.indexByCell.has(cellIndex)
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return true
      }
    }
    return false
  }

  hasAny(cellIndices: readonly number[] | U32, count = cellIndices.length): boolean {
    const length = Math.min(count, cellIndices.length)
    if (length <= 0 || this.size === 0) {
      return false
    }
    if (!this.indexByCell && (length <= 4 || this.size <= 16)) {
      for (let inputIndex = 0; inputIndex < length; inputIndex += 1) {
        const cellIndex = cellIndices[inputIndex]!
        for (let collectionIndex = 0; collectionIndex < this.size; collectionIndex += 1) {
          if (this.getCellIndexAt(collectionIndex) === cellIndex) {
            return true
          }
        }
      }
      return false
    }
    for (let index = 0; index < length; index += 1) {
      if (this.has(cellIndices[index]!)) {
        return true
      }
    }
    return false
  }

  addDelta(cellIndex: number, delta: number): void {
    this.addDeltaWithKind(cellIndex, delta, undefined)
  }

  addScalarDelta(cellIndex: number, delta: number): void {
    this.addDeltaWithKind(cellIndex, delta, 'scalar')
  }

  private addDeltaWithKind(cellIndex: number, delta: number, kind: 'scalar' | undefined): void {
    this.validatedScalarDeltaSize = -1
    this.cleanScalarDeltaSize = -1
    this.trustedDirectScalarDeltaSize = -1
    if (this.constantDelta !== undefined) {
      const existingIndex = this.findIndex(cellIndex)
      if (existingIndex === -1 && Object.is(this.constantDelta, delta)) {
        this.materializeSharedCellIndices()
        if (this.indexByCell) {
          this.indexByCell.set(cellIndex, this.cellIndices.length)
        }
        this.cellIndices.push(cellIndex)
        this.deltaCount += 1
        if (kind === 'scalar') {
          this.scalarDeltaCount += 1
        }
        return
      }
    }
    this.materializeConstantDeltas()
    const index = this.ensureIndex(cellIndex)
    this.deltas ??= []
    this.deltaAssigned ??= []
    this.scalarDeltaAssigned ??= []
    if (!this.deltaAssigned[index]) {
      this.deltaAssigned[index] = true
      this.deltaCount += 1
      this.deltas[index] = delta
      if (kind === 'scalar') {
        this.scalarDeltaAssigned[index] = true
        this.scalarDeltaCount += 1
      }
      return
    }
    this.deltas[index] = (this.deltas[index] ?? 0) + delta
    if (kind !== 'scalar' && this.scalarDeltaAssigned[index]) {
      this.scalarDeltaAssigned[index] = false
      this.scalarDeltaCount -= 1
    }
  }

  appendDeltas(cellIndices: readonly number[] | U32, deltas: readonly number[], kind?: 'scalar'): void {
    if (cellIndices.length === 0) {
      return
    }
    this.validatedScalarDeltaSize = -1
    this.cleanScalarDeltaSize = -1
    this.trustedDirectScalarDeltaSize = -1
    if (this.size !== 0) {
      if (cellIndices.length > 16 || this.size > 16) {
        this.appendPreparedDeltas(cellIndices, deltas, kind)
        return
      }
      for (let index = 0; index < cellIndices.length; index += 1) {
        this.addDeltaWithKind(cellIndices[index]!, deltas[index]!, kind)
      }
      return
    }
    this.sharedCellIndices = cellIndices
    this.deltas = []
    this.deltaAssigned = []
    this.scalarDeltaAssigned = kind === 'scalar' ? [] : undefined
    for (let index = 0; index < cellIndices.length; index += 1) {
      this.deltas[index] = deltas[index]!
      this.deltaAssigned[index] = true
      if (kind === 'scalar') {
        this.scalarDeltaAssigned![index] = true
      }
    }
    this.deltaCount = cellIndices.length
    this.scalarDeltaCount = kind === 'scalar' ? cellIndices.length : 0
  }

  appendConstantDelta(cellIndices: readonly number[] | U32, delta: number, kind?: 'scalar'): void {
    if (cellIndices.length === 0) {
      return
    }
    this.validatedScalarDeltaSize = -1
    this.cleanScalarDeltaSize = -1
    this.trustedDirectScalarDeltaSize = -1
    if (this.size !== 0) {
      if (cellIndices.length > 16 || this.size > 16) {
        this.appendPreparedConstantDelta(cellIndices, delta, kind)
        return
      }
      for (let index = 0; index < cellIndices.length; index += 1) {
        this.addDeltaWithKind(cellIndices[index]!, delta, kind)
      }
      return
    }
    this.sharedCellIndices = cellIndices
    this.constantDelta = delta
    this.deltaCount = cellIndices.length
    this.scalarDeltaCount = kind === 'scalar' ? cellIndices.length : 0
  }

  hasDelta(cellIndex: number): boolean {
    const index = this.findIndex(cellIndex)
    if (index !== -1 && this.constantDelta !== undefined) {
      return true
    }
    return index !== -1 && this.deltaAssigned?.[index] === true
  }

  getDelta(cellIndex: number): number | undefined {
    const index = this.findIndex(cellIndex)
    if (index !== -1 && this.constantDelta !== undefined) {
      return this.constantDelta
    }
    if (index === -1 || this.deltaAssigned?.[index] !== true) {
      return undefined
    }
    return this.deltas?.[index]
  }

  getDeltaAt(index: number): number | undefined {
    if (this.constantDelta !== undefined && index >= 0 && index < this.size) {
      return this.constantDelta
    }
    if (this.deltaAssigned?.[index] !== true) {
      return undefined
    }
    return this.deltas?.[index]
  }

  addCurrentResult(cellIndex: number, result: DirectScalarCurrentOperand): void {
    this.validatedScalarDeltaSize = -1
    this.cleanScalarDeltaSize = -1
    this.trustedDirectScalarDeltaSize = -1
    this.materializeConstantDeltas()
    const index = this.ensureIndex(cellIndex)
    this.currentResults ??= []
    this.currentResultAssigned ??= []
    this.currentResultAssigned[index] = true
    this.currentResults[index] = result
  }

  getCurrentResult(cellIndex: number): DirectScalarCurrentOperand | undefined {
    const index = this.findIndex(cellIndex)
    if (index === -1 || this.currentResultAssigned?.[index] !== true) {
      return undefined
    }
    return this.currentResults?.[index]
  }

  getCurrentResultAt(index: number): DirectScalarCurrentOperand | undefined {
    if (this.currentResultAssigned?.[index] !== true) {
      return undefined
    }
    return this.currentResults?.[index]
  }

  hasCompleteDeltas(): boolean {
    return this.size > 0 && this.deltaCount === this.size
  }

  getConstantScalarDelta(): number | undefined {
    return this.constantDelta !== undefined &&
      this.scalarDeltaCount === this.size &&
      this.deltaCount === this.size &&
      this.currentResultAssigned === undefined
      ? this.constantDelta
      : undefined
  }

  hasCompleteScalarDeltas(): boolean {
    return this.size > 0 && this.deltaCount === this.size && this.scalarDeltaCount === this.size && this.currentResultAssigned === undefined
  }

  markScalarDeltaCellsValidated(): void {
    this.validatedScalarDeltaSize = this.size
  }

  markScalarDeltaCellsCleanNumber(): void {
    this.cleanScalarDeltaSize = this.size
  }

  markScalarDeltaCellsTrustedDirectScalarFormulas(): void {
    this.trustedDirectScalarDeltaSize = this.size
  }

  hasValidatedScalarDeltaCells(): boolean {
    return this.validatedScalarDeltaSize === this.size && this.hasCompleteScalarDeltas()
  }

  hasCleanScalarDeltaCells(): boolean {
    return this.cleanScalarDeltaSize === this.size && this.validatedScalarDeltaSize === this.size && this.hasCompleteScalarDeltas()
  }

  hasTrustedDirectScalarDeltaCells(): boolean {
    return this.trustedDirectScalarDeltaSize === this.size && this.hasCompleteScalarDeltas()
  }

  getScalarDeltaAt(index: number): number | undefined {
    if (this.constantDelta !== undefined && this.scalarDeltaCount === this.size && index >= 0 && index < this.size) {
      return this.constantDelta
    }
    if (this.scalarDeltaAssigned?.[index] !== true) {
      return undefined
    }
    return this.deltas?.[index]
  }

  getCellIndexAt(index: number): number {
    return (this.sharedCellIndices ?? this.cellIndices)[index]!
  }

  getCellIndicesForRead(): readonly number[] | U32 {
    return this.sharedCellIndices ?? this.cellIndices
  }

  markDirectRangeInputCovered(cellIndex: number): void {
    if (this.directRangeCoveredInputCellSet) {
      this.directRangeCoveredInputCellSet.add(cellIndex)
      return
    }
    const covered = (this.directRangeCoveredInputCellIndices ??= [])
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return
      }
    }
    covered.push(cellIndex)
    if (covered.length > 16) {
      this.directRangeCoveredInputCellSet = new Set(covered)
    }
  }

  hasCoveredDirectRangeInput(cellIndex: number): boolean {
    if (this.directRangeCoveredInputCellSet) {
      return this.directRangeCoveredInputCellSet.has(cellIndex)
    }
    const covered = this.directRangeCoveredInputCellIndices
    if (!covered) {
      return false
    }
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return true
      }
    }
    return false
  }

  markDirectFormulaInputCovered(cellIndex: number): void {
    if (this.directFormulaCoveredInputCellSet) {
      this.directFormulaCoveredInputCellSet.add(cellIndex)
      return
    }
    const covered = (this.directFormulaCoveredInputCellIndices ??= [])
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return
      }
    }
    covered.push(cellIndex)
    if (covered.length > 16) {
      this.directFormulaCoveredInputCellSet = new Set(covered)
    }
  }

  hasCoveredDirectFormulaInput(cellIndex: number): boolean {
    if (this.directFormulaCoveredInputCellSet) {
      return this.directFormulaCoveredInputCellSet.has(cellIndex)
    }
    const covered = this.directFormulaCoveredInputCellIndices
    if (!covered) {
      return false
    }
    for (let index = 0; index < covered.length; index += 1) {
      if (covered[index] === cellIndex) {
        return true
      }
    }
    return false
  }

  forEach(fn: (cellIndex: number) => void): void {
    for (let index = 0; index < this.size; index += 1) {
      fn(this.getCellIndexAt(index))
    }
  }

  forEachIndexed(fn: (cellIndex: number, index: number) => void): void {
    for (let index = 0; index < this.size; index += 1) {
      fn(this.getCellIndexAt(index), index)
    }
  }

  private findIndex(cellIndex: number): number {
    if (!this.indexByCell && this.size > 16) {
      this.materializeIndexByCell()
    }
    const mappedIndex = this.indexByCell?.get(cellIndex)
    if (mappedIndex !== undefined) {
      return mappedIndex
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return index
      }
    }
    return -1
  }

  private ensureIndex(cellIndex: number): number {
    if (!this.indexByCell && this.size > 16) {
      this.materializeIndexByCell()
    }
    if (this.indexByCell) {
      const mappedIndex = this.indexByCell.get(cellIndex)
      if (mappedIndex !== undefined) {
        return mappedIndex
      }
      this.materializeSharedCellIndices()
      const index = this.cellIndices.length
      this.indexByCell.set(cellIndex, index)
      this.cellIndices.push(cellIndex)
      return index
    }
    for (let index = 0; index < this.size; index += 1) {
      if (this.getCellIndexAt(index) === cellIndex) {
        return index
      }
    }
    this.materializeSharedCellIndices()
    const index = this.cellIndices.length
    this.cellIndices.push(cellIndex)
    if (this.size > 16) {
      this.materializeIndexByCell()
    }
    return index
  }

  private materializeIndexByCell(): void {
    this.indexByCell = new Map()
    for (let index = 0; index < this.size; index += 1) {
      this.indexByCell.set(this.getCellIndexAt(index), index)
    }
  }

  private appendPreparedDeltas(cellIndices: readonly number[] | U32, deltas: readonly number[], kind: 'scalar' | undefined): void {
    this.prepareForBulkDeltaAppend()
    for (let index = 0; index < cellIndices.length; index += 1) {
      this.addPreparedDelta(cellIndices[index]!, deltas[index]!, kind)
    }
  }

  private appendPreparedConstantDelta(cellIndices: readonly number[] | U32, delta: number, kind: 'scalar' | undefined): void {
    this.prepareForBulkDeltaAppend()
    for (let index = 0; index < cellIndices.length; index += 1) {
      this.addPreparedDelta(cellIndices[index]!, delta, kind)
    }
  }

  private prepareForBulkDeltaAppend(): void {
    this.materializeConstantDeltas()
    this.materializeSharedCellIndices()
    this.materializeIndexByCell()
    this.deltas ??= []
    this.deltaAssigned ??= []
    this.scalarDeltaAssigned ??= []
  }

  private addPreparedDelta(cellIndex: number, delta: number, kind: 'scalar' | undefined): void {
    const mappedIndex = this.indexByCell?.get(cellIndex)
    if (mappedIndex === undefined) {
      const index = this.cellIndices.length
      this.indexByCell!.set(cellIndex, index)
      this.cellIndices.push(cellIndex)
      this.deltas![index] = delta
      this.deltaAssigned![index] = true
      this.deltaCount += 1
      if (kind === 'scalar') {
        this.scalarDeltaAssigned![index] = true
        this.scalarDeltaCount += 1
      }
      return
    }
    if (!this.deltaAssigned![mappedIndex]) {
      this.deltaAssigned![mappedIndex] = true
      this.deltaCount += 1
      this.deltas![mappedIndex] = delta
      if (kind === 'scalar') {
        this.scalarDeltaAssigned![mappedIndex] = true
        this.scalarDeltaCount += 1
      }
      return
    }
    this.deltas![mappedIndex] = (this.deltas![mappedIndex] ?? 0) + delta
    if (kind !== 'scalar' && this.scalarDeltaAssigned![mappedIndex]) {
      this.scalarDeltaAssigned![mappedIndex] = false
      this.scalarDeltaCount -= 1
    }
  }

  private materializeSharedCellIndices(): void {
    const sharedCellIndices = this.sharedCellIndices
    if (!sharedCellIndices) {
      return
    }
    this.cellIndices.length = sharedCellIndices.length
    for (let index = 0; index < sharedCellIndices.length; index += 1) {
      this.cellIndices[index] = sharedCellIndices[index]!
    }
    this.sharedCellIndices = undefined
  }

  private materializeConstantDeltas(): void {
    if (this.constantDelta === undefined) {
      return
    }
    const delta = this.constantDelta
    const scalarDeltaCount = this.scalarDeltaCount
    const size = this.size
    this.constantDelta = undefined
    this.deltas = []
    this.deltaAssigned = []
    this.scalarDeltaAssigned = scalarDeltaCount === size ? [] : undefined
    for (let index = 0; index < size; index += 1) {
      this.deltas[index] = delta
      this.deltaAssigned[index] = true
      if (this.scalarDeltaAssigned) {
        this.scalarDeltaAssigned[index] = true
      }
    }
    this.deltaCount = size
    this.scalarDeltaCount = this.scalarDeltaAssigned ? size : 0
  }
}
