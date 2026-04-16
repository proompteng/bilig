import { MAX_COLS, MAX_ROWS, MAX_WASM_RANGE_CELLS, type RangeIndex } from '@bilig/protocol'
import type { CellAddress, CellRangeAddress, RangeAddress } from '@bilig/formula'
import { makeCellEntity, makeRangeEntity } from './entity-ids.js'
import { EdgeArena, type EdgeSlice } from './edge-arena.js'

export interface RangeDescriptor {
  index: RangeIndex
  sheetId: number
  kind: RangeAddress['kind']
  row1: number
  col1: number
  row2: number
  col2: number
  membersOffset: number
  membersLength: number
  formulaMembersOffset: number
  formulaMembersLength: number
  dependencySourcesOffset: number
  dependencySourcesLength: number
  refCount: number
  dynamic: boolean
  parentRangeIndex: RangeIndex | undefined
}

export interface RangeMaterializer {
  ensureCell(sheetId: number, row: number, col: number): number
  forEachSheetCell(sheetId: number, fn: (cellIndex: number, row: number, col: number) => void): void
  isFormulaCell?(cellIndex: number): boolean
}

interface DynamicRangeIndex {
  sheetId: number
  rangeIndex: number
}

export interface RegisteredCellRange {
  rangeIndex: number
  cellRange: CellRangeAddress
  materialized: boolean
}

export class RangeRegistry {
  private readonly descriptors: RangeDescriptor[] = []
  private readonly byKey = new Map<string, RangeIndex>()
  private readonly dynamicBySheet = new Map<number, DynamicRangeIndex[]>()
  private readonly members = new EdgeArena()
  private readonly memberSlices: EdgeSlice[] = []
  private readonly formulaMembers = new EdgeArena()
  private readonly formulaMemberSlices: EdgeSlice[] = []
  private readonly dependencySources = new EdgeArena()
  private readonly dependencySourceSlices: EdgeSlice[] = []

  get size(): number {
    return this.descriptors.length
  }

  reset(): void {
    this.descriptors.length = 0
    this.byKey.clear()
    this.dynamicBySheet.clear()
    this.members.reset()
    this.memberSlices.length = 0
    this.formulaMembers.reset()
    this.formulaMemberSlices.length = 0
    this.dependencySources.reset()
    this.dependencySourceSlices.length = 0
  }

  intern(sheetId: number, range: RangeAddress, materializer: RangeMaterializer): RegisteredCellRange {
    const descriptorKey = keyForRange(sheetId, range)
    const existingIndex = this.byKey.get(descriptorKey)
    if (existingIndex !== undefined) {
      const existing = this.descriptors[existingIndex]!
      existing.refCount += 1
      return {
        rangeIndex: existingIndex,
        cellRange: toCellRange(range),
        materialized: false,
      }
    }

    const cellRange = toCellRange(range)
    const dynamic = range.kind !== 'cells'
    const descriptor: RangeDescriptor = {
      index: this.descriptors.length,
      sheetId,
      kind: range.kind,
      row1: cellRange.start.row,
      col1: cellRange.start.col,
      row2: cellRange.end.row,
      col2: cellRange.end.col,
      membersOffset: 0,
      membersLength: 0,
      formulaMembersOffset: 0,
      formulaMembersLength: 0,
      dependencySourcesOffset: 0,
      dependencySourcesLength: 0,
      refCount: 1,
      dynamic,
      parentRangeIndex: undefined,
    }

    const sourceReuse =
      range.kind === 'cells' ? resolveCellRangeDependencyReuse(this.descriptors, this.byKey, sheetId, cellRange) : undefined
    const isFormulaCell =
      materializer.isFormulaCell === undefined ? undefined : (cellIndex: number): boolean => materializer.isFormulaCell!(cellIndex)
    const memberIndices =
      range.kind === 'cells'
        ? materializeBoundedMembers(sheetId, cellRange, materializer, this, sourceReuse)
        : materializeDynamicMembers(sheetId, cellRange, range.kind, materializer)
    const dependencySourceEntities =
      range.kind === 'cells'
        ? materializeCellRangeDependencySources(sourceReuse, memberIndices)
        : materializeDynamicDependencySources(memberIndices)
    const formulaMemberIndices = materializeFormulaMembers(
      memberIndices,
      isFormulaCell,
      sourceReuse ? this.getFormulaMembersView(sourceReuse.parentRangeIndex) : undefined,
      sourceReuse ? this.getMembersView(sourceReuse.parentRangeIndex).length : 0,
    )
    const memberSlice = this.members.replace(this.members.empty(), memberIndices)
    const formulaMemberSlice = this.formulaMembers.replace(this.formulaMembers.empty(), formulaMemberIndices)
    const dependencySourceSlice = this.dependencySources.replace(this.dependencySources.empty(), dependencySourceEntities)
    this.memberSlices[descriptor.index] = memberSlice
    this.formulaMemberSlices[descriptor.index] = formulaMemberSlice
    this.dependencySourceSlices[descriptor.index] = dependencySourceSlice
    syncDescriptorMembers(descriptor, memberSlice)
    syncDescriptorFormulaMembers(descriptor, formulaMemberSlice)
    syncDescriptorDependencySources(descriptor, dependencySourceSlice)
    if (sourceReuse) {
      descriptor.parentRangeIndex = sourceReuse.parentRangeIndex
      this.descriptors[sourceReuse.parentRangeIndex]!.refCount += 1
    }
    this.descriptors.push(descriptor)
    this.byKey.set(descriptorKey, descriptor.index)

    if (dynamic) {
      const entries = this.dynamicBySheet.get(sheetId) ?? []
      entries.push({ sheetId, rangeIndex: descriptor.index })
      this.dynamicBySheet.set(sheetId, entries)
    }

    return {
      rangeIndex: descriptor.index,
      cellRange,
      materialized: true,
    }
  }

  release(rangeIndex: RangeIndex): { removed: boolean; members: Uint32Array } {
    const descriptor = this.descriptors[rangeIndex]
    if (!descriptor) {
      return { removed: false, members: new Uint32Array() }
    }
    const memberSlice = this.memberSlices[rangeIndex] ?? this.members.empty()

    descriptor.refCount -= 1
    if (descriptor.refCount > 0) {
      return { removed: false, members: this.members.read(memberSlice) }
    }

    this.byKey.delete(keyForDescriptor(descriptor))
    const members = this.members.read(memberSlice)
    this.members.free(memberSlice)
    this.memberSlices[rangeIndex] = this.members.empty()
    const formulaMemberSlice = this.formulaMemberSlices[rangeIndex] ?? this.formulaMembers.empty()
    this.formulaMembers.free(formulaMemberSlice)
    this.formulaMemberSlices[rangeIndex] = this.formulaMembers.empty()
    const dependencySourceSlice = this.dependencySourceSlices[rangeIndex] ?? this.dependencySources.empty()
    this.dependencySources.free(dependencySourceSlice)
    this.dependencySourceSlices[rangeIndex] = this.dependencySources.empty()
    syncDescriptorMembers(descriptor, this.memberSlices[rangeIndex])
    syncDescriptorFormulaMembers(descriptor, this.formulaMemberSlices[rangeIndex])
    syncDescriptorDependencySources(descriptor, this.dependencySourceSlices[rangeIndex])
    if (descriptor.dynamic) {
      const dynamic = this.dynamicBySheet.get(descriptor.sheetId)
      if (dynamic) {
        this.dynamicBySheet.set(
          descriptor.sheetId,
          dynamic.filter((entry) => entry.rangeIndex !== rangeIndex),
        )
      }
    }
    const parentRangeIndex = descriptor.parentRangeIndex
    descriptor.parentRangeIndex = undefined
    descriptor.refCount = 0
    if (parentRangeIndex !== undefined) {
      this.release(parentRangeIndex)
    }
    return { removed: true, members }
  }

  getDescriptor(rangeIndex: RangeIndex): RangeDescriptor {
    const descriptor = this.descriptors[rangeIndex]
    if (!descriptor) {
      throw new Error(`Unknown range index: ${rangeIndex}`)
    }
    return descriptor
  }

  getMembers(rangeIndex: RangeIndex): Uint32Array {
    return this.members.read(this.memberSlices[rangeIndex] ?? this.members.empty())
  }

  getDependencySourceEntities(rangeIndex: RangeIndex): Uint32Array {
    return this.dependencySources.read(this.dependencySourceSlices[rangeIndex] ?? this.dependencySources.empty())
  }

  getFormulaMembers(rangeIndex: RangeIndex): Uint32Array {
    return this.formulaMembers.read(this.formulaMemberSlices[rangeIndex] ?? this.formulaMembers.empty())
  }

  getMembersView(rangeIndex: RangeIndex): Uint32Array {
    return this.members.readView(this.memberSlices[rangeIndex] ?? this.members.empty())
  }

  getFormulaMembersView(rangeIndex: RangeIndex): Uint32Array {
    return this.formulaMembers.readView(this.formulaMemberSlices[rangeIndex] ?? this.formulaMembers.empty())
  }

  getMemberPoolView(): Uint32Array {
    return this.members.view()
  }

  addDynamicMember(sheetId: number, row: number, col: number, cellIndex: number, includeAsFormulaMember = false): RangeIndex[] {
    const entries = this.dynamicBySheet.get(sheetId)
    if (!entries || entries.length === 0) {
      return []
    }

    const matched: RangeIndex[] = []
    for (let index = 0; index < entries.length; index += 1) {
      const rangeIndex = entries[index]!.rangeIndex
      const descriptor = this.descriptors[rangeIndex]!
      if (!matchesDynamicRange(descriptor, row, col)) {
        continue
      }
      const currentSlice = this.memberSlices[rangeIndex] ?? this.members.empty()
      const nextMembers = this.members.appendUnique(currentSlice, cellIndex)
      if (nextMembers.ptr !== currentSlice.ptr || nextMembers.len !== currentSlice.len) {
        this.memberSlices[rangeIndex] = nextMembers
        syncDescriptorMembers(descriptor, nextMembers)
        if (includeAsFormulaMember) {
          const currentFormulaSlice = this.formulaMemberSlices[rangeIndex] ?? this.formulaMembers.empty()
          const nextFormulaMembers = this.formulaMembers.appendUnique(currentFormulaSlice, cellIndex)
          this.formulaMemberSlices[rangeIndex] = nextFormulaMembers
          syncDescriptorFormulaMembers(descriptor, nextFormulaMembers)
        }
        matched.push(rangeIndex)
      }
    }
    return matched
  }

  expandToCells(rangeIndex: RangeIndex): Uint32Array {
    return this.getMembersView(rangeIndex)
  }

  refresh(
    rangeIndex: RangeIndex,
    materializer: RangeMaterializer,
  ): { oldDependencySources: Uint32Array; newDependencySources: Uint32Array } {
    const descriptor = this.getDescriptor(rangeIndex)
    const oldDependencySources = this.getDependencySourceEntities(rangeIndex)
    let memberIndices: Uint32Array
    let formulaMemberIndices: Uint32Array
    let dependencySourceEntities: Uint32Array
    if (descriptor.kind === 'cells') {
      const range = descriptorToCellRangeAddress(descriptor)
      const sourceReuse = resolveCellRangeDependencyReuse(this.descriptors, this.byKey, descriptor.sheetId, range)
      const isFormulaCell =
        materializer.isFormulaCell === undefined ? undefined : (cellIndex: number): boolean => materializer.isFormulaCell!(cellIndex)
      memberIndices = materializeBoundedMembers(descriptor.sheetId, range, materializer, this, sourceReuse)
      dependencySourceEntities = materializeCellRangeDependencySources(sourceReuse, memberIndices)
      formulaMemberIndices = materializeFormulaMembers(
        memberIndices,
        isFormulaCell,
        sourceReuse ? this.getFormulaMembersView(sourceReuse.parentRangeIndex) : undefined,
        sourceReuse ? this.getMembersView(sourceReuse.parentRangeIndex).length : 0,
      )
    } else {
      const range = descriptorToCellRangeAddress(descriptor)
      memberIndices = materializeDynamicMembers(descriptor.sheetId, range, descriptor.kind, materializer)
      dependencySourceEntities = materializeDynamicDependencySources(memberIndices)
      const isFormulaCell =
        materializer.isFormulaCell === undefined ? undefined : (cellIndex: number): boolean => materializer.isFormulaCell!(cellIndex)
      formulaMemberIndices = materializeFormulaMembers(memberIndices, isFormulaCell)
    }
    const memberSlice = this.members.replace(this.memberSlices[rangeIndex] ?? this.members.empty(), memberIndices)
    const formulaMemberSlice = this.formulaMembers.replace(
      this.formulaMemberSlices[rangeIndex] ?? this.formulaMembers.empty(),
      formulaMemberIndices,
    )
    const dependencySourceSlice = this.dependencySources.replace(
      this.dependencySourceSlices[rangeIndex] ?? this.dependencySources.empty(),
      dependencySourceEntities,
    )
    this.memberSlices[rangeIndex] = memberSlice
    this.formulaMemberSlices[rangeIndex] = formulaMemberSlice
    this.dependencySourceSlices[rangeIndex] = dependencySourceSlice
    syncDescriptorMembers(descriptor, memberSlice)
    syncDescriptorFormulaMembers(descriptor, formulaMemberSlice)
    syncDescriptorDependencySources(descriptor, dependencySourceSlice)
    return {
      oldDependencySources,
      newDependencySources: this.getDependencySourceEntities(rangeIndex),
    }
  }
}

function materializeBoundedMembers(
  sheetId: number,
  range: CellRangeAddress,
  materializer: RangeMaterializer,
  registry: RangeRegistry,
  reuse:
    | {
        parentRangeIndex: RangeIndex
        tailStartRow: number
        tailEndRow: number
        tailStartCol: number
        tailEndCol: number
      }
    | undefined,
): Uint32Array {
  if (reuse) {
    const tailCells = materializeBoundedTailCells(sheetId, reuse, materializer)
    const parentMembers = registry.getMembersView(reuse.parentRangeIndex)
    const members = new Uint32Array(parentMembers.length + tailCells.length)
    members.set(parentMembers, 0)
    members.set(tailCells, parentMembers.length)
    return members
  }
  const rowCount = range.end.row - range.start.row + 1
  const colCount = range.end.col - range.start.col + 1
  const memberCount = rowCount * colCount
  if (memberCount > MAX_WASM_RANGE_CELLS) {
    throw new Error(`Bounded range exceeds fast-path cap: ${memberCount}`)
  }
  const members = new Uint32Array(memberCount)
  let cursor = 0
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    for (let col = range.start.col; col <= range.end.col; col += 1) {
      members[cursor] = materializer.ensureCell(sheetId, row, col)
      cursor += 1
    }
  }
  return members
}

function materializeDynamicMembers(
  sheetId: number,
  range: CellRangeAddress,
  kind: RangeAddress['kind'],
  materializer: RangeMaterializer,
): Uint32Array {
  let matchCount = 0
  materializer.forEachSheetCell(sheetId, (_cellIndex, row, col) => {
    if (kind === 'rows') {
      if (row >= range.start.row && row <= range.end.row) {
        matchCount += 1
      }
      return
    }
    if (col >= range.start.col && col <= range.end.col) {
      matchCount += 1
    }
  })
  const matches = new Uint32Array(matchCount)
  let cursor = 0
  materializer.forEachSheetCell(sheetId, (cellIndex, row, col) => {
    if (kind === 'rows') {
      if (row >= range.start.row && row <= range.end.row) {
        matches[cursor] = cellIndex
        cursor += 1
      }
      return
    }
    if (col >= range.start.col && col <= range.end.col) {
      matches[cursor] = cellIndex
      cursor += 1
    }
  })
  return matches.toSorted()
}

function materializeDynamicDependencySources(memberIndices: Uint32Array): Uint32Array {
  const dependencySources = new Uint32Array(memberIndices.length)
  for (let index = 0; index < memberIndices.length; index += 1) {
    dependencySources[index] = makeCellEntity(memberIndices[index]!)
  }
  return dependencySources
}

function materializeFormulaMembers(
  memberIndices: Uint32Array,
  isFormulaCell: ((cellIndex: number) => boolean) | undefined,
  parentFormulaMembers?: Uint32Array,
  parentMemberCount = 0,
): Uint32Array {
  if (!isFormulaCell) {
    return parentFormulaMembers ? Uint32Array.from(parentFormulaMembers) : new Uint32Array()
  }
  let appendedFormulaCount = 0
  const startIndex = parentMemberCount
  for (let index = startIndex; index < memberIndices.length; index += 1) {
    if (isFormulaCell(memberIndices[index]!)) {
      appendedFormulaCount += 1
    }
  }
  const parentCount = parentFormulaMembers?.length ?? 0
  const formulaMembers = new Uint32Array(parentCount + appendedFormulaCount)
  if (parentFormulaMembers && parentCount > 0) {
    formulaMembers.set(parentFormulaMembers, 0)
  }
  let cursor = parentCount
  for (let index = startIndex; index < memberIndices.length; index += 1) {
    const cellIndex = memberIndices[index]!
    if (!isFormulaCell(cellIndex)) {
      continue
    }
    formulaMembers[cursor] = cellIndex
    cursor += 1
  }
  return formulaMembers
}

function materializeCellRangeDependencySources(
  reuse:
    | {
        parentRangeIndex: RangeIndex
        tailStartRow: number
        tailEndRow: number
        tailStartCol: number
        tailEndCol: number
      }
    | undefined,
  memberIndices: Uint32Array,
): Uint32Array {
  if (reuse) {
    const tailMemberCount = (reuse.tailEndRow - reuse.tailStartRow + 1) * (reuse.tailEndCol - reuse.tailStartCol + 1)
    const dependencySources = new Uint32Array(tailMemberCount + 1)
    dependencySources[0] = makeRangeEntity(reuse.parentRangeIndex)
    const tailStartIndex = memberIndices.length - tailMemberCount
    for (let index = 0; index < tailMemberCount; index += 1) {
      dependencySources[index + 1] = makeCellEntity(memberIndices[tailStartIndex + index]!)
    }
    return dependencySources
  }
  return materializeDynamicDependencySources(memberIndices)
}

function resolveCellRangeDependencyReuse(
  descriptors: readonly RangeDescriptor[],
  byKey: ReadonlyMap<string, RangeIndex>,
  sheetId: number,
  range: CellRangeAddress,
):
  | {
      parentRangeIndex: RangeIndex
      tailStartRow: number
      tailEndRow: number
      tailStartCol: number
      tailEndCol: number
    }
  | undefined {
  const width = range.end.col - range.start.col + 1
  const height = range.end.row - range.start.row + 1
  if (width === 1 && height > 1) {
    const prefixKey = keyForRange(sheetId, {
      kind: 'cells',
      start: range.start,
      end: { ...range.end, row: range.end.row - 1 },
    })
    const prefixRangeIndex = byKey.get(prefixKey)
    if (prefixRangeIndex !== undefined && descriptors[prefixRangeIndex]?.refCount) {
      return {
        parentRangeIndex: prefixRangeIndex,
        tailStartRow: range.end.row,
        tailEndRow: range.end.row,
        tailStartCol: range.start.col,
        tailEndCol: range.end.col,
      }
    }
  }
  if (height === 1 && width > 1) {
    const prefixKey = keyForRange(sheetId, {
      kind: 'cells',
      start: range.start,
      end: { ...range.end, col: range.end.col - 1 },
    })
    const prefixRangeIndex = byKey.get(prefixKey)
    if (prefixRangeIndex !== undefined && descriptors[prefixRangeIndex]?.refCount) {
      return {
        parentRangeIndex: prefixRangeIndex,
        tailStartRow: range.start.row,
        tailEndRow: range.end.row,
        tailStartCol: range.end.col,
        tailEndCol: range.end.col,
      }
    }
  }
  return undefined
}

function materializeBoundedTailCells(
  sheetId: number,
  reuse: {
    tailStartRow: number
    tailEndRow: number
    tailStartCol: number
    tailEndCol: number
  },
  materializer: RangeMaterializer,
): Uint32Array {
  const tailRowCount = reuse.tailEndRow - reuse.tailStartRow + 1
  const tailColCount = reuse.tailEndCol - reuse.tailStartCol + 1
  const tailCells = new Uint32Array(tailRowCount * tailColCount)
  let cursor = 0
  for (let row = reuse.tailStartRow; row <= reuse.tailEndRow; row += 1) {
    for (let col = reuse.tailStartCol; col <= reuse.tailEndCol; col += 1) {
      tailCells[cursor] = materializer.ensureCell(sheetId, row, col)
      cursor += 1
    }
  }
  return tailCells
}

function matchesDynamicRange(descriptor: RangeDescriptor, row: number, col: number): boolean {
  if (descriptor.kind === 'rows') {
    return row >= descriptor.row1 && row <= descriptor.row2
  }
  if (descriptor.kind === 'cols') {
    return col >= descriptor.col1 && col <= descriptor.col2
  }
  return false
}

function syncDescriptorMembers(descriptor: RangeDescriptor, slice: EdgeSlice): void {
  descriptor.membersOffset = slice.ptr < 0 ? 0 : slice.ptr
  descriptor.membersLength = slice.len
}

function syncDescriptorFormulaMembers(descriptor: RangeDescriptor, slice: EdgeSlice): void {
  descriptor.formulaMembersOffset = slice.ptr < 0 ? 0 : slice.ptr
  descriptor.formulaMembersLength = slice.len
}

function syncDescriptorDependencySources(descriptor: RangeDescriptor, slice: EdgeSlice): void {
  descriptor.dependencySourcesOffset = slice.ptr < 0 ? 0 : slice.ptr
  descriptor.dependencySourcesLength = slice.len
}

function toCellRange(range: RangeAddress): CellRangeAddress {
  if (range.kind === 'cells') {
    return range
  }
  if (range.kind === 'rows') {
    const cellRange: CellRangeAddress = {
      kind: 'cells',
      start: toCellLikeAddress(range.sheetName, range.start.row, 0),
      end: toCellLikeAddress(range.sheetName, range.end.row, MAX_COLS - 1),
    }
    if (range.sheetName !== undefined) {
      cellRange.sheetName = range.sheetName
    }
    return cellRange
  }
  const cellRange: CellRangeAddress = {
    kind: 'cells',
    start: toCellLikeAddress(range.sheetName, 0, range.start.col),
    end: toCellLikeAddress(range.sheetName, MAX_ROWS - 1, range.end.col),
  }
  if (range.sheetName !== undefined) {
    cellRange.sheetName = range.sheetName
  }
  return cellRange
}

function descriptorToCellRangeAddress(descriptor: RangeDescriptor): CellRangeAddress {
  return {
    kind: 'cells',
    start: toCellLikeAddress(undefined, descriptor.row1, descriptor.col1),
    end: toCellLikeAddress(undefined, descriptor.row2, descriptor.col2),
  }
}

function toCellLikeAddress(sheetName: string | undefined, row: number, col: number): CellAddress {
  const address: CellAddress = {
    row,
    col,
    text: '',
  }
  if (sheetName !== undefined) {
    address.sheetName = sheetName
  }
  return address
}

function keyForRange(sheetId: number, range: RangeAddress): string {
  if (range.kind === 'cells') {
    return `cells:${sheetId}:${range.start.row}:${range.start.col}:${range.end.row}:${range.end.col}`
  }
  if (range.kind === 'rows') {
    return `rows:${sheetId}:${range.start.row}:${range.end.row}`
  }
  return `cols:${sheetId}:${range.start.col}:${range.end.col}`
}

function keyForDescriptor(descriptor: RangeDescriptor): string {
  if (descriptor.kind === 'cells') {
    return `cells:${descriptor.sheetId}:${descriptor.row1}:${descriptor.col1}:${descriptor.row2}:${descriptor.col2}`
  }
  if (descriptor.kind === 'rows') {
    return `rows:${descriptor.sheetId}:${descriptor.row1}:${descriptor.row2}`
  }
  return `cols:${descriptor.sheetId}:${descriptor.col1}:${descriptor.col2}`
}
