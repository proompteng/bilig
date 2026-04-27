import type { StructuralAxisTransform } from '@bilig/formula'

export type FormulaFamilyId = number
export type FormulaFamilyRunId = number

export type FormulaFamilyRunAxis = 'row' | 'column'

export interface FormulaFamilyKey {
  readonly sheetId: number
  readonly templateId: number
  readonly shapeKey: string
}

export interface FormulaFamilyMember {
  readonly cellIndex: number
  readonly row: number
  readonly col: number
}

export interface FormulaFamilyMemberRun {
  readonly id: FormulaFamilyRunId
  readonly axis: FormulaFamilyRunAxis
  readonly fixedIndex: number
  readonly start: number
  readonly end: number
  readonly step: number
  readonly cellIndices: readonly number[]
}

export interface FormulaFamily {
  readonly id: FormulaFamilyId
  readonly sheetId: number
  readonly templateId: number
  readonly shapeKey: string
  readonly runs: readonly FormulaFamilyMemberRun[]
}

export interface FormulaFamilyMembership {
  readonly familyId: FormulaFamilyId
  readonly runId: FormulaFamilyRunId
}

export interface FormulaFamilyStats {
  readonly familyCount: number
  readonly runCount: number
  readonly memberCount: number
}

export interface FormulaFamilyStructuralSourceTransform {
  readonly ownerSheetName: string
  readonly targetSheetName: string
  readonly transform: StructuralAxisTransform
  readonly preservesValue: boolean
}

export interface FormulaFamilyStructuralSourceTransformEntry {
  readonly cellIndices: readonly number[]
  readonly transform: FormulaFamilyStructuralSourceTransform
}

export interface FormulaFamilyStore {
  readonly upsertFormula: (args: FormulaFamilyKey & FormulaFamilyMember) => FormulaFamilyMembership
  readonly unregisterFormula: (cellIndex: number) => boolean
  readonly getMembership: (cellIndex: number) => FormulaFamilyMembership | undefined
  readonly countSheetMembers: (sheetId: number) => number
  readonly forEachFamily: (fn: (family: FormulaFamily) => void) => void
  readonly setStructuralSourceTransform: (familyId: FormulaFamilyId, transform: FormulaFamilyStructuralSourceTransform) => void
  readonly getStructuralSourceTransform: (cellIndex: number) => FormulaFamilyStructuralSourceTransform | undefined
  readonly consumeStructuralSourceTransforms: () => FormulaFamilyStructuralSourceTransformEntry[]
  readonly getStats: () => FormulaFamilyStats
  readonly listFamilies: () => FormulaFamily[]
  readonly invalidateSheet: (sheetId: number) => void
  readonly applyStructuralInvalidation: (args: {
    readonly sheetId: number
    readonly axis: 'row' | 'column'
    readonly start: number
    readonly end: number
  }) => void
  readonly clear: () => void
}

interface MutableFormulaFamilyMemberRun {
  id: FormulaFamilyRunId
  axis: FormulaFamilyRunAxis
  fixedIndex: number
  start: number
  end: number
  step: number
  cellIndices: number[]
}

interface MutableFormulaFamily {
  id: FormulaFamilyId
  sheetId: number
  templateId: number
  shapeKey: string
  key: string
  runs: MutableFormulaFamilyMemberRun[]
  rowRunsByFixedIndex: Map<number, MutableFormulaFamilyMemberRun[]>
  columnRunsByFixedIndex: Map<number, MutableFormulaFamilyMemberRun[]>
  singletonRunsByRow: Map<number, MutableFormulaFamilyMemberRun[]>
}

interface FormulaFamilyCellRecord extends FormulaFamilyKey, FormulaFamilyMember {}

function keyForFormulaFamily(args: FormulaFamilyKey): string {
  return `${args.sheetId}\t${args.templateId}\t${args.shapeKey}`
}

export function createFormulaFamilyStore(): FormulaFamilyStore {
  const familiesById = new Map<FormulaFamilyId, MutableFormulaFamily>()
  const familyIdByKey = new Map<string, FormulaFamilyId>()
  const cellRecords = new Map<number, FormulaFamilyCellRecord>()
  const memberships = new Map<number, FormulaFamilyMembership>()
  const sheetMemberCounts = new Map<number, number>()
  const structuralSourceTransforms = new Map<FormulaFamilyId, FormulaFamilyStructuralSourceTransform>()
  let nextFamilyId = 1
  let nextRunId = 1

  const getOrCreateFamily = (args: FormulaFamilyKey): MutableFormulaFamily => {
    const key = keyForFormulaFamily(args)
    const existingId = familyIdByKey.get(key)
    if (existingId !== undefined) {
      return familiesById.get(existingId)!
    }
    const family: MutableFormulaFamily = {
      id: nextFamilyId,
      sheetId: args.sheetId,
      templateId: args.templateId,
      shapeKey: args.shapeKey,
      key,
      runs: [],
      rowRunsByFixedIndex: new Map(),
      columnRunsByFixedIndex: new Map(),
      singletonRunsByRow: new Map(),
    }
    nextFamilyId += 1
    familiesById.set(family.id, family)
    familyIdByKey.set(key, family.id)
    return family
  }

  const makeRun = (
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
  ): MutableFormulaFamilyMemberRun => {
    const sorted = sortRunMembers(axis, members)
    const first = sorted[0]!
    const last = sorted[sorted.length - 1]!
    return {
      id: nextRunId++,
      axis,
      fixedIndex,
      start: axis === 'row' ? first.row : first.col,
      end: axis === 'row' ? last.row : last.col,
      step: inferRunStep(axis, sorted),
      cellIndices: sorted.map((member) => member.cellIndex),
    }
  }

  const replaceRunWithMembers = (
    family: MutableFormulaFamily,
    runIndex: number,
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
    memberCellIndex: number,
  ): FormulaFamilyMembership => {
    const previousRun = family.runs[runIndex]
    if (previousRun) {
      unindexRun(family, previousRun)
    }
    const run = makeRun(axis, fixedIndex, members)
    family.runs.splice(runIndex, 1, run)
    indexRun(family, run)
    run.cellIndices.forEach((cellIndex) => {
      memberships.set(cellIndex, { familyId: family.id, runId: run.id })
    })
    return memberships.get(memberCellIndex)!
  }

  const appendMemberToRun = (
    family: MutableFormulaFamily,
    run: MutableFormulaFamilyMemberRun,
    member: FormulaFamilyMember,
  ): FormulaFamilyMembership => {
    const wasSingleton = run.cellIndices.length === 1
    if (wasSingleton) {
      removeMapArrayValue(family.singletonRunsByRow, runRowStart(run), run)
    }
    const memberIndex = run.axis === 'row' ? member.row : member.col
    if (memberIndex < run.start) {
      run.start = memberIndex
      run.cellIndices.unshift(member.cellIndex)
    } else {
      run.end = Math.max(run.end, memberIndex)
      run.cellIndices.push(member.cellIndex)
    }
    const membership = { familyId: family.id, runId: run.id }
    memberships.set(member.cellIndex, membership)
    return membership
  }

  const splitRunAfterRemoval = (family: MutableFormulaFamily, runIndex: number, removedCellIndex: number): void => {
    const run = family.runs[runIndex]
    if (!run) {
      return
    }
    const remainingMembers = run.cellIndices
      .filter((cellIndex) => cellIndex !== removedCellIndex)
      .flatMap((cellIndex): FormulaFamilyMember[] => {
        const record = cellRecords.get(cellIndex)
        return record ? [{ cellIndex, row: record.row, col: record.col }] : []
      })
    unindexRun(family, run)
    family.runs.splice(runIndex, 1)
    if (remainingMembers.length === 0) {
      return
    }
    const groups = groupRunMembersByStep(run.axis, remainingMembers, run.step)
    groups.forEach((group, offset) => {
      const nextRun = makeRun(run.axis, run.fixedIndex, group)
      family.runs.splice(runIndex + offset, 0, nextRun)
      indexRun(family, nextRun)
      nextRun.cellIndices.forEach((cellIndex) => {
        memberships.set(cellIndex, { familyId: family.id, runId: nextRun.id })
      })
    })
  }

  const unregisterFormula = (cellIndex: number): boolean => {
    const membership = memberships.get(cellIndex)
    const record = cellRecords.get(cellIndex)
    if (!membership || !record) {
      return false
    }
    memberships.delete(cellIndex)
    cellRecords.delete(cellIndex)
    const sheetMemberCount = sheetMemberCounts.get(record.sheetId) ?? 0
    if (sheetMemberCount <= 1) {
      sheetMemberCounts.delete(record.sheetId)
    } else {
      sheetMemberCounts.set(record.sheetId, sheetMemberCount - 1)
    }
    const family = familiesById.get(membership.familyId)
    if (!family) {
      return true
    }
    const runIndex = family.runs.findIndex((run) => run.id === membership.runId)
    splitRunAfterRemoval(family, runIndex, cellIndex)
    if (family.runs.length === 0) {
      familiesById.delete(family.id)
      familyIdByKey.delete(family.key)
      structuralSourceTransforms.delete(family.id)
    }
    return true
  }

  return {
    upsertFormula(args) {
      unregisterFormula(args.cellIndex)
      const family = getOrCreateFamily(args)
      const member: FormulaFamilyMember = { cellIndex: args.cellIndex, row: args.row, col: args.col }
      cellRecords.set(args.cellIndex, { ...args })
      sheetMemberCounts.set(args.sheetId, (sheetMemberCounts.get(args.sheetId) ?? 0) + 1)

      for (const run of candidateRunsForMember(family, member)) {
        const runIndex = family.runs.indexOf(run)
        if (runIndex < 0) {
          continue
        }
        const maybeMembership = tryMergeRun(family, runIndex, run, member, appendMemberToRun, replaceRunWithMembers, cellRecords)
        if (maybeMembership) {
          return maybeMembership
        }
      }
      for (let runIndex = 0; runIndex < family.runs.length; runIndex += 1) {
        const run = family.runs[runIndex]!
        const maybeMembership = tryMergeRun(family, runIndex, run, member, appendMemberToRun, replaceRunWithMembers, cellRecords)
        if (maybeMembership) {
          return maybeMembership
        }
      }

      const run = makeRun('row', args.col, [member])
      family.runs.push(run)
      indexRun(family, run)
      const membership = { familyId: family.id, runId: run.id }
      memberships.set(args.cellIndex, membership)
      return membership
    },
    unregisterFormula,
    getMembership(cellIndex) {
      return memberships.get(cellIndex)
    },
    countSheetMembers(sheetId) {
      return sheetMemberCounts.get(sheetId) ?? 0
    },
    forEachFamily(fn) {
      familiesById.forEach((family) => {
        fn({
          id: family.id,
          sheetId: family.sheetId,
          templateId: family.templateId,
          shapeKey: family.shapeKey,
          runs: family.runs,
        })
      })
    },
    setStructuralSourceTransform(familyId, transform) {
      if (familiesById.has(familyId)) {
        structuralSourceTransforms.set(familyId, transform)
      }
    },
    getStructuralSourceTransform(cellIndex) {
      const membership = memberships.get(cellIndex)
      return membership ? structuralSourceTransforms.get(membership.familyId) : undefined
    },
    consumeStructuralSourceTransforms() {
      const entries: FormulaFamilyStructuralSourceTransformEntry[] = []
      structuralSourceTransforms.forEach((transform, familyId) => {
        const family = familiesById.get(familyId)
        if (!family) {
          return
        }
        entries.push({
          cellIndices: family.runs.flatMap((run) => run.cellIndices),
          transform,
        })
      })
      structuralSourceTransforms.clear()
      return entries
    },
    getStats() {
      let runCount = 0
      familiesById.forEach((family) => {
        runCount += family.runs.length
      })
      return {
        familyCount: familiesById.size,
        runCount,
        memberCount: memberships.size,
      }
    },
    listFamilies() {
      return [...familiesById.values()]
        .toSorted((left, right) => left.sheetId - right.sheetId || left.templateId - right.templateId || left.id - right.id)
        .map((family) => ({
          id: family.id,
          sheetId: family.sheetId,
          templateId: family.templateId,
          shapeKey: family.shapeKey,
          runs: family.runs.map((run) => ({ ...run, cellIndices: [...run.cellIndices] })),
        }))
    },
    invalidateSheet(sheetId) {
      ;[...cellRecords.values()]
        .filter((record) => record.sheetId === sheetId)
        .forEach((record) => {
          unregisterFormula(record.cellIndex)
        })
    },
    applyStructuralInvalidation(args) {
      ;[...cellRecords.values()]
        .filter((record) => {
          if (record.sheetId !== args.sheetId) {
            return false
          }
          const axisIndex = args.axis === 'row' ? record.row : record.col
          return axisIndex >= args.start && axisIndex < args.end
        })
        .forEach((record) => {
          unregisterFormula(record.cellIndex)
        })
    },
    clear() {
      familiesById.clear()
      familyIdByKey.clear()
      cellRecords.clear()
      memberships.clear()
      sheetMemberCounts.clear()
      structuralSourceTransforms.clear()
    },
  }
}

function indexRun(family: MutableFormulaFamily, run: MutableFormulaFamilyMemberRun): void {
  appendMapArray(run.axis === 'row' ? family.rowRunsByFixedIndex : family.columnRunsByFixedIndex, run.fixedIndex, run)
  if (run.cellIndices.length === 1) {
    appendMapArray(family.singletonRunsByRow, runRowStart(run), run)
  }
}

function unindexRun(family: MutableFormulaFamily, run: MutableFormulaFamilyMemberRun): void {
  removeMapArrayValue(run.axis === 'row' ? family.rowRunsByFixedIndex : family.columnRunsByFixedIndex, run.fixedIndex, run)
  if (run.cellIndices.length === 1) {
    removeMapArrayValue(family.singletonRunsByRow, runRowStart(run), run)
  }
}

function runRowStart(run: MutableFormulaFamilyMemberRun): number {
  return run.axis === 'row' ? run.start : run.fixedIndex
}

function appendMapArray<K, V>(map: Map<K, V[]>, key: K, value: V): void {
  const existing = map.get(key)
  if (existing) {
    existing.push(value)
    return
  }
  map.set(key, [value])
}

function removeMapArrayValue<K, V>(map: Map<K, V[]>, key: K, value: V): void {
  const existing = map.get(key)
  if (!existing) {
    return
  }
  const index = existing.indexOf(value)
  if (index >= 0) {
    existing.splice(index, 1)
  }
  if (existing.length === 0) {
    map.delete(key)
  }
}

function candidateRunsForMember(family: MutableFormulaFamily, member: FormulaFamilyMember): MutableFormulaFamilyMemberRun[] {
  const candidates: MutableFormulaFamilyMemberRun[] = []
  const seen = new Set<FormulaFamilyRunId>()
  const appendCandidates = (runs: readonly MutableFormulaFamilyMemberRun[] | undefined): void => {
    runs?.forEach((run) => {
      if (seen.has(run.id)) {
        return
      }
      seen.add(run.id)
      candidates.push(run)
    })
  }
  appendCandidates(family.rowRunsByFixedIndex.get(member.col))
  appendCandidates(family.columnRunsByFixedIndex.get(member.row))
  appendCandidates(family.singletonRunsByRow.get(member.row))
  return candidates
}

function tryMergeRun(
  family: MutableFormulaFamily,
  runIndex: number,
  run: MutableFormulaFamilyMemberRun,
  member: FormulaFamilyMember,
  appendMemberToRun: (
    family: MutableFormulaFamily,
    run: MutableFormulaFamilyMemberRun,
    member: FormulaFamilyMember,
  ) => FormulaFamilyMembership,
  replaceRunWithMembers: (
    family: MutableFormulaFamily,
    runIndex: number,
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
    memberCellIndex: number,
  ) => FormulaFamilyMembership,
  cellRecords: ReadonlyMap<number, FormulaFamilyCellRecord>,
): FormulaFamilyMembership | undefined {
  if (run.axis === 'row' && run.fixedIndex === member.col) {
    if (canAppendStridedRunMember(run, member.row)) {
      return appendMemberToRun(family, run, member)
    }
    const membership = tryReshapeStridedRun(family, runIndex, run, member, replaceRunWithMembers, cellRecords)
    if (membership) {
      return membership
    }
  }
  if (run.axis === 'column' && run.fixedIndex === member.row) {
    if (canAppendStridedRunMember(run, member.col)) {
      return appendMemberToRun(family, run, member)
    }
    const membership = tryReshapeStridedRun(family, runIndex, run, member, replaceRunWithMembers, cellRecords)
    if (membership) {
      return membership
    }
  }
  if (run.cellIndices.length === 1) {
    const existingRecord = cellRecords.get(run.cellIndices[0]!)
    if (!existingRecord) {
      return undefined
    }
    const existing = {
      cellIndex: existingRecord.cellIndex,
      row: existingRecord.row,
      col: existingRecord.col,
    }
    if (existing.col === member.col && existing.row !== member.row) {
      return replaceRunWithMembers(family, runIndex, 'row', member.col, [existing, member], member.cellIndex)
    }
    if (existing.row === member.row && existing.col !== member.col) {
      return replaceRunWithMembers(family, runIndex, 'column', member.row, [existing, member], member.cellIndex)
    }
  }
  return undefined
}

function inferRunStep(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): number {
  if (members.length < 2) {
    return 1
  }
  const first = members[0]!
  const second = members[1]!
  return Math.max(1, axis === 'row' ? second.row - first.row : second.col - first.col)
}

function canAppendStridedRunMember(run: MutableFormulaFamilyMemberRun, memberIndex: number): boolean {
  return memberIndex === run.start - run.step || memberIndex === run.end + run.step
}

function tryReshapeStridedRun(
  family: MutableFormulaFamily,
  runIndex: number,
  run: MutableFormulaFamilyMemberRun,
  member: FormulaFamilyMember,
  replaceRunWithMembers: (
    family: MutableFormulaFamily,
    runIndex: number,
    axis: FormulaFamilyRunAxis,
    fixedIndex: number,
    members: readonly FormulaFamilyMember[],
    memberCellIndex: number,
  ) => FormulaFamilyMembership,
  cellRecords: ReadonlyMap<number, FormulaFamilyCellRecord>,
): FormulaFamilyMembership | undefined {
  const memberIndex = run.axis === 'row' ? member.row : member.col
  if (run.cellIndices.length !== 2 || memberIndex <= run.start || memberIndex >= run.end) {
    return undefined
  }
  const existingMembers = run.cellIndices.flatMap((cellIndex): FormulaFamilyMember[] => {
    const record = cellRecords.get(cellIndex)
    return record ? [{ cellIndex, row: record.row, col: record.col }] : []
  })
  if (existingMembers.length !== run.cellIndices.length) {
    return undefined
  }
  const members = [...existingMembers, member]
  if (!isUniformRun(run.axis, members)) {
    return undefined
  }
  return replaceRunWithMembers(family, runIndex, run.axis, run.fixedIndex, members, member.cellIndex)
}

function isUniformRun(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): boolean {
  const sorted = sortRunMembers(axis, members)
  if (sorted.length < 2) {
    return true
  }
  const first = sorted[0]!
  const second = sorted[1]!
  const step = axis === 'row' ? second.row - first.row : second.col - first.col
  if (step <= 0) {
    return false
  }
  for (let index = 2; index < sorted.length; index += 1) {
    const previous = sorted[index - 1]!
    const current = sorted[index]!
    const delta = axis === 'row' ? current.row - previous.row : current.col - previous.col
    if (delta !== step) {
      return false
    }
  }
  return true
}

function sortRunMembers(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): FormulaFamilyMember[] {
  return [...members].toSorted((left, right) =>
    axis === 'row' ? left.row - right.row || left.cellIndex - right.cellIndex : left.col - right.col || left.cellIndex - right.cellIndex,
  )
}

function groupRunMembersByStep(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[], step: number): FormulaFamilyMember[][] {
  const sorted = sortRunMembers(axis, members)
  const groups: FormulaFamilyMember[][] = []
  for (const member of sorted) {
    const current = groups[groups.length - 1]
    const previous = current?.[current.length - 1]
    const memberIndex = axis === 'row' ? member.row : member.col
    const previousIndex = previous ? (axis === 'row' ? previous.row : previous.col) : undefined
    if (!current || previousIndex === undefined || memberIndex !== previousIndex + step) {
      groups.push([member])
      continue
    }
    current.push(member)
  }
  return groups
}
