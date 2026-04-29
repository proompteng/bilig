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

export interface FormulaFamilyRunUpsertArgs extends FormulaFamilyKey {
  readonly members: readonly FormulaFamilyMember[]
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
  readonly upsertFormulaRun: (args: FormulaFamilyRunUpsertArgs) => FormulaFamilyMembership[]
  readonly registerFormulaRun: (args: FormulaFamilyRunUpsertArgs) => void
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
  rowAppendRunByFixedIndex: Array<MutableFormulaFamilyMemberRun | undefined>
  recentAppendRun: MutableFormulaFamilyMemberRun | undefined
}

interface FormulaFamilyCellRecord extends FormulaFamilyKey, FormulaFamilyMember {}

interface FormulaFamilyRunDescriptor {
  readonly axis: FormulaFamilyRunAxis
  readonly fixedIndex: number
  readonly start: number
  readonly end: number
  readonly step: number
  readonly members: readonly FormulaFamilyMember[]
}

type FormulaRunFastPath =
  | {
      readonly kind: 'append'
      readonly run: MutableFormulaFamilyMemberRun
    }
  | {
      readonly kind: 'create'
    }

function keyForFormulaFamily(args: FormulaFamilyKey): string {
  return `${args.sheetId}\t${args.templateId}\t${args.shapeKey}`
}

export function createFormulaFamilyStore(): FormulaFamilyStore {
  const familiesById = new Map<FormulaFamilyId, MutableFormulaFamily>()
  const familyIdByKey = new Map<string, FormulaFamilyId>()
  const recentFamilyByTemplateId = new Map<number, MutableFormulaFamily>()
  const cellRecords: Array<FormulaFamilyCellRecord | undefined> = []
  const membershipFamilyIds: number[] = []
  const membershipRunIds: number[] = []
  const sheetMemberCounts = new Map<number, number>()
  const structuralSourceTransforms = new Map<FormulaFamilyId, FormulaFamilyStructuralSourceTransform>()
  const noMemberships: FormulaFamilyMembership[] = []
  let memberCount = 0
  let nextFamilyId = 1
  let nextRunId = 1

  const setMembership = (cellIndex: number, familyId: FormulaFamilyId, runId: FormulaFamilyRunId): FormulaFamilyMembership => {
    membershipFamilyIds[cellIndex] = familyId
    membershipRunIds[cellIndex] = runId
    return { familyId, runId }
  }

  const getMembershipRecord = (cellIndex: number): FormulaFamilyMembership | undefined => {
    const familyId = membershipFamilyIds[cellIndex] ?? 0
    if (familyId === 0) {
      return undefined
    }
    return { familyId, runId: membershipRunIds[cellIndex]! }
  }

  const getExistingFamily = (args: FormulaFamilyKey): MutableFormulaFamily | undefined => {
    const recent = recentFamilyByTemplateId.get(args.templateId)
    if (recent && recent.sheetId === args.sheetId && recent.shapeKey === args.shapeKey && familiesById.get(recent.id) === recent) {
      return recent
    }
    const key = keyForFormulaFamily(args)
    const existingId = familyIdByKey.get(key)
    if (existingId === undefined) {
      return undefined
    }
    const existing = familiesById.get(existingId)
    if (existing) {
      recentFamilyByTemplateId.set(args.templateId, existing)
    }
    return existing
  }

  const getOrCreateFamily = (args: FormulaFamilyKey): MutableFormulaFamily => {
    const existing = getExistingFamily(args)
    if (existing) {
      return existing
    }
    const key = keyForFormulaFamily(args)
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
      rowAppendRunByFixedIndex: [],
      recentAppendRun: undefined,
    }
    nextFamilyId += 1
    familiesById.set(family.id, family)
    familyIdByKey.set(key, family.id)
    recentFamilyByTemplateId.set(args.templateId, family)
    return family
  }

  const recordFormulaMember = (key: FormulaFamilyKey, member: FormulaFamilyMember): void => {
    cellRecords[member.cellIndex] = {
      sheetId: key.sheetId,
      templateId: key.templateId,
      shapeKey: key.shapeKey,
      cellIndex: member.cellIndex,
      row: member.row,
      col: member.col,
    }
    memberCount += 1
    sheetMemberCounts.set(key.sheetId, (sheetMemberCounts.get(key.sheetId) ?? 0) + 1)
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
    family.recentAppendRun = run
    run.cellIndices.forEach((cellIndex) => {
      setMembership(cellIndex, family.id, run.id)
    })
    return getMembershipRecord(memberCellIndex)!
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
    family.recentAppendRun = run
    return setMembership(member.cellIndex, family.id, run.id)
  }

  const splitRunAfterRemoval = (family: MutableFormulaFamily, runIndex: number, removedCellIndex: number): void => {
    const run = family.runs[runIndex]
    if (!run) {
      return
    }
    const remainingMembers = run.cellIndices
      .filter((cellIndex) => cellIndex !== removedCellIndex)
      .flatMap((cellIndex): FormulaFamilyMember[] => {
        const record = cellRecords[cellIndex]
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
        setMembership(cellIndex, family.id, nextRun.id)
      })
    })
  }

  const unregisterFormula = (cellIndex: number): boolean => {
    const membership = getMembershipRecord(cellIndex)
    const record = cellRecords[cellIndex]
    if (!membership || !record) {
      return false
    }
    membershipFamilyIds[cellIndex] = 0
    membershipRunIds[cellIndex] = 0
    cellRecords[cellIndex] = undefined
    memberCount -= 1
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
      if (recentFamilyByTemplateId.get(record.templateId) === family) {
        recentFamilyByTemplateId.delete(record.templateId)
      }
      structuralSourceTransforms.delete(family.id)
    }
    return true
  }

  const upsertFormula = (args: FormulaFamilyKey & FormulaFamilyMember): FormulaFamilyMembership => {
    if ((membershipFamilyIds[args.cellIndex] ?? 0) !== 0) {
      unregisterFormula(args.cellIndex)
    }
    const family = getOrCreateFamily(args)
    const member: FormulaFamilyMember = args
    recordFormulaMember(args, member)

    const rowAppendRun = family.rowAppendRunByFixedIndex[member.col]
    if (rowAppendRun && canAppendStridedRunMember(rowAppendRun, member.row)) {
      return appendMemberToRun(family, rowAppendRun, member)
    }

    const recentRun = family.recentAppendRun
    if (recentRun?.axis === 'row' && recentRun.fixedIndex === member.col && canAppendStridedRunMember(recentRun, member.row)) {
      return appendMemberToRun(family, recentRun, member)
    }

    const rowRuns = family.rowRunsByFixedIndex.get(member.col)
    if (rowRuns?.length === 1) {
      const run = rowRuns[0]!
      if (canAppendStridedRunMember(run, member.row)) {
        return appendMemberToRun(family, run, member)
      }
    }

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
    family.recentAppendRun = run
    return setMembership(args.cellIndex, family.id, run.id)
  }

  const fallbackUpsertFormulaRun = (args: FormulaFamilyRunUpsertArgs): FormulaFamilyMembership[] => {
    args.members.forEach((member) => {
      upsertFormula({ sheetId: args.sheetId, templateId: args.templateId, shapeKey: args.shapeKey, ...member })
    })
    return args.members.map((member) => getMembershipRecord(member.cellIndex)!)
  }

  const fallbackRegisterFormulaRun = (args: FormulaFamilyRunUpsertArgs): void => {
    args.members.forEach((member) => {
      upsertFormula({ sheetId: args.sheetId, templateId: args.templateId, shapeKey: args.shapeKey, ...member })
    })
  }

  const tryRegisterFreshOrderedUniformRun = (args: FormulaFamilyRunUpsertArgs): boolean => {
    const descriptor = describeFreshOrderedUniformRun(args.members, membershipFamilyIds)
    if (!descriptor) {
      return false
    }
    const family = getOrCreateFamily(args)
    if (!canRegisterFreshOrderedRunWithoutMerging(family, descriptor, args.members)) {
      return false
    }
    const cellIndices: number[] = []
    cellIndices.length = args.members.length
    const run: MutableFormulaFamilyMemberRun = {
      id: nextRunId++,
      axis: descriptor.axis,
      fixedIndex: descriptor.fixedIndex,
      start: descriptor.start,
      end: descriptor.end,
      step: descriptor.step,
      cellIndices,
    }
    for (let index = 0; index < args.members.length; index += 1) {
      const member = args.members[index]!
      recordFormulaMember(args, member)
      run.cellIndices[index] = member.cellIndex
      membershipFamilyIds[member.cellIndex] = family.id
      membershipRunIds[member.cellIndex] = run.id
    }
    family.runs.push(run)
    indexRun(family, run)
    family.recentAppendRun = run
    return true
  }

  const tryUpsertFormulaRunDescriptors = (
    args: FormulaFamilyRunUpsertArgs,
    descriptors: readonly FormulaFamilyRunDescriptor[],
    materializeMemberships = true,
  ): FormulaFamilyMembership[] | undefined => {
    const existingFamily = getExistingFamily(args)
    const fastPaths: FormulaRunFastPath[] = []
    for (const descriptor of descriptors) {
      const fastPath = existingFamily ? getFormulaRunFastPath(existingFamily, descriptor) : { kind: 'create' as const }
      if (!fastPath) {
        return undefined
      }
      fastPaths.push(fastPath)
    }
    const family = existingFamily ?? getOrCreateFamily(args)
    args.members.forEach((member) => {
      recordFormulaMember(args, member)
    })
    descriptors.forEach((descriptor, index) => {
      const fastPath = fastPaths[index]!
      const run = fastPath.kind === 'append' ? fastPath.run : makeRun(descriptor.axis, descriptor.fixedIndex, descriptor.members)
      if (fastPath.kind === 'append') {
        appendFormulaRunMembers(family, run, descriptor)
      } else {
        family.runs.push(run)
        indexRun(family, run)
        family.recentAppendRun = run
      }
      descriptor.members.forEach((member) => {
        setMembership(member.cellIndex, family.id, run.id)
      })
    })
    if (!materializeMemberships) {
      return noMemberships
    }
    return args.members.map((member) => getMembershipRecord(member.cellIndex)!)
  }

  const upsertFormulaRun = (args: FormulaFamilyRunUpsertArgs): FormulaFamilyMembership[] => {
    const inOrderDescriptors = describeFreshFormulaRunMemberSegmentsInOrder(args.members, membershipFamilyIds)
    if (inOrderDescriptors) {
      const memberships = tryUpsertFormulaRunDescriptors(args, inOrderDescriptors)
      if (memberships) {
        return memberships
      }
    }
    if (hasDuplicateRunMemberCellIndex(args.members)) {
      return fallbackUpsertFormulaRun(args)
    }
    if (args.members.some((member) => (membershipFamilyIds[member.cellIndex] ?? 0) !== 0)) {
      return fallbackUpsertFormulaRun(args)
    }
    const descriptor = describeFormulaRunMembers(args.members)
    if (!descriptor) {
      return fallbackUpsertFormulaRun(args)
    }
    return tryUpsertFormulaRunDescriptors(args, [descriptor]) ?? fallbackUpsertFormulaRun(args)
  }

  const registerFormulaRun = (args: FormulaFamilyRunUpsertArgs): void => {
    if (tryRegisterFreshOrderedUniformRun(args)) {
      return
    }
    const inOrderDescriptors = describeFreshFormulaRunMemberSegmentsInOrder(args.members, membershipFamilyIds)
    if (inOrderDescriptors && tryUpsertFormulaRunDescriptors(args, inOrderDescriptors, false)) {
      return
    }
    if (hasDuplicateRunMemberCellIndex(args.members)) {
      fallbackRegisterFormulaRun(args)
      return
    }
    if (args.members.some((member) => (membershipFamilyIds[member.cellIndex] ?? 0) !== 0)) {
      fallbackRegisterFormulaRun(args)
      return
    }
    const descriptor = describeFormulaRunMembers(args.members)
    if (!descriptor || !tryUpsertFormulaRunDescriptors(args, [descriptor], false)) {
      fallbackRegisterFormulaRun(args)
    }
  }

  return {
    upsertFormula,
    upsertFormulaRun,
    registerFormulaRun,
    unregisterFormula,
    getMembership(cellIndex) {
      return getMembershipRecord(cellIndex)
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
      const membership = getMembershipRecord(cellIndex)
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
        memberCount,
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
      const removedCellIndices: number[] = []
      for (let cellIndex = 0; cellIndex < cellRecords.length; cellIndex += 1) {
        const record = cellRecords[cellIndex]
        if (record?.sheetId === sheetId) {
          removedCellIndices.push(record.cellIndex)
        }
      }
      removedCellIndices.forEach((cellIndex) => {
        unregisterFormula(cellIndex)
      })
    },
    applyStructuralInvalidation(args) {
      const removedCellIndices: number[] = []
      for (let cellIndex = 0; cellIndex < cellRecords.length; cellIndex += 1) {
        const record = cellRecords[cellIndex]
        if (record?.sheetId !== args.sheetId) {
          continue
        }
        const axisIndex = args.axis === 'row' ? record.row : record.col
        if (axisIndex >= args.start && axisIndex < args.end) {
          removedCellIndices.push(record.cellIndex)
        }
      }
      removedCellIndices.forEach((cellIndex) => {
        unregisterFormula(cellIndex)
      })
    },
    clear() {
      familiesById.clear()
      familyIdByKey.clear()
      recentFamilyByTemplateId.clear()
      cellRecords.length = 0
      membershipFamilyIds.length = 0
      membershipRunIds.length = 0
      sheetMemberCounts.clear()
      structuralSourceTransforms.clear()
      memberCount = 0
    },
  }
}

function indexRun(family: MutableFormulaFamily, run: MutableFormulaFamilyMemberRun): void {
  appendMapArray(run.axis === 'row' ? family.rowRunsByFixedIndex : family.columnRunsByFixedIndex, run.fixedIndex, run)
  if (run.axis === 'row') {
    family.rowAppendRunByFixedIndex[run.fixedIndex] = run
  }
  if (run.cellIndices.length === 1) {
    appendMapArray(family.singletonRunsByRow, runRowStart(run), run)
  }
}

function unindexRun(family: MutableFormulaFamily, run: MutableFormulaFamilyMemberRun): void {
  removeMapArrayValue(run.axis === 'row' ? family.rowRunsByFixedIndex : family.columnRunsByFixedIndex, run.fixedIndex, run)
  if (run.axis === 'row' && family.rowAppendRunByFixedIndex[run.fixedIndex] === run) {
    family.rowAppendRunByFixedIndex[run.fixedIndex] = undefined
  }
  if (run.cellIndices.length === 1) {
    removeMapArrayValue(family.singletonRunsByRow, runRowStart(run), run)
  }
  if (family.recentAppendRun === run) {
    family.recentAppendRun = undefined
  }
}

function hasDuplicateRunMemberCellIndex(members: readonly FormulaFamilyMember[]): boolean {
  return new Set(members.map((member) => member.cellIndex)).size !== members.length
}

function describeFreshOrderedUniformRun(
  members: readonly FormulaFamilyMember[],
  membershipFamilyIds: readonly number[],
): Omit<FormulaFamilyRunDescriptor, 'members'> | undefined {
  const first = members[0]
  if (!first || (membershipFamilyIds[first.cellIndex] ?? 0) !== 0) {
    return undefined
  }
  if (members.length === 1) {
    return {
      axis: 'row',
      fixedIndex: first.col,
      start: first.row,
      end: first.row,
      step: 1,
    }
  }
  const second = members[1]!
  if ((membershipFamilyIds[second.cellIndex] ?? 0) !== 0 || second.cellIndex <= first.cellIndex) {
    return undefined
  }
  let axis: FormulaFamilyRunAxis
  let fixedIndex: number
  let previousIndex: number
  let step: number
  if (second.col === first.col && second.row > first.row) {
    axis = 'row'
    fixedIndex = first.col
    previousIndex = first.row
    step = second.row - first.row
  } else if (second.row === first.row && second.col > first.col) {
    axis = 'column'
    fixedIndex = first.row
    previousIndex = first.col
    step = second.col - first.col
  } else {
    return undefined
  }
  for (let index = 1; index < members.length; index += 1) {
    const member = members[index]!
    if ((membershipFamilyIds[member.cellIndex] ?? 0) !== 0 || member.cellIndex <= members[index - 1]!.cellIndex) {
      return undefined
    }
    const memberIndex = axis === 'row' ? member.row : member.col
    if (memberIndex !== previousIndex + step || (axis === 'row' ? member.col !== fixedIndex : member.row !== fixedIndex)) {
      return undefined
    }
    previousIndex = memberIndex
  }
  return {
    axis,
    fixedIndex,
    start: axis === 'row' ? first.row : first.col,
    end: previousIndex,
    step,
  }
}

function describeFreshFormulaRunMemberSegmentsInOrder(
  members: readonly FormulaFamilyMember[],
  membershipFamilyIds: readonly number[],
): FormulaFamilyRunDescriptor[] | undefined {
  const first = members[0]
  if (!first || (membershipFamilyIds[first.cellIndex] ?? 0) !== 0) {
    return undefined
  }
  const second = members[1]!
  if (!second) {
    return undefined
  }
  if ((membershipFamilyIds[second.cellIndex] ?? 0) !== 0 || second.cellIndex <= first.cellIndex) {
    return undefined
  }
  let axis: FormulaFamilyRunAxis
  let fixedIndex: number
  let previousIndex: number
  let step: number
  if (second.col === first.col && second.row > first.row) {
    axis = 'row'
    fixedIndex = first.col
    previousIndex = first.row
    step = second.row - first.row
  } else if (second.row === first.row && second.col > first.col) {
    axis = 'column'
    fixedIndex = first.row
    previousIndex = first.col
    step = second.col - first.col
  } else {
    return undefined
  }
  const indices = [previousIndex]
  for (let index = 1; index < members.length; index += 1) {
    const member = members[index]!
    if ((membershipFamilyIds[member.cellIndex] ?? 0) !== 0) {
      return undefined
    }
    if (index > 0 && member.cellIndex <= members[index - 1]!.cellIndex) {
      return undefined
    }
    if (axis === 'row') {
      if (member.col !== fixedIndex || member.row <= previousIndex) {
        return undefined
      }
      previousIndex = member.row
    } else {
      if (member.row !== fixedIndex || member.col <= previousIndex) {
        return undefined
      }
      previousIndex = member.col
    }
    indices.push(previousIndex)
  }
  const lastIndex = indices[indices.length - 1]!
  if (isUniformIndexRun(indices, step)) {
    return [
      {
        axis,
        fixedIndex,
        start: axis === 'row' ? first.row : first.col,
        end: lastIndex,
        step,
        members,
      },
    ]
  }
  for (let cycleLength = 2; cycleLength <= Math.min(8, Math.floor(members.length / 2)); cycleLength += 1) {
    const descriptors = describeDeinterleavedFormulaRunSegments(axis, fixedIndex, members, cycleLength)
    if (descriptors) {
      return descriptors
    }
  }
  return undefined
}

function describeDeinterleavedFormulaRunSegments(
  axis: FormulaFamilyRunAxis,
  fixedIndex: number,
  members: readonly FormulaFamilyMember[],
  cycleLength: number,
): FormulaFamilyRunDescriptor[] | undefined {
  const descriptors: FormulaFamilyRunDescriptor[] = []
  for (let offset = 0; offset < cycleLength; offset += 1) {
    const segment: FormulaFamilyMember[] = []
    for (let index = offset; index < members.length; index += cycleLength) {
      segment.push(members[index]!)
    }
    const indices = segment.map((member) => (axis === 'row' ? member.row : member.col))
    const step = indices[1]! - indices[0]!
    if (step <= 0 || !isUniformIndexRun(indices, step)) {
      return undefined
    }
    descriptors.push({
      axis,
      fixedIndex,
      start: indices[0]!,
      end: indices[indices.length - 1]!,
      step,
      members: segment,
    })
  }
  return descriptors
}

function isUniformIndexRun(indices: readonly number[], step: number): boolean {
  for (let index = 1; index < indices.length; index += 1) {
    if (indices[index]! !== indices[index - 1]! + step) {
      return false
    }
  }
  return true
}

function describeFormulaRunMembers(members: readonly FormulaFamilyMember[]): FormulaFamilyRunDescriptor | undefined {
  const first = members[0]
  if (!first) {
    return undefined
  }
  if (members.length === 1) {
    return {
      axis: 'row',
      fixedIndex: first.col,
      start: first.row,
      end: first.row,
      step: 1,
      members: [first],
    }
  }
  const isRowRun = members.every((member) => member.col === first.col)
  const isColumnRun = members.every((member) => member.row === first.row)
  if (isRowRun === isColumnRun) {
    return undefined
  }
  const axis: FormulaFamilyRunAxis = isRowRun ? 'row' : 'column'
  const sorted = sortRunMembers(axis, members)
  const firstIndex = axis === 'row' ? sorted[0]!.row : sorted[0]!.col
  const secondIndex = axis === 'row' ? sorted[1]!.row : sorted[1]!.col
  const step = secondIndex - firstIndex
  if (step <= 0) {
    return undefined
  }
  for (let index = 2; index < sorted.length; index += 1) {
    const previous = sorted[index - 1]!
    const current = sorted[index]!
    const previousIndex = axis === 'row' ? previous.row : previous.col
    const currentIndex = axis === 'row' ? current.row : current.col
    if (currentIndex - previousIndex !== step) {
      return undefined
    }
  }
  return {
    axis,
    fixedIndex: axis === 'row' ? first.col : first.row,
    start: firstIndex,
    end: axis === 'row' ? sorted[sorted.length - 1]!.row : sorted[sorted.length - 1]!.col,
    step,
    members: sorted,
  }
}

function getFormulaRunFastPath(family: MutableFormulaFamily, descriptor: FormulaFamilyRunDescriptor): FormulaRunFastPath | undefined {
  const targetRuns =
    descriptor.axis === 'row'
      ? family.rowRunsByFixedIndex.get(descriptor.fixedIndex)
      : family.columnRunsByFixedIndex.get(descriptor.fixedIndex)
  let appendRun: MutableFormulaFamilyMemberRun | undefined
  let appendRunCount = 0
  targetRuns?.forEach((run) => {
    if (!canAppendFormulaRunDescriptor(run, descriptor)) {
      return
    }
    appendRun = run
    appendRunCount += 1
  })
  if (appendRunCount > 1) {
    return undefined
  }
  for (const member of descriptor.members) {
    for (const candidate of candidateRunsForMember(family, member)) {
      if (candidate !== appendRun) {
        return undefined
      }
    }
  }
  return appendRun ? { kind: 'append', run: appendRun } : { kind: 'create' }
}

function canAppendFormulaRunDescriptor(run: MutableFormulaFamilyMemberRun, descriptor: FormulaFamilyRunDescriptor): boolean {
  return (
    run.axis === descriptor.axis &&
    run.fixedIndex === descriptor.fixedIndex &&
    run.step === descriptor.step &&
    (descriptor.start === run.end + run.step || descriptor.end === run.start - run.step)
  )
}

function appendFormulaRunMembers(
  family: MutableFormulaFamily,
  run: MutableFormulaFamilyMemberRun,
  descriptor: FormulaFamilyRunDescriptor,
): void {
  if (run.cellIndices.length === 1) {
    removeMapArrayValue(family.singletonRunsByRow, runRowStart(run), run)
  }
  const cellIndices = descriptor.members.map((member) => member.cellIndex)
  if (descriptor.end < run.start) {
    run.start = descriptor.start
    run.cellIndices.unshift(...cellIndices)
  } else {
    run.end = descriptor.end
    run.cellIndices.push(...cellIndices)
  }
  family.recentAppendRun = run
}

function canRegisterFreshOrderedRunWithoutMerging(
  family: MutableFormulaFamily,
  descriptor: Omit<FormulaFamilyRunDescriptor, 'members'>,
  members: readonly FormulaFamilyMember[],
): boolean {
  const targetRuns =
    descriptor.axis === 'row'
      ? family.rowRunsByFixedIndex.get(descriptor.fixedIndex)
      : family.columnRunsByFixedIndex.get(descriptor.fixedIndex)
  if (targetRuns && targetRuns.length > 0) {
    return false
  }
  return members.every((member) => candidateRunsForMember(family, member).length === 0)
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
  cellRecords: readonly (FormulaFamilyCellRecord | undefined)[],
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
    const existingRecord = cellRecords[run.cellIndices[0]!]
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
  cellRecords: readonly (FormulaFamilyCellRecord | undefined)[],
): FormulaFamilyMembership | undefined {
  const memberIndex = run.axis === 'row' ? member.row : member.col
  if (run.cellIndices.length !== 2 || memberIndex <= run.start || memberIndex >= run.end) {
    return undefined
  }
  const existingMembers = run.cellIndices.flatMap((cellIndex): FormulaFamilyMember[] => {
    const record = cellRecords[cellIndex]
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
