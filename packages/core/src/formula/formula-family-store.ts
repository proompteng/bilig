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

export interface FormulaFamilyStore {
  readonly upsertFormula: (args: FormulaFamilyKey & FormulaFamilyMember) => FormulaFamilyMembership
  readonly unregisterFormula: (cellIndex: number) => boolean
  readonly getMembership: (cellIndex: number) => FormulaFamilyMembership | undefined
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
  cellIndices: number[]
}

interface MutableFormulaFamily {
  id: FormulaFamilyId
  sheetId: number
  templateId: number
  shapeKey: string
  key: string
  runs: MutableFormulaFamilyMemberRun[]
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
    const run = makeRun(axis, fixedIndex, members)
    family.runs.splice(runIndex, 1, run)
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
    family.runs.splice(runIndex, 1)
    if (remainingMembers.length === 0) {
      return
    }
    const groups = groupContiguousRunMembers(run.axis, remainingMembers)
    groups.forEach((group, offset) => {
      const nextRun = makeRun(run.axis, run.fixedIndex, group)
      family.runs.splice(runIndex + offset, 0, nextRun)
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
    const family = familiesById.get(membership.familyId)
    if (!family) {
      return true
    }
    const runIndex = family.runs.findIndex((run) => run.id === membership.runId)
    splitRunAfterRemoval(family, runIndex, cellIndex)
    if (family.runs.length === 0) {
      familiesById.delete(family.id)
      familyIdByKey.delete(family.key)
    }
    return true
  }

  return {
    upsertFormula(args) {
      unregisterFormula(args.cellIndex)
      const family = getOrCreateFamily(args)
      const member: FormulaFamilyMember = { cellIndex: args.cellIndex, row: args.row, col: args.col }
      cellRecords.set(args.cellIndex, { ...args })

      for (let runIndex = 0; runIndex < family.runs.length; runIndex += 1) {
        const run = family.runs[runIndex]!
        const maybeMembership = tryMergeRun(family, runIndex, run, member, appendMemberToRun, replaceRunWithMembers, cellRecords)
        if (maybeMembership) {
          return maybeMembership
        }
      }

      const run = makeRun('row', args.col, [member])
      family.runs.push(run)
      const membership = { familyId: family.id, runId: run.id }
      memberships.set(args.cellIndex, membership)
      return membership
    },
    unregisterFormula,
    getMembership(cellIndex) {
      return memberships.get(cellIndex)
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
    },
  }
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
  if (run.axis === 'row' && run.fixedIndex === member.col && member.row >= run.start - 1 && member.row <= run.end + 1) {
    return appendMemberToRun(family, run, member)
  }
  if (run.axis === 'column' && run.fixedIndex === member.row && member.col >= run.start - 1 && member.col <= run.end + 1) {
    return appendMemberToRun(family, run, member)
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
    if (existing.col === member.col && Math.abs(existing.row - member.row) === 1) {
      return replaceRunWithMembers(family, runIndex, 'row', member.col, [existing, member], member.cellIndex)
    }
    if (existing.row === member.row && Math.abs(existing.col - member.col) === 1) {
      return replaceRunWithMembers(family, runIndex, 'column', member.row, [existing, member], member.cellIndex)
    }
  }
  return undefined
}

function sortRunMembers(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): FormulaFamilyMember[] {
  return [...members].toSorted((left, right) =>
    axis === 'row' ? left.row - right.row || left.cellIndex - right.cellIndex : left.col - right.col || left.cellIndex - right.cellIndex,
  )
}

function groupContiguousRunMembers(axis: FormulaFamilyRunAxis, members: readonly FormulaFamilyMember[]): FormulaFamilyMember[][] {
  const sorted = sortRunMembers(axis, members)
  const groups: FormulaFamilyMember[][] = []
  for (const member of sorted) {
    const current = groups[groups.length - 1]
    const previous = current?.[current.length - 1]
    const memberIndex = axis === 'row' ? member.row : member.col
    const previousIndex = previous ? (axis === 'row' ? previous.row : previous.col) : undefined
    if (!current || previousIndex === undefined || memberIndex !== previousIndex + 1) {
      groups.push([member])
      continue
    }
    current.push(member)
  }
  return groups
}
