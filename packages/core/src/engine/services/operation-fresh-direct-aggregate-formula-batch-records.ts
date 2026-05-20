import type { CompiledFormula } from '@bilig/formula'
import { buildFormulaFamilyShapeKey } from '../../formula/formula-family-deps.js'
import type {
  FormulaFamilyFreshUniformRunRegistrationArgs,
  FormulaFamilyMember,
  FormulaFamilyRunUpsertArgs,
} from '../../formula/formula-family-store.js'
import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeFormula } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'

export interface FreshDirectAggregateFormulaEntry {
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId: number
  readonly aggregateKind: RuntimeDirectAggregateDescriptor['aggregateKind']
  readonly aggregateRowStart: number
  readonly aggregateRowEnd: number
  readonly aggregateColStart: number
  readonly aggregateColEnd: number
  readonly resultOffset: number | undefined
  readonly result: DirectScalarCurrentOperand
}

export type FreshDirectAggregateFormulaEntrySeed = Omit<FreshDirectAggregateFormulaEntry, 'result'>

export interface FreshFormulaEntryPosition {
  readonly row: number
  readonly col: number
}

export interface ContiguousSingleColumnFormulaBatch {
  readonly rowStart: number
  readonly col: number
}

export interface FreshFormulaFamilyRunRegistrationArgs {
  readonly state: {
    readonly formulas: {
      get(cellIndex: number): RuntimeFormula | undefined
    }
  }
  readonly registerFreshFormulaFamilyRun: ((run: FormulaFamilyFreshUniformRunRegistrationArgs) => boolean) | undefined
  readonly upsertFormulaFamilyRun: ((run: FormulaFamilyRunUpsertArgs) => void) | undefined
}

export function createFreshFormulaInstanceList(count: number): FormulaInstanceSnapshot[] {
  const records: FormulaInstanceSnapshot[] = []
  records.length = count
  return records
}

export function registerFreshDirectAggregateFormulaFamilyRun(
  args: FreshFormulaFamilyRunRegistrationArgs,
  sheetId: number,
  entries: readonly FreshDirectAggregateFormulaEntry[],
  cellIndices: readonly number[] | Uint32Array,
): void {
  registerFreshBoundFormulaFamilyRun(args, sheetId, entries, cellIndices)
}

export function registerFreshBoundFormulaFamilyRun(
  args: FreshFormulaFamilyRunRegistrationArgs,
  sheetId: number,
  entries: readonly FreshFormulaEntryPosition[],
  cellIndices: readonly number[] | Uint32Array,
): void {
  const upsertFormulaFamilyRun = args.upsertFormulaFamilyRun
  if (upsertFormulaFamilyRun === undefined || entries.length === 0) {
    return
  }
  const firstRegistration = readBoundFormulaFamilyRegistration(args, cellIndices[0])
  if (firstRegistration === undefined) {
    return
  }

  let uniformSingleColumnRun = true
  let sameFamily = true
  for (let index = 1; index < entries.length; index += 1) {
    const entry = entries[index]!
    const registration = readBoundFormulaFamilyRegistration(args, cellIndices[index])
    if (
      registration === undefined ||
      registration.templateId !== firstRegistration.templateId ||
      registration.shapeKey !== firstRegistration.shapeKey
    ) {
      sameFamily = false
      uniformSingleColumnRun = false
      break
    }
    const firstEntry = entries[0]!
    if (entry.col !== firstEntry.col || entry.row !== firstEntry.row + index) {
      uniformSingleColumnRun = false
    }
  }

  const firstEntry = entries[0]!
  if (
    sameFamily &&
    uniformSingleColumnRun &&
    args.registerFreshFormulaFamilyRun?.({
      sheetId,
      templateId: firstRegistration.templateId,
      shapeKey: firstRegistration.shapeKey,
      axis: 'row',
      fixedIndex: firstEntry.col,
      start: firstEntry.row,
      step: 1,
      cellIndices,
    })
  ) {
    return
  }

  if (sameFamily) {
    upsertFormulaFamilyRun({
      sheetId,
      templateId: firstRegistration.templateId,
      shapeKey: firstRegistration.shapeKey,
      members: materializeFormulaFamilyMembers(entries, cellIndices, 0, entries.length),
    })
    return
  }

  const groups = new Map<string, { templateId: number; shapeKey: string; members: FormulaFamilyMember[] }>()
  for (let index = 0; index < entries.length; index += 1) {
    const registration = readBoundFormulaFamilyRegistration(args, cellIndices[index])
    if (registration === undefined) {
      continue
    }
    const key = `${registration.templateId}\t${registration.shapeKey}`
    let group = groups.get(key)
    if (group === undefined) {
      group = { templateId: registration.templateId, shapeKey: registration.shapeKey, members: [] }
      groups.set(key, group)
    }
    const entry = entries[index]!
    group.members.push({ cellIndex: cellIndices[index]!, row: entry.row, col: entry.col })
  }
  groups.forEach((group) => {
    upsertFormulaFamilyRun({
      sheetId,
      templateId: group.templateId,
      shapeKey: group.shapeKey,
      members: group.members,
    })
  })
}

export function getContiguousSingleColumnFormulaBatch(
  entries: readonly FreshFormulaEntryPosition[],
): ContiguousSingleColumnFormulaBatch | undefined {
  const first = entries[0]
  if (first === undefined) {
    return undefined
  }
  for (let index = 1; index < entries.length; index += 1) {
    const entry = entries[index]!
    if (entry.col !== first.col || entry.row !== first.row + index) {
      return undefined
    }
  }
  return { rowStart: first.row, col: first.col }
}

function readBoundFormulaFamilyRegistration(
  args: FreshFormulaFamilyRunRegistrationArgs,
  cellIndex: number | undefined,
): { readonly templateId: number; readonly shapeKey: string } | undefined {
  if (cellIndex === undefined) {
    return undefined
  }
  const formula = args.state.formulas.get(cellIndex)
  if (formula === undefined || formula.templateId === undefined) {
    return undefined
  }
  return {
    templateId: formula.templateId,
    shapeKey: directAggregateRuntimeFormulaFamilyShapeKey(formula),
  }
}

function directAggregateRuntimeFormulaFamilyShapeKey(formula: RuntimeFormula): string {
  return buildFormulaFamilyShapeKey({
    compiled: formula.compiled,
    dependencyCount: formula.dependencyIndices.length,
    rangeDependencyCount: formula.rangeDependencies.length,
    directAggregateKind: formula.directAggregate?.aggregateKind,
    directLookupKind: formula.directLookup?.kind,
    directScalarKind: formula.directScalar?.kind,
    directCriteriaKind: formula.directCriteria?.aggregateKind,
  })
}

function materializeFormulaFamilyMembers(
  entries: readonly FreshFormulaEntryPosition[],
  cellIndices: readonly number[] | Uint32Array,
  start: number,
  end: number,
): FormulaFamilyMember[] {
  const members: FormulaFamilyMember[] = []
  members.length = end - start
  for (let index = start; index < end; index += 1) {
    const entry = entries[index]!
    members[index - start] = { cellIndex: cellIndices[index]!, row: entry.row, col: entry.col }
  }
  return members
}
