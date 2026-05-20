import type { EdgeArena } from '../../edge-arena.js'
import { makeCellEntity } from '../../entity-ids.js'
import { markFormulaCellBound } from './formula-binding-cell-flags.js'
import { buildDirectScalarDescriptor } from './formula-binding-direct-scalar.js'
import { appendFreshFormulaDependencyReverseEdges } from './formula-binding-install.js'
import type { FormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import { makeUnmanagedCompiledPlan } from './formula-binding-plan-helpers.js'
import type {
  CreateEngineFormulaBindingServiceArgs,
  FreshDirectScalarFormulaBindingInput,
  FreshDirectScalarFormulaBindingMember,
} from './formula-binding-service-types.js'
import type { RuntimeDirectScalarDescriptor, RuntimeDirectScalarOperand, RuntimeFormula } from '../runtime-state.js'

const EMPTY_U32 = new Uint32Array(0)

export function bindFreshDirectScalarFormulaRun(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly edgeArena: EdgeArena
  readonly formulaMemberCounts: FormulaBindingMemberCounts
  readonly appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void
  readonly trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void
  readonly run: FreshDirectScalarFormulaBindingInput
}): void {
  if ('member' in args.run) {
    assertFreshDirectScalarFormulaCell(args.serviceArgs, args.run.cellIndex)
    bindFreshDirectScalarFormulaMember(
      args.serviceArgs,
      args.edgeArena,
      args.formulaMemberCounts,
      args.appendKnownUniqueReverseEdge,
      args.trackFormulaSheetIndexes,
      args.run.sheetId,
      args.run.ownerSheetName,
      args.run.cellIndex,
      args.run.member,
    )
    return
  }

  if (args.run.cellIndices.length !== args.run.members.length) {
    throw new Error('Expected fresh direct scalar formula cell index count to match member count')
  }

  for (let index = 0; index < args.run.members.length; index += 1) {
    assertFreshDirectScalarFormulaCell(args.serviceArgs, args.run.cellIndices[index]!)
  }

  for (let index = 0; index < args.run.members.length; index += 1) {
    bindFreshDirectScalarFormulaMember(
      args.serviceArgs,
      args.edgeArena,
      args.formulaMemberCounts,
      args.appendKnownUniqueReverseEdge,
      args.trackFormulaSheetIndexes,
      args.run.sheetId,
      args.run.ownerSheetName,
      args.run.cellIndices[index]!,
      args.run.members[index]!,
    )
  }
}

function assertFreshDirectScalarFormulaCell(serviceArgs: CreateEngineFormulaBindingServiceArgs, cellIndex: number): void {
  if ((serviceArgs.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0) {
    throw new Error('Expected fresh direct scalar formula cell')
  }
}

function bindFreshDirectScalarFormulaMember(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  edgeArena: EdgeArena,
  formulaMemberCounts: FormulaBindingMemberCounts,
  appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void,
  trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void,
  sheetId: number,
  ownerSheetName: string,
  cellIndex: number,
  member: FreshDirectScalarFormulaBindingMember,
): void {
  const directScalar = buildDirectScalarDescriptor({
    compiled: member.compiled,
    ownerSheetName,
    ownerSheetId: sheetId,
    workbook: serviceArgs.state.workbook,
    ensureCellTracked: serviceArgs.ensureCellTracked,
    ensureCellTrackedByCoords: serviceArgs.ensureCellTrackedByCoords,
  })
  if (directScalar === undefined) {
    throw new Error('Expected fresh direct scalar formula descriptor')
  }
  const dependencies = materializeFreshDirectScalarDependencies(member.compiled, directScalar)
  if (dependencies === undefined) {
    throw new Error('Expected fresh direct scalar dependencies')
  }
  const dependencyEntities = edgeArena.replace(edgeArena.empty(), dependencies.dependencyEntities)
  const runtimeFormula: RuntimeFormula = {
    cellIndex,
    formulaSlotId: 0,
    planId: 0,
    templateId: member.templateId,
    source: member.source,
    compiled: member.compiled,
    plan: makeUnmanagedCompiledPlan(member.source, member.compiled, member.templateId),
    dependencyIndices: dependencies.dependencyIndices,
    dependencyEntities,
    rangeDependencies: EMPTY_U32,
    graphRangeDependencies: EMPTY_U32,
    runtimeProgram: EMPTY_U32,
    constants: member.compiled.constants,
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: member.compiled.constants.length,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar,
    directCriteria: undefined,
  }
  const formulaSlotId = serviceArgs.state.formulas.set(cellIndex, runtimeFormula)
  runtimeFormula.formulaSlotId = formulaSlotId
  formulaMemberCounts.increment(sheetId, member.col)
  markFormulaCellBound(serviceArgs.state.workbook.cellStore, cellIndex, member.compiled.mode)
  appendFreshFormulaDependencyReverseEdges(dependencies.dependencyEntities, makeCellEntity(cellIndex), appendKnownUniqueReverseEdge)
  trackFormulaSheetIndexes(cellIndex, ownerSheetName, member.compiled)
}

function materializeFreshDirectScalarDependencies(
  compiled: FreshDirectScalarFormulaBindingMember['compiled'],
  directScalar: RuntimeDirectScalarDescriptor,
): { readonly dependencyIndices: Uint32Array; readonly dependencyEntities: Uint32Array } | undefined {
  if (
    compiled.symbolicRanges.length !== 0 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0
  ) {
    return undefined
  }
  const dependencyIndices = new Uint32Array(Math.min(compiled.symbolicRefs.length, 2))
  const dependencyEntities = new Uint32Array(compiled.symbolicRefs.length)
  let dependencyIndexCount = 0
  let dependencyEntityCount = 0
  const appendOperand = (operand: RuntimeDirectScalarOperand): boolean => {
    if (operand.kind === 'literal-number') {
      return true
    }
    if (operand.kind === 'error') {
      return false
    }
    const cellIndex = operand.cellIndex
    let seen = false
    for (let existingIndex = 0; existingIndex < dependencyIndexCount; existingIndex += 1) {
      if (dependencyIndices[existingIndex] === cellIndex) {
        seen = true
        break
      }
    }
    if (!seen) {
      dependencyIndices[dependencyIndexCount] = cellIndex
      dependencyIndexCount += 1
    }
    dependencyEntities[dependencyEntityCount] = makeCellEntity(cellIndex)
    dependencyEntityCount += 1
    return true
  }
  const matched =
    directScalar.kind === 'abs'
      ? appendOperand(directScalar.operand)
      : appendOperand(directScalar.left) && appendOperand(directScalar.right)
  if (!matched || dependencyEntityCount !== compiled.symbolicRefs.length) {
    return undefined
  }
  return {
    dependencyIndices:
      dependencyIndexCount === dependencyIndices.length ? dependencyIndices : dependencyIndices.subarray(0, dependencyIndexCount),
    dependencyEntities,
  }
}
