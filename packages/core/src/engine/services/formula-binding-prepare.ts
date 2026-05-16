import { formulaContainsDateSystemSensitiveBuiltin, type CompiledFormula } from '@bilig/formula'
import { FormulaMode, Opcode } from '@bilig/protocol'
import { resolveRuntimeDirectLookupBinding } from '../direct-vector-lookup.js'
import {
  type CompiledPlanRecord,
  type MaterializedDependencies,
  type RuntimeDirectAggregateDescriptor,
  type RuntimeDirectCriteriaDescriptor,
  type RuntimeDirectLookupDescriptor,
  type RuntimeDirectScalarDescriptor,
  UNRESOLVED_WASM_OPERAND,
} from '../runtime-state.js'
import { collectDirectApproximateLookupCandidates, collectIndexedExactLookupCandidates } from './formula-binding-lookup-candidates.js'
import { buildDirectScalarDescriptor, tryParseDependencyCellAddress } from './formula-binding-direct-scalar.js'
import {
  buildDirectAggregateDescriptor,
  buildDirectCriteriaDescriptor,
  buildDirectLookupDescriptor,
  hasLookupPlanInstruction,
  type ParsedCompiledFormula,
} from './formula-binding-direct-descriptors.js'
import type { FormulaBindingDependencyMaterializer } from './formula-binding-dependency-materializer.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

const PUSH_CELL_OPCODE = Number(Opcode.PushCell)
const PUSH_RANGE_OPCODE = Number(Opcode.PushRange)
const PUSH_STRING_OPCODE = Number(Opcode.PushString)
const EMPTY_RUNTIME_PROGRAM = new Uint32Array(0)

function shouldEvaluateMetadataCellAliasInJs(compiled: ParsedCompiledFormula): boolean {
  return compiled.symbolicNames.length > 0 && compiled.optimizedAst.kind === 'CellRef'
}

function normalizeMetadataCellAliasMode(compiled: ParsedCompiledFormula): ParsedCompiledFormula {
  return shouldEvaluateMetadataCellAliasInJs(compiled) ? { ...compiled, mode: FormulaMode.JsOnly } : compiled
}

function shouldEvaluateWorkbookDateSystemInJs(compiled: ParsedCompiledFormula, dateSystem: string | undefined): boolean {
  return dateSystem === '1904' && formulaContainsDateSystemSensitiveBuiltin(compiled.ast)
}

export interface PreparedFormulaBinding {
  readonly compiled: ParsedCompiledFormula
  readonly dependencies: MaterializedDependencies
  readonly directLookup: RuntimeDirectLookupDescriptor | undefined
  readonly directAggregate: RuntimeDirectAggregateDescriptor | undefined
  readonly directScalar: RuntimeDirectScalarDescriptor | undefined
  readonly directCriteria: RuntimeDirectCriteriaDescriptor | undefined
  readonly runtimeProgram: Uint32Array
  readonly plan: CompiledPlanRecord
  readonly templateId: number | undefined
  readonly indexedExactLookupCandidates: ReturnType<typeof collectIndexedExactLookupCandidates>
  readonly directApproximateLookupCandidates: ReturnType<typeof collectDirectApproximateLookupCandidates>
}

export function prepareFormulaBindingFromCompiled(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly cellIndex: number
  readonly ownerSheetName: string
  readonly source: string
  readonly compiledInput: ParsedCompiledFormula
  readonly templateId: number | undefined
  readonly normalizeLookupCompileMode: (compiled: ParsedCompiledFormula) => ParsedCompiledFormula
  readonly dependencyMaterializer: FormulaBindingDependencyMaterializer
  readonly ensureDependencyBuildCapacity: (
    cellCapacity: number,
    dependencyCapacity: number,
    symbolicRefCapacity?: number,
    symbolicRangeCapacity?: number,
  ) => void
  readonly directAggregateContainsOwnerCell: (directAggregate: RuntimeDirectAggregateDescriptor | undefined, cellIndex: number) => boolean
  readonly makeUnmanagedCompiledPlan: (source: string, compiled: CompiledFormula, templateId: number | undefined) => CompiledPlanRecord
}): PreparedFormulaBinding {
  const serviceArgs = args.serviceArgs
  const ownerSheetId = serviceArgs.state.workbook.getSheet(args.ownerSheetName)?.id
  const workbookDateSystem = serviceArgs.state.workbook.getCalculationSettings().dateSystem
  const requiresWorkbookDateSystemJs = shouldEvaluateWorkbookDateSystemInJs(args.compiledInput, workbookDateSystem)
  const compiled = normalizeMetadataCellAliasMode(
    requiresWorkbookDateSystemJs
      ? { ...args.normalizeLookupCompileMode(args.compiledInput), mode: FormulaMode.JsOnly }
      : args.normalizeLookupCompileMode(args.compiledInput),
  )
  const hasLookupInstruction = hasLookupPlanInstruction(compiled.jsPlan)
  const directLookupBinding =
    !requiresWorkbookDateSystemJs && hasLookupInstruction
      ? resolveRuntimeDirectLookupBinding(compiled.jsPlan, args.ownerSheetName)
      : undefined
  const directScalar = requiresWorkbookDateSystemJs
    ? undefined
    : buildDirectScalarDescriptor({
        compiled,
        ownerSheetName: args.ownerSheetName,
        ownerSheetId,
        workbook: serviceArgs.state.workbook,
        ensureCellTracked: serviceArgs.ensureCellTracked,
        ensureCellTrackedByCoords: serviceArgs.ensureCellTrackedByCoords,
      })
  const directAggregateCandidate =
    directScalar === undefined && !requiresWorkbookDateSystemJs
      ? buildDirectAggregateDescriptor({
          compiled,
          ownerSheetName: args.ownerSheetName,
          workbook: serviceArgs.state.workbook,
          regionGraph: serviceArgs.regionGraph,
        })
      : undefined
  const directAggregate = args.directAggregateContainsOwnerCell(directAggregateCandidate, args.cellIndex)
    ? undefined
    : directAggregateCandidate
  const directCriteria =
    directScalar === undefined && directAggregate === undefined && !requiresWorkbookDateSystemJs
      ? buildDirectCriteriaDescriptor({
          compiled,
          source: args.source,
          ownerSheetName: args.ownerSheetName,
          workbook: serviceArgs.state.workbook,
          ensureCellTracked: serviceArgs.ensureCellTracked,
          regionGraph: serviceArgs.regionGraph,
        })
      : undefined
  const indexedExactLookupCandidates =
    hasLookupInstruction && serviceArgs.state.getUseColumnIndex() ? collectIndexedExactLookupCandidates(compiled.optimizedAst) : []
  const directApproximateLookupCandidates = hasLookupInstruction ? collectDirectApproximateLookupCandidates(compiled.optimizedAst) : []
  const directScalarDependencies = args.dependencyMaterializer.materializeDirectScalarDependencies(compiled, directScalar)
  const directAggregateDependencies =
    directScalarDependencies === undefined
      ? args.dependencyMaterializer.materializeDirectAggregateDependencies(compiled, directAggregate)
      : undefined
  const directCriteriaDependencies =
    directScalarDependencies === undefined && directAggregateDependencies === undefined
      ? args.dependencyMaterializer.materializeDirectCriteriaDependencies(compiled, directCriteria)
      : undefined
  const dependencies =
    directScalarDependencies ??
    directAggregateDependencies ??
    directCriteriaDependencies ??
    args.dependencyMaterializer.materializeDependencies(args.ownerSheetName, compiled, directAggregate, directLookupBinding)
  const directLookup = directLookupBinding
    ? buildDirectLookupDescriptor({
        compiled,
        ownerSheetName: args.ownerSheetName,
        workbook: serviceArgs.state.workbook,
        ensureCellTracked: serviceArgs.ensureCellTracked,
        exactLookup: serviceArgs.exactLookup,
        sortedLookup: serviceArgs.sortedLookup,
      })
    : undefined

  if (directScalarDependencies === undefined) {
    args.ensureDependencyBuildCapacity(
      serviceArgs.state.workbook.cellStore.size + 1,
      compiled.deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
    )
    for (let index = 0; index < compiled.symbolicRefs.length; index += 1) {
      const parsedRef = compiled.parsedSymbolicRefs?.[index]
      if (parsedRef && parsedRef.sheetName === undefined) {
        serviceArgs.getSymbolicRefBindings()[index] =
          parsedRef.row !== undefined && parsedRef.col !== undefined && ownerSheetId !== undefined
            ? serviceArgs.ensureCellTrackedByCoords(ownerSheetId, parsedRef.row, parsedRef.col)
            : serviceArgs.ensureCellTracked(args.ownerSheetName, parsedRef.address)
        continue
      }
      const ref = compiled.symbolicRefs[index]!
      const [qualifiedSheetName, qualifiedAddress] = ref.includes('!') ? ref.split('!') : [undefined, ref]
      const fallbackAddress = tryParseDependencyCellAddress(qualifiedAddress, qualifiedSheetName)?.text
      if (fallbackAddress === undefined) {
        serviceArgs.getSymbolicRefBindings()[index] = UNRESOLVED_WASM_OPERAND
        continue
      }
      const sheetName =
        parsedRef?.sheetName ??
        qualifiedSheetName ??
        serviceArgs.state.workbook.getSheetNameById(serviceArgs.state.workbook.cellStore.sheetIds[args.cellIndex]!)
      if ((parsedRef?.sheetName ?? qualifiedSheetName) && !serviceArgs.state.workbook.getSheet(sheetName)) {
        serviceArgs.getSymbolicRefBindings()[index] = UNRESOLVED_WASM_OPERAND
        continue
      }
      serviceArgs.getSymbolicRefBindings()[index] = serviceArgs.ensureCellTracked(sheetName, parsedRef?.address ?? fallbackAddress)
    }
  }

  const directOnlyRuntimeProgram =
    (directScalar !== undefined || directAggregate !== undefined || directCriteria !== undefined) &&
    !compiled.volatile &&
    !compiled.producesSpill
  const literalStringIds = directOnlyRuntimeProgram ? [] : compiled.symbolicStrings.map((value) => serviceArgs.state.strings.intern(value))
  const runtimeProgram =
    directOnlyRuntimeProgram || compiled.program.length === 0 ? EMPTY_RUNTIME_PROGRAM : new Uint32Array(compiled.program.length)
  if (runtimeProgram.length > 0) {
    runtimeProgram.set(compiled.program)
    for (let index = 0; index < compiled.program.length; index += 1) {
      const instruction = compiled.program[index]!
      const opcode = instruction >>> 24
      const operand = instruction & 0x00ff_ffff
      if (opcode === PUSH_CELL_OPCODE) {
        const targetIndex = operand < compiled.symbolicRefs.length ? (serviceArgs.getSymbolicRefBindings()[operand] ?? 0) : 0
        runtimeProgram[index] = (PUSH_CELL_OPCODE << 24) | (targetIndex & 0x00ff_ffff)
        continue
      }
      if (opcode === PUSH_RANGE_OPCODE) {
        const targetIndex = operand < dependencies.symbolicRangeCount ? (dependencies.symbolicRangeIndices[operand] ?? 0) : 0
        runtimeProgram[index] = (PUSH_RANGE_OPCODE << 24) | (targetIndex & 0x00ff_ffff)
        continue
      }
      if (opcode === PUSH_STRING_OPCODE) {
        const stringId = operand < literalStringIds.length ? (literalStringIds[operand] ?? 0) : 0
        runtimeProgram[index] = (PUSH_STRING_OPCODE << 24) | (stringId & 0x00ff_ffff)
      }
    }
  }

  return {
    compiled,
    dependencies,
    directLookup,
    directAggregate,
    directScalar,
    directCriteria,
    runtimeProgram,
    plan: directOnlyRuntimeProgram
      ? args.makeUnmanagedCompiledPlan(args.source, compiled, args.templateId)
      : serviceArgs.compiledPlans.intern(args.source, compiled, args.templateId),
    templateId: args.templateId,
    indexedExactLookupCandidates,
    directApproximateLookupCandidates,
  }
}
