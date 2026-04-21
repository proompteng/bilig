import type { CompiledFormula } from '@bilig/formula'

export interface FormulaFamilyShapeKeyInput {
  readonly compiled: Pick<CompiledFormula, 'mode' | 'deps' | 'symbolicRanges' | 'symbolicNames' | 'symbolicTables' | 'symbolicSpills'>
  readonly dependencyCount: number
  readonly rangeDependencyCount: number
  readonly directAggregateKind?: string | undefined
  readonly directLookupKind?: string | undefined
  readonly directScalarKind?: string | undefined
  readonly directCriteriaKind?: string | undefined
}

export function buildFormulaFamilyShapeKey(input: FormulaFamilyShapeKeyInput): string {
  return JSON.stringify({
    mode: input.compiled.mode,
    deps: input.compiled.deps.length,
    symbolicRanges: input.compiled.symbolicRanges.length,
    symbolicNames: input.compiled.symbolicNames.length,
    symbolicTables: input.compiled.symbolicTables.length,
    symbolicSpills: input.compiled.symbolicSpills.length,
    dependencyCount: input.dependencyCount,
    rangeDependencyCount: input.rangeDependencyCount,
    directAggregateKind: input.directAggregateKind ?? null,
    directLookupKind: input.directLookupKind ?? null,
    directScalarKind: input.directScalarKind ?? null,
    directCriteriaKind: input.directCriteriaKind ?? null,
  })
}
