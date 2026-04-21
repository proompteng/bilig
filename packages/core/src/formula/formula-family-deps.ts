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
  return [
    input.compiled.mode,
    input.compiled.deps.length,
    input.compiled.symbolicRanges.length,
    input.compiled.symbolicNames.length,
    input.compiled.symbolicTables.length,
    input.compiled.symbolicSpills.length,
    input.dependencyCount,
    input.rangeDependencyCount,
    input.directAggregateKind ?? '',
    input.directLookupKind ?? '',
    input.directScalarKind ?? '',
    input.directCriteriaKind ?? '',
  ].join('|')
}
