import type { CompiledFormula } from '@bilig/formula'
import type { CompiledPlanRecord, RuntimeDirectScalarDescriptor } from '../runtime-state.js'

export function formulaBindingErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function canUseDirectOnlyRuntimeProgram(compiled: CompiledFormula, directScalar: RuntimeDirectScalarDescriptor | undefined): boolean {
  return directScalar !== undefined && !compiled.volatile && !compiled.producesSpill
}

export function makeUnmanagedCompiledPlan(source: string, compiled: CompiledFormula, templateId: number | undefined): CompiledPlanRecord {
  return {
    id: 0,
    source,
    compiled,
    ...(templateId !== undefined ? { templateId } : {}),
  }
}

export function canRetainUnmanagedCompiledPlan(
  existingPlanId: number,
  compiled: CompiledFormula,
  directScalar: RuntimeDirectScalarDescriptor | undefined,
): boolean {
  return existingPlanId === 0 && canUseDirectOnlyRuntimeProgram(compiled, directScalar)
}
