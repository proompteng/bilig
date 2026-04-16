import type { CellValue } from '@bilig/protocol'
import type { EvaluationResult } from './runtime-values.js'

export const externalFunctionSurfaces = ['cube', 'web', 'host', 'external-data', 'add-in'] as const

export type ExternalFunctionSurface = (typeof externalFunctionSurfaces)[number]

export type ExternalScalarFunction = (...args: CellValue[]) => CellValue

export interface ExternalRangeFunctionArgument {
  kind: 'range'
  values: CellValue[]
  refKind: 'cells' | 'rows' | 'cols'
  rows: number
  cols: number
}

export type ExternalLookupFunctionArgument = CellValue | ExternalRangeFunctionArgument
export type ExternalLookupFunction = (...args: ExternalLookupFunctionArgument[]) => EvaluationResult

export type ExternalFunctionBinding =
  | { kind: 'scalar'; implementation: ExternalScalarFunction }
  | { kind: 'lookup'; implementation: ExternalLookupFunction }

export interface ExternalFunctionAdapter {
  readonly surface: ExternalFunctionSurface
  resolveFunction(name: string): ExternalFunctionBinding | undefined
}

const adapters = new Map<ExternalFunctionSurface, ExternalFunctionAdapter>()

function normalizeFunctionName(name: string): string {
  return name.trim().toUpperCase()
}

function resolveExternalFunction(name: string): ExternalFunctionBinding | undefined {
  const normalized = normalizeFunctionName(name)
  if (!normalized) {
    return undefined
  }

  for (const adapter of adapters.values()) {
    const binding = adapter.resolveFunction(normalized)
    if (binding) {
      return binding
    }
  }
  return undefined
}

export function installExternalFunctionAdapter(adapter: ExternalFunctionAdapter): void {
  adapters.set(adapter.surface, adapter)
}

export function removeExternalFunctionAdapter(surface: ExternalFunctionSurface): void {
  adapters.delete(surface)
}

export function clearExternalFunctionAdapters(): void {
  adapters.clear()
}

export function listExternalFunctionAdapterSurfaces(): ExternalFunctionSurface[] {
  return [...adapters.keys()]
}

export function hasExternalFunction(name: string): boolean {
  return resolveExternalFunction(name) !== undefined
}

export function getExternalScalarFunction(name: string): ExternalScalarFunction | undefined {
  const binding = resolveExternalFunction(name)
  return binding?.kind === 'scalar' ? binding.implementation : undefined
}

export function getExternalLookupFunction(name: string): ExternalLookupFunction | undefined {
  const binding = resolveExternalFunction(name)
  return binding?.kind === 'lookup' ? binding.implementation : undefined
}
