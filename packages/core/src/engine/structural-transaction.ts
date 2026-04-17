import type { StructuralAxisTransform } from '@bilig/formula'
import type { SheetGridAxisRemapScope } from '../sheet-grid.js'

export interface StructuralRemappedCell {
  readonly cellIndex: number
  readonly fromRow: number
  readonly fromCol: number
  readonly toRow: number | undefined
  readonly toCol: number | undefined
}

export interface StructuralInvalidationSpan {
  readonly axis: 'row' | 'column'
  readonly start: number
  readonly end: number
}

export interface StructuralTransaction {
  readonly sheetName: string
  readonly sheetId: number
  readonly transform: StructuralAxisTransform
  readonly scope: SheetGridAxisRemapScope
  readonly remappedCells: readonly StructuralRemappedCell[]
  readonly removedCellIndices: readonly number[]
  readonly invalidationSpans: readonly StructuralInvalidationSpan[]
}

export function structuralScopeForTransform(transform: StructuralAxisTransform): SheetGridAxisRemapScope {
  switch (transform.kind) {
    case 'insert':
    case 'delete':
      return { start: transform.start }
    case 'move':
      if (transform.target < transform.start) {
        return { start: transform.target, end: transform.start + transform.count }
      }
      if (transform.target > transform.start) {
        return { start: transform.start, end: transform.target + transform.count }
      }
      return { start: transform.start, end: transform.start + transform.count }
    default: {
      const exhaustive: never = transform
      return exhaustive
    }
  }
}

export function buildStructuralTransaction(input: {
  readonly sheetName: string
  readonly sheetId: number
  readonly transform: StructuralAxisTransform
  readonly remappedCells: readonly StructuralRemappedCell[]
}): StructuralTransaction {
  const scope = structuralScopeForTransform(input.transform)
  const removedCellIndices = input.remappedCells
    .filter((entry) => entry.toRow === undefined || entry.toCol === undefined)
    .map((entry) => entry.cellIndex)

  return {
    sheetName: input.sheetName,
    sheetId: input.sheetId,
    transform: input.transform,
    scope,
    remappedCells: input.remappedCells,
    removedCellIndices,
    invalidationSpans: [
      {
        axis: input.transform.axis,
        start: scope.start,
        end:
          input.transform.kind === 'move'
            ? (scope.end ?? input.transform.start + input.transform.count)
            : input.transform.start + input.transform.count,
      },
    ],
  }
}
