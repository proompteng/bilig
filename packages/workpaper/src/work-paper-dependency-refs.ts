import { parseCellAddress, parseRangeAddress } from '@bilig/formula'
import { assertRange } from './work-paper-runtime-helpers.js'
import type { WorkPaperCellAddress, WorkPaperCellRange, WorkPaperDependencyRef } from './work-paper-types.js'

export interface WorkPaperDependencyResolver {
  readonly defaultSheetName: () => string
  readonly requireSheetId: (name: string) => number
}

export function toWorkPaperDependencyRefs(values: readonly string[], resolver: WorkPaperDependencyResolver): WorkPaperDependencyRef[] {
  return values.map((value) => toWorkPaperDependencyRef(value, resolver))
}

export function collectWorkPaperRangeDependencies(input: {
  readonly range: WorkPaperCellRange
  readonly readDependencies: (address: WorkPaperCellAddress) => readonly string[]
  readonly resolver: WorkPaperDependencyResolver
}): WorkPaperDependencyRef[] {
  assertRange(input.range)
  const seen = new Set<string>()
  const collected: WorkPaperDependencyRef[] = []
  for (let row = input.range.start.row; row <= input.range.end.row; row += 1) {
    for (let col = input.range.start.col; col <= input.range.end.col; col += 1) {
      const address = { sheet: input.range.start.sheet, row, col }
      toWorkPaperDependencyRefs(input.readDependencies(address), input.resolver).forEach((dependency) => {
        const key = workPaperDependencyKey(dependency)
        if (seen.has(key)) {
          return
        }
        seen.add(key)
        collected.push(dependency)
      })
    }
  }
  return collected
}

function toWorkPaperDependencyRef(value: string, resolver: WorkPaperDependencyResolver): WorkPaperDependencyRef {
  try {
    const parsedCell = parseCellAddress(value)
    return {
      kind: 'cell',
      address: {
        sheet: resolver.requireSheetId(parsedCell.sheetName ?? resolver.defaultSheetName()),
        row: parsedCell.row,
        col: parsedCell.col,
      },
    }
  } catch {
    try {
      const parsedRange = parseRangeAddress(value)
      if (parsedRange.kind === 'cells') {
        const sheet = resolver.requireSheetId(parsedRange.sheetName ?? resolver.defaultSheetName())
        return {
          kind: 'range',
          range: {
            start: {
              sheet,
              row: parsedRange.start.row,
              col: parsedRange.start.col,
            },
            end: {
              sheet,
              row: parsedRange.end.row,
              col: parsedRange.end.col,
            },
          },
        }
      }
    } catch {
      return { kind: 'name', name: value }
    }
  }
  return { kind: 'name', name: value }
}

function workPaperDependencyKey(dependency: WorkPaperDependencyRef): string {
  if (dependency.kind === 'cell') {
    return `cell:${dependency.address.sheet}:${dependency.address.row}:${dependency.address.col}`
  }
  if (dependency.kind === 'range') {
    return `range:${dependency.range.start.sheet}:${dependency.range.start.row}:${dependency.range.start.col}:${dependency.range.end.row}:${dependency.range.end.col}`
  }
  return `name:${dependency.name}`
}
