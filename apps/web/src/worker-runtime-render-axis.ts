import type { WorkbookAxisEntrySnapshot } from '@bilig/protocol'

export interface RenderedAxisState {
  readonly sizes: Record<number, number>
  readonly sortedOverrides: Array<readonly [number, number]>
  readonly version: number
}

export function buildRenderedAxisState(entries: readonly WorkbookAxisEntrySnapshot[], defaultSize: number): RenderedAxisState {
  const sizes: Record<number, number> = {}
  const sortedOverrides: Array<readonly [number, number]> = []
  let hash = entries.length === 0 ? 0 : 2_166_136_261
  for (const entry of [...entries].toSorted((left, right) => left.index - right.index)) {
    const renderedSize = entry.hidden ? 0 : (entry.size ?? defaultSize)
    if (renderedSize !== defaultSize) {
      sizes[entry.index] = renderedSize
      sortedOverrides.push([entry.index, renderedSize] as const)
    }
    hash = mixRevisionInteger(hash, entry.index)
    hash = mixRevisionInteger(hash, Math.round((entry.size ?? -1) * 1_000))
    hash = mixRevisionInteger(hash, entry.hidden ? 1 : 0)
  }
  return {
    sizes,
    sortedOverrides,
    version: hash >>> 0,
  }
}

export function resolveRevision(value: number | undefined): number {
  return Number.isInteger(value) && value !== undefined && value >= 0 ? value : 0
}

export function buildFreezeVersion(freezeRows: number, freezeCols: number): number {
  return mixRevisionInteger(mixRevisionInteger(2_166_136_261, freezeRows), freezeCols)
}

function mixRevisionInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}
