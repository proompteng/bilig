import type { WorkPaperSheetDimensions } from './work-paper-types.js'

export function applyCachedSheetDimensionInsertion(args: {
  readonly axis: 'row' | 'column'
  readonly cache: Map<number, WorkPaperSheetDimensions>
  readonly count: number
  readonly sheetId: number
  readonly start: number
}): boolean {
  const cached = args.cache.get(args.sheetId)
  if (!cached) {
    return false
  }
  if (args.count <= 0) {
    return true
  }
  if (args.axis === 'row') {
    if (args.start < cached.height) {
      cached.height += args.count
    }
    return true
  }
  if (args.start < cached.width) {
    cached.width += args.count
  }
  return true
}

export function applyCachedSheetDimensionDeletion(args: {
  readonly axis: 'row' | 'column'
  readonly cache: Map<number, WorkPaperSheetDimensions>
  readonly count: number
  readonly sheetId: number
  readonly start: number
}): boolean {
  const cached = args.cache.get(args.sheetId)
  if (!cached) {
    return false
  }
  if (args.count <= 0) {
    return true
  }
  if (args.axis === 'row') {
    if (args.start >= cached.height) {
      return true
    }
    if (args.start + args.count >= cached.height) {
      return false
    }
    cached.height -= args.count
    return true
  }
  if (args.start >= cached.width) {
    return true
  }
  if (args.start + args.count >= cached.width) {
    return false
  }
  cached.width -= args.count
  return true
}

export function applyCachedSheetDimensionMove(args: {
  readonly axis: 'row' | 'column'
  readonly cache: Map<number, WorkPaperSheetDimensions>
  readonly count: number
  readonly sheetId: number
  readonly start: number
  readonly target: number
}): boolean {
  const cached = args.cache.get(args.sheetId)
  if (!cached) {
    return false
  }
  if (args.count <= 0 || args.start === args.target) {
    return true
  }
  const extent = args.axis === 'row' ? cached.height : cached.width
  if (args.start >= extent || args.target >= extent) {
    return false
  }
  return args.start + args.count < extent
}
