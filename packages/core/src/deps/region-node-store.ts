export type RegionId = number

export interface SingleColumnRegionNode {
  readonly id: RegionId
  readonly kind: 'single-column'
  readonly sheetId: number
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly col: number
  readonly length: number
}

export interface RegionNodeStore {
  readonly internSingleColumnRegion: (args: {
    readonly sheetId: number
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly col: number
  }) => RegionId
  readonly get: (regionId: RegionId) => SingleColumnRegionNode | undefined
}

export function createRegionNodeStore(): RegionNodeStore {
  const nodes: SingleColumnRegionNode[] = []
  const bySheet = new Map<number, Map<number, Map<number, Map<number, RegionId>>>>()

  return {
    internSingleColumnRegion(args) {
      let byCol = bySheet.get(args.sheetId)
      if (!byCol) {
        byCol = new Map()
        bySheet.set(args.sheetId, byCol)
      }
      let byRowStart = byCol.get(args.col)
      if (!byRowStart) {
        byRowStart = new Map()
        byCol.set(args.col, byRowStart)
      }
      let byRowEnd = byRowStart.get(args.rowStart)
      if (!byRowEnd) {
        byRowEnd = new Map()
        byRowStart.set(args.rowStart, byRowEnd)
      }
      const existing = byRowEnd.get(args.rowEnd)
      if (existing !== undefined) {
        return existing
      }
      const id = nodes.length
      const node: SingleColumnRegionNode = {
        id,
        kind: 'single-column',
        sheetId: args.sheetId,
        sheetName: args.sheetName,
        rowStart: args.rowStart,
        rowEnd: args.rowEnd,
        col: args.col,
        length: args.rowEnd - args.rowStart + 1,
      }
      nodes.push(node)
      byRowEnd.set(args.rowEnd, id)
      return id
    },
    get(regionId) {
      return nodes[regionId]
    },
  }
}
