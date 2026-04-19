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

function keyForSingleColumnRegion(args: {
  readonly sheetId: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly col: number
}): string {
  return `${args.sheetId}\t${args.col}\t${args.rowStart}\t${args.rowEnd}`
}

export function createRegionNodeStore(): RegionNodeStore {
  const nodes: SingleColumnRegionNode[] = []
  const byKey = new Map<string, RegionId>()

  return {
    internSingleColumnRegion(args) {
      const key = keyForSingleColumnRegion(args)
      const existing = byKey.get(key)
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
      byKey.set(key, id)
      return id
    },
    get(regionId) {
      return nodes[regionId]
    },
  }
}
