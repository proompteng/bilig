import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { packTileKey53 } from '../renderer-v3/tile-key.js'
import { TileResidencyV3, type TileUpsertInputV3 } from '../renderer-v3/tile-residency.js'

interface FuzzTilePayload {
  readonly label: string
}

interface FuzzTileResources {
  readonly resourceId: number
}

function tileInput(input: {
  readonly rowTile: number
  readonly colTile: number
  readonly byteSizeCpu: number
  readonly byteSizeGpu: number
  readonly packet?: FuzzTilePayload | null
  readonly resources?: FuzzTileResources | null
}): TileUpsertInputV3<FuzzTilePayload, FuzzTileResources> {
  return {
    sheetOrdinal: 1,
    rowTile: input.rowTile,
    colTile: input.colTile,
    dprBucket: 1,
    axisSeqX: 1,
    axisSeqY: 1,
    freezeSeq: 1,
    valueSeq: 1,
    styleSeq: 1,
    textSeq: 1,
    rectSeq: 1,
    key: packTileKey53({ sheetOrdinal: 1, rowTile: input.rowTile, colTile: input.colTile, dprBucket: 1 }),
    byteSizeCpu: input.byteSizeCpu,
    byteSizeGpu: input.byteSizeGpu,
    ...(input.packet === undefined ? {} : { packet: input.packet }),
    ...(input.resources === undefined ? {} : { resources: input.resources }),
  }
}

describe('tile residency fuzz', () => {
  it('should treat explicit null upsert fields as resource clearing operations', async () => {
    await runProperty({
      suite: 'grid/tile-residency/explicit-null-clears-payloads',
      arbitrary: fc.record({
        rowTile: fc.integer({ min: 0, max: 12 }),
        colTile: fc.integer({ min: 0, max: 12 }),
        packetLabel: fc.string({ minLength: 1, maxLength: 12 }),
        resourceId: fc.integer({ min: 1, max: 10_000 }),
        byteSizeCpu: fc.integer({ min: 1, max: 5_000 }),
        byteSizeGpu: fc.integer({ min: 1, max: 5_000 }),
      }),
      predicate: async ({ rowTile, colTile, packetLabel, resourceId, byteSizeCpu, byteSizeGpu }) => {
        const residency = new TileResidencyV3<FuzzTilePayload, FuzzTileResources>()
        const initial = residency.upsert(
          tileInput({
            rowTile,
            colTile,
            byteSizeCpu,
            byteSizeGpu,
            packet: { label: packetLabel },
            resources: { resourceId },
          }),
        )

        expect(initial.packet).toEqual({ label: packetLabel })
        expect(initial.resources).toEqual({ resourceId })

        const cleared = residency.upsert(
          tileInput({
            rowTile,
            colTile,
            byteSizeCpu,
            byteSizeGpu,
            packet: null,
            resources: null,
          }),
        )

        expect(cleared.packet).toBeNull()
        expect(cleared.resources).toBeNull()
      },
      parameters: { numRuns: 120 },
    })
  })
})
