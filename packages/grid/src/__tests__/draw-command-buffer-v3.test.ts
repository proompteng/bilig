import { describe, expect, it } from 'vitest'
import { packTileKey53 } from '../renderer-v3/tile-key.js'
import { DrawCommandBufferV3 } from '../renderer-v3/draw-command-buffer.js'

describe('DrawCommandBufferV3', () => {
  it('separates content tile identity from pane placement', () => {
    const tileKey = packTileKey53({ sheetOrdinal: 1, rowTile: 0, colTile: 0, dprBucket: 1 })
    const buffer = new DrawCommandBufferV3()

    buffer.beginFrame()
    buffer.addBodyPlacement({
      tileKey,
      clipWidth: 200,
      clipHeight: 120,
      translateX: -40,
      translateY: -16,
    })
    buffer.addPlacement({
      tileKey,
      pane: 'frozenRows',
      clipX: 0,
      clipY: 0,
      clipWidth: 200,
      clipHeight: 24,
      translateX: -40,
      translateY: 0,
      z: 1,
    })

    expect(buffer.snapshot()).toEqual({
      frameSeq: 1,
      placements: [
        {
          tileKey,
          pane: 'body',
          clipX: 0,
          clipY: 0,
          clipWidth: 200,
          clipHeight: 120,
          translateX: -40,
          translateY: -16,
          z: 0,
        },
        {
          tileKey,
          pane: 'frozenRows',
          clipX: 0,
          clipY: 0,
          clipWidth: 200,
          clipHeight: 24,
          translateX: -40,
          translateY: 0,
          z: 1,
        },
      ],
    })
  })

  it('starts each frame with an empty placement list', () => {
    const buffer = new DrawCommandBufferV3()

    buffer.beginFrame()
    buffer.addBodyPlacement({
      tileKey: packTileKey53({ sheetOrdinal: 1, rowTile: 0, colTile: 0, dprBucket: 1 }),
      clipWidth: 1,
      clipHeight: 1,
      translateX: 0,
      translateY: 0,
    })
    buffer.beginFrame()

    expect(buffer.snapshot()).toEqual({
      frameSeq: 2,
      placements: [],
    })
  })
})
