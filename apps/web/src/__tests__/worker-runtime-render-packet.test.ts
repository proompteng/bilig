import { describe, expect, it } from 'vitest'
import { GRID_SCENE_PACKET_V2_VERSION } from '../../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { packWorkerGridScenePacket } from '../worker-runtime-render-packet.js'

describe('worker-runtime-render-packet', () => {
  it('packs worker scene data into transferable typed arrays', () => {
    const packet = packWorkerGridScenePacket({
      generation: 3,
      paneId: 'body',
      sheetName: 'Sheet1',
      surfaceSize: { width: 400, height: 200 },
      viewport: { rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 1 },
      gpuScene: {
        fillRects: [{ x: 1, y: 2, width: 3, height: 4, color: { r: 1, g: 0.5, b: 0.25, a: 1 } }],
        borderRects: [{ x: 5, y: 6, width: 7, height: 8, color: { r: 0, g: 0, b: 0, a: 1 } }],
      },
      textScene: {
        items: [
          {
            align: 'left',
            clipInsetBottom: 0,
            clipInsetLeft: 1,
            clipInsetRight: 2,
            clipInsetTop: 3,
            color: '#000000',
            font: '400 11px sans-serif',
            fontSize: 11,
            height: 20,
            strike: false,
            text: 'A',
            underline: false,
            width: 100,
            wrap: false,
            x: 10,
            y: 11,
          },
        ],
      },
    })

    expect(packet.rectInstances).toBeInstanceOf(Float32Array)
    expect(packet.rects).toBeInstanceOf(Float32Array)
    expect(packet.magic).toBe('bilig.grid.scene.v2')
    expect(packet.version).toBe(GRID_SCENE_PACKET_V2_VERSION)
    expect(packet.sheetName).toBe('Sheet1')
    expect(packet.key).toMatchObject({
      axisVersionX: 0,
      axisVersionY: 0,
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetName: 'Sheet1',
      valueVersion: 0,
      styleVersion: 0,
    })
    expect(packet.surfaceSize).toEqual({ width: 400, height: 200 })
    expect(packet.textMetrics).toBeInstanceOf(Float32Array)
    expect(packet.rectCount).toBe(2)
    expect(packet.fillRectCount).toBe(1)
    expect(packet.borderRectCount).toBe(1)
    expect(packet.textCount).toBe(1)
    expect(Array.from(packet.rectInstances.slice(0, 20))).toEqual([1, 2, 3, 4, 1, 0.5, 0.25, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 400, 200])
    expect(Array.from(packet.rectInstances.slice(20, 40))).toEqual([5, 6, 7, 8, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 400, 200])
    expect(Array.from(packet.rects.slice(0, 8))).toEqual([1, 2, 3, 4, 1, 0.5, 0.25, 1])
    expect(Array.from(packet.textMetrics.slice(0, 8))).toEqual([10, 11, 100, 20, 3, 2, 0, 1])
  })
})
