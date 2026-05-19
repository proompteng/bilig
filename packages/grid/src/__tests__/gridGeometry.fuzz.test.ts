import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { createGridGeometrySnapshot } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'

describe('grid geometry fuzz', () => {
  it('should keep generated camera geometry clamped and hit-test visible cells', async () => {
    await runProperty({
      suite: 'grid/geometry/camera-hit-test-invariants',
      arbitrary: fc.record({
        scrollLeft: fc.integer({ min: -2_000, max: 20_000 }),
        scrollTop: fc.integer({ min: -2_000, max: 20_000 }),
        hostWidth: fc.integer({ min: 80, max: 1_600 }),
        hostHeight: fc.integer({ min: 80, max: 1_200 }),
        freezeRows: fc.integer({ min: 0, max: 8 }),
        freezeCols: fc.integer({ min: 0, max: 8 }),
        col: fc.integer({ min: 0, max: 20 }),
        row: fc.integer({ min: 0, max: 40 }),
      }),
      predicate: async (input) => {
        const snapshot = createGridGeometrySnapshot({
          ...input,
          seq: 1,
          sheetName: 'Sheet1',
          dpr: 2,
          gridMetrics: getGridMetrics(),
          updatedAt: 1_000,
        })

        expect(snapshot.camera.bodyScrollX).toBeGreaterThanOrEqual(0)
        expect(snapshot.camera.bodyScrollY).toBeGreaterThanOrEqual(0)
        expect(snapshot.camera.bodyViewportWidth).toBeGreaterThanOrEqual(0)
        expect(snapshot.camera.bodyViewportHeight).toBeGreaterThanOrEqual(0)
        expect(snapshot.camera.panes).toHaveLength(9)

        const rect = snapshot.cellScreenRect(input.col, input.row)
        if (rect) {
          expect(rect.width).toBeGreaterThan(0)
          expect(rect.height).toBeGreaterThan(0)
        }

        const paneRect =
          snapshot.cellScreenRectForPane(input.col, input.row, 'frozen-cells') ??
          snapshot.cellScreenRectForPane(input.col, input.row, 'frozen-rows') ??
          snapshot.cellScreenRectForPane(input.col, input.row, 'frozen-columns') ??
          snapshot.cellScreenRectForPane(input.col, input.row, 'body')
        if (paneRect) {
          expect(snapshot.hitTestScreenPoint({ x: paneRect.x + paneRect.width / 2, y: paneRect.y + paneRect.height / 2 })).toEqual({
            col: input.col,
            row: input.row,
          })
        }
      },
      parameters: { numRuns: 120 },
    })
  })
})
