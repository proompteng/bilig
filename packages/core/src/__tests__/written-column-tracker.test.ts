import { describe, expect, it } from 'vitest'
import { createWrittenColumnTracker, markWrittenColumn, materializeWrittenColumns } from '../written-column-tracker.js'

describe('written column tracker', () => {
  it('materializes sorted unique columns across bitset and sparse ranges', () => {
    const tracker = createWrittenColumnTracker()
    for (const col of [35, 2, 0, 63, 35, 29, 30, 2]) {
      markWrittenColumn(tracker, col)
    }

    expect(materializeWrittenColumns(tracker)).toEqual(Uint32Array.from([0, 2, 29, 30, 35, 63]))
  })
})
