import { describe, expect, it } from 'vitest'
import {
  countWorkbookAgentRangeColumns,
  countWorkbookAgentRangeRows,
  countWorkbookAgentRangesCells,
  createWorkbookAgentRangeChunkPlan,
  ensureWorkbookAgentRangeCellLimit,
  workbookAgentRangesIntersect,
} from './workbook-agent-range-chunks.js'

describe('workbook agent range chunks', () => {
  it('counts total cells across multiple ranges', () => {
    expect(
      countWorkbookAgentRangesCells([
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C3' },
      ]),
    ).toBe(7)
  })

  it('normalizes row and column counts for reversed ranges', () => {
    expect(
      countWorkbookAgentRangeRows({
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'B2',
      }),
    ).toBe(4)
    expect(
      countWorkbookAgentRangeColumns({
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'B2',
      }),
    ).toBe(3)
  })

  it('detects intersecting ranges only on the same sheet', () => {
    expect(
      workbookAgentRangesIntersect(
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'D4',
        },
      ),
    ).toBe(true)
    expect(
      workbookAgentRangesIntersect(
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'D4',
        },
      ),
    ).toBe(false)
    expect(
      workbookAgentRangesIntersect(
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        {
          sheetName: 'Sheet2',
          startAddress: 'A1',
          endAddress: 'B2',
        },
      ),
    ).toBe(false)
  })

  it('rejects ranges that exceed the tool cell limit', () => {
    expect(() =>
      ensureWorkbookAgentRangeCellLimit(
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'C3',
        },
        4,
      ),
    ).toThrow('Range Sheet1!A1:C3 has 9 cells; tool limit is 4 cells per call')
  })

  it('creates stable chunk plans for oversized ranges', () => {
    const plan = createWorkbookAgentRangeChunkPlan(
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B5',
      },
      4,
    )
    expect(plan.totalCells).toBe(10)
    expect(plan.chunkCount).toBe(3)
    expect(plan.chunks.map((chunk) => `${chunk.startAddress}:${chunk.endAddress}`)).toEqual(['A1:B2', 'A3:B4', 'A5:B5'])
  })
})
