import { describe, expect, it } from 'vitest'
import { buildWorkbookAgentContext, singleCellAgentSelectionRange } from '../workbook-agent-context.js'

describe('workbook agent context', () => {
  it('builds agent context from explicit selection geometry instead of a parsed label', () => {
    expect(
      buildWorkbookAgentContext({
        selection: {
          sheetName: 'Sheet1',
          address: 'B18',
        },
        selectionRange: {
          startAddress: 'A6',
          endAddress: 'H15',
        },
        viewport: {
          rowStart: 5,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      }),
    ).toEqual({
      selection: {
        sheetName: 'Sheet1',
        address: 'B18',
        range: {
          startAddress: 'A6',
          endAddress: 'H15',
        },
      },
      viewport: {
        rowStart: 5,
        rowEnd: 20,
        colStart: 0,
        colEnd: 10,
      },
    })
  })

  it('creates a single-cell range fallback from the active address', () => {
    expect(
      singleCellAgentSelectionRange({
        sheetName: 'Sheet7',
        address: 'F9',
      }),
    ).toEqual({
      startAddress: 'F9',
      endAddress: 'F9',
    })
  })
})
