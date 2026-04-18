import { describe, expect, it } from 'vitest'
import { buildWorkbookAgentContext, createSingleCellSelectionSnapshot } from '../workbook-agent-context.js'

describe('workbook agent context', () => {
  it('builds agent context from one authoritative selection snapshot', () => {
    expect(
      buildWorkbookAgentContext({
        selection: {
          sheetName: 'Sheet1',
          address: 'B18',
          kind: 'range',
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

  it('creates a single-cell selection snapshot from the active address', () => {
    expect(
      createSingleCellSelectionSnapshot({
        sheetName: 'Sheet7',
        address: 'F9',
      }),
    ).toEqual({
      sheetName: 'Sheet7',
      address: 'F9',
      kind: 'cell',
      range: {
        startAddress: 'F9',
        endAddress: 'F9',
      },
    })
  })
})
