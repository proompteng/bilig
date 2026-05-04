import { describe, expect, it } from 'vitest'
import {
  resolveWorkbookAgentInspectionTarget,
  resolveWorkbookAgentSelectionRange,
  resolveWorkbookAgentVisibleRange,
  workbookAgentViewportAroundAddress,
  workbookAgentViewportToRange,
} from './workbook-agent-context-geometry.js'

describe('workbook agent context geometry', () => {
  const context = {
    selection: {
      sheetName: 'Sheet1',
      address: 'B2',
      range: {
        startAddress: 'B2',
        endAddress: 'D4',
      },
    },
    viewport: {
      rowStart: 1,
      rowEnd: 20,
      colStart: 1,
      colEnd: 5,
    },
  }

  it('derives selection and visible ranges from workbook ui context', () => {
    expect(resolveWorkbookAgentSelectionRange(context)).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'D4',
    })
    expect(resolveWorkbookAgentVisibleRange(context)).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'F21',
    })
    expect(workbookAgentViewportToRange('Sheet1', context.viewport)).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'F21',
    })
  })

  it('reuses context selection when inspection args are partial', () => {
    expect(resolveWorkbookAgentInspectionTarget(context, {})).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      range: {
        startAddress: 'B2',
        endAddress: 'D4',
      },
    })
  })

  it('recenters the viewport around a target address', () => {
    expect(
      workbookAgentViewportAroundAddress('Sheet1', 'D10', {
        rowStart: 1,
        rowEnd: 20,
        colStart: 1,
        colEnd: 5,
      }),
    ).toEqual({
      rowStart: 9,
      rowEnd: 28,
      colStart: 3,
      colEnd: 7,
    })
  })
})
