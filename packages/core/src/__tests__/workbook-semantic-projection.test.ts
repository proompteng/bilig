import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  normalizeWorkbookSnapshotForSemanticComparison,
  projectWorkbookSemanticSnapshot,
  workbookSemanticSnapshotsEqual,
} from '../semantics/index.js'

const baseSnapshot: WorkbookSnapshot = {
  version: 1,
  workbook: {
    name: 'semantic-fixture',
  },
  sheets: [
    {
      name: 'Sheet1',
      order: 0,
      cells: [],
    },
  ],
}

describe('workbook semantic projection', () => {
  it('normalizes metadata ordering and equivalent range coverage for engine comparisons', () => {
    const snapshot: WorkbookSnapshot = {
      ...baseSnapshot,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          styles: [
            { id: 'style-b', font: { bold: true } },
            { id: 'style-a', fill: { backgroundColor: '#ffffff' } },
          ],
          definedNames: [
            { name: 'Totals', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'C2' } },
            { name: 'Inputs', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            rows: [
              { id: 'row-2', index: 2, size: 24 },
              { id: 'row-1', index: 1, size: 18 },
            ],
            styleRanges: [
              { range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C2' }, styleId: 'style-a' },
              { range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, styleId: 'style-a' },
            ],
          },
          cells: [
            { address: 'B1', value: 2 },
            { address: 'A1', value: 1 },
          ],
        },
      ],
    }

    const normalized = normalizeWorkbookSnapshotForSemanticComparison(snapshot)

    expect(normalized.workbook.metadata?.styles?.map((style) => style.id)).toEqual(['style-a', 'style-b'])
    expect(normalized.workbook.metadata?.definedNames?.map((definedName) => definedName.name)).toEqual(['Inputs', 'Totals'])
    expect(normalized.sheets[0]?.metadata?.rows).toEqual([
      { index: 1, size: 18 },
      { index: 2, size: 24 },
    ])
    expect(normalized.sheets[0]?.metadata?.styleRanges).toEqual([
      { range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C2' }, styleId: 'style-a' },
    ])
  })

  it('projects stable workbook semantics independent of generated style ids and defaults', () => {
    const left: WorkbookSnapshot = {
      ...baseSnapshot,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          styles: [{ id: 'left-style', fill: { backgroundColor: '#dbeafe' }, font: { bold: true } }],
          charts: [
            {
              id: 'chart-1',
              sheetName: 'Sheet1',
              address: 'E1',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
              chartType: 'column',
              rows: 8,
              cols: 4,
            },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            styleRanges: [{ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, styleId: 'left-style' }],
            freezePane: { rows: 1, cols: 1, topLeftCell: 'B2', activePane: 'bottomRight' },
          },
          cells: [
            { address: 'B1', formula: 'A1*2' },
            { address: 'A1', value: 12, format: '$0.00' },
          ],
        },
      ],
    }
    const right: WorkbookSnapshot = {
      ...left,
      workbook: {
        name: 'semantic-fixture',
        metadata: {
          styles: [{ id: 'right-style', fill: { backgroundColor: '#dbeafe' }, font: { bold: true } }],
          charts: [
            {
              id: 'chart-1',
              sheetName: 'Sheet1',
              address: 'E1',
              source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
              chartType: 'column',
              seriesOrientation: 'columns',
              rows: 8,
              cols: 4,
            },
          ],
        },
      },
      sheets: [
        {
          name: 'Sheet1',
          order: 0,
          metadata: {
            styleRanges: [{ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, styleId: 'right-style' }],
            freezePane: { rows: 1, cols: 1 },
          },
          cells: [
            { address: 'A1', value: 12, format: '$0.00' },
            { address: 'B1', formula: 'A1*2' },
          ],
        },
      ],
    }

    expect(projectWorkbookSemanticSnapshot(left)).toEqual(projectWorkbookSemanticSnapshot(right))
    expect(workbookSemanticSnapshotsEqual(left, right)).toBe(true)
  })
})
