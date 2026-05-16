import { describe, expect, it } from 'vitest'
import { normalizeWorkbookChangeRowModel } from '../workbook-change-row.js'

const baseRow = {
  revision: 12,
  actorUserId: 'alex@example.com',
  clientMutationId: 'mutation-12',
  eventKind: 'setCellValue',
  summary: 'Updated Sheet1!A1',
  sheetId: 1,
  sheetName: 'Sheet1',
  anchorAddress: 'A1',
  rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
  undoBundleJson: {
    kind: 'engineOps',
    ops: [],
  },
  revertedByRevision: null,
  revertsRevision: null,
  createdAt: 1_768_348_800_000,
} as const

describe('workbook change row model', () => {
  it('normalizes valid Zero workbook_change rows', () => {
    expect(normalizeWorkbookChangeRowModel(baseRow)).toMatchObject({
      revision: 12,
      actorUserId: 'alex@example.com',
      eventKind: 'setCellValue',
      rangeJsonInvalid: false,
    })
  })

  it('rejects rows with event kinds outside the shared workbook event model', () => {
    expect(
      normalizeWorkbookChangeRowModel({
        ...baseRow,
        eventKind: 'legacyPatch',
      }),
    ).toBeNull()
  })

  it('preserves malformed range metadata as an explicit trust failure', () => {
    expect(
      normalizeWorkbookChangeRowModel({
        ...baseRow,
        eventKind: 'insertRows',
        anchorAddress: 'A3',
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'row-band' },
      }),
    ).toMatchObject({
      eventKind: 'insertRows',
      rangeJson: null,
      rangeJsonInvalid: true,
    })
  })
})
