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

  it('rejects unsafe or non-positive history revision identifiers', () => {
    for (const patch of [
      { revision: 0 },
      { revision: -1 },
      { revision: Number.MAX_SAFE_INTEGER + 1 },
      { revertedByRevision: 0 },
      { revertsRevision: -3 },
      { sheetId: 0 },
    ]) {
      expect(
        normalizeWorkbookChangeRowModel({
          ...baseRow,
          ...patch,
        }),
      ).toBeNull()
    }
  })

  it('keeps backfilled epoch-zero timestamps valid while rejecting negative timestamps', () => {
    expect(
      normalizeWorkbookChangeRowModel({
        ...baseRow,
        createdAt: 0,
      }),
    ).toMatchObject({
      createdAt: 0,
    })
    expect(
      normalizeWorkbookChangeRowModel({
        ...baseRow,
        createdAt: -1,
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

  it('marks persisted ranges with malformed addresses as explicit trust failures', () => {
    expect(
      normalizeWorkbookChangeRowModel({
        ...baseRow,
        rangeJson: { sheetName: 'Sheet1', startAddress: 'A0', endAddress: 'A1' },
      }),
    ).toMatchObject({
      rangeJson: null,
      rangeJsonInvalid: true,
    })
  })
})
