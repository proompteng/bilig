import { describe, expect, it } from 'vitest'
import type { WorkbookAxisEntryRecord } from '../workbook-metadata-types.js'
import {
  getAxisMetadataRecord,
  listAxisEntries,
  materializeAxisEntries,
  moveAxisEntries,
  snapshotAxisEntriesInRange,
  spliceAxisEntries,
  syncAxisMetadataBucket,
} from '../workbook-axis-records.js'

function createEntryFactory(prefix: string) {
  let nextId = 1
  return (): WorkbookAxisEntryRecord => ({
    id: `${prefix}-${nextId++}`,
    size: null,
    hidden: null,
  })
}

describe('workbook axis records', () => {
  it('materializes sparse entries and snapshots only defined records', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = []
    const createEntry = createEntryFactory('row')

    expect(materializeAxisEntries(entries, 1, 2, createEntry)).toEqual([
      { id: 'row-1', index: 1 },
      { id: 'row-2', index: 2 },
    ])

    entries[1]!.size = 120
    entries[2]!.hidden = true

    expect(listAxisEntries(entries)).toEqual([
      { id: 'row-1', index: 1, size: 120 },
      { id: 'row-2', index: 2, hidden: true },
    ])
    expect(snapshotAxisEntriesInRange(entries, 0, 4)).toEqual([
      { id: 'row-1', index: 1, size: 120 },
      { id: 'row-2', index: 2, hidden: true },
    ])
  })

  it('splices and moves sparse axis entries while preserving provided ids', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = []
    const createEntry = createEntryFactory('column')

    expect(spliceAxisEntries(entries, 0, 0, 2, createEntry, [{ id: 'column-existing', index: 1, size: 90, hidden: false }])).toEqual([])
    expect(listAxisEntries(entries)).toEqual([{ id: 'column-existing', index: 1, size: 90, hidden: false }])

    moveAxisEntries(entries, 1, 1, 0, createEntry)
    expect(listAxisEntries(entries)).toEqual([{ id: 'column-existing', index: 0, size: 90, hidden: false }])
  })

  it('coalesces contiguous metadata ranges and rejects mixed ranges', () => {
    const entries: Array<WorkbookAxisEntryRecord | undefined> = [
      { id: 'row-1', size: 30, hidden: false },
      { id: 'row-2', size: 30, hidden: false },
      undefined,
      { id: 'row-3', size: 30, hidden: true },
    ]
    const bucket = new Map()

    syncAxisMetadataBucket(bucket, 'Sheet1', entries)

    expect([...bucket.values()].toSorted((left, right) => left.start - right.start)).toEqual([
      { sheetName: 'Sheet1', start: 0, count: 2, size: 30, hidden: false },
      { sheetName: 'Sheet1', start: 3, count: 1, size: 30, hidden: true },
    ])
    expect(getAxisMetadataRecord(entries, 'Sheet1', 0, 2)).toEqual({
      sheetName: 'Sheet1',
      start: 0,
      count: 2,
      size: 30,
      hidden: false,
    })
    expect(getAxisMetadataRecord(entries, 'Sheet1', 0, 4)).toBeUndefined()
  })
})
