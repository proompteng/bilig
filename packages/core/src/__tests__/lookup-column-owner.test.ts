import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { createEngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'
import {
  applyLookupColumnOwnerLiteralWrite,
  buildLookupColumnOwner,
  findExactMatchInRange,
  isLookupColumnOwner,
  sliceOffsetBounds,
  summarizeApproximateRange,
  summarizeExactRange,
  supportsNumericApproximateRange,
  supportsTextApproximateRange,
} from '../engine/services/lookup-column-owner.js'

function setStoredCellValue(workbook: WorkbookStore, strings: StringPool, sheetName: string, address: string, value: CellValue): void {
  const cellIndex = workbook.ensureCell(sheetName, address)
  workbook.cellStore.setValue(cellIndex, value, value.tag === ValueTag.String ? strings.intern(value.value) : 0)
}

describe('lookup column owner helpers', () => {
  it('builds numeric owners with exact and approximate summaries', () => {
    const workbook = new WorkbookStore('lookup-owner-numeric')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')
    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, { tag: ValueTag.Number, value })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const owner = buildLookupColumnOwner({
      owner: runtimeColumnStore.getColumnOwner({ sheetName: 'Sheet1', col: 0 }),
      normalizeStringId: runtimeColumnStore.normalizeStringId,
    })

    expect(isLookupColumnOwner(owner)).toBe(true)
    expect(owner?.rowStart).toBe(0)
    expect(owner?.rowEnd).toBe(3)
    expect(sliceOffsetBounds(owner!, 1, 3)).toEqual({ start: 1, end: 3 })
    expect(sliceOffsetBounds(owner!, -1, 1)).toBeUndefined()
    expect(findExactMatchInRange(owner!, 'n:5', 0, 3, 1)).toBe(2)
    expect(findExactMatchInRange(owner!, 'n:5', 0, 3, -1)).toBe(2)
    expect(findExactMatchInRange(owner!, 'n:99', 0, 3, 1)).toBeUndefined()

    expect(summarizeExactRange(owner!, 0, 3)).toEqual({
      comparableKind: 'numeric',
      uniformStart: 1,
      uniformStep: 2,
    })
    expect(summarizeApproximateRange(owner!, 0, 3)).toEqual({
      comparableKind: 'numeric',
      uniformStart: 1,
      uniformStep: 2,
      sortedAscending: true,
      sortedDescending: false,
    })
    expect(supportsNumericApproximateRange(owner!, 0, 3, 1)).toBe(true)
    expect(supportsNumericApproximateRange(owner!, 0, 3, -1)).toBe(false)
  })

  it('updates text owners in place and tracks text-sorted summaries', () => {
    const workbook = new WorkbookStore('lookup-owner-text')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')
    ;['apple', 'banana', 'pear'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, { tag: ValueTag.String, value })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const owner = buildLookupColumnOwner({
      owner: runtimeColumnStore.getColumnOwner({ sheetName: 'Sheet1', col: 0 }),
      normalizeStringId: runtimeColumnStore.normalizeStringId,
    })
    expect(owner).toBeDefined()
    expect(summarizeExactRange(owner!, 0, 2)).toEqual({
      comparableKind: 'text',
      uniformStart: undefined,
      uniformStep: undefined,
    })
    expect(summarizeApproximateRange(owner!, 0, 2)).toEqual({
      comparableKind: 'text',
      uniformStart: undefined,
      uniformStep: undefined,
      sortedAscending: true,
      sortedDescending: false,
    })
    expect(supportsTextApproximateRange(owner!, 0, 2, 1)).toBe(true)
    expect(supportsTextApproximateRange(owner!, 0, 2, -1)).toBe(false)
    expect(findExactMatchInRange(owner!, 's:PEAR', 0, 2, 1)).toBe(2)

    const bananaId = strings.intern('banana')
    const blueberryId = strings.intern('blueberry')
    expect(
      applyLookupColumnOwnerLiteralWrite({
        owner: owner!,
        write: {
          row: 1,
          oldValue: { tag: ValueTag.String, value: 'banana' },
          newValue: { tag: ValueTag.String, value: 'blueberry' },
          oldStringId: bananaId,
          newStringId: blueberryId,
        },
        normalizeStringId: runtimeColumnStore.normalizeStringId,
      }),
    ).toBe(true)
    expect(findExactMatchInRange(owner!, 's:BLUEBERRY', 0, 2, 1)).toBe(1)
    expect(summarizeApproximateRange(owner!, 0, 2)?.sortedAscending).toBe(true)
    expect(
      applyLookupColumnOwnerLiteralWrite({
        owner: owner!,
        write: {
          row: 5,
          oldValue: { tag: ValueTag.String, value: 'pear' },
          newValue: { tag: ValueTag.String, value: 'plum' },
        },
        normalizeStringId: runtimeColumnStore.normalizeStringId,
      }),
    ).toBe(false)
  })

  it('patches numeric approximate summaries incrementally after monotonic writes', () => {
    const workbook = new WorkbookStore('lookup-owner-numeric-incremental')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')
    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, { tag: ValueTag.Number, value })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const owner = buildLookupColumnOwner({
      owner: runtimeColumnStore.getColumnOwner({ sheetName: 'Sheet1', col: 0 }),
      normalizeStringId: runtimeColumnStore.normalizeStringId,
    })

    expect(owner).toBeDefined()
    expect(summarizeApproximateRange(owner!, 0, 3)).toEqual({
      comparableKind: 'numeric',
      uniformStart: 1,
      uniformStep: 2,
      sortedAscending: true,
      sortedDescending: false,
    })
    expect(owner!.summariesDirty).toBe(false)

    expect(
      applyLookupColumnOwnerLiteralWrite({
        owner: owner!,
        write: {
          row: 3,
          oldValue: { tag: ValueTag.Number, value: 7 },
          newValue: { tag: ValueTag.Number, value: 9 },
        },
        normalizeStringId: runtimeColumnStore.normalizeStringId,
      }),
    ).toBe(true)
    expect(owner!.summariesDirty).toBe(false)
    expect(summarizeApproximateRange(owner!, 0, 3)).toEqual({
      comparableKind: 'numeric',
      uniformStart: undefined,
      uniformStep: undefined,
      sortedAscending: true,
      sortedDescending: false,
    })
  })

  it('handles mixed and invalid owners conservatively', () => {
    const workbook = new WorkbookStore('lookup-owner-mixed')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')
    setStoredCellValue(workbook, strings, 'Sheet1', 'A1', { tag: ValueTag.Number, value: 1 })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', { tag: ValueTag.String, value: 'pear' })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A3', { tag: ValueTag.Error, code: ErrorCode.Div0 })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const owner = buildLookupColumnOwner({
      owner: runtimeColumnStore.getColumnOwner({ sheetName: 'Sheet1', col: 0 }),
      normalizeStringId: runtimeColumnStore.normalizeStringId,
    })

    expect(owner).toBeDefined()
    expect(summarizeExactRange(owner!, 0, 2)).toEqual({
      comparableKind: 'mixed',
      uniformStart: undefined,
      uniformStep: undefined,
    })
    expect(summarizeApproximateRange(owner!, 0, 2)).toEqual({
      comparableKind: undefined,
      uniformStart: undefined,
      uniformStep: undefined,
      sortedAscending: false,
      sortedDescending: false,
    })
    expect(supportsNumericApproximateRange(owner!, 0, 2, 1)).toBe(false)
    expect(supportsTextApproximateRange(owner!, 0, 2, 1)).toBe(false)
    expect(summarizeExactRange(owner!, 10, 11)).toBeUndefined()
    expect(summarizeApproximateRange(owner!, 10, 11)).toBeUndefined()
    expect(
      buildLookupColumnOwner({
        owner: {
          sheetName: 'Sheet1',
          col: 0,
          columnVersion: 0,
          structureVersion: 0,
          sheetColumnVersions: new Uint32Array(0),
          pages: new Map(),
        },
        normalizeStringId: runtimeColumnStore.normalizeStringId,
      }),
    ).toBeUndefined()
    expect(isLookupColumnOwner(null)).toBe(false)
    expect(isLookupColumnOwner({})).toBe(false)
  })
})
