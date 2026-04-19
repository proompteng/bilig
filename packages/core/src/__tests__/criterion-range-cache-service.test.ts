import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createDepPatternStore } from '../deps/dep-pattern-store.js'
import { createRegionGraph } from '../deps/region-graph.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { createCriterionRangeCacheService } from '../engine/services/criterion-range-cache-service.js'
import { createEngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'

function setStoredCellValue(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetName: string,
  address: string,
  value: { tag: ValueTag.String; value: string } | { tag: ValueTag.Number; value: number } | { tag: ValueTag.Boolean; value: boolean },
): void {
  const cellIndex = workbook.ensureCell(sheetName, address)
  workbook.cellStore.setValue(cellIndex, value, value.tag === ValueTag.String ? strings.intern(value.value) : 0)
}

function isErrorValue(value: unknown): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return typeof value === 'object' && value !== null && 'tag' in value && value.tag === ValueTag.Error && 'code' in value
}

describe('createCriterionRangeCacheService', () => {
  it('rejects empty criteria requests', () => {
    const workbook = new WorkbookStore('criteria-cache-empty')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    expect(criterionCache.getOrBuildMatchingRows({ criteriaPairs: [] })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('reuses matching row sets for identical requests and invalidates on source writes', () => {
    const workbook = new WorkbookStore('criteria-cache')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;['A', 'B', 'A', 'B', 'A', 'B'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const first = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
      ],
    })
    if (isErrorValue(first)) {
      throw new Error(`unexpected criteria cache error: ${first.code}`)
    }
    expect([...first.rows]).toEqual([0, 2, 4])

    const second = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
      ],
    })
    if (isErrorValue(second)) {
      throw new Error(`unexpected criteria cache error: ${second.code}`)
    }
    expect(second.rows).toBe(first.rows)

    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.String,
      value: 'A',
    })

    const third = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
      ],
    })
    if (isErrorValue(third)) {
      throw new Error(`unexpected criteria cache error: ${third.code}`)
    }
    expect(third.rows).not.toBe(first.rows)
    expect([...third.rows]).toEqual([0, 1, 2, 4])
  })

  it('supports compiled operator criteria and validates matching range lengths', () => {
    const workbook = new WorkbookStore('criteria-cache-operators')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3, 4, 5, 6].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })
    ;['x', 'x', 'y', 'x', 'y', 'x'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const matching = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: '>2', stringId: strings.intern('>2') },
        },
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
          criteria: { tag: ValueTag.String, value: 'x', stringId: strings.intern('x') },
        },
      ],
    })
    if (isErrorValue(matching)) {
      throw new Error(`unexpected criteria cache error: ${matching.code}`)
    }
    expect([...matching.rows]).toEqual([3, 5])

    const invalid = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: '>2', stringId: strings.intern('>2') },
        },
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 4, col: 1, length: 5 },
          criteria: { tag: ValueTag.String, value: 'x', stringId: strings.intern('x') },
        },
      ],
    })
    expect(invalid).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers remaining numeric comparator branches and non-numeric rejections', () => {
    const workbook = new WorkbookStore('criteria-cache-comparators')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3, 4, 5, 6].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A7', {
      tag: ValueTag.String,
      value: 'skip',
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const gteMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 6, col: 0, length: 7 },
          criteria: { tag: ValueTag.String, value: '>=4', stringId: strings.intern('>=4') },
        },
      ],
    })
    if (isErrorValue(gteMatches)) {
      throw new Error(`unexpected >= criteria error: ${gteMatches.code}`)
    }
    expect([...gteMatches.rows]).toEqual([3, 4, 5])

    const lteMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 6, col: 0, length: 7 },
          criteria: { tag: ValueTag.String, value: '<=2', stringId: strings.intern('<=2') },
        },
      ],
    })
    if (isErrorValue(lteMatches)) {
      throw new Error(`unexpected <= criteria error: ${lteMatches.code}`)
    }
    expect([...lteMatches.rows]).toEqual([0, 1])

    const lessThanMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 6, col: 0, length: 7 },
          criteria: { tag: ValueTag.String, value: '<4', stringId: strings.intern('<4') },
        },
      ],
    })
    if (isErrorValue(lessThanMatches)) {
      throw new Error(`unexpected < criteria error: ${lessThanMatches.code}`)
    }
    expect([...lessThanMatches.rows]).toEqual([0, 1, 2])
  })

  it('matches empty, boolean, wildcard, and error-heavy criteria correctly', () => {
    const workbook = new WorkbookStore('criteria-cache-generic')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    workbook.ensureCell('Sheet1', 'A1')
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.String,
      value: '',
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A3', {
      tag: ValueTag.Boolean,
      value: false,
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A4', {
      tag: ValueTag.Number,
      value: 0,
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A5', {
      tag: ValueTag.String,
      value: 'northwest',
    })
    const errorCell = workbook.ensureCell('Sheet1', 'A6')
    workbook.cellStore.setValue(errorCell, { tag: ValueTag.Error, code: ErrorCode.Ref })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const emptyMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.Empty },
        },
      ],
    })
    if (isErrorValue(emptyMatches)) {
      throw new Error(`unexpected empty-match error: ${emptyMatches.code}`)
    }
    expect([...emptyMatches.rows]).toEqual([0, 1])

    const falseMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.Boolean, value: false },
        },
      ],
    })
    if (isErrorValue(falseMatches)) {
      throw new Error(`unexpected false-match error: ${falseMatches.code}`)
    }
    expect([...falseMatches.rows]).toEqual([0, 2, 3])

    const wildcardMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'north*', stringId: strings.intern('north*') },
        },
      ],
    })
    if (isErrorValue(wildcardMatches)) {
      throw new Error(`unexpected wildcard-match error: ${wildcardMatches.code}`)
    }
    expect([...wildcardMatches.rows]).toEqual([4])

    const errorCriterion = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.Error, code: ErrorCode.Name },
        },
      ],
    })
    if (isErrorValue(errorCriterion)) {
      throw new Error(`unexpected error-criterion failure: ${errorCriterion.code}`)
    }
    expect([...errorCriterion.rows]).toEqual([])
  })
})
