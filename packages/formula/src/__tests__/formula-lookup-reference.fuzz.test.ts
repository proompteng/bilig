import * as fc from 'fast-check'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import { getLookupBuiltin, normalizeExactLookupNumber, type RangeBuiltinArgument } from '../index.js'

const MATCH = getLookupBuiltin('MATCH')!
const XMATCH = getLookupBuiltin('XMATCH')!
const XLOOKUP = getLookupBuiltin('XLOOKUP')!

const scalarArbitrary = fc.oneof<CellValue>(
  fc.integer({ min: -8, max: 8 }).map(num),
  fc.constantFrom('alpha', 'BRAVO', 'charlie', '').map(text),
  fc.boolean().map(bool),
  fc.constant(empty()),
)

const exactLookupCaseArbitrary = fc.array(scalarArbitrary, { minLength: 1, maxLength: 8 }).chain((values) =>
  fc.record({
    values: fc.constant(values),
    lookupIndex: fc.integer({ min: 0, max: values.length - 1 }),
    searchMode: fc.constantFrom(1, -1),
  }),
)

const sortedNumericLookupCaseArbitrary = fc
  .uniqueArray(fc.integer({ min: -20, max: 20 }), {
    minLength: 1,
    maxLength: 8,
  })
  .map((values) => values.toSorted((left, right) => left - right))
  .chain((values) =>
    fc.record({
      values: fc.constant(values),
      lookupValue: fc.integer({ min: -25, max: 25 }),
    }),
  )

describe('formula lookup reference fuzz', () => {
  it('keeps exact MATCH, XMATCH, and XLOOKUP aligned with scalar comparison rules', async () => {
    await runProperty({
      suite: 'formula/lookup-reference/exact-vector-oracle',
      arbitrary: exactLookupCaseArbitrary,
      predicate: ({ values, lookupIndex, searchMode }) => {
        const lookupValue = values[lookupIndex]!
        const lookupRange = columnRange(values)
        const returnRange = columnRange(values.map((_value, index) => num(index + 100)))
        const firstMatchIndex = findExactMatchIndex(values, lookupValue, 1)
        const expectedIndex = findExactMatchIndex(values, lookupValue, searchMode)

        expect(firstMatchIndex).not.toBe(-1)
        expect(expectedIndex).not.toBe(-1)
        expect(MATCH(lookupValue, lookupRange, num(0))).toEqual(num(firstMatchIndex + 1))
        expect(XMATCH(lookupValue, lookupRange, num(0), num(searchMode))).toEqual(num(expectedIndex + 1))
        expect(XLOOKUP(lookupValue, lookupRange, returnRange, text('missing'), num(0), num(searchMode))).toEqual(num(expectedIndex + 100))
      },
    })
  })

  it('keeps approximate MATCH and XLOOKUP binary modes aligned on sorted numeric vectors', async () => {
    await runProperty({
      suite: 'formula/lookup-reference/approximate-numeric-oracle',
      arbitrary: sortedNumericLookupCaseArbitrary,
      predicate: ({ values, lookupValue }) => {
        const ascendingRange = columnRange(values.map(num))
        const descending = values.toReversed()
        const descendingRange = columnRange(descending.map(num))
        const ascendingReturnRange = columnRange(values.map((value) => text(`a${value}`)))
        const descendingReturnRange = columnRange(descending.map((value) => text(`d${value}`)))
        const lowerBoundIndex = findLastIndex(values, (value) => value <= lookupValue)
        const upperBoundDescendingIndex = findLastIndex(descending, (value) => value >= lookupValue)

        assertPositionOrNotAvailable(MATCH(num(lookupValue), ascendingRange, num(1)), lowerBoundIndex)
        assertPositionOrNotAvailable(XMATCH(num(lookupValue), ascendingRange, num(-1), num(2)), lowerBoundIndex)
        expect(XLOOKUP(num(lookupValue), ascendingRange, ascendingReturnRange, text('missing'), num(-1), num(2))).toEqual(
          lowerBoundIndex === -1 ? text('missing') : text(`a${values[lowerBoundIndex]}`),
        )

        assertPositionOrNotAvailable(MATCH(num(lookupValue), descendingRange, num(-1)), upperBoundDescendingIndex)
        assertPositionOrNotAvailable(XMATCH(num(lookupValue), descendingRange, num(1), num(-2)), upperBoundDescendingIndex)
        expect(XLOOKUP(num(lookupValue), descendingRange, descendingReturnRange, text('missing'), num(1), num(-2))).toEqual(
          upperBoundDescendingIndex === -1 ? text('missing') : text(`d${descending[upperBoundDescendingIndex]}`),
        )
      },
    })
  })
})

function columnRange(values: CellValue[]): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows: values.length, cols: 1 }
}

function findExactMatchIndex(values: readonly CellValue[], lookupValue: CellValue, searchMode: 1 | -1): number {
  const first = searchMode === -1 ? values.length - 1 : 0
  const last = searchMode === -1 ? -1 : values.length
  const step = searchMode === -1 ? -1 : 1
  for (let index = first; index !== last; index += step) {
    if (sameLookupScalar(values[index], lookupValue)) {
      return index
    }
  }
  return -1
}

function sameLookupScalar(left: CellValue, right: CellValue): boolean {
  if ((left.tag === ValueTag.String || left.tag === ValueTag.Empty) && (right.tag === ValueTag.String || right.tag === ValueTag.Empty)) {
    return lookupText(left).toUpperCase() === lookupText(right).toUpperCase()
  }
  const leftNumber = lookupNumber(left)
  const rightNumber = lookupNumber(right)
  return (
    leftNumber !== undefined &&
    rightNumber !== undefined &&
    normalizeExactLookupNumber(leftNumber) === normalizeExactLookupNumber(rightNumber)
  )
}

function findLastIndex<T>(values: readonly T[], predicate: (value: T) => boolean): number {
  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (predicate(values[index])) {
      return index
    }
  }
  return -1
}

function assertPositionOrNotAvailable(actual: unknown, expectedIndex: number): void {
  expect(actual).toEqual(expectedIndex === -1 ? err(ErrorCode.NA) : num(expectedIndex + 1))
}

function lookupText(value: CellValue): string {
  if (value.tag === ValueTag.String) {
    return value.value
  }
  if (value.tag === ValueTag.Empty) {
    return ''
  }
  return ''
}

function lookupNumber(value: CellValue): number | undefined {
  if (value.tag === ValueTag.Number) {
    return value.value
  }
  if (value.tag === ValueTag.Boolean) {
    return value.value ? 1 : 0
  }
  if (value.tag === ValueTag.Empty) {
    return 0
  }
  return undefined
}

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function bool(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function empty(): CellValue {
  return { tag: ValueTag.Empty }
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}
