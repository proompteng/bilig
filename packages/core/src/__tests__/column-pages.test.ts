import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  createRuntimeColumnPage,
  decodeValueTag,
  materializeRuntimeColumnPageValue,
  patchRuntimeColumnPageValue,
  readRuntimeColumnPageEntry,
} from '../indexes/column-pages.js'

describe('column-pages', () => {
  it('creates zeroed pages and treats missing tags as empty', () => {
    const page = createRuntimeColumnPage(10, 4)

    expect(page.rowStart).toBe(10)
    expect(page.tags).toHaveLength(4)
    expect(page.numbers).toHaveLength(4)
    expect(page.stringIds).toHaveLength(4)
    expect(page.errors).toHaveLength(4)
    expect(decodeValueTag(undefined)).toBe(ValueTag.Empty)
    expect(readRuntimeColumnPageEntry(page, 10)).toEqual({
      rawTag: ValueTag.Empty,
      number: 0,
      stringId: 0,
      error: 0,
    })
    expect(materializeRuntimeColumnPageValue(page, 10, () => undefined)).toEqual({ tag: ValueTag.Empty })
  })

  it('decodes every supported tag and falls back to empty for unknown values', () => {
    expect(decodeValueTag(0)).toBe(ValueTag.Empty)
    expect(decodeValueTag(1)).toBe(ValueTag.Number)
    expect(decodeValueTag(2)).toBe(ValueTag.Boolean)
    expect(decodeValueTag(3)).toBe(ValueTag.String)
    expect(decodeValueTag(4)).toBe(ValueTag.Error)
    expect(decodeValueTag(99)).toBe(ValueTag.Empty)
  })

  it('patches and materializes number, boolean, string, and error values', () => {
    const page = createRuntimeColumnPage(20, 5)

    patchRuntimeColumnPageValue({
      page,
      row: 20,
      value: { tag: ValueTag.Number, value: -0 },
    })
    patchRuntimeColumnPageValue({
      page,
      row: 21,
      value: { tag: ValueTag.Boolean, value: true },
    })
    patchRuntimeColumnPageValue({
      page,
      row: 22,
      value: { tag: ValueTag.String, value: 'alpha', stringId: 7 },
      stringId: 7,
    })
    patchRuntimeColumnPageValue({
      page,
      row: 23,
      value: { tag: ValueTag.Error, code: ErrorCode.Ref },
    })

    expect(readRuntimeColumnPageEntry(page, 20)).toEqual({
      rawTag: ValueTag.Number,
      number: 0,
      stringId: 0,
      error: 0,
    })
    expect(materializeRuntimeColumnPageValue(page, 20, () => undefined)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(materializeRuntimeColumnPageValue(page, 21, () => undefined)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(materializeRuntimeColumnPageValue(page, 22, (stringId) => (stringId === 7 ? 'alpha' : undefined))).toEqual({
      tag: ValueTag.String,
      value: 'alpha',
      stringId: 7,
    })
    expect(materializeRuntimeColumnPageValue(page, 23, () => undefined)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
  })

  it('clears patched entries and preserves empty string semantics for zero string ids', () => {
    const page = createRuntimeColumnPage(30, 3)

    patchRuntimeColumnPageValue({
      page,
      row: 30,
      value: { tag: ValueTag.String, value: '', stringId: 0 },
    })
    expect(materializeRuntimeColumnPageValue(page, 30, () => 'ignored')).toEqual({
      tag: ValueTag.String,
      value: '',
      stringId: 0,
    })

    patchRuntimeColumnPageValue({
      page,
      row: 30,
      value: { tag: ValueTag.Empty },
    })
    expect(readRuntimeColumnPageEntry(page, 30)).toEqual({
      rawTag: ValueTag.Empty,
      number: 0,
      stringId: 0,
      error: 0,
    })
    expect(materializeRuntimeColumnPageValue(page, 30, () => undefined)).toEqual({ tag: ValueTag.Empty })
  })
})
