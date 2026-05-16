import { describe, expect, it } from 'vitest'
import {
  normalizeNumberFormatInput,
  normalizeStylePatch,
  type ServerCellNumberFormatInput,
  type ServerCellStylePatchInput,
} from '../server-mutator-format-payloads.js'

describe('server mutator format payload normalization', () => {
  it('omits undefined style patch fields while preserving explicit clears', () => {
    const patch = {
      fill: {
        backgroundColor: undefined,
      },
      font: {
        family: 'Inter',
        size: undefined,
        bold: null,
        italic: false,
      },
      alignment: {
        horizontal: 'center',
        vertical: undefined,
        wrap: true,
        indent: null,
      },
      borders: {
        top: {
          style: 'solid',
          weight: undefined,
          color: '#111111',
        },
        right: null,
      },
    } satisfies ServerCellStylePatchInput

    expect(normalizeStylePatch(patch)).toEqual({
      fill: {},
      font: {
        family: 'Inter',
        bold: null,
        italic: false,
      },
      alignment: {
        horizontal: 'center',
        wrap: true,
        indent: null,
      },
      borders: {
        top: {
          style: 'solid',
          color: '#111111',
        },
        right: null,
      },
    })
  })

  it('preserves top-level style patch clears', () => {
    expect(
      normalizeStylePatch({
        fill: null,
        font: null,
        alignment: null,
        borders: null,
      }),
    ).toEqual({
      fill: null,
      font: null,
      alignment: null,
      borders: null,
    })
  })

  it('returns string number formats unchanged', () => {
    expect(normalizeNumberFormatInput('$#,##0.00')).toBe('$#,##0.00')
  })

  it('omits undefined number-format preset fields', () => {
    const format = {
      kind: 'currency',
      currency: 'USD',
      decimals: 2,
      useGrouping: undefined,
      negativeStyle: 'parentheses',
      zeroStyle: undefined,
    } satisfies ServerCellNumberFormatInput

    expect(normalizeNumberFormatInput(format)).toEqual({
      kind: 'currency',
      currency: 'USD',
      decimals: 2,
      negativeStyle: 'parentheses',
    })
  })
})
