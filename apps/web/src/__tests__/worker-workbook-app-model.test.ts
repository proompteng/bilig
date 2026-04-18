import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  emptyCellSnapshot,
  parseEditorInput,
  parseSelectionTarget,
  parsedEditorInputFromSnapshot,
  parsedEditorInputMatchesSnapshot,
  sameCellContent,
  toEditorValue,
} from '../worker-workbook-app-model.js'

describe('worker workbook app model', () => {
  it('normalizes snapshots into parsed editor input shapes', () => {
    const formulaCell = {
      ...emptyCellSnapshot('Sheet1', 'A1'),
      formula: 'SUM(B1:B3)',
      value: { tag: ValueTag.Number, value: 42 },
    }
    const booleanCell = {
      ...emptyCellSnapshot('Sheet1', 'A2'),
      input: true,
      value: { tag: ValueTag.Boolean, value: true },
    }
    const errorCell = {
      ...emptyCellSnapshot('Sheet1', 'A3'),
      value: { tag: ValueTag.Error, code: ErrorCode.Div0 },
    }

    expect(parsedEditorInputFromSnapshot(formulaCell)).toEqual({
      kind: 'formula',
      formula: 'SUM(B1:B3)',
    })
    expect(parsedEditorInputFromSnapshot(booleanCell)).toEqual({
      kind: 'value',
      value: true,
    })
    expect(parsedEditorInputFromSnapshot(errorCell)).toEqual({
      kind: 'value',
      value: '#DIV/0!',
    })
  })

  it('matches parsed editor input against authoritative snapshots', () => {
    const numericCell = {
      ...emptyCellSnapshot('Sheet1', 'B4'),
      input: 17,
      value: { tag: ValueTag.Number, value: 17 },
    }
    const formulaCell = {
      ...emptyCellSnapshot('Sheet1', 'B5'),
      formula: 'A1+A2',
      value: { tag: ValueTag.Number, value: 9 },
    }

    expect(parsedEditorInputMatchesSnapshot({ kind: 'value', value: 17 }, numericCell)).toBe(true)
    expect(parsedEditorInputMatchesSnapshot({ kind: 'formula', formula: 'A1+A2' }, formulaCell)).toBe(true)
    expect(parsedEditorInputMatchesSnapshot({ kind: 'clear' }, formulaCell)).toBe(false)
  })

  it('treats style or version-only drift as the same cell content', () => {
    const baseCell = {
      ...emptyCellSnapshot('Sheet1', 'C7'),
      input: 'remote',
      value: { tag: ValueTag.String, value: 'remote' },
      version: 1,
    }
    const styleOnlyUpdate = {
      ...baseCell,
      styleId: 'style-2',
      version: 2,
    }
    const contentChange = {
      ...baseCell,
      input: 'local',
      value: { tag: ValueTag.String, value: 'local' },
      version: 3,
    }

    expect(sameCellContent(baseCell, styleOnlyUpdate)).toBe(true)
    expect(sameCellContent(baseCell, contentChange)).toBe(false)
  })

  it('resolves cell-ref defined names in the name box', () => {
    expect(
      parseSelectionTarget('TaxRate', 'Sheet1', [
        {
          name: 'TaxRate',
          value: {
            kind: 'cell-ref',
            sheetName: 'Sheet2',
            address: 'C4',
          },
        },
      ]),
    ).toEqual({
      kind: 'cell',
      sheetName: 'Sheet2',
      address: 'C4',
      range: {
        startAddress: 'C4',
        endAddress: 'C4',
      },
    })
  })

  it('preserves whitespace in literal inputs while still parsing formulas and compact scalars', () => {
    expect(parseEditorInput('  SKU  ')).toEqual({
      kind: 'value',
      value: '  SKU  ',
    })
    expect(parseEditorInput('=SUM(A1:A3)')).toEqual({
      kind: 'formula',
      formula: 'SUM(A1:A3)',
    })
    expect(parseEditorInput('  =SUM(A1:A3)')).toEqual({
      kind: 'value',
      value: '  =SUM(A1:A3)',
    })
    expect(parseEditorInput('42')).toEqual({
      kind: 'value',
      value: 42,
    })
    expect(parseEditorInput(' 42 ')).toEqual({
      kind: 'value',
      value: ' 42 ',
    })
    expect(parseEditorInput('TRUE')).toEqual({
      kind: 'value',
      value: true,
    })
    expect(parseEditorInput(' TRUE ')).toEqual({
      kind: 'value',
      value: ' TRUE ',
    })
    expect(parseEditorInput('')).toEqual({
      kind: 'clear',
    })
    expect(parseEditorInput('   ')).toEqual({
      kind: 'value',
      value: '   ',
    })
  })

  it('keeps formula source text visible even when the evaluated value is an error', () => {
    expect(
      toEditorValue({
        ...emptyCellSnapshot('Sheet1', 'F7'),
        formula: 'SUM(A1:A3)',
        value: { tag: ValueTag.Error, code: ErrorCode.Value },
      }),
    ).toBe('=SUM(A1:A3)')
  })

  it('resolves cell ranges, row ranges, column ranges, quoted sheets, and range-ref defined names', () => {
    expect(parseSelectionTarget('B2:D8', 'Sheet1')).toEqual({
      kind: 'range',
      sheetName: 'Sheet1',
      address: 'B2',
      range: {
        startAddress: 'B2',
        endAddress: 'D8',
      },
    })

    expect(parseSelectionTarget('B:D', 'Sheet1')).toEqual({
      kind: 'column',
      sheetName: 'Sheet1',
      address: 'B1',
      range: {
        startAddress: 'B1',
        endAddress: 'D1048576',
      },
    })

    expect(parseSelectionTarget('2:5', 'Sheet1')).toEqual({
      kind: 'row',
      sheetName: 'Sheet1',
      address: 'A2',
      range: {
        startAddress: 'A2',
        endAddress: 'XFD5',
      },
    })

    expect(parseSelectionTarget("'Ops Team'!C3:E7", 'Sheet1')).toEqual({
      kind: 'range',
      sheetName: 'Ops Team',
      address: 'C3',
      range: {
        startAddress: 'C3',
        endAddress: 'E7',
      },
    })

    expect(
      parseSelectionTarget('DataBlock', 'Sheet1', [
        {
          name: 'DataBlock',
          value: {
            kind: 'range-ref',
            sheetName: 'Sheet2',
            startAddress: 'B4',
            endAddress: 'D9',
          },
        },
      ]),
    ).toEqual({
      kind: 'range',
      sheetName: 'Sheet2',
      address: 'B4',
      range: {
        startAddress: 'B4',
        endAddress: 'D9',
      },
    })
  })
})
