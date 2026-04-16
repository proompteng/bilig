import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  absoluteAddress,
  cellTypeCode,
  currentCellReference,
  referenceColumnNumber,
  referenceRowNumber,
  sheetIndexByName,
  sheetNames,
} from '../js-evaluator-reference-context.js'
import type { EvaluationContext, ReferenceOperand } from '../js-evaluator.js'

const context: EvaluationContext = {
  sheetName: 'Sheet2',
  currentAddress: 'C4',
  resolveCell: () => ({ tag: ValueTag.Empty }),
  resolveRange: () => [],
  listSheetNames: () => ['Sheet1', 'Sheet2', 'Summary'],
}

describe('js evaluator reference context', () => {
  it('resolves row, column, and absolute addresses from refs and current context', () => {
    expect(currentCellReference(context)).toEqual({
      kind: 'cell',
      sheetName: 'Sheet2',
      address: 'C4',
    })
    expect(referenceRowNumber(undefined, context)).toBe(4)
    expect(referenceColumnNumber(undefined, context)).toBe(3)
    expect(absoluteAddress(undefined, context)).toBe('$C$4')

    const cellRef: ReferenceOperand = { kind: 'cell', address: 'B7' }
    const rangeRef: ReferenceOperand = { kind: 'range', start: 'D3', end: 'F5', refKind: 'cells' }
    const rowRef: ReferenceOperand = { kind: 'row', address: '9' }
    const columnRef: ReferenceOperand = { kind: 'col', address: 'AA' }

    expect(referenceRowNumber(cellRef, context)).toBe(7)
    expect(referenceColumnNumber(cellRef, context)).toBe(2)
    expect(absoluteAddress(cellRef, context)).toBe('$B$7')

    expect(referenceRowNumber(rangeRef, context)).toBe(3)
    expect(referenceColumnNumber(rangeRef, context)).toBe(4)
    expect(absoluteAddress(rangeRef, context)).toBe('$D$3')

    expect(referenceRowNumber(rowRef, context)).toBe(9)
    expect(referenceColumnNumber(rowRef, context)).toBeUndefined()

    expect(referenceRowNumber(columnRef, context)).toBeUndefined()
    expect(referenceColumnNumber(columnRef, context)).toBe(27)
  })

  it('resolves sheet indexes case-insensitively and falls back to current sheet', () => {
    expect(sheetNames(context)).toEqual(['Sheet1', 'Sheet2', 'Summary'])
    expect(sheetIndexByName('sheet2', context)).toBe(2)
    expect(sheetIndexByName(' SUMMARY ', context)).toBe(3)
    expect(sheetIndexByName('Missing', context)).toBeUndefined()
  })

  it('maps cell types to Excel-compatible CELL(type) markers', () => {
    expect(cellTypeCode({ tag: ValueTag.Empty })).toBe('b')
    expect(cellTypeCode({ tag: ValueTag.String, value: 'x', stringId: 0 })).toBe('l')
    expect(cellTypeCode({ tag: ValueTag.Number, value: 1 })).toBe('v')
    expect(cellTypeCode({ tag: ValueTag.Boolean, value: true })).toBe('v')
    expect(cellTypeCode({ tag: ValueTag.Error, code: ErrorCode.Ref })).toBe('v')
  })
})
