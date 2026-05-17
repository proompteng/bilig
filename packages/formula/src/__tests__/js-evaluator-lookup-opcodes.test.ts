import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { evaluateLookupMatchOpcode } from '../js-evaluator-lookup-opcodes.js'
import type { EvaluationContext, JsPlanInstruction, StackValue } from '../js-evaluator-types.js'

type ExactMatchInstruction = Extract<JsPlanInstruction, { opcode: 'lookup-exact-match' }>
type ApproximateMatchInstruction = Extract<JsPlanInstruction, { opcode: 'lookup-approximate-match' }>

const exactInstruction: ExactMatchInstruction = {
  opcode: 'lookup-exact-match',
  callee: 'MATCH',
  start: 'A1',
  end: 'A3',
  startRow: 0,
  endRow: 2,
  startCol: 0,
  endCol: 0,
  refKind: 'cells',
  searchMode: 1,
}

const approximateInstruction: ApproximateMatchInstruction = {
  opcode: 'lookup-approximate-match',
  callee: 'XMATCH',
  start: 'A1',
  end: 'A3',
  startRow: 0,
  endRow: 2,
  startCol: 0,
  endCol: 0,
  refKind: 'cells',
  matchMode: 1,
}

function baseContext(overrides: Partial<EvaluationContext> = {}): EvaluationContext {
  return {
    sheetName: 'Sheet1',
    resolveCell: () => empty(),
    resolveRange: () => [num(1), num(2), num(3)],
    ...overrides,
  }
}

describe('js evaluator lookup opcodes', () => {
  it('uses direct exact lookup handlers without materializing ranges', () => {
    const resolveExactVectorMatch = vi.fn(() => ({ handled: true as const, position: 2 }))
    const resolveRange = vi.fn(() => {
      throw new Error('range should stay unmaterialized')
    })
    const noteExactLookupDirect = vi.fn()

    expect(
      evaluateLookupMatchOpcode({
        instruction: exactInstruction,
        lookupOperand: scalar(num(2)),
        context: baseContext({ resolveExactVectorMatch, resolveRange, noteExactLookupDirect }),
      }),
    ).toEqual(scalar(num(2)))
    expect(resolveExactVectorMatch).toHaveBeenCalledWith({
      lookupValue: num(2),
      sheetName: 'Sheet1',
      start: 'A1',
      end: 'A3',
      startRow: 0,
      endRow: 2,
      startCol: 0,
      endCol: 0,
      searchMode: 1,
    })
    expect(noteExactLookupDirect).toHaveBeenCalledTimes(1)
    expect(resolveRange).not.toHaveBeenCalled()
  })

  it('materializes wildcard MATCH through the builtin fallback instead of trusting exact handlers', () => {
    const resolveExactVectorMatch = vi.fn(() => {
      throw new Error('wildcard MATCH must not use exact-vector shortcut')
    })
    const resolveRange = vi.fn(() => [text('alpha'), text('bravo'), text('beta')])
    const noteExactLookupFallback = vi.fn()
    const noteRangeMaterialization = vi.fn()

    expect(
      evaluateLookupMatchOpcode({
        instruction: exactInstruction,
        lookupOperand: scalar(text('b*')),
        context: baseContext({ resolveExactVectorMatch, resolveRange, noteExactLookupFallback, noteRangeMaterialization }),
      }),
    ).toEqual(scalar(num(2)))
    expect(resolveExactVectorMatch).not.toHaveBeenCalled()
    expect(noteExactLookupFallback).toHaveBeenCalledTimes(1)
    expect(noteRangeMaterialization).toHaveBeenCalledWith(3)
    expect(resolveRange).toHaveBeenCalledTimes(1)
  })

  it('uses direct approximate handlers and preserves #N/A misses', () => {
    const resolveApproximateVectorMatch = vi.fn(() => ({ handled: true as const, position: undefined }))
    const resolveRange = vi.fn(() => {
      throw new Error('range should stay unmaterialized')
    })

    expect(
      evaluateLookupMatchOpcode({
        instruction: approximateInstruction,
        lookupOperand: scalar(num(4)),
        context: baseContext({ resolveApproximateVectorMatch, resolveRange }),
      }),
    ).toEqual(scalar(err(ErrorCode.NA)))
    expect(resolveApproximateVectorMatch).toHaveBeenCalledWith({
      lookupValue: num(4),
      sheetName: 'Sheet1',
      start: 'A1',
      end: 'A3',
      startRow: 0,
      endRow: 2,
      startCol: 0,
      endCol: 0,
      matchMode: 1,
    })
    expect(resolveRange).not.toHaveBeenCalled()
  })
})

function scalar(value: CellValue): StackValue {
  return { kind: 'scalar', value }
}

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function empty(): CellValue {
  return { tag: ValueTag.Empty }
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}
