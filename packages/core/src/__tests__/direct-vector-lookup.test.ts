import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@bilig/protocol'
import { resolveRuntimeDirectLookupBinding } from '../engine/direct-vector-lookup.js'

describe('resolveRuntimeDirectLookupBinding', () => {
  it('returns exact bindings for supported single-column exact lookup plans', () => {
    expect(
      resolveRuntimeDirectLookupBinding(
        [
          { opcode: 'push-cell', address: '$D$1' },
          {
            opcode: 'lookup-exact-match',
            start: 'A2',
            end: 'A6',
            startRow: 1,
            endRow: 5,
            startCol: 0,
            endCol: 0,
            searchMode: -1,
          },
          { opcode: 'return' },
        ],
        'Sheet1',
      ),
    ).toEqual({
      kind: 'exact',
      operandSheetName: 'Sheet1',
      operandAddress: 'D1',
      lookupSheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 5,
      col: 0,
      searchMode: -1,
    })
  })

  it('returns approximate bindings for supported single-column approximate lookup plans', () => {
    expect(
      resolveRuntimeDirectLookupBinding(
        [
          { opcode: 'push-cell', address: 'D1', sheetName: 'Inputs' },
          {
            opcode: 'lookup-approximate-match',
            start: 'C2',
            end: 'C6',
            startRow: 1,
            endRow: 5,
            startCol: 2,
            endCol: 2,
            matchMode: 1,
            sheetName: 'Lookup',
          },
          { opcode: 'return' },
        ],
        'Sheet1',
      ),
    ).toEqual({
      kind: 'approximate',
      operandSheetName: 'Inputs',
      operandAddress: 'D1',
      lookupSheetName: 'Lookup',
      rowStart: 1,
      rowEnd: 5,
      col: 2,
      matchMode: 1,
    })
  })

  it('rejects unsupported operand and lookup instruction shapes', () => {
    const rejectedPlans: Array<readonly unknown[]> = [
      [],
      [
        null,
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-unknown' },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [{ opcode: 'push-number', value: 5 }, { opcode: 'return' }],
      [{ opcode: 'push-cell', address: 'D1' }, null, { opcode: 'return' }],
      [
        { opcode: 'push-cell', address: 'D1' },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: '0',
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-number', value: 5 },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-boolean', value: true },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-string', value: 'x' },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-error', code: ErrorCode.Div0 },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-name', name: 'criterion' },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-cell', address: 'D1' },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'B3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 1,
          searchMode: 1,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-cell', address: 'D1' },
        {
          opcode: 'lookup-exact-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          searchMode: 2,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-cell', address: 'D1' },
        {
          opcode: 'lookup-approximate-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          matchMode: 0,
        },
        { opcode: 'return' },
      ],
      [
        { opcode: 'push-cell', address: 'D1' },
        {
          opcode: 'lookup-approximate-match',
          start: 'A1',
          end: 'A3',
          startRow: 0,
          endRow: 2,
          startCol: 0,
          endCol: 0,
          matchMode: 1,
        },
        { opcode: 'noop' },
      ],
    ]

    for (const plan of rejectedPlans) {
      expect(resolveRuntimeDirectLookupBinding(plan, 'Sheet1')).toBeUndefined()
    }
  })
})
