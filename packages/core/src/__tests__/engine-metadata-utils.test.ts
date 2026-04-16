import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { FormulaNode } from '@bilig/formula'
import {
  definedNameValueToCellValue,
  definedNameValuesEqual,
  renameDefinedNameValueSheet,
  resolveMetadataReferencesInAst,
  spillDependencyKeyFromRef,
  tableDependencyKey,
} from '../engine-metadata-utils.js'
import { StringPool } from '../string-pool.js'

describe('engine metadata utils', () => {
  it('converts and compares defined-name snapshots predictably', () => {
    const strings = new StringPool()

    expect(definedNameValueToCellValue({ kind: 'scalar', value: 'hello' }, strings)).toEqual({
      tag: ValueTag.String,
      value: 'hello',
      stringId: strings.intern('hello'),
    })
    expect(
      definedNameValueToCellValue(
        {
          kind: 'cell-ref',
          sheetName: 'Sheet1',
          address: 'A1',
        },
        strings,
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(
      definedNameValuesEqual(
        { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      ),
    ).toBe(true)
    expect(
      definedNameValuesEqual(
        { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C2' },
      ),
    ).toBe(false)
  })

  it('renames defined-name sheet references without touching structured refs', () => {
    expect(renameDefinedNameValueSheet({ kind: 'cell-ref', sheetName: 'Old Sheet', address: 'A1' }, 'Old Sheet', 'New Sheet')).toEqual({
      kind: 'cell-ref',
      sheetName: 'New Sheet',
      address: 'A1',
    })
    expect(
      renameDefinedNameValueSheet({ kind: 'formula', formula: "='Old Sheet'!A1+SUM('Old Sheet'!B:B)" }, 'Old Sheet', 'New Sheet'),
    ).toEqual({
      kind: 'formula',
      formula: "='New Sheet'!A1+SUM('New Sheet'!B:B)",
    })
    expect(
      renameDefinedNameValueSheet({ kind: 'structured-ref', tableName: 'Sales', columnName: 'Amount' }, 'Old Sheet', 'New Sheet'),
    ).toEqual({
      kind: 'structured-ref',
      tableName: 'Sales',
      columnName: 'Amount',
    })
  })

  it('resolves names, structured refs, and spill refs across nested formulas', () => {
    const formula: FormulaNode = {
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'NameRef', name: 'Rate' },
      right: {
        kind: 'BinaryExpr',
        operator: '+',
        left: { kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' },
        right: { kind: 'SpillRef', sheetName: 'Sheet1', ref: 'A1' },
      },
    }

    expect(
      resolveMetadataReferencesInAst(formula, {
        resolveName: (name) => (name === 'Rate' ? { kind: 'scalar', value: 2 } : undefined),
        resolveStructuredReference: (tableName, columnName) =>
          tableName === 'Sales' && columnName === 'Amount'
            ? {
                kind: 'RangeRef',
                refKind: 'cells',
                sheetName: 'Data',
                start: 'B2',
                end: 'B4',
              }
            : undefined,
        resolveSpillReference: (sheetName, address) =>
          sheetName === 'Sheet1' && address === 'A1' ? { kind: 'CellRef', sheetName: 'Sheet1', ref: 'C3' } : undefined,
      }),
    ).toEqual({
      fullyResolved: true,
      substituted: true,
      node: {
        kind: 'BinaryExpr',
        operator: '+',
        left: { kind: 'NumberLiteral', value: 2 },
        right: {
          kind: 'BinaryExpr',
          operator: '+',
          left: {
            kind: 'RangeRef',
            refKind: 'cells',
            sheetName: 'Data',
            start: 'B2',
            end: 'B4',
          },
          right: { kind: 'CellRef', sheetName: 'Sheet1', ref: 'C3' },
        },
      },
    })
  })

  it('treats broken formula metadata and unresolved scalar names predictably', () => {
    expect(
      resolveMetadataReferencesInAst(
        { kind: 'NameRef', name: 'BrokenFormulaObject' },
        {
          resolveName: (name) => (name === 'BrokenFormulaObject' ? { kind: 'formula', formula: '=(' } : undefined),
          resolveStructuredReference: () => undefined,
          resolveSpillReference: () => undefined,
        },
      ),
    ).toEqual({
      fullyResolved: true,
      substituted: true,
      node: { kind: 'ErrorLiteral', code: ErrorCode.Value },
    })

    expect(
      resolveMetadataReferencesInAst(
        { kind: 'NameRef', name: 'BrokenFormulaString' },
        {
          resolveName: (name) => (name === 'BrokenFormulaString' ? '=(' : undefined),
          resolveStructuredReference: () => undefined,
          resolveSpillReference: () => undefined,
        },
      ),
    ).toEqual({
      fullyResolved: true,
      substituted: true,
      node: { kind: 'ErrorLiteral', code: ErrorCode.Value },
    })

    const unresolved = resolveMetadataReferencesInAst(
      { kind: 'NameRef', name: 'PendingNull' },
      {
        resolveName: (name) => (name === 'PendingNull' ? { kind: 'scalar', value: null } : undefined),
        resolveStructuredReference: () => undefined,
        resolveSpillReference: () => undefined,
      },
    )
    expect(unresolved.fullyResolved).toBe(false)
    expect(unresolved.substituted).toBe(false)
    expect(unresolved.node).toEqual({ kind: 'NameRef', name: 'PendingNull' })
  })

  it('propagates missing refs and cycles through unary, call, and invoke nodes', () => {
    const resolved = resolveMetadataReferencesInAst(
      {
        kind: 'InvokeExpr',
        callee: { kind: 'NameRef', name: 'Loop' },
        args: [
          {
            kind: 'UnaryExpr',
            operator: '-',
            argument: { kind: 'StructuredRef', tableName: 'MissingTable', columnName: 'Amount' },
          },
          {
            kind: 'CallExpr',
            callee: 'SUM',
            args: [{ kind: 'SpillRef', sheetName: 'Sheet1', ref: 'B2' }],
          },
        ],
      },
      {
        resolveName: (name) => (name === 'Loop' ? { kind: 'formula', formula: '=Loop' } : undefined),
        resolveStructuredReference: () => undefined,
        resolveSpillReference: () => undefined,
      },
    )

    expect(resolved).toEqual({
      fullyResolved: true,
      substituted: true,
      node: {
        kind: 'InvokeExpr',
        callee: { kind: 'ErrorLiteral', code: ErrorCode.Cycle },
        args: [
          {
            kind: 'UnaryExpr',
            operator: '-',
            argument: { kind: 'ErrorLiteral', code: ErrorCode.Ref },
          },
          {
            kind: 'CallExpr',
            callee: 'SUM',
            args: [{ kind: 'ErrorLiteral', code: ErrorCode.Ref }],
          },
        ],
      },
    })
  })

  it('normalizes quoted spill references with escaped apostrophes', () => {
    expect(spillDependencyKeyFromRef("'O''Brien''s Sheet'!A1", 'Sheet1')).toBe("O'Brien's Sheet!A1")
    expect(spillDependencyKeyFromRef('B2', 'Sheet1')).toBe('Sheet1!B2')
    expect(tableDependencyKey('sales')).toBe(tableDependencyKey('Sales'))
  })
})
