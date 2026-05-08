import { describe, expect, it } from 'vitest'

import { importCsv } from '../index.js'

function cellValuesByAddress(csv: string): Map<string, unknown> {
  const imported = importCsv(csv, 'accounting-identifiers.csv')
  const cells = imported.snapshot.sheets[0]?.cells ?? []
  return new Map(cells.map((cell) => [cell.address, 'formula' in cell ? cell.formula : cell.value]))
}

describe('GitHub issue #28 CSV leading-zero identifiers', () => {
  it('preserves leading-zero accounting identifiers as text', () => {
    const values = cellValuesByAddress(
      [
        'Account,CostCenter,Invoice,RoutingNumber,Amount',
        '0010,00042,INV-0007,021000021,125.50',
        '"0100","007","000123","011000015",50',
      ].join('\n'),
    )

    expect(values.get('A2')).toBe('0010')
    expect(values.get('B2')).toBe('00042')
    expect(values.get('C2')).toBe('INV-0007')
    expect(values.get('D2')).toBe('021000021')
    expect(values.get('E2')).toBe(125.5)
    expect(values.get('A3')).toBe('0100')
    expect(values.get('B3')).toBe('007')
    expect(values.get('C3')).toBe('000123')
    expect(values.get('D3')).toBe('011000015')
    expect(values.get('E3')).toBe(50)
  })
})
