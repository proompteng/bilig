import { describe, expect, it } from 'vitest'
import {
  buildTranslatedCellReferenceMap,
  formatParsedCellReference,
  formatParsedRangeReference,
  translateCellReference,
  translateColumnReference,
  translateParsedRangeReference,
  translateQualifiedDependencyReference,
  translateQualifiedRangeReference,
  translateRangeAddress,
  translateRowReference,
} from '../formula-reference-translation.js'
import { parseRangeAddress } from '../addressing.js'

describe('formula reference translation', () => {
  it('translates local, qualified, and parsed references', () => {
    expect(translateCellReference('$A1', 2, 4)).toBe('$A3')
    expect(translateColumnReference('C', 2)).toBe('E')
    expect(translateRowReference('$4', 3)).toBe('$4')
    expect(translateQualifiedRangeReference("'Sheet 2'!A1:B2", 1, 1)).toBe("'Sheet 2'!B2:C3")
    expect(translateQualifiedRangeReference('Jan:Mar!B2:C2', 1, 1)).toBe('Jan:Mar!C3:D3')
    expect(translateQualifiedRangeReference("'Jan 2026':'Mar 2026'!$B2:C$2", 1, 1)).toBe("'Jan 2026':'Mar 2026'!$B3:D$2")
    expect(translateQualifiedDependencyReference('Jan:Mar!2:4', 2, 0)).toBe('Jan:Mar!4:6')
    expect(translateQualifiedDependencyReference('Jan:Mar!B:D', 0, 2)).toBe('Jan:Mar!D:F')
    expect(formatParsedCellReference({ address: 'A1', sheetName: 'Sheet 2', explicitSheet: true })).toBe("'Sheet 2'!A1")

    const translatedRows = translateParsedRangeReference(
      {
        kind: 'range',
        address: '2:4',
        refKind: 'rows',
        startAddress: '2',
        endAddress: '4',
        startRow: 1,
        endRow: 3,
        startCol: 0,
        endCol: 0,
      },
      2,
      0,
    )
    expect(formatParsedRangeReference(translatedRows)).toBe('4:6')
  })

  it('builds instruction maps from original parsed references', () => {
    const source = { address: 'A1', row: 0, col: 0, rowAbsolute: false, colAbsolute: false }
    const target = { address: 'B2', row: 1, col: 1, rowAbsolute: false, colAbsolute: false }

    expect(buildTranslatedCellReferenceMap([source], [target]).get('\tA1')).toBe(target)
    expect(translateRangeAddress(parseRangeAddress('A1:B2'), 1, 1).start.text).toBe('B2')
  })
})
