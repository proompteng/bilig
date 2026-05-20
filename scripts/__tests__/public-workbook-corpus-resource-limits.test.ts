import { describe, expect, it } from 'vitest'

import {
  formulaOracleResourceLimitPreflight,
  importResourceLimitPreflight,
  inspectFormulaOracleDependencyFootprint,
} from '../public-workbook-corpus-resource-limits.ts'
import type { WorkbookSnapshot } from '../../packages/protocol/src/types.ts'
import type { PublicWorkbookArtifact } from '../public-workbook-corpus-types.ts'
import { emptyFeatureCounts, type WorkbookFootprint } from '../public-workbook-corpus-workbook.ts'

describe('public workbook corpus resource limit preflights', () => {
  it('skips formula oracle when sparse formulas reference broad dependency ranges', () => {
    const snapshot = workbookWithFormula(
      'SUMPRODUCT((Data!$A$2:$A$99311=Working!$A$3)*(Data!$B$2:$B$99311=Working!$A5)*(Data!$D$2:$D$99311))',
    )

    const footprint = inspectFormulaOracleDependencyFootprint(snapshot)
    const preflight = formulaOracleResourceLimitPreflight(snapshot)

    expect(footprint).toMatchObject({
      formulaCellCount: 1,
      dependencyReferenceCount: 5,
      maxDependencyCellReferences: 99310,
      maxDependencyReference: 'Data!A2:A99311',
      unparseableDependencyReferenceCount: 0,
    })
    expect(footprint.totalDependencyCellReferences).toBe(297_932)
    expect(preflight).toMatchObject({
      classification: 'xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>2000000dependencyCells',
      evidence: expect.arrayContaining([
        'rss-limit-phase=formula-oracle',
        'formula-oracle-dependency-footprint=297932',
        'formula-oracle-largest-dependency=Data!A2:A99311:99310',
      ]),
    })
  })

  it('keeps formula oracle enabled for compact formulas', () => {
    const snapshot = workbookWithFormula('SUM(Data!A1:A10)')

    expect(inspectFormulaOracleDependencyFootprint(snapshot)).toMatchObject({
      formulaCellCount: 1,
      dependencyReferenceCount: 1,
      totalDependencyCellReferences: 10,
      maxDependencyCellReferences: 10,
      maxDependencyReference: 'Data!A1:A10',
    })
    expect(formulaOracleResourceLimitPreflight(snapshot)).toBeNull()
  })

  it('skips formula oracle before dependency parsing when formula count is too high', () => {
    const preflight = formulaOracleResourceLimitPreflight(workbookWithManyFormulas(2_001))

    expect(preflight).toMatchObject({
      classification: 'xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>2000formulas',
      evidence: expect.arrayContaining(['rss-limit-phase=formula-oracle', 'formula-oracle-formula-count=2001']),
    })
  })

  it('allows large simple value-only XLSX imports through the product fast path budget', () => {
    expect(importResourceLimitPreflight(workbookArtifact(), workbookFootprint({ cellCount: 406_431, valueCellCount: 406_431 }))).toBeNull()
  })

  it('allows large value-only XLSX imports with workbook defined names through the product fast path budget', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint(
          { cellCount: 262_869, valueCellCount: 262_869, definedNameCount: 2 },
          { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
        ),
      ),
    ).toBeNull()
  })

  it('keeps the stricter import preflight when the large simple product importer is blocked', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint(
          { cellCount: 262_869, valueCellCount: 262_869, definedNameCount: 2 },
          { largeSimpleXlsxImport: { eligible: false, blockers: ['unsupported-package-parts=2'] } },
        ),
      )?.evidence,
    ).toEqual(expect.arrayContaining(['Public corpus verification import preflight limit exceeded: cell-count 262869 > 200000']))
  })

  it('keeps the stricter import preflight for large formula workbooks', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint({ cellCount: 406_431, valueCellCount: 406_431, formulaCellCount: 1 }),
      )?.evidence,
    ).toEqual(expect.arrayContaining(['Public corpus verification import preflight limit exceeded: cell-count 406431 > 200000']))
  })

  it('allows large XLSX imports with a small simple-formula set through the product fast path budget', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint(
          { cellCount: 361_614, valueCellCount: 361_614, formulaCellCount: 92 },
          { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
        ),
      ),
    ).toBeNull()
  })

  it('allows large XLSX imports with many supported formulas through the product fast path budget', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint(
          { cellCount: 342_986, valueCellCount: 342_986, formulaCellCount: 46_205 },
          { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
        ),
      ),
    ).toBeNull()
  })

  it('allows large XLSX imports with tables through the product fast path budget when compatibility is proven', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint(
          { cellCount: 228_782, valueCellCount: 228_782, tableCount: 1 },
          { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
        ),
      ),
    ).toBeNull()
  })

  it('allows large XLSX imports with conditional formatting through the product fast path budget when compatibility is proven', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint(
          { cellCount: 246_086, valueCellCount: 246_086, conditionalFormatCount: 20 },
          { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
        ),
      ),
    ).toBeNull()
  })

  it('allows large XLSX imports with supported data validations through the product fast path budget when compatibility is proven', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact(),
        workbookFootprint(
          { cellCount: 246_086, valueCellCount: 246_086, dataValidationCount: 20 },
          { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
        ),
      ),
    ).toBeNull()
  })

  it('does not apply the SheetJS sheet/package import budget to proven large simple XLSX imports', () => {
    expect(
      importResourceLimitPreflight(
        workbookArtifact({ byteSize: 2_482_697 }),
        workbookFootprint(
          { cellCount: 246_084, valueCellCount: 246_084, sheetCount: 55, conditionalFormatCount: 20 },
          { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
        ),
      ),
    ).toBeNull()
  })

  it('still rejects simple workbooks above the large fast path budget', () => {
    expect(
      importResourceLimitPreflight(workbookArtifact(), workbookFootprint({ cellCount: 800_000, valueCellCount: 800_000 }))?.evidence,
    ).toEqual(expect.arrayContaining(['Public corpus verification import preflight limit exceeded: cell-count 800000 > 750000']))
  })
})

function workbookWithFormula(formula: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'resource-limit-formulas' },
    sheets: [
      {
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 1 },
          { address: 'B1', value: 2 },
          { address: 'D1', value: 3 },
        ],
      },
      {
        name: 'Working',
        order: 1,
        cells: [
          { address: 'A3', value: 'England' },
          { address: 'A5', value: 'total' },
          { address: 'B5', formula, value: 42 },
        ],
      },
    ],
  }
}

function workbookWithManyFormulas(count: number): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'many-formulas' },
    sheets: [
      {
        name: 'Data',
        order: 0,
        cells: Array.from({ length: count }, (_value, index) => ({
          address: `A${String(index + 1)}`,
          formula: '1+1',
          value: 2,
        })),
      },
    ],
  }
}

function workbookArtifact(input: { readonly byteSize?: number } = {}): PublicWorkbookArtifact {
  return {
    id: 'artifact',
    sourceId: 'source',
    sourceUrl: 'https://example.com/workbook.xlsx',
    downloadUrl: 'https://example.com/workbook.xlsx',
    fileName: 'workbook.xlsx',
    sha256: '0'.repeat(64),
    byteSize: input.byteSize ?? 1024 * 1024,
    cachePath: 'workbook.xlsx',
    workbookFingerprint: '1'.repeat(64),
    fetchedAt: '2026-05-17T00:00:00.000Z',
    license: {
      title: 'Test',
      evidenceUrl: 'https://example.com/license',
      spdxId: 'CC0-1.0',
    },
  }
}

function workbookFootprint(
  counts: Partial<WorkbookFootprint['featureCounts']>,
  options: Pick<WorkbookFootprint, 'largeSimpleXlsxImport'> = {},
): WorkbookFootprint {
  const featureCounts = {
    ...emptyFeatureCounts(),
    sheetCount: 1,
    ...counts,
  }
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: 'workbook',
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: featureCounts.cellCount,
          columnCount: 1,
          nonEmptyCellCount: featureCounts.cellCount,
          usedRange:
            featureCounts.cellCount > 0 ? { startRow: 0, startColumn: 0, endRow: featureCounts.cellCount - 1, endColumn: 0 } : null,
        },
      ],
    },
    externalWorkbookReferences: [],
    ...options,
  }
}
