import {
  addressToCell,
  aggregateOptionsOracleAddresses,
  blockedSpillReferenceOracleAddresses,
  buildAggregateOptionsOracleWorkbook,
  buildDynamicSpillOracleWorkbook,
  buildFutureFunctionOracleWorkbook,
  buildIndexImplicitIntersectionOracleWorkbook,
  buildOffsetImplicitIntersectionOracleWorkbook,
  buildOracleWorkbook,
  buildShrinkingSpillReferenceOracleWorkbook,
  buildSingleImplicitIntersectionOracleWorkbook,
  buildSpillReferenceOracleWorkbook,
  buildStructuralMoveColumnOracleWorkbook,
  buildTextsplitErrorOracleWorkbook,
  describe,
  dynamicSpillOracleAddresses,
  expect,
  expectedAggregateOptionsOracleCells,
  expectedBlockedSpillReferenceOracleValues,
  expectedDynamicSpillOracleValues,
  expectedFutureFunctionOracleCells,
  expectedIndexImplicitIntersectionOracleCells,
  expectedOffsetImplicitIntersectionOracleCells,
  expectedOracleCells,
  expectedShrinkingSpillReferenceOracleCells,
  expectedSingleImplicitIntersectionOracleValues,
  expectedSpillReferenceOracleCells,
  expectedTextsplitErrorOracleCells,
  expectedUnblockedSpillReferenceOracleValues,
  exportXlsx,
  futureFunctionOracleAddresses,
  importXlsx,
  indexImplicitIntersectionConfig,
  indexImplicitIntersectionOracleAddresses,
  it,
  normalizedCellValue,
  offsetImplicitIntersectionOracleAddresses,
  oracleFormulaAddresses,
  readRuntimeImage,
  shrinkingSpillReferenceOracleAddresses,
  singleImplicitIntersectionOracleAddresses,
  spillReferenceOracleAddresses,
  structuralMoveColumnFormulaOracleCell,
  textsplitErrorOracleAddresses,
  ValueTag,
  workbookConfig,
  WorkPaper,
} from './macos-desktop-excel-xlsx-oracle-test-helpers.js'

describe('macOS Desktop Excel XLSX oracle formula and spill compatibility', () => {
  it('exports and reimports the oracle fixture through the headless workbook path', () => {
    const workbook = buildOracleWorkbook()
    try {
      expect(oracleFormulaAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(oracleFormulaAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address))))).toEqual(
          expectedOracleCells.map((expected) => expected.value),
        )
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports and reimports AGGREGATE option semantics through the headless workbook path', () => {
    const workbook = buildAggregateOptionsOracleWorkbook()
    try {
      expect(aggregateOptionsOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedAggregateOptionsOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-aggregate-options-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          aggregateOptionsOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedAggregateOptionsOracleCells.map((expected) => expected.value))
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports and reimports standalone INDEX implicit-intersection formulas without spill metadata', () => {
    const workbook = buildIndexImplicitIntersectionOracleWorkbook()
    try {
      expect(
        indexImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedIndexImplicitIntersectionOracleCells.map((expected) => expected.value))
      expect(workbook.engine.getSpillRanges()).toEqual([])

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-index-implicit-intersection-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, indexImplicitIntersectionConfig)
      try {
        expect(
          indexImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedIndexImplicitIntersectionOracleCells.map((expected) => expected.value))
        expect(reimported.engine.getSpillRanges()).toEqual([])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports and reimports standalone OFFSET implicit-intersection formulas without spill metadata', () => {
    const workbook = buildOffsetImplicitIntersectionOracleWorkbook()
    try {
      expect(
        offsetImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedOffsetImplicitIntersectionOracleCells.map((expected) => expected.value))
      expect(workbook.engine.getSpillRanges()).toEqual([])

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-offset-implicit-intersection-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, indexImplicitIntersectionConfig)
      try {
        expect(
          offsetImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedOffsetImplicitIntersectionOracleCells.map((expected) => expected.value))
        expect(reimported.engine.getSpillRanges()).toEqual([])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports Excel future functions in a Desktop Excel-compatible XLSX shape', () => {
    const workbook = buildFutureFunctionOracleWorkbook()
    try {
      expect(futureFunctionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedFutureFunctionOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-future-function-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          futureFunctionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedFutureFunctionOracleCells.map((expected) => expected.value))
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports SINGLE implicit-intersection formulas through Desktop Excel-compatible XLSX', () => {
    const workbook = buildSingleImplicitIntersectionOracleWorkbook()
    try {
      expect(
        singleImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedSingleImplicitIntersectionOracleValues.map((expected) => expected.value))

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-single-implicit-intersection-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          singleImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedSingleImplicitIntersectionOracleValues.map((expected) => expected.value))
        expect(reimported.getCellFormula(addressToCell('C1'))).toBe('=SINGLE(A1:A3)')
        expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(SINGLE(A1:A3))')
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports native dynamic-array spill caches through the headless XLSX path', () => {
    const workbook = buildDynamicSpillOracleWorkbook()
    try {
      expect(dynamicSpillOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedDynamicSpillOracleValues.map((expected) => expected.value),
      )

      const snapshot = workbook.exportSnapshot()
      expect(
        readRuntimeImage(snapshot)
          ?.cellValues?.filter((cellValue) => cellValue.sheetName === 'Cases' && cellValue.col === 1)
          .map(({ row, value }) => ({ row, value })),
      ).toEqual([
        { row: 0, value: { tag: ValueTag.Number, value: 2 } },
        { row: 1, value: { tag: ValueTag.Number, value: 4 } },
        { row: 2, value: { tag: ValueTag.Number, value: 6 } },
      ])

      const imported = importXlsx(exportXlsx(snapshot), 'headless-dynamic-spill-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(dynamicSpillOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address))))).toEqual(
          expectedDynamicSpillOracleValues.map((expected) => expected.value),
        )
        expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports spill-reference consumer formulas through Desktop Excel-compatible XLSX', () => {
    const workbook = buildSpillReferenceOracleWorkbook()
    try {
      expect(spillReferenceOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedSpillReferenceOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-spill-reference-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          spillReferenceOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedSpillReferenceOracleCells.map((expected) => expected.value))
        expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(B1#)')
        expect(reimported.getCellFormula(addressToCell('E1'))).toBe('=ROWS(B1#)')
        expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=INDEX(B1#,2)')
        expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('keeps spill-reference consumers valid when a dynamic-array owner shrinks to one cell', () => {
    const workbook = buildShrinkingSpillReferenceOracleWorkbook()
    try {
      workbook.setCellContents(addressToCell('A1'), 1)

      expect(
        shrinkingSpillReferenceOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedShrinkingSpillReferenceOracleCells.map((expected) => expected.value))
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 1 }])

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-shrinking-spill-reference-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          shrinkingSpillReferenceOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedShrinkingSpillReferenceOracleCells.map((expected) => expected.value))
        expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(B1#)')
        expect(reimported.getCellFormula(addressToCell('E1'))).toBe('=ROWS(B1#)')
        expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=IFERROR(INDEX(B1#,2),"missing")')
        expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 1 }])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('blocks spill-reference consumers when a spill child is authored and unblocks after clear', () => {
    const workbook = buildShrinkingSpillReferenceOracleWorkbook()
    try {
      workbook.setCellContents(addressToCell('B2'), 99)

      expect(
        blockedSpillReferenceOracleAddresses.map((address) => ({
          address,
          value: normalizedCellValue(workbook.getCellValue(addressToCell(address))),
        })),
      ).toEqual(expectedBlockedSpillReferenceOracleValues)
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 1 }])

      const blockedImported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-blocked-spill-reference-oracle.xlsx')
      const blockedReimported = WorkPaper.buildFromSnapshot(blockedImported.snapshot, workbookConfig)
      try {
        expect(
          blockedSpillReferenceOracleAddresses.map((address) => ({
            address,
            value: normalizedCellValue(blockedReimported.getCellValue(addressToCell(address))),
          })),
        ).toEqual(expectedBlockedSpillReferenceOracleValues)
        expect(blockedImported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 1 }])
      } finally {
        blockedReimported.dispose()
      }

      workbook.setCellContents(addressToCell('B2'), null)

      expect(
        blockedSpillReferenceOracleAddresses.map((address) => ({
          address,
          value: normalizedCellValue(workbook.getCellValue(addressToCell(address))),
        })),
      ).toEqual(expectedUnblockedSpillReferenceOracleValues)
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
    } finally {
      workbook.dispose()
    }
  })

  it('exports and reimports TEXTSPLIT error spill children through the headless XLSX path', () => {
    const workbook = buildTextsplitErrorOracleWorkbook()
    try {
      expect(textsplitErrorOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedTextsplitErrorOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-textsplit-error-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          textsplitErrorOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedTextsplitErrorOracleCells.map((expected) => expected.value))
        expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'C1', rows: 2, cols: 2 }])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('rewrites moved-out column ranges like Desktop Excel', () => {
    const workbook = buildStructuralMoveColumnOracleWorkbook()
    const sheetId = workbook.getSheetId('Cases')!
    try {
      workbook.moveColumns(sheetId, 1, 1, 4)
      expect(workbook.getCellFormula(addressToCell('F1'))).toBe('=SUM(B1:B1)')
      expect(normalizedCellValue(workbook.getCellValue(addressToCell('F1')))).toEqual(structuralMoveColumnFormulaOracleCell.value)

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-structural-move-column-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=SUM(B1:B1)')
        expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual(structuralMoveColumnFormulaOracleCell.value)
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('keeps dynamic-array spill metadata valid after structural row inserts', () => {
    const workbook = buildDynamicSpillOracleWorkbook()
    const sheetId = workbook.getSheetId('Cases')!
    try {
      workbook.addRows(sheetId, 0, 1)
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B2', rows: 3, cols: 1 }])
      expect(['B2', 'B3', 'B4'].map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedDynamicSpillOracleValues.map((expected) => expected.value),
      )
      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-structural-dynamic-spill-oracle.xlsx')
      expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B2', rows: 3, cols: 1 }])
    } finally {
      workbook.dispose()
    }
  })

  it('rematerializes dynamic-array spill metadata after structural row edits through spill children', () => {
    const assertVerticalRowEdit = (edit: (workbook: WorkPaper, sheetId: number) => void): void => {
      const workbook = buildShrinkingSpillReferenceOracleWorkbook()
      const sheetId = workbook.getSheetId('Cases')!
      try {
        edit(workbook, sheetId)
        expect(
          blockedSpillReferenceOracleAddresses.map((address) => ({
            address,
            value: normalizedCellValue(workbook.getCellValue(addressToCell(address))),
          })),
        ).toEqual(expectedUnblockedSpillReferenceOracleValues)
        expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])

        const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-structural-row-edit-spill-oracle.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            blockedSpillReferenceOracleAddresses.map((address) => ({
              address,
              value: normalizedCellValue(reimported.getCellValue(addressToCell(address))),
            })),
          ).toEqual(expectedUnblockedSpillReferenceOracleValues)
          expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
        } finally {
          reimported.dispose()
        }
      } finally {
        workbook.dispose()
      }
    }

    assertVerticalRowEdit((workbook, sheetId) => workbook.addRows(sheetId, 1, 1))
    assertVerticalRowEdit((workbook, sheetId) => workbook.removeRows(sheetId, 1, 1))
  })
})
