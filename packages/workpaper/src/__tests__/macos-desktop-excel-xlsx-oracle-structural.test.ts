import type { NormalizedFormulaValue } from './macos-desktop-excel-xlsx-oracle-test-helpers.js'
import {
  addressToCell,
  aggregateOptionsOracleAddresses,
  blockedSpillReferenceOracleAddresses,
  buildAggregateOptionsOracleWorkbook,
  buildBiligDataTableXlsx,
  buildBiligOneVariableDataTableXlsx,
  buildChooseArrayIndexOracleWorkbook,
  buildDynamicSpillOracleWorkbook,
  buildFutureFunctionOracleWorkbook,
  buildHeadlessExcelComparisons,
  buildHorizontalStructuralSpillOracleWorkbook,
  buildIndexImplicitIntersectionOracleWorkbook,
  buildOffsetImplicitIntersectionOracleWorkbook,
  buildOracleWorkbook,
  buildReportSummary,
  buildShrinkingSpillReferenceOracleWorkbook,
  buildSingleImplicitIntersectionOracleWorkbook,
  buildSpillReferenceOracleWorkbook,
  buildStructuralMoveColumnOracleWorkbook,
  buildTableColumnDeleteDefinedNameOracleEngine,
  buildTableColumnDeleteOracleEngine,
  buildTableColumnInsertOracleEngine,
  buildTableEmptyBodyOracleEngine,
  buildTableHeaderRenameDefinedNameOracleEngine,
  buildTableHeaderRenameOracleEngine,
  buildTextsplitErrorOracleWorkbook,
  buildTwoDimensionalStructuralSpillOracleWorkbook,
  chooseArrayIndexOracleAddresses,
  dataTableFormulasWarning,
  dataTableOracleAddresses,
  describe,
  dynamicSpillOracleAddresses,
  ErrorCode,
  expect,
  expectedAggregateOptionsOracleCells,
  expectedChooseArrayIndexOracleValues,
  expectedDataTableImportedFormulaByAddress,
  expectedDataTableOracleValues,
  expectedDesktopExcelBlockedSpillReferenceValues,
  expectedDesktopExcelHorizontalStructuralSpillValues,
  expectedDesktopExcelShrinkingSpillReferenceOracleCells,
  expectedDesktopExcelSingleImplicitIntersectionOracleCells,
  expectedDynamicSpillOracleValues,
  expectedFutureFunctionOracleCells,
  expectedHorizontalStructuralSpillOracleValues,
  expectedIndexImplicitIntersectionOracleCells,
  expectedOffsetImplicitIntersectionOracleCells,
  expectedOneVariableDataTableImportedFormulaByAddress,
  expectedOneVariableDataTableOracleValues,
  expectedOracleCells,
  expectedShrinkingSpillReferenceOracleCells,
  expectedSingleImplicitIntersectionOracleValues,
  expectedSpillReferenceOracleCells,
  expectedTextsplitErrorOracleCells,
  expectedUnblockedSpillReferenceOracleValues,
  exportXlsx,
  futureFunctionOracleAddresses,
  horizontalStructuralSpillOracleAddresses,
  importXlsx,
  indexImplicitIntersectionConfig,
  indexImplicitIntersectionOracleAddresses,
  isMacosExcelInstalled,
  it,
  join,
  mkdtempSync,
  normalizedCellValue,
  offsetImplicitIntersectionOracleAddresses,
  oneVariableDataTableOracleAddresses,
  oracleFormulaAddresses,
  readFileSync,
  rmSync,
  runMacosExcelInspectionOracle,
  runMacosExcelStructuralOperationOracle,
  shrinkingSpillReferenceOracleAddresses,
  singleImplicitIntersectionOracleAddresses,
  spillReferenceOracleAddresses,
  structuralMoveColumnFormulaOracleCell,
  tableEmptyBodyOracleCell,
  textsplitErrorOracleAddresses,
  tmpdir,
  ValueTag,
  workbookConfig,
  WorkPaper,
  writeFileSync,
} from './macos-desktop-excel-xlsx-oracle-test-helpers.js'

describe('macOS Desktop Excel XLSX oracle structural and table compatibility', () => {
  it('rematerializes horizontal dynamic-array spill metadata after structural column edits through spill children', () => {
    const assertHorizontalColumnEdit = (edit: (workbook: WorkPaper, sheetId: number) => void): void => {
      const workbook = buildHorizontalStructuralSpillOracleWorkbook()
      const sheetId = workbook.getSheetId('Cases')!
      try {
        edit(workbook, sheetId)
        expect(
          horizontalStructuralSpillOracleAddresses.map((address) => ({
            address,
            value: normalizedCellValue(workbook.getCellValue(addressToCell(address))),
          })),
        ).toEqual(expectedHorizontalStructuralSpillOracleValues)
        expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 3 }])

        const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-structural-column-edit-spill-oracle.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            horizontalStructuralSpillOracleAddresses.map((address) => ({
              address,
              value: normalizedCellValue(reimported.getCellValue(addressToCell(address))),
            })),
          ).toEqual(expectedHorizontalStructuralSpillOracleValues)
          expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 3 }])
        } finally {
          reimported.dispose()
        }
      } finally {
        workbook.dispose()
      }
    }

    assertHorizontalColumnEdit((workbook, sheetId) => workbook.addColumns(sheetId, 2, 1))
    assertHorizontalColumnEdit((workbook, sheetId) => workbook.removeColumns(sheetId, 2, 1))
  })

  it('rematerializes vertical dynamic-array spill metadata after moving the owner row', () => {
    const workbook = buildShrinkingSpillReferenceOracleWorkbook()
    const sheetId = workbook.getSheetId('Cases')!
    try {
      workbook.moveRows(sheetId, 0, 1, 2)
      expect(workbook.getCellFormula(addressToCell('B3'))).toBe('=SEQUENCE(A3,1,1,1)')
      expect(workbook.getCellFormula(addressToCell('D3'))).toBe('=SUM(B3#)')
      expect(workbook.getCellFormula(addressToCell('E3'))).toBe('=ROWS(B3#)')
      expect(workbook.getCellFormula(addressToCell('F3'))).toBe('=IFERROR(INDEX(B3#,2),"missing")')
      expect(
        ['B3', 'B4', 'B5', 'D3', 'E3', 'F3'].map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual([
        { kind: 'number', value: 1 },
        { kind: 'number', value: 2 },
        { kind: 'number', value: 3 },
        { kind: 'number', value: 6 },
        { kind: 'number', value: 3 },
        { kind: 'number', value: 2 },
      ])
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B3', rows: 3, cols: 1 }])
    } finally {
      workbook.dispose()
    }
  })

  it('rematerializes horizontal dynamic-array spill metadata after moving the owner column', () => {
    const workbook = buildHorizontalStructuralSpillOracleWorkbook()
    const sheetId = workbook.getSheetId('Cases')!
    try {
      workbook.moveColumns(sheetId, 1, 1, 2)
      expect(workbook.getCellFormula(addressToCell('C1'))).toBe('=SEQUENCE(1,3,1,1)')
      expect(workbook.getCellFormula(addressToCell('A3'))).toBe('=SUM(C1#)')
      expect(workbook.getCellFormula(addressToCell('A4'))).toBe('=COLUMNS(C1#)')
      expect(workbook.getCellFormula(addressToCell('A5'))).toBe('=IFERROR(INDEX(C1#,1,2),"missing")')
      expect(
        ['C1', 'D1', 'E1', 'A3', 'A4', 'A5'].map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual([
        { kind: 'number', value: 1 },
        { kind: 'number', value: 2 },
        { kind: 'number', value: 3 },
        { kind: 'number', value: 6 },
        { kind: 'number', value: 3 },
        { kind: 'number', value: 2 },
      ])
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'C1', rows: 1, cols: 3 }])
    } finally {
      workbook.dispose()
    }
  })

  it('keeps two-dimensional spill metadata valid after moving the owner column', () => {
    const workbook = buildTwoDimensionalStructuralSpillOracleWorkbook()
    const sheetId = workbook.getSheetId('Cases')!
    try {
      workbook.moveColumns(sheetId, 1, 1, 3)
      expect(workbook.getCellFormula(addressToCell('D2'))).toBe('=SEQUENCE(2,3,1,1)')
      expect(workbook.getCellFormula(addressToCell('G2'))).toBe('=SUM(D2#)')
      expect(workbook.getCellFormula(addressToCell('G3'))).toBe('=ROWS(D2#)')
      expect(workbook.getCellFormula(addressToCell('G4'))).toBe('=COLUMNS(D2#)')
      expect(workbook.getCellFormula(addressToCell('G5'))).toBe('=IFERROR(INDEX(D2#,2,2),"missing")')
      expect(
        ['D2', 'E2', 'F2', 'D3', 'E3', 'F3', 'G2', 'G3', 'G4', 'G5'].map((address) =>
          normalizedCellValue(workbook.getCellValue(addressToCell(address))),
        ),
      ).toEqual([
        { kind: 'number', value: 1 },
        { kind: 'number', value: 2 },
        { kind: 'number', value: 3 },
        { kind: 'number', value: 4 },
        { kind: 'number', value: 5 },
        { kind: 'number', value: 6 },
        { kind: 'number', value: 21 },
        { kind: 'number', value: 2 },
        { kind: 'number', value: 3 },
        { kind: 'number', value: 5 },
      ])
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'D2', rows: 2, cols: 3 }])
    } finally {
      workbook.dispose()
    }
  })

  it('generates Excel-compatible table headers when inserting columns inside tables', async () => {
    const engine = await buildTableColumnInsertOracleEngine()

    engine.insertColumns('Data', 1, 1)

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'D3',
      columnNames: ['Region', 'Column1', 'Revenue', 'Margin'],
    })
    expect(engine.getCellValue('Data', 'B1')).toMatchObject({ tag: ValueTag.String, value: 'Column1' })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(Sales[Margin])')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-column-insert-oracle.xlsx')
    expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'D3',
      columnNames: ['Region', 'Column1', 'Revenue', 'Margin'],
    })
  })

  it('rewrites deleted table-column structured references before XLSX export', async () => {
    const engine = await buildTableColumnDeleteOracleEngine()

    engine.deleteColumns('Data', 1, 1)

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Margin'],
    })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(#REF!)')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Margin])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-column-delete-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B3',
        columnNames: ['Region', 'Margin'],
      })
      expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(#REF!)')
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual({
        kind: 'error',
        value: String(ErrorCode.Ref),
      })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 5 })
    } finally {
      reimported.dispose()
    }
  })

  it('renames table headers and structured references before XLSX export', async () => {
    const engine = await buildTableHeaderRenameOracleEngine()

    engine.setCellValue('Data', 'B1', 'Revenue')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Revenue', 'Margin'],
    })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Revenue])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(Sales[Margin])')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-header-rename-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'C3',
        columnNames: ['Region', 'Revenue', 'Margin'],
      })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual({ kind: 'number', value: 5 })
    } finally {
      reimported.dispose()
    }
  })

  it('rewrites table-column defined names to #REF! before XLSX export', async () => {
    const engine = await buildTableColumnDeleteDefinedNameOracleEngine()

    engine.deleteColumns('Data', 1, 1)

    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'formula', formula: '=#REF!' },
    })
    expect(engine.getDefinedName('SalesAmountFormula')).toEqual({
      name: 'SalesAmountFormula',
      value: { kind: 'formula', formula: '=#REF!' },
    })
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-column-delete-defined-name-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
        { name: 'SalesAmount', value: { kind: 'formula', formula: '=#REF!' } },
        { name: 'SalesAmountFormula', value: { kind: 'formula', formula: '=#REF!' } },
      ])
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual({ kind: 'error', value: String(ErrorCode.Ref) })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'error', value: String(ErrorCode.Ref) })
    } finally {
      reimported.dispose()
    }
  })

  it('renames table-column defined names before XLSX export', async () => {
    const engine = await buildTableHeaderRenameDefinedNameOracleEngine()

    engine.setCellValue('Data', 'B1', 'Revenue')

    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'structured-ref', tableName: 'Sales', columnName: 'Revenue' },
    })
    expect(engine.getDefinedName('SalesAmountFormula')).toEqual({
      name: 'SalesAmountFormula',
      value: { kind: 'formula', formula: '=Sales[Revenue]' },
    })
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 30 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-header-rename-defined-name-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
        { name: 'SalesAmount', value: { kind: 'formula', formula: '=Sales[Revenue]' } },
        { name: 'SalesAmountFormula', value: { kind: 'formula', formula: '=Sales[Revenue]' } },
      ])
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual({ kind: 'number', value: 30 })
    } finally {
      reimported.dispose()
    }
  })

  it('routes WorkPaper table header edits through structured reference rewrites', async () => {
    const engine = await buildTableHeaderRenameOracleEngine()
    const workbook = WorkPaper.buildFromSnapshot(engine.exportSnapshot(), workbookConfig)
    try {
      workbook.setCellContents(addressToCell('B1'), 'Revenue')

      expect(workbook.getCellFormula(addressToCell('E1'))).toBe('=SUM(Sales[Revenue])')
      expect(normalizedCellValue(workbook.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
      expect(workbook.exportSnapshot().workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'C3',
        columnNames: ['Region', 'Revenue', 'Margin'],
      })
    } finally {
      workbook.dispose()
    }
  })

  it('keeps table structured-reference aggregates valid when deleting the only data row', async () => {
    const engine = await buildTableEmptyBodyOracleEngine()

    engine.deleteRows('Data', 1, 1)

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Region', 'Amount'],
    })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(Sales[Amount])')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Number, value: 0 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-empty-body-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B2',
        columnNames: ['Region', 'Amount'],
      })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual(tableEmptyBodyOracleCell.value)
    } finally {
      reimported.dispose()
    }
  })

  it('imports @bilig/xlsx two-variable data-table outputs into headless calculable formulas', () => {
    const imported = importXlsx(buildBiligDataTableXlsx(), 'headless-bilig-xlsx-data-table-oracle.xlsx')
    expect(imported.warnings).not.toContain(dataTableFormulasWarning)
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas).toEqual([
      {
        address: 'C3',
        formulaXml: '<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2"/>',
      },
    ])

    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(dataTableOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedDataTableOracleValues.map((expected) => expected.value),
      )
      for (const [address, formula] of expectedDataTableImportedFormulaByAddress) {
        expect(workbook.getCellFormula(addressToCell(address))).toBe(formula)
      }
    } finally {
      workbook.dispose()
    }
  })

  it('imports @bilig/xlsx one-variable data-table outputs into headless calculable formulas', () => {
    const imported = importXlsx(buildBiligOneVariableDataTableXlsx(), 'headless-bilig-xlsx-one-variable-data-table-oracle.xlsx')
    expect(imported.warnings).not.toContain(dataTableFormulasWarning)
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas).toEqual([
      {
        address: 'C2',
        formulaXml: '<f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1"/>',
      },
      {
        address: 'B6',
        formulaXml: '<f t="dataTable" ref="B6:B8" dt2D="0" dtr="0" r1="A1"/>',
      },
    ])

    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(
        oneVariableDataTableOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedOneVariableDataTableOracleValues.map((expected) => expected.value))
      for (const [address, formula] of expectedOneVariableDataTableImportedFormulaByAddress) {
        expect(workbook.getCellFormula(addressToCell(address))).toBe(formula)
      }
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips fresh Desktop Excel recalculation caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-oracle.xlsx')
        const workbook = buildOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: oracleFormulaAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          const comparisons = buildHeadlessExcelComparisons(reimported, excelResult.cells)
          const summary = buildReportSummary({
            workbooks: [
              {
                id: 'headless-oracle',
                workbook: 'headless-oracle.xlsx',
                elapsedMs: 0,
                formulaCells: comparisons.length,
                status: 'ok',
                comparisons,
              },
            ],
          })

          expect(comparisons.map((comparison) => comparison.classification)).toEqual(
            oracleFormulaAddresses.map(() => 'bilig_matches_excel'),
          )
          expect(summary).toMatchObject({
            biligVsFreshExcelMatchRate: 1,
            comparableFormulaCells: oracleFormulaAddresses.length,
            realBiligMismatches: 0,
            totalFormulaCells: oracleFormulaAddresses.length,
          })
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips Desktop Excel AGGREGATE option semantics into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-aggregate-options-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-aggregate-options-oracle.xlsx')
        const workbook = buildAggregateOptionsOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: aggregateOptionsOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedAggregateOptionsOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-aggregate-options-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          const comparisons = buildHeadlessExcelComparisons(reimported, excelResult.cells, 'headless-aggregate-options-oracle')
          const summary = buildReportSummary({
            workbooks: [
              {
                id: 'headless-aggregate-options-oracle',
                workbook: 'headless-aggregate-options-oracle.xlsx',
                elapsedMs: 0,
                formulaCells: comparisons.length,
                status: 'ok',
                comparisons,
              },
            ],
          })

          expect(comparisons.map((comparison) => comparison.classification)).toEqual(
            aggregateOptionsOracleAddresses.map(() => 'bilig_matches_excel'),
          )
          expect(summary).toMatchObject({
            biligVsFreshExcelMatchRate: 1,
            comparableFormulaCells: aggregateOptionsOracleAddresses.length,
            realBiligMismatches: 0,
            totalFormulaCells: aggregateOptionsOracleAddresses.length,
          })
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel standalone INDEX implicit-intersection semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-index-implicit-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-index-implicit-intersection-oracle.xlsx')
        const workbook = buildIndexImplicitIntersectionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Sheet1',
          formulaCells: [],
          inspectCells: indexImplicitIntersectionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedIndexImplicitIntersectionOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-index-implicit-intersection-recalculated.xlsx')
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
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel standalone OFFSET implicit-intersection semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-offset-implicit-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-offset-implicit-intersection-oracle.xlsx')
        const workbook = buildOffsetImplicitIntersectionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Sheet1',
          formulaCells: [],
          inspectCells: offsetImplicitIntersectionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedOffsetImplicitIntersectionOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-offset-implicit-intersection-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, indexImplicitIntersectionConfig)
        try {
          expect(
            offsetImplicitIntersectionOracleAddresses.map((address) =>
              normalizedCellValue(reimported.getCellValue(addressToCell(address))),
            ),
          ).toEqual(expectedOffsetImplicitIntersectionOracleCells.map((expected) => expected.value))
          expect(reimported.engine.getSpillRanges()).toEqual([])
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips fresh Desktop Excel future-function recalculation caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-future-function-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-future-function-oracle.xlsx')
        const workbook = buildFutureFunctionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: futureFunctionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedFutureFunctionOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-future-function-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          const comparisons = buildHeadlessExcelComparisons(reimported, excelResult.cells, 'headless-future-function-oracle')
          const summary = buildReportSummary({
            workbooks: [
              {
                id: 'headless-future-function-oracle',
                workbook: 'headless-future-function-oracle.xlsx',
                elapsedMs: 0,
                formulaCells: comparisons.length,
                status: 'ok',
                comparisons,
              },
            ],
          })

          expect(comparisons.map((comparison) => comparison.classification)).toEqual(
            futureFunctionOracleAddresses.map(() => 'bilig_matches_excel'),
          )
          expect(summary).toMatchObject({
            biligVsFreshExcelMatchRate: 1,
            comparableFormulaCells: futureFunctionOracleAddresses.length,
            realBiligMismatches: 0,
            totalFormulaCells: futureFunctionOracleAddresses.length,
          })
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel SINGLE implicit-intersection semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-single-implicit-intersection-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-single-implicit-intersection-oracle.xlsx')
        const workbook = buildSingleImplicitIntersectionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: singleImplicitIntersectionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedDesktopExcelSingleImplicitIntersectionOracleCells)

        const imported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-single-implicit-intersection-oracle-recalculated.xlsx',
        )
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            singleImplicitIntersectionOracleAddresses.map((address) =>
              normalizedCellValue(reimported.getCellValue(addressToCell(address))),
            ),
          ).toEqual(expectedSingleImplicitIntersectionOracleValues.map((expected) => expected.value))
          expect(reimported.getCellFormula(addressToCell('C1'))).toBe('=SINGLE(A1:A3)')
          expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(SINGLE(A1:A3))')
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips Desktop Excel native dynamic-array spill caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-dynamic-spill-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-dynamic-spill-oracle.xlsx')
        const workbook = buildDynamicSpillOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: dynamicSpillOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedDynamicSpillOracleValues)
        expect(excelResult.cells[0]?.formula).toBe('=MAP(A1:A3,LAMBDA(x,x*2))')

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-dynamic-spill-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            dynamicSpillOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedDynamicSpillOracleValues.map((expected) => expected.value))
          expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel spill-reference formula consumer semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-spill-reference-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-spill-reference-oracle.xlsx')
        const workbook = buildSpillReferenceOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: spillReferenceOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedSpillReferenceOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-spill-reference-oracle-recalculated.xlsx')
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
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel spill-reference semantics after a dynamic-array shrink to one cell',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-shrinking-spill-reference-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-shrinking-spill-reference-oracle.xlsx')
        const workbook = buildShrinkingSpillReferenceOracleWorkbook()
        try {
          workbook.setCellContents(addressToCell('A1'), 1)
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: shrinkingSpillReferenceOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedDesktopExcelShrinkingSpillReferenceOracleCells)

        const imported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-shrinking-spill-reference-oracle-recalculated.xlsx',
        )
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
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel spill-reference semantics after authoring and clearing a spill child',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-blocked-spill-reference-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-blocked-spill-reference-oracle.xlsx')
        const workbook = buildShrinkingSpillReferenceOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const blockedExcelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Cases',
          operations: [{ kind: 'setCellValue', address: 'B2', value: 99 }],
          inspectCells: blockedSpillReferenceOracleAddresses,
          saveWorkbook: true,
        })

        expect(blockedExcelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(
          expectedDesktopExcelBlockedSpillReferenceValues,
        )
        expect(blockedExcelResult.cells[0]?.formula).toBe('=SEQUENCE(A1,1,1,1)')
        expect(blockedExcelResult.cells[3]?.formula).toBe('=SUM(B1#)')
        expect(blockedExcelResult.cells[4]?.formula).toBe('=ROWS(B1#)')
        expect(blockedExcelResult.cells[5]?.formula).toBe('=IFERROR(INDEX(B1#,2),"missing")')

        const blockedImported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-blocked-spill-reference-oracle-recalculated.xlsx',
        )
        const blockedReimported = WorkPaper.buildFromSnapshot(blockedImported.snapshot, workbookConfig)
        try {
          expect(blockedImported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 1 }])
          expect(normalizedCellValue(blockedReimported.getCellValue(addressToCell('B2')))).toEqual({ kind: 'number', value: 99 })
          expect(normalizedCellValue(blockedReimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 1 })
          expect(normalizedCellValue(blockedReimported.getCellValue(addressToCell('F1')))).toEqual({
            kind: 'string',
            value: 'missing',
          })
        } finally {
          blockedReimported.dispose()
        }

        const unblockedExcelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Cases',
          operations: [{ kind: 'clearCell', address: 'B2' }],
          inspectCells: blockedSpillReferenceOracleAddresses,
          saveWorkbook: true,
        })

        expect(unblockedExcelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(
          expectedUnblockedSpillReferenceOracleValues,
        )

        const unblockedImported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-unblocked-spill-reference-oracle-recalculated.xlsx',
        )
        const unblockedReimported = WorkPaper.buildFromSnapshot(unblockedImported.snapshot, workbookConfig)
        try {
          expect(
            blockedSpillReferenceOracleAddresses.map((address) => ({
              address,
              value: normalizedCellValue(unblockedReimported.getCellValue(addressToCell(address))),
            })),
          ).toEqual(expectedUnblockedSpillReferenceOracleValues)
          expect(unblockedImported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
        } finally {
          unblockedReimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel spill rematerialization after row edits through spill children',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-structural-delete-spill-oracle-'))
      try {
        for (const { fileName, operation } of [
          { fileName: 'headless-structural-insert-spill-oracle.xlsx', operation: { kind: 'insertRows' as const, range: '2:2' } },
          { fileName: 'headless-structural-delete-spill-oracle.xlsx', operation: { kind: 'deleteRows' as const, range: '2:2' } },
        ]) {
          const workbookPath = join(tempDir, fileName)
          const workbook = buildShrinkingSpillReferenceOracleWorkbook()
          try {
            writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
          } finally {
            workbook.dispose()
          }

          const excelResult = runMacosExcelStructuralOperationOracle({
            workbookPath,
            worksheetName: 'Cases',
            operations: [operation],
            inspectCells: blockedSpillReferenceOracleAddresses,
            saveWorkbook: true,
          })

          expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedUnblockedSpillReferenceOracleValues)
          expect(excelResult.cells[0]?.formula).toBe('=SEQUENCE(A1,1,1,1)')
          expect(excelResult.cells[3]?.formula).toBe('=SUM(B1#)')
          expect(excelResult.cells[4]?.formula).toBe('=ROWS(B1#)')
          expect(excelResult.cells[5]?.formula).toBe('=IFERROR(INDEX(B1#,2),"missing")')

          const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), `recalculated-${fileName}`)
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
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel horizontal spill rematerialization after column edits through spill children',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-structural-column-spill-oracle-'))
      try {
        for (const { fileName, operation } of [
          {
            fileName: 'headless-structural-insert-horizontal-spill-oracle.xlsx',
            operation: { kind: 'insertColumns' as const, range: 'C:C' },
          },
          {
            fileName: 'headless-structural-delete-horizontal-spill-oracle.xlsx',
            operation: { kind: 'deleteColumns' as const, range: 'C:C' },
          },
        ]) {
          const workbookPath = join(tempDir, fileName)
          const workbook = buildHorizontalStructuralSpillOracleWorkbook()
          try {
            writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
          } finally {
            workbook.dispose()
          }

          const excelResult = runMacosExcelStructuralOperationOracle({
            workbookPath,
            worksheetName: 'Cases',
            operations: [operation],
            inspectCells: horizontalStructuralSpillOracleAddresses,
            saveWorkbook: true,
          })

          expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(
            expectedDesktopExcelHorizontalStructuralSpillValues,
          )
          expect(excelResult.cells[0]?.formula).toBe('=SEQUENCE(1,3,1,1)')
          expect(excelResult.cells[4]?.formula).toBe('=SUM(B1#)')
          expect(excelResult.cells[5]?.formula).toBe('=COLUMNS(B1#)')
          expect(excelResult.cells[6]?.formula).toBe('=IFERROR(INDEX(B1#,1,2),"missing")')

          const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), `recalculated-${fileName}`)
          const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
          try {
            expect(
              horizontalStructuralSpillOracleAddresses.map((address) => ({
                address,
                value: normalizedCellValue(reimported.getCellValue(addressToCell(address))),
              })),
            ).toEqual(expectedHorizontalStructuralSpillOracleValues)
            expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 1, cols: 3 }])
          } finally {
            reimported.dispose()
          }
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel spill rematerialization after moving dynamic-array owners',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const cases = [
        {
          fileName: 'headless-structural-move-owner-row-spill-oracle.xlsx',
          buildWorkbook: buildShrinkingSpillReferenceOracleWorkbook,
          operations: [{ kind: 'moveRows' as const, sourceRange: '1:1', destinationRange: '4:4' }],
          inspectCells: ['B3', 'B4', 'B5', 'D3', 'E3', 'F3'],
          expectedCells: [
            { address: 'B3', formula: '=SEQUENCE(A3,1,1,1)', value: { kind: 'number', value: 1 } },
            { address: 'B4', value: { kind: 'number', value: 2 } },
            { address: 'B5', value: { kind: 'number', value: 3 } },
            { address: 'D3', formula: '=SUM(B3#)', value: { kind: 'number', value: 6 } },
            { address: 'E3', formula: '=ROWS(B3#)', value: { kind: 'number', value: 3 } },
            { address: 'F3', formula: '=IFERROR(INDEX(B3#,2),"missing")', value: { kind: 'number', value: 2 } },
          ],
          expectedSpills: [{ sheetName: 'Cases', address: 'B3', rows: 3, cols: 1 }],
        },
        {
          fileName: 'headless-structural-move-owner-column-spill-oracle.xlsx',
          buildWorkbook: buildHorizontalStructuralSpillOracleWorkbook,
          operations: [{ kind: 'moveColumns' as const, sourceRange: 'B:B', destinationRange: 'D:D' }],
          inspectCells: ['C1', 'D1', 'E1', 'A3', 'A4', 'A5'],
          expectedCells: [
            { address: 'C1', formula: '=SEQUENCE(1,3,1,1)', value: { kind: 'number', value: 1 } },
            { address: 'D1', value: { kind: 'number', value: 2 } },
            { address: 'E1', value: { kind: 'number', value: 3 } },
            { address: 'A3', formula: '=SUM(C1#)', value: { kind: 'number', value: 6 } },
            { address: 'A4', formula: '=COLUMNS(C1#)', value: { kind: 'number', value: 3 } },
            { address: 'A5', formula: '=IFERROR(INDEX(C1#,1,2),"missing")', value: { kind: 'number', value: 2 } },
          ],
          expectedSpills: [{ sheetName: 'Cases', address: 'C1', rows: 1, cols: 3 }],
        },
        {
          fileName: 'headless-structural-move-2d-owner-column-spill-oracle.xlsx',
          buildWorkbook: buildTwoDimensionalStructuralSpillOracleWorkbook,
          operations: [{ kind: 'moveColumns' as const, sourceRange: 'B:B', destinationRange: 'E:E' }],
          inspectCells: ['D2', 'E2', 'F2', 'D3', 'E3', 'F3', 'G2', 'G3', 'G4', 'G5'],
          expectedCells: [
            { address: 'D2', formula: '=SEQUENCE(2,3,1,1)', value: { kind: 'number', value: 1 } },
            { address: 'E2', value: { kind: 'number', value: 2 } },
            { address: 'F2', value: { kind: 'number', value: 3 } },
            { address: 'D3', value: { kind: 'number', value: 4 } },
            { address: 'E3', value: { kind: 'number', value: 5 } },
            { address: 'F3', value: { kind: 'number', value: 6 } },
            { address: 'G2', formula: '=SUM(D2#)', value: { kind: 'number', value: 21 } },
            { address: 'G3', formula: '=ROWS(D2#)', value: { kind: 'number', value: 2 } },
            { address: 'G4', formula: '=COLUMNS(D2#)', value: { kind: 'number', value: 3 } },
            { address: 'G5', formula: '=IFERROR(INDEX(D2#,2,2),"missing")', value: { kind: 'number', value: 5 } },
          ],
          expectedSpills: [{ sheetName: 'Cases', address: 'D2', rows: 2, cols: 3 }],
        },
      ] as const

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-structural-owner-move-spill-oracle-'))
      try {
        for (const testCase of cases) {
          const workbookPath = join(tempDir, testCase.fileName)
          const workbook = testCase.buildWorkbook()
          try {
            writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
          } finally {
            workbook.dispose()
          }

          const excelResult = runMacosExcelStructuralOperationOracle({
            workbookPath,
            worksheetName: 'Cases',
            operations: testCase.operations,
            inspectCells: testCase.inspectCells,
            saveWorkbook: true,
          })

          const actualCells = excelResult.cells.map(({ address, formula, value }) => {
            const actualCell: { address: string; formula?: string; value: NormalizedFormulaValue } = { address, value }
            if (formula) {
              actualCell.formula = formula
            }
            return actualCell
          })
          expect(actualCells).toEqual(testCase.expectedCells)

          const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), `recalculated-${testCase.fileName}`)
          const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
          try {
            expect(
              testCase.inspectCells.map((address) => ({
                address,
                value: normalizedCellValue(reimported.getCellValue(addressToCell(address))),
              })),
            ).toEqual(testCase.expectedCells.map(({ address, value }) => ({ address, value })))
            expect(imported.snapshot.workbook.metadata?.spills).toEqual(testCase.expectedSpills)
          } finally {
            reimported.dispose()
          }
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips Desktop Excel TEXTSPLIT error spill-child caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-textsplit-error-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-textsplit-error-oracle.xlsx')
        const workbook = buildTextsplitErrorOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: textsplitErrorOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedTextsplitErrorOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-textsplit-error-oracle-recalculated.xlsx')
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
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel CHOOSE array-index virtual table semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-choose-array-index-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-choose-array-index-oracle.xlsx')
        const workbook = buildChooseArrayIndexOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'ChooseRef',
          formulaCells: [],
          inspectCells: chooseArrayIndexOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedChooseArrayIndexOracleValues)
        expect(excelResult.cells[0]?.formula).toBe('=CHOOSE({1,2},A1:A3,B1:B3)')

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-choose-array-index-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            chooseArrayIndexOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedChooseArrayIndexOracleValues.map((expected) => expected.value))
          expect(imported.snapshot.workbook.metadata?.spills).toEqual([
            { sheetName: 'ChooseRef', address: 'E1', rows: 3, cols: 2 },
            { sheetName: 'ChooseRef', address: 'H2', rows: 3, cols: 1 },
          ])
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel column-move formula rewrite semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-structural-move-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-structural-move-column-oracle.xlsx')
        const workbook = buildStructuralMoveColumnOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Cases',
          operations: [{ kind: 'moveColumns', sourceRange: 'B:B', destinationRange: 'F:F' }],
          inspectCells: ['F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual([structuralMoveColumnFormulaOracleCell])

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-structural-move-column-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=SUM(B1:B1)')
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual(structuralMoveColumnFormulaOracleCell.value)
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )
})
