import {
  addressToCell,
  buildDataTableOracleWorkbook,
  buildOneVariableDataTableOracleWorkbook,
  buildTableColumnDeleteDefinedNameOracleEngine,
  buildTableColumnDeleteOracleEngine,
  buildTableColumnInsertOracleEngine,
  buildTableEmptyBodyOracleEngine,
  buildTableHeaderRenameDefinedNameOracleEngine,
  buildTableHeaderRenameOracleEngine,
  dataTableFormulasWarning,
  dataTableOracleAddresses,
  describe,
  ErrorCode,
  expect,
  expectedDataTableImportedFormulaByAddress,
  expectedDataTableOracleValues,
  expectedOneVariableDataTableImportedFormulaByAddress,
  expectedOneVariableDataTableOracleValues,
  exportXlsx,
  importXlsx,
  isMacosExcelInstalled,
  it,
  join,
  mkdtempSync,
  normalizedCellValue,
  oneVariableDataTableOracleAddresses,
  readFileSync,
  rmSync,
  runMacosExcelStructuralOperationOracle,
  tableColumnDeleteDefinedNameOracleFormulaCells,
  tableColumnDeleteOracleFormulaCells,
  tableColumnInsertOracleCells,
  tableEmptyBodyOracleCell,
  tableHeaderRenameDefinedNameOracleCells,
  tableHeaderRenameOracleCells,
  tmpdir,
  workbookConfig,
  WorkPaper,
  writeFileSync,
} from './macos-desktop-excel-xlsx-oracle-test-helpers.js'

describe('macOS Desktop Excel XLSX oracle gated native Excel compatibility', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table column-insert structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-column-insert-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-column-insert-oracle.xlsx')
        const engine = await buildTableColumnInsertOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertColumns', range: 'B:B' }],
          inspectCells: ['B1', 'F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(tableColumnInsertOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-column-insert-oracle-recalculated.xlsx')
        expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
          name: 'Sales',
          sheetName: 'Data',
          startAddress: 'A1',
          endAddress: 'D3',
          columnNames: ['Region', 'Column1', 'Revenue', 'Margin'],
        })
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table header-rename structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-header-rename-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-header-rename-oracle.xlsx')
        const engine = await buildTableHeaderRenameOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'setCellValue', address: 'B1', value: 'Revenue' }],
          inspectCells: ['B1', 'E1', 'F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(tableHeaderRenameOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-header-rename-oracle-recalculated.xlsx')
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
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table-header defined-name rename semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-header-defined-name-rename-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-header-defined-name-rename-oracle.xlsx')
        const engine = await buildTableHeaderRenameDefinedNameOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'setCellValue', address: 'B1', value: 'Revenue' }],
          inspectCells: ['B1', 'E1', 'F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(tableHeaderRenameDefinedNameOracleCells)

        const imported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-table-header-defined-name-rename-oracle-recalculated.xlsx',
        )
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
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table column-delete structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-column-delete-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-column-delete-oracle.xlsx')
        const engine = await buildTableColumnDeleteOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteColumns', range: 'B:B' }],
          inspectCells: ['D1', 'E1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, formula }) => ({ address, formula }))).toEqual(tableColumnDeleteOracleFormulaCells)
        expect(excelResult.cells[1]).toMatchObject({
          address: 'E1',
          rawValue: 'number\t5.0',
          value: { kind: 'number', value: 5 },
        })

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-column-delete-oracle-recalculated.xlsx')
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
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table-column defined-name deletion semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-column-defined-name-delete-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-column-defined-name-delete-oracle.xlsx')
        const engine = await buildTableColumnDeleteDefinedNameOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteColumns', range: 'B:B' }],
          inspectCells: ['D1', 'E1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, formula }) => ({ address, formula }))).toEqual(
          tableColumnDeleteDefinedNameOracleFormulaCells,
        )

        const imported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-table-column-defined-name-delete-oracle-recalculated.xlsx',
        )
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
            { name: 'SalesAmount', value: { kind: 'formula', formula: '=#REF!' } },
            { name: 'SalesAmountFormula', value: { kind: 'formula', formula: '=#REF!' } },
          ])
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual({
            kind: 'error',
            value: String(ErrorCode.Ref),
          })
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({
            kind: 'error',
            value: String(ErrorCode.Ref),
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
    'matches Desktop Excel empty-table-body structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-empty-body-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-empty-body-oracle.xlsx')
        const engine = await buildTableEmptyBodyOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteRows', range: '2:2' }],
          inspectCells: ['D1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual([tableEmptyBodyOracleCell])

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-empty-body-oracle-recalculated.xlsx')
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
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'imports Desktop Excel native data-table outputs into headless formulas',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-data-table-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-data-table-oracle.xlsx')
        const workbook = buildDataTableOracleWorkbook(false)
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'DataTable',
          operations: [{ kind: 'createDataTable', range: 'B2:D4', rowInput: 'A1', columnInput: 'A2' }],
          inspectCells: dataTableOracleAddresses,
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedDataTableOracleValues)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-data-table-oracle-recalculated.xlsx')
        expect(imported.warnings).not.toContain(dataTableFormulasWarning)
        expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas[0]?.address).toBe('C3')
        expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas[0]?.formulaXml).toContain('t="dataTable"')

        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(dataTableOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address))))).toEqual(
            expectedDataTableOracleValues.map((expected) => expected.value),
          )
          for (const [address, formula] of expectedDataTableImportedFormulaByAddress) {
            expect(reimported.getCellFormula(addressToCell(address))).toBe(formula)
          }
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
    'imports Desktop Excel native one-variable data-table outputs into headless formulas',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-one-variable-data-table-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-one-variable-data-table-oracle.xlsx')
        const workbook = buildOneVariableDataTableOracleWorkbook(false)
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'DataTable',
          operations: [
            { kind: 'createDataTable', range: 'B1:D2', rowInput: 'A1' },
            { kind: 'createDataTable', range: 'A5:B8', columnInput: 'A1' },
          ],
          inspectCells: oneVariableDataTableOracleAddresses,
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedOneVariableDataTableOracleValues)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-one-variable-data-table-oracle-recalculated.xlsx')
        expect(imported.warnings).not.toContain(dataTableFormulasWarning)
        expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas.map(({ address }) => address)).toEqual(['C2', 'B6'])

        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            oneVariableDataTableOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedOneVariableDataTableOracleValues.map((expected) => expected.value))
          for (const [address, formula] of expectedOneVariableDataTableImportedFormulaByAddress) {
            expect(reimported.getCellFormula(addressToCell(address))).toBe(formula)
          }
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
