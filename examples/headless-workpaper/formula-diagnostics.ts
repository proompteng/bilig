import { WorkPaper, type WorkPaperCellAddress, type WorkPaperFormulaDiagnostic } from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

const workbook = WorkPaper.buildFromSheets(
  {
    Tax: [
      ['Metric', 'Value', 'Date serial', 'Date label'],
      ['Cash flow 0', -100000, 45292, '2024-01-01'],
      ['Cash flow 1', 25000, 45658, '2025-01-01'],
      ['Cash flow 2', 35000, 46023, '2026-01-01'],
      ['Cash flow 3', 45000, 46388, '2027-01-01'],
      ['Invalid XIRR', '=XIRR(B2:B5,D2:D5)', null, null],
      ['Valid XIRR', '=XIRR(B2:B5,C2:C5)', null, null],
    ],
  },
  { maxRows: 1000, maxColumns: 100, useColumnIndex: true },
)

const taxSheet = requireSheet(workbook, 'Tax')
const invalid = { sheet: taxSheet, row: 5, col: 1 }
const valid = { sheet: taxSheet, row: 6, col: 1 }
const invalidDiagnostic = requireFirstDiagnostic(workbook.getCellFormulaDiagnostics(invalid))

const output = {
  verified: true,
  invalidDisplay: workbook.getCellDisplayValue(invalid),
  invalidDiagnostics: [
    {
      code: invalidDiagnostic.code,
      functionName: invalidDiagnostic.functionName,
      errorText: invalidDiagnostic.errorText,
      references: invalidDiagnostic.references,
    },
  ],
  validDisplay: workbook.getCellDisplayValue(valid),
  validValue: readNumber(workbook, valid, 'Valid XIRR'),
}

assertOutput(output, invalidDiagnostic)
console.log(JSON.stringify(output, null, 2))

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function requireFirstDiagnostic(diagnostics: readonly WorkPaperFormulaDiagnostic[]): WorkPaperFormulaDiagnostic {
  const diagnostic = diagnostics[0]
  if (diagnostic === undefined) {
    throw new Error('Expected formula diagnostics for invalid XIRR')
  }
  return diagnostic
}

function readNumber(workpaper: WorkPaperInstance, address: WorkPaperCellAddress, label: string): number {
  const cell = workpaper.getCellValue(address)
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertOutput(actual: typeof output, diagnostic: WorkPaperFormulaDiagnostic): void {
  const expected = {
    verified: true,
    invalidDisplay: '#VALUE!',
    invalidDiagnostics: [
      {
        code: 'financial-unsupported-date-coercion',
        functionName: 'XIRR',
        errorText: '#VALUE!',
        references: ['Tax!D2:D5', 'Tax!D2'],
      },
    ],
    validDisplay: '0.02256857579464',
    validValue: 0.02256857579463996,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected formula diagnostic WorkPaper result: ${JSON.stringify(actual)}`)
  }

  if (typeof diagnostic.message !== 'string' || !diagnostic.message.includes('Use numeric Excel serial dates')) {
    throw new Error(`Unexpected formula diagnostic message: ${JSON.stringify(diagnostic)}`)
  }
}
