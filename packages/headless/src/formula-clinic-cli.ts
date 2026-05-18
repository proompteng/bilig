import { existsSync, readFileSync, statSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { WorkPaper } from './work-paper.js'

export interface FormulaClinicCellRef {
  readonly sheetName: string | null
  readonly a1: string
}

export interface FormulaClinicFormulaSample {
  readonly sheet: string
  readonly cell: string
  readonly formula: string
}

export interface FormulaClinicImportedWorkbook {
  readonly snapshot: WorkbookSnapshot
  readonly sheetNames: readonly string[]
  readonly warnings: readonly string[]
}

export type FormulaClinicImportXlsx = (bytes: Uint8Array, fileName: string) => FormulaClinicImportedWorkbook

export interface FormulaClinicCliOptions {
  readonly cells: readonly FormulaClinicCellRef[]
  readonly evaluationTimeoutMs: number
  readonly filePath?: string
  readonly help: boolean
  readonly maxFormulaSamples: number
}

export interface FormulaClinicCliHost {
  readonly argv: readonly string[]
  readonly importXlsx: FormulaClinicImportXlsx
  readonly packageVersion?: string
  readonly readFile?: (path: string) => Uint8Array
  readonly statFileSizeBytes?: (path: string) => number
  readonly writeStderr?: (text: string) => void
  readonly writeStdout?: (text: string) => void
}

interface FormulaClinicReport {
  readonly error: string | null
  readonly fileName: string
  readonly fileSizeBytes: number
  readonly formulaCellCount: number
  readonly formulaSamples: readonly FormulaClinicFormulaSample[]
  readonly nonEmptyCellCount: number
  readonly packageVersion: string
  readonly readback: readonly string[]
  readonly sheetNames: readonly string[]
  readonly status: 'failed' | 'imported'
  readonly warningMessages: readonly string[]
}

export function runFormulaClinicCli(host: FormulaClinicCliHost): number {
  const writeStdout = host.writeStdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = host.writeStderr ?? ((text: string) => process.stderr.write(text))
  let options: FormulaClinicCliOptions

  try {
    options = parseFormulaClinicCliArgs(host.argv)
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n\n${formulaClinicHelpText()}`)
    return 1
  }

  if (options.help) {
    writeStdout(formulaClinicHelpText())
    return 0
  }

  if (options.filePath === undefined) {
    writeStderr(`Missing workbook path.\n\n${formulaClinicHelpText()}`)
    return 1
  }

  const packageVersion = host.packageVersion ?? readHeadlessPackageVersion()
  const filePath = options.filePath
  const fileName = basename(filePath)
  let fileSizeBytes = 0
  let bytes: Uint8Array

  try {
    fileSizeBytes = host.statFileSizeBytes?.(filePath) ?? statSync(filePath).size
    bytes = host.readFile?.(filePath) ?? new Uint8Array(readFileSync(filePath))
  } catch (error) {
    writeStdout(
      renderFormulaClinicReport({
        status: 'failed',
        packageVersion,
        fileName,
        fileSizeBytes,
        sheetNames: [],
        warningMessages: [],
        formulaCellCount: 0,
        nonEmptyCellCount: 0,
        formulaSamples: [],
        readback: [],
        error: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
      }),
    )
    return 1
  }

  try {
    const imported = host.importXlsx(bytes, fileName)
    const formulaSamples = collectFormulaSamples(imported.snapshot.sheets, options.maxFormulaSamples)
    const formulaCellCount = imported.snapshot.sheets.reduce(
      (count, sheet) => count + sheet.cells.filter((cell) => typeof cell.formula === 'string').length,
      0,
    )
    const nonEmptyCellCount = imported.snapshot.sheets.reduce((count, sheet) => count + sheet.cells.length, 0)
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
      evaluationTimeoutMs: options.evaluationTimeoutMs,
      useColumnIndex: true,
    })

    try {
      const readback = options.cells.map((cell) => readRequestedCell(workbook, imported.sheetNames, cell))
      writeStdout(
        renderFormulaClinicReport({
          status: 'imported',
          packageVersion,
          fileName,
          fileSizeBytes,
          sheetNames: imported.sheetNames,
          warningMessages: imported.warnings,
          formulaCellCount,
          nonEmptyCellCount,
          formulaSamples,
          readback,
          error: null,
        }),
      )
      return 0
    } finally {
      workbook.dispose()
    }
  } catch (error) {
    writeStdout(
      renderFormulaClinicReport({
        status: 'failed',
        packageVersion,
        fileName,
        fileSizeBytes,
        sheetNames: [],
        warningMessages: [],
        formulaCellCount: 0,
        nonEmptyCellCount: 0,
        formulaSamples: [],
        readback: [],
        error: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
      }),
    )
    return 1
  }
}

export function parseFormulaClinicCliArgs(args: readonly string[]): FormulaClinicCliOptions {
  let cells = ''
  let evaluationTimeoutMs = 30_000
  let filePath: string | undefined
  let help = false
  let maxFormulaSamples = 12

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      continue
    }
    if (arg === '--help' || arg === '-h') {
      help = true
      continue
    }
    if (arg === '--cells') {
      cells = readRequiredOptionValue(args, index, '--cells')
      index += 1
      continue
    }
    if (arg === '--formula-samples') {
      maxFormulaSamples = readPositiveIntegerOptionValue(args, index, '--formula-samples')
      index += 1
      continue
    }
    if (arg === '--timeout-ms') {
      evaluationTimeoutMs = readPositiveIntegerOptionValue(args, index, '--timeout-ms')
      index += 1
      continue
    }
    if (arg.startsWith('-')) {
      throw new Error(`Unknown bilig-formula-clinic argument: ${arg}`)
    }
    if (filePath !== undefined) {
      throw new Error(`Unexpected extra workbook path: ${arg}`)
    }
    filePath = arg
  }

  const parsedCells = cells
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map(parseCellRef)

  if (filePath === undefined) {
    return { cells: parsedCells, evaluationTimeoutMs, help, maxFormulaSamples }
  }
  return { cells: parsedCells, evaluationTimeoutMs, filePath, help, maxFormulaSamples }
}

export function formulaClinicHelpText(): string {
  return [
    'Usage: bilig-formula-clinic ./workbook.xlsx --cells "Summary!B7,Inputs!B2"',
    '',
    'Options:',
    '  --cells <refs>              Comma-separated cells to read after import.',
    '  --formula-samples <count>   Formula examples to include. Default: 12.',
    '  --timeout-ms <ms>           WorkPaper evaluation timeout. Default: 30000.',
    '  -h, --help                  Print this help text.',
    '',
    'The command reads the workbook locally and prints a Markdown issue body.',
    'It does not upload files or send workbook contents anywhere.',
    '',
  ].join('\n')
}

export function renderFormulaClinicReport(report: FormulaClinicReport): string {
  return `# Bilig formula clinic report

## Summary

- Package: \`@bilig/headless@${report.packageVersion}\`
- File: \`${report.fileName}\`
- File size: ${report.fileSizeBytes.toString()} bytes
- Status: ${report.status}
- Sheets: ${report.sheetNames.length > 0 ? report.sheetNames.map((sheet) => `\`${sheet}\``).join(', ') : 'not imported'}
- Non-empty imported cells: ${report.nonEmptyCellCount.toString()}
- Formula cells: ${report.formulaCellCount.toString()}

## Import warnings

${report.warningMessages.length > 0 ? report.warningMessages.map((warning) => `- ${warning}`).join('\n') : '- None reported'}

## Formula samples

${
  report.formulaSamples.length > 0
    ? report.formulaSamples.map((sample) => `- \`${sample.sheet}!${sample.cell}\`: \`${sample.formula}\``).join('\n')
    : '- No formula samples captured'
}

## Requested readback

${report.readback.length > 0 ? report.readback.map((line) => `- ${line}`).join('\n') : '- No cells requested. Re-run with `--cells "Sheet1!A1,Summary!B7"` if readback matters.'}

## Failure

${report.error === null ? '- None' : `- ${report.error}`}

## Expected behavior

Paste the value you expected from Excel, LibreOffice, Microsoft Graph, an existing service, or a manual check.

## Actual blocker

Paste what is wrong: stale cached value, wrong shared-formula expansion, unsupported formula, import warning, timeout, restore mismatch, or missing API.

## Public fixture permission

- [ ] This reduced case is public, anonymized, and safe to add to tests, examples, docs, or corpus files.
- [ ] The expected value comes from Excel, another trusted system, or a manual check.
- [ ] The report can be reproduced without a private workbook or account.

Privacy note: do not paste confidential workbook contents, customer data, financial models, or files that cannot be redistributed.
`
}

function readRequestedCell(
  workbook: ReturnType<typeof WorkPaper.buildFromSnapshot>,
  sheetNames: readonly string[],
  cell: FormulaClinicCellRef,
): string {
  const sheetName = cell.sheetName ?? sheetNames[0]
  if (sheetName === undefined) {
    return `\`${cell.a1}\`: no sheets were imported`
  }
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    return `\`${sheetName}!${cell.a1}\`: sheet was not imported`
  }
  const address = parseA1Address(cell.a1)
  if (address === null) {
    return `\`${sheetName}!${cell.a1}\`: invalid A1 address`
  }
  const value = workbook.getCellDisplayValue({ sheet: sheetId, row: address.row, col: address.col })
  const formula = workbook.getCellFormula({ sheet: sheetId, row: address.row, col: address.col })
  return `\`${sheetName}!${cell.a1}\`: value \`${value}\`${formula ? `, formula \`${formula}\`` : ''}`
}

function collectFormulaSamples(
  sheets: readonly { readonly name: string; readonly cells: readonly { readonly address: string; readonly formula?: string }[] }[],
  sampleLimit: number,
): FormulaClinicFormulaSample[] {
  const samples: FormulaClinicFormulaSample[] = []
  for (const sheet of sheets) {
    for (const cell of sheet.cells) {
      if (typeof cell.formula !== 'string') {
        continue
      }
      samples.push({ sheet: sheet.name, cell: cell.address, formula: cell.formula })
      if (samples.length >= sampleLimit) {
        return samples
      }
    }
  }
  return samples
}

function parseCellRef(input: string): FormulaClinicCellRef {
  const bangIndex = input.lastIndexOf('!')
  if (bangIndex < 0) {
    return { sheetName: null, a1: input }
  }
  return {
    sheetName: unquoteSheetName(input.slice(0, bangIndex)),
    a1: input.slice(bangIndex + 1),
  }
}

function unquoteSheetName(input: string): string {
  const trimmed = input.trim()
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'")
  }
  return trimmed
}

function parseA1Address(a1: string): { readonly row: number; readonly col: number } | null {
  const match = /^([A-Z]+)([1-9][0-9]*)$/i.exec(a1.trim())
  if (match === null) {
    return null
  }
  const letters = match[1]?.toUpperCase() ?? ''
  const row = Number(match[2]) - 1
  let col = 0
  for (const letter of letters) {
    col = col * 26 + (letter.charCodeAt(0) - 64)
  }
  return { row, col: col - 1 }
}

function readRequiredOptionValue(args: readonly string[], index: number, optionName: string): string {
  const value = args[index + 1]
  if (value === undefined || value.trim().length === 0 || value.startsWith('-')) {
    throw new Error(`${optionName} requires a value`)
  }
  return value
}

function readPositiveIntegerOptionValue(args: readonly string[], index: number, optionName: string): number {
  const raw = readRequiredOptionValue(args, index, optionName)
  const parsed = Number(raw)
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new Error(`${optionName} requires a positive integer`)
  }
  return parsed
}

function readHeadlessPackageVersion(): string {
  try {
    const entrypointPath = fileURLToPath(import.meta.resolve('@bilig/headless'))
    const packageJson: unknown = JSON.parse(readFileSync(findPackageJsonPath(entrypointPath), 'utf8'))
    if (typeof packageJson === 'object' && packageJson !== null && 'version' in packageJson && typeof packageJson.version === 'string') {
      return packageJson.version
    }
    return 'unknown'
  } catch {
    return 'unknown'
  }
}

function findPackageJsonPath(entrypointPath: string): string {
  let current = dirname(entrypointPath)
  for (let index = 0; index < 8; index += 1) {
    const candidate = join(current, 'package.json')
    if (existsSync(candidate)) {
      return candidate
    }
    const parent = dirname(current)
    if (parent === current) {
      break
    }
    current = parent
  }
  throw new Error('Could not resolve @bilig/headless package.json')
}
