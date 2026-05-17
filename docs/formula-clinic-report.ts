import { basename, dirname, join } from 'node:path'
import { existsSync, readFileSync, statSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

import { WorkPaper } from '@bilig/headless'
import { importXlsx } from '@bilig/headless/xlsx'

type CellRef = {
  readonly sheetName: string | null
  readonly a1: string
}

type FormulaSample = {
  readonly sheet: string
  readonly cell: string
  readonly formula: string
}

const args = process.argv.slice(2)
const filePath = firstPositionalArg(args)
const requestedCells = readOption(args, '--cells')
  .split(',')
  .map((entry) => entry.trim())
  .filter(Boolean)
  .map(parseCellRef)
const maxFormulaSamples = readPositiveIntegerOption(args, '--formula-samples', 12)
const evaluationTimeoutMs = readPositiveIntegerOption(args, '--timeout-ms', 30_000)

if (args.includes('--help') || filePath === null) {
  printHelp()
  process.exit(filePath === null ? 1 : 0)
}

const packageVersion = readHeadlessPackageVersion()
const fileStats = statSync(filePath)
const fileName = basename(filePath)
const bytes = new Uint8Array(readFileSync(filePath))

try {
  const imported = importXlsx(bytes, fileName)
  const formulaSamples = collectFormulaSamples(imported.snapshot.sheets, maxFormulaSamples)
  const formulaCellCount = imported.snapshot.sheets.reduce(
    (count, sheet) => count + sheet.cells.filter((cell) => typeof cell.formula === 'string').length,
    0,
  )
  const nonEmptyCellCount = imported.snapshot.sheets.reduce((count, sheet) => count + sheet.cells.length, 0)
  const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
    evaluationTimeoutMs,
    useColumnIndex: true,
  })

  try {
    const readback = requestedCells.map((cell) => readRequestedCell(workbook, imported.sheetNames, cell))
    printReport({
      status: 'imported',
      packageVersion,
      fileName,
      fileSizeBytes: fileStats.size,
      sheetNames: imported.sheetNames,
      warningMessages: imported.warnings,
      formulaCellCount,
      nonEmptyCellCount,
      formulaSamples,
      readback,
      error: null,
    })
  } finally {
    workbook.dispose()
  }
} catch (error) {
  printReport({
    status: 'failed',
    packageVersion,
    fileName,
    fileSizeBytes: fileStats.size,
    sheetNames: [],
    warningMessages: [],
    formulaCellCount: 0,
    nonEmptyCellCount: 0,
    formulaSamples: [],
    readback: [],
    error: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
  })
  process.exitCode = 1
}

function printHelp(): void {
  console.log(`Bilig formula clinic report

Usage:
  npx tsx formula-clinic-report.ts ./workbook.xlsx --cells "Summary!B7,Inputs!B2"

Options:
  --cells <refs>              Comma-separated cells to read after import.
  --formula-samples <count>   Formula examples to include. Default: 12.
  --timeout-ms <ms>           WorkPaper evaluation timeout. Default: 30000.

This script reads the workbook locally and prints a Markdown issue body.
It does not upload files or send workbook contents anywhere.`)
}

function printReport(report: {
  readonly status: 'imported' | 'failed'
  readonly packageVersion: string
  readonly fileName: string
  readonly fileSizeBytes: number
  readonly sheetNames: readonly string[]
  readonly warningMessages: readonly string[]
  readonly formulaCellCount: number
  readonly nonEmptyCellCount: number
  readonly formulaSamples: readonly FormulaSample[]
  readonly readback: readonly string[]
  readonly error: string | null
}): void {
  console.log(`# Bilig formula clinic report

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

Privacy note: do not paste confidential workbook contents, customer data, financial models, or files that cannot be redistributed.`)
}

function readRequestedCell(workbook: ReturnType<typeof WorkPaper.buildFromSnapshot>, sheetNames: readonly string[], cell: CellRef): string {
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
  const value = workbook.getCellValue({ sheet: sheetId, row: address.row, col: address.col })
  const formula = workbook.getCellFormula({ sheet: sheetId, row: address.row, col: address.col })
  return `\`${sheetName}!${cell.a1}\`: value ${formatCellValue(value)}${formula ? `, formula \`${formula}\`` : ''}`
}

function collectFormulaSamples(
  sheets: readonly { readonly name: string; readonly cells: readonly { readonly address: string; readonly formula?: string }[] }[],
  sampleLimit: number,
): FormulaSample[] {
  const samples: FormulaSample[] = []
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

function formatCellValue(value: unknown): string {
  if (typeof value === 'object' && value !== null && 'tag' in value) {
    return `\`${JSON.stringify(value)}\``
  }
  return `\`${String(value)}\``
}

function parseCellRef(input: string): CellRef {
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

function readOption(argv: readonly string[], optionName: string): string {
  const index = argv.indexOf(optionName)
  if (index < 0) {
    return ''
  }
  return argv[index + 1] ?? ''
}

function readPositiveIntegerOption(argv: readonly string[], optionName: string, fallback: number): number {
  const raw = readOption(argv, optionName)
  if (raw === '') {
    return fallback
  }
  const parsed = Number(raw)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback
}

function firstPositionalArg(argv: readonly string[]): string | null {
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    if (arg === undefined) {
      continue
    }
    if (arg.startsWith('--')) {
      index += 1
      continue
    }
    return arg
  }
  return null
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
