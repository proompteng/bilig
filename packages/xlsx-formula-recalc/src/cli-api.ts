import { readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'

import { ValueTag, type LiteralInput } from '@bilig/protocol'
import {
  assertXlsxExternalWorkbookByteInputWithinLimit,
  inspectXlsxCacheFileStreamingNative,
  replaceXlsxWorksheetCellXml,
  StreamingNativeXlsxRecalcError,
  type XlsxExternalWorkbookInput,
  writeSimpleXlsxWorkbook,
} from '@bilig/xlsx'

import { recalculateXlsxFileToFile } from './file-recalc.js'
import {
  defaultGithubActionPackageVersionForHelp,
  parseGithubActionWorkflowArgs,
  printGithubActionWorkflow,
  type CliInspectLimit,
} from './cli-github-action.js'
import type {
  XlsxFormulaRecalcCellValue,
  XlsxFormulaRecalcEdit,
  XlsxFormulaRecalcEngine,
  XlsxFormulaRecalcFallbackPolicy,
} from './types.js'

interface CliExternalWorkbook {
  readonly path: string
  readonly target?: string
}

interface CliOptions {
  readonly mode: 'file' | 'demo'
  readonly inputPath: string | undefined
  readonly outputPath: string
  readonly edits: readonly XlsxFormulaRecalcEdit[]
  readonly reads: readonly string[]
  readonly externalWorkbooks: readonly CliExternalWorkbook[]
  readonly inspect: boolean
  readonly inspectLimit: CliInspectLimit
  readonly engine?: XlsxFormulaRecalcEngine
  readonly maxRssBytes?: number
  readonly fallbackPolicy?: XlsxFormulaRecalcFallbackPolicy
  readonly json: boolean
}

export interface XlsxFormulaRecalcCliContext {
  readonly commandName?: string
  readonly stdout?: (text: string) => void
  readonly stderr?: (text: string) => void
}

const defaultInspectFormulaLimit = 2000
const cacheDoctorCommandName = 'xlsx-cache-doctor'
const printGithubActionOption = '--print-github-action'
const defaultGithubActionPackageVersion = defaultGithubActionPackageVersionForHelp()
const githubActionParseHelpers = {
  defaultInspectFormulaLimit,
  parseInspectLimit,
  requireNextArg,
}

export function runXlsxFormulaRecalcCli(args: readonly string[], context: XlsxFormulaRecalcCliContext = {}): number {
  const commandName = context.commandName ?? 'xlsx-recalc'
  const writeStdout = context.stdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = context.stderr ?? ((text: string) => process.stderr.write(text))

  try {
    if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
      printHelp(commandName, writeStdout)
      return 0
    }

    if (commandName === cacheDoctorCommandName && args.includes(printGithubActionOption)) {
      printGithubActionWorkflow(parseGithubActionWorkflowArgs(args, commandName, githubActionParseHelpers), writeStdout)
      return 0
    }

    const options = parseCliArgs(normalizeCliArgsForCommand(args, commandName), commandName)
    const inputName =
      options.mode === 'demo'
        ? commandName === cacheDoctorCommandName
          ? 'bilig-cache-doctor-stale-demo.xlsx'
          : 'bilig-formula-recalc-demo.xlsx'
        : basename(requireInputPath(options))
    const externalWorkbooks = readExternalWorkbookInputs(options.externalWorkbooks)
    if (options.inspect) {
      if (options.mode === 'file') {
        throw new Error(
          'File-mode XLSX inspection requires runXlsxFormulaRecalcCliAsync so the file-backed streaming-native inspector is used.',
        )
      }
      const input = inputBytesForCli(options, commandName)
      printInspectionSummary({ input, inputName, externalWorkbooks, options, writeStdout })
      return 0
    }
    if (options.mode === 'file') {
      throw new Error(
        'File-mode XLSX recalculation requires runXlsxFormulaRecalcCliAsync so the file-backed streaming-native engine is used.',
      )
    }
    const result = recalculateDemoXlsxToFile(options)

    const summary = {
      mode: options.mode,
      input: options.inputPath ?? 'generated demo workbook',
      output: options.outputPath,
      edits: options.edits.length,
      externalWorkbooks: externalWorkbooks.length,
      reads: result.reads,
      warnings: result.warnings,
      ...(result.diagnostics ? { diagnostics: result.diagnostics } : {}),
      commandSucceeded: true,
      recalculationCompleted: true,
      excelParity: 'not_proven',
      ...(options.mode === 'demo'
        ? {
            expectedReadback: { 'Summary!B2': 72_000 },
            expectedValueMatched: numericReadValue(result.reads['Summary!B2']) === 72_000,
          }
        : {}),
    }

    if (options.json) {
      writeStdout(`${JSON.stringify(summary, null, 2)}\n`)
    } else {
      writeStdout(`Recalculated ${summary.input} -> ${options.outputPath}\n`)
      for (const [target, value] of Object.entries(result.reads)) {
        writeStdout(`${target}: ${JSON.stringify(value)}\n`)
      }
      if (result.warnings.length > 0) {
        writeStdout(`Warnings: ${result.warnings.length.toString()}\n`)
      }
    }
    return 0
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    return 1
  }
}

export async function runXlsxFormulaRecalcCliAsync(args: readonly string[], context: XlsxFormulaRecalcCliContext = {}): Promise<number> {
  const commandName = context.commandName ?? 'xlsx-recalc'
  const writeStdout = context.stdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = context.stderr ?? ((text: string) => process.stderr.write(text))

  try {
    if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
      printHelp(commandName, writeStdout)
      return 0
    }

    if (commandName === cacheDoctorCommandName && args.includes(printGithubActionOption)) {
      printGithubActionWorkflow(parseGithubActionWorkflowArgs(args, commandName, githubActionParseHelpers), writeStdout)
      return 0
    }

    const options = parseCliArgs(normalizeCliArgsForCommand(args, commandName), commandName)
    const inputName =
      options.mode === 'demo'
        ? commandName === cacheDoctorCommandName
          ? 'bilig-cache-doctor-stale-demo.xlsx'
          : 'bilig-formula-recalc-demo.xlsx'
        : basename(requireInputPath(options))
    const externalWorkbooks = readExternalWorkbookInputs(options.externalWorkbooks)
    if (options.inspect) {
      await printInspectionSummaryAsync({ inputName, externalWorkbooks, options, writeStdout, commandName })
      return 0
    }
    const result =
      options.mode === 'file'
        ? await recalculateXlsxFileToFile(requireInputPath(options), {
            fileName: inputName,
            ...(externalWorkbooks.length > 0 ? { externalWorkbooks } : {}),
            edits: options.edits,
            reads: options.reads,
            ...(options.engine === undefined ? { engine: 'auto' as const } : { engine: options.engine }),
            ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
            ...(options.fallbackPolicy === undefined ? {} : { fallbackPolicy: options.fallbackPolicy }),
            outputPath: options.outputPath,
          })
        : recalculateDemoXlsxToFile(options)

    const summary = {
      mode: options.mode,
      input: options.inputPath ?? 'generated demo workbook',
      output: options.outputPath,
      edits: options.edits.length,
      externalWorkbooks: externalWorkbooks.length,
      reads: result.reads,
      warnings: result.warnings,
      ...(result.diagnostics ? { diagnostics: result.diagnostics } : {}),
      commandSucceeded: true,
      recalculationCompleted: true,
      excelParity: 'not_proven',
      ...(options.mode === 'demo'
        ? {
            expectedReadback: { 'Summary!B2': 72_000 },
            expectedValueMatched: numericReadValue(result.reads['Summary!B2']) === 72_000,
          }
        : {}),
    }

    if (options.json) {
      writeStdout(`${JSON.stringify(summary, null, 2)}\n`)
    } else {
      writeStdout(`Recalculated ${summary.input} -> ${options.outputPath}\n`)
      for (const [target, value] of Object.entries(result.reads)) {
        writeStdout(`${target}: ${JSON.stringify(value)}\n`)
      }
      if (result.warnings.length > 0) {
        writeStdout(`Warnings: ${result.warnings.length.toString()}\n`)
      }
    }
    return 0
  } catch (error) {
    if (args.includes('--json')) {
      writeStdout(`${JSON.stringify(cliErrorSummary(error), null, 2)}\n`)
    } else {
      writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
    }
    return 1
  }
}

function normalizeCliArgsForCommand(args: readonly string[], commandName: string): readonly string[] {
  if (commandName !== cacheDoctorCommandName || args.includes('--inspect') || hasExplicitRecalcIntent(args)) {
    return args
  }
  return [...args, '--inspect']
}

function hasExplicitRecalcIntent(args: readonly string[]): boolean {
  return args.includes('--read') || args.includes('--out') || args.includes('-o')
}

function parseCliArgs(args: readonly string[], commandName: string): CliOptions {
  const demo = args.includes('--demo')
  const inputPath = demo ? undefined : args[0]
  if (!demo && (!inputPath || inputPath.startsWith('-'))) {
    throw new Error('Expected input XLSX path or --demo')
  }

  const edits: XlsxFormulaRecalcEdit[] = []
  const reads: string[] = []
  const externalWorkbooks: CliExternalWorkbook[] = []
  let outputPath: string | undefined
  let inspect = false
  let inspectLimit: CliOptions['inspectLimit'] = defaultInspectFormulaLimit
  let engine: XlsxFormulaRecalcEngine | undefined
  let maxRssBytes: number | undefined
  let fallbackPolicy: XlsxFormulaRecalcFallbackPolicy | undefined
  let json = false

  for (let index = demo ? 0 : 1; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error(`Unexpected missing ${commandName} argument`)
    }
    switch (arg) {
      case '--demo':
        break
      case '--set':
        edits.push(parseEdit(requireNextArg(args, index, '--set')))
        index += 1
        break
      case '--read':
        reads.push(requireNextArg(args, index, '--read'))
        index += 1
        break
      case '--external-workbook':
        externalWorkbooks.push({ path: requireNextArg(args, index, '--external-workbook') })
        index += 1
        break
      case '--external-workbook-target':
        externalWorkbooks.push({
          path: requireNextArg(args, index, '--external-workbook-target'),
          target: requireNextArg(args, index + 1, '--external-workbook-target target'),
        })
        index += 2
        break
      case '--inspect':
        inspect = true
        break
      case '--inspect-limit':
        inspectLimit = parseInspectLimit(requireNextArg(args, index, '--inspect-limit'))
        index += 1
        break
      case '--timeout-ms':
        throw unsupportedTimeoutOptionError()
      case '--engine':
        engine = parseEngineOption(requireNextArg(args, index, '--engine'))
        index += 1
        break
      case '--fallback-policy':
        fallbackPolicy = parseFallbackPolicyOption(requireNextArg(args, index, '--fallback-policy'))
        index += 1
        break
      case '--max-rss-bytes':
        maxRssBytes = parsePositiveIntegerOption(requireNextArg(args, index, '--max-rss-bytes'), '--max-rss-bytes')
        index += 1
        break
      case '--max-rss-mb':
        maxRssBytes = parsePositiveIntegerOption(requireNextArg(args, index, '--max-rss-mb'), '--max-rss-mb') * 1024 * 1024
        index += 1
        break
      case '--out':
      case '-o':
        outputPath = requireNextArg(args, index, arg)
        index += 1
        break
      case '--json':
        json = true
        break
      default:
        throw new Error(`Unknown ${commandName} option: ${arg}`)
    }
  }

  return {
    mode: demo ? 'demo' : 'file',
    inputPath,
    outputPath:
      outputPath ?? (demo ? 'bilig-formula-recalc-demo.xlsx' : defaultOutputPath(requireDefined(inputPath, 'Expected input XLSX path'))),
    edits: edits.length > 0 ? edits : demoDefaultEdits(demo),
    reads: reads.length > 0 ? reads : demoDefaultReads(demo),
    externalWorkbooks,
    inspect,
    inspectLimit,
    ...(engine === undefined ? {} : { engine }),
    ...(maxRssBytes === undefined ? {} : { maxRssBytes }),
    ...(fallbackPolicy === undefined ? {} : { fallbackPolicy }),
    json,
  }
}

function readExternalWorkbookInputs(workbooks: readonly CliExternalWorkbook[]): XlsxExternalWorkbookInput[] {
  return workbooks.map((workbook) => {
    const byteLength = statSync(workbook.path).size
    assertXlsxExternalWorkbookByteInputWithinLimit(byteLength, workbook.path)
    return {
      bytes: readFileSync(workbook.path),
      fileName: basename(workbook.path),
      ...(workbook.target ? { target: workbook.target } : {}),
    }
  })
}

interface PrintInspectionSummaryInput {
  readonly input: Uint8Array | Buffer
  readonly inputName: string
  readonly externalWorkbooks: readonly XlsxExternalWorkbookInput[]
  readonly options: CliOptions
  readonly writeStdout: (text: string) => void
}

interface PrintInspectionSummaryAsyncInput {
  readonly inputName: string
  readonly externalWorkbooks: readonly XlsxExternalWorkbookInput[]
  readonly options: CliOptions
  readonly writeStdout: (text: string) => void
  readonly commandName: string
}

interface PrintableInspectionSummary {
  readonly schemaVersion: string
  readonly sheetNames: readonly string[]
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly uninspectedFormulaCellCount: number
  readonly inspectionLimit: CliOptions['inspectLimit']
  readonly staleCachedFormulaCount: number
  readonly cacheStatusSummary: {
    readonly inspected: number
    readonly stale: number
    readonly fresh: number
    readonly missingCache: number
    readonly unsupportedRecalculation: number
  }
  readonly suggestedReads: readonly string[]
  readonly formulas: readonly unknown[]
  readonly warnings: readonly string[]
  readonly diagnostics?: unknown
  readonly inspectionCompleted: true
  readonly recalculationCompleted: boolean
  readonly excelParity: 'not_proven'
}

async function printInspectionSummaryAsync(args: PrintInspectionSummaryAsyncInput): Promise<void> {
  if (args.options.mode === 'file') {
    if (args.externalWorkbooks.length > 0) {
      throw new Error('streaming-native inspection does not support external workbook companions')
    }
    const inspection = await inspectXlsxCacheFileStreamingNative(requireInputPath(args.options), {
      inspectLimit: args.options.inspectLimit,
      ...(args.options.maxRssBytes === undefined ? {} : { maxRssBytes: args.options.maxRssBytes }),
    })
    writeInspectionSummary(inspection, args)
    return
  }
  if (args.options.mode === 'demo') {
    printInspectionSummary({ ...args, input: inputBytesForCli(args.options, args.commandName) })
    return
  }
  throw new Error('Unexpected XLSX inspection mode')
}

function printInspectionSummary(args: PrintInspectionSummaryInput): void {
  const inspection = buildDemoInspectionSummary(args.options)
  writeInspectionSummary(inspection, args)
}

function writeInspectionSummary(
  inspection: PrintableInspectionSummary,
  args: Pick<PrintInspectionSummaryInput, 'options' | 'writeStdout' | 'externalWorkbooks'>,
): void {
  const summary = {
    schemaVersion: inspection.schemaVersion,
    mode: args.options.mode,
    input: args.options.inputPath ?? 'generated demo workbook',
    edits: args.options.edits.length,
    externalWorkbooks: args.externalWorkbooks.length,
    sheetNames: inspection.sheetNames,
    formulaCellCount: inspection.formulaCellCount,
    inspectedFormulaCellCount: inspection.inspectedFormulaCellCount,
    uninspectedFormulaCellCount: inspection.uninspectedFormulaCellCount,
    inspectionLimit: inspection.inspectionLimit,
    staleCachedFormulaCount: inspection.staleCachedFormulaCount,
    cacheStatusSummary: inspection.cacheStatusSummary,
    suggestedReads: inspection.suggestedReads,
    formulas: inspection.formulas,
    warnings: inspection.warnings,
    ...(inspection.diagnostics ? { diagnostics: inspection.diagnostics } : {}),
    commandSucceeded: true,
    inspectionCompleted: inspection.inspectionCompleted,
    recalculationCompleted: inspection.recalculationCompleted,
    excelParity: inspection.excelParity,
  }

  if (args.options.json) {
    args.writeStdout(`${JSON.stringify(summary, null, 2)}\n`)
    return
  }

  args.writeStdout(`Inspected ${summary.input}\n`)
  args.writeStdout(`Sheets: ${summary.sheetNames.join(', ')}\n`)
  args.writeStdout(`Formula cells: ${summary.formulaCellCount.toString()}\n`)
  args.writeStdout(`Inspected formula cells: ${summary.inspectedFormulaCellCount.toString()}\n`)
  args.writeStdout(`Uninspected formula cells: ${summary.uninspectedFormulaCellCount.toString()}\n`)
  args.writeStdout(`Stale cached formula cells: ${summary.staleCachedFormulaCount.toString()}\n`)
  args.writeStdout(`Fresh cached formula cells: ${summary.cacheStatusSummary.fresh.toString()}\n`)
  args.writeStdout(`Missing cached formula values: ${summary.cacheStatusSummary.missingCache.toString()}\n`)
  args.writeStdout(`Unsupported recalculation results: ${summary.cacheStatusSummary.unsupportedRecalculation.toString()}\n`)
  if (summary.suggestedReads.length > 0) {
    args.writeStdout(`Suggested reads: ${summary.suggestedReads.join(', ')}\n`)
    args.writeStdout(`${suggestedRecalcCommand(args.options, summary.suggestedReads)}\n`)
  }
  if (summary.warnings.length > 0) {
    args.writeStdout(`Warnings: ${summary.warnings.length.toString()}\n`)
  }
  args.writeStdout('If the important formula is missing, rerun with --inspect-limit all and --read <Sheet!Cell>.\n')
}

function numericReadValue(value: XlsxFormulaRecalcCellValue | undefined): number | undefined {
  return value !== undefined &&
    typeof value === 'object' &&
    value !== null &&
    'tag' in value &&
    value.tag === ValueTag.Number &&
    'value' in value
    ? value.value
    : undefined
}

function suggestedRecalcCommand(options: CliOptions, reads: readonly string[]): string {
  const command = ['xlsx-recalc']
  if (options.mode === 'demo') {
    command.push('--demo')
  } else if (options.inputPath) {
    command.push(shellQuote(options.inputPath))
  }
  for (const edit of options.edits) {
    command.push('--set', shellQuote(`${edit.target}=${String(edit.value)}`))
  }
  for (const read of reads.slice(0, 5)) {
    command.push('--read', shellQuote(read))
  }
  command.push('--json')
  return command.join(' ')
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function inputBytesForCli(options: CliOptions, commandName: string): Uint8Array {
  if (options.mode !== 'demo') {
    throw new Error('File-mode XLSX byte input is disabled; use runXlsxFormulaRecalcCliAsync for file-backed native XLSX work.')
  }
  return commandName === cacheDoctorCommandName && options.inspect ? buildStaleCacheDoctorDemoWorkbookBytes() : buildDemoWorkbookBytes()
}

function buildDemoWorkbookBytes(values: { readonly units?: number; readonly price?: number; readonly revenue?: number } = {}): Uint8Array {
  const units = values.units ?? 40
  const price = values.price ?? 1200
  const revenue = values.revenue ?? units * price
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Inputs',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'Metric' },
          { address: 'B1', row: 0, col: 1, value: 'Value' },
          { address: 'A2', row: 1, col: 0, value: 'Units' },
          { address: 'B2', row: 1, col: 1, value: units },
          { address: 'A3', row: 2, col: 0, value: 'Price' },
          { address: 'B3', row: 2, col: 1, value: price },
        ],
      },
      {
        name: 'Summary',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'Metric' },
          { address: 'B1', row: 0, col: 1, value: 'Value' },
          { address: 'A2', row: 1, col: 0, value: 'Revenue' },
          { address: 'B2', row: 1, col: 1, formula: 'Inputs!B2*Inputs!B3', value: revenue },
        ],
      },
    ],
  })
}

function buildStaleCacheDoctorDemoWorkbookBytes(): Uint8Array {
  return replaceWorksheetCellXml(
    buildDemoWorkbookBytes(),
    'xl/worksheets/sheet2.xml',
    'B2',
    '<c r="B2"><f>Inputs!B2*Inputs!B3</f><v>60000</v></c>',
  )
}

function replaceWorksheetCellXml(bytes: Uint8Array, path: string, address: string, replacement: string): Uint8Array {
  return replaceXlsxWorksheetCellXml(bytes, {
    path,
    address,
    replacement,
    missingMessage: `Demo XLSX is missing ${path} ${address}`,
  })
}

function recalculateDemoXlsxToFile(options: CliOptions): {
  readonly bytesWritten: number
  readonly warnings: readonly []
  readonly sheetNames: readonly ['Inputs', 'Summary']
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly []
  readonly diagnostics?: undefined
} {
  const values = demoWorkbookValuesAfterEdits(options.edits)
  const xlsx = buildDemoWorkbookBytes(values)
  writeFileSync(options.outputPath, xlsx)
  return {
    bytesWritten: xlsx.byteLength,
    warnings: [],
    sheetNames: ['Inputs', 'Summary'],
    reads: Object.fromEntries(options.reads.map((read) => [read, demoReadValue(read, values)])),
    changes: [],
  }
}

function buildDemoInspectionSummary(options: CliOptions): PrintableInspectionSummary {
  const values = demoWorkbookValuesAfterEdits(options.edits)
  const inspectionLimit = options.inspectLimit
  const inspectedFormulaCellCount = inspectionLimit === 'all' || inspectionLimit > 0 ? 1 : 0
  const formula =
    inspectedFormulaCellCount > 0
      ? [
          {
            target: 'Summary!B2',
            formula: '=Inputs!B2*Inputs!B3',
            cachedValue: 60_000,
            literalRecalculatedValue: values.revenue,
            cacheStatus: values.revenue === 60_000 ? 'fresh' : 'stale',
            staleCachedValue: values.revenue !== 60_000,
          },
        ]
      : []
  return {
    schemaVersion: 'xlsx-cache-doctor.v1',
    sheetNames: ['Inputs', 'Summary'],
    formulaCellCount: 1,
    inspectedFormulaCellCount,
    uninspectedFormulaCellCount: 1 - inspectedFormulaCellCount,
    inspectionLimit,
    staleCachedFormulaCount: values.revenue === 60_000 || inspectedFormulaCellCount === 0 ? 0 : 1,
    cacheStatusSummary: {
      inspected: inspectedFormulaCellCount,
      stale: values.revenue === 60_000 || inspectedFormulaCellCount === 0 ? 0 : 1,
      fresh: values.revenue === 60_000 && inspectedFormulaCellCount > 0 ? 1 : 0,
      missingCache: 0,
      unsupportedRecalculation: 0,
    },
    suggestedReads: inspectedFormulaCellCount > 0 ? ['Summary!B2'] : [],
    formulas: formula,
    warnings: [],
    inspectionCompleted: true,
    recalculationCompleted: true,
    excelParity: 'not_proven',
  }
}

function demoWorkbookValuesAfterEdits(edits: readonly XlsxFormulaRecalcEdit[]): {
  readonly units: number
  readonly price: number
  readonly revenue: number
} {
  let units = 40
  let price = 1200
  for (const edit of edits) {
    const target = normalizeDemoTarget(edit.target)
    if (target === 'Inputs!B2' && typeof edit.value === 'number') {
      units = edit.value
    }
    if (target === 'Inputs!B3' && typeof edit.value === 'number') {
      price = edit.value
    }
  }
  return { units, price, revenue: units * price }
}

function normalizeDemoTarget(target: string): string {
  return target.replace(/^'Inputs'!/u, 'Inputs!').replace(/^'Summary'!/u, 'Summary!')
}

function demoReadValue(
  target: string,
  values: { readonly units: number; readonly price: number; readonly revenue: number },
): XlsxFormulaRecalcCellValue {
  switch (normalizeDemoTarget(target)) {
    case 'Inputs!B2':
      return { tag: ValueTag.Number, value: values.units }
    case 'Inputs!B3':
      return { tag: ValueTag.Number, value: values.price }
    case 'Summary!B2':
      return { tag: ValueTag.Number, value: values.revenue }
    default:
      return { tag: ValueTag.Empty }
  }
}

function demoDefaultEdits(enabled: boolean): readonly XlsxFormulaRecalcEdit[] {
  return enabled
    ? [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ]
    : []
}

function demoDefaultReads(enabled: boolean): readonly string[] {
  return enabled ? ['Summary!B2'] : []
}

function parseEdit(raw: string): XlsxFormulaRecalcEdit {
  const separator = raw.indexOf('=')
  if (separator <= 0) {
    throw new Error(`Expected --set value in Target=Value form, received: ${raw}`)
  }
  return {
    target: raw.slice(0, separator),
    value: parseRawCellContent(raw.slice(separator + 1)),
  }
}

function parseRawCellContent(raw: string): LiteralInput {
  if (raw === 'null') {
    return null
  }
  if (raw === 'true') {
    return true
  }
  if (raw === 'false') {
    return false
  }
  if (/^-?(?:\d+|\d*\.\d+)(?:e[+-]?\d+)?$/iu.test(raw)) {
    return Number(raw)
  }
  return raw
}

function parseInspectLimit(raw: string): CliOptions['inspectLimit'] {
  if (raw === 'all') {
    return raw
  }
  const value = Number(raw)
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`Expected --inspect-limit to be "all" or a positive integer, received: ${raw}`)
  }
  return value
}

function parsePositiveIntegerOption(raw: string, option: string): number {
  const value = Number(raw)
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`Expected ${option} to be a positive integer, received: ${raw}`)
  }
  return value
}

function parseEngineOption(raw: string): XlsxFormulaRecalcEngine {
  if (raw === 'auto' || raw === 'streaming-native') {
    return raw
  }
  if (raw === 'workpaper') {
    throw legacyWorkPaperCliPathError('engine')
  }
  throw new Error(`Expected --engine to be "auto" or "streaming-native", received: ${raw}`)
}

function parseFallbackPolicyOption(raw: string): XlsxFormulaRecalcFallbackPolicy {
  if (raw === 'error') {
    return raw
  }
  if (raw === 'workpaper') {
    throw legacyWorkPaperCliPathError('fallbackPolicy')
  }
  throw new Error(`Expected --fallback-policy to be "error", received: ${raw}`)
}

function requireNextArg(args: readonly string[], index: number, option: string): string {
  const value = args[index + 1]
  if (!value || value.startsWith('--')) {
    throw new Error(`Expected value after ${option}`)
  }
  return value
}

function requireInputPath(options: CliOptions): string {
  return requireDefined(options.inputPath, 'Expected input XLSX path')
}

function requireDefined(value: string | undefined, message: string): string {
  if (!value) {
    throw new Error(message)
  }
  return value
}

function defaultOutputPath(inputPath: string): string {
  const extension = extname(inputPath)
  const base = extension.length > 0 ? basename(inputPath, extension) : basename(inputPath)
  return join(dirname(inputPath), `${base}.recalculated${extension || '.xlsx'}`)
}

function printHelp(commandName: string, writeStdout: (text: string) => void): void {
  if (commandName === cacheDoctorCommandName) {
    writeStdout(`Usage: ${commandName} <input.xlsx> [options]
       ${commandName} --demo [--json]

Diagnose stale cached formula values without writing an output XLSX. This is
the memorable alias for: xlsx-recalc <input.xlsx> --inspect.

Options:
  --demo                  Generate a tiny workbook and inspect its formula cache.
  ${printGithubActionOption} <workbook-glob>
                          Print a ready-to-commit GitHub Actions workflow that uses proompteng/bilig@v1.
  --set <Sheet!A1=value>  Edit an input cell before diagnosis. Repeatable.
  --inspect-limit <all|n> Formula cells to recompute during inspection. Defaults to ${defaultInspectFormulaLimit}.
  --engine <auto|streaming-native>
                          Select the recalculation engine. File-to-file CLI defaults to auto, which tries streaming-native.
  --fallback-policy <error>
                          Decide whether unsupported streaming-native jobs fail. WorkPaper bytes APIs live under @bilig/workpaper/xlsx.
  --max-rss-mb <n>        Fail streaming-native if resident memory exceeds this cap.
  --fail-on-stale <true|false>
                          With ${printGithubActionOption}, decide whether the generated workflow fails pull requests. Defaults to false.
  --changed-files-only <true|false>
                          With ${printGithubActionOption}, inspect only changed XLSX files. Defaults to true.
  --json-output <path>    With ${printGithubActionOption}, set the JSON report path. Defaults to \${{ runner.temp }}/xlsx-cache-doctor.json.
  --markdown-output <path>
                          With ${printGithubActionOption}, set the Markdown report path. Defaults to \${{ runner.temp }}/xlsx-cache-doctor.md.
  --package-version <version>
                          With ${printGithubActionOption}, pin @bilig/xlsx-formula-recalc in the generated workflow. Defaults to ${defaultGithubActionPackageVersion}.
  --workflow-name <name>  With ${printGithubActionOption}, set the generated workflow name.
  --external-workbook <path>
                          Supply a companion XLSX for external-link cache refresh. Repeatable.
  --external-workbook-target <path> <target>
                          Supply a companion XLSX for an exact Excel link target. Repeatable.
  --json                  Print a JSON summary.
  --help, -h              Show this help.
`)
    return
  }

  writeStdout(`Usage: ${commandName} <input.xlsx> [options]
       ${commandName} --demo [--json] [--out demo.recalculated.xlsx]

Options:
  --demo                  Generate a tiny workbook, edit inputs, recalculate, and write proof XLSX.
  --set <Sheet!A1=value>  Edit an input cell before recalculation. Repeatable.
  --read <Sheet!A1>       Read a recalculated cell after edits. Repeatable.
  --inspect               Inspect formula cells, stale cached values, and suggested --read targets.
  --inspect-limit <all|n> Formula cells to recompute during inspection. Defaults to ${defaultInspectFormulaLimit}.
  --engine <auto|streaming-native>
                          Select the recalculation engine. File-to-file CLI defaults to auto, which tries streaming-native.
  --fallback-policy <error>
                          Decide whether unsupported streaming-native jobs fail. WorkPaper bytes APIs live under @bilig/workpaper/xlsx.
  --max-rss-mb <n>        Fail streaming-native if resident memory exceeds this cap.
  --external-workbook <path>
                          Supply a companion XLSX for external-link cache refresh. Repeatable.
  --external-workbook-target <path> <target>
                          Supply a companion XLSX for an exact Excel link target. Repeatable.
  --out, -o <path>        Output XLSX path. Defaults to <input>.recalculated.xlsx.
  --json                  Print a JSON summary.
  --help, -h              Show this help.
`)
}

function cliErrorSummary(error: unknown): Record<string, unknown> {
  return {
    commandSucceeded: false,
    recalculationCompleted: false,
    error: error instanceof Error ? error.message : String(error),
    ...(error instanceof StreamingNativeXlsxRecalcError ? { diagnostics: error.diagnostics } : {}),
  }
}

function legacyWorkPaperCliPathError(optionName: 'engine' | 'fallbackPolicy'): Error {
  return new Error(
    `The primary xlsx-recalc file CLI no longer loads or exports WorkPaper through ${optionName}; use @bilig/workpaper for WorkPaper workflows.`,
  )
}

function unsupportedTimeoutOptionError(): Error {
  return new Error(
    'The streaming-native xlsx-recalc CLI does not support --timeout-ms; use @bilig/workpaper for WorkPaper evaluation timeout controls.',
  )
}
