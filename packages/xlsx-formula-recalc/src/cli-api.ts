import { readFileSync, writeFileSync } from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'

import type { RawCellContent } from '@bilig/headless'
import { exportXlsx, recalculateXlsx, WorkPaper, type XlsxFormulaRecalcEdit } from './index.js'

interface CliOptions {
  readonly mode: 'file' | 'demo'
  readonly inputPath: string | undefined
  readonly outputPath: string
  readonly edits: readonly XlsxFormulaRecalcEdit[]
  readonly reads: readonly string[]
  readonly json: boolean
}

export interface XlsxFormulaRecalcCliContext {
  readonly commandName?: string
  readonly stdout?: (text: string) => void
  readonly stderr?: (text: string) => void
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

    const options = parseCliArgs(args, commandName)
    const input = options.mode === 'demo' ? buildDemoWorkbookBytes() : readFileSync(requireInputPath(options))
    const inputName = options.mode === 'demo' ? 'bilig-formula-recalc-demo.xlsx' : basename(requireInputPath(options))
    const result = recalculateXlsx(input, {
      fileName: inputName,
      edits: options.edits,
      reads: options.reads,
    })
    writeFileSync(options.outputPath, result.xlsx)

    const summary = {
      mode: options.mode,
      input: options.inputPath ?? 'generated demo workbook',
      output: options.outputPath,
      edits: options.edits.length,
      reads: result.reads,
      warnings: result.warnings,
      verified: true,
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

function parseCliArgs(args: readonly string[], commandName: string): CliOptions {
  const demo = args.includes('--demo')
  const inputPath = demo ? undefined : args[0]
  if (!demo && (!inputPath || inputPath.startsWith('-'))) {
    throw new Error('Expected input XLSX path or --demo')
  }

  const edits: XlsxFormulaRecalcEdit[] = []
  const reads: string[] = []
  let outputPath: string | undefined
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
    json,
  }
}

function buildDemoWorkbookBytes(): Uint8Array {
  const sourceWorkbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Units', 40],
      ['Price', 1200],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Revenue', '=Inputs!B2*Inputs!B3'],
    ],
  })
  try {
    return exportXlsx(sourceWorkbook.exportSnapshot())
  } finally {
    sourceWorkbook.dispose()
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

function parseRawCellContent(raw: string): RawCellContent {
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
  writeStdout(`Usage: ${commandName} <input.xlsx> [options]
       ${commandName} --demo [--json] [--out demo.recalculated.xlsx]

Options:
  --demo                  Generate a tiny workbook, edit inputs, recalculate, and write proof XLSX.
  --set <Sheet!A1=value>  Edit an input cell before recalculation. Repeatable.
  --read <Sheet!A1>       Read a recalculated cell after edits. Repeatable.
  --out, -o <path>        Output XLSX path. Defaults to <input>.recalculated.xlsx.
  --json                  Print a JSON summary.
  --help, -h              Show this help.
`)
}
