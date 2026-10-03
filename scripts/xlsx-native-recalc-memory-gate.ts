#!/usr/bin/env bun

import { spawn } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

import { strToU8, zipSync } from 'fflate'

import { readFlagArg, readMegabytesArg, readNumberArg, readStringArg } from './xlsx-fixture-corpus-cli.ts'
import { formatByteSize, startChildRssWatchdog, terminateChildProcess } from './xlsx-fixture-corpus-process.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const mib = 1024 * 1024
const defaultCacheDir = join(rootDir, '.cache', 'xlsx-native-recalc-memory-gate')
const defaultCliPath = join(rootDir, 'packages', '@bilig/xlsx-formula-recalc', 'dist', 'cli.js')
const defaultSyntheticRowCount = 50_000
const defaultTimeoutMs = 180_000
const rssCheckIntervalMs = 10
const noop = (): void => undefined

export const nativeRecalcMemoryGateBudgets = {
  syntheticRowChainMaxRssBytes: 320 * mib,
  issue442MaxRssBytes: 350 * mib,
} as const

export interface NativeRecalcGateTarget {
  readonly id: string
  readonly label: string
  readonly inputPath: string
  readonly outputPath: string
  readonly maxRssBytes: number
  readonly edits: readonly { readonly target: string; readonly value: number | string }[]
  readonly reads: readonly string[]
  readonly expectedReads: Readonly<Record<string, number | string>>
}

export interface NativeRecalcMemoryGateResult {
  readonly id: string
  readonly label: string
  readonly status: 'passed' | 'failed' | 'skipped'
  readonly maxRssBytes: number
  readonly peakRssBytes: number | null
  readonly diagnosticsMaxObservedRssBytes?: number
  readonly fallbackUsed?: boolean
  readonly patchedCacheCount?: number
  readonly nativeKernelFormulaCellCount?: number
  readonly values?: Readonly<Record<string, unknown>>
  readonly reason?: string
}

interface ChildRunResult {
  readonly exitCode: number | null
  readonly signal: string | null
  readonly stdout: string
  readonly stderr: string
  readonly peakRssBytes: number | null
  readonly rssLimitExceededBytes?: number
  readonly timedOut: boolean
}

interface SyntheticWorkbookArtifact {
  readonly id: string
  readonly fileName: string
  readonly filePath: string
  readonly sha256: string
  readonly byteSize: number
  readonly editTarget: string
  readonly reads: readonly string[]
  readonly expectedReads: Readonly<Record<string, number>>
}

async function main(): Promise<void> {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const outputDir = resolve(readStringArg('--output-dir', join(cacheDir, 'outputs')))
  const cliPath = resolve(readStringArg('--cli', defaultCliPath))
  const nodeBin = readStringArg('--node-bin', process.env['BILIG_XLSX_NATIVE_RECALC_NODE'] ?? 'node')
  const syntheticRows = readNumberArg('--synthetic-rows', defaultSyntheticRowCount)
  const timeoutMs = readNumberArg('--timeout-ms', defaultTimeoutMs)
  const syntheticMaxRssBytes = readMegabytesArg('--synthetic-max-rss-mb', nativeRecalcMemoryGateBudgets.syntheticRowChainMaxRssBytes)
  const issue442MaxRssBytes = readMegabytesArg('--issue-442-max-rss-mb', nativeRecalcMemoryGateBudgets.issue442MaxRssBytes)
  const issue442Path = readStringArg('--issue-442-path', '')
  const syntheticIssue442 = readFlagArg('--synthetic-issue-442')
  const requireIssue442 = readFlagArg('--require-issue-442')
  const issue442Only = readFlagArg('--issue-442-only')
  mkdirSync(outputDir, { recursive: true })

  const results: NativeRecalcMemoryGateResult[] = []
  if (!existsSync(cliPath)) {
    results.push(failedSetupResult('xlsx-native-recalc-cli', 'xlsx-recalc dist CLI', syntheticMaxRssBytes, `CLI not found: ${cliPath}`))
  } else {
    const targets: NativeRecalcGateTarget[] = []
    if (!issue442Only) {
      const synthetic = writeSyntheticNativeRecalcWorkbook(cacheDir, { rowCount: syntheticRows })
      targets.push({
        id: synthetic.id,
        label: synthetic.fileName,
        inputPath: synthetic.filePath,
        outputPath: join(outputDir, 'synthetic-row-chain.native.xlsx'),
        maxRssBytes: syntheticMaxRssBytes,
        edits: [{ target: synthetic.editTarget, value: 16 }],
        reads: synthetic.reads,
        expectedReads: synthetic.expectedReads,
      })
    }
    if (syntheticIssue442) {
      const synthetic = writeSyntheticIssue442NativeRecalcWorkbook(cacheDir)
      targets.push({
        id: synthetic.id,
        label: synthetic.fileName,
        inputPath: synthetic.filePath,
        outputPath: join(outputDir, 'synthetic-issue-442.native.xlsx'),
        maxRssBytes: issue442MaxRssBytes,
        edits: [{ target: synthetic.editTarget, value: 16 }],
        reads: synthetic.reads,
        expectedReads: synthetic.expectedReads,
      })
    } else if (issue442Path) {
      targets.push({
        id: 'issue-442-ocha-native-recalc',
        label: 'ocha-operational-partners-presence-jan-sep-2024.xlsx',
        inputPath: resolve(issue442Path),
        outputPath: join(outputDir, 'issue-442-ocha.native.xlsx'),
        maxRssBytes: issue442MaxRssBytes,
        edits: [{ target: 'Data!R57152', value: 16 }],
        reads: ['Data!U57152', 'Data!V57152'],
        expectedReads: { 'Data!U57152': 168.75, 'Data!V57152': 28.125 },
      })
    } else if (requireIssue442) {
      results.push(
        failedSetupResult('issue-442-ocha-native-recalc', 'issue 442 OCHA workbook', issue442MaxRssBytes, 'missing --issue-442-path'),
      )
    } else {
      results.push(skippedResult('issue-442-ocha-native-recalc', 'issue 442 OCHA workbook', issue442MaxRssBytes, 'no --issue-442-path'))
    }

    for (const target of targets) {
      // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential targets keep host memory bounded and make RSS attribution clear.
      results.push(await runNativeRecalcGateTarget(target, cliPath, { nodeBin, timeoutMs }))
    }
  }

  const failed = results.filter((result) => result.status === 'failed')
  process.stdout.write(
    `${JSON.stringify(
      {
        mode: 'xlsx-native-recalc-memory-gate',
        targets: {
          syntheticRowChainMaxRssBytes: syntheticMaxRssBytes,
          issue442MaxRssBytes,
        },
        results,
      },
      null,
      2,
    )}\n`,
  )
  if (failed.length > 0) {
    throw new Error(
      `XLSX native recalc memory gate failed: ${failed.map((result) => `${result.id} ${result.reason ?? ''}`.trim()).join('; ')}`,
    )
  }
}

export function buildNativeRecalcCliArgs(target: NativeRecalcGateTarget): string[] {
  const args = [
    target.inputPath,
    '--out',
    target.outputPath,
    '--engine',
    'streaming-native',
    '--fallback-policy',
    'error',
    '--max-rss-bytes',
    String(target.maxRssBytes),
    '--json',
  ]
  for (const edit of target.edits) {
    args.push('--set', `${edit.target}=${String(edit.value)}`)
  }
  for (const read of target.reads) {
    args.push('--read', read)
  }
  return args
}

export async function runNativeRecalcGateTarget(
  target: NativeRecalcGateTarget,
  cliPath: string,
  options: { readonly nodeBin: string; readonly timeoutMs: number },
): Promise<NativeRecalcMemoryGateResult> {
  if (!existsSync(target.inputPath)) {
    return failedSetupResult(target.id, target.label, target.maxRssBytes, `input not found: ${target.inputPath}`)
  }
  const child = await runChildProcess(options.nodeBin, [cliPath, ...buildNativeRecalcCliArgs(target)], {
    maxRssBytes: target.maxRssBytes,
    timeoutMs: options.timeoutMs,
  })
  const peakRssBytes = child.peakRssBytes
  if (child.rssLimitExceededBytes !== undefined) {
    return failedRuntimeResult(
      target,
      peakRssBytes,
      `peak RSS ${formatByteSize(child.rssLimitExceededBytes)} exceeded ${formatByteSize(target.maxRssBytes)}`,
    )
  }
  if (child.timedOut) {
    return failedRuntimeResult(target, peakRssBytes, `worker timed out after ${String(options.timeoutMs)}ms`)
  }
  if (child.exitCode !== 0) {
    return failedRuntimeResult(
      target,
      peakRssBytes,
      `worker exited with ${child.signal ? `signal ${child.signal}` : `code ${String(child.exitCode)}`}: ${compactOutput(child.stderr || child.stdout)}`,
    )
  }
  try {
    const summary = parseCliSummary(child.stdout)
    const values = readSummaryValues(summary, target.reads)
    const mismatchedRead = firstMismatchedRead(values, target.expectedReads)
    if (mismatchedRead) {
      return failedRuntimeResult(target, peakRssBytes, mismatchedRead, values)
    }
    const diagnostics = readDiagnostics(summary)
    if (diagnostics?.engineMode !== 'streaming-native') {
      return failedRuntimeResult(
        target,
        peakRssBytes,
        `expected streaming-native diagnostics, received ${diagnostics?.engineMode ?? 'missing'}`,
        values,
      )
    }
    if (diagnostics.fallbackUsed !== false) {
      return failedRuntimeResult(target, peakRssBytes, 'streaming-native diagnostics did not report fallbackUsed=false', values)
    }
    if ((diagnostics.nativeKernelFormulaCellCount ?? 0) <= 0) {
      return failedRuntimeResult(target, peakRssBytes, 'streaming-native did not report native kernel formula evaluation', values)
    }
    return {
      id: target.id,
      label: target.label,
      status: 'passed',
      maxRssBytes: target.maxRssBytes,
      peakRssBytes,
      values,
      diagnosticsMaxObservedRssBytes: diagnostics.maxObservedRssBytes,
      fallbackUsed: diagnostics.fallbackUsed,
      patchedCacheCount: diagnostics.patchedCacheCount,
      nativeKernelFormulaCellCount: diagnostics.nativeKernelFormulaCellCount,
    }
  } catch (error) {
    return failedRuntimeResult(
      target,
      peakRssBytes,
      `worker returned invalid JSON: ${error instanceof Error ? error.message : String(error)}`,
    )
  }
}

export function writeSyntheticNativeRecalcWorkbook(
  cacheDir: string,
  options: { readonly rowCount?: number } = {},
): SyntheticWorkbookArtifact {
  const rowCount = options.rowCount ?? defaultSyntheticRowCount
  const filesDir = join(cacheDir, 'files')
  mkdirSync(filesDir, { recursive: true })
  const editRow = rowCount
  const worksheetXml = syntheticNativeRecalcWorksheetXml(rowCount)
  const bytes = zipSync({
    '[Content_Types].xml': strToU8(contentTypesXml()),
    '_rels/.rels': strToU8(rootRelationshipsXml()),
    'xl/workbook.xml': strToU8(workbookXml()),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRelationshipsXml()),
    'xl/worksheets/sheet1.xml': strToU8(worksheetXml),
  })
  const sha256 = createHash('sha256').update(bytes).digest('hex')
  const fileName = `synthetic-native-row-chain-${String(rowCount)}.xlsx`
  const filePath = join(filesDir, `${sha256}.xlsx`)
  writeFileSync(filePath, bytes)
  return {
    id: `synthetic-native-row-chain-${String(rowCount)}`,
    fileName,
    filePath,
    sha256,
    byteSize: bytes.byteLength,
    editTarget: `Data!A${String(editRow)}`,
    reads: [`Data!C${String(editRow)}`, `Data!B${String(editRow)}`],
    expectedReads: {
      [`Data!C${String(editRow)}`]: 168.75,
      [`Data!B${String(editRow)}`]: 28.125,
    },
  }
}

export function writeSyntheticIssue442NativeRecalcWorkbook(cacheDir: string): SyntheticWorkbookArtifact {
  const rowCount = 57_152
  const filesDir = join(cacheDir, 'files')
  mkdirSync(filesDir, { recursive: true })
  const worksheetXml = syntheticIssue442WorksheetXml(rowCount)
  const bytes = zipSync({
    '[Content_Types].xml': strToU8(contentTypesXml()),
    '_rels/.rels': strToU8(rootRelationshipsXml()),
    'xl/workbook.xml': strToU8(workbookXml()),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRelationshipsXml()),
    'xl/worksheets/sheet1.xml': strToU8(worksheetXml),
  })
  const sha256 = createHash('sha256').update(bytes).digest('hex')
  const fileName = `synthetic-issue-442-native-recalc-${String(rowCount)}.xlsx`
  const filePath = join(filesDir, `${sha256}.xlsx`)
  writeFileSync(filePath, bytes)
  return {
    id: 'synthetic-issue-442-native-recalc',
    fileName,
    filePath,
    sha256,
    byteSize: bytes.byteLength,
    editTarget: `Data!R${String(rowCount)}`,
    reads: [`Data!U${String(rowCount)}`, `Data!V${String(rowCount)}`],
    expectedReads: {
      [`Data!U${String(rowCount)}`]: 168.75,
      [`Data!V${String(rowCount)}`]: 28.125,
    },
  }
}

function syntheticNativeRecalcWorksheetXml(rowCount: number): string {
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    rows.push(
      `<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c><c r="B${String(
        row,
      )}"><f>A${String(row)}*1.7578125</f><v>0</v></c><c r="C${String(row)}"><f>B${String(row)}*6</f><v>0</v></c></row>`,
    )
  }
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<dimension ref="A1:C${String(rowCount)}"/>`,
    `<sheetData>${rows.join('')}</sheetData>`,
    '</worksheet>',
  ].join('')
}

function syntheticIssue442WorksheetXml(rowCount: number): string {
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    rows.push(
      `<row r="${String(row)}"><c r="R${String(row)}"><v>${String(row)}</v></c><c r="U${String(
        row,
      )}"><f>R${String(row)}*10.546875</f><v>0</v></c><c r="V${String(row)}"><f>U${String(row)}/6</f><v>0</v></c></row>`,
    )
  }
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<dimension ref="R1:V${String(rowCount)}"/>`,
    `<sheetData>${rows.join('')}</sheetData>`,
    '</worksheet>',
  ].join('')
}

function contentTypesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`
}

function rootRelationshipsXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
}

function workbookXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>'
}

function workbookRelationshipsXml(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
}

function runChildProcess(
  command: string,
  args: readonly string[],
  options: { readonly maxRssBytes: number; readonly timeoutMs: number },
): Promise<ChildRunResult> {
  return new Promise((resolvePromise) => {
    const child = spawn(command, args, {
      cwd: rootDir,
      detached: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let peakRssBytes = 0
    let settled = false
    let rssLimitExceededBytes: number | undefined
    let timedOut = false
    let stopRssWatchdog = noop
    const finish = (result: Omit<ChildRunResult, 'stdout' | 'stderr' | 'peakRssBytes' | 'rssLimitExceededBytes' | 'timedOut'>): void => {
      if (settled) {
        return
      }
      settled = true
      clearTimeout(timer)
      stopRssWatchdog()
      // oxlint-disable-next-line eslint-plugin-promise(no-multiple-resolved) -- `settled` gates close/error/watchdog races before resolving.
      resolvePromise({
        ...result,
        stdout,
        stderr,
        peakRssBytes: peakRssBytes || null,
        ...(rssLimitExceededBytes === undefined ? {} : { rssLimitExceededBytes }),
        timedOut,
      })
    }
    stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes: options.maxRssBytes,
      intervalMs: rssCheckIntervalMs,
      onSample: (rssBytes) => {
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
      },
      onLimitExceeded: (rssBytes) => {
        rssLimitExceededBytes = rssBytes
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
        terminateChildProcess(child, 'SIGTERM', { processGroup: true })
      },
    })
    const timer = setTimeout(() => {
      timedOut = true
      terminateChildProcess(child, 'SIGTERM', { processGroup: true })
    }, options.timeoutMs)
    timer.unref()
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
    })
    child.on('error', (error) => {
      stderr += error.message
      finish({ exitCode: 1, signal: null })
    })
    child.on('close', (exitCode, signal) => finish({ exitCode, signal }))
  })
}

function parseCliSummary(stdout: string): Readonly<Record<string, unknown>> {
  const parsed: unknown = JSON.parse(stdout.trim())
  return asRecord(parsed)
}

function readSummaryValues(summary: Readonly<Record<string, unknown>>, reads: readonly string[]): Readonly<Record<string, unknown>> {
  const readsObject = asRecord(summary['reads'])
  const values: Record<string, unknown> = {}
  for (const read of reads) {
    const cell = asRecord(readsObject[read])
    values[read] = cell['value']
  }
  return values
}

function readDiagnostics(summary: Readonly<Record<string, unknown>>): {
  readonly engineMode?: string
  readonly fallbackUsed?: boolean
  readonly maxObservedRssBytes?: number
  readonly patchedCacheCount?: number
  readonly nativeKernelFormulaCellCount?: number
} {
  const diagnostics = asRecord(summary['diagnostics'])
  const formulaCounts = asRecord(diagnostics['formulaCounts'])
  return {
    ...(typeof diagnostics['engineMode'] === 'string' ? { engineMode: diagnostics['engineMode'] } : {}),
    ...(typeof diagnostics['fallbackUsed'] === 'boolean' ? { fallbackUsed: diagnostics['fallbackUsed'] } : {}),
    ...(typeof diagnostics['maxObservedRssBytes'] === 'number' ? { maxObservedRssBytes: diagnostics['maxObservedRssBytes'] } : {}),
    ...(typeof diagnostics['patchedCacheCount'] === 'number' ? { patchedCacheCount: diagnostics['patchedCacheCount'] } : {}),
    ...(typeof formulaCounts['nativeKernelFormulaCellCount'] === 'number'
      ? { nativeKernelFormulaCellCount: formulaCounts['nativeKernelFormulaCellCount'] }
      : {}),
  }
}

function asRecord(value: unknown): Readonly<Record<string, unknown>> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error('expected object')
  }
  const record: Record<string, unknown> = {}
  for (const [key, entry] of Object.entries(value)) {
    record[key] = entry
  }
  return record
}

function firstMismatchedRead(
  values: Readonly<Record<string, unknown>>,
  expectedReads: Readonly<Record<string, number | string>>,
): string | null {
  for (const [target, expected] of Object.entries(expectedReads)) {
    const actual = values[target]
    if (typeof expected === 'number' && typeof actual === 'number' && Math.abs(actual - expected) <= 1e-9) {
      continue
    }
    if (actual === expected) {
      continue
    }
    return `read ${target} expected ${String(expected)}, received ${String(actual)}`
  }
  return null
}

function skippedResult(id: string, label: string, maxRssBytes: number, reason: string): NativeRecalcMemoryGateResult {
  return {
    id,
    label,
    status: 'skipped',
    maxRssBytes,
    peakRssBytes: null,
    reason,
  }
}

function failedSetupResult(id: string, label: string, maxRssBytes: number, reason: string): NativeRecalcMemoryGateResult {
  return {
    id,
    label,
    status: 'failed',
    maxRssBytes,
    peakRssBytes: null,
    reason,
  }
}

function failedRuntimeResult(
  target: NativeRecalcGateTarget,
  peakRssBytes: number | null,
  reason: string,
  values?: Readonly<Record<string, unknown>>,
): NativeRecalcMemoryGateResult {
  return {
    id: target.id,
    label: target.label,
    status: 'failed',
    maxRssBytes: target.maxRssBytes,
    peakRssBytes,
    ...(values === undefined ? {} : { values }),
    reason,
  }
}

function compactOutput(value: string): string {
  const compacted = value.replace(/\s+/gu, ' ').trim()
  return compacted.length > 500 ? `${compacted.slice(0, 500)}...` : compacted
}

if (import.meta.url === `file://${process.argv[1]}`) {
  await main()
}
