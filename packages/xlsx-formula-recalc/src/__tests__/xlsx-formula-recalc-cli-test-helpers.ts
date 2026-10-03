import { readFileSync } from 'node:fs'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { parse as parseYaml } from 'yaml'

import { WorkPaper, exportXlsx } from '@bilig/workpaper/xlsx'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
export const packageVersion = readPackageVersion()

interface CliSummary {
  readonly mode: string
  readonly externalWorkbooks: number
  readonly reads: Readonly<Record<string, { readonly value: unknown }>>
  readonly warnings: readonly string[]
  readonly diagnostics?: {
    readonly engineMode?: string
    readonly fallbackUsed?: boolean
    readonly externalWorkbookHydration?: Record<string, unknown>
  }
  readonly commandSucceeded: boolean
  readonly recalculationCompleted: boolean
  readonly expectedReadback?: Readonly<Record<string, number>>
  readonly expectedValueMatched?: boolean
  readonly excelParity: 'not_proven'
}

interface NoSheetJsChildOutput {
  readonly exitCode: number
  readonly stderr: string
  readonly before: readonly string[]
  readonly after: readonly string[]
  readonly summary: CliSummary
}

interface CliErrorSummary {
  readonly commandSucceeded: false
  readonly recalculationCompleted: false
  readonly error: string
}

interface CliInspectionSummary {
  readonly mode: string
  readonly schemaVersion: 'xlsx-cache-doctor.v1'
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly uninspectedFormulaCellCount: number
  readonly inspectionLimit: number | 'all'
  readonly staleCachedFormulaCount: number
  readonly cacheStatusSummary: {
    readonly inspected: number
    readonly stale: number
    readonly fresh: number
    readonly missingCache: number
    readonly unsupportedRecalculation: number
  }
  readonly suggestedReads: readonly string[]
  readonly formulas: ReadonlyArray<{
    readonly target: string
    readonly formula: string
    readonly cachedValue?: unknown
    readonly literalRecalculatedValue?: unknown
    readonly cacheStatus: 'fresh' | 'stale' | 'missing-cache' | 'unsupported-recalculation'
    readonly staleCachedValue: boolean | null
  }>
  readonly commandSucceeded: boolean
  readonly inspectionCompleted: boolean
  readonly recalculationCompleted: boolean
  readonly excelParity: 'not_proven'
}

interface WorkbookCompatibilityReportForTest {
  readonly schemaVersion: string
  readonly verified: boolean
  readonly diagnostics: {
    readonly engineMode: string
    readonly fallbackUsed: boolean
    readonly inputBytes: number
    readonly phaseRssPeaks: readonly unknown[]
    readonly maxObservedRssBytes: number
    readonly sheetCount: number
    readonly targetRowCount: number
    readonly editCount: number
    readonly readCount: number
    readonly formulaCounts: {
      readonly scannedFormulaCellCount: number
      readonly targetedFormulaCellCount: number
      readonly evaluatedFormulaCellCount: number
      readonly patchedFormulaCacheCount: number
      readonly unsupportedFormulaCellCount: number
      readonly nativeKernelFormulaCellCount: number
      readonly nativeKernelBatchCount: number
    }
    readonly patchedCacheCount: number
    readonly unsupportedReason?: string
  }
  readonly workbook: {
    readonly formulaCellCount: number
  }
  readonly findings: {
    readonly unsupportedFunctions: readonly { readonly name: string; readonly count: number }[]
    readonly volatileFunctions: readonly { readonly name: string; readonly count: number }[]
    readonly staleCachedFormulas: { readonly count: number }
    readonly missingCachedFormulaValues: { readonly count: number }
    readonly unsupportedRecalculations: { readonly count: number }
  }
  readonly risk: {
    readonly level: string
  }
  readonly cacheInspection: {
    readonly inspectionLimit: number | 'all'
    readonly uninspectedFormulaCellCount: number
  }
  readonly recalculationCompleted: boolean
  readonly excelParity: 'not_proven'
  readonly limitations: readonly string[]
}

export function readCliSummary(stdout: string): CliSummary {
  const parsed: unknown = JSON.parse(stdout)
  if (typeof parsed !== 'object' || parsed === null) {
    throw new Error(`Expected CLI summary object, received ${stdout}`)
  }
  const mode = Reflect.get(parsed, 'mode')
  const externalWorkbooks = Reflect.get(parsed, 'externalWorkbooks')
  const reads = Reflect.get(parsed, 'reads')
  const warnings = Reflect.get(parsed, 'warnings')
  const diagnostics = Reflect.get(parsed, 'diagnostics')
  const commandSucceeded = Reflect.get(parsed, 'commandSucceeded')
  const recalculationCompleted = Reflect.get(parsed, 'recalculationCompleted')
  const expectedReadback = Reflect.get(parsed, 'expectedReadback')
  const expectedValueMatched = Reflect.get(parsed, 'expectedValueMatched')
  const excelParity = Reflect.get(parsed, 'excelParity')
  if (
    typeof mode !== 'string' ||
    typeof externalWorkbooks !== 'number' ||
    typeof reads !== 'object' ||
    reads === null ||
    !Array.isArray(warnings) ||
    typeof commandSucceeded !== 'boolean' ||
    typeof recalculationCompleted !== 'boolean' ||
    typeof excelParity !== 'string'
  ) {
    throw new Error(`Unexpected CLI summary shape: ${stdout}`)
  }
  const parsedDiagnostics = readCliSummaryDiagnostics(diagnostics)
  return {
    mode,
    externalWorkbooks,
    reads: readCliSummaryReads(reads),
    warnings: warnings.filter((warning): warning is string => typeof warning === 'string'),
    ...(parsedDiagnostics ? { diagnostics: parsedDiagnostics } : {}),
    commandSucceeded,
    recalculationCompleted,
    ...(readNumericRecord(expectedReadback) ? { expectedReadback: readNumericRecord(expectedReadback) } : {}),
    ...(typeof expectedValueMatched === 'boolean' ? { expectedValueMatched } : {}),
    excelParity: excelParity === 'not_proven' ? excelParity : 'not_proven',
  }
}

export function readNoSheetJsChildOutput(stdout: string): NoSheetJsChildOutput {
  const parsed: unknown = JSON.parse(stdout)
  const record = requireRecord(parsed)
  return {
    exitCode: requireNumber(record['exitCode']),
    stderr: requireString(record['stderr']),
    before: requireStringArray(record['before']),
    after: requireStringArray(record['after']),
    summary: readCliSummary(JSON.stringify(record['summary'])),
  }
}

export function readCliErrorSummary(stdout: string): CliErrorSummary {
  const record = requireRecord(JSON.parse(stdout))
  return {
    commandSucceeded: requireFalse(record['commandSucceeded']),
    recalculationCompleted: requireFalse(record['recalculationCompleted']),
    error: requireString(record['error']),
  }
}

export function readCliInspectionSummary(stdout: string): CliInspectionSummary {
  const parsed: unknown = JSON.parse(stdout)
  if (!isRecord(parsed)) {
    throw new Error(`Expected CLI inspection summary object, received ${stdout}`)
  }
  const mode = parsed['mode']
  const schemaVersion = parsed['schemaVersion']
  const formulaCellCount = parsed['formulaCellCount']
  const inspectedFormulaCellCount = parsed['inspectedFormulaCellCount']
  const uninspectedFormulaCellCount = parsed['uninspectedFormulaCellCount']
  const inspectionLimit = parsed['inspectionLimit']
  const staleCachedFormulaCount = parsed['staleCachedFormulaCount']
  const cacheStatusSummary = parsed['cacheStatusSummary']
  const suggestedReads = parsed['suggestedReads']
  const formulas = parsed['formulas']
  const commandSucceeded = parsed['commandSucceeded']
  const inspectionCompleted = parsed['inspectionCompleted']
  const recalculationCompleted = parsed['recalculationCompleted']
  const excelParity = parsed['excelParity']
  if (
    typeof mode !== 'string' ||
    schemaVersion !== 'xlsx-cache-doctor.v1' ||
    typeof formulaCellCount !== 'number' ||
    typeof inspectedFormulaCellCount !== 'number' ||
    typeof uninspectedFormulaCellCount !== 'number' ||
    !isInspectionLimit(inspectionLimit) ||
    typeof staleCachedFormulaCount !== 'number' ||
    !isCliCacheStatusSummary(cacheStatusSummary) ||
    !Array.isArray(suggestedReads) ||
    !Array.isArray(formulas) ||
    typeof commandSucceeded !== 'boolean' ||
    typeof inspectionCompleted !== 'boolean' ||
    typeof recalculationCompleted !== 'boolean' ||
    excelParity !== 'not_proven'
  ) {
    throw new Error(`Unexpected CLI inspection summary shape: ${stdout}`)
  }
  return {
    mode,
    schemaVersion,
    formulaCellCount,
    inspectedFormulaCellCount,
    uninspectedFormulaCellCount,
    inspectionLimit,
    staleCachedFormulaCount,
    cacheStatusSummary,
    suggestedReads: suggestedReads.filter((read): read is string => typeof read === 'string'),
    formulas: formulas.filter(isCliInspectionFormula),
    commandSucceeded,
    inspectionCompleted,
    recalculationCompleted,
    excelParity,
  }
}

export function readWorkbookCompatibilityReport(stdout: string): WorkbookCompatibilityReportForTest {
  const parsed: unknown = JSON.parse(stdout)
  if (!isRecord(parsed)) {
    throw new Error(`Expected workbook compatibility report object, received ${stdout}`)
  }
  const workbook = parsed['workbook']
  const findings = parsed['findings']
  const risk = parsed['risk']
  const cacheInspection = parsed['cacheInspection']
  const diagnostics = requireRecord(parsed['diagnostics'])
  const formulaCounts = requireRecord(diagnostics['formulaCounts'])
  const limitations = parsed['limitations']
  if (!isRecord(workbook) || !isRecord(findings) || !isRecord(risk) || !isRecord(cacheInspection) || !Array.isArray(limitations)) {
    throw new Error(`Unexpected workbook compatibility report shape: ${stdout}`)
  }
  return {
    schemaVersion: requireString(parsed['schemaVersion']),
    verified: parsed['verified'] === true,
    diagnostics: {
      engineMode: requireString(diagnostics['engineMode']),
      fallbackUsed: diagnostics['fallbackUsed'] === true,
      inputBytes: requireNumber(diagnostics['inputBytes']),
      phaseRssPeaks: Array.isArray(diagnostics['phaseRssPeaks']) ? diagnostics['phaseRssPeaks'] : [],
      maxObservedRssBytes: requireNumber(diagnostics['maxObservedRssBytes']),
      sheetCount: requireNumber(diagnostics['sheetCount']),
      targetRowCount: requireNumber(diagnostics['targetRowCount']),
      editCount: requireNumber(diagnostics['editCount']),
      readCount: requireNumber(diagnostics['readCount']),
      formulaCounts: {
        scannedFormulaCellCount: requireNumber(formulaCounts['scannedFormulaCellCount']),
        targetedFormulaCellCount: requireNumber(formulaCounts['targetedFormulaCellCount']),
        evaluatedFormulaCellCount: requireNumber(formulaCounts['evaluatedFormulaCellCount']),
        patchedFormulaCacheCount: requireNumber(formulaCounts['patchedFormulaCacheCount']),
        unsupportedFormulaCellCount: requireNumber(formulaCounts['unsupportedFormulaCellCount']),
        nativeKernelFormulaCellCount: requireNumber(formulaCounts['nativeKernelFormulaCellCount']),
        nativeKernelBatchCount: requireNumber(formulaCounts['nativeKernelBatchCount']),
      },
      patchedCacheCount: requireNumber(diagnostics['patchedCacheCount']),
      ...(typeof diagnostics['unsupportedReason'] === 'string' ? { unsupportedReason: diagnostics['unsupportedReason'] } : {}),
    },
    workbook: {
      formulaCellCount: requireNumber(workbook['formulaCellCount']),
    },
    findings: {
      unsupportedFunctions: requireNamedCounts(findings['unsupportedFunctions']),
      volatileFunctions: requireNamedCounts(findings['volatileFunctions']),
      staleCachedFormulas: requireCountObject(findings['staleCachedFormulas']),
      missingCachedFormulaValues: requireCountObject(findings['missingCachedFormulaValues']),
      unsupportedRecalculations: requireCountObject(findings['unsupportedRecalculations']),
    },
    risk: {
      level: requireString(risk['level']),
    },
    cacheInspection: {
      inspectionLimit: cacheInspection['inspectionLimit'] === 'all' ? 'all' : requireNumber(cacheInspection['inspectionLimit']),
      uninspectedFormulaCellCount: requireNumber(cacheInspection['uninspectedFormulaCellCount']),
    },
    recalculationCompleted: parsed['recalculationCompleted'] === true,
    excelParity: parsed['excelParity'] === 'not_proven' ? 'not_proven' : 'not_proven',
    limitations: limitations.filter((limitation): limitation is string => typeof limitation === 'string'),
  }
}

export function readGeneratedWorkflow(stdout: string): Record<string, unknown> {
  return requireRecord(parseYaml(stdout))
}

function readPackageVersion(): string {
  const parsed: unknown = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'))
  if (!isRecord(parsed) || Array.isArray(parsed)) {
    throw new Error('Expected package.json object')
  }
  const version = parsed['version']
  if (typeof version !== 'string' || version.length === 0) {
    throw new Error('Expected package.json version')
  }
  return version
}

export function requireRecord(value: unknown): Record<string, unknown> {
  if (!isRecord(value) || Array.isArray(value)) {
    throw new Error(`Expected object, received ${typeof value}`)
  }
  return value
}

function requireString(value: unknown): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected string, received ${typeof value}`)
  }
  return value
}

function requireStringArray(value: unknown): readonly string[] {
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected string array, received ${typeof value}`)
  }
  return value
}

function requireFalse(value: unknown): false {
  if (value !== false) {
    throw new Error(`Expected false, received ${typeof value}`)
  }
  return value
}

function requireNumber(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`Expected finite number, received ${typeof value}`)
  }
  return value
}

function requireCountObject(value: unknown): { readonly count: number } {
  const record = requireRecord(value)
  return { count: requireNumber(record['count']) }
}

function requireNamedCounts(value: unknown): readonly { readonly name: string; readonly count: number }[] {
  return requireRecordArray(value).map((entry) => ({
    name: requireString(entry['name']),
    count: requireNumber(entry['count']),
  }))
}

export function requireRecordArray(value: unknown): readonly Record<string, unknown>[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected array, received ${typeof value}`)
  }
  return value.map(requireRecord)
}

function isInspectionLimit(value: unknown): value is CliInspectionSummary['inspectionLimit'] {
  return value === 'all' || typeof value === 'number'
}

function isCliInspectionFormula(value: unknown): value is CliInspectionSummary['formulas'][number] {
  return (
    isRecord(value) &&
    typeof value['target'] === 'string' &&
    typeof value['formula'] === 'string' &&
    isCliCacheStatus(value['cacheStatus']) &&
    (typeof value['staleCachedValue'] === 'boolean' || value['staleCachedValue'] === null)
  )
}

function isCliCacheStatus(value: unknown): value is CliInspectionSummary['formulas'][number]['cacheStatus'] {
  return value === 'fresh' || value === 'stale' || value === 'missing-cache' || value === 'unsupported-recalculation'
}

function isCliCacheStatusSummary(value: unknown): value is CliInspectionSummary['cacheStatusSummary'] {
  return (
    isRecord(value) &&
    typeof value['inspected'] === 'number' &&
    typeof value['stale'] === 'number' &&
    typeof value['fresh'] === 'number' &&
    typeof value['missingCache'] === 'number' &&
    typeof value['unsupportedRecalculation'] === 'number'
  )
}

function readCliSummaryDiagnostics(value: unknown): CliSummary['diagnostics'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const externalWorkbookHydration = value['externalWorkbookHydration']
  const engineMode = value['engineMode']
  const fallbackUsed = value['fallbackUsed']
  const parsed = {
    ...(typeof engineMode === 'string' ? { engineMode } : {}),
    ...(typeof fallbackUsed === 'boolean' ? { fallbackUsed } : {}),
    ...(isRecord(externalWorkbookHydration) ? { externalWorkbookHydration } : {}),
  }
  return Object.keys(parsed).length > 0 ? parsed : undefined
}

function readNumericRecord(value: unknown): Readonly<Record<string, number>> | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const parsed: Record<string, number> = {}
  for (const [entryKey, entryValue] of Object.entries(value)) {
    if (typeof entryValue !== 'number') {
      return undefined
    }
    parsed[entryKey] = entryValue
  }
  return parsed
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readCliSummaryReads(value: object): CliSummary['reads'] {
  const reads: Record<string, { readonly value: unknown }> = {}
  for (const [target, cellValue] of Object.entries(value)) {
    if (typeof cellValue !== 'object' || cellValue === null || !Reflect.has(cellValue, 'value')) {
      throw new Error(`Unexpected CLI read value for ${target}`)
    }
    reads[target] = {
      value: Reflect.get(cellValue, 'value'),
    }
  }
  return reads
}

export function readFileBytes(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path))
}

export function readCachedFormulaValue(bytes: Uint8Array, sheetPath: string, address: string): string | null {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const match = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)
  return match?.[1] ?? null
}

export function buildStaleFormulaCacheWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Sheet1: [
      ['Input', 'Output'],
      [2, '=A2*10'],
    ],
  })
  try {
    return replaceWorksheetCellXml(
      exportXlsx(workbook.exportSnapshot()),
      'xl/worksheets/sheet1.xml',
      'B2',
      '<c r="B2"><f>A2*10</f><v>999</v></c>',
    )
  } finally {
    workbook.dispose()
  }
}

export function buildMissingFormulaCacheWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Sheet1: [
      ['Input', 'Output'],
      [2, '=A2*10'],
    ],
  })
  try {
    return replaceWorksheetCellXml(exportXlsx(workbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f></c>')
  } finally {
    workbook.dispose()
  }
}

export function buildManyFormulaCacheWorkbook(options: { readonly formulaCount?: number } = {}): Uint8Array {
  const formulaCount = options.formulaCount ?? 60
  const staleRow = formulaCount + 1
  const rows: Array<[number | string, number | string]> = [['Input', 'Output']]
  for (let row = 2; row <= staleRow; row += 1) {
    rows.push([row - 1, `=A${row}*10`])
  }
  const workbook = WorkPaper.buildFromSheets({
    Sheet1: rows,
  })
  try {
    return replaceWorksheetCellXml(
      exportXlsx(workbook.exportSnapshot()),
      'xl/worksheets/sheet1.xml',
      `B${staleRow}`,
      `<c r="B${staleRow}"><f>A${staleRow}*10</f><v>999</v></c>`,
    )
  } finally {
    workbook.dispose()
  }
}

export function buildProviderBackedRiskWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Risks: [
      ['Signal', 'Formula'],
      ['Market data', '=GOOGLEFINANCE("GOOG","price")'],
      ['CSV import', '=IMPORTDATA("https://example.com/data.csv")'],
      ['HTML import', '=IMPORTHTML("https://example.com","table",1)'],
      ['Range import', '=IMPORTRANGE("source","Revenue!B2")'],
      ['Translate', '=TRANSLATE("hello","en","es")'],
    ],
  })
  try {
    return exportXlsx(workbook.exportSnapshot())
  } finally {
    workbook.dispose()
  }
}

export function buildExternalSourceWorkbook(rates: readonly [number, number, number]): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Rates: [
      ['SKU', 'Rate'],
      ['A', rates[0]],
      ['B', rates[1]],
      ['C', rates[2]],
    ],
  })
  try {
    return exportXlsx(workbook.exportSnapshot())
  } finally {
    workbook.dispose()
  }
}

function replaceWorksheetCellXml(bytes: Uint8Array, path: string, address: string, replacement: string): Uint8Array {
  const zip = unzipSync(bytes)
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  zip[path] = strToU8(xml.replace(new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>[\\s\\S]*?<\\/c>`, 'u'), replacement))
  return zipSync(zip)
}

export function buildExternalLinkRangeCacheWorkbook(target: string, options: { readonly lookupFormulas?: boolean } = {}): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [
      [null, 2, 120],
      [null, null, 40],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array()).replace(
      /<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u,
      '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>120</v></c>',
    )
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      options.lookupFormulas === false
        ? sheetXml
        : sheetXml.replace(
            /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
            "<c r=\"C2\"><f>_xlfn.XLOOKUP(&quot;B&quot;,'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1</f><v>40</v></c>",
          ),
    )
    zip['xl/workbook.xml'] = strToU8(
      ensureRelationshipNamespace(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())).replace(
        '</sheets>',
        '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
      ),
    )
    zip['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
        '</Relationships>',
        '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
      ),
    )
    zip['xl/externalLinks/externalLink5.xml'] = strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
        '<externalBook r:id="rId1">',
        '<sheetNames><sheetName val="Rates"/></sheetNames>',
        '<sheetDataSet><sheetData sheetId="0">',
        '<row r="2"><cell r="A2" t="str"><v>A</v></cell><cell r="B2"><v>10</v></cell></row>',
        '<row r="3"><cell r="A3" t="str"><v>B</v></cell><cell r="B3"><v>20</v></cell></row>',
        '<row r="4"><cell r="A4" t="str"><v>C</v></cell><cell r="B4"><v>30</v></cell></row>',
        '</sheetData></sheetDataSet>',
        '</externalBook>',
        '</externalLink>',
      ].join(''),
    )
    zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
        '</Relationships>',
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function ensureRelationshipNamespace(xml: string): string {
  return xml.replace(/<workbook\b([^>]*)>/u, (match) =>
    match.includes('xmlns:r=') ? match : match.replace('>', ` xmlns:r="${officeRelationshipNamespace}">`),
  )
}

export function readExternalLinkCacheCellValue(bytes: Uint8Array, address: string): string | null {
  const xml = strFromU8(unzipSync(bytes)['xl/externalLinks/externalLink5.xml'] ?? new Uint8Array())
  const match = new RegExp(`<cell\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/cell>`, 'u').exec(xml)
  return match?.[1] ?? null
}
