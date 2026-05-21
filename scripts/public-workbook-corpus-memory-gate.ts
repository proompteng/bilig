#!/usr/bin/env bun

import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

import { strToU8, zipSync } from 'fflate'

import { parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { readFlagArg, readStringArg } from './public-workbook-corpus-cli.ts'
import { formatByteSize } from './public-workbook-corpus-process.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase } from './public-workbook-corpus-types.ts'
import { verifyCachedWorkbookArtifactIsolated } from './public-workbook-corpus-verify.ts'

interface MemoryGateTarget {
  readonly artifactId: string
  readonly maxRssBytes: number
  readonly label: string
}

interface MemoryGateResult {
  readonly id: string
  readonly label: string
  readonly status: 'passed' | 'failed' | 'skipped'
  readonly maxRssBytes: number
  readonly peakRssBytes: number | null
  readonly cells?: number
  readonly reason?: string
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const mib = 1024 * 1024
const publicWorkbookMaxRssBytes = 112 * mib
// GitHub's Linux x64 Bun runner can sample the 750k compact-inspect path just above
// the shared 112 MiB target because RSS is page-granular and includes runtime
// allocator noise. Keep that headroom explicit and fixture-specific.
const synthetic750kMaxRssBytes = 113 * mib
const syntheticRepeatedStringMaxRssBytes = 112 * mib
const syntheticDuplicateSharedStringMaxRssBytes = 112 * mib
const syntheticMixedRichSharedStringMaxRssBytes = 112 * mib
const syntheticFormulaHeavyMaxRssBytes = 112 * mib
const syntheticCachedExternalFormulaMaxRssBytes = 112 * mib
const hardMaxRssBytes = 192 * mib
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultSyntheticCacheDir = join(rootDir, '.cache', 'public-workbook-corpus-memory-gate')
const verifyTimeoutMs = 180_000
const verifyMaxCellCount = 1_500_000
const rssCheckIntervalMs = 10
const publicTargets: readonly MemoryGateTarget[] = [
  { artifactId: 'workbook-f3c2a05d7d838a75', label: '0nContract_Register_Jan_2026.xlsx', maxRssBytes: publicWorkbookMaxRssBytes },
  { artifactId: 'workbook-5db97e9230dbaf6b', label: 'noibyfarmsize_fr.xlsx', maxRssBytes: publicWorkbookMaxRssBytes },
  { artifactId: 'workbook-ca98068307263914', label: 'sfgsme-efcpme-tables-2023-eng.xlsx', maxRssBytes: publicWorkbookMaxRssBytes },
]

async function main(): Promise<void> {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const syntheticCacheDir = resolve(readStringArg('--synthetic-cache-dir', defaultSyntheticCacheDir))
  const requirePublic = readFlagArg('--require-public')
  const syntheticOnly = readFlagArg('--synthetic-only')
  const results: MemoryGateResult[] = []

  if (!syntheticOnly) {
    results.push(...(await runPublicWorkbookGates({ cacheDir, manifestPath, requirePublic })))
  }
  results.push(await runSynthetic750kGate(syntheticCacheDir))
  results.push(await runSyntheticRepeatedStringGate(syntheticCacheDir))
  results.push(await runSyntheticDuplicateSharedStringGate(syntheticCacheDir))
  results.push(await runSyntheticMixedRichSharedStringGate(syntheticCacheDir))
  results.push(await runSyntheticFormulaHeavyGate(syntheticCacheDir))
  results.push(await runSyntheticCachedExternalFormulaGate(syntheticCacheDir))

  const failed = results.filter((result) => result.status === 'failed')
  process.stdout.write(
    `${JSON.stringify(
      {
        mode: 'public-workbook-corpus-memory-gate',
        targets: {
          publicWorkbookMaxRssBytes,
          synthetic750kMaxRssBytes,
          syntheticRepeatedStringMaxRssBytes,
          syntheticDuplicateSharedStringMaxRssBytes,
          syntheticMixedRichSharedStringMaxRssBytes,
          syntheticFormulaHeavyMaxRssBytes,
          syntheticCachedExternalFormulaMaxRssBytes,
          hardMaxRssBytes,
        },
        results,
      },
      null,
      2,
    )}\n`,
  )
  if (failed.length > 0) {
    throw new Error(
      `Public workbook memory gate failed: ${failed.map((result) => `${result.id} ${result.reason ?? ''}`.trim()).join('; ')}`,
    )
  }
}

async function runPublicWorkbookGates(args: {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly requirePublic: boolean
}): Promise<MemoryGateResult[]> {
  if (!existsSync(args.manifestPath)) {
    if (args.requirePublic) {
      return publicTargets.map((target) => failedMissingPublicResult(target, `manifest not found: ${args.manifestPath}`))
    }
    return publicTargets.map((target) => skippedResult(target, `manifest not found: ${args.manifestPath}`))
  }
  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(args.manifestPath, 'utf8')))
  const artifactsById = new Map(manifest.artifacts.map((artifact) => [artifact.id, artifact]))
  const results: MemoryGateResult[] = []
  for (const target of publicTargets) {
    const artifact = artifactsById.get(target.artifactId)
    const cachePath = artifact ? join(args.cacheDir, artifact.cachePath) : ''
    if (!artifact || !existsSync(cachePath)) {
      const reason = artifact ? `cache file not found: ${cachePath}` : `artifact not found in manifest: ${target.artifactId}`
      results.push(args.requirePublic ? failedMissingPublicResult(target, reason) : skippedResult(target, reason))
      continue
    }
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential workers keep the host RSS bounded while enforcing per-workbook gates.
    results.push(await runGateTarget(target, artifact, args.cacheDir, args.manifestPath))
  }
  return results
}

async function runSynthetic750kGate(cacheDir: string): Promise<MemoryGateResult> {
  const artifact = writeSynthetic750kWorkbook(cacheDir)
  return runGateTarget(
    {
      artifactId: artifact.id,
      label: artifact.fileName,
      maxRssBytes: synthetic750kMaxRssBytes,
    },
    artifact,
    cacheDir,
    join(cacheDir, 'manifest.json'),
  )
}

async function runSyntheticRepeatedStringGate(cacheDir: string): Promise<MemoryGateResult> {
  const artifact = writeSyntheticRepeatedStringWorkbook(cacheDir)
  return runGateTarget(
    {
      artifactId: artifact.id,
      label: artifact.fileName,
      maxRssBytes: syntheticRepeatedStringMaxRssBytes,
    },
    artifact,
    cacheDir,
    join(cacheDir, 'manifest.json'),
  )
}

async function runSyntheticDuplicateSharedStringGate(cacheDir: string): Promise<MemoryGateResult> {
  const artifact = writeSyntheticDuplicateSharedStringWorkbook(cacheDir)
  return runGateTarget(
    {
      artifactId: artifact.id,
      label: artifact.fileName,
      maxRssBytes: syntheticDuplicateSharedStringMaxRssBytes,
    },
    artifact,
    cacheDir,
    join(cacheDir, 'manifest.json'),
  )
}

async function runSyntheticMixedRichSharedStringGate(cacheDir: string): Promise<MemoryGateResult> {
  const artifact = writeSyntheticMixedRichSharedStringWorkbook(cacheDir)
  return runGateTarget(
    {
      artifactId: artifact.id,
      label: artifact.fileName,
      maxRssBytes: syntheticMixedRichSharedStringMaxRssBytes,
    },
    artifact,
    cacheDir,
    join(cacheDir, 'manifest.json'),
  )
}

async function runSyntheticFormulaHeavyGate(cacheDir: string): Promise<MemoryGateResult> {
  const artifact = writeSyntheticFormulaHeavyWorkbook(cacheDir)
  return runGateTarget(
    {
      artifactId: artifact.id,
      label: artifact.fileName,
      maxRssBytes: syntheticFormulaHeavyMaxRssBytes,
    },
    artifact,
    cacheDir,
    join(cacheDir, 'manifest.json'),
  )
}

async function runSyntheticCachedExternalFormulaGate(cacheDir: string): Promise<MemoryGateResult> {
  const artifact = writeSyntheticCachedExternalFormulaWorkbook(cacheDir)
  return runGateTarget(
    {
      artifactId: artifact.id,
      label: artifact.fileName,
      maxRssBytes: syntheticCachedExternalFormulaMaxRssBytes,
    },
    artifact,
    cacheDir,
    join(cacheDir, 'manifest.json'),
  )
}

async function runGateTarget(
  target: MemoryGateTarget,
  artifact: PublicWorkbookArtifact,
  cacheDir: string,
  manifestPath: string,
): Promise<MemoryGateResult> {
  const maxRssBytes = Math.min(target.maxRssBytes, hardMaxRssBytes)
  const verified = await verifyCachedWorkbookArtifactIsolated({
    artifact,
    cacheDir,
    manifestPath,
    runStructuralSmoke: false,
    timeoutMs: verifyTimeoutMs,
    maxRssBytes,
    maxCellCount: verifyMaxCellCount,
    rssCheckIntervalMs,
  })
  return memoryGateResult(target, verified, maxRssBytes)
}

function memoryGateResult(target: MemoryGateTarget, verified: PublicWorkbookCorpusCase, maxRssBytes: number): MemoryGateResult {
  const peakRssBytes = verified.peakRssBytes ?? null
  if (!verified.passed) {
    return {
      id: target.artifactId,
      label: target.label,
      status: 'failed',
      maxRssBytes,
      peakRssBytes,
      cells: verified.featureCounts?.cellCount,
      reason: `verification did not pass; status=${verified.status}`,
    }
  }
  if (peakRssBytes === null) {
    return {
      id: target.artifactId,
      label: target.label,
      status: 'failed',
      maxRssBytes,
      peakRssBytes,
      cells: verified.featureCounts?.cellCount,
      reason: 'peak RSS was not sampled',
    }
  }
  if (peakRssBytes > maxRssBytes) {
    return {
      id: target.artifactId,
      label: target.label,
      status: 'failed',
      maxRssBytes,
      peakRssBytes,
      cells: verified.featureCounts?.cellCount,
      reason: `peak RSS ${formatByteSize(peakRssBytes)} exceeded ${formatByteSize(maxRssBytes)}`,
    }
  }
  return {
    id: target.artifactId,
    label: target.label,
    status: 'passed',
    maxRssBytes,
    peakRssBytes,
    cells: verified.featureCounts?.cellCount,
  }
}

function skippedResult(target: MemoryGateTarget, reason: string): MemoryGateResult {
  return {
    id: target.artifactId,
    label: target.label,
    status: 'skipped',
    maxRssBytes: target.maxRssBytes,
    peakRssBytes: null,
    reason,
  }
}

function failedMissingPublicResult(target: MemoryGateTarget, reason: string): MemoryGateResult {
  return {
    ...skippedResult(target, reason),
    status: 'failed',
  }
}

function writeSynthetic750kWorkbook(cacheDir: string): PublicWorkbookArtifact {
  const rowCount = 150_000
  const columnCount = 5
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    const cells: string[] = []
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${String.fromCharCode(65 + column)}${String(row)}`
      cells.push(`<c r="${address}"><v>${String(row * (column + 1))}</v></c>`)
    }
    rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
  }
  return writeSyntheticWorkbookArtifact(cacheDir, {
    id: 'synthetic-750k-memory-v2',
    fileName: 'synthetic-750k-memory-v2.xlsx',
    sourceId: 'synthetic-memory-v2',
    worksheetXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<dimension ref="A1:E${String(rowCount)}"/>`,
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join(''),
  })
}

function writeSyntheticRepeatedStringWorkbook(cacheDir: string): PublicWorkbookArtifact {
  const rowCount = 25_000
  const columnCount = 5
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    const cells: string[] = []
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${String.fromCharCode(65 + column)}${String(row)}`
      cells.push(`<c r="${address}" t="inlineStr"><is><t>Repeated vendor label</t></is></c>`)
    }
    rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
  }
  return writeSyntheticWorkbookArtifact(cacheDir, {
    id: 'synthetic-repeated-string-memory-v4',
    fileName: 'synthetic-repeated-string-memory-v4.xlsx',
    sourceId: 'synthetic-repeated-string-memory-v4',
    worksheetXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<dimension ref="A1:E${String(rowCount)}"/>`,
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join(''),
  })
}

function writeSyntheticDuplicateSharedStringWorkbook(cacheDir: string): PublicWorkbookArtifact {
  const rowCount = 25_000
  const columnCount = 5
  const rows: string[] = []
  const sharedStrings: string[] = []
  let sharedStringIndex = 0
  for (let row = 1; row <= rowCount; row += 1) {
    const cells: string[] = []
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${String.fromCharCode(65 + column)}${String(row)}`
      cells.push(`<c r="${address}" t="s"><v>${String(sharedStringIndex)}</v></c>`)
      sharedStrings.push('<si><t>Repeated vendor label</t></si>')
      sharedStringIndex += 1
    }
    rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
  }
  return writeSyntheticWorkbookArtifact(cacheDir, {
    id: 'synthetic-duplicate-shared-string-memory-v4',
    fileName: 'synthetic-duplicate-shared-string-memory-v4.xlsx',
    sourceId: 'synthetic-duplicate-shared-string-memory-v4',
    sharedStringsXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${String(sharedStringIndex)}" uniqueCount="1">`,
      sharedStrings.join(''),
      '</sst>',
    ].join(''),
    worksheetXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<dimension ref="A1:E${String(rowCount)}"/>`,
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join(''),
  })
}

function writeSyntheticMixedRichSharedStringWorkbook(cacheDir: string): PublicWorkbookArtifact {
  const rowCount = 25_000
  const columnCount = 5
  const rows: string[] = []
  const sharedStrings: string[] = []
  let sharedStringIndex = 0
  const richSharedStringIndex = rowCount * columnCount - 1
  for (let row = 1; row <= rowCount; row += 1) {
    const cells: string[] = []
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${String.fromCharCode(65 + column)}${String(row)}`
      cells.push(`<c r="${address}" t="s"><v>${String(sharedStringIndex)}</v></c>`)
      sharedStrings.push(
        sharedStringIndex === richSharedStringIndex
          ? '<si><r><rPr><b/></rPr><t>Rich vendor label</t></r></si>'
          : '<si><t>Repeated vendor label</t></si>',
      )
      sharedStringIndex += 1
    }
    rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
  }
  return writeSyntheticWorkbookArtifact(cacheDir, {
    id: 'synthetic-mixed-rich-shared-string-memory-v4',
    fileName: 'synthetic-mixed-rich-shared-string-memory-v4.xlsx',
    sourceId: 'synthetic-mixed-rich-shared-string-memory-v4',
    sharedStringsXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${String(sharedStringIndex)}" uniqueCount="2">`,
      sharedStrings.join(''),
      '</sst>',
    ].join(''),
    worksheetXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<dimension ref="A1:E${String(rowCount)}"/>`,
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join(''),
  })
}

function writeSyntheticFormulaHeavyWorkbook(cacheDir: string): PublicWorkbookArtifact {
  const rowCount = 25_000
  const columnCount = 5
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    const cells: string[] = []
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${String.fromCharCode(65 + column)}${String(row)}`
      cells.push(`<c r="${address}"><f>1+1</f><v>2</v></c>`)
    }
    rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
  }
  return writeSyntheticWorkbookArtifact(cacheDir, {
    id: 'synthetic-formula-heavy-memory-v4',
    fileName: 'synthetic-formula-heavy-memory-v4.xlsx',
    sourceId: 'synthetic-formula-heavy-memory-v4',
    worksheetXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<dimension ref="A1:E${String(rowCount)}"/>`,
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join(''),
  })
}

function writeSyntheticCachedExternalFormulaWorkbook(cacheDir: string): PublicWorkbookArtifact {
  const rowCount = 25_000
  const columnCount = 5
  const formula = "'[external.xlsx]Sheet1'!A1"
  const rows: string[] = []
  const calcChainCells: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    const cells: string[] = []
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${String.fromCharCode(65 + column)}${String(row)}`
      cells.push(`<c r="${address}"><f>${formula}</f><v>${String(row + column)}</v></c>`)
      calcChainCells.push(`<c r="${address}" i="1"/>`)
    }
    rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
  }
  return writeSyntheticWorkbookArtifact(cacheDir, {
    id: 'synthetic-cached-external-formula-memory-v4',
    fileName: 'synthetic-cached-external-formula-memory-v4.xlsx',
    sourceId: 'synthetic-cached-external-formula-memory-v4',
    worksheetXml: [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<dimension ref="A1:E${String(rowCount)}"/>`,
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join(''),
    extraEntries: {
      'xl/calcChain.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${calcChainCells.join('')}</calcChain>`),
      'docProps/padding.bin': deterministicBytes(1_200_000),
    },
  })
}

function writeSyntheticWorkbookArtifact(
  cacheDir: string,
  input: {
    readonly fileName: string
    readonly id: string
    readonly sourceId: string
    readonly worksheetXml?: string
    readonly sheets?: readonly {
      readonly name: string
      readonly path: string
      readonly worksheetXml: string
    }[]
    readonly sharedStringsXml?: string
    readonly extraEntries?: Readonly<Record<string, Uint8Array>>
  },
): PublicWorkbookArtifact {
  const filesDir = join(cacheDir, 'files')
  mkdirSync(filesDir, { recursive: true })
  const sheets = input.sheets ?? [{ name: 'Data', path: 'xl/worksheets/sheet1.xml', worksheetXml: input.worksheetXml ?? '' }]
  const bytes = zipSync({
    '[Content_Types].xml': strToU8(
      contentTypesXml(sheets, input.sharedStringsXml !== undefined, input.extraEntries?.['xl/calcChain.xml'] !== undefined),
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    ),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets
    .map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${String(index + 1)}" r:id="rId${String(index + 1)}"/>`)
    .join('')}</sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets
  .map(
    (sheet, index) =>
      `<Relationship Id="rId${String(index + 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${sheet.path.slice('xl/'.length)}"/>`,
  )
  .join('')}
</Relationships>`),
    ...(input.sharedStringsXml !== undefined ? { 'xl/sharedStrings.xml': strToU8(input.sharedStringsXml) } : {}),
    ...Object.fromEntries(sheets.map((sheet) => [sheet.path, strToU8(sheet.worksheetXml)])),
    ...input.extraEntries,
  })
  const sha256 = createHash('sha256').update(bytes).digest('hex')
  const cachePath = `files/${sha256}.xlsx`
  writeFileSync(join(cacheDir, cachePath), bytes)
  const artifact: PublicWorkbookArtifact = {
    id: input.id,
    sourceId: input.sourceId,
    sourceUrl: `https://example.invalid/${input.fileName}`,
    downloadUrl: `https://example.invalid/${input.fileName}`,
    fileName: input.fileName,
    cachePath,
    sha256,
    byteSize: bytes.byteLength,
    workbookFingerprint: `synthetic-${sha256}`,
    fetchedAt: new Date(0).toISOString(),
    license: { title: 'Synthetic test fixture', url: 'https://example.invalid/license', spdx: 'LicenseRef-Synthetic' },
  }
  writeSyntheticManifest(cacheDir, artifact)
  return artifact
}

function contentTypesXml(sheets: readonly { readonly path: string }[], hasSharedStrings: boolean, hasCalcChain = false): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheets
    .map(
      (sheet) =>
        `<Override PartName="/${sheet.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('')}
  ${
    hasSharedStrings
      ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
      : ''
  }
  ${
    hasCalcChain
      ? '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
      : ''
  }
</Types>`
}

function deterministicBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  let state = 0x12345678
  for (let index = 0; index < length; index += 1) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    bytes[index] = (state >>> 24) & 0xff
  }
  return bytes
}

function writeSyntheticManifest(cacheDir: string, artifact: PublicWorkbookArtifact): void {
  writeFileSync(
    join(cacheDir, 'manifest.json'),
    JSON.stringify(
      {
        schemaVersion: 1,
        generatedAt: new Date(0).toISOString(),
        targetWorkbookCount: 1,
        sources: [
          {
            id: artifact.sourceId,
            kind: 'direct-url',
            sourceUrl: artifact.sourceUrl,
            downloadUrl: artifact.downloadUrl,
            fileName: artifact.fileName,
            discoveredAt: artifact.fetchedAt,
            license: artifact.license,
          },
        ],
        artifacts: [artifact],
      },
      null,
      2,
    ),
  )
}

if (import.meta.url === `file://${process.argv[1]}`) {
  await main()
}
