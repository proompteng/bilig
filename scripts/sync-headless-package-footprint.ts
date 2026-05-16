#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { formatJsonForRepo } from './scorecard-format.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const packageDir = join(repoRoot, 'packages', 'headless')
const outputPath = join(repoRoot, 'docs', 'headless-package-footprint.json')
const checkMode = process.argv.includes('--check')

const requiredDescription = 'Formula WorkPaper runtime for Node.js services and agent tools with JSON persistence and formula readback.'
const requiredKeywords = [
  'workpaper',
  'workbook-api',
  'formula-engine',
  'spreadsheet-formulas',
  'node',
  'typescript',
  'agent-tools',
  'mcp',
  'json-persistence',
  'xlsx',
] as const
const forbiddenDescriptionFragments = [
  'Headless spreadsheet engine',
  'XLSX import/export, agent tools, workbook JSON persistence, and service-side spreadsheet automation',
] as const

interface PackageManifest {
  readonly name: string
  readonly version: string
  readonly description: string
  readonly keywords: readonly string[]
  readonly dependencies: Readonly<Record<string, string>>
  readonly engines: {
    readonly node: string
  }
  readonly bin: Readonly<Record<string, string>>
  readonly exports: Readonly<Record<string, unknown>>
}

interface PackFile {
  readonly path: string
  readonly size: number
}

interface PackResult {
  readonly size: number
  readonly unpackedSize: number
  readonly entryCount: number
  readonly files: readonly PackFile[]
}

interface HeadlessPackageFootprint {
  readonly schemaVersion: 1
  readonly package: {
    readonly name: string
    readonly version: string
    readonly description: string
    readonly nodeEngine: string
    readonly keywordCount: number
    readonly keywords: readonly string[]
    readonly dependencyNames: readonly string[]
    readonly hasXlsxSubpath: boolean
    readonly hasMcpBinary: boolean
  }
  readonly npmPackDryRun: {
    readonly tarballBytes: number
    readonly unpackedBytes: number
    readonly entryCount: number
    readonly readmeBytes: number
    readonly packageJsonBytes: number
  }
}

function asRecord(value: unknown, context: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${context} must be an object`)
  }
  return Object.fromEntries(Object.entries(value))
}

function readString(record: Record<string, unknown>, key: string, context: string): string {
  const value = record[key]
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`${context}.${key} must be a non-empty string`)
  }
  return value
}

function readStringArray(record: Record<string, unknown>, key: string, context: string): readonly string[] {
  const value = record[key]
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string' && entry.length > 0)) {
    throw new Error(`${context}.${key} must be a non-empty string array`)
  }
  return value
}

async function readPackageManifest(): Promise<PackageManifest> {
  const parsed = asRecord(JSON.parse(await readFile(join(packageDir, 'package.json'), 'utf8')) as unknown, 'packages/headless/package.json')
  const engines = asRecord(parsed['engines'], 'packages/headless/package.json.engines')
  return {
    name: readString(parsed, 'name', 'packages/headless/package.json'),
    version: readString(parsed, 'version', 'packages/headless/package.json'),
    description: readString(parsed, 'description', 'packages/headless/package.json'),
    keywords: readStringArray(parsed, 'keywords', 'packages/headless/package.json'),
    dependencies: asStringRecord(parsed['dependencies'], 'packages/headless/package.json.dependencies'),
    engines: {
      node: readString(engines, 'node', 'packages/headless/package.json.engines'),
    },
    bin: asStringRecord(parsed['bin'], 'packages/headless/package.json.bin'),
    exports: asRecord(parsed['exports'], 'packages/headless/package.json.exports'),
  }
}

function asStringRecord(value: unknown, context: string): Readonly<Record<string, string>> {
  const record = asRecord(value, context)
  const strings: Record<string, string> = {}
  for (const [key, entry] of Object.entries(record)) {
    if (typeof entry !== 'string' || entry.length === 0) {
      throw new Error(`${context}.${key} must be a non-empty string`)
    }
    strings[key] = entry
  }
  return strings
}

function runPackDryRun(): PackResult {
  const result = spawnSync('npm', ['pack', '--dry-run', '--json'], {
    cwd: packageDir,
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`npm pack --dry-run failed:\n${result.stderr}`)
  }
  const parsed = JSON.parse(result.stdout) as unknown
  if (!Array.isArray(parsed) || parsed.length !== 1) {
    throw new Error('npm pack --dry-run --json must return exactly one package result')
  }
  const packageResult = asRecord(parsed[0], 'npm pack result')
  const files = packageResult['files']
  if (!Array.isArray(files)) {
    throw new Error('npm pack result.files must be an array')
  }
  return {
    size: readFiniteNumber(packageResult, 'size', 'npm pack result'),
    unpackedSize: readFiniteNumber(packageResult, 'unpackedSize', 'npm pack result'),
    entryCount: readFiniteNumber(packageResult, 'entryCount', 'npm pack result'),
    files: files.map((entry, index) => {
      const file = asRecord(entry, `npm pack result.files[${index.toString()}]`)
      return {
        path: readString(file, 'path', `npm pack result.files[${index.toString()}]`),
        size: readFiniteNumber(file, 'size', `npm pack result.files[${index.toString()}]`),
      }
    }),
  }
}

function readFiniteNumber(record: Record<string, unknown>, key: string, context: string): number {
  const value = record[key]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`${context}.${key} must be a finite number`)
  }
  return value
}

function assertPositioning(manifest: PackageManifest): void {
  if (manifest.description !== requiredDescription) {
    throw new Error(`packages/headless/package.json description must be: ${requiredDescription}`)
  }
  for (const fragment of forbiddenDescriptionFragments) {
    if (manifest.description.includes(fragment)) {
      throw new Error(`packages/headless/package.json description must not include stale positioning: ${fragment}`)
    }
  }
  if (manifest.keywords.length > 24) {
    throw new Error(`packages/headless/package.json has ${manifest.keywords.length.toString()} keywords; keep npm metadata compressed`)
  }
  for (const keyword of requiredKeywords) {
    if (!manifest.keywords.includes(keyword)) {
      throw new Error(`packages/headless/package.json keywords must include ${keyword}`)
    }
  }
}

function formatBytes(bytes: number): string {
  if (bytes >= 1_000_000) {
    return `${(bytes / 1_000_000).toFixed(2)} MB`
  }
  if (bytes >= 1_000) {
    return `${Math.round(bytes / 1_000).toString()} kB`
  }
  return `${bytes.toString()} B`
}

function buildFootprint(manifest: PackageManifest, pack: PackResult): HeadlessPackageFootprint {
  const readme = pack.files.find((file) => file.path === 'README.md')
  const packageJson = pack.files.find((file) => file.path === 'package.json')
  if (!readme || !packageJson) {
    throw new Error('npm pack dry run must include README.md and package.json')
  }
  return {
    schemaVersion: 1,
    package: {
      name: manifest.name,
      version: manifest.version,
      description: manifest.description,
      nodeEngine: manifest.engines.node,
      keywordCount: manifest.keywords.length,
      keywords: manifest.keywords,
      dependencyNames: Object.keys(manifest.dependencies).toSorted(),
      hasXlsxSubpath: Object.hasOwn(manifest.exports, './xlsx'),
      hasMcpBinary: Object.hasOwn(manifest.bin, 'bilig-workpaper-mcp'),
    },
    npmPackDryRun: {
      tarballBytes: pack.size,
      unpackedBytes: pack.unpackedSize,
      entryCount: pack.entryCount,
      readmeBytes: readme.size,
      packageJsonBytes: packageJson.size,
    },
  }
}

function renderMarkdownBlock(footprint: HeadlessPackageFootprint): string {
  return [
    '<!-- headless-package-footprint:start -->',
    '',
    `Current checked npm footprint for \`${footprint.package.name}@${footprint.package.version}\`:`,
    '',
    `- Pack dry run: \`${formatBytes(footprint.npmPackDryRun.tarballBytes)}\` tarball, \`${formatBytes(footprint.npmPackDryRun.unpackedBytes)}\` unpacked, \`${footprint.npmPackDryRun.entryCount.toString()}\` package entries.`,
    '- Boundary: the main import is the WorkPaper formula/JSON runtime; XLSX',
    '  import/export stays behind the `@bilig/headless/xlsx` subpath; MCP is the',
    '  `bilig-workpaper-mcp` binary wrapper.',
    `- Runtime: Node \`${footprint.package.nodeEngine}\`; Node 22 support waits for release CI coverage.`,
    '<!-- headless-package-footprint:end -->',
  ].join('\n')
}

function replaceMarkdownBlock(source: string, renderedBlock: string, path: string): string {
  const pattern = /<!-- headless-package-footprint:start -->[\s\S]*?<!-- headless-package-footprint:end -->/u
  if (!pattern.test(source)) {
    throw new Error(`${path} is missing the headless package footprint block`)
  }
  return source.replace(pattern, renderedBlock)
}

async function syncMarkdownBlocks(footprint: HeadlessPackageFootprint): Promise<void> {
  const renderedBlock = renderMarkdownBlock(footprint)
  const paths = ['README.md', join('packages', 'headless', 'README.md')] as const
  await Promise.all(
    paths.map(async (path) => {
      const absolutePath = join(repoRoot, path)
      const current = await readFile(absolutePath, 'utf8')
      const next = replaceMarkdownBlock(current, renderedBlock, path)
      if (checkMode && current !== next) {
        throw new Error(`${path} package footprint block is out of date. Run: pnpm headless:footprint:generate`)
      }
      if (!checkMode && current !== next) {
        await writeFile(absolutePath, next)
      }
    }),
  )
}

async function buildCurrentFootprint(): Promise<HeadlessPackageFootprint> {
  const manifest = await readPackageManifest()
  assertPositioning(manifest)
  return buildFootprint(manifest, runPackDryRun())
}

async function writeUntilStable(attempt = 0, previousJson?: string): Promise<HeadlessPackageFootprint> {
  if (attempt >= 5) {
    const footprint = await buildCurrentFootprint()
    const renderedJson = formatJsonForRepo(`${JSON.stringify(footprint, null, 2)}\n`)
    await writeFile(outputPath, renderedJson)
    await syncMarkdownBlocks(footprint)
    const stabilized = await buildCurrentFootprint()
    const stabilizedJson = formatJsonForRepo(`${JSON.stringify(stabilized, null, 2)}\n`)
    if (stabilizedJson !== renderedJson) {
      throw new Error('headless package footprint did not stabilize after updating generated README blocks')
    }
    return stabilized
  }

  const footprint = await buildCurrentFootprint()
  const renderedJson = formatJsonForRepo(`${JSON.stringify(footprint, null, 2)}\n`)
  await writeFile(outputPath, renderedJson)
  await syncMarkdownBlocks(footprint)
  if (previousJson === renderedJson) {
    return footprint
  }
  return writeUntilStable(attempt + 1, renderedJson)
}

if (checkMode) {
  const footprint = await buildCurrentFootprint()
  const renderedJson = formatJsonForRepo(`${JSON.stringify(footprint, null, 2)}\n`)
  const current = await readFile(outputPath, 'utf8')
  if (current !== renderedJson) {
    throw new Error('docs/headless-package-footprint.json is out of date. Run: pnpm headless:footprint:generate')
  }
  await syncMarkdownBlocks(footprint)
  console.log(
    JSON.stringify(
      {
        ok: true,
        outputPath,
        package: `${footprint.package.name}@${footprint.package.version}`,
        tarball: formatBytes(footprint.npmPackDryRun.tarballBytes),
        unpacked: formatBytes(footprint.npmPackDryRun.unpackedBytes),
        entryCount: footprint.npmPackDryRun.entryCount,
      },
      null,
      2,
    ),
  )
} else {
  const footprint = await writeUntilStable()
  console.log(
    JSON.stringify(
      {
        mode: 'write',
        outputPath,
        package: `${footprint.package.name}@${footprint.package.version}`,
        tarball: formatBytes(footprint.npmPackDryRun.tarballBytes),
        unpacked: formatBytes(footprint.npmPackDryRun.unpackedBytes),
        entryCount: footprint.npmPackDryRun.entryCount,
      },
      null,
      2,
    ),
  )
}
