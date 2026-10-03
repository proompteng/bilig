import { readdirSync, readFileSync, statSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import * as workbookRoot from '../index.js'
import { checkWorkbookCommandRequest, checkWorkbookFeaturePlugin, defineWorkbookFeaturePlugin } from '../features-public.js'

interface PackageManifest {
  readonly dependencies?: Record<string, string>
  readonly devDependencies?: Record<string, string>
  readonly peerDependencies?: Record<string, string>
  readonly optionalDependencies?: Record<string, string>
  readonly exports?: Record<string, unknown>
  readonly files?: readonly string[]
  readonly scripts?: Record<string, string>
}

interface WorkspaceResolutionEntry {
  readonly packageDir: string
  readonly sourceEntry: string
}

const sourceRoot = fileURLToPath(new URL('../', import.meta.url))
const testRoot = fileURLToPath(new URL('./', import.meta.url))
const packageJsonPath = new URL('../../package.json', import.meta.url)
const readmePath = new URL('../../README.md', import.meta.url)
const workspaceResolutionPath = new URL('../../../../workspace-resolution.generated.json', import.meta.url)

const bannedRuntimeDependencies = Object.freeze([
  '@bilig/core',
  '@bilig/workpaper',
  '@bilig/agent-api',
  '@bilig/web',
  '@bilig/grid',
  'zod',
  'effect',
] as const)

const businessExampleTerms = Object.freeze([
  'findTable',
  'findRows',
  'Item',
  'Quantity',
  'Rate',
  'Status',
  'Total',
  'ready',
  'revenue',
  'prepaid',
  'forecast',
  'quote',
] as const)

function readPackageManifest(): PackageManifest {
  const parsed: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('package.json did not parse as an object')
  }
  return parsed as PackageManifest
}

function isWorkspaceResolutionEntry(value: unknown): value is WorkspaceResolutionEntry {
  return (
    typeof value === 'object' &&
    value !== null &&
    !Array.isArray(value) &&
    typeof Reflect.get(value, 'packageDir') === 'string' &&
    typeof Reflect.get(value, 'sourceEntry') === 'string'
  )
}

function readWorkspaceResolution(): Record<string, WorkspaceResolutionEntry> {
  const parsed: unknown = JSON.parse(readFileSync(workspaceResolutionPath, 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('workspace-resolution.generated.json did not parse as an object')
  }
  const resolution: Record<string, WorkspaceResolutionEntry> = {}
  for (const [key, value] of Object.entries(parsed)) {
    if (!isWorkspaceResolutionEntry(value)) {
      throw new Error(`workspace-resolution.generated.json entry ${key} is invalid`)
    }
    resolution[key] = value
  }
  return resolution
}

function walkSourceFiles(root: string): readonly string[] {
  const entries = readdirSync(root, { withFileTypes: true })
  const files: string[] = []
  for (const entry of entries) {
    if (entry.name === '__tests__' || entry.name.endsWith('.test.ts')) {
      continue
    }
    const path = `${root}/${entry.name}`
    if (entry.isDirectory()) {
      files.push(...walkSourceFiles(path))
      continue
    }
    if (entry.isFile() && entry.name.endsWith('.ts')) {
      files.push(path)
    }
  }
  return files
}

function walkTestFiles(root: string): readonly string[] {
  const files: string[] = []
  for (const entry of readdirSync(root, { withFileTypes: true })) {
    const path = `${root}/${entry.name}`
    if (entry.isDirectory()) {
      files.push(...walkTestFiles(path))
      continue
    }
    if (entry.isFile() && entry.name.endsWith('.test.ts')) {
      files.push(path)
    }
  }
  return files
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function importPattern(packageName: string): RegExp {
  const escaped = escapeRegExp(packageName)
  return new RegExp(`(?:from\\s+['"]${escaped}(?:/[^'"]*)?['"]|import\\(\\s*['"]${escaped}(?:/[^'"]*)?['"]\\s*\\))`)
}

function shapeExample(readme: string): string {
  const sectionStart = readme.indexOf('## The Shape')
  const sectionEnd = readme.indexOf('## Which Package')
  expect(sectionStart).toBeGreaterThanOrEqual(0)
  expect(sectionEnd).toBeGreaterThan(sectionStart)

  const section = readme.slice(sectionStart, sectionEnd)
  const match = /```ts\n([\s\S]*?)\n```/.exec(section)
  expect(match?.[1]).toBeDefined()
  return match?.[1] ?? ''
}

describe('@bilig/workbook package boundary', () => {
  it('keeps runtime dependencies boring and workbook-local', () => {
    const manifest = readPackageManifest()

    expect(manifest.dependencies).toEqual({
      '@bilig/formula': 'workspace:*',
      '@bilig/protocol': 'workspace:*',
    })

    const dependencyFields = [manifest.dependencies, manifest.devDependencies, manifest.peerDependencies, manifest.optionalDependencies]
    for (const dependencies of dependencyFields) {
      const names = new Set(Object.keys(dependencies ?? {}))
      for (const banned of bannedRuntimeDependencies) {
        expect(names.has(banned)).toBe(false)
      }
    }
  })

  it('keeps public source imports out of app, engine, ui, and validator packages', () => {
    const files = walkSourceFiles(sourceRoot)
    expect(files.length).toBeGreaterThan(0)

    const importPatterns = bannedRuntimeDependencies.map((dependency) => ({
      dependency,
      pattern: importPattern(dependency),
    }))

    for (const file of files) {
      expect(statSync(file).isFile()).toBe(true)
      const source = readFileSync(file, 'utf8')
      for (const { dependency, pattern } of importPatterns) {
        expect(pattern.test(source), `${file} imports ${dependency}`).toBe(false)
      }
      expect(source.includes('apps/'), `${file} imports app code`).toBe(false)
    }
  })

  it('publishes layered subpath exports without replacing the root barrel', () => {
    const exportsMap = readPackageManifest().exports
    const workspaceResolution = readWorkspaceResolution()
    const expectedSubpaths = Object.freeze([
      '.',
      './model',
      './prepare',
      './find',
      './check',
      './formula',
      './verify',
      './runtime',
      './features',
      './testing',
      './command',
      './schema',
    ] as const)

    expect(exportsMap).toMatchObject({
      '.': expect.any(Object),
      './model': expect.any(Object),
      './prepare': expect.any(Object),
      './find': expect.any(Object),
      './check': expect.any(Object),
      './formula': expect.any(Object),
      './verify': expect.any(Object),
      './runtime': expect.any(Object),
      './features': expect.any(Object),
      './testing': expect.any(Object),
      './command': expect.any(Object),
      './schema': expect.any(Object),
    })

    expect(Object.keys(exportsMap ?? {}).toSorted()).toEqual([...expectedSubpaths].toSorted())
    for (const subpath of expectedSubpaths) {
      const packageName = subpath === '.' ? '@bilig/workbook' : `@bilig/workbook/${subpath.slice(2)}`
      expect(workspaceResolution[packageName], `${packageName} missing from workspace resolution`).toMatchObject({
        packageDir: 'packages/workbook',
        sourceEntry: expect.stringMatching(/^packages\/workbook\/src\/.+\.ts$/),
      })
    }
    expect(readPackageManifest().files).toContain('fixtures')
  })

  it('runs every package-local test from the package test script', () => {
    const manifest = readPackageManifest()
    const testScript = manifest.scripts?.['test'] ?? ''
    const testFiles = walkTestFiles(testRoot)
      .map((file) => file.slice(sourceRoot.length).replaceAll('\\', '/').replaceAll('//', '/'))
      .toSorted()

    expect(testScript).toContain('scripts/run-vitest.ts --run')
    expect(testFiles.length).toBeGreaterThan(0)
    for (const file of testFiles) {
      expect(testScript, `${file} is missing from the package-local test script`).toContain(`packages/workbook/src/${file}`)
    }
  })

  it('keeps runtime feature extension helpers on the advanced features subpath', () => {
    expect('defineWorkbookFeaturePlugin' in workbookRoot).toBe(false)
    expect('checkWorkbookFeaturePlugin' in workbookRoot).toBe(false)
    expect('workbookProjectionInterceptorPoints' in workbookRoot).toBe(false)
    expect('workbookUiContributionSlots' in workbookRoot).toBe(false)
    expect('isWorkbookProjectionInterceptorPoint' in workbookRoot).toBe(false)
    expect('isWorkbookUiContributionSlot' in workbookRoot).toBe(false)
    expect('checkWorkbookCommandRequest' in workbookRoot).toBe(true)
    expect('checkWorkbookCommandResultForBundle' in workbookRoot).toBe(true)

    expect(checkWorkbookCommandRequest({ featureId: 'tables', commandId: 'create' })).toMatchObject({
      status: 'valid',
    })
    expect(
      checkWorkbookFeaturePlugin(
        defineWorkbookFeaturePlugin({
          id: 'tables',
          version: '1',
          commands: [],
          projectionInterceptors: [],
          uiContributions: [],
        }),
      ),
    ).toMatchObject({
      status: 'valid',
    })
  })

  it('keeps the first README path neutral and strict-proof oriented', () => {
    const readme = readFileSync(readmePath, 'utf8')
    const example = shapeExample(readme)

    expect(readme).toContain('## Use These First')
    expect(readme).toContain('## Mental Model')
    expect(readme).toContain('| Package')
    expect(readme).toContain('Choose when')
    expect(readme).toContain('Do not use for')
    expect(readme).toContain('| `@bilig/workbook`')
    expect(readme).toContain('Defining generic workbook intent')
    expect(readme).toContain('| `@bilig/workpaper`')
    expect(readme).toContain('Running workbook tools, MCP, or product workflows')
    expect(readme).toContain('| `@bilig/workpaper`')
    expect(readme).toContain('Owning workbook state inside Node')
    expect(readme).toContain('| `@bilig/core`')
    expect(readme).toContain('Implementing calculation or mutation internals')
    expect(readme).toContain('Generic command request, bundle, result, and receipt validators are available on')
    expect(readme).toContain('plugin registration, projection interceptors, and UI contribution metadata live')
    expect(readme).not.toContain('The main API is intentionally small')
    expect(readme).not.toContain('They are intentionally not on the root import path')
    expect(readme).not.toContain('examples/workbook-agent-model')
    expect(readme.split(/\r?\n/).length).toBeLessThanOrEqual(260)

    expect(example).toContain("workbook.findName('input')")
    expect(example).toContain("workbook.findName('factor')")
    expect(example).toContain("workbook.findName('result')")
    expect(example).toContain("prepareWorkbookAction(model, 'calculate')")
    expect(example).toContain('runWorkbookPlan(prepared.planData, adapter, { strict: true })')
    for (const term of businessExampleTerms) {
      expect(example.includes(term), `README first example contains ${term}`).toBe(false)
    }
  })
})
