#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))

const retainedExampleManifests = new Set([
  'examples/headless-workpaper/package.json',
  'examples/recalc-bridge-workflows/package.json',
  'examples/serverless-workpaper-api/package.json',
  'examples/xlsx-recalculation-node/package.json',
  'integrations/n8n-nodes-workpaper/package.json',
])

const allowedScorecardPaths = new Set(['.github/workflows/scorecard-code-scanning.yml', '.github/workflows/scorecard.yml'])

const forbiddenTrackedPathPatterns = [
  /^(?:coverage|playwright-report|test-results|artifacts\/fuzz|\.tmp)(?:\/|$)/u,
  /(?:^|\/)\.DS_Store$/u,
  /\.tsbuildinfo$/u,
  /(?:^|\/)(?:npm|pnpm)-debug\.log/u,
  /^internal\/growth\//u,
  /^packages\/storage-browser(?:\/|$)/u,
  /product-hunt/u,
  /community-growth/u,
  /community-launch/u,
  /public-workbook-corpus/u,
  /completion-audit/u,
  /active-not-achieved/u,
  /dominance/u,
  /uv\.lock$/u,
] as const

const publicDocPathPatterns = [
  /^README\.md$/u,
  /^docs\/.*\.(?:md|html|txt|json|xml)$/u,
  /^packages\/[^/]+\/(?:README|AGENTS|SKILL)\.md$/u,
  /^examples\/[^/]+\/README\.md$/u,
] as const

const browserImportGraphEntrypoints = [
  'apps/web/src/workbook-import-preview.worker.ts',
  'packages/excel-import/src/browser.ts',
  'packages/xlsx/src/browser.ts',
] as const

const nodeBuiltinSpecifiers = new Set([
  'node:fs',
  'node:path',
  'node:crypto',
  'node:zlib',
  'node:stream',
  'node:os',
  'fs',
  'path',
  'crypto',
  'zlib',
  'stream',
  'os',
])

const browserPackageEntrypoints = new Map<string, string>([
  ['@bilig/core', 'packages/core/src/index.ts'],
  ['@bilig/excel-import/browser', 'packages/excel-import/src/browser.ts'],
  ['@bilig/formula', 'packages/formula/src/index.ts'],
  ['@bilig/formula/external-function-adapter', 'packages/formula/src/external-function-adapter.ts'],
  ['@bilig/protocol', 'packages/protocol/src/index.ts'],
  ['@bilig/xlsx/browser', 'packages/xlsx/src/browser.ts'],
])

function gitTrackedFiles(): string[] {
  return execFileSync('git', ['ls-files', '-z'], { cwd: repoRoot, encoding: 'utf8' })
    .split('\0')
    .filter((path) => path.length > 0)
    .toSorted()
}

function readText(path: string): string {
  return readFileSync(join(repoRoot, path), 'utf8')
}

function packageScripts(): Record<string, string> {
  const packageJson: unknown = JSON.parse(readText('package.json'))
  if (typeof packageJson !== 'object' || packageJson === null || Array.isArray(packageJson)) {
    throw new Error('package.json must be a JSON object')
  }
  const scripts = Reflect.get(packageJson, 'scripts')
  if (typeof scripts !== 'object' || scripts === null || Array.isArray(scripts)) {
    return {}
  }
  return Object.fromEntries(Object.entries(scripts).filter((entry): entry is [string, string] => typeof entry[1] === 'string'))
}

function checkTrackedPaths(trackedFiles: readonly string[]): string[] {
  const violations: string[] = []
  for (const path of trackedFiles) {
    if (path.endsWith('package-lock.json') && path !== 'integrations/n8n-nodes-workpaper/package-lock.json') {
      violations.push(`${path}: npm lockfile is not owned by standalone n8n CI`)
    }
    if (!allowedScorecardPaths.has(path) && path.includes('scorecard')) {
      violations.push(`${path}: internal scorecard paths are forbidden outside OpenSSF workflows`)
    }
    for (const pattern of forbiddenTrackedPathPatterns) {
      if (pattern.test(path)) {
        violations.push(`${path}: forbidden tracked repository garbage path`)
      }
    }
    if (path === 'packages/create-workpaper/agent-overlay/package.json') {
      continue
    }
    if ((path.endsWith('/package.json') || path === 'package.json') && path !== 'package.json') {
      const isExampleOrIntegration = path.startsWith('examples/') || path.startsWith('integrations/')
      if (isExampleOrIntegration && !retainedExampleManifests.has(path)) {
        violations.push(`${path}: example/integration package manifest is not CI-owned`)
      }
    }
  }
  return violations
}

function checkPackageScripts(): string[] {
  const violations: string[] = []
  const scripts = packageScripts()
  for (const [name, command] of Object.entries(scripts)) {
    const scriptText = `${name} ${command}`
    for (const forbidden of [
      'community:growth',
      'docs:launch-assets',
      'docs:launch-demo',
      'product-hunt',
      'public-workbook-corpus',
      'completion-audit',
      'dominance',
    ] as const) {
      if (scriptText.includes(forbidden)) {
        violations.push(`package.json scripts.${name}: forbidden script surface ${forbidden}`)
      }
    }
  }
  for (const required of ['repo-garbage:check', 'examples:check'] as const) {
    if (!scripts[required]) {
      violations.push(`package.json scripts.${required}: required cleanup gate is missing`)
    }
  }
  if (!scripts['analyze:quality']?.includes('repo-garbage:check')) {
    violations.push('package.json scripts.analyze:quality: must include repo-garbage:check')
  }
  return violations
}

function checkPublicDocs(trackedFiles: readonly string[]): string[] {
  const violations: string[] = []
  for (const path of trackedFiles) {
    if (!publicDocPathPatterns.some((pattern) => pattern.test(path))) {
      continue
    }
    const source = readText(path)
    if (source.includes('/Users/gregkonush')) {
      violations.push(`${path}: public docs must not contain local absolute user paths`)
    }
    if (/\/Users\/[^)\]\s]+/u.test(source)) {
      violations.push(`${path}: public docs must not contain /Users absolute paths`)
    }
  }
  return violations
}

function checkAgentOverlayGeneration(trackedFiles: readonly string[]): string[] {
  const violations: string[] = []
  const generator = readText('scripts/sync-agent-discovery-docs.ts')
  const overlayFiles = trackedFiles.filter((path) => path.startsWith('packages/create-workpaper/agent-overlay/'))
  for (const path of overlayFiles) {
    if (!generator.includes(path)) {
      violations.push(`${path}: create-workpaper agent overlay file is not generated by sync-agent-discovery-docs.ts`)
    }
  }
  return violations
}

function checkWorkspacePackages(trackedFiles: readonly string[]): string[] {
  const violations: string[] = []
  const packageDirs = new Set(
    trackedFiles.filter((path) => path.startsWith('packages/') && path.endsWith('/package.json')).map((path) => dirname(path)),
  )
  const sourceDirs = new Set(
    trackedFiles
      .filter((path) => path.startsWith('packages/') && path.includes('/src/'))
      .map((path) => path.split('/').slice(0, 2).join('/')),
  )
  for (const dir of sourceDirs) {
    if (!packageDirs.has(dir)) {
      violations.push(`${dir}: package source directory has no package.json`)
    }
  }
  for (const dir of packageDirs) {
    if (dir === 'packages/create-workpaper' || dir === 'packages/create-workpaper/agent-overlay') {
      continue
    }
    if (!existsSync(join(repoRoot, dir, 'tsconfig.json'))) {
      violations.push(`${dir}: package directory has package.json but no tsconfig.json`)
    }
  }
  return violations
}

function runtimeImportSpecifiers(source: string): string[] {
  const specifiers: string[] = []
  for (const match of source.matchAll(/^\s*import\s+(?!type\b)(?:[^'"]*?\s+from\s+)?["']([^'"]+)["']/gmu)) {
    const specifier = match[1]
    if (specifier) {
      specifiers.push(specifier)
    }
  }
  for (const match of source.matchAll(/^\s*export\s+(?!type\b)[^'"]*?\s+from\s+["']([^'"]+)["']/gmu)) {
    const specifier = match[1]
    if (specifier) {
      specifiers.push(specifier)
    }
  }
  for (const match of source.matchAll(/\bimport\(\s*["']([^'"]+)["']\s*\)/gu)) {
    const specifier = match[1]
    if (specifier) {
      specifiers.push(specifier)
    }
  }
  return specifiers
}

function resolveBrowserImport(fromPath: string, specifier: string): string | null {
  if (specifier.startsWith('.')) {
    const base = join(dirname(fromPath), specifier).replaceAll('\\', '/')
    for (const candidate of [base, `${base}.ts`, `${base}.tsx`, `${base}/index.ts`]) {
      if (existsSync(join(repoRoot, candidate))) {
        return candidate
      }
    }
    return null
  }
  return browserPackageEntrypoints.get(specifier) ?? null
}

function checkBrowserImportGraphs(): string[] {
  const violations: string[] = []
  const visited = new Set<string>()
  const queue = [...browserImportGraphEntrypoints]

  for (let index = 0; index < queue.length; index += 1) {
    const path = queue[index]
    if (!path || visited.has(path)) {
      continue
    }
    visited.add(path)
    const source = readText(path)
    for (const specifier of runtimeImportSpecifiers(source)) {
      if (nodeBuiltinSpecifiers.has(specifier)) {
        violations.push(`${path}: browser import graph statically imports ${specifier}`)
        continue
      }
      const resolved = resolveBrowserImport(path, specifier)
      if (resolved && !visited.has(resolved)) {
        queue.push(resolved)
      }
    }
  }

  return violations
}

function main(): void {
  const trackedFiles = gitTrackedFiles()
  const violations = [
    ...checkTrackedPaths(trackedFiles),
    ...checkPackageScripts(),
    ...checkPublicDocs(trackedFiles),
    ...checkAgentOverlayGeneration(trackedFiles),
    ...checkWorkspacePackages(trackedFiles),
    ...checkBrowserImportGraphs(),
  ]

  if (violations.length > 0) {
    console.error(`Repository garbage policy failed:\n${violations.map((violation) => `- ${violation}`).join('\n')}`)
    process.exit(1)
  }
  console.log('Repository garbage policy passed.')
}

main()
