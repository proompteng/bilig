import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)

function collectFilePaths(value: unknown): string[] {
  if (typeof value === 'string') {
    return /^(?:apps|packages|scripts|docs|examples|integrations)\/[^*?[\]]+\.(?:ts|tsx)$/u.test(value) ? [value] : []
  }
  if (Array.isArray(value)) {
    return value.flatMap(collectFilePaths)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.values(value).flatMap(collectFilePaths)
  }
  return []
}

function readRootKnipIgnoreFiles(): string[] {
  const config: unknown = JSON.parse(readFileSync(resolve(repoRoot, 'knip.json'), 'utf8'))
  if (typeof config !== 'object' || config === null || Array.isArray(config)) {
    throw new Error('knip.json must be an object')
  }
  const workspaces = Reflect.get(config, 'workspaces')
  const rootWorkspace = typeof workspaces === 'object' && workspaces !== null ? Reflect.get(workspaces, '.') : undefined
  const ignoreFiles = typeof rootWorkspace === 'object' && rootWorkspace !== null ? Reflect.get(rootWorkspace, 'ignoreFiles') : undefined
  if (!Array.isArray(ignoreFiles) || !ignoreFiles.every((entry): entry is string => typeof entry === 'string')) {
    throw new Error('knip.json root workspace ignoreFiles must be a string array')
  }
  return ignoreFiles
}

describe('repository tooling configuration', () => {
  it('does not exempt removed source files from the Oxlint policy', () => {
    const config: unknown = JSON.parse(readFileSync(resolve(repoRoot, '.oxlintrc.json'), 'utf8'))
    const configuredFiles = collectFilePaths(config)

    expect(configuredFiles.filter((path) => !existsSync(resolve(repoRoot, path)))).toEqual([])
  })

  it('does not retain Knip ignoreFiles patterns reported as redundant', () => {
    const redundantPatterns = new Set([
      '**/__tests__/**',
      '**/*.test.ts?(x)',
      'e2e/**',
      'playwright.config.ts',
      'playwright.prod.config.ts',
    ])

    expect(readRootKnipIgnoreFiles().filter((pattern) => redundantPatterns.has(pattern))).toEqual([])
  })

  it('keeps exact paths in the headless package workflow current', () => {
    const workflow = readFileSync(resolve(repoRoot, '.github/workflows/headless-package.yml'), 'utf8')
    const pullRequestPaths = workflow.match(/pull_request:\n\s+paths:\n(?<paths>[\s\S]*?)\n\s+workflow_dispatch:/u)?.groups?.paths ?? ''
    const releasePaths = workflow.match(/release_paths=\(\n(?<paths>[\s\S]*?)\n\s+\)/u)?.groups?.paths ?? ''
    const configuredPaths = [
      ...[...pullRequestPaths.matchAll(/^\s+- ['"]([^'"]+)['"]$/gmu)].map((match) => match[1]),
      ...[...releasePaths.matchAll(/^\s+((?:apps|packages|scripts|docs|examples|integrations|skills)\/[^\s]+)$/gmu)].map(
        (match) => match[1],
      ),
    ].filter((path) => !path.includes('*') && !path.includes('?'))

    expect(configuredPaths.filter((path) => !existsSync(resolve(repoRoot, path)))).toEqual([])
  })

  it('does not ignore retired product build directories in the Docker context', () => {
    const ignoredPaths = readFileSync(resolve(repoRoot, '.dockerignore'), 'utf8')
      .split(/\r?\n/u)
      .map((line) => line.trim())

    expect(ignoredPaths).not.toContain('apps/playground/dist')
    expect(ignoredPaths).not.toContain('apps/sync-server/dist')
  })
})
