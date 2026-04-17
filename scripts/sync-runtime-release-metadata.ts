#!/usr/bin/env bun

import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

import { loadRuntimePackages, parseStableSemver } from './runtime-package-set.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const args = new Map<string, string | true>()
for (let index = 0; index < process.argv.length; index += 1) {
  const value = process.argv[index]
  if (!value || value === '--' || !value.startsWith('--')) {
    continue
  }
  const key = value.slice(2)
  const nextValue = process.argv[index + 1]
  if (!nextValue || nextValue.startsWith('--')) {
    args.set(key, true)
    continue
  }
  args.set(key, nextValue)
  index += 1
}

const version = readRequiredStringArg('version')
const notesFile = readRequiredStringArg('notes-file')
parseStableSemver(version)

const notesMarkdown = readFileSync(resolve(rootDir, notesFile), 'utf8').trim()
const runtimePackages = loadRuntimePackages(rootDir)

for (const runtimePackage of runtimePackages) {
  const packageJsonPath = join(rootDir, runtimePackage.dir, 'package.json')
  const manifest = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  manifest.version = version
  writeFileSync(packageJsonPath, `${JSON.stringify(manifest, null, 2)}\n`)
}

const changelogPath = join(rootDir, 'packages/headless/CHANGELOG.md')
const existingChangelog = readFileSync(changelogPath, 'utf8')
const releaseHeading = `## ${version}`
if (!existingChangelog.includes(releaseHeading)) {
  const header = '# Changelog\n\nAll notable changes to `@bilig/headless` will be documented in this file.\n'
  const nextContent = `${header}\n${releaseHeading}\n\n${notesMarkdown}\n\n${stripHeader(existingChangelog)}`
  writeFileSync(changelogPath, nextContent)
}

console.log(
  JSON.stringify(
    {
      version,
      updatedPackages: runtimePackages.map((runtimePackage) => runtimePackage.name),
      changelogPath,
    },
    null,
    2,
  ),
)

function stripHeader(content: string): string {
  const lines = content.trim().split('\n')
  const filtered = lines.filter(
    (line) =>
      line.trim() !== '# Changelog' &&
      line.trim() !== 'All notable changes to `@bilig/headless` will be documented in this file.' &&
      line.trim() !== 'This package uses release-please to manage versioned release notes.',
  )
  return `${filtered.join('\n').trim()}\n`
}

function readRequiredStringArg(name: string): string {
  const value = args.get(name)
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`--${name} is required`)
  }
  return value.trim()
}
