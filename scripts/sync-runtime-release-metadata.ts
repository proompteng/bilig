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
  const intro = extractChangelogIntro(existingChangelog)
  const normalizedNotes = normalizeReleaseNotes(version, notesMarkdown)
  const nextContent = `${intro}\n\n${releaseHeading}\n\n${normalizedNotes}\n\n${stripExistingReleaseSections(existingChangelog)}`
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

function extractChangelogIntro(content: string): string {
  const lines = content.trim().split('\n')
  const releaseIndex = lines.findIndex((line) => line.startsWith('## '))
  const introLines = releaseIndex >= 0 ? lines.slice(0, releaseIndex) : lines
  return introLines.join('\n').trim()
}

function stripExistingReleaseSections(content: string): string {
  const lines = content.trim().split('\n')
  const firstReleaseIndex = lines.findIndex((line) => line.startsWith('## '))
  if (firstReleaseIndex < 0) {
    return ''
  }
  return lines.slice(firstReleaseIndex).join('\n').trim()
}

function normalizeReleaseNotes(releaseVersion: string, notes: string): string {
  const lines = notes.trim().split('\n')
  const filtered = lines.filter((line, index) => !(index === 0 && line.trim() === `# Libraries v${releaseVersion}`))
  return filtered.join('\n').trim()
}

function readRequiredStringArg(name: string): string {
  const value = args.get(name)
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`--${name} is required`)
  }
  return value.trim()
}
