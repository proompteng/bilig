#!/usr/bin/env bun

import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { loadRuntimePackages, parseStableSemver } from './runtime-package-set.ts'

export interface SyncRuntimePackageVersionsOptions {
  rootDir: string
  version: string
}

export interface SyncRuntimePackageVersionsResult {
  version: string
  updatedFiles: string[]
  updatedPackages: string[]
}

export function syncRuntimePackageVersions(options: SyncRuntimePackageVersionsOptions): SyncRuntimePackageVersionsResult {
  const version = options.version.trim()
  parseStableSemver(version)

  const updatedFiles: string[] = []
  const runtimePackages = loadRuntimePackages(options.rootDir)

  for (const runtimePackage of runtimePackages) {
    const packageJsonPath = join(options.rootDir, runtimePackage.dir, 'package.json')
    const manifest = readJsonRecord(packageJsonPath)
    manifest['version'] = version
    if (writeJsonIfChanged(packageJsonPath, manifest)) {
      updatedFiles.push(packageJsonPath)
    }
  }

  syncHeadlessMcpServerVersion(options.rootDir, version, updatedFiles)

  return {
    version,
    updatedFiles,
    updatedPackages: runtimePackages.map((runtimePackage) => runtimePackage.name),
  }
}

function syncHeadlessMcpServerVersion(rootDir: string, version: string, updatedFiles: string[]): void {
  const serverJsonPath = join(rootDir, 'packages/headless/server.json')
  const serverJson = readJsonRecord(serverJsonPath)
  serverJson['version'] = version

  const npmPackage = findNpmPackageEntry(serverJson, '@bilig/headless')
  if (!npmPackage) {
    throw new Error('packages/headless/server.json must include an npm package entry for @bilig/headless')
  }
  npmPackage['version'] = version

  if (writeJsonIfChanged(serverJsonPath, serverJson)) {
    updatedFiles.push(serverJsonPath)
  }
}

function findNpmPackageEntry(serverJson: Record<string, unknown>, packageName: string): Record<string, unknown> | undefined {
  const packages = serverJson['packages']
  if (!Array.isArray(packages)) {
    return undefined
  }
  return packages.find(
    (entry): entry is Record<string, unknown> => isRecord(entry) && entry['registryType'] === 'npm' && entry['identifier'] === packageName,
  )
}

function readJsonRecord(path: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isRecord(parsed)) {
    throw new Error(`Expected JSON object in ${path}`)
  }
  return parsed
}

function writeJsonIfChanged(path: string, value: Record<string, unknown>): boolean {
  const nextContent = `${JSON.stringify(value, null, 2)}\n`
  if (readFileSync(path, 'utf8') === nextContent) {
    return false
  }
  writeFileSync(path, nextContent)
  return true
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readRequiredStringArg(args: Map<string, string | true>, name: string): string {
  const value = args.get(name)
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`--${name} is required`)
  }
  return value.trim()
}

function parseArgs(argv: readonly string[]): Map<string, string | true> {
  const args = new Map<string, string | true>()
  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index]
    if (!value || value === '--' || !value.startsWith('--')) {
      continue
    }
    const key = value.slice(2)
    const nextValue = argv[index + 1]
    if (!nextValue || nextValue.startsWith('--')) {
      args.set(key, true)
      continue
    }
    args.set(key, nextValue)
    index += 1
  }
  return args
}

function isDirectInvocation(): boolean {
  const scriptPath = process.argv[1]
  return Boolean(scriptPath) && import.meta.url === pathToFileURL(resolve(scriptPath)).href
}

if (isDirectInvocation()) {
  const args = parseArgs(process.argv.slice(2))
  const version = readRequiredStringArg(args, 'version')
  const rootDir = resolve(new URL('..', import.meta.url).pathname)
  const result = syncRuntimePackageVersions({ rootDir, version })
  console.log(JSON.stringify(result, null, 2))
}
