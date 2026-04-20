import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

export interface WorkspaceResolutionEntry {
  readonly packageDir: string
  readonly sourceEntry: string
}

export type WorkspaceResolutionMap = Record<string, WorkspaceResolutionEntry>

export const workspaceRootDir = fileURLToPath(new URL('..', import.meta.url))
export const workspaceResolutionJsonPath = join(workspaceRootDir, 'workspace-resolution.generated.json')
export const workspaceResolutionTsconfigPath = join(workspaceRootDir, 'tsconfig.workspace-paths.json')

function normalizePath(value: string): string {
  return value.replaceAll('\\', '/')
}

function sortedEntries<T>(entries: Iterable<readonly [string, T]>): Array<readonly [string, T]> {
  return [...entries].toSorted(([left], [right]) => left.localeCompare(right))
}

function stripKnownModuleExtension(value: string): string {
  if (value.endsWith('.d.ts')) {
    return value.slice(0, -'.d.ts'.length)
  }
  return value.replace(/\.(?:[cm]?js|tsx?|jsx?)$/, '')
}

function resolveSourceEntryFromExportTarget(packageDir: string, packageDirRelative: string, exportTarget: string): string | null {
  const normalizedTarget = normalizePath(exportTarget)
  const candidateSourceStem = normalizedTarget.startsWith('./dist/')
    ? stripKnownModuleExtension(normalizedTarget.slice('./dist/'.length))
    : normalizedTarget.startsWith('./src/')
      ? stripKnownModuleExtension(normalizedTarget.slice('./src/'.length))
      : null
  if (!candidateSourceStem) {
    return null
  }

  for (const extension of ['.ts', '.tsx'] as const) {
    const relativeSourceEntry = normalizePath(`src/${candidateSourceStem}${extension}`)
    if (existsSync(join(packageDir, relativeSourceEntry))) {
      return normalizePath(`${packageDirRelative}/${relativeSourceEntry}`)
    }
  }
  return null
}

function resolveExportTargetPath(exportValue: unknown): string | null {
  if (typeof exportValue === 'string') {
    return exportValue
  }
  if (typeof exportValue !== 'object' || exportValue === null) {
    return null
  }

  for (const key of ['import', 'types', 'default'] as const) {
    const nestedValue = Reflect.get(exportValue, key)
    if (typeof nestedValue === 'string') {
      return nestedValue
    }
  }
  return null
}

function scanWorkspaceExportEntries(
  packageName: string,
  packageDir: string,
  packageDirRelative: string,
  packageJsonValue: object,
): Array<readonly [string, WorkspaceResolutionEntry]> {
  const exportsValue = Reflect.get(packageJsonValue, 'exports')
  if (typeof exportsValue !== 'object' || exportsValue === null) {
    return []
  }

  const entries: Array<readonly [string, WorkspaceResolutionEntry]> = []
  for (const [subpath, exportValue] of Object.entries(exportsValue)) {
    if (!subpath.startsWith('./')) {
      continue
    }
    const exportTarget = resolveExportTargetPath(exportValue)
    if (!exportTarget) {
      continue
    }
    const sourceEntry = resolveSourceEntryFromExportTarget(packageDir, packageDirRelative, exportTarget)
    if (!sourceEntry) {
      continue
    }
    entries.push([
      `${packageName}/${subpath.slice(2)}`,
      {
        packageDir: normalizePath(packageDirRelative),
        sourceEntry,
      },
    ])
  }
  return entries
}

export function scanWorkspaceResolution(rootDir = workspaceRootDir): WorkspaceResolutionMap {
  const packagesDir = join(rootDir, 'packages')
  const entries: Array<readonly [string, WorkspaceResolutionEntry]> = []
  for (const directoryEntry of readdirSync(packagesDir, { withFileTypes: true })) {
    if (!directoryEntry.isDirectory()) {
      continue
    }
    const packageDir = join(packagesDir, directoryEntry.name)
    const packageJsonPath = join(packageDir, 'package.json')
    const sourceEntryPath = join(packageDir, 'src', 'index.ts')
    if (!existsSync(packageJsonPath) || !existsSync(sourceEntryPath)) {
      continue
    }
    const packageJsonValue = JSON.parse(readFileSync(packageJsonPath, 'utf8')) as unknown
    if (typeof packageJsonValue !== 'object' || packageJsonValue === null) {
      continue
    }
    const packageName = Reflect.get(packageJsonValue, 'name')
    if (typeof packageName !== 'string' || !packageName.startsWith('@bilig/')) {
      continue
    }
    const packageDirRelative = normalizePath(`packages/${directoryEntry.name}`)
    entries.push([
      packageName,
      {
        packageDir: packageDirRelative,
        sourceEntry: normalizePath(`${packageDirRelative}/src/index.ts`),
      },
    ])
    entries.push(...scanWorkspaceExportEntries(packageName, packageDir, packageDirRelative, packageJsonValue))
  }
  return Object.fromEntries(sortedEntries(entries))
}

function readWorkspaceResolution(rootDir = workspaceRootDir): WorkspaceResolutionMap {
  const workspaceResolutionPath = join(rootDir, 'workspace-resolution.generated.json')
  if (!existsSync(workspaceResolutionPath)) {
    throw new Error(`Missing workspace resolution artifact at ${workspaceResolutionPath}. Run bun scripts/gen-workspace-resolution.ts.`)
  }
  const parsed = JSON.parse(readFileSync(workspaceResolutionPath, 'utf8')) as unknown
  if (typeof parsed !== 'object' || parsed === null) {
    throw new Error('Workspace resolution artifact is not an object')
  }
  const entries: Array<readonly [string, WorkspaceResolutionEntry]> = []
  for (const [packageName, value] of Object.entries(parsed)) {
    if (typeof packageName !== 'string' || !packageName.startsWith('@bilig/')) {
      continue
    }
    if (
      typeof value !== 'object' ||
      value === null ||
      typeof value['packageDir'] !== 'string' ||
      typeof value['sourceEntry'] !== 'string'
    ) {
      throw new Error(`Invalid workspace resolution entry for ${packageName}`)
    }
    entries.push([
      packageName,
      {
        packageDir: value['packageDir'],
        sourceEntry: value['sourceEntry'],
      },
    ])
  }
  return Object.fromEntries(sortedEntries(entries))
}

function createWorkspaceAliasMap(rootDir = workspaceRootDir, resolution = readWorkspaceResolution(rootDir)): Record<string, string> {
  return Object.fromEntries(
    sortedEntries(Object.entries(resolution).map(([packageName, entry]) => [packageName, join(rootDir, entry.sourceEntry)])),
  )
}

export function createViteAliasRecord(
  extraAliases: Record<string, string> = {},
  rootDir = workspaceRootDir,
): Array<{ find: string; replacement: string }> {
  const explicitAliases = Object.entries(extraAliases).map(([find, replacement]) => ({
    find,
    replacement,
  }))
  const generatedAliases = Object.entries(createWorkspaceAliasMap(rootDir)).map(([find, replacement]) => ({
    find,
    replacement,
  }))
  return [...explicitAliases, ...generatedAliases].toSorted(
    (left, right) => right.find.length - left.find.length || left.find.localeCompare(right.find),
  )
}

export function createVitestAliasEntries(
  extraAliases: Array<{ find: string; replacement: string }> = [],
  rootDir = workspaceRootDir,
): Array<{ find: string; replacement: string }> {
  const generatedAliases = Object.entries(createWorkspaceAliasMap(rootDir)).map(([find, replacement]) => ({
    find,
    replacement,
  }))
  return [...extraAliases, ...generatedAliases]
}

export function createTsconfigPaths(resolution = readWorkspaceResolution(workspaceRootDir)): Record<string, string[]> {
  return Object.fromEntries(
    sortedEntries(Object.entries(resolution).map(([packageName, entry]) => [packageName, [`./${entry.sourceEntry}`]])),
  )
}
