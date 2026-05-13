import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

export function syncStagedMcpServerMetadata(packageName: string, stagedPackageDir: string, targetVersion: string): void {
  const manifest = readPackageManifest(stagedPackageDir)
  if (!shouldValidateMcpMetadata(packageName, manifest)) {
    return
  }

  const serverJsonPath = join(stagedPackageDir, 'server.json')
  const serverJson = readMcpServerJson(serverJsonPath)
  serverJson['version'] = targetVersion

  const npmPackage = findNpmPackageEntry(serverJson, manifest['name'])
  if (!npmPackage) {
    throw new Error(`Staged ${packageName} server.json must include an npm package entry for ${String(manifest['name'])}`)
  }
  npmPackage['version'] = targetVersion

  writeFileSync(serverJsonPath, `${JSON.stringify(serverJson, null, 2)}\n`)
}

export function validateStagedMcpServerMetadata(packageName: string, stagedPackageDir: string, expectedVersion: string): void {
  const manifest = readPackageManifest(stagedPackageDir)
  if (!shouldValidateMcpMetadata(packageName, manifest)) {
    return
  }

  const serverJsonPath = join(stagedPackageDir, 'server.json')
  const serverJson = readMcpServerJson(serverJsonPath)

  if (serverJson['name'] !== manifest['mcpName']) {
    throw new Error(
      `Staged ${packageName} server.json name must match package.json mcpName: ${String(serverJson['name'])} !== ${String(
        manifest['mcpName'],
      )}`,
    )
  }
  if (serverJson['version'] !== expectedVersion) {
    throw new Error(
      `Staged ${packageName} server.json version must match package version: ${String(serverJson['version'])} !== ${expectedVersion}`,
    )
  }

  const npmPackage = findNpmPackageEntry(serverJson, manifest['name'])
  if (!npmPackage) {
    throw new Error(`Staged ${packageName} server.json must include an npm package entry for ${String(manifest['name'])}`)
  }
  if (npmPackage['version'] !== expectedVersion) {
    throw new Error(
      `Staged ${packageName} server.json npm package version must match package version: ${String(npmPackage['version'])} !== ${expectedVersion}`,
    )
  }
}

function shouldValidateMcpMetadata(packageName: string, manifest: Record<string, unknown>): manifest is { name: string; mcpName: string } {
  if (packageName !== '@bilig/headless') {
    return false
  }
  return typeof manifest['name'] === 'string' && typeof manifest['mcpName'] === 'string'
}

function readPackageManifest(stagedPackageDir: string): Record<string, unknown> {
  const manifestPath = join(stagedPackageDir, 'package.json')
  return readJsonRecord(manifestPath)
}

function readMcpServerJson(serverJsonPath: string): Record<string, unknown> {
  if (!existsSync(serverJsonPath)) {
    throw new Error(`Staged MCP package is missing server.json: ${serverJsonPath}`)
  }
  return readJsonRecord(serverJsonPath)
}

function readJsonRecord(path: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isRecord(parsed)) {
    throw new Error(`Expected JSON object in ${path}`)
  }
  return parsed
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
