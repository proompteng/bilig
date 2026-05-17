import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

import { parseStrictBooleanEnvFlag } from './strict-env.js'

export const webPreviewBuildMetadataFilename = 'bilig-build-metadata.json'
export const webPreviewBuildMetadataSchemaVersion = 1

export interface WebPreviewBuildMetadata {
  readonly schemaVersion: typeof webPreviewBuildMetadataSchemaVersion
  readonly remoteSyncEnabled: boolean
}

export function createWebPreviewBuildMetadata(env: { readonly VITE_BILIG_REMOTE_SYNC?: string | undefined }): WebPreviewBuildMetadata {
  return {
    schemaVersion: webPreviewBuildMetadataSchemaVersion,
    remoteSyncEnabled: parseStrictBooleanEnvFlag(env.VITE_BILIG_REMOTE_SYNC, 'VITE_BILIG_REMOTE_SYNC', true),
  }
}

export function resolveExpectedWebPreviewRemoteSyncEnabled(env: { readonly BILIG_E2E_REMOTE_SYNC?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.BILIG_E2E_REMOTE_SYNC, 'BILIG_E2E_REMOTE_SYNC', true)
}

export function parseWebPreviewBuildMetadata(raw: string): WebPreviewBuildMetadata {
  const parsed = JSON.parse(raw) as unknown
  if (typeof parsed !== 'object' || parsed === null) {
    throw new Error(`${webPreviewBuildMetadataFilename} must contain a JSON object.`)
  }

  const schemaVersion = Reflect.get(parsed, 'schemaVersion')
  const remoteSyncEnabled = Reflect.get(parsed, 'remoteSyncEnabled')
  if (schemaVersion !== webPreviewBuildMetadataSchemaVersion || typeof remoteSyncEnabled !== 'boolean') {
    throw new Error(
      `${webPreviewBuildMetadataFilename} must contain schemaVersion=${webPreviewBuildMetadataSchemaVersion} and boolean remoteSyncEnabled.`,
    )
  }

  return {
    schemaVersion,
    remoteSyncEnabled,
  }
}

export function readWebPreviewBuildMetadata(distDir: string): WebPreviewBuildMetadata | null {
  const path = join(distDir, webPreviewBuildMetadataFilename)
  if (!existsSync(path)) {
    return null
  }
  return parseWebPreviewBuildMetadata(readFileSync(path, 'utf8'))
}

export function assertReusableWebPreviewBundleMatchesEnv(
  distDir: string,
  env: {
    readonly BILIG_E2E_REMOTE_SYNC?: string | undefined
  },
): void {
  const metadata = readWebPreviewBuildMetadata(distDir)
  if (!metadata) {
    throw new Error(
      `Refusing to reuse preview web bundle without ${webPreviewBuildMetadataFilename}; rebuild it with BILIG_DEV_WEB_PREVIEW_BUILD=1.`,
    )
  }

  const expectedRemoteSyncEnabled = resolveExpectedWebPreviewRemoteSyncEnabled(env)
  if (metadata.remoteSyncEnabled !== expectedRemoteSyncEnabled) {
    throw new Error(
      `Refusing to reuse preview web bundle built with remoteSyncEnabled=${String(
        metadata.remoteSyncEnabled,
      )}; current BILIG_E2E_REMOTE_SYNC expects remoteSyncEnabled=${String(expectedRemoteSyncEnabled)}. Rebuild the bundle.`,
    )
  }
}
