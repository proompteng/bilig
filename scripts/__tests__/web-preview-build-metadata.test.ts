import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { describe, expect, it } from 'vitest'

import {
  assertReusableWebPreviewBundleMatchesEnv,
  createWebPreviewBuildMetadata,
  webPreviewBuildMetadataFilename,
} from '../web-preview-build-metadata.js'

function withTempDist(run: (distDir: string) => void): void {
  const distDir = mkdtempSync(join(tmpdir(), 'bilig-web-dist-'))
  try {
    run(distDir)
  } finally {
    rmSync(distDir, { recursive: true, force: true })
  }
}

function writeMetadata(distDir: string, metadata: unknown): void {
  writeFileSync(join(distDir, webPreviewBuildMetadataFilename), `${JSON.stringify(metadata)}\n`, 'utf8')
}

describe('web preview build metadata', () => {
  it('records whether the bundle was built for remote sync', () => {
    expect(createWebPreviewBuildMetadata({ VITE_BILIG_REMOTE_SYNC: '0' }).remoteSyncEnabled).toBe(false)
    expect(createWebPreviewBuildMetadata({ VITE_BILIG_REMOTE_SYNC: 'false' }).remoteSyncEnabled).toBe(false)
    expect(createWebPreviewBuildMetadata({ VITE_BILIG_REMOTE_SYNC: '1' }).remoteSyncEnabled).toBe(true)
    expect(createWebPreviewBuildMetadata({}).remoteSyncEnabled).toBe(true)
  })

  it('allows reuse when the local stack expects the same sync mode', () => {
    withTempDist((distDir) => {
      writeMetadata(distDir, { schemaVersion: 1, remoteSyncEnabled: false })

      expect(() => assertReusableWebPreviewBundleMatchesEnv(distDir, { BILIG_E2E_REMOTE_SYNC: '0' })).not.toThrow()
    })
  })

  it('rejects stale preview bundles built for a different sync mode', () => {
    withTempDist((distDir) => {
      writeMetadata(distDir, { schemaVersion: 1, remoteSyncEnabled: true })

      expect(() => assertReusableWebPreviewBundleMatchesEnv(distDir, { BILIG_E2E_REMOTE_SYNC: '0' })).toThrow(
        'Refusing to reuse preview web bundle built with remoteSyncEnabled=true',
      )
    })
  })

  it('rejects bundle reuse when metadata is missing', () => {
    withTempDist((distDir) => {
      expect(() => assertReusableWebPreviewBundleMatchesEnv(distDir, { BILIG_E2E_REMOTE_SYNC: '0' })).toThrow(
        `Refusing to reuse preview web bundle without ${webPreviewBuildMetadataFilename}`,
      )
    })
  })
})
