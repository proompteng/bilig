import { spawnSync, type SpawnSyncReturns } from 'node:child_process'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { workspaceRootDir } from './workspace-resolution.js'

interface SpawnSyncLike {
  (
    command: string,
    args: readonly string[],
    options: {
      cwd: string
      env: NodeJS.ProcessEnv | undefined
      stdio: 'pipe'
    },
  ): SpawnSyncReturns<Buffer>
}

interface EnsureWasmKernelArtifactOptions {
  readonly rootDir?: string
  readonly existsSync?: (path: string) => boolean
  readonly spawnSync?: SpawnSyncLike
  readonly env?: NodeJS.ProcessEnv
}

export function resolveWasmKernelArtifactPath(rootDir = workspaceRootDir): string {
  return join(rootDir, 'packages/wasm-kernel/build/release.wasm')
}

function formatSpawnFailure(result: SpawnSyncReturns<Buffer>): string {
  if (result.error) {
    return result.error.message
  }
  const stderr = result.stderr.toString().trim()
  const stdout = result.stdout.toString().trim()
  return stderr || stdout || `Exited with status ${String(result.status)}`
}

export function ensureWasmKernelArtifact(options: EnsureWasmKernelArtifactOptions = {}): string {
  const rootDir = options.rootDir ?? workspaceRootDir
  const artifactPath = resolveWasmKernelArtifactPath(rootDir)
  const artifactExists = options.existsSync ?? existsSync
  if (artifactExists(artifactPath)) {
    return artifactPath
  }

  const spawnBuild = options.spawnSync ?? spawnSync
  const buildResult = spawnBuild('pnpm', ['wasm:build'], {
    cwd: rootDir,
    env: options.env ?? process.env,
    stdio: 'pipe',
  })
  if (buildResult.status !== 0) {
    throw new Error(`Failed to build wasm kernel artifact: ${formatSpawnFailure(buildResult)}`)
  }
  if (!artifactExists(artifactPath)) {
    throw new Error(`pnpm wasm:build completed but did not create '${artifactPath}'.`)
  }
  return artifactPath
}
