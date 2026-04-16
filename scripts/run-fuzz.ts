#!/usr/bin/env bun

import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join, resolve } from 'node:path'

type FuzzMode = 'default' | 'main' | 'nightly' | 'replay'

function parseMode(value: string | undefined): FuzzMode {
  if (value === 'main' || value === 'nightly' || value === 'replay') {
    return value
  }
  return 'default'
}

function runCommand(command: string[], extraEnv: Record<string, string>): void {
  const result = Bun.spawnSync(command, {
    stdin: 'inherit',
    stdout: 'inherit',
    stderr: 'inherit',
    env: {
      ...process.env,
      ...extraEnv,
    },
  })
  if (result.exitCode !== 0) {
    process.exit(result.exitCode ?? 1)
  }
}

function listVitestFuzzFiles(): string[] {
  return ['packages', 'apps'].flatMap((root) => walkFuzzFiles(root)).toSorted((left, right) => left.localeCompare(right))
}

function parseReplayKind(filePath: string): string | null {
  if (!existsSync(filePath)) {
    throw new Error(`Replay fixture does not exist: ${filePath}`)
  }
  const parsed = JSON.parse(readFileSync(filePath, 'utf8')) as unknown
  return typeof parsed === 'object' && parsed !== null && typeof parsed['kind'] === 'string' ? parsed['kind'] : null
}

const args = process.argv.slice(2)
const mode = parseMode(args[0])
const replayFixture = mode === 'replay' ? args.slice(1).find((value) => value !== '--') : undefined

if (mode === 'replay' && !replayFixture) {
  console.error('Usage: pnpm test:fuzz:replay -- <fixture-path>')
  process.exit(1)
}

const resolvedReplayFixture = replayFixture ? resolve(replayFixture) : null
const replayKind = resolvedReplayFixture ? parseReplayKind(resolvedReplayFixture) : null
const env = {
  BILIG_FUZZ_PROFILE: mode,
  BILIG_FUZZ_CAPTURE: '1',
  ...(resolvedReplayFixture ? { BILIG_FUZZ_REPLAY: resolvedReplayFixture } : {}),
}

const vitestFuzzFiles = listVitestFuzzFiles()
runCommand(['pnpm', 'exec', 'vitest', 'run', ...vitestFuzzFiles], env)

if (!resolvedReplayFixture || replayKind === 'browser') {
  runCommand(['bun', 'scripts/run-browser-tests.ts', '--grep', '@fuzz-browser'], {
    ...env,
    BILIG_FUZZ_BROWSER: '1',
  })
}

function walkFuzzFiles(root: string): string[] {
  if (!existsSync(root)) {
    return []
  }

  const files: string[] = []
  for (const entry of readdirSync(root, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name.startsWith('.')) {
      continue
    }

    const relativePath = join(root, entry.name)
    if (entry.isDirectory()) {
      files.push(...walkFuzzFiles(relativePath))
      continue
    }

    if (entry.isFile() && relativePath.endsWith('.fuzz.test.ts')) {
      files.push(relativePath)
    }
  }

  return files
}
