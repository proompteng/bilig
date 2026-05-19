#!/usr/bin/env bun

import { existsSync, readdirSync } from 'node:fs'
import { availableParallelism } from 'node:os'
import { join, resolve } from 'node:path'
import { buildVitestFuzzCommand, parseFuzzMode, type FuzzMode } from './run-fuzz-config.js'

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

function assertReplayFixtureExists(filePath: string): void {
  if (!existsSync(filePath)) {
    throw new Error(`Replay fixture does not exist: ${filePath}`)
  }
}

const args = process.argv.slice(2)
const mode = parseFuzzModeOrExit(args[0])
const replayFixture = mode === 'replay' ? args.slice(1).find((value) => value !== '--') : undefined

if (mode === 'replay' && !replayFixture) {
  console.error('Usage: pnpm test:fuzz -- replay <fixture-path>')
  process.exit(1)
}

const resolvedReplayFixture = replayFixture ? resolve(replayFixture) : null
if (resolvedReplayFixture) {
  assertReplayFixtureExists(resolvedReplayFixture)
}
const env = {
  BILIG_FUZZ_PROFILE: mode,
  BILIG_FUZZ_CAPTURE: '1',
  ...(resolvedReplayFixture ? { BILIG_FUZZ_REPLAY: resolvedReplayFixture } : {}),
}

const vitestFuzzFiles = listVitestFuzzFiles()
runCommand(buildVitestFuzzCommand(vitestFuzzFiles, availableParallelism()), env)

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

function parseFuzzModeOrExit(value: string | undefined): FuzzMode {
  try {
    return parseFuzzMode(value)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
}
