#!/usr/bin/env bun

import { execFile } from 'node:child_process'
import { existsSync } from 'node:fs'
import { promisify } from 'node:util'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const exec = promisify(execFile)
const rootDir = dirname(dirname(fileURLToPath(import.meta.url)))
const ascBinaryName = process.platform === 'win32' ? 'asc.cmd' : 'asc'
const ascCandidates = [resolve(rootDir, 'node_modules/.bin', ascBinaryName), resolve(rootDir, '../../node_modules/.bin', ascBinaryName)]

function resolveAscBinary(): string {
  for (const candidate of ascCandidates) {
    if (existsSync(candidate)) {
      return candidate
    }
  }

  throw new Error(`Unable to locate AssemblyScript compiler. Tried: ${ascCandidates.join(', ')}`)
}

await exec(resolveAscBinary(), ['assembly/index.ts', '--target', 'release'], {
  cwd: rootDir,
})
