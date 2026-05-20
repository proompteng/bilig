#!/usr/bin/env bun

import { writeFootprintWorkerResult } from './public-workbook-corpus-worker-commands.ts'

const defaultVerifyMaxRssBytes = 1536 * 1024 * 1024

if (process.argv[2] !== 'footprint-worker') {
  throw new Error('Expected footprint-worker command')
}

await writeFootprintWorkerResult({
  filePath: readStringArg('--file', ''),
  fileName: readStringArg('--file-name', 'workbook.xlsx'),
  verifyMaxRssBytes: readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes),
})

function readStringArg(name: string, fallback: string): string {
  for (const [index, arg] of process.argv.entries()) {
    if (arg === name) {
      const value = process.argv[index + 1]
      if (!value || value.startsWith('--')) {
        throw new Error(`Expected ${name} to have a value`)
      }
      return value
    }
    if (arg.startsWith(`${name}=`)) {
      const value = arg.slice(name.length + 1)
      if (value.length === 0) {
        throw new Error(`Expected ${name} to have a value`)
      }
      return value
    }
  }
  return fallback
}

function readMegabytesArg(name: string, fallbackBytes: number): number {
  const raw = readStringArg(name, String(Math.ceil(fallbackBytes / 1024 / 1024)))
  const parsed = Number(raw)
  if (!/^\d+$/u.test(raw) || !Number.isSafeInteger(parsed) || parsed <= 0) {
    throw new Error(`Expected ${name} to be a positive integer number of MiB`)
  }
  return parsed * 1024 * 1024
}
