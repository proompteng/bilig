#!/usr/bin/env bun

import { readdir, readFile, stat } from 'node:fs/promises'
import path from 'node:path'

import { parseSourceMaxLines } from './source-file-size-config.js'

const root = process.cwd()
const maxLines = parseSourceMaxLines(process.env['BILIG_SOURCE_MAX_LINES'])
const roots = ['apps', 'packages', 'scripts', 'e2e']
const ignoredDirNames = new Set(['node_modules', 'dist', 'build', 'generated', '__tests__', 'coverage'])
const ignoredFileSuffixes = ['.test.ts', '.test.tsx', '.pw.ts', '.d.ts']
const sourceExtensions = new Set(['.ts', '.tsx'])

interface FileSizeRecord {
  readonly lineCount: number
  readonly relativePath: string
}

const checkedFiles: FileSizeRecord[] = []
const violations: FileSizeRecord[] = []

await Promise.all(roots.map(async (relativeRoot) => walk(path.join(root, relativeRoot))))

checkedFiles.sort((left, right) => right.lineCount - left.lineCount || left.relativePath.localeCompare(right.relativePath))
violations.sort((left, right) => right.lineCount - left.lineCount || left.relativePath.localeCompare(right.relativePath))

if (violations.length > 0) {
  console.error(`Source file size check failed. Limit is ${String(maxLines)} lines for non-test, non-generated TypeScript files.`)
  for (const violation of violations) {
    console.error(`- ${violation.relativePath}: ${String(violation.lineCount)} lines`)
  }
  process.exit(1)
}

const largest = checkedFiles[0]
console.log(
  `Source file size check passed (${String(checkedFiles.length)} files, max ${
    largest ? `${largest.relativePath} at ${String(largest.lineCount)} lines` : '0 lines'
  }).`,
)

async function walk(currentPath: string): Promise<void> {
  const currentStat = await stat(currentPath)
  if (currentStat.isDirectory()) {
    if (ignoredDirNames.has(path.basename(currentPath))) {
      return
    }
    const entries = await readdir(currentPath, { withFileTypes: true })
    await Promise.all(entries.map(async (entry) => walk(path.join(currentPath, entry.name))))
    return
  }

  if (!sourceExtensions.has(path.extname(currentPath))) {
    return
  }

  const relativePath = path.relative(root, currentPath)
  if (ignoredFileSuffixes.some((suffix) => relativePath.endsWith(suffix))) {
    return
  }

  const lineCount = countLines(await readFile(currentPath, 'utf8'))
  const record = { lineCount, relativePath }
  checkedFiles.push(record)
  if (lineCount > maxLines) {
    violations.push(record)
  }
}

function countLines(content: string): number {
  if (content.length === 0) {
    return 0
  }
  return content.split(/\r\n|\r|\n/).length
}
