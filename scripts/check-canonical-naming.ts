#!/usr/bin/env bun

import { readdir, readFile, stat } from 'node:fs/promises'
import path from 'node:path'

const root = process.cwd()
const roots = ['packages', 'apps', 'scripts', 'docs', 'e2e']
const allowedHistoricalSegments = ['/history/', '/historical/']
const ignoredDirNames = new Set(['node_modules', 'dist', 'coverage', '.git', '.turbo'])
const ignoredExtensions = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp', '.wasm', '.tsbuildinfo'])
const token = (...parts) => parts.join('')
const bannedPatterns = [
  { label: token('top', '50'), regex: new RegExp(token('top', '50'), 'gi') },
  { label: token('top', '100'), regex: new RegExp(token('top', '100'), 'gi') },
  { label: token('Top', ' ', '50'), regex: new RegExp(token('Top', ' ', '50'), 'g') },
  { label: token('Top', ' ', '100'), regex: new RegExp(token('Top', ' ', '100'), 'g') },
  {
    label: token('top', '100', '-', 'canonical'),
    regex: new RegExp(token('top', '100', '-', 'canonical'), 'g'),
  },
  {
    label: token('post', '-', 'top', '100'),
    regex: new RegExp(token('post', '-', 'top', '100'), 'g'),
  },
  {
    label: token('formula', '-', 'top', '100'),
    regex: new RegExp(token('formula', '-', 'top', '100'), 'g'),
  },
]

const violations = []

await Promise.all(roots.map(async (relativeRoot) => walk(path.join(root, relativeRoot))))

if (violations.length > 0) {
  console.error('Canonical naming check failed:')
  for (const violation of violations) {
    console.error(`- ${violation}`)
  }
  process.exit(1)
}

async function walk(currentPath) {
  const currentStat = await stat(currentPath)
  if (currentStat.isDirectory()) {
    if (ignoredDirNames.has(path.basename(currentPath))) {
      return
    }
    const entries = await readdir(currentPath, { withFileTypes: true })
    await Promise.all(entries.map(async (entry) => walk(path.join(currentPath, entry.name))))
    return
  }

  if (ignoredExtensions.has(path.extname(currentPath))) {
    return
  }

  const relativePath = path.relative(root, currentPath)
  if (allowedHistoricalSegments.some((segment) => relativePath.includes(segment))) {
    return
  }

  for (const pattern of bannedPatterns) {
    if (pattern.regex.test(relativePath)) {
      violations.push(`${relativePath}: banned path token '${pattern.label}'`)
    }
    pattern.regex.lastIndex = 0
  }

  const content = await readFile(currentPath, 'utf8')
  for (const pattern of bannedPatterns) {
    if (pattern.regex.test(content)) {
      violations.push(`${relativePath}: banned content token '${pattern.label}'`)
    }
    pattern.regex.lastIndex = 0
  }
}
