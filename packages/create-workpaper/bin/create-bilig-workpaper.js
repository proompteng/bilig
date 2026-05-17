#!/usr/bin/env node
import { cp, mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)))
const templateRoot = join(packageRoot, 'template')

const args = process.argv.slice(2)
const wantsHelp = args.includes('--help') || args.includes('-h')
const force = args.includes('--force')
const targetArg = args.find((arg) => !arg.startsWith('-'))

if (wantsHelp) {
  printHelp()
  process.exit(0)
}

const targetDirectory = resolve(process.cwd(), targetArg ?? 'bilig-workpaper-starter')
const projectName = normalizePackageName(targetArg ?? 'bilig-workpaper-starter')

await ensureWritableTarget(targetDirectory, force)
await copyTemplate(targetDirectory, projectName)

console.log(`Created ${relative(process.cwd(), targetDirectory) || '.'}`)
console.log('')
console.log('Next:')
console.log(`  cd ${relative(process.cwd(), targetDirectory) || '.'}`)
console.log('  npm install')
console.log('  npm run smoke')
console.log('')
console.log('Expected smoke output includes:')
console.log('  "verified": true')

function printHelp() {
  console.log(`@bilig/create-workpaper

Usage:
  npm create @bilig/workpaper@latest <directory>
  npm exec @bilig/create-workpaper@latest <directory>

Options:
  --force   Allow writing into an existing directory.
  -h, --help
`)
}

async function ensureWritableTarget(directory, allowExisting) {
  try {
    const existing = await stat(directory)
    if (!existing.isDirectory()) {
      throw new Error(`${directory} exists and is not a directory`)
    }
    const entries = await readdir(directory)
    if (entries.length > 0 && !allowExisting) {
      throw new Error(`${directory} is not empty. Re-run with --force to write into it.`)
    }
  } catch (error) {
    if (error?.code === 'ENOENT') {
      await mkdir(directory, { recursive: true })
      return
    }
    throw error
  }
}

async function copyTemplate(outputDirectory, packageName) {
  await cp(templateRoot, outputDirectory, {
    recursive: true,
    filter: async (source) => {
      const sourceStat = await stat(source)
      if (sourceStat.isDirectory()) {
        return true
      }
      const relativePath = relative(templateRoot, source)
      const targetPath = join(outputDirectory, relativePath)
      const text = await readFile(source, 'utf8')
      await mkdir(dirname(targetPath), { recursive: true })
      await writeFile(targetPath, text.replaceAll('__PROJECT_NAME__', packageName))
      return false
    },
  })
}

function normalizePackageName(name) {
  const parts = name.split(/[\\/]/)
  let base = 'bilig-workpaper-starter'
  for (let index = parts.length - 1; index >= 0; index -= 1) {
    if (parts[index] !== '') {
      base = parts[index]
      break
    }
  }
  const normalized = base
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '')

  return normalized || 'bilig-workpaper-starter'
}
