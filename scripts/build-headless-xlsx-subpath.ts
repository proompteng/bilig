#!/usr/bin/env bun

import { cpSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const excelImportDistDir = join(rootDir, 'packages', 'excel-import', 'dist')
const headlessDistDir = join(rootDir, 'packages', 'headless', 'dist')
const bundledXlsxDistDir = join(headlessDistDir, 'xlsx-internal')

if (!existsSync(join(excelImportDistDir, 'index.js')) || !existsSync(join(excelImportDistDir, 'index.d.ts'))) {
  throw new Error('Build @bilig/excel-import before building the @bilig/headless XLSX subpath')
}

mkdirSync(headlessDistDir, { recursive: true })
rmSync(bundledXlsxDistDir, { recursive: true, force: true })
cpSync(excelImportDistDir, bundledXlsxDistDir, { recursive: true })

writeFileSync(join(headlessDistDir, 'xlsx.js'), "export * from './xlsx-internal/index.js'\n")
writeFileSync(join(headlessDistDir, 'xlsx.d.ts'), "export * from './xlsx-internal/index.js'\n")
writeFileSync(
  join(headlessDistDir, 'formula-clinic-bin.js'),
  `#!/usr/bin/env node
import { importXlsx } from './xlsx.js'
import { runFormulaClinicCli } from './formula-clinic-cli.js'

process.exitCode = runFormulaClinicCli({
  argv: process.argv.slice(2),
  importXlsx,
  writeStderr: (text) => process.stderr.write(text),
  writeStdout: (text) => process.stdout.write(text),
})
`,
)
