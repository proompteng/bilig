#!/usr/bin/env node
import { runFormulaClinicCli } from '@bilig/headless/cli'
import { importXlsx } from './xlsx.js'

process.exitCode = runFormulaClinicCli({
  argv: process.argv.slice(2),
  importXlsx,
})
