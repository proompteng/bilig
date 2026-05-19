#!/usr/bin/env node

import { runXlsxFormulaRecalcCli } from './cli-api.js'

process.exitCode = runXlsxFormulaRecalcCli(process.argv.slice(2))
