#!/usr/bin/env node

import { runXlsxFormulaRecalcCli } from '@bilig/xlsx-formula-recalc/cli-api'

process.exitCode = runXlsxFormulaRecalcCli(process.argv.slice(2), {
  commandName: 'exceljs-recalc',
})
