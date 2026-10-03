#!/usr/bin/env node

import { runXlsxFormulaRecalcCliAsync } from '@bilig/xlsx-formula-recalc/cli-api'

process.exitCode = await runXlsxFormulaRecalcCliAsync(process.argv.slice(2), {
  commandName: 'exceljs-recalc',
})
