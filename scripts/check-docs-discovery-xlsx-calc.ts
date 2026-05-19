import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireIncludes } from './check-docs-discovery-core.ts'

export async function requireXlsxCalcAlternativeDiscovery(docsRoot: string): Promise<void> {
  const content = await readFile(join(docsRoot, 'xlsx-calc-alternative-node-workbook-recalculation.md'), 'utf8')

  for (const required of [
    'title: xlsx-calc alternative for Node workbook recalculation',
    'canonical_url: https://proompteng.github.io/bilig/xlsx-calc-alternative-node-workbook-recalculation.html',
    'cd bilig/examples/xlsx-recalculation-node',
    '"exportedReimportMatchesAfter": true',
    '"formulasSurvivedXlsxRoundTrip": true',
    'npx --package xlsx-formula-recalc xlsx-recalc',
    'packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json',
    'WorkPaper mean wins: `4/4`',
    'WorkPaper p95 wins: `4/4`',
    'coverage note: this is a limited SheetJS-style workbook-wide comparison',
    'https://github.com/fabiooshiro/xlsx-calc',
    'https://docs.sheetjs.com/docs/csf/features/formulae/',
    'star the',
  ] as const) {
    requireIncludes(content, required, 'docs/xlsx-calc-alternative-node-workbook-recalculation.md')
  }
}
