function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

function requireNotIncludes(haystack: string, needle: string, context: string): void {
  if (haystack.includes(needle)) {
    throw new Error(`${context} must not include ${needle}`)
  }
}

function requireAllIncludes(haystack: string, needles: string[], context: string): void {
  for (const needle of needles) {
    requireIncludes(haystack, needle, context)
  }
}

function requireNoIncludes(haystack: string, needles: string[], context: string): void {
  for (const needle of needles) {
    requireNotIncludes(haystack, needle, context)
  }
}

export function requireHomepageDiscovery(index: string, siteCss: string): void {
  requireAllIncludes(
    index,
    [
      '<link rel="canonical" href="https://proompteng.github.io/bilig/" />',
      '<link rel="sitemap" type="application/xml" href="https://proompteng.github.io/bilig/sitemap.xml" />',
      '<link rel="alternate" type="text/plain" href="https://proompteng.github.io/bilig/llms.txt" title="llms.txt" />',
      '"@type": "SoftwareSourceCode"',
      '"codeRepository": "https://github.com/proompteng/bilig"',
      '<title>bilig - Workbook Formulas for TypeScript Services</title>',
      '<meta name="robots" content="index, follow, max-image-preview:large" />',
      '<link rel="icon" type="image/svg+xml" href="./assets/favicon.svg" />',
      '<meta property="og:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-08-2" />',
      '<meta property="og:image:alt" content="bilig headless spreadsheet engine workbook and TypeScript code preview" />',
      '<meta name="twitter:card" content="summary_large_image" />',
      '<meta name="twitter:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-08-2" />',
      '<meta name="twitter:image:alt" content="bilig headless spreadsheet engine workbook and TypeScript code preview" />',
      '<link rel="stylesheet" href="./assets/site.css?v=2026-05-14-1" />',
      '<link rel="stylesheet" href="./assets/product-demo.css?v=2026-05-13-4" />',
      'Workbook formulas for TypeScript services.',
      'Keep a calculation as cells and formulas when that is the clearest model.',
      'Copy the `.ts` files and run them.',
      'Change an input cell and read the dependent total.',
      '89 starter issues',
      'Small docs, examples, adapters, and tests.',
      'Revenue.workpaper',
      'One workbook object. TypeScript changes it and reads the value it produced.',
      'Use a workbook when the formula is the thing you ship.',
      'If a model is already easier to review in rows and formulas',
      'Change the input.',
      'Run this before reading more.',
      'The smoke test changes one customer count',
      'curl -fsSLo eval.ts \\\n  https://proompteng.github.io/bilig/npm-eval.ts',
      'More TypeScript examples live in',
      'Use it from the place that already owns the workflow.',
      'Node route, queue worker, CLI, or MCP server: load the workbook',
      'The benchmark artifact includes the caveat.',
      "It compares WorkPaper with HyperFormula on the repo's headless formula workloads.",
      'It is useful evidence, not a promise about',
      'every Excel file, every formula, or the browser grid.',
      'WorkPaper vs HyperFormula',
      'mean latency',
      'latest artifact',
      'pnpm workpaper:bench:competitive:check',
      '<strong>46/46</strong>',
      'WorkPaper has the lower mean latency on every comparable row.',
      'One lookup workload is slower at p95 by',
      '<code>1.043x</code>',
      'Browser rendering speed, full Excel compatibility, or arbitrary customer workbooks.',
      'packages/benchmarks/baselines/<wbr />workpaper-vs-hyperformula.json',
      'Headless formula workloads that both engines can run.',
      'Benchmark notes',
      'Compatibility gaps',
      'Starter issues',
      'Choose the path that matches your job.',
      'Start with npm, a service route, an agent tool, or the engine comparison.',
      'Take a starter issue that improves the examples.',
      'Docs grouped by the work in front of you.',
      'Open the install path, service recipe, agent adapter, or comparison page you need right now.',
      '<h3>Run</h3>',
      '<h3>Build</h3>',
      '<h3>Agents</h3>',
      '<h3>Decide</h3>',
      'Run it on one calculation you care about.',
    ],
    'docs/index.html',
  )

  requireAllIncludes(
    siteCss,
    [
      '--font-body: ui-sans-serif',
      '--font-display: ui-sans-serif',
      "--font-mono: 'SFMono-Regular'",
      'grid-template-columns: minmax(0, 480px) minmax(0, 610px);',
      '.hero-notes',
      '.proof-facts',
    ],
    'docs/assets/site.css',
  )

  requireNoIncludes(
    index,
    [
      'bilig-hero-workbook-api.png?v=2026-05-08-2',
      '<link rel="stylesheet" href="./assets/fonts.css',
      'Run the TypeScript example',
      'Build a workbook in Node, change inputs through code',
      'For pricing, budgets, imports, payouts, and agent tools that still work like a workbook.',
      'Build the sheet, write the input,',
      'Use it for pricing calculators, budget checks, imports, payouts, and agent tools when the formula should stay in a workbook.',
      'Build the sheet, edit the input, read the cell, save JSON.',
      'Examples are `.ts` files with real imports.',
      'Change B2 and read the dependent total.',
      'Small first PRs are kept open.',
      'The workbook is data plus formulas. The code edits it and reads the result.',
      'Use a workbook when formulas are the clearest source of truth.',
      'No hidden browser, no screen scraping',
      'Run the benchmark, then decide.',
      'bilig is ahead on this checked-in WorkPaper suite.',
      '46/46 mean wins, with the caveat beside it.',
      'good news for this suite, not a blank check',
      'Run the benchmark before you depend on it.',
      'Fast where the benchmark says so. Clear where it does not.',
      'Speed claims are cheap. Run the benchmark.',
      'The benchmark command and JSON artifact are in the repo.',
      'Before you quote the number, run it.',
      'The result, baseline JSON, and p95 caveat are all in the repo.',
      'The current artifact shows 46 of 46 comparable mean rows ahead of HyperFormula',
      'Current result: 46/46 mean rows.',
      'The benchmark lives in the repo.',
      'Run the command, read the JSON, and check the slower row before you use the number in your own decision.',
      'The checked-in benchmark compares WorkPaper against HyperFormula-style workloads.',
      'One p95 caveat is listed beside it.',
      'Comparable benchmark rows only, measured by mean latency.',
      '<dd>46 / 46 ahead</dd>',
      'The JSON artifact changes in review when the benchmark moves.',
      'run the benchmark locally before depending on the headline row.',
      'Coverage and caveats',
      'Excel behavior not covered yet',
      'Small TypeScript tasks',
      'The speed claim is deliberately narrow',
      'Run it yourself.',
      'the slower p95 row named beside the result.',
      'Comparable benchmark rows only. Treat the number like a benchmark, not a slogan.',
      'Mean latency only. This is not a promise that every workbook, formula, or p95 row is faster.',
      '<dd>46 / 46 comparable mean rows</dd>',
      'The checked-in JSON is the thing to inspect in review, not a screenshot of the number.',
      'The artifact lives in the repo so benchmark drift shows up in review.',
      'That slower row stays visible because it matters if your workload looks like it.',
      'Rerun the command before using the number in your own docs.',
      'Artifact summary, caveat, and rerun instructions.',
      'Run the benchmark yourself.',
      'Do not use this as a blanket speed claim.',
      'Pick the closest starting point.',
      'Docs, examples, issues.',
      'Small docs, examples, and integration tasks for first PRs.',
      'Known gaps are documented in public before they become surprises.',
      'Everything public, including the rough edges.',
      'Try the package before you trust the page.',
      'Public project numbers',
      '<span>Stars</span>',
      '<strong>24</strong>',
      'No trust-me homepage claims',
      'trust-me',
      'What this benchmark says.',
      'The checked-in run has WorkPaper faster on mean latency for 46 of 46 comparable workloads.',
      '<dd>46 / 46</dd>',
      'Mean latency wins on the comparable rows in the current WorkPaper vs HyperFormula suite.',
      'CI checks this file, so benchmark drift shows up as a normal review diff.',
      'The approximate-lookup duplicate case is slower at p95. Benchmark your own workbook if that pattern matters.',
      'Read the benchmark notes',
      'Pick a starter issue',
      'One claim, with the caveat beside it.',
      'The benchmark is public. So are the gaps.',
      '46/46 mean wins. One p95 row is slower.',
      'The current WorkPaper vs HyperFormula artifact puts WorkPaper ahead on mean latency for this suite.',
      'Current checked-in result',
      '<strong>46 / 46</strong>',
      'Lower mean latency on every comparable row in the current artifact.',
      'Comparable headless workloads',
      'No UI-performance claim. No Excel-compatibility claim.',
      'Committed JSON, not a screenshot. Benchmark changes show up as normal diffs.',
      'p95 is 1.043x slower. If your workbook relies on approximate lookups with duplicates, test it first.',
      'Read what the numbers mean',
      'Check the Excel gaps',
      'Good first issues',
      'Public project signals',
      '<strong>40 starter tasks</strong>',
      '90 starter issues',
      '<strong>0.13.9</strong>',
      'Read those before you depend on the package.',
      'launch essay',
    ],
    'docs/index.html',
  )

  requireNotIncludes(siteCss, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'Bilig Condensed', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'Bilig Sans', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'Bilig Mono', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'border-left: 1px solid rgba(255, 250, 240, 0.16);', 'docs/assets/site.css')
}
