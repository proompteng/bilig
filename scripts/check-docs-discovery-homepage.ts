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
      '<title>bilig - Spreadsheet Formulas for TypeScript Services</title>',
      '<meta name="robots" content="index, follow, max-image-preview:large" />',
      '<link rel="icon" type="image/svg+xml" href="./assets/favicon.svg" />',
      '<meta property="og:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-08-2" />',
      '<meta property="og:image:alt" content="bilig headless spreadsheet engine workbook and TypeScript code preview" />',
      '<meta name="twitter:card" content="summary_large_image" />',
      '<meta name="twitter:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-08-2" />',
      '<meta name="twitter:image:alt" content="bilig headless spreadsheet engine workbook and TypeScript code preview" />',
      '<link rel="stylesheet" href="./assets/fonts.css?v=2026-05-13-1" />',
      '<link rel="stylesheet" href="./assets/site.css?v=2026-05-13-27" />',
      '<link rel="stylesheet" href="./assets/product-demo.css?v=2026-05-13-3" />',
      'Spreadsheet formulas in TypeScript.',
      'When a calculation already makes sense as rows, cells, and formulas',
      'No browser grid, no screenshot parsing.',
      'The examples are runnable `.ts` files.',
      'Write B2, then read the total that changed.',
      '89 starter issues',
      'Docs, examples, and adapters for first PRs.',
      'Revenue.workpaper',
      'One workbook object. Plain TypeScript in, calculated values out.',
      'For product code that still needs spreadsheet math.',
      'bilig is for the part of a service where a workbook model is easier to inspect',
      'Change the cell your workflow owns.',
      'Copy the smoke test into a clean folder.',
      'More TypeScript examples live in',
      'Put the workbook behind the transport you already use.',
      'Node route, queue worker, CLI, or MCP server',
      '46/46 on the current mean-latency suite.',
      'That is the whole claim: WorkPaper is ahead on the comparable mean rows',
      'pnpm workpaper:bench:competitive:check',
      '<dd>46 / 46 mean rows</dd>',
      'Comparable WorkPaper vs HyperFormula workloads only.',
      'packages/benchmarks/baselines/<wbr />workpaper-vs-hyperformula.json',
      'Committed in the repo and checked by CI',
      '<code>lookup-approximate-duplicates</code> p95 1.043x',
      'If approximate lookup with duplicate keys matters to you',
      'How to read the benchmark',
      'Excel gaps',
      'Good first issues',
      'If you are evaluating it, start here.',
      'The useful entry point depends on whether you are testing npm',
      'Take a starter issue that improves the examples.',
      'Docs by job, not by package name.',
      'Open the install path, service recipe, agent adapter, or comparison page that matches the work in front of you.',
      '<h3>Run</h3>',
      '<h3>Build</h3>',
      '<h3>Agents</h3>',
      '<h3>Decide</h3>',
      'Try it on one real calculation.',
    ],
    'docs/index.html',
  )

  requireAllIncludes(
    siteCss,
    [
      "--font-body: 'Bilig Sans'",
      "--font-display: 'Bilig Condensed'",
      "--font-mono: 'Bilig Mono'",
      'grid-template-columns: minmax(0, 455px) minmax(0, 645px);',
      '.hero-notes',
      '.proof-command,',
    ],
    'docs/assets/site.css',
  )

  requireNoIncludes(
    index,
    [
      'bilig-hero-workbook-api.png?v=2026-05-08-2',
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
      'Run the benchmark before you depend on it.',
      'Fast where the benchmark says so. Clear where it does not.',
      'Speed claims are cheap. Run the benchmark.',
      'The benchmark command and JSON artifact are in the repo.',
      'Before you quote the number, run it.',
      'The result, baseline JSON, and p95 caveat are all in the repo.',
      'The benchmark is checked in.',
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
      'Check the Excel gaps',
      'Pick a starter issue',
      'One claim, with the caveat beside it.',
      'The benchmark is public. So are the gaps.',
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
  requireNotIncludes(siteCss, '--font-display: ui-sans-serif, -apple-system', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'border-left: 1px solid rgba(255, 250, 240, 0.16);', 'docs/assets/site.css')
}
