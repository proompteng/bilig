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
      '<title>bilig - Spreadsheet Formulas for Node.js Services and Agent Tools</title>',
      '<meta name="robots" content="index, follow, max-image-preview:large" />',
      '<link rel="icon" type="image/svg+xml" href="./assets/favicon.svg" />',
      '<link rel="stylesheet" href="./assets/fonts.css?v=2026-05-13-1" />',
      '<link rel="stylesheet" href="./assets/site.css?v=2026-05-13-24" />',
      '<link rel="stylesheet" href="./assets/product-demo.css?v=2026-05-13-1" />',
      'Revenue.workpaper',
      'When a feature still depends on formulas, build a WorkPaper in TypeScript',
      '<span>Install</span>',
      '<strong>The headless package is public on npm.</strong>',
      '<strong>Run the benchmark from the repo.</strong>',
      '<span>Contribute</span>',
      '<strong>88 small issues are open for first pull requests.</strong>',
      'Speed claims are cheap. Run the benchmark.',
      'The current run is checked in: 46 of 46 comparable mean rows ahead of HyperFormula',
      'plus the p95 row where bilig loses.',
      'Comparable rows only, measured by mean latency.',
      'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
      'pnpm workpaper:bench:competitive:check',
      '<dd>46 / 46 ahead</dd>',
      'The JSON changes in review when the benchmark moves.',
      'p95 1.043x',
      'test locally before quoting the headline number.',
      '<code>lookup-approximate-duplicates</code>',
      'Coverage and caveats',
      'Small TypeScript tasks',
      'Start from the job you need done.',
      'Each route points to a runnable TypeScript example',
      'Install from npm in a blank folder.',
      'Put a WorkPaper behind a Node endpoint.',
      'Expose workbook reads and writes as tools.',
      'Check the alternatives before switching.',
      'Read the Excel gaps before importing real workbooks.',
      'Take a small issue that improves the public examples.',
      'Open the page for the thing you are building.',
      'without reading a launch essay first.',
      '<h3>Run</h3>',
      '<h3>Build</h3>',
      '<h3>Agents</h3>',
      '<h3>Decide</h3>',
      'Install it. Change one cell. Check the total.',
    ],
    'docs/index.html',
  )

  requireAllIncludes(
    siteCss,
    [
      '--font-display: ui-sans-serif, -apple-system',
      'max-width: 470px;',
      'grid-template-columns: minmax(100px, 0.34fr) minmax(0, 1fr);',
      'border-bottom: 1px solid rgba(255, 250, 240, 0.14);',
    ],
    'docs/assets/site.css',
  )

  requireNoIncludes(
    index,
    [
      'bilig-hero-workbook-api.png?v=2026-05-08-2',
      'Run the TypeScript example',
      'Build a workbook in Node, change inputs through code',
      'Run the benchmark before you depend on it.',
      'Fast where the benchmark says so. Clear where it does not.',
      'The benchmark command and JSON artifact are in the repo.',
      'The benchmark is checked in.',
      'The current artifact shows 46 of 46 comparable mean rows ahead of HyperFormula',
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
      'One claim, with the caveat beside it.',
      'The benchmark is public. So are the gaps.',
      'Public project signals',
      '<strong>40 starter tasks</strong>',
      '<strong>0.13.9</strong>',
    ],
    'docs/index.html',
  )

  requireNotIncludes(siteCss, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'border-left: 1px solid rgba(255, 250, 240, 0.16);', 'docs/assets/site.css')
}
