import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { docsSiteSources } from './check-docs-discovery-site-sources.ts'

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

function requireHomepageLocalHtmlLinksHaveSources(index: string, docsRoot: string): void {
  const sourceFilesByPage = new Map(docsSiteSources.map(([urlPath, sourceFile]) => [urlPath, sourceFile]))
  const missingSources: string[] = []
  const localHtmlHrefs = Array.from(index.matchAll(/href="([^"]+\.html(?:#[^"]*)?)"/g))
    .map((match) => match[1])
    .filter((href) => href.startsWith('./') || href.startsWith('/bilig/'))

  for (const href of localHtmlHrefs) {
    const page = href
      .replace(/^\.\//, '')
      .replace(/^\/bilig\//, '')
      .split('#')[0]
    const sourceFile = sourceFilesByPage.get(page)

    if (sourceFile === undefined) {
      missingSources.push(`${href}: missing from docsSiteSources`)
      continue
    }

    const sourcePath = join(docsRoot, sourceFile)
    if (!existsSync(sourcePath)) {
      missingSources.push(`${href}: missing source docs/${sourceFile}`)
      continue
    }

    if (page.endsWith('.html') && sourceFile.endsWith('.md')) {
      const sourceText = readFileSync(sourcePath, 'utf8')
      if (!sourceText.startsWith('---\n')) {
        missingSources.push(`${href}: docs/${sourceFile} is missing YAML front matter, so GitHub Pages will not emit ${page}`)
      }
    }
  }

  if (missingSources.length > 0) {
    throw new Error(`docs/index.html has local HTML links without published sources:\n${missingSources.join('\n')}`)
  }
}

export function requireHomepageDiscovery(index: string, siteCss: string, productCss: string, docsRoot: string): void {
  const siteNav = readFileSync(join(docsRoot, 'assets', 'site-nav.js'), 'utf8')

  requireHomepageLocalHtmlLinksHaveSources(index, docsRoot)

  requireAllIncludes(
    index,
    [
      '<link rel="canonical" href="https://proompteng.github.io/bilig/" />',
      '<link rel="sitemap" type="application/xml" href="https://proompteng.github.io/bilig/sitemap.xml" />',
      '<link rel="help" href="https://context7.com/proompteng/bilig" title="Ask Bilig docs on Context7" />',
      '<link rel="alternate" type="text/plain" href="https://proompteng.github.io/bilig/llms.txt" title="llms.txt" />',
      '"@type": "SoftwareSourceCode"',
      '"codeRepository": "https://github.com/proompteng/bilig"',
      '<title>bilig - WorkPaper Runtime for Node</title>',
      'content="Run workbook formulas from Node: edit cells, recalculate, read outputs, and save the result as JSON."',
      'content="WorkPaper runtime, WorkPaper formula readback, workbook formulas Node.js, backend workbook formulas, tool-host workbook automation, MCP tool contracts, MCP WorkPaper tools, JSON WorkPaper persistence, @bilig/workpaper, Node service workbook automation"',
      '<meta name="robots" content="index, follow, max-image-preview:large" />',
      '<link rel="icon" type="image/svg+xml" href="./assets/favicon.svg" />',
      '<meta property="og:title" content="bilig - WorkPaper Runtime for Node" />',
      'property="og:description"',
      'content="Edit workbook inputs, recalculate formulas, read outputs, and persist JSON from Node without spreadsheet UI automation."',
      '<meta property="og:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-15-2" />',
      '<meta property="og:image:alt" content="bilig WorkPaper runtime preview" />',
      '<meta name="twitter:card" content="summary_large_image" />',
      '<meta name="twitter:title" content="bilig - WorkPaper Runtime for Node" />',
      'name="twitter:description"',
      '<meta name="twitter:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-15-2" />',
      '<meta name="twitter:image:alt" content="bilig WorkPaper runtime preview" />',
      '<link rel="stylesheet" href="./assets/fonts.css?v=2026-05-14-1" />',
      '<link rel="stylesheet" href="./assets/site.css?v=2026-05-30-10" />',
      '<link rel="stylesheet" href="./assets/product-demo.css?v=2026-05-15-3" />',
      '<script src="./assets/site-nav.js?v=2026-05-30-11"></script>',
      '<script type="module" src="./assets/hero-scene.js?v=2026-05-15-1"></script>',
      '<p class="eyebrow">WorkPaper runtime</p>',
      '<h1 id="hero-title" class="hero-title">Run workbook rules in Node.</h1>',
      'Bilig gives services and tools a workbook API for pricing models, quote approval, payout checks, import validation, forecasts,',
      'and formula-backed workflow steps.',
      'WorkPaper formula readback before backend code or tool hosts use the result',
      'MCP tool contracts',
      'Use the workbook report when a saved file stays in the loop.',
      './eval-workpaper-service.html',
      './eval-agent-mcp.html',
      './agent-adoption-kit.html',
      'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
      'Edits an input cell, recalculates a formula, saves JSON, restores it, and returns <code>verified: true</code>.',
      'Start with <a href="https://www.npmjs.com/package/@bilig/workpaper">@bilig/workpaper</a>',
      '<a href="#runtime">Runtime</a>',
      '<a href="#install">Install</a>',
      '<a href="#mcp">MCP</a>',
      '<a href="#compatibility">Compatibility</a>',
      '<a href="#docs">Docs</a>',
      '<a class="button primary" href="https://www.npmjs.com/package/@bilig/workpaper">Install WorkPaper</a>',
      '<a class="button" href="./eval-workpaper-service.html">Run WorkPaper</a>',
      '<a class="button" href="./eval-agent-mcp.html">MCP tools</a>',
      '<a class="button" href="./workbook-compatibility-report.html">Workbook report</a>',
      '<a class="button" href="https://github.com/proompteng/bilig">Star on GitHub</a>',
      '<strong>WorkPaper run verified</strong>',
      '<dt>Formula</dt>',
      '<dd><code>Summary!B2</code></dd>',
      '<dt>Before</dt>',
      '<dt>After</dt>',
      '<h2 id="runtime-title">Teams that keep business rules in workbook formulas.</h2>',
      'Pricing, payout, approval, import, and forecast rules often start in cells.',
      "const canonicalDocsRoot = 'https://proompteng.github.io/bilig/'",
      "const isRawFilePreview = window.location.protocol === 'file:'",
      'const isCanonicalGeneratedDocsHost',
      "window.location.hostname === 'proompteng.github.io' && window.location.pathname.startsWith('/bilig/')",
      'document.querySelectorAll',
      'assets/site-nav.js',
      'isGeneratedDocsHref',
      'markdownHrefForLocalSource',
      'canonicalHrefForGeneratedDocs',
      'rewriteLinks',
      'if (isRawFilePreview) {',
      'rewriteLinks(markdownHrefForLocalSource)',
      'if (!isCanonicalGeneratedDocsHost) {',
      'rewriteLinks(canonicalHrefForGeneratedDocs)',
      'new URL(href.slice(2), canonicalDocsRoot).toString()',
      "href.replace(/\\.html(?=$|[?#])/, '.md')",
      '<code>npm install @bilig/workpaper</code>',
      'Put a WorkPaper behind a Node route.',
      '<code>xlsx-recalc --read Summary!B7</code>',
      'bilig-hero-ambient.png?v=2026-05-15-1',
      'Not another spreadsheet UI. Not just a formula engine.',
      'File libraries move workbook bytes. Formula engines calculate formulas.',
      'Keep the business rule readable.',
      'Let the service change the inputs.',
      'Persist the workbook state.',
      'Run the WorkPaper loop first.',
      'No account, no repo clone, no spreadsheet UI.',
      'npm exec --yes --package @bilig/workpaper@latest --',
      'bilig-evaluate --door workpaper-service --json',
      '"door": <span class="str">"workpaper-service"</span>',
      '"afterRestore": <span class="num">38400</span>',
      'Tool host? Do not drive Excel, LibreOffice, Google Sheets, or a browser grid.',
      'bilig-evaluate --door agent-mcp --json',
      'Put the workbook behind the code that owns the workflow.',
      'Route handler, queue worker, CLI, or MCP server: load the workbook',
      'Use Bilig when the workbook is the rulebook.',
      'Put WorkPaper behind a Node API.',
      'Expose tools, not a spreadsheet screen.',
      'Compatibility stays explicit.',
      'WorkPaper is the source of truth for workbook-shaped service logic.',
      'Saved XLSX files still need import and recalculation checks before production use.',
      'Excel oracle harness',
      'Workbook Compatibility Report',
      'XLSX formula recalculation',
      'Formula bug clinic',
      'Compatibility gaps',
      './production-adoption-checklist-headless-workpaper.html',
      'https://github.com/proompteng/bilig/blob/main/SECURITY.md',
      'https://github.com/proompteng/bilig/blob/main/SUPPORT.md',
      'Pick the path that matches the job.',
      'Start with WorkPaper when the service owns the rule.',
      'Install the WorkPaper package.',
      'Review saved-workbook compatibility before import.',
      'Run a WorkPaper in a service.',
      'Give tool hosts a workbook API.',
      'Take a starter issue that improves the examples.',
      'Start with code/test picks, example tasks, adapters, or focused docs coverage.',
      '<h3>Run</h3>',
      '@bilig/workpaper npm package',
      '<h3>Build</h3>',
      '<h3>Tool hosts</h3>',
      'MCP and framework tool chooser',
      'workbook runtime intent API',
      'Context7 indexed docs',
      '<h3>Decide</h3>',
      'Try it on the workbook rule your code already depends on.',
      'Run one WorkPaper through Node:',
    ],
    'docs/index.html',
  )

  requireAllIncludes(
    siteNav,
    [
      'syncTopbarHeight',
      'ResizeObserver',
      'correctCurrentHashScroll',
      "window.addEventListener('load', correctCurrentHashScroll)",
      'await document.fonts.ready',
    ],
    'docs/assets/site-nav.js',
  )

  requireAllIncludes(
    siteCss,
    [
      "--font-body: 'Bilig Sans'",
      "--font-display: 'Bilig Sans'",
      "--font-mono: 'Bilig Mono'",
      'grid-template-columns: minmax(440px, 0.78fr) minmax(520px, 1fr);',
      'grid-template-columns: minmax(360px, 0.92fr) minmax(0, 1.08fr);',
      'min-width: 0;',
      'section[id] {\n  scroll-margin-top:',
      'scroll-padding-top: calc(var(--topbar-height) + 18px);',
      '@media (max-width: 389px) {\n  :root {\n    --topbar-height: 126px;',
      'flex-wrap: nowrap;',
      'overflow-x: auto;',
      '.nav a {\n  display: inline-flex;',
      '.hero-command',
      '.hero-note',
      'overflow-wrap: anywhere;',
      '@media (max-width: 1180px) {',
      'grid-template-columns: repeat(3, minmax(0, 1fr));',
      '.link-stack .path-link em {\n    grid-column: auto;',
      '.proof-result {\n  display: grid;\n  grid-template-columns: minmax(0, 1fr);',
      'width: fit-content;',
      '.proof-metric {\n  display: grid;',
      '.proof-result p {\n  min-width: 0;',
      '.proof-facts',
      '.page-main',
      '.markdown-page',
    ],
    'docs/assets/site.css',
  )

  requireAllIncludes(
    productCss,
    [
      '.hero-media {\n  position: relative;',
      '.hero-ambient,',
      '.hero-canvas,',
      '.hero-fallback {',
      '.action-report-card {',
      '.report-card-top strong {',
      '.hero-media::after',
      '.action-report-card dl {',
    ],
    'docs/assets/product-demo.css',
  )

  requireNoIncludes(
    index,
    [
      'bilig-hero-workbook-api.png?v=2026-05-08-2',
      'bilig-hero-workbook-api.png?v=2026-05-14-4',
      'bilig-hero-workbook-api.png?v=2026-05-14-6',
      'localPreviewHostnames',
      'isLocalHttpPreview',
      'Run the smoke test',
      '<h1 id="hero-title" class="hero-title" translate="no">bilig</h1>',
      'Headless workbooks for TypeScript services.',
      'Keep the spreadsheet shape when cells and formulas are still the clearest way to explain the calculation.',
      'Run the maintained eval before reading more docs.',
      'The smoke test changes one input',
      'TypeScript smoke test',
      '>eval.ts<',
      'curl -fsSLo eval.ts',
      'npx tsx eval.ts',
      'npm-only smoke test',
      'npm smoke test feedback',
      'If the smoke test matches your workflow',
      'TypeScript package',
      'Examples are real <code>.ts</code> files with public imports.',
      'The npm package ships a stdio server for tool integrations.',
      'github-social-preview.png?v=2026-05-08-2',
      'github-social-preview.png?v=2026-05-14-5',
      "window.location.hostname === 'localhost'",
      "window.location.hostname === '127.0.0.1'",
      "window.location.hostname === '::1'",
      'localPreviewHosts',
      'isLocalStaticPreview',
      'const localHtmlLinks',
      'const probeAndRewriteMissingLocalHtml',
      'publicDocsHostnames',
      'isHostedSourcePreview',
      'isSourcePreview',
      'Run workbook formulas in TypeScript.',
      'Use a WorkPaper when a calculation is easier to review as cells and formulas.',
      'Runnable .ts examples',
      'The docs link to real TypeScript files.',
      'Set an input cell and read the recalculated total.',
      'no browser required',
      'Use it when the spreadsheet is the model.',
      'Pricing tables, import checks, payouts, and budget rules often make the most sense as rows and formulas.',
      'Run the smoke test in a blank Node project.',
      'Put the workbook where the calculation happens.',
      'It does not claim every workbook will be faster.',
      'Pick the path that matches your codebase.',
      'Docs by job, not package name.',
      'Spreadsheet logic for TypeScript services.',
      'When a pricing rule, budget check, or payout model is easiest to express as cells and formulas',
      'The docs link to runnable TypeScript files.',
      'Edit, then read',
      'Change one input and inspect the dependent total.',
      'Small docs, examples, adapters, and tests are open.',
      'Run the TypeScript example',
      'Build a workbook in Node, change inputs through code',
      'For pricing, budgets, imports, payouts, and tools that still work like a workbook.',
      'Build the sheet, write the input,',
      'Use it for pricing calculators, budget checks, imports, payouts, and tools when the formula should stay in a workbook.',
      'Build the sheet, edit the input, read the cell, save JSON.',
      'Examples are `.ts` files with real imports.',
      'Change B2 and read the dependent total.',
      'Small first PRs are kept open.',
      'The workbook is data plus formulas. The code edits it and reads the result.',
      'Workbook formulas for TypeScript services.',
      'Keep a calculation as cells and formulas when that is the clearest model.',
      'Copy the `.ts` files and run them.',
      'Change an input cell and read the dependent total.',
      'One workbook object. TypeScript changes it and reads the value it produced.',
      'Use a workbook when formulas are the clearest source of truth.',
      'Use a workbook when the formula is the thing you ship.',
      'If a model is already easier to review in rows and formulas',
      'Change the input.',
      'Run this before reading more.',
      'The smoke test changes one customer count',
      'Use it from the place that already owns the workflow.',
      'Node route, queue worker, CLI, or MCP server: load the workbook',
      'The number is public. So is the caveat.',
      'not a blanket',
      'promise for every workbook.',
      'The approximate-lookup duplicate case is slower',
      'Browser grid performance, Excel compatibility, and your own workbook shape.',
      'It covers shared headless formula workloads,',
      'compatibility, browser rendering, or your own workbook.',
      'No hidden browser, no screen scraping',
      'bilig is ahead on this checked-in WorkPaper suite.',
      'good news for this suite, not a blank check',
      'Before you quote the number, run it.',
      'Run the command, read the JSON, and check the slower row before you use the number in your own decision.',
      'Coverage and caveats',
      'Excel behavior not covered yet',
      'Small TypeScript tasks',
      'The speed claim is deliberately narrow',
      'Run it yourself.',
      'That slower row stays visible because it matters if your workload looks like it.',
      'Rerun the command before using the number in your own docs.',
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
      'Pick a starter issue',
      'One claim, with the caveat beside it.',
      'Comparable headless workloads',
      'No UI-performance claim. No Excel-compatibility claim.',
      'Read what the numbers mean',
      'Check the Excel gaps',
      'Good first issues',
      'Public project signals',
      '<strong>40 starter tasks</strong>',
      '90 starter issues',
      '105 starter issues',
      '<strong>0.13.9</strong>',
      'Read those before you depend on the package.',
      'launch essay',
      '<a class="button" href="./eval-workpaper-service.html">Build an API</a>',
      '<a class="button" href="./eval-agent-mcp.html">Use with agents</a>',
      '<title>bilig - Formulas for TypeScript</title>',
      '<h1 id="hero-title" class="hero-title">Formulas for TypeScript.</h1>',
      '<title>bilig - Formula workbooks for Node services</title>',
      '<h1 id="hero-title" class="hero-title">Formula workbooks for Node services.</h1>',
      '<strong>Quote approved</strong>',
      '<h2 id="market-title">Teams that still audit rules in cells.</h2>',
      '"https://dev.to/gregkonush/why-agents-need-workbook-apis-instead-of-spreadsheet-screenshots-3d61"',
    ],
    'docs/index.html',
  )

  requireNotIncludes(siteCss, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'grid-template-columns: minmax(150px, 0.34fr) minmax(0, 1fr);', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'grid-template-columns: minmax(210px, max-content) minmax(0, 1fr);', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'border-left: 1px solid rgba(255, 250, 240, 0.16);', 'docs/assets/site.css')
}
