function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

export function requireXlsxCorpusVerifierDiscovery(content: string): void {
  for (const required of [
    'not on a vague',
    'Run It Against Your Files',
    'Run The Excel Oracle Harness',
    'pnpm workpaper:xlsx-oracle -- prepare-oracle /path/to/workbooks "$OUT"',
    'cache-diagnostic.json',
    'excel-oracle-report.json',
    'missing_excel_oracle',
    'Put It In CI',
    'pnpm workpaper:xlsx-corpus:check -- /path/to/workbooks',
    'Turn A Miss Into A Contribution',
    'https://github.com/proompteng/bilig/issues/new/choose',
    'https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only',
    'https://github.com/proompteng/bilig/blob/main/packages/headless/README.md',
  ] as const) {
    requireIncludes(content, required, 'docs/xlsx-corpus-verifier-walkthrough.md')
  }
}
