import { readFile, stat } from 'node:fs/promises'

export function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

export function requireNotIncludes(haystack: string, needle: string, context: string): void {
  if (haystack.includes(needle)) {
    throw new Error(`${context} must not include ${needle}`)
  }
}

export async function requirePublishedSource(path: string): Promise<void> {
  await requireFile(path)

  if (!path.endsWith('.md')) {
    return
  }

  const frontMatter = getFrontMatter(await readFile(path, 'utf8'))
  if (frontMatter !== undefined && /^published:\s*false\s*$/m.test(frontMatter)) {
    throw new Error(`${path} is listed in the sitemap but has published: false`)
  }
}

export function extractSitemapUrls(sitemap: string): string[] {
  return [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map((match) => match[1] ?? '')
}

export function extractNpmRunScripts(readme: string): string[] {
  const scripts = new Set<string>()

  // Match the command form used throughout the headless example README,
  // including optional npm flags such as `npm run --silent script-name`.
  for (const match of readme.matchAll(/\bnpm\s+run(?:\s+--[\w-]+)*\s+([\w:-]+)/g)) {
    const script = match[1]
    if (script !== undefined) {
      scripts.add(script)
    }
  }

  return [...scripts].toSorted()
}

export function requirePackageKeywords(packageJson: string, requiredKeywords: readonly string[], context: string): void {
  const manifest: unknown = JSON.parse(packageJson)

  if (typeof manifest !== 'object' || manifest === null || !('keywords' in manifest)) {
    throw new Error(`${context} is missing a keywords array`)
  }

  const { keywords } = manifest
  if (!Array.isArray(keywords) || !keywords.every((keyword) => typeof keyword === 'string')) {
    throw new Error(`${context} keywords must be an array of strings`)
  }

  for (const requiredKeyword of requiredKeywords) {
    if (!keywords.includes(requiredKeyword)) {
      throw new Error(`${context} is missing discovery keyword: ${requiredKeyword}`)
    }
  }
}

export function requireDocumentedScriptsExist(readme: string, packageJson: string, context: string): void {
  const scripts = getPackageScripts(packageJson, 'examples/headless-workpaper/package.json')

  for (const documentedScript of extractNpmRunScripts(readme)) {
    if (!(documentedScript in scripts)) {
      throw new Error(`${context} documents missing package.json script: npm run ${documentedScript}`)
    }
  }
}

export function requireNoUnsupportedGoogleSheetsTenXClaims(scorecardJson: string, publicSources: Record<string, string>): void {
  const scorecard = JSON.parse(scorecardJson) as unknown
  const googleSheetsGatePassed =
    typeof scorecard === 'object' &&
    scorecard !== null &&
    'overallGoogleSheets10xStatus' in scorecard &&
    typeof scorecard.overallGoogleSheets10xStatus === 'object' &&
    scorecard.overallGoogleSheets10xStatus !== null &&
    'passed' in scorecard.overallGoogleSheets10xStatus &&
    scorecard.overallGoogleSheets10xStatus.passed === true
  if (googleSheetsGatePassed) {
    return
  }

  const forbiddenClaims = [
    '10x faster than Google Sheets',
    '10x better than Google Sheets',
    '10x more responsive than Google Sheets',
    '10x Google Sheets',
    'beats Google Sheets by 10x',
    'beat Google Sheets by 10x',
  ]
  for (const [sourceName, source] of Object.entries(publicSources)) {
    const normalizedSource = source.toLowerCase()
    const match = forbiddenClaims.find((claim) => normalizedSource.includes(claim.toLowerCase()))
    if (match) {
      throw new Error(
        `${sourceName} contains unsupported broad Google Sheets 10x wording "${match}" while overallGoogleSheets10xStatus is not passed`,
      )
    }
  }
}

export async function requireFile(path: string): Promise<void> {
  const info = await stat(path)
  if (!info.isFile()) {
    throw new Error(`${path} is not a file`)
  }
}

function getFrontMatter(content: string): string | undefined {
  if (!content.startsWith('---\n')) {
    return undefined
  }

  const end = content.indexOf('\n---', 4)
  if (end === -1) {
    return undefined
  }

  return content.slice(4, end)
}

function getPackageScripts(packageJson: string, context: string): Record<string, unknown> {
  const manifest: unknown = JSON.parse(packageJson)

  if (typeof manifest !== 'object' || manifest === null || !('scripts' in manifest)) {
    throw new Error(`${context} is missing a scripts object`)
  }

  const { scripts } = manifest
  if (typeof scripts !== 'object' || scripts === null || Array.isArray(scripts)) {
    throw new Error(`${context} scripts must be an object`)
  }

  return scripts
}
