import { readdir, writeFile } from 'node:fs/promises'
import { extname, join } from 'node:path'
import { versionedStaticReferenceExtensions, versionedStaticReferenceRoots } from './agent-discovery-constants.ts'
import { readTextFileIfExists } from './read-if-exists.ts'

interface StaticReferenceSyncOptions {
  readonly checkOnly: boolean
  readonly headlessPackageSpec: string
  readonly headlessPackageVersion: string
  readonly mcpbReleaseAssetUrl: string
  readonly mcpbReleaseChecksumUrl: string
  readonly repoRoot: string
  readonly workbookPackageSpec: string
  readonly workpaperPackageSpec: string
}

async function collectVersionedStaticReferenceFiles(repoRoot: string, relativePath: string): Promise<string[]> {
  const absolutePath = join(repoRoot, relativePath)
  try {
    const entries = await readdir(absolutePath, { withFileTypes: true })
    const nested = await Promise.all(
      entries.map((entry) => {
        const entryRelativePath = `${relativePath}/${entry.name}`
        return entry.isDirectory() ? collectVersionedStaticReferenceFiles(repoRoot, entryRelativePath) : [entryRelativePath]
      }),
    )
    return nested.flat().filter((entryRelativePath) => versionedStaticReferenceExtensions.has(extname(entryRelativePath)))
  } catch (error) {
    if (error instanceof Error && 'code' in error && error.code === 'ENOTDIR') {
      return versionedStaticReferenceExtensions.has(extname(relativePath)) ? [relativePath] : []
    }
    if (error instanceof Error && 'code' in error && error.code === 'ENOENT') {
      return []
    }
    throw error
  }
}

function syncVersionedStaticReferenceLine(line: string, options: StaticReferenceSyncOptions): string {
  const stableSemverPattern = String.raw`\d+\.\d+\.\d+`
  const {
    headlessPackageSpec,
    headlessPackageVersion,
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
    workbookPackageSpec,
    workpaperPackageSpec,
  } = options
  if (line.includes('Current checked npm footprint for')) {
    return line.replace(new RegExp(`@bilig/workpaper@${stableSemverPattern}`, 'g'), headlessPackageSpec)
  }
  return line
    .replace(new RegExp(`(npm exec --package )@bilig/workpaper@${stableSemverPattern}`, 'g'), `$1${workpaperPackageSpec}`)
    .replace(new RegExp(`(npm install )@bilig/workbook@${stableSemverPattern}`, 'g'), `$1${workbookPackageSpec}`)
    .replace(new RegExp(`(npm install )@bilig/workpaper@${stableSemverPattern}`, 'g'), `$1${workpaperPackageSpec}`)
    .replace(new RegExp(`("--package",\\s*")@bilig/workpaper@${stableSemverPattern}(")`, 'g'), `$1${workpaperPackageSpec}$2`)
    .replace(new RegExp(`('--package',\\s*')@bilig/workpaper@${stableSemverPattern}(')`, 'g'), `$1${workpaperPackageSpec}$2`)
    .replace(new RegExp(`^(\\s*)"@bilig/workpaper@${stableSemverPattern}"(,?\\s*)$`, 'g'), `$1"${workpaperPackageSpec}"$2`)
    .replace(new RegExp(`^(\\s*)'@bilig/workpaper@${stableSemverPattern}'(,?\\s*)$`, 'g'), `$1'${workpaperPackageSpec}'$2`)
    .replace(
      new RegExp(`(Current checked npm footprint for \`)@bilig/workpaper@${stableSemverPattern}(\`)`, 'g'),
      `$1${headlessPackageSpec}$2`,
    )
    .replace(new RegExp(`npm latest is \`${stableSemverPattern}\``, 'g'), `npm latest is \`${headlessPackageVersion}\``)
    .replace(new RegExp(`npm latest is \`@bilig/workpaper@${stableSemverPattern}\``, 'g'), `npm latest is \`${workpaperPackageSpec}\``)
    .replace(new RegExp(`npm latest \`@bilig/workpaper@${stableSemverPattern}\``, 'g'), `npm latest \`${workpaperPackageSpec}\``)
    .replace(new RegExp(`\`@bilig/workpaper@${stableSemverPattern}\``, 'g'), `\`${workpaperPackageSpec}\``)
    .replace(new RegExp(`\`@bilig/workbook@${stableSemverPattern}\``, 'g'), `\`${workbookPackageSpec}\``)
    .replace(
      new RegExp(`io\\.github\\.proompteng/bilig-workpaper@${stableSemverPattern}`, 'g'),
      `io.github.proompteng/bilig-workpaper@${headlessPackageVersion}`,
    )
    .replace(new RegExp(`(now points reviewers at \`)@bilig/workpaper@${stableSemverPattern}(\`)`, 'g'), `$1${workpaperPackageSpec}$2`)
    .replace(
      new RegExp(
        `https://github\\.com/proompteng/bilig/releases/download/libraries-v${stableSemverPattern}/bilig-workpaper\\.mcpb\\.sha256`,
        'g',
      ),
      mcpbReleaseChecksumUrl,
    )
    .replace(
      new RegExp(`https://github\\.com/proompteng/bilig/releases/download/libraries-v${stableSemverPattern}/bilig-workpaper\\.mcpb`, 'g'),
      mcpbReleaseAssetUrl,
    )
}

function syncVersionedStaticReferenceContent(content: string, options: StaticReferenceSyncOptions): string {
  return content
    .split(/(?<=\n)/)
    .map((line) => syncVersionedStaticReferenceLine(line, options))
    .join('')
}

export async function syncVersionedStaticReferences(options: StaticReferenceSyncOptions): Promise<string[]> {
  const targetFiles = [
    ...new Set(
      (
        await Promise.all(
          versionedStaticReferenceRoots.map((relativePath) => collectVersionedStaticReferenceFiles(options.repoRoot, relativePath)),
        )
      ).flat(),
    ),
  ].toSorted()
  const staleTargets: string[] = []

  await Promise.all(
    targetFiles.map(async (relativePath) => {
      const absolutePath = join(options.repoRoot, relativePath)
      const existing = await readTextFileIfExists(absolutePath)
      if (existing === undefined) {
        return
      }

      const nextContent = syncVersionedStaticReferenceContent(existing, options)
      if (nextContent === existing) {
        return
      }

      if (options.checkOnly) {
        staleTargets.push(relativePath)
        return
      }

      await writeFile(absolutePath, nextContent)
    }),
  )

  return staleTargets.toSorted()
}
