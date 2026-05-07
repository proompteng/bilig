import { createRequire } from 'node:module'

const requirePackageJson = createRequire(import.meta.url)
const packageManifest: unknown = requirePackageJson('../package.json')

export const WORKPAPER_VERSION = readWorkPaperPackageVersion(packageManifest)

export function readWorkPaperPackageVersion(manifest: unknown): string {
  if (typeof manifest !== 'object' || manifest === null) {
    throw new Error('Expected @bilig/headless package.json to contain a non-empty version string')
  }
  const version = Reflect.get(manifest, 'version')
  if (typeof version !== 'string' || version.trim().length === 0) {
    throw new Error('Expected @bilig/headless package.json to contain a non-empty version string')
  }
  return version
}
