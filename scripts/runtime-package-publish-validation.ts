import { spawnSync } from 'node:child_process'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

const textDecoder = new TextDecoder()

export function validateStagedRuntimePackageVersion(packageName: string, stagedPackageDir: string, expectedVersion: string): void {
  if (packageName !== '@bilig/headless') {
    return
  }
  const versionModuleUrl = pathToFileURL(join(stagedPackageDir, 'dist/work-paper-version.js')).href
  const script = [
    `const versionModule = await import(${JSON.stringify(versionModuleUrl)});`,
    `if (versionModule.WORKPAPER_VERSION !== ${JSON.stringify(expectedVersion)}) {`,
    '  throw new Error(',
    `    ${JSON.stringify('Staged @bilig/headless WorkPaper.version does not match package version: ')} +`,
    `      String(versionModule.WORKPAPER_VERSION) + ${JSON.stringify(' !== ')} + ${JSON.stringify(expectedVersion)},`,
    '  );',
    '}',
  ].join('\n')
  runCommand('node', ['--input-type=module', '--eval', script])
}

function runCommand(command: string, args: string[]): void {
  const result = spawnSync(command, args, {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (result.status !== 0) {
    const stderr = textDecoder.decode(result.stderr).trim()
    throw new Error(`Command failed: ${command} ${args.join(' ')}${stderr ? `\n${stderr}` : ''}`)
  }
}
