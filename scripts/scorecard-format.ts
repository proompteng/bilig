import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

export function formatJsonForRepo(args: {
  readonly rootDir: string
  readonly serializedJson: string
  readonly tempPrefix: string
}): string {
  const tempDir = mkdtempSync(join(tmpdir(), `${args.tempPrefix}-`))
  const tempFilePath = join(tempDir, 'scorecard.json')
  writeFileSync(tempFilePath, args.serializedJson)
  const oxfmtPath = join(args.rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(`Unable to format generated scorecard: ${new TextDecoder().decode(formatResult.stderr).trim()}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}
