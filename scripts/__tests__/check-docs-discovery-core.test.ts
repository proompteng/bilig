import { describe, expect, it } from 'vitest'
import {
  extractNpmRunScripts,
  requireDocumentIncludes,
  requireDocumentNotIncludes,
  requireDocumentedScriptsExist,
  requireDocumentsInclude,
  requireDocumentsNotInclude,
  requirePackageScriptsDocumented,
} from '../check-docs-discovery-core.ts'

describe('docs discovery core guards', () => {
  it('requires every requested proof string in every document', () => {
    const documents = [
      { path: 'README.md', content: 'alpha beta gamma' },
      { path: 'docs/index.html', content: 'beta gamma delta' },
    ]

    expect(() => requireDocumentsInclude(documents, ['beta', 'gamma'])).not.toThrow()
    expect(() => requireDocumentsInclude(documents, ['alpha'])).toThrow('docs/index.html is missing alpha')
    expect(() => requireDocumentIncludes(documents[0], ['delta'])).toThrow('README.md is missing delta')
  })

  it('rejects forbidden proof strings with the owning document path', () => {
    const documents = [
      { path: 'packages/workpaper/README.md', content: 'safe relative link' },
      { path: 'docs/llms.txt', content: 'safe public link' },
    ]

    expect(() => requireDocumentsNotInclude(documents, ['../../docs'])).not.toThrow()
    expect(() => requireDocumentsNotInclude(documents, ['safe public link'])).toThrow('docs/llms.txt must not include safe public link')
    expect(() => requireDocumentNotIncludes(documents[0], ['safe relative link'])).toThrow(
      'packages/workpaper/README.md must not include safe relative link',
    )
  })

  it('extracts concrete example script commands from npm and pnpm forms', () => {
    const readme = [
      '`npm start`',
      '`npm run agent:verify`',
      '`npm run --silent markdown-report`',
      '`pnpm --dir examples/headless-workpaper run csv-shaped`',
      '`pnpm --dir examples/headless-workpaper run <script>`',
    ].join('\n')

    expect(extractNpmRunScripts(readme)).toEqual(['agent:verify', 'csv-shaped', 'markdown-report', 'start'])
  })

  it('guards both documented scripts and package scripts in the headless example README', () => {
    const scripts = {
      'agent:verify': 'node agent-writeback-verification.ts',
      'markdown-report': 'node markdown-report.ts',
      start: 'node revenue-plan.ts',
      typecheck: 'tsc --noEmit',
    }
    const packageJson = JSON.stringify({ scripts })
    const readme = [
      '| Quick revenue workbook | `npm start` | formulas |',
      '| Agent writeback check | `npm run agent:verify` | readback |',
      '## Markdown Report Output',
      '```sh',
      'npm run markdown-report',
      '```',
    ].join('\n')

    expect(() => requireDocumentedScriptsExist(readme, packageJson, 'README.md')).not.toThrow()
    expect(() =>
      requirePackageScriptsDocumented(readme, packageJson, 'README.md', {
        ignoredScripts: ['typecheck'],
      }),
    ).not.toThrow()

    const packageWithUndocumentedScript = JSON.stringify({
      scripts: {
        ...scripts,
        'json-records': 'node json-records-input.ts',
      },
    })

    expect(() => requirePackageScriptsDocumented(readme, packageWithUndocumentedScript, 'README.md')).toThrow(
      'README.md is missing README coverage for package.json script: json-records',
    )
    expect(() => requireDocumentedScriptsExist('npm run missing-example', packageJson, 'README.md')).toThrow(
      'README.md documents missing package.json script: npm run missing-example',
    )
  })
})
