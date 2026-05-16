import { describe, expect, it } from 'vitest'
import {
  requireDocumentIncludes,
  requireDocumentNotIncludes,
  requireDocumentsInclude,
  requireDocumentsNotInclude,
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
      { path: 'packages/headless/README.md', content: 'safe relative link' },
      { path: 'docs/llms.txt', content: 'safe public link' },
    ]

    expect(() => requireDocumentsNotInclude(documents, ['../../docs'])).not.toThrow()
    expect(() => requireDocumentsNotInclude(documents, ['safe public link'])).toThrow('docs/llms.txt must not include safe public link')
    expect(() => requireDocumentNotIncludes(documents[0], ['safe relative link'])).toThrow(
      'packages/headless/README.md must not include safe relative link',
    )
  })
})
