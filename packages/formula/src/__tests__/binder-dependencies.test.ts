import { describe, expect, it } from 'vitest'
import { parseFormula } from '../parser.js'
import { collectFormulaDependencyMetadata } from '../binder-dependencies.js'

describe('formula binder dependency metadata', () => {
  it('separates lexical names from external symbols while expanding bounded OFFSET targets', () => {
    const metadata = collectFormulaDependencyMetadata(
      parseFormula('LET(local, A1, LAMBDA(param, OFFSET(I27, MATCH($C$24, $B$28:$B$30, 0), 0) + param + local + ExternalName)(B1))'),
    )

    expect(metadata.deps).toEqual(['A1', 'I27', '$C$24', 'B28:B30', 'I28', 'I29', 'I30', 'B1'])
    expect(metadata.symbolicNames).toEqual(['ExternalName'])
    expect(metadata.symbolicTables).toEqual([])
    expect(metadata.symbolicSpills).toEqual([])
  })
})
