import { describe, expect, it } from 'vitest'
import { compileFormula } from '../compiler.js'
import { translateCompiledFormula } from '../translation.js'

describe('OFFSET dependency binding', () => {
  it('adds possible MATCH-selected target cells as static dependencies', () => {
    const compiled = compileFormula('OFFSET(I27,MATCH($C$24,$B$28:$B$30,0),0)')

    expect(compiled.deps).toEqual(['I27', '$C$24', 'B28:B30', 'I28', 'I29', 'I30'])
  })

  it('keeps absolute MATCH lookup ranges fixed when translating OFFSET formula templates', () => {
    const compiled = compileFormula('OFFSET(I27,MATCH($C$24,$B$28:$B$30,0),0)')
    const translated = translateCompiledFormula(compiled, 0, 1, 'OFFSET(J27,MATCH($C$24,$B$28:$B$30,0),0)').compiled

    expect(translated.deps).toEqual(['J27', '$C$24', 'B28:B30', 'J28', 'J29', 'J30'])
    expect(translated.deps).not.toContain('C28:C30')
    expect(translated.jsPlan).toContainEqual(
      expect.objectContaining({
        opcode: 'lookup-exact-match',
        start: '$B$28',
        end: '$B$30',
        startCol: 1,
        endCol: 1,
      }),
    )
  })
})
