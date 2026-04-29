import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getTextBuiltin } from '../builtins/text.js'

describe('text core builtins', () => {
  it('should support core cleanup and composition helpers', () => {
    // Arrange
    const CLEAN = getTextBuiltin('CLEAN')!
    const CONCAT = getTextBuiltin('CONCAT')!
    const CONCATENATE = getTextBuiltin('CONCATENATE')!
    const PROPER = getTextBuiltin('PROPER')!
    const EXACT = getTextBuiltin('EXACT')!
    const TRIM = getTextBuiltin('TRIM')!

    // Act
    const cleaned = CLEAN(text('a\u0001b\u007fc'))
    const concatenated = CONCAT(text('a'), number(1), text('b'))
    const concatenatedLegacy = CONCATENATE(text('x'), text('y'))
    const proper = PROPER(text('hELLO, wORLD'))
    const exact = EXACT(text('Alpha'), text('Alpha'))
    const trimmed = TRIM(text('  alpha   beta  '))

    // Assert
    expect(cleaned).toEqual(text('abc'))
    expect(concatenated).toEqual(text('a1b'))
    expect(concatenatedLegacy).toEqual(text('xy'))
    expect(proper).toEqual(text('Hello, World'))
    expect(exact).toEqual(bool(true))
    expect(trimmed).toEqual(text('alpha beta'))
  })

  it('should support localization and baht text helpers', () => {
    // Arrange
    const ASC = getTextBuiltin('ASC')!
    const JIS = getTextBuiltin('JIS')!
    const DBCS = getTextBuiltin('DBCS')!
    const BAHTTEXT = getTextBuiltin('BAHTTEXT')!
    const PHONETIC = getTextBuiltin('PHONETIC')!

    // Act
    const asc = ASC(text('ＡＢＣ　１２３'))
    const jis = JIS(text('ABC 123'))
    const dbcs = DBCS(text('ｶﾞｷﾞｸﾞ'))
    const baht = BAHTTEXT(number(1234))
    const phonetic = PHONETIC(text('カタカナ'))

    // Assert
    expect(asc).toEqual(text('ABC 123'))
    expect(jis).toEqual(text('ＡＢＣ　１２３'))
    expect(dbcs).toEqual(text('ガギグ'))
    expect(baht).toEqual(text('หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน'))
    expect(phonetic).toEqual(text('カタカナ'))
  })

  it('should return value errors for missing required args and keep explicit errors', () => {
    // Arrange
    const CONCAT = getTextBuiltin('CONCAT')!
    const CLEAN = getTextBuiltin('CLEAN')!
    const ASC = getTextBuiltin('ASC')!
    const JIS = getTextBuiltin('JIS')!
    const DBCS = getTextBuiltin('DBCS')!
    const PHONETIC = getTextBuiltin('PHONETIC')!
    const BAHTTEXT = getTextBuiltin('BAHTTEXT')!
    const CONCATENATE = getTextBuiltin('CONCATENATE')!
    const PROPER = getTextBuiltin('PROPER')!
    const EXACT = getTextBuiltin('EXACT')!
    const TRIM = getTextBuiltin('TRIM')!

    // Act
    const concatMissing = CONCAT()
    const cleanMissing = CLEAN()
    const ascMissing = ASC()
    const jisMissing = JIS()
    const dbcsMissing = DBCS()
    const phoneticMissing = PHONETIC()
    const bahtBad = BAHTTEXT(text('bad'))
    const concatenateMissing = CONCATENATE()
    const properMissing = PROPER()
    const exactMissing = EXACT(text('left'))
    const trimMissing = TRIM()
    const concatError = CONCAT(err(ErrorCode.Ref), text('x'))
    const cleanError = CLEAN(err(ErrorCode.Ref))
    const ascError = ASC(err(ErrorCode.Ref))
    const jisError = JIS(err(ErrorCode.Ref))
    const dbcsError = DBCS(err(ErrorCode.Ref))
    const phoneticError = PHONETIC(err(ErrorCode.NA))
    const bahtError = BAHTTEXT(err(ErrorCode.Ref))
    const concatenateError = CONCATENATE(err(ErrorCode.Ref), text('x'))
    const properError = PROPER(err(ErrorCode.Ref))
    const exactError = EXACT(err(ErrorCode.Ref), text('x'))
    const trimError = TRIM(err(ErrorCode.Ref))

    // Assert
    expect(concatMissing).toEqual(valueError())
    expect(cleanMissing).toEqual(valueError())
    expect(ascMissing).toEqual(valueError())
    expect(jisMissing).toEqual(valueError())
    expect(dbcsMissing).toEqual(valueError())
    expect(phoneticMissing).toEqual(valueError())
    expect(bahtBad).toEqual(valueError())
    expect(concatenateMissing).toEqual(valueError())
    expect(properMissing).toEqual(valueError())
    expect(exactMissing).toEqual(valueError())
    expect(trimMissing).toEqual(valueError())
    expect(concatError).toEqual(err(ErrorCode.Ref))
    expect(cleanError).toEqual(err(ErrorCode.Ref))
    expect(ascError).toEqual(err(ErrorCode.Ref))
    expect(jisError).toEqual(err(ErrorCode.Ref))
    expect(dbcsError).toEqual(err(ErrorCode.Ref))
    expect(phoneticError).toEqual(err(ErrorCode.NA))
    expect(bahtError).toEqual(err(ErrorCode.Ref))
    expect(concatenateError).toEqual(err(ErrorCode.Ref))
    expect(properError).toEqual(err(ErrorCode.Ref))
    expect(exactError).toEqual(err(ErrorCode.Ref))
    expect(trimError).toEqual(err(ErrorCode.Ref))
  })

  it('should cover unicode width conversion and baht text edge branches', () => {
    // Arrange
    const ASC = getTextBuiltin('ASC')!
    const JIS = getTextBuiltin('JIS')!
    const BAHTTEXT = getTextBuiltin('BAHTTEXT')!

    // Act
    const fullWidth = JIS(text('A 😀 ｶﾞ｡'))
    const halfWidth = ASC(text('Ａ　😀　ガ。'))
    const zeroBaht = BAHTTEXT(number(0))
    const negativeSatang = BAHTTEXT(number(-21.25))
    const recursiveMillion = BAHTTEXT(number(1_234_567.89))
    const infinite = BAHTTEXT(number(Number.POSITIVE_INFINITY))
    const tooLarge = BAHTTEXT(number(10_000_000_000_000))

    // Assert
    expect(fullWidth).toEqual(text('Ａ　😀　ガ。'))
    expect(halfWidth).toEqual(text('A 😀 ｶﾞ｡'))
    expect(zeroBaht).toEqual(text('ศูนย์บาทถ้วน'))
    expect(negativeSatang).toEqual(text('ลบยี่สิบเอ็ดบาทยี่สิบห้าสตางค์'))
    expect(recursiveMillion).toEqual(text('หนึ่งล้านสองแสนสามหมื่นสี่พันห้าร้อยหกสิบเจ็ดบาทแปดสิบเก้าสตางค์'))
    expect(infinite).toEqual(valueError())
    expect(tooLarge).toEqual(valueError())
  })
})

// Helpers
function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function bool(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function valueError(): CellValue {
  return err(ErrorCode.Value)
}
