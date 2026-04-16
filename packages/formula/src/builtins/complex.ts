import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

interface ComplexBuiltinHelpers {
  toNumber(this: void, value: CellValue): number | undefined
  numberResult(this: void, value: number): CellValue
  valueError(this: void): CellValue
}

type ComplexSuffix = 'i' | 'j'

interface ComplexNumber {
  real: number
  imaginary: number
  suffix: ComplexSuffix
}

export function createComplexBuiltins({ toNumber, numberResult, valueError }: ComplexBuiltinHelpers): Record<string, Builtin> {
  return {
    COMPLEX: (realArg, imaginaryArg = { tag: ValueTag.Number, value: 0 }, suffixArg) => {
      const real = toNumber(realArg)
      const imaginary = toNumber(imaginaryArg)
      if (real === undefined || imaginary === undefined) {
        return valueError()
      }
      let suffix: ComplexSuffix = 'i'
      if (suffixArg !== undefined) {
        if (suffixArg.tag !== ValueTag.String) {
          return valueError()
        }
        const normalized = suffixArg.value.trim().toLowerCase()
        if (normalized !== 'i' && normalized !== 'j') {
          return valueError()
        }
        suffix = normalized
      }
      return complexResult({
        real: normalizeSignedZero(real),
        imaginary: normalizeSignedZero(imaginary),
        suffix,
      })
    },
    IMREAL: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : numberResult(parsed.real)
    },
    IMAGINARY: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : numberResult(parsed.imaginary)
    },
    IMABS: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : numberResult(complexMagnitude(parsed))
    },
    IMARGUMENT: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : numberResult(complexArgument(parsed))
    },
    IMCONJUGATE: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : complexResult(complexConjugate(parsed))
    },
    IMSUM: (...args) => {
      const parsed = args.map((arg) => parseComplexNumber(arg, toNumber))
      if (parsed.length === 0 || parsed.some((value) => value === undefined)) {
        return valueError()
      }
      const values: ComplexNumber[] = []
      for (const value of parsed) {
        if (value === undefined) {
          return valueError()
        }
        values.push(value)
      }
      return complexResult(
        values.reduce((sum, value) => addComplex(sum, value), {
          real: 0,
          imaginary: 0,
          suffix: commonComplexSuffix(values),
        }),
      )
    },
    IMSUB: (leftArg, rightArg) => {
      const left = parseComplexNumber(leftArg, toNumber)
      const right = parseComplexNumber(rightArg, toNumber)
      return left === undefined || right === undefined ? valueError() : complexResult(subtractComplex(left, right))
    },
    IMPRODUCT: (...args) => {
      const parsed = args.map((arg) => parseComplexNumber(arg, toNumber))
      if (parsed.length === 0 || parsed.some((value) => value === undefined)) {
        return valueError()
      }
      const values: ComplexNumber[] = []
      for (const value of parsed) {
        if (value === undefined) {
          return valueError()
        }
        values.push(value)
      }
      return complexResult(
        values.reduce((product, value) => multiplyComplex(product, value), {
          real: 1,
          imaginary: 0,
          suffix: commonComplexSuffix(values),
        }),
      )
    },
    IMDIV: (leftArg, rightArg) => {
      const left = parseComplexNumber(leftArg, toNumber)
      const right = parseComplexNumber(rightArg, toNumber)
      const quotient = left === undefined || right === undefined ? undefined : divideComplex(left, right)
      return quotient === undefined ? (left !== undefined && right !== undefined ? div0Error() : valueError()) : complexResult(quotient)
    },
    IMEXP: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : complexResult(complexExp(parsed))
    },
    IMLN: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const logged = parsed === undefined ? undefined : complexLn(parsed)
      return logged === undefined ? valueError() : complexResult(logged)
    },
    IMLOG10: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const logged = parsed === undefined ? undefined : complexLn(parsed)
      return logged === undefined
        ? valueError()
        : complexResult({
            real: normalizeSignedZero(logged.real / Math.log(10)),
            imaginary: normalizeSignedZero(logged.imaginary / Math.log(10)),
            suffix: logged.suffix,
          })
    },
    IMLOG2: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const logged = parsed === undefined ? undefined : complexLn(parsed)
      return logged === undefined
        ? valueError()
        : complexResult({
            real: normalizeSignedZero(logged.real / Math.log(2)),
            imaginary: normalizeSignedZero(logged.imaginary / Math.log(2)),
            suffix: logged.suffix,
          })
    },
    IMPOWER: (valueArg, powerArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const exponent = toNumber(powerArg)
      const powered = parsed === undefined || exponent === undefined ? undefined : complexPower(parsed, exponent)
      return powered === undefined ? valueError() : complexResult(powered)
    },
    IMSQRT: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : complexResult(complexSqrt(parsed))
    },
    IMSIN: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : complexResult(complexSin(parsed))
    },
    IMCOS: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : complexResult(complexCos(parsed))
    },
    IMTAN: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      if (parsed === undefined) {
        return valueError()
      }
      const quotient = divideComplex(complexSin(parsed), complexCos(parsed))
      return quotient === undefined ? div0Error() : complexResult(quotient)
    },
    IMSINH: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : complexResult(complexSinh(parsed))
    },
    IMCOSH: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      return parsed === undefined ? valueError() : complexResult(complexCosh(parsed))
    },
    IMSEC: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const reciprocal = parsed === undefined ? undefined : reciprocalComplex(complexCos(parsed))
      return reciprocal === undefined ? (parsed === undefined ? valueError() : div0Error()) : complexResult(reciprocal)
    },
    IMCSC: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const reciprocal = parsed === undefined ? undefined : reciprocalComplex(complexSin(parsed))
      return reciprocal === undefined ? (parsed === undefined ? valueError() : div0Error()) : complexResult(reciprocal)
    },
    IMCOT: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      if (parsed === undefined) {
        return valueError()
      }
      const quotient = divideComplex(complexCos(parsed), complexSin(parsed))
      return quotient === undefined ? div0Error() : complexResult(quotient)
    },
    IMSECH: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const reciprocal = parsed === undefined ? undefined : reciprocalComplex(complexCosh(parsed))
      return reciprocal === undefined ? (parsed === undefined ? valueError() : div0Error()) : complexResult(reciprocal)
    },
    IMCSCH: (valueArg) => {
      const parsed = parseComplexNumber(valueArg, toNumber)
      const reciprocal = parsed === undefined ? undefined : reciprocalComplex(complexSinh(parsed))
      return reciprocal === undefined ? (parsed === undefined ? valueError() : div0Error()) : complexResult(reciprocal)
    },
  }
}

function div0Error(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Div0 }
}

function normalizeSignedZero(value: number): number {
  return Object.is(value, -0) || Math.abs(value) < 1e-12 ? 0 : value
}

function parseComplexNumber(value: CellValue, toNumber: ComplexBuiltinHelpers['toNumber']): ComplexNumber | undefined {
  if (value.tag === ValueTag.Error) {
    return undefined
  }
  if (value.tag === ValueTag.Number || value.tag === ValueTag.Boolean || value.tag === ValueTag.Empty) {
    const numeric = toNumber(value)
    return numeric === undefined ? undefined : { real: normalizeSignedZero(numeric), imaginary: 0, suffix: 'i' }
  }
  const raw = value.value.trim()
  if (raw === '') {
    return undefined
  }
  const normalized = raw.toLowerCase()
  const suffix = normalized.endsWith('j') ? 'j' : normalized.endsWith('i') ? 'i' : undefined
  if (!suffix) {
    const real = Number(raw)
    return Number.isFinite(real) ? { real: normalizeSignedZero(real), imaginary: 0, suffix: 'i' } : undefined
  }
  const body = normalized.slice(0, -1)
  let splitIndex = -1
  for (let index = 1; index < body.length; index += 1) {
    const char = body[index]!
    const previous = body[index - 1]!
    if ((char === '+' || char === '-') && previous !== 'e') {
      splitIndex = index
    }
  }
  if (splitIndex !== -1) {
    const real = Number(body.slice(0, splitIndex))
    const imaginaryToken = body.slice(splitIndex)
    const imaginary = imaginaryToken === '+' ? 1 : imaginaryToken === '-' ? -1 : Number(imaginaryToken)
    if (!Number.isFinite(real) || !Number.isFinite(imaginary)) {
      return undefined
    }
    return {
      real: normalizeSignedZero(real),
      imaginary: normalizeSignedZero(imaginary),
      suffix,
    }
  }
  const imaginary = body === '' || body === '+' ? 1 : body === '-' ? -1 : Number(body)
  if (!Number.isFinite(imaginary)) {
    return undefined
  }
  return { real: 0, imaginary: normalizeSignedZero(imaginary), suffix }
}

function complexToString(value: ComplexNumber): string {
  const real = normalizeSignedZero(value.real)
  const imaginary = normalizeSignedZero(value.imaginary)
  if (imaginary === 0) {
    return `${real}`
  }
  const imagMagnitude = Math.abs(imaginary)
  const imagPart = imagMagnitude === 1 ? value.suffix : `${imagMagnitude}${value.suffix}`
  if (real === 0) {
    return imaginary < 0 ? `-${imagPart}` : imagPart
  }
  return `${real}${imaginary < 0 ? '-' : '+'}${imagPart}`
}

function complexResult(value: ComplexNumber): CellValue {
  return { tag: ValueTag.String, value: complexToString(value), stringId: 0 }
}

function commonComplexSuffix(values: readonly ComplexNumber[]): ComplexSuffix {
  return values.some((value) => value.suffix === 'j') ? 'j' : 'i'
}

function addComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(left.real + right.real),
    imaginary: normalizeSignedZero(left.imaginary + right.imaginary),
    suffix: commonComplexSuffix([left, right]),
  }
}

function subtractComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(left.real - right.real),
    imaginary: normalizeSignedZero(left.imaginary - right.imaginary),
    suffix: commonComplexSuffix([left, right]),
  }
}

function multiplyComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(left.real * right.real - left.imaginary * right.imaginary),
    imaginary: normalizeSignedZero(left.real * right.imaginary + left.imaginary * right.real),
    suffix: commonComplexSuffix([left, right]),
  }
}

function divideComplex(left: ComplexNumber, right: ComplexNumber): ComplexNumber | undefined {
  const denominator = right.real ** 2 + right.imaginary ** 2
  if (denominator === 0) {
    return undefined
  }
  return {
    real: normalizeSignedZero((left.real * right.real + left.imaginary * right.imaginary) / denominator),
    imaginary: normalizeSignedZero((left.imaginary * right.real - left.real * right.imaginary) / denominator),
    suffix: commonComplexSuffix([left, right]),
  }
}

function reciprocalComplex(value: ComplexNumber): ComplexNumber | undefined {
  return divideComplex({ real: 1, imaginary: 0, suffix: value.suffix }, value)
}

function complexMagnitude(value: ComplexNumber): number {
  return Math.hypot(value.real, value.imaginary)
}

function complexArgument(value: ComplexNumber): number {
  return Math.atan2(value.imaginary, value.real)
}

function complexConjugate(value: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(value.real),
    imaginary: normalizeSignedZero(-value.imaginary),
    suffix: value.suffix,
  }
}

function complexExp(value: ComplexNumber): ComplexNumber {
  const scale = Math.exp(value.real)
  return {
    real: normalizeSignedZero(scale * Math.cos(value.imaginary)),
    imaginary: normalizeSignedZero(scale * Math.sin(value.imaginary)),
    suffix: value.suffix,
  }
}

function complexLn(value: ComplexNumber): ComplexNumber | undefined {
  const magnitude = complexMagnitude(value)
  if (magnitude === 0) {
    return undefined
  }
  return {
    real: normalizeSignedZero(Math.log(magnitude)),
    imaginary: normalizeSignedZero(complexArgument(value)),
    suffix: value.suffix,
  }
}

function complexPower(value: ComplexNumber, exponent: number): ComplexNumber | undefined {
  const magnitude = complexMagnitude(value)
  if (magnitude === 0) {
    return exponent === 0 ? undefined : { real: 0, imaginary: 0, suffix: value.suffix }
  }
  const angle = complexArgument(value)
  const scaledMagnitude = magnitude ** exponent
  return {
    real: normalizeSignedZero(scaledMagnitude * Math.cos(exponent * angle)),
    imaginary: normalizeSignedZero(scaledMagnitude * Math.sin(exponent * angle)),
    suffix: value.suffix,
  }
}

function complexSqrt(value: ComplexNumber): ComplexNumber {
  const magnitude = complexMagnitude(value)
  const real = Math.sqrt((magnitude + value.real) / 2)
  const imaginarySign = value.imaginary < 0 ? -1 : 1
  const imaginary = imaginarySign * Math.sqrt(Math.max(0, (magnitude - value.real) / 2))
  return {
    real: normalizeSignedZero(real),
    imaginary: normalizeSignedZero(imaginary),
    suffix: value.suffix,
  }
}

function complexSin(value: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(Math.sin(value.real) * Math.cosh(value.imaginary)),
    imaginary: normalizeSignedZero(Math.cos(value.real) * Math.sinh(value.imaginary)),
    suffix: value.suffix,
  }
}

function complexCos(value: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(Math.cos(value.real) * Math.cosh(value.imaginary)),
    imaginary: normalizeSignedZero(-Math.sin(value.real) * Math.sinh(value.imaginary)),
    suffix: value.suffix,
  }
}

function complexSinh(value: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(Math.sinh(value.real) * Math.cos(value.imaginary)),
    imaginary: normalizeSignedZero(Math.cosh(value.real) * Math.sin(value.imaginary)),
    suffix: value.suffix,
  }
}

function complexCosh(value: ComplexNumber): ComplexNumber {
  return {
    real: normalizeSignedZero(Math.cosh(value.real) * Math.cos(value.imaginary)),
    imaginary: normalizeSignedZero(Math.sinh(value.real) * Math.sin(value.imaginary)),
    suffix: value.suffix,
  }
}
