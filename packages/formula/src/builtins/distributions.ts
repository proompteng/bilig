import { erfApprox, populationStandardDeviation, sampleStandardDeviation } from './statistics.js'

const LANCZOS_G = 7
const LANCZOS_COEFFICIENTS = [
  676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406, 12.507343278686905, -0.13857109526572012,
  9.984369578019572e-6, 1.5056327351493116e-7,
] as const

const BESSEL_EPSILON = 1e-8

export function logGamma(value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    return Number.NaN
  }
  let sum = 0.9999999999998099
  const shifted = value - 1
  LANCZOS_COEFFICIENTS.forEach((coefficient, index) => {
    sum += coefficient / (shifted + index + 1)
  })
  const t = shifted + LANCZOS_G + 0.5
  return 0.5 * Math.log(2 * Math.PI) + (shifted + 0.5) * Math.log(t) - t + Math.log(sum)
}

function gammaReal(value: number): number {
  if (!Number.isFinite(value) || (value <= 0 && value === Math.trunc(value))) {
    return Number.NaN
  }
  if (value < 0.5) {
    const sine = Math.sin(Math.PI * value)
    return sine === 0 ? Number.NaN : Math.PI / (sine * gammaReal(1 - value))
  }
  return Math.exp(logGamma(value))
}

function besselSeries(order: number, x: number, alternating: boolean): number {
  const half = x / 2
  let term = half ** order / gammaReal(order + 1)
  if (!Number.isFinite(term)) {
    return Number.NaN
  }
  let sum = term
  for (let index = 0; index < 400; index += 1) {
    const denominator = (index + 1) * (index + order + 1)
    if (denominator === 0) {
      return Number.NaN
    }
    term *= ((alternating ? -1 : 1) * half * half) / denominator
    sum += term
    if (Math.abs(term) <= Math.abs(sum) * 1e-15) {
      break
    }
  }
  return sum
}

export function besselJValue(x: number, order: number): number {
  if (x === 0) {
    return order === 0 ? 1 : 0
  }
  const absolute = Math.abs(x)
  const result = besselSeries(order, absolute, true)
  return x < 0 && order % 2 === 1 ? -result : result
}

export function besselIValue(x: number, order: number): number {
  if (x === 0) {
    return order === 0 ? 1 : 0
  }
  const absolute = Math.abs(x)
  const result = besselSeries(order, absolute, false)
  return x < 0 && order % 2 === 1 ? -result : result
}

export function besselYValue(x: number, order: number): number {
  const shiftedOrder = order + BESSEL_EPSILON
  return (
    (besselSeries(shiftedOrder, x, true) * Math.cos(Math.PI * shiftedOrder) - besselSeries(-shiftedOrder, x, true)) /
    Math.sin(Math.PI * shiftedOrder)
  )
}

export function besselKValue(x: number, order: number): number {
  const shiftedOrder = order + BESSEL_EPSILON
  return ((Math.PI / 2) * (besselSeries(-shiftedOrder, x, false) - besselSeries(shiftedOrder, x, false))) / Math.sin(Math.PI * shiftedOrder)
}

export function gammaFunction(value: number): number {
  if (!Number.isFinite(value) || (Number.isInteger(value) && value <= 0)) {
    return Number.NaN
  }
  if (value < 0.5) {
    const sine = Math.sin(Math.PI * value)
    if (sine === 0) {
      return Number.NaN
    }
    return Math.PI / (sine * gammaFunction(1 - value))
  }
  return Math.exp(logGamma(value))
}

function regularizedLowerGamma(shape: number, x: number): number {
  if (!Number.isFinite(shape) || !Number.isFinite(x) || shape <= 0 || x < 0) {
    return Number.NaN
  }
  if (x === 0) {
    return 0
  }
  const logGammaShape = logGamma(shape)
  if (!Number.isFinite(logGammaShape)) {
    return Number.NaN
  }
  if (x < shape + 1) {
    let term = 1 / shape
    let sum = term
    for (let iteration = 1; iteration < 1000; iteration += 1) {
      term *= x / (shape + iteration)
      sum += term
      if (Math.abs(term) <= Math.abs(sum) * 1e-14) {
        break
      }
    }
    return sum * Math.exp(-x + shape * Math.log(x) - logGammaShape)
  }

  let b = x + 1 - shape
  let c = 1 / 1e-300
  let d = 1 / b
  let h = d
  for (let iteration = 1; iteration < 1000; iteration += 1) {
    const factor = -iteration * (iteration - shape)
    b += 2
    d = factor * d + b
    if (Math.abs(d) < 1e-300) {
      d = 1e-300
    }
    c = b + factor / c
    if (Math.abs(c) < 1e-300) {
      c = 1e-300
    }
    d = 1 / d
    const delta = d * c
    h *= delta
    if (Math.abs(delta - 1) <= 1e-14) {
      break
    }
  }
  return 1 - Math.exp(-x + shape * Math.log(x) - logGammaShape) * h
}

export function regularizedUpperGamma(shape: number, x: number): number {
  const lower = regularizedLowerGamma(shape, x)
  return Number.isFinite(lower) ? 1 - lower : Number.NaN
}

function logBeta(alpha: number, beta: number): number {
  return logGamma(alpha) + logGamma(beta) - logGamma(alpha + beta)
}

function betaContinuedFraction(x: number, alpha: number, beta: number): number {
  const maxIterations = 200
  const epsilon = 1e-14
  const tiny = 1e-300
  const qab = alpha + beta
  const qap = alpha + 1
  const qam = alpha - 1
  let c = 1
  let d = 1 - (qab * x) / qap
  if (Math.abs(d) < tiny) {
    d = tiny
  }
  d = 1 / d
  let h = d
  for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
    const step = iteration * 2
    let factor = (iteration * (beta - iteration) * x) / ((qam + step) * (alpha + step))
    d = 1 + factor * d
    if (Math.abs(d) < tiny) {
      d = tiny
    }
    c = 1 + factor / c
    if (Math.abs(c) < tiny) {
      c = tiny
    }
    d = 1 / d
    h *= d * c

    factor = (-(alpha + iteration) * (qab + iteration) * x) / ((alpha + step) * (qap + step))
    d = 1 + factor * d
    if (Math.abs(d) < tiny) {
      d = tiny
    }
    c = 1 + factor / c
    if (Math.abs(c) < tiny) {
      c = tiny
    }
    d = 1 / d
    const delta = d * c
    h *= delta
    if (Math.abs(delta - 1) <= epsilon) {
      break
    }
  }
  return h
}

function regularizedBeta(x: number, alpha: number, beta: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(alpha) || !Number.isFinite(beta) || alpha <= 0 || beta <= 0 || x < 0 || x > 1) {
    return Number.NaN
  }
  if (x === 0) {
    return 0
  }
  if (x === 1) {
    return 1
  }
  const logTerm = alpha * Math.log(x) + beta * Math.log(1 - x) - logBeta(alpha, beta)
  if (!Number.isFinite(logTerm)) {
    return Number.NaN
  }
  const front = Math.exp(logTerm)
  if (x < (alpha + 1) / (alpha + beta + 2)) {
    return (front * betaContinuedFraction(x, alpha, beta)) / alpha
  }
  return 1 - (front * betaContinuedFraction(1 - x, beta, alpha)) / beta
}

export function betaDistributionDensity(x: number, alpha: number, beta: number, lowerBound = 0, upperBound = 1): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !Number.isFinite(lowerBound) ||
    !Number.isFinite(upperBound) ||
    alpha <= 0 ||
    beta <= 0 ||
    upperBound <= lowerBound ||
    x < lowerBound ||
    x > upperBound
  ) {
    return Number.NaN
  }
  const scale = upperBound - lowerBound
  const normalized = (x - lowerBound) / scale
  if (normalized === 0) {
    if (alpha === 1) {
      return beta / scale
    }
    return alpha < 1 ? Number.POSITIVE_INFINITY : 0
  }
  if (normalized === 1) {
    if (beta === 1) {
      return alpha / scale
    }
    return beta < 1 ? Number.POSITIVE_INFINITY : 0
  }
  return Math.exp((alpha - 1) * Math.log(normalized) + (beta - 1) * Math.log(1 - normalized) - logBeta(alpha, beta) - Math.log(scale))
}

export function betaDistributionCdf(x: number, alpha: number, beta: number, lowerBound = 0, upperBound = 1): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !Number.isFinite(lowerBound) ||
    !Number.isFinite(upperBound) ||
    alpha <= 0 ||
    beta <= 0 ||
    upperBound <= lowerBound ||
    x < lowerBound ||
    x > upperBound
  ) {
    return Number.NaN
  }
  return regularizedBeta((x - lowerBound) / (upperBound - lowerBound), alpha, beta)
}

function inverseRegularizedBeta(probability: number, alpha: number, beta: number): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !(probability > 0 && probability < 1) ||
    alpha <= 0 ||
    beta <= 0
  ) {
    return undefined
  }
  let lower = 0
  let upper = 1
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const midpoint = (lower + upper) / 2
    const cdf = regularizedBeta(midpoint, alpha, beta)
    if (!Number.isFinite(cdf)) {
      return undefined
    }
    if (cdf < probability) {
      lower = midpoint
    } else {
      upper = midpoint
    }
  }
  return (lower + upper) / 2
}

export function betaDistributionInverse(
  probability: number,
  alpha: number,
  beta: number,
  lowerBound = 0,
  upperBound = 1,
): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !Number.isFinite(lowerBound) ||
    !Number.isFinite(upperBound) ||
    !(probability > 0 && probability < 1) ||
    alpha <= 0 ||
    beta <= 0 ||
    upperBound <= lowerBound
  ) {
    return undefined
  }
  const normalized = inverseRegularizedBeta(probability, alpha, beta)
  return normalized === undefined ? undefined : lowerBound + (upperBound - lowerBound) * normalized
}

export function fDistributionDensity(x: number, degreesFreedom1: number, degreesFreedom2: number): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(degreesFreedom1) ||
    !Number.isFinite(degreesFreedom2) ||
    x < 0 ||
    degreesFreedom1 < 1 ||
    degreesFreedom2 < 1
  ) {
    return Number.NaN
  }
  if (x === 0) {
    return degreesFreedom1 === 2 ? 1 : degreesFreedom1 < 2 ? Number.POSITIVE_INFINITY : 0
  }
  const a = degreesFreedom1 / 2
  const b = degreesFreedom2 / 2
  return Math.exp(
    a * Math.log(degreesFreedom1) +
      b * Math.log(degreesFreedom2) +
      (a - 1) * Math.log(x) -
      (a + b) * Math.log(degreesFreedom1 * x + degreesFreedom2) -
      logBeta(a, b),
  )
}

export function fDistributionCdf(x: number, degreesFreedom1: number, degreesFreedom2: number): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(degreesFreedom1) ||
    !Number.isFinite(degreesFreedom2) ||
    x < 0 ||
    degreesFreedom1 < 1 ||
    degreesFreedom2 < 1
  ) {
    return Number.NaN
  }
  const a = degreesFreedom1 / 2
  const b = degreesFreedom2 / 2
  const transformed = (degreesFreedom1 * x) / (degreesFreedom1 * x + degreesFreedom2)
  return regularizedBeta(transformed, a, b)
}

export function inverseFDistribution(probability: number, degreesFreedom1: number, degreesFreedom2: number): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(degreesFreedom1) ||
    !Number.isFinite(degreesFreedom2) ||
    !(probability > 0 && probability < 1) ||
    degreesFreedom1 < 1 ||
    degreesFreedom2 < 1
  ) {
    return undefined
  }
  const transformed = inverseRegularizedBeta(probability, degreesFreedom1 / 2, degreesFreedom2 / 2)
  if (transformed === undefined || transformed >= 1) {
    return undefined
  }
  return (degreesFreedom2 * transformed) / (degreesFreedom1 * (1 - transformed))
}

export function studentTDensity(x: number, degreesFreedom: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(degreesFreedom) || degreesFreedom < 1) {
    return Number.NaN
  }
  const halfDegrees = degreesFreedom / 2
  return Math.exp(
    logGamma((degreesFreedom + 1) / 2) -
      logGamma(halfDegrees) -
      0.5 * (Math.log(degreesFreedom) + Math.log(Math.PI)) -
      ((degreesFreedom + 1) / 2) * Math.log(1 + (x * x) / degreesFreedom),
  )
}

export function studentTCdf(x: number, degreesFreedom: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(degreesFreedom) || degreesFreedom < 1) {
    return Number.NaN
  }
  if (x === 0) {
    return 0.5
  }
  const transformed = degreesFreedom / (degreesFreedom + x * x)
  const tail = regularizedBeta(transformed, degreesFreedom / 2, 0.5)
  if (!Number.isFinite(tail)) {
    return Number.NaN
  }
  return x > 0 ? 1 - tail / 2 : tail / 2
}

export function inverseStudentT(probability: number, degreesFreedom: number): number | undefined {
  if (!Number.isFinite(probability) || !Number.isFinite(degreesFreedom) || !(probability > 0 && probability < 1) || degreesFreedom < 1) {
    return undefined
  }
  if (probability === 0.5) {
    return 0
  }
  if (probability < 0.5) {
    const mirrored = inverseStudentT(1 - probability, degreesFreedom)
    return mirrored === undefined ? undefined : -mirrored
  }
  let lower = 0
  let upper = 1
  let upperCdf = studentTCdf(upper, degreesFreedom)
  while (Number.isFinite(upperCdf) && upperCdf < probability && upper < 1e10) {
    lower = upper
    upper *= 2
    upperCdf = studentTCdf(upper, degreesFreedom)
  }
  if (!Number.isFinite(upperCdf)) {
    return undefined
  }
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const midpoint = (lower + upper) / 2
    const cdf = studentTCdf(midpoint, degreesFreedom)
    if (!Number.isFinite(cdf)) {
      return undefined
    }
    if (cdf < probability) {
      lower = midpoint
    } else {
      upper = midpoint
    }
  }
  return (lower + upper) / 2
}

function logCombination(total: number, chosen: number): number {
  if (!Number.isInteger(total) || !Number.isInteger(chosen) || chosen < 0 || chosen > total) {
    return Number.NaN
  }
  return logGamma(total + 1) - logGamma(chosen + 1) - logGamma(total - chosen + 1)
}

export function binomialProbability(successes: number, trials: number, probability: number): number {
  if (
    !Number.isInteger(successes) ||
    !Number.isInteger(trials) ||
    successes < 0 ||
    trials < 0 ||
    successes > trials ||
    probability < 0 ||
    probability > 1
  ) {
    return Number.NaN
  }
  if (probability === 0) {
    return successes === 0 ? 1 : 0
  }
  if (probability === 1) {
    return successes === trials ? 1 : 0
  }
  return Math.exp(logCombination(trials, successes) + successes * Math.log(probability) + (trials - successes) * Math.log(1 - probability))
}

export function negativeBinomialProbability(failures: number, successes: number, probability: number): number {
  if (!Number.isInteger(failures) || !Number.isInteger(successes) || failures < 0 || successes <= 0 || probability < 0 || probability > 1) {
    return Number.NaN
  }
  if (probability === 0) {
    return 0
  }
  if (probability === 1) {
    return failures === 0 ? 1 : 0
  }
  return Math.exp(
    logCombination(failures + successes - 1, failures) + failures * Math.log(1 - probability) + successes * Math.log(probability),
  )
}

export function hypergeometricProbability(
  sampleSuccesses: number,
  sampleSize: number,
  populationSuccesses: number,
  populationSize: number,
): number {
  if (
    !Number.isInteger(sampleSuccesses) ||
    !Number.isInteger(sampleSize) ||
    !Number.isInteger(populationSuccesses) ||
    !Number.isInteger(populationSize) ||
    sampleSuccesses < 0 ||
    sampleSize < 0 ||
    populationSuccesses < 0 ||
    populationSize <= 0 ||
    sampleSize > populationSize ||
    populationSuccesses > populationSize
  ) {
    return Number.NaN
  }
  const minimum = Math.max(0, sampleSize - (populationSize - populationSuccesses))
  const maximum = Math.min(sampleSize, populationSuccesses)
  if (sampleSuccesses < minimum || sampleSuccesses > maximum) {
    return 0
  }
  return Math.exp(
    logCombination(populationSuccesses, sampleSuccesses) +
      logCombination(populationSize - populationSuccesses, sampleSize - sampleSuccesses) -
      logCombination(populationSize, sampleSize),
  )
}

export function poissonProbability(events: number, mean: number): number {
  if (!Number.isInteger(events) || events < 0 || !Number.isFinite(mean) || mean < 0) {
    return Number.NaN
  }
  if (mean === 0) {
    return events === 0 ? 1 : 0
  }
  return Math.exp(events * Math.log(mean) - mean - logGamma(events + 1))
}

export function gammaDistributionDensity(x: number, alpha: number, beta: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(alpha) || !Number.isFinite(beta)) {
    return Number.NaN
  }
  if (x < 0 || alpha <= 0 || beta <= 0) {
    return Number.NaN
  }
  if (x === 0) {
    if (alpha === 1) {
      return 1 / beta
    }
    return alpha < 1 ? Number.POSITIVE_INFINITY : 0
  }
  return Math.exp((alpha - 1) * Math.log(x) - x / beta - logGamma(alpha) - alpha * Math.log(beta))
}

export function gammaDistributionCdf(x: number, alpha: number, beta: number): number {
  return regularizedLowerGamma(alpha, x / beta)
}

export function inverseGammaDistribution(probability: number, alpha: number, beta: number): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !(probability > 0 && probability < 1) ||
    !(alpha > 0) ||
    !(beta > 0)
  ) {
    return undefined
  }

  let estimate = alpha * beta
  if (!(estimate > 0) || !Number.isFinite(estimate)) {
    estimate = 1
  }

  let lower = 0
  let upper = Math.max(estimate, 1)
  let upperCdf = gammaDistributionCdf(upper, alpha, beta)
  for (let iteration = 0; iteration < 64 && upperCdf < probability; iteration += 1) {
    upper *= 2
    upperCdf = gammaDistributionCdf(upper, alpha, beta)
  }
  if (!(upperCdf >= probability)) {
    return undefined
  }

  let current = Math.min(Math.max(estimate, lower), upper)
  for (let iteration = 0; iteration < 20; iteration += 1) {
    const cdf = gammaDistributionCdf(current, alpha, beta)
    if (!Number.isFinite(cdf)) {
      break
    }
    if (cdf < probability) {
      lower = current
    } else {
      upper = current
    }
    const density = gammaDistributionDensity(current, alpha, beta)
    if (!(density > 0) || !Number.isFinite(density)) {
      current = (lower + upper) / 2
      continue
    }
    const next = current - (cdf - probability) / density
    current = Number.isFinite(next) && next > lower && next < upper ? next : (lower + upper) / 2
  }

  for (let iteration = 0; iteration < 60; iteration += 1) {
    const midpoint = (lower + upper) / 2
    const cdf = gammaDistributionCdf(midpoint, alpha, beta)
    if (!Number.isFinite(cdf)) {
      return undefined
    }
    if (cdf < probability) {
      lower = midpoint
    } else {
      upper = midpoint
    }
  }

  return (lower + upper) / 2
}

export function chiSquareDensity(x: number, degreesFreedom: number): number {
  return gammaDistributionDensity(x, degreesFreedom / 2, 2)
}

export function chiSquareCdf(x: number, degreesFreedom: number): number {
  return gammaDistributionCdf(x, degreesFreedom / 2, 2)
}

export function inverseChiSquare(probability: number, degreesFreedom: number): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(degreesFreedom) ||
    !(probability > 0 && probability < 1) ||
    !(degreesFreedom >= 1)
  ) {
    return undefined
  }

  const z = inverseStandardNormal(probability)
  const approximationFactor = z === undefined ? Number.NaN : 1 - 2 / (9 * degreesFreedom) + z * Math.sqrt(2 / (9 * degreesFreedom))
  let estimate =
    Number.isFinite(approximationFactor) && approximationFactor > 0 ? degreesFreedom * approximationFactor ** 3 : degreesFreedom
  if (!(estimate > 0) || !Number.isFinite(estimate)) {
    estimate = Math.max(degreesFreedom, 1)
  }

  let lower = 0
  let upper = Math.max(estimate, 1)
  let upperCdf = chiSquareCdf(upper, degreesFreedom)
  for (let iteration = 0; iteration < 64 && upperCdf < probability; iteration += 1) {
    upper *= 2
    upperCdf = chiSquareCdf(upper, degreesFreedom)
  }
  if (!(upperCdf >= probability)) {
    return undefined
  }

  let current = Math.min(Math.max(estimate, lower), upper)
  for (let iteration = 0; iteration < 20; iteration += 1) {
    const cdf = chiSquareCdf(current, degreesFreedom)
    if (!Number.isFinite(cdf)) {
      break
    }
    if (cdf < probability) {
      lower = current
    } else {
      upper = current
    }
    const density = chiSquareDensity(current, degreesFreedom)
    if (!(density > 0) || !Number.isFinite(density)) {
      current = (lower + upper) / 2
      continue
    }
    const next = current - (cdf - probability) / density
    current = Number.isFinite(next) && next > lower && next < upper ? next : (lower + upper) / 2
  }

  for (let iteration = 0; iteration < 60; iteration += 1) {
    const midpoint = (lower + upper) / 2
    const cdf = chiSquareCdf(midpoint, degreesFreedom)
    if (!Number.isFinite(cdf)) {
      return undefined
    }
    if (cdf < probability) {
      lower = midpoint
    } else {
      upper = midpoint
    }
  }

  return (lower + upper) / 2
}

export function standardNormalPdf(value: number): number {
  return Math.exp(-(value * value) / 2) / Math.sqrt(2 * Math.PI)
}

export function standardNormalCdf(value: number): number {
  return 0.5 * (1 + erfApprox(value / Math.SQRT2))
}

export function inverseStandardNormal(probability: number): number | undefined {
  if (!(probability > 0 && probability < 1)) {
    return undefined
  }
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239]
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572]
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783]
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416]
  const lower = 0.02425
  const upper = 1 - lower

  if (probability < lower) {
    const q = Math.sqrt(-2 * Math.log(probability))
    return (
      (((((c[0]! * q + c[1]!) * q + c[2]!) * q + c[3]!) * q + c[4]!) * q + c[5]!) /
      ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1)
    )
  }
  if (probability > upper) {
    const q = Math.sqrt(-2 * Math.log(1 - probability))
    return -(
      (((((c[0]! * q + c[1]!) * q + c[2]!) * q + c[3]!) * q + c[4]!) * q + c[5]!) /
      ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1)
    )
  }
  const q = probability - 0.5
  const r = q * q
  return (
    ((((((a[0]! * r + a[1]!) * r + a[2]!) * r + a[3]!) * r + a[4]!) * r + a[5]!) * q) /
    (((((b[0]! * r + b[1]!) * r + b[2]!) * r + b[3]!) * r + b[4]!) * r + 1)
  )
}

export function skewSample(numbers: number[]): number | undefined {
  if (numbers.length < 3) {
    return undefined
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length
  const stddev = sampleStandardDeviation(numbers)
  if (!(stddev > 0)) {
    return undefined
  }
  const moment3 = numbers.reduce((sum, value) => sum + (value - mean) ** 3, 0)
  const n = numbers.length
  return (n * moment3) / ((n - 1) * (n - 2) * stddev ** 3)
}

export function skewPopulation(numbers: number[]): number | undefined {
  if (numbers.length === 0) {
    return undefined
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length
  const stddev = populationStandardDeviation(numbers)
  if (!(stddev > 0)) {
    return undefined
  }
  const moment3 = numbers.reduce((sum, value) => sum + (value - mean) ** 3, 0) / numbers.length
  return moment3 / stddev ** 3
}

export function kurtosis(numbers: number[]): number | undefined {
  if (numbers.length < 4) {
    return undefined
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length
  const stddev = sampleStandardDeviation(numbers)
  if (!(stddev > 0)) {
    return undefined
  }
  const n = numbers.length
  const sum4 = numbers.reduce((sum, value) => sum + ((value - mean) / stddev) ** 4, 0)
  return (n * (n + 1) * sum4) / ((n - 1) * (n - 2) * (n - 3)) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
}

export function percentileNormal(mean: number, standardDeviation: number, value: number): number {
  return standardNormalCdf((value - mean) / standardDeviation)
}

export function inverseNormal(probability: number, mean: number, standardDeviation: number): number | undefined {
  const z = inverseStandardNormal(probability)
  return z === undefined ? undefined : mean + standardDeviation * z
}
