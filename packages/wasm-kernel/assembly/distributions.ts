export function erfApprox(value: f64): f64 {
  const sign = value < 0 ? -1.0 : 1.0;
  const absolute = Math.abs(value);
  const t = 1.0 / (1.0 + 0.3275911 * absolute);
  const y =
    1.0 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-absolute * absolute);
  return sign * y;
}

export function standardNormalPdf(value: f64): f64 {
  return Math.exp(-(value * value) / 2.0) / Math.sqrt(2.0 * Math.PI);
}

export function standardNormalCdf(value: f64): f64 {
  return 0.5 * (1.0 + erfApprox(value / Math.sqrt(2.0)));
}

export function inverseStandardNormal(value: f64): f64 {
  if (!(value > 0.0 && value < 1.0)) {
    return NaN;
  }
  const lower = 0.02425;
  const upper = 1.0 - lower;

  if (value < lower) {
    const q = Math.sqrt(-2.0 * Math.log(value));
    return (
      (((((-0.007784894002430293 * q - 0.3223964580411365) * q - 2.400758277161838) * q -
        2.549732539343734) *
        q +
        4.374664141464968) *
        q +
        2.938163982698783) /
      ((((0.007784695709041462 * q + 0.3224671290700398) * q + 2.445134137142996) * q +
        3.754408661907416) *
        q +
        1.0)
    );
  }

  if (value > upper) {
    const q = Math.sqrt(-2.0 * Math.log(1.0 - value));
    return -(
      (((((-0.007784894002430293 * q - 0.3223964580411365) * q - 2.400758277161838) * q -
        2.549732539343734) *
        q +
        4.374664141464968) *
        q +
        2.938163982698783) /
      ((((0.007784695709041462 * q + 0.3224671290700398) * q + 2.445134137142996) * q +
        3.754408661907416) *
        q +
        1.0)
    );
  }

  const q = value - 0.5;
  const r = q * q;
  return (
    ((((((-39.69683028665376 * r + 220.9460984245205) * r - 275.9285104469687) * r +
      138.357751867269) *
      r -
      30.66479806614716) *
      r +
      2.506628277459239) *
      q) /
    (((((-54.47609879822406 * r + 161.5858368580409) * r - 155.6989798598866) * r +
      66.80131188771972) *
      r -
      13.28068155288572) *
      r +
      1.0)
  );
}

export function logGamma(value: f64): f64 {
  if (!isFinite(value) || value <= 0.0) {
    return NaN;
  }
  const shifted = value - 1.0;
  let sum = 0.9999999999998099;
  sum += 676.5203681218851 / (shifted + 1.0);
  sum += -1259.1392167224028 / (shifted + 2.0);
  sum += 771.3234287776531 / (shifted + 3.0);
  sum += -176.6150291621406 / (shifted + 4.0);
  sum += 12.507343278686905 / (shifted + 5.0);
  sum += -0.13857109526572012 / (shifted + 6.0);
  sum += 9.984369578019572e-6 / (shifted + 7.0);
  sum += 1.5056327351493116e-7 / (shifted + 8.0);
  const t = shifted + 7.5;
  return 0.5 * Math.log(2.0 * Math.PI) + (shifted + 0.5) * Math.log(t) - t + Math.log(sum);
}

export function gammaFunction(value: f64): f64 {
  if (!isFinite(value) || (value <= 0.0 && value == Math.floor(value))) {
    return NaN;
  }
  if (value < 0.5) {
    const sine = Math.sin(Math.PI * value);
    if (sine == 0.0) {
      return NaN;
    }
    return Math.PI / (sine * gammaFunction(1.0 - value));
  }
  return Math.exp(logGamma(value));
}

const BESSEL_EPSILON: f64 = 1e-8;

function besselSeries(order: f64, x: f64, alternating: bool): f64 {
  const half = x / 2.0;
  let term = Math.pow(half, order) / gammaFunction(order + 1.0);
  if (!isFinite(term)) {
    return NaN;
  }
  let sum = term;
  for (let index = 0; index < 400; index += 1) {
    const denominator = <f64>(index + 1) * (<f64>index + order + 1.0);
    if (denominator == 0.0) {
      return NaN;
    }
    term *= ((alternating ? -1.0 : 1.0) * half * half) / denominator;
    sum += term;
    if (Math.abs(term) <= Math.abs(sum) * 1e-15) {
      break;
    }
  }
  return sum;
}

export function besselJValue(x: f64, order: i32): f64 {
  if (x == 0.0) {
    return order == 0 ? 1.0 : 0.0;
  }
  const absolute = Math.abs(x);
  let result = besselSeries(<f64>order, absolute, true);
  if (x < 0.0 && (order & 1) == 1) {
    result = -result;
  }
  return result;
}

export function besselIValue(x: f64, order: i32): f64 {
  if (x == 0.0) {
    return order == 0 ? 1.0 : 0.0;
  }
  const absolute = Math.abs(x);
  let result = besselSeries(<f64>order, absolute, false);
  if (x < 0.0 && (order & 1) == 1) {
    result = -result;
  }
  return result;
}

export function besselYValue(x: f64, order: i32): f64 {
  const shiftedOrder = <f64>order + BESSEL_EPSILON;
  return (
    (besselSeries(shiftedOrder, x, true) * Math.cos(Math.PI * shiftedOrder) -
      besselSeries(-shiftedOrder, x, true)) /
    Math.sin(Math.PI * shiftedOrder)
  );
}

export function besselKValue(x: f64, order: i32): f64 {
  const shiftedOrder = <f64>order + BESSEL_EPSILON;
  return (
    ((Math.PI / 2.0) *
      (besselSeries(-shiftedOrder, x, false) - besselSeries(shiftedOrder, x, false))) /
    Math.sin(Math.PI * shiftedOrder)
  );
}

export function regularizedLowerGamma(shape: f64, x: f64): f64 {
  if (!isFinite(shape) || !isFinite(x) || shape <= 0.0 || x < 0.0) {
    return NaN;
  }
  if (x == 0.0) {
    return 0.0;
  }
  const logGammaShape = logGamma(shape);
  if (!isFinite(logGammaShape)) {
    return NaN;
  }
  if (x < shape + 1.0) {
    let term = 1.0 / shape;
    let sum = term;
    for (let iteration = 1; iteration < 1000; iteration += 1) {
      term *= x / (shape + <f64>iteration);
      sum += term;
      if (Math.abs(term) <= Math.abs(sum) * 1e-14) {
        break;
      }
    }
    return sum * Math.exp(-x + shape * Math.log(x) - logGammaShape);
  }

  let b = x + 1.0 - shape;
  let c = 1.0 / 1e-300;
  let d = 1.0 / b;
  let h = d;
  for (let iteration = 1; iteration < 1000; iteration += 1) {
    const factor = -(<f64>iteration) * (<f64>iteration - shape);
    b += 2.0;
    d = factor * d + b;
    if (Math.abs(d) < 1e-300) {
      d = 1e-300;
    }
    c = b + factor / c;
    if (Math.abs(c) < 1e-300) {
      c = 1e-300;
    }
    d = 1.0 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1.0) <= 1e-14) {
      break;
    }
  }
  return 1.0 - Math.exp(-x + shape * Math.log(x) - logGammaShape) * h;
}

export function regularizedUpperGamma(shape: f64, x: f64): f64 {
  const lower = regularizedLowerGamma(shape, x);
  return isFinite(lower) ? 1.0 - lower : NaN;
}

function logBeta(alpha: f64, beta: f64): f64 {
  return logGamma(alpha) + logGamma(beta) - logGamma(alpha + beta);
}

function betaContinuedFraction(x: f64, alpha: f64, beta: f64): f64 {
  const maxIterations = 200;
  const epsilon = 1e-14;
  const tiny = 1e-300;
  const qab = alpha + beta;
  const qap = alpha + 1.0;
  const qam = alpha - 1.0;
  let c = 1.0;
  let d = 1.0 - (qab * x) / qap;
  if (Math.abs(d) < tiny) {
    d = tiny;
  }
  d = 1.0 / d;
  let h = d;
  for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
    const step = <f64>(iteration * 2);
    let factor = (<f64>iteration * (beta - <f64>iteration) * x) / ((qam + step) * (alpha + step));
    d = 1.0 + factor * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1.0 + factor / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1.0 / d;
    h *= d * c;

    factor =
      (-(alpha + <f64>iteration) * (qab + <f64>iteration) * x) / ((alpha + step) * (qap + step));
    d = 1.0 + factor * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1.0 + factor / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1.0 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1.0) <= epsilon) {
      break;
    }
  }
  return h;
}

function regularizedBeta(x: f64, alpha: f64, beta: f64): f64 {
  if (
    !isFinite(x) ||
    !isFinite(alpha) ||
    !isFinite(beta) ||
    alpha <= 0.0 ||
    beta <= 0.0 ||
    x < 0.0 ||
    x > 1.0
  ) {
    return NaN;
  }
  if (x == 0.0) {
    return 0.0;
  }
  if (x == 1.0) {
    return 1.0;
  }
  const logTerm = alpha * Math.log(x) + beta * Math.log(1.0 - x) - logBeta(alpha, beta);
  if (!isFinite(logTerm)) {
    return NaN;
  }
  const front = Math.exp(logTerm);
  if (x < (alpha + 1.0) / (alpha + beta + 2.0)) {
    return (front * betaContinuedFraction(x, alpha, beta)) / alpha;
  }
  return 1.0 - (front * betaContinuedFraction(1.0 - x, beta, alpha)) / beta;
}

export function betaDistributionDensity(
  x: f64,
  alpha: f64,
  beta: f64,
  lowerBound: f64 = 0.0,
  upperBound: f64 = 1.0,
): f64 {
  if (
    !isFinite(x) ||
    !isFinite(alpha) ||
    !isFinite(beta) ||
    !isFinite(lowerBound) ||
    !isFinite(upperBound) ||
    alpha <= 0.0 ||
    beta <= 0.0 ||
    upperBound <= lowerBound ||
    x < lowerBound ||
    x > upperBound
  ) {
    return NaN;
  }
  const scale = upperBound - lowerBound;
  const normalized = (x - lowerBound) / scale;
  if (normalized == 0.0) {
    if (alpha == 1.0) {
      return beta / scale;
    }
    return alpha < 1.0 ? Infinity : 0.0;
  }
  if (normalized == 1.0) {
    if (beta == 1.0) {
      return alpha / scale;
    }
    return beta < 1.0 ? Infinity : 0.0;
  }
  return Math.exp(
    (alpha - 1.0) * Math.log(normalized) +
      (beta - 1.0) * Math.log(1.0 - normalized) -
      logBeta(alpha, beta) -
      Math.log(scale),
  );
}

export function betaDistributionCdf(
  x: f64,
  alpha: f64,
  beta: f64,
  lowerBound: f64 = 0.0,
  upperBound: f64 = 1.0,
): f64 {
  if (
    !isFinite(x) ||
    !isFinite(alpha) ||
    !isFinite(beta) ||
    !isFinite(lowerBound) ||
    !isFinite(upperBound) ||
    alpha <= 0.0 ||
    beta <= 0.0 ||
    upperBound <= lowerBound ||
    x < lowerBound ||
    x > upperBound
  ) {
    return NaN;
  }
  return regularizedBeta((x - lowerBound) / (upperBound - lowerBound), alpha, beta);
}

function inverseRegularizedBeta(probability: f64, alpha: f64, beta: f64): f64 {
  if (
    !isFinite(probability) ||
    !isFinite(alpha) ||
    !isFinite(beta) ||
    !(probability > 0.0 && probability < 1.0) ||
    alpha <= 0.0 ||
    beta <= 0.0
  ) {
    return NaN;
  }
  let lower = 0.0;
  let upper = 1.0;
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const midpoint = (lower + upper) / 2.0;
    const cdf = regularizedBeta(midpoint, alpha, beta);
    if (!isFinite(cdf)) {
      return NaN;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }
  return (lower + upper) / 2.0;
}

export function betaDistributionInverse(
  probability: f64,
  alpha: f64,
  beta: f64,
  lowerBound: f64 = 0.0,
  upperBound: f64 = 1.0,
): f64 {
  if (
    !isFinite(probability) ||
    !isFinite(alpha) ||
    !isFinite(beta) ||
    !isFinite(lowerBound) ||
    !isFinite(upperBound) ||
    !(probability > 0.0 && probability < 1.0) ||
    alpha <= 0.0 ||
    beta <= 0.0 ||
    upperBound <= lowerBound
  ) {
    return NaN;
  }
  const normalized = inverseRegularizedBeta(probability, alpha, beta);
  return isNaN(normalized) ? NaN : lowerBound + (upperBound - lowerBound) * normalized;
}

export function fDistributionDensity(x: f64, degreesFreedom1: f64, degreesFreedom2: f64): f64 {
  if (
    !isFinite(x) ||
    !isFinite(degreesFreedom1) ||
    !isFinite(degreesFreedom2) ||
    x < 0.0 ||
    degreesFreedom1 < 1.0 ||
    degreesFreedom2 < 1.0
  ) {
    return NaN;
  }
  if (x == 0.0) {
    return degreesFreedom1 == 2.0 ? 1.0 : degreesFreedom1 < 2.0 ? Infinity : 0.0;
  }
  const a = degreesFreedom1 / 2.0;
  const b = degreesFreedom2 / 2.0;
  return Math.exp(
    a * Math.log(degreesFreedom1) +
      b * Math.log(degreesFreedom2) +
      (a - 1.0) * Math.log(x) -
      (a + b) * Math.log(degreesFreedom1 * x + degreesFreedom2) -
      logBeta(a, b),
  );
}

export function fDistributionCdf(x: f64, degreesFreedom1: f64, degreesFreedom2: f64): f64 {
  if (
    !isFinite(x) ||
    !isFinite(degreesFreedom1) ||
    !isFinite(degreesFreedom2) ||
    x < 0.0 ||
    degreesFreedom1 < 1.0 ||
    degreesFreedom2 < 1.0
  ) {
    return NaN;
  }
  const a = degreesFreedom1 / 2.0;
  const b = degreesFreedom2 / 2.0;
  const transformed = (degreesFreedom1 * x) / (degreesFreedom1 * x + degreesFreedom2);
  return regularizedBeta(transformed, a, b);
}

export function inverseFDistribution(
  probability: f64,
  degreesFreedom1: f64,
  degreesFreedom2: f64,
): f64 {
  if (
    !isFinite(probability) ||
    !isFinite(degreesFreedom1) ||
    !isFinite(degreesFreedom2) ||
    !(probability > 0.0 && probability < 1.0) ||
    degreesFreedom1 < 1.0 ||
    degreesFreedom2 < 1.0
  ) {
    return NaN;
  }
  const transformed = inverseRegularizedBeta(
    probability,
    degreesFreedom1 / 2.0,
    degreesFreedom2 / 2.0,
  );
  if (isNaN(transformed) || transformed >= 1.0) {
    return NaN;
  }
  return (degreesFreedom2 * transformed) / (degreesFreedom1 * (1.0 - transformed));
}

export function studentTDensity(x: f64, degreesFreedom: f64): f64 {
  if (!isFinite(x) || !isFinite(degreesFreedom) || degreesFreedom < 1.0) {
    return NaN;
  }
  const halfDegrees = degreesFreedom / 2.0;
  return Math.exp(
    logGamma((degreesFreedom + 1.0) / 2.0) -
      logGamma(halfDegrees) -
      0.5 * (Math.log(degreesFreedom) + Math.log(Math.PI)) -
      ((degreesFreedom + 1.0) / 2.0) * Math.log(1.0 + (x * x) / degreesFreedom),
  );
}

export function studentTCdf(x: f64, degreesFreedom: f64): f64 {
  if (!isFinite(x) || !isFinite(degreesFreedom) || degreesFreedom < 1.0) {
    return NaN;
  }
  if (x == 0.0) {
    return 0.5;
  }
  const transformed = degreesFreedom / (degreesFreedom + x * x);
  const tail = regularizedBeta(transformed, degreesFreedom / 2.0, 0.5);
  if (!isFinite(tail)) {
    return NaN;
  }
  return x > 0.0 ? 1.0 - tail / 2.0 : tail / 2.0;
}

export function inverseStudentT(probability: f64, degreesFreedom: f64): f64 {
  if (
    !isFinite(probability) ||
    !isFinite(degreesFreedom) ||
    !(probability > 0.0 && probability < 1.0) ||
    degreesFreedom < 1.0
  ) {
    return NaN;
  }
  if (probability == 0.5) {
    return 0.0;
  }
  if (probability < 0.5) {
    const mirrored = inverseStudentT(1.0 - probability, degreesFreedom);
    return isNaN(mirrored) ? NaN : -mirrored;
  }
  let lower = 0.0;
  let upper = 1.0;
  let upperCdf = studentTCdf(upper, degreesFreedom);
  while (isFinite(upperCdf) && upperCdf < probability && upper < 1e10) {
    lower = upper;
    upper *= 2.0;
    upperCdf = studentTCdf(upper, degreesFreedom);
  }
  if (!isFinite(upperCdf)) {
    return NaN;
  }
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const midpoint = (lower + upper) / 2.0;
    const cdf = studentTCdf(midpoint, degreesFreedom);
    if (!isFinite(cdf)) {
      return NaN;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }
  return (lower + upper) / 2.0;
}

function logCombination(total: i32, chosen: i32): f64 {
  if (chosen < 0 || chosen > total) {
    return NaN;
  }
  return (
    logGamma(<f64>total + 1.0) - logGamma(<f64>chosen + 1.0) - logGamma(<f64>(total - chosen) + 1.0)
  );
}

export function binomialProbability(successes: i32, trials: i32, probability: f64): f64 {
  if (successes < 0 || trials < 0 || successes > trials || probability < 0.0 || probability > 1.0) {
    return NaN;
  }
  if (probability == 0.0) {
    return successes == 0 ? 1.0 : 0.0;
  }
  if (probability == 1.0) {
    return successes == trials ? 1.0 : 0.0;
  }
  return Math.exp(
    logCombination(trials, successes) +
      <f64>successes * Math.log(probability) +
      <f64>(trials - successes) * Math.log(1.0 - probability),
  );
}

export function negativeBinomialProbability(failures: i32, successes: i32, probability: f64): f64 {
  if (failures < 0 || successes <= 0 || probability < 0.0 || probability > 1.0) {
    return NaN;
  }
  if (probability == 0.0) {
    return 0.0;
  }
  if (probability == 1.0) {
    return failures == 0 ? 1.0 : 0.0;
  }
  return Math.exp(
    logCombination(failures + successes - 1, failures) +
      <f64>failures * Math.log(1.0 - probability) +
      <f64>successes * Math.log(probability),
  );
}

export function hypergeometricProbability(
  sampleSuccesses: i32,
  sampleSize: i32,
  populationSuccesses: i32,
  populationSize: i32,
): f64 {
  if (
    sampleSuccesses < 0 ||
    sampleSize < 0 ||
    populationSuccesses < 0 ||
    populationSize <= 0 ||
    sampleSize > populationSize ||
    populationSuccesses > populationSize
  ) {
    return NaN;
  }
  const minimum = max<i32>(0, sampleSize - (populationSize - populationSuccesses));
  const maximum = min<i32>(sampleSize, populationSuccesses);
  if (sampleSuccesses < minimum || sampleSuccesses > maximum) {
    return 0.0;
  }
  return Math.exp(
    logCombination(populationSuccesses, sampleSuccesses) +
      logCombination(populationSize - populationSuccesses, sampleSize - sampleSuccesses) -
      logCombination(populationSize, sampleSize),
  );
}

export function poissonProbability(events: i32, mean: f64): f64 {
  if (events < 0 || !isFinite(mean) || mean < 0.0) {
    return NaN;
  }
  if (mean == 0.0) {
    return events == 0 ? 1.0 : 0.0;
  }
  return Math.exp(<f64>events * Math.log(mean) - mean - logGamma(<f64>events + 1.0));
}

export function gammaDistributionDensity(x: f64, alpha: f64, beta: f64): f64 {
  if (
    !isFinite(x) ||
    !isFinite(alpha) ||
    !isFinite(beta) ||
    x < 0.0 ||
    alpha <= 0.0 ||
    beta <= 0.0
  ) {
    return NaN;
  }
  if (x == 0.0) {
    if (alpha == 1.0) {
      return 1.0 / beta;
    }
    return alpha < 1.0 ? Infinity : 0.0;
  }
  return Math.exp(
    (alpha - 1.0) * Math.log(x) - x / beta - logGamma(alpha) - alpha * Math.log(beta),
  );
}

export function gammaDistributionCdf(x: f64, alpha: f64, beta: f64): f64 {
  return regularizedLowerGamma(alpha, x / beta);
}

export function inverseGammaDistribution(probability: f64, alpha: f64, beta: f64): f64 {
  if (
    !isFinite(probability) ||
    !isFinite(alpha) ||
    !isFinite(beta) ||
    !(probability > 0.0 && probability < 1.0) ||
    !(alpha > 0.0) ||
    !(beta > 0.0)
  ) {
    return NaN;
  }

  let estimate = alpha * beta;
  if (!(estimate > 0.0) || !isFinite(estimate)) {
    estimate = 1.0;
  }

  let lower = 0.0;
  let upper = max<f64>(estimate, 1.0);
  let upperCdf = gammaDistributionCdf(upper, alpha, beta);
  for (let iteration = 0; iteration < 64 && upperCdf < probability; iteration += 1) {
    upper *= 2.0;
    upperCdf = gammaDistributionCdf(upper, alpha, beta);
  }
  if (!(upperCdf >= probability)) {
    return NaN;
  }

  let current = min<f64>(max<f64>(estimate, lower), upper);
  for (let iteration = 0; iteration < 20; iteration += 1) {
    const cdf = gammaDistributionCdf(current, alpha, beta);
    if (!isFinite(cdf)) {
      break;
    }
    if (cdf < probability) {
      lower = current;
    } else {
      upper = current;
    }
    const density = gammaDistributionDensity(current, alpha, beta);
    if (!(density > 0.0) || !isFinite(density)) {
      current = (lower + upper) / 2.0;
      continue;
    }
    const next = current - (cdf - probability) / density;
    current = isFinite(next) && next > lower && next < upper ? next : (lower + upper) / 2.0;
  }

  for (let iteration = 0; iteration < 60; iteration += 1) {
    const midpoint = (lower + upper) / 2.0;
    const cdf = gammaDistributionCdf(midpoint, alpha, beta);
    if (!isFinite(cdf)) {
      return NaN;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }

  return (lower + upper) / 2.0;
}

export function chiSquareDensity(x: f64, degreesFreedom: f64): f64 {
  return gammaDistributionDensity(x, degreesFreedom / 2.0, 2.0);
}

export function chiSquareCdf(x: f64, degreesFreedom: f64): f64 {
  return gammaDistributionCdf(x, degreesFreedom / 2.0, 2.0);
}

export function inverseChiSquare(probability: f64, degreesFreedom: f64): f64 {
  if (
    !isFinite(probability) ||
    !isFinite(degreesFreedom) ||
    !(probability > 0.0 && probability < 1.0) ||
    !(degreesFreedom >= 1.0)
  ) {
    return NaN;
  }

  const z = inverseStandardNormal(probability);
  const approximationFactor = isNaN(z)
    ? NaN
    : 1.0 - 2.0 / (9.0 * degreesFreedom) + z * Math.sqrt(2.0 / (9.0 * degreesFreedom));
  let estimate =
    isFinite(approximationFactor) && approximationFactor > 0.0
      ? degreesFreedom * Math.pow(approximationFactor, 3.0)
      : degreesFreedom;
  if (!(estimate > 0.0) || !isFinite(estimate)) {
    estimate = max<f64>(degreesFreedom, 1.0);
  }

  let lower = 0.0;
  let upper = max<f64>(estimate, 1.0);
  let upperCdf = chiSquareCdf(upper, degreesFreedom);
  for (let iteration = 0; iteration < 64 && upperCdf < probability; iteration += 1) {
    upper *= 2.0;
    upperCdf = chiSquareCdf(upper, degreesFreedom);
  }
  if (!(upperCdf >= probability)) {
    return NaN;
  }

  let current = min<f64>(max<f64>(estimate, lower), upper);
  for (let iteration = 0; iteration < 20; iteration += 1) {
    const cdf = chiSquareCdf(current, degreesFreedom);
    if (!isFinite(cdf)) {
      break;
    }
    if (cdf < probability) {
      lower = current;
    } else {
      upper = current;
    }
    const density = chiSquareDensity(current, degreesFreedom);
    if (!(density > 0.0) || !isFinite(density)) {
      current = (lower + upper) / 2.0;
      continue;
    }
    const next = current - (cdf - probability) / density;
    current = isFinite(next) && next > lower && next < upper ? next : (lower + upper) / 2.0;
  }

  for (let iteration = 0; iteration < 60; iteration += 1) {
    const midpoint = (lower + upper) / 2.0;
    const cdf = chiSquareCdf(midpoint, degreesFreedom);
    if (!isFinite(cdf)) {
      return NaN;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }

  return (lower + upper) / 2.0;
}
