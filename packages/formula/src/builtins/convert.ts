import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { CellValue } from '@bilig/protocol'

type ConvertGroup =
  | 'mass'
  | 'distance'
  | 'time'
  | 'pressure'
  | 'force'
  | 'energy'
  | 'power'
  | 'magnetism'
  | 'temperature'
  | 'volume'
  | 'area'
  | 'information'
  | 'speed'

interface ConvertScaleUnit {
  kind: 'scale'
  group: ConvertGroup
  toBase: number
  metricPrefix?: boolean | undefined
  binaryPrefix?: boolean | undefined
  prefixPower?: number | undefined
}

interface ConvertTemperatureUnit {
  kind: 'temperature'
  group: 'temperature'
  unit: 'C' | 'F' | 'K' | 'Rank' | 'Reau'
}

type ConvertUnit = ConvertScaleUnit | ConvertTemperatureUnit

interface PrefixSpec {
  symbol: string
  factor: number
}

const metricPrefixes: readonly PrefixSpec[] = [
  { symbol: 'da', factor: 1e1 },
  { symbol: 'Y', factor: 1e24 },
  { symbol: 'Z', factor: 1e21 },
  { symbol: 'E', factor: 1e18 },
  { symbol: 'P', factor: 1e15 },
  { symbol: 'T', factor: 1e12 },
  { symbol: 'G', factor: 1e9 },
  { symbol: 'M', factor: 1e6 },
  { symbol: 'k', factor: 1e3 },
  { symbol: 'h', factor: 1e2 },
  { symbol: 'e', factor: 1e1 },
  { symbol: 'd', factor: 1e-1 },
  { symbol: 'c', factor: 1e-2 },
  { symbol: 'm', factor: 1e-3 },
  { symbol: 'u', factor: 1e-6 },
  { symbol: 'n', factor: 1e-9 },
  { symbol: 'p', factor: 1e-12 },
  { symbol: 'f', factor: 1e-15 },
  { symbol: 'a', factor: 1e-18 },
  { symbol: 'z', factor: 1e-21 },
  { symbol: 'y', factor: 1e-24 },
]

const binaryPrefixes: readonly PrefixSpec[] = [
  { symbol: 'Yi', factor: 2 ** 80 },
  { symbol: 'Zi', factor: 2 ** 70 },
  { symbol: 'Ei', factor: 2 ** 60 },
  { symbol: 'Pi', factor: 2 ** 50 },
  { symbol: 'Ti', factor: 2 ** 40 },
  { symbol: 'Gi', factor: 2 ** 30 },
  { symbol: 'Mi', factor: 2 ** 20 },
  { symbol: 'ki', factor: 2 ** 10 },
]

const exactConvertUnits = new Map<string, ConvertUnit>()

function addScaleUnit(
  aliases: readonly string[],
  group: ConvertGroup,
  toBase: number,
  options: Pick<ConvertScaleUnit, 'metricPrefix' | 'binaryPrefix' | 'prefixPower'> = {},
): void {
  for (const alias of aliases) {
    exactConvertUnits.set(alias, {
      kind: 'scale',
      group,
      toBase,
      metricPrefix: options.metricPrefix,
      binaryPrefix: options.binaryPrefix,
      prefixPower: options.prefixPower ?? 1,
    })
  }
}

function addTemperatureUnit(aliases: readonly string[], unit: ConvertTemperatureUnit['unit']): void {
  for (const alias of aliases) {
    exactConvertUnits.set(alias, { kind: 'temperature', group: 'temperature', unit })
  }
}

addScaleUnit(['g'], 'mass', 1, { metricPrefix: true })
addScaleUnit(['sg'], 'mass', 14593.902937206363)
addScaleUnit(['lbm'], 'mass', 453.59237)
addScaleUnit(['u'], 'mass', 1.660538782e-24)
addScaleUnit(['ozm'], 'mass', 28.349523125)
addScaleUnit(['grain'], 'mass', 0.06479891)
addScaleUnit(['cwt', 'shweight'], 'mass', 45359.237)
addScaleUnit(['uk_cwt', 'lcwt', 'hweight'], 'mass', 50802.34544)
addScaleUnit(['stone'], 'mass', 6350.29318)
addScaleUnit(['ton'], 'mass', 907184.74)
addScaleUnit(['uk_ton', 'LTON', 'brton'], 'mass', 1016046.9088)

addScaleUnit(['m'], 'distance', 1, { metricPrefix: true })
addScaleUnit(['mi'], 'distance', 1609.344)
addScaleUnit(['Nmi'], 'distance', 1852)
addScaleUnit(['in'], 'distance', 0.0254)
addScaleUnit(['ft'], 'distance', 0.3048)
addScaleUnit(['yd'], 'distance', 0.9144)
addScaleUnit(['ang'], 'distance', 1e-10)
addScaleUnit(['ell'], 'distance', 1.143)
addScaleUnit(['ly'], 'distance', 9.4607304725808e15)
addScaleUnit(['parsec', 'pc'], 'distance', 3.085677581491367e16)
addScaleUnit(['Picapt', 'Pica'], 'distance', 0.0254 / 72)
addScaleUnit(['pica'], 'distance', 0.0254 / 6)
addScaleUnit(['survey_mi'], 'distance', 1609.3472186944373)

addScaleUnit(['yr'], 'time', 31557600)
addScaleUnit(['day', 'd'], 'time', 86400)
addScaleUnit(['hr'], 'time', 3600)
addScaleUnit(['mn', 'min'], 'time', 60)
addScaleUnit(['sec', 's'], 'time', 1, { metricPrefix: true })

addScaleUnit(['Pa', 'p'], 'pressure', 1, { metricPrefix: true })
addScaleUnit(['atm', 'at'], 'pressure', 101325)
addScaleUnit(['mmHg', 'Torr'], 'pressure', 101325 / 760)
addScaleUnit(['psi'], 'pressure', 6894.757293168361)

addScaleUnit(['N'], 'force', 1, { metricPrefix: true })
addScaleUnit(['dyn', 'dy'], 'force', 1e-5)
addScaleUnit(['lbf'], 'force', 4.4482216152605)

addScaleUnit(['J'], 'energy', 1, { metricPrefix: true })
addScaleUnit(['e'], 'energy', 1e-7)
addScaleUnit(['c'], 'energy', 4.184)
addScaleUnit(['cal'], 'energy', 4.1868)
addScaleUnit(['eV', 'ev'], 'energy', 1.602176487e-19, { metricPrefix: true })
addScaleUnit(['HPh', 'hh'], 'energy', 2684519.537696173)
addScaleUnit(['Wh', 'wh'], 'energy', 3600)
addScaleUnit(['flb'], 'energy', 0.3048 * 4.4482216152605)
addScaleUnit(['BTU', 'btu'], 'energy', 1055.05585262)

addScaleUnit(['HP', 'h'], 'power', 745.6998715822701)
addScaleUnit(['PS'], 'power', 735.49875)
addScaleUnit(['W', 'w'], 'power', 1, { metricPrefix: true })

addScaleUnit(['T'], 'magnetism', 1, { metricPrefix: true })
addScaleUnit(['ga'], 'magnetism', 1e-4)

addTemperatureUnit(['C', 'cel'], 'C')
addTemperatureUnit(['F', 'fah'], 'F')
addTemperatureUnit(['K', 'kel'], 'K')
addTemperatureUnit(['Rank'], 'Rank')
addTemperatureUnit(['Reau'], 'Reau')

addScaleUnit(['tsp'], 'volume', 4.92892159375e-6)
addScaleUnit(['tspm'], 'volume', 5e-6)
addScaleUnit(['tbs'], 'volume', 1.478676478125e-5)
addScaleUnit(['oz'], 'volume', 2.95735295625e-5)
addScaleUnit(['cup'], 'volume', 0.0002365882365)
addScaleUnit(['pt', 'us_pt'], 'volume', 0.000473176473)
addScaleUnit(['uk_pt'], 'volume', 0.00056826125)
addScaleUnit(['qt'], 'volume', 0.000946352946)
addScaleUnit(['uk_qt'], 'volume', 0.0011365225)
addScaleUnit(['gal'], 'volume', 0.003785411784)
addScaleUnit(['uk_gal'], 'volume', 0.00454609)
addScaleUnit(['l', 'L'], 'volume', 0.001, { metricPrefix: true })
addScaleUnit(['lt'], 'volume', 0.001)
addScaleUnit(['ang3', 'ang^3'], 'volume', 1e-30)
addScaleUnit(['barrel'], 'volume', 0.158987294928)
addScaleUnit(['bushel'], 'volume', 0.03523907016688)
addScaleUnit(['ft3', 'ft^3'], 'volume', 0.028316846592)
addScaleUnit(['in3', 'in^3'], 'volume', 1.6387064e-5)
addScaleUnit(['ly3', 'ly^3'], 'volume', 8.46786664623715e47)
addScaleUnit(['m3', 'm^3'], 'volume', 1, { metricPrefix: true, prefixPower: 3 })
addScaleUnit(['mi3', 'mi^3'], 'volume', 4.16818182544058e9)
addScaleUnit(['yd3', 'yd^3'], 'volume', 0.764554857984)
addScaleUnit(['Nmi3', 'Nmi^3'], 'volume', 6.350012608e9)
addScaleUnit(['Picapt3', 'Picapt^3', 'Pica3', 'Pica^3'], 'volume', 4.390243770290368e-11)
addScaleUnit(['GRT', 'regton'], 'volume', 2.8316846592)
addScaleUnit(['MTON'], 'volume', 1.13267386368)

addScaleUnit(['uk_acre'], 'area', 4046.8564224)
addScaleUnit(['us_acre'], 'area', 4046.872609874252)
addScaleUnit(['ang2', 'ang^2'], 'area', 1e-20)
addScaleUnit(['ar'], 'area', 100)
addScaleUnit(['ft2', 'ft^2'], 'area', 0.09290304)
addScaleUnit(['ha'], 'area', 10000)
addScaleUnit(['in2', 'in^2'], 'area', 0.00064516)
addScaleUnit(['ly2', 'ly^2'], 'area', 8.95054210748189e31)
addScaleUnit(['m2', 'm^2'], 'area', 1, { metricPrefix: true, prefixPower: 2 })
addScaleUnit(['Morgen'], 'area', 2500)
addScaleUnit(['mi2', 'mi^2'], 'area', 2589988.110336)
addScaleUnit(['Nmi2', 'Nmi^2'], 'area', 3429904)
addScaleUnit(['Picapt2', 'Pica2', 'Pica^2', 'Picapt^2'], 'area', 1.244707012345679e-7)
addScaleUnit(['yd2', 'yd^2'], 'area', 0.83612736)

addScaleUnit(['bit'], 'information', 1, { metricPrefix: true, binaryPrefix: true })
addScaleUnit(['byte'], 'information', 8, { metricPrefix: true, binaryPrefix: true })

addScaleUnit(['admkn'], 'speed', 0.5147733333333333)
addScaleUnit(['kn'], 'speed', 1852 / 3600)
addScaleUnit(['m/h', 'm/hr'], 'speed', 1 / 3600, { metricPrefix: true })
addScaleUnit(['m/s', 'm/sec'], 'speed', 1, { metricPrefix: true })
addScaleUnit(['mph'], 'speed', 1609.344 / 3600)

const euroRates = new Map<string, number>([
  ['BEF', 40.3399],
  ['LUF', 40.3399],
  ['DEM', 1.95583],
  ['ESP', 166.386],
  ['FRF', 6.55957],
  ['IEP', 0.787564],
  ['ITL', 1936.27],
  ['NLG', 2.20371],
  ['ATS', 13.7603],
  ['PTE', 200.482],
  ['FIM', 5.94573],
  ['GRD', 340.75],
  ['SIT', 239.64],
  ['EUR', 1],
])

const euroCalculationPrecision = new Map<string, number>([
  ['BEF', 0],
  ['LUF', 0],
  ['DEM', 2],
  ['ESP', 0],
  ['FRF', 2],
  ['IEP', 2],
  ['ITL', 0],
  ['NLG', 2],
  ['ATS', 2],
  ['PTE', 0],
  ['FIM', 2],
  ['GRD', 0],
  ['SIT', 2],
  ['EUR', 2],
])

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function errorResult(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function toNumber(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
      return undefined
    default:
      return undefined
  }
}

function toInteger(value: CellValue | undefined): number | undefined {
  const numeric = toNumber(value)
  return numeric !== undefined && Number.isFinite(numeric) ? Math.trunc(numeric) : undefined
}

function toBoolean(value: CellValue | undefined, fallback: boolean): boolean | undefined {
  if (value === undefined) {
    return fallback
  }
  if (value.tag === ValueTag.Boolean) {
    return value.value
  }
  const numeric = toNumber(value)
  return numeric === undefined ? undefined : numeric !== 0
}

function unitText(value: CellValue | undefined): string | undefined {
  return value?.tag === ValueTag.String ? value.value : undefined
}

function roundToPlaces(value: number, places: number): number {
  const scale = 10 ** places
  return Math.round(value * scale) / scale
}

function roundToSignificantDigits(value: number, digits: number): number {
  if (value === 0 || !Number.isFinite(value)) {
    return value
  }
  const exponent = Math.floor(Math.log10(Math.abs(value)))
  const scale = 10 ** (digits - exponent - 1)
  return Math.round(value * scale) / scale
}

function resolvePrefixedUnit(
  unitCode: string,
  prefixes: readonly PrefixSpec[],
  predicate: (unit: ConvertScaleUnit) => boolean,
): ConvertScaleUnit | undefined {
  for (const prefix of prefixes) {
    if (!unitCode.startsWith(prefix.symbol) || unitCode.length <= prefix.symbol.length) {
      continue
    }
    const suffix = unitCode.slice(prefix.symbol.length)
    const exact = exactConvertUnits.get(suffix)
    if (exact?.kind !== 'scale' || !predicate(exact)) {
      continue
    }
    return {
      kind: 'scale',
      group: exact.group,
      toBase: exact.toBase * prefix.factor ** (exact.prefixPower ?? 1),
      metricPrefix: exact.metricPrefix,
      binaryPrefix: exact.binaryPrefix,
      prefixPower: exact.prefixPower,
    }
  }
  return undefined
}

function resolveConvertUnit(unitCode: string): ConvertUnit | undefined {
  const exact = exactConvertUnits.get(unitCode)
  if (exact) {
    return exact
  }
  return (
    resolvePrefixedUnit(unitCode, binaryPrefixes, (unit) => unit.binaryPrefix === true) ??
    resolvePrefixedUnit(unitCode, metricPrefixes, (unit) => unit.metricPrefix === true)
  )
}

function temperatureToKelvin(unit: ConvertTemperatureUnit['unit'], value: number): number {
  switch (unit) {
    case 'C':
      return value + 273.15
    case 'F':
      return (value + 459.67) * (5 / 9)
    case 'K':
      return value
    case 'Rank':
      return value * (5 / 9)
    case 'Reau':
      return value * 1.25 + 273.15
  }
}

function kelvinToTemperature(unit: ConvertTemperatureUnit['unit'], kelvin: number): number {
  switch (unit) {
    case 'C':
      return kelvin - 273.15
    case 'F':
      return kelvin * (9 / 5) - 459.67
    case 'K':
      return kelvin
    case 'Rank':
      return kelvin * (9 / 5)
    case 'Reau':
      return (kelvin - 273.15) * 0.8
  }
}

function normalizeConvertResult(value: number): number {
  if (Object.is(value, -0)) {
    return 0
  }
  const nearestInteger = Math.round(value)
  if (Math.abs(value - nearestInteger) <= 1e-12 * Math.max(1, Math.abs(value))) {
    return nearestInteger
  }
  return roundToSignificantDigits(value, 14)
}

export function convertBuiltin(
  numberArg: CellValue | undefined,
  fromUnitArg: CellValue | undefined,
  toUnitArg: CellValue | undefined,
): CellValue {
  const numeric = toNumber(numberArg)
  const fromUnitCode = unitText(fromUnitArg)
  const toUnitCode = unitText(toUnitArg)
  if (numeric === undefined || fromUnitCode === undefined || toUnitCode === undefined) {
    return errorResult(ErrorCode.Value)
  }
  const fromUnit = resolveConvertUnit(fromUnitCode)
  const toUnit = resolveConvertUnit(toUnitCode)
  if (fromUnit === undefined || toUnit === undefined) {
    return errorResult(ErrorCode.NA)
  }
  if (fromUnit.kind === 'temperature' || toUnit.kind === 'temperature') {
    if (fromUnit.kind !== 'temperature' || toUnit.kind !== 'temperature') {
      return errorResult(ErrorCode.NA)
    }
    return numberResult(normalizeConvertResult(kelvinToTemperature(toUnit.unit, temperatureToKelvin(fromUnit.unit, numeric))))
  }
  if (fromUnit.group !== toUnit.group) {
    return errorResult(ErrorCode.NA)
  }
  return numberResult(normalizeConvertResult((numeric * fromUnit.toBase) / toUnit.toBase))
}

export function euroconvertBuiltin(
  numberArg: CellValue | undefined,
  sourceArg: CellValue | undefined,
  targetArg: CellValue | undefined,
  fullPrecisionArg?: CellValue,
  triangulationPrecisionArg?: CellValue,
): CellValue {
  const numeric = toNumber(numberArg)
  const sourceCode = unitText(sourceArg)
  const targetCode = unitText(targetArg)
  const fullPrecision = toBoolean(fullPrecisionArg, false)
  const triangulationPrecision = triangulationPrecisionArg === undefined ? undefined : toInteger(triangulationPrecisionArg)
  if (
    numeric === undefined ||
    sourceCode === undefined ||
    targetCode === undefined ||
    fullPrecision === undefined ||
    (triangulationPrecisionArg !== undefined && (triangulationPrecision === undefined || triangulationPrecision < 3))
  ) {
    return errorResult(ErrorCode.Value)
  }
  const sourceRate = euroRates.get(sourceCode)
  const targetRate = euroRates.get(targetCode)
  if (sourceRate === undefined || targetRate === undefined) {
    return errorResult(ErrorCode.Value)
  }
  if (sourceCode === targetCode) {
    return numberResult(numeric)
  }

  let euroValue = sourceCode === 'EUR' ? numeric : numeric / sourceRate
  if (sourceCode !== 'EUR' && triangulationPrecision !== undefined) {
    euroValue = roundToSignificantDigits(euroValue, triangulationPrecision)
  }

  const rawResult = targetCode === 'EUR' ? euroValue : euroValue * targetRate
  if (fullPrecision) {
    return numberResult(rawResult)
  }

  return numberResult(roundToPlaces(rawResult, euroCalculationPrecision.get(targetCode) ?? 2))
}
