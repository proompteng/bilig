export const CONVERT_GROUP_TEMPERATURE = 9

const CONVERT_GROUP_INVALID = 0
const CONVERT_GROUP_MASS = 1
const CONVERT_GROUP_DISTANCE = 2
const CONVERT_GROUP_TIME = 3
const CONVERT_GROUP_PRESSURE = 4
const CONVERT_GROUP_FORCE = 5
const CONVERT_GROUP_ENERGY = 6
const CONVERT_GROUP_POWER = 7
const CONVERT_GROUP_MAGNETISM = 8
const CONVERT_GROUP_VOLUME = 10
const CONVERT_GROUP_AREA = 11
const CONVERT_GROUP_INFORMATION = 12
const CONVERT_GROUP_SPEED = 13

const CONVERT_TEMP_NONE = 0
const CONVERT_TEMP_C = 1
const CONVERT_TEMP_F = 2
const CONVERT_TEMP_K = 3
const CONVERT_TEMP_RANK = 4
const CONVERT_TEMP_REAU = 5

export let resolvedConvertGroup = CONVERT_GROUP_INVALID
export let resolvedConvertFactor = NaN
export let resolvedConvertTemperature = CONVERT_TEMP_NONE

function setResolvedConvertScaleUnit(group: i32, factor: f64): bool {
  resolvedConvertGroup = group
  resolvedConvertFactor = factor
  resolvedConvertTemperature = CONVERT_TEMP_NONE
  return true
}

function setResolvedConvertTemperatureUnit(code: i32): bool {
  resolvedConvertGroup = CONVERT_GROUP_TEMPERATURE
  resolvedConvertFactor = NaN
  resolvedConvertTemperature = code
  return true
}

function resolveConvertExactUnit(unitText: string): bool {
  if (unitText == 'g') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 1.0)
  if (unitText == 'sg') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 14593.902937206363)
  if (unitText == 'lbm') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 453.59237)
  if (unitText == 'u') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 1.660538782e-24)
  if (unitText == 'ozm') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 28.349523125)
  if (unitText == 'grain') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 0.06479891)
  if (unitText == 'cwt' || unitText == 'shweight') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 45359.237)
  }
  if (unitText == 'uk_cwt' || unitText == 'lcwt' || unitText == 'hweight') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 50802.34544)
  }
  if (unitText == 'stone') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 6350.29318)
  if (unitText == 'ton') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 907184.74)
  if (unitText == 'uk_ton' || unitText == 'LTON' || unitText == 'brton') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 1016046.9088)
  }

  if (unitText == 'm') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1.0)
  if (unitText == 'mi') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1609.344)
  if (unitText == 'Nmi') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1852.0)
  if (unitText == 'in') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.0254)
  if (unitText == 'ft') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.3048)
  if (unitText == 'yd') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.9144)
  if (unitText == 'ang') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1e-10)
  if (unitText == 'ell') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1.143)
  if (unitText == 'ly') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 9.4607304725808e15)
  if (unitText == 'parsec' || unitText == 'pc') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 3.085677581491367e16)
  }
  if (unitText == 'Picapt' || unitText == 'Pica') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.0254 / 72.0)
  }
  if (unitText == 'pica') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.0254 / 6.0)
  if (unitText == 'survey_mi') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1609.3472186944373)
  }

  if (unitText == 'yr') return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 31557600.0)
  if (unitText == 'day' || unitText == 'd') return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 86400.0)
  if (unitText == 'hr') return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 3600.0)
  if (unitText == 'mn' || unitText == 'min') return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 60.0)
  if (unitText == 'sec' || unitText == 's') return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 1.0)

  if (unitText == 'Pa' || unitText == 'p') return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 1.0)
  if (unitText == 'atm' || unitText == 'at') return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 101325.0)
  if (unitText == 'mmHg' || unitText == 'Torr') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 101325.0 / 760.0)
  }
  if (unitText == 'psi') return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 6894.757293168361)

  if (unitText == 'N') return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, 1.0)
  if (unitText == 'dyn' || unitText == 'dy') return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, 1e-5)
  if (unitText == 'lbf') return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, 4.4482216152605)

  if (unitText == 'J') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.0)
  if (unitText == 'e') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1e-7)
  if (unitText == 'c') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 4.184)
  if (unitText == 'cal') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 4.1868)
  if (unitText == 'eV' || unitText == 'ev') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.602176487e-19)
  }
  if (unitText == 'HPh' || unitText == 'hh') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 2684519.537696173)
  }
  if (unitText == 'Wh' || unitText == 'wh') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 3600.0)
  if (unitText == 'flb') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.3558179483314004)
  if (unitText == 'BTU' || unitText == 'btu') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1055.05585262)
  }

  if (unitText == 'HP' || unitText == 'h') return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, 745.6998715822701)
  if (unitText == 'PS') return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, 735.49875)
  if (unitText == 'W' || unitText == 'w') return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, 1.0)

  if (unitText == 'T') return setResolvedConvertScaleUnit(CONVERT_GROUP_MAGNETISM, 1.0)
  if (unitText == 'ga') return setResolvedConvertScaleUnit(CONVERT_GROUP_MAGNETISM, 1e-4)

  if (unitText == 'C' || unitText == 'cel') return setResolvedConvertTemperatureUnit(CONVERT_TEMP_C)
  if (unitText == 'F' || unitText == 'fah') return setResolvedConvertTemperatureUnit(CONVERT_TEMP_F)
  if (unitText == 'K' || unitText == 'kel') return setResolvedConvertTemperatureUnit(CONVERT_TEMP_K)
  if (unitText == 'Rank') return setResolvedConvertTemperatureUnit(CONVERT_TEMP_RANK)
  if (unitText == 'Reau') return setResolvedConvertTemperatureUnit(CONVERT_TEMP_REAU)

  if (unitText == 'tsp') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 4.92892159375e-6)
  if (unitText == 'tspm') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 5e-6)
  if (unitText == 'tbs') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.478676478125e-5)
  if (unitText == 'oz') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 2.95735295625e-5)
  if (unitText == 'cup') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.0002365882365)
  if (unitText == 'pt' || unitText == 'us_pt') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.000473176473)
  if (unitText == 'uk_pt') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.00056826125)
  if (unitText == 'qt') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.000946352946)
  if (unitText == 'uk_qt') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.0011365225)
  if (unitText == 'gal') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.003785411784)
  if (unitText == 'uk_gal') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.00454609)
  if (unitText == 'l' || unitText == 'L' || unitText == 'lt') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.001)
  if (unitText == 'ang3' || unitText == 'ang^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1e-30)
  if (unitText == 'barrel') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.158987294928)
  if (unitText == 'bushel') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.03523907016688)
  if (unitText == 'ft3' || unitText == 'ft^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.028316846592)
  if (unitText == 'in3' || unitText == 'in^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.6387064e-5)
  if (unitText == 'ly3' || unitText == 'ly^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 8.46786664623715e47)
  if (unitText == 'm3' || unitText == 'm^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.0)
  if (unitText == 'mi3' || unitText == 'mi^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 4.16818182544058e9)
  if (unitText == 'yd3' || unitText == 'yd^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.764554857984)
  if (unitText == 'Nmi3' || unitText == 'Nmi^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 6.350012608e9)
  if (unitText == 'Picapt3' || unitText == 'Picapt^3' || unitText == 'Pica3' || unitText == 'Pica^3') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 4.390243770290368e-11)
  }
  if (unitText == 'GRT' || unitText == 'regton') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 2.8316846592)
  if (unitText == 'MTON') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.13267386368)

  if (unitText == 'uk_acre') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 4046.8564224)
  if (unitText == 'us_acre') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 4046.872609874252)
  if (unitText == 'ang2' || unitText == 'ang^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 1e-20)
  if (unitText == 'ar') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 100.0)
  if (unitText == 'ft2' || unitText == 'ft^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 0.09290304)
  if (unitText == 'ha') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 10000.0)
  if (unitText == 'in2' || unitText == 'in^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 0.00064516)
  if (unitText == 'ly2' || unitText == 'ly^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 8.95054210748189e31)
  if (unitText == 'm2' || unitText == 'm^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 1.0)
  if (unitText == 'Morgen') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 2500.0)
  if (unitText == 'mi2' || unitText == 'mi^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 2589988.110336)
  if (unitText == 'Nmi2' || unitText == 'Nmi^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 3429904.0)
  if (unitText == 'Picapt2' || unitText == 'Pica2' || unitText == 'Pica^2' || unitText == 'Picapt^2') {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 1.244707012345679e-7)
  }
  if (unitText == 'yd2' || unitText == 'yd^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 0.83612736)

  if (unitText == 'bit') return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 1.0)
  if (unitText == 'byte') return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 8.0)

  if (unitText == 'admkn') return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 0.5147733333333333)
  if (unitText == 'kn') return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1852.0 / 3600.0)
  if (unitText == 'm/h' || unitText == 'm/hr') return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1.0 / 3600.0)
  if (unitText == 'm/s' || unitText == 'm/sec') return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1.0)
  if (unitText == 'mph') return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1609.344 / 3600.0)

  return false
}

function resolveConvertMetricSuffix(suffix: string, prefixFactor: f64): bool {
  if (suffix == 'g') return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, prefixFactor)
  if (suffix == 'm') return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, prefixFactor)
  if (suffix == 'sec' || suffix == 's') return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, prefixFactor)
  if (suffix == 'Pa' || suffix == 'p') return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, prefixFactor)
  if (suffix == 'N') return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, prefixFactor)
  if (suffix == 'J') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, prefixFactor)
  if (suffix == 'eV' || suffix == 'ev') return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.602176487e-19 * prefixFactor)
  if (suffix == 'W' || suffix == 'w') return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, prefixFactor)
  if (suffix == 'T') return setResolvedConvertScaleUnit(CONVERT_GROUP_MAGNETISM, prefixFactor)
  if (suffix == 'l' || suffix == 'L') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.001 * prefixFactor)
  if (suffix == 'm3' || suffix == 'm^3') return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, Math.pow(prefixFactor, 3.0))
  if (suffix == 'm2' || suffix == 'm^2') return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, Math.pow(prefixFactor, 2.0))
  if (suffix == 'bit') return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, prefixFactor)
  if (suffix == 'byte') return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 8.0 * prefixFactor)
  if (suffix == 'm/h' || suffix == 'm/hr') return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, prefixFactor / 3600.0)
  if (suffix == 'm/s' || suffix == 'm/sec') return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, prefixFactor)
  return false
}

function resolveConvertBinarySuffix(suffix: string, prefixFactor: f64): bool {
  if (suffix == 'bit') return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, prefixFactor)
  if (suffix == 'byte') return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 8.0 * prefixFactor)
  return false
}

function resolveConvertMetricPrefixedUnit(unitText: string): bool {
  if (unitText.length <= 1) return false
  if (unitText.startsWith('da')) return resolveConvertMetricSuffix(unitText.slice(2), 1e1)
  if (unitText.startsWith('Y')) return resolveConvertMetricSuffix(unitText.slice(1), 1e24)
  if (unitText.startsWith('Z')) return resolveConvertMetricSuffix(unitText.slice(1), 1e21)
  if (unitText.startsWith('E')) return resolveConvertMetricSuffix(unitText.slice(1), 1e18)
  if (unitText.startsWith('P')) return resolveConvertMetricSuffix(unitText.slice(1), 1e15)
  if (unitText.startsWith('T')) return resolveConvertMetricSuffix(unitText.slice(1), 1e12)
  if (unitText.startsWith('G')) return resolveConvertMetricSuffix(unitText.slice(1), 1e9)
  if (unitText.startsWith('M')) return resolveConvertMetricSuffix(unitText.slice(1), 1e6)
  if (unitText.startsWith('k')) return resolveConvertMetricSuffix(unitText.slice(1), 1e3)
  if (unitText.startsWith('h')) return resolveConvertMetricSuffix(unitText.slice(1), 1e2)
  if (unitText.startsWith('e')) return resolveConvertMetricSuffix(unitText.slice(1), 1e1)
  if (unitText.startsWith('d')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-1)
  if (unitText.startsWith('c')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-2)
  if (unitText.startsWith('m')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-3)
  if (unitText.startsWith('u')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-6)
  if (unitText.startsWith('n')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-9)
  if (unitText.startsWith('p')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-12)
  if (unitText.startsWith('f')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-15)
  if (unitText.startsWith('a')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-18)
  if (unitText.startsWith('z')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-21)
  if (unitText.startsWith('y')) return resolveConvertMetricSuffix(unitText.slice(1), 1e-24)
  return false
}

function resolveConvertBinaryPrefixedUnit(unitText: string): bool {
  if (unitText.length <= 2) return false
  if (unitText.startsWith('Yi')) return resolveConvertBinarySuffix(unitText.slice(2), 1208925819614629174706176.0)
  if (unitText.startsWith('Zi')) return resolveConvertBinarySuffix(unitText.slice(2), 1180591620717411303424.0)
  if (unitText.startsWith('Ei')) return resolveConvertBinarySuffix(unitText.slice(2), 1152921504606846976.0)
  if (unitText.startsWith('Pi')) return resolveConvertBinarySuffix(unitText.slice(2), 1125899906842624.0)
  if (unitText.startsWith('Ti')) return resolveConvertBinarySuffix(unitText.slice(2), 1099511627776.0)
  if (unitText.startsWith('Gi')) return resolveConvertBinarySuffix(unitText.slice(2), 1073741824.0)
  if (unitText.startsWith('Mi')) return resolveConvertBinarySuffix(unitText.slice(2), 1048576.0)
  if (unitText.startsWith('ki')) return resolveConvertBinarySuffix(unitText.slice(2), 1024.0)
  return false
}

export function resolveConvertUnit(unitText: string): bool {
  return resolveConvertExactUnit(unitText) || resolveConvertBinaryPrefixedUnit(unitText) || resolveConvertMetricPrefixedUnit(unitText)
}

export function convertTemperatureToKelvin(unitCode: i32, value: f64): f64 {
  if (unitCode == CONVERT_TEMP_C) return value + 273.15
  if (unitCode == CONVERT_TEMP_F) return (value + 459.67) * (5.0 / 9.0)
  if (unitCode == CONVERT_TEMP_K) return value
  if (unitCode == CONVERT_TEMP_RANK) return value * (5.0 / 9.0)
  if (unitCode == CONVERT_TEMP_REAU) return value * 1.25 + 273.15
  return NaN
}

export function convertKelvinToTemperature(unitCode: i32, value: f64): f64 {
  if (unitCode == CONVERT_TEMP_C) return value - 273.15
  if (unitCode == CONVERT_TEMP_F) return value * (9.0 / 5.0) - 459.67
  if (unitCode == CONVERT_TEMP_K) return value
  if (unitCode == CONVERT_TEMP_RANK) return value * (9.0 / 5.0)
  if (unitCode == CONVERT_TEMP_REAU) return (value - 273.15) * 0.8
  return NaN
}
