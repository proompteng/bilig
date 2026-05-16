import type { InstallBenchmarkCorpusResult } from './worker-runtime.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function isBenchmarkCorpusViewport(value: unknown): value is InstallBenchmarkCorpusResult['primaryViewport'] {
  if (!isRecord(value)) {
    return false
  }
  return (
    typeof value['sheetName'] === 'string' &&
    value['sheetName'].length > 0 &&
    isNonNegativeInteger(value['rowStart']) &&
    isNonNegativeInteger(value['rowEnd']) &&
    value['rowEnd'] >= value['rowStart'] &&
    isNonNegativeInteger(value['colStart']) &&
    isNonNegativeInteger(value['colEnd']) &&
    value['colEnd'] >= value['colStart']
  )
}

export function isInstallBenchmarkCorpusResult(value: unknown): value is InstallBenchmarkCorpusResult {
  if (!isRecord(value)) {
    return false
  }
  return (
    typeof value['id'] === 'string' &&
    value['id'].length > 0 &&
    isNonNegativeInteger(value['materializedCellCount']) &&
    value['materializedCellCount'] > 0 &&
    isBenchmarkCorpusViewport(value['primaryViewport'])
  )
}
