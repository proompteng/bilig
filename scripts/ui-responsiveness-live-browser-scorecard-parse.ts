import type { NumericSummary } from '../packages/benchmarks/src/stats.js'
import {
  arrayField,
  asObject,
  booleanField,
  literalField,
  numberField,
  objectField,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'
import type {
  SameCorpusCapture,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  SameCorpusCaptureVerifiedCell,
  UiResponsivenessLiveBrowserCase,
  UiResponsivenessLiveBrowserScorecard,
  UiResponsivenessLiveBrowserVendor,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
  UiResponsivenessSameCorpusProof,
  UiResponsivenessSameCorpusWorkload,
} from './gen-ui-responsiveness-live-browser-scorecard.ts'
import type {
  SameCorpusPixelGridProof,
  SameCorpusProductPixelGridProof,
  SameCorpusScenarioProof,
  SameCorpusScreenshotProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import { isUiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export function parseUiResponsivenessLiveBrowserScorecard(value: Record<string, unknown>): UiResponsivenessLiveBrowserScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const benchmarkViewport = objectField(benchmark, 'viewport')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'ui-responsiveness-live-browser-timing'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-public-browser-playwright'),
      browserEngine: literalField(source, 'browserEngine', 'chromium'),
      measuredOperation: literalField(source, 'measuredOperation', 'public-workbook-load-and-viewport-scroll'),
    },
    benchmark: {
      sampleCount: numberField(benchmark, 'sampleCount'),
      viewport: {
        width: numberField(benchmarkViewport, 'width'),
        height: numberField(benchmarkViewport, 'height'),
      },
      samplingOrder: literalField(benchmark, 'samplingOrder', 'google-sheets-then-microsoft-excel-web'),
    },
    summary: {
      directBrowserTimingCaptured: booleanField(summary, 'directBrowserTimingCaptured'),
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredVendorCount: numberField(summary, 'requiredVendorCount'),
      capturedVendors: stringArrayField(summary, 'capturedVendors').map(parseVendor),
      limitations: stringArrayField(summary, 'limitations'),
    },
    cases: arrayField(value, 'cases').map(parseBrowserCase),
    sameCorpusProof: parseSameCorpusProof(objectField(value, 'sameCorpusProof')),
  }
}

export function parseSameCorpusCapture(value: Record<string, unknown>): SameCorpusCapture {
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'ui-responsiveness-same-corpus-capture'),
    sampleCount: numberField(value, 'sampleCount'),
    limitations: stringArrayField(value, 'limitations'),
    cases: arrayField(value, 'cases').map(parseSameCorpusCaptureCase),
  }
}

function parseBrowserCase(value: unknown): UiResponsivenessLiveBrowserCase {
  const record = asObject(value, 'UI responsiveness live browser case')
  return {
    id: stringField(record, 'id'),
    vendor: parseVendor(stringField(record, 'vendor')),
    product: stringField(record, 'product'),
    sourceUrl: stringField(record, 'sourceUrl'),
    finalUrl: stringField(record, 'finalUrl'),
    title: stringField(record, 'title'),
    accessMode: parseAccessMode(stringField(record, 'accessMode')),
    workload: literalField(record, 'workload', 'open-public-workbook-and-scroll-viewport'),
    sampleCount: numberField(record, 'sampleCount'),
    loadToReadyMs: parseNumericSummary(objectField(record, 'loadToReadyMs')),
    scrollResponseMs: parseNumericSummary(objectField(record, 'scrollResponseMs')),
    postScrollFrameMs: parseNumericSummary(objectField(record, 'postScrollFrameMs')),
    passed: booleanField(record, 'passed'),
    limitations: stringArrayField(record, 'limitations'),
  }
}

function parseSameCorpusProof(value: Record<string, unknown>): UiResponsivenessSameCorpusProof {
  return {
    captured: booleanField(value, 'captured'),
    evidenceKind: parseSameCorpusEvidenceKind(stringField(value, 'evidenceKind')),
    requiredProductCount: numberField(value, 'requiredProductCount'),
    requiredCaseCount: numberField(value, 'requiredCaseCount'),
    tenXMeanAndP95CaseCount: numberField(value, 'tenXMeanAndP95CaseCount'),
    coveredCorpusCaseIds: stringArrayField(value, 'coveredCorpusCaseIds'),
    limitations: stringArrayField(value, 'limitations'),
    cases: arrayField(value, 'cases').map(parseSameCorpusCase),
  }
}

function parseSameCorpusCase(value: unknown): UiResponsivenessSameCorpusCase {
  const record = asObject(value, 'UI responsiveness same-corpus case')
  const microsoftExcelWeb = Object.hasOwn(record, 'microsoftExcelWeb')
    ? parseSameCorpusMeasurement(objectField(record, 'microsoftExcelWeb'))
    : undefined
  const biligToMicrosoftExcelWebMeanRatio = optionalNumberField(record, 'biligToMicrosoftExcelWebMeanRatio')
  const biligToMicrosoftExcelWebP95Ratio = optionalNumberField(record, 'biligToMicrosoftExcelWebP95Ratio')
  const biligToGoogleSheetsScrollEventMeanRatio = optionalNumberField(record, 'biligToGoogleSheetsScrollEventMeanRatio')
  const biligToGoogleSheetsScrollEventP95Ratio = optionalNumberField(record, 'biligToGoogleSheetsScrollEventP95Ratio')
  const biligToMicrosoftExcelWebScrollEventMeanRatio = optionalNumberField(record, 'biligToMicrosoftExcelWebScrollEventMeanRatio')
  const biligToMicrosoftExcelWebScrollEventP95Ratio = optionalNumberField(record, 'biligToMicrosoftExcelWebScrollEventP95Ratio')
  const tenXMeanAndP95Metric = optionalSameCorpusTenXMetric(record, 'tenXMeanAndP95Metric')
  const postOperationFrameGuardrailPassed = optionalBooleanField(record, 'postOperationFrameGuardrailPassed')
  const scrollMovementGuardrailPassed = optionalBooleanField(record, 'scrollMovementGuardrailPassed')
  return {
    id: stringField(record, 'id'),
    corpusCaseId: stringField(record, 'corpusCaseId'),
    materializedCells: numberField(record, 'materializedCells'),
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    sampleCount: numberField(record, 'sampleCount'),
    bilig: parseSameCorpusMeasurement(objectField(record, 'bilig')),
    googleSheets: parseSameCorpusMeasurement(objectField(record, 'googleSheets')),
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
    biligToGoogleSheetsMeanRatio: numberField(record, 'biligToGoogleSheetsMeanRatio'),
    biligToGoogleSheetsP95Ratio: numberField(record, 'biligToGoogleSheetsP95Ratio'),
    ...(biligToMicrosoftExcelWebMeanRatio !== undefined ? { biligToMicrosoftExcelWebMeanRatio } : {}),
    ...(biligToMicrosoftExcelWebP95Ratio !== undefined ? { biligToMicrosoftExcelWebP95Ratio } : {}),
    ...(biligToGoogleSheetsScrollEventMeanRatio !== undefined ? { biligToGoogleSheetsScrollEventMeanRatio } : {}),
    ...(biligToGoogleSheetsScrollEventP95Ratio !== undefined ? { biligToGoogleSheetsScrollEventP95Ratio } : {}),
    ...(biligToMicrosoftExcelWebScrollEventMeanRatio !== undefined ? { biligToMicrosoftExcelWebScrollEventMeanRatio } : {}),
    ...(biligToMicrosoftExcelWebScrollEventP95Ratio !== undefined ? { biligToMicrosoftExcelWebScrollEventP95Ratio } : {}),
    ...(tenXMeanAndP95Metric ? { tenXMeanAndP95Metric } : {}),
    scenarioProof: parseSameCorpusScenarioProof(objectField(record, 'scenarioProof')),
    tenXMeanAndP95AgainstGoogleSheets: booleanField(record, 'tenXMeanAndP95AgainstGoogleSheets'),
    ...(Object.hasOwn(record, 'tenXMeanAndP95AgainstMicrosoftExcelWeb')
      ? { tenXMeanAndP95AgainstMicrosoftExcelWeb: booleanField(record, 'tenXMeanAndP95AgainstMicrosoftExcelWeb') }
      : {}),
    ...(postOperationFrameGuardrailPassed !== undefined ? { postOperationFrameGuardrailPassed } : {}),
    ...(scrollMovementGuardrailPassed !== undefined ? { scrollMovementGuardrailPassed } : {}),
    passed: booleanField(record, 'passed'),
  }
}

function parseSameCorpusMeasurement(value: Record<string, unknown>): UiResponsivenessSameCorpusMeasurement {
  const scrollEventResponseMs = optionalNumericSummary(value, 'scrollEventResponseMs')
  const scrollMovementPx = optionalNumericSummary(value, 'scrollMovementPx')
  return {
    product: parseSameCorpusProduct(stringField(value, 'product')),
    source: stringField(value, 'source'),
    operationResponseMs: parseNumericSummary(objectField(value, 'operationResponseMs')),
    postOperationFrameMs: parseNumericSummary(objectField(value, 'postOperationFrameMs')),
    ...(scrollEventResponseMs ? { scrollEventResponseMs } : {}),
    ...(scrollMovementPx ? { scrollMovementPx } : {}),
    corpusVerification: parseSameCorpusVerification(objectField(value, 'corpusVerification')),
    limitations: stringArrayField(value, 'limitations'),
  }
}

function parseSameCorpusCaptureCase(value: unknown): SameCorpusCaptureCase {
  const record = asObject(value, 'UI responsiveness same-corpus capture case')
  const microsoftExcelWeb = Object.hasOwn(record, 'microsoftExcelWeb')
    ? parseSameCorpusCaptureMeasurement(objectField(record, 'microsoftExcelWeb'), 'microsoft-excel-web')
    : undefined
  return {
    id: stringField(record, 'id'),
    corpusCaseId: stringField(record, 'corpusCaseId'),
    materializedCells: numberField(record, 'materializedCells'),
    workload: parseSameCorpusWorkload(stringField(record, 'workload')),
    scenarioProof: parseSameCorpusScenarioProof(objectField(record, 'scenarioProof')),
    bilig: parseSameCorpusCaptureMeasurement(objectField(record, 'bilig'), 'bilig'),
    googleSheets: parseSameCorpusCaptureMeasurement(objectField(record, 'googleSheets'), 'google-sheets'),
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
  }
}

function parseSameCorpusCaptureMeasurement(
  value: Record<string, unknown>,
  product: UiResponsivenessSameCorpusProduct,
): SameCorpusCaptureMeasurement {
  const parsedProduct = parseSameCorpusProduct(stringField(value, 'product'))
  if (parsedProduct !== product) {
    throw new Error(`UI responsiveness same-corpus capture product mismatch: expected ${product}, got ${parsedProduct}`)
  }
  return {
    product: parsedProduct,
    source: stringField(value, 'source'),
    operationResponseMsSamples: numericArrayField(value, 'operationResponseMsSamples'),
    postOperationFrameMsSamples: numericArrayField(value, 'postOperationFrameMsSamples'),
    ...(Object.hasOwn(value, 'scrollEventResponseMsSamples')
      ? { scrollEventResponseMsSamples: numericArrayField(value, 'scrollEventResponseMsSamples') }
      : {}),
    ...(Object.hasOwn(value, 'scrollMovementPxSamples')
      ? { scrollMovementPxSamples: numericArrayField(value, 'scrollMovementPxSamples') }
      : {}),
    corpusVerification: parseSameCorpusVerification(objectField(value, 'corpusVerification')),
    limitations: stringArrayField(value, 'limitations'),
  }
}

function parseSameCorpusVerification(value: Record<string, unknown>): SameCorpusCaptureCorpusVerification {
  return {
    verified: booleanField(value, 'verified'),
    method: parseSameCorpusVerificationMethod(stringField(value, 'method')),
    sheetName: stringField(value, 'sheetName'),
    materializedCells: numberField(value, 'materializedCells'),
    checkedCells: arrayField(value, 'checkedCells').map(parseSameCorpusVerifiedCell),
  }
}

function parseSameCorpusScenarioProof(value: Record<string, unknown>): SameCorpusScenarioProof {
  const microsoftExcelWebMeanMs = optionalNumberField(value, 'microsoftExcelWebMeanMs')
  const microsoftExcelWebP95Ms = optionalNumberField(value, 'microsoftExcelWebP95Ms')
  const microsoftExcelWebMeanRatio = optionalNumberField(value, 'microsoftExcelWebMeanRatio')
  const microsoftExcelWebP95Ratio = optionalNumberField(value, 'microsoftExcelWebP95Ratio')
  return {
    biligMeanMs: numberField(value, 'biligMeanMs'),
    biligP95Ms: numberField(value, 'biligP95Ms'),
    googleMeanMs: numberField(value, 'googleMeanMs'),
    googleP95Ms: numberField(value, 'googleP95Ms'),
    ...(microsoftExcelWebMeanMs !== undefined ? { microsoftExcelWebMeanMs } : {}),
    ...(microsoftExcelWebP95Ms !== undefined ? { microsoftExcelWebP95Ms } : {}),
    meanRatio: numberField(value, 'meanRatio'),
    p95Ratio: numberField(value, 'p95Ratio'),
    ...(microsoftExcelWebMeanRatio !== undefined ? { microsoftExcelWebMeanRatio } : {}),
    ...(microsoftExcelWebP95Ratio !== undefined ? { microsoftExcelWebP95Ratio } : {}),
    screenshotProof: parseSameCorpusScreenshotProof(objectField(value, 'screenshotProof')),
    pixelGridProof: parseSameCorpusPixelGridProof(objectField(value, 'pixelGridProof')),
  }
}

function parseSameCorpusScreenshotProof(value: Record<string, unknown>): SameCorpusScreenshotProof {
  return {
    captured: booleanField(value, 'captured'),
    requiredProducts: stringArrayField(value, 'requiredProducts').map(parseSameCorpusProduct),
    artifactPaths: stringArrayField(value, 'artifactPaths'),
    missingProducts: stringArrayField(value, 'missingProducts').map(parseSameCorpusProduct),
  }
}

function parseSameCorpusPixelGridProof(value: Record<string, unknown>): SameCorpusPixelGridProof {
  return {
    captured: booleanField(value, 'captured'),
    requiredProducts: stringArrayField(value, 'requiredProducts').map(parseSameCorpusProduct),
    products: arrayField(value, 'products').map(parseSameCorpusProductPixelGridProof),
    missingProducts: stringArrayField(value, 'missingProducts').map(parseSameCorpusProduct),
  }
}

function parseSameCorpusProductPixelGridProof(value: unknown): SameCorpusProductPixelGridProof {
  const record = asObject(value, 'UI responsiveness same-corpus product pixel grid proof')
  return {
    product: parseSameCorpusProduct(stringField(record, 'product')),
    captured: booleanField(record, 'captured'),
    method: parseSameCorpusPixelGridMethod(stringField(record, 'method')),
    viewportPixelWidth: numberField(record, 'viewportPixelWidth'),
    viewportPixelHeight: numberField(record, 'viewportPixelHeight'),
    evidence: stringArrayField(record, 'evidence'),
  }
}

function parseSameCorpusVerifiedCell(value: unknown): SameCorpusCaptureVerifiedCell {
  const record = asObject(value, 'UI responsiveness same-corpus verified cell')
  return {
    address: stringField(record, 'address'),
    expected: stringField(record, 'expected'),
    actual: stringField(record, 'actual'),
  }
}

function parseNumericSummary(value: Record<string, unknown>): NumericSummary {
  return {
    samples: arrayField(value, 'samples').map((entry) => {
      if (typeof entry !== 'number' || !Number.isFinite(entry)) {
        throw new Error('Expected numeric summary samples to contain finite numbers')
      }
      return entry
    }),
    min: numberField(value, 'min'),
    median: numberField(value, 'median'),
    p95: numberField(value, 'p95'),
    max: numberField(value, 'max'),
    mean: numberField(value, 'mean'),
  }
}

function numericArrayField(value: Record<string, unknown>, key: string): number[] {
  return arrayField(value, key).map((entry) => {
    if (typeof entry !== 'number' || !Number.isFinite(entry) || entry < 0) {
      throw new Error(`Expected ${key} to contain finite non-negative numbers`)
    }
    return entry
  })
}

function optionalNumericSummary(value: Record<string, unknown>, key: string): NumericSummary | undefined {
  return Object.hasOwn(value, key) ? parseNumericSummary(objectField(value, key)) : undefined
}

function optionalNumberField(value: Record<string, unknown>, key: string): number | undefined {
  return Object.hasOwn(value, key) ? numberField(value, key) : undefined
}

function optionalBooleanField(value: Record<string, unknown>, key: string): boolean | undefined {
  return Object.hasOwn(value, key) ? booleanField(value, key) : undefined
}

function optionalSameCorpusTenXMetric(
  value: Record<string, unknown>,
  key: string,
): UiResponsivenessSameCorpusCase['tenXMeanAndP95Metric'] | undefined {
  if (!Object.hasOwn(value, key)) {
    return undefined
  }
  const metric = stringField(value, key)
  if (metric === 'operationResponseMs' || metric === 'scrollEventResponseMs') {
    return metric
  }
  throw new Error(`Unexpected UI responsiveness same-corpus 10x metric: ${metric}`)
}

function parseVendor(value: string): UiResponsivenessLiveBrowserVendor {
  if (value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser vendor: ${value}`)
}

function parseSameCorpusEvidenceKind(value: string): UiResponsivenessSameCorpusProof['evidenceKind'] {
  if (value === 'same-corpus-browser-capture' || value === 'not-captured') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus evidence kind: ${value}`)
}

function parseSameCorpusProduct(value: string): UiResponsivenessSameCorpusProduct {
  if (value === 'bilig' || value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus product: ${value}`)
}

function parseSameCorpusVerificationMethod(value: string): SameCorpusCaptureCorpusVerification['method'] {
  if (value === 'bilig-benchmark-state' || value === 'google-sheets-xlsx-export' || value === 'microsoft-excel-web-source-xlsx') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus verification method: ${value}`)
}

function parseSameCorpusPixelGridMethod(value: string): SameCorpusProductPixelGridProof['method'] {
  if (value === 'typegpu-visible-canvas' || value === 'google-sheets-visible-grid' || value === 'excel-web-visible-grid') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus pixel grid proof method: ${value}`)
}

function parseSameCorpusWorkload(value: string): UiResponsivenessSameCorpusWorkload {
  if (isUiResponsivenessSameCorpusWorkload(value)) {
    return value
  }
  throw new Error(`Unexpected UI responsiveness same-corpus workload: ${value}`)
}

function parseAccessMode(value: string): UiResponsivenessLiveBrowserCase['accessMode'] {
  if (value === 'public-comment-only' || value === 'public-view-only' || value === 'public-office-web-viewer') {
    return value
  }
  throw new Error(`Unexpected UI responsiveness live browser access mode: ${value}`)
}
