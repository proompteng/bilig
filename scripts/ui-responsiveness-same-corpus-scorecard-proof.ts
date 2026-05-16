import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import { validateSameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

export type UiResponsivenessSameCorpusProduct = 'bilig' | 'google-sheets' | 'microsoft-excel-web'

export interface UiResponsivenessSameCorpusMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: string
  readonly operationResponseMs: NumericSummary
  readonly postOperationFrameMs: NumericSummary
  readonly scrollEventResponseMs?: NumericSummary
  readonly scrollMovementPx?: NumericSummary
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface UiResponsivenessSameCorpusCase {
  readonly id: string
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleCount: number
  readonly bilig: UiResponsivenessSameCorpusMeasurement
  readonly googleSheets: UiResponsivenessSameCorpusMeasurement
  readonly microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement | undefined
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio?: number | undefined
  readonly biligToMicrosoftExcelWebP95Ratio?: number | undefined
  readonly biligToGoogleSheetsScrollEventMeanRatio?: number
  readonly biligToGoogleSheetsScrollEventP95Ratio?: number
  readonly biligToMicrosoftExcelWebScrollEventMeanRatio?: number
  readonly biligToMicrosoftExcelWebScrollEventP95Ratio?: number
  readonly tenXMeanAndP95Metric?: 'operationResponseMs' | 'scrollEventResponseMs'
  readonly scenarioProof: SameCorpusScenarioProof
  readonly tenXMeanAndP95AgainstGoogleSheets: boolean
  readonly tenXMeanAndP95AgainstMicrosoftExcelWeb?: boolean | undefined
  readonly postOperationFrameGuardrailPassed?: boolean
  readonly scrollMovementGuardrailPassed?: boolean
  readonly passed: boolean
}

export interface UiResponsivenessSameCorpusProof {
  readonly captured: boolean
  readonly evidenceKind: 'same-corpus-browser-capture' | 'not-captured'
  readonly requiredProductCount: number
  readonly requiredCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly coveredCorpusCaseIds: string[]
  readonly limitations: string[]
  readonly cases: UiResponsivenessSameCorpusCase[]
}

export interface SameCorpusCapture {
  readonly schemaVersion: 1
  readonly suite: 'ui-responsiveness-same-corpus-capture'
  readonly sampleCount: number
  readonly limitations: string[]
  readonly cases: SameCorpusCaptureCase[]
}

export interface SameCorpusCaptureCase {
  readonly id: string
  readonly corpusCaseId: string
  readonly materializedCells: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly scenarioProof: SameCorpusScenarioProof
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb?: SameCorpusCaptureMeasurement | undefined
}

export interface SameCorpusCaptureMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly source: string
  readonly operationResponseMsSamples: number[]
  readonly postOperationFrameMsSamples: number[]
  readonly scrollEventResponseMsSamples?: number[]
  readonly scrollMovementPxSamples?: number[]
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly limitations: string[]
}

export interface SameCorpusCaptureVerifiedCell {
  readonly address: string
  readonly expected: string
  readonly actual: string
}

export interface SameCorpusCaptureCorpusVerification {
  readonly verified: boolean
  readonly method: 'bilig-benchmark-state' | 'google-sheets-xlsx-export' | 'microsoft-excel-web-source-xlsx'
  readonly sheetName: string
  readonly materializedCells: number
  readonly checkedCells: readonly SameCorpusCaptureVerifiedCell[]
}

const requiredSameCorpusWorkloads = requiredUiResponsivenessSameCorpusWorkloads
const sameCorpusSampleCount = 3

export function buildMissingSameCorpusProof(): UiResponsivenessSameCorpusProof {
  return {
    captured: false,
    evidenceKind: 'not-captured',
    requiredProductCount: 2,
    requiredCaseCount: requiredSameCorpusWorkloads.length,
    tenXMeanAndP95CaseCount: 0,
    coveredCorpusCaseIds: [],
    limitations: ['Same-corpus live browser timing against Bilig and Google Sheets has not been captured yet.'],
    cases: [],
  }
}

export function buildSameCorpusProof(capture: SameCorpusCapture): UiResponsivenessSameCorpusProof {
  validateSameCorpusCapture(capture)
  const cases = capture.cases.map(buildSameCorpusCase)
  const proof: UiResponsivenessSameCorpusProof = {
    captured: true,
    evidenceKind: 'same-corpus-browser-capture',
    requiredProductCount: 2,
    requiredCaseCount: requiredSameCorpusWorkloads.length,
    tenXMeanAndP95CaseCount: cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length,
    coveredCorpusCaseIds: [...new Set(cases.map((entry) => entry.corpusCaseId))].toSorted(),
    limitations: [...capture.limitations],
    cases,
  }
  validateSameCorpusProof(proof)
  return proof
}

export function validateSameCorpusProof(proof: UiResponsivenessSameCorpusProof): void {
  if (proof.requiredProductCount !== 2) {
    throw new Error('UI responsiveness same-corpus Google Sheets proof must compare Bilig and Google Sheets')
  }
  if (proof.requiredCaseCount !== requiredSameCorpusWorkloads.length) {
    throw new Error('UI responsiveness same-corpus proof required case count is stale')
  }
  if (!proof.captured) {
    if (proof.evidenceKind !== 'not-captured' || proof.cases.length !== 0) {
      throw new Error('UI responsiveness same-corpus proof has stale not-captured metadata')
    }
    if (proof.limitations.length === 0) {
      throw new Error('UI responsiveness same-corpus proof must disclose that capture is missing')
    }
    return
  }
  if (proof.evidenceKind !== 'same-corpus-browser-capture') {
    throw new Error('UI responsiveness same-corpus proof has stale capture metadata')
  }
  for (const workload of requiredSameCorpusWorkloads) {
    if (!proof.cases.some((entry) => entry.workload === workload)) {
      throw new Error(`UI responsiveness same-corpus proof is missing required workload: ${workload}`)
    }
  }
  if (proof.cases.length !== proof.requiredCaseCount) {
    throw new Error('UI responsiveness same-corpus proof must include every required captured case')
  }
  const tenXCaseCount = proof.cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length
  if (proof.tenXMeanAndP95CaseCount !== tenXCaseCount) {
    throw new Error('UI responsiveness same-corpus proof 10x case count is stale')
  }
  const coveredCorpusCaseIds = [...new Set(proof.cases.map((entry) => entry.corpusCaseId))].toSorted()
  if (JSON.stringify(proof.coveredCorpusCaseIds) !== JSON.stringify(coveredCorpusCaseIds)) {
    throw new Error('UI responsiveness same-corpus proof covered corpus IDs are stale')
  }
  for (const entry of proof.cases) {
    validateSameCorpusCase(entry)
  }
}

function validateSameCorpusCapture(capture: SameCorpusCapture): void {
  if (capture.sampleCount < sameCorpusSampleCount) {
    throw new Error('UI responsiveness same-corpus capture must contain at least 3 samples per product')
  }
  if (capture.cases.length === 0) {
    throw new Error('UI responsiveness same-corpus capture must include at least one case')
  }
  for (const entry of capture.cases) {
    const measurements = [entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])]
    const hasAnyScrollEventSamples = measurements.some(
      (measurement) => measurement.scrollEventResponseMsSamples !== undefined || measurement.scrollMovementPxSamples !== undefined,
    )
    const requiresScrollEventSamples = uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload) || hasAnyScrollEventSamples
    for (const measurement of measurements) {
      if (
        measurement.operationResponseMsSamples.length < capture.sampleCount ||
        measurement.postOperationFrameMsSamples.length < capture.sampleCount
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few samples for ${entry.id}`)
      }
      if (
        requiresScrollEventSamples &&
        ((measurement.scrollEventResponseMsSamples?.length ?? 0) < capture.sampleCount ||
          (measurement.scrollMovementPxSamples?.length ?? 0) < capture.sampleCount)
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few scroll-event samples for ${entry.id}`)
      }
      validateSameCorpusCaptureVerification(measurement.corpusVerification, measurement.product, entry.materializedCells, entry.id)
    }
  }
}

function buildSameCorpusCase(captureCase: SameCorpusCaptureCase): UiResponsivenessSameCorpusCase {
  const bilig = buildSameCorpusMeasurement(captureCase.bilig)
  const googleSheets = buildSameCorpusMeasurement(captureCase.googleSheets)
  const microsoftExcelWeb = captureCase.microsoftExcelWeb ? buildSameCorpusMeasurement(captureCase.microsoftExcelWeb) : undefined
  const biligToGoogleSheetsMeanRatio = ratio(bilig.operationResponseMs.mean, googleSheets.operationResponseMs.mean)
  const biligToGoogleSheetsP95Ratio = ratio(bilig.operationResponseMs.p95, googleSheets.operationResponseMs.p95)
  const biligToMicrosoftExcelWebMeanRatio = microsoftExcelWeb
    ? ratio(bilig.operationResponseMs.mean, microsoftExcelWeb.operationResponseMs.mean)
    : undefined
  const biligToMicrosoftExcelWebP95Ratio = microsoftExcelWeb
    ? ratio(bilig.operationResponseMs.p95, microsoftExcelWeb.operationResponseMs.p95)
    : undefined
  const scrollEventMetrics = sameCorpusScrollEventMetrics(bilig, googleSheets, microsoftExcelWeb)
  const comparedProducts = [bilig, googleSheets, ...(microsoftExcelWeb ? [microsoftExcelWeb] : [])]
  const postOperationFrameGuardrailPassed = comparedProducts.every(
    (entry) => entry.postOperationFrameMs.p95 > 0 && entry.postOperationFrameMs.p95 <= 50,
  )
  const scrollMovementGuardrailPassed =
    scrollEventMetrics !== null && comparedProducts.every((entry) => (entry.scrollMovementPx?.min ?? 0) >= 1)
  const requiresScrollEventMetric = uiSameCorpusWorkloadRequiresScrollEventEvidence(captureCase.workload)
  const timingMetricPassedAgainstGoogleSheets = requiresScrollEventMetric
    ? scrollEventMetrics !== null &&
      scrollEventMetrics.biligToGoogleSheetsMeanRatio <= 0.1 &&
      scrollEventMetrics.biligToGoogleSheetsP95Ratio <= 0.1 &&
      scrollMovementGuardrailPassed
    : biligToGoogleSheetsMeanRatio <= 0.1 && biligToGoogleSheetsP95Ratio <= 0.1
  const timingMetricPassedAgainstMicrosoftExcelWeb = microsoftExcelWeb
    ? requiresScrollEventMetric
      ? scrollEventMetrics !== null &&
        scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio <= 0.1 &&
        scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio <= 0.1 &&
        scrollMovementGuardrailPassed
      : (biligToMicrosoftExcelWebMeanRatio ?? Number.POSITIVE_INFINITY) <= 0.1 &&
        (biligToMicrosoftExcelWebP95Ratio ?? Number.POSITIVE_INFINITY) <= 0.1
    : undefined
  const visualProofGuardrailPassed = captureCase.scenarioProof.screenshotProof.captured && captureCase.scenarioProof.pixelGridProof.captured
  const tenXMeanAndP95AgainstGoogleSheets =
    timingMetricPassedAgainstGoogleSheets && postOperationFrameGuardrailPassed && visualProofGuardrailPassed
  const tenXMeanAndP95AgainstMicrosoftExcelWeb =
    timingMetricPassedAgainstMicrosoftExcelWeb === undefined
      ? undefined
      : timingMetricPassedAgainstMicrosoftExcelWeb && postOperationFrameGuardrailPassed && visualProofGuardrailPassed
  return {
    id: captureCase.id,
    corpusCaseId: captureCase.corpusCaseId,
    materializedCells: captureCase.materializedCells,
    workload: captureCase.workload,
    sampleCount: Math.min(
      bilig.operationResponseMs.samples.length,
      googleSheets.operationResponseMs.samples.length,
      ...(microsoftExcelWeb ? [microsoftExcelWeb.operationResponseMs.samples.length] : []),
    ),
    bilig,
    googleSheets,
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
    biligToGoogleSheetsMeanRatio,
    biligToGoogleSheetsP95Ratio,
    ...(biligToMicrosoftExcelWebMeanRatio !== undefined ? { biligToMicrosoftExcelWebMeanRatio } : {}),
    ...(biligToMicrosoftExcelWebP95Ratio !== undefined ? { biligToMicrosoftExcelWebP95Ratio } : {}),
    ...(requiresScrollEventMetric && scrollEventMetrics
      ? {
          biligToGoogleSheetsScrollEventMeanRatio: scrollEventMetrics.biligToGoogleSheetsMeanRatio,
          biligToGoogleSheetsScrollEventP95Ratio: scrollEventMetrics.biligToGoogleSheetsP95Ratio,
          ...(microsoftExcelWeb
            ? {
                biligToMicrosoftExcelWebScrollEventMeanRatio: scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio,
                biligToMicrosoftExcelWebScrollEventP95Ratio: scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio,
              }
            : {}),
          tenXMeanAndP95Metric: 'scrollEventResponseMs' as const,
          scrollMovementGuardrailPassed,
        }
      : { tenXMeanAndP95Metric: 'operationResponseMs' as const }),
    scenarioProof: { ...captureCase.scenarioProof },
    postOperationFrameGuardrailPassed,
    tenXMeanAndP95AgainstGoogleSheets,
    ...(tenXMeanAndP95AgainstMicrosoftExcelWeb !== undefined ? { tenXMeanAndP95AgainstMicrosoftExcelWeb } : {}),
    passed: tenXMeanAndP95AgainstGoogleSheets,
  }
}

function buildSameCorpusMeasurement(capture: SameCorpusCaptureMeasurement): UiResponsivenessSameCorpusMeasurement {
  return {
    product: capture.product,
    source: capture.source,
    operationResponseMs: summarizeNumbers(capture.operationResponseMsSamples),
    postOperationFrameMs: summarizeNumbers(capture.postOperationFrameMsSamples),
    ...(capture.scrollEventResponseMsSamples ? { scrollEventResponseMs: summarizeNumbers(capture.scrollEventResponseMsSamples) } : {}),
    ...(capture.scrollMovementPxSamples ? { scrollMovementPx: summarizeNumbers(capture.scrollMovementPxSamples) } : {}),
    corpusVerification: cloneSameCorpusVerification(capture.corpusVerification),
    limitations: [...capture.limitations],
  }
}

function sameCorpusScrollEventMetrics(
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): {
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio: number
  readonly biligToMicrosoftExcelWebP95Ratio: number
} | null {
  if (
    !bilig.scrollEventResponseMs ||
    !googleSheets.scrollEventResponseMs ||
    (microsoftExcelWeb && !microsoftExcelWeb.scrollEventResponseMs)
  ) {
    return null
  }
  return {
    biligToGoogleSheetsMeanRatio: ratio(bilig.scrollEventResponseMs.mean, googleSheets.scrollEventResponseMs.mean),
    biligToGoogleSheetsP95Ratio: ratio(bilig.scrollEventResponseMs.p95, googleSheets.scrollEventResponseMs.p95),
    biligToMicrosoftExcelWebMeanRatio: microsoftExcelWeb
      ? ratio(bilig.scrollEventResponseMs.mean, microsoftExcelWeb.scrollEventResponseMs!.mean)
      : Number.POSITIVE_INFINITY,
    biligToMicrosoftExcelWebP95Ratio: microsoftExcelWeb
      ? ratio(bilig.scrollEventResponseMs.p95, microsoftExcelWeb.scrollEventResponseMs!.p95)
      : Number.POSITIVE_INFINITY,
  }
}

function ratio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}

function validateSameCorpusCase(entry: UiResponsivenessSameCorpusCase): void {
  if (entry.materializedCells <= 0 || !Number.isInteger(entry.materializedCells)) {
    throw new Error(`UI responsiveness same-corpus case has invalid materialized cell count: ${entry.id}`)
  }
  validateSameCorpusMeasurement(entry.bilig, 'bilig', entry.id)
  validateSameCorpusMeasurement(entry.googleSheets, 'google-sheets', entry.id)
  if (entry.microsoftExcelWeb) {
    validateSameCorpusMeasurement(entry.microsoftExcelWeb, 'microsoft-excel-web', entry.id)
  }
  if (
    uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload) &&
    ![entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])].every((measurement) =>
      hasSameCorpusScrollEvidence(measurement),
    )
  ) {
    throw new Error(`UI responsiveness same-corpus proof is missing scroll-event evidence for ${entry.id}`)
  }
  const comparableSampleCount = Math.min(
    entry.bilig.operationResponseMs.samples.length,
    entry.googleSheets.operationResponseMs.samples.length,
    ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb.operationResponseMs.samples.length] : []),
  )
  if (entry.sampleCount !== comparableSampleCount || comparableSampleCount < sameCorpusSampleCount) {
    throw new Error(`UI responsiveness same-corpus case has too few comparable samples: ${entry.id}`)
  }
  const googleSheetsMeanRatio = ratio(entry.bilig.operationResponseMs.mean, entry.googleSheets.operationResponseMs.mean)
  const googleSheetsP95Ratio = ratio(entry.bilig.operationResponseMs.p95, entry.googleSheets.operationResponseMs.p95)
  const microsoftExcelWebMeanRatio = entry.microsoftExcelWeb
    ? ratio(entry.bilig.operationResponseMs.mean, entry.microsoftExcelWeb.operationResponseMs.mean)
    : undefined
  const microsoftExcelWebP95Ratio = entry.microsoftExcelWeb
    ? ratio(entry.bilig.operationResponseMs.p95, entry.microsoftExcelWeb.operationResponseMs.p95)
    : undefined
  if (
    entry.biligToGoogleSheetsMeanRatio !== googleSheetsMeanRatio ||
    entry.biligToGoogleSheetsP95Ratio !== googleSheetsP95Ratio ||
    entry.biligToMicrosoftExcelWebMeanRatio !== microsoftExcelWebMeanRatio ||
    entry.biligToMicrosoftExcelWebP95Ratio !== microsoftExcelWebP95Ratio
  ) {
    throw new Error(`UI responsiveness same-corpus ratio is stale: ${entry.id}`)
  }
  const scrollEventMetrics = sameCorpusScrollEventMetrics(entry.bilig, entry.googleSheets, entry.microsoftExcelWeb)
  const comparedProducts = [entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])]
  const postOperationFrameGuardrailPassed = comparedProducts.every(
    (measurement) => measurement.postOperationFrameMs.p95 > 0 && measurement.postOperationFrameMs.p95 <= 50,
  )
  const scrollMovementGuardrailPassed =
    scrollEventMetrics !== null && comparedProducts.every((measurement) => (measurement.scrollMovementPx?.min ?? 0) >= 1)
  if (scrollEventMetrics) {
    if (
      entry.biligToGoogleSheetsScrollEventMeanRatio !== scrollEventMetrics.biligToGoogleSheetsMeanRatio ||
      entry.biligToGoogleSheetsScrollEventP95Ratio !== scrollEventMetrics.biligToGoogleSheetsP95Ratio ||
      (entry.microsoftExcelWeb &&
        (entry.biligToMicrosoftExcelWebScrollEventMeanRatio !== scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio ||
          entry.biligToMicrosoftExcelWebScrollEventP95Ratio !== scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio))
    ) {
      throw new Error(`UI responsiveness same-corpus scroll-event ratio is stale: ${entry.id}`)
    }
  }
  if (
    entry.postOperationFrameGuardrailPassed !== undefined &&
    entry.postOperationFrameGuardrailPassed !== postOperationFrameGuardrailPassed
  ) {
    throw new Error(`UI responsiveness same-corpus post-frame guardrail is stale: ${entry.id}`)
  }
  if (entry.scrollMovementGuardrailPassed !== undefined && entry.scrollMovementGuardrailPassed !== scrollMovementGuardrailPassed) {
    throw new Error(`UI responsiveness same-corpus scroll-movement guardrail is stale: ${entry.id}`)
  }
  const requiresScrollEventMetric = uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload)
  const expectedMetric = requiresScrollEventMetric ? 'scrollEventResponseMs' : 'operationResponseMs'
  if (entry.tenXMeanAndP95Metric !== expectedMetric) {
    throw new Error(`UI responsiveness same-corpus metric is stale: ${entry.id}`)
  }
  validateSameCorpusScenarioProof(entry.scenarioProof, entry.id, entry.bilig, entry.googleSheets, entry.microsoftExcelWeb)
  const visualProofGuardrailPassed = entry.scenarioProof.screenshotProof.captured && entry.scenarioProof.pixelGridProof.captured
  const timingMetricPassedAgainstGoogleSheets = requiresScrollEventMetric
    ? scrollEventMetrics !== null &&
      scrollEventMetrics.biligToGoogleSheetsMeanRatio <= 0.1 &&
      scrollEventMetrics.biligToGoogleSheetsP95Ratio <= 0.1 &&
      scrollMovementGuardrailPassed
    : googleSheetsMeanRatio <= 0.1 && googleSheetsP95Ratio <= 0.1
  const timingMetricPassedAgainstMicrosoftExcelWeb = entry.microsoftExcelWeb
    ? requiresScrollEventMetric
      ? scrollEventMetrics !== null &&
        scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio <= 0.1 &&
        scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio <= 0.1 &&
        scrollMovementGuardrailPassed
      : (microsoftExcelWebMeanRatio ?? Number.POSITIVE_INFINITY) <= 0.1 && (microsoftExcelWebP95Ratio ?? Number.POSITIVE_INFINITY) <= 0.1
    : undefined
  const tenXAgainstGoogleSheets = timingMetricPassedAgainstGoogleSheets && postOperationFrameGuardrailPassed && visualProofGuardrailPassed
  const tenXAgainstMicrosoftExcelWeb =
    timingMetricPassedAgainstMicrosoftExcelWeb === undefined
      ? undefined
      : timingMetricPassedAgainstMicrosoftExcelWeb && postOperationFrameGuardrailPassed && visualProofGuardrailPassed
  if (
    entry.tenXMeanAndP95AgainstGoogleSheets !== tenXAgainstGoogleSheets ||
    entry.tenXMeanAndP95AgainstMicrosoftExcelWeb !== tenXAgainstMicrosoftExcelWeb ||
    entry.passed !== tenXAgainstGoogleSheets
  ) {
    throw new Error(`UI responsiveness same-corpus pass flag is stale: ${entry.id}`)
  }
}

function hasSameCorpusScrollEvidence(measurement: UiResponsivenessSameCorpusMeasurement): boolean {
  return Boolean(
    measurement.scrollEventResponseMs &&
    measurement.scrollMovementPx &&
    measurement.scrollEventResponseMs.samples.length >= sameCorpusSampleCount &&
    measurement.scrollMovementPx.samples.length >= sameCorpusSampleCount &&
    measurement.scrollMovementPx.min >= 1,
  )
}

function validateSameCorpusMeasurement(
  measurement: UiResponsivenessSameCorpusMeasurement,
  product: UiResponsivenessSameCorpusProduct,
  caseId: string,
): void {
  if (measurement.product !== product) {
    throw new Error(`UI responsiveness same-corpus product mismatch for ${caseId}`)
  }
  if (measurement.source.length === 0) {
    throw new Error(`UI responsiveness same-corpus source is missing for ${caseId}`)
  }
  validateSummary(measurement.operationResponseMs, `${caseId} ${product} operationResponseMs`)
  validateSummary(measurement.postOperationFrameMs, `${caseId} ${product} postOperationFrameMs`)
  if (measurement.scrollEventResponseMs) {
    validateSummary(measurement.scrollEventResponseMs, `${caseId} ${product} scrollEventResponseMs`)
  }
  if (measurement.scrollMovementPx) {
    validateSummary(measurement.scrollMovementPx, `${caseId} ${product} scrollMovementPx`)
  }
  validateSameCorpusCaptureVerification(measurement.corpusVerification, product, null, caseId)
}

function validateSameCorpusCaptureVerification(
  verification: SameCorpusCaptureCorpusVerification,
  product: UiResponsivenessSameCorpusProduct,
  expectedMaterializedCells: number | null,
  caseId: string,
): void {
  if (!verification.verified) {
    throw new Error(`UI responsiveness same-corpus verification is not marked verified for ${caseId} ${product}`)
  }
  if (expectedMaterializedCells !== null && verification.materializedCells !== expectedMaterializedCells) {
    throw new Error(`UI responsiveness same-corpus verification materialized cell count mismatch for ${caseId} ${product}`)
  }
  if (product === 'bilig' && verification.method !== 'bilig-benchmark-state') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'google-sheets' && verification.method !== 'google-sheets-xlsx-export') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'microsoft-excel-web' && verification.method !== 'microsoft-excel-web-source-xlsx') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product !== 'bilig' && verification.checkedCells.length < 3) {
    throw new Error(`UI responsiveness same-corpus verification must check at least 3 cells for ${caseId} ${product}`)
  }
  for (const cell of verification.checkedCells) {
    if (cell.address.trim().length === 0 || cell.expected !== cell.actual) {
      throw new Error(`UI responsiveness same-corpus verification cell mismatch for ${caseId} ${product}`)
    }
  }
}

function cloneSameCorpusVerification(verification: SameCorpusCaptureCorpusVerification): SameCorpusCaptureCorpusVerification {
  return {
    verified: verification.verified,
    method: verification.method,
    sheetName: verification.sheetName,
    materializedCells: verification.materializedCells,
    checkedCells: verification.checkedCells.map((cell) => ({ ...cell })),
  }
}

function validateSummary(summary: NumericSummary, label: string): void {
  if (summary.samples.length < sameCorpusSampleCount) {
    throw new Error(`UI responsiveness same-corpus scorecard has too few samples for ${label}`)
  }
  for (const value of [summary.min, summary.median, summary.p95, summary.max, summary.mean, ...summary.samples]) {
    if (!Number.isFinite(value) || value < 0) {
      throw new Error(`UI responsiveness same-corpus scorecard has invalid numeric summary for ${label}`)
    }
  }
}
