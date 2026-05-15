export type HeadlessComparisonCoverageTier = 'workbook-wide' | 'scalar-formula'

export interface ExtraHeadlessComparisonEngineSummary {
  readonly artifactPath: string
  readonly comparableWorkloadCount: number
  readonly coverageNote?: string
  readonly coverageTier: HeadlessComparisonCoverageTier
  readonly engineName: string
  readonly generatedAt: string
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly version: string
  readonly workloadFamilies: readonly string[]
}
