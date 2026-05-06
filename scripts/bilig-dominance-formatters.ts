import type {
  CompetitiveFamilySummary,
  HeadedBrowserFrameP95Contract,
  LargeWorkbookSloMeasurement,
  RatioSummary,
} from './bilig-dominance-scorecard-types.ts'

export function familyWinSummary(family: CompetitiveFamilySummary): string {
  const worstMean =
    family.worstWorkpaperToHyperFormulaMeanRatio === null || family.worstMeanRatioWorkload === null
      ? 'no comparable mean ratio'
      : `worst mean ratio ${family.worstWorkpaperToHyperFormulaMeanRatio} on ${family.worstMeanRatioWorkload}`
  const worstP95 =
    family.worstWorkpaperToHyperFormulaP95Ratio === null || family.worstP95RatioWorkload === null
      ? 'no comparable p95 ratio'
      : `worst p95 ratio ${family.worstWorkpaperToHyperFormulaP95Ratio} on ${family.worstP95RatioWorkload}`
  return `${family.family}: WorkPaper ${family.workpaperWins}/${family.comparableCount}, HyperFormula ${family.hyperformulaWins}/${family.comparableCount}; ${worstMean}; ${worstP95}`
}

export function formulaMissingFunctionBlockers(missingOfficeFunctionCount: number): string[] {
  return missingOfficeFunctionCount > 0
    ? [`${String(missingOfficeFunctionCount)} Office-listed functions are still missing from the runtime inventory`]
    : []
}

export function sloSummary(measurement: LargeWorkbookSloMeasurement): string {
  return `${measurement.id}: ${measurement.metric} ${measurement.actualP95}ms against ${measurement.budgetP95}ms SLO (${measurement.sampleCount} samples)`
}

export function headedBrowserContractSummary(contract: HeadedBrowserFrameP95Contract): string {
  return `${contract.id}: ${contract.metric} budget ${contract.budgetP95}ms on ${contract.materializedCells} materialized cells via ${contract.command}`
}

export function formatRatio(summary: RatioSummary): string {
  return `${summary.production}/${summary.total} (${summary.percent}%)`
}

export function formatList(values: readonly string[]): string {
  return values.length === 0 ? 'none' : values.join(', ')
}
