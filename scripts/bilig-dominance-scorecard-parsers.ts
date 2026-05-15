import type {
  CompetitiveArtifact,
  CompetitiveFamilySummary,
  CompetitiveResult,
  FormulaDominanceSnapshot,
  HyperFormulaSurfaceSnapshot,
  RatioSummary,
} from './bilig-dominance-scorecard-types.ts'
import {
  arrayField,
  asObject,
  booleanField,
  literalField,
  numberField,
  objectField,
  optionalNumberField,
  optionalStringField,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'

export function parseFormulaDominanceSnapshot(value: Record<string, unknown>): FormulaDominanceSnapshot {
  const formulaBreadth = objectField(value, 'formulaBreadth')
  const canonical = objectField(value, 'canonical')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    formulaBreadth: {
      officeListed: ratioField(formulaBreadth, 'officeListed'),
      tracked: ratioField(formulaBreadth, 'tracked'),
      missingOfficeFunctions: stringArrayField(formulaBreadth, 'missingOfficeFunctions'),
    },
    canonical: {
      summary: ratioField(canonical, 'summary'),
      nonProductionRows: arrayField(canonical, 'nonProductionRows'),
    },
  }
}

export function parseSurfaceSnapshot(value: Record<string, unknown>): HyperFormulaSurfaceSnapshot {
  const classSurface = objectField(value, 'classSurface')
  return {
    hyperFormulaVersion: stringField(value, 'hyperFormulaVersion'),
    hyperFormulaCommit: stringField(value, 'hyperFormulaCommit'),
    classSurface: {
      staticMembers: stringArrayField(classSurface, 'staticMembers'),
      staticMethods: stringArrayField(classSurface, 'staticMethods'),
      instanceAccessors: stringArrayField(classSurface, 'instanceAccessors'),
      instanceMethods: stringArrayField(classSurface, 'instanceMethods'),
    },
    configKeys: stringArrayField(value, 'configKeys'),
  }
}

export function parseCompetitiveArtifact(value: Record<string, unknown>): CompetitiveArtifact {
  const engines = objectField(value, 'engines')
  const hyperformula = objectField(engines, 'hyperformula')
  const scorecard = objectField(value, 'scorecard')
  return {
    generatedAt: stringField(value, 'generatedAt'),
    engines: {
      hyperformula: {
        commit: stringField(hyperformula, 'commit'),
        version: stringField(hyperformula, 'version'),
      },
    },
    families: arrayField(value, 'families').map(parseCompetitiveFamily),
    results: arrayField(value, 'results').map(parseCompetitiveResult),
    scorecard: {
      comparableCount: numberField(scorecard, 'comparableCount'),
      directionalMeanRatioGeomean: numberField(scorecard, 'directionalMeanRatioGeomean'),
      directionalP95RatioGeomean: numberField(scorecard, 'directionalP95RatioGeomean'),
      hyperformulaWins: numberField(scorecard, 'hyperformulaWins'),
      worstMeanRatioWorkload: stringField(scorecard, 'worstMeanRatioWorkload'),
      worstP95RatioWorkload: stringField(scorecard, 'worstP95RatioWorkload'),
      worstWorkpaperToHyperFormulaMeanRatio: numberField(scorecard, 'worstWorkpaperToHyperFormulaMeanRatio'),
      worstWorkpaperToHyperFormulaP95Ratio: numberField(scorecard, 'worstWorkpaperToHyperFormulaP95Ratio'),
      workpaperWins: numberField(scorecard, 'workpaperWins'),
    },
  }
}

function parseCompetitiveFamily(value: unknown): CompetitiveFamilySummary {
  const family = asObject(value, 'competitive family')
  const parsed: CompetitiveFamilySummary = {
    comparableCount: numberField(family, 'comparableCount'),
    family: stringField(family, 'family'),
    hyperformulaWins: numberField(family, 'hyperformulaWins'),
    scorecardEligible: booleanField(family, 'scorecardEligible'),
    workpaperWins: numberField(family, 'workpaperWins'),
    worstMeanRatioWorkload: optionalStringField(family, 'worstMeanRatioWorkload'),
    worstP95RatioWorkload: optionalStringField(family, 'worstP95RatioWorkload'),
    worstWorkpaperToHyperFormulaMeanRatio: optionalNumberField(family, 'worstWorkpaperToHyperFormulaMeanRatio'),
    worstWorkpaperToHyperFormulaP95Ratio: optionalNumberField(family, 'worstWorkpaperToHyperFormulaP95Ratio'),
  }
  if ('workloads' in family) {
    parsed.workloads = stringArrayField(family, 'workloads')
  }
  return parsed
}

function parseCompetitiveResult(value: unknown): CompetitiveResult {
  const result = asObject(value, 'competitive result')
  const comparison = result['comparison']
  return {
    comparable: booleanField(result, 'comparable'),
    workload: stringField(result, 'workload'),
    comparison:
      comparison === undefined
        ? undefined
        : {
            workpaperToHyperFormulaMeanRatio: numberField(
              asObject(comparison, 'competitive comparison'),
              'workpaperToHyperFormulaMeanRatio',
            ),
            workpaperToHyperFormulaP95Ratio: numberField(asObject(comparison, 'competitive comparison'), 'workpaperToHyperFormulaP95Ratio'),
          },
  }
}

function ratioField(value: Record<string, unknown>, field: string): RatioSummary {
  const ratioValue = objectField(value, field)
  return {
    percent: numberField(ratioValue, 'percent'),
    production: numberField(ratioValue, 'production'),
    total: numberField(ratioValue, 'total'),
  }
}
