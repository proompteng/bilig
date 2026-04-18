import type { CompiledFormula } from '@bilig/formula'
import { createTemplateBank, type FormulaTemplateResolution, type FormulaTemplateSnapshot } from '../../formula/template-bank.js'
import type { EngineCounters } from '../../perf/engine-counters.js'

export interface EngineFormulaTemplateNormalizationService {
  readonly clear: () => void
  readonly reset: () => void
  readonly compileForCell: (source: string, ownerRow: number, ownerCol: number) => CompiledFormula
  readonly resolveForCell: (source: string, ownerRow: number, ownerCol: number) => FormulaTemplateResolution
  readonly resolveByTemplateId: (
    templateId: number,
    source: string,
    ownerRow: number,
    ownerCol: number,
  ) => FormulaTemplateResolution | undefined
  readonly listTemplates: () => FormulaTemplateSnapshot[]
  readonly hydrateTemplates: (snapshots: readonly FormulaTemplateSnapshot[]) => void
}

export function createEngineFormulaTemplateNormalizationService(args?: {
  readonly counters?: EngineCounters
}): EngineFormulaTemplateNormalizationService {
  const templateBank = createTemplateBank(args)

  return {
    clear() {
      templateBank.clearTransientCache()
    },
    reset() {
      templateBank.reset()
    },
    compileForCell(source, ownerRow, ownerCol) {
      return templateBank.resolve(source, ownerRow, ownerCol).compiled
    },
    resolveForCell(source, ownerRow, ownerCol) {
      return templateBank.resolve(source, ownerRow, ownerCol)
    },
    resolveByTemplateId(templateId, source, ownerRow, ownerCol) {
      return templateBank.resolveById(templateId, source, ownerRow, ownerCol)
    },
    listTemplates() {
      return templateBank.list()
    },
    hydrateTemplates(snapshots) {
      templateBank.hydrate(snapshots)
    },
  }
}
