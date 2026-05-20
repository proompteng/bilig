export type ImportExportSemanticDisposition = 'preserved' | 'unsupported' | 'external' | 'declined-runtime'

export interface ImportExportSemanticLedgerEntry {
  readonly feature: string
  readonly disposition: ImportExportSemanticDisposition
  readonly reason: string
}

export const importExportSemanticLossLedger: readonly ImportExportSemanticLedgerEntry[] = [
  {
    feature: 'xlsx.macros.execution',
    disposition: 'declined-runtime',
    reason: 'Bilig preserves macro payload metadata but intentionally never executes workbook macros.',
  },
]

export function buildImportExportSemanticLedger(coveredFeatures: readonly string[]): readonly ImportExportSemanticLedgerEntry[] {
  const lossFeatures = new Set(importExportSemanticLossLedger.map((entry) => entry.feature))
  return [
    ...coveredFeatures
      .filter((feature) => !lossFeatures.has(feature))
      .map((feature) => ({
        feature,
        disposition: importExportCoveredFeatureDisposition(feature),
        reason: importExportCoveredFeatureReason(feature),
      })),
    ...importExportSemanticLossLedger,
  ].toSorted((left, right) => left.feature.localeCompare(right.feature))
}

export function importExportUnsupportedFeatures(
  ledger: readonly ImportExportSemanticLedgerEntry[] = importExportSemanticLossLedger,
): readonly string[] {
  return ledger
    .filter((entry) => entry.disposition === 'unsupported')
    .map((entry) => entry.feature)
    .toSorted()
}

export function importExportDeclinedRuntimeFeatures(
  ledger: readonly ImportExportSemanticLedgerEntry[] = importExportSemanticLossLedger,
): readonly string[] {
  return ledger
    .filter((entry) => entry.disposition === 'declined-runtime')
    .map((entry) => entry.feature)
    .toSorted()
}

function importExportCoveredFeatureDisposition(feature: string): ImportExportSemanticDisposition {
  return feature.startsWith('external.') ? 'external' : 'preserved'
}

function importExportCoveredFeatureReason(feature: string): string {
  if (feature.startsWith('external.')) {
    return 'Tracked by the official Google Sheets and Microsoft Excel import/export comparison artifact.'
  }
  return 'Preserved by required import/export fidelity case evidence.'
}
