import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import type { PublicWorkbookExternalReferenceSummary } from './public-workbook-corpus-types.ts'

export function summarizeExternalWorkbookReferences(snapshot: WorkbookSnapshot): PublicWorkbookExternalReferenceSummary | undefined {
  const dependencies = snapshot.workbook.metadata?.unsupportedFormulaDependencies ?? []
  const externalWorkbookReferences = snapshot.workbook.metadata?.externalWorkbookReferences ?? []
  if (dependencies.length === 0 && externalWorkbookReferences.length === 0) {
    return undefined
  }
  const linkedWorkbookKeys = new Set<string>()
  for (const linkedWorkbook of externalWorkbookReferences) {
    linkedWorkbookKeys.add(
      `${String(linkedWorkbook.bookIndex)}\t${linkedWorkbook.target ?? ''}\t${linkedWorkbook.packagePath ?? ''}\t${linkedWorkbook.workbookName ?? ''}`,
    )
  }
  for (const dependency of dependencies) {
    for (const linkedWorkbook of dependency.linkedWorkbooks) {
      linkedWorkbookKeys.add(
        `${String(linkedWorkbook.bookIndex)}\t${linkedWorkbook.target ?? ''}\t${linkedWorkbook.packagePath ?? ''}\t${linkedWorkbook.workbookName ?? ''}`,
      )
    }
  }
  return {
    linkedWorkbookCount: linkedWorkbookKeys.size,
    formulaDependencyCount: dependencies.length,
    cachedValueDependencyCount: dependencies.filter((dependency) => dependency.cachedValuesUsed).length,
  }
}

export function unsupportedWorkbookMetadataEvidence(
  snapshot: WorkbookSnapshot,
  validation: { readonly skippedUnsupportedFormulaCount: number },
): readonly string[] {
  const evidence: string[] = []
  const unsupportedFormulaDependencies = snapshot.workbook.metadata?.unsupportedFormulaDependencies ?? []
  const externalWorkbookReferences = summarizeExternalWorkbookReferences(snapshot)
  if (externalWorkbookReferences) {
    evidence.push(`external-workbook-links=${String(externalWorkbookReferences.linkedWorkbookCount)}`)
    evidence.push(`external-workbook-formula-dependencies=${String(externalWorkbookReferences.formulaDependencyCount)}`)
    evidence.push(`external-workbook-cached-value-dependencies=${String(externalWorkbookReferences.cachedValueDependencyCount)}`)
    evidence.push(
      ...(snapshot.workbook.metadata?.externalWorkbookReferences ?? [])
        .slice(0, 10)
        .map(
          (entry) => `external-workbook=${entry.workbookName ?? entry.target ?? entry.packagePath ?? `book#${String(entry.bookIndex)}`}`,
        ),
    )
  }
  if (unsupportedFormulaDependencies.length > 0) {
    if (validation.skippedUnsupportedFormulaCount > 0) {
      evidence.push(`formula-oracle-skipped-unsupported-external-formulas=${String(validation.skippedUnsupportedFormulaCount)}`)
    }
    evidence.push(
      ...unsupportedFormulaDependencies
        .slice(0, 10)
        .map(
          (entry) =>
            `external-formula=${entry.sheetName}!${entry.address} linked=${formatLinkedWorkbookReferences(
              entry.linkedWorkbooks,
            )} cached=${String(entry.cachedValuesUsed)} resolved=${String(entry.resolvedExternalReferenceCount)} unresolved=${String(
              entry.unresolvedExternalReferenceCount,
            )}`,
        ),
    )
  }
  const unsupportedPivots = snapshot.workbook.metadata?.unsupportedPivots ?? []
  if (unsupportedPivots.length > 0) {
    evidence.push(`unsupported-pivots=${String(unsupportedPivots.length)}`)
    evidence.push(
      ...unsupportedPivots
        .slice(0, 10)
        .map(
          (entry) =>
            `unsupported-pivot=${entry.kind}:${entry.sheetName ?? '<workbook>'}!${entry.address ?? '<unknown>'}:cache=${String(
              entry.cacheId ?? '<unknown>',
            )}`,
        ),
    )
  }
  return evidence
}

export function formatLinkedWorkbookReferences(
  linkedWorkbooks: readonly {
    readonly bookIndex: number
    readonly workbookName?: string
    readonly target?: string
    readonly packagePath?: string
  }[],
): string {
  if (linkedWorkbooks.length === 0) {
    return '[]'
  }
  return `[${linkedWorkbooks
    .slice(0, 3)
    .map((entry) => entry.workbookName ?? entry.target ?? entry.packagePath ?? `book#${String(entry.bookIndex)}`)
    .join(',')}]`
}
