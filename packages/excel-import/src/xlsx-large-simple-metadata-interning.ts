import type {
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  CellRangeRef,
  LiteralInput,
  WorkbookAutoFilterSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
} from '@bilig/protocol'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import type { LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'

export function internLargeSimpleWorksheetMetadata(
  metadata: LargeSimpleWorksheetScannedMetadata | undefined,
  stringPool: ImportedWorkbookStringPool | undefined,
): LargeSimpleWorksheetScannedMetadata | undefined {
  if (!metadata || !stringPool) {
    return metadata
  }
  const intern = (value: string): string => stringPool.intern(value)
  return {
    ...(metadata.cellMetadataRefs && metadata.cellMetadataRefs.length > 0
      ? {
          cellMetadataRefs: metadata.cellMetadataRefs.map((ref) => ({
            address: intern(ref.address),
            ...(ref.cm ? { cm: intern(ref.cm) } : {}),
            ...(ref.vm ? { vm: intern(ref.vm) } : {}),
          })),
        }
      : {}),
    ...(metadata.columns ? { columns: internAxisMetadata(metadata.columns, intern) } : {}),
    ...(metadata.conditionalFormats && metadata.conditionalFormats.length > 0
      ? { conditionalFormats: metadata.conditionalFormats.map((format) => internConditionalFormat(format, intern)) }
      : {}),
    ...(metadata.conditionalFormattingXml && metadata.conditionalFormattingXml.length > 0
      ? { conditionalFormattingXml: metadata.conditionalFormattingXml.map(intern) }
      : {}),
    ...(metadata.controlArtifacts
      ? {
          controlArtifacts: {
            controlsXml: intern(metadata.controlArtifacts.controlsXml),
            worksheetRootOpenTag: intern(metadata.controlArtifacts.worksheetRootOpenTag),
            ...(metadata.controlArtifacts.legacyDrawingRelationshipId
              ? { legacyDrawingRelationshipId: intern(metadata.controlArtifacts.legacyDrawingRelationshipId) }
              : {}),
          },
        }
      : {}),
    ...(metadata.dataValidations && metadata.dataValidations.length > 0
      ? { dataValidations: metadata.dataValidations.map((validation) => internDataValidation(validation, intern)) }
      : {}),
    ...(metadata.drawingRelationshipId ? { drawingRelationshipId: intern(metadata.drawingRelationshipId) } : {}),
    ...(metadata.legacyDrawingRelationshipId ? { legacyDrawingRelationshipId: intern(metadata.legacyDrawingRelationshipId) } : {}),
    ...(metadata.filters && metadata.filters.length > 0 ? { filters: metadata.filters.map((filter) => internFilter(filter, intern)) } : {}),
    ...(metadata.hyperlinks && metadata.hyperlinks.length > 0
      ? {
          hyperlinks: metadata.hyperlinks.map((hyperlink) => ({
            ref: intern(hyperlink.ref),
            ...(hyperlink.relationshipId ? { relationshipId: intern(hyperlink.relationshipId) } : {}),
            ...(hyperlink.location ? { location: intern(hyperlink.location) } : {}),
            ...(hyperlink.tooltip ? { tooltip: intern(hyperlink.tooltip) } : {}),
            ...(hyperlink.display ? { display: intern(hyperlink.display) } : {}),
          })),
        }
      : {}),
    ...(metadata.rows ? { rows: internAxisMetadata(metadata.rows, intern) } : {}),
    ...(metadata.merges && metadata.merges.length > 0
      ? {
          merges: metadata.merges.map((range) => ({
            startAddress: intern(range.startAddress),
            endAddress: intern(range.endAddress),
          })),
        }
      : {}),
    ...(metadata.printPageSetup
      ? {
          printPageSetup: {
            ...(metadata.printPageSetup.printOptionsXml ? { printOptionsXml: intern(metadata.printPageSetup.printOptionsXml) } : {}),
            ...(metadata.printPageSetup.pageMarginsXml ? { pageMarginsXml: intern(metadata.printPageSetup.pageMarginsXml) } : {}),
            ...(metadata.printPageSetup.pageSetupXml ? { pageSetupXml: intern(metadata.printPageSetup.pageSetupXml) } : {}),
            ...(metadata.printPageSetup.headerFooterXml ? { headerFooterXml: intern(metadata.printPageSetup.headerFooterXml) } : {}),
            ...(metadata.printPageSetup.rowBreaksXml ? { rowBreaksXml: intern(metadata.printPageSetup.rowBreaksXml) } : {}),
            ...(metadata.printPageSetup.colBreaksXml ? { colBreaksXml: intern(metadata.printPageSetup.colBreaksXml) } : {}),
          },
        }
      : {}),
    ...(metadata.sheetFormatPr ? { sheetFormatPr: metadata.sheetFormatPr } : {}),
    ...(metadata.sheetSlicerListExtXml ? { sheetSlicerListExtXml: intern(metadata.sheetSlicerListExtXml) } : {}),
    ...(metadata.tableRelationshipIds && metadata.tableRelationshipIds.length > 0
      ? { tableRelationshipIds: metadata.tableRelationshipIds.map(intern) }
      : {}),
  }
}

function internAxisMetadata(
  axis: { readonly entries: readonly WorkbookAxisEntrySnapshot[]; readonly metadata: readonly WorkbookAxisMetadataSnapshot[] },
  intern: (value: string) => string,
): { readonly entries: WorkbookAxisEntrySnapshot[]; readonly metadata: WorkbookAxisMetadataSnapshot[] } {
  return {
    entries: axis.entries.map((entry) => ({
      ...entry,
      id: intern(entry.id),
    })),
    metadata: [...axis.metadata],
  }
}

function internConditionalFormat(
  format: WorkbookConditionalFormatSnapshot,
  intern: (value: string) => string,
): WorkbookConditionalFormatSnapshot {
  return {
    ...format,
    id: intern(format.id),
    range: internRange(format.range, intern),
    rule: internConditionalFormatRule(format.rule, intern),
  }
}

function internConditionalFormatRule(
  rule: WorkbookConditionalFormatSnapshot['rule'],
  intern: (value: string) => string,
): WorkbookConditionalFormatSnapshot['rule'] {
  switch (rule.kind) {
    case 'cellIs':
      return { ...rule, values: rule.values.map((value) => internLiteral(value, intern)) }
    case 'formula':
      return { ...rule, formula: intern(rule.formula) }
    case 'textContains':
      return { ...rule, text: intern(rule.text) }
    case 'blanks':
    case 'notBlanks':
      return rule
  }
}

function internDataValidation(
  validation: WorkbookDataValidationSnapshot,
  intern: (value: string) => string,
): WorkbookDataValidationSnapshot {
  return {
    ...validation,
    range: internRange(validation.range, intern),
    rule: internDataValidationRule(validation.rule, intern),
    ...(validation.promptTitle ? { promptTitle: intern(validation.promptTitle) } : {}),
    ...(validation.promptMessage ? { promptMessage: intern(validation.promptMessage) } : {}),
    ...(validation.errorTitle ? { errorTitle: intern(validation.errorTitle) } : {}),
    ...(validation.errorMessage ? { errorMessage: intern(validation.errorMessage) } : {}),
  }
}

function internDataValidationRule(
  rule: WorkbookDataValidationSnapshot['rule'],
  intern: (value: string) => string,
): WorkbookDataValidationSnapshot['rule'] {
  switch (rule.kind) {
    case 'list':
      return {
        ...rule,
        ...(rule.values ? { values: rule.values.map((value) => internLiteral(value, intern)) } : {}),
        ...(rule.source ? { source: internValidationSource(rule.source, intern) } : {}),
      }
    case 'checkbox':
      return {
        ...rule,
        ...(rule.checkedValue !== undefined ? { checkedValue: internLiteral(rule.checkedValue, intern) } : {}),
        ...(rule.uncheckedValue !== undefined ? { uncheckedValue: internLiteral(rule.uncheckedValue, intern) } : {}),
      }
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength':
      return { ...rule, values: rule.values.map((value) => internLiteral(value, intern)) }
    case 'any':
      return rule
  }
}

function internValidationSource(
  source: NonNullable<Extract<WorkbookDataValidationSnapshot['rule'], { kind: 'list' }>['source']>,
  intern: (value: string) => string,
): NonNullable<Extract<WorkbookDataValidationSnapshot['rule'], { kind: 'list' }>['source']> {
  switch (source.kind) {
    case 'cell-ref':
      return { ...source, sheetName: intern(source.sheetName), address: intern(source.address) }
    case 'range-ref':
      return {
        ...source,
        sheetName: intern(source.sheetName),
        startAddress: intern(source.startAddress),
        endAddress: intern(source.endAddress),
      }
    case 'named-range':
      return { ...source, name: intern(source.name) }
    case 'structured-ref':
      return { ...source, tableName: intern(source.tableName), columnName: intern(source.columnName) }
  }
}

function internFilter(filter: WorkbookAutoFilterSnapshot, intern: (value: string) => string): WorkbookAutoFilterSnapshot {
  return {
    ...internRange(filter, intern),
    ...(filter.criteria
      ? {
          criteria: filter.criteria.map((criterion) => ({
            ...criterion,
            ...(criterion.filters
              ? {
                  filters: {
                    ...criterion.filters,
                    values: criterion.filters.values.map(intern),
                  },
                }
              : {}),
            ...(criterion.customFilters
              ? {
                  customFilters: {
                    ...criterion.customFilters,
                    filters: criterion.customFilters.filters.map((customFilter) => ({
                      ...customFilter,
                      value: intern(customFilter.value),
                    })),
                  },
                }
              : {}),
          })),
        }
      : {}),
  }
}

function internRange(range: CellRangeRef, intern: (value: string) => string): CellRangeRef {
  return {
    sheetName: intern(range.sheetName),
    startAddress: intern(range.startAddress),
    endAddress: intern(range.endAddress),
  }
}

function internLiteral(value: LiteralInput, intern: (value: string) => string): LiteralInput {
  return typeof value === 'string' ? intern(value) : value
}
