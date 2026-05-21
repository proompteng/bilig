import type {
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookAutoFilterSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
} from '@bilig/protocol'
import { readLargeSimpleAutoFiltersFromBytes } from './xlsx-large-simple-autofilter-byte-scan.js'
import {
  countLargeSimpleConditionalFormattingSqrefRangesFromRootTag,
  readLargeSimpleConditionalFormattingFromBytes,
  wrapLargeSimpleConditionalFormatRuleBytes,
} from './xlsx-large-simple-conditional-format-byte-scan.js'
import {
  countLargeSimpleDataValidationsFromBytes,
  readLargeSimpleDataValidationsFromBytes,
} from './xlsx-large-simple-data-validation-byte-scan.js'
import { readLargeSimpleSheetHyperlinkRefsFromBytes } from './xlsx-large-simple-hyperlinks.js'
import { appendLargeSimplePrintPageSetupElement, isLargeSimplePrintPageSetupElementName } from './xlsx-large-simple-printer-settings.js'
import { rowTagHasMetadataAttribute } from './xlsx-large-simple-row-metadata-scan.js'
import {
  appendLargeSimpleColumnMetadataFromBytes,
  appendLargeSimpleRowMetadataTagFromBytes,
  readLargeSimpleDrawingRelationshipIdTagFromBytes,
  readLargeSimpleMergeRefsFromBytes,
  readLargeSimpleSheetFormatPrTagFromBytes,
  readLargeSimpleTableRelationshipIdsFromBytes,
} from './xlsx-large-simple-metadata-byte-scan.js'
import { decodeBytes, encodeCellAddress } from './xlsx-large-simple-xml-byte-utils.js'
import { countOpeningTags, readXmlAttributeFromTag } from './xlsx-large-simple-worksheet-stream-xml.js'
import { readElementAttribute, readSlicerListExtensionXml } from './xlsx-large-simple-worksheet-stream-cell-readers.js'
import type {
  LargeSimpleWorksheetCellMetadataRef,
  LargeSimpleWorksheetMergeRef,
  LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'

export type StreamedMetadataElement = 'cols' | 'mergeCells' | 'tableParts'

export interface ActiveConditionalFormatting {
  readonly rootTag: Uint8Array
  ruleSeen: boolean
}

interface ConditionalFormattingScan {
  readonly ruleCount: number
  readonly conditionalFormats?: readonly WorkbookConditionalFormatSnapshot[]
  readonly artifactXml?: string
}

export class LargeSimpleWorksheetStreamMetadataCollector {
  mergeCount = 0
  conditionalFormatCount = 0
  dataValidationCount = 0
  tableCount = 0

  private columnEntries: WorkbookAxisEntrySnapshot[] | undefined
  private columnMetadata: WorkbookAxisMetadataSnapshot[] | undefined
  private conditionalFormats: WorkbookConditionalFormatSnapshot[] | undefined
  private conditionalFormatIdCounter = 0
  private conditionalFormattingXml: string[] | undefined
  private dataValidations: WorkbookDataValidationSnapshot[] | undefined
  private controlArtifactsXml: string[] | undefined
  private worksheetRootOpenTag: string | undefined
  private legacyDrawingRelationshipId: string | undefined
  private cellMetadataRefs: LargeSimpleWorksheetCellMetadataRef[] | undefined
  private drawingRelationshipId: string | undefined
  private filters: LargeSimpleWorksheetScannedMetadata['filters']
  private hyperlinks: LargeSimpleWorksheetScannedMetadata['hyperlinks']
  private rowEntries: WorkbookAxisEntrySnapshot[] | undefined
  private rowMetadata: WorkbookAxisMetadataSnapshot[] | undefined
  private mergeRefs: LargeSimpleWorksheetMergeRef[] | undefined
  private printPageSetup: LargeSimpleWorksheetScannedMetadata['printPageSetup']
  private pivotTableDefinitionsXml: string[] | undefined
  private sheetFormatPr: LargeSimpleWorksheetScannedMetadata['sheetFormatPr']
  private sheetSlicerListExtXml: string | undefined
  private tableRelationshipIds: string[] | undefined

  constructor(
    private readonly sheetName: string | undefined,
    private readonly retainMetadataXml: boolean,
  ) {}

  setWorksheetRootOpenTag(xml: string): void {
    this.worksheetRootOpenTag = xml
  }

  buildMetadataScan(): LargeSimpleWorksheetScannedMetadata | undefined {
    const columns =
      (this.columnEntries?.length ?? 0) > 0 || (this.columnMetadata?.length ?? 0) > 0
        ? { entries: this.columnEntries ?? [], metadata: this.columnMetadata ?? [] }
        : undefined
    const rows =
      (this.rowEntries?.length ?? 0) > 0 || (this.rowMetadata?.length ?? 0) > 0
        ? { entries: this.rowEntries ?? [], metadata: this.rowMetadata ?? [] }
        : undefined
    const metadata: LargeSimpleWorksheetScannedMetadata = {
      ...(this.cellMetadataRefs && this.cellMetadataRefs.length > 0 ? { cellMetadataRefs: this.cellMetadataRefs } : {}),
      ...(columns ? { columns } : {}),
      ...(this.conditionalFormats && this.conditionalFormats.length > 0 ? { conditionalFormats: this.conditionalFormats } : {}),
      ...(this.conditionalFormattingXml && this.conditionalFormattingXml.length > 0
        ? { conditionalFormattingXml: this.conditionalFormattingXml }
        : {}),
      ...(this.controlArtifactsXml && this.controlArtifactsXml.length > 0 && this.worksheetRootOpenTag
        ? {
            controlArtifacts: {
              controlsXml: this.controlArtifactsXml.join(''),
              worksheetRootOpenTag: this.worksheetRootOpenTag,
              ...(this.legacyDrawingRelationshipId ? { legacyDrawingRelationshipId: this.legacyDrawingRelationshipId } : {}),
            },
          }
        : {}),
      ...(this.dataValidations && this.dataValidations.length > 0 ? { dataValidations: this.dataValidations } : {}),
      ...(this.drawingRelationshipId ? { drawingRelationshipId: this.drawingRelationshipId } : {}),
      ...(this.legacyDrawingRelationshipId ? { legacyDrawingRelationshipId: this.legacyDrawingRelationshipId } : {}),
      ...(this.filters && this.filters.length > 0 ? { filters: this.filters } : {}),
      ...(this.hyperlinks && this.hyperlinks.length > 0 ? { hyperlinks: this.hyperlinks } : {}),
      ...(this.pivotTableDefinitionsXml && this.pivotTableDefinitionsXml.length > 0
        ? { pivotTableDefinitionsXml: this.pivotTableDefinitionsXml.join('') }
        : {}),
      ...(rows ? { rows } : {}),
      ...(this.mergeRefs && this.mergeRefs.length > 0 ? { merges: this.mergeRefs } : {}),
      ...(this.printPageSetup ? { printPageSetup: this.printPageSetup } : {}),
      ...(this.sheetFormatPr ? { sheetFormatPr: this.sheetFormatPr } : {}),
      ...(this.sheetSlicerListExtXml ? { sheetSlicerListExtXml: this.sheetSlicerListExtXml } : {}),
      ...(this.tableRelationshipIds && this.tableRelationshipIds.length > 0 ? { tableRelationshipIds: this.tableRelationshipIds } : {}),
    }
    return Object.keys(metadata).length > 0 ? metadata : undefined
  }

  collectRowMetadata(buffer: Uint8Array, nameEnd: number, tagEnd: number, currentRow: number): void {
    if (!rowTagHasMetadataAttribute(buffer, nameEnd, tagEnd)) {
      return
    }
    this.rowEntries ??= []
    this.rowMetadata ??= []
    appendLargeSimpleRowMetadataTagFromBytes(this.rowEntries, this.rowMetadata, buffer, nameEnd, tagEnd, currentRow)
  }

  collectCellMetadataRef(buffer: Uint8Array, row: number, column: number, nameEnd: number, tagEnd: number): void {
    if (!this.retainMetadataXml) {
      return
    }
    const cm = readXmlAttributeFromTag(buffer, nameEnd, tagEnd, 'cm')
    const vm = readXmlAttributeFromTag(buffer, nameEnd, tagEnd, 'vm')
    if (!cm && !vm) {
      return
    }
    this.cellMetadataRefs ??= []
    this.cellMetadataRefs.push({
      address: encodeCellAddress(row, column),
      ...(cm ? { cm } : {}),
      ...(vm ? { vm } : {}),
    })
  }

  collectTypedMetadataElement(localName: string, buffer: Uint8Array, startIndex: number, endIndex: number): boolean {
    if (localName === 'mergeCells') {
      const refs = readLargeSimpleMergeRefsFromBytes(buffer, startIndex, endIndex)
      this.mergeCount += refs.length
      if (refs.length > 0) {
        this.mergeRefs ??= []
        this.mergeRefs.push(...refs)
      }
      return true
    }
    if (localName === 'cols') {
      this.columnEntries ??= []
      this.columnMetadata ??= []
      appendLargeSimpleColumnMetadataFromBytes(this.columnEntries, this.columnMetadata, buffer, startIndex, endIndex)
      return true
    }
    if (localName === 'sheetFormatPr') {
      this.sheetFormatPr = readLargeSimpleSheetFormatPrTagFromBytes(buffer, startIndex, endIndex) ?? this.sheetFormatPr
      return true
    }
    if (localName === 'extLst') {
      const sheetSlicerListExtXml = readSlicerListExtensionXml(decodeBytes(buffer, startIndex, endIndex))
      if (!sheetSlicerListExtXml) {
        return false
      }
      this.sheetSlicerListExtXml = sheetSlicerListExtXml
      return true
    }
    if (localName === 'drawing') {
      this.drawingRelationshipId = readLargeSimpleDrawingRelationshipIdTagFromBytes(buffer, startIndex, endIndex)
      return true
    }
    if (localName === 'pivotTableDefinition') {
      this.pivotTableDefinitionsXml ??= []
      this.pivotTableDefinitionsXml.push(decodeBytes(buffer, startIndex, endIndex))
      return true
    }
    if (localName === 'legacyDrawing') {
      const tagXml = decodeBytes(buffer, startIndex, endIndex)
      this.legacyDrawingRelationshipId =
        readElementAttribute(tagXml, 'r:id') ?? readElementAttribute(tagXml, 'id') ?? this.legacyDrawingRelationshipId
      return true
    }
    if (localName === 'controls' || localName === 'oleObjects') {
      this.controlArtifactsXml ??= []
      this.controlArtifactsXml.push(decodeBytes(buffer, startIndex, endIndex))
      return true
    }
    if (localName === 'autoFilter') {
      if (!this.sheetName) {
        return false
      }
      const filters = readLargeSimpleAutoFiltersFromBytes(this.sheetName, buffer, startIndex, endIndex)
      this.filters = [...(this.filters ?? []), ...filters]
      return true
    }
    if (localName === 'hyperlinks') {
      const refs = readLargeSimpleSheetHyperlinkRefsFromBytes(buffer, startIndex, endIndex)
      if (refs === null) {
        return false
      }
      this.hyperlinks = [...(this.hyperlinks ?? []), ...refs]
      return true
    }
    if (isLargeSimplePrintPageSetupElementName(localName)) {
      this.printPageSetup ??= {}
      appendLargeSimplePrintPageSetupElement(this.printPageSetup, localName, decodeBytes(buffer, startIndex, endIndex))
      return true
    }
    if (localName === 'tableParts') {
      const relationshipIds = readLargeSimpleTableRelationshipIdsFromBytes(buffer, startIndex, endIndex)
      this.tableCount += relationshipIds.length
      if (relationshipIds.length > 0) {
        this.tableRelationshipIds ??= []
        this.tableRelationshipIds.push(...relationshipIds)
      }
      return true
    }
    if (localName === 'conditionalFormatting') {
      if (!this.sheetName) {
        return false
      }
      this.appendConditionalFormattingScan(
        readLargeSimpleConditionalFormattingFromBytes(this.sheetName, buffer, startIndex, endIndex, this.conditionalFormatIdCounter + 1),
      )
      return true
    }
    return false
  }

  addAutoFilter(filter: WorkbookAutoFilterSnapshot | null): void {
    if (this.retainMetadataXml && filter) {
      this.filters = [...(this.filters ?? []), filter]
    }
  }

  countMetadataElement(localName: string, buffer: Uint8Array, contentStart: number, contentEnd: number): void {
    if (localName === 'conditionalFormatting') {
      this.conditionalFormatCount += Math.max(1, countOpeningTags(buffer, contentStart, contentEnd, 'cfRule'))
      return
    }
    if (localName === 'mergeCells') {
      this.mergeCount += countOpeningTags(buffer, contentStart, contentEnd, 'mergeCell')
    } else if (localName === 'tableParts') {
      this.tableCount += countOpeningTags(buffer, contentStart, contentEnd, 'tablePart')
    } else if (localName === 'dataValidations') {
      this.dataValidationCount += countOpeningTags(buffer, contentStart, contentEnd, 'dataValidation')
    }
  }

  collectActiveMetadataTag(
    activeElement: StreamedMetadataElement,
    localName: string,
    buffer: Uint8Array,
    tagStart: number,
    nameEnd: number,
    tagEnd: number,
  ): void {
    if (activeElement === 'cols' && localName === 'col') {
      this.collectColumnMetadataTag(buffer, tagStart, tagEnd)
      return
    }
    if (activeElement === 'mergeCells' && localName === 'mergeCell') {
      this.collectMergeCellTag(buffer, nameEnd, tagEnd)
      return
    }
    if (activeElement === 'tableParts' && localName === 'tablePart') {
      const relationshipId =
        readXmlAttributeFromTag(buffer, nameEnd, tagEnd, 'r:id') ?? readXmlAttributeFromTag(buffer, nameEnd, tagEnd, 'id')
      if (relationshipId) {
        this.tableCount += 1
        if (this.retainMetadataXml) {
          this.tableRelationshipIds ??= []
          this.tableRelationshipIds.push(relationshipId)
        }
      }
    }
  }

  collectActiveHyperlinkTag(buffer: Uint8Array, startIndex: number, tagEnd: number): boolean {
    if (!this.retainMetadataXml) {
      return true
    }
    const refs = readLargeSimpleSheetHyperlinkRefsFromBytes(buffer, startIndex, tagEnd + 1)
    if (refs === null) {
      return false
    }
    if (refs.length > 0) {
      this.hyperlinks = [...(this.hyperlinks ?? []), ...refs]
    }
    return true
  }

  countConditionalFormatRootRule(rootTag: Uint8Array): void {
    this.conditionalFormatCount += countLargeSimpleConditionalFormattingSqrefRangesFromRootTag(rootTag)
  }

  collectConditionalFormattingRule(rootTag: Uint8Array, buffer: Uint8Array, startIndex: number, endIndex: number): boolean {
    if (!this.sheetName) {
      return false
    }
    const scanBytes = wrapLargeSimpleConditionalFormatRuleBytes(rootTag, buffer, startIndex, endIndex)
    this.appendConditionalFormattingScan(
      readLargeSimpleConditionalFormattingFromBytes(
        this.sheetName,
        scanBytes,
        0,
        scanBytes.byteLength,
        this.conditionalFormatIdCounter + 1,
      ),
    )
    return true
  }

  collectDataValidationElement(buffer: Uint8Array, startIndex: number, endIndex: number): boolean {
    if (this.retainMetadataXml) {
      if (!this.sheetName) {
        return false
      }
      const validations = readLargeSimpleDataValidationsFromBytes(this.sheetName, buffer, startIndex, endIndex)
      if (validations === null) {
        return false
      }
      this.dataValidationCount += validations.length
      if (validations.length > 0) {
        this.dataValidations ??= []
        this.dataValidations.push(...validations)
      }
      return true
    }
    const count = countLargeSimpleDataValidationsFromBytes(this.sheetName ?? 'Sheet1', buffer, startIndex, endIndex)
    if (count === null) {
      return false
    }
    this.dataValidationCount += count
    return true
  }

  private collectColumnMetadataTag(buffer: Uint8Array, tagStart: number, tagEnd: number): void {
    if (!this.retainMetadataXml) {
      return
    }
    this.columnEntries ??= []
    this.columnMetadata ??= []
    appendLargeSimpleColumnMetadataFromBytes(this.columnEntries, this.columnMetadata, buffer, tagStart, tagEnd + 1)
  }

  private collectMergeCellTag(buffer: Uint8Array, nameEnd: number, tagEnd: number): void {
    const ref = readXmlAttributeFromTag(buffer, nameEnd, tagEnd, 'ref')
    const [startAddress, endAddress] = ref?.split(':') ?? []
    if (!startAddress || !endAddress || startAddress === endAddress) {
      return
    }
    this.mergeCount += 1
    if (this.retainMetadataXml) {
      this.mergeRefs ??= []
      this.mergeRefs.push({ startAddress, endAddress })
    }
  }

  private appendConditionalFormattingScan(scan: ConditionalFormattingScan): void {
    this.conditionalFormatCount += scan.ruleCount
    if (scan.conditionalFormats && scan.conditionalFormats.length > 0) {
      this.conditionalFormats ??= []
      this.conditionalFormats.push(...scan.conditionalFormats)
      this.conditionalFormatIdCounter += scan.conditionalFormats.length
    }
    if (scan.artifactXml) {
      this.conditionalFormattingXml ??= []
      this.conditionalFormattingXml.push(scan.artifactXml)
    }
  }
}
