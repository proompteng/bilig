import type { Effect } from 'effect'
import type {
  CellRangeRef,
  LiteralInput,
  WorkbookAutoFilterSnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookChartSnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookImageSnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookNoteSnapshot,
  WorkbookPivotSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookShapeSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookTableSnapshot,
  WorkbookVolatileContextSnapshot,
} from '@bilig/protocol'
import type {
  WorkbookCalculationSettingsRecord,
  WorkbookChartRecord,
  WorkbookCommentThreadRecord,
  WorkbookConditionalFormatRecord,
  WorkbookDataValidationRecord,
  WorkbookDefinedNameRecord,
  WorkbookFilterRecord,
  WorkbookFreezePaneRecord,
  WorkbookImageRecord,
  WorkbookMacroPayloadRecord,
  WorkbookMergeRangeRecord,
  WorkbookNoteRecord,
  WorkbookPivotRecord,
  WorkbookPropertyRecord,
  WorkbookRangeProtectionRecord,
  WorkbookSheetProtectionRecord,
  WorkbookShapeRecord,
  WorkbookSortKeyRecord,
  WorkbookSortRecord,
  WorkbookSpillRecord,
  WorkbookTableRecord,
  WorkbookVolatileContextRecord,
} from './workbook-metadata-types.js'

export class WorkbookMetadataError extends Error {
  readonly _tag = 'WorkbookMetadataError'
  override readonly cause: unknown

  constructor(args: { message: string; cause: unknown }) {
    super(args.message)
    this.name = 'WorkbookMetadataError'
    this.cause = args.cause
  }
}

export interface WorkbookMetadataService {
  readonly renameSheet: (oldSheetName: string, newSheetName: string) => Effect.Effect<void, WorkbookMetadataError>
  readonly deleteSheetRecords: (sheetName: string) => Effect.Effect<void, WorkbookMetadataError>
  readonly reset: () => Effect.Effect<void, WorkbookMetadataError>
  readonly setWorkbookProperty: (
    key: string,
    value: LiteralInput,
  ) => Effect.Effect<WorkbookPropertyRecord | undefined, WorkbookMetadataError>
  readonly getWorkbookProperty: (key: string) => Effect.Effect<WorkbookPropertyRecord | undefined, WorkbookMetadataError>
  readonly listWorkbookProperties: () => Effect.Effect<WorkbookPropertyRecord[], WorkbookMetadataError>
  readonly setMacroPayload: (record: WorkbookMacroPayloadSnapshot) => Effect.Effect<WorkbookMacroPayloadRecord, WorkbookMetadataError>
  readonly listMacroPayloads: () => Effect.Effect<WorkbookMacroPayloadRecord[], WorkbookMetadataError>
  readonly setCalculationSettings: (
    settings: WorkbookCalculationSettingsSnapshot,
  ) => Effect.Effect<WorkbookCalculationSettingsRecord, WorkbookMetadataError>
  readonly getCalculationSettings: () => Effect.Effect<WorkbookCalculationSettingsRecord, WorkbookMetadataError>
  readonly setVolatileContext: (
    context: WorkbookVolatileContextSnapshot,
  ) => Effect.Effect<WorkbookVolatileContextRecord, WorkbookMetadataError>
  readonly getVolatileContext: () => Effect.Effect<WorkbookVolatileContextRecord, WorkbookMetadataError>
  readonly setDefinedName: (
    name: string,
    value: WorkbookDefinedNameValueSnapshot,
    scopeSheetName?: string,
  ) => Effect.Effect<WorkbookDefinedNameRecord, WorkbookMetadataError>
  readonly getDefinedName: (
    name: string,
    scopeSheetName?: string,
  ) => Effect.Effect<WorkbookDefinedNameRecord | undefined, WorkbookMetadataError>
  readonly deleteDefinedName: (name: string, scopeSheetName?: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listDefinedNames: () => Effect.Effect<WorkbookDefinedNameRecord[], WorkbookMetadataError>
  readonly setTable: (record: WorkbookTableSnapshot) => Effect.Effect<WorkbookTableRecord, WorkbookMetadataError>
  readonly getTable: (name: string) => Effect.Effect<WorkbookTableRecord | undefined, WorkbookMetadataError>
  readonly deleteTable: (name: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listTables: () => Effect.Effect<WorkbookTableRecord[], WorkbookMetadataError>
  readonly setFreezePane: (
    sheetName: string,
    rows: number,
    cols: number,
    options?: Pick<WorkbookFreezePaneSnapshot, 'topLeftCell' | 'activePane'>,
  ) => Effect.Effect<WorkbookFreezePaneRecord, WorkbookMetadataError>
  readonly getFreezePane: (sheetName: string) => Effect.Effect<WorkbookFreezePaneRecord | undefined, WorkbookMetadataError>
  readonly clearFreezePane: (sheetName: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly setMergeRange: (range: CellRangeRef) => Effect.Effect<WorkbookMergeRangeRecord, WorkbookMetadataError>
  readonly setMergeRanges: (
    sheetName: string,
    ranges: readonly CellRangeRef[],
  ) => Effect.Effect<WorkbookMergeRangeRecord[], WorkbookMetadataError>
  readonly getMergeRange: (sheetName: string, address: string) => Effect.Effect<WorkbookMergeRangeRecord | undefined, WorkbookMetadataError>
  readonly getMergeRangeByRange: (range: CellRangeRef) => Effect.Effect<WorkbookMergeRangeRecord | undefined, WorkbookMetadataError>
  readonly clearMergeRanges: (range: CellRangeRef) => Effect.Effect<WorkbookMergeRangeRecord[], WorkbookMetadataError>
  readonly listMergeRanges: (sheetName: string) => Effect.Effect<WorkbookMergeRangeRecord[], WorkbookMetadataError>
  readonly setSheetProtection: (
    record: WorkbookSheetProtectionSnapshot,
  ) => Effect.Effect<WorkbookSheetProtectionRecord, WorkbookMetadataError>
  readonly getSheetProtection: (sheetName: string) => Effect.Effect<WorkbookSheetProtectionRecord | undefined, WorkbookMetadataError>
  readonly clearSheetProtection: (sheetName: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly setFilter: (sheetName: string, range: WorkbookAutoFilterSnapshot) => Effect.Effect<WorkbookFilterRecord, WorkbookMetadataError>
  readonly getFilter: (sheetName: string, range: CellRangeRef) => Effect.Effect<WorkbookFilterRecord | undefined, WorkbookMetadataError>
  readonly deleteFilter: (sheetName: string, range: CellRangeRef) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listFilters: (sheetName: string) => Effect.Effect<WorkbookFilterRecord[], WorkbookMetadataError>
  readonly setSort: (
    sheetName: string,
    range: CellRangeRef,
    keys: readonly WorkbookSortKeyRecord[],
  ) => Effect.Effect<WorkbookSortRecord, WorkbookMetadataError>
  readonly getSort: (sheetName: string, range: CellRangeRef) => Effect.Effect<WorkbookSortRecord | undefined, WorkbookMetadataError>
  readonly deleteSort: (sheetName: string, range: CellRangeRef) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listSorts: (sheetName: string) => Effect.Effect<WorkbookSortRecord[], WorkbookMetadataError>
  readonly setDataValidation: (record: WorkbookDataValidationSnapshot) => Effect.Effect<WorkbookDataValidationRecord, WorkbookMetadataError>
  readonly getDataValidation: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<WorkbookDataValidationRecord | undefined, WorkbookMetadataError>
  readonly deleteDataValidation: (sheetName: string, range: CellRangeRef) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listDataValidations: (sheetName: string) => Effect.Effect<WorkbookDataValidationRecord[], WorkbookMetadataError>
  readonly setConditionalFormat: (
    record: WorkbookConditionalFormatSnapshot,
  ) => Effect.Effect<WorkbookConditionalFormatRecord, WorkbookMetadataError>
  readonly getConditionalFormat: (id: string) => Effect.Effect<WorkbookConditionalFormatRecord | undefined, WorkbookMetadataError>
  readonly deleteConditionalFormat: (id: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listConditionalFormats: (sheetName: string) => Effect.Effect<WorkbookConditionalFormatRecord[], WorkbookMetadataError>
  readonly setRangeProtection: (
    record: WorkbookRangeProtectionSnapshot,
  ) => Effect.Effect<WorkbookRangeProtectionRecord, WorkbookMetadataError>
  readonly getRangeProtection: (id: string) => Effect.Effect<WorkbookRangeProtectionRecord | undefined, WorkbookMetadataError>
  readonly deleteRangeProtection: (id: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listRangeProtections: (sheetName: string) => Effect.Effect<WorkbookRangeProtectionRecord[], WorkbookMetadataError>
  readonly setCommentThread: (record: WorkbookCommentThreadSnapshot) => Effect.Effect<WorkbookCommentThreadRecord, WorkbookMetadataError>
  readonly getCommentThread: (
    sheetName: string,
    address: string,
  ) => Effect.Effect<WorkbookCommentThreadRecord | undefined, WorkbookMetadataError>
  readonly deleteCommentThread: (sheetName: string, address: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listCommentThreads: (sheetName: string) => Effect.Effect<WorkbookCommentThreadRecord[], WorkbookMetadataError>
  readonly setNote: (record: WorkbookNoteSnapshot) => Effect.Effect<WorkbookNoteRecord, WorkbookMetadataError>
  readonly getNote: (sheetName: string, address: string) => Effect.Effect<WorkbookNoteRecord | undefined, WorkbookMetadataError>
  readonly deleteNote: (sheetName: string, address: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listNotes: (sheetName: string) => Effect.Effect<WorkbookNoteRecord[], WorkbookMetadataError>
  readonly setSpill: (
    sheetName: string,
    address: string,
    rows: number,
    cols: number,
  ) => Effect.Effect<WorkbookSpillRecord, WorkbookMetadataError>
  readonly getSpill: (sheetName: string, address: string) => Effect.Effect<WorkbookSpillRecord | undefined, WorkbookMetadataError>
  readonly deleteSpill: (sheetName: string, address: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listSpills: () => Effect.Effect<WorkbookSpillRecord[], WorkbookMetadataError>
  readonly setPivot: (record: WorkbookPivotSnapshot) => Effect.Effect<WorkbookPivotRecord, WorkbookMetadataError>
  readonly getPivot: (sheetName: string, address: string) => Effect.Effect<WorkbookPivotRecord | undefined, WorkbookMetadataError>
  readonly getPivotByKey: (key: string) => Effect.Effect<WorkbookPivotRecord | undefined, WorkbookMetadataError>
  readonly deletePivot: (sheetName: string, address: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly hasPivots: () => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listPivots: () => Effect.Effect<WorkbookPivotRecord[], WorkbookMetadataError>
  readonly setChart: (record: WorkbookChartSnapshot) => Effect.Effect<WorkbookChartRecord, WorkbookMetadataError>
  readonly getChart: (id: string) => Effect.Effect<WorkbookChartRecord | undefined, WorkbookMetadataError>
  readonly deleteChart: (id: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listCharts: () => Effect.Effect<WorkbookChartRecord[], WorkbookMetadataError>
  readonly setImage: (record: WorkbookImageSnapshot) => Effect.Effect<WorkbookImageRecord, WorkbookMetadataError>
  readonly getImage: (id: string) => Effect.Effect<WorkbookImageRecord | undefined, WorkbookMetadataError>
  readonly deleteImage: (id: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listImages: () => Effect.Effect<WorkbookImageRecord[], WorkbookMetadataError>
  readonly setShape: (record: WorkbookShapeSnapshot) => Effect.Effect<WorkbookShapeRecord, WorkbookMetadataError>
  readonly getShape: (id: string) => Effect.Effect<WorkbookShapeRecord | undefined, WorkbookMetadataError>
  readonly deleteShape: (id: string) => Effect.Effect<boolean, WorkbookMetadataError>
  readonly listShapes: () => Effect.Effect<WorkbookShapeRecord[], WorkbookMetadataError>
}
