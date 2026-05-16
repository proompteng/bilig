import { defineQueriesWithType, defineQuery, defineQueryWithType } from '@rocicorp/zero'
import { z } from 'zod'
import type { schema } from './schema.js'
import { safeNonNegativeIntegerSchema, safePositiveIntegerSchema } from './integer-schemas.js'
import { zql } from './zql.js'

const defineQueries = defineQueriesWithType<typeof schema>()
const defineUserQuery = defineQueryWithType<typeof schema, { readonly userID: string }>()

export const workbookQueryArgsSchema = z.object({
  documentId: z.string().min(1),
})

export const workbookThreadArgsSchema = workbookQueryArgsSchema.extend({
  threadId: z.string().min(1),
})

const workbookViewerArgsSchema = workbookQueryArgsSchema.extend({
  currentUserId: z.string().min(1),
})

const workbookThreadViewerArgsSchema = workbookThreadArgsSchema.extend({
  currentUserId: z.string().min(1),
})

const workbookSheetArgsSchema = workbookQueryArgsSchema
  .extend({
    sheetId: safePositiveIntegerSchema.optional(),
    sheetName: z.string().min(1).optional(),
  })
  .refine((args) => (args.sheetId !== undefined) !== (args.sheetName !== undefined), {
    message: 'exactly one of sheetId or sheetName is required',
  })

function scopeCellInputQueryToSheet(args: z.infer<typeof workbookSheetArgsSchema>) {
  const query = zql.cells.where('workbookId', args.documentId)
  return args.sheetName !== undefined
    ? query.where('sheetName', args.sheetName)
    : query.where((expression) => expression.exists('sheet', (sheetQuery) => sheetQuery.where('sheetId', args.sheetId ?? 0)))
}

function scopeCellEvalQueryToSheet(args: z.infer<typeof workbookSheetArgsSchema>) {
  const query = zql.cell_eval.where('workbookId', args.documentId)
  return args.sheetName !== undefined
    ? query.where('sheetName', args.sheetName)
    : query.where((expression) => expression.exists('sheet', (sheetQuery) => sheetQuery.where('sheetId', args.sheetId ?? 0)))
}

function scopeRowMetadataQueryToSheet(args: z.infer<typeof workbookSheetArgsSchema>) {
  const query = zql.row_metadata.where('workbookId', args.documentId)
  return args.sheetName !== undefined
    ? query.where('sheetName', args.sheetName)
    : query.where((expression) => expression.exists('sheet', (sheetQuery) => sheetQuery.where('sheetId', args.sheetId ?? 0)))
}

function scopeColumnMetadataQueryToSheet(args: z.infer<typeof workbookSheetArgsSchema>) {
  const query = zql.column_metadata.where('workbookId', args.documentId)
  return args.sheetName !== undefined
    ? query.where('sheetName', args.sheetName)
    : query.where((expression) => expression.exists('sheet', (sheetQuery) => sheetQuery.where('sheetId', args.sheetId ?? 0)))
}

export const workbookCellArgsSchema = workbookSheetArgsSchema.extend({
  address: z.string().min(1),
})

export const workbookTileArgsSchema = workbookSheetArgsSchema
  .extend({
    rowStart: safeNonNegativeIntegerSchema,
    rowEnd: safeNonNegativeIntegerSchema,
    colStart: safeNonNegativeIntegerSchema,
    colEnd: safeNonNegativeIntegerSchema,
  })
  .refine((args) => args.rowEnd >= args.rowStart && args.colEnd >= args.colStart, {
    message: 'tile end must be greater than or equal to tile start',
  })

export const workbookRowTileArgsSchema = workbookSheetArgsSchema
  .extend({
    rowStart: safeNonNegativeIntegerSchema,
    rowEnd: safeNonNegativeIntegerSchema,
  })
  .refine((args) => args.rowEnd >= args.rowStart, {
    message: 'row tile end must be greater than or equal to row tile start',
  })

export const workbookColumnTileArgsSchema = workbookSheetArgsSchema
  .extend({
    colStart: safeNonNegativeIntegerSchema,
    colEnd: safeNonNegativeIntegerSchema,
  })
  .refine((args) => args.colEnd >= args.colStart, {
    message: 'column tile end must be greater than or equal to column tile start',
  })

const workbookGet = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) => zql.workbooks.where('id', documentId).one())

const sheetByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.sheets.where('workbookId', documentId).orderBy('sortOrder', 'asc'),
)

const cellInputOne = defineQuery(workbookCellArgsSchema, ({ args }) =>
  scopeCellInputQueryToSheet(args).where('address', args.address).one(),
)

const cellInputTile = defineQuery(workbookTileArgsSchema, ({ args }) =>
  scopeCellInputQueryToSheet(args)
    .where('rowNum', '>=', args.rowStart)
    .where('rowNum', '<=', args.rowEnd)
    .where('colNum', '>=', args.colStart)
    .where('colNum', '<=', args.colEnd)
    .orderBy('rowNum', 'asc')
    .orderBy('colNum', 'asc'),
)

const cellEvalOne = defineQuery(workbookCellArgsSchema, ({ args }) => scopeCellEvalQueryToSheet(args).where('address', args.address).one())

const cellEvalTile = defineQuery(workbookTileArgsSchema, ({ args }) =>
  scopeCellEvalQueryToSheet(args)
    .where('rowNum', '>=', args.rowStart)
    .where('rowNum', '<=', args.rowEnd)
    .where('colNum', '>=', args.colStart)
    .where('colNum', '<=', args.colEnd)
    .orderBy('rowNum', 'asc')
    .orderBy('colNum', 'asc'),
)

const sheetRowTile = defineQuery(workbookRowTileArgsSchema, ({ args }) =>
  scopeRowMetadataQueryToSheet(args)
    .where('startIndex', '>=', args.rowStart)
    .where('startIndex', '<=', args.rowEnd)
    .orderBy('startIndex', 'asc'),
)

const sheetColTile = defineQuery(workbookColumnTileArgsSchema, ({ args }) =>
  scopeColumnMetadataQueryToSheet(args)
    .where('startIndex', '>=', args.colStart)
    .where('startIndex', '<=', args.colEnd)
    .orderBy('startIndex', 'asc'),
)

const cellStyleByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.cell_styles.where('workbookId', documentId).orderBy('styleId', 'asc'),
)

const numberFormatByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.cell_number_formats.where('workbookId', documentId).orderBy('formatId', 'asc'),
)

const presenceCoarseByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.presence_coarse.where('workbookId', documentId).orderBy('updatedAt', 'desc'),
)

const workbookChangeByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.workbook_change.where('workbookId', documentId).orderBy('createdAt', 'desc').orderBy('revision', 'desc'),
)

const workbookChatThreadByWorkbook = defineUserQuery(workbookQueryArgsSchema, ({ args: { documentId }, ctx }) =>
  zql.workbook_chat_thread
    .where('workbookId', documentId)
    .where((expression) => expression.or(expression.cmp('ownerUserId', ctx.userID), expression.cmp('scope', 'shared')))
    .orderBy('updatedAtUnixMs', 'desc'),
)

const visibleWorkbookChatThreadByWorkbook = defineQuery(workbookViewerArgsSchema, ({ args: { currentUserId, documentId } }) =>
  zql.workbook_chat_thread
    .where('workbookId', documentId)
    .where((expression) => expression.or(expression.cmp('ownerUserId', currentUserId), expression.cmp('scope', 'shared')))
    .orderBy('updatedAtUnixMs', 'desc'),
)

const workbookWorkflowRunByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_workflow_run
    .where('workbookId', args.documentId)
    .where('threadId', args.threadId)
    .where((expression) =>
      expression.or(
        expression.cmp('startedByUserId', ctx.userID),
        expression.exists('chatThreads', (threadQuery) => threadQuery.where('scope', 'shared')),
      ),
    )
    .orderBy('updatedAtUnixMs', 'desc'),
)

const visibleWorkbookWorkflowRunByThread = defineQuery(workbookThreadViewerArgsSchema, ({ args }) =>
  zql.workbook_workflow_run
    .where('workbookId', args.documentId)
    .where('threadId', args.threadId)
    .where((expression) =>
      expression.or(
        expression.cmp('startedByUserId', args.currentUserId),
        expression.exists('chatThreads', (threadQuery) => threadQuery.where('scope', 'shared')),
      ),
    )
    .orderBy('updatedAtUnixMs', 'desc'),
)

export const queries = defineQueries({
  workbook: {
    get: workbookGet,
  },
  workbooks: {
    get: workbookGet,
  },
  sheet: {
    byWorkbook: sheetByWorkbook,
  },
  sheets: {
    byWorkbook: sheetByWorkbook,
  },
  cellInput: {
    one: cellInputOne,
    tile: cellInputTile,
  },
  cells: {
    one: cellInputOne,
    tile: cellInputTile,
  },
  cellEval: {
    one: cellEvalOne,
    tile: cellEvalTile,
  },
  cellRender: {
    one: cellEvalOne,
    tile: cellEvalTile,
  },
  sheetRow: {
    tile: sheetRowTile,
  },
  rowMetadata: {
    tile: sheetRowTile,
  },
  sheetCol: {
    tile: sheetColTile,
  },
  columnMetadata: {
    tile: sheetColTile,
  },
  cellStyle: {
    byWorkbook: cellStyleByWorkbook,
  },
  numberFormat: {
    byWorkbook: numberFormatByWorkbook,
  },
  presenceCoarse: {
    byWorkbook: presenceCoarseByWorkbook,
  },
  presence: {
    byWorkbook: presenceCoarseByWorkbook,
  },
  workbookChange: {
    byWorkbook: workbookChangeByWorkbook,
  },
  workbookChanges: {
    byWorkbook: workbookChangeByWorkbook,
  },
  workbookChatThread: {
    byWorkbook: workbookChatThreadByWorkbook,
    visibleByWorkbook: visibleWorkbookChatThreadByWorkbook,
  },
  workbookAgentThread: {
    byWorkbook: workbookChatThreadByWorkbook,
    visibleByWorkbook: visibleWorkbookChatThreadByWorkbook,
  },
  workbookWorkflowRun: {
    byThread: workbookWorkflowRunByThread,
    visibleByThread: visibleWorkbookWorkflowRunByThread,
  },
  workbookAgentWorkflowRun: {
    byThread: workbookWorkflowRunByThread,
    visibleByThread: visibleWorkbookWorkflowRunByThread,
  },
})
