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

const limitedWorkbookQueryArgsSchema = workbookQueryArgsSchema.extend({
  limit: safePositiveIntegerSchema.optional(),
})

const workbookChangeRevisionArgsSchema = workbookQueryArgsSchema.extend({
  revision: safePositiveIntegerSchema,
})

const workbookChangeAfterRevisionArgsSchema = workbookQueryArgsSchema.extend({
  revision: safeNonNegativeIntegerSchema,
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

const workbookChangeOne = defineQuery(workbookChangeRevisionArgsSchema, ({ args }) =>
  zql.workbook_change.where('workbookId', args.documentId).where('revision', args.revision).one(),
)

const workbookChangeByWorkbook = defineQuery(limitedWorkbookQueryArgsSchema, ({ args }) => {
  const query = zql.workbook_change.where('workbookId', args.documentId).orderBy('createdAt', 'desc').orderBy('revision', 'desc')
  return args.limit === undefined ? query : query.limit(args.limit)
})

const workbookChangeHistoryByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.workbook_change.where('workbookId', documentId).orderBy('revision', 'asc'),
)

const workbookChangeAfterRevision = defineQuery(workbookChangeAfterRevisionArgsSchema, ({ args }) =>
  zql.workbook_change.where('workbookId', args.documentId).where('revision', '>', args.revision).orderBy('revision', 'asc'),
)

const workbookChatThreadByWorkbook = defineUserQuery(workbookQueryArgsSchema, ({ args: { documentId }, ctx }) =>
  zql.workbook_chat_thread
    .where('workbookId', documentId)
    .where((expression) => expression.or(expression.cmp('ownerUserId', ctx.userID), expression.cmp('scope', 'shared')))
    .orderBy('updatedAtUnixMs', 'desc'),
)

const visibleWorkbookChatThreadByWorkbook = defineUserQuery(workbookQueryArgsSchema, ({ args: { documentId }, ctx }) =>
  zql.workbook_chat_thread
    .where('workbookId', documentId)
    .where((expression) => expression.or(expression.cmp('ownerUserId', ctx.userID), expression.cmp('scope', 'shared')))
    .orderBy('updatedAtUnixMs', 'desc'),
)

const workbookChatItemByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_chat_item
    .where('workbookId', args.documentId)
    .where('threadId', args.threadId)
    .where((expression) =>
      expression.exists('thread', (threadQuery) =>
        threadQuery.where((threadExpression) =>
          threadExpression.or(threadExpression.cmp('ownerUserId', ctx.userID), threadExpression.cmp('scope', 'shared')),
        ),
      ),
    )
    .orderBy('sortOrder', 'asc')
    .orderBy('entryId', 'asc'),
)

const workbookChatToolCallByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_chat_tool_call
    .where('workbookId', args.documentId)
    .where('threadId', args.threadId)
    .where((expression) =>
      expression.exists('thread', (threadQuery) =>
        threadQuery.where((threadExpression) =>
          threadExpression.or(threadExpression.cmp('ownerUserId', ctx.userID), threadExpression.cmp('scope', 'shared')),
        ),
      ),
    )
    .orderBy('sortOrder', 'asc')
    .orderBy('entryId', 'asc'),
)

const workbookReviewQueueItemByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_review_queue_item
    .where('workbookId', args.documentId)
    .where('threadId', args.threadId)
    .where((expression) =>
      expression.exists('thread', (threadQuery) =>
        threadQuery.where((threadExpression) =>
          threadExpression.or(threadExpression.cmp('ownerUserId', ctx.userID), threadExpression.cmp('scope', 'shared')),
        ),
      ),
    )
    .orderBy('createdAtUnixMs', 'asc')
    .orderBy('reviewItemId', 'asc'),
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

const visibleWorkbookWorkflowRunByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
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

const workbookWorkflowStepByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_workflow_step
    .where('workbookId', args.documentId)
    .where((expression) =>
      expression.exists('workflowRun', (runQuery) =>
        runQuery.where('threadId', args.threadId).where((runExpression) =>
          runExpression.or(
            runExpression.cmp('startedByUserId', ctx.userID),
            runExpression.exists('chatThreads', (threadQuery) => threadQuery.where('scope', 'shared')),
          ),
        ),
      ),
    )
    .orderBy('runId', 'asc')
    .orderBy('stepOrder', 'asc'),
)

const workbookWorkflowArtifactByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_workflow_artifact
    .where('workbookId', args.documentId)
    .where((expression) =>
      expression.exists('workflowRun', (runQuery) =>
        runQuery.where('threadId', args.threadId).where((runExpression) =>
          runExpression.or(
            runExpression.cmp('startedByUserId', ctx.userID),
            runExpression.exists('chatThreads', (threadQuery) => threadQuery.where('scope', 'shared')),
          ),
        ),
      ),
    )
    .orderBy('runId', 'asc'),
)

const workbookAgentRunByWorkbook = defineUserQuery(workbookQueryArgsSchema, ({ args: { documentId }, ctx }) =>
  zql.workbook_agent_run.where('workbookId', documentId).where('actorUserId', ctx.userID).orderBy('appliedAtUnixMs', 'desc'),
)

const workbookAgentRunByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_agent_run
    .where('workbookId', args.documentId)
    .where('threadId', args.threadId)
    .where((expression) =>
      expression.or(
        expression.cmp('actorUserId', ctx.userID),
        expression.exists('chatThreads', (threadQuery) => threadQuery.where('scope', 'shared')),
      ),
    )
    .orderBy('appliedAtUnixMs', 'desc'),
)

const visibleWorkbookAgentRunByThread = defineUserQuery(workbookThreadArgsSchema, ({ args, ctx }) =>
  zql.workbook_agent_run
    .where('workbookId', args.documentId)
    .where('threadId', args.threadId)
    .where((expression) =>
      expression.or(
        expression.cmp('actorUserId', ctx.userID),
        expression.exists('chatThreads', (threadQuery) => threadQuery.where('scope', 'shared')),
      ),
    )
    .orderBy('appliedAtUnixMs', 'desc'),
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
    one: workbookChangeOne,
    byWorkbook: workbookChangeByWorkbook,
    historyByWorkbook: workbookChangeHistoryByWorkbook,
    afterRevision: workbookChangeAfterRevision,
  },
  workbookChanges: {
    one: workbookChangeOne,
    byWorkbook: workbookChangeByWorkbook,
    historyByWorkbook: workbookChangeHistoryByWorkbook,
    afterRevision: workbookChangeAfterRevision,
  },
  workbookChatThread: {
    byWorkbook: workbookChatThreadByWorkbook,
    visibleByWorkbook: visibleWorkbookChatThreadByWorkbook,
  },
  workbookChatItem: {
    byThread: workbookChatItemByThread,
  },
  workbookChatToolCall: {
    byThread: workbookChatToolCallByThread,
  },
  workbookReviewQueueItem: {
    byThread: workbookReviewQueueItemByThread,
  },
  workbookAgentThread: {
    byWorkbook: workbookChatThreadByWorkbook,
    visibleByWorkbook: visibleWorkbookChatThreadByWorkbook,
  },
  workbookAgentRun: {
    byWorkbook: workbookAgentRunByWorkbook,
    byThread: workbookAgentRunByThread,
    visibleByThread: visibleWorkbookAgentRunByThread,
  },
  workbookWorkflowRun: {
    byThread: workbookWorkflowRunByThread,
    visibleByThread: visibleWorkbookWorkflowRunByThread,
  },
  workbookWorkflowStep: {
    byThread: workbookWorkflowStepByThread,
  },
  workbookWorkflowArtifact: {
    byThread: workbookWorkflowArtifactByThread,
  },
  workbookAgentWorkflowRun: {
    byThread: workbookWorkflowRunByThread,
    visibleByThread: visibleWorkbookWorkflowRunByThread,
  },
})
