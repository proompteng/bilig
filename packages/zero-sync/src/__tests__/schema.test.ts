import { describe, expect, it } from 'vitest'
import { schema, zeroSchemaColumnNamesByTable, zeroSchemaServerColumnNamesByTable, zeroSchemaTableNames } from '../schema'

describe('zero sync schema', () => {
  it('exports replicated table metadata from the shared schema model', () => {
    expect(zeroSchemaTableNames).toEqual(Object.keys(schema.tables))
    expect(zeroSchemaColumnNamesByTable.workbooks).toEqual(Object.keys(schema.tables.workbooks.columns))
    expect(zeroSchemaServerColumnNamesByTable.cell_styles).toEqual(['workbook_id', 'style_id', 'record_json', 'hash', 'created_at'])
    expect(zeroSchemaColumnNamesByTable.workbook_agent_run).toEqual(Object.keys(schema.tables.workbook_agent_run.columns))
    expect(zeroSchemaColumnNamesByTable.workbook_chat_item).toEqual(Object.keys(schema.tables.workbook_chat_item.columns))
    expect(zeroSchemaColumnNamesByTable.workbook_chat_tool_call).toEqual(Object.keys(schema.tables.workbook_chat_tool_call.columns))
    expect(zeroSchemaColumnNamesByTable.workbook_review_queue_item).toEqual(Object.keys(schema.tables.workbook_review_queue_item.columns))
    expect(zeroSchemaColumnNamesByTable.workbook_workflow_run).toEqual(Object.keys(schema.tables.workbook_workflow_run.columns))
    expect(zeroSchemaColumnNamesByTable.workbook_workflow_step).toEqual(Object.keys(schema.tables.workbook_workflow_step.columns))
    expect(zeroSchemaColumnNamesByTable.workbook_workflow_artifact).toEqual(Object.keys(schema.tables.workbook_workflow_artifact.columns))
  })

  it('maps workbooks.updated_at as a numeric timestamp', () => {
    expect(schema.tables.workbooks.columns.updatedAt.type).toBe('number')
    expect(schema.tables.workbooks.columns.updatedAt.serverName).toBe('updated_at')
    expect('snapshot' in schema.tables.workbooks.columns).toBe(false)
    expect('replicaSnapshot' in schema.tables.workbooks.columns).toBe(false)
  })

  it('exposes the current workbook projection tables', () => {
    expect(schema.tables.sheets.columns.sheetId.serverName).toBe('sheet_id')
    expect(schema.tables.cells.columns.rowNum.serverName).toBe('row_num')
    expect(schema.tables.cells.columns.styleId.serverName).toBe('style_id')
    expect(schema.tables.cell_eval.columns.calcRevision.serverName).toBe('calc_revision')
    expect(schema.tables.cell_eval.columns.styleId.serverName).toBe('style_id')
    expect(schema.tables.cell_eval.columns.formatId.serverName).toBe('format_id')
    expect('cell_styles' in schema.tables).toBe(true)
    expect('cell_number_formats' in schema.tables).toBe(true)
    expect('row_metadata' in schema.tables).toBe(true)
    expect('column_metadata' in schema.tables).toBe(true)
    expect('defined_names' in schema.tables).toBe(true)
    expect('workbook_version' in schema.tables).toBe(false)
    expect('workbook_metadata' in schema.tables).toBe(false)
    expect('calculation_settings' in schema.tables).toBe(false)
  })

  it('maps durable workbook chat thread summaries to the UI contract', () => {
    expect(schema.tables.workbook_chat_thread.columns.ownerUserId.serverName).toBe('actor_user_id')
    expect(schema.tables.workbook_chat_thread.columns.executionPolicy.serverName).toBe('execution_policy')
    expect(schema.tables.workbook_chat_thread.columns.context.serverName).toBe('context_json')
    expect(schema.tables.workbook_chat_thread.columns.reviewQueueItemCount.serverName).toBe('review_queue_item_count')
    expect(schema.tables.workbook_chat_thread.columns.latestEntryText.serverName).toBe('latest_entry_text')
  })

  it('relates durable chat child rows to their parent thread visibility model', () => {
    for (const tableName of ['workbook_chat_item', 'workbook_chat_tool_call', 'workbook_review_queue_item'] as const) {
      expect(schema.relationships[tableName].thread).toEqual([
        {
          sourceField: ['workbookId', 'threadId', 'actorUserId'],
          destField: ['workbookId', 'threadId', 'ownerUserId'],
          destSchema: 'workbook_chat_thread',
          cardinality: 'one',
        },
      ])
    }
  })

  it('relates applied agent runs to chat thread visibility rows', () => {
    expect(schema.relationships.workbook_agent_run.ownerChatThreads).toEqual([
      {
        sourceField: ['workbookId', 'threadId', 'actorUserId'],
        destField: ['workbookId', 'threadId', 'ownerUserId'],
        destSchema: 'workbook_chat_thread',
        cardinality: 'many',
      },
    ])
  })

  it('relates workflow runs to chat thread visibility rows', () => {
    expect(schema.relationships.workbook_workflow_run.chatThreads).toEqual([
      {
        sourceField: ['workbookId', 'threadId'],
        destField: ['workbookId', 'threadId'],
        destSchema: 'workbook_chat_thread',
        cardinality: 'many',
      },
    ])
    expect(schema.relationships.workbook_workflow_run.ownerChatThreads).toEqual([
      {
        sourceField: ['workbookId', 'threadId', 'startedByUserId'],
        destField: ['workbookId', 'threadId', 'ownerUserId'],
        destSchema: 'workbook_chat_thread',
        cardinality: 'many',
      },
    ])
  })

  it('relates workflow child rows to their parent run visibility model', () => {
    expect(schema.relationships.workbook_workflow_step.workflowRun).toEqual([
      {
        sourceField: ['runId'],
        destField: ['runId'],
        destSchema: 'workbook_workflow_run',
        cardinality: 'one',
      },
    ])
    expect(schema.relationships.workbook_workflow_artifact.workflowRun).toEqual([
      {
        sourceField: ['runId'],
        destField: ['runId'],
        destSchema: 'workbook_workflow_run',
        cardinality: 'one',
      },
    ])
  })

  it('relates sheet-owned projection rows through the shared sheet id model', () => {
    expect(schema.relationships.cells.sheet).toEqual([
      {
        sourceField: ['workbookId', 'sheetName'],
        destField: ['workbookId', 'name'],
        destSchema: 'sheets',
        cardinality: 'one',
      },
    ])
    expect(schema.relationships.cell_eval.sheet).toEqual([
      {
        sourceField: ['workbookId', 'sheetName'],
        destField: ['workbookId', 'name'],
        destSchema: 'sheets',
        cardinality: 'one',
      },
    ])
    expect(schema.relationships.row_metadata.sheet).toEqual([
      {
        sourceField: ['workbookId', 'sheetName'],
        destField: ['workbookId', 'name'],
        destSchema: 'sheets',
        cardinality: 'one',
      },
    ])
    expect(schema.relationships.column_metadata.sheet).toEqual([
      {
        sourceField: ['workbookId', 'sheetName'],
        destField: ['workbookId', 'name'],
        destSchema: 'sheets',
        cardinality: 'one',
      },
    ])
  })
})
