# Workbook Agent Platform Design

## Status

Draft for implementation.

This document defines the production target for the workbook agent platform in `bilig`.

The current system already has:

- workbook reads and writes through direct tools
- workflow execution
- execution policy concepts
- execution records

The current system still carries preview-first runtime structure in the main path. This design removes that structure completely and replaces it with an execution-first architecture.

## Goal

Build a workbook agent platform that can operate a serious spreadsheet product safely and directly.

Target product behavior:

- private-thread work executes directly under session policy
- shared-thread work uses owner review when policy requires it
- the agent reads authoritative workbook state, not UI-derived approximations
- the UI centers on applied work, undo, and review queue state
- the runtime has one write engine, one review model, and one durable execution model

Target engineering outcome:

- no preview-first runtime path
- no legacy behavior authority
- no duplicate write engines
- no user-facing implementation slang
- no long-term compatibility branches in the live path

## Problem Statement

The current workbook agent implementation is structurally split.

Primary issues:

- `executionPolicy` exists, but private-thread behavior still passes through preview-oriented bundle staging
- `pendingBundle` is still the main mutation carrier for too many flows
- prompt and UI language still reflect internal implementation concepts instead of a clean product model
- the agent tool surface is strong for reads and basic writes, but incomplete for high-value spreadsheet operations
- workbook state access is still not rich enough for serious autonomous editing across tables, validation, analytics objects, and collaboration objects

That architecture produces the wrong product shape:

- users see review cards for normal private-thread edits
- the model describes review mechanics instead of workbook outcomes
- tool capability gaps prevent complete end-to-end spreadsheet work

## Design Principles

### 1. One behavioral authority

`executionPolicy` is the only runtime authority for mutation routing.

Runtime decisions do not branch on:

- `approvalMode`
- ad hoc prompt wording
- UI state
- thread-specific exceptions

### 2. One workbook-state authority

Workbook state comes from runtime and grid semantics only.

The agent never depends on:

- name-box strings
- display-only selection labels
- inferred style state
- duplicated local caches as the source of truth

### 3. One write engine

All mutating operations flow through one change-set executor:

- direct tools
- durable workflows
- replay
- review apply
- undo and redo

### 4. One durable execution model

Executed work is represented by `executionRecords`.

Review-required work is represented by `reviewQueueItems`.

There is no third mutation model in the runtime.

### 5. Product language only

User-visible language must speak in workbook terms:

- applied changes
- owner review
- undo available
- workbook panel
- selection range

The product never tells users to use a rail or apply a preview bundle.

### 6. Delete as you cut over

Legacy code is removed as each replacement lands.

This design does not allow a permanent dual-runtime model.

## Required Product Behavior

### Private threads

Private-thread sessions execute workbook edits directly.

Examples:

- "remove this fill"
- "normalize these headers"
- "fill formulas down"
- "hide this row"
- "create a chart from this table"

Expected result:

- one turn
- one execution record
- one authoritative diff
- undo available

### Shared threads

Shared-thread sessions use `ownerReview` by default.

Expected result:

- low-risk work follows session policy
- medium/high-risk work becomes a review queue item
- owner review applies or returns the change set

### Agent comprehension

The agent can directly inspect:

- selection geometry
- visible range
- used range
- tables
- charts
- pivot tables
- validations
- conditional formats
- comments and notes
- protection state
- hidden and grouped axes
- freeze panes
- merged ranges

## Domain Model

The following contracts become first-class and authoritative.

### Session and routing

- `WorkbookAgentExecutionPolicy`
- `WorkbookAgentSessionSnapshot`
- `WorkbookReviewQueueItem`
- `WorkbookExecutionRecord`

### Workbook state

- `WorkbookSelectionSnapshot`
- `WorkbookViewportSnapshot`
- `WorkbookCellSnapshot`
- `WorkbookRangeSnapshot`
- `WorkbookSheetSnapshot`
- `WorkbookWorkbookSnapshot`

### Semantic workbook objects

- `WorkbookNamedRangeSnapshot`
- `WorkbookTableSnapshot`
- `WorkbookChartSnapshot`
- `WorkbookPivotSnapshot`
- `WorkbookConditionalFormatSnapshot`
- `WorkbookValidationRuleSnapshot`
- `WorkbookCommentThreadSnapshot`
- `WorkbookProtectionSnapshot`

### Change execution

- `WorkbookSemanticSelector`
- `WorkbookChangeOperation`
- `WorkbookChangeSet`
- `WorkbookChangeDiff`
- `WorkbookUndoPayload`

## Semantic Selector Model

The agent should not operate only on A1 coordinates.

`WorkbookSemanticSelector` must support:

- `a1Range`
- `namedRange`
- `table`
- `tableColumn`
- `currentSelection`
- `currentRegion`
- `visibleRows`
- `rowQuery`
- `columnQuery`

Selectors resolve through one service:

- `WorkbookSelectorResolver`

The resolver returns authoritative workbook objects with explicit revision identity. Every read and write path depends on this resolver.

## Tool Surface

The workbook toolset should be layered, typed, and workbook-native.

### Layer 1: context and navigation

- `get_context`
- `read_selection`
- `read_visible_range`
- `read_workbook`
- `get_sheet_view`
- `get_used_range`
- `get_current_region`
- `list_sheets`
- `select_sheet`
- `select_range`

### Layer 2: rich reads and inspection

- `read_range`
- `inspect_cell`
- `search_workbook`
- `trace_dependencies`
- `list_named_ranges`
- `list_tables`
- `list_charts`
- `list_pivots`
- `get_data_validation`
- `get_conditional_formats`
- `get_comments`
- `get_row_metadata`
- `get_column_metadata`
- `get_merged_ranges`

Read payloads must include the full semantic snapshot where relevant:

- raw input
- displayed value
- formula
- cell type
- number format
- style metadata
- hyperlink metadata
- merge metadata
- validation metadata
- conditional formatting matches
- hidden and filtered state

### Layer 3: write primitives

- `write_range`
- `clear_range`
- `copy_range`
- `move_range`
- `fill_range`
- `insert_rows`
- `delete_rows`
- `insert_columns`
- `delete_columns`
- `insert_cells`
- `delete_cells`
- `merge_range`
- `unmerge_range`
- `paste_special`
- `set_hyperlinks`
- `set_checkboxes`

### Layer 4: formatting and view

- `format_range`
- `set_number_format`
- `set_alignment`
- `set_borders`
- `set_fill`
- `set_font`
- `set_wrap`
- `set_indentation`
- `update_row_metadata`
- `update_column_metadata`
- `freeze_panes`
- `set_sheet_view_options`
- `add_conditional_format`
- `update_conditional_format`
- `remove_conditional_format`

### Layer 5: workbook structure

- `create_sheet`
- `rename_sheet`
- `delete_sheet`
- `duplicate_sheet`
- `move_sheet`
- `hide_sheet`
- `unhide_sheet`
- `create_named_range`
- `update_named_range`
- `delete_named_range`
- `create_table`
- `resize_table`
- `delete_table`
- `group_rows`
- `ungroup_rows`
- `group_columns`
- `ungroup_columns`

### Layer 6: transformation and cleanup

- `sort_range`
- `filter_range`
- `create_filter_view`
- `clear_filter`
- `remove_duplicates`
- `find_replace`
- `text_to_columns`
- `split_column`
- `merge_columns`
- `normalize_whitespace`
- `normalize_case`
- `parse_numbers_dates`
- `infer_headers`
- `select_rows_by_key`

### Layer 7: formula and calculation

- `set_formula`
- `fill_formulas_down`
- `fill_formulas_across`
- `recalc_sheet`
- `recalc_workbook`
- `evaluate_formula`
- `explain_formula`
- `get_formula_ast`
- `find_formula_issues`
- `repair_formula_issues`
- `convert_formulas_to_values`
- `detect_circular_references`
- `find_volatile_hotspots`

### Layer 8: analytics and collaboration

- `create_chart`
- `update_chart`
- `delete_chart`
- `create_pivot_table`
- `update_pivot_table`
- `delete_pivot_table`
- `create_sparkline`
- `update_sparkline`
- `add_slicer`
- `remove_slicer`
- `add_comment`
- `reply_comment`
- `resolve_comment`
- `delete_comment`
- `add_note`
- `update_note`
- `delete_note`

### Layer 9: validation, protection, media, exchange

- `create_data_validation`
- `update_data_validation`
- `remove_data_validation`
- `list_data_validation_rules`
- `protect_sheet`
- `unprotect_sheet`
- `protect_range`
- `unprotect_range`
- `lock_cells`
- `unlock_cells`
- `hide_formulas`
- `get_protection_status`
- `insert_image`
- `move_image`
- `delete_image`
- `insert_shape`
- `update_shape`
- `delete_shape`
- `insert_text_box`
- `import_csv`
- `export_selection`
- `export_sheet`
- `export_workbook`
- `refresh_external_data`
- `get_connection_status`

### Layer 10: audit and AI operations

- `begin_change_set`
- `preview_change_set`
- `commit_change_set`
- `discard_change_set`
- `undo_change_set`
- `redo_change_set`
- `get_change_diff`
- `read_recent_changes`
- `diff_revisions`
- `scan_broken_references`
- `scan_hidden_rows_affecting_results`
- `scan_inconsistent_formulas`
- `scan_used_range_bloat`
- `scan_performance_hotspots`
- `verify_invariants`

## Backend Architecture

The backend should be split into explicit services with narrow ownership.

### 1. WorkbookInspectionService

Responsibilities:

- build authoritative workbook object snapshots
- expose cell, range, table, chart, pivot, validation, and protection reads
- serve tool reads and workflow reads from one object model

Primary files to introduce:

- `apps/bilig/src/codex-app/workbook-agent-inspection-service.ts`
- `apps/bilig/src/codex-app/workbook-agent-object-snapshots.ts`

### 2. WorkbookSelectorResolver

Responsibilities:

- resolve semantic selectors against workbook state
- validate selector scope and revision identity
- return stable object references for change execution

Primary files to introduce:

- `apps/bilig/src/codex-app/workbook-selector-resolver.ts`

### 3. WorkbookChangeSetBuilder

Responsibilities:

- convert tool requests and workflow output into typed `WorkbookChangeSet`
- validate requested operations before execution
- compute risk class from operation semantics

Primary files to introduce:

- `apps/bilig/src/codex-app/workbook-change-set-builder.ts`

### 4. WorkbookChangeExecutor

Responsibilities:

- execute one atomic change set
- compute authoritative diff
- persist undo payload
- append execution records
- emit publication/update events

Primary files to introduce:

- `apps/bilig/src/zero/workbook-change-executor.ts`
- `apps/bilig/src/zero/workbook-change-diff.ts`
- `apps/bilig/src/zero/workbook-undo-store.ts`

This becomes the only write engine.

### 5. WorkbookReviewQueueService

Responsibilities:

- create review queue items
- record owner approvals and returns
- transition queue item state
- apply approved change sets through the same executor

Primary files to introduce:

- `apps/bilig/src/codex-app/workbook-review-queue-service.ts`
- `apps/bilig/src/zero/workbook-review-queue-store.ts`

### 6. WorkbookExecutionHistoryService

Responsibilities:

- read and persist execution history
- expose replay and undo state
- maintain thread summaries from execution history

Primary files to introduce:

- `apps/bilig/src/codex-app/workbook-execution-history-service.ts`
- `apps/bilig/src/zero/workbook-execution-history-store.ts`

### 7. WorkbookToolRegistry

Responsibilities:

- define the workbook tool surface
- map tool arguments into selectors and change sets
- keep tool responses consistent across direct tools and workflows

Primary files to introduce:

- `apps/bilig/src/codex-app/workbook-tool-registry.ts`
- `apps/bilig/src/codex-app/workbook-tool-read-handlers.ts`
- `apps/bilig/src/codex-app/workbook-tool-write-handlers.ts`

### 8. WorkbookWorkflowRuntime

Responsibilities:

- orchestrate durable workflows only
- produce either a direct change set or a report artifact
- never own a separate write path

Existing file to reshape:

- `apps/bilig/src/codex-app/workbook-agent-workflow-runtime.ts`

## Web Architecture

### Workbook panel state

Current state shaping in `use-workbook-agent-pane.tsx` is too broad.

Split into:

- `use-workbook-agent-session-state.ts`
- `use-workbook-agent-execution-history.ts`
- `use-workbook-agent-review-queue.ts`
- `use-workbook-agent-composer.ts`

### Workbook panel rendering

`WorkbookAgentPanel.tsx` should render three primary content families:

- live conversation
- execution history
- review queue

Private-thread UI:

- execution history rows
- direct rerun
- undo available

Shared-thread UI:

- review queue items
- owner decisions
- recommendation metadata

There should be no private-thread `Apply` and `Dismiss` flow after cutover.

## Persistence Model

### New durable tables

Introduce:

- `workbook_execution_record`
- `workbook_review_queue_item`
- `workbook_review_decision`
- `workbook_undo_payload`

Retain:

- timeline entries
- workflow runs

Delete:

- `workbook_pending_bundle`

### Session snapshot

Session snapshot must contain:

- `executionPolicy`
- `entries`
- `executionRecords`
- `reviewQueueItems`
- `workflowRuns`

Session snapshot must not contain:

- `pendingBundle`
- `approvalMode`

## Cutover Strategy

This design does not allow a long-lived mixed runtime.

Cutover sequence:

1. land new contracts and storage
2. add one-time migration to transform legacy staged state into:
   - execution records
   - review queue items
3. cut runtime reads to the new model only
4. cut runtime writes to the new model only
5. remove legacy endpoints, stores, and UI

Migration handling is one-time data transformation, not runtime fallback behavior.

Post-cutover runtime code does not branch on legacy fields.

## Legacy Deletion Plan

The following code and concepts are scheduled for deletion as part of the implementation:

- `approvalMode` as runtime behavior authority
- `pendingBundle` as runtime mutation state
- `workbook_pending_bundle` table
- preview-first private-thread UI
- apply/dismiss bundle endpoints
- prompt instructions that say `rail`
- preview-specific timeline copy
- duplicate write paths outside the change-set executor

Deletion must occur in the same phase that replaces the old behavior.

## Module Refactor Plan

The current large files should be reduced by responsibility.

### Backend

Refactor:

- `apps/bilig/src/codex-app/workbook-agent-service.ts`
- `apps/bilig/src/codex-app/workbook-agent-tools.ts`
- `apps/bilig/src/codex-app/workbook-agent-workflow-runtime.ts`
- `apps/bilig/src/zero/workbook-chat-thread-store.ts`

Into:

- inspection
- selectors
- change-set building
- execution
- review queue
- execution history
- workflows

### Web

Refactor:

- `apps/web/src/use-workbook-agent-pane.tsx`
- `apps/web/src/WorkbookAgentPanel.tsx`

Into:

- session state
- review queue state
- execution history state
- composer state
- rendering primitives

## Testing Standard

This program requires explicit proof at each layer.

### Unit

- selector resolution
- object snapshot generation
- change-set building
- diff generation
- undo payload capture
- review queue transitions
- execution record persistence

### Integration

- private-thread direct apply from tools
- private-thread replay
- shared-thread owner review
- workflow-generated change sets
- table and validation mutation preservation
- chart and pivot object round-trips

### Browser

- range selection is reported exactly
- style inspection is exact
- private-thread edit applies in one turn
- shared-thread review queue behaves correctly
- execution history renders applied work and rerun actions
- undo and redo are visible and correct

### Correctness and invariants

Every write family requires invariant tests:

- formulas remain correct after row and column insertion
- table references resize correctly
- validation rules stay attached to the intended cells
- conditional formats survive structure edits
- hidden and grouped rows remain coherent
- chart and pivot bindings survive source updates

## Execution Plan

### Phase 1: contracts and selectors

- introduce canonical contracts
- implement selector resolver
- update inspection payloads to semantic snapshots

Exit gate:

- reads and tool schemas operate on selector and snapshot contracts only

### Phase 2: single change-set executor

- implement `WorkbookChangeSet`
- implement `WorkbookChangeExecutor`
- implement `WorkbookChangeDiff`
- implement undo payload persistence

Exit gate:

- all direct writes use the executor

### Phase 3: execution history and review queue

- implement execution history store
- implement review queue store
- migrate session snapshots

Exit gate:

- execution records and review queue items are the only durable mutation models

### Phase 4: tool surface expansion

- add missing structural tools
- add sort/filter
- add validation
- add named range and table CRUD
- add comments and notes

Exit gate:

- the agent can complete typical operational workbook tasks end to end

### Phase 5: analytics and advanced workbook objects

- add charts
- add pivot tables
- add sparklines
- add slicers
- add media and protection

Exit gate:

- dashboard and reporting workbooks are fully operable by the agent

### Phase 6: panel rewrite and cutover

- rewrite the workbook panel around execution history and review queue
- remove private-thread review cards
- remove preview-first language

Exit gate:

- private threads are execution-first in UI and runtime

### Phase 7: legacy deletion

- delete legacy tables, endpoints, models, and code paths
- remove stale tests and fixtures

Exit gate:

- no runtime code references preview-first mutation state

## Acceptance Criteria

The program is complete when all of the following are true:

- private-thread work executes directly under session policy
- shared-thread review is explicit and isolated to review queue items
- the agent can inspect workbook structure, style, validation, collaboration, and analytics objects directly
- all workbook writes use one change-set executor
- execution history and undo are first-class product surfaces
- `approvalMode` and `pendingBundle` are absent from runtime behavior and durable writes
- no user-facing copy references rails or preview bundles
- all affected packages pass typecheck, lint, focused unit tests, browser tests, and invariant tests

## Final Standard

The final system should behave like a serious spreadsheet operator, not a cell-edit macro runner.

That means:

- semantic targeting
- rich workbook reads
- exact writes
- atomic execution
- durable history
- review as a policy state
- zero live legacy mutation code

This is the standard for implementation.
